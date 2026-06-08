# Calc engine: streaming + fusion design note

Status: research synthesis → recommendation → **v1 landed** for aggregates. 2026-06-07.

## Implemented (v1)
`src/Spreadsheet/Formula.fs` now has the streaming primitives from this note:
`RangeView` (a `[<Struct>]` pull enumerator over the sparse cell map),
`Fold<'T,'S,'R>` (one-pass reducer), `run`, `both` (two folds in one
traversal), `numAggFold`/`countFold`, and `RangeAccess`/`rangeRead` (index
access for parallel-range zip). `foldArgs` streams a fold across an aggregate's
whole argument list without building a list.

Migrated to streaming (no per-range `CellValue list`):
- **Aggregates** — SUM, PRODUCT, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK
  (AVERAGE = SUM+COUNT in one pass, via `foldArgs`/`numAggFold`).
- **Criteria family** — COUNTIF, SUMIF, AVERAGEIF, SUMIFS, COUNTIFS, MAXIFS,
  MINIFS, AVERAGEIFS (parallel ranges zipped by index via `rangeRead`).
- **Lookup family** — MATCH, VLOOKUP, HLOOKUP, XLOOKUP, XMATCH (scans stream
  the lookup line and read the result cell directly; `xMatchIndex` streams).

Still on the list path (need all values / sorting): MEDIAN, STDEV, STDEVP, VAR,
LARGE, SMALL, SUMPRODUCT.

### Range-granular dependencies (Calc.fs)
Resolves open question #1: the per-cell-edge reification was a real, severe
bottleneck. A single `SUM(A1:A100000)` measured **778 ms / 100k graph edges /
~29 MB / 210 ms per edit**. Fix: ranges larger than `rangeDepThreshold` (2048
cells) are tracked as a `DepRect` (rectangle) instead of one edge per cell
(`Formula.referencesRanged`; `Calc.CalcMetadata.DepRects`). The **rebuilder**
dirties a rect-subscriber when an edited cell is contained in its rectangle
(`markDirtyForCell` containment scan, no reverse edges); the **scheduler**
(`topologicalSort` via `rebuildCalcChain.getDeps`) adds ordering edges only to
the *formula* cells inside a rect (few), so a large aggregate still computes
after formulas it covers — without enumerating data cells. Same formula now:
**36 ms / 0 edges / ~0 MB / 21 ms per edit**, identical results. Small ranges
(≤ threshold) keep the explicit-edge path unchanged. Verified:
`tests/formula/stage5-calc.fsx` (16 assertions incl. in-rect edit → dirty →
recalc).

### Dynamic arrays (Batch 2, on these primitives)
`ArrVal` (2-D value) + `evalArray` + operator **broadcasting**; functions
SEQUENCE/RANDARRAY/UNIQUE/SORT/SORTBY/FILTER/TRANSPOSE. Arrays are consumed by
the streaming aggregates via the same `foldArgs` (e.g. `SUM(FILTER(…))` folds
the produced array — no separate path), collapse to top-left in scalar context,
and **spill** into neighbouring cells for rendering (`spillArrays`, `#SPILL!`
on collision). Verified: `stage8-arrays.fsx` (25) + headless render of
`synthetic-dynamic-arrays.xlsx` (SORT/UNIQUE/FILTER/SEQUENCE incl. a 2-D spill).

Verified overall: `tests/formula` stages 1–8, **269 assertions**, all green;
behavior byte-identical to the old paths for non-array formulas. New functions
render correctly headlessly. Next: interactive spill (clear-on-recalc +
spill-cell deps in Calc.fs), canvas/PDF `#SPILL!`, LAMBDA + first-class
functions, and open questions #2–#4 below.

Motivation: today the calc engine has cell-level incremental recalc (dependency
graph + dirty propagation + topological recompute, sparse
`Map<CellAddress, CellData>`), but **value computation is eager**:
`expandRangeRef` / `rangeValues` / `flattenArgs` materialize a `CellValue list`
for every range on every eval (`SUM(A1:A100000)` → 100k-element list;
`SUM(A:A)+AVERAGE(A:A)` → two traversals, no sharing), and range references
reify **one dependency-graph edge per cell**. We want to stop reifying large
parts of the sheet into RAM and to make the upcoming dynamic-array spill
(FILTER/SORT/UNIQUE/SEQUENCE) streaming-native — while staying simple and
F#-idiomatic.

This note is the output of a deep-research pass (sources cited inline; all
theory claims came back 3-0 on adversarial verification except the
Applicative/Monadic split, 2-1).

## The unification (three theories, one picture)

The key realization: **incremental dependency tracking and fused value
computation are orthogonal axes**, and our engine already lives on the first.

1. **Spreadsheet recalc *is* a build/dataflow system.** *Build Systems à la
   Carte* (Mokhov/Mitchell/Peyton Jones, ICFP'18/JFP'20) models Excel as
   `excel :: Build Monad …` and shows every incremental engine factors into a
   **scheduler** (what order to (re)compute) × **rebuilder** (whether a cell
   needs recompute) — independently recombinable. Our topological recompute =
   scheduler; our dirty-bit graph = rebuilder. *We are already this design.*
   [build-systems-final.pdf]
2. **Dependencies split Applicative vs Monadic.** Static deps readable from the
   formula (`SUM(A1:A100)`, arithmetic) vs dynamic deps that depend on computed
   values (`INDIRECT`, `OFFSET`). The Applicative end is exactly where one-pass
   **applicative folds** apply; the Monadic end is where deps must be
   re-discovered after computing intermediates (Excel uses a restarting
   scheduler for this; we use topological — a tunable knob, not a fixed cost).
   [build-systems-final.pdf]
3. **Demand-driven incrementality (Adapton).** Only recompute *observed*
   outputs; forced results aren't recomputed until re-forced. Adapton's own
   running example is a spreadsheet with two panes where only the visible one
   propagates. → off-screen cells, hidden panes, and **un-displayed spill
   ranges should cost nothing** until demanded. [Adapton, PLDI'14]
4. **Streaming fusion is the value-computation axis.** Deforestation (Wadler) /
   `foldr`+`build` short-cut fusion (Gill/Launchbury/Peyton Jones) cancel
   intermediate "glue" lists via one local rewrite — but **don't fuse left
   folds or zips**. *Stream fusion* (Coutts/Leshchinsky/Stewart, ICFP'07)
   represents a stream as a non-recursive stepper `s -> Step a s` with a `Skip`
   constructor and fuses the **wider class spreadsheets need: left folds
   (sum/avg accumulators) and zips (combining two ranges)**. *strymonas*
   (Kiselyov et al., POPL'17) achieves *complete* deforestation (zero main-loop
   allocation) via staging — the high ceiling, not the v1 tool.
   [stream-fusion.pdf; 1612.06668]
5. **Push vs pull is not a free choice.** Pull is required for **zip / parallel
   multi-range consumption** (`MAXIFS`, `SUMPRODUCT`, any two-range operator);
   push "stumbles on zip — cannot do it at constant per-element cost." Push
   excels at filter/fan-out. strymonas therefore takes **pull as the base**.
   [1612.06668]
6. **.NET reality.** Naive `seq`/`IEnumerable` pull pays **3 virtual calls per
   element** (`MoveNext`, `Current`, the lambda); library-assembled streams
   (push or pull) can be up to ~2 orders of magnitude slower than hand loops
   from residual tuple/closure overhead. → on the hot path use **struct
   enumerators + `inline`/SRTP**, not plain `seq`. The real cost of today's
   eager path is *allocation + multi-pass*, which fusion removes.
   [Clash of the Lambdas, ICOOOLPS'14; 1612.06668]
7. **Transducers / collection operators are the decoupled surface.** A
   transducer (`reducing-fn -> reducing-fn`, composed with `>>`) is independent
   of input/output, runs single-pass with no intermediate aggregates, and the
   *same* transducer drives pull and push. Naiad/differential dataflow shows a
   LINQ-shaped (`SelectMany`/`GroupBy`) surface can drive an incremental engine
   on .NET — validates a LINQ/transducer-shaped operator API. [Clojure
   transducers; Naiad SOSP'13]

## Recommendation (simple, F#-native, efficient)

A **hybrid** that changes value computation only and leaves the incremental
layer intact:

### 1. Producer — lazy pull `RangeView` (struct enumerator over the sparse map)
Never a `CellValue list`, never plain `seq` on the hot path.

```fsharp
/// A lazy rectangular window over a sheet's sparse cells. Yields unwrapped
/// CellValues row-major without materializing a list. Two iteration modes:
///   dense  — every cell incl. blanks (CellValue.Empty), for positional/zip ops
///   sparse — only populated cells, for blank-ignoring aggregates (SUM/COUNT)
[<Struct>]
type RangeView =
    { Model: SpreadsheetModel; SR: int; SC: int; ER: int; EC: int }
    member v.GetEnumerator() = RangeEnumerator(v)   // struct, pattern-based foreach

and [<Struct>] RangeEnumerator = …  // mutable r,c; MoveNext/Current are struct methods
```

`inline` consumers + struct enumerator let the JIT inline the per-element step —
no `IEnumerable` virtual dispatch, no per-element allocation. (Refinement: the
*sparse* mode iterates only the map entries inside the rectangle, so `SUM(A:A)`
costs O(populated) not O(1,048,576).)

### 2. Reducer — composable one-pass applicative folds (the headline win)
`SUM(A:A)+AVERAGE(A:A)` becomes **one traversal** sharing accumulators.

```fsharp
[<Struct>]
type Fold<'T,'S,'R> = { Init: 'S; Step: 'S -> 'T -> 'S; Extract: 'S -> 'R }

let inline run (v: RangeView) (f: Fold<_,_,_>) =
    let mutable s = f.Init
    let mutable e = v.GetEnumerator()
    while e.MoveNext() do s <- f.Step s e.Current
    f.Extract s

/// Run two folds over ONE traversal (state stays a struct → no boxing).
let inline both (a: Fold<'T,'SA,'RA>) (b: Fold<'T,'SB,'RB>) =
    { Init = struct(a.Init, b.Init)
      Step = fun (struct(x,y)) t -> struct(a.Step x t, b.Step y t)
      Extract = fun (struct(x,y)) -> struct(a.Extract x, b.Extract y) }

let sumF   = { Init = 0.0; Step = (fun s x -> s + num x); Extract = id }
let countF = { Init = 0;   Step = (fun n x -> if isNum x then n+1 else n); Extract = id }
let averageF =                               // SUM and COUNT in one pass
    let c = both sumF countF
    { Init = c.Init; Step = c.Step
      Extract = fun s -> let struct(sm,n) = c.Extract s in if n=0 then div0 else sm/float n }
```

`MAXIFS`/`SUMIFS`/`SUMPRODUCT` = a fold whose `Step` pulls the parallel
criteria/weight ranges via **pull-side zip** (cheap; the reason pull is the
base). This is strictly simpler than full stream fusion and gets the single-pass
multi-aggregate win immediately. [foldl library; haskellforall composable folds]

### 3. Spill — transducer stages, lazy until forced
- `FILTER` / `UNIQUE` are filter/dedup **transducer stages** composed onto the
  fold/producer; they stream (UNIQUE keeps a `HashSet`, still one pass).
- `SEQUENCE` is a pure producer; `SORT` is a **barrier** producer (a sort can't
  stream — it must materialize; call that out, don't pretend otherwise).
- A spill result stays a **lazy `RangeView` placed at the anchor**, materialized
  to cells only when displayed/forced (Adapton demand-driven) → off-screen
  spills are free.

### 4. Incremental layer — unchanged
Keep the topological scheduler + dirty-graph rebuilder; they are already the
Build-Systems-à-la-Carte factoring, and fusion is **orthogonal** (it changes how
a cell's value is computed, not how dependencies are tracked).

### 5. Escape hatch — defer
strymonas-style staged codegen / LinqOptimizer-style runtime fusion gave F# 14×
on `sumOfSquares` historically, but modern .NET (`Span<T>`, ref structs, JIT
inlining) narrows the gap. Reach for it **only if profiling** shows the struct
enumerator + applicative-fold path is still hot. [LinqOptimizer; strymonas]

## Tradeoffs / caveats
- **SORT can't be fused** (inherent barrier). FILTER-of-FILTER (nested spill) is
  the case most likely to *lose* fusion — `concatMap`/nested-stream fusion is the
  known hard case (didn't fuse in GHC for years). Verify nesting specifically.
- Two recommendation pieces are **synthesized**, not directly cited: the
  applicative-fold encoding (mechanism inferred from the transducer/stream-fusion
  sources) and the `RangeView`-over-sparse-map struct enumerator (our
  application). Everything else rests on peer-reviewed primary sources.
- Performance multipliers in the sources are 2014–2017 (OCaml/Scala/Java/.NET);
  the *qualitative* lesson (avoid allocation + multi-pass + virtual dispatch)
  transfers, exact numbers won't — **benchmark on 1k / 100k / full-column**
  before adopting any codegen.

## Open questions (decide before/early in Batch 2)
1. **Where's the real bottleneck today** — list allocation/GC, double traversal,
   or per-cell range-edge reification? If whole-column deps dominate, we also
   need a **range-granular dependency** (a rectangular subscription, à la
   differential dataflow) — not just fused values.
2. **Structural-edit incrementality for spill** (insert/delete rows feeding
   SORT/FILTER): is cell-level dirtying enough, or do we need Adapton's
   reorderable DCG / differential deltas? (Simple-vs-powerful line for
   SORT/UNIQUE is unsettled.)
3. **INDIRECT/OFFSET × fusion**: discovering a Monadic dependency mid-fold —
   restart (Excel-style) or defer?
4. **Does modern .NET close the abstraction gap** enough that struct-enumerator
   + applicative-fold matches hand loops, making staged codegen permanently
   unnecessary here? Needs a microbenchmark.

## Sources
- Build Systems à la Carte — https://www.microsoft.com/en-us/research/wp-content/uploads/2018/03/build-systems-final.pdf
- Adapton (PLDI'14) — https://www.cs.tufts.edu/~jfoster/papers/cs-tr-5027.pdf
- Incremental λ-calculus / theory of changes — https://arxiv.org/abs/1312.0658
- Naiad / timely & differential dataflow (SOSP'13) — https://dl.acm.org/doi/10.1145/2517349.2522738
- Clojure transducers — https://clojure.org/reference/transducers
- Blelloch, collection-oriented (NESL) — https://www.cs.cmu.edu/~guyb/papers/collections.pdf
- Stream Fusion (ICFP'07) — https://www.cs.tufts.edu/~nr/cs257/archive/duncan-coutts/stream-fusion.pdf
- Stream Fusion, to Completeness / strymonas (POPL'17) — https://arxiv.org/pdf/1612.06668
- Clash of the Lambdas (ICOOOLPS'14) — https://yanniss.github.io/lambdas-icooolps14.pdf
- foldr/build short-cut deforestation — https://dl.acm.org/doi/10.1145/165180.165214
- Wadler, deforestation — https://homepages.inf.ed.ac.uk/wadler/topics/deforestation.html
- Control.Foldl (composable folds) — https://hackage.haskell.org/package/foldl · https://haskellforall.com/2013/08/composable-streaming-folds
- Nessos Streams (F#/.NET push fusion) — https://github.com/nessos/Streams
- Push vs pull loop fusion in query engines (JFP) — https://www.cambridge.org/core/journals/journal-of-functional-programming/article/push-versus-pullbased-loop-fusion-in-query-engines/D67AE4899E87F4B5102F859B0FC02045
- High-performance F# (struct enumerators, inlining) — https://www.bartoszsypytkowski.com/writing-high-performance-f-code/
