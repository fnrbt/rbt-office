#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open System.Diagnostics
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-26s  %s" label detail)

// --- Perf 1: deep dependency chain --------------------------------------
// A1 = 1, A2 = A1+1, … An = A(n-1)+1. Recalc must topo-sort and evaluate
// the whole chain in one pass; the answer at the tail is n.
let chainLen = 5000
let chainCells =
    [ for r in 0 .. chainLen - 1 ->
        let v =
            if r = 0 then CellValue.Formula ("=1", CellValue.Empty)
            else CellValue.Formula (sprintf "=A%d+1" r, CellValue.Empty)
        { Row = r; Col = 0 }, cell v ]
    |> Map.ofList
let chainModel = { createEmpty (chainLen + 10) 10 with Cells = chainCells }

let sw = Stopwatch.StartNew()
let chainResult = recalculateAllFormulas None None chainModel
sw.Stop()
let chainMs = sw.Elapsed.TotalMilliseconds
let tailVal =
    match getCellValue { Row = chainLen - 1; Col = 0 } chainResult with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "chain-correct" (tailVal = CellValue.Number (float chainLen))
   (sprintf "tail expected %d got %A" chainLen tailVal)
ok "chain-perf" (chainMs < 2000.0) (sprintf "%d-cell chain took %.0f ms (budget 2000)" chainLen chainMs)
printfn "  deep chain: %d cells in %.0f ms" chainLen chainMs

// --- Perf 2: wide aggregate over a large range --------------------------
// B1..B20000 = 1, and C1 = SUM(B1:B20000). The range reference must not
// blow up — recalc should stay well under budget.
let wideN = 20000
let wideCells =
    [ for r in 0 .. wideN - 1 -> { Row = r; Col = 1 }, cell (CellValue.Number 1.0) ]
    @ [ { Row = 0; Col = 2 }, cell (CellValue.Formula (sprintf "=SUM(B1:B%d)" wideN, CellValue.Empty)) ]
    |> Map.ofList
let wideModel = { createEmpty (wideN + 10) 10 with Cells = wideCells }

let sw2 = Stopwatch.StartNew()
let wideResult = recalculateAllFormulas None None wideModel
sw2.Stop()
let wideMs = sw2.Elapsed.TotalMilliseconds
let sumVal =
    match getCellValue { Row = 0; Col = 2 } wideResult with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "wide-correct" (sumVal = CellValue.Number (float wideN))
   (sprintf "SUM expected %d got %A" wideN sumVal)
ok "wide-perf" (wideMs < 2000.0) (sprintf "wide SUM took %.0f ms (budget 2000)" wideMs)
printfn "  wide aggregate: SUM over %d cells in %.0f ms" wideN wideMs

// --- Perf 3: re-evaluating a single formula many times ------------------
// Tight loop over evaluateFormula — catches per-eval allocation regressions.
let evalModel = { createEmpty 100 100 with Cells = wideCells }
let iters = 20000
let sw3 = Stopwatch.StartNew()
for _ in 1 .. iters do
    evaluateFormula "=B1+B2*B3-B4/B5" evalModel None |> ignore
sw3.Stop()
let perEvalUs = sw3.Elapsed.TotalMilliseconds * 1000.0 / float iters
ok "eval-perf" (perEvalUs < 200.0) (sprintf "%.1f us/eval (budget 200)" perEvalUs)
printfn "  single-formula eval: %.1f us/eval over %d iters" perEvalUs iters

// --- Correctness: cycle detection does not hang -------------------------
let cycleCells =
    [ { Row = 0; Col = 5 }, cell (CellValue.Formula ("=F2+1", CellValue.Empty))
      { Row = 1; Col = 5 }, cell (CellValue.Formula ("=F1+1", CellValue.Empty)) ]
    |> Map.ofList
let cycleModel = { createEmpty 100 100 with Cells = cycleCells }
let cycleResult = recalculateAllFormulas None None cycleModel
let f1 =
    match getCellValue { Row = 0; Col = 5 } cycleResult with
    | CellValue.Formula (_, v) -> v | v -> v
// (Not calculated, as in Excel without iterative calculation: the value it had - none, 0.)
ok "cycle-detected" (f1 = CellValue.Number 0.0)
   (sprintf "expected 0 (a formula on a circular reference is not calculated), got %A" f1)

printfn ""
printfn "Stage 4 perf/correctness tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
