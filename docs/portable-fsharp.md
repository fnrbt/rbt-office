# F# that compiles well to .NET and JavaScript

The same F# runs on .NET and, through Fable, in the browser. Most of it compiles to good code on
both, but a few spellings that are free on .NET become library calls, allocations or interpreters
in JavaScript. This note lists them, cheapest replacement first. The helpers are in
`src/Office.Formats/src/Portable.fs` (module `Portable`, referenced by Word and Spreadsheet).
Each helper's .NET branch is exactly the spelling it replaces, so .NET behaviour and speed stay as
they were.

Only change code that a profile shows is hot (see *Measuring* below), or where the replacement is a
mechanical one-for-one. The list comes from CPU profiles of opening, scrolling, editing and typing,
2026-11-18 (Fable 4.28), checked against the output of Fable 5.27 (`.config/dotnet-tools.json`) on the
upgrade, 2026-11-19: what Fable 5 changed is said where it applies (*Fable 5* below).

## Helpers

| Instead of | Write | Why (in JavaScript) |
|---|---|---|
| `xs.RemoveAt (xs.Count - 1)` | `Portable.Stack.dropLast xs` | `RemoveAt` is `splice(i, 1)`, which allocates the removed items as an array: ~17x slower than `pop()`. Every explicit-stack walk (formula trees, the XML parser, the dependency graph) did this per node. |
| `let x = xs.[last] in xs.RemoveAt last` | `let x = Portable.Stack.pop xs` | The same. |
| `match d.TryGetValue k with \| true, v -> ... \| false, _ -> ...` | `match Portable.Dict.tryFind k d with \| ValueSome v -> ... \| ValueNone -> ...` | The out-parameter becomes a reference cell plus a tuple (Fable 4 also gave the cell two closures; Fable 5 a plain `FSharpRef(0)`), and the key is looked up twice (`has`, then `get`). `tryFind` is one `Map.get` and allocates nothing (voption is erased in Fable; on .NET it is a struct). |
| `s.Replace("a", "b")` (string arguments) | `s \|> Portable.Str.replace "a" "b"` | Fable's Replace builds a `RegExp` from an escaped copy of the pattern on every call (and reads `$&`, `$1` in the replacement as patterns). `replace` is `indexOf`, then `split`/`join`. The pattern must not be empty. |
| `sprintf "%s %fpx" a x`, `String.Format` | `a + " " + Portable.Num.toFixed 6 x + "px"`, `string i` for `%d` | sprintf is a format-string interpreter in JavaScript, run on every call. `string` of an int and `+` are native. `Num.toFixed` is `%.Nf` exactly (`toFixed` in JavaScript, `F<N>` invariant on .NET). Keep sprintf for messages and anything off the hot path. |

## Conventions

- **Descending and stepped loops**: `for i = hi downto lo do` and `for i in lo .. hi do` compile to
  plain loops, and under Fable 5 so does `for i in hi .. -1 .. lo`. A step other than 1 or -1
  (`for i in lo .. 2 .. hi`) is still a range enumerator (`rangeDouble`) with a call per step (Fable 4:
  -1 too). Comprehensions are enumerators whatever the step: `[| for i in 0 .. n - 1 -> f i |]` is
  `toArray(delay(map(rangeDouble ...)))`; `Array.init n f` is a loop.
- **Long F# lists**: walking a list in Fable costs a call per step (`.Head` / `.Tail` are
  `FSharpList__get_Head` / `get_Tail`, which check for the empty list; under Fable 4 `.Tail` went
  through `Option.value`), and the profile of opening a long document had `List.js` and `Option.js` as
  its largest library costs (Fable 4). Don't re-walk a long list on every pass: keep the result (by the
  list's reference: lists are immutable), or keep an array beside it.
- **`int64` / `uint64`**: are `BigInt` in JavaScript (Fable 4 and 5) - every operation a library call and a
  heap value. Use `int` where the range allows; for an integral `float` shown as text with
  `abs x < 1e15`, `(int64 x).ToString()` can be the float's own text (JavaScript prints
  integral doubles below 1e21 without an exponent). Measured: not hot in the interactions
  profiled (`BigInt.js` < 0.1 %), so the existing uses are left.
- **`decimal`**: is a library type in JavaScript (`Decimal.js`); keep it to money formatting off hot
  paths. (Rbt.Vb's Currency is one: under Fable 4.28 its arithmetic failed in the browser with
  `x.div is not a function`; under 5.27 it works - `macro-currency` in `tools/xlsx-browser/features.js`.)
- **`Math.Round`**: Fable 5 honours `MidpointRounding` (`roundToMode`); Fable 4 ignored it.
- **Structural equality and comparison**: `a = b` and `compare a b` on records, unions and
  tuples are generic library calls that walk every field (`recordEquals`, `Util.compare`). A
  `Map`/`Set` keyed by a record or tuple compares structurally at every node of the tree. For
  hot keys give the type `[<CustomEquality; CustomComparison>]` with a hand-written `CompareTo`
  (see `CellAddress` in `src/Spreadsheet/Model.fs`), key a `Dictionary` by an `int` (`row * 16384 + col`),
  or test `obj.ReferenceEquals(a, b) ||` first when the two are often the same object.
- **Fieldless union cases**: under Fable 5 a case without fields (`Empty`, `None`-like cases of
  your own unions) is one shared object (`static Empty = new CellValue(7, [])`); Fable 4 made a
  new one at every use. Nothing to write differently any more; a hand-made shared value is no longer
  needed for speed.
- **Strings indexed by position**: under Fable 5 `s.[i]` is `getCharAtIndex(s, i)`, a call that checks
  the bounds and throws as .NET does (Fable 4: `s[i]`, `undefined` beyond the end). In a scanner that
  reads every character, `s.[i]` costs a call per character; leave it unless a profile shows it.
- **Lists across bundles**: each bundle of `dist` has its own fable-library, and Fable 5's list
  equality and comparison test `instanceof FSharpList`, so a list made by the office-formats bundle
  was never equal to one of the spreadsheet bundle (a host that reads with one and imports with the
  other: `InstallImportedWorkbook`). `scripts/patch-fable-library.mjs` (run by `build:fable`)
  restores Fable 4's walk of any list. Records, unions, `Map` and `Set` of two bundles were never
  equal (their tests compare constructors): convert reader values into the model's own types.
- **Function values in options and records**: Fable keeps a function stored in an option, a
  record or a module value curried, and wraps it (`uncurry5`, plus a `WeakMap` entry) each time
  it is passed where an uncurried function is expected. In a per-cell or per-glyph call, pass a
  lambda written at the call, or take the function as a plain argument.
- **JS objects built for interop**: `createObj (fieldsA @ fieldsB)` builds and copies a list
  of pairs per object; a literal `createObj [ ... ]` becomes an object literal. For an object
  with optional parts, `let o = obj ()` then `o?Name <- value` per field (in the same order) is
  cheap. Reading such an object, an `inline` reader with a literal name (`o?Name`) is a cached
  property load per call site; one function reading `o?(name)` for every name is not.
- **Exceptions as control flow**: a `try ... with` that is hit often is slow on both backends
  (and in JavaScript the exception object captures a stack). Use options/`Result`, or
  `TryParse`-style functions.
- **Timers**: `setTimeout(f, 0)` is clamped to 4 ms after a few nested calls. Use a
  `MessageChannel` or `queueMicrotask` (or `requestIdleCallback` for background work) to chain
  slices.
- **Don't name the backend's types in shared code.** Shared code names what it means (`int` for
  counts and indices, `float` for measures); widths that only one backend has cheaply (`int64`,
  `decimal`, `bigint`) stay at the edges, where a value is parsed or written.

## Writing a helper

A helper is `inline`, with `#if FABLE_COMPILER` around a native expression (`emitJsExpr`) and
the replaced spelling in the `#else`. Two traps:

- Fable 5 coerces an emitted expression whose F# type is `int` with `| 0` (and so on for the
  other integer types; Fable 4.28 left at least a non-inline `emitJsExpr` as it was): an `undefined` meaning
  "not there" becomes `0`. Type such an expression `obj` and unbox once it is known to be there
  (see `Dict.tryFind`). Fable 5 also adds `| 0` to the result of a library call typed `int`
  (`List.fold f 0 xs | 0`).
- Arguments substituted into the emitted text more than once are evaluated more than once:
  bind them with `let` first.

## Measuring

The browser harnesses print the functions with the most self time; with `PROFILE_OUT=file.json`
they also save the raw CPU profile of that phase for further analysis:

```sh
PLAYWRIGHT=.../playwright/index.mjs PROFILE_OUT=/tmp/load.json node tools/word-browser/perf.mjs --profile load
PLAYWRIGHT=... PROFILE_OUT=/tmp/scroll.json node tools/xlsx-browser/perf.mjs --profile scroll --steps 0 tests/real-world/dashboard-sales-analysis.xlsx
```

Library time is the samples whose function is in `fable_modules/fable-library-js.*` (the folder is
named by the Fable version: `fable-library-js.5.27.0`; don't write the version into a check, find it
as `tools/word-browser/checks/xml-scan.js` does); attribute
each to the nearest caller outside the library to find the F# function (Fable keeps F# names;
`Foo__Bar_Z...` is member `Bar` of type `Foo`).
