#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

open System.Diagnostics
open System.Reflection
open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Formula
open RedBat.Spreadsheet.Layout
open RedBat.Spreadsheet.CellText
open RedBat.Spreadsheet.Paint

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-26s  %s" label detail)

let measuredWrap =
    wrapLines (Some (fun (s: string) -> float s.Length)) 24.0 11.0 "alpha beta gamma"
ok "wrap-measured-boundary" (measuredWrap = [ "alpha beta"; "gamma" ])
   (sprintf "expected exact measured wrap got %A" measuredWrap)

let sectionFormat = { defaultFormat with NumberFormat = Some "\"semi;colon\";[Red]0.0;\"zero\"" }
let sectionText value = cellDisplayText None 120.0 sectionFormat (CellValue.Number value)
let sectionActual = [ for _ in 1 .. 3 -> sectionText 1.0, sectionText -2.5, sectionText 0.0 ]
let sectionExpected = [ for _ in 1 .. 3 -> "semi;colon", "2.5", "zero" ]
ok "format-section-split-cache"
   (sectionActual = sectionExpected)
   "custom number-format sections with quoted semicolons should be stable"
let largeNumberFormat index =
    "\"" + System.String('x', 5000) + sprintf "%04d" index + "\";[Red][<0]0.00;0.00"
let largeNumberFormatRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for index in 0 .. 1999 do
        let code = largeNumberFormat index
        let format = { defaultFormat with NumberFormat = Some code }
        cellDisplayText None 120.0 format (CellValue.Number 1.0) |> ignore
        numberFormatColorForValue code -1.0 |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "format-cache-bounded-source-memory"
   (largeNumberFormatRetainedBytes < 40000000L)
   (sprintf "large number-format caches retained %.1f MB" (float largeNumberFormatRetainedBytes / 1024.0 / 1024.0))
let oversizedNumberFormatWeakReference () =
    let code = "[Red];" + System.String('x', 5100000)
    numberFormatColorForValue code -1.0 |> ignore
    System.WeakReference(code)
let oversizedNumberFormatWeak = oversizedNumberFormatWeakReference ()
let oversizedNumberFormatCollected =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    not oversizedNumberFormatWeak.IsAlive
ok "format-cache-skips-oversized-entry"
   oversizedNumberFormatCollected
   "oversized number-format key should not be retained by caches"
let derivedOversizedNumberFormatWeakReference () =
    let payload = System.String('x', 1300000)
    let code = "[Red]" + payload + ";" + payload
    numberFormatColorForValue code 1.0 |> ignore
    System.WeakReference(code)
let derivedOversizedNumberFormatWeak = derivedOversizedNumberFormatWeakReference ()
let derivedOversizedNumberFormatCollected =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    not derivedOversizedNumberFormatWeak.IsAlive
ok "format-cache-skips-derived-entry"
   derivedOversizedNumberFormatCollected
   "number-format cache should count generated section text, not only the source key"
let manyEmptySectionsFormatWeakReference () =
    let code = System.String(';', 220000)
    let format = { defaultFormat with NumberFormat = Some code }
    cellDisplayText None 120.0 format (CellValue.Number 1.0) |> ignore
    System.WeakReference(code)
let manyEmptySectionsFormatWeak = manyEmptySectionsFormatWeakReference ()
let manyEmptySectionsFormatCollected =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    not manyEmptySectionsFormatWeak.IsAlive
ok "format-cache-skips-many-empty-sections"
   manyEmptySectionsFormatCollected
   "split-section cache should count retained section-list nodes, not only text"

let conditionalFormatCode = "[Red][<0](#,##0.00);[Blue][>=1000]#,##0.00;0.00"
let conditionalFormat = { defaultFormat with NumberFormat = Some conditionalFormatCode }
let conditionalText value = cellDisplayText None 120.0 conditionalFormat (CellValue.Number value)
let conditionalActual =
    [ for _ in 1 .. 3 ->
        conditionalText -2.5,
        conditionalText 1250.0,
        conditionalText 10.0,
        numberFormatColorForValue conditionalFormatCode -2.5,
        numberFormatColorForValue conditionalFormatCode 1250.0,
        numberFormatColorForValue conditionalFormatCode 10.0 ]
let conditionalExpected =
    [ for _ in 1 .. 3 ->
        "(2.50)", "1,250.00", "10.00", Some "#FF0000", Some "#0000FF", None ]
ok "format-condition-cache"
   (conditionalActual = conditionalExpected)
   "conditional number-format sections and colours should be stable"

let repeatedLongColorFormat =
    "[Red][<0]" + String.replicate 5200 "0" + ";[Blue][>=1000]" + String.replicate 120 "#" + "0;0"
numberFormatColorForValue repeatedLongColorFormat -12.34 |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let repeatedLongColorSw = Stopwatch.StartNew()
let mutable repeatedLongColorHits = 0
for _pass in 1 .. 100 do
    for i in 1 .. 5000 do
        match numberFormatColorForValue repeatedLongColorFormat (float -i) with
        | Some _ -> repeatedLongColorHits <- repeatedLongColorHits + 1
        | None -> ()
repeatedLongColorSw.Stop()
ok "format-color-large-cache-hit-perf"
   (repeatedLongColorSw.Elapsed.TotalMilliseconds < 650.0 && repeatedLongColorHits = 500000)
   (sprintf "500k color lookups took %.1f ms, hits=%d" repeatedLongColorSw.Elapsed.TotalMilliseconds repeatedLongColorHits)

let manyFormatTexts () =
    [ for i in 0 .. 1199 ->
        let format =
            { defaultFormat with
                NumberFormat = Some (sprintf "[>=%d]#,##0.00;[Red](#,##0.00);0.00" (i + 1)) }
        cellDisplayText None 120.0 format (CellValue.Number 1234.56) ]
let manyFormatFirst = manyFormatTexts ()
let manyFormatSecond = manyFormatTexts ()
ok "format-cache-many-distinct"
   (manyFormatFirst = manyFormatSecond && manyFormatFirst.Length = 1200)
   "distinct custom format cache entries should stay stable past the old clear threshold"

let nestedFormulaDepth = 2000
let nestedFormula =
    "=SUM("
    + String.replicate nestedFormulaDepth "TRANSPOSE("
    + "SEQUENCE(10)"
    + String.replicate nestedFormulaDepth ")"
    + ")"
let nestedFormulaSw = Stopwatch.StartNew()
let nestedFormulaResult = evaluateFormula nestedFormula (createEmpty 1 1) None
nestedFormulaSw.Stop()
ok "parse-nesting-safe-error"
   (nestedFormulaResult = CellValue.Error "Formula nesting too deep")
   (sprintf "expected safe nesting error, got %A" nestedFormulaResult)
ok "parse-nesting-safe-perf"
   (nestedFormulaSw.Elapsed.TotalMilliseconds < 500.0)
   (sprintf "deeply nested formula should fail before stack exhaustion, took %.0f ms" nestedFormulaSw.Elapsed.TotalMilliseconds)
let touchOversizedFormula () =
    let formula = "=" + System.String('A', 5100000)
    evaluateFormula formula (createEmpty 1 1) None |> ignore
let oversizedFormulaRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    touchOversizedFormula ()
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "parse-cache-skips-oversized-entry"
   (oversizedFormulaRetainedBytes < 8000000L)
   (sprintf "oversized formula cache retained %.1f MB" (float oversizedFormulaRetainedBytes / 1024.0 / 1024.0))
let derivedOversizedFormulaWeakReference () =
    let formula = "=\"" + System.String('F', 3000000) + "\""
    evaluateFormula formula (createEmpty 1 1) None |> ignore
    System.WeakReference(formula)
let derivedOversizedFormulaWeak = derivedOversizedFormulaWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "parse-cache-skips-derived-entry"
   (not derivedOversizedFormulaWeak.IsAlive)
   "parse cache should count retained AST text, not only the formula source"

let subtotalProjectedFilterRows = 100000
let subtotalProjectedFilterFormula = "=SUBTOTAL(9,A1:A100000)"
let subtotalProjectedFilterModel =
    { createEmpty subtotalProjectedFilterRows 2 with
        Cells =
            [ for row in 0 .. subtotalProjectedFilterRows - 1 ->
                { Row = row; Col = 0 }, cell (CellValue.Number 1.0) ]
            |> Map.ofList
        Filters = Map.ofList [ 1, { Column = 1; VisibleValues = Set.singleton "keep" } ]
        FilterHiddenRows = [| for row in 1 .. 2 .. subtotalProjectedFilterRows - 1 -> row |] }
evaluateFormula subtotalProjectedFilterFormula subtotalProjectedFilterModel None |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let subtotalProjectedFilterBefore = System.GC.GetAllocatedBytesForCurrentThread()
let subtotalProjectedFilterSw = Stopwatch.StartNew()
let subtotalProjectedFilterResult = evaluateFormula subtotalProjectedFilterFormula subtotalProjectedFilterModel None
subtotalProjectedFilterSw.Stop()
let subtotalProjectedFilterBytes = System.GC.GetAllocatedBytesForCurrentThread() - subtotalProjectedFilterBefore
ok "subtotal-projected-filter-perf"
   (subtotalProjectedFilterResult = CellValue.Number 50000.0
    && subtotalProjectedFilterSw.Elapsed.TotalMilliseconds < 120.0)
   (sprintf "projected filtered SUBTOTAL over %d rows took %.1f ms and returned %A"
        subtotalProjectedFilterRows
        subtotalProjectedFilterSw.Elapsed.TotalMilliseconds
        subtotalProjectedFilterResult)
ok "subtotal-projected-filter-allocation"
   (subtotalProjectedFilterBytes < 3000000L)
   (sprintf "projected filtered SUBTOTAL allocated %.1f MB"
        (float subtotalProjectedFilterBytes / 1024.0 / 1024.0))

let subtotalWideFilterRows = 100000
let subtotalWideFilterColumnCount = 64
let subtotalWideFilterFormula = "=SUBTOTAL(9,A1:A100000)"
let subtotalWideFilterModel =
    { createEmpty subtotalWideFilterRows (subtotalWideFilterColumnCount + 1) with
        Cells =
            [ for row in 0 .. subtotalWideFilterRows - 1 do
                yield { Row = row; Col = 0 }, cell (CellValue.Number 1.0)
                yield { Row = row; Col = 1 }, cell (CellValue.String (if row % 3 = 0 then "keep" else "drop")) ]
            |> Map.ofList
        Filters =
            [ for col in 1 .. subtotalWideFilterColumnCount ->
                col, { Column = col; VisibleValues = Set.ofList [ "keep"; "" ] } ]
            |> Map.ofList }
evaluateFormula subtotalWideFilterFormula subtotalWideFilterModel None |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let subtotalWideFilterBefore = System.GC.GetAllocatedBytesForCurrentThread()
let subtotalWideFilterSw = Stopwatch.StartNew()
let subtotalWideFilterResult = evaluateFormula subtotalWideFilterFormula subtotalWideFilterModel None
subtotalWideFilterSw.Stop()
let subtotalWideFilterBytes = System.GC.GetAllocatedBytesForCurrentThread() - subtotalWideFilterBefore
ok "subtotal-wide-filter-perf"
   (subtotalWideFilterResult = CellValue.Number 33334.0
    && subtotalWideFilterSw.Elapsed.TotalMilliseconds < 140.0)
   (sprintf "wide filtered SUBTOTAL over %d rows and %d filters took %.1f ms and returned %A"
        subtotalWideFilterRows
        subtotalWideFilterColumnCount
        subtotalWideFilterSw.Elapsed.TotalMilliseconds
        subtotalWideFilterResult)
ok "subtotal-wide-filter-allocation"
   (subtotalWideFilterBytes < 2500000L)
   (sprintf "wide filtered SUBTOTAL allocated %.1f MB"
        (float subtotalWideFilterBytes / 1024.0 / 1024.0))

let wideDependencyRefCount = 60000
let wideDependencyFormula =
    "=" + String.concat "+" [ for row in 1 .. wideDependencyRefCount -> sprintf "A%d" row ]
let wideDependencyCells =
    [ for row in 0 .. wideDependencyRefCount - 1 do
        yield { Row = row; Col = 0 }, cell (CellValue.Number 1.0)
      yield { Row = 0; Col = 1 }, cell (CellValue.Formula (wideDependencyFormula, CellValue.Empty)) ]
    |> Map.ofList
let wideDependencyModel = { createEmpty (wideDependencyRefCount + 10) 4 with Cells = wideDependencyCells }
let wideDependencySw = Stopwatch.StartNew()
let wideDependencyResult = recalculateAllFormulas None None wideDependencyModel
wideDependencySw.Stop()
let wideDependencyValue =
    match getCellValue { Row = 0; Col = 1 } wideDependencyResult with
    | CellValue.Formula (_, cached) -> cached
    | other -> other
ok "wide-dependency-walk-correct"
   (wideDependencyValue = CellValue.Number (float wideDependencyRefCount))
   (sprintf "wide dependency formula expected %d got %A" wideDependencyRefCount wideDependencyValue)
ok "wide-dependency-walk-perf"
   (wideDependencySw.Elapsed.TotalMilliseconds < 1500.0)
   (sprintf "wide dependency formula should not recurse/overflow, took %.0f ms" wideDependencySw.Elapsed.TotalMilliseconds)
let wideShiftRefCount = 60000
let wideShiftFormula =
    "=" + String.concat "+" [| for row in 1 .. wideShiftRefCount -> sprintf "A%d" row |]
let wideShiftSw = Stopwatch.StartNew()
let wideShiftOk, wideShiftDetail =
    match parseFormula wideShiftFormula with
    | Ok expr ->
        let shifted = shiftExpr 1 1 expr
        let refs = referencesInExpr None shifted
        let firstShifted = Set.contains { Row = 1; Col = 1 } refs
        let lastShifted = Set.contains { Row = wideShiftRefCount; Col = 1 } refs
        refs.Count = wideShiftRefCount && firstShifted && lastShifted,
        sprintf "refs=%d first=%b last=%b" refs.Count firstShifted lastShifted
    | Result.Error error ->
        false, error
wideShiftSw.Stop()
ok "wide-shift-expr-correct"
   wideShiftOk
   wideShiftDetail
ok "wide-shift-expr-perf"
   (wideShiftSw.Elapsed.TotalMilliseconds < 1500.0)
   (sprintf "wide shiftExpr should not recurse/overflow, took %.0f ms; %s" wideShiftSw.Elapsed.TotalMilliseconds wideShiftDetail)

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

let reverseChainLen = 20000
let reverseChainCells =
    [ for r in 0 .. reverseChainLen - 1 ->
        let v =
            if r = reverseChainLen - 1 then CellValue.Formula ("=1", CellValue.Empty)
            else CellValue.Formula (sprintf "=A%d+1" (r + 2), CellValue.Empty)
        { Row = r; Col = 0 }, cell v ]
    |> Map.ofList
let reverseChainModel = { createEmpty (reverseChainLen + 10) 10 with Cells = reverseChainCells }
let reverseChainSw = Stopwatch.StartNew()
let reverseChainResult = recalculateAllFormulas None None reverseChainModel
reverseChainSw.Stop()
let reverseChainTop =
    match getCellValue { Row = 0; Col = 0 } reverseChainResult with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "reverse-chain-correct" (reverseChainTop = CellValue.Number (float reverseChainLen))
   (sprintf "top expected %d got %A" reverseChainLen reverseChainTop)
ok "reverse-chain-perf"
   (reverseChainSw.Elapsed.TotalMilliseconds < 3000.0)
   (sprintf "%d-cell reverse chain took %.0f ms" reverseChainLen reverseChainSw.Elapsed.TotalMilliseconds)

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

// Large sparse numeric ranges should ignore missing cells just like the dense
// range path. This also exercises the sparse aggregate scan used for broad
// ranges in large sheets.
let sparseRangeCells =
    [ { Row = 0; Col = 0 }, cell (CellValue.Number 1.0)
      { Row = 9999; Col = 0 }, cell (CellValue.Number 2.0)
      { Row = 0; Col = 1 }, cell (CellValue.Formula ("=SUM(A1:A10000)", CellValue.Empty))
      { Row = 1; Col = 1 }, cell (CellValue.Formula ("=AVERAGE(A1:A10000)", CellValue.Empty)) ]
    |> Map.ofList
let sparseRangeResult = recalculateAllFormulas None None { createEmpty 10010 4 with Cells = sparseRangeCells }
let sparseSum =
    match getCellValue { Row = 0; Col = 1 } sparseRangeResult with
    | CellValue.Formula (_, v) -> v
    | v -> v
let sparseAverage =
    match getCellValue { Row = 1; Col = 1 } sparseRangeResult with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "sparse-range-sum" (sparseSum = CellValue.Number 3.0)
   (sprintf "SUM over sparse range expected 3, got %A" sparseSum)
ok "sparse-range-average" (sparseAverage = CellValue.Number 1.5)
   (sprintf "AVERAGE over sparse range expected 1.5, got %A" sparseAverage)

// Large ranges should not be expanded into a dependency edge per cell, but
// formula cells inside the range still have to evaluate before the aggregate.
let rangeOrderCells =
    [ { Row = 0; Col = 0 }, cell (CellValue.Formula ("=SUM(D1:D5000)", CellValue.Empty))
      { Row = 4999; Col = 3 }, cell (CellValue.Formula ("=7", CellValue.Empty)) ]
    |> Map.ofList
let rangeOrderModel = { createEmpty 5010 8 with Cells = rangeOrderCells }
let rangeOrderResult = recalculateAllFormulas None None rangeOrderModel
let rangeOrderValue =
    match getCellValue { Row = 0; Col = 0 } rangeOrderResult with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "large-range-order" (rangeOrderValue = CellValue.Number 7.0)
   (sprintf "SUM over large range should see formula inside range, got %A" rangeOrderValue)
let wideRangeOrderCells =
    [ { Row = 0; Col = 0 }, cell (CellValue.Formula ("=SUM(A4:ZZ4)", CellValue.Empty))
      { Row = 3; Col = 701 }, cell (CellValue.Formula ("=11", CellValue.Empty)) ]
    |> Map.ofList
let wideRangeOrderModel = { createEmpty 10 800 with Cells = wideRangeOrderCells }
let wideRangeOrderResult = recalculateAllFormulas None None wideRangeOrderModel
let wideRangeOrderValue =
    match getCellValue { Row = 0; Col = 0 } wideRangeOrderResult with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "wide-large-range-order" (wideRangeOrderValue = CellValue.Number 11.0)
   (sprintf "SUM over wide large range should see formula inside range, got %A" wideRangeOrderValue)

let formulaPresenceCellsA =
    [ for i in 0 .. 20 -> { Row = i; Col = 0 }, cell (CellValue.Number (float i)) ]
    |> Map.ofList
let formulaPresenceCellsB =
    [ for i in 0 .. 20 -> { Row = i; Col = 1 }, cell (CellValue.String (string i)) ]
    |> Map.ofList
ok "formula-presence-empty-a" (not (cellsHaveFormulaCells formulaPresenceCellsA))
   "first formula-presence map should have no formulas"
cellsHaveFormulaCells formulaPresenceCellsB |> ignore
ok "formula-presence-cache-keeps-first" (cellsFormulaCountIfKnown formulaPresenceCellsA = Some 0)
   "formula-presence cache should retain the first cells map after caching another"
let formulaPresenceManySources =
    [| for sourceIndex in 0 .. 23 ->
        [ for index in 0 .. 19999 ->
            { Row = index; Col = sourceIndex }, cell (CellValue.Number (float index)) ]
        |> Map.ofList |]
let formulaPresenceManyRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for cells in formulaPresenceManySources do
            if cellsHaveFormulaCells cells then
                total <- total + 1
    total
formulaPresenceManyRun 20 |> ignore
let formulaPresenceManySw = Stopwatch.StartNew()
let formulaPresenceManyTotal = formulaPresenceManyRun 20
formulaPresenceManySw.Stop()
ok "formula-presence-cache-many-sources"
   (formulaPresenceManySw.Elapsed.TotalMilliseconds < 30.0 && formulaPresenceManyTotal = 0)
   (sprintf "formula-presence cache should retain many cell maps, took %.1f ms" formulaPresenceManySw.Elapsed.TotalMilliseconds)
let formulaPresenceRetainedCellCount = 20000
let formulaPresenceRetainedModelCount = 80
let formulaPresenceRetainedCells sourceIndex =
    [ for index in 0 .. formulaPresenceRetainedCellCount - 1 ->
        { Row = index; Col = sourceIndex % 4 }, cell (CellValue.Number (float index)) ]
    |> Map.ofList
let formulaPresenceRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for sourceIndex in 0 .. formulaPresenceRetainedModelCount - 1 do
        cellsHaveFormulaCells (formulaPresenceRetainedCells sourceIndex) |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "formula-presence-cache-bounded-source-memory"
   (formulaPresenceRetainedBytes < 80000000L)
   (sprintf "formula-presence cache retained %.1f MB" (float formulaPresenceRetainedBytes / 1024.0 / 1024.0))
let formulaPresenceTextRetainedSourceCount = 80
let formulaPresenceTextRetainedText sourceIndex =
    sprintf "formula-presence-retained-text-%03d-%s" sourceIndex (System.String('x', 512 * 1024))
let formulaPresenceTextRetainedCells sourceIndex =
    [ { Row = 0; Col = 0 }, cell (CellValue.String (formulaPresenceTextRetainedText sourceIndex)) ]
    |> Map.ofList
let formulaPresenceTextRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for sourceIndex in 0 .. formulaPresenceTextRetainedSourceCount - 1 do
        cellsHaveFormulaCells (formulaPresenceTextRetainedCells sourceIndex) |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "formula-presence-cache-bounded-text-memory"
   (formulaPresenceTextRetainedBytes < 40000000L)
   (sprintf "formula-presence text cache retained %.1f MB" (float formulaPresenceTextRetainedBytes / 1024.0 / 1024.0))
let formulaPresenceOversizedSourceCellsWeakReference () =
    let payload = System.String('F', 6 * 1024 * 1024)
    let cells =
        [ { Row = 0; Col = 0 }, cell (CellValue.String payload) ]
        |> Map.ofList
    cellsHaveFormulaCells cells |> ignore
    System.WeakReference(cells)
let formulaPresenceOversizedSourceCellsWeak = formulaPresenceOversizedSourceCellsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "formula-presence-cache-skips-oversized-source"
   (not formulaPresenceOversizedSourceCellsWeak.IsAlive)
   "formula-presence cache should not retain an oversized source cell map"
let formulaPresenceOversizedCellCountCells =
    seq {
        for index in 0 .. 500000 do
            yield { Row = index; Col = 0 }, cell (CellValue.Number (float index))
    }
    |> Map.ofSeq
let formulaPresenceOversizedCellCountSw = Stopwatch.StartNew()
let formulaPresenceOversizedCellCountHasFormula = cellsHaveFormulaCells formulaPresenceOversizedCellCountCells
formulaPresenceOversizedCellCountSw.Stop()
ok "formula-presence-oversized-cell-count-fast"
   (not formulaPresenceOversizedCellCountHasFormula
    && cellsFormulaCountIfKnown formulaPresenceOversizedCellCountCells = None
    && formulaPresenceOversizedCellCountSw.Elapsed.TotalMilliseconds < 300.0)
   (sprintf "oversized numeric formula-presence source took %.1f ms" formulaPresenceOversizedCellCountSw.Elapsed.TotalMilliseconds)
let deepCachedFormulaValueDepth = 200000
let mutable deepCachedFormulaValue = CellValue.String "leaf"
let mutable deepCachedFormulaNumber = CellValue.Number 7.0
for depth = deepCachedFormulaValueDepth - 1 downto 0 do
    deepCachedFormulaValue <- CellValue.Formula ("=A1", deepCachedFormulaValue)
    deepCachedFormulaNumber <- CellValue.Formula ("=B1", deepCachedFormulaNumber)
let deepCachedFormulaCells =
    [ { Row = 0; Col = 0 }, cell deepCachedFormulaValue ]
    |> Map.ofList
let deepCachedFormulaModel = { createEmpty 100 10 with Cells = deepCachedFormulaCells }
let deepCachedFormulaViewport = defaultViewport 500.0 300.0
let deepCachedFormulaSw = Stopwatch.StartNew()
let deepCachedFormulaDisplay = displayValue deepCachedFormulaValue
let deepCachedFormulaText = cellDisplayText None 120.0 defaultFormat deepCachedFormulaValue
let deepCachedFormulaExpr = valueToExpr deepCachedFormulaValue
let deepCachedFormulaPivot =
    RedBat.Spreadsheet.Pivot.groupBy
        [| [| deepCachedFormulaValue |] |]
        [| [| deepCachedFormulaNumber |] |]
        RedBat.Spreadsheet.Pivot.Sum
let deepCachedFormulaHasFormula = cellsHaveFormulaCells deepCachedFormulaCells
let deepCachedFormulaVisible = visibleCells deepCachedFormulaModel deepCachedFormulaViewport
let deepCachedFormulaPainted = buildViewportPaintItems deepCachedFormulaModel deepCachedFormulaViewport
deepCachedFormulaSw.Stop()
ok "deep cached formula value stack"
   (deepCachedFormulaSw.Elapsed.TotalMilliseconds < 1000.0
    && deepCachedFormulaDisplay = "leaf"
    && deepCachedFormulaText = "leaf"
    && deepCachedFormulaExpr = Str "leaf"
    && deepCachedFormulaPivot.Length = 2
    && deepCachedFormulaPivot.[0].[1] = CellValue.Number 7.0
    && deepCachedFormulaHasFormula
    && deepCachedFormulaVisible.Length > 0
    && deepCachedFormulaPainted.Length > 0)
   (sprintf "deep cached formula value paths took %.1f ms" deepCachedFormulaSw.Elapsed.TotalMilliseconds)
let deepCachedFormulaTextCountSw = Stopwatch.StartNew()
let deepCachedFormulaTextCount = cellDataTextCharCountFor 1024 (cell deepCachedFormulaValue)
deepCachedFormulaTextCountSw.Stop()
ok "cell text count caps deep formula"
   (deepCachedFormulaTextCount = 1025
    && deepCachedFormulaTextCountSw.Elapsed.TotalMilliseconds < 25.0)
   (sprintf "deep formula text count %d took %.1f ms" deepCachedFormulaTextCount deepCachedFormulaTextCountSw.Elapsed.TotalMilliseconds)
let deepCriteriaCachedFormulaModel =
    { createEmpty 2 3 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.String "leaf")
              { Row = 0; Col = 1 }, cell deepCachedFormulaValue ]
            |> Map.ofList }
let deepCriteriaCachedFormulaSw = Stopwatch.StartNew()
let deepCriteriaCachedFormulaResult =
    evaluateFormula "=COUNTIF(A1:A1,B1)" deepCriteriaCachedFormulaModel None
deepCriteriaCachedFormulaSw.Stop()
ok "criteria cache deep formula value stack"
   (deepCriteriaCachedFormulaResult = CellValue.Number 1.0
    && deepCriteriaCachedFormulaSw.Elapsed.TotalMilliseconds < 500.0)
   (sprintf "deep formula criteria result %A took %.1f ms" deepCriteriaCachedFormulaResult deepCriteriaCachedFormulaSw.Elapsed.TotalMilliseconds)
let distinctCriteriaPairs = 512
let distinctCriteriaFormula =
    let args =
        [ for i in 0 .. distinctCriteriaPairs - 1 do
            yield "A1:A1"
            yield sprintf "\"miss-%04d\"" i ]
        |> String.concat ","
    "=COUNTIFS(" + args + ")"
let distinctCriteriaModel =
    { createEmpty 2 2 with
        Cells = [ { Row = 0; Col = 0 }, cell (CellValue.String "miss-0000") ] |> Map.ofList }
let distinctCriteriaSw = Stopwatch.StartNew()
let distinctCriteriaResult = evaluateFormula distinctCriteriaFormula distinctCriteriaModel None
distinctCriteriaSw.Stop()
ok "criteria cache many distinct predicates"
   (distinctCriteriaResult = CellValue.Number 0.0
    && distinctCriteriaSw.Elapsed.TotalMilliseconds < 500.0)
   (sprintf "distinct criteria result %A took %.1f ms" distinctCriteriaResult distinctCriteriaSw.Elapsed.TotalMilliseconds)
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

// Layout prefix caches must preserve row/column coordinates and invalidate
// when the immutable override maps change.
let layoutModel =
    { createEmpty 100 100 with
        RowHeights = Map.ofList [ 0, 10.0; 2, 30.0 ]
        ColumnWidths = Map.ofList [ 0, 50.0; 2, 75.0 ] }
ok "layout-row-prefix" (abs (rowTop layoutModel 4 - 90.0) < 0.001)
   (sprintf "rowTop expected 90 got %.3f" (rowTop layoutModel 4))
ok "layout-col-prefix" (abs (columnLeft layoutModel 4 - 325.0) < 0.001)
   (sprintf "columnLeft expected 325 got %.3f" (columnLeft layoutModel 4))
let layoutModelChanged =
    { layoutModel with
        RowHeights = Map.add 1 50.0 layoutModel.RowHeights
        ColumnWidths = Map.add 1 120.0 layoutModel.ColumnWidths }
ok "layout-row-prefix-invalidates" (abs (rowTop layoutModelChanged 4 - 115.0) < 0.001)
   (sprintf "rowTop changed expected 115 got %.3f" (rowTop layoutModelChanged 4))
ok "layout-col-prefix-invalidates" (abs (columnLeft layoutModelChanged 4 - 345.0) < 0.001)
   (sprintf "columnLeft changed expected 345 got %.3f" (columnLeft layoutModelChanged 4))
let layoutPrefixPerfModelA =
    { createEmpty 25000 25000 with
        RowHeights = [ for i in 0 .. 19999 -> i, 20.0 + float (i % 5) ] |> Map.ofList
        ColumnWidths = [ for i in 0 .. 19999 -> i, 90.0 + float (i % 7) ] |> Map.ofList }
let layoutPrefixPerfModelB =
    { createEmpty 25000 25000 with
        RowHeights = [ for i in 0 .. 19999 -> i, 22.0 + float (i % 5) ] |> Map.ofList
        ColumnWidths = [ for i in 0 .. 19999 -> i, 92.0 + float (i % 7) ] |> Map.ofList }
rowTop layoutPrefixPerfModelA 20000 |> ignore
columnLeft layoutPrefixPerfModelA 20000 |> ignore
let layoutPrefixSw = Stopwatch.StartNew()
let mutable layoutPrefixChecksum = 0
for i in 1 .. 300 do
    let model = if i % 2 = 0 then layoutPrefixPerfModelA else layoutPrefixPerfModelB
    layoutPrefixChecksum <- layoutPrefixChecksum ^^^ int (rowTop model 20000)
    layoutPrefixChecksum <- layoutPrefixChecksum ^^^ int (columnLeft model 20000)
layoutPrefixSw.Stop()
ok "layout-prefix-cache-keeps-first-source"
   (layoutPrefixSw.Elapsed.TotalMilliseconds < 300.0 && layoutPrefixChecksum >= 0)
   (sprintf "row/column prefix caches should retain alternating override maps, took %.1f ms" layoutPrefixSw.Elapsed.TotalMilliseconds)
let prefixManyModelCount = 24
let prefixManyOverrideCount = 5000
let prefixManyModels =
    [| for modelIndex in 0 .. prefixManyModelCount - 1 ->
        { createEmpty 8000 8000 with
            RowHeights = [ for i in 0 .. prefixManyOverrideCount - 1 -> i, 20.0 + float ((i + modelIndex) % 5) ] |> Map.ofList
            ColumnWidths = [ for i in 0 .. prefixManyOverrideCount - 1 -> i, 90.0 + float ((i + modelIndex) % 7) ] |> Map.ofList } |]
let prefixManyRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in prefixManyModels do
            total <- total ^^^ int (rowTop model prefixManyOverrideCount)
            total <- total ^^^ int (columnLeft model prefixManyOverrideCount)
    total
prefixManyRun 3 |> ignore
let prefixManySw = Stopwatch.StartNew()
let prefixManyTotal = prefixManyRun 120
prefixManySw.Stop()
ok "layout-prefix-cache-many-models"
   (prefixManySw.Elapsed.TotalMilliseconds < 15.0 && prefixManyTotal >= 0)
   (sprintf "row/column prefix caches should retain many override maps, took %.1f ms" prefixManySw.Elapsed.TotalMilliseconds)
let hiddenPrefixManyModels =
    [| for modelIndex in 0 .. prefixManyModelCount - 1 ->
        { createEmpty 8000 8000 with
            RowHeights = [ for i in 0 .. prefixManyOverrideCount - 1 -> i, 20.0 + float ((i + modelIndex) % 5) ] |> Map.ofList
            Filters = Map.ofList [ 0, { Column = 0; VisibleValues = Set.singleton "keep" } ]
            FilterHiddenRows = [| for row in 1 .. 2 .. prefixManyOverrideCount - 1 -> row |] } |]
let hiddenPrefixManyRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in hiddenPrefixManyModels do
            total <- total ^^^ int (rowTop model prefixManyOverrideCount)
    total
hiddenPrefixManyRun 3 |> ignore
let hiddenPrefixManySw = Stopwatch.StartNew()
let hiddenPrefixManyTotal = hiddenPrefixManyRun 120
hiddenPrefixManySw.Stop()
ok "layout-hidden-prefix-cache-many-models"
   (hiddenPrefixManySw.Elapsed.TotalMilliseconds < 15.0 && hiddenPrefixManyTotal >= 0)
   (sprintf "hidden row prefix cache should retain many filtered row-height maps, took %.1f ms" hiddenPrefixManySw.Elapsed.TotalMilliseconds)
let layoutPrefixOversizedSourceWeakReference () =
    let widths =
        [ for col in 0 .. 1000000 ->
            col, 90.0 + float (col % 7) ]
        |> Map.ofList
    let model =
        { createEmpty 1 1000002 with
            ColumnWidths = widths }
    columnLeft model 1000001 |> ignore
    System.WeakReference(widths)
let layoutPrefixOversizedSourceWeak = layoutPrefixOversizedSourceWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "layout-prefix-cache-skips-oversized-source"
   (not layoutPrefixOversizedSourceWeak.IsAlive)
   "layout prefix cache should not retain an oversized override map"
let layoutHiddenPrefixOversizedOverridesWeakReference () =
    let heights =
        [ for row in 0 .. 1000000 ->
            row, 20.0 + float (row % 5) ]
        |> Map.ofList
    let hiddenRows = [| 1 |]
    let model =
        { createEmpty 1000002 1 with
            RowHeights = heights
            Filters = Map.ofList [ 0, { Column = 0; VisibleValues = Set.singleton "keep" } ]
            FilterHiddenRows = hiddenRows }
    let sw = Stopwatch.StartNew()
    rowTop model 2 |> ignore
    sw.Stop()
    System.WeakReference(heights), sw.Elapsed.TotalMilliseconds
let layoutHiddenPrefixOversizedOverridesWeak, layoutHiddenPrefixOversizedOverridesElapsed =
    layoutHiddenPrefixOversizedOverridesWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "layout-hidden-prefix-skips-oversized-overrides"
   (not layoutHiddenPrefixOversizedOverridesWeak.IsAlive
    && layoutHiddenPrefixOversizedOverridesElapsed < 20.0)
   (sprintf "hidden prefix cache should not retain oversized row-height maps, took %.1f ms" layoutHiddenPrefixOversizedOverridesElapsed)
let layoutHiddenPrefixOversizedRowsWeakReference () =
    let hiddenRows = [| for row in 0 .. 1000000 -> row |]
    let model =
        { createEmpty 1000002 1 with
            Filters = Map.ofList [ 0, { Column = 0; VisibleValues = Set.singleton "keep" } ]
            FilterHiddenRows = hiddenRows }
    rowTop model 1000001 |> ignore
    System.WeakReference(hiddenRows)
let layoutHiddenPrefixOversizedRowsWeak = layoutHiddenPrefixOversizedRowsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "layout-hidden-prefix-cache-skips-oversized-source"
   (not layoutHiddenPrefixOversizedRowsWeak.IsAlive)
   "hidden prefix cache should not retain an oversized hidden-row array"
let layoutHiddenPrefixDefaultRowsWeakReference () =
    let hiddenRows = [| for row in 1 .. 2 .. 9999 -> row |]
    let model =
        { createEmpty 10000 1 with
            Filters = Map.ofList [ 0, { Column = 0; VisibleValues = Set.singleton "keep" } ]
            FilterHiddenRows = hiddenRows }
    rowTop model 9999 |> ignore
    System.WeakReference(hiddenRows)
let layoutHiddenPrefixDefaultRowsWeak = layoutHiddenPrefixDefaultRowsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "layout-hidden-prefix-skips-default-source"
   (not layoutHiddenPrefixDefaultRowsWeak.IsAlive)
   "hidden prefix cache should not retain default-height hidden rows"

let filteredLayoutPerfRows = 30000
let filteredLayoutPerfModel =
    { createEmpty filteredLayoutPerfRows 100 with
        RowHeights =
            [ for row in 0 .. filteredLayoutPerfRows - 1 -> row, 18.0 + float (row % 11) ]
            |> Map.ofList
        Filters = Map.ofList [ 0, { Column = 0; VisibleValues = Set.singleton "keep" } ]
        FilterHiddenRows = [| for row in 1 .. 2 .. filteredLayoutPerfRows - 1 -> row |] }
let filteredLayoutDeepRow = 26000
let filteredLayoutDeepScrollY = rowTop filteredLayoutPerfModel filteredLayoutDeepRow
let filteredLayoutViewport =
    { defaultViewport 900.0 520.0 with
        ScrollX = 120.0
        ScrollY = 30.0 + filteredLayoutDeepScrollY }
let filteredLayoutRange = visibleRange filteredLayoutPerfModel filteredLayoutViewport
ok "layout-filtered-visible-start"
   (filteredLayoutRange.StartRow = filteredLayoutDeepRow && getRowHeight filteredLayoutRange.StartRow filteredLayoutPerfModel > 0.0)
   (sprintf "filtered deep viewport should start at visible row %d, got %A" filteredLayoutDeepRow filteredLayoutRange)
let filteredLayoutSw = Stopwatch.StartNew()
let mutable filteredLayoutChecksum = 0
for i in 0 .. 199 do
    let viewport =
        { filteredLayoutViewport with
            ScrollY = 30.0 + filteredLayoutDeepScrollY + float (i % 16) }
    let range = visibleRange filteredLayoutPerfModel viewport
    filteredLayoutChecksum <- filteredLayoutChecksum ^^^ range.StartRow ^^^ range.EndRowExclusive
filteredLayoutSw.Stop()
ok "layout-filtered-visible-perf"
   (filteredLayoutSw.Elapsed.TotalMilliseconds < 300.0 && filteredLayoutChecksum >= 0)
   (sprintf "deep filtered visibleRange should avoid scanning hidden/custom rows, took %.1f ms" filteredLayoutSw.Elapsed.TotalMilliseconds)

let layoutViewport = { defaultViewport 500.0 300.0 with ScrollX = 60.0; ScrollY = 40.0 }
let visibleCellsA = visibleCells layoutModel layoutViewport
let visibleCellsB = visibleCells layoutModel layoutViewport
ok "layout-visible-cache-reuse" (obj.ReferenceEquals(visibleCellsA, visibleCellsB))
   "visibleCells should reuse cached list for same model and viewport geometry"
let visibleCellsChangedViewport = visibleCells layoutModel { layoutViewport with ScrollX = 70.0 }
ok "layout-visible-cache-viewport-invalidates" (not (obj.ReferenceEquals(visibleCellsA, visibleCellsChangedViewport)))
   "visibleCells cache should invalidate when viewport geometry changes"
ok "layout-visible-cache-keeps-first-viewport" (obj.ReferenceEquals(visibleCellsA, visibleCells layoutModel layoutViewport))
   "visibleCells should retain the first viewport after caching another viewport"
let visibleCellsChangedModel = visibleCells layoutModelChanged layoutViewport
ok "layout-visible-cache-model-invalidates" (not (obj.ReferenceEquals(visibleCellsA, visibleCellsChangedModel)))
   "visibleCells cache should invalidate when model reference changes"
ok "layout-visible-cache-keeps-first-model" (obj.ReferenceEquals(visibleCellsA, visibleCells layoutModel layoutViewport))
   "visibleCells should retain the first model after caching another model"
let headersA = headers layoutModel layoutViewport
let headersB = headers layoutModel layoutViewport
ok "layout-headers-cache-reuse" (obj.ReferenceEquals(headersA, headersB))
   "headers should reuse cached list for same model and viewport geometry"
let headersChangedViewport = headers layoutModel { layoutViewport with ScrollY = 50.0 }
ok "layout-headers-cache-viewport-invalidates" (not (obj.ReferenceEquals(headersA, headersChangedViewport)))
   "headers cache should invalidate when viewport geometry changes"
ok "layout-headers-cache-keeps-first-viewport" (obj.ReferenceEquals(headersA, headers layoutModel layoutViewport))
   "headers should retain the first viewport after caching another viewport"
let headersChangedModel = headers layoutModelChanged layoutViewport
ok "layout-headers-cache-model-invalidates" (not (obj.ReferenceEquals(headersA, headersChangedModel)))
   "headers cache should invalidate when model reference changes"
ok "layout-headers-cache-keeps-first-model" (obj.ReferenceEquals(headersA, headers layoutModel layoutViewport))
   "headers should retain the first model after caching another model"
let oversizedLayoutCacheModel = createEmpty 210000 1
let oversizedLayoutCacheViewport =
    defaultViewport 100.0 (30.0 + 25.0 * 200001.0)
let oversizedVisibleCacheSkipped () =
    let sw = Stopwatch.StartNew()
    let cellsA = visibleCells oversizedLayoutCacheModel oversizedLayoutCacheViewport
    let cellsB = visibleCells oversizedLayoutCacheModel oversizedLayoutCacheViewport
    sw.Stop()
    cellsA.Length, obj.ReferenceEquals(cellsA, cellsB), sw.Elapsed.TotalMilliseconds
let oversizedVisibleCount, oversizedVisibleReused, oversizedVisibleElapsed = oversizedVisibleCacheSkipped ()
ok "layout-visible-cache-skips-oversized-entry"
   (oversizedVisibleCount > 200000 && not oversizedVisibleReused)
   (sprintf "oversized visible-cell entry should not be cached, count=%d reused=%b" oversizedVisibleCount oversizedVisibleReused)
ok "layout-visible-cache-oversized-perf"
   (oversizedVisibleElapsed < 180.0)
   (sprintf "oversized visible-cell cache skip should avoid extra list traversals, took %.1f ms" oversizedVisibleElapsed)
let oversizedHeadersCacheSkipped () =
    let headersA = headers oversizedLayoutCacheModel oversizedLayoutCacheViewport
    let headersB = headers oversizedLayoutCacheModel oversizedLayoutCacheViewport
    headersA.Length, obj.ReferenceEquals(headersA, headersB)
let oversizedHeaderCount, oversizedHeadersReused = oversizedHeadersCacheSkipped ()
ok "layout-headers-cache-skips-oversized-entry"
   (oversizedHeaderCount > 200000 && not oversizedHeadersReused)
   (sprintf "oversized header entry should not be cached, count=%d reused=%b" oversizedHeaderCount oversizedHeadersReused)
let manyViewportLayoutModel =
    { createEmpty 5000 200 with
        Cells =
            [ for index in 0 .. 23 ->
                { Row = index * 20; Col = index % 10 }, cell (CellValue.String (string index)) ]
            |> Map.ofList }
let manyViewportLayoutViewports =
    [| for index in 0 .. 23 ->
        { defaultViewport 2200.0 1600.0 with
            ScrollX = float (index % 6) * 120.0
            ScrollY = float (index / 6) * 420.0 } |]
let manyViewportLayoutRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for viewport in manyViewportLayoutViewports do
            total <- total + (visibleCells manyViewportLayoutModel viewport).Length
            total <- total + (headers manyViewportLayoutModel viewport).Length
    total
manyViewportLayoutRun 20 |> ignore
let manyViewportLayoutSw = Stopwatch.StartNew()
let manyViewportLayoutTotal = manyViewportLayoutRun 20
manyViewportLayoutSw.Stop()
ok "layout-viewport-cache-many-viewports"
   (manyViewportLayoutSw.Elapsed.TotalMilliseconds < 50.0 && manyViewportLayoutTotal > 0)
   (sprintf "visibleCells/headers should retain repeated scroll positions, took %.1f ms" manyViewportLayoutSw.Elapsed.TotalMilliseconds)
let manyModelLayoutCount = 24
let manyModelLayoutModels =
    [| for modelIndex in 0 .. manyModelLayoutCount - 1 ->
        { createEmpty 5000 200 with
            Cells =
                [ for index in 0 .. 23 ->
                    { Row = index * 20; Col = (index + modelIndex) % 10 }, cell (CellValue.String (string index)) ]
                |> Map.ofList } |]
let manyModelLayoutViewport =
    { defaultViewport 2200.0 1600.0 with
        ScrollX = 120.0
        ScrollY = 420.0 }
let manyModelLayoutRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in manyModelLayoutModels do
            total <- total + (visibleCells model manyModelLayoutViewport).Length
            total <- total + (headers model manyModelLayoutViewport).Length
    total
manyModelLayoutRun 3 |> ignore
let manyModelLayoutSw = Stopwatch.StartNew()
let manyModelLayoutTotal = manyModelLayoutRun 30
manyModelLayoutSw.Stop()
ok "layout-viewport-cache-many-models"
   (manyModelLayoutSw.Elapsed.TotalMilliseconds < 35.0 && manyModelLayoutTotal > 0)
   (sprintf "visibleCells/headers should retain repeated model sources, took %.1f ms" manyModelLayoutSw.Elapsed.TotalMilliseconds)
let manyModelLayoutHitSw = Stopwatch.StartNew()
let manyModelLayoutHitTotal = manyModelLayoutRun 200
manyModelLayoutHitSw.Stop()
ok "layout-viewport-cache-hit-fast"
   (manyModelLayoutHitSw.Elapsed.TotalMilliseconds < 70.0 && manyModelLayoutHitTotal > 0)
   (sprintf "cached visibleCells/headers hits should not reorder source caches, took %.1f ms" manyModelLayoutHitSw.Elapsed.TotalMilliseconds)
let layoutSelectionSourceMetricsModel =
    { createEmpty 1000 100 with
        Cells =
            [ for row in 0 .. 999 do
                for col in 0 .. 99 do
                    { Row = row; Col = col }, cell (CellValue.String "cell text payload") ]
            |> Map.ofList }
let layoutSelectionSourceMetricsViewport = defaultViewport 1200.0 800.0
let layoutSelectionSourceMetricsRun passes =
    let mutable total = 0
    for index in 1 .. passes do
        let model =
            { layoutSelectionSourceMetricsModel with
                Selection = SingleCell (index % 50, index % 20) }
        total <- total + (visibleCells model layoutSelectionSourceMetricsViewport).Length
        total <- total + (headers model layoutSelectionSourceMetricsViewport).Length
    total
layoutSelectionSourceMetricsRun 3 |> ignore
let layoutSelectionSourceMetricsSw = Stopwatch.StartNew()
let layoutSelectionSourceMetricsTotal = layoutSelectionSourceMetricsRun 30
layoutSelectionSourceMetricsSw.Stop()
ok "layout-selection-source-metrics-cache"
   (layoutSelectionSourceMetricsSw.Elapsed.TotalMilliseconds < 150.0 && layoutSelectionSourceMetricsTotal > 0)
   (sprintf "selection-only layout should not rescan source text, took %.1f ms" layoutSelectionSourceMetricsSw.Elapsed.TotalMilliseconds)
let layoutCacheRetainedTableCount = 10000
let layoutCacheRetainedModelCount = 80
let layoutCacheRetainedStyle =
    { Name = "LayoutImported"
      ShowFirstColumn = false
      ShowLastColumn = false
      ShowRowStripes = true
      ShowColumnStripes = false
      HeaderFill = None
      StripeFill = None }
let layoutCacheRetainedModel modelIndex =
    { createEmpty 100 20 with
        Tables =
            [ for index in 0 .. layoutCacheRetainedTableCount - 1 ->
                { Name = sprintf "Layout_%02d_%04d" modelIndex index
                  StartRow = index + 1
                  StartCol = 1
                  EndRow = index + 2
                  EndCol = 4
                  HeaderRowCount = 1
                  Style = Some layoutCacheRetainedStyle
                  AutoFilter = None
                  HasTotals = false } ] }
let layoutCacheRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. layoutCacheRetainedModelCount - 1 do
        let model = layoutCacheRetainedModel modelIndex
        visibleCells model layoutViewport |> ignore
        headers model layoutViewport |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "layout-viewport-cache-bounded-source-memory"
   (layoutCacheRetainedBytes < 80000000L)
   (sprintf "layout viewport caches retained %.1f MB" (float layoutCacheRetainedBytes / 1024.0 / 1024.0))
let layoutTextCacheRetainedModelCount = 80
let layoutTextCacheRetainedValue modelIndex =
    sprintf "layout-text-retained-%02d-%s" modelIndex (System.String('x', 512 * 1024))
let layoutTextCacheRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. layoutTextCacheRetainedModelCount - 1 do
        let model =
            { createEmpty 100 100 with
                Cells =
                    Map.ofList [
                        { Row = 0; Col = 0 }, cell (CellValue.String (layoutTextCacheRetainedValue modelIndex))
                    ] }
        visibleCells model layoutViewport |> ignore
        headers model layoutViewport |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "layout-viewport-cache-bounded-text-memory"
   (layoutTextCacheRetainedBytes < 40000000L)
   (sprintf "layout viewport text caches retained %.1f MB" (float layoutTextCacheRetainedBytes / 1024.0 / 1024.0))
let layoutOversizedSourceModelWeakReference () =
    let payload = System.String('L', 6 * 1024 * 1024)
    let model =
        { createEmpty 100 100 with
            Cells = Map.ofList [ { Row = 0; Col = 0 }, cell (CellValue.String payload) ] }
    visibleCells model layoutViewport |> ignore
    headers model layoutViewport |> ignore
    System.WeakReference(model)
let layoutOversizedSourceModelWeak = layoutOversizedSourceModelWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "layout-viewport-cache-skips-oversized-source"
   (not layoutOversizedSourceModelWeak.IsAlive)
   "oversized layout source model should not be retained by viewport geometry caches"
let oversizedItemSourceCellCount = 200001
let oversizedItemSourceModel =
    { createEmpty 1 oversizedItemSourceCellCount with
        Cells =
            [ for col in 0 .. oversizedItemSourceCellCount - 1 ->
                { Row = 0; Col = col }, cell (CellValue.Number 1.0) ]
            |> Map.ofList }
let oversizedMergedCellSourceCount = 1000000
let oversizedMergedCellSource =
    [ for index in 0 .. oversizedMergedCellSourceCount - 1 ->
        { StartRow = 1000000 + index
          StartCol = 0
          EndRow = 1000000 + index
          EndCol = 0 } ]
let oversizedMergedCellSourceModel =
    { createEmpty 100 100 with MergedCells = oversizedMergedCellSource }
let oversizedMergedCellSourceViewport = defaultViewport 0.0 0.0
let oversizedTextSourceCellCount = 200000
let oversizedTextSourceManyCellsModel =
    let payload = System.String('S', 6 * 1024 * 1024)
    let headerFooter = { defaultHeaderFooter with OddHeader = Some payload }
    { createEmpty 1 oversizedTextSourceCellCount with
        PageSetup = { defaultPageSetup with HeaderFooter = headerFooter }
        Cells =
            [ for col in 0 .. oversizedTextSourceCellCount - 1 ->
                { Row = 0; Col = col }, cell (CellValue.Number 1.0) ]
            |> Map.ofList }
let oversizedFilterTextSourceManyCellsModel =
    let payload = System.String('F', 6 * 1024 * 1024)
    { createEmpty 1 oversizedTextSourceCellCount with
        Filters = Map.ofList [ 0, { Column = 0; VisibleValues = Set.singleton payload } ]
        Cells =
            [ for col in 0 .. oversizedTextSourceCellCount - 1 ->
                { Row = 0; Col = col }, cell (CellValue.Number 1.0) ]
            |> Map.ofList }
let layoutOversizedItemSourceSw = Stopwatch.StartNew()
visibleCells oversizedItemSourceModel layoutViewport |> ignore
headers oversizedItemSourceModel layoutViewport |> ignore
layoutOversizedItemSourceSw.Stop()
ok "layout-oversized-item-source-fast"
   (layoutOversizedItemSourceSw.Elapsed.TotalMilliseconds < 750.0)
   (sprintf "layout oversized item source took %.1f ms" layoutOversizedItemSourceSw.Elapsed.TotalMilliseconds)
let layoutOversizedMergedSourceMs, layoutOversizedMergedSourceTotal =
    let sw = Stopwatch.StartNew()
    let mutable total = 0
    for _ in 1 .. 20 do
        total <- total + (headers oversizedMergedCellSourceModel oversizedMergedCellSourceViewport).Length
    sw.Stop()
    sw.Elapsed.TotalMilliseconds, total
ok "layout-oversized-merge-source-fast"
   (layoutOversizedMergedSourceTotal = 20
    && layoutOversizedMergedSourceMs < 120.0)
   (sprintf "layout oversized merge source took %.1f ms" layoutOversizedMergedSourceMs)
let layoutOversizedTextSourceSw = Stopwatch.StartNew()
visibleCells oversizedTextSourceManyCellsModel layoutViewport |> ignore
headers oversizedTextSourceManyCellsModel layoutViewport |> ignore
layoutOversizedTextSourceSw.Stop()
ok "layout-oversized-text-source-fast"
   (layoutOversizedTextSourceSw.Elapsed.TotalMilliseconds < 600.0)
   (sprintf "layout oversized text source took %.1f ms" layoutOversizedTextSourceSw.Elapsed.TotalMilliseconds)
let layoutOversizedFilterTextSourceSw = Stopwatch.StartNew()
for _ in 1 .. 12 do
    visibleCells oversizedFilterTextSourceManyCellsModel layoutViewport |> ignore
    headers oversizedFilterTextSourceManyCellsModel layoutViewport |> ignore
layoutOversizedFilterTextSourceSw.Stop()
ok "layout-oversized-filter-text-source-fast"
   (layoutOversizedFilterTextSourceSw.Elapsed.TotalMilliseconds < 500.0)
   (sprintf "layout oversized filter text source took %.1f ms" layoutOversizedFilterTextSourceSw.Elapsed.TotalMilliseconds)
let layoutOversizedPageSetupModelWeakReference () =
    let payload = System.String('G', 6 * 1024 * 1024)
    let headerFooter = { defaultHeaderFooter with OddHeader = Some payload }
    let model =
        { createEmpty 100 100 with
            PageSetup = { defaultPageSetup with HeaderFooter = headerFooter } }
    visibleCells model layoutViewport |> ignore
    headers model layoutViewport |> ignore
    System.WeakReference(model)
let layoutOversizedPageSetupModelWeak = layoutOversizedPageSetupModelWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "layout-viewport-cache-skips-oversized-page-setup"
   (not layoutOversizedPageSetupModelWeak.IsAlive)
   "layout viewport caches should not retain models whose header/footer metadata exceeds the source text budget"
let layoutOversizedPageSetupItemsMs, layoutOversizedPageSetupItemsTotal =
    let pageSetup =
        { defaultPageSetup with
            ManualRowBreaks = [ for index in 0 .. 1000000 -> index ] }
    let model = { createEmpty 100 100 with PageSetup = pageSetup }
    let sw = Stopwatch.StartNew()
    let mutable total = 0
    for _ in 1 .. 20 do
        total <- total + (visibleCells model layoutViewport).Length
        total <- total + (headers model layoutViewport).Length
    sw.Stop()
    sw.Elapsed.TotalMilliseconds, total
ok "layout-oversized-page-setup-items-fast"
   (layoutOversizedPageSetupItemsTotal > 0 && layoutOversizedPageSetupItemsMs < 350.0)
   (sprintf "layout oversized page setup items took %.1f ms" layoutOversizedPageSetupItemsMs)
let layoutOversizedClipboardModelWeakReference () =
    let payload = System.String('K', 6 * 1024 * 1024)
    let clipboard =
        { Cells = [ 0, 0, cell (CellValue.String payload) ]
          IsCut = false
          SourceStartRow = 0
          SourceStartCol = 0 }
    let model = { createEmpty 100 100 with Clipboard = Some clipboard }
    visibleCells model layoutViewport |> ignore
    headers model layoutViewport |> ignore
    System.WeakReference(model)
let layoutOversizedClipboardModelWeak = layoutOversizedClipboardModelWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "layout-viewport-cache-skips-oversized-clipboard"
   (not layoutOversizedClipboardModelWeak.IsAlive)
   "layout viewport caches should not retain models whose clipboard payload exceeds the source text budget"
let layoutOversizedHistoryModelWeakReference () =
    let payload = System.String('J', 6 * 1024 * 1024)
    let historyCommand =
        { Undo = (fun model -> if payload.Length < 0 then model else model)
          Redo = (fun model -> if payload.Length < 0 then model else model)
          Description = "layout-history-payload"
          RetainedItemCount = 1
          RetainedTextCharCount = payload.Length }
    let model =
        { createEmpty 100 100 with
            History = { emptyHistory with Past = [ historyCommand ] } }
    visibleCells model layoutViewport |> ignore
    headers model layoutViewport |> ignore
    System.WeakReference(model)
let layoutOversizedHistoryModelWeak = layoutOversizedHistoryModelWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "layout-viewport-cache-skips-oversized-history"
   (not layoutOversizedHistoryModelWeak.IsAlive)
   "layout viewport caches should not retain models whose history payload exceeds the source text budget"

let sparsePaintModel =
    { createEmpty 100 100 with
        Cells = Map.ofList [ { Row = 0; Col = 0 }, cell (CellValue.String "visible") ] }
let viewportPaintA = buildViewportPaintItems sparsePaintModel layoutViewport
let viewportPaintB = buildViewportPaintItems sparsePaintModel layoutViewport
ok "paint-viewport-cache-reuse" (obj.ReferenceEquals(viewportPaintA, viewportPaintB))
   "buildViewportPaintItems should reuse cached list for same model and viewport"
let viewportPaintChangedViewport = buildViewportPaintItems sparsePaintModel { layoutViewport with ScrollX = layoutViewport.ScrollX + 1.0 }
ok "paint-viewport-cache-viewport-invalidates" (not (obj.ReferenceEquals(viewportPaintA, viewportPaintChangedViewport)))
   "buildViewportPaintItems cache should invalidate when viewport geometry changes"
ok "paint-viewport-cache-keeps-first-viewport" (obj.ReferenceEquals(viewportPaintA, buildViewportPaintItems sparsePaintModel layoutViewport))
   "buildViewportPaintItems should retain the first viewport after caching another viewport"
let sparsePaintModelChanged = { sparsePaintModel with Selection = Range (2, 2, 3, 3) }
let viewportPaintChangedModelForRetention = buildViewportPaintItems sparsePaintModelChanged layoutViewport
ok "paint-viewport-cache-retention-invalidates" (not (obj.ReferenceEquals(viewportPaintA, viewportPaintChangedModelForRetention)))
   "buildViewportPaintItems cache should invalidate when model reference changes"
ok "paint-viewport-cache-keeps-first-model" (obj.ReferenceEquals(viewportPaintA, buildViewportPaintItems sparsePaintModel layoutViewport))
   "buildViewportPaintItems should retain the first model after caching another model"
let viewportPaintIterItems =
    let items = System.Collections.Generic.List<ViewportPaintItem>()
    iterViewportPaintItems sparsePaintModel layoutViewport (fun zone item -> items.Add { Zone = zone; Item = item })
    [ for item in items -> item ]
ok "paint-viewport-iter-build-match" (viewportPaintIterItems = buildViewportPaintItems sparsePaintModel layoutViewport)
   "iterViewportPaintItems should emit the cached buildViewportPaintItems sequence"
let staticPaintIterItems =
    let items = System.Collections.Generic.List<PaintItem>()
    iterPaintItems sparsePaintModel layoutViewport items.Add
    [ for item in items -> item ]
ok "paint-static-iter-build-match" (staticPaintIterItems = buildPaintItems sparsePaintModel layoutViewport)
   "iterPaintItems should emit the same static paint sequence as buildPaintItems"
let viewportPaintChangedModel = buildViewportPaintItems { sparsePaintModel with Selection = SingleCell (1, 1) } layoutViewport
ok "paint-viewport-cache-model-invalidates" (not (obj.ReferenceEquals(viewportPaintA, viewportPaintChangedModel)))
   "buildViewportPaintItems cache should invalidate when model reference changes"
let paintMeasureA (_: string) (_: float) (_: string option) (_: bool) (_: bool) = 10.0
let paintMeasureB (_: string) (_: float) (_: string option) (_: bool) (_: bool) = 11.0
let paintMeasureViewportA = { layoutViewport with Measurer = Some paintMeasureA }
let viewportPaintMeasuredA = buildViewportPaintItems sparsePaintModel paintMeasureViewportA
let viewportPaintMeasuredB = buildViewportPaintItems sparsePaintModel paintMeasureViewportA
ok "paint-viewport-cache-measurer-reuse" (obj.ReferenceEquals(viewportPaintMeasuredA, viewportPaintMeasuredB))
   "buildViewportPaintItems should reuse cache for same measurer instance"
let viewportPaintChangedMeasurer = buildViewportPaintItems sparsePaintModel { layoutViewport with Measurer = Some paintMeasureB }
ok "paint-viewport-cache-measurer-invalidates" (not (obj.ReferenceEquals(viewportPaintMeasuredA, viewportPaintChangedMeasurer)))
   "buildViewportPaintItems cache should invalidate when measurer changes"
let oversizedPaintCacheModel = createEmpty 70000 1
let oversizedPaintCacheViewport =
    defaultViewport 100.0 (30.0 + 25.0 * 70000.0)
let oversizedViewportPaintCacheSkipped () =
    let sw = Stopwatch.StartNew()
    let itemsA = buildViewportPaintItems oversizedPaintCacheModel oversizedPaintCacheViewport
    let itemsB = buildViewportPaintItems oversizedPaintCacheModel oversizedPaintCacheViewport
    sw.Stop()
    itemsA.Length, obj.ReferenceEquals(itemsA, itemsB), sw.Elapsed.TotalMilliseconds
let oversizedViewportPaintCount, oversizedViewportPaintReused, oversizedViewportPaintElapsed = oversizedViewportPaintCacheSkipped ()
ok "paint-viewport-cache-skips-oversized-entry"
   (oversizedViewportPaintCount > 200000 && not oversizedViewportPaintReused)
   (sprintf "oversized viewport paint entry should not be cached, count=%d reused=%b" oversizedViewportPaintCount oversizedViewportPaintReused)
ok "paint-viewport-cache-oversized-perf"
   (oversizedViewportPaintElapsed < 220.0)
   (sprintf "oversized viewport paint cache skip should avoid extra list traversals, took %.1f ms" oversizedViewportPaintElapsed)
let oversizedStaticPaintCacheSkipped () =
    let sw = Stopwatch.StartNew()
    let itemsA = buildPaintItems oversizedPaintCacheModel oversizedPaintCacheViewport
    let itemsB = buildPaintItems oversizedPaintCacheModel oversizedPaintCacheViewport
    sw.Stop()
    itemsA.Length, obj.ReferenceEquals(itemsA, itemsB), sw.Elapsed.TotalMilliseconds
let oversizedStaticPaintCount, oversizedStaticPaintReused, oversizedStaticPaintElapsed = oversizedStaticPaintCacheSkipped ()
ok "paint-static-cache-skips-oversized-entry"
   (oversizedStaticPaintCount > 200000 && not oversizedStaticPaintReused)
   (sprintf "oversized static paint entry should not be cached, count=%d reused=%b" oversizedStaticPaintCount oversizedStaticPaintReused)
ok "paint-static-cache-oversized-perf"
   (oversizedStaticPaintElapsed < 180.0)
   (sprintf "oversized static paint cache skip should avoid extra list traversals, took %.1f ms" oversizedStaticPaintElapsed)
let manyViewportPaintModel =
    { createEmpty 5000 200 with
        Cells =
            [ for index in 0 .. 23 ->
                { Row = index * 20; Col = index % 10 }, cell (CellValue.String (string index)) ]
            |> Map.ofList }
let manyViewportPaintViewports =
    [| for index in 0 .. 23 ->
        { defaultViewport 2200.0 1600.0 with
            ScrollX = float (index % 6) * 120.0
            ScrollY = float (index / 6) * 420.0 } |]
let manyViewportPaintRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for viewport in manyViewportPaintViewports do
            total <- total + (buildViewportPaintItems manyViewportPaintModel viewport).Length
    total
manyViewportPaintRun 10 |> ignore
let manyViewportPaintSw = Stopwatch.StartNew()
let manyViewportPaintTotal = manyViewportPaintRun 10
manyViewportPaintSw.Stop()
ok "paint-viewport-cache-many-viewports"
   (manyViewportPaintSw.Elapsed.TotalMilliseconds < 50.0 && manyViewportPaintTotal > 0)
   (sprintf "buildViewportPaintItems should retain repeated scroll positions, took %.1f ms" manyViewportPaintSw.Elapsed.TotalMilliseconds)
let manyModelPaintRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in manyModelLayoutModels do
            total <- total + (buildViewportPaintItems model manyModelLayoutViewport).Length
    total
manyModelPaintRun 3 |> ignore
let manyModelPaintSw = Stopwatch.StartNew()
let manyModelPaintTotal = manyModelPaintRun 30
manyModelPaintSw.Stop()
ok "paint-viewport-cache-many-models"
   (manyModelPaintSw.Elapsed.TotalMilliseconds < 35.0 && manyModelPaintTotal > 0)
   (sprintf "buildViewportPaintItems should retain repeated model sources, took %.1f ms" manyModelPaintSw.Elapsed.TotalMilliseconds)
let manyModelPaintHitRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in manyModelLayoutModels do
            total <- total + (buildViewportPaintItems model manyModelLayoutViewport).Length
            total <- total + (buildPaintItems model manyModelLayoutViewport).Length
    total
manyModelPaintHitRun 5 |> ignore
let manyModelPaintHitSw = Stopwatch.StartNew()
let manyModelPaintHitTotal = manyModelPaintHitRun 200
manyModelPaintHitSw.Stop()
ok "paint-cache-hit-fast"
   (manyModelPaintHitSw.Elapsed.TotalMilliseconds < 55.0 && manyModelPaintHitTotal > 0)
   (sprintf "cached viewport/static paint hits should not reorder source caches, took %.1f ms" manyModelPaintHitSw.Elapsed.TotalMilliseconds)
let paintSelectionSourceMetricsModel =
    { createEmpty 1000 100 with
        Cells =
            [ for row in 0 .. 999 do
                for col in 0 .. 99 do
                    { Row = row; Col = col }, cell (CellValue.String "cell text payload") ]
            |> Map.ofList }
let paintSelectionSourceMetricsViewport = defaultViewport 1200.0 800.0
let paintSelectionSourceMetricsRun passes =
    let mutable total = 0
    for index in 1 .. passes do
        let model =
            { paintSelectionSourceMetricsModel with
                Selection = SingleCell (index % 50, index % 20) }
        total <- total + (buildViewportPaintItems model paintSelectionSourceMetricsViewport).Length
    total
paintSelectionSourceMetricsRun 3 |> ignore
let paintSelectionSourceMetricsSw = Stopwatch.StartNew()
let paintSelectionSourceMetricsTotal = paintSelectionSourceMetricsRun 30
paintSelectionSourceMetricsSw.Stop()
ok "paint-selection-source-metrics-cache"
   (paintSelectionSourceMetricsSw.Elapsed.TotalMilliseconds < 250.0 && paintSelectionSourceMetricsTotal > 0)
   (sprintf "selection-only paint should not rescan source text, took %.1f ms" paintSelectionSourceMetricsSw.Elapsed.TotalMilliseconds)
let viewportPaintRetainedTableCount = 10000
let viewportPaintRetainedModelCount = 80
let viewportPaintRetainedStyle =
    { Name = "ViewportImported"
      ShowFirstColumn = false
      ShowLastColumn = false
      ShowRowStripes = true
      ShowColumnStripes = false
      HeaderFill = None
      StripeFill = None }
let viewportPaintRetainedModel modelIndex =
    { createEmpty 100 20 with
        Tables =
            [ for index in 0 .. viewportPaintRetainedTableCount - 1 ->
                { Name = sprintf "Viewport_%02d_%04d" modelIndex index
                  StartRow = index + 1
                  StartCol = 1
                  EndRow = index + 2
                  EndCol = 4
                  HeaderRowCount = 1
                  Style = Some viewportPaintRetainedStyle
                  AutoFilter = None
                  HasTotals = false } ] }
let viewportPaintRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. viewportPaintRetainedModelCount - 1 do
        buildViewportPaintItems (viewportPaintRetainedModel modelIndex) layoutViewport |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "paint-viewport-cache-bounded-source-memory"
   (viewportPaintRetainedBytes < 80000000L)
   (sprintf "viewport paint cache retained %.1f MB" (float viewportPaintRetainedBytes / 1024.0 / 1024.0))
let paintTextCacheRetainedModelCount = 80
let paintTextCacheRetainedValue modelIndex =
    sprintf "paint-text-retained-%02d-%s" modelIndex (System.String('x', 512 * 1024))
let paintTextCacheRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. paintTextCacheRetainedModelCount - 1 do
        let model =
            { createEmpty 1000 100 with
                Cells =
                    Map.ofList [
                        { Row = 900; Col = 0 }, cell (CellValue.String (paintTextCacheRetainedValue modelIndex))
                    ] }
        buildViewportPaintItems model layoutViewport |> ignore
        buildPaintItems model layoutViewport |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "paint-cache-bounded-text-memory"
   (paintTextCacheRetainedBytes < 50000000L)
   (sprintf "paint caches retained %.1f MB of text payloads" (float paintTextCacheRetainedBytes / 1024.0 / 1024.0))
let paintDuplicateEntryTextModel =
    let payload = System.String('P', 4 * 1024 * 1024) + "\nsecond-line"
    { createEmpty 1000 100 with
        Cells = Map.ofList [ { Row = 0; Col = 0 }, cell (CellValue.String payload) ] }
let paintDuplicateEntryViewportA = defaultViewport 500.0 300.0
let paintDuplicateEntryViewportB = { paintDuplicateEntryViewportA with ScrollX = 1.0 }
let staticPaintDuplicateEntryTextWeakReference () =
    let firstItems = buildPaintItems paintDuplicateEntryTextModel paintDuplicateEntryViewportA
    buildPaintItems paintDuplicateEntryTextModel paintDuplicateEntryViewportB |> ignore
    System.WeakReference(firstItems)
let staticPaintDuplicateEntryTextWeak = staticPaintDuplicateEntryTextWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "paint-static-cache-skips-duplicated-entry-text"
   (not staticPaintDuplicateEntryTextWeak.IsAlive)
   "static paint cache should evict older geometry entries when retained item text exceeds budget"
let viewportPaintDuplicateEntryTextWeakReference () =
    let firstItems = buildViewportPaintItems paintDuplicateEntryTextModel paintDuplicateEntryViewportA
    buildViewportPaintItems paintDuplicateEntryTextModel paintDuplicateEntryViewportB |> ignore
    System.WeakReference(firstItems)
let viewportPaintDuplicateEntryTextWeak = viewportPaintDuplicateEntryTextWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "paint-viewport-cache-skips-duplicated-entry-text"
   (not viewportPaintDuplicateEntryTextWeak.IsAlive)
   "viewport paint cache should evict older geometry entries when retained item text exceeds budget"
let paintOversizedSourceModelWeakReference () =
    let payload = System.String('P', 6 * 1024 * 1024)
    let model =
        { createEmpty 100 100 with
            Cells = Map.ofList [ { Row = 0; Col = 0 }, cell (CellValue.String payload) ] }
    buildViewportPaintItems model layoutViewport |> ignore
    buildPaintItems model layoutViewport |> ignore
    System.WeakReference(model)
let paintOversizedSourceModelWeak = paintOversizedSourceModelWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "paint-cache-skips-oversized-source"
   (not paintOversizedSourceModelWeak.IsAlive)
   "oversized paint source model should not be retained by paint caches"
let paintOversizedItemSourceSw = Stopwatch.StartNew()
buildViewportPaintItems oversizedItemSourceModel layoutViewport |> ignore
buildPaintItems oversizedItemSourceModel layoutViewport |> ignore
paintOversizedItemSourceSw.Stop()
ok "paint-oversized-item-source-fast"
   (paintOversizedItemSourceSw.Elapsed.TotalMilliseconds < 1000.0)
   (sprintf "paint oversized item source took %.1f ms" paintOversizedItemSourceSw.Elapsed.TotalMilliseconds)
let paintOversizedMergedSourceMs, paintOversizedMergedSourceTotal =
    let sw = Stopwatch.StartNew()
    let mutable total = 0
    for _ in 1 .. 20 do
        total <- total + (buildViewportPaintItems oversizedMergedCellSourceModel oversizedMergedCellSourceViewport).Length
        total <- total + (buildPaintItems oversizedMergedCellSourceModel oversizedMergedCellSourceViewport).Length
    sw.Stop()
    sw.Elapsed.TotalMilliseconds, total
ok "paint-oversized-merge-source-fast"
   (paintOversizedMergedSourceTotal > 0
    && paintOversizedMergedSourceMs < 200.0)
   (sprintf "paint oversized merge source took %.1f ms" paintOversizedMergedSourceMs)
let paintOversizedTextSourceSw = Stopwatch.StartNew()
buildViewportPaintItems oversizedTextSourceManyCellsModel layoutViewport |> ignore
buildPaintItems oversizedTextSourceManyCellsModel layoutViewport |> ignore
paintOversizedTextSourceSw.Stop()
ok "paint-oversized-text-source-fast"
   (paintOversizedTextSourceSw.Elapsed.TotalMilliseconds < 800.0)
   (sprintf "paint oversized text source took %.1f ms" paintOversizedTextSourceSw.Elapsed.TotalMilliseconds)
let paintOversizedValidationTextSourceModel =
    let payload = System.String('Y', 6 * 1024 * 1024)
    { createEmpty 1 oversizedTextSourceCellCount with
        Validations =
            Map.ofList [
                { Row = 0; Col = 0 },
                { Rule = ValidationRule.ListValues [ payload ]
                  ErrorMessage = None
                  InputMessage = None }
            ]
        Cells =
            [ for col in 0 .. oversizedTextSourceCellCount - 1 ->
                { Row = 0; Col = col }, cell (CellValue.Number 1.0) ]
            |> Map.ofList }
let paintOversizedValidationTextSourceSw = Stopwatch.StartNew()
for _ in 1 .. 8 do
    buildViewportPaintItems paintOversizedValidationTextSourceModel layoutViewport |> ignore
    buildPaintItems paintOversizedValidationTextSourceModel layoutViewport |> ignore
paintOversizedValidationTextSourceSw.Stop()
ok "paint-oversized-validation-text-source-fast"
   (paintOversizedValidationTextSourceSw.Elapsed.TotalMilliseconds < 700.0)
   (sprintf "paint oversized validation text source took %.1f ms" paintOversizedValidationTextSourceSw.Elapsed.TotalMilliseconds)
let paintOversizedMetadataModelWeakReference () =
    let payload = System.String('V', 6 * 1024 * 1024)
    let model =
        { createEmpty 100 100 with
            Cells = Map.ofList [ { Row = 0; Col = 0 }, cell (CellValue.String "tiny") ]
            Validations =
                Map.ofList [
                    { Row = 0; Col = 0 },
                    { Rule = ValidationRule.ListValues [ payload ]
                      ErrorMessage = Some payload
                      InputMessage = Some "input" }
                ] }
    buildViewportPaintItems model layoutViewport |> ignore
    buildPaintItems model layoutViewport |> ignore
    System.WeakReference(model)
let paintOversizedMetadataModelWeak = paintOversizedMetadataModelWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "paint-cache-skips-oversized-metadata"
   (not paintOversizedMetadataModelWeak.IsAlive)
   "paint caches should not retain models whose validation metadata exceeds the source text budget"
let paintOversizedPageSetupModelWeakReference () =
    let payload = System.String('H', 6 * 1024 * 1024)
    let headerFooter = { defaultHeaderFooter with OddHeader = Some payload }
    let model =
        { createEmpty 100 100 with
            PageSetup = { defaultPageSetup with HeaderFooter = headerFooter } }
    buildViewportPaintItems model layoutViewport |> ignore
    buildPaintItems model layoutViewport |> ignore
    System.WeakReference(model)
let paintOversizedPageSetupModelWeak = paintOversizedPageSetupModelWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "paint-cache-skips-oversized-page-setup"
   (not paintOversizedPageSetupModelWeak.IsAlive)
   "paint caches should not retain models whose header/footer metadata exceeds the source text budget"
let paintOversizedPageSetupItemsMs, paintOversizedPageSetupItemsTotal =
    let pageSetup =
        { defaultPageSetup with
            ManualRowBreaks = [ for index in 0 .. 1000000 -> index ] }
    let model = { createEmpty 100 100 with PageSetup = pageSetup }
    let sw = Stopwatch.StartNew()
    let mutable total = 0
    for _ in 1 .. 20 do
        total <- total + (buildViewportPaintItems model layoutViewport).Length
        total <- total + (buildPaintItems model layoutViewport).Length
    sw.Stop()
    sw.Elapsed.TotalMilliseconds, total
ok "paint-oversized-page-setup-items-fast"
   (paintOversizedPageSetupItemsTotal > 0 && paintOversizedPageSetupItemsMs < 450.0)
   (sprintf "paint oversized page setup items took %.1f ms" paintOversizedPageSetupItemsMs)
let paintOversizedClipboardModelWeakReference () =
    let payload = System.String('B', 6 * 1024 * 1024)
    let clipboard =
        { Cells = [ 0, 0, cell (CellValue.String payload) ]
          IsCut = false
          SourceStartRow = 0
          SourceStartCol = 0 }
    let model = { createEmpty 100 100 with Clipboard = Some clipboard }
    buildViewportPaintItems model layoutViewport |> ignore
    buildPaintItems model layoutViewport |> ignore
    System.WeakReference(model)
let paintOversizedClipboardModelWeak = paintOversizedClipboardModelWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "paint-cache-skips-oversized-clipboard"
   (not paintOversizedClipboardModelWeak.IsAlive)
   "paint caches should not retain models whose clipboard payload exceeds the source text budget"
let paintOversizedHistoryModelWeakReference () =
    let payload = System.String('Y', 6 * 1024 * 1024)
    let historyCommand =
        { Undo = (fun model -> if payload.Length < 0 then model else model)
          Redo = (fun model -> if payload.Length < 0 then model else model)
          Description = "paint-history-payload"
          RetainedItemCount = 1
          RetainedTextCharCount = payload.Length }
    let model =
        { createEmpty 100 100 with
            History = { emptyHistory with Past = [ historyCommand ] } }
    buildViewportPaintItems model layoutViewport |> ignore
    buildPaintItems model layoutViewport |> ignore
    System.WeakReference(model)
let paintOversizedHistoryModelWeak = paintOversizedHistoryModelWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "paint-cache-skips-oversized-history"
   (not paintOversizedHistoryModelWeak.IsAlive)
   "paint caches should not retain models whose history payload exceeds the source text budget"
let paintFilterProjectionOversizedSourceModelWeakReference () =
    let payload = System.String('F', 6 * 1024 * 1024)
    let model =
        { createEmpty 2 1 with
            Cells =
                Map.ofList [
                    { Row = 0; Col = 0 }, cell (CellValue.String payload)
                    { Row = 1; Col = 0 }, cell (CellValue.String "keep")
                ] }
        |> applyFilter 0 (Set.singleton "keep")
    buildViewportPaintItems model layoutViewport |> ignore
    buildPaintItems model layoutViewport |> ignore
    System.WeakReference(model)
let paintFilterProjectionOversizedSourceModelWeak = paintFilterProjectionOversizedSourceModelWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "paint-filter-projection-cache-skips-oversized-source"
   (not paintFilterProjectionOversizedSourceModelWeak.IsAlive)
   "paint filter projection cache should not retain an oversized source model"
[<System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)>]
let paintFilterProjectionRetainedSourceModelWeakReference () =
    let model =
        { createEmpty 1000 20 with
            Cells =
                Map.ofList [
                    { Row = 0; Col = 0 }, cell (CellValue.String "header")
                    { Row = 1; Col = 0 }, cell (CellValue.String "keep")
                    { Row = 2; Col = 0 }, cell (CellValue.String "drop")
                ] }
        |> applyFilter 0 (Set.singleton "keep")
    iterPaintItems model layoutViewport ignore
    System.WeakReference(model)
let paintFilterProjectionRetainedSourceModelWeak = paintFilterProjectionRetainedSourceModelWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "paint-filter-projection-cache-releases-source"
   (not paintFilterProjectionRetainedSourceModelWeak.IsAlive)
   "paint filter projection cache should not retain a small filtered source model forever"
let paintFilterProjectionOversizedProjectedElapsed =
    let rowCount = 199998
    let cells =
        [ for row in 0 .. rowCount - 1 ->
            { Row = row; Col = 0 }, cell (CellValue.String "drop") ]
        |> Map.ofList
    let filters = Map.ofList [ 0, { Column = 0; VisibleValues = Set.singleton "keep" } ]
    let warmHiddenRowsModel =
        { createEmpty rowCount 1 with
            Cells = cells
            Filters = filters }
    applyFilterRowHeights warmHiddenRowsModel |> ignore
    let warmPaintModel =
        { createEmpty rowCount 1 with
            Cells = cells
            Filters = filters }
    buildPaintItems warmPaintModel layoutViewport |> ignore
    let sw = Stopwatch.StartNew()
    for _ in 1 .. 5 do
        let model =
            { createEmpty rowCount 1 with
                Cells = cells
                Filters = filters }
        buildPaintItems model layoutViewport |> ignore
    sw.Stop()
    sw.Elapsed.TotalMilliseconds
ok "paint-filter-projection-oversized-projected-fast"
   (paintFilterProjectionOversizedProjectedElapsed < 400.0)
   (sprintf "paint filter projection oversized projected took %.1f ms" paintFilterProjectionOversizedProjectedElapsed)
let sparsePaintTexts =
    buildPaintItems sparsePaintModel (defaultViewport 700.0 500.0)
    |> List.choose (function
        | DrawText text -> Some text.Text
        | _ -> None)
let sparseVisibleTextCount = sparsePaintTexts |> List.filter ((=) "visible") |> List.length
ok "paint-sparse-visible-text" (sparseVisibleTextCount = 1 && not (sparsePaintTexts |> List.exists ((=) "")))
   (sprintf "unexpected sparse paint text items %A" sparsePaintTexts)
let offscreenSparsePaint =
    buildPaintItems
        { createEmpty 1000 100 with
            Cells = Map.ofList [ { Row = 500; Col = 0 }, cell (CellValue.String "far") ] }
        (defaultViewport 700.0 500.0)
let offscreenSparseTexts =
    offscreenSparsePaint
    |> List.choose (function | DrawText text -> Some text.Text | _ -> None)
ok "paint-sparse-offscreen-text" (not (offscreenSparseTexts |> List.exists ((=) "far")))
   (sprintf "offscreen sparse text should not paint, got %A" offscreenSparseTexts)
ok "paint-sparse-selection-fill"
   (offscreenSparsePaint |> List.exists (function | FillRect (_, "#E8F5E9") -> true | _ -> false))
   "sparse default paint path should keep the selected-cell fill"
let styledSparsePaint =
    buildPaintItems
        { createEmpty 100 100 with
            Cells =
                Map.ofList [
                    { Row = 1; Col = 0 },
                    { Value = CellValue.String "styled"
                      Format = { defaultFormat with BgColor = Some "#123456" }
                      Comment = None } ] }
        (defaultViewport 700.0 500.0)
ok "paint-sparse-styled-fallback"
   (styledSparsePaint |> List.exists (function | FillRect (_, "#123456") -> true | _ -> false))
   "sparse cells with non-default formatting should fall back to the full painter"

let explicitBorderNoMergePaint =
    let bottomBorder = Some { Style = "thin"; Color = Some "#7F7F7F" }
    buildPaintItems
        { createEmpty 2 2 with
            Cells =
                Map.ofList [
                    { Row = 0; Col = 0 }, { cell (CellValue.String "a") with Format = { defaultFormat with BorderBottom = bottomBorder } }
                    { Row = 0; Col = 1 }, cell (CellValue.String "b")
                    { Row = 1; Col = 0 }, cell (CellValue.String "c")
                    { Row = 1; Col = 1 }, cell (CellValue.String "d") ] }
        (defaultViewport 300.0 120.0)
let explicitBorderLightGridlines =
    explicitBorderNoMergePaint
    |> List.choose (function
        | GridLine (x1, y1, x2, y2, "#a6a6a6", w) -> Some (x1, y1, x2, y2, w)
        | _ -> None)
ok "paint-border-gridline-dedup"
   ((explicitBorderLightGridlines |> Set.ofList |> Set.count) = explicitBorderLightGridlines.Length)
   (sprintf "explicit-border paint should not emit duplicate light gridlines, got %d items/%d unique"
        explicitBorderLightGridlines.Length
        (explicitBorderLightGridlines |> Set.ofList |> Set.count))

let tablePaintStyle =
    { Name = "TableStyleMedium2"
      ShowFirstColumn = false
      ShowLastColumn = false
      ShowRowStripes = true
      ShowColumnStripes = false
      HeaderFill = Some "#112233"
      StripeFill = None }
let tablePaintModel =
    { createEmpty 20 20 with
        Tables =
            [ { Name = "StyledPaintTable"
                StartRow = 1
                StartCol = 1
                EndRow = 4
                EndCol = 3
                HeaderRowCount = 1
                Style = Some tablePaintStyle
                AutoFilter = None
                HasTotals = false } ] }
let tablePaintHasHeaderFill =
    buildViewportPaintItems tablePaintModel layoutViewport
    |> List.exists (fun painted ->
        match painted.Item with
        | FillRect (_, color) -> color = "#112233"
        | _ -> false)
ok "paint-table-style-cache-miss" tablePaintHasHeaderFill
   "render-time table style detection should invalidate when the table list changes"
let renderStyleCacheImportedStyle =
    { Name = "TableStyleMedium2"
      ShowFirstColumn = false
      ShowLastColumn = false
      ShowRowStripes = true
      ShowColumnStripes = false
      HeaderFill = None
      StripeFill = None }
let renderStyleCacheTableCount = 3000
let renderStyleCacheModels =
    [| for modelIndex in 0 .. 23 ->
        { createEmpty 100 20 with
            Tables =
                [ for index in 0 .. renderStyleCacheTableCount - 1 ->
                    { Name = sprintf "Imported_%02d_%04d" modelIndex index
                      StartRow = index + 1
                      StartCol = 1
                      EndRow = index + 2
                      EndCol = 4
                      HeaderRowCount = 1
                      Style = Some renderStyleCacheImportedStyle
                      AutoFilter = None
                      HasTotals = false } ] } |]
let renderStyleCacheRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in renderStyleCacheModels do
            iterPaintItems model layoutViewport (fun _ -> total <- total + 1)
    total
renderStyleCacheRun 30 |> ignore
let renderStyleCacheSw = Stopwatch.StartNew()
let renderStyleCacheTotal = renderStyleCacheRun 120
renderStyleCacheSw.Stop()
ok "paint-render-table-style-cache-many-models"
   (renderStyleCacheSw.Elapsed.TotalMilliseconds < 80.0 && renderStyleCacheTotal > 0)
   (sprintf "render-time table style detection should retain many table-list snapshots, took %.1f ms" renderStyleCacheSw.Elapsed.TotalMilliseconds)
let renderStyleCacheRetainedTableCount = 10000
let renderStyleCacheRetainedModelCount = 80
let renderStyleCacheRetainedModel modelIndex =
    { createEmpty 100 20 with
        Tables =
            [ for index in 0 .. renderStyleCacheRetainedTableCount - 1 ->
                { Name = sprintf "Retained_%02d_%04d" modelIndex index
                  StartRow = index + 1
                  StartCol = 1
                  EndRow = index + 2
                  EndCol = 4
                  HeaderRowCount = 1
                  Style = Some renderStyleCacheImportedStyle
                  AutoFilter = None
                  HasTotals = false } ] }
let renderStyleCacheRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. renderStyleCacheRetainedModelCount - 1 do
        buildPaintItems (renderStyleCacheRetainedModel modelIndex) layoutViewport |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "paint-render-table-style-cache-bounded-memory"
   (renderStyleCacheRetainedBytes < 80000000L)
   (sprintf "render-time table style caches retained %.1f MB" (float renderStyleCacheRetainedBytes / 1024.0 / 1024.0))
let renderStyleCacheOversizedTablesWeakReference () =
    let tableCount = 200001
    let tables =
        [ for index in 0 .. tableCount - 1 ->
            { Name = sprintf "Oversized_%06d" index
              StartRow = index + 1
              StartCol = 1
              EndRow = index + 2
              EndCol = 4
              HeaderRowCount = 1
              Style = Some renderStyleCacheImportedStyle
              AutoFilter = None
              HasTotals = false } ]
    let model = { createEmpty 100 20 with Tables = tables }
    let sw = Stopwatch.StartNew()
    let items = buildPaintItems model layoutViewport
    sw.Stop()
    System.WeakReference(tables), sw.Elapsed.TotalMilliseconds, items.Length
let renderStyleCacheOversizedTablesWeak, renderStyleCacheOversizedTablesMs, renderStyleCacheOversizedTablesItemCount =
    renderStyleCacheOversizedTablesWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "paint-render-table-style-cache-skips-oversized-source"
   (not renderStyleCacheOversizedTablesWeak.IsAlive)
   "render-time table style cache should not retain an oversized table list"
ok "paint-render-table-style-oversized-source-fast"
   (renderStyleCacheOversizedTablesItemCount > 0 && renderStyleCacheOversizedTablesMs < 250.0)
   (sprintf "render-time table style oversized source took %.1f ms" renderStyleCacheOversizedTablesMs)
let renderStyleCacheOversizedTableTextWeakReference () =
    let payload = System.String('R', 6 * 1024 * 1024)
    let tables =
        [ { Name = payload
            StartRow = 1
            StartCol = 1
            EndRow = 2
            EndCol = 4
            HeaderRowCount = 1
            Style = Some renderStyleCacheImportedStyle
            AutoFilter = None
            HasTotals = false } ]
    let model = { createEmpty 100 20 with Tables = tables }
    buildPaintItems model layoutViewport |> ignore
    System.WeakReference(tables)
let renderStyleCacheOversizedTableTextWeak = renderStyleCacheOversizedTableTextWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "paint-render-table-style-cache-skips-oversized-text"
   (not renderStyleCacheOversizedTableTextWeak.IsAlive)
   "render-time table style cache should not retain an oversized table payload"
let tableStylePerfTable prefix row =
    { Name = sprintf "%s%i" prefix row
      StartRow = row
      StartCol = 1
      EndRow = row + 1
      EndCol = 4
      HeaderRowCount = 1
      Style = None
      AutoFilter = None
      HasTotals = false }
let tableStylePerfModel prefix =
    { createEmpty 100 100 with
        Tables = [ for row in 1 .. 50000 -> tableStylePerfTable prefix row ] }
let tableStylePerfA = tableStylePerfModel "table-style-a"
let tableStylePerfB = tableStylePerfModel "table-style-b"
buildPaintItems tableStylePerfA layoutViewport |> ignore
buildPaintItems tableStylePerfB layoutViewport |> ignore
let tableStylePerfSw = Stopwatch.StartNew()
for _ in 1 .. 60 do
    buildPaintItems tableStylePerfA layoutViewport |> ignore
    buildPaintItems tableStylePerfB layoutViewport |> ignore
tableStylePerfSw.Stop()
ok "paint-table-style-cache-keeps-first-list"
   (tableStylePerfSw.Elapsed.TotalMilliseconds < 25.0)
   (sprintf "render-time table style cache should retain alternating table lists, took %.1f ms" tableStylePerfSw.Elapsed.TotalMilliseconds)

let sortedModel =
    { createEmpty 10 3 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 1.0)
              { Row = 0; Col = 1 }, cell (CellValue.String "a")
              { Row = 1; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 1; Col = 1 }, cell (CellValue.String "b") ]
            |> Map.ofList }
let sortedAgain = sortByColumn 0 true sortedModel
ok "sort-identity-preserves-cells"
   (obj.ReferenceEquals(sortedAgain.Cells, sortedModel.Cells)
    && sortedAgain.SortState = Some { Column = 0; Ascending = true })
   "already-sorted contiguous rows should not rebuild the cell map"
ok "sort-identity-preserves-model"
   (obj.ReferenceEquals(sortByColumn 0 true sortedAgain, sortedAgain))
   "repeating the same already-sorted sort should return the same model"
let sortedTwoColumnModel =
    { sortedModel with
        Cells =
            sortedModel.Cells
            |> Map.add { Row = 0; Col = 2 } (cell (CellValue.Number 1.0))
            |> Map.add { Row = 1; Col = 2 } (cell (CellValue.Number 2.0))
        SortState = Some { Column = 0; Ascending = true } }
let sortedTwoColumnCol0 = sortByColumn 0 true sortedTwoColumnModel
let sortedTwoColumnCol2 = sortByColumn 2 true { sortedTwoColumnModel with SortState = Some { Column = 2; Ascending = true } }
ok "sort-identity-cache-keeps-first-column"
   (obj.ReferenceEquals(sortByColumn 0 true sortedTwoColumnCol0, sortedTwoColumnCol0)
    && obj.ReferenceEquals(sortByColumn 2 true sortedTwoColumnCol2, sortedTwoColumnCol2))
   "identity sort cache should retain more than one sorted column for the same cells"
let sortedTwoColumnSwitchBack = sortByColumn 0 true sortedTwoColumnCol2
ok "sort-identity-cache-switches-state"
   (obj.ReferenceEquals(sortedTwoColumnSwitchBack.Cells, sortedTwoColumnModel.Cells)
    && sortedTwoColumnSwitchBack.SortState = Some { Column = 0; Ascending = true })
   "cached identity sort should avoid rebuilding cells when only the sort state changes"
let sortedWideColumnModel =
    { createEmpty 20 20 with
        Cells =
            [ for row in 0 .. 9 do
                for col in 0 .. 15 ->
                    { Row = row; Col = col }, cell (CellValue.Number (float row)) ]
            |> Map.ofList }
let sortedWideColumn0 = sortByColumn 0 true sortedWideColumnModel
for col in 1 .. 15 do
    sortByColumn col true sortedWideColumnModel |> ignore
ok "sort-identity-cache-keeps-many-columns"
   (obj.ReferenceEquals(sortByColumn 0 true sortedWideColumn0, sortedWideColumn0))
   "identity sort cache should retain the first sorted column after caching many other columns"
let sortedOtherModel =
    { createEmpty 10 3 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 10.0)
              { Row = 1; Col = 0 }, cell (CellValue.Number 20.0) ]
            |> Map.ofList }
let sortedOtherAgain = sortByColumn 0 true sortedOtherModel
ok "sort-identity-cache-keeps-first-source"
   (obj.ReferenceEquals(sortByColumn 0 true sortedAgain, sortedAgain)
    && sortedOtherAgain.SortState = Some { Column = 0; Ascending = true })
   "identity sort cache should retain the first cell map after caching another"
let sortIdentityManyModelCount = 24
let sortIdentityManyRowCount = 8000
let sortIdentityManyModels =
    [| for modelIndex in 0 .. sortIdentityManyModelCount - 1 ->
        { createEmpty sortIdentityManyRowCount 2 with
            Cells =
                [ for row in 0 .. sortIdentityManyRowCount - 1 do
                    yield { Row = row; Col = 0 }, cell (CellValue.Number (float row))
                    yield { Row = row; Col = 1 }, cell (CellValue.String (sprintf "sort-%02d-%05d" modelIndex row)) ]
                |> Map.ofList
            SortState = Some { Column = 0; Ascending = true } } |]
let sortIdentityManyRun passes =
    let mutable same = 0
    for _pass in 1 .. passes do
        for model in sortIdentityManyModels do
            if obj.ReferenceEquals(sortByColumn 0 true model, model) then
                same <- same + 1
    same
sortIdentityManyRun 3 |> ignore
let sortIdentityManySw = Stopwatch.StartNew()
let sortIdentityManySame = sortIdentityManyRun 30
sortIdentityManySw.Stop()
ok "sort-identity-cache-many-models"
   (sortIdentityManySw.Elapsed.TotalMilliseconds < 35.0
    && sortIdentityManySame = sortIdentityManyModelCount * 30)
   (sprintf "identity sort cache should retain many cell maps, took %.1f ms" sortIdentityManySw.Elapsed.TotalMilliseconds)
let sortIdentityRetainedRowCount = 20000
let sortIdentityRetainedModelCount = 80
let sortIdentityRetainedModel modelIndex =
    { createEmpty sortIdentityRetainedRowCount 2 with
        Cells =
            [ for row in 0 .. sortIdentityRetainedRowCount - 1 do
                yield { Row = row; Col = 0 }, cell (CellValue.Number (float row))
                yield { Row = row; Col = 1 }, cell (CellValue.String (sprintf "sort-retained-%02d-%05d" modelIndex row)) ]
            |> Map.ofList
        SortState = Some { Column = 0; Ascending = true } }
let sortIdentityRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. sortIdentityRetainedModelCount - 1 do
        sortByColumn 0 true (sortIdentityRetainedModel modelIndex) |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "sort-identity-cache-bounded-source-memory"
   (sortIdentityRetainedBytes < 160000000L)
   (sprintf "sort identity cache retained %.1f MB" (float sortIdentityRetainedBytes / 1024.0 / 1024.0))
let sortIdentityTextRetainedModelCount = 80
let sortIdentityTextRetainedValue modelIndex =
    sprintf "sort-retained-text-%03d-%s" modelIndex (System.String('x', 512 * 1024))
let sortIdentityTextRetainedModel modelIndex =
    { createEmpty 10 2 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 0.0)
              { Row = 0; Col = 1 }, cell (CellValue.String (sortIdentityTextRetainedValue modelIndex)) ]
            |> Map.ofList
        SortState = Some { Column = 0; Ascending = true } }
let sortIdentityTextRetainedBytes, sortIdentityTextRetainedTotal =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable total = 0
    for modelIndex in 0 .. sortIdentityTextRetainedModelCount - 1 do
        let sorted = sortByColumn 0 true (sortIdentityTextRetainedModel modelIndex)
        if sorted.SortState.IsSome then
            total <- total + 1
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before, total
ok "sort-identity-cache-bounded-text-memory"
   (sortIdentityTextRetainedBytes < 40000000L
    && sortIdentityTextRetainedTotal = sortIdentityTextRetainedModelCount)
   (sprintf "sort identity text cache retained %.1f MB" (float sortIdentityTextRetainedBytes / 1024.0 / 1024.0))
let sortIdentityOversizedSourceCellsWeakReference () =
    let payload = System.String('S', 6 * 1024 * 1024)
    let cells =
        [ { Row = 0; Col = 0 }, cell (CellValue.Number 0.0)
          { Row = 0; Col = 1 }, cell (CellValue.String payload) ]
        |> Map.ofList
    let model =
        { createEmpty 2 2 with
            Cells = cells
            SortState = Some { Column = 0; Ascending = true } }
    sortByColumn 0 true model |> ignore
    System.WeakReference(cells)
let sortIdentityOversizedSourceCellsWeak = sortIdentityOversizedSourceCellsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "sort-identity-cache-skips-oversized-source"
   (not sortIdentityOversizedSourceCellsWeak.IsAlive)
   "sort identity cache should not retain an oversized source cell map"
let sortIdentityAllocationRowCount = 50000
let sortIdentityAllocationModel =
    { createEmpty sortIdentityAllocationRowCount 2 with
        Cells =
            [ for row in 0 .. sortIdentityAllocationRowCount - 1 do
                yield { Row = row; Col = 0 }, cell (CellValue.Number (float row))
                yield { Row = row; Col = 1 }, cell (CellValue.Number (float row)) ]
            |> Map.ofList
        SortState = Some { Column = 0; Ascending = true } }
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let sortIdentityAllocationBefore = System.GC.GetAllocatedBytesForCurrentThread()
let sortIdentityAllocationResult = sortByColumn 0 true sortIdentityAllocationModel
let sortIdentityAllocationBytes =
    System.GC.GetAllocatedBytesForCurrentThread() - sortIdentityAllocationBefore
ok "sort-identity-streams-noop"
   (obj.ReferenceEquals(sortIdentityAllocationResult.Cells, sortIdentityAllocationModel.Cells)
    && sortIdentityAllocationBytes < 2500000L)
   (sprintf "already-sorted identity sort allocated %d bytes" sortIdentityAllocationBytes)
let sortedDesc = sortByColumn 0 false sortedModel
ok "sort-desc-reorders-cells"
   (not (obj.ReferenceEquals(sortedDesc.Cells, sortedModel.Cells))
    && getCellValue { Row = 0; Col = 1 } sortedDesc = CellValue.String "b"
    && getCellValue { Row = 1; Col = 1 } sortedDesc = CellValue.String "a")
   "descending sort should still reorder rows"
let descSortedModel =
    { createEmpty 10 3 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 0; Col = 1 }, cell (CellValue.String "b")
              { Row = 1; Col = 0 }, cell (CellValue.Number 1.0)
              { Row = 1; Col = 1 }, cell (CellValue.String "a") ]
            |> Map.ofList }
let descSortedAgain = sortByColumn 0 false descSortedModel
ok "sort-desc-identity-preserves-cells"
   (obj.ReferenceEquals(descSortedAgain.Cells, descSortedModel.Cells)
    && descSortedAgain.SortState = Some { Column = 0; Ascending = false })
   "already-descending contiguous rows should not rebuild the cell map"
ok "sort-desc-identity-preserves-model"
   (obj.ReferenceEquals(sortByColumn 0 false descSortedAgain, descSortedAgain))
   "repeating the same already-descending sort should return the same model"
let gappedSortedModel =
    { createEmpty 10 3 with
        Cells =
            [ { Row = 1; Col = 0 }, cell (CellValue.Number 1.0)
              { Row = 1; Col = 1 }, cell (CellValue.String "a")
              { Row = 3; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 3; Col = 1 }, cell (CellValue.String "b") ]
            |> Map.ofList }
let gappedSorted = sortByColumn 0 true gappedSortedModel
ok "sort-gapped-still-compacts"
   (not (obj.ReferenceEquals(gappedSorted.Cells, gappedSortedModel.Cells))
    && getCellValue { Row = 0; Col = 1 } gappedSorted = CellValue.String "a"
    && getCellValue { Row = 1; Col = 1 } gappedSorted = CellValue.String "b")
   "already-sorted gapped rows should still compact to contiguous sorted rows"
let missingSortColumnModel =
    { createEmpty 10 3 with
        Cells =
            [ { Row = 0; Col = 1 }, cell (CellValue.String "a")
              { Row = 1; Col = 1 }, cell (CellValue.String "b") ]
            |> Map.ofList }
let missingSortColumnSorted = sortByColumn 0 true missingSortColumnModel
ok "sort-missing-column-empty-identity"
   (obj.ReferenceEquals(missingSortColumnSorted.Cells, missingSortColumnModel.Cells)
    && missingSortColumnSorted.SortState = Some { Column = 0; Ascending = true })
   "rows without sort-column cells should sort by Empty without a cell-map rebuild when already contiguous"

// usedBounds is cached by immutable cell/merge identity. Repeated reads should
// preserve the value, while cell edits and merged-range edits must invalidate.
let usedBoundsBase = createEmpty 10 10
ok "usedbounds-empty" (usedBounds usedBoundsBase = None)
   (sprintf "empty usedBounds got %A" (usedBounds usedBoundsBase))
let usedBoundsWithCell =
    usedBoundsBase
    |> setCell { Row = 3; Col = 4 } (cell (CellValue.Number 1.0))
ok "usedbounds-cell" (usedBounds usedBoundsWithCell = Some (3, 4) && usedBounds usedBoundsWithCell = Some (3, 4))
   (sprintf "cell usedBounds got %A" (usedBounds usedBoundsWithCell))
let usedBoundsWithMerge = usedBoundsWithCell |> addMergedCell 6 7 8 9
ok "usedbounds-merge-invalidates" (usedBounds usedBoundsWithMerge = Some (8, 9))
   (sprintf "merge usedBounds got %A" (usedBounds usedBoundsWithMerge))
let usedBoundsAfterClear = usedBoundsWithMerge |> clearCell { Row = 3; Col = 4 }
ok "usedbounds-cell-clear-invalidates" (usedBounds usedBoundsAfterClear = Some (8, 9))
   (sprintf "clear usedBounds got %A" (usedBounds usedBoundsAfterClear))
let usedBoundsManyModelCount = 24
let usedBoundsManyRowCount = 8000
let usedBoundsManyModels =
    [| for modelIndex in 0 .. usedBoundsManyModelCount - 1 ->
        { createEmpty usedBoundsManyRowCount 8 with
            Cells =
                [ for row in 0 .. usedBoundsManyRowCount - 1 ->
                    { Row = row; Col = row % 8 }, cell (CellValue.String (sprintf "bounds-%02d-%05d" modelIndex row)) ]
                |> Map.ofList } |]
let usedBoundsManyRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in usedBoundsManyModels do
            match usedBounds model with
            | Some (row, col) -> total <- total + row + col
            | None -> ()
    total
usedBoundsManyRun 3 |> ignore
let usedBoundsManySw = Stopwatch.StartNew()
let usedBoundsManyTotal = usedBoundsManyRun 60
usedBoundsManySw.Stop()
ok "usedbounds-cache-many-models"
   (usedBoundsManySw.Elapsed.TotalMilliseconds < 15.0 && usedBoundsManyTotal > 0)
   (sprintf "usedBounds cache should retain many cell maps, took %.1f ms" usedBoundsManySw.Elapsed.TotalMilliseconds)
let lastUsedColumnManyRun passes =
    let mutable total = 0
    let targetRow = usedBoundsManyRowCount - 1
    for _pass in 1 .. passes do
        for model in usedBoundsManyModels do
            match lastUsedColumnInRow targetRow model with
            | Some col -> total <- total + col
            | None -> ()
    total
lastUsedColumnManyRun 3 |> ignore
let lastUsedColumnManySw = Stopwatch.StartNew()
let lastUsedColumnManyTotal = lastUsedColumnManyRun 60
lastUsedColumnManySw.Stop()
ok "last-used-column-cache-many-models"
   (lastUsedColumnManySw.Elapsed.TotalMilliseconds < 15.0 && lastUsedColumnManyTotal > 0)
   (sprintf "last used column cache should retain many cell maps, took %.1f ms" lastUsedColumnManySw.Elapsed.TotalMilliseconds)
let lastUsedDistinctRowCount = 100000
let lastUsedDistinctRows =
    [| for index in 0 .. 99 -> index * 997 |]
let lastUsedDistinctModel =
    { createEmpty lastUsedDistinctRowCount 20 with
        Cells =
            [ for row in 0 .. lastUsedDistinctRowCount - 1 ->
                { Row = row; Col = row % 20 }, cell (CellValue.Number (float row)) ]
            |> Map.ofList }
let lastUsedDistinctRowsRun () =
    let mutable total = 0
    for row in lastUsedDistinctRows do
        match lastUsedColumnInRow row lastUsedDistinctModel with
        | Some col -> total <- total + col
        | None -> ()
    total
lastUsedDistinctRowsRun () |> ignore
let lastUsedDistinctSw = Stopwatch.StartNew()
let lastUsedDistinctTotal = lastUsedDistinctRowsRun ()
lastUsedDistinctSw.Stop()
ok "last-used-column-cache-distinct-rows"
   (lastUsedDistinctSw.Elapsed.TotalMilliseconds < 5.0 && lastUsedDistinctTotal > 0)
   (sprintf "last used column cache should serve distinct rows without rescanning, took %.1f ms" lastUsedDistinctSw.Elapsed.TotalMilliseconds)
let lastUsedMergeDistinctMergeCount = 50000
let lastUsedMergeDistinctRows =
    [| for index in 0 .. 199 -> index * 251 |]
let lastUsedMergeDistinctModel =
    { createEmpty (lastUsedMergeDistinctMergeCount + 1) 32 with
        MergedCells =
            [ for row in 0 .. lastUsedMergeDistinctMergeCount - 1 ->
                { StartRow = row
                  StartCol = row % 8
                  EndRow = row
                  EndCol = (row % 8) + 12 } ] }
let lastUsedMergeDistinctRowsRun () =
    let mutable total = 0
    for row in lastUsedMergeDistinctRows do
        match lastUsedColumnInRow row lastUsedMergeDistinctModel with
        | Some col -> total <- total + col
        | None -> ()
    total
lastUsedMergeDistinctRowsRun () |> ignore
let lastUsedMergeDistinctSw = Stopwatch.StartNew()
let lastUsedMergeDistinctTotal = lastUsedMergeDistinctRowsRun ()
lastUsedMergeDistinctSw.Stop()
ok "last-used-column-cache-distinct-merge-rows"
   (lastUsedMergeDistinctSw.Elapsed.TotalMilliseconds < 5.0 && lastUsedMergeDistinctTotal > 0)
   (sprintf "last used column merge cache should serve distinct rows without rescanning, took %.1f ms" lastUsedMergeDistinctSw.Elapsed.TotalMilliseconds)
let maxUsedColumnManyModels =
    [| for modelIndex in 0 .. usedBoundsManyModelCount - 1 ->
        { createEmpty usedBoundsManyRowCount 16 with
            Cells =
                [ for row in 0 .. usedBoundsManyRowCount - 1 ->
                    { Row = row; Col = row % 8 }, cell (CellValue.String (sprintf "maxcol-%02d-%05d" modelIndex row)) ]
                |> Map.ofList } |]
let maxUsedColumnManyRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in maxUsedColumnManyModels do
            let inserted = insertColumn 12 model
            if obj.ReferenceEquals(inserted.Cells, model.Cells) then
                total <- total + inserted.ColCount
    total
maxUsedColumnManyRun 3 |> ignore
let maxUsedColumnManySw = Stopwatch.StartNew()
let maxUsedColumnManyTotal = maxUsedColumnManyRun 60
maxUsedColumnManySw.Stop()
ok "max-used-column-cache-many-models"
   (maxUsedColumnManySw.Elapsed.TotalMilliseconds < 15.0 && maxUsedColumnManyTotal > 0)
   (sprintf "max used column cache should retain many cell maps, took %.1f ms" maxUsedColumnManySw.Elapsed.TotalMilliseconds)
let boundsCacheRetainedRowCount = 20000
let boundsCacheRetainedModelCount = 80
let boundsCacheRetainedModel modelIndex =
    { createEmpty boundsCacheRetainedRowCount 16 with
        Cells =
            [ for row in 0 .. boundsCacheRetainedRowCount - 1 ->
                { Row = row; Col = row % 8 }, cell (CellValue.String (sprintf "bounds-retained-%02d-%05d" modelIndex row)) ]
            |> Map.ofList }
let boundsCacheRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. boundsCacheRetainedModelCount - 1 do
        let model = boundsCacheRetainedModel modelIndex
        usedBounds model |> ignore
        lastUsedColumnInRow (boundsCacheRetainedRowCount - 1) model |> ignore
        insertColumn 12 model |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "bounds-column-caches-bounded-source-memory"
   (boundsCacheRetainedBytes < 160000000L)
   (sprintf "bounds/used-column caches retained %.1f MB" (float boundsCacheRetainedBytes / 1024.0 / 1024.0))
let boundsCacheSourceTextRetainedModel modelIndex =
    { createEmpty 4 16 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.String (System.String(char (65 + (modelIndex % 26)), 2000000))) ]
            |> Map.ofList }
let boundsCacheSourceTextRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. 39 do
        let model = boundsCacheSourceTextRetainedModel modelIndex
        usedBounds model |> ignore
        lastUsedColumnInRow 0 model |> ignore
        insertColumn 12 model |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "bounds-column-caches-source-text-memory"
   (boundsCacheSourceTextRetainedBytes < 30000000L)
   (sprintf "bounds/used-column source text retained %.1f MB" (float boundsCacheSourceTextRetainedBytes / 1024.0 / 1024.0))
let usedBoundsOversizedSourceCellsWeakReference () =
    let payload = System.String('U', 6 * 1024 * 1024)
    let cells =
        [ { Row = 0; Col = 0 }, cell (CellValue.String payload) ]
        |> Map.ofList
    let model = { createEmpty 4 16 with Cells = cells }
    usedBounds model |> ignore
    System.WeakReference(cells)
let usedBoundsOversizedSourceCellsWeak = usedBoundsOversizedSourceCellsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "usedbounds-cache-skips-oversized-source"
   (not usedBoundsOversizedSourceCellsWeak.IsAlive)
   "usedBounds cache should not retain an oversized source cell map"
let lastUsedColumnOversizedSourceCellsWeakReference () =
    let payload = System.String('L', 6 * 1024 * 1024)
    let cells =
        [ { Row = 0; Col = 0 }, cell (CellValue.String payload) ]
        |> Map.ofList
    let model = { createEmpty 4 16 with Cells = cells }
    lastUsedColumnInRow 0 model |> ignore
    System.WeakReference(cells)
let lastUsedColumnOversizedSourceCellsWeak = lastUsedColumnOversizedSourceCellsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "last-used-column-cache-skips-oversized-source"
   (not lastUsedColumnOversizedSourceCellsWeak.IsAlive)
   "last-used-column cache should not retain an oversized source cell map"
let maxUsedColumnOversizedSourceCellsWeakReference () =
    let payload = System.String('M', 6 * 1024 * 1024)
    let cells =
        [ { Row = 0; Col = 0 }, cell (CellValue.String payload) ]
        |> Map.ofList
    let model = { createEmpty 4 16 with Cells = cells }
    insertColumn 12 model |> ignore
    System.WeakReference(cells)
let maxUsedColumnOversizedSourceCellsWeak = maxUsedColumnOversizedSourceCellsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "max-used-column-cache-skips-oversized-source"
   (not maxUsedColumnOversizedSourceCellsWeak.IsAlive)
   "max-used-column cache should not retain an oversized source cell map"

// Recalc's no-formula guard caches by immutable cell-map identity. It must
// propagate through set/clear without suppressing newly added formulas or
// keeping stale formula presence after the last formula is removed.
let noFormulaCached =
    createEmpty 10 5
    |> setCell { Row = 0; Col = 0 } (cell (CellValue.Number 1.0))
    |> recalcAndSpill None None
let noFormulaDirectRecalc = recalculateAllFormulas None None noFormulaCached
ok "recalc-direct-no-formula-cache"
   (obj.ReferenceEquals(noFormulaDirectRecalc.Cells, noFormulaCached.Cells))
   "direct full recalc should reuse known no-formula cell map"
let formulaAfterCachedNoFormula =
    noFormulaCached
    |> setCell { Row = 0; Col = 1 } (cell (CellValue.Formula ("=A1+2", CellValue.Empty)))
    |> recalcAndSpill None None
let b1AfterCachedNoFormula =
    match getCellValue { Row = 0; Col = 1 } formulaAfterCachedNoFormula with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "recalc-guard-formula-add" (b1AfterCachedNoFormula = CellValue.Number 3.0)
   (sprintf "new formula after cached no-formula state got %A" b1AfterCachedNoFormula)
let noFormulaAfterClear =
    formulaAfterCachedNoFormula
    |> clearCell { Row = 0; Col = 1 }
    |> recalcAndSpill None None
ok "recalc-guard-formula-clear" (Map.containsKey { Row = 0; Col = 1 } noFormulaAfterClear.Cells |> not)
   "clearing only formula should leave B1 empty after guarded recalc"
let formulaCountAfterPartialClear =
    { createEmpty 10 5 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Formula ("=1", CellValue.Empty))
              { Row = 0; Col = 1 }, cell (CellValue.Formula ("=A1+1", CellValue.Empty)) ]
            |> Map.ofList }
    |> recalculateAllFormulas None None
    |> clearCell { Row = 0; Col = 1 }
    |> setCell { Row = 0; Col = 0 } (cell (CellValue.Formula ("=2", CellValue.Empty)))
    |> recalcAndSpill None None
let formulaCountA1 =
    match getCellValue { Row = 0; Col = 0 } formulaCountAfterPartialClear with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "recalc-guard-formula-count" (formulaCountA1 = CellValue.Number 2.0)
   (sprintf "remaining formula after partial clear should recalc to 2, got %A" formulaCountA1)
let clearCellCacheCount = 8000
let clearCellCacheCells =
    seq {
        for row in 0 .. clearCellCacheCount - 1 do
            yield { Row = row; Col = 0 }, cell (CellValue.Formula (sprintf "=%d" row, CellValue.Number (float row)))
    }
    |> Map.ofSeq
let mutable clearCellCacheModel = { createEmpty (clearCellCacheCount + 1) 2 with Cells = clearCellCacheCells }
cellsHaveFormulaCells clearCellCacheModel.Cells |> ignore
let clearCellCacheSw = Stopwatch.StartNew()
for row in 0 .. clearCellCacheCount - 1 do
    clearCellCacheModel <- clearCell { Row = row; Col = 0 } clearCellCacheModel
clearCellCacheSw.Stop()
ok "clear-cell-formula-cache-time"
   (clearCellCacheModel.Cells.Count = 0
    && not (cellsHaveFormulaCells clearCellCacheModel.Cells)
    && clearCellCacheSw.Elapsed.TotalMilliseconds < 1500.0)
   (sprintf "cleared %d formulas in %.1f ms, remaining=%d hasFormulas=%b"
        clearCellCacheCount
        clearCellCacheSw.Elapsed.TotalMilliseconds
        clearCellCacheModel.Cells.Count
        (cellsHaveFormulaCells clearCellCacheModel.Cells))

let settledFormulaModel =
    { createEmpty 10 5 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 1.0)
              { Row = 0; Col = 1 }, cell (CellValue.Formula ("=A1+2", CellValue.Number 3.0)) ]
            |> Map.ofList }
let settledFormulaRecalc = recalculateAllFormulas None None settledFormulaModel
ok "recalc-unchanged-preserves-cells"
   (obj.ReferenceEquals(settledFormulaModel.Cells, settledFormulaRecalc.Cells))
   "unchanged formula results should not rebuild the cell map"
ok "recalc-unchanged-preserves-model"
   (obj.ReferenceEquals(settledFormulaModel, settledFormulaRecalc))
   "unchanged formula results should not rebuild the spreadsheet model"
let changedFormulaRecalc =
    settledFormulaModel
    |> setCell { Row = 0; Col = 0 } (cell (CellValue.Number 5.0))
    |> recalculateAllFormulas None None
let changedFormulaValue =
    match getCellValue { Row = 0; Col = 1 } changedFormulaRecalc with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "recalc-changed-still-updates" (changedFormulaValue = CellValue.Number 7.0)
   (sprintf "changed source should update formula to 7, got %A" changedFormulaValue)
let changedFormulaTextRecalc =
    settledFormulaModel
    |> setCell { Row = 0; Col = 1 } (cell (CellValue.Formula ("=A1+3", CellValue.Empty)))
    |> recalculateAllFormulas None None
let changedFormulaTextValue =
    match getCellValue { Row = 0; Col = 1 } changedFormulaTextRecalc with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "recalc-cache-formula-text-change" (changedFormulaTextValue = CellValue.Number 4.0)
   (sprintf "changed formula text should parse/evaluate to 4, got %A" changedFormulaTextValue)

// Recalc reuses parse results within one pass for repeated formula text.
// Both successful parses and parse errors must keep the same cell-visible result.
let duplicateFormulaModel =
    { createEmpty 10 5 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 0; Col = 1 }, cell (CellValue.Formula ("=A1+3", CellValue.Empty))
              { Row = 1; Col = 1 }, cell (CellValue.Formula ("=A1+3", CellValue.Empty))
              { Row = 0; Col = 2 }, cell (CellValue.Formula ("=SUM(", CellValue.Empty))
              { Row = 1; Col = 2 }, cell (CellValue.Formula ("=SUM(", CellValue.Empty)) ]
            |> Map.ofList }
let duplicateFormulaResult = recalculateAllFormulas None None duplicateFormulaModel
let duplicateValue addr =
    match getCellValue addr duplicateFormulaResult with
    | CellValue.Formula (_, v) -> v
    | v -> v
let duplicateSuccess = duplicateValue { Row = 1; Col = 1 }
let duplicateParseError = duplicateValue { Row = 1; Col = 2 }
ok "recalc-parse-cache-success" (duplicateSuccess = CellValue.Number 5.0)
   (sprintf "expected 5 got %A" duplicateSuccess)
ok "recalc-parse-cache-error" (match duplicateParseError with CellValue.Error _ -> true | _ -> false)
   (sprintf "expected parse error got %A" duplicateParseError)
let recalcLocalCacheChurnCount = 9000
let recalcLocalCacheChurnModel =
    { createEmpty (recalcLocalCacheChurnCount + 1) 4 with
        Cells =
            [ yield { Row = 0; Col = 0 }, cell (CellValue.Number 10.0)
              for row in 1 .. recalcLocalCacheChurnCount ->
                { Row = row; Col = 1 }, cell (CellValue.Formula (sprintf "=A$1+%d" row, CellValue.Empty)) ]
            |> Map.ofList }
let recalcLocalCacheChurnResult = recalculateAllFormulas None None recalcLocalCacheChurnModel
let recalcLocalCacheChurnValue =
    match getCellValue { Row = recalcLocalCacheChurnCount; Col = 1 } recalcLocalCacheChurnResult with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "recalc-local-caches-churn" (recalcLocalCacheChurnValue = CellValue.Number (float (recalcLocalCacheChurnCount + 10)))
   (sprintf "expected local cache churn result %d got %A" (recalcLocalCacheChurnCount + 10) recalcLocalCacheChurnValue)

// Repeated formula text can also share dependency extraction. The duplicate
// formulas below must still evaluate after every formula cell in the range.
let duplicateDependencyModel =
    { createEmpty 10 5 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Formula ("=1", CellValue.Empty))
              { Row = 1; Col = 0 }, cell (CellValue.Formula ("=A1+1", CellValue.Empty))
              { Row = 0; Col = 1 }, cell (CellValue.Formula ("=SUM(A1:A2)+A1", CellValue.Empty))
              { Row = 1; Col = 1 }, cell (CellValue.Formula ("=SUM(A1:A2)+A1", CellValue.Empty)) ]
            |> Map.ofList }
let duplicateDependencyResult = recalculateAllFormulas None None duplicateDependencyModel
let duplicateDependencyValue =
    match getCellValue { Row = 1; Col = 1 } duplicateDependencyResult with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "recalc-dependency-cache" (duplicateDependencyValue = CellValue.Number 4.0)
   (sprintf "expected duplicate range dependency result 4 got %A" duplicateDependencyValue)

// Repeated criteria formulas should share criteria predicate compilation within
// a recalc pass while preserving wildcard semantics.
let duplicateCriteriaModel =
    { createEmpty 20 5 with
        Cells =
            [ for r in 0 .. 5 ->
                { Row = r; Col = 0 }, cell (CellValue.String (if r % 2 = 0 then "apple" else "banana"))
              for r in 0 .. 2 ->
                { Row = r; Col = 2 }, cell (CellValue.Formula ("=COUNTIF(A1:A6,\"app*\")", CellValue.Empty)) ]
            |> Map.ofList }
let duplicateCriteriaResult = recalculateAllFormulas None None duplicateCriteriaModel
let duplicateCriteriaValue =
    match getCellValue { Row = 2; Col = 2 } duplicateCriteriaResult with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "recalc-criteria-cache" (duplicateCriteriaValue = CellValue.Number 3.0)
   (sprintf "expected wildcard COUNTIF result 3 got %A" duplicateCriteriaValue)
let criteriaCacheChurnCount = 2100
let criteriaCacheChurnModel =
    { createEmpty (criteriaCacheChurnCount + 1) 4 with
        Cells =
            [ yield { Row = 0; Col = 0 }, cell (CellValue.String "needle")
              for row in 1 .. criteriaCacheChurnCount ->
                { Row = row; Col = 1 }, cell (CellValue.Formula (sprintf "=COUNTIF(A$1:A$1,\"miss-%04d\")" row, CellValue.Empty)) ]
            |> Map.ofList }
let criteriaCacheChurnResult = recalculateAllFormulas None None criteriaCacheChurnModel
let criteriaCacheChurnValue =
    match getCellValue { Row = criteriaCacheChurnCount; Col = 1 } criteriaCacheChurnResult with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "recalc-criteria-cache-churn" (criteriaCacheChurnValue = CellValue.Number 0.0)
   (sprintf "expected criteria cache churn result 0 got %A" criteriaCacheChurnValue)

let rangeReadCacheStaticFlags = BindingFlags.Static ||| BindingFlags.Public ||| BindingFlags.NonPublic
let rangeReadCacheInstanceFlags = BindingFlags.Instance ||| BindingFlags.Public ||| BindingFlags.NonPublic
let formulaModuleType = typeof<RangeAccess>.DeclaringType
let rangeReadCacheMaterializeAfter =
    formulaModuleType.GetMethod("get_rangeReadCacheMaterializeAfter", rangeReadCacheStaticFlags).Invoke(null, [||]) :?> int
let rangeReadCacheMaxMaterializedValues =
    formulaModuleType.GetMethod("get_rangeReadCacheMaxMaterializedValues", rangeReadCacheStaticFlags).Invoke(null, [||]) :?> int
let rangeReadCachedMethod =
    formulaModuleType.GetMethod("rangeReadCached", rangeReadCacheStaticFlags)
let rangeReadCacheType =
    formulaModuleType.GetNestedType("RangeReadCache", BindingFlags.Public ||| BindingFlags.NonPublic)
let rangeReadCacheReadMethod =
    rangeReadCacheType.GetMethod("Read", rangeReadCacheInstanceFlags)
let rangeReadCacheValuesField =
    rangeReadCacheType.GetField("values", rangeReadCacheInstanceFlags)
let rangeReadSource count =
    { M =
        { createEmpty count 1 with
            Cells = [ { Row = count - 1; Col = 0 }, cell (CellValue.Number (float count)) ] |> Map.ofList }
      SR = 0
      SC = 0
      W = 1
      Count = count }
let createRangeReadCache source count =
    rangeReadCachedMethod.Invoke(null, [| box source; box count |])
let readRangeReadCache cache k =
    rangeReadCacheReadMethod.Invoke(cache, [| box k |]) :?> CellValue
let rangeReadCacheValues cache =
    rangeReadCacheValuesField.GetValue(cache) :?> CellValue[]

let smallRangeReadCount = rangeReadCacheMaterializeAfter + 1
let smallRangeReadCache = createRangeReadCache (rangeReadSource smallRangeReadCount) smallRangeReadCount
let mutable smallRangeReadLastValue = CellValue.Empty
for k in 0 .. rangeReadCacheMaterializeAfter do
    smallRangeReadLastValue <- readRangeReadCache smallRangeReadCache k
let smallRangeReadValues = rangeReadCacheValues smallRangeReadCache
ok "range-read-cache-small-materializes"
   (not (isNull smallRangeReadValues)
    && smallRangeReadValues.Length = smallRangeReadCount
    && smallRangeReadLastValue = CellValue.Number (float smallRangeReadCount))
   (sprintf "small repeated range reads should materialize, values=%b last=%A" (not (isNull smallRangeReadValues)) smallRangeReadLastValue)

let largeRangeReadCount = rangeReadCacheMaxMaterializedValues + 1
let largeRangeReadCache = createRangeReadCache (rangeReadSource largeRangeReadCount) largeRangeReadCount
let mutable largeRangeReadLastValue = CellValue.Empty
for k in 0 .. rangeReadCacheMaterializeAfter do
    let index = if k = rangeReadCacheMaterializeAfter then largeRangeReadCount - 1 else k
    largeRangeReadLastValue <- readRangeReadCache largeRangeReadCache index
let largeRangeReadValues = rangeReadCacheValues largeRangeReadCache
ok "range-read-cache-large-streams"
   (isNull largeRangeReadValues && largeRangeReadLastValue = CellValue.Number (float largeRangeReadCount))
   (sprintf "large repeated range reads should stream, values=%b last=%A" (not (isNull largeRangeReadValues)) largeRangeReadLastValue)

// Repeated wildcard lookup formulas should share wildcard regex translation
// while preserving XMATCH wildcard semantics.
let duplicateWildcardLookupModel =
    { createEmpty 20 5 with
        Cells =
            [ yield { Row = 0; Col = 0 }, cell (CellValue.String "apple")
              yield { Row = 1; Col = 0 }, cell (CellValue.String "banana")
              yield { Row = 2; Col = 0 }, cell (CellValue.String "berry")
              for r in 0 .. 2 ->
                { Row = r; Col = 2 }, cell (CellValue.Formula ("=XMATCH(\"ban*\",A1:A3,2)", CellValue.Empty)) ]
            |> Map.ofList }
let duplicateWildcardLookupResult = recalculateAllFormulas None None duplicateWildcardLookupModel
let duplicateWildcardLookupValue =
    match getCellValue { Row = 2; Col = 2 } duplicateWildcardLookupResult with
    | CellValue.Formula (_, v) -> v
    | v -> v
ok "recalc-wildcard-lookup-cache" (duplicateWildcardLookupValue = CellValue.Number 2.0)
   (sprintf "expected wildcard XMATCH result 2 got %A" duplicateWildcardLookupValue)

let wildcardCachePatternCount = 300
let wildcardCacheModel =
    { createEmpty (wildcardCachePatternCount + 5) 2 with
        Cells =
            [ for r in 0 .. wildcardCachePatternCount - 1 ->
                { Row = r; Col = 0 }, cell (CellValue.String (sprintf "item%03d-value" r)) ]
            |> Map.ofList }
let wildcardCacheValues () =
    [ for i in 0 .. wildcardCachePatternCount - 1 ->
        evaluateFormula (sprintf "=COUNTIF(A1:A%d,\"item%03d*\")" wildcardCachePatternCount i) wildcardCacheModel None ]
let wildcardCacheFirst = wildcardCacheValues ()
let wildcardCacheSecond = wildcardCacheValues ()
ok "wildcard-cache-many-distinct"
   (wildcardCacheFirst = wildcardCacheSecond
    && List.forall ((=) (CellValue.Number 1.0)) wildcardCacheSecond)
   "distinct wildcard criteria should remain stable past the old regex clear threshold"
let largeWildcardModel =
    { createEmpty 1 1 with
        Cells = Map.ofList [ { Row = 0; Col = 0 }, cell (CellValue.String "needle") ] }
let largeWildcardPattern index =
    System.String('x', 10000) + sprintf "%04d*" index
let largeWildcardRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for index in 0 .. 1999 do
        evaluateFormula (sprintf "=COUNTIF(A1:A1,\"%s\")" (largeWildcardPattern index)) largeWildcardModel None |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "wildcard-regex-cache-bounded-source-memory"
   (largeWildcardRetainedBytes < 40000000L)
   (sprintf "large wildcard regex cache retained %.1f MB" (float largeWildcardRetainedBytes / 1024.0 / 1024.0))
let oversizedTranslatedWildcardWeakReference () =
    let pattern = System.String('.', 350000) + "*"
    let model =
        { createEmpty 1 2 with
            Cells =
                [ { Row = 0; Col = 0 }, cell (CellValue.String "needle")
                  { Row = 0; Col = 1 }, cell (CellValue.String pattern) ]
                |> Map.ofList }
    evaluateFormula "=COUNTIF(A1:A1,B1)" model None |> ignore
    System.WeakReference(pattern)
let oversizedTranslatedWildcardWeak = oversizedTranslatedWildcardWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "wildcard-regex-cache-skips-translated-pattern"
   (not oversizedTranslatedWildcardWeak.IsAlive)
   "wildcard regex cache should count translated regex text, not only the source pattern"
let largeFormulaSource index =
    "=" + "\"" + System.String('x', 5000) + sprintf "%04d" index + "\""
let largeFormulaRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for index in 0 .. 1999 do
        evaluateFormula (largeFormulaSource index) wildcardCacheModel None |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "formula-parse-cache-bounded-source-memory"
   (largeFormulaRetainedBytes < 40000000L)
   (sprintf "large formula parse cache retained %.1f MB" (float largeFormulaRetainedBytes / 1024.0 / 1024.0))
let parseFormulaSharedMethod =
    formulaModuleType.GetMethod("parseFormulaShared", BindingFlags.NonPublic ||| BindingFlags.Static)
let collectFormulaCellsMethod =
    formulaModuleType.GetMethod("collectFormulaCells", BindingFlags.NonPublic ||| BindingFlags.Static)
let largeFormulaCacheHitSource index =
    let terms = [| for i in 1 .. 1800 -> string ((i + index) % 97 + 1) |]
    "=" + System.String.Join("+", terms)
if isNull parseFormulaSharedMethod then
    ok "large-formula-parse-cache-hit-perf"
       false
       "expected private parseFormulaShared method"
else
    let cacheHitFormulas = [| for index in 0 .. 63 -> largeFormulaCacheHitSource index |]
    for formula in cacheHitFormulas do
        parseFormulaSharedMethod.Invoke(null, [| box formula |]) |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let cacheHitSw = Stopwatch.StartNew()
    let mutable cacheHitCount = 0
    for _pass in 1 .. 3000 do
        for formula in cacheHitFormulas do
            if not (isNull (parseFormulaSharedMethod.Invoke(null, [| box formula |]))) then
                cacheHitCount <- cacheHitCount + 1
    cacheHitSw.Stop()
    ok "large-formula-parse-cache-hit-perf"
       (cacheHitSw.Elapsed.TotalMilliseconds < 140.0 && cacheHitCount = 192000)
       (sprintf "large formula parse cache hits took %.1f ms" cacheHitSw.Elapsed.TotalMilliseconds)
if isNull collectFormulaCellsMethod then
    ok "recalc-local-parse-cache-hit-perf"
       false
       "expected private collectFormulaCells method"
else
    let repeatedFormula = largeFormulaCacheHitSource 0
    let repeatedFormulaCells =
        [ for row in 0 .. 15999 ->
            { Row = row; Col = 0 }, cell (CellValue.Formula (repeatedFormula, CellValue.Empty)) ]
        |> Map.ofList
    let repeatedFormulaModel = { createEmpty 16000 1 with Cells = repeatedFormulaCells }
    collectFormulaCellsMethod.Invoke(null, [| box repeatedFormulaModel |]) |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let recalcParseSw = Stopwatch.StartNew()
    let mutable recalcParseHits = 0
    for _pass in 1 .. 12 do
        let result = collectFormulaCellsMethod.Invoke(null, [| box repeatedFormulaModel |])
        recalcParseHits <-
            recalcParseHits
            + (result.GetType().GetProperty("Count").GetValue(result) :?> int)
    recalcParseSw.Stop()
    ok "recalc-local-parse-cache-hit-perf"
       (recalcParseSw.Elapsed.TotalMilliseconds < 120.0 && recalcParseHits = 192000)
       (sprintf "recalc local parse cache hits took %.1f ms" recalcParseSw.Elapsed.TotalMilliseconds)
let recalcReferenceFormula = "=" + String.replicate 5200 " " + "0"
let recalcReferenceCells =
    [ for row in 0 .. 19999 ->
        { Row = row; Col = 0 }, cell (CellValue.Formula (recalcReferenceFormula, CellValue.Empty)) ]
    |> Map.ofList
let recalcReferenceModel = { createEmpty 20000 1 with Cells = recalcReferenceCells }
recalculateAllFormulas None None recalcReferenceModel |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let recalcReferenceSw = Stopwatch.StartNew()
let mutable recalcReferenceCellsSeen = 0
for _pass in 1 .. 10 do
    let result = recalculateAllFormulas None None recalcReferenceModel
    recalcReferenceCellsSeen <- recalcReferenceCellsSeen + result.Cells.Count
recalcReferenceSw.Stop()
ok "recalc-reference-cache-hit-perf"
   (recalcReferenceSw.Elapsed.TotalMilliseconds < 520.0 && recalcReferenceCellsSeen = 200000)
   (sprintf "recalc reference cache hits took %.1f ms" recalcReferenceSw.Elapsed.TotalMilliseconds)

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
ok "cycle-detected" (match f1 with CellValue.Error _ -> true | _ -> false)
   (sprintf "expected cycle error, got %A" f1)

printfn ""
printfn "Stage 4 perf/correctness tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
