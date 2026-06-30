#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Formula

// Build a tiny model with some seed cells: A1=10, A2=20, A3=text "hi",
// A4=TRUE, B1=3.
let cells =
    [ { Row=0; Col=0 }, { Value = CellValue.Number 10.0; Format = defaultFormat; Comment = None }
      { Row=1; Col=0 }, { Value = CellValue.Number 20.0; Format = defaultFormat; Comment = None }
      { Row=2; Col=0 }, { Value = CellValue.String "hi"; Format = defaultFormat; Comment = None }
      { Row=3; Col=0 }, { Value = CellValue.Boolean true; Format = defaultFormat; Comment = None }
      { Row=0; Col=1 }, { Value = CellValue.Number 3.0; Format = defaultFormat; Comment = None } ]
    |> Map.ofList
let model = { createEmpty 100 26 with Cells = cells }

let mutable pass = 0
let mutable fail = 0
let check label formula (expected: CellValue) =
    let got = evaluateFormula formula model None
    if got = expected then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-28s  %-22s  expected %A  got %A" label formula expected got

let n x = CellValue.Number x
let s x = CellValue.String x
let b x = CellValue.Boolean x
let e x = CellValue.Error x
let checkText label got expected =
    if got = expected then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-28s  expected %A  got %A" label expected got

let allocatedBytesFor f =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetAllocatedBytesForCurrentThread()
    let result = f()
    System.GC.GetAllocatedBytesForCurrentThread() - before, result

// arithmetic + precedence
check "add"            "=1+2"          (n 3.0)
check "mul-prec"       "=2+3*4"        (n 14.0)
check "paren"          "=(2+3)*4"      (n 20.0)
check "div"            "=10/4"         (n 2.5)
check "div0"           "=1/0"          (e "DIV/0!")
check "exp"            "=2^10"         (n 1024.0)
check "exp-left-assoc" "=2^3^2"        (n 64.0)
check "unary-neg-exp"  "=-2^2"         (n 4.0)        // Excel: (-2)^2
check "neg"            "=-5"           (n -5.0)
check "percent"        "=50%"          (n 0.5)
check "percent-expr"   "=10+5%"        (n 10.05)
check "concat"         "=\"a\"&\"b\""  (s "ab")
check "concat-num"     "=1&2"          (s "12")
check "concat-prec"    "=1+2&3"        (s "33")       // (1+2)&3

// comparisons
check "eq-true"        "=1=1"          (b true)
check "eq-false"       "=1=2"          (b false)
check "ne"             "=1<>2"         (b true)
check "lt"             "=1<2"          (b true)
check "ge"             "=2>=2"         (b true)
check "cmp-prec"       "=1+1=2"        (b true)       // (1+1)=2
check "str-cmp-ci"     "=\"ABC\"=\"abc\""  (b true)   // case-insensitive

// cell refs
check "ref"            "=A1"           (n 10.0)
check "ref-add"        "=A1+A2"        (n 30.0)
check "abs-ref"        "=$A$1"         (n 10.0)
check "mixed-ref"      "=$A1+A$2"      (n 30.0)
let longBinaryRefCount = 12000
let longBinaryRefModel =
    { createEmpty (longBinaryRefCount + 10) 4 with
        Cells =
            [ for row in 0 .. longBinaryRefCount - 1 ->
                { Row = row; Col = 0 }, { Value = CellValue.Number 1.0; Format = defaultFormat; Comment = None } ]
            |> Map.ofList }
let longBinaryRefFormula =
    "=" + String.concat "+" [ for row in 1 .. longBinaryRefCount -> sprintf "A%d" row ]
let longBinaryRefSw = System.Diagnostics.Stopwatch.StartNew()
let longBinaryRefResult = evaluateFormula longBinaryRefFormula longBinaryRefModel None
longBinaryRefSw.Stop()
checkText "long binary ref formula"
    (longBinaryRefResult = n (float longBinaryRefCount)
     && longBinaryRefSw.Elapsed.TotalMilliseconds < 1000.0)
    true

// coercion
check "text-num-add"   "=\"5\"+3"      (n 8.0)
check "bool-add"       "=TRUE+1"       (n 2.0)
check "text-fail"      "=\"x\"+1"      (e "VALUE!")
check "blank-add"      "=Z99+5"        (n 5.0)        // blank cell → 0

// error propagation + literals
check "err-lit"        "=#REF!"        (e "REF!")
check "err-prop"       "=#N/A+1"       (e "N/A")
check "err-prop-deep"  "=(1/0)*2"      (e "DIV/0!")

// functions
check "sum-range"      "=SUM(A1:A2)"   (n 30.0)
check "sum-mixed"      "=SUM(A1:A4)"   (n 31.0)       // 10+20+TRUE(1), "hi" ignored
check "average"        "=AVERAGE(A1:A2)" (n 15.0)
check "count"          "=COUNT(A1:A4)" (n 2.0)        // only 2 numbers
check "min"            "=MIN(A1:A2)"   (n 10.0)
check "max"            "=MAX(A1:B1)"   (n 10.0)
check "if-true"        "=IF(1>0,\"y\",\"n\")"  (s "y")
check "if-false"       "=IF(A1>100,1,2)"       (n 2.0)
check "if-nested"      "=IF(A1>5,IF(B1>5,1,2),3)" (n 2.0)
check "sum-of-func"    "=SUM(A1:A2)+B1"  (n 33.0)

let csvModel =
    { createEmpty 10 10 with
        Cells =
            [ { Row = 0; Col = 0 }, { Value = CellValue.String "a,b"; Format = defaultFormat; Comment = None }
              { Row = 1; Col = 2 }, { Value = CellValue.Number 5.0; Format = defaultFormat; Comment = None } ]
            |> Map.ofList }
checkText "csv sparse shape" (exportCSV csvModel) "\"a,b\",,\n,,5"
let csvExportA = exportCSV csvModel
let csvExportB = exportCSV csvModel
checkText "csv export cached identity" (obj.ReferenceEquals(csvExportA, csvExportB)) true
let csvModelEdited = setCell { Row = 0; Col = 1 } { Value = CellValue.String "new"; Format = defaultFormat; Comment = None } csvModel
checkText "csv export cache invalidates edit" (exportCSV csvModelEdited) "\"a,b\",new,\n,,5"
checkText "csv export cache keeps first source" (obj.ReferenceEquals(exportCSV csvModel, csvExportA)) true
let csvQuoteModel =
    { createEmpty 2 2 with
        Cells =
            [ { Row = 0; Col = 0 }, { Value = CellValue.String "say \"hi\""; Format = defaultFormat; Comment = None }
              { Row = 0; Col = 1 }, { Value = CellValue.String "plain"; Format = defaultFormat; Comment = None } ]
            |> Map.ofList }
checkText "csv quote escaping" (exportCSV csvQuoteModel) "\"say \"\"hi\"\"\",plain"
let csvFarColumn = 500000
let csvFarColumnModel =
    { createEmpty 1 (csvFarColumn + 1) with
        Cells =
            [ { Row = 0; Col = csvFarColumn },
              { Value = CellValue.Number 1.0; Format = defaultFormat; Comment = None } ]
            |> Map.ofList }
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let csvFarColumnBefore = System.GC.GetAllocatedBytesForCurrentThread()
let csvFarColumnSw = System.Diagnostics.Stopwatch.StartNew()
let csvFarColumnText = exportCSV csvFarColumnModel
csvFarColumnSw.Stop()
let csvFarColumnBytes = System.GC.GetAllocatedBytesForCurrentThread() - csvFarColumnBefore
checkText "csv export far column shape"
    (csvFarColumnText.Length, csvFarColumnText.[csvFarColumn])
    (csvFarColumn + 1, '1')
checkText "csv export far column perf"
    (csvFarColumnSw.Elapsed.TotalMilliseconds < 50.0 && csvFarColumnBytes < 4000000L)
    true
let csvWideTrailingRows = 80
let csvWideTrailingCol = 8000
let csvWideTrailingModel =
    { createEmpty csvWideTrailingRows (csvWideTrailingCol + 1) with
        Cells =
            [ yield { Row = 0; Col = csvWideTrailingCol },
                    { Value = CellValue.Number 9.0; Format = defaultFormat; Comment = None }
              for row in 1 .. csvWideTrailingRows - 1 ->
                  { Row = row; Col = 0 },
                  { Value = CellValue.Number (float row); Format = defaultFormat; Comment = None } ]
            |> Map.ofList }
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let csvWideTrailingBefore = System.GC.GetAllocatedBytesForCurrentThread()
let csvWideTrailingSw = System.Diagnostics.Stopwatch.StartNew()
let csvWideTrailingText = exportCSV csvWideTrailingModel
csvWideTrailingSw.Stop()
let csvWideTrailingBytes = System.GC.GetAllocatedBytesForCurrentThread() - csvWideTrailingBefore
checkText "csv export wide trailing shape"
    (csvWideTrailingText.[csvWideTrailingCol], csvWideTrailingText.[csvWideTrailingCol + 1], csvWideTrailingText.[csvWideTrailingCol + 2])
    ('9', '\n', '1')
checkText "csv export wide trailing perf"
    (csvWideTrailingSw.Elapsed.TotalMilliseconds < 50.0 && csvWideTrailingBytes < 4000000L)
    true
let csvImportBase =
    { createEmpty 4 4 with
        Cells =
            [ { Row = 1; Col = 1 }, { Value = CellValue.String "clear-me"; Format = defaultFormat; Comment = None } ]
            |> Map.ofList }
let csvText = "\"a,b\",\"say \"\"hi\"\"\",=A1+2,true\n42,,false,text"
let csvImported = importCSV csvText 0 0 csvImportBase
checkText "csv import quoted comma" (displayValue (getCellValue { Row = 0; Col = 0 } csvImported)) "a,b"
checkText "csv import quoted quote" (displayValue (getCellValue { Row = 0; Col = 1 } csvImported)) "say \"hi\""
checkText "csv import formula" (match getCellValue { Row = 0; Col = 2 } csvImported with Formula (f, _) -> f | other -> sprintf "%A" other) "=A1+2"
checkText "csv import bool" (match getCellValue { Row = 0; Col = 3 } csvImported with Boolean b -> b | _ -> false) true
checkText "csv import number" (displayValue (getCellValue { Row = 1; Col = 0 } csvImported)) "42"
checkText "csv import false" (match getCellValue { Row = 1; Col = 2 } csvImported with Boolean b -> b | _ -> true) false
checkText "csv import empty clears" (getCellValue { Row = 1; Col = 1 } csvImported) Empty
checkText "csv import same identity" (obj.ReferenceEquals(importCSV csvText 0 0 csvImported, csvImported)) true
let csvImportedWithChanges, csvImportOldCells, csvImportNewCells =
    importCSVWithChanges csvText 0 0 csvImportBase
let csvImportChangeAt addr pairs =
    pairs
    |> Seq.tryFind (fun (candidate, _) -> candidate = addr)
    |> Option.map snd
checkText "csv import changes count" (csvImportOldCells.Count, csvImportNewCells.Count) (8, 8)
checkText "csv import changes model" (getCellValue { Row = 0; Col = 0 } csvImportedWithChanges) (s "a,b")
checkText "csv import changes old clear"
    (csvImportChangeAt { Row = 1; Col = 1 } csvImportOldCells |> Option.bind id |> Option.map (fun c -> c.Value))
    (Some (s "clear-me"))
checkText "csv import changes new clear"
    (csvImportChangeAt { Row = 1; Col = 1 } csvImportNewCells)
    (Some None)
let csvImportSameWithChanges, csvImportSameOldCells, csvImportSameNewCells =
    importCSVWithChanges csvText 0 0 csvImported
checkText "csv import changes same identity" (obj.ReferenceEquals(csvImportSameWithChanges, csvImported)) true
checkText "csv import changes same old" csvImportSameOldCells.Count 0
checkText "csv import changes same new" csvImportSameNewCells.Count 0
let csvOtherText = "left,right\n1,2"
let csvImportedOther = importCSV csvOtherText 0 0 csvImportBase
checkText "csv import other text" (displayValue (getCellValue { Row = 0; Col = 0 } csvImportedOther)) "left"
let csvImportedAgainFromBase = importCSV csvText 0 0 csvImportBase
checkText "csv import source cache cells" (obj.ReferenceEquals(csvImportedAgainFromBase.Cells, csvImported.Cells)) true
let csvImportBaseMoved = { csvImportBase with Selection = SingleCell (3, 3) }
let csvImportedMoved = importCSV csvText 0 0 csvImportBaseMoved
checkText "csv import source cache preserves selection" csvImportedMoved.Selection (SingleCell (3, 3))
let csvEmptyOnlyImportBase = createEmpty 20 20
checkText "csv import empty source blank identity"
    (obj.ReferenceEquals(importCSV ",,,\n,,," 0 0 csvEmptyOnlyImportBase, csvEmptyOnlyImportBase))
    true
let csvSparseEmptySourceText (rows: int) (cols: int) =
    let sb = System.Text.StringBuilder(rows * cols)
    for row in 0 .. rows - 1 do
        if row > 0 then
            sb.Append('\n') |> ignore
        sb.Append(row) |> ignore
        for _col in 1 .. cols - 1 do
            sb.Append(',') |> ignore
    sb.ToString()
checkText "csv parse quoted"
    (parseCSV "\"a,b\",\"say \"\"hi\"\"\"\n42,")
    [ [ "a,b"; "say \"hi\"" ]; [ "42"; "" ] ]
let csvParsePerfInput = csvSparseEmptySourceText 20000 8
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let csvParsePerfBefore = System.GC.GetAllocatedBytesForCurrentThread()
let csvParsePerfSw = System.Diagnostics.Stopwatch.StartNew()
let csvParsePerfRows = parseCSV csvParsePerfInput
csvParsePerfSw.Stop()
let csvParsePerfBytes = System.GC.GetAllocatedBytesForCurrentThread() - csvParsePerfBefore
checkText "csv parse large shape"
    (csvParsePerfRows.Length, csvParsePerfRows |> List.sumBy List.length)
    (20000, 160000)
checkText "csv parse large perf"
    (csvParsePerfSw.Elapsed.TotalMilliseconds < 80.0
     && csvParsePerfBytes < 16L * 1024L * 1024L)
    true
let csvSparseEmptySourceRows = 6000
let csvSparseEmptySourceCols = 40
let csvSparseEmptySourceInput = csvSparseEmptySourceText csvSparseEmptySourceRows csvSparseEmptySourceCols
let csvSparseEmptySourceBase = createEmpty (csvSparseEmptySourceRows + 1) (csvSparseEmptySourceCols + 1)
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let csvSparseEmptySourceBefore = System.GC.GetAllocatedBytesForCurrentThread()
let csvSparseEmptySourceSw = System.Diagnostics.Stopwatch.StartNew()
let csvSparseEmptySourceImported = importCSV csvSparseEmptySourceInput 0 0 csvSparseEmptySourceBase
csvSparseEmptySourceSw.Stop()
let csvSparseEmptySourceBytes = System.GC.GetAllocatedBytesForCurrentThread() - csvSparseEmptySourceBefore
checkText "csv import empty source sparse shape"
    (csvSparseEmptySourceImported.Cells.Count,
     getCellValue { Row = csvSparseEmptySourceRows - 1; Col = 0 } csvSparseEmptySourceImported)
    (csvSparseEmptySourceRows, CellValue.Number (float (csvSparseEmptySourceRows - 1)))
checkText "csv import empty source sparse perf"
    (csvSparseEmptySourceSw.Elapsed.TotalMilliseconds < 80.0)
    true
checkText "csv import empty source sparse allocation"
    (csvSparseEmptySourceBytes < 8000000L)
    true

let csvCacheSourceCount = 24
let csvCacheRows = 500
let csvCacheCols = 8
let csvCacheText sourceIndex =
    String.concat "\n" [
        for row in 0 .. csvCacheRows - 1 ->
            String.concat "," [
                for col in 0 .. csvCacheCols - 1 ->
                    string (sourceIndex * 1000000 + row * csvCacheCols + col)
            ]
    ]
let csvCacheExportModels =
    [| for sourceIndex in 0 .. csvCacheSourceCount - 1 ->
        { createEmpty (csvCacheRows + 5) (csvCacheCols + 5) with
            Cells =
                [ for row in 0 .. csvCacheRows - 1 do
                    for col in 0 .. csvCacheCols - 1 ->
                        { Row = row; Col = col },
                        { Value = CellValue.Number (float (sourceIndex * 1000000 + row * csvCacheCols + col)); Format = defaultFormat; Comment = None } ]
                |> Map.ofList } |]
let csvCacheImportTexts = [| for sourceIndex in 0 .. csvCacheSourceCount - 1 -> csvCacheText sourceIndex |]
let csvCacheImportBase = createEmpty (csvCacheRows + 5) (csvCacheCols + 5)
let csvCacheExportRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in csvCacheExportModels do
            total <- total + (exportCSV model).Length
    total
let csvCacheImportRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for text in csvCacheImportTexts do
            total <- total + (importCSV text 0 0 csvCacheImportBase).Cells.Count
    total
csvCacheExportRun 6 |> ignore
csvCacheImportRun 6 |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let csvCacheSw = System.Diagnostics.Stopwatch.StartNew()
let csvCacheExportTotal = csvCacheExportRun 30
let csvCacheImportTotal = csvCacheImportRun 30
csvCacheSw.Stop()
checkText "csv cache many sources perf"
    (csvCacheSw.Elapsed.TotalMilliseconds < 30.0 && csvCacheExportTotal > 0 && csvCacheImportTotal > 0)
    true
let csvTrimOrderCell value =
    { Value = CellValue.String value; Format = defaultFormat; Comment = None }
let csvTrimOrderTextModel label payloadLength =
    let prefix = sprintf "%s-" label
    let text = prefix + System.String(char (65 + (payloadLength % 26)), payloadLength)
    { createEmpty 2 1 with
        Cells = [ { Row = 0; Col = 0 }, csvTrimOrderCell text ] |> Map.ofList }
let csvExportTrimOrderBytes, csvExportTrimOrderTotal =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable total = 0
    for index in 0 .. 7 do
        total <- total + (exportCSV (csvTrimOrderTextModel (sprintf "old-%02d" index) 2000000)).Length
    total <- total + (exportCSV (csvTrimOrderTextModel "current-large" 6000000)).Length
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before, total
checkText "csv export cache trim current memory"
    (csvExportTrimOrderBytes < 28000000L
     && csvExportTrimOrderTotal > 0)
    true
let csvTrimOrderImportSourceModel label payloadLength =
    let prefix = sprintf "%s-" label
    let text = prefix + System.String(char (70 + (payloadLength % 20)), payloadLength)
    { createEmpty 3 1 with
        Cells = [ { Row = 1; Col = 0 }, csvTrimOrderCell text ] |> Map.ofList }
let csvImportTrimOrderBytes, csvImportTrimOrderTotal =
    let currentImportCellCount () =
        let currentCsv = System.String('Z', 6000000)
        let imported = importCSV currentCsv 0 0 (createEmpty 2 1)
        imported.Cells.Count
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable total = 0
    for index in 0 .. 7 do
        let imported = importCSV "1" 0 0 (csvTrimOrderImportSourceModel (sprintf "old-%02d" index) 2000000)
        total <- total + imported.Cells.Count
    total <- total + currentImportCellCount ()
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before, total
checkText "csv import cache trim current memory"
    (csvImportTrimOrderBytes < 28000000L
     && csvImportTrimOrderTotal = 17)
    true
let csvExportRetainedModelCount = 80
let csvExportRetainedRowCount = 20000
let csvExportRetainedModel modelIndex =
    { createEmpty csvExportRetainedRowCount 1 with
        Cells =
            [ for row in 0 .. csvExportRetainedRowCount - 1 ->
                { Row = row; Col = 0 }, { Value = CellValue.String ""; Format = defaultFormat; Comment = None } ]
            |> Map.ofList
        SearchQuery = Some (sprintf "csv-export-retained-%02d" modelIndex) }
let csvExportRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. csvExportRetainedModelCount - 1 do
        exportCSV (csvExportRetainedModel modelIndex) |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
checkText "csv export cache bounded source memory"
    (csvExportRetainedBytes < 140000000L)
    true
let csvExportSourceTextModelCount = 24
let csvExportSourceTextPayload modelIndex =
    String.replicate 2000000 (string (char (65 + (modelIndex % 26))))
let csvExportSourceTextModel modelIndex =
    { createEmpty 2 1 with
        Cells =
            [ { Row = 0; Col = 0 },
              { Value = CellValue.Number (float modelIndex)
                Format = defaultFormat
                Comment = Some { Text = csvExportSourceTextPayload modelIndex; Author = Some "perf" } } ]
            |> Map.ofList }
let csvExportSourceTextBytes, csvExportSourceTextTotal =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable total = 0
    for modelIndex in 0 .. csvExportSourceTextModelCount - 1 do
        total <- total + (exportCSV (csvExportSourceTextModel modelIndex)).Length
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before, total
checkText "csv export cache source text memory"
    (csvExportSourceTextBytes < 30000000L
     && csvExportSourceTextTotal > 0)
    true
let csvImportSourceTextModelCount = 24
let csvImportSourceTextPayload modelIndex =
    String.replicate 2000000 (string (char (65 + (modelIndex % 26))))
let csvImportSourceTextModel modelIndex =
    { createEmpty 3 1 with
        Cells =
            [ { Row = 1; Col = 0 }, { Value = CellValue.String (csvImportSourceTextPayload modelIndex); Format = defaultFormat; Comment = None } ]
            |> Map.ofList }
let csvImportSourceTextBytes, csvImportSourceTextTotal =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable total = 0
    for modelIndex in 0 .. csvImportSourceTextModelCount - 1 do
        let imported = importCSV "1" 0 0 (csvImportSourceTextModel modelIndex)
        total <- total + imported.Cells.Count
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before, total
checkText "csv import cache source text memory"
    (csvImportSourceTextBytes < 30000000L
     && csvImportSourceTextTotal = csvImportSourceTextModelCount * 2)
    true
let csvExportOversizedSourceCellsWeakReference () =
    let payload = System.String('E', 6 * 1024 * 1024)
    let cells =
        [ { Row = 0; Col = 0 },
          { Value = CellValue.Number 1.0
            Format = defaultFormat
            Comment = Some { Text = payload; Author = Some "perf" } } ]
        |> Map.ofList
    let model = { createEmpty 2 1 with Cells = cells }
    exportCSV model |> ignore
    System.WeakReference(cells)
let csvExportOversizedSourceCellsWeak = csvExportOversizedSourceCellsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
checkText "csv export cache skips oversized source"
    csvExportOversizedSourceCellsWeak.IsAlive
    false
let csvExportOversizedTextSkipsSourceElapsed =
    let sourceCellCount = 500000
    let farCol = 6000000
    let sourceCell =
        { Value = CellValue.Number 1.0
          Format =
            { defaultFormat with
                FontName = Some "f"
                TextColor = Some "t"
                BgColor = Some "b"
                HorizontalAlign = Some "h"
                VerticalAlign = Some "v" }
          Comment = None }
    let cells =
        [ for col in 0 .. sourceCellCount - 2 do
            yield { Row = 0; Col = col }, sourceCell
          yield { Row = 0; Col = farCol }, sourceCell ]
        |> Map.ofList
    let model = { createEmpty 1 (farCol + 1) with Cells = cells }
    usedBounds model |> ignore
    let sw = System.Diagnostics.Stopwatch.StartNew()
    exportCSV model |> ignore
    sw.Stop()
    sw.Elapsed.TotalMilliseconds
checkText "csv export oversized text skips source"
    (csvExportOversizedTextSkipsSourceElapsed < 1500.0)
    true
let csvImportOversizedSourceCellsWeakReference () =
    let payload = System.String('I', 6 * 1024 * 1024)
    let cells =
        [ { Row = 1; Col = 0 }, { Value = CellValue.String payload; Format = defaultFormat; Comment = None } ]
        |> Map.ofList
    let model = { createEmpty 3 1 with Cells = cells }
    importCSV "1" 0 0 model |> ignore
    System.WeakReference(cells)
let csvImportOversizedSourceCellsWeak = csvImportOversizedSourceCellsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
checkText "csv import cache skips oversized source"
    csvImportOversizedSourceCellsWeak.IsAlive
    false
let csvImportLargeNoopBlankText = System.String(',', 300000)
let csvImportLargeNoopBlankBytes, csvImportLargeNoopBlankModel =
    allocatedBytesFor (fun () -> importCSV csvImportLargeNoopBlankText 0 0 (createEmpty 1 1))
checkText "csv import large noop blank identity" (csvImportLargeNoopBlankModel.Cells.Count) 0
checkText "csv import large noop blank allocation" (csvImportLargeNoopBlankBytes < 100000L) true
let csvImportLargeNoopBlankWeakReference () =
    let csvText = System.String(',', 300000)
    importCSV csvText 0 0 (createEmpty 1 1) |> ignore
    System.WeakReference(csvText)
let csvImportLargeNoopBlankWeak = csvImportLargeNoopBlankWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
checkText "csv import skips large noop blank cache"
    csvImportLargeNoopBlankWeak.IsAlive
    false
let csvImportOversizedSourceCellCountWeakReference () =
    let sourceCellCount = 250001
    let cells =
        [ for col in 0 .. sourceCellCount - 1 ->
            { Row = 0; Col = col }, { Value = CellValue.Number 1.0; Format = defaultFormat; Comment = None } ]
        |> Map.ofList
    let model = { createEmpty 1 sourceCellCount with Cells = cells }
    let csvText = System.String(',', sourceCellCount - 1)
    importCSV csvText 0 0 model |> ignore
    System.WeakReference(cells)
let csvImportOversizedSourceCellCountWeak = csvImportOversizedSourceCellCountWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
checkText "csv import cache skips oversized source cells"
    csvImportOversizedSourceCellCountWeak.IsAlive
    false
let csvImportOversizedTextWeakReference () =
    let rec formulaChain depth value =
        if depth = 0 then value
        else formulaChain (depth - 1) (CellValue.Formula ("", value))
    let sourceRows = 80000
    let sourceCell = { Value = formulaChain 128 (CellValue.String "x"); Format = defaultFormat; Comment = None }
    let cells =
        [ for row in 0 .. sourceRows - 1 ->
            { Row = row; Col = 0 }, sourceCell ]
        |> Map.ofList
    let model = { createEmpty sourceRows 1 with Cells = cells }
    let csvText = System.String('Z', 6 * 1024 * 1024)
    let sw = System.Diagnostics.Stopwatch.StartNew()
    importCSV csvText 0 0 model |> ignore
    sw.Stop()
    checkText "csv import oversized text skips source" (sw.Elapsed.TotalMilliseconds < 1500.0) true
    System.WeakReference(csvText)
let csvImportOversizedTextWeak = csvImportOversizedTextWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
checkText "csv import cache skips oversized text"
    csvImportOversizedTextWeak.IsAlive
    false

printfn ""
printfn "Stage 1 formula tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
