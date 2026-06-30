#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// AutoFilter model: unique values for the checklist, apply, and visible rows.

open System.Diagnostics
open RedBat.Spreadsheet.Model

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let S = CellValue.String
// Col 0 (rows 0-4): East, West, East, North, West ; Col 1 = sales
let model =
    { createEmpty 20 8 with
        Cells =
            ([ "East";"West";"East";"North";"West" ] |> List.mapi (fun i v -> { Row=i; Col=0 }, cell (S v)))
            @ ([ 10.;20.;30.;40.;50. ] |> List.mapi (fun i v -> { Row=i; Col=1 }, cell (CellValue.Number v)))
            |> Map.ofList }

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-26s expected %A got %A" label expected got

check "unique values"   (columnUniqueValues 0 model) [ "East"; "North"; "West" ]
check "raw unique values" (getColumnUniqueValues 0 model) [ "East"; "North"; "West" ]
check "unique col1 count" (List.length (columnUniqueValues 1 model)) 5
let modelWithBlank =
    model |> setCell { Row = 5; Col = 0 } (cell CellValue.Empty)
check "raw unique keeps blank" (getColumnUniqueValues 0 modelWithBlank |> List.contains "") true
check "unique skips blank" (columnUniqueValues 0 modelWithBlank) [ "East"; "North"; "West" ]
let rawUniqueCached = getColumnUniqueValues 0 modelWithBlank
let noBlankUniqueCached = columnUniqueValues 0 modelWithBlank
check "raw unique cached repeat" (obj.ReferenceEquals(getColumnUniqueValues 0 modelWithBlank, rawUniqueCached)) true
check "unique no blank cached repeat" (obj.ReferenceEquals(columnUniqueValues 0 modelWithBlank, noBlankUniqueCached)) true
let otherColumnUniqueCached = columnUniqueValues 1 modelWithBlank
check "unique other col count" otherColumnUniqueCached.Length 5
check "raw unique survives other col" (obj.ReferenceEquals(getColumnUniqueValues 0 modelWithBlank, rawUniqueCached)) true
check "set same cell identity" (obj.ReferenceEquals(setCell { Row = 5; Col = 0 } (cell CellValue.Empty) modelWithBlank, modelWithBlank)) true
check "clear missing identity" (obj.ReferenceEquals(clearCell { Row = 19; Col = 7 } modelWithBlank, modelWithBlank)) true
let formattedOnce = setCellFormat { Row = 0; Col = 0 } (fun format -> { format with TextColor = Some "#C00000" }) modelWithBlank
check "format existing changes" ((getCell { Row = 0; Col = 0 } formattedOnce).Value.Format.TextColor) (Some "#C00000")
check "format same identity" (obj.ReferenceEquals(setCellFormat { Row = 0; Col = 0 } (fun format -> { format with TextColor = Some "#C00000" }) formattedOnce, formattedOnce)) true
check "format missing default identity" (obj.ReferenceEquals(setCellFormat { Row = 19; Col = 7 } id modelWithBlank, modelWithBlank)) true
let formattedMissing = setCellFormat { Row = 19; Col = 7 } (fun format -> { format with BgColor = Some "#FFFF00" }) modelWithBlank
check "format missing creates" ((getCell { Row = 19; Col = 7 } formattedMissing).Value.Format.BgColor) (Some "#FFFF00")
check "text color same identity" (obj.ReferenceEquals(setCellTextColor { Row = 0; Col = 0 } "#C00000" formattedOnce, formattedOnce)) true
let bgOnce = setCellBgColor { Row = 0; Col = 0 } "#FFFF00" modelWithBlank
check "bg color changes" ((getCell { Row = 0; Col = 0 } bgOnce).Value.Format.BgColor) (Some "#FFFF00")
check "bg color same identity" (obj.ReferenceEquals(setCellBgColor { Row = 0; Col = 0 } "#FFFF00" bgOnce, bgOnce)) true
let numFmtOnce = setCellNumberFormat { Row = 0; Col = 1 } "currency" modelWithBlank
check "number format changes" ((getCell { Row = 0; Col = 1 } numFmtOnce).Value.Format.NumberFormat) (Some "currency")
check "number format same identity" (obj.ReferenceEquals(setCellNumberFormat { Row = 0; Col = 1 } "currency" numFmtOnce, numFmtOnce)) true
let alignOnce = setCellHorizontalAlign { Row = 0; Col = 1 } "center" modelWithBlank
check "align changes" ((getCell { Row = 0; Col = 1 } alignOnce).Value.Format.HorizontalAlign) (Some "center")
check "align same identity" (obj.ReferenceEquals(setCellHorizontalAlign { Row = 0; Col = 1 } "center" alignOnce, alignOnce)) true
let alignMissing = setCellHorizontalAlign { Row = 19; Col = 6 } "right" modelWithBlank
check "align missing creates" ((getCell { Row = 19; Col = 6 } alignMissing).Value.Format.HorizontalAlign) (Some "right")
let borderOnce = setCellBorder { Row = 0; Col = 1 } "all" "thin" "#000000" modelWithBlank
check "border changes" ((getCell { Row = 0; Col = 1 } borderOnce).Value.Format.BorderTop) (Some { Style = "thin"; Color = Some "#000000" })
check "border same identity" (obj.ReferenceEquals(setCellBorder { Row = 0; Col = 1 } "all" "thin" "#000000" borderOnce, borderOnce)) true
check "border invalid side identity" (obj.ReferenceEquals(setCellBorder { Row = 0; Col = 1 } "diagonal" "thin" "#000000" borderOnce, borderOnce)) true
let modelWithSouth =
    modelWithBlank |> setCell { Row = 6; Col = 0 } (cell (S "South"))
check "unique invalidates edit" (columnUniqueValues 0 modelWithSouth) [ "East"; "North"; "South"; "West" ]
let errorUniqueModel =
    { createEmpty 10 2 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Error "N/A")
              { Row = 1; Col = 0 }, cell (CellValue.Error "VALUE!")
              { Row = 2; Col = 0 }, cell (CellValue.Error "N/A") ]
            |> Map.ofList }
check "unique error values" (columnUniqueValues 0 errorUniqueModel) [ "#N/A"; "#VALUE!" ]
let errorUniqueTable =
    { Name = "Errors"
      StartRow = 1
      StartCol = 1
      EndRow = 3
      EndCol = 1
      HeaderRowCount = 0
      Style = None
      AutoFilter = None
      HasTotals = false }
check "table unique error values" (tableColumnUniqueValues errorUniqueTable 0 errorUniqueModel) [ "#N/A"; "#VALUE!" ]
check "unique cache keeps first source" (obj.ReferenceEquals(columnUniqueValues 0 modelWithBlank, noBlankUniqueCached)) true
let oversizedColumnUniqueCacheSkipped () =
    let model =
        { createEmpty 100001 1 with
            Cells =
                [ for row in 0 .. 100000 ->
                    { Row = row; Col = 0 }, cell (S (sprintf "u%06d" row)) ]
                |> Map.ofList }
    let valuesA = columnUniqueValues 0 model
    let valuesB = columnUniqueValues 0 model
    valuesA.Length, obj.ReferenceEquals(valuesA, valuesB)
let oversizedUniqueCount, oversizedUniqueReused = oversizedColumnUniqueCacheSkipped ()
check "unique cache skips oversized column"
    (oversizedUniqueCount = 100001 && not oversizedUniqueReused)
    true
let oversizedTableUniqueCacheSkipped () =
    let table =
        { Name = "OversizedUnique"
          StartRow = 1
          StartCol = 1
          EndRow = 200001
          EndCol = 1
          HeaderRowCount = 0
          Style = None
          AutoFilter = Some (1, 1)
          HasTotals = false }
    let model =
        { createEmpty 200001 1 with
            Tables = [ table ]
            Cells =
                [ for row in 0 .. 200000 ->
                    { Row = row; Col = 0 }, cell (S (sprintf "tu%06d" row)) ]
                |> Map.ofList }
    let valuesA = tableColumnUniqueValues table 0 model
    let valuesB = tableColumnUniqueValues table 0 model
    valuesA.Length, obj.ReferenceEquals(valuesA, valuesB)
let oversizedTableUniqueCount, oversizedTableUniqueReused = oversizedTableUniqueCacheSkipped ()
check "unique cache skips oversized table column"
    (oversizedTableUniqueCount > 200000 && not oversizedTableUniqueReused)
    true
let wideUniqueModel =
    { createEmpty 60 20 with
        Cells =
            [ for row in 0 .. 49 do
                for col in 0 .. 15 ->
                    { Row = row; Col = col }, cell (S (sprintf "c%02d-%02d" col (row % 10))) ]
            |> Map.ofList }
let wideUniqueFirst = columnUniqueValues 0 wideUniqueModel
for col in 1 .. 15 do
    columnUniqueValues col wideUniqueModel |> ignore
check "unique cache keeps many columns" (obj.ReferenceEquals(columnUniqueValues 0 wideUniqueModel, wideUniqueFirst)) true
let uniqueManyModelCount = 24
let uniqueManyRowCount = 8000
let uniqueManyModels =
    [| for modelIndex in 0 .. uniqueManyModelCount - 1 ->
        { createEmpty uniqueManyRowCount 2 with
            Cells =
                [ for row in 0 .. uniqueManyRowCount - 1 ->
                    { Row = row; Col = 0 }, cell (S (sprintf "unique-%02d" (row % 32))) ]
                |> Map.ofList } |]
let uniqueManyRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in uniqueManyModels do
            total <- total + (columnUniqueValues 0 model).Length
    total
uniqueManyRun 3 |> ignore
let uniqueManySw = Stopwatch.StartNew()
let uniqueManyTotal = uniqueManyRun 60
uniqueManySw.Stop()
check "unique cache many models"
    (uniqueManySw.Elapsed.TotalMilliseconds < 5.0 && uniqueManyTotal > 0)
    true
let uniqueCacheRetainedRowCount = 20000
let uniqueCacheRetainedModelCount = 80
let uniqueCacheRetainedTable rows : Table =
    { Name = "UniqueRetained"
      StartRow = 1
      StartCol = 1
      EndRow = rows
      EndCol = 1
      HeaderRowCount = 1
      Style = None
      AutoFilter = Some (1, 1)
      HasTotals = false }
let uniqueCacheRetainedModel modelIndex =
    let table = uniqueCacheRetainedTable uniqueCacheRetainedRowCount
    { createEmpty uniqueCacheRetainedRowCount 2 with
        Tables = [ table ]
        Cells =
            [ for row in 0 .. uniqueCacheRetainedRowCount - 1 ->
                { Row = row; Col = 0 }, cell (S (sprintf "unique-retained-%02d" ((row + modelIndex) % 8))) ]
            |> Map.ofList }
let uniqueCacheRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. uniqueCacheRetainedModelCount - 1 do
        let model = uniqueCacheRetainedModel modelIndex
        columnUniqueValues 0 model |> ignore
        tableColumnUniqueValues (List.head model.Tables) 0 model |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "unique caches bounded source memory" (uniqueCacheRetainedBytes < 160000000L) true
let uniqueCacheSourceTextTable =
    { Name = "UniqueSourceText"
      StartRow = 1
      StartCol = 1
      EndRow = 2
      EndCol = 1
      HeaderRowCount = 1
      Style = None
      AutoFilter = Some (1, 1)
      HasTotals = false }
let uniqueCacheSourceTextModel modelIndex =
    { createEmpty 4 1 with
        Tables = [ uniqueCacheSourceTextTable ]
        Cells =
            [ { Row = 0; Col = 0 }, cell (S "Header")
              { Row = 1; Col = 0 }, cell (S (System.String(char (65 + (modelIndex % 26)), 2000000))) ]
            |> Map.ofList }
let uniqueCacheSourceTextBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. 39 do
        let model = uniqueCacheSourceTextModel modelIndex
        columnUniqueValues 0 model |> ignore
        tableColumnUniqueValues (List.head model.Tables) 0 model |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "unique caches source text memory" (uniqueCacheSourceTextBytes < 30000000L) true
let uniqueCachesOversizedSourceWeakReference () =
    let table = uniqueCacheRetainedTable 500001
    let cells =
        [ for row in 0 .. 500000 ->
            { Row = row; Col = 0 }, cell (S "same") ]
        |> Map.ofList
    let model =
        { createEmpty 500001 1 with
            Tables = [ table ]
            Cells = cells }
    columnUniqueValues 0 model |> ignore
    tableColumnUniqueValues table 0 model |> ignore
    System.WeakReference(cells)
let uniqueCachesOversizedSourceWeak = uniqueCachesOversizedSourceWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "unique caches skip oversized source"
    uniqueCachesOversizedSourceWeak.IsAlive
    false
let uniqueCachesOversizedTextWeakReference () =
    let table = uniqueCacheRetainedTable 2
    let payload = System.String('U', 6 * 1024 * 1024)
    let cells =
        [ { Row = 0; Col = 0 }, cell (S payload)
          { Row = 1; Col = 0 }, cell (S "same") ]
        |> Map.ofList
    let model =
        { createEmpty 2 1 with
            Tables = [ table ]
            Cells = cells }
    columnUniqueValues 0 model |> ignore
    tableColumnUniqueValues table 0 model |> ignore
    System.WeakReference(cells)
let uniqueCachesOversizedTextWeak = uniqueCachesOversizedTextWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "unique caches skip oversized text"
    uniqueCachesOversizedTextWeak.IsAlive
    false
let uniqueCachesGeneratedTextWeakReference () =
    let payload = System.String('E', 4 * 1024 * 1024)
    let cells =
        [ { Row = 1; Col = 0 }, cell (CellValue.Error payload) ]
        |> Map.ofList
    let model =
        { createEmpty 3 1 with
            Cells = cells }
    let tableWithHeaderRows headerRows =
        { uniqueCacheRetainedTable 2 with
            Name = sprintf "GeneratedUniqueText%d" headerRows
            StartRow = 2 - headerRows
            EndRow = 2
            HeaderRowCount = headerRows }
    let firstValues = tableColumnUniqueValues (tableWithHeaderRows 1) 0 model
    tableColumnUniqueValues (tableWithHeaderRows 2) 0 model |> ignore
    System.WeakReference(firstValues)
let uniqueCachesGeneratedTextWeak = uniqueCachesGeneratedTextWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "table unique cache skips generated text"
    uniqueCachesGeneratedTextWeak.IsAlive
    false
let uniqueCachesOversizedResultSkipSourceElapsed =
    let table = uniqueCacheRetainedTable 100001
    let format =
        { defaultFormat with
            FontName = Some "UniquePerfFont"
            NumberFormat = Some (System.String('0', 256))
            BgColor = Some "#123456"
            TextColor = Some "#654321"
            HorizontalAlign = Some "left"
            VerticalAlign = Some "middle" }
    let cells =
        [ for row in 0 .. 100000 ->
            { Row = row; Col = 0 },
            { Value = S (sprintf "oversized-result-%06d" row)
              Format = format
              Comment = Some { Text = "unique result cache source text"; Author = Some "perf" } } ]
        |> Map.ofList
    let model =
        { createEmpty 100001 1 with
            Tables = [ table ]
            Cells = cells }
    let sw = Stopwatch.StartNew()
    columnUniqueValues 0 model |> ignore
    tableColumnUniqueValues table 0 model |> ignore
    sw.Stop()
    sw.Elapsed.TotalMilliseconds
if uniqueCachesOversizedResultSkipSourceElapsed < 1500.0 then
    pass <- pass + 1
else
    fail <- fail + 1
    printfn
        "FAIL %-26s expected under %.1f ms got %.1f ms"
        "unique oversized skip source"
        1500.0
        uniqueCachesOversizedResultSkipSourceElapsed

// no filter → all rows visible
check "no filter all visible" (getVisibleRows model |> Set.count) 5

// filter col0 to {East} → rows 0 and 2
let fEast = applyFilter 0 (Set.ofList [ "East" ]) model
let fEastRows = getVisibleRows fEast
check "filter East rows"  fEastRows (Set.ofList [ 0; 2 ])
check "same filter identity" (obj.ReferenceEquals(applyFilter 0 (Set.ofList [ "East" ]) fEast, fEast)) true

// filter col0 to {West, North} → rows 1,3,4
let fWN = applyFilter 0 (Set.ofList [ "West"; "North" ]) model
check "filter West+North"  (getVisibleRows fWN) (Set.ofList [ 1; 3; 4 ])
check "filter visible cache keeps first" (obj.ReferenceEquals(getVisibleRows fEast, fEastRows)) true

// AND across two columns: East AND sales in {30}
let fAnd = model |> applyFilter 0 (Set.ofList [ "East" ]) |> applyFilter 1 (Set.ofList [ "30" ])
check "filter AND two cols" (getVisibleRows fAnd) (Set.ofList [ 2 ])
let fEastWithChangedCells =
    { fEast with Cells = Map.add { Row = 1; Col = 0 } (cell (S "East")) fEast.Cells }
check "filter visible cache invalidates cells" (getVisibleRows fEastWithChangedCells) (Set.ofList [ 0; 1; 2 ])
check "filter visible cache keeps first source" (obj.ReferenceEquals(getVisibleRows fEast, fEastRows)) true
let fNorth = applyFilter 0 (Set.ofList [ "North" ]) model
check "filter visible cache invalidates filter" (getVisibleRows fNorth) (Set.ofList [ 3 ])
let oversizedVisibleRowsCacheSkipped () =
    let model =
        { createEmpty 200001 1 with
            Cells =
                [ for row in 0 .. 200000 ->
                    { Row = row; Col = 0 }, cell (S "keep") ]
                |> Map.ofList }
        |> applyFilter 0 (Set.singleton "keep")
    let rowsA = getVisibleRows model
    let rowsB = getVisibleRows model
    rowsA.Count, obj.ReferenceEquals(rowsA, rowsB)
let oversizedVisibleRowsCount, oversizedVisibleRowsReused = oversizedVisibleRowsCacheSkipped ()
check "filter visible cache skips oversized rows"
    (oversizedVisibleRowsCount > 200000 && not oversizedVisibleRowsReused)
    true
let oversizedFilterProjectionCacheSkipped () =
    let values =
        [ for index in 0 .. 100000 ->
            sprintf "oversized-filter-%06d" index ]
        |> Set.ofList
    let model =
        { createEmpty 2 1 with
            Cells =
                [ { Row = 0; Col = 0 }, cell (S "header")
                  { Row = 1; Col = 0 }, cell (S "hidden") ]
                |> Map.ofList }
        |> applyFilter 0 values
    let projectedA = applyFilterRowHeights model
    let projectedB = applyFilterRowHeights model
    projectedA.FilterHiddenRows.Length, obj.ReferenceEquals(projectedA.FilterHiddenRows, projectedB.FilterHiddenRows)
let oversizedProjectionHiddenCount, oversizedProjectionReused = oversizedFilterProjectionCacheSkipped ()
check "filter projection cache skips oversized key"
    (oversizedProjectionHiddenCount = 1 && not oversizedProjectionReused)
    true
let filterRowCachesOversizedKeySkipSourceElapsed =
    let values =
        [ for index in 0 .. 100000 ->
            sprintf "oversized-key-%06d" index ]
        |> Set.ofList
    let format =
        { defaultFormat with
            FontName = Some "FilterPerfFont"
            NumberFormat = Some (System.String('0', 256))
            BgColor = Some "#223344"
            TextColor = Some "#ddeeff" }
    let cells =
        [ for row in 0 .. 79999 ->
            { Row = row; Col = 0 },
            { Value = S "drop"
              Format = format
              Comment = Some { Text = "filter row cache source text"; Author = Some "perf" } } ]
        |> Map.ofList
    let model =
        { createEmpty 80000 1 with Cells = cells }
        |> applyFilter 0 values
    let sw = Stopwatch.StartNew()
    applyFilterRowHeights model |> ignore
    getVisibleRows model |> ignore
    sw.Stop()
    sw.Elapsed.TotalMilliseconds
if filterRowCachesOversizedKeySkipSourceElapsed < 1500.0 then
    pass <- pass + 1
else
    fail <- fail + 1
    printfn
        "FAIL %-26s expected under %.1f ms got %.1f ms"
        "filter oversized skips source"
        1500.0
        filterRowCachesOversizedKeySkipSourceElapsed
let filterRowCachesOversizedFilterTextSkipSourceElapsed =
    let values = Set.singleton (System.String('K', 6 * 1024 * 1024))
    let cells =
        [ for row in 0 .. 199999 ->
            { Row = row; Col = 0 }, cell (S "drop") ]
        |> Map.ofList
    let model =
        { createEmpty 200000 1 with Cells = cells }
        |> applyFilter 0 values
    let sw = Stopwatch.StartNew()
    applyFilterRowHeights model |> ignore
    getVisibleRows model |> ignore
    sw.Stop()
    sw.Elapsed.TotalMilliseconds
if filterRowCachesOversizedFilterTextSkipSourceElapsed < 1000.0 then
    pass <- pass + 1
else
    fail <- fail + 1
    printfn
        "FAIL %-26s expected under %.1f ms got %.1f ms"
        "filter text skips source"
        1000.0
        filterRowCachesOversizedFilterTextSkipSourceElapsed
let filterRowCachesOversizedSourceWeakReference () =
    let cells =
        [ for row in 0 .. 500000 ->
            let value = if row = 0 then "keep" else "drop"
            { Row = row; Col = 0 }, cell (S value) ]
        |> Map.ofList
    let model =
        { createEmpty 500001 1 with Cells = cells }
        |> applyFilter 0 (Set.singleton "keep")
    getVisibleRows model |> ignore
    applyFilterRowHeights model |> ignore
    System.WeakReference(cells)
let filterRowCachesOversizedSourceWeak = filterRowCachesOversizedSourceWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "filter row caches skip oversized source"
    filterRowCachesOversizedSourceWeak.IsAlive
    false
let filterRowCachesOversizedTextWeakReference () =
    let payload = System.String('F', 6 * 1024 * 1024)
    let cells =
        [ { Row = 0; Col = 0 }, cell (S payload)
          { Row = 1; Col = 0 }, cell (S "keep") ]
        |> Map.ofList
    let model =
        { createEmpty 2 1 with Cells = cells }
        |> applyFilter 0 (Set.singleton "keep")
    getVisibleRows model |> ignore
    applyFilterRowHeights model |> ignore
    System.WeakReference(cells)
let filterRowCachesOversizedTextWeak = filterRowCachesOversizedTextWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "filter row caches skip oversized text"
    filterRowCachesOversizedTextWeak.IsAlive
    false
let manyFilterCacheModel =
    { createEmpty 80 2 with
        Cells =
            [ for row in 0 .. 63 ->
                { Row = row; Col = 0 }, cell (S (sprintf "group-%02d" (row % 16))) ]
            |> Map.ofList }
let manyFilterModels =
    [| for i in 0 .. 15 ->
        manyFilterCacheModel |> applyFilter 0 (Set.singleton (sprintf "group-%02d" i)) |]
let manyVisibleRowsFirst = getVisibleRows manyFilterModels.[0]
for i in 1 .. 15 do
    getVisibleRows manyFilterModels.[i] |> ignore
check "filter visible cache keeps many filters" (obj.ReferenceEquals(getVisibleRows manyFilterModels.[0], manyVisibleRowsFirst)) true
let aggregateFilterKeyValueCount = 50001
let aggregateFilterKeyValues prefix =
    [ for index in 0 .. aggregateFilterKeyValueCount - 1 ->
        sprintf "%s-%05d" prefix index ]
    |> Set.ofList
let aggregateFilterKeyValuesA = aggregateFilterKeyValues "a"
let aggregateFilterKeyValuesB = aggregateFilterKeyValues "b"
let aggregateVisibleRowsBaseModel =
    { createEmpty 1 1 with
        Cells = [ { Row = 0; Col = 0 }, cell (S "a-00000") ] |> Map.ofList }
let aggregateVisibleRowsModelA =
    { aggregateVisibleRowsBaseModel with
        Filters = Map.ofList [ 0, { Column = 0; VisibleValues = aggregateFilterKeyValuesA } ] }
let aggregateVisibleRowsModelB =
    { aggregateVisibleRowsBaseModel with
        Filters = Map.ofList [ 0, { Column = 0; VisibleValues = aggregateFilterKeyValuesB } ] }
let aggregateVisibleRowsA = getVisibleRows aggregateVisibleRowsModelA
getVisibleRows aggregateVisibleRowsModelB |> ignore
let aggregateVisibleRowsAAgain = getVisibleRows aggregateVisibleRowsModelA
check "filter visible cache keeps aggregate-sized keys"
    (aggregateVisibleRowsA = Set.singleton 0 && obj.ReferenceEquals(aggregateVisibleRowsA, aggregateVisibleRowsAAgain))
    true
let rowFilterManyModelCount = 24
let rowFilterManyRowCount = 5000
let rowFilterManyModels =
    [| for modelIndex in 0 .. rowFilterManyModelCount - 1 ->
        { createEmpty rowFilterManyRowCount 2 with
            Cells =
                [ for row in 0 .. rowFilterManyRowCount - 1 ->
                    { Row = row; Col = 0 }, cell (S (sprintf "drop-%02d-%05d" modelIndex row)) ]
                |> Map.ofList }
        |> applyFilter 0 (Set.singleton "keep") |]
let rowFilterVisibleRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in rowFilterManyModels do
            total <- total + (getVisibleRows model).Count
    total
rowFilterVisibleRun 3 |> ignore
let rowFilterVisibleSw = Stopwatch.StartNew()
let rowFilterVisibleTotal = rowFilterVisibleRun 60
rowFilterVisibleSw.Stop()
check "filter visible cache many models"
    (rowFilterVisibleSw.Elapsed.TotalMilliseconds < 6.0 && rowFilterVisibleTotal = 0)
    true

// remove filter restores
check "remove filter"     (getVisibleRows (removeFilter 0 fEast) |> Set.count) 5
check "missing filter identity" (obj.ReferenceEquals(removeFilter 9 model, model)) true

// render-side row collapsing: filtered-out rows get height 0, header (row 0)
// and visible rows keep their height. Filter {East}: row 0 (header, East) and
// row 2 (East) stay; rows 1,3,4 collapse to 0.
let collapsed = applyFilterRowHeights fEast
check "header row kept"   (getRowHeight 0 collapsed) 25.0
check "visible row kept"  (getRowHeight 2 collapsed) 25.0
check "hidden row1 zero"  (getRowHeight 1 collapsed) 0.0
check "hidden row3 zero"  (getRowHeight 3 collapsed) 0.0
check "hidden row4 zero"  (getRowHeight 4 collapsed) 0.0
let collapsedAgain = applyFilterRowHeights collapsed
check "filter projection cache identity" (obj.ReferenceEquals(collapsed.FilterHiddenRows, collapsedAgain.FilterHiddenRows)) true
applyFilterRowHeights fEastWithChangedCells |> ignore
let collapsedAfterSourceSwitch = applyFilterRowHeights fEast
check "filter projection keeps first source" (obj.ReferenceEquals(collapsedAfterSourceSwitch.FilterHiddenRows, collapsed.FilterHiddenRows)) true
check "same projected filter identity" (obj.ReferenceEquals(applyFilter 0 (Set.ofList [ "East" ]) collapsed, collapsed)) true
let allRowsFilter = applyFilter 0 (Set.ofList [ "East"; "North"; "West" ]) model
check "filter projection no hidden identity" (obj.ReferenceEquals(applyFilterRowHeights allRowsFilter, allRowsFilter)) true
let directNorth =
    { collapsed with
        Filters = Map.add 0 { Column = 0; VisibleValues = Set.ofList [ "North" ] } collapsed.Filters }
let collapsedNorth = applyFilterRowHeights directNorth
check "filter projection invalidates filter" (collapsedNorth.FilterHiddenRows) [| 1; 2; 4 |]
let collapsedEastAgain = applyFilterRowHeights fEast
check "filter projection keeps first" (obj.ReferenceEquals(collapsedEastAgain.FilterHiddenRows, collapsed.FilterHiddenRows)) true
let manyFilterProjectionFirst = applyFilterRowHeights manyFilterModels.[0]
for i in 1 .. 15 do
    applyFilterRowHeights manyFilterModels.[i] |> ignore
check "filter projection keeps many filters"
    (obj.ReferenceEquals((applyFilterRowHeights manyFilterModels.[0]).FilterHiddenRows, manyFilterProjectionFirst.FilterHiddenRows)) true
let aggregateFilterKeyBaseModel =
    { createEmpty 2 1 with
        Cells = [ { Row = 1; Col = 0 }, cell (S "hidden") ] |> Map.ofList }
let aggregateFilterKeyModelA =
    { aggregateFilterKeyBaseModel with
        Filters = Map.ofList [ 0, { Column = 0; VisibleValues = aggregateFilterKeyValuesA } ] }
let aggregateFilterKeyModelB =
    { aggregateFilterKeyBaseModel with
        Filters = Map.ofList [ 0, { Column = 0; VisibleValues = aggregateFilterKeyValuesB } ] }
let aggregateFilterKeyProjectedA = applyFilterRowHeights aggregateFilterKeyModelA
applyFilterRowHeights aggregateFilterKeyModelB |> ignore
let aggregateFilterKeyProjectedAAgain = applyFilterRowHeights aggregateFilterKeyModelA
check "filter projection keeps aggregate-sized keys"
    (obj.ReferenceEquals(aggregateFilterKeyProjectedA.FilterHiddenRows, aggregateFilterKeyProjectedAAgain.FilterHiddenRows))
    true
let sparseMultiFilterRowCount = 100000
let sparseMultiFilterColumnCount = 64
let sparseMultiFilterModel =
    { createEmpty sparseMultiFilterRowCount sparseMultiFilterColumnCount with
        Cells =
            [ for row in 0 .. sparseMultiFilterRowCount - 1 ->
                { Row = row; Col = 0 }, cell (S (if row % 3 = 0 then "keep" else "drop")) ]
            |> Map.ofList
        Filters =
            [ for col in 0 .. sparseMultiFilterColumnCount - 1 ->
                col, { Column = col; VisibleValues = Set.singleton "keep" } ]
            |> Map.ofList }
let sparseMultiFilterSw = Stopwatch.StartNew()
let sparseMultiFilterProjected = applyFilterRowHeights sparseMultiFilterModel
sparseMultiFilterSw.Stop()
check "filter projection sparse many columns"
    (sparseMultiFilterProjected.FilterHiddenRows.Length = sparseMultiFilterRowCount - 1
     && sparseMultiFilterSw.Elapsed.TotalMilliseconds < 120.0)
    true
let sparseBlankAllowedFilterModel =
    { sparseMultiFilterModel with
        Filters =
            [ for col in 0 .. sparseMultiFilterColumnCount - 1 ->
                col, { Column = col; VisibleValues = Set.ofList [ "keep"; "" ] } ]
            |> Map.ofList }
let sparseBlankAllowedFilterSw = Stopwatch.StartNew()
let sparseBlankAllowedFilterProjected = applyFilterRowHeights sparseBlankAllowedFilterModel
sparseBlankAllowedFilterSw.Stop()
check "filter projection sparse blank-allowed"
    (sparseBlankAllowedFilterProjected.FilterHiddenRows.Length = 66666
     && sparseBlankAllowedFilterSw.Elapsed.TotalMilliseconds < 90.0)
    true
let denseSingleFilterProjectionBytes =
    let rowCount = 100000
    let model =
        { createEmpty rowCount 1 with
            Cells =
                [ for row in 0 .. rowCount - 1 ->
                    { Row = row; Col = 0 }, cell (S (if row % 3 = 0 then "keep" else "drop")) ]
                |> Map.ofList
            Filters = Map.ofList [ 0, { Column = 0; VisibleValues = Set.singleton "keep" } ] }
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetAllocatedBytesForCurrentThread()
    let sw = Stopwatch.StartNew()
    let projected = applyFilterRowHeights model
    sw.Stop()
    let bytes = System.GC.GetAllocatedBytesForCurrentThread() - before
    check "filter projection dense single count" projected.FilterHiddenRows.Length 66666
    check "filter projection dense single perf" (sw.Elapsed.TotalMilliseconds < 90.0) true
    bytes
check "filter projection dense single allocation" (denseSingleFilterProjectionBytes < 2000000L) true
let rowFilterProjectionRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in rowFilterManyModels do
            total <- total + (applyFilterRowHeights model).FilterHiddenRows.Length
    total
rowFilterProjectionRun 3 |> ignore
let rowFilterProjectionSw = Stopwatch.StartNew()
let rowFilterProjectionTotal = rowFilterProjectionRun 60
rowFilterProjectionSw.Stop()
check "filter projection cache many models"
    (rowFilterProjectionSw.Elapsed.TotalMilliseconds < 8.0 && rowFilterProjectionTotal > 0)
    true
let rowFilterRetainedModelCount = 80
let rowFilterRetainedRowCount = 20000
let rowFilterRetainedModel modelIndex =
    { createEmpty rowFilterRetainedRowCount 1 with
        Cells =
            [ for row in 0 .. rowFilterRetainedRowCount - 1 ->
                { Row = row; Col = 0 }, cell (S (sprintf "keep-%02d" modelIndex)) ]
            |> Map.ofList }
let rowFilterRetainedWithFilter values model =
    { model with Filters = Map.ofList [ 0, { Column = 0; VisibleValues = values } ] }
let rowFilterRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. rowFilterRetainedModelCount - 1 do
        let model = rowFilterRetainedModel modelIndex
        applyFilterRowHeights (rowFilterRetainedWithFilter (Set.singleton (sprintf "keep-%02d" modelIndex)) model) |> ignore
        getVisibleRows (rowFilterRetainedWithFilter Set.empty model) |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "filter row caches bounded source memory" (rowFilterRetainedBytes < 160000000L) true
let rowFilterSourceTextModel modelIndex =
    { createEmpty 4 2 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (S "keep")
              { Row = 0; Col = 1 }, cell (S (System.String(char (65 + (modelIndex % 26)), 2000000))) ]
            |> Map.ofList }
    |> applyFilter 0 (Set.singleton "keep")
let rowFilterSourceTextBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. 39 do
        let model = rowFilterSourceTextModel modelIndex
        applyFilterRowHeights model |> ignore
        getVisibleRows model |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "filter row caches source text memory" (rowFilterSourceTextBytes < 30000000L) true
let filterKeyRetainedModelCount = 80
let filterKeyRetainedValueCount = 10000
let filterKeyRetainedValues modelIndex =
    [ for index in 0 .. filterKeyRetainedValueCount - 1 ->
        sprintf "filter-key-retained-%02d-%05d" modelIndex index ]
    |> Set.ofList
let filterKeyRetainedModel modelIndex =
    { createEmpty 1 1 with
        Filters = Map.ofList [ 0, { Column = 0; VisibleValues = filterKeyRetainedValues modelIndex } ] }
let filterKeyRetainedBytes, filterKeyRetainedTotal =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable total = 0
    for modelIndex in 0 .. filterKeyRetainedModelCount - 1 do
        let model = filterKeyRetainedModel modelIndex
        total <- total + (applyFilterRowHeights model).FilterHiddenRows.Length
        total <- total + (getVisibleRows model).Count
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before, total
check "filter key cache retained memory" (filterKeyRetainedBytes < 60000000L && filterKeyRetainedTotal = 0) true
let filterKeySourceTextModelCount = 40
let filterKeySourceTextValue modelIndex =
    String.replicate 2000000 (string (char (65 + (modelIndex % 26))))
let filterKeySourceTextModel modelIndex =
    { createEmpty 1 1 with
        Filters = Map.ofList [ 0, { Column = 0; VisibleValues = Set.singleton (filterKeySourceTextValue modelIndex) } ] }
let filterKeySourceTextBytes, filterKeySourceTextTotal =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable total = 0
    for modelIndex in 0 .. filterKeySourceTextModelCount - 1 do
        let model = filterKeySourceTextModel modelIndex
        total <- total + (applyFilterRowHeights model).FilterHiddenRows.Length
        total <- total + (getVisibleRows model).Count
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before, total
check "filter key cache source text memory" (filterKeySourceTextBytes < 30000000L && filterKeySourceTextTotal = 0) true
// no filters → untouched (no-op): all rows keep default height
check "no-filter no-op"   (getRowHeight 1 (applyFilterRowHeights model)) 25.0

let replaceModel =
    { createEmpty 10 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (S "target Target")
              { Row = 1; Col = 0 }, cell (S "no match")
              { Row = 2; Col = 0 }, cell (CellValue.Number 42.0) ]
            |> Map.ofList }
let replacedInsensitive = replaceAll "target" "x" false replaceModel
check "replace all insensitive" (getCellValue { Row = 0; Col = 0 } replacedInsensitive) (S "x x")
check "replace all no match" (getCellValue { Row = 1; Col = 0 } replacedInsensitive) (S "no match")
check "replace all missing no-op" (obj.ReferenceEquals(replaceAll "missing" "x" false replaceModel, replaceModel)) true
let replacedInsensitiveWithCount, replacedInsensitiveCount = replaceAllWithCount "target" "x" false replaceModel
check "replace all count insensitive" replacedInsensitiveCount 1
check "replace all count text" (getCellValue { Row = 0; Col = 0 } replacedInsensitiveWithCount) (S "x x")
let replaceAllChangedModel, replaceAllChangedCount, replaceAllOldCells, replaceAllNewCells =
    replaceAllWithChanges "target" "x" false replaceModel
let replaceAllOldAddr, replaceAllOldCell = replaceAllOldCells.[0]
let replaceAllNewAddr, replaceAllNewCell = replaceAllNewCells.[0]
check "replace all changes count" replaceAllChangedCount 1
check "replace all changes old count" replaceAllOldCells.Count 1
check "replace all changes new count" replaceAllNewCells.Count 1
check "replace all changes text" (getCellValue { Row = 0; Col = 0 } replaceAllChangedModel) (S "x x")
check "replace all changes old addr" replaceAllOldAddr { Row = 0; Col = 0 }
check "replace all changes new addr" replaceAllNewAddr { Row = 0; Col = 0 }
check "replace all changes old value" (replaceAllOldCell |> Option.map (fun c -> c.Value)) (Some (S "target Target"))
check "replace all changes new value" (replaceAllNewCell |> Option.map (fun c -> c.Value)) (Some (S "x x"))
let replaceMissingWithCount, replaceMissingCount = replaceAllWithCount "missing" "x" false replaceModel
check "replace all count missing" replaceMissingCount 0
check "replace all count missing no-op" (obj.ReferenceEquals(replaceMissingWithCount, replaceModel)) true
let replaceMissingChangedModel, replaceMissingChangedCount, replaceMissingOldCells, replaceMissingNewCells =
    replaceAllWithChanges "missing" "x" false replaceModel
check "replace all changes missing count" replaceMissingChangedCount 0
check "replace all changes missing identity" (obj.ReferenceEquals(replaceMissingChangedModel, replaceModel)) true
check "replace all changes missing old" replaceMissingOldCells.Count 0
check "replace all changes missing new" replaceMissingNewCells.Count 0
let replacedSensitive = replaceAll "target" "x" true replaceModel
check "replace all sensitive" (getCellValue { Row = 0; Col = 0 } replacedSensitive) (S "x Target")
check "replace all same sensitive no-op" (obj.ReferenceEquals(replaceAll "target" "target" true replaceModel, replaceModel)) true
let replaceSameWithCount, replaceSameCount = replaceAllWithCount "target" "target" true replaceModel
check "replace all count same sensitive" replaceSameCount 0
check "replace all count same sensitive no-op" (obj.ReferenceEquals(replaceSameWithCount, replaceModel)) true
let lowerOnlyReplaceModel =
    { createEmpty 5 2 with
        Cells = [ { Row = 0; Col = 0 }, cell (S "target target") ] |> Map.ofList }
check "replace all same insensitive no-op" (obj.ReferenceEquals(replaceAll "target" "target" false lowerOnlyReplaceModel, lowerOnlyReplaceModel)) true
let replaceSameInsensitiveWithCount, replaceSameInsensitiveCount = replaceAllWithCount "target" "target" false lowerOnlyReplaceModel
check "replace all count same insensitive" replaceSameInsensitiveCount 0
check "replace all count same insensitive no-op" (obj.ReferenceEquals(replaceSameInsensitiveWithCount, lowerOnlyReplaceModel)) true
check "replace selection same insensitive none" (replaceAtSelection "target" "target" false { lowerOnlyReplaceModel with Selection = SingleCell (0, 0) }).IsNone true
let regexLiteralReplaceModel =
    { createEmpty 5 2 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (S "a.b")
              { Row = 1; Col = 0 }, cell (S "a[b")
              { Row = 2; Col = 0 }, cell (S "target") ]
            |> Map.ofList }
check "replace all dot literal" (getCellValue { Row = 0; Col = 0 } (replaceAll "." "X" false regexLiteralReplaceModel)) (S "aXb")
check "replace all bracket literal" (getCellValue { Row = 1; Col = 0 } (replaceAll "[" "X" false regexLiteralReplaceModel)) (S "aXb")
check "replace all dollar literal" (getCellValue { Row = 2; Col = 0 } (replaceAll "target" "$1" false regexLiteralReplaceModel)) (S "$1")
let replacedSelectionBracket = replaceAtSelection "[" "X" false { regexLiteralReplaceModel with Selection = SingleCell (1, 0) }
check "replace selection bracket literal" (replacedSelectionBracket |> Option.map (getCellValue { Row = 1; Col = 0 })) (Some (S "aXb"))
let replacedSelectionDollar = replaceAtSelection "target" "$1" false { regexLiteralReplaceModel with Selection = SingleCell (2, 0) }
check "replace selection dollar literal" (replacedSelectionDollar |> Option.map (getCellValue { Row = 2; Col = 0 })) (Some (S "$1"))
check "replace all number text" (getCellValue { Row = 2; Col = 0 } (replaceAll "42" "forty-two" true replaceModel)) (S "forty-two")
check "search cells insensitive" (searchCells "TARGET" false replaceModel |> Array.toList) [ { Row = 0; Col = 0 } ]
check "search cells empty no results" (searchCells "" false replaceModel |> Array.toList) []
check "search cells number digit" (searchCells "42" true replaceModel |> Array.toList) [ { Row = 2; Col = 0 } ]
check "search cells number text impossible" (searchCells "missing" false replaceModel |> Array.toList) []
let cachedSearchFirst = searchCells "TARGET" false replaceModel
let cachedSearchSecond = searchCells "TARGET" false replaceModel
check "search cells cache identity" (obj.ReferenceEquals(cachedSearchFirst, cachedSearchSecond)) true
let cachedSearchCaseFirst = searchCells "target" true replaceModel
let cachedSearchOtherQuery = searchCells "42" true replaceModel
let cachedSearchCaseSecond = searchCells "target" true replaceModel
check "search cells cache keeps first query" (obj.ReferenceEquals(cachedSearchCaseFirst, cachedSearchCaseSecond)) true
check "search cells cache keeps other query" (obj.ReferenceEquals(cachedSearchOtherQuery, searchCells "42" true replaceModel)) true
let editedSearchModel = replaceModel |> setCell { Row = 1; Col = 0 } (cell (S "TARGET"))
check "search cells cache invalidates edit" (searchCells "TARGET" false editedSearchModel |> Array.toList) [ { Row = 0; Col = 0 }; { Row = 1; Col = 0 } ]
check "search cells cache keeps first source" (obj.ReferenceEquals(searchCells "TARGET" false replaceModel, cachedSearchFirst)) true
let searchCellsCacheableMissRows = 500000
let searchCellsCacheableMissModel =
    { createEmpty searchCellsCacheableMissRows 1 with
        Cells =
            seq {
                for row in 0 .. searchCellsCacheableMissRows - 1 do
                    yield { Row = row; Col = 0 }, cell (S "same")
            }
            |> Map.ofSeq }
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let searchCellsCacheableMissBefore = System.GC.GetAllocatedBytesForCurrentThread()
let searchCellsCacheableMissSw = Stopwatch.StartNew()
let searchCellsCacheableMissResults = searchCells "not-present-cacheable" false searchCellsCacheableMissModel
searchCellsCacheableMissSw.Stop()
let searchCellsCacheableMissBytes =
    System.GC.GetAllocatedBytesForCurrentThread() - searchCellsCacheableMissBefore
check "search cells cacheable miss count" searchCellsCacheableMissResults.Length 0
check "search cells cacheable miss perf" (searchCellsCacheableMissSw.Elapsed.TotalMilliseconds < 240.0) true
check "search cells cacheable miss allocation" (searchCellsCacheableMissBytes < 100000L) true
check "search cells cacheable miss cached"
    (obj.ReferenceEquals(searchCells "not-present-cacheable" false searchCellsCacheableMissModel, searchCellsCacheableMissResults))
    true
let oversizedSearchCacheSkipped () =
    let format =
        { defaultFormat with
            FontName = Some "SearchPerfFont"
            NumberFormat = Some (System.String('0', 256))
            BgColor = Some "#334455"
            TextColor = Some "#e0e0e0" }
    let model =
        { createEmpty 200001 1 with
            Cells =
                [ for row in 0 .. 200000 ->
                    { Row = row; Col = 0 },
                    { Value = S "x"
                      Format = format
                      Comment = Some { Text = "search cache source text"; Author = Some "perf" } } ]
                |> Map.ofList }
    let sw = Stopwatch.StartNew()
    let resultsA = searchCells "x" true model
    let resultsB = searchCells "x" true model
    sw.Stop()
    resultsA.Length, obj.ReferenceEquals(resultsA, resultsB), sw.Elapsed.TotalMilliseconds
let oversizedSearchCount, oversizedSearchReused, oversizedSearchElapsed = oversizedSearchCacheSkipped ()
check "search cells cache skips oversized results"
    (oversizedSearchCount > 200000 && not oversizedSearchReused)
    true
if oversizedSearchElapsed < 1500.0 then
    pass <- pass + 1
else
    fail <- fail + 1
    printfn
        "FAIL %-26s expected under %.1f ms got %.1f ms"
        "search oversized skips source"
        1500.0
        oversizedSearchElapsed
let searchCellsOversizedSourceWeakReference () =
    let cells =
        [ for row in 0 .. 500000 ->
            { Row = row; Col = 0 }, cell (S "same") ]
        |> Map.ofList
    let model = { createEmpty 500001 1 with Cells = cells }
    searchCells "missing" false model |> ignore
    System.WeakReference(cells)
let searchCellsOversizedSourceWeak = searchCellsOversizedSourceWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "search cells cache skips oversized source"
    searchCellsOversizedSourceWeak.IsAlive
    false
let searchCellsOversizedTextWeakReference () =
    let payload = System.String('Q', 6 * 1024 * 1024)
    let cells =
        [ { Row = 0; Col = 0 }, cell (S payload) ]
        |> Map.ofList
    let model = { createEmpty 1 1 with Cells = cells }
    searchCells "missing" false model |> ignore
    System.WeakReference(cells)
let searchCellsOversizedTextWeak = searchCellsOversizedTextWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "search cells cache skips oversized text"
    searchCellsOversizedTextWeak.IsAlive
    false
let searchCellsOversizedQueryWeakReference () =
    let payload = System.String('Z', 6 * 1024 * 1024)
    let model =
        { createEmpty 1 1 with
            Cells = [ { Row = 0; Col = 0 }, cell (S "tiny") ] |> Map.ofList }
    searchCells payload false model |> ignore
    System.WeakReference(payload)
let searchCellsOversizedQueryWeak = searchCellsOversizedQueryWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "search cells cache skips oversized query"
    searchCellsOversizedQueryWeak.IsAlive
    false
let manySearchQueriesModel =
    { createEmpty 40 2 with
        Cells =
            [ for row in 0 .. 39 ->
                { Row = row; Col = 0 }, cell (S (sprintf "needle-%02d" row)) ]
            |> Map.ofList }
let manySearchFirst = searchCells "missing-00" false manySearchQueriesModel
for i in 1 .. 15 do
    searchCells (sprintf "missing-%02d" i) false manySearchQueriesModel |> ignore
check "search cells cache keeps many queries" (obj.ReferenceEquals(searchCells "missing-00" false manySearchQueriesModel, manySearchFirst)) true
let searchManyModelCount = 24
let searchManyCellCount = 8000
let searchManyModels =
    [| for modelIndex in 0 .. searchManyModelCount - 1 ->
        { createEmpty searchManyCellCount 2 with
            Cells =
                [ for row in 0 .. searchManyCellCount - 1 ->
                    { Row = row; Col = 0 }, cell (S (sprintf "search-many-%02d-%05d" modelIndex row)) ]
                |> Map.ofList } |]
let searchManyRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in searchManyModels do
            total <- total + (searchCells "missing" false model).Length
    total
searchManyRun 3 |> ignore
let searchManySw = Stopwatch.StartNew()
let searchManyTotal = searchManyRun 60
searchManySw.Stop()
check "search cells cache many models"
    (searchManySw.Elapsed.TotalMilliseconds < 20.0 && searchManyTotal = 0)
    true
let searchCellsRetainedModelCount = 80
let searchCellsRetainedRowCount = 20000
let searchCellsRetainedModel modelIndex =
    { createEmpty searchCellsRetainedRowCount 1 with
        Cells =
            [ for row in 0 .. searchCellsRetainedRowCount - 1 ->
                { Row = row; Col = 0 }, cell (S (sprintf "search-retained-%02d-%05d" modelIndex row)) ]
            |> Map.ofList }
let searchCellsRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. searchCellsRetainedModelCount - 1 do
        searchCells "missing" false (searchCellsRetainedModel modelIndex) |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "search cells cache bounded source memory" (searchCellsRetainedBytes < 180000000L) true
let searchCellsSourceTextRetainedModel modelIndex =
    { createEmpty 4 1 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (S (System.String(char (65 + (modelIndex % 26)), 2000000))) ]
            |> Map.ofList }
let searchCellsSourceTextRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. 39 do
        searchCells "not-present" false (searchCellsSourceTextRetainedModel modelIndex) |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "search cells cache source text memory" (searchCellsSourceTextRetainedBytes < 30000000L) true
let directMissingReplaceModel =
    { createEmpty 5 2 with
        Cells = [ { Row = 0; Col = 0 }, cell (S "alpha beta") ] |> Map.ofList }
let directMissingReplaceA, directMissingReplaceCountA = replaceAllWithCount "absent-direct" "x" false directMissingReplaceModel
let otherDirectMissingReplace, otherDirectMissingCount = replaceAllWithCount "still-absent" "x" false directMissingReplaceModel
let directMissingReplaceB, directMissingReplaceCountB = replaceAllWithCount "absent-direct" "x" false directMissingReplaceModel
check "replace all other missing count" otherDirectMissingCount 0
check "replace all other missing identity" (obj.ReferenceEquals(otherDirectMissingReplace, directMissingReplaceModel)) true
check "replace all direct missing first count" directMissingReplaceCountA 0
check "replace all direct missing second count" directMissingReplaceCountB 0
check "replace all direct missing identity" (obj.ReferenceEquals(directMissingReplaceB, directMissingReplaceModel)) true
let directMissingChangedReplaceModel =
    directMissingReplaceModel |> setCell { Row = 1; Col = 0 } (cell (S "absent-direct"))
let directMissingChangedReplace, directMissingChangedReplaceCount = replaceAllWithCount "absent-direct" "x" false directMissingChangedReplaceModel
check "replace all direct missing cache invalidates count" directMissingChangedReplaceCount 1
check "replace all direct missing cache invalidates text" (getCellValue { Row = 1; Col = 0 } directMissingChangedReplace) (S "x")
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let replaceCacheableMissingBefore = System.GC.GetAllocatedBytesForCurrentThread()
let replaceCacheableMissingSw = Stopwatch.StartNew()
let replaceCacheableMissingModel, replaceCacheableMissingCount =
    replaceAllWithCount "replace-not-present-cacheable" "x" false searchCellsCacheableMissModel
replaceCacheableMissingSw.Stop()
let replaceCacheableMissingBytes =
    System.GC.GetAllocatedBytesForCurrentThread() - replaceCacheableMissingBefore
check "replace all cacheable missing count" replaceCacheableMissingCount 0
check "replace all cacheable missing identity"
    (obj.ReferenceEquals(replaceCacheableMissingModel, searchCellsCacheableMissModel))
    true
check "replace all cacheable missing perf" (replaceCacheableMissingSw.Elapsed.TotalMilliseconds < 240.0) true
check "replace all cacheable missing allocation" (replaceCacheableMissingBytes < 100000L) true
let replaceCacheableMissingCached, replaceCacheableMissingCachedCount =
    replaceAllWithCount "replace-not-present-cacheable" "x" false searchCellsCacheableMissModel
check "replace all cacheable missing cached count" replaceCacheableMissingCachedCount 0
check "replace all cacheable missing cached identity"
    (obj.ReferenceEquals(replaceCacheableMissingCached, searchCellsCacheableMissModel))
    true
let replaceBooleanSkipRows = 250000
let replaceBooleanSkipModel =
    { createEmpty replaceBooleanSkipRows 1 with
        Cells =
            [ for row in 0 .. replaceBooleanSkipRows - 1 ->
                { Row = row; Col = 0 }, cell (CellValue.Boolean (row % 2 = 0)) ]
            |> Map.ofList }
let replaceBooleanSkipSw = Stopwatch.StartNew()
let replaceBooleanSkipResult, replaceBooleanSkipCount =
    replaceAllWithCount "xyz" "x" false replaceBooleanSkipModel
replaceBooleanSkipSw.Stop()
check "replace all boolean impossible count" replaceBooleanSkipCount 0
check "replace all boolean impossible identity" (obj.ReferenceEquals(replaceBooleanSkipResult, replaceBooleanSkipModel)) true
if replaceBooleanSkipSw.Elapsed.TotalMilliseconds < 250.0 then
    pass <- pass + 1
else
    fail <- fail + 1
    printfn
        "FAIL %-26s expected under %.1f ms got %.1f ms"
        "replace boolean skip perf"
        250.0
        replaceBooleanSkipSw.Elapsed.TotalMilliseconds
let replaceIdentityCacheRows = 30000
let replaceIdentityCacheModel =
    { createEmpty replaceIdentityCacheRows 1 with
        Cells =
            [ for row in 0 .. replaceIdentityCacheRows - 1 ->
                { Row = row; Col = 0 }, cell (S "target") ]
            |> Map.ofList }
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let replaceIdentityWarmBefore = System.GC.GetAllocatedBytesForCurrentThread()
let replaceIdentityWarm, replaceIdentityWarmCount = replaceAllWithCount "target" "target" false replaceIdentityCacheModel
let replaceIdentityWarmBytes = System.GC.GetAllocatedBytesForCurrentThread() - replaceIdentityWarmBefore
check "replace all identity cache warm count" replaceIdentityWarmCount 0
check "replace all identity cache warm identity" (obj.ReferenceEquals(replaceIdentityWarm, replaceIdentityCacheModel)) true
check "replace all identity cache warm allocation" (replaceIdentityWarmBytes < 200000L) true
let replaceIdentityCacheRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        let replaced, count = replaceAllWithCount "target" "target" false replaceIdentityCacheModel
        if not (obj.ReferenceEquals(replaced, replaceIdentityCacheModel)) then
            failwith "replace identity cache changed the source model"
        total <- total + count
    total
replaceIdentityCacheRun 5 |> ignore
let replaceIdentityCacheSw = Stopwatch.StartNew()
let replaceIdentityCacheTotal = replaceIdentityCacheRun 80
replaceIdentityCacheSw.Stop()
check "replace all identity cache count" replaceIdentityCacheTotal 0
if replaceIdentityCacheSw.Elapsed.TotalMilliseconds < 25.0 then
    pass <- pass + 1
else
    fail <- fail + 1
    printfn
        "FAIL %-26s expected under %.1f ms got %.1f ms"
        "replace identity cache perf"
        25.0
        replaceIdentityCacheSw.Elapsed.TotalMilliseconds
let replaceIdentityChangedModel =
    replaceIdentityCacheModel |> setCell { Row = 0; Col = 0 } (cell (S "Target"))
let replaceIdentityChanged, replaceIdentityChangedCount = replaceAllWithCount "target" "target" false replaceIdentityChangedModel
check "replace all identity cache invalidates count" replaceIdentityChangedCount 1
check "replace all identity cache invalidates text" (getCellValue { Row = 0; Col = 0 } replaceIdentityChanged) (S "target")
let replaceIdentityOversizedTextWeakReference () =
    let payload = "target" + System.String('Q', 6 * 1024 * 1024)
    let cells =
        [ { Row = 0; Col = 0 }, cell (S payload) ]
        |> Map.ofList
    let model = { createEmpty 1 1 with Cells = cells }
    replaceAllWithCount "target" "target" false model |> ignore
    System.WeakReference(cells)
let replaceIdentityOversizedTextWeak = replaceIdentityOversizedTextWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "replace identity cache skips oversized text"
    replaceIdentityOversizedTextWeak.IsAlive
    false
let startedSearchModel = startSearch "target" false replaceModel
check "start search query" startedSearchModel.SearchQuery (Some "target")
check "start search case" startedSearchModel.SearchCaseSensitive (Some false)
check "start search no-op identity" (obj.ReferenceEquals(startSearch "target" false startedSearchModel, startedSearchModel)) true
let caseChangedSearchModel = startSearch "target" true startedSearchModel
check "start search case invalidates" (caseChangedSearchModel.SearchResults |> Array.toList) [ { Row = 0; Col = 0 } ]
check "start search case changed" caseChangedSearchModel.SearchCaseSensitive (Some true)
let singleResultSearchModel =
    { startSearch "target" true replaceModel with
        Selection = SingleCell (0, 0) }
check "next search same identity" (nextSearchResult singleResultSearchModel |> Option.exists (fun next -> obj.ReferenceEquals(next, singleResultSearchModel))) true
check "replace all empty no-op" (obj.ReferenceEquals(replaceAll "" "x" false replaceModel, replaceModel)) true
check "replace all empty sensitive no-op" (obj.ReferenceEquals(replaceAll "" "x" true replaceModel, replaceModel)) true
let replaceEmptyWithCount, replaceEmptyCount = replaceAllWithCount "" "x" false replaceModel
check "replace all count empty" replaceEmptyCount 0
check "replace all count empty no-op" (obj.ReferenceEquals(replaceEmptyWithCount, replaceModel)) true
check "replace selection empty none" (replaceAtSelection "" "x" false { replaceModel with Selection = SingleCell (0, 0) }).IsNone true
check "replace selection empty sensitive none" (replaceAtSelection "" "x" true { replaceModel with Selection = SingleCell (0, 0) }).IsNone true
check "replace selection same sensitive none" (replaceAtSelection "target" "target" true { replaceModel with Selection = SingleCell (0, 0) }).IsNone true
let numericSearchModel =
    { createEmpty 5 2 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number System.Double.NaN)
              { Row = 1; Col = 0 }, cell (CellValue.Number System.Double.PositiveInfinity)
              { Row = 2; Col = 0 }, cell (S "missing") ]
            |> Map.ofList }
check "search cells nan preserved" (searchCells "nan" false numericSearchModel |> Array.toList) [ { Row = 0; Col = 0 } ]
check "search cells infinity preserved" (searchCells "INF" false numericSearchModel |> Array.toList) [ { Row = 1; Col = 0 } ]
check "search cells impossible still finds string" (searchCells "missing" false numericSearchModel |> Array.toList) [ { Row = 2; Col = 0 } ]
let booleanSearchModel =
    { createEmpty 5 2 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Boolean true)
              { Row = 1; Col = 0 }, cell (CellValue.Boolean false)
              { Row = 2; Col = 0 }, cell (S "missing") ]
            |> Map.ofList }
check "search cells true preserved" (searchCells "TRUE" true booleanSearchModel |> Array.toList) [ { Row = 0; Col = 0 } ]
check "search cells false preserved" (searchCells "false" false booleanSearchModel |> Array.toList) [ { Row = 1; Col = 0 } ]
check "search cells impossible skips boolean finds string" (searchCells "missing" false booleanSearchModel |> Array.toList) [ { Row = 2; Col = 0 } ]
let errorSearchModel =
    { createEmpty 5 2 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Error "N/A")
              { Row = 1; Col = 0 }, cell (CellValue.Error "VALUE!")
              { Row = 2; Col = 0 }, cell (S "missing") ]
            |> Map.ofList }
check "search cells error hash" (searchCells "#" true errorSearchModel |> Array.toList) [ { Row = 0; Col = 0 }; { Row = 1; Col = 0 } ]
check "search cells error code" (searchCells "#N/A" false errorSearchModel |> Array.toList) [ { Row = 0; Col = 0 } ]
check "search cells error raw code" (searchCells "value" false errorSearchModel |> Array.toList) [ { Row = 1; Col = 0 } ]
check "search cells impossible skips error finds string" (searchCells "missing" false errorSearchModel |> Array.toList) [ { Row = 2; Col = 0 } ]
let numericOnlyReplaceModel =
    { createEmpty 5 2 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 123.0)
              { Row = 1; Col = 0 }, cell (CellValue.Number 456.0) ]
            |> Map.ofList }
let errorOnlyReplaceModel =
    { createEmpty 5 2 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Error "N/A")
              { Row = 1; Col = 0 }, cell (CellValue.Error "VALUE!") ]
            |> Map.ofList }
let leadingStringPrimitiveReplaceModel =
    { createEmpty 5 2 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (S "ordinary")
              { Row = 1; Col = 0 }, cell (CellValue.Number 123.0)
              { Row = 2; Col = 0 }, cell (CellValue.Error "N/A") ]
            |> Map.ofList }
check "replace all numeric impossible no-op" (obj.ReferenceEquals(replaceAll "missing" "x" false numericOnlyReplaceModel, numericOnlyReplaceModel)) true
check "replace all error impossible no-op" (obj.ReferenceEquals(replaceAll "missing" "x" false errorOnlyReplaceModel, errorOnlyReplaceModel)) true
check "replace all mixed primitive string" (getCellValue { Row = 2; Col = 0 } (replaceAll "missing" "x" false numericSearchModel)) (S "x")
check "replace all leading string primitive no-op" (obj.ReferenceEquals(replaceAll "missing" "x" false leadingStringPrimitiveReplaceModel, leadingStringPrimitiveReplaceModel)) true

let denseReplaceModel =
    { createEmpty 4096 1 with
        Cells =
            [ for row in 0 .. 4095 ->
                { Row = row; Col = 0 }, cell (S "target") ]
            |> Map.ofList }
let denseReplaced = replaceAll "target" "x" true denseReplaceModel
check "replace all dense first" (getCellValue { Row = 0; Col = 0 } denseReplaced) (S "x")
check "replace all dense last" (getCellValue { Row = 4095; Col = 0 } denseReplaced) (S "x")
let denseChanged, denseChangedCount, denseOldCells, denseNewCells =
    replaceAllWithChanges "target" "x" true denseReplaceModel
check "replace all dense changes count" denseChangedCount 4096
check "replace all dense old count" denseOldCells.Count 4096
check "replace all dense new count" denseNewCells.Count 4096
check "replace all dense changes last" (getCellValue { Row = 4095; Col = 0 } denseChanged) (S "x")
let sparseEdgeReplaceModel =
    { createEmpty 5000 1 with
        Cells =
            [ for row in 0 .. 4999 ->
                let text = if row = 0 || row = 4999 then "target" else "other"
                { Row = row; Col = 0 }, cell (S text) ]
            |> Map.ofList }
let sparseEdgeReplaced = replaceAll "target" "x" true sparseEdgeReplaceModel
check "replace all sparse edge first" (getCellValue { Row = 0; Col = 0 } sparseEdgeReplaced) (S "x")
check "replace all sparse edge middle" (getCellValue { Row = 2500; Col = 0 } sparseEdgeReplaced) (S "other")
check "replace all sparse edge last" (getCellValue { Row = 4999; Col = 0 } sparseEdgeReplaced) (S "x")

let sparseCopyModel =
    { createEmpty 1000 10 with
        Cells =
            [ { Row = 10; Col = 2 }, cell (S "a")
              { Row = 500; Col = 5 }, cell (S "b")
              { Row = 999; Col = 9 }, cell (S "c") ]
            |> Map.ofList }
let sparseClip = serializeSelection (Range (0, 0, 999, 9)) sparseCopyModel false
check "sparse copy count" sparseClip.Cells.Length 3
check "sparse copy offsets"
    (sparseClip.Cells |> List.map (fun (r, c, data) -> r, c, data.Value))
    [ (10, 2, S "a"); (500, 5, S "b"); (999, 9, S "c") ]
let sparseOffsetClip = serializeSelection (Range (10, 2, 999, 9)) sparseCopyModel false
check "sparse copy relative offsets"
    (sparseOffsetClip.Cells |> List.map (fun (r, c, data) -> r, c, data.Value))
    [ (0, 0, S "a"); (490, 3, S "b"); (989, 7, S "c") ]

let formulaClip =
    { Cells =
        [ 0, 0, cell (CellValue.Formula ("=A1+B2", CellValue.Empty))
          1, 0, cell (CellValue.Formula ("=A1+B2", CellValue.Empty)) ]
      IsCut = false
      SourceStartRow = 0
      SourceStartCol = 0 }
let pastedFormulaModel = pasteClipboard { Row = 2; Col = 1 } formulaClip (createEmpty 10 10)
check "paste formula adjusted first" (getCellValue { Row = 2; Col = 1 } pastedFormulaModel) (CellValue.Formula ("=B3+C4", CellValue.Empty))
check "paste formula adjusted repeat" (getCellValue { Row = 3; Col = 1 } pastedFormulaModel) (CellValue.Formula ("=B3+C4", CellValue.Empty))

let valueClip =
    { Cells =
        [ 0, 0, cell (S "north")
          1, 2, cell (CellValue.Number 12.0) ]
      IsCut = false
      SourceStartRow = 0
      SourceStartCol = 0 }
let pastedValueModel = pasteClipboard { Row = 4; Col = 3 } valueClip (createEmpty 10 10)
check "paste values empty first" (getCellValue { Row = 4; Col = 3 } pastedValueModel) (S "north")
check "paste values empty second" (getCellValue { Row = 5; Col = 5 } pastedValueModel) (CellValue.Number 12.0)
check "paste values same identity" (obj.ReferenceEquals(pasteClipboard { Row = 4; Col = 3 } valueClip pastedValueModel, pastedValueModel)) true
check "paste values repeat identity" (obj.ReferenceEquals(pasteClipboard { Row = 4; Col = 3 } valueClip pastedValueModel, pastedValueModel)) true
let dualPasteModel =
    pastedValueModel
    |> pasteClipboard { Row = 6; Col = 3 } valueClip
pasteClipboard { Row = 4; Col = 3 } valueClip dualPasteModel |> ignore
pasteClipboard { Row = 6; Col = 3 } valueClip dualPasteModel |> ignore
check "paste values cache keeps first target" (obj.ReferenceEquals(pasteClipboard { Row = 4; Col = 3 } valueClip dualPasteModel, dualPasteModel)) true
check "paste values cache keeps first source" (obj.ReferenceEquals(pasteClipboard { Row = 4; Col = 3 } valueClip pastedValueModel, pastedValueModel)) true

let pasteChangedResultRows = 20000
let pasteChangedResultCols = 3
let pasteChangedResultClip =
    { Cells =
        [ for row in 0 .. pasteChangedResultRows - 1 do
            for col in 0 .. pasteChangedResultCols - 1 ->
                row, col, cell (CellValue.Number (float (row * pasteChangedResultCols + col))) ]
      IsCut = false
      SourceStartRow = 0
      SourceStartCol = 0 }
let pasteChangedResultTarget = { Row = 0; Col = 0 }
let pasteChangedResultModel =
    pasteClipboard
        pasteChangedResultTarget
        pasteChangedResultClip
        (createEmpty (pasteChangedResultRows + 2) (pasteChangedResultCols + 2))
let pasteChangedResultSw = Stopwatch.StartNew()
let pasteChangedResultRepeat = pasteClipboard pasteChangedResultTarget pasteChangedResultClip pasteChangedResultModel
pasteChangedResultSw.Stop()
check "paste changed result identity" (obj.ReferenceEquals(pasteChangedResultRepeat, pasteChangedResultModel)) true
if pasteChangedResultSw.Elapsed.TotalMilliseconds < 25.0 then
    pass <- pass + 1
else
    fail <- fail + 1
    printfn "FAIL %-26s expected under %.1f ms got %.1f ms" "paste changed result perf" 25.0 pasteChangedResultSw.Elapsed.TotalMilliseconds

let pasteIdentityCacheClipRows = 500
let pasteIdentityCacheClipCols = 4
let pasteIdentityCacheTargets =
    [| for index in 0 .. 23 -> { Row = 0; Col = index * (pasteIdentityCacheClipCols + 1) } |]
let pasteIdentityCacheClip =
    { Cells =
        [ for row in 0 .. pasteIdentityCacheClipRows - 1 do
            for col in 0 .. pasteIdentityCacheClipCols - 1 ->
                row, col, cell (CellValue.Number (float (row * pasteIdentityCacheClipCols + col))) ]
      IsCut = false
      SourceStartRow = 0
      SourceStartCol = 0 }
let pasteIdentityCacheModel =
    pasteIdentityCacheTargets
    |> Array.fold
        (fun model target -> pasteClipboard target pasteIdentityCacheClip model)
        (createEmpty (pasteIdentityCacheClipRows + 5) (pasteIdentityCacheTargets.Length * (pasteIdentityCacheClipCols + 1) + pasteIdentityCacheClipCols))
let pasteIdentityCacheRun passes =
    let mutable hits = 0
    for _pass in 1 .. passes do
        for target in pasteIdentityCacheTargets do
            let pasted = pasteClipboard target pasteIdentityCacheClip pasteIdentityCacheModel
            if obj.ReferenceEquals(pasted, pasteIdentityCacheModel) then
                hits <- hits + 1
            else
                failwithf "paste changed model at %A" target
    hits
pasteIdentityCacheRun 8 |> ignore
let pasteIdentityCacheSw = Stopwatch.StartNew()
let pasteIdentityCacheHits = pasteIdentityCacheRun 8
pasteIdentityCacheSw.Stop()
check "paste values cache many targets hits" pasteIdentityCacheHits (pasteIdentityCacheTargets.Length * 8)
if pasteIdentityCacheSw.Elapsed.TotalMilliseconds < 75.0 then
    pass <- pass + 1
else
    fail <- fail + 1
    printfn "FAIL %-26s expected under %.1f ms got %.1f ms" "paste cache many targets perf" 75.0 pasteIdentityCacheSw.Elapsed.TotalMilliseconds
let pasteIdentityManyModelCount = 24
let pasteIdentityManyModels =
    [| for modelIndex in 0 .. pasteIdentityManyModelCount - 1 ->
        let target = { Row = 0; Col = modelIndex % 3 }
        pasteClipboard target pasteIdentityCacheClip (createEmpty (pasteIdentityCacheClipRows + 5) (pasteIdentityCacheClipCols + 8)) |]
let pasteIdentityManyTargets =
    [| for modelIndex in 0 .. pasteIdentityManyModelCount - 1 -> { Row = 0; Col = modelIndex % 3 } |]
let pasteIdentityManyRun passes =
    let mutable hits = 0
    for _pass in 1 .. passes do
        for index in 0 .. pasteIdentityManyModels.Length - 1 do
            let pasted = pasteClipboard pasteIdentityManyTargets.[index] pasteIdentityCacheClip pasteIdentityManyModels.[index]
            if obj.ReferenceEquals(pasted, pasteIdentityManyModels.[index]) then
                hits <- hits + 1
            else
                failwithf "paste changed many-model source at %d" index
    hits
pasteIdentityManyRun 3 |> ignore
let pasteIdentityManySw = Stopwatch.StartNew()
let pasteIdentityManyHits = pasteIdentityManyRun 60
pasteIdentityManySw.Stop()
check "paste cache many models hits" pasteIdentityManyHits (pasteIdentityManyModelCount * 60)
if pasteIdentityManySw.Elapsed.TotalMilliseconds < 4.0 then
    pass <- pass + 1
else
    fail <- fail + 1
    printfn "FAIL %-26s expected under %.1f ms got %.1f ms" "paste cache many models perf" 4.0 pasteIdentityManySw.Elapsed.TotalMilliseconds
let pasteIdentityRetainedRowCount = 10000
let pasteIdentityRetainedModelCount = 80
let pasteIdentityRetainedClip col =
    { Cells = [ for row in 0 .. pasteIdentityRetainedRowCount - 1 -> row, 0, cell (CellValue.Number (float row)) ]
      IsCut = false
      SourceStartRow = 0
      SourceStartCol = col }
let pasteIdentityRetainedModel modelIndex =
    let col = modelIndex % 4
    { createEmpty pasteIdentityRetainedRowCount 8 with
        Cells =
            [ for row in 0 .. pasteIdentityRetainedRowCount - 1 ->
                { Row = row; Col = col }, cell (CellValue.Number (float row)) ]
            |> Map.ofList }
let pasteIdentityRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. pasteIdentityRetainedModelCount - 1 do
        let col = modelIndex % 4
        pasteClipboard { Row = 0; Col = col } (pasteIdentityRetainedClip col) (pasteIdentityRetainedModel modelIndex) |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "paste cache bounded source memory" (pasteIdentityRetainedBytes < 100000000L) true
let pasteIdentitySourceTextPayload modelIndex =
    System.String(char (65 + (modelIndex % 26)), 2000000)
let pasteIdentitySourceTextModel modelIndex =
    { createEmpty 4 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (S (pasteIdentitySourceTextPayload modelIndex)) ]
            |> Map.ofList }
let pasteIdentitySourceTextClip modelIndex =
    { Cells = [ 0, 0, cell (S (pasteIdentitySourceTextPayload modelIndex)) ]
      IsCut = false
      SourceStartRow = 0
      SourceStartCol = 0 }
let pasteIdentitySourceTextBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. 39 do
        pasteClipboard { Row = 0; Col = 0 } (pasteIdentitySourceTextClip modelIndex) (pasteIdentitySourceTextModel modelIndex) |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "paste cache source text memory" (pasteIdentitySourceTextBytes < 40000000L) true
let pasteIdentityOversizedChangedResultWeakReference () =
    let payload = System.String('P', 6 * 1024 * 1024)
    let clip =
        { Cells = [ 0, 0, cell (S payload) ]
          IsCut = false
          SourceStartRow = 0
          SourceStartCol = 0 }
    let model = pasteClipboard { Row = 0; Col = 0 } clip (createEmpty 4 4)
    getCellValue { Row = 0; Col = 0 } model |> ignore
    System.WeakReference(payload)
let pasteIdentityOversizedChangedResultWeak = pasteIdentityOversizedChangedResultWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "paste cache skips oversized changed"
    pasteIdentityOversizedChangedResultWeak.IsAlive
    false
let pasteIdentityOversizedClipboardWeakReference () =
    let cells =
        [ for row in 0 .. 250000 ->
            row, 0, cell (S "x") ]
    let clip =
        { Cells = cells
          IsCut = false
          SourceStartRow = 0
          SourceStartCol = 0 }
    let sw = Stopwatch.StartNew()
    pasteClipboard { Row = 10; Col = 0 } clip (createEmpty 1 1) |> ignore
    sw.Stop()
    if sw.Elapsed.TotalMilliseconds < 250.0 then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-26s expected under %.1f ms got %.1f ms" "paste oversized clipboard perf" 250.0 sw.Elapsed.TotalMilliseconds
    System.WeakReference(cells)
let pasteIdentityOversizedClipboardWeak = pasteIdentityOversizedClipboardWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "paste cache skips oversized clipboard"
    pasteIdentityOversizedClipboardWeak.IsAlive
    false
let pasteIdentityOversizedClipboardSkipsSourceTextElapsed =
    let rec formulaChain depth value =
        if depth = 0 then value
        else formulaChain (depth - 1) (CellValue.Formula ("", value))
    let sourceRows = 80000
    let sourceCell = cell (formulaChain 128 (S "x"))
    let sourceCells =
        [ for row in 0 .. sourceRows - 1 ->
            { Row = row; Col = 0 }, sourceCell ]
        |> Map.ofList
    let clip =
        { Cells =
            [ for row in 0 .. 200000 ->
                row, 0, cell (CellValue.Number 1.0) ]
          IsCut = false
          SourceStartRow = 0
          SourceStartCol = 0 }
    let model = { createEmpty sourceRows 1 with Cells = sourceCells }
    let sw = Stopwatch.StartNew()
    pasteClipboard { Row = sourceRows + 10; Col = 0 } clip model |> ignore
    sw.Stop()
    sw.Elapsed.TotalMilliseconds
if pasteIdentityOversizedClipboardSkipsSourceTextElapsed < 500.0 then
    pass <- pass + 1
else
    fail <- fail + 1
    printfn
        "FAIL %-26s expected under %.1f ms got %.1f ms"
        "paste oversized skips source"
        500.0
        pasteIdentityOversizedClipboardSkipsSourceTextElapsed

let structuralModel =
    { createEmpty 100 100 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (S "a")
              { Row = 10; Col = 10 }, cell (S "b") ]
            |> Map.ofList
        Selection = SingleCell (0, 0) }
check "row default height identity" (obj.ReferenceEquals(setRowHeight 20 25.0 structuralModel, structuralModel)) true
check "col default width identity" (obj.ReferenceEquals(setColumnWidth 20 100.0 structuralModel, structuralModel)) true
check "row negative height identity" (obj.ReferenceEquals(setRowHeight -1 32.0 structuralModel, structuralModel)) true
check "col negative width identity" (obj.ReferenceEquals(setColumnWidth -1 140.0 structuralModel, structuralModel)) true
let sizedStructuralModel =
    structuralModel
    |> setRowHeight 20 32.0
    |> setColumnWidth 20 140.0
check "row same height identity" (obj.ReferenceEquals(setRowHeight 20 32.0 sizedStructuralModel, sizedStructuralModel)) true
check "col same width identity" (obj.ReferenceEquals(setColumnWidth 20 140.0 sizedStructuralModel, sizedStructuralModel)) true
let resetSizedStructuralModel =
    sizedStructuralModel
    |> setRowHeight 20 25.0
    |> setColumnWidth 20 100.0
check "row reset default value" (getRowHeight 20 resetSizedStructuralModel) 25.0
check "col reset default value" (getColumnWidth 20 resetSizedStructuralModel) 100.0
check "row reset removes override" (Map.containsKey 20 resetSizedStructuralModel.RowHeights) false
check "col reset removes override" (Map.containsKey 20 resetSizedStructuralModel.ColumnWidths) false
let explicitDefaultSizeModel =
    { structuralModel with
        RowHeights = Map.ofList [ 20, 25.0 ]
        ColumnWidths = Map.ofList [ 20, 100.0 ] }
let normalizedDefaultSizeModel =
    explicitDefaultSizeModel
    |> setRowHeight 20 25.0
    |> setColumnWidth 20 100.0
check "explicit row default removed" normalizedDefaultSizeModel.RowHeights.Count 0
check "explicit col default removed" normalizedDefaultSizeModel.ColumnWidths.Count 0
check "clear empty range identity" (obj.ReferenceEquals(clearSelectionCells (Range (50, 50, 60, 60)) structuralModel, structuralModel)) true
let clearedSelectionModel = clearSelectionCells (Range (0, 0, 5, 5)) structuralModel
check "clear selection removes hit" (getCellValue { Row = 0; Col = 0 } clearedSelectionModel) Empty
check "clear selection preserves miss" (getCellValue { Row = 10; Col = 10 } clearedSelectionModel) (S "b")
let formulaClearModel =
    createEmpty 10000 1000
    |> setCell { Row = 1; Col = 1 } { Value = Formula ("=1", CellValue.Number 1.0); Format = defaultFormat; Comment = None }
    |> setCell { Row = 9999; Col = 999 } { Value = S "outside"; Format = defaultFormat; Comment = None }
let formulaClearedModel = clearSelectionCells (Range (0, 0, 5000, 500)) formulaClearModel
check "clear sparse formula removed" (getCellValue { Row = 1; Col = 1 } formulaClearedModel) Empty
check "clear sparse outside preserved" (getCellValue { Row = 9999; Col = 999 } formulaClearedModel) (S "outside")
check "clear sparse formula cache" (cellsHaveFormulaCells formulaClearedModel.Cells) false
check "clear sparse repeat identity" (obj.ReferenceEquals(clearSelectionCells (Range (0, 0, 5000, 500)) formulaClearedModel, formulaClearedModel)) true
let allCellsClearedModel = clearSelectionCells (Range (0, 0, 9999, 999)) formulaClearModel
check "clear all populated cells" allCellsClearedModel.Cells.Count 0
check "clear all formula cache" (cellsHaveFormulaCells allCellsClearedModel.Cells) false
check "clear all repeat identity" (obj.ReferenceEquals(clearSelectionCells (Range (0, 0, 9999, 999)) allCellsClearedModel, allCellsClearedModel)) true
let clearDenseQuarterRows = 400
let clearDenseQuarterCols = 100
let clearDenseQuarterModel =
    { createEmpty (clearDenseQuarterRows + 10) (clearDenseQuarterCols + 10) with
        Cells =
            [ for row in 0 .. clearDenseQuarterRows - 1 do
                for col in 0 .. clearDenseQuarterCols - 1 ->
                    { Row = row; Col = col }, cell (CellValue.Number (float (row * clearDenseQuarterCols + col))) ]
            |> Map.ofList }
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let clearDenseQuarterBefore = System.GC.GetAllocatedBytesForCurrentThread()
let clearDenseQuarterResult = clearSelectionCells (Range (0, 0, 99, clearDenseQuarterCols - 1)) clearDenseQuarterModel
let clearDenseQuarterBytes = System.GC.GetAllocatedBytesForCurrentThread() - clearDenseQuarterBefore
check "clear dense quarter count" clearDenseQuarterResult.Cells.Count 30000
check "clear dense quarter allocation" (clearDenseQuarterBytes < 8000000L) true
let clearSparseWideRows = 60000
let clearSparseWideModel =
    { createEmpty (clearSparseWideRows + 10) 4 with
        Cells =
            [ for row in 0 .. clearSparseWideRows - 1 ->
                { Row = row; Col = 0 }, cell (CellValue.Number (float row)) ]
            |> Map.ofList }
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let clearSparseWideBefore = System.GC.GetAllocatedBytesForCurrentThread()
let clearSparseWideResult = clearSelectionCells (Range (0, 0, (clearSparseWideRows / 2) - 1, 2)) clearSparseWideModel
let clearSparseWideBytes = System.GC.GetAllocatedBytesForCurrentThread() - clearSparseWideBefore
check "clear sparse wide count" clearSparseWideResult.Cells.Count (clearSparseWideRows / 2)
check "clear sparse wide allocation" (clearSparseWideBytes < 20000000L) true
let clearMostlyAllCount = 100000
let clearMostlyAllModel =
    { createEmpty (clearMostlyAllCount + 20) 1 with
        Cells =
            seq {
                for row in 0 .. clearMostlyAllCount - 2 do
                    yield { Row = row; Col = 0 }, cell (CellValue.Number (float row))
                yield { Row = clearMostlyAllCount + 10; Col = 0 }, cell (S "outside")
            }
            |> Map.ofSeq }
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let clearMostlyAllBefore = System.GC.GetAllocatedBytesForCurrentThread()
let clearMostlyAllSw = Stopwatch.StartNew()
let clearMostlyAllResult = clearSelectionCells (Range (0, 0, clearMostlyAllCount - 2, 0)) clearMostlyAllModel
clearMostlyAllSw.Stop()
let clearMostlyAllBytes = System.GC.GetAllocatedBytesForCurrentThread() - clearMostlyAllBefore
check "clear mostly all count" clearMostlyAllResult.Cells.Count 1
check "clear mostly all outside" (getCellValue { Row = clearMostlyAllCount + 10; Col = 0 } clearMostlyAllResult) (S "outside")
if clearMostlyAllSw.Elapsed.TotalMilliseconds < 60.0 && clearMostlyAllBytes < 1000000L then
    pass <- pass + 1
else
    fail <- fail + 1
    printfn
        "FAIL %-26s expected under %.1f ms and %d bytes got %.1f ms and %d bytes"
        "clear mostly all perf"
        60.0
        1000000L
        clearMostlyAllSw.Elapsed.TotalMilliseconds
        clearMostlyAllBytes
cellsHaveFormulaCells clearMostlyAllModel.Cells |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let clearMostlyAllCachedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let clearMostlyAllCachedSw = Stopwatch.StartNew()
let clearMostlyAllCachedResult = clearSelectionCells (Range (0, 0, clearMostlyAllCount - 2, 0)) clearMostlyAllModel
clearMostlyAllCachedSw.Stop()
let clearMostlyAllCachedBytes = System.GC.GetAllocatedBytesForCurrentThread() - clearMostlyAllCachedBefore
check "clear mostly all cached count" clearMostlyAllCachedResult.Cells.Count 1
check "clear mostly all cached formula" (cellsHaveFormulaCells clearMostlyAllCachedResult.Cells) false
if clearMostlyAllCachedSw.Elapsed.TotalMilliseconds < 60.0 && clearMostlyAllCachedBytes < 1000000L then
    pass <- pass + 1
else
    fail <- fail + 1
    printfn
        "FAIL %-26s expected under %.1f ms and %d bytes got %.1f ms and %d bytes"
        "clear mostly cached perf"
        60.0
        1000000L
        clearMostlyAllCachedSw.Elapsed.TotalMilliseconds
        clearMostlyAllCachedBytes
let insertedRowAfterUsed = insertRow 50 structuralModel
check "insert row after used same cells" (obj.ReferenceEquals(insertedRowAfterUsed.Cells, structuralModel.Cells)) true
check "insert row after used count" insertedRowAfterUsed.RowCount 101
let deletedRowAfterUsed = deleteRow 50 structuralModel
check "delete row after used same cells" (obj.ReferenceEquals(deletedRowAfterUsed.Cells, structuralModel.Cells)) true
check "delete row after used count" deletedRowAfterUsed.RowCount 99
let insertedColAfterUsed = insertColumn 50 structuralModel
check "insert col after used same cells" (obj.ReferenceEquals(insertedColAfterUsed.Cells, structuralModel.Cells)) true
check "insert col after used count" insertedColAfterUsed.ColCount 101
let deletedColAfterUsed = deleteColumn 50 structuralModel
check "delete col after used same cells" (obj.ReferenceEquals(deletedColAfterUsed.Cells, structuralModel.Cells)) true
check "delete col after used count" deletedColAfterUsed.ColCount 99
check "delete row negative identity" (obj.ReferenceEquals(deleteRow -1 structuralModel, structuralModel)) true
check "delete row past end identity" (obj.ReferenceEquals(deleteRow structuralModel.RowCount structuralModel, structuralModel)) true
check "delete col negative identity" (obj.ReferenceEquals(deleteColumn -1 structuralModel, structuralModel)) true
check "delete col past end identity" (obj.ReferenceEquals(deleteColumn structuralModel.ColCount structuralModel, structuralModel)) true
check "insert row negative identity" (obj.ReferenceEquals(insertRow -1 structuralModel, structuralModel)) true
check "insert row past end identity" (obj.ReferenceEquals(insertRow (structuralModel.RowCount + 1) structuralModel, structuralModel)) true
check "insert col negative identity" (obj.ReferenceEquals(insertColumn -1 structuralModel, structuralModel)) true
check "insert col past end identity" (obj.ReferenceEquals(insertColumn (structuralModel.ColCount + 1) structuralModel, structuralModel)) true
let structuralModelWide = structuralModel |> setCell { Row = 0; Col = 80 } (cell (S "wide"))
let insertedColAfterWide = insertColumn 50 structuralModelWide
check "insert col cache invalidates" (getCellValue { Row = 0; Col = 81 } insertedColAfterWide) (S "wide")
let deletedColAfterWide = deleteColumn 50 structuralModelWide
check "delete col cache invalidates" (getCellValue { Row = 0; Col = 79 } deletedColAfterWide) (S "wide")
let insertedColAfterUsedAgain = insertColumn 50 structuralModel
check "insert col cache keeps first source" (obj.ReferenceEquals(insertedColAfterUsedAgain.Cells, structuralModel.Cells)) true
let sparseTailRowModel =
    { createEmpty 120 4 with
        Cells =
            ([ for row in 0 .. 7 -> { Row = row; Col = 0 }, cell (S (sprintf "head%d" row)) ]
             @ [ { Row = 90; Col = 0 }, cell (S "tail-a")
                 { Row = 95; Col = 0 }, cell (S "tail-b") ])
            |> Map.ofList }
let sparseTailRowInserted = insertRow 90 sparseTailRowModel
check "insert row sparse tail preserve head" (getCellValue { Row = 7; Col = 0 } sparseTailRowInserted) (S "head7")
check "insert row sparse tail shift first" (getCellValue { Row = 91; Col = 0 } sparseTailRowInserted) (S "tail-a")
check "insert row sparse tail shift last" (getCellValue { Row = 96; Col = 0 } sparseTailRowInserted) (S "tail-b")
let sparseTailRowDeleted = deleteRow 90 sparseTailRowModel
check "delete row sparse tail removes row" (getCellValue { Row = 90; Col = 0 } sparseTailRowDeleted) Empty
check "delete row sparse tail shifts last" (getCellValue { Row = 94; Col = 0 } sparseTailRowDeleted) (S "tail-b")
let sparseTailColModel =
    { createEmpty 4 120 with
        Cells =
            ([ for col in 0 .. 7 -> { Row = 0; Col = col }, cell (S (sprintf "head%d" col)) ]
             @ [ { Row = 0; Col = 90 }, cell (S "tail-a")
                 { Row = 0; Col = 95 }, cell (S "tail-b") ])
            |> Map.ofList }
let sparseTailColInserted = insertColumn 90 sparseTailColModel
check "insert col sparse tail preserve head" (getCellValue { Row = 0; Col = 7 } sparseTailColInserted) (S "head7")
check "insert col sparse tail shift first" (getCellValue { Row = 0; Col = 91 } sparseTailColInserted) (S "tail-a")
check "insert col sparse tail shift last" (getCellValue { Row = 0; Col = 96 } sparseTailColInserted) (S "tail-b")
let sparseTailColDeleted = deleteColumn 90 sparseTailColModel
check "delete col sparse tail removes col" (getCellValue { Row = 0; Col = 90 } sparseTailColDeleted) Empty
check "delete col sparse tail shifts last" (getCellValue { Row = 0; Col = 94 } sparseTailColDeleted) (S "tail-b")
check "used bounds empty" (usedBounds (createEmpty 10 10)) None
check "used bounds cells" (usedBounds structuralModel) (Some (10, 10))
check "last used row existing" (lastUsedColumnInRow 10 structuralModel) (Some 10)
check "last used row cached repeat" (lastUsedColumnInRow 10 structuralModel) (Some 10)
check "last used row missing" (lastUsedColumnInRow 9 structuralModel) None
check "last used row keeps first row" (lastUsedColumnInRow 10 structuralModel) (Some 10)
check "last used row missing cached repeat" (lastUsedColumnInRow 9 structuralModel) None
let structuralModelEdited = structuralModel |> setCell { Row = 10; Col = 12 } (cell (CellValue.Number 2.0))
check "last used row edit invalidates" (lastUsedColumnInRow 10 structuralModelEdited) (Some 12)
let mergedLastUsedModel =
    { structuralModel with
        MergedCells =
            [ { StartRow = 20; StartCol = 2; EndRow = 22; EndCol = 8 } ] }
check "last used row merge" (lastUsedColumnInRow 21 mergedLastUsedModel) (Some 8)
let mergedLastUsedChanged =
    { mergedLastUsedModel with
        MergedCells =
            [ { StartRow = 20; StartCol = 2; EndRow = 22; EndCol = 9 } ] }
check "last used row merge invalidates" (lastUsedColumnInRow 21 mergedLastUsedChanged) (Some 9)
check "last used row cache keeps first source" (lastUsedColumnInRow 10 structuralModel) (Some 10)
let lastUsedMergeRetainedModelCount = 80
let lastUsedMergeRetainedMergeCount = 10000
let lastUsedMergeRetainedMerges modelIndex =
    [ for index in 0 .. lastUsedMergeRetainedMergeCount - 1 ->
        { StartRow = index
          StartCol = modelIndex
          EndRow = index
          EndCol = modelIndex + 16 } ]
let lastUsedMergeRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. lastUsedMergeRetainedModelCount - 1 do
        let model =
            { createEmpty (lastUsedMergeRetainedMergeCount + 1) 32 with
                MergedCells = lastUsedMergeRetainedMerges modelIndex }
        lastUsedColumnInRow (lastUsedMergeRetainedMergeCount / 2) model |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "last used merge cache retained memory" (lastUsedMergeRetainedBytes < 20000000L) true
let mergedOnlyBoundsModel =
    { createEmpty 30 12 with
        MergedCells =
            [ { StartRow = 20; StartCol = 2; EndRow = 22; EndCol = 8 } ] }
check "duplicate merge identity" (obj.ReferenceEquals(addMergedCell 20 2 22 8 mergedOnlyBoundsModel, mergedOnlyBoundsModel)) true
check "missing unmerge identity" (obj.ReferenceEquals(removeMergedCell { Row = 99; Col = 99 } mergedOnlyBoundsModel, mergedOnlyBoundsModel)) true
let unmergeHeadTail =
    [ for index in 0 .. 9999 ->
        { StartRow = index + 1
          StartCol = 0
          EndRow = index + 1
          EndCol = 2 } ]
let unmergeHead =
    { StartRow = 0
      StartCol = 0
      EndRow = 0
      EndCol = 2 }
let unmergeHeadModel =
    { createEmpty 10002 4 with MergedCells = unmergeHead :: unmergeHeadTail }
let unmergeHeadBytes, unmergeHeadResult =
    let before = System.GC.GetAllocatedBytesForCurrentThread()
    let result = removeMergedCell { Row = 0; Col = 1 } unmergeHeadModel
    System.GC.GetAllocatedBytesForCurrentThread() - before, result
check "unmerge head reuses tail" (obj.ReferenceEquals(unmergeHeadResult.MergedCells, unmergeHeadTail)) true
check "unmerge head allocation" (unmergeHeadBytes < 100000L) true
let unmergeMiddlePrefixCount = 20000
let unmergeMiddleTailCount = 20000
let unmergeMiddlePrefix =
    [ for index in 0 .. unmergeMiddlePrefixCount - 1 ->
        { StartRow = index
          StartCol = 0
          EndRow = index
          EndCol = 2 } ]
let unmergeMiddleTarget =
    { StartRow = unmergeMiddlePrefixCount
      StartCol = 0
      EndRow = unmergeMiddlePrefixCount
      EndCol = 2 }
let unmergeMiddleTail =
    [ for index in unmergeMiddlePrefixCount + 1 .. unmergeMiddlePrefixCount + unmergeMiddleTailCount ->
        { StartRow = index
          StartCol = 0
          EndRow = index
          EndCol = 2 } ]
let unmergeMiddleModel =
    { createEmpty (unmergeMiddlePrefixCount + unmergeMiddleTailCount + 2) 4 with
        MergedCells = unmergeMiddlePrefix @ (unmergeMiddleTarget :: unmergeMiddleTail) }
let unmergeMiddleBytes, unmergeMiddleResult =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetAllocatedBytesForCurrentThread()
    let result = removeMergedCell { Row = unmergeMiddlePrefixCount; Col = 1 } unmergeMiddleModel
    System.GC.GetAllocatedBytesForCurrentThread() - before, result
check "unmerge middle shares tail"
    (obj.ReferenceEquals(List.skip unmergeMiddlePrefixCount unmergeMiddleResult.MergedCells, unmergeMiddleTail))
    true
check "unmerge middle allocation" (unmergeMiddleBytes < 1200000L) true
let unmergeTailBytes, unmergeTailResult =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetAllocatedBytesForCurrentThread()
    let result = removeMergedCell { Row = unmergeMiddlePrefixCount + unmergeMiddleTailCount; Col = 1 } unmergeMiddleModel
    System.GC.GetAllocatedBytesForCurrentThread() - before, result
check "unmerge tail count" unmergeTailResult.MergedCells.Length (unmergeMiddlePrefixCount + unmergeMiddleTailCount)
check "unmerge tail allocation" (unmergeTailBytes < 2000000L) true
check "used bounds merge only" (usedBounds mergedOnlyBoundsModel) (Some (22, 8))
check "used bounds cache keeps first source" (usedBounds structuralModel) (Some (10, 10))
let usedBoundsMergeRetainedModelCount = 80
let usedBoundsMergeRetainedMergeCount = 10000
let usedBoundsMergeRetainedMerges modelIndex =
    [ for index in 0 .. usedBoundsMergeRetainedMergeCount - 1 ->
        { StartRow = index
          StartCol = modelIndex
          EndRow = index
          EndCol = modelIndex + 16 } ]
let usedBoundsMergeRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. usedBoundsMergeRetainedModelCount - 1 do
        let model =
            { createEmpty (usedBoundsMergeRetainedMergeCount + 1) 32 with
                MergedCells = usedBoundsMergeRetainedMerges modelIndex }
        usedBounds model |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "used bounds merge cache retained memory" (usedBoundsMergeRetainedBytes < 20000000L) true
check "find merged hit" (findMergedCell { Row = 21; Col = 5 } mergedOnlyBoundsModel) (Some (List.head mergedOnlyBoundsModel.MergedCells))
check "find merged miss" (findMergedCell { Row = 21; Col = 9 } mergedOnlyBoundsModel) None
let overlapFirst = { StartRow = 1; StartCol = 1; EndRow = 3; EndCol = 3 }
let overlapSecond = { StartRow = 2; StartCol = 2; EndRow = 4; EndCol = 4 }
let overlapModel = { createEmpty 10 10 with MergedCells = [ overlapFirst; overlapSecond ] }
check "find merged list order" (findMergedCell { Row = 2; Col = 2 } overlapModel) (Some overlapFirst)
check "visible merges order" (mergedCellsIntersectingRange 2 3 2 3 overlapModel) [ overlapFirst; overlapSecond ]
let changedMerge = { StartRow = 7; StartCol = 7; EndRow = 8; EndCol = 8 }
let changedMergeModel = { overlapModel with MergedCells = [ changedMerge ] }
check "find merged cache invalidates old" (findMergedCell { Row = 2; Col = 2 } changedMergeModel) None
check "find merged cache invalidates new" (findMergedCell { Row = 7; Col = 7 } changedMergeModel) (Some changedMerge)
check "find merged cache keeps first source" (findMergedCell { Row = 2; Col = 2 } overlapModel) (Some overlapFirst)
let unsortedVisibleFirst = { StartRow = 8; StartCol = 1; EndRow = 8; EndCol = 1 }
let unsortedVisibleSecond = { StartRow = 1; StartCol = 1; EndRow = 2; EndCol = 2 }
let unsortedMergeModel = { createEmpty 12 12 with MergedCells = [ unsortedVisibleFirst; unsortedVisibleSecond ] }
check "visible merges unsorted order" (mergedCellsIntersectingRange 1 9 1 3 unsortedMergeModel) [ unsortedVisibleFirst; unsortedVisibleSecond ]
check "visible merges cache invalidates" (mergedCellsIntersectingRange 7 9 7 9 changedMergeModel) [ changedMerge ]
check "visible merges cache keeps first source" (mergedCellsIntersectingRange 1 9 1 3 unsortedMergeModel) [ unsortedVisibleFirst; unsortedVisibleSecond ]
let oversizedMergeRangeCacheSkipped () =
    let mergeCount = 200001
    let model =
        { createEmpty (mergeCount + 1) 2 with
            MergedCells =
                [ for row in 0 .. mergeCount - 1 ->
                    { StartRow = row
                      StartCol = 0
                      EndRow = row
                      EndCol = 1 } ] }
    let mergesA = mergedCellsIntersectingRange 0 mergeCount 0 2 model
    let mergesB = mergedCellsIntersectingRange 0 mergeCount 0 2 model
    mergesA.Length, obj.ReferenceEquals(mergesA, mergesB)
let oversizedMergeRangeCount, oversizedMergeRangeReused = oversizedMergeRangeCacheSkipped ()
check "merge range cache skips oversized result"
    (oversizedMergeRangeCount > 200000 && not oversizedMergeRangeReused)
    true
let mergeCacheModelCount = 24
let mergeCacheMergeCount = 1000
let mergeCacheMerges modelIndex =
    [ for index in mergeCacheMergeCount - 1 .. -1 .. 0 ->
        let row = index * 2
        { StartRow = row
          StartCol = modelIndex % 5
          EndRow = row
          EndCol = (modelIndex % 5) + 2 } ]
let mergeCacheModels =
    [| for modelIndex in 0 .. mergeCacheModelCount - 1 ->
        { createEmpty (mergeCacheMergeCount * 2 + 5) 16 with
            MergedCells = mergeCacheMerges modelIndex } |]
let mergeCacheRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in mergeCacheModels do
            total <- total + (mergedCellsIntersectingRange 0 (mergeCacheMergeCount * 2) 0 16 model).Length
            match findMergedCell { Row = mergeCacheMergeCount; Col = 2 } model with
            | Some _ -> total <- total + 1
            | None -> ()
    total
mergeCacheRun 3 |> ignore
let mergeCacheSw = Stopwatch.StartNew()
let mergeCacheTotal = mergeCacheRun 60
mergeCacheSw.Stop()
check "merge cache many models"
    (mergeCacheSw.Elapsed.TotalMilliseconds < 25.0 && mergeCacheTotal > 0)
    true
let mergeCacheRetainedModelCount = 80
let mergeCacheRetainedMergeCount = 20000
let mergeCacheRetainedMerges modelIndex =
    [ for index in mergeCacheRetainedMergeCount - 1 .. -1 .. 0 ->
        let row = index * 2
        { StartRow = row
          StartCol = modelIndex % 5
          EndRow = row
          EndCol = (modelIndex % 5) + 2 } ]
let mergeCacheRetainedModel modelIndex =
    { createEmpty (mergeCacheRetainedMergeCount * 2 + 10) 16 with
        MergedCells = mergeCacheRetainedMerges modelIndex }
let mergeCacheRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. mergeCacheRetainedModelCount - 1 do
        let model = mergeCacheRetainedModel modelIndex
        findMergedCell { Row = mergeCacheRetainedMergeCount; Col = 2 } model |> ignore
        mergedCellsIntersectingRange 0 1 15 16 model |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "merge cache bounded source memory" (mergeCacheRetainedBytes < 160000000L) true
let mergeCacheOversizedSourceMergesWeakReference () =
    let mergeCount = 100001
    let merges =
        [ for index in mergeCount - 1 .. -1 .. 0 ->
            { StartRow = index
              StartCol = 0
              EndRow = index
              EndCol = 1 } ]
    let model = { createEmpty (mergeCount + 1) 4 with MergedCells = merges }
    findMergedCell { Row = 0; Col = 0 } model |> ignore
    mergedCellsIntersectingRange 0 1 0 2 model |> ignore
    System.WeakReference(merges)
let mergeCacheOversizedSourceMergesWeak = mergeCacheOversizedSourceMergesWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "merge cache skips oversized source"
    mergeCacheOversizedSourceMergesWeak.IsAlive
    false
let oversizedUnsortedMergeRangeAllocatedBytes () =
    let mergeCount = 120001
    let merges =
        [ for index in mergeCount - 1 .. -1 .. 0 ->
            { StartRow = index
              StartCol = 0
              EndRow = index
              EndCol = 1 } ]
    let model = { createEmpty (mergeCount + 1) 4 with MergedCells = merges }
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetAllocatedBytesForCurrentThread()
    let result = mergedCellsIntersectingRange 0 1 0 2 model
    let allocated = System.GC.GetAllocatedBytesForCurrentThread() - before
    allocated, result
let oversizedUnsortedMergeAllocated, oversizedUnsortedMergeResult =
    oversizedUnsortedMergeRangeAllocatedBytes ()
check "merge range oversized unsorted result"
    oversizedUnsortedMergeResult
    [ { StartRow = 0; StartCol = 0; EndRow = 0; EndCol = 1 } ]
check "merge range oversized unsorted avoids index"
    (oversizedUnsortedMergeAllocated < 1000000L)
    true
let noopCommand = { Undo = id; Redo = id; Description = "noop"; RetainedItemCount = 1; RetainedTextCharCount = 1 }
let noHistoryModel =
    { structuralModel with History = { Past = []; Future = []; MaxSize = 0 } }
check "push history disabled identity" (obj.ReferenceEquals(pushCommand noopCommand noHistoryModel, noHistoryModel)) true
let noHistoryWithFuture =
    { structuralModel with History = { Past = []; Future = [ noopCommand ]; MaxSize = 0 } }
let noHistoryCleared = pushCommand noopCommand noHistoryWithFuture
check "push history disabled clears future" (noHistoryCleared.History.Future.IsEmpty) true
check "push history disabled clears past" (noHistoryCleared.History.Past.IsEmpty) true
let oversizedHistoryRetainedItemCount = commandRetainedSumItemCounts [ 600000 ]
let mutable oversizedHistoryTextCountEvaluated = false
let oversizedHistoryRetainedTextCount =
    commandRetainedTextCharCountForItemCount oversizedHistoryRetainedItemCount (fun () ->
        oversizedHistoryTextCountEvaluated <- true
        1)
check "history text count skips oversized items"
    (oversizedHistoryRetainedTextCount > 1 && not oversizedHistoryTextCountEvaluated)
    true
let oversizedHistoryInitialTextCount = commandRetainedTextLength (System.String('H', 21 * 1024 * 1024))
let mutable oversizedHistoryInitialTextEvaluated = false
let oversizedHistoryInitialRetainedTextCount =
    commandRetainedTextCharCountForItemCountWithInitial 1 oversizedHistoryInitialTextCount (fun () ->
        oversizedHistoryInitialTextEvaluated <- true
        1)
check "history text count skips oversized initial text"
    (oversizedHistoryInitialRetainedTextCount > 1 && not oversizedHistoryInitialTextEvaluated)
    true
let historyLargeCellMapWeakReference () =
    let mutable model = { createEmpty 1 1 with History = { Past = []; Future = []; MaxSize = 50 } }
    let mutable firstCellsWeak = Unchecked.defaultof<System.WeakReference>
    for commandIndex in 0 .. 5 do
        let cells =
            [ for row in 0 .. 119999 ->
                { Row = row; Col = 0 }, cell (Number (float (row + commandIndex))) ]
            |> Map.ofList
        if commandIndex = 0 then
            firstCellsWeak <- System.WeakReference(cells)
        let command = {
            Undo = fun m -> { m with Cells = cells }
            Redo = id
            Description = "large cell map"
            RetainedItemCount = commandRetainedCellMapItemCount cells
            RetainedTextCharCount = commandRetainedCellMapTextCharCount cells
        }
        model <- pushCommand command { model with Cells = Map.empty }
    firstCellsWeak, model.History.Past.Length
let historyLargeCellMapWeak, historyLargeCellMapRetainedCount =
    historyLargeCellMapWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "push history trims large commands" (historyLargeCellMapRetainedCount < 6) true
check "push history releases trimmed cells" historyLargeCellMapWeak.IsAlive false
let historyCellMapTextMetricCacheElapsed =
    let cells =
        [ for row in 0 .. 119999 ->
            { Row = row; Col = 0 }, cell (Number (float row)) ]
        |> Map.ofList
    let model = { createEmpty 120000 4 with Cells = cells }
    commandRetainedCellMapTextCharCount cells |> ignore
    let updated =
        setCell
            { Row = 1; Col = 1 }
            (cell (Number 1.0))
            model
    let sw = System.Diagnostics.Stopwatch.StartNew()
    let retainedTextCharCount = commandRetainedCellMapTextCharCount updated.Cells
    sw.Stop()
    check "history cell map text metric value" retainedTextCharCount 0
    sw.Elapsed.TotalMilliseconds
check "history cell map text metric cache"
    (historyCellMapTextMetricCacheElapsed < 30.0)
    true

printfn ""
printfn "Stage 16 filter tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
