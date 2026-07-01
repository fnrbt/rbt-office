#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// Structured references (rich grammar) + SUBTOTAL / AGGREGATE.

open RedBat.Spreadsheet.Model
module L = RedBat.Spreadsheet.Layout
module P = RedBat.Spreadsheet.Paint
module F = RedBat.Spreadsheet.Formula
module WB = RedBat.Spreadsheet.Workbook

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let S = CellValue.String
let N = CellValue.Number

// Sheet layout (0-indexed):
//   row0 header:  Region | Amount | Sales Amount
//   row1..4 data: East/10/100, West/20/200, East/30/300, North/40/400
//   row5 totals:  "Total"| =SUBTOTAL(9,B2:B5)->100 | (blank)
let baseCells =
    [ { Row=0; Col=0 }, cell (S "Region");   { Row=0; Col=1 }, cell (S "Amount");   { Row=0; Col=2 }, cell (S "Sales Amount")
      { Row=1; Col=0 }, cell (S "East");      { Row=1; Col=1 }, cell (N 10.);        { Row=1; Col=2 }, cell (N 100.)
      { Row=2; Col=0 }, cell (S "West");      { Row=2; Col=1 }, cell (N 20.);        { Row=2; Col=2 }, cell (N 200.)
      { Row=3; Col=0 }, cell (S "East");      { Row=3; Col=1 }, cell (N 30.);        { Row=3; Col=2 }, cell (N 300.)
      { Row=4; Col=0 }, cell (S "North");     { Row=4; Col=1 }, cell (N 40.);        { Row=4; Col=2 }, cell (N 400.)
      { Row=5; Col=0 }, cell (S "Total")
      // totals cell is itself a SUBTOTAL formula (cached 100) — must be skipped
      // by an outer SUBTOTAL ranging over it.
      { Row=5; Col=1 }, cell (CellValue.Formula ("SUBTOTAL(9,B2:B5)", N 100.)) ]
    |> Map.ofList

let model = { createEmpty 20 8 with Cells = baseCells }

let table : WB.Table =
    { Name = "Sales"; SheetId = "sheet1"
      StartRow = 0; StartCol = 0; EndRow = 5; EndCol = 2
      Columns = [ "Region"; "Amount"; "Sales Amount" ]
      HasHeaders = true; HasTotals = true }

let wb : WB.Workbook =
    { Sheets = Map.ofList [ "sheet1", ({ Id="sheet1"; Name="Sheet1"; Model=model; Visible=true }: WB.Sheet) ]
      SheetOrder = [ "sheet1" ]; ActiveSheet = "sheet1"
      Tables = Map.ofList [ "Sales", table ]; Names = Map.empty }

let ctxAt (m: SpreadsheetModel) (cur: CellAddress option) : F.EvalContext =
    { Model = m; Workbook = Some wb; Sheet = Some "sheet1"; CriteriaCache = None; CurrentCell = cur }

let evalIn (m: SpreadsheetModel) (cur: CellAddress option) (s: string) : CellValue =
    match F.parseFormula s with
    | Ok e -> F.eval e (ctxAt m cur)
    | Result.Error msg -> CellValue.Error msg

let evalS s = evalIn model None s
let num = function CellValue.Number n -> n | v -> nan
let str = function CellValue.String s -> s | CellValue.Number n -> string n | v -> sprintf "%A" v

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-30s expected %A got %A" label expected got

let checkAllocUnder label (got: int64) (limit: int64) =
    if got < limit then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-30s expected < %d bytes got %d" label limit got

let allocatedBytesFor (f: unit -> 'a) =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetAllocatedBytesForCurrentThread()
    let result = f()
    System.GC.KeepAlive result
    System.GC.GetAllocatedBytesForCurrentThread() - before

check "same active sheet identity" (obj.ReferenceEquals(WB.setActiveSheet "sheet1" wb, wb)) true
check "same sheet name identity" (obj.ReferenceEquals(WB.renameSheet "sheet1" "Sheet1" wb, wb)) true
check "same sheet order identity" (obj.ReferenceEquals(WB.reorderSheets [ "sheet1" ] wb, wb)) true
check "same table definition identity" (obj.ReferenceEquals(WB.defineTable table wb, wb)) true
check "delete missing table identity" (obj.ReferenceEquals(WB.deleteTable "Missing" wb, wb)) true
let missingSheetWorkbook =
    { wb with
        Sheets = wb.Sheets |> Map.add "sheet2" { Id = "sheet2"; Name = "Sheet2"; Model = model; Visible = true }
        SheetOrder = [ "sheet1"; "sheet2" ] }
check "delete missing sheet identity"
    (match WB.deleteSheet "missing" missingSheetWorkbook with
     | Some next -> obj.ReferenceEquals(next, missingSheetWorkbook)
     | None -> false)
    true
check "same sheet model identity" (obj.ReferenceEquals(WB.updateSheetModel "sheet1" model wb, wb)) true

let largeSheetCount = 50000
let largeSheetModel = createEmpty 1 1
let largeSheetIds = [ for i in 1 .. largeSheetCount -> sprintf "sheet%d" i ]
let largeWorkbook: WB.Workbook =
    let sheets =
        largeSheetIds
        |> List.map (fun id -> id, ({ Id = id; Name = id; Model = largeSheetModel; Visible = true }: WB.Sheet))
        |> Map.ofList
    { Sheets = sheets
      SheetOrder = largeSheetIds
      ActiveSheet = "sheet1"
      Tables = Map.empty
      Names = Map.empty }
let largeWorkbookTables =
    [ for i in 1 .. largeSheetCount ->
        let name = sprintf "Table%d" i
        name, ({ table with Name = name }: WB.Table) ]
    |> Map.ofList
let largeTableWorkbook = { largeWorkbook with Tables = largeWorkbookTables }

checkAllocUnder "addSheet large allocation" (allocatedBytesFor (fun () -> WB.addSheet "Next" largeWorkbook)) 2400000L
checkAllocUnder "deleteSheet large allocation" (allocatedBytesFor (fun () -> WB.deleteSheet "sheet25000" largeWorkbook)) 2400000L
checkAllocUnder "deleteSheet missing allocation" (allocatedBytesFor (fun () -> WB.deleteSheet "missing" largeWorkbook)) 100000L
checkAllocUnder "getSheetNames large allocation" (allocatedBytesFor (fun () -> WB.getSheetNames largeWorkbook)) 5200000L
let largeSheetNames = WB.getSheetNames largeWorkbook
check "getSheetNames cached identity" (obj.ReferenceEquals(WB.getSheetNames largeWorkbook, largeSheetNames)) true
let cachedSheetNamesAllocated =
    allocatedBytesFor (fun () ->
        for _ in 1 .. 200 do
            WB.getSheetNames largeWorkbook |> ignore)
checkAllocUnder "getSheetNames cached allocation" cachedSheetNamesAllocated 100000L
checkAllocUnder "getTableNames large allocation" (allocatedBytesFor (fun () -> WB.getTableNames largeTableWorkbook)) 2400000L
let largeTableNames = WB.getTableNames largeTableWorkbook
check "getTableNames cached identity" (obj.ReferenceEquals(WB.getTableNames largeTableWorkbook, largeTableNames)) true
let cachedTableNamesAllocated =
    allocatedBytesFor (fun () ->
        for _ in 1 .. 200 do
            WB.getTableNames largeTableWorkbook |> ignore)
checkAllocUnder "getTableNames cached allocation" cachedTableNamesAllocated 100000L

let workbookNameCacheWorkbookCount = 24
let workbookNameCacheNameCount = 2000
let workbookNameCacheWorkbooks =
    [| for workbookIndex in 0 .. workbookNameCacheWorkbookCount - 1 ->
        let sheetIds = [ for index in 0 .. workbookNameCacheNameCount - 1 -> sprintf "cache-sheet-%02d-%04d" workbookIndex index ]
        let sheets =
            sheetIds
            |> List.map (fun id -> id, ({ Id = id; Name = id; Model = largeSheetModel; Visible = true }: WB.Sheet))
            |> Map.ofList
        let tables =
            [ for index in 0 .. workbookNameCacheNameCount - 1 ->
                let name = sprintf "CacheTable_%02d_%04d" workbookIndex index
                name, ({ table with Name = name; SheetId = sheetIds.[index % sheetIds.Length] }: WB.Table) ]
            |> Map.ofList
        ({ Sheets = sheets
           SheetOrder = sheetIds
           ActiveSheet = sheetIds.[0]
           Tables = tables
           Names = Map.empty }: WB.Workbook) |]
let workbookNameCacheRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for workbook in workbookNameCacheWorkbooks do
            total <- total + (WB.getSheetNames workbook).Length
            total <- total + (WB.getTableNames workbook).Length
    total
workbookNameCacheRun 6 |> ignore
let workbookNameCacheSw = System.Diagnostics.Stopwatch.StartNew()
let workbookNameCacheTotal = workbookNameCacheRun 30
workbookNameCacheSw.Stop()
check "workbook names cache many workbooks"
    (workbookNameCacheSw.Elapsed.TotalMilliseconds < 80.0 && workbookNameCacheTotal > 0)
    true

let workbookNameSourceCycleWorkbookCount = 80
let workbookNameSourceCycleNameCount = 200
let workbookNameSourceCycleWorkbooks =
    [| for workbookIndex in 0 .. workbookNameSourceCycleWorkbookCount - 1 ->
        let sheetIds =
            [ for index in 0 .. workbookNameSourceCycleNameCount - 1 ->
                sprintf "source-sheet-%02d-%04d" workbookIndex index ]
        let sheets =
            sheetIds
            |> List.map (fun id ->
                id,
                ({ Id = id
                   Name = id
                   Model = largeSheetModel
                   Visible = true }: WB.Sheet))
            |> Map.ofList
        let tables =
            [ for index in 0 .. workbookNameSourceCycleNameCount - 1 ->
                let name = sprintf "SourceTable_%02d_%04d" workbookIndex index
                name, ({ table with Name = name; SheetId = sheetIds.[index % sheetIds.Length] }: WB.Table) ]
            |> Map.ofList
        ({ Sheets = sheets
           SheetOrder = sheetIds
           ActiveSheet = sheetIds.[0]
           Tables = tables
           Names = Map.empty }: WB.Workbook) |]
for workbook in workbookNameSourceCycleWorkbooks do
    WB.getSheetNames workbook |> ignore
    WB.getTableNames workbook |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let workbookNameSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let workbookNameSourceCycleSw = System.Diagnostics.Stopwatch.StartNew()
let mutable workbookNameSourceCycleTotal = 0
for _pass in 1 .. 20 do
    for workbook in workbookNameSourceCycleWorkbooks do
        workbookNameSourceCycleTotal <- workbookNameSourceCycleTotal + (WB.getSheetNames workbook).Length
        workbookNameSourceCycleTotal <- workbookNameSourceCycleTotal + (WB.getTableNames workbook).Length
workbookNameSourceCycleSw.Stop()
let workbookNameSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - workbookNameSourceCycleAllocatedBefore
check "workbook names cache source cycle"
    (workbookNameSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && workbookNameSourceCycleAllocated < 512L * 1024L
     && workbookNameSourceCycleTotal = 640000)
    true
let workbookNameRetainedWorkbookCount = 80
let workbookNameRetainedNameCount = 500
let workbookNameRetainedLongName prefix workbookIndex nameIndex =
    sprintf
        "%s_%03d_%04d_%s"
        prefix
        workbookIndex
        nameIndex
        (System.String('x', 1024))
let workbookNameRetainedSheetWorkbook workbookIndex : WB.Workbook =
    let sheetIds =
        [ for index in 0 .. workbookNameRetainedNameCount - 1 ->
            sprintf "retained-sheet-%03d-%04d" workbookIndex index ]
    let sheets =
        sheetIds
        |> List.mapi (fun index id ->
            id,
            ({ Id = id
               Name = workbookNameRetainedLongName "SheetName" workbookIndex index
               Model = largeSheetModel
               Visible = true }: WB.Sheet))
        |> Map.ofList
    { Sheets = sheets
      SheetOrder = sheetIds
      ActiveSheet = List.head sheetIds
      Tables = Map.empty
      Names = Map.empty }
let workbookNameRetainedTableWorkbook workbookIndex : WB.Workbook =
    let tables =
        [ for index in 0 .. workbookNameRetainedNameCount - 1 ->
            let name = workbookNameRetainedLongName "TableName" workbookIndex index
            name, ({ table with Name = name }: WB.Table) ]
        |> Map.ofList
    { largeWorkbook with
        Sheets =
            Map.ofList
                [ "sheet1",
                  ({ Id = "sheet1"
                     Name = "Sheet1"
                     Model = largeSheetModel
                     Visible = true }: WB.Sheet) ]
        SheetOrder = [ "sheet1" ]
        ActiveSheet = "sheet1"
        Tables = tables }
let workbookNameTrimOrderLongName prefix nameIndex =
    sprintf "%s_%04d_%s" prefix nameIndex (System.String('z', 12000))
let workbookNameTrimOrderSheetWorkbook () : WB.Workbook =
    let sheetIds =
        [ for index in 0 .. workbookNameRetainedNameCount - 1 ->
            sprintf "trim-sheet-%04d" index ]
    let sheets =
        sheetIds
        |> List.mapi (fun index id ->
            id,
            ({ Id = id
               Name = workbookNameTrimOrderLongName "TrimSheetName" index
               Model = largeSheetModel
               Visible = true }: WB.Sheet))
        |> Map.ofList
    { Sheets = sheets
      SheetOrder = sheetIds
      ActiveSheet = List.head sheetIds
      Tables = Map.empty
      Names = Map.empty }
let workbookNameTrimOrderTableWorkbook () : WB.Workbook =
    let tables =
        [ for index in 0 .. workbookNameRetainedNameCount - 1 ->
            let name = workbookNameTrimOrderLongName "TrimTableName" index
            name, ({ table with Name = name }: WB.Table) ]
        |> Map.ofList
    { largeWorkbook with
        Sheets =
            Map.ofList
                [ "sheet1",
                  ({ Id = "sheet1"
                     Name = "Sheet1"
                     Model = largeSheetModel
                     Visible = true }: WB.Sheet) ]
        SheetOrder = [ "sheet1" ]
        ActiveSheet = "sheet1"
        Tables = tables }
let workbookSheetNamesTrimOrderBytes, workbookSheetNamesTrimOrderTotal =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable total = 0
    for workbookIndex in 0 .. 7 do
        total <- total + (WB.getSheetNames (workbookNameRetainedSheetWorkbook (1000 + workbookIndex))).Length
    total <- total + (WB.getSheetNames (workbookNameTrimOrderSheetWorkbook())).Length
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before, total
check "workbook sheet names trim current memory"
    (workbookSheetNamesTrimOrderBytes < 18000000L
     && workbookSheetNamesTrimOrderTotal = 9 * workbookNameRetainedNameCount)
    true
let workbookTableNamesTrimOrderBytes, workbookTableNamesTrimOrderTotal =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable total = 0
    for workbookIndex in 0 .. 7 do
        total <- total + (WB.getTableNames (workbookNameRetainedTableWorkbook (1000 + workbookIndex))).Length
    total <- total + (WB.getTableNames (workbookNameTrimOrderTableWorkbook())).Length
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before, total
check "workbook table names trim current memory"
    (workbookTableNamesTrimOrderBytes < 18000000L
     && workbookTableNamesTrimOrderTotal = 9 * workbookNameRetainedNameCount)
    true
let workbookSheetNamesRetainedBytes, workbookSheetNamesRetainedTotal =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable total = 0
    for workbookIndex in 0 .. workbookNameRetainedWorkbookCount - 1 do
        total <- total + (WB.getSheetNames (workbookNameRetainedSheetWorkbook workbookIndex)).Length
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before, total
check "workbook sheet names cache retained memory"
    (workbookSheetNamesRetainedBytes < 60000000L
     && workbookSheetNamesRetainedTotal = workbookNameRetainedWorkbookCount * workbookNameRetainedNameCount)
    true
let workbookTableNamesRetainedBytes, workbookTableNamesRetainedTotal =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable total = 0
    for workbookIndex in 0 .. workbookNameRetainedWorkbookCount - 1 do
        total <- total + (WB.getTableNames (workbookNameRetainedTableWorkbook workbookIndex)).Length
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before, total
check "workbook table names cache retained memory"
    (workbookTableNamesRetainedBytes < 60000000L
     && workbookTableNamesRetainedTotal = workbookNameRetainedWorkbookCount * workbookNameRetainedNameCount)
    true
let workbookSheetNamesOversizedSourceWeakReference () =
    let sheetId = "oversized-sheet"
    let sheetName = System.String('S', 6 * 1024 * 1024)
    let sheets =
        Map.ofList [
            sheetId,
            ({ Id = sheetId; Name = sheetName; Model = largeSheetModel; Visible = true }: WB.Sheet)
        ]
    let sheetOrder = [ sheetId ]
    let workbook: WB.Workbook =
        { Sheets = sheets
          SheetOrder = sheetOrder
          ActiveSheet = sheetId
          Tables = Map.empty
          Names = Map.empty }
    WB.getSheetNames workbook |> ignore
    System.WeakReference(sheets)
let workbookSheetNamesOversizedSourceWeak = workbookSheetNamesOversizedSourceWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "workbook sheet names cache skips oversized source"
    workbookSheetNamesOversizedSourceWeak.IsAlive
    false
let workbookSheetNamesOversizedNameSkipsModelScanMs, workbookSheetNamesOversizedNameSkipsModelScanCount =
    let sourceCellCount = 100000
    let cells =
        seq {
            for row in 0 .. sourceCellCount - 1 do
                yield { Row = row; Col = 0 }, cell (N 1.0)
        }
        |> Map.ofSeq
    let sheetId = "oversized-name-model-scan"
    let sheets =
        Map.ofList [
            sheetId,
            ({ Id = sheetId
               Name = System.String('N', 6 * 1024 * 1024)
               Model = { largeSheetModel with Cells = cells }
               Visible = true }: WB.Sheet)
        ]
    let workbook: WB.Workbook =
        { Sheets = sheets
          SheetOrder = [ sheetId ]
          ActiveSheet = sheetId
          Tables = Map.empty
          Names = Map.empty }
    let sw = System.Diagnostics.Stopwatch.StartNew()
    let mutable total = 0
    for _ in 1 .. 40 do
        total <- total + (WB.getSheetNames workbook).Length
    sw.Stop()
    sw.Elapsed.TotalMilliseconds, total
check "workbook sheet names oversized name skips model scan"
    (workbookSheetNamesOversizedNameSkipsModelScanCount = 40
     && workbookSheetNamesOversizedNameSkipsModelScanMs < 120.0)
    true
let workbookSheetNamesOversizedFilterTextSkipsModelScanMs, workbookSheetNamesOversizedFilterTextSkipsModelScanCount =
    let sourceCellCount = 100000
    let payload = System.String('F', 6 * 1024 * 1024)
    let cells =
        seq {
            for row in 0 .. sourceCellCount - 1 do
                yield { Row = row; Col = 0 }, cell (N 1.0)
        }
        |> Map.ofSeq
    let sheetId = "oversized-filter-text-model-scan"
    let model =
        { largeSheetModel with
            Cells = cells
            Filters = Map.ofList [ 0, { Column = 0; VisibleValues = Set.singleton payload } ] }
    let workbook: WB.Workbook =
        { Sheets =
            Map.ofList [
                sheetId,
                ({ Id = sheetId
                   Name = "OversizedFilterText"
                   Model = model
                   Visible = true }: WB.Sheet)
            ]
          SheetOrder = [ sheetId ]
          ActiveSheet = sheetId
          Tables = Map.empty
          Names = Map.empty }
    let sw = System.Diagnostics.Stopwatch.StartNew()
    let mutable total = 0
    for _ in 1 .. 40 do
        total <- total + (WB.getSheetNames workbook).Length
    sw.Stop()
    sw.Elapsed.TotalMilliseconds, total
check "workbook sheet names oversized filter text skips model scan"
    (workbookSheetNamesOversizedFilterTextSkipsModelScanCount = 40
     && workbookSheetNamesOversizedFilterTextSkipsModelScanMs < 120.0)
    true
let workbookSheetNamesOversizedModelWeakReference () =
    let sheetId = "small-sheet"
    let model =
        { largeSheetModel with
            Cells = Map.ofList [ { Row = 0; Col = 0 }, cell (S (System.String('M', 6 * 1024 * 1024))) ] }
    let sheets =
        Map.ofList [
            sheetId,
            ({ Id = sheetId; Name = "Small"; Model = model; Visible = true }: WB.Sheet)
        ]
    let workbook: WB.Workbook =
        { Sheets = sheets
          SheetOrder = [ sheetId ]
          ActiveSheet = sheetId
          Tables = Map.empty
          Names = Map.empty }
    WB.getSheetNames workbook |> ignore
    System.WeakReference(sheets)
let workbookSheetNamesOversizedModelWeak = workbookSheetNamesOversizedModelWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "workbook sheet names cache skips oversized model"
    workbookSheetNamesOversizedModelWeak.IsAlive
    false
let workbookSheetNamesOversizedPageSetupItemsWeakReference () =
    let sheetId = "page-setup-items"
    let pageSetup =
        { defaultPageSetup with
            ManualRowBreaks = [ for index in 0 .. 500000 -> index ] }
    let model = { largeSheetModel with PageSetup = pageSetup }
    let sheets =
        Map.ofList [
            sheetId,
            ({ Id = sheetId; Name = "PageSetupItems"; Model = model; Visible = true }: WB.Sheet)
        ]
    let workbook: WB.Workbook =
        { Sheets = sheets
          SheetOrder = [ sheetId ]
          ActiveSheet = sheetId
          Tables = Map.empty
          Names = Map.empty }
    WB.getSheetNames workbook |> ignore
    System.WeakReference(sheets)
let workbookSheetNamesOversizedPageSetupItemsWeak = workbookSheetNamesOversizedPageSetupItemsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "workbook sheet names cache skips oversized page setup items"
    workbookSheetNamesOversizedPageSetupItemsWeak.IsAlive
    false
let workbookSheetNamesOversizedMergedSourceMs, workbookSheetNamesOversizedMergedSourceCount =
    let mergedCells =
        [ for index in 0 .. 1000000 ->
            { StartRow = 1000000 + index
              StartCol = 0
              EndRow = 1000000 + index
              EndCol = 0 } ]
    let sheetId = "oversized-merged-source"
    let model = { largeSheetModel with MergedCells = mergedCells }
    let workbook: WB.Workbook =
        { Sheets =
            Map.ofList [
                sheetId,
                ({ Id = sheetId; Name = "OversizedMerged"; Model = model; Visible = true }: WB.Sheet)
            ]
          SheetOrder = [ sheetId ]
          ActiveSheet = sheetId
          Tables = Map.empty
          Names = Map.empty }
    let sw = System.Diagnostics.Stopwatch.StartNew()
    let mutable total = 0
    for _ in 1 .. 20 do
        total <- total + (WB.getSheetNames workbook).Length
    sw.Stop()
    sw.Elapsed.TotalMilliseconds, total
check "workbook sheet names oversized merged source fast"
    (workbookSheetNamesOversizedMergedSourceCount = 20
     && workbookSheetNamesOversizedMergedSourceMs < 180.0)
    true
let workbookSheetNamesOversizedItemSourceMs, workbookSheetNamesOversizedItemSourceCount =
    let sourceCellCount = 500001
    let cells =
        seq {
            for row in 0 .. sourceCellCount - 1 do
                yield { Row = row; Col = 0 }, cell (N 1.0)
        }
        |> Map.ofSeq
    let sheetId = "oversized-items"
    let model = { largeSheetModel with Cells = cells }
    let workbook: WB.Workbook =
        { Sheets =
            Map.ofList [
                sheetId,
                ({ Id = sheetId; Name = "OversizedItems"; Model = model; Visible = true }: WB.Sheet)
            ]
          SheetOrder = [ sheetId ]
          ActiveSheet = sheetId
          Tables = Map.empty
          Names = Map.empty }
    let sw = System.Diagnostics.Stopwatch.StartNew()
    let names = WB.getSheetNames workbook
    sw.Stop()
    sw.Elapsed.TotalMilliseconds, names.Length
check "workbook sheet names oversized item source fast"
    (workbookSheetNamesOversizedItemSourceCount = 1 && workbookSheetNamesOversizedItemSourceMs < 45.0)
    true
let workbookSheetNamesOversizedOrderWeakReference () =
    let sheetId = "small-sheet"
    let sheets =
        Map.ofList [
            sheetId,
            ({ Id = sheetId; Name = "Small"; Model = largeSheetModel; Visible = true }: WB.Sheet)
        ]
    let sheetOrder = [ System.String('O', 6 * 1024 * 1024) ]
    let workbook: WB.Workbook =
        { Sheets = sheets
          SheetOrder = sheetOrder
          ActiveSheet = sheetId
          Tables = Map.empty
          Names = Map.empty }
    WB.getSheetNames workbook |> ignore
    System.WeakReference(sheetOrder)
let workbookSheetNamesOversizedOrderWeak = workbookSheetNamesOversizedOrderWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "workbook sheet names cache skips oversized order"
    workbookSheetNamesOversizedOrderWeak.IsAlive
    false
let workbookSheetNamesInvalidOversizedOrderMs, workbookSheetNamesInvalidOversizedOrderTotal =
    let sheetCount = 100000
    let sheetIds = [ for index in 0 .. sheetCount - 1 -> sprintf "invalid-order-sheet-%06d" index ]
    let sheets =
        sheetIds
        |> List.map (fun id ->
            id,
            ({ Id = id
               Name = id
               Model = largeSheetModel
               Visible = true }: WB.Sheet))
        |> Map.ofList
    let workbook: WB.Workbook =
        { Sheets = sheets
          SheetOrder = [ System.String('I', 6 * 1024 * 1024) ]
          ActiveSheet = sheetIds.[0]
          Tables = Map.empty
          Names = Map.empty }
    let sw = System.Diagnostics.Stopwatch.StartNew()
    let mutable total = 0
    for _ in 1 .. 20 do
        total <- total + (WB.getSheetNames workbook).Length
    sw.Stop()
    sw.Elapsed.TotalMilliseconds, total
check "workbook sheet names invalid oversized order fast"
    (workbookSheetNamesInvalidOversizedOrderTotal = 0
     && workbookSheetNamesInvalidOversizedOrderMs < 120.0)
    true
let workbookTableNamesOversizedSourceWeakReference () =
    let tableName = System.String('T', 6 * 1024 * 1024)
    let tables =
        Map.ofList [
            tableName,
            ({ table with Name = tableName }: WB.Table)
        ]
    let workbook =
        { wb with
            Tables = tables
            Names = Map.empty }
    WB.getTableNames workbook |> ignore
    System.WeakReference(tables)
let workbookTableNamesOversizedSourceWeak = workbookTableNamesOversizedSourceWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "workbook table names cache skips oversized source"
    workbookTableNamesOversizedSourceWeak.IsAlive
    false
let workbookTableNamesOversizedMapNameSkipsColumnsMs, workbookTableNamesOversizedMapNameSkipsColumnsCount =
    let columns = [ for index in 0 .. 100000 -> sprintf "Column%06d" index ]
    let tableName = System.String('K', 6 * 1024 * 1024)
    let tables =
        Map.ofList [
            tableName,
            ({ table with
                Name = "SmallTableName"
                Columns = columns }: WB.Table)
        ]
    let workbook =
        { wb with
            Tables = tables
            Names = Map.empty }
    let sw = System.Diagnostics.Stopwatch.StartNew()
    let mutable total = 0
    for _ in 1 .. 40 do
        total <- total + (WB.getTableNames workbook).Length
    sw.Stop()
    sw.Elapsed.TotalMilliseconds, total
check "workbook table names oversized map name skips columns"
    (workbookTableNamesOversizedMapNameSkipsColumnsCount = 40
     && workbookTableNamesOversizedMapNameSkipsColumnsMs < 120.0)
    true
let workbookTableNamesOversizedColumnsWeakReference () =
    let tables =
        Map.ofList [
            "SmallTable",
            ({ table with
                Name = "SmallTable"
                Columns = [ System.String('C', 6 * 1024 * 1024) ] }: WB.Table)
        ]
    let workbook =
        { wb with
            Tables = tables
            Names = Map.empty }
    WB.getTableNames workbook |> ignore
    System.WeakReference(tables)
let workbookTableNamesOversizedColumnsWeak = workbookTableNamesOversizedColumnsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "workbook table names cache skips oversized columns"
    workbookTableNamesOversizedColumnsWeak.IsAlive
    false
let workbookTableNamesOversizedColumnCountWeakReference () =
    let columns = [ for _ in 0 .. 500000 -> "" ]
    let tables =
        Map.ofList [
            "ManyEmptyColumns",
            ({ table with
                Name = "ManyEmptyColumns"
                Columns = columns }: WB.Table)
        ]
    let workbook =
        { wb with
            Tables = tables
            Names = Map.empty }
    let sw = System.Diagnostics.Stopwatch.StartNew()
    let names = WB.getTableNames workbook
    sw.Stop()
    System.WeakReference(columns), sw.Elapsed.TotalMilliseconds, names.Length
let workbookTableNamesOversizedColumnCountWeak, workbookTableNamesOversizedColumnCountMs, workbookTableNamesOversizedColumnCountNames =
    workbookTableNamesOversizedColumnCountWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "workbook table names cache skips oversized column count"
    workbookTableNamesOversizedColumnCountWeak.IsAlive
    false
check "workbook table names oversized column count fast"
    (workbookTableNamesOversizedColumnCountNames = 1 && workbookTableNamesOversizedColumnCountMs < 45.0)
    true

// --- structured-ref grammar ---
check "Table[Col] scalar (first)" (num (evalS "Sales[Amount]")) 10.0
check "SUM Table[Col]"            (num (evalS "SUM(Sales[Amount])")) 100.0
check "multi-word column"         (num (evalS "SUM(Sales[Sales Amount])")) 1000.0
check "#Data only"                (num (evalS "SUM(Sales[[#Data],[Amount]])")) 100.0
check "#Headers scalar"           (str (evalS "Sales[[#Headers],[Amount]]")) "Amount"
check "#Totals scalar"            (num (evalS "Sales[[#Totals],[Amount]]")) 100.0
check "column span sum"           (num (evalS "SUM(Sales[[Amount]:[Sales Amount]])")) 1100.0
check "#All sum (header text ignored)" (num (evalS "SUM(Sales[[#All],[Amount]])")) 200.0  // data 100 + totals 100

// --- this-row [@Col] (qualified + unqualified), driven by CurrentCell ---
check "qualified @ row2"  (num (evalIn model (Some { Row=2; Col=0 }) "Sales[@Amount]")) 20.0
check "unqualified @ row3" (num (evalIn model (Some { Row=3; Col=0 }) "[@Amount]")) 30.0
check "@ off-table -> #VALUE!" (evalIn model (Some { Row=0; Col=0 }) "Sales[@Amount]") (CellValue.Error "VALUE!")

// --- SUBTOTAL ---
check "SUBTOTAL 9 sum"        (num (evalS "SUBTOTAL(9,B2:B5)")) 100.0
check "SUBTOTAL 1 average"    (num (evalS "SUBTOTAL(1,B2:B5)")) 25.0
check "SUBTOTAL 2 count"      (num (evalS "SUBTOTAL(2,B2:B5)")) 4.0
check "SUBTOTAL 4 max"        (num (evalS "SUBTOTAL(4,B2:B5)")) 40.0
// nested SUBTOTAL in B6 is skipped → B2:B6 still sums to 100, not 200
check "SUBTOTAL skips nested" (num (evalS "SUBTOTAL(9,B2:B6)")) 100.0
check "plain SUM counts nested" (num (evalS "SUM(B2:B6)")) 200.0
let sparseSubtotal =
    { createEmpty 20010 4 with
        Cells =
            [ { Row=0; Col=0 }, cell (N 10.0)
              { Row=9999; Col=0 }, cell (N 20.0)
              { Row=14999; Col=0 }, cell (S "x")
              { Row=19999; Col=0 }, cell (CellValue.Formula ("SUBTOTAL(9,A1:A2)", N 999.0)) ]
            |> Map.ofList }
check "SUBTOTAL sparse sum"     (num (evalIn sparseSubtotal None "SUBTOTAL(9,A1:A20000)")) 30.0
check "SUBTOTAL sparse average" (num (evalIn sparseSubtotal None "SUBTOTAL(1,A1:A20000)")) 15.0
check "SUBTOTAL sparse count"   (num (evalIn sparseSubtotal None "SUBTOTAL(2,A1:A20000)")) 2.0
check "SUBTOTAL sparse counta"  (num (evalIn sparseSubtotal None "SUBTOTAL(3,A1:A20000)")) 3.0
let wideNestedSubtotalArgCount = 200000
let wideNestedSubtotalArgs = Array.create (wideNestedSubtotalArgCount + 1) "1"
wideNestedSubtotalArgs.[wideNestedSubtotalArgCount] <- "SUBTOTAL(9,A1:A2)"
let wideNestedSubtotalFormula =
    "SUM(" + System.String.Join(",", wideNestedSubtotalArgs) + ")"
let wideNestedSubtotalModel =
    { createEmpty 4 4 with
        Cells =
            [ { Row=0; Col=0 }, cell (CellValue.Formula (wideNestedSubtotalFormula, N 123.0))
              { Row=1; Col=0 }, cell (N 10.0) ]
            |> Map.ofList }
let wideNestedSubtotalSw = System.Diagnostics.Stopwatch.StartNew()
let wideNestedSubtotalResult = evalIn wideNestedSubtotalModel None "SUBTOTAL(9,A1:A2)"
wideNestedSubtotalSw.Stop()
if wideNestedSubtotalResult = N 10.0 && wideNestedSubtotalSw.Elapsed.TotalMilliseconds < 3000.0 then
    pass <- pass + 1
else
    fail <- fail + 1
    printfn "FAIL %-30s expected %A in under 3000 ms got %A in %.0f ms" "SUBTOTAL wide nested scan" (N 10.0) wideNestedSubtotalResult wideNestedSubtotalSw.Elapsed.TotalMilliseconds

// manually-hidden row (height 0): 109 excludes it, 9 includes it
let hidden = { model with RowHeights = Map.ofList [ 3, 0.0 ] }   // row idx 3 = value 30
check "SUBTOTAL 109 skips hidden" (num (evalIn hidden None "SUBTOTAL(109,B2:B5)")) 70.0
check "SUBTOTAL 9 keeps hidden"   (num (evalIn hidden None "SUBTOTAL(9,B2:B5)")) 100.0

// filter-hidden rows excluded by both 9 and 109
let filtered = applyFilter 0 (Set.ofList [ "East" ]) model   // keeps rows 1,3 (East)
check "SUBTOTAL 9 honors filter" (num (evalIn filtered None "SUBTOTAL(9,B2:B5)")) 40.0  // 10 + 30
let filterRight =
    let cells =
        seq {
            yield { Row=0; Col=0 }, cell (S "Amount")
            yield { Row=0; Col=1 }, cell (S "Region")
            for r = 1 to 1500 do
                yield { Row=r; Col=0 }, cell (N 1.0)
                yield { Row=r; Col=1 }, cell (S (if r % 2 = 0 then "Keep" else "Drop"))
        }
        |> Map.ofSeq
    { createEmpty 1502 3 with Cells = cells }
    |> applyFilter 1 (Set.singleton "Keep")
check "SUBTOTAL sparse filter to right" (num (evalIn filterRight None "SUBTOTAL(9,A2:A1501)")) 750.0

// --- AGGREGATE ---
let withErr = { model with Cells = Map.add { Row=3; Col=1 } (cell (CellValue.Error "DIV/0!")) model.Cells }
check "AGGREGATE 9 opt6 ignores errors" (num (evalIn withErr None "AGGREGATE(9,6,B2:B5)")) 70.0  // 10+20+40
check "AGGREGATE 9 opt4 sees error"     (evalIn withErr None "AGGREGATE(9,4,B2:B5)") (CellValue.Error "DIV/0!")
check "AGGREGATE 1 average"             (num (evalS "AGGREGATE(1,0,B2:B5)")) 25.0
check "AGGREGATE sparse skips nested"   (num (evalIn sparseSubtotal None "AGGREGATE(9,0,A1:A20000)")) 30.0
check "AGGREGATE sparse includes nested" (num (evalIn sparseSubtotal None "AGGREGATE(9,6,A1:A20000)")) 1029.0

// --- bridge helper: makeTableFromRegion (1-indexed XLSX region → 0-indexed) ---
// Columns omitted → read from the header row of the model.
let bridged = WB.makeTableFromRegion "sheet1" "Sales" 1 1 6 3 1 0 [] model
check "bridge name"       bridged.Name "Sales"
check "bridge 0-indexed start" (bridged.StartRow, bridged.StartCol) (0, 0)
check "bridge 0-indexed end"   (bridged.EndRow, bridged.EndCol) (5, 2)
check "bridge hasHeaders"  bridged.HasHeaders true
check "bridge cols from header" bridged.Columns [ "Region"; "Amount"; "Sales Amount" ]
// Explicit column names win over header cells.
let bridged2 = WB.makeTableFromRegion "sheet1" "T" 1 1 6 3 1 1 [ "A"; "B"; "C" ] model
check "bridge explicit cols" bridged2.Columns [ "A"; "B"; "C" ]
check "bridge hasTotals"   bridged2.HasTotals true

// --- createTable: register a real table + SUBTOTAL totals row ---
// Fresh data block: header A1:C1, data rows 2-4.
let dataModel =
    { createEmpty 20 8 with
        Cells =
            [ { Row=0; Col=0 }, cell (S "Region"); { Row=0; Col=1 }, cell (S "Q1"); { Row=0; Col=2 }, cell (S "Q2")
              { Row=1; Col=0 }, cell (S "East");   { Row=1; Col=1 }, cell (N 100.); { Row=1; Col=2 }, cell (N 200.)
              { Row=2; Col=0 }, cell (S "West");   { Row=2; Col=1 }, cell (N 150.); { Row=2; Col=2 }, cell (N 250.) ]
            |> Map.ofList }
let created = createTable "Sales" 0 0 2 2 true dataModel
let ct = List.exactlyOne created.Tables
check "createTable name"      ct.Name "Sales"
check "createTable 1-indexed"  (ct.StartRow, ct.StartCol, ct.EndCol) (1, 1, 3)
check "createTable totals grows range" ct.EndRow 4   // header(1)+data(2,3)+totals(4)
check "createTable hasTotals"  ct.HasTotals true
check "createTable autofilter" ct.AutoFilter (Some (1, 3))
check "createTable style"      (ct.Style |> Option.map (fun s -> s.Name)) (Some "TableStyleMedium2")
check "createTable totals label" (str (getCellValue { Row=3; Col=0 } created)) "Total"
// totals cells hold SUBTOTAL formulas
let totalsFormula =
    match getCellValue { Row=3; Col=1 } created with CellValue.Formula (t, _) -> t | _ -> ""
check "createTable totals formula" totalsFormula "SUBTOTAL(109,B2:B3)"

let largeTableModel =
    let tables = [ for i in 1 .. 50000 -> { ct with Name = sprintf "Table%d" i } ]
    { dataModel with Tables = tables }
checkAllocUnder "createTable append allocation" (allocatedBytesFor (fun () -> createTable "Next" 0 0 2 2 false largeTableModel)) 2400000L

let paintHasFill color model =
    let mutable found = false
    P.iterViewportPaintItems model (L.defaultViewport 500.0 240.0) (fun _ item ->
        match item with
        | P.FillRect (_, fill) when fill = color -> found <- true
        | _ -> ())
    found
let paintModel = { created with Selection = SingleCell(10, 7) }
check "paint table header fill" (paintHasFill "#4472C4" paintModel) true
check "paint table stripe fill" (paintHasFill "#D9E1F2" paintModel) true
let paintTextCacheRetainedModelCount = 80
let paintTextCacheLongValue modelIndex =
    sprintf "paint-text-retained-%03d-%s" modelIndex (System.String('x', 512 * 1024))
let paintTextCacheModel modelIndex =
    { createEmpty 10 5 with
        Cells =
            Map.ofList
                [ { Row = 0; Col = 0 }, cell (S (paintTextCacheLongValue modelIndex)) ] }
let paintTextCacheViewport = L.defaultViewport 500.0 240.0
let paintTextCacheRetainedBytes build =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable total = 0
    for modelIndex in 0 .. paintTextCacheRetainedModelCount - 1 do
        total <- total + build (paintTextCacheModel modelIndex)
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before, total
let paintViewportTextCacheBytes, paintViewportTextCacheTotal =
    paintTextCacheRetainedBytes (fun model -> (P.buildViewportPaintItems model paintTextCacheViewport).Length)
check "viewport paint text cache retained memory"
    (paintViewportTextCacheBytes < 40000000L && paintViewportTextCacheTotal > 0)
    true
let paintStaticTextCacheBytes, paintStaticTextCacheTotal =
    paintTextCacheRetainedBytes (fun model -> (P.buildPaintItems model paintTextCacheViewport).Length)
check "static paint text cache retained memory"
    (paintStaticTextCacheBytes < 40000000L && paintStaticTextCacheTotal > 0)
    true

// --- table-aware filtering (filter-button scope) ---
check "tableAt header cell" (tableAt 0 0 created |> Option.isSome) true
check "tableAt data cell"   (tableAt 2 1 created |> Option.isSome) true
check "tableAt outside"     (tableAt 12 12 created |> Option.isSome) false
let tableUniqueCol0 = tableColumnUniqueValues ct 0 created
check "tableColUnique col0" tableUniqueCol0 [ "East"; "West" ]
check "tableColUnique excludes totals" (tableColumnUniqueValues ct 1 created) [ "100"; "150" ]
check "tableColUnique cached repeat" (tableColumnUniqueValues ct 0 created) [ "East"; "West" ]
check "tableColUnique survives other col" (obj.ReferenceEquals(tableColumnUniqueValues ct 0 created, tableUniqueCol0)) true
let tableUniqueEdited = created |> setCell { Row=2; Col=0 } (cell (S "North"))
check "tableColUnique invalidates edit" (tableColumnUniqueValues ct 0 tableUniqueEdited) [ "East"; "North" ]
check "tableColUnique keeps first source" (obj.ReferenceEquals(tableColumnUniqueValues ct 0 created, tableUniqueCol0)) true
let wideUniqueTable =
    { Name = "WideUnique"
      StartRow = 1
      StartCol = 1
      EndRow = 12
      EndCol = 24
      HeaderRowCount = 1
      Style = None
      AutoFilter = Some (1, 24)
      HasTotals = false }
let wideUniqueModel =
    { createEmpty 16 28 with
        Tables = [ wideUniqueTable ]
        Cells =
            [ for row in 0 .. 11 do
                for col in 0 .. 23 ->
                    let value =
                        if row = 0 then S (sprintf "H%02d" col)
                        else S (sprintf "c%02d-%02d" col (row % 4))
                    { Row = row; Col = col }, cell value ]
            |> Map.ofList }
let wideTableUniqueFirst = tableColumnUniqueValues wideUniqueTable 0 wideUniqueModel
for col in 1 .. 23 do
    tableColumnUniqueValues wideUniqueTable col wideUniqueModel |> ignore
check "tableColUnique keeps many columns" (obj.ReferenceEquals(tableColumnUniqueValues wideUniqueTable 0 wideUniqueModel, wideTableUniqueFirst)) true
let tableUniqueManyModelCount = 24
let tableUniqueManyRowCount = 5000
let tableUniqueManyTable rows : RedBat.Spreadsheet.Model.Table =
    { Name = "ManyUnique"
      StartRow = 1
      StartCol = 1
      EndRow = rows
      EndCol = 1
      HeaderRowCount = 1
      Style = None
      AutoFilter = Some (1, 1)
      HasTotals = false }
let tableUniqueManyModels =
    [| for modelIndex in 0 .. tableUniqueManyModelCount - 1 ->
        let table = tableUniqueManyTable tableUniqueManyRowCount
        { createEmpty tableUniqueManyRowCount 2 with
            Tables = [ table ]
            Cells =
                [ for row in 0 .. tableUniqueManyRowCount - 1 ->
                    { Row = row; Col = 0 }, cell (S (sprintf "table-unique-%02d" (row % 32))) ]
                |> Map.ofList } |]
let tableUniqueManyRun passes =
    let mutable total = 0
    for _pass in 1 .. passes do
        for model in tableUniqueManyModels do
            total <- total + (tableColumnUniqueValues (List.head model.Tables) 0 model).Length
    total
tableUniqueManyRun 3 |> ignore
let tableUniqueManySw = System.Diagnostics.Stopwatch.StartNew()
let tableUniqueManyTotal = tableUniqueManyRun 30
tableUniqueManySw.Stop()
check "tableColUnique many models"
    (tableUniqueManySw.Elapsed.TotalMilliseconds < 35.0 && tableUniqueManyTotal > 0)
    true
let tableUniqueNoTotals =
    tableColumnUniqueValues { ct with HasTotals = false } 0 (created |> setCell { Row=3; Col=0 } (cell (S "South")))
check "tableColUnique table flags invalidate" tableUniqueNoTotals [ "East"; "South"; "West" ]

// --- applyTableStyles: bakes header/stripe styling for created tables only ---
let fmtAt (m: SpreadsheetModel) r c = (getCell { Row=r; Col=c } m |> Option.get).Format
let styledM = applyTableStyles created
let hdr = fmtAt styledM 0 0
check "style header fill"   hdr.BgColor (Some "#4472C4")
check "style header white"  hdr.TextColor (Some "#FFFFFF")
check "style header bold"    hdr.Bold true
check "style header filter btn" hdr.FilterButton true
check "style stripe row0"   (fmtAt styledM 1 0).BgColor (Some "#D9E1F2")   // East, even data idx
check "style no stripe row1" (fmtAt styledM 2 0).BgColor None               // West, odd data idx
let restyledM = applyTableStyles styledM
check "style reapply same cells" (obj.ReferenceEquals(restyledM.Cells, styledM.Cells)) true
let restyledSelection = applyTableStyles { styledM with Selection = SingleCell(7, 7) }
check "style reapply preserves selection" restyledSelection.Selection (SingleCell(7, 7))
check "style reapply selection cells" (obj.ReferenceEquals(restyledSelection.Cells, styledM.Cells)) true
// Imported-style table (HeaderFill None) is left untouched.
let importedStyleModel =
    { dataModel with
        Tables = [ { Name="T"; StartRow=1; StartCol=1; EndRow=3; EndCol=3; HeaderRowCount=1
                     Style = Some { Name="TableStyleMedium7"; ShowFirstColumn=false; ShowLastColumn=false; ShowRowStripes=true; ShowColumnStripes=false; HeaderFill=None; StripeFill=None }
                     AutoFilter = Some (1,3); HasTotals=false } ] }
check "style no-op when HeaderFill None" (fmtAt (applyTableStyles importedStyleModel) 0 0).BgColor None
check "style cache keeps first source" (obj.ReferenceEquals((applyTableStyles created).Cells, styledM.Cells)) true
let tableStyleDenseRows = 1800
let tableStyleDenseCols = 30
let tableStyleDenseStyle =
    Some { Name = "TableStyleMedium2"; ShowFirstColumn = false; ShowLastColumn = false; ShowRowStripes = true; ShowColumnStripes = false; HeaderFill = Some "#4472C4"; StripeFill = Some "#D9E1F2" }
let tableStyleDenseTable : RedBat.Spreadsheet.Model.Table =
    { Name = "DenseStyle"
      StartRow = 1
      StartCol = 1
      EndRow = tableStyleDenseRows
      EndCol = tableStyleDenseCols
      HeaderRowCount = 1
      Style = tableStyleDenseStyle
      AutoFilter = Some (1, tableStyleDenseCols)
      HasTotals = false }
let tableStyleDenseModel =
    { createEmpty tableStyleDenseRows tableStyleDenseCols with
        Cells =
            [ for col in 0 .. tableStyleDenseCols - 1 ->
                { Row = 0; Col = col }, cell (S (sprintf "H%d" col)) ]
            |> Map.ofList
        Tables = [ tableStyleDenseTable ] }
let tableStyleDenseSw = System.Diagnostics.Stopwatch.StartNew()
let tableStyleDenseStyled = applyTableStyles tableStyleDenseModel
tableStyleDenseSw.Stop()
check "style dense batch perf"
    (tableStyleDenseSw.Elapsed.TotalMilliseconds < 500.0
     && tableStyleDenseStyled.Cells.Count > tableStyleDenseCols * 100)
    true
let tableStyleManyModelCount = 24
let tableStyleManyRowCount = 1200
let tableStyleManyColCount = 10
let tableStyleManyStyle =
    Some { Name = "TableStyleMedium2"; ShowFirstColumn = false; ShowLastColumn = false; ShowRowStripes = true; ShowColumnStripes = false; HeaderFill = Some "#4472C4"; StripeFill = Some "#D9E1F2" }
let tableStyleManyTable modelIndex : RedBat.Spreadsheet.Model.Table =
    { Name = sprintf "StyleCache%d" modelIndex
      StartRow = 1
      StartCol = 1
      EndRow = tableStyleManyRowCount
      EndCol = tableStyleManyColCount
      HeaderRowCount = 1
      Style = tableStyleManyStyle
      AutoFilter = Some (1, tableStyleManyColCount)
      HasTotals = false }
let tableStyleManyModels =
    [| for modelIndex in 0 .. tableStyleManyModelCount - 1 ->
        { createEmpty tableStyleManyRowCount tableStyleManyColCount with
            Cells =
                [ for col in 0 .. tableStyleManyColCount - 1 ->
                    { Row = 0; Col = col }, cell (S (sprintf "H%d" col)) ]
                |> Map.ofList
            Tables = [ tableStyleManyTable modelIndex ] } |]
let tableStyleManyStyledModels = tableStyleManyModels |> Array.map applyTableStyles
let tableStyleManyRun passes =
    let mutable total = 0
    let mutable keptIdentity = true
    for _pass in 1 .. passes do
        for model in tableStyleManyStyledModels do
            let styled = applyTableStyles model
            keptIdentity <- keptIdentity && obj.ReferenceEquals(styled.Cells, model.Cells)
            total <- total + styled.Cells.Count
    total, keptIdentity
tableStyleManyRun 3 |> ignore
let tableStyleManySw = System.Diagnostics.Stopwatch.StartNew()
let tableStyleManyTotal, tableStyleManyKeptIdentity = tableStyleManyRun 60
tableStyleManySw.Stop()
let tableStyleManyOk =
    tableStyleManySw.Elapsed.TotalMilliseconds < 150.0 && tableStyleManyKeptIdentity && tableStyleManyTotal > 0
if not tableStyleManyOk then
    printfn "style cache many models details: %.1f ms kept=%b total=%d" tableStyleManySw.Elapsed.TotalMilliseconds tableStyleManyKeptIdentity tableStyleManyTotal
check "style cache many models"
    tableStyleManyOk
    true
let tableStyleRetainedModelCount = 80
let tableStyleRetainedTableCount = 10000
let tableStyleRetainedTable modelIndex tableIndex : RedBat.Spreadsheet.Model.Table =
    { Name = sprintf "ImportedNoStyle_%03d_%05d" modelIndex tableIndex
      StartRow = tableIndex + 1
      StartCol = 1
      EndRow = tableIndex + 1
      EndCol = 1
      HeaderRowCount = 1
      Style = None
      AutoFilter = None
      HasTotals = false }
let tableStyleRetainedModel modelIndex =
    { createEmpty (tableStyleRetainedTableCount + 1) 2 with
        Tables =
            [ for tableIndex in 0 .. tableStyleRetainedTableCount - 1 ->
                tableStyleRetainedTable modelIndex tableIndex ] }
let tableStyleRetainedBytes, tableStyleRetainedTotal =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable total = 0
    for modelIndex in 0 .. tableStyleRetainedModelCount - 1 do
        let styled = applyTableStyles (tableStyleRetainedModel modelIndex)
        total <- total + styled.Cells.Count + styled.Tables.Length
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before, total
check "style cache retained table memory"
    (tableStyleRetainedBytes < 50000000L
     && tableStyleRetainedTotal = tableStyleRetainedModelCount * tableStyleRetainedTableCount)
    true

let tableStyleSourceTextModelCount = 24
let tableStyleSourceTextPayload modelIndex =
    String.replicate 2000000 (string (char (65 + (modelIndex % 26))))
let tableStyleSourceTextTable modelIndex : RedBat.Spreadsheet.Model.Table =
    { Name = sprintf "SourceTextStyle_%03d" modelIndex
      StartRow = 1
      StartCol = 1
      EndRow = 2
      EndCol = 1
      HeaderRowCount = 1
      Style = tableStyleManyStyle
      AutoFilter = Some (1, 1)
      HasTotals = false }
let tableStyleSourceTextModel modelIndex =
    { createEmpty 3 1 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (S (tableStyleSourceTextPayload modelIndex)) ]
            |> Map.ofList
        Tables = [ tableStyleSourceTextTable modelIndex ] }
let tableStyleSourceTextBytes, tableStyleSourceTextTotal =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable total = 0
    for modelIndex in 0 .. tableStyleSourceTextModelCount - 1 do
        let styled = applyTableStyles (tableStyleSourceTextModel modelIndex)
        total <- total + styled.Cells.Count
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before, total
check "style cache source text memory"
    (tableStyleSourceTextBytes < 30000000L
     && tableStyleSourceTextTotal = tableStyleSourceTextModelCount * 2)
    true

let tableStyleOversizedSourceCellsWeakReference () =
    let cells =
        seq {
            for row in 0 .. 200000 do
                yield { Row = row; Col = 0 }, cell (S "x")
        }
        |> Map.ofSeq
    let tables =
        [ { Name = "OversizedStyleSourceCells"
            StartRow = 1
            StartCol = 1
            EndRow = 1
            EndCol = 1
            HeaderRowCount = 1
            Style = None
            AutoFilter = None
            HasTotals = false } ]
    let model = { createEmpty 200001 1 with Cells = cells; Tables = tables }
    applyTableStyles model |> ignore
    System.WeakReference(cells)
let tableStyleOversizedSourceCellsWeak = tableStyleOversizedSourceCellsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "style cache skips oversized source cells"
    tableStyleOversizedSourceCellsWeak.IsAlive
    false

let tableStyleOversizedTableNameWeakReference () =
    let payload = System.String('N', 6 * 1024 * 1024)
    let tables =
        [ { Name = payload
            StartRow = 1
            StartCol = 1
            EndRow = 2
            EndCol = 1
            HeaderRowCount = 1
            Style = tableStyleManyStyle
            AutoFilter = Some (1, 1)
            HasTotals = false } ]
    let model = { createEmpty 3 1 with Tables = tables }
    applyTableStyles model |> ignore
    System.WeakReference(tables)
let tableStyleOversizedTableNameWeak = tableStyleOversizedTableNameWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "style cache skips oversized table name"
    tableStyleOversizedTableNameWeak.IsAlive
    false

let tableStyleOversizedStyleTextWeakReference () =
    let payload = System.String('S', 6 * 1024 * 1024)
    let oversizedStyle =
        Some { Name = payload
               ShowFirstColumn = false
               ShowLastColumn = false
               ShowRowStripes = true
               ShowColumnStripes = false
               HeaderFill = Some "#4472C4"
               StripeFill = Some "#D9E1F2" }
    let tables =
        [ { Name = "OversizedStyleText"
            StartRow = 1
            StartCol = 1
            EndRow = 2
            EndCol = 1
            HeaderRowCount = 1
            Style = oversizedStyle
            AutoFilter = Some (1, 1)
            HasTotals = false } ]
    let model = { createEmpty 3 1 with Tables = tables }
    applyTableStyles model |> ignore
    System.WeakReference(tables)
let tableStyleOversizedStyleTextWeak = tableStyleOversizedStyleTextWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "style cache skips oversized style text"
    tableStyleOversizedStyleTextWeak.IsAlive
    false

let tableStyleOversizedStyleTextSkipsCellsMs, tableStyleOversizedStyleTextSkipsCellsCount, tableStyleOversizedStyleTextSkipsCellsExpected =
    let sourceCellCount = 100000
    let payload = System.String('T', 6 * 1024 * 1024)
    let cells =
        seq {
            for row in 0 .. sourceCellCount - 1 do
                yield { Row = row; Col = 0 }, cell (N 1.0)
        }
        |> Map.ofSeq
    let oversizedStyle =
        Some { Name = payload
               ShowFirstColumn = false
               ShowLastColumn = false
               ShowRowStripes = true
               ShowColumnStripes = false
               HeaderFill = Some "#4472C4"
               StripeFill = Some "#D9E1F2" }
    let table =
        { Name = "OversizedStyleTextCells"
          StartRow = 1
          StartCol = 1
          EndRow = 2
          EndCol = 1
          HeaderRowCount = 1
          Style = oversizedStyle
          AutoFilter = Some (1, 1)
          HasTotals = false }
    let model = { createEmpty (sourceCellCount + 1) 1 with Cells = cells; Tables = [ table ] }
    let sw = System.Diagnostics.Stopwatch.StartNew()
    let mutable total = 0
    for _ in 1 .. 40 do
        total <- total + (applyTableStyles model).Cells.Count
    sw.Stop()
    sw.Elapsed.TotalMilliseconds, total, 40 * sourceCellCount
check "style cache oversized style text skips cells"
    (tableStyleOversizedStyleTextSkipsCellsCount = tableStyleOversizedStyleTextSkipsCellsExpected
     && tableStyleOversizedStyleTextSkipsCellsMs < 160.0)
    true

printfn ""
printfn "Stage 17 table tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
