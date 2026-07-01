#r "/home/dave/Dev/browser/renderer/bin/Debug/net10.0/Renderer.dll"
#r "/home/dave/Dev/browser/widgets/bin/Debug/net10.0/Widgets.dll"
#r "/home/dave/Dev/browser/office/office-native/bin/Debug/net10.0/OfficeShell.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"
#r "/home/dave/Dev/browser/office/spreadsheet-native/bin/Debug/net10.0/SpreadsheetWidgets.dll"

open RedBat.Spreadsheet.Model
open System.Diagnostics
open System.Reflection

module Editor = Composit.Ui.SpreadsheetEditor

let mutable pass = 0
let mutable fail = 0

let check label got expected =
    if got = expected then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-38s expected %A got %A" label expected got

let cell value : CellData = { Value = value; Format = defaultFormat; Comment = None }

let editorType = typeof<Editor.SpreadsheetEditorState>.Assembly.GetType("Composit.Ui.SpreadsheetEditor")
let staticBindingFlags = BindingFlags.NonPublic ||| BindingFlags.Public ||| BindingFlags.Static
let gridType = typeof<Editor.SpreadsheetEditorState>.Assembly.GetType("Composit.Ui.SpreadsheetGrid")
let cellCommands =
    gridType.GetMethod("cellCommands", staticBindingFlags)
let cellCommandsInViewport =
    gridType.GetMethod("cellCommandsInViewport", staticBindingFlags)
let measureText =
    gridType.GetMethod("measureText", staticBindingFlags)
let setFocus =
    editorType.GetMethod("setFocus", staticBindingFlags)
let extendFocus =
    editorType.GetMethod("extendFocus", staticBindingFlags)
let startEditing =
    editorType.GetMethod("startEditing", staticBindingFlags)
let commitEdit =
    editorType.GetMethod("commitEdit", staticBindingFlags)
let usedBoundsCached =
    editorType.GetMethod("usedBoundsCached", staticBindingFlags)

let makeLargeCells commandIndex =
    [ for row in 0 .. 119999 ->
        { Row = row; Col = 0 }, cell (CellValue.Number (float (row + commandIndex))) ]
    |> Map.ofList

let nativeHistoryLargeSnapshotWeakReference () =
    let firstCells = makeLargeCells 0
    let initialModel = { createEmpty 120000 1 with Cells = firstCells }
    let state = Editor.createState [| initialModel |] [| "Sheet1" |] 0
    let firstCellsWeak = System.WeakReference(firstCells)

    for commandIndex in 1 .. 6 do
        let cells = makeLargeCells commandIndex
        Editor.applyChange state (fun model -> { model with Cells = cells }) |> ignore

    let retainedSnapshots = state.UndoStack.Value.Length
    let retainedCells =
        state.UndoStack.Value
        |> List.sumBy (fun snapshot ->
            snapshot.Sheets
            |> Array.sumBy (fun sheet -> sheet.Cells.Count))

    firstCellsWeak, retainedSnapshots, retainedCells

let largeHistoryWeak, largeHistoryRetainedSnapshots, largeHistoryRetainedCells =
    nativeHistoryLargeSnapshotWeakReference ()

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

check "native spreadsheet history trims large snapshots"
    (largeHistoryRetainedSnapshots < 6 && largeHistoryRetainedCells <= 500000)
    true
check "native spreadsheet history releases trimmed cells"
    largeHistoryWeak.IsAlive
    false

let nativeHistoryLargeMetadataWeakReference () =
    let metadataPayload index =
        System.String(char (int 'A' + (index % 26)), 3 * 1024 * 1024)
    let largeMetadataModel index =
        let payload = metadataPayload index
        let tableStyle =
            { Name = "Style" + payload
              ShowFirstColumn = false
              ShowLastColumn = false
              ShowRowStripes = true
              ShowColumnStripes = false
              HeaderFill = None
              StripeFill = None }
        let headerFooter = { defaultHeaderFooter with OddHeader = Some payload }
        { createEmpty 1 1 with
            PageSetup = { defaultPageSetup with HeaderFooter = headerFooter }
            Tables =
                [ { Name = "Table" + payload
                    StartRow = 1
                    StartCol = 1
                    EndRow = 1
                    EndCol = 1
                    HeaderRowCount = 1
                    Style = Some tableStyle
                    AutoFilter = None
                    HasTotals = false } ]
            Validations =
                Map.ofList [
                    { Row = 0; Col = 0 },
                    { Rule = ValidationRule.ListValues [ payload ]
                      ErrorMessage = Some payload
                      InputMessage = Some payload }
                ] }
    let initialModel = largeMetadataModel 0
    let state = Editor.createState [| initialModel |] [| "Sheet1" |] 0
    let firstModelWeak = System.WeakReference(initialModel)

    for commandIndex in 1 .. 4 do
        let nextModel = largeMetadataModel commandIndex
        Editor.applyChange state (fun _ -> nextModel) |> ignore

    let retainedSnapshots = state.UndoStack.Value.Length
    let retainedTextChars = state.UndoStack.Value |> List.sumBy (fun snapshot -> snapshot.RetainedTextCharCount)
    firstModelWeak, retainedSnapshots, retainedTextChars

let metadataHistoryWeak, metadataHistoryRetainedSnapshots, metadataHistoryRetainedTextChars =
    nativeHistoryLargeMetadataWeakReference ()

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

check "native spreadsheet history trims metadata"
    (metadataHistoryRetainedSnapshots < 4 && metadataHistoryRetainedTextChars <= 20000000)
    true
check "native spreadsheet history releases metadata"
    metadataHistoryWeak.IsAlive
    false

if isNull cellCommands || isNull cellCommandsInViewport then
    fail <- fail + 1
    printfn "FAIL %-38s expected grid cellCommands helper" "native grid reflection"
else
    let rendererRows = 120
    let rendererCols = 40
    let rendererCells =
        [ for row in 0 .. rendererRows - 1 do
            for col in 0 .. rendererCols - 1 do
                { Row = row; Col = col }, cell (CellValue.String (sprintf "R%dC%d" row col)) ]
        |> Map.ofList
    let rendererModel = { createEmpty rendererRows rendererCols with Cells = rendererCells }
    let invokeRendererCommands () =
        cellCommands.Invoke(null, [| box rendererModel; box 2400.0; box 2400.0 |]) :?> Renderer.Cmd list
    let invokeRendererCommandsWide () =
        cellCommands.Invoke(null, [| box rendererModel; box 2420.0; box 2400.0 |]) :?> Renderer.Cmd list
    let rendererCommandsA = invokeRendererCommands ()
    let rendererCommandsB = invokeRendererCommands ()
    check "native grid command cache reuse"
        (obj.ReferenceEquals(rendererCommandsA, rendererCommandsB))
        true
    let rendererCommandsWide = invokeRendererCommandsWide ()
    let rendererCommandsAAfterWide = invokeRendererCommands ()
    check "native grid command cache retains viewports"
        (obj.ReferenceEquals(rendererCommandsA, rendererCommandsAAfterWide)
         && rendererCommandsWide.Length > 0)
        true

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let rendererCommandsSw = Stopwatch.StartNew()
    let mutable rendererCommandTotal = 0
    for _ in 1 .. 10 do
        rendererCommandTotal <- rendererCommandTotal + (invokeRendererCommands ()).Length
    rendererCommandsSw.Stop()
    check "native grid command cache perf"
        (rendererCommandsSw.Elapsed.TotalMilliseconds < 30.0
         && rendererCommandTotal = rendererCommandsA.Length * 10)
        true

    invokeRendererCommands () |> ignore
    invokeRendererCommandsWide () |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let alternatingAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let alternatingSw = Stopwatch.StartNew()
    let mutable alternatingTotal = 0
    for index in 1 .. 40 do
        alternatingTotal <-
            alternatingTotal
            + if index % 2 = 0 then (invokeRendererCommandsWide ()).Length else (invokeRendererCommands ()).Length
    alternatingSw.Stop()
    let alternatingAllocated = System.GC.GetAllocatedBytesForCurrentThread() - alternatingAllocatedBefore
    check "native grid command cache alternates"
        (alternatingSw.Elapsed.TotalMilliseconds < 30.0
         && alternatingAllocated < 1024L * 1024L
         && alternatingTotal = (rendererCommandsA.Length + rendererCommandsWide.Length) * 20)
        true

    invokeRendererCommands () |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let movingSelectionAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let movingSelectionSw = Stopwatch.StartNew()
    let mutable movingSelectionTotal = 0
    for index in 0 .. 39 do
        let model =
            { rendererModel with
                Selection = SingleCell (index % rendererRows, (index / rendererRows) % rendererCols) }
        movingSelectionTotal <-
            movingSelectionTotal
            + (cellCommands.Invoke(null, [| box model; box 2400.0; box 2400.0 |]) :?> Renderer.Cmd list).Length
    movingSelectionSw.Stop()
    let movingSelectionAllocated = System.GC.GetAllocatedBytesForCurrentThread() - movingSelectionAllocatedBefore
    check "native grid moving selection cache"
        (movingSelectionSw.Elapsed.TotalMilliseconds < 120.0
         && movingSelectionAllocated < 8L * 1024L * 1024L
         && movingSelectionTotal = rendererCommandsA.Length * 40)
        true

    let scrollRows = 20000
    let scrollCells =
        [ for row in 0 .. scrollRows - 1 ->
            { Row = row; Col = 0 }, cell (CellValue.Number (float row)) ]
        |> Map.ofList
    let scrollModel = { createEmpty scrollRows 20 with Cells = scrollCells }
    let invokeViewportCommands scrollY =
        cellCommandsInViewport.Invoke(null, [| box scrollModel; box 0.0; box scrollY; box 900.0; box 600.0 |]) :?> Renderer.Cmd list
    let viewportCommandsA = invokeViewportCommands 200000.0
    let viewportCommandsB = invokeViewportCommands 200000.0
    check "native grid viewport command size"
        (viewportCommandsA.Length = viewportCommandsB.Length
         && viewportCommandsA.Length < 1000)
        true

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let viewportAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let viewportSw = Stopwatch.StartNew()
    let mutable viewportTotal = 0
    for index in 0 .. 19 do
        viewportTotal <- viewportTotal + (invokeViewportCommands (200000.0 + float (index % 2) * 420.0)).Length
    viewportSw.Stop()
    let viewportAllocated = System.GC.GetAllocatedBytesForCurrentThread() - viewportAllocatedBefore
    check "native grid viewport render bounded"
        (viewportSw.Elapsed.TotalMilliseconds < 80.0
         && viewportAllocated < 16L * 1024L * 1024L
         && viewportTotal < 10000)
        true

if isNull measureText then
    fail <- fail + 1
    printfn "FAIL %-38s expected grid measureText helper" "native text measure reflection"
else
    let invokeMeasure text =
        measureText.Invoke(null, [| box text; box 11.0; box (Some "Calibri"); box false; box false |]) :?> float
    let firstWidth = invokeMeasure "Repeated spreadsheet text 12345"

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let measureSw = Stopwatch.StartNew()
    let mutable measuredTotal = 0.0
    for _ in 1 .. 20000 do
        measuredTotal <- measuredTotal + invokeMeasure "Repeated spreadsheet text 12345"
    measureSw.Stop()
    check "native text measure cache stable"
        (firstWidth > 0.0 && abs (measuredTotal - firstWidth * 20000.0) < 0.01)
        true
    check "native text measure cache perf"
        (measureSw.Elapsed.TotalMilliseconds < 50.0)
        true

let stablePaintMeasure (text: string) (fontSize: float) (fontName: string option) (bold: bool) (italic: bool) =
    float text.Length * fontSize * (if bold then 0.65 else 0.6)
    + (if italic then 1.0 else 0.0)
    + float (fontName |> Option.defaultValue "").Length * 0.01
let paintCacheRows = 2000
let paintCacheCols = 50
let paintCacheCells =
    [ for row in 0 .. paintCacheRows - 1 do
        for col in 0 .. paintCacheCols - 1 do
            { Row = row; Col = col }, cell (CellValue.String (sprintf "R%dC%d" row col)) ]
    |> Map.ofList
let paintCacheModel = { createEmpty paintCacheRows paintCacheCols with Cells = paintCacheCells }
let paintCacheViewport : RedBat.Spreadsheet.Layout.Viewport =
    { ScrollX = 0.0
      ScrollY = 0.0
      Width = 900.0
      Height = 600.0
      HeaderHeight = 20.0
      RowHeaderWidth = 28.0
      Measurer = Some stablePaintMeasure }
let firstPaintItems = RedBat.Spreadsheet.Paint.buildViewportPaintItems paintCacheModel paintCacheViewport
let secondPaintItems = RedBat.Spreadsheet.Paint.buildViewportPaintItems paintCacheModel paintCacheViewport
check "native paint viewport cache reuses"
    (obj.ReferenceEquals(firstPaintItems, secondPaintItems))
    true
let firstVisibleCells = RedBat.Spreadsheet.Layout.visibleCells paintCacheModel paintCacheViewport
let secondVisibleCells = RedBat.Spreadsheet.Layout.visibleCells paintCacheModel paintCacheViewport
let firstHeaders = RedBat.Spreadsheet.Layout.headers paintCacheModel paintCacheViewport
let secondHeaders = RedBat.Spreadsheet.Layout.headers paintCacheModel paintCacheViewport
check "native viewport geometry cache reuses"
    (obj.ReferenceEquals(firstVisibleCells, secondVisibleCells)
     && obj.ReferenceEquals(firstHeaders, secondHeaders))
    true

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let paintCacheAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let paintCacheSw = Stopwatch.StartNew()
let mutable paintCacheTotal = 0
for _ in 1 .. 40 do
    paintCacheTotal <- paintCacheTotal + (RedBat.Spreadsheet.Paint.buildViewportPaintItems paintCacheModel paintCacheViewport).Length
paintCacheSw.Stop()
let paintCacheAllocated = System.GC.GetAllocatedBytesForCurrentThread() - paintCacheAllocatedBefore
check "native paint viewport cache perf"
    (paintCacheSw.Elapsed.TotalMilliseconds < 20.0
     && paintCacheAllocated < 1024L * 1024L
     && paintCacheTotal = firstPaintItems.Length * 40)
    true

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let geometryHotHitAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let geometryHotHitSw = Stopwatch.StartNew()
let mutable geometryHotHitTotal = 0
for _ in 1 .. 50000 do
    geometryHotHitTotal <-
        geometryHotHitTotal
        + (RedBat.Spreadsheet.Layout.visibleCells paintCacheModel paintCacheViewport).Length
        + (RedBat.Spreadsheet.Layout.headers paintCacheModel paintCacheViewport).Length
geometryHotHitSw.Stop()
let geometryHotHitAllocated = System.GC.GetAllocatedBytesForCurrentThread() - geometryHotHitAllocatedBefore
check "native viewport geometry cache hot hits"
    (geometryHotHitSw.Elapsed.TotalMilliseconds < 140.0
     && geometryHotHitAllocated < 512L * 1024L
     && geometryHotHitTotal = (firstVisibleCells.Length + firstHeaders.Length) * 50000)
    true

let prefixCoordinateRows = 10000
let prefixCoordinateModel =
    { createEmpty 120000 120000 with
        RowHeights =
            [ for row in 0 .. prefixCoordinateRows - 1 -> row, 20.0 + float (row % 7) ]
            |> Map.ofList
        ColumnWidths =
            [ for col in 0 .. prefixCoordinateRows - 1 -> col, 80.0 + float (col % 5) ]
            |> Map.ofList }
RedBat.Spreadsheet.Layout.rowTop prefixCoordinateModel 80000 |> ignore
RedBat.Spreadsheet.Layout.columnLeft prefixCoordinateModel 80000 |> ignore
let prefixCoordinateHotHitRun count =
    let mutable total = 0.0
    for index in 1 .. count do
        let offset = index % 1000
        total <-
            total
            + RedBat.Spreadsheet.Layout.rowTop prefixCoordinateModel (70000 + offset)
            + RedBat.Spreadsheet.Layout.columnLeft prefixCoordinateModel (70000 + offset)
    total
prefixCoordinateHotHitRun 1000 |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let prefixCoordinateAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let prefixCoordinateSw = Stopwatch.StartNew()
let prefixCoordinateTotal = prefixCoordinateHotHitRun 200000
prefixCoordinateSw.Stop()
let prefixCoordinateAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - prefixCoordinateAllocatedBefore
check "native prefix coordinate cache hot hits"
    (prefixCoordinateSw.Elapsed.TotalMilliseconds < 140.0
     && prefixCoordinateAllocated < 1024L * 1024L
     && prefixCoordinateTotal > 0.0)
    true

let shallowPrefixCoordinateModel =
    { createEmpty 120000 120000 with
        RowHeights =
            [ for row in 0 .. 9999 -> row, 18.0 + float (row % 11) ]
            |> Map.ofList
        ColumnWidths =
            [ for col in 0 .. 9999 -> col, 76.0 + float (col % 13) ]
            |> Map.ofList }
let shallowPrefixCoordinateRun count =
    let mutable total = 0.0
    for index in 1 .. count do
        let offset = index % 250
        let coord = 3000 + offset
        total <-
            total
            + RedBat.Spreadsheet.Layout.rowTop shallowPrefixCoordinateModel coord
            + RedBat.Spreadsheet.Layout.columnLeft shallowPrefixCoordinateModel coord
    total
shallowPrefixCoordinateRun 1000 |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let shallowPrefixCoordinateAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let shallowPrefixCoordinateSw = Stopwatch.StartNew()
let shallowPrefixCoordinateTotal = shallowPrefixCoordinateRun 10000
shallowPrefixCoordinateSw.Stop()
let shallowPrefixCoordinateAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - shallowPrefixCoordinateAllocatedBefore
check "native prefix coordinate shallow cache"
    (shallowPrefixCoordinateSw.Elapsed.TotalMilliseconds < 40.0
     && shallowPrefixCoordinateAllocated < 128L * 1024L
     && shallowPrefixCoordinateTotal > 0.0)
    true

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let paintHotHitAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let paintHotHitSw = Stopwatch.StartNew()
let mutable paintHotHitTotal = 0
for _ in 1 .. 50000 do
    paintHotHitTotal <- paintHotHitTotal + (RedBat.Spreadsheet.Paint.buildViewportPaintItems paintCacheModel paintCacheViewport).Length
paintHotHitSw.Stop()
let paintHotHitAllocated = System.GC.GetAllocatedBytesForCurrentThread() - paintHotHitAllocatedBefore
check "native paint viewport cache hot hits"
    (paintHotHitSw.Elapsed.TotalMilliseconds < 160.0
     && paintHotHitAllocated < 3L * 1024L * 1024L
     && paintHotHitTotal = firstPaintItems.Length * 50000)
    true

let sparsePaintCells =
    [ for row in 0 .. 1999 ->
        { Row = row; Col = 0 }, cell (CellValue.String (sprintf "R%d" row)) ]
    |> Map.ofList
let sparsePaintModel = { createEmpty 20000 50 with Cells = sparsePaintCells }
let sparsePaintPartsA =
    RedBat.Spreadsheet.Paint.tryBuildSparseDefaultPaintParts sparsePaintModel paintCacheViewport
let sparsePaintPartsB =
    RedBat.Spreadsheet.Paint.tryBuildSparseDefaultPaintParts sparsePaintModel paintCacheViewport
let sparsePaintPartsItemCount (parts: RedBat.Spreadsheet.Paint.SparseDefaultPaintParts) =
    parts.GridItems.Length + parts.TextItems.Length
let sparsePaintPartsReuse =
    match sparsePaintPartsA, sparsePaintPartsB with
    | Some partsA, Some partsB ->
        obj.ReferenceEquals(sparsePaintPartsA, sparsePaintPartsB)
        && obj.ReferenceEquals(partsA, partsB)
    | _ -> false
check "native sparse default paint parts reuses"
    sparsePaintPartsReuse
    true

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let sparsePaintPartsAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let sparsePaintPartsSw = Stopwatch.StartNew()
let mutable sparsePaintPartsTotal = 0
for _ in 1 .. 100000 do
    match RedBat.Spreadsheet.Paint.tryBuildSparseDefaultPaintParts sparsePaintModel paintCacheViewport with
    | Some parts ->
        sparsePaintPartsTotal <- sparsePaintPartsTotal + sparsePaintPartsItemCount parts
    | None -> ()
sparsePaintPartsSw.Stop()
let sparsePaintPartsAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - sparsePaintPartsAllocatedBefore
check "native sparse default paint parts hot hits"
    (sparsePaintPartsSw.Elapsed.TotalMilliseconds < 80.0
     && sparsePaintPartsAllocated < 8L * 1024L * 1024L
     && sparsePaintPartsTotal > 0)
    true

let sparsePaintPartsAltViewport = { paintCacheViewport with ScrollY = 420.0 }
let sparsePaintPartsAltA =
    RedBat.Spreadsheet.Paint.tryBuildSparseDefaultPaintParts sparsePaintModel paintCacheViewport
let sparsePaintPartsAltB =
    RedBat.Spreadsheet.Paint.tryBuildSparseDefaultPaintParts sparsePaintModel sparsePaintPartsAltViewport
let sparsePaintPartsAltExpected =
    match sparsePaintPartsAltA, sparsePaintPartsAltB with
    | Some partsA, Some partsB -> Some (sparsePaintPartsItemCount partsA + sparsePaintPartsItemCount partsB)
    | _ -> None

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let sparsePaintPartsAltAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let sparsePaintPartsAltSw = Stopwatch.StartNew()
let mutable sparsePaintPartsAltTotal = 0
for index in 1 .. 200000 do
    let viewport = if index % 2 = 0 then paintCacheViewport else sparsePaintPartsAltViewport
    match RedBat.Spreadsheet.Paint.tryBuildSparseDefaultPaintParts sparsePaintModel viewport with
    | Some parts ->
        sparsePaintPartsAltTotal <- sparsePaintPartsAltTotal + sparsePaintPartsItemCount parts
    | None -> ()
sparsePaintPartsAltSw.Stop()
let sparsePaintPartsAltAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - sparsePaintPartsAltAllocatedBefore
check "native sparse default paint parts alternates"
    (sparsePaintPartsAltSw.Elapsed.TotalMilliseconds < 120.0
     && sparsePaintPartsAltAllocated < 16L * 1024L * 1024L
     && (match sparsePaintPartsAltExpected with
         | Some expected -> sparsePaintPartsAltTotal = expected * 100000
         | None -> false))
    true

let sparsePaintPartsCycleViewports =
    [| paintCacheViewport
       sparsePaintPartsAltViewport
       { paintCacheViewport with ScrollY = 840.0 } |]
let sparsePaintPartsCycleCounts =
    sparsePaintPartsCycleViewports
    |> Array.map (fun viewport ->
        RedBat.Spreadsheet.Paint.tryBuildSparseDefaultPaintParts sparsePaintModel viewport
        |> Option.map sparsePaintPartsItemCount)
let mutable sparsePaintPartsCycleExpectedTotal = 0
for index in 1 .. 200000 do
    match sparsePaintPartsCycleCounts.[index % sparsePaintPartsCycleCounts.Length] with
    | Some count -> sparsePaintPartsCycleExpectedTotal <- sparsePaintPartsCycleExpectedTotal + count
    | None -> ()

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let sparsePaintPartsCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let sparsePaintPartsCycleSw = Stopwatch.StartNew()
let mutable sparsePaintPartsCycleTotal = 0
for index in 1 .. 200000 do
    match RedBat.Spreadsheet.Paint.tryBuildSparseDefaultPaintParts sparsePaintModel sparsePaintPartsCycleViewports.[index % sparsePaintPartsCycleViewports.Length] with
    | Some parts ->
        sparsePaintPartsCycleTotal <- sparsePaintPartsCycleTotal + sparsePaintPartsItemCount parts
    | None -> ()
sparsePaintPartsCycleSw.Stop()
let sparsePaintPartsCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - sparsePaintPartsCycleAllocatedBefore
check "native sparse default paint parts three-position cycle"
    (sparsePaintPartsCycleSw.Elapsed.TotalMilliseconds < 140.0
     && sparsePaintPartsCycleAllocated < 16L * 1024L * 1024L
     && (sparsePaintPartsCycleCounts |> Array.forall Option.isSome)
     && sparsePaintPartsCycleTotal = sparsePaintPartsCycleExpectedTotal
     && sparsePaintPartsCycleTotal > 0)
    true

if isNull usedBoundsCached then
    fail <- fail + 1
    printfn "FAIL %-38s expected editor usedBoundsCached helper" "native used bounds reflection"
else
    let usedBoundsRows = 520000
    let usedBoundsCells =
        [ for row in 0 .. usedBoundsRows - 1 ->
            { Row = row; Col = row % 12 }, cell (CellValue.Number (float row)) ]
        |> Map.ofList
    let usedBoundsModel = { createEmpty usedBoundsRows 20 with Cells = usedBoundsCells }
    let usedBoundsState = Editor.createState [| usedBoundsModel |] [| "Sheet1" |] 0
    let firstBounds =
        usedBoundsCached.Invoke(null, [| box usedBoundsState; box usedBoundsModel |]) :?> (int * int) option
    check "native used bounds cache value"
        firstBounds
        (Some (usedBoundsRows - 1, 11))

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let usedBoundsAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let usedBoundsSw = Stopwatch.StartNew()
    let mutable usedBoundsTotal = 0
    for _ in 1 .. 20 do
        match usedBoundsCached.Invoke(null, [| box usedBoundsState; box usedBoundsModel |]) :?> (int * int) option with
        | Some (row, col) -> usedBoundsTotal <- usedBoundsTotal + row + col
        | None -> ()
    usedBoundsSw.Stop()
    let usedBoundsAllocated = System.GC.GetAllocatedBytesForCurrentThread() - usedBoundsAllocatedBefore
    check "native used bounds cache perf"
        (usedBoundsSw.Elapsed.TotalMilliseconds < 20.0
         && usedBoundsAllocated < 1024L * 1024L
         && usedBoundsTotal = (usedBoundsRows - 1 + 11) * 20)
        true

let lastUsedColumnRows = 100000
let lastUsedColumnCells =
    [ for row in 0 .. lastUsedColumnRows - 1 ->
        { Row = row; Col = row % 50 }, cell (CellValue.Number (float row)) ]
    |> Map.ofList
let lastUsedColumnModel = { createEmpty lastUsedColumnRows 60 with Cells = lastUsedColumnCells }
let lastUsedColumnQueryRows = [| for row in 1000 .. 1199 -> row |]
for row in lastUsedColumnQueryRows do
    lastUsedColumnInRow row lastUsedColumnModel |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let lastUsedColumnAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let lastUsedColumnSw = Stopwatch.StartNew()
let mutable lastUsedColumnTotal = 0
for index in 1 .. 500000 do
    match lastUsedColumnInRow lastUsedColumnQueryRows.[index % lastUsedColumnQueryRows.Length] lastUsedColumnModel with
    | Some col -> lastUsedColumnTotal <- lastUsedColumnTotal + col
    | None -> lastUsedColumnTotal <- lastUsedColumnTotal - 1
lastUsedColumnSw.Stop()
let lastUsedColumnAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - lastUsedColumnAllocatedBefore
check "native last used column row cycle cache"
    (lastUsedColumnSw.Elapsed.TotalMilliseconds < 80.0
     && lastUsedColumnAllocated < 512L * 1024L
     && lastUsedColumnTotal = 12250000)
    true

let filterCacheCells =
    [ { Row = 0; Col = 0 }, cell (CellValue.String "header")
      { Row = 1; Col = 0 }, cell (CellValue.String "keep")
      { Row = 2; Col = 0 }, cell (CellValue.String "hide")
      { Row = 3; Col = 0 }, cell (CellValue.String "keep")
      { Row = 4; Col = 1 }, cell (CellValue.String "other") ]
    |> Map.ofList

let filterCacheModel =
    { createEmpty 5 2 with Cells = filterCacheCells }
    |> applyFilter 0 (Set.ofList [ "keep" ])

let filterCacheProjected = applyFilterRowHeights filterCacheModel
let filterCacheProjectedAgain = applyFilterRowHeights filterCacheProjected
check "native filter hidden rows value"
    (filterCacheProjected.FilterHiddenRows |> Array.toList)
    [ 2; 4 ]
check "native filter hidden rows reuses"
    (obj.ReferenceEquals(filterCacheProjected.FilterHiddenRows, filterCacheProjectedAgain.FilterHiddenRows))
    true

let filterCacheVisibleRows = getVisibleRows filterCacheProjected
let filterCacheVisibleRowsAgain = getVisibleRows filterCacheProjected
check "native visible rows cache value"
    filterCacheVisibleRows
    (Set.ofList [ 1; 3 ])
check "native visible rows cache reuses"
    (obj.ReferenceEquals(filterCacheVisibleRows, filterCacheVisibleRowsAgain))
    true

let searchCacheCells =
    [ { Row = 0; Col = 0 }, cell (CellValue.String "header")
      { Row = 1; Col = 0 }, cell (CellValue.String "needle")
      { Row = 2; Col = 0 }, cell (CellValue.String "other")
      { Row = 3; Col = 1 }, cell (CellValue.String "Needle case") ]
    |> Map.ofList

let searchCacheModel = { createEmpty 4 2 with Cells = searchCacheCells }
let searchCacheResults = searchCells "needle" false searchCacheModel
let searchCacheResultsAgain = searchCells "needle" false searchCacheModel
check "native search cache value"
    (searchCacheResults |> Array.toList)
    [ { Row = 1; Col = 0 }; { Row = 3; Col = 1 } ]
check "native search cache reuses"
    (obj.ReferenceEquals(searchCacheResults, searchCacheResultsAgain))
    true

let uniqueValuesCells =
    [ { Row = 0; Col = 0 }, cell (CellValue.String "red")
      { Row = 1; Col = 0 }, cell (CellValue.String "blue")
      { Row = 2; Col = 0 }, cell (CellValue.String "")
      { Row = 3; Col = 0 }, cell (CellValue.String "red") ]
    |> Map.ofList

let uniqueValuesModel = { createEmpty 4 1 with Cells = uniqueValuesCells }
let uniqueValues = getColumnUniqueValues 0 uniqueValuesModel
let uniqueValuesAgain = getColumnUniqueValues 0 uniqueValuesModel
check "native column unique values cache value"
    uniqueValues
    [ ""; "blue"; "red" ]
check "native column unique values reuses"
    (obj.ReferenceEquals(uniqueValues, uniqueValuesAgain))
    true
columnUniqueValues 0 uniqueValuesModel |> ignore
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let uniqueValuesAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let uniqueValuesSw = Stopwatch.StartNew()
let mutable uniqueValuesTotal = 0
for _pass in 1 .. 200000 do
    uniqueValuesTotal <- uniqueValuesTotal + (getColumnUniqueValues 0 uniqueValuesModel).Length
    uniqueValuesTotal <- uniqueValuesTotal + (columnUniqueValues 0 uniqueValuesModel).Length
uniqueValuesSw.Stop()
let uniqueValuesAllocated = System.GC.GetAllocatedBytesForCurrentThread() - uniqueValuesAllocatedBefore
check "native column unique values hot hit"
    (uniqueValuesSw.Elapsed.TotalMilliseconds < 250.0
     && uniqueValuesAllocated < 2L * 1024L * 1024L
     && uniqueValuesTotal = 1000000)
    true

let tableUniqueValuesCells =
    [ { Row = 0; Col = 0 }, cell (CellValue.String "Color")
      { Row = 1; Col = 0 }, cell (CellValue.String "red")
      { Row = 2; Col = 0 }, cell (CellValue.String "blue")
      { Row = 3; Col = 0 }, cell (CellValue.String "red") ]
    |> Map.ofList
let tableUniqueValuesTable =
    { Name = "Colors"
      StartRow = 1
      StartCol = 1
      EndRow = 4
      EndCol = 1
      HeaderRowCount = 1
      Style = None
      AutoFilter = None
      HasTotals = false }
let tableUniqueValuesModel = { createEmpty 4 1 with Cells = tableUniqueValuesCells; Tables = [ tableUniqueValuesTable ] }
let tableUniqueValues = tableColumnUniqueValues tableUniqueValuesTable 0 tableUniqueValuesModel
let tableUniqueValuesAgain = tableColumnUniqueValues tableUniqueValuesTable 0 tableUniqueValuesModel
check "native table unique values cache value"
    tableUniqueValues
    [ "blue"; "red" ]
check "native table unique values reuses"
    (obj.ReferenceEquals(tableUniqueValues, tableUniqueValuesAgain))
    true

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let tableUniqueValuesAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let tableUniqueValuesSw = Stopwatch.StartNew()
let mutable tableUniqueValuesTotal = 0
for _pass in 1 .. 200000 do
    tableUniqueValuesTotal <- tableUniqueValuesTotal + (tableColumnUniqueValues tableUniqueValuesTable 0 tableUniqueValuesModel).Length
tableUniqueValuesSw.Stop()
let tableUniqueValuesAllocated = System.GC.GetAllocatedBytesForCurrentThread() - tableUniqueValuesAllocatedBefore
check "native table unique values hot hit"
    (tableUniqueValuesSw.Elapsed.TotalMilliseconds < 80.0
     && tableUniqueValuesAllocated < 512L * 1024L
     && tableUniqueValuesTotal = tableUniqueValues.Length * 200000)
    true

let sortIdentityCells =
    [ { Row = 0; Col = 0 }, cell (CellValue.Number 1.0)
      { Row = 1; Col = 0 }, cell (CellValue.Number 2.0)
      { Row = 2; Col = 0 }, cell (CellValue.Number 3.0) ]
    |> Map.ofList

let sortIdentityModel = { createEmpty 3 1 with Cells = sortIdentityCells }
let sortIdentitySorted = sortByColumn 0 true sortIdentityModel
let sortIdentitySortedAgain = sortByColumn 0 true sortIdentitySorted
check "native sort identity cache value"
    sortIdentitySorted.SortState
    (Some { Column = 0; Ascending = true })
check "native sort identity cache reuses"
    (obj.ReferenceEquals(sortIdentitySorted, sortIdentitySortedAgain))
    true

let exportCsvCells =
    [ { Row = 0; Col = 0 }, cell (CellValue.String "Name")
      { Row = 0; Col = 1 }, cell (CellValue.String "Note")
      { Row = 1; Col = 0 }, cell (CellValue.String "Ada")
      { Row = 1; Col = 1 }, cell (CellValue.String "comma,quote\"") ]
    |> Map.ofList

let exportCsvModel = { createEmpty 2 2 with Cells = exportCsvCells }
let exportCsvText = exportCSV exportCsvModel
let exportCsvTextAgain = exportCSV exportCsvModel
check "native export csv cache value"
    exportCsvText
    "Name,Note\nAda,\"comma,quote\"\"\""
check "native export csv cache reuses"
    (obj.ReferenceEquals(exportCsvText, exportCsvTextAgain))
    true

let importCsvText = "Name,Score\nAda,42\nBob,17"
let importCsvImported = importCSV importCsvText 0 0 (createEmpty 3 2)
let importCsvImportedAgain = importCSV importCsvText 0 0 importCsvImported
check "native import csv cache reuses"
    (obj.ReferenceEquals(importCsvImported, importCsvImportedAgain))
    true

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let importCsvAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let importCsvSw = Stopwatch.StartNew()
let mutable importCsvSame = 0
for _pass in 1 .. 200000 do
    if obj.ReferenceEquals(importCSV importCsvText 0 0 importCsvImported, importCsvImported) then
        importCsvSame <- importCsvSame + 1
importCsvSw.Stop()
let importCsvAllocated = System.GC.GetAllocatedBytesForCurrentThread() - importCsvAllocatedBefore
check "native import csv hot hit"
    (importCsvSw.Elapsed.TotalMilliseconds < 150.0
     && importCsvAllocated < 512L * 1024L
     && importCsvSame = 200000)
    true

let formulaPresenceCells =
    [ { Row = 0; Col = 0 }, cell (CellValue.Number 1.0)
      { Row = 1; Col = 0 }, cell (CellValue.String "plain") ]
    |> Map.ofList

check "native formula presence cache value"
    (cellsHaveFormulaCells formulaPresenceCells)
    false
check "native formula presence cache count"
    (cellsFormulaCountIfKnown formulaPresenceCells)
    (Some 0)

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let formulaPresenceProbeCount = 200000
let formulaPresenceAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let formulaPresenceSw = Stopwatch.StartNew()
let mutable formulaPresenceTotal = 0
for _ in 1 .. formulaPresenceProbeCount do
    if cellsHaveFormulaCells formulaPresenceCells then
        formulaPresenceTotal <- formulaPresenceTotal + 1
    match cellsFormulaCountIfKnown formulaPresenceCells with
    | Some count -> formulaPresenceTotal <- formulaPresenceTotal + count
    | None -> formulaPresenceTotal <- formulaPresenceTotal - 1
formulaPresenceSw.Stop()
let formulaPresenceAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - formulaPresenceAllocatedBefore
check "native formula presence cache hot hits"
    (formulaPresenceSw.Elapsed.TotalMilliseconds < 80.0
     && formulaPresenceAllocated < 1024L * 1024L
     && formulaPresenceTotal = 0)
    true

let conditionalFormatRows = 20000
let conditionalFormatCells =
    [ for row in 0 .. conditionalFormatRows - 1 ->
        { Row = row; Col = 0 }, cell (CellValue.Number (float row)) ]
    |> Map.ofList
let conditionalFormatModel = { createEmpty conditionalFormatRows 4 with Cells = conditionalFormatCells }
let conditionalDataBarModel =
    applyDataBars 0 0 (conditionalFormatRows - 1) 0 "#63C384" conditionalFormatModel
let conditionalColorScaleModel =
    applyColorScale 0 0 (conditionalFormatRows - 1) 0 "#FF0000" None "#00FF00" conditionalFormatModel
let conditionalIconSetModel =
    applyIconSet 0 0 (conditionalFormatRows - 1) 0 3 conditionalFormatModel
check "native data bar no-op identity"
    (obj.ReferenceEquals(
        applyDataBars 0 0 (conditionalFormatRows - 1) 0 "#63C384" conditionalDataBarModel,
        conditionalDataBarModel))
    true
check "native color scale no-op identity"
    (obj.ReferenceEquals(
        applyColorScale 0 0 (conditionalFormatRows - 1) 0 "#FF0000" None "#00FF00" conditionalColorScaleModel,
        conditionalColorScaleModel))
    true
check "native icon set no-op identity"
    (obj.ReferenceEquals(
        applyIconSet 0 0 (conditionalFormatRows - 1) 0 3 conditionalIconSetModel,
        conditionalIconSetModel))
    true

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let conditionalFormatAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let conditionalFormatSw = Stopwatch.StartNew()
let mutable conditionalFormatHits = 0
for _ in 1 .. 100000 do
    if obj.ReferenceEquals(
        applyDataBars 0 0 (conditionalFormatRows - 1) 0 "#63C384" conditionalDataBarModel,
        conditionalDataBarModel) then
        conditionalFormatHits <- conditionalFormatHits + 1
    if obj.ReferenceEquals(
        applyColorScale 0 0 (conditionalFormatRows - 1) 0 "#FF0000" None "#00FF00" conditionalColorScaleModel,
        conditionalColorScaleModel) then
        conditionalFormatHits <- conditionalFormatHits + 1
    if obj.ReferenceEquals(
        applyIconSet 0 0 (conditionalFormatRows - 1) 0 3 conditionalIconSetModel,
        conditionalIconSetModel) then
        conditionalFormatHits <- conditionalFormatHits + 1
conditionalFormatSw.Stop()
let conditionalFormatAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - conditionalFormatAllocatedBefore
check "native conditional format no-op hot hits"
    (conditionalFormatSw.Elapsed.TotalMilliseconds < 80.0
     && conditionalFormatAllocated < 4L * 1024L * 1024L
     && conditionalFormatHits = 300000)
    true

let conditionalNoNumberCells =
    [ for row in 0 .. conditionalFormatRows - 1 ->
        { Row = row; Col = 0 }, cell (CellValue.String (sprintf "text-%d" row)) ]
    |> Map.ofList
let conditionalNoNumberModel = { createEmpty conditionalFormatRows 4 with Cells = conditionalNoNumberCells }
applyDataBars 0 0 (conditionalFormatRows - 1) 0 "#63C384" conditionalNoNumberModel |> ignore
applyDataBars 0 0 (conditionalFormatRows - 1) 0 "#FF0000" conditionalNoNumberModel |> ignore
applyColorScale 0 0 (conditionalFormatRows - 1) 0 "#FF0000" None "#00FF00" conditionalNoNumberModel |> ignore
applyColorScale 0 0 (conditionalFormatRows - 1) 0 "#0000FF" None "#00FF00" conditionalNoNumberModel |> ignore
applyIconSet 0 0 (conditionalFormatRows - 1) 0 3 conditionalNoNumberModel |> ignore
applyIconSet 0 0 (conditionalFormatRows - 1) 0 4 conditionalNoNumberModel |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let conditionalNoNumberAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let conditionalNoNumberSw = Stopwatch.StartNew()
let mutable conditionalNoNumberHits = 0
for index in 1 .. 100000 do
    let dataBarColor = if index % 2 = 0 then "#63C384" else "#FF0000"
    if obj.ReferenceEquals(
        applyDataBars 0 0 (conditionalFormatRows - 1) 0 dataBarColor conditionalNoNumberModel,
        conditionalNoNumberModel) then
        conditionalNoNumberHits <- conditionalNoNumberHits + 1
    let scaleLowColor = if index % 2 = 0 then "#FF0000" else "#0000FF"
    if obj.ReferenceEquals(
        applyColorScale 0 0 (conditionalFormatRows - 1) 0 scaleLowColor None "#00FF00" conditionalNoNumberModel,
        conditionalNoNumberModel) then
        conditionalNoNumberHits <- conditionalNoNumberHits + 1
    let iconCount = if index % 2 = 0 then 3 else 4
    if obj.ReferenceEquals(
        applyIconSet 0 0 (conditionalFormatRows - 1) 0 iconCount conditionalNoNumberModel,
        conditionalNoNumberModel) then
        conditionalNoNumberHits <- conditionalNoNumberHits + 1
conditionalNoNumberSw.Stop()
let conditionalNoNumberAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - conditionalNoNumberAllocatedBefore
check "native conditional format alternating no-number cache"
    (conditionalNoNumberSw.Elapsed.TotalMilliseconds < 80.0
     && conditionalNoNumberAllocated < 512L * 1024L
     && conditionalNoNumberHits = 300000)
    true

let conditionalNoNumberDataBarColors = [| "#63C384"; "#FF0000"; "#0000FF" |]
let conditionalNoNumberScaleLowColors = [| "#FF0000"; "#0000FF"; "#63C384" |]
let conditionalNoNumberIconCounts = [| 3; 4; 5 |]
for color in conditionalNoNumberDataBarColors do
    applyDataBars 0 0 (conditionalFormatRows - 1) 0 color conditionalNoNumberModel |> ignore
for color in conditionalNoNumberScaleLowColors do
    applyColorScale 0 0 (conditionalFormatRows - 1) 0 color None "#00FF00" conditionalNoNumberModel |> ignore
for count in conditionalNoNumberIconCounts do
    applyIconSet 0 0 (conditionalFormatRows - 1) 0 count conditionalNoNumberModel |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let conditionalNoNumberCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let conditionalNoNumberCycleSw = Stopwatch.StartNew()
let mutable conditionalNoNumberCycleHits = 0
for index in 1 .. 100000 do
    let variant = index % 3
    if obj.ReferenceEquals(
        applyDataBars 0 0 (conditionalFormatRows - 1) 0 conditionalNoNumberDataBarColors.[variant] conditionalNoNumberModel,
        conditionalNoNumberModel) then
        conditionalNoNumberCycleHits <- conditionalNoNumberCycleHits + 1
    if obj.ReferenceEquals(
        applyColorScale 0 0 (conditionalFormatRows - 1) 0 conditionalNoNumberScaleLowColors.[variant] None "#00FF00" conditionalNoNumberModel,
        conditionalNoNumberModel) then
        conditionalNoNumberCycleHits <- conditionalNoNumberCycleHits + 1
    if obj.ReferenceEquals(
        applyIconSet 0 0 (conditionalFormatRows - 1) 0 conditionalNoNumberIconCounts.[variant] conditionalNoNumberModel,
        conditionalNoNumberModel) then
        conditionalNoNumberCycleHits <- conditionalNoNumberCycleHits + 1
conditionalNoNumberCycleSw.Stop()
let conditionalNoNumberCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - conditionalNoNumberCycleAllocatedBefore
check "native conditional format no-number three-position cycle"
    (conditionalNoNumberCycleSw.Elapsed.TotalMilliseconds < 60.0
     && conditionalNoNumberCycleAllocated < 512L * 1024L
     && conditionalNoNumberCycleHits = 300000)
    true

let pasteIdentityCells =
    [ { Row = 1; Col = 1 }, cell (CellValue.String "same") ]
    |> Map.ofList
let pasteIdentityModel = { createEmpty 4 4 with Cells = pasteIdentityCells }
let pasteIdentityClipboard = serializeSelection (SingleCell (1, 1)) pasteIdentityModel false
let pasteIdentityFirst = pasteClipboard { Row = 1; Col = 1 } pasteIdentityClipboard pasteIdentityModel
let pasteIdentitySecond = pasteClipboard { Row = 1; Col = 1 } pasteIdentityClipboard pasteIdentityFirst
check "native paste identity keeps model"
    (obj.ReferenceEquals(pasteIdentityModel, pasteIdentityFirst))
    true
check "native paste identity cache reuses"
    (obj.ReferenceEquals(pasteIdentityFirst, pasteIdentitySecond))
    true

if isNull setFocus || isNull extendFocus || isNull startEditing || isNull commitEdit then
    fail <- fail + 1
    printfn "FAIL %-38s expected editor helpers" "native editor reflection"
else
    let cells =
        [ for row in 0 .. 99999 ->
            { Row = row; Col = 0 }, cell (CellValue.Number (float row)) ]
        |> Map.ofList
    let focusState =
        Editor.createState [| { createEmpty 100000 20 with Cells = cells; Selection = SingleCell (0, 0) } |] [| "Sheet1" |] 0
    let initialFocusSheets = focusState.Sheets.Value
    let focusSw = System.Diagnostics.Stopwatch.StartNew()
    for _ in 1 .. 100 do
        setFocus.Invoke(null, [| box focusState; box 0; box 0 |]) |> ignore
    focusSw.Stop()
    check "native focus no-op keeps sheets"
        (obj.ReferenceEquals(initialFocusSheets, focusState.Sheets.Value))
        true
    check "native focus no-op perf"
        (focusSw.Elapsed.TotalMilliseconds < 80.0)
        true

    let rangeState =
        Editor.createState [| { createEmpty 100000 20 with Cells = cells; Selection = Range (0, 0, 10, 5) } |] [| "Sheet1" |] 0
    let initialRangeSheets = rangeState.Sheets.Value
    let rangeSw = System.Diagnostics.Stopwatch.StartNew()
    for _ in 1 .. 100 do
        extendFocus.Invoke(null, [| box rangeState; box 10; box 5 |]) |> ignore
    rangeSw.Stop()
    check "native extend no-op keeps sheets"
        (obj.ReferenceEquals(initialRangeSheets, rangeState.Sheets.Value))
        true
    check "native extend no-op perf"
        (rangeSw.Elapsed.TotalMilliseconds < 80.0)
        true

    let commitState =
        Editor.createState [| { createEmpty 100000 20 with Cells = cells; Selection = SingleCell (50000, 0) } |] [| "Sheet1" |] 0
    startEditing.Invoke(null, [| box commitState; box 50000; box 0; box None |]) |> ignore
    let initialCommitSheets = commitState.Sheets.Value
    let commitSw = Stopwatch.StartNew()
    commitEdit.Invoke(null, [| box commitState |]) |> ignore
    commitSw.Stop()
    check "native commit unchanged keeps sheets"
        (obj.ReferenceEquals(initialCommitSheets, commitState.Sheets.Value))
        true
    check "native commit unchanged avoids undo"
        commitState.UndoStack.Value.Length
        0
    check "native commit unchanged perf"
        (commitSw.Elapsed.TotalMilliseconds < 20.0)
        true

printfn ""
printfn "Spreadsheet native editor cache tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
