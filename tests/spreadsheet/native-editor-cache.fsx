#r "/home/dave/Dev/browser/renderer/bin/Debug/net10.0/Renderer.dll"
#r "/home/dave/Dev/browser/widgets/bin/Debug/net10.0/Widgets.dll"
#r "/home/dave/Dev/browser/office/office-native/bin/Debug/net10.0/OfficeShell.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"
#r "/home/dave/Dev/browser/office/spreadsheet-native/bin/Debug/net10.0/SpreadsheetWidgets.dll"

open RedBat.Spreadsheet.Model
open System.Diagnostics
open System.Reflection

module Editor = Composit.Ui.SpreadsheetEditor

type UsedBoundsCachedDelegate = delegate of Editor.SpreadsheetEditorState * SpreadsheetModel -> (int * int) option

type CellCommandsDelegate = delegate of SpreadsheetModel * float * float -> Renderer.Cmd list

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
let cellCommandsDelegate =
    if isNull cellCommands then
        Unchecked.defaultof<CellCommandsDelegate>
    else
        cellCommands.CreateDelegate(typeof<CellCommandsDelegate>) :?> CellCommandsDelegate
let cellCommandsInViewport =
    gridType.GetMethod("cellCommandsInViewport", staticBindingFlags)
let measureText =
    gridType.GetMethod("measureText", staticBindingFlags)
let measureTextDelegate =
    if isNull measureText then
        null
    else
        measureText.CreateDelegate(typeof<System.Func<string, float, string option, bool, bool, float>>)
        :?> System.Func<string, float, string option, bool, bool, float>
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
let usedBoundsCachedDelegate =
    if isNull usedBoundsCached then
        Unchecked.defaultof<UsedBoundsCachedDelegate>
    else
        usedBoundsCached.CreateDelegate(typeof<UsedBoundsCachedDelegate>) :?> UsedBoundsCachedDelegate
let paintModuleType = typeof<SpreadsheetModel>.Assembly.GetType("RedBat.Spreadsheet.Paint")
let hasAnyRenderTimeTableStyle =
    paintModuleType.GetMethod("hasAnyRenderTimeTableStyle", staticBindingFlags)

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
         && alternatingAllocated < 128L * 1024L
         && alternatingTotal = (rendererCommandsA.Length + rendererCommandsWide.Length) * 20)
        true

    invokeRendererCommands () |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let rendererCommandsHotAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let rendererCommandsHotSw = Stopwatch.StartNew()
    let mutable rendererCommandsHotTotal = 0
    for _ in 1 .. 100000 do
        rendererCommandsHotTotal <-
            rendererCommandsHotTotal
            + cellCommandsDelegate.Invoke(rendererModel, 2400.0, 2400.0).Length
    rendererCommandsHotSw.Stop()
    let rendererCommandsHotAllocated =
        System.GC.GetAllocatedBytesForCurrentThread() - rendererCommandsHotAllocatedBefore
    check "native grid command cache hot hits"
        (rendererCommandsHotSw.Elapsed.TotalMilliseconds < 200.0
         && rendererCommandsHotAllocated < 64L * 1024L
         && rendererCommandsHotTotal = rendererCommandsA.Length * 100000)
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
        (movingSelectionSw.Elapsed.TotalMilliseconds < 80.0
         && movingSelectionAllocated < 512L * 1024L
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
        (viewportSw.Elapsed.TotalMilliseconds < 40.0
         && viewportAllocated < 1024L * 1024L
         && viewportTotal < 10000)
        true

    let sparseFrameRows = 20000
    let sparseFrameCells =
        [ for row in 0 .. 1999 ->
            { Row = row; Col = 0 }, cell (CellValue.Number (float row)) ]
        |> Map.ofList
    let sparseFrameModel = { createEmpty sparseFrameRows 20 with Cells = sparseFrameCells }
    let sparseFrameViewports =
        [| for index in 0 .. 79 ->
            0.0, float index * 420.0, 900.0, 600.0 |]
    let invokeSparseFrameViewport (scrollX, scrollY, viewportWidth, viewportHeight) =
        cellCommandsInViewport.Invoke(
            null,
            [| box sparseFrameModel; box scrollX; box scrollY; box viewportWidth; box viewportHeight |])
        :?> Renderer.Cmd list

    for viewport in sparseFrameViewports do
        invokeSparseFrameViewport viewport |> ignore

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let sparseFrameAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let sparseFrameSw = Stopwatch.StartNew()
    let mutable sparseFrameTotal = 0
    for _pass in 1 .. 20 do
        for viewport in sparseFrameViewports do
            sparseFrameTotal <- sparseFrameTotal + (invokeSparseFrameViewport viewport).Length
    sparseFrameSw.Stop()
    let sparseFrameAllocated = System.GC.GetAllocatedBytesForCurrentThread() - sparseFrameAllocatedBefore
    check "native grid sparse frame entry cycle"
        (sparseFrameSw.Elapsed.TotalMilliseconds < 80.0
         && sparseFrameAllocated < 768L * 1024L
         && sparseFrameTotal = 6440)
        true

    let denseRendererRows = 2000
    let denseRendererCols = 80
    let denseRendererCells =
        [ for row in 0 .. denseRendererRows - 1 do
            for col in 0 .. denseRendererCols - 1 do
                { Row = row; Col = col }, cell (CellValue.Number (float (row + col))) ]
        |> Map.ofList
    let denseRendererModel =
        { createEmpty denseRendererRows denseRendererCols with
            Cells = denseRendererCells
            RichRuns =
                Map.ofList [
                    { Row = 0; Col = 0 },
                    [ "1", { emptyRunOverride with Bold = Some true } ]
                ] }
    let denseRendererViewports =
        [| for index in 0 .. 79 ->
            0.0, float index * 120.0, 900.0, 600.0 |]
    let invokeDenseRendererViewport (scrollX, scrollY, viewportWidth, viewportHeight) =
        cellCommandsInViewport.Invoke(
            null,
            [| box denseRendererModel; box scrollX; box scrollY; box viewportWidth; box viewportHeight |])
        :?> Renderer.Cmd list

    for viewport in denseRendererViewports do
        invokeDenseRendererViewport viewport |> ignore

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let denseRendererAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let denseRendererSw = Stopwatch.StartNew()
    let mutable denseRendererTotal = 0
    for _pass in 1 .. 20 do
        for viewport in denseRendererViewports do
            denseRendererTotal <- denseRendererTotal + (invokeDenseRendererViewport viewport).Length
    denseRendererSw.Stop()
    let denseRendererAllocated = System.GC.GetAllocatedBytesForCurrentThread() - denseRendererAllocatedBefore
    check "native grid renderer command entry cycle"
        (denseRendererSw.Elapsed.TotalMilliseconds < 40.0
         && denseRendererAllocated < 1024L * 1024L
         && denseRendererTotal = 582820)
        true

if isNull measureText then
    fail <- fail + 1
    printfn "FAIL %-38s expected grid measureText helper" "native text measure reflection"
else
    let measureFontName = Some "Calibri"
    let invokeMeasure text =
        measureTextDelegate.Invoke(text, 11.0, measureFontName, false, false)
    let firstWidth = invokeMeasure "Repeated spreadsheet text 12345"

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let mutable measureBestElapsed = System.Double.PositiveInfinity
    let mutable measuredTotal = 0.0
    for _sample in 1 .. 3 do
        let measureSw = Stopwatch.StartNew()
        let mutable sampleTotal = 0.0
        for _ in 1 .. 20000 do
            sampleTotal <- sampleTotal + invokeMeasure "Repeated spreadsheet text 12345"
        measureSw.Stop()
        if measureSw.Elapsed.TotalMilliseconds < measureBestElapsed then
            measureBestElapsed <- measureSw.Elapsed.TotalMilliseconds
            measuredTotal <- sampleTotal
    check "native text measure cache stable"
        (firstWidth > 0.0 && abs (measuredTotal - firstWidth * 20000.0) < 0.01)
        true
    check "native text measure cache perf"
        (measureBestElapsed < 50.0)
        true

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let measureAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let measureAllocSw = Stopwatch.StartNew()
    let mutable measureAllocTotal = 0.0
    for _ in 1 .. 100000 do
        measureAllocTotal <- measureAllocTotal + invokeMeasure "Repeated spreadsheet text 12345"
    measureAllocSw.Stop()
    let measureAllocated = System.GC.GetAllocatedBytesForCurrentThread() - measureAllocatedBefore
    check "native text measure cache alloc"
        (measureAllocSw.Elapsed.TotalMilliseconds < 80.0
         && measureAllocated < 64L * 1024L
         && abs (measureAllocTotal - firstWidth * 100000.0) < 0.01)
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

let geometrySourceCycleCount = 80
let geometrySourceCycleRows = 200
let geometrySourceCycleCols = 10
let geometrySourceCycleModels =
    [| for sourceIndex in 0 .. geometrySourceCycleCount - 1 ->
        let cells =
            [ for row in 0 .. geometrySourceCycleRows - 1 do
                for col in 0 .. geometrySourceCycleCols - 1 do
                    { Row = row; Col = col },
                    cell (CellValue.String (sprintf "S%02dR%03dC%02d" sourceIndex row col)) ]
            |> Map.ofList
        { createEmpty geometrySourceCycleRows geometrySourceCycleCols with Cells = cells } |]
for model in geometrySourceCycleModels do
    RedBat.Spreadsheet.Layout.visibleCells model paintCacheViewport |> ignore
    RedBat.Spreadsheet.Layout.headers model paintCacheViewport |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let geometrySourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let geometrySourceCycleSw = Stopwatch.StartNew()
let mutable geometrySourceCycleTotal = 0
for _pass in 1 .. 20 do
    for model in geometrySourceCycleModels do
        geometrySourceCycleTotal <-
            geometrySourceCycleTotal
            + (RedBat.Spreadsheet.Layout.visibleCells model paintCacheViewport).Length
            + (RedBat.Spreadsheet.Layout.headers model paintCacheViewport).Length
geometrySourceCycleSw.Stop()
let geometrySourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - geometrySourceCycleAllocatedBefore
check "native viewport geometry source cycle"
    (geometrySourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && geometrySourceCycleAllocated < 512L * 1024L
     && geometrySourceCycleTotal = 400000)
    true

let geometryEntryCycleViewports: RedBat.Spreadsheet.Layout.Viewport array =
    [| for index in 0 .. 79 ->
        { paintCacheViewport with
            ScrollX = float (index % 5) * 48.0
            ScrollY = float index * 420.0 } |]
for viewport in geometryEntryCycleViewports do
    RedBat.Spreadsheet.Layout.visibleCells paintCacheModel viewport |> ignore
    RedBat.Spreadsheet.Layout.headers paintCacheModel viewport |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let geometryEntryCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let geometryEntryCycleSw = Stopwatch.StartNew()
let mutable geometryEntryCycleTotal = 0
for _pass in 1 .. 20 do
    for viewport in geometryEntryCycleViewports do
        geometryEntryCycleTotal <-
            geometryEntryCycleTotal
            + (RedBat.Spreadsheet.Layout.visibleCells paintCacheModel viewport).Length
            + (RedBat.Spreadsheet.Layout.headers paintCacheModel viewport).Length
geometryEntryCycleSw.Stop()
let geometryEntryCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - geometryEntryCycleAllocatedBefore
check "native viewport geometry entry cycle"
    (geometryEntryCycleSw.Elapsed.TotalMilliseconds < 80.0
     && geometryEntryCycleAllocated < 512L * 1024L
     && geometryEntryCycleTotal = 445560)
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
    (prefixCoordinateSw.Elapsed.TotalMilliseconds < 45.0
     && prefixCoordinateAllocated < 64L * 1024L
     && prefixCoordinateTotal > 0.0)
    true

let prefixCoordinateCycleModelCount = 80
let prefixCoordinateCycleOverrideCount = 5000
let prefixCoordinateCycleModels =
    [| for modelIndex in 0 .. prefixCoordinateCycleModelCount - 1 ->
        { createEmpty 8000 8000 with
            RowHeights =
                [ for row in 0 .. prefixCoordinateCycleOverrideCount - 1 ->
                    row, 20.0 + float ((row + modelIndex) % 5) ]
                |> Map.ofList
            ColumnWidths =
                [ for col in 0 .. prefixCoordinateCycleOverrideCount - 1 ->
                    col, 90.0 + float ((col + modelIndex) % 7) ]
                |> Map.ofList } |]
for model in prefixCoordinateCycleModels do
    RedBat.Spreadsheet.Layout.rowTop model prefixCoordinateCycleOverrideCount |> ignore
    RedBat.Spreadsheet.Layout.columnLeft model prefixCoordinateCycleOverrideCount |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let prefixCoordinateCyclePasses = 20
let prefixCoordinateCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let prefixCoordinateCycleSw = Stopwatch.StartNew()
let mutable prefixCoordinateCycleTotal = 0.0
for _pass in 1 .. prefixCoordinateCyclePasses do
    for model in prefixCoordinateCycleModels do
        prefixCoordinateCycleTotal <-
            prefixCoordinateCycleTotal
            + RedBat.Spreadsheet.Layout.rowTop model prefixCoordinateCycleOverrideCount
            + RedBat.Spreadsheet.Layout.columnLeft model prefixCoordinateCycleOverrideCount
prefixCoordinateCycleSw.Stop()
let prefixCoordinateCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - prefixCoordinateCycleAllocatedBefore
check "native prefix coordinate source cycle"
    (prefixCoordinateCycleSw.Elapsed.TotalMilliseconds < 80.0
     && prefixCoordinateCycleAllocated < 512L * 1024L
     && abs (prefixCoordinateCycleTotal - 919999820.0) < 0.001)
    true

let hiddenPrefixCoordinateCycleModels =
    [| for modelIndex in 0 .. prefixCoordinateCycleModelCount - 1 ->
        { createEmpty 8000 8000 with
            RowHeights =
                [ for row in 0 .. prefixCoordinateCycleOverrideCount - 1 ->
                    row, 20.0 + float ((row + modelIndex) % 5) ]
                |> Map.ofList
            Filters =
                Map.ofList [
                    0, { Column = 0; VisibleValues = Set.singleton "keep" }
                ]
            FilterHiddenRows =
                [| for row in 1 .. 2 .. prefixCoordinateCycleOverrideCount - 1 -> row |] } |]
for model in hiddenPrefixCoordinateCycleModels do
    RedBat.Spreadsheet.Layout.rowTop model prefixCoordinateCycleOverrideCount |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let hiddenPrefixCoordinateCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let hiddenPrefixCoordinateCycleSw = Stopwatch.StartNew()
let mutable hiddenPrefixCoordinateCycleTotal = 0.0
for _pass in 1 .. prefixCoordinateCyclePasses do
    for model in hiddenPrefixCoordinateCycleModels do
        hiddenPrefixCoordinateCycleTotal <-
            hiddenPrefixCoordinateCycleTotal
            + RedBat.Spreadsheet.Layout.rowTop model prefixCoordinateCycleOverrideCount
hiddenPrefixCoordinateCycleSw.Stop()
let hiddenPrefixCoordinateCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - hiddenPrefixCoordinateCycleAllocatedBefore
check "native hidden prefix coordinate source cycle"
    (hiddenPrefixCoordinateCycleSw.Elapsed.TotalMilliseconds < 80.0
     && hiddenPrefixCoordinateCycleAllocated < 512L * 1024L
     && abs (hiddenPrefixCoordinateCycleTotal - 88000000.0) < 0.001)
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
    (shallowPrefixCoordinateSw.Elapsed.TotalMilliseconds < 3.0
     && shallowPrefixCoordinateAllocated < 64L * 1024L
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

for model in geometrySourceCycleModels do
    RedBat.Spreadsheet.Paint.buildViewportPaintItems model paintCacheViewport |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let paintSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let paintSourceCycleSw = Stopwatch.StartNew()
let mutable paintSourceCycleTotal = 0
for _pass in 1 .. 20 do
    for model in geometrySourceCycleModels do
        paintSourceCycleTotal <-
            paintSourceCycleTotal
            + (RedBat.Spreadsheet.Paint.buildViewportPaintItems model paintCacheViewport).Length
paintSourceCycleSw.Stop()
let paintSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - paintSourceCycleAllocatedBefore
check "native paint viewport source cycle"
    (paintSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && paintSourceCycleAllocated < 512L * 1024L
     && paintSourceCycleTotal = 568000)
    true

for viewport in geometryEntryCycleViewports do
    RedBat.Spreadsheet.Paint.buildViewportPaintItems paintCacheModel viewport |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let paintEntryCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let paintEntryCycleSw = Stopwatch.StartNew()
let mutable paintEntryCycleTotal = 0
for _pass in 1 .. 20 do
    for viewport in geometryEntryCycleViewports do
        paintEntryCycleTotal <-
            paintEntryCycleTotal
            + (RedBat.Spreadsheet.Paint.buildViewportPaintItems paintCacheModel viewport).Length
paintEntryCycleSw.Stop()
let paintEntryCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - paintEntryCycleAllocatedBefore
check "native paint viewport entry cycle"
    (paintEntryCycleSw.Elapsed.TotalMilliseconds < 80.0
     && paintEntryCycleAllocated < 256L * 1024L
     && paintEntryCycleTotal = 618020)
    true

for model in geometrySourceCycleModels do
    RedBat.Spreadsheet.Paint.buildPaintItems model paintCacheViewport |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let staticPaintSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let staticPaintSourceCycleSw = Stopwatch.StartNew()
let mutable staticPaintSourceCycleTotal = 0
for _pass in 1 .. 20 do
    for model in geometrySourceCycleModels do
        staticPaintSourceCycleTotal <-
            staticPaintSourceCycleTotal
            + (RedBat.Spreadsheet.Paint.buildPaintItems model paintCacheViewport).Length
staticPaintSourceCycleSw.Stop()
let staticPaintSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - staticPaintSourceCycleAllocatedBefore
check "native static paint source cycle"
    (staticPaintSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && staticPaintSourceCycleAllocated < 512L * 1024L
     && staticPaintSourceCycleTotal = 568000)
    true

for viewport in geometryEntryCycleViewports do
    RedBat.Spreadsheet.Paint.buildPaintItems paintCacheModel viewport |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let staticPaintEntryCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let staticPaintEntryCycleSw = Stopwatch.StartNew()
let mutable staticPaintEntryCycleTotal = 0
for _pass in 1 .. 20 do
    for viewport in geometryEntryCycleViewports do
        staticPaintEntryCycleTotal <-
            staticPaintEntryCycleTotal
            + (RedBat.Spreadsheet.Paint.buildPaintItems paintCacheModel viewport).Length
staticPaintEntryCycleSw.Stop()
let staticPaintEntryCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - staticPaintEntryCycleAllocatedBefore
check "native static paint entry cycle"
    (staticPaintEntryCycleSw.Elapsed.TotalMilliseconds < 80.0
     && staticPaintEntryCycleAllocated < 256L * 1024L
     && staticPaintEntryCycleTotal = 618020)
    true

if isNull hasAnyRenderTimeTableStyle then
    fail <- fail + 1
    printfn "FAIL %-38s expected private table style cache helper" "native table style cache reflection"
else
    let tableStyleSourceCycleStyle =
        { Name = "TableStyleMedium2"
          ShowFirstColumn = false
          ShowLastColumn = false
          ShowRowStripes = true
          ShowColumnStripes = false
          HeaderFill = Some "#4472c4"
          StripeFill = Some "#d9eaf7" }
    let tableStyleSourceCycleSources =
        [| for sourceIndex in 0 .. 79 ->
            [ for tableIndex in 0 .. 999 ->
                { Name = sprintf "native-table-style-%02d-%04d" sourceIndex tableIndex
                  StartRow = tableIndex + 1
                  StartCol = 1
                  EndRow = tableIndex + 3
                  EndCol = 4
                  HeaderRowCount = 1
                  Style =
                    if tableIndex = 999 then
                        Some tableStyleSourceCycleStyle
                    else
                        None
                  AutoFilter = None
                  HasTotals = false } ] |]

    for tables in tableStyleSourceCycleSources do
        hasAnyRenderTimeTableStyle.Invoke(null, [| box tables |]) |> ignore

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let tableStyleSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let tableStyleSourceCycleSw = Stopwatch.StartNew()
    let mutable tableStyleSourceCycleHits = 0
    for _pass in 1 .. 20 do
        for tables in tableStyleSourceCycleSources do
            if hasAnyRenderTimeTableStyle.Invoke(null, [| box tables |]) :?> bool then
                tableStyleSourceCycleHits <- tableStyleSourceCycleHits + 1
    tableStyleSourceCycleSw.Stop()
    let tableStyleSourceCycleAllocated =
        System.GC.GetAllocatedBytesForCurrentThread() - tableStyleSourceCycleAllocatedBefore
    check "native table style source cycle"
        (tableStyleSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
         && tableStyleSourceCycleAllocated < 512L * 1024L
         && tableStyleSourceCycleHits = 1600)
        true

let tableStyleCellsEntryCycleRowCount = 100
let tableStyleCellsEntryCycleColCount = 10
let tableStyleCellsEntryCycleModelCount = 80
let tableStyleCellsEntryCycleStyle =
    Some {
        Name = "TableStyleMedium2"
        ShowFirstColumn = false
        ShowLastColumn = false
        ShowRowStripes = true
        ShowColumnStripes = false
        HeaderFill = Some "#4472C4"
        StripeFill = Some "#D9E1F2"
    }
let tableStyleCellsEntryCycleTable modelIndex : Table =
    { Name = sprintf "NativeEntryStyle%02d" modelIndex
      StartRow = 1
      StartCol = 1
      EndRow = tableStyleCellsEntryCycleRowCount
      EndCol = tableStyleCellsEntryCycleColCount
      HeaderRowCount = 1
      Style = tableStyleCellsEntryCycleStyle
      AutoFilter = Some (1, tableStyleCellsEntryCycleColCount)
      HasTotals = false }
let tableStyleCellsEntryCycleStyledModels =
    [| for modelIndex in 0 .. tableStyleCellsEntryCycleModelCount - 1 ->
        { createEmpty tableStyleCellsEntryCycleRowCount tableStyleCellsEntryCycleColCount with
            Tables = [ tableStyleCellsEntryCycleTable modelIndex ] }
        |> applyTableStyles |]

for model in tableStyleCellsEntryCycleStyledModels do
    applyTableStyles model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let tableStyleCellsEntryCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let tableStyleCellsEntryCycleSw = Stopwatch.StartNew()
let mutable tableStyleCellsEntryCycleSame = 0
let mutable tableStyleCellsEntryCycleTotal = 0
for _pass in 1 .. 20 do
    for model in tableStyleCellsEntryCycleStyledModels do
        let styled = applyTableStyles model
        if obj.ReferenceEquals(styled.Cells, model.Cells) then
            tableStyleCellsEntryCycleSame <- tableStyleCellsEntryCycleSame + 1
        tableStyleCellsEntryCycleTotal <- tableStyleCellsEntryCycleTotal + styled.Cells.Count
tableStyleCellsEntryCycleSw.Stop()
let tableStyleCellsEntryCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - tableStyleCellsEntryCycleAllocatedBefore
check "native table style cells entry cycle"
    (tableStyleCellsEntryCycleSw.Elapsed.TotalMilliseconds < 80.0
     && tableStyleCellsEntryCycleAllocated < 512L * 1024L
     && tableStyleCellsEntryCycleSame = tableStyleCellsEntryCycleModelCount * 20
     && tableStyleCellsEntryCycleTotal = 816000)
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
     && sparsePaintPartsAllocated < 64L * 1024L
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
     && sparsePaintPartsAltAllocated < 512L * 1024L
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
     && sparsePaintPartsCycleAllocated < 512L * 1024L
     && (sparsePaintPartsCycleCounts |> Array.forall Option.isSome)
     && sparsePaintPartsCycleTotal = sparsePaintPartsCycleExpectedTotal
     && sparsePaintPartsCycleTotal > 0)
    true

let sparsePaintPartsEntryCycleViewports: RedBat.Spreadsheet.Layout.Viewport array =
    [| for index in 0 .. 79 ->
        { paintCacheViewport with
            ScrollX = 0.0
            ScrollY = float index * 420.0 } |]
let sparsePaintPartsEntryCycleCounts =
    sparsePaintPartsEntryCycleViewports
    |> Array.map (fun viewport ->
        RedBat.Spreadsheet.Paint.tryBuildSparseDefaultPaintParts sparsePaintModel viewport
        |> Option.map sparsePaintPartsItemCount)

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let sparsePaintPartsEntryCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let sparsePaintPartsEntryCycleSw = Stopwatch.StartNew()
let mutable sparsePaintPartsEntryCycleTotal = 0
for _pass in 1 .. 20 do
    for viewport in sparsePaintPartsEntryCycleViewports do
        match RedBat.Spreadsheet.Paint.tryBuildSparseDefaultPaintParts sparsePaintModel viewport with
        | Some parts ->
            sparsePaintPartsEntryCycleTotal <- sparsePaintPartsEntryCycleTotal + sparsePaintPartsItemCount parts
        | None -> ()
sparsePaintPartsEntryCycleSw.Stop()
let sparsePaintPartsEntryCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - sparsePaintPartsEntryCycleAllocatedBefore
check "native sparse default paint parts entry cycle"
    (sparsePaintPartsEntryCycleSw.Elapsed.TotalMilliseconds < 80.0
     && sparsePaintPartsEntryCycleAllocated < 2L * 1024L * 1024L
     && (sparsePaintPartsEntryCycleCounts |> Array.forall Option.isSome)
     && sparsePaintPartsEntryCycleTotal = 96920)
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
        usedBoundsCachedDelegate.Invoke(usedBoundsState, usedBoundsModel)
    check "native used bounds cache value"
        firstBounds
        (Some (usedBoundsRows - 1, 11))

    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()

    let usedBoundsAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
    let usedBoundsSw = Stopwatch.StartNew()
    let mutable usedBoundsTotal = 0L
    for _ in 1 .. 100000 do
        match usedBoundsCachedDelegate.Invoke(usedBoundsState, usedBoundsModel) with
        | Some (row, col) -> usedBoundsTotal <- usedBoundsTotal + int64 row + int64 col
        | None -> ()
    usedBoundsSw.Stop()
    let usedBoundsAllocated = System.GC.GetAllocatedBytesForCurrentThread() - usedBoundsAllocatedBefore
    check "native used bounds cache perf"
        (usedBoundsSw.Elapsed.TotalMilliseconds < 80.0
         && usedBoundsAllocated < 64L * 1024L
         && usedBoundsTotal = int64 (usedBoundsRows - 1 + 11) * 100000L)
        true

let usedBoundsCycleSourceCount = 80
let usedBoundsCycleCellCount = 5000
let usedBoundsCycleModels =
    [| for sourceIndex in 0 .. usedBoundsCycleSourceCount - 1 ->
        let cells =
            [ for row in 0 .. usedBoundsCycleCellCount - 1 ->
                let col = if row = usedBoundsCycleCellCount - 1 then 40 + sourceIndex % 7 else sourceIndex % 5
                { Row = row; Col = col }, cell (CellValue.Number (float (row + sourceIndex))) ]
            |> Map.ofList
        { createEmpty usedBoundsCycleCellCount 64 with Cells = cells } |]

for model in usedBoundsCycleModels do
    usedBounds model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let usedBoundsCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let usedBoundsCycleSw = Stopwatch.StartNew()
let mutable usedBoundsCycleTotal = 0
for _pass in 1 .. 20 do
    for model in usedBoundsCycleModels do
        match usedBounds model with
        | Some (row, col) -> usedBoundsCycleTotal <- usedBoundsCycleTotal + row + col
        | None -> usedBoundsCycleTotal <- usedBoundsCycleTotal - 1
usedBoundsCycleSw.Stop()
let usedBoundsCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - usedBoundsCycleAllocatedBefore
check "native used bounds source cycle"
    (usedBoundsCycleSw.Elapsed.TotalMilliseconds < 80.0
     && usedBoundsCycleAllocated < 512L * 1024L
     && usedBoundsCycleTotal = 8067080)
    true

let maxUsedColumnCycleInsertCol = 60
for model in usedBoundsCycleModels do
    insertColumn maxUsedColumnCycleInsertCol model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let maxUsedColumnCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let maxUsedColumnCycleSw = Stopwatch.StartNew()
let mutable maxUsedColumnCycleTotal = 0
let mutable maxUsedColumnCycleSameCells = 0
for _pass in 1 .. 20 do
    for model in usedBoundsCycleModels do
        let updated = insertColumn maxUsedColumnCycleInsertCol model
        maxUsedColumnCycleTotal <- maxUsedColumnCycleTotal + updated.ColCount
        if obj.ReferenceEquals(updated.Cells, model.Cells) then
            maxUsedColumnCycleSameCells <- maxUsedColumnCycleSameCells + 1
maxUsedColumnCycleSw.Stop()
let maxUsedColumnCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - maxUsedColumnCycleAllocatedBefore
check "native max used column source cycle"
    (maxUsedColumnCycleSw.Elapsed.TotalMilliseconds < 80.0
     && maxUsedColumnCycleAllocated < 512L * 1024L
     && maxUsedColumnCycleTotal = 104000
     && maxUsedColumnCycleSameCells = 1600)
    true

let lastUsedColumnSourceCycleQueryRow = usedBoundsCycleCellCount - 1
for model in usedBoundsCycleModels do
    lastUsedColumnInRow lastUsedColumnSourceCycleQueryRow model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let lastUsedColumnSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let lastUsedColumnSourceCycleSw = Stopwatch.StartNew()
let mutable lastUsedColumnSourceCycleTotal = 0
for _pass in 1 .. 20 do
    for model in usedBoundsCycleModels do
        match lastUsedColumnInRow lastUsedColumnSourceCycleQueryRow model with
        | Some col -> lastUsedColumnSourceCycleTotal <- lastUsedColumnSourceCycleTotal + col
        | None -> lastUsedColumnSourceCycleTotal <- lastUsedColumnSourceCycleTotal - 1
lastUsedColumnSourceCycleSw.Stop()
let lastUsedColumnSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - lastUsedColumnSourceCycleAllocatedBefore
check "native last used column source cycle"
    (lastUsedColumnSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && lastUsedColumnSourceCycleAllocated < 512L * 1024L
     && lastUsedColumnSourceCycleTotal = 68680)
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

let lastUsedColumnEntryCollisionMergeCount = 60000
let lastUsedColumnEntryCollisionQueryRows = [| for index in 0 .. 79 -> index * 256 |]
let lastUsedColumnEntryCollisionMerges =
    [ for index in 0 .. lastUsedColumnEntryCollisionMergeCount - 1 ->
        let row = index * 2
        { StartRow = row
          StartCol = index % 8
          EndRow = row + 1
          EndCol = (index % 8) + 12 } ]
let lastUsedColumnEntryCollisionModel =
    { createEmpty (lastUsedColumnEntryCollisionMergeCount * 2 + 10) 32 with
        MergedCells = lastUsedColumnEntryCollisionMerges }

for row in lastUsedColumnEntryCollisionQueryRows do
    lastUsedColumnInRow row lastUsedColumnEntryCollisionModel |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let lastUsedColumnEntryCollisionAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let lastUsedColumnEntryCollisionSw = Stopwatch.StartNew()
let mutable lastUsedColumnEntryCollisionTotal = 0
for _pass in 1 .. 20 do
    for row in lastUsedColumnEntryCollisionQueryRows do
        match lastUsedColumnInRow row lastUsedColumnEntryCollisionModel with
        | Some col -> lastUsedColumnEntryCollisionTotal <- lastUsedColumnEntryCollisionTotal + col
        | None -> lastUsedColumnEntryCollisionTotal <- lastUsedColumnEntryCollisionTotal - 1
lastUsedColumnEntryCollisionSw.Stop()
let lastUsedColumnEntryCollisionAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - lastUsedColumnEntryCollisionAllocatedBefore
check "native last used column entry collision cycle"
    (lastUsedColumnEntryCollisionSw.Elapsed.TotalMilliseconds < 80.0
     && lastUsedColumnEntryCollisionAllocated < 512L * 1024L
     && lastUsedColumnEntryCollisionTotal = 19200)
    true

let mergeRangeCycleMergeCount = 40000
let mergeRangeCycleQueryCount = 80
let mergeRangeCycleMerges =
    [ for index in mergeRangeCycleMergeCount - 1 .. -1 .. 0 ->
        { StartRow = index * 2
          StartCol = index % 12
          EndRow = index * 2
          EndCol = (index % 12) + 1 } ]
let mergeRangeCycleModel =
    { createEmpty (mergeRangeCycleMergeCount * 2 + 10) 20 with MergedCells = mergeRangeCycleMerges }
let mergeRangeCycleQueries =
    [| for index in 0 .. mergeRangeCycleQueryCount - 1 ->
        let startRow = mergeRangeCycleMergeCount + index * 7
        struct (startRow, startRow + 30, 0, 20) |]
for struct (startRow, endRow, startCol, endCol) in mergeRangeCycleQueries do
    mergedCellsIntersectingRange startRow endRow startCol endCol mergeRangeCycleModel |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let mergeRangeCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let mergeRangeCycleSw = Stopwatch.StartNew()
let mutable mergeRangeCycleTotal = 0
for _pass in 1 .. 20 do
    for struct (startRow, endRow, startCol, endCol) in mergeRangeCycleQueries do
        mergeRangeCycleTotal <-
            mergeRangeCycleTotal
            + (mergedCellsIntersectingRange startRow endRow startCol endCol mergeRangeCycleModel).Length
mergeRangeCycleSw.Stop()
let mergeRangeCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - mergeRangeCycleAllocatedBefore
check "native merge range query cycle cache"
    (mergeRangeCycleSw.Elapsed.TotalMilliseconds < 80.0
     && mergeRangeCycleAllocated < 16L * 1024L
     && mergeRangeCycleTotal = 24000)
    true

let mergeSourceCycleCount = 80
let mergeSourceCycleMergeCount = 1000
let mergeSourceCycleModels =
    [| for modelIndex in 0 .. mergeSourceCycleCount - 1 ->
        let startCol = modelIndex % 5
        let merges =
            [ for index in mergeSourceCycleMergeCount - 1 .. -1 .. 0 ->
                let row = index * 2
                { StartRow = row
                  StartCol = startCol
                  EndRow = row
                  EndCol = startCol + 2 } ]
        { createEmpty (mergeSourceCycleMergeCount * 2 + 5) 16 with
            MergedCells = merges },
        { Row = mergeSourceCycleMergeCount; Col = startCol + 1 } |]
for model, lookupAddress in mergeSourceCycleModels do
    mergedCellsIntersectingRange 0 (mergeSourceCycleMergeCount * 2) 0 16 model |> ignore
    findMergedCell lookupAddress model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let mergeSourceCyclePasses = 20
let mergeSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let mergeSourceCycleSw = Stopwatch.StartNew()
let mutable mergeSourceCycleTotal = 0
for _pass in 1 .. mergeSourceCyclePasses do
    for model, lookupAddress in mergeSourceCycleModels do
        mergeSourceCycleTotal <-
            mergeSourceCycleTotal
            + (mergedCellsIntersectingRange 0 (mergeSourceCycleMergeCount * 2) 0 16 model).Length
        match findMergedCell lookupAddress model with
        | Some _ -> mergeSourceCycleTotal <- mergeSourceCycleTotal + 1
        | None -> ()
mergeSourceCycleSw.Stop()
let mergeSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - mergeSourceCycleAllocatedBefore
check "native merge source cycle cache"
    (mergeSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && mergeSourceCycleAllocated < 512L * 1024L
     && mergeSourceCycleTotal = 1601600)
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
let filterCacheProjectedCached = applyFilterRowHeights filterCacheModel
check "native filter hidden rows value"
    (filterCacheProjected.FilterHiddenRows |> Array.toList)
    [ 2; 4 ]
check "native filter hidden rows reuses"
    (obj.ReferenceEquals(filterCacheProjected.FilterHiddenRows, filterCacheProjectedAgain.FilterHiddenRows))
    true
check "native filter projection reuses"
    (obj.ReferenceEquals(filterCacheProjected, filterCacheProjectedCached))
    true

let filterCacheVisibleRows = getVisibleRows filterCacheProjected
let filterCacheVisibleRowsAgain = getVisibleRows filterCacheProjected
check "native visible rows cache value"
    filterCacheVisibleRows
    (Set.ofList [ 1; 3 ])
check "native visible rows cache reuses"
    (obj.ReferenceEquals(filterCacheVisibleRows, filterCacheVisibleRowsAgain))
    true

let filterRowCycleModelCount = 80
let filterRowCycleRowCount = 5000
let filterRowCycleModels =
    [| for modelIndex in 0 .. filterRowCycleModelCount - 1 ->
        { createEmpty filterRowCycleRowCount 1 with
            Cells =
                [ for row in 0 .. filterRowCycleRowCount - 1 ->
                    { Row = row; Col = 0 }, cell (CellValue.String (sprintf "drop-%02d-%04d" modelIndex row)) ]
                |> Map.ofList }
        |> applyFilter 0 (Set.singleton "keep") |]
for model in filterRowCycleModels do
    applyFilterRowHeights model |> ignore
    getVisibleRows model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let filterRowCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let filterRowCycleSw = Stopwatch.StartNew()
let mutable filterRowCycleTotal = 0
for _pass in 1 .. 20 do
    for model in filterRowCycleModels do
        filterRowCycleTotal <- filterRowCycleTotal + (applyFilterRowHeights model).FilterHiddenRows.Length
        filterRowCycleTotal <- filterRowCycleTotal + (getVisibleRows model).Count
filterRowCycleSw.Stop()
let filterRowCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - filterRowCycleAllocatedBefore
check "native filter row source cycle"
    (filterRowCycleSw.Elapsed.TotalMilliseconds < 80.0
     && filterRowCycleAllocated < 128L * 1024L
     && filterRowCycleTotal = 7998400)
    true

let filterEntryCycleFilterCount = 80
let filterEntryCycleCells =
    [ for row in 0 .. filterRowCycleRowCount - 1 ->
        { Row = row; Col = 0 },
        cell (CellValue.String (sprintf "bucket-%02d" (row % filterEntryCycleFilterCount))) ]
    |> Map.ofList
let filterEntryCycleSourceModel =
    { createEmpty filterRowCycleRowCount 1 with Cells = filterEntryCycleCells }
let filterEntryCycleModels =
    [| for index in 0 .. filterEntryCycleFilterCount - 1 ->
        filterEntryCycleSourceModel
        |> applyFilter 0 (Set.singleton (sprintf "bucket-%02d" index)) |]

for model in filterEntryCycleModels do
    applyFilterRowHeights model |> ignore
    getVisibleRows model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let filterEntryCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let filterEntryCycleSw = Stopwatch.StartNew()
let mutable filterEntryCycleTotal = 0
for _pass in 1 .. 20 do
    for model in filterEntryCycleModels do
        filterEntryCycleTotal <- filterEntryCycleTotal + (applyFilterRowHeights model).FilterHiddenRows.Length
        filterEntryCycleTotal <- filterEntryCycleTotal + (getVisibleRows model).Count
filterEntryCycleSw.Stop()
let filterEntryCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - filterEntryCycleAllocatedBefore
check "native filter row entry cycle"
    (filterEntryCycleSw.Elapsed.TotalMilliseconds < 80.0
     && filterEntryCycleAllocated < 128L * 1024L
     && filterEntryCycleTotal = 7998420)
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

let searchEmptyCycleRows = 10000
let searchEmptyCycleCells =
    [ for row in 0 .. searchEmptyCycleRows - 1 ->
        { Row = row; Col = 0 }, cell (CellValue.String (sprintf "value-%05d" row)) ]
    |> Map.ofList
let searchEmptyCycleModel = { createEmpty searchEmptyCycleRows 4 with Cells = searchEmptyCycleCells }
let searchEmptyCycleQueries = [| for index in 0 .. 199 -> sprintf "missing-%03d" index |]
for query in searchEmptyCycleQueries do
    searchCells query false searchEmptyCycleModel |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let searchEmptyCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let searchEmptyCycleSw = Stopwatch.StartNew()
let mutable searchEmptyCycleTotal = 0
for index in 1 .. 1000 do
    searchEmptyCycleTotal <-
        searchEmptyCycleTotal
        + (searchCells searchEmptyCycleQueries.[index % searchEmptyCycleQueries.Length] false searchEmptyCycleModel).Length
searchEmptyCycleSw.Stop()
let searchEmptyCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - searchEmptyCycleAllocatedBefore
check "native search empty query cycle cache"
    (searchEmptyCycleSw.Elapsed.TotalMilliseconds < 80.0
     && searchEmptyCycleAllocated < 512L * 1024L
     && searchEmptyCycleTotal = 0)
    true

let searchSourceCycleCount = 80
let searchSourceCycleRows = 3000
let searchSourceCycleQuery = "missing-needle"
let searchSourceCycleModels =
    [| for sourceIndex in 0 .. searchSourceCycleCount - 1 ->
        let cells =
            [ for row in 0 .. searchSourceCycleRows - 1 ->
                { Row = row; Col = sourceIndex % 4 },
                cell (CellValue.String (sprintf "value-%04d-%04d" sourceIndex row)) ]
            |> Map.ofList
        { createEmpty searchSourceCycleRows 4 with Cells = cells } |]

for model in searchSourceCycleModels do
    searchCells searchSourceCycleQuery false model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let searchSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let searchSourceCycleSw = Stopwatch.StartNew()
let mutable searchSourceCycleTotal = 0
for _pass in 1 .. 20 do
    for model in searchSourceCycleModels do
        searchSourceCycleTotal <-
            searchSourceCycleTotal + (searchCells searchSourceCycleQuery false model).Length
searchSourceCycleSw.Stop()
let searchSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - searchSourceCycleAllocatedBefore
check "native search source cycle cache"
    (searchSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && searchSourceCycleAllocated < 512L * 1024L
     && searchSourceCycleTotal = 0)
    true

let replaceIdentityCycleCount = 80
let replaceIdentityCycleRows = 3000
let replaceIdentityCycleModels =
    [| for sourceIndex in 0 .. replaceIdentityCycleCount - 1 ->
        let cells =
            [ for row in 0 .. replaceIdentityCycleRows - 1 ->
                { Row = row; Col = sourceIndex % 4 },
                cell (CellValue.String "NEEDLE") ]
            |> Map.ofList
        { createEmpty replaceIdentityCycleRows 4 with Cells = cells } |]

for model in replaceIdentityCycleModels do
    replaceAll "needle" "NEEDLE" false model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let replaceIdentityCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let replaceIdentityCycleSw = Stopwatch.StartNew()
let mutable replaceIdentityCycleSame = 0
for _pass in 1 .. 20 do
    for model in replaceIdentityCycleModels do
        if obj.ReferenceEquals(replaceAll "needle" "NEEDLE" false model, model) then
            replaceIdentityCycleSame <- replaceIdentityCycleSame + 1
replaceIdentityCycleSw.Stop()
let replaceIdentityCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - replaceIdentityCycleAllocatedBefore
check "native replace identity source cycle"
    (replaceIdentityCycleSw.Elapsed.TotalMilliseconds < 80.0
     && replaceIdentityCycleAllocated < 16L * 1024L
     && replaceIdentityCycleSame = 1600)
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

let uniqueValuesCycleRows = 250
let uniqueValuesCycleCols = 80
let uniqueValuesCycleCells =
    [ for row in 0 .. uniqueValuesCycleRows - 1 do
        for col in 0 .. uniqueValuesCycleCols - 1 ->
            { Row = row; Col = col }, cell (CellValue.String (sprintf "C%02d-V%02d" col (row % 20))) ]
    |> Map.ofList
let uniqueValuesCycleModel = { createEmpty uniqueValuesCycleRows uniqueValuesCycleCols with Cells = uniqueValuesCycleCells }
for col in 0 .. uniqueValuesCycleCols - 1 do
    columnUniqueValues col uniqueValuesCycleModel |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let uniqueValuesCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let uniqueValuesCycleSw = Stopwatch.StartNew()
let mutable uniqueValuesCycleTotal = 0
for index in 1 .. 800 do
    uniqueValuesCycleTotal <- uniqueValuesCycleTotal + (columnUniqueValues (index % uniqueValuesCycleCols) uniqueValuesCycleModel).Length
uniqueValuesCycleSw.Stop()
let uniqueValuesCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - uniqueValuesCycleAllocatedBefore
check "native column unique values column cycle"
    (uniqueValuesCycleSw.Elapsed.TotalMilliseconds < 80.0
     && uniqueValuesCycleAllocated < 512L * 1024L
     && uniqueValuesCycleTotal = 16000)
    true

let uniqueValuesSourceCycleCount = 80
let uniqueValuesSourceCycleRows = 1000
let uniqueValuesSourceCycleModels =
    [| for sourceIndex in 0 .. uniqueValuesSourceCycleCount - 1 ->
        let targetCol = sourceIndex % 4
        let cells =
            [ for row in 0 .. uniqueValuesSourceCycleRows - 1 ->
                { Row = row; Col = targetCol },
                cell (CellValue.String (sprintf "u%04d" row)) ]
            |> Map.ofList
        { createEmpty uniqueValuesSourceCycleRows 4 with Cells = cells }, targetCol |]

for model, col in uniqueValuesSourceCycleModels do
    columnUniqueValues col model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let uniqueValuesSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let uniqueValuesSourceCycleSw = Stopwatch.StartNew()
let mutable uniqueValuesSourceCycleTotal = 0
for _pass in 1 .. 20 do
    for model, col in uniqueValuesSourceCycleModels do
        uniqueValuesSourceCycleTotal <- uniqueValuesSourceCycleTotal + (columnUniqueValues col model).Length
uniqueValuesSourceCycleSw.Stop()
let uniqueValuesSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - uniqueValuesSourceCycleAllocatedBefore
check "native column unique values source cycle"
    (uniqueValuesSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && uniqueValuesSourceCycleAllocated < 512L * 1024L
     && uniqueValuesSourceCycleTotal = 1600000)
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

let tableUniqueValuesCycleRows = 250
let tableUniqueValuesCycleCols = 80
let tableUniqueValuesCycleCells =
    [ for col in 0 .. tableUniqueValuesCycleCols - 1 do
        yield { Row = 0; Col = col }, cell (CellValue.String (sprintf "Header-%02d" col))
        for row in 1 .. tableUniqueValuesCycleRows - 1 ->
            { Row = row; Col = col }, cell (CellValue.String (sprintf "C%02d-V%02d" col (row % 20))) ]
    |> Map.ofList
let tableUniqueValuesCycleTable =
    { Name = "CycleValues"
      StartRow = 1
      StartCol = 1
      EndRow = tableUniqueValuesCycleRows
      EndCol = tableUniqueValuesCycleCols
      HeaderRowCount = 1
      Style = None
      AutoFilter = None
      HasTotals = false }
let tableUniqueValuesCycleModel =
    { createEmpty tableUniqueValuesCycleRows tableUniqueValuesCycleCols with
        Cells = tableUniqueValuesCycleCells
        Tables = [ tableUniqueValuesCycleTable ] }
for col in 0 .. tableUniqueValuesCycleCols - 1 do
    tableColumnUniqueValues tableUniqueValuesCycleTable col tableUniqueValuesCycleModel |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let tableUniqueValuesCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let tableUniqueValuesCycleSw = Stopwatch.StartNew()
let mutable tableUniqueValuesCycleTotal = 0
for index in 1 .. 800 do
    tableUniqueValuesCycleTotal <-
        tableUniqueValuesCycleTotal
        + (tableColumnUniqueValues tableUniqueValuesCycleTable (index % tableUniqueValuesCycleCols) tableUniqueValuesCycleModel).Length
tableUniqueValuesCycleSw.Stop()
let tableUniqueValuesCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - tableUniqueValuesCycleAllocatedBefore
check "native table unique values column cycle"
    (tableUniqueValuesCycleSw.Elapsed.TotalMilliseconds < 80.0
     && tableUniqueValuesCycleAllocated < 512L * 1024L
     && tableUniqueValuesCycleTotal = 16000)
    true

let tableUniqueValuesSourceCycleCount = 80
let tableUniqueValuesSourceCycleRows = 2000
let tableUniqueValuesSourceCycleModels =
    [| for sourceIndex in 0 .. tableUniqueValuesSourceCycleCount - 1 ->
        let targetCol = sourceIndex % 4
        let cells =
            [ yield { Row = 0; Col = targetCol }, cell (CellValue.String "Header")
              for row in 1 .. tableUniqueValuesSourceCycleRows - 1 ->
                { Row = row; Col = targetCol },
                cell (CellValue.String (sprintf "u%04d" row)) ]
            |> Map.ofList
        let table =
            { Name = sprintf "SourceCycle%d" sourceIndex
              StartRow = 1
              StartCol = 1
              EndRow = tableUniqueValuesSourceCycleRows
              EndCol = 4
              HeaderRowCount = 1
              Style = None
              AutoFilter = None
              HasTotals = false }
        { createEmpty tableUniqueValuesSourceCycleRows 4 with Cells = cells; Tables = [ table ] },
        table,
        targetCol |]

for model, table, col in tableUniqueValuesSourceCycleModels do
    tableColumnUniqueValues table col model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let tableUniqueValuesSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let tableUniqueValuesSourceCycleSw = Stopwatch.StartNew()
let mutable tableUniqueValuesSourceCycleTotal = 0
for _pass in 1 .. 20 do
    for model, table, col in tableUniqueValuesSourceCycleModels do
        tableUniqueValuesSourceCycleTotal <-
            tableUniqueValuesSourceCycleTotal + (tableColumnUniqueValues table col model).Length
tableUniqueValuesSourceCycleSw.Stop()
let tableUniqueValuesSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - tableUniqueValuesSourceCycleAllocatedBefore
check "native table unique values source cycle"
    (tableUniqueValuesSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && tableUniqueValuesSourceCycleAllocated < 512L * 1024L
     && tableUniqueValuesSourceCycleTotal = 3198400)
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

let sortIdentityCycleModelCount = 80
let sortIdentityCycleRowCount = 3000
let sortIdentityCycleModels =
    [| for modelIndex in 0 .. sortIdentityCycleModelCount - 1 ->
        { createEmpty sortIdentityCycleRowCount 2 with
            Cells =
                [ for row in 0 .. sortIdentityCycleRowCount - 1 do
                    yield { Row = row; Col = 0 }, cell (CellValue.Number (float row))
                    yield { Row = row; Col = 1 }, cell (CellValue.String (sprintf "M%02d-R%04d" modelIndex row)) ]
                |> Map.ofList
            SortState = Some { Column = 0; Ascending = true } } |]
for model in sortIdentityCycleModels do
    sortByColumn 0 true model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let sortIdentityCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let sortIdentityCycleSw = Stopwatch.StartNew()
let mutable sortIdentityCycleSame = 0
for _pass in 1 .. 20 do
    for model in sortIdentityCycleModels do
        if obj.ReferenceEquals(sortByColumn 0 true model, model) then
            sortIdentityCycleSame <- sortIdentityCycleSame + 1
sortIdentityCycleSw.Stop()
let sortIdentityCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - sortIdentityCycleAllocatedBefore
check "native sort identity source cycle"
    (sortIdentityCycleSw.Elapsed.TotalMilliseconds < 80.0
     && sortIdentityCycleAllocated < 16L * 1024L
     && sortIdentityCycleSame = sortIdentityCycleModelCount * 20)
    true

let sortIdentityEntryCycleRowCount = 1000
let sortIdentityEntryCycleColumnCount = 80
let sortIdentityEntryCycleCells =
    [ for row in 0 .. sortIdentityEntryCycleRowCount - 1 do
        for col in 0 .. sortIdentityEntryCycleColumnCount - 1 do
            yield { Row = row; Col = col }, cell (CellValue.Number (float row)) ]
    |> Map.ofList
let sortIdentityEntryCycleModels =
    [| for col in 0 .. sortIdentityEntryCycleColumnCount - 1 ->
        { createEmpty sortIdentityEntryCycleRowCount sortIdentityEntryCycleColumnCount with
            Cells = sortIdentityEntryCycleCells
            SortState = Some { Column = col; Ascending = true } } |]

for col in 0 .. sortIdentityEntryCycleColumnCount - 1 do
    sortByColumn col true sortIdentityEntryCycleModels.[col] |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let sortIdentityEntryCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let sortIdentityEntryCycleSw = Stopwatch.StartNew()
let mutable sortIdentityEntryCycleSame = 0
for _pass in 1 .. 20 do
    for col in 0 .. sortIdentityEntryCycleColumnCount - 1 do
        if obj.ReferenceEquals(sortByColumn col true sortIdentityEntryCycleModels.[col], sortIdentityEntryCycleModels.[col]) then
            sortIdentityEntryCycleSame <- sortIdentityEntryCycleSame + 1
sortIdentityEntryCycleSw.Stop()
let sortIdentityEntryCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - sortIdentityEntryCycleAllocatedBefore
check "native sort identity entry cycle"
    (sortIdentityEntryCycleSw.Elapsed.TotalMilliseconds < 80.0
     && sortIdentityEntryCycleAllocated < 16L * 1024L
     && sortIdentityEntryCycleSame = sortIdentityEntryCycleColumnCount * 20)
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

let exportCsvCycleSourceCount = 80
let exportCsvCycleRowCount = 1000
let exportCsvCycleModels =
    [| for sourceIndex in 0 .. exportCsvCycleSourceCount - 1 ->
        let cells =
            [ for row in 0 .. exportCsvCycleRowCount - 1 ->
                let col = if row = exportCsvCycleRowCount - 1 then 10 + sourceIndex % 3 else sourceIndex % 3
                { Row = row; Col = col }, cell (CellValue.Number (float (row + sourceIndex))) ]
            |> Map.ofList
        { createEmpty exportCsvCycleRowCount 16 with Cells = cells } |]

for model in exportCsvCycleModels do
    exportCSV model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let exportCsvCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let exportCsvCycleSw = Stopwatch.StartNew()
let mutable exportCsvCycleTotalLength = 0
for _pass in 1 .. 20 do
    for model in exportCsvCycleModels do
        exportCsvCycleTotalLength <- exportCsvCycleTotalLength + (exportCSV model).Length
exportCsvCycleSw.Stop()
let exportCsvCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - exportCsvCycleAllocatedBefore
check "native export csv source cycle"
    (exportCsvCycleSw.Elapsed.TotalMilliseconds < 80.0
     && exportCsvCycleAllocated < 512L * 1024L
     && exportCsvCycleTotalLength = 23943700)
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

let importCsvCycleSourceCount = 80
let importCsvCycleRowCount = 3000
let importCsvCycleText =
    [ for row in 0 .. importCsvCycleRowCount - 1 -> string row ]
    |> String.concat "\n"
let importCsvCycleModels =
    [| for _sourceIndex in 0 .. importCsvCycleSourceCount - 1 ->
        let cells =
            [ for row in 0 .. importCsvCycleRowCount - 1 ->
                { Row = row; Col = 0 }, cell (CellValue.Number (float row)) ]
            |> Map.ofList
        { createEmpty importCsvCycleRowCount 1 with Cells = cells } |]

for model in importCsvCycleModels do
    importCSV importCsvCycleText 0 0 model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let importCsvCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let importCsvCycleSw = Stopwatch.StartNew()
let mutable importCsvCycleSame = 0
for _pass in 1 .. 20 do
    for model in importCsvCycleModels do
        if obj.ReferenceEquals(importCSV importCsvCycleText 0 0 model, model) then
            importCsvCycleSame <- importCsvCycleSame + 1
importCsvCycleSw.Stop()
let importCsvCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - importCsvCycleAllocatedBefore
check "native import csv source cycle"
    (importCsvCycleSw.Elapsed.TotalMilliseconds < 80.0
     && importCsvCycleAllocated < 512L * 1024L
     && importCsvCycleSame = 1600)
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

let formulaPresenceCycleSourceCount = 80
let formulaPresenceCycleCellCount = 5000
let formulaPresenceCycleSources =
    [| for sourceIndex in 0 .. formulaPresenceCycleSourceCount - 1 ->
        [ for row in 0 .. formulaPresenceCycleCellCount - 1 ->
            { Row = row; Col = sourceIndex % 4 }, cell (CellValue.Number (float row)) ]
        |> Map.ofList |]
for cells in formulaPresenceCycleSources do
    cellsHaveFormulaCells cells |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let formulaPresenceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let formulaPresenceCycleSw = Stopwatch.StartNew()
let mutable formulaPresenceCycleTotal = 0
for _pass in 1 .. 20 do
    for cells in formulaPresenceCycleSources do
        if cellsHaveFormulaCells cells then
            formulaPresenceCycleTotal <- formulaPresenceCycleTotal + 1
        match cellsFormulaCountIfKnown cells with
        | Some count -> formulaPresenceCycleTotal <- formulaPresenceCycleTotal + count
        | None -> formulaPresenceCycleTotal <- formulaPresenceCycleTotal - 1
formulaPresenceCycleSw.Stop()
let formulaPresenceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - formulaPresenceCycleAllocatedBefore
check "native formula presence source cycle"
    (formulaPresenceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && formulaPresenceCycleAllocated < 512L * 1024L
     && formulaPresenceCycleTotal = 0)
    true

let recalcIndependentFormulaCount = 5000
let recalcIndependentCells =
    [ for row in 0 .. recalcIndependentFormulaCount - 1 do
        yield { Row = row; Col = 0 }, cell (CellValue.Number (float row))
        yield { Row = row; Col = 1 }, cell (CellValue.Formula (sprintf "=1+%d" (row % 20), CellValue.Empty)) ]
    |> Map.ofList
let recalcIndependentModel =
    { createEmpty recalcIndependentFormulaCount 2 with
        Cells = recalcIndependentCells }
RedBat.Spreadsheet.Formula.recalculateAllFormulas None None recalcIndependentModel |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let recalcIndependentAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let recalcIndependentSw = Stopwatch.StartNew()
let mutable recalcIndependentTotal = 0
for _pass in 1 .. 20 do
    let recalculated =
        RedBat.Spreadsheet.Formula.recalculateAllFormulas None None recalcIndependentModel
    recalcIndependentTotal <- recalcIndependentTotal + recalculated.Cells.Count
recalcIndependentSw.Stop()
let recalcIndependentAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - recalcIndependentAllocatedBefore
check "native independent formula recalc fast path"
    (recalcIndependentSw.Elapsed.TotalMilliseconds < 30.0
     && recalcIndependentAllocated < 12L * 1024L
     && recalcIndependentTotal = recalcIndependentModel.Cells.Count * 20)
    true

let recalcCriteriaFormulaCount = 6000
let recalcCriteriaSourceRows = 8
let recalcCriteriaCells =
    [ for row in 0 .. recalcCriteriaSourceRows - 1 do
        yield { Row = row; Col = 0 }, cell (CellValue.Number (float row))
      for row in 0 .. recalcCriteriaFormulaCount - 1 do
        yield
            { Row = row; Col = 1 },
            cell (CellValue.Formula ("=COUNTIF($A$1:$A$8,\">=4\")", CellValue.Empty)) ]
    |> Map.ofList
let recalcCriteriaModel =
    { createEmpty (recalcCriteriaFormulaCount + recalcCriteriaSourceRows + 4) 4 with
        Cells = recalcCriteriaCells }
RedBat.Spreadsheet.Formula.recalculateAllFormulas None None recalcCriteriaModel |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let recalcCriteriaAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let recalcCriteriaSw = Stopwatch.StartNew()
let mutable recalcCriteriaTotal = 0.0
for _pass in 1 .. 30 do
    let recalculated =
        RedBat.Spreadsheet.Formula.recalculateAllFormulas None None recalcCriteriaModel
    match Map.tryFind { Row = 0; Col = 1 } recalculated.Cells with
    | Some { Value = CellValue.Formula (_, CellValue.Number value) } ->
        recalcCriteriaTotal <- recalcCriteriaTotal + value
    | _ ->
        recalcCriteriaTotal <- recalcCriteriaTotal - 1000000.0
recalcCriteriaSw.Stop()
let recalcCriteriaAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - recalcCriteriaAllocatedBefore
check "native independent criteria recalc stable model"
    (recalcCriteriaSw.Elapsed.TotalMilliseconds < 20.0
     && recalcCriteriaAllocated < 32L * 1024L
     && recalcCriteriaTotal = 120.0)
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

let conditionalFormatCycleSourceCount = 80
let conditionalFormatCycleRows = 2000
let conditionalFormatCycleModels =
    [| for sourceIndex in 0 .. conditionalFormatCycleSourceCount - 1 ->
        let cells =
            [ for row in 0 .. conditionalFormatCycleRows - 1 ->
                { Row = row; Col = 0 }, cell (CellValue.String (sprintf "source-%d-row-%d" sourceIndex row)) ]
            |> Map.ofList
        { createEmpty conditionalFormatCycleRows 1 with Cells = cells } |]

for model in conditionalFormatCycleModels do
    applyDataBars 0 0 (conditionalFormatCycleRows - 1) 0 "#63C384" model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let conditionalFormatSourceCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let conditionalFormatSourceCycleSw = Stopwatch.StartNew()
let mutable conditionalFormatSourceCycleHits = 0
for _pass in 1 .. 20 do
    for model in conditionalFormatCycleModels do
        if obj.ReferenceEquals(
            applyDataBars 0 0 (conditionalFormatCycleRows - 1) 0 "#63C384" model,
            model) then
            conditionalFormatSourceCycleHits <- conditionalFormatSourceCycleHits + 1
conditionalFormatSourceCycleSw.Stop()
let conditionalFormatSourceCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - conditionalFormatSourceCycleAllocatedBefore
check "native conditional format source cycle"
    (conditionalFormatSourceCycleSw.Elapsed.TotalMilliseconds < 80.0
     && conditionalFormatSourceCycleAllocated < 512L * 1024L
     && conditionalFormatSourceCycleHits = 1600)
    true

let conditionalFormatCycleKeyColors =
    [| for index in 0 .. 79 ->
        sprintf "#%06X" ((index * 7919) &&& 0xFFFFFF) |]
let conditionalFormatCycleKeyModel = conditionalFormatCycleModels.[0]

for color in conditionalFormatCycleKeyColors do
    applyDataBars 0 0 (conditionalFormatCycleRows - 1) 0 color conditionalFormatCycleKeyModel |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let conditionalFormatEntryCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let conditionalFormatEntryCycleSw = Stopwatch.StartNew()
let mutable conditionalFormatEntryCycleHits = 0
for _pass in 1 .. 20 do
    for color in conditionalFormatCycleKeyColors do
        if obj.ReferenceEquals(
            applyDataBars 0 0 (conditionalFormatCycleRows - 1) 0 color conditionalFormatCycleKeyModel,
            conditionalFormatCycleKeyModel) then
            conditionalFormatEntryCycleHits <- conditionalFormatEntryCycleHits + 1
conditionalFormatEntryCycleSw.Stop()
let conditionalFormatEntryCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - conditionalFormatEntryCycleAllocatedBefore
check "native conditional format entry cycle"
    (conditionalFormatEntryCycleSw.Elapsed.TotalMilliseconds < 80.0
     && conditionalFormatEntryCycleAllocated < 512L * 1024L
     && conditionalFormatEntryCycleHits = 1600)
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

let pasteIdentityCycleCount = 80
let pasteIdentityCycleRows = 2000
let pasteIdentityCycleModels =
    [| for sourceIndex in 0 .. pasteIdentityCycleCount - 1 ->
        let targetCol = sourceIndex % 4
        let cells =
            [ for row in 0 .. pasteIdentityCycleRows - 1 ->
                { Row = row; Col = targetCol },
                cell (CellValue.String (sprintf "same-%04d-%04d" sourceIndex row)) ]
            |> Map.ofList
        let model = { createEmpty pasteIdentityCycleRows 4 with Cells = cells }
        let clipboard = serializeSelection (SingleCell (0, targetCol)) model false
        model, clipboard, targetCol |]

for model, clipboard, targetCol in pasteIdentityCycleModels do
    pasteClipboard { Row = 0; Col = targetCol } clipboard model |> ignore

System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()

let pasteIdentityCycleAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let pasteIdentityCycleSw = Stopwatch.StartNew()
let mutable pasteIdentityCycleSame = 0
for _pass in 1 .. 20 do
    for model, clipboard, targetCol in pasteIdentityCycleModels do
        if obj.ReferenceEquals(pasteClipboard { Row = 0; Col = targetCol } clipboard model, model) then
            pasteIdentityCycleSame <- pasteIdentityCycleSame + 1
pasteIdentityCycleSw.Stop()
let pasteIdentityCycleAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - pasteIdentityCycleAllocatedBefore
check "native paste identity source cycle"
    (pasteIdentityCycleSw.Elapsed.TotalMilliseconds < 80.0
     && pasteIdentityCycleAllocated < 16L * 1024L
     && pasteIdentityCycleSame = 1600)
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
