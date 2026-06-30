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
