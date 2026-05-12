module WordEditorApp

open System
open System.IO
open Renderer
open Platform
open Platform.X11Window
open Composit.Ui.Reactive
open Composit.Ui.Widget
open Composit.Ui.Layout

module WordUi = RedBat.Word.Ui
module Doc = RedBat.Word.Document
module Editor = Composit.Ui.WordEditor
module Chrome = Composit.Ui.WordUiNative

let private existingFile path =
    if File.Exists path then Some path else None

let private wordFontPath () =
    [
        Environment.GetEnvironmentVariable("RB_WORD_FONT_DIR")
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "artifacts/fonts/windows"))
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../artifacts/fonts/windows"))
        Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "../../artifacts/fonts/windows"))
    ]
    |> List.choose (fun dir ->
        if String.IsNullOrWhiteSpace dir then None
        else Path.Combine(dir, "calibri.ttf") |> existingFile)
    |> List.tryHead

let private loadDocx (path: string) : Doc.Document option =
    let bytes = File.ReadAllBytes path
    match DocxReader.loadBytesNative bytes with
    | Some parsed -> Some (NativeRunner.importDoc parsed)
    | None -> None

let private emptyDoc () : Doc.Document =
    let paraId = Guid.NewGuid().ToString("N")
    let para : Doc.ParagraphBlock = {
        Id = paraId
        Buffer = { Rope = RedBat.Word.Rope.ofString ""; Runs = [] }
        Metadata = Doc.defaultMetadata
    }
    {
        Blocks = [ Doc.Paragraph para ]
        Selection = Some { Anchor = { BlockId = paraId; Offset = 0 }
                           Focus  = { BlockId = paraId; Offset = 0 } }
        PageSetup = Doc.defaultPageSetup
        Styles = []
        Headers = []
        Footers = []
        HeaderParts = []
        FooterParts = []
    }

// ── Editor state mutation with undo tracking ───────────────────────

let private applyChange (state: Editor.WordEditorState) (transform: Doc.Document -> Doc.Document) =
    let before = state.Document.Value
    let after = transform before
    if not (System.Object.ReferenceEquals(before, after)) && after <> before then
        state.UndoStack.Value <- before :: state.UndoStack.Value
        state.RedoStack.Value <- []
        state.Document.Value <- after
        Platform.requestRedraw()

// ── ShellState / ShellActions wiring ───────────────────────────────

let private shellState (_state: Editor.WordEditorState) : WordUi.ShellState =
    let constFalse () = false
    {
        IsBoldActive = constFalse
        IsItalicActive = constFalse
        IsUnderlineActive = constFalse
        IsAllCapsActive = constFalse
        IsSmallCapsActive = constFalse
        IsStrikeActive = constFalse
        Alignment = fun () -> "left"
        BlockStyle = fun () -> "normal"
        ListStyle = fun () -> "none"
        NumberFormat = fun () -> "decimal"
        VertAlign = fun () -> "baseline"
        TextColor = fun () -> "#000000"
        HighlightColor = fun () -> "#FFFFFF"
        FontFamily = fun () -> "Calibri"
        FontSize = fun () -> "11"
        IsTrackChangesActive = constFalse
        ParagraphStyles = fun () -> WordUi.paragraphStyles
        CharacterStyles = fun () -> []
    }

let private shellActions (state: Editor.WordEditorState) (status: Signal<string>) : WordUi.ShellActions =
    let apply transform () = applyChange state transform
    let applyArg transform arg = applyChange state (transform arg)
    let say message () = status.Value <- message
    let sayArg message _ = status.Value <- message
    let alignment = function
        | "center" -> Doc.alignCenter
        | "right" -> Doc.alignRight
        | "justify" -> Doc.alignJustify
        | _ -> Doc.alignLeft
    let heading level =
        if level <= 0 then Doc.clearHeading
        else Doc.setHeadingLevel level
    {
        PrintPdf = say "Print PDF not wired in native shell"
        OpenFindReplace = say "Find/replace dialog not wired"
        OpenImagePicker = say "Image picker not wired"
        OpenImageSize = say "Image size dialog not wired"
        OpenTextBoxStyle = say "Text-box style dialog not wired"
        OpenPageSetup = say "Page setup dialog not wired"
        OpenHeaderFooter = say "Header/footer dialog not wired"
        InsertDateTime = apply (fun doc ->
            Doc.insertText (DateTime.Now.ToString("yyyy-MM-dd HH:mm")) doc)
        InsertPageNumber = apply (Doc.insertField "PAGE" "1")
        InsertTextBox = apply (Doc.insertTextBox "Text box")
        OpenCaption = say "Caption dialog not wired"
        OpenCrossReference = apply (Doc.insertCrossReferenceToLatestCaption "Figure")
        OpenBookmark = say "Bookmark dialog not wired"
        OpenContentControl = say "Content-control dialog not wired"
        OpenField = say "Field dialog not wired"
        RefreshFields = say "Refresh fields not wired"
        OpenEquation = apply (Doc.insertEquation "E = mc^2")
        OpenRevision = say "Revision dialog not wired"
        ToggleTrackChanges = say "Track changes not wired"
        SelectNextRevision = say "No revisions"
        SelectPreviousRevision = say "No revisions"
        AcceptSelectedRevision = say "No revisions"
        RejectSelectedRevision = say "No revisions"
        AcceptAllRevisions = say "No revisions"
        RejectAllRevisions = say "No revisions"
        InsertSectionBreak = sayArg "Section break not wired"
        OpenAnnotation = sayArg "Annotation dialog not wired"
        OpenParagraphStyle = say "Paragraph style dialog not wired"
        OpenTableCellStyle = say "Cell style dialog not wired"
        OpenTableMargins = say "Cell margins dialog not wired"
        OpenInsertTable = say "Insert-table dialog not wired"
        InsertDefaultTable = apply (Doc.insertTable 2 3)
        InsertTableRow = apply Doc.insertTableRowAfterSelection
        InsertTableColumn = apply Doc.insertTableColumnAfterSelection
        DeleteTableRow = apply Doc.deleteSelectedTableRow
        DeleteTableColumn = apply Doc.deleteSelectedTableColumn
        ApplyHeading = applyArg heading
        ApplyNormalParagraph = apply Doc.clearHeading
        ApplyParagraphStyle = applyArg Doc.setParagraphStyle
        ApplyCharacterStyle = applyArg Doc.setCharacterStyle
        ToggleBullet = apply Doc.toggleBulletList
        ToggleNumbered = apply Doc.toggleNumberedList
        ApplyNumberFormat = applyArg Doc.setNumberFormat
        IncreaseListIndent = apply Doc.increaseIndent
        DecreaseListIndent = apply Doc.decreaseIndent
        ToggleBold = apply Doc.toggleBold
        ToggleItalic = apply Doc.toggleItalic
        ToggleUnderline = apply Doc.toggleUnderline
        ToggleAllCaps = apply Doc.toggleAllCaps
        ToggleSmallCaps = apply Doc.toggleSmallCaps
        ToggleStrike = apply Doc.toggleStrike
        ClearFormatting = apply Doc.clearFormatting
        ApplyVertAlign = fun value -> applyChange state (Doc.setVertAlign (Some value))
        ApplyAlignment = applyArg alignment
        ApplyFontFamily = applyArg Doc.setFontFamily
        ApplyFontSize = fun value ->
            match Double.TryParse value with
            | true, parsed -> applyChange state (Doc.setFontSize parsed)
            | false, _ -> ()
        ApplyTextColor = fun value -> applyChange state (Doc.setColor (Some value))
        ApplyHighlight = fun value -> applyChange state (Doc.setHighlight (Some value))
    }

// ── App state and event loop ───────────────────────────────────────

type AppState = {
    mutable ShiftDown: bool
    mutable CtrlDown: bool
    mutable AltDown: bool
    mutable LayoutRoot: LayoutNode option
    mutable Width: int
    mutable Height: int
    mutable LastHover: (bool -> unit) option
    mutable LastMouseX: float
    mutable LastMouseY: float
    mutable LastCmds: Cmd list
    Window: IWindow
    EditorState: Editor.WordEditorState
    RenderState: Chrome.ShellRenderState
    ControlServer: ControlServer.ControlServer
}

let private handleInput (state: AppState) (ev: InputEvent) =
    match ev with
    | KeyDown(keycode, _) ->
        match keycode with
        | 50 | 62 -> state.ShiftDown <- true
        | 37 | 105 -> state.CtrlDown <- true
        | 64 | 108 -> state.AltDown <- true
        | _ -> ()

        if not (KeyMap.isModifier keycode) then
            match KeyMap.translate keycode state.ShiftDown with
            | Some key ->
                // Escape closes an open menu first; otherwise falls through to editor.
                if key = "Escape" && state.RenderState.OpenMenu.Value.IsSome then
                    state.RenderState.OpenMenu.Value <- None
                    state.Window.RequestRedraw()
                else
                    let ke : KeyEvent = { Key = key; Shift = state.ShiftDown; Ctrl = state.CtrlDown; Alt = state.AltDown }
                    match state.LayoutRoot with
                    | Some root ->
                        let _ = dispatchKey root ke
                        state.Window.RequestRedraw()
                    | None -> ()
            | None -> ()

    | KeyUp(keycode, _) ->
        match keycode with
        | 50 | 62 -> state.ShiftDown <- false
        | 37 | 105 -> state.CtrlDown <- false
        | 64 | 108 -> state.AltDown <- false
        | _ -> ()

    | MouseButton(button, pressed, x, y) ->
        state.LastMouseX <- x
        state.LastMouseY <- y
        match state.LayoutRoot with
        | Some root when button = 0 ->
            if pressed then
                let (onClick, _) = hitTestAll root x y
                onClick |> Option.iter (fun f -> f())
                match pointerHitTestAll root x y with
                | Some(h, lx, ly) -> h (PointerDown(lx, ly))
                | None -> ()
            else
                match pointerHitTestAll root x y with
                | Some(h, lx, ly) -> h (PointerUp(lx, ly))
                | None -> ()
            state.Window.RequestRedraw()
        | _ -> ()

    | MouseMove(x, y) ->
        state.LastMouseX <- x
        state.LastMouseY <- y
        match state.LayoutRoot with
        | Some root ->
            let (_, onHover) = hitTestAll root x y
            state.LastHover |> Option.iter (fun f -> f false)
            onHover |> Option.iter (fun f -> f true)
            state.LastHover <- onHover
            match pointerHitTestAll root x y with
            | Some(h, lx, ly) -> h (PointerMove(lx, ly))
            | None -> ()
            state.Window.RequestRedraw()
        | None -> ()

    | MouseScroll(_, dy) ->
        match state.LayoutRoot with
        | Some root ->
            match scrollHitTestAll root state.LastMouseX state.LastMouseY with
            | Some(_, sy) ->
                sy.Value <- sy.Peek() + dy * 60.0
                state.Window.RequestRedraw()
            | None -> ()
        | None -> ()

    | Resize(w, h) ->
        state.Width <- w
        state.Height <- h

    | Close ->
        state.Window.Close()

    | MouseLeave ->
        state.LastHover |> Option.iter (fun f -> f false)
        state.LastHover <- None

let private renderFrame (state: AppState) (buildWidget: unit -> Widget) (w: int) (h: int) : Cmd list =
    state.ControlServer.ProcessCommands()
    try
        state.Width <- w
        state.Height <- h
        Animation.tick()
        TextEdit.tickCaret 0.016

        let widgetTree = buildWidget()
        let bounds = { X = 0.0; Y = 0.0; Width = float w; Height = float h }
        clearOverlays()
        initOverlayBounds bounds
        let root = layout widgetTree bounds
        state.LayoutRoot <- Some root

        let mainCmds = paint root
        let overlayCmds = getOverlays() |> List.collect paint
        let bg = Composit.Ui.OfficeShell.palette WordUi.chrome
        let cmds = [ Clear bg.PaperDesk; yield! mainCmds; yield! overlayCmds ]
        state.LastCmds <- cmds
        cmds
    with ex ->
        eprintfn "[render] %O" ex
        [ Clear (Color.rgb 255uy 0uy 0uy) ]

// ── Main ───────────────────────────────────────────────────────────

[<EntryPoint>]
let main argv =
    let sw = Diagnostics.Stopwatch.StartNew()
    Platform.now <- fun () -> sw.Elapsed.TotalSeconds

    wordFontPath () |> Option.iter Composit.Ui.Layout.initFont

    let doc =
        match argv |> Array.tryHead with
        | Some path ->
            match loadDocx path with
            | Some parsed -> parsed
            | None ->
                eprintfn "Failed to parse %s; opening blank document" path
                emptyDoc ()
        | None ->
            emptyDoc ()

    let editor = Editor.createState doc
    let status = signal "Ready"
    let renderState = Chrome.closed
    let state = shellState editor
    let actions = shellActions editor status

    /// Status string. Falls back to the editor's current cursor / selection
    /// summary if no transient message has been set (the `Ready` literal).
    let statusText () =
        let doc = editor.Document.Value
        match status.Value, doc.Selection with
        | message, _ when message <> "Ready" && message <> "" -> message
        | _, Some sel when not sel.IsCollapsed ->
            let text = Doc.getSelectionText doc
            sprintf "%d characters selected" text.Length
        | _, _ -> "Ready"

    let title =
        match argv |> Array.tryHead with
        | Some path -> sprintf "RedBat Word - %s" (Path.GetFileName path)
        | None -> "RedBat Word - Untitled"

    let chrome = WordUi.chrome
    let workspaceWidth = chrome.DocumentTrackWidth
    // Chrome bands: title 32 + menu 32 + toolbar (32 per row * 2) + ruler 22 + status 22, plus desk padding 48.
    let bodyOverhead = 32.0 + 32.0 + 64.0 + 22.0 + 22.0 + 48.0

    let initialWidth = int chrome.WindowWidth
    let initialHeight = 1000
    let x11win = X11Window(title, initialWidth, initialHeight)
    let window = x11win :> IWindow

    let controlServer = new ControlServer.ControlServer()

    let appState = {
        ShiftDown = false; CtrlDown = false; AltDown = false
        LayoutRoot = None
        Width = initialWidth; Height = initialHeight
        LastHover = None
        LastMouseX = 0.0; LastMouseY = 0.0
        LastCmds = []
        Window = window
        EditorState = editor
        RenderState = renderState
        ControlServer = controlServer
    }

    let buildWidget () =
        let bodyHeight = max 200.0 (float appState.Height - bodyOverhead)
        let editorWidget = Editor.scrollWidget workspaceWidth bodyHeight (fun _ -> None) editor
        let workspace = Chrome.documentWorkspace editorWidget
        Chrome.viewWith renderState title statusText state actions workspace

    Platform.requestRedraw <- fun () -> window.RequestRedraw()
    Platform.setTimeout <- fun ms cb -> window.SetTimeout(ms, cb)
    Platform.clearTimeout <- fun id -> window.ClearTimeout(id)
    Platform.onTextInputFocus <- fun _ -> ()
    Platform.onTextInputBlur <- fun () -> ()
    Platform.onCursorRectChange <- fun _ -> ()

    // Control server for screenshots / synthetic input (parity with widgets/x11-viewer).
    let rasterBackend = Raster.create()
    controlServer.ScreenshotProvider <- Some (fun () ->
        let cmds = appState.LastCmds
        if cmds.IsEmpty then Array.empty
        else
            rasterBackend.Init appState.Width appState.Height
            rasterBackend.Submit cmds
            let fb = rasterBackend.Finish()
            PngEncode.encode fb)
    controlServer.ClickAtHandler <- Some (fun (x, y) ->
        handleInput appState (MouseButton(0, true, x, y))
        handleInput appState (MouseButton(0, false, x, y)))
    controlServer.SendKeyHandler <- Some (fun keycode pressed ->
        if pressed then handleInput appState (KeyDown(keycode, 0u))
        else handleInput appState (KeyUp(keycode, 0u)))
    controlServer.ExitHandler <- Some (fun () -> window.Close())
    controlServer.OnCommandReceived <- Some (fun () -> window.RequestRedraw())
    controlServer.Start()

    Focus.request editor.FocusId

    window.Run(
        (fun ev -> handleInput appState ev),
        (fun w h -> renderFrame appState buildWidget w h))

    (controlServer :> System.IDisposable).Dispose()
    0
