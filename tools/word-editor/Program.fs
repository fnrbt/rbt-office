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

// ── Recent files ──────────────────────────────────────────────────

let private recentFilesPath =
    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".config/redbat-word/recent.txt")

let private loadRecentFiles () : string list =
    if File.Exists recentFilesPath then
        try File.ReadAllLines(recentFilesPath) |> Array.toList |> List.filter (fun s -> s.Trim() <> "")
        with _ -> []
    else []

let private saveRecentFiles (list: string list) =
    try
        Directory.CreateDirectory(Path.GetDirectoryName recentFilesPath) |> ignore
        File.WriteAllLines(recentFilesPath, list |> List.toArray)
    with _ -> ()

let private recentFiles : Signal<string list> = signal (loadRecentFiles ())

/// Push a path to the top of the recent-files list (dedupe, cap at 12,
/// persist).
let private rememberRecent (path: string) =
    let abs = Path.GetFullPath path
    let next =
        abs :: (recentFiles.Value |> List.filter (fun p -> p <> abs))
        |> List.truncate 12
    recentFiles.Value <- next
    saveRecentFiles next

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

// ── Minimal native .docx writer ───────────────────────────────────
//
// Walks Document.Blocks, emits Word's wordprocessingml XML for each
// paragraph (with bold / italic / underline / color / font / alignment
// preserved on each run), and bundles the four required parts into a
// zip. Tables, images, headers/footers, and section breaks are NOT
// preserved — but a round-trip through Open → Save keeps the visible
// formatted text intact for the common case.

module DocxWrite =
    let private escapeXml (s: string) =
        let sb = System.Text.StringBuilder()
        for c in s do
            match c with
            | '<' -> sb.Append "&lt;" |> ignore
            | '>' -> sb.Append "&gt;" |> ignore
            | '&' -> sb.Append "&amp;" |> ignore
            | '"' -> sb.Append "&quot;" |> ignore
            | _ -> sb.Append c |> ignore
        sb.ToString()

    let private runXml (text: string) (style: RedBat.Word.TextBuffer.TextStyle) =
        let rPr = System.Text.StringBuilder()
        rPr.Append "<w:rPr>" |> ignore
        if style.Bold then rPr.Append "<w:b/><w:bCs/>" |> ignore
        if style.Italic then rPr.Append "<w:i/><w:iCs/>" |> ignore
        if style.Underline then rPr.Append "<w:u w:val=\"single\"/>" |> ignore
        if style.Strike then rPr.Append "<w:strike/>" |> ignore
        match style.FontFamily with
        | Some f ->
            let f = escapeXml f
            rPr.Append (sprintf "<w:rFonts w:ascii=\"%s\" w:hAnsi=\"%s\" w:cs=\"%s\"/>" f f f) |> ignore
        | None -> ()
        // Font size in DOCX is in half-points.
        let halfPt = int (style.FontSize * 2.0)
        rPr.Append (sprintf "<w:sz w:val=\"%d\"/><w:szCs w:val=\"%d\"/>" halfPt halfPt) |> ignore
        match style.Color with
        | Some hex ->
            let v = if hex.StartsWith "#" then hex.Substring(1) else hex
            rPr.Append (sprintf "<w:color w:val=\"%s\"/>" v) |> ignore
        | None -> ()
        rPr.Append "</w:rPr>" |> ignore
        sprintf "<w:r>%s<w:t xml:space=\"preserve\">%s</w:t></w:r>" (string rPr) (escapeXml text)

    let private paragraphXml (para: Doc.ParagraphBlock) =
        let pPr = System.Text.StringBuilder()
        pPr.Append "<w:pPr>" |> ignore
        match para.Metadata.Alignment with
        | Doc.Center -> pPr.Append "<w:jc w:val=\"center\"/>" |> ignore
        | Doc.Right -> pPr.Append "<w:jc w:val=\"right\"/>" |> ignore
        | Doc.Justify -> pPr.Append "<w:jc w:val=\"both\"/>" |> ignore
        | _ -> ()
        pPr.Append "</w:pPr>" |> ignore

        let text = RedBat.Word.TextBuffer.toString para.Buffer
        let runs = System.Text.StringBuilder()
        let mutable lastEnd = 0
        for run in para.Buffer.Runs do
            match run with
            | RedBat.Word.TextBuffer.Text tr ->
                let s = tr.Start
                let e = tr.Start + tr.Length
                if e > s && s < text.Length then
                    let slice = text.Substring(s, min tr.Length (text.Length - s))
                    runs.Append (runXml slice tr.Style) |> ignore
                lastEnd <- max lastEnd e
            | _ -> ()
        // Fallback: if no runs covered the text, dump as a single default run.
        if runs.Length = 0 && text <> "" then
            runs.Append (runXml text RedBat.Word.TextBuffer.defaultStyle) |> ignore
        sprintf "<w:p>%s%s</w:p>" (string pPr) (string runs)

    let private headingXml (level: int) (para: Doc.ParagraphBlock) =
        let inner = paragraphXml para
        // Inject the heading style id after w:pPr opening — quick splice.
        let styleId = sprintf "Heading%d" level
        let injected = sprintf "<w:pStyle w:val=\"%s\"/>" styleId
        let i = inner.IndexOf "<w:pPr>"
        if i >= 0 then
            inner.Substring(0, i + 7) + injected + inner.Substring(i + 7)
        else
            inner

    let private blocksXml (blocks: Doc.Block list) =
        blocks
        |> List.choose (function
            | Doc.Paragraph p -> Some (paragraphXml p)
            | Doc.Heading (level, p) -> Some (headingXml level p)
            | Doc.Table _ | Doc.TextBox _ -> None) // not preserved
        |> String.concat ""

    let private documentXml (doc: Doc.Document) =
        sprintf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\
<w:body>%s</w:body></w:document>" (blocksXml doc.Blocks)

    let private contentTypesXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
<Default Extension=\"xml\" ContentType=\"application/xml\"/>\
<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\
</Types>"

    let private rootRelsXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\
</Relationships>"

    /// Pack the document as a .docx and write to `path`.
    let save (doc: Doc.Document) (path: string) =
        let files = [
            { ZipWriter.Path = "[Content_Types].xml"; ZipWriter.Content = contentTypesXml }
            { ZipWriter.Path = "_rels/.rels"; ZipWriter.Content = rootRelsXml }
            { ZipWriter.Path = "word/document.xml"; ZipWriter.Content = documentXml doc }
        ]
        let bytes = ZipWriter.createPackageBytes files
        File.WriteAllBytes(path, bytes)

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

/// Number of mutations applied since the last save / load. Reset to 0 by
/// SaveFile; bumped by applyChange. The title bar uses this for the "•"
/// unsaved indicator.
let private editsSinceSave : Signal<int> = signal 0

let private applyChange (state: Editor.WordEditorState) (transform: Doc.Document -> Doc.Document) =
    let before = state.Document.Value
    let after = transform before
    if not (System.Object.ReferenceEquals(before, after)) && after <> before then
        state.UndoStack.Value <- before :: state.UndoStack.Value
        state.RedoStack.Value <- []
        state.Document.Value <- after
        editsSinceSave.Value <- editsSinceSave.Value + 1
        Platform.requestRedraw()

// ── ShellState / ShellActions wiring ───────────────────────────────

/// Lookup the style and metadata under the caret. Falls back to defaults
/// when the cursor is in an empty document or between blocks. For a
/// non-collapsed selection we sample the run just inside the selection
/// start, so "did I bold the selection" lights up the B button.
let private styleAtCursor (state: Editor.WordEditorState) =
    let doc = state.Document.Value
    let probe (para: Doc.ParagraphBlock) (sel: Doc.Selection) =
        let lo = min sel.Anchor.Offset sel.Focus.Offset
        let hi = max sel.Anchor.Offset sel.Focus.Offset
        let bufLen = RedBat.Word.TextBuffer.length para.Buffer
        // Prefer the char just before the upper end of the range (or before
        // the cursor when collapsed); fall back to the first char.
        let p1 = if hi > 0 then hi - 1 else 0
        let p2 = if bufLen > 0 then min p1 (bufLen - 1) else 0
        let chosen = if sel.IsCollapsed then p2 else max lo (min p1 (bufLen - 1))
        RedBat.Word.TextBuffer.getStyleAt chosen para.Buffer
    match doc.Selection with
    | None -> RedBat.Word.TextBuffer.defaultStyle, Doc.defaultMetadata, None
    | Some sel ->
        match Doc.findBlock doc sel.Focus.BlockId with
        | Some (Doc.Paragraph para)        -> probe para sel, para.Metadata, None
        | Some (Doc.Heading (level, para)) -> probe para sel, para.Metadata, Some level
        | _ -> RedBat.Word.TextBuffer.defaultStyle, Doc.defaultMetadata, None

let private shellState (state: Editor.WordEditorState) (openPath: string -> unit) : WordUi.ShellState =
    let style () = let s, _, _ = styleAtCursor state in s
    let meta  () = let _, m, _ = styleAtCursor state in m
    let heading () = let _, _, h = styleAtCursor state in h
    let alignmentName () =
        match (meta ()).Alignment with
        | Doc.Center -> "center"
        | Doc.Right -> "right"
        | Doc.Justify -> "justify"
        | _ -> "left"
    let blockStyleName () =
        match heading () with
        | Some 1 -> "heading-1"
        | Some 2 -> "heading-2"
        | Some 3 -> "heading-3"
        | Some _ -> "heading-1"
        | None -> "normal"
    let listStyleName () =
        match (meta ()).ListStyle with
        | Doc.Bullet -> "bullet"
        | Doc.Numbered -> "numbered"
        | _ -> "none"
    let vertAlignName () =
        match (style ()).VertAlign with
        | Some "superscript" -> "superscript"
        | Some "subscript" -> "subscript"
        | _ -> "baseline"
    {
        IsBoldActive = fun () -> (style ()).Bold
        IsItalicActive = fun () -> (style ()).Italic
        IsUnderlineActive = fun () -> (style ()).Underline
        IsAllCapsActive = fun () -> (style ()).AllCaps
        IsSmallCapsActive = fun () -> (style ()).SmallCaps
        IsStrikeActive = fun () -> (style ()).Strike
        Alignment = alignmentName
        BlockStyle = blockStyleName
        ListStyle = listStyleName
        NumberFormat = fun () -> (meta ()).ListNumberFormat |> Option.defaultValue "decimal"
        VertAlign = vertAlignName
        TextColor = fun () -> (style ()).Color |> Option.defaultValue "#000000"
        HighlightColor = fun () -> (style ()).Highlight |> Option.defaultValue "#FFFFFF"
        FontFamily = fun () -> (style ()).FontFamily |> Option.defaultValue "Calibri"
        FontSize = fun () -> sprintf "%g" (style ()).FontSize
        IsTrackChangesActive = fun () -> false
        ParagraphStyles = fun () -> WordUi.paragraphStyles
        CharacterStyles = fun () -> []
        RecentFiles = fun () -> recentFiles.Value
        OpenRecentFile = openPath
    }

let private shellActions
        (state: Editor.WordEditorState)
        (status: Signal<string>)
        (currentPath: Signal<string option>)
        (renderState: Chrome.ShellRenderState)
        (loadFromPath: string -> unit) : WordUi.ShellActions =
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

    // Dialog-driven actions.
    let openLinkDialog () =
        WordUi.insertLinkDialog (fun url ->
            applyChange state (Doc.setHyperlink (Some url))
            status.Value <- sprintf "Linked to %s" url)
        |> Composit.Ui.OfficeShell.openDialog renderState
    let openImageDialog () =
        WordUi.insertImageDialog (fun src ->
            applyChange state (Doc.insertImageAtCursor src (Some 320.0) (Some 240.0) None)
            status.Value <- sprintf "Inserted image %s" src)
        |> Composit.Ui.OfficeShell.openDialog renderState
    let openInsertTableDialog () =
        WordUi.insertTableDialog (fun rows cols ->
            applyChange state (Doc.insertTable rows cols)
            status.Value <- sprintf "Inserted %dx%d table" rows cols)
        |> Composit.Ui.OfficeShell.openDialog renderState
    let openFindReplaceDialog () =
        WordUi.findReplaceDialog
            (fun q matchCase ->
                if q = "" then
                    status.Value <- "Enter a search term"
                    ""
                else
                    let doc = state.Document.Value
                    let all = Doc.findText q matchCase doc
                    if all.IsEmpty then
                        status.Value <- sprintf "No matches for \"%s\"" q
                        ""
                    else
                        // Pick the first match after the current cursor (wrap).
                        let cursor =
                            match doc.Selection with
                            | Some sel -> sel.Focus
                            | None ->
                                match doc.Blocks |> List.tryHead with
                                | Some (Doc.Paragraph p) -> { Doc.BlockId = p.Id; Offset = 0 }
                                | Some (Doc.Heading (_, p)) -> { Doc.BlockId = p.Id; Offset = 0 }
                                | _ -> { Doc.BlockId = ""; Offset = 0 }
                        let after =
                            all
                            |> List.tryFind (fun m ->
                                m.BlockId = cursor.BlockId && m.Start >= cursor.Offset)
                            |> Option.orElseWith (fun () ->
                                // No more in current block; pick first match in
                                // any later block, otherwise wrap to first overall.
                                let blockOrder =
                                    doc.Blocks
                                    |> List.choose (function
                                        | Doc.Paragraph p -> Some p.Id
                                        | Doc.Heading (_, p) -> Some p.Id
                                        | _ -> None)
                                let curIdx =
                                    blockOrder |> List.tryFindIndex (fun id -> id = cursor.BlockId)
                                    |> Option.defaultValue -1
                                all |> List.tryFind (fun m ->
                                    let mIdx = blockOrder |> List.tryFindIndex (fun id -> id = m.BlockId) |> Option.defaultValue -1
                                    mIdx > curIdx))
                            |> Option.defaultValue (List.head all)
                        applyChange state (fun d ->
                            Doc.setSelection
                                { BlockId = after.BlockId; Offset = after.Start }
                                { BlockId = after.BlockId; Offset = after.End }
                                d)
                        let pos = all |> List.findIndex (fun m -> m.BlockId = after.BlockId && m.Start = after.Start)
                        status.Value <- sprintf "Match %d of %d for \"%s\"" (pos + 1) all.Length q
                        "")
            (fun q r matchCase ->
                let nextDoc, count = Doc.replaceAllText q r matchCase state.Document.Value
                if count > 0 then applyChange state (fun _ -> nextDoc)
                status.Value <- sprintf "Replaced %d occurrence(s)" count
                "")
        |> Composit.Ui.OfficeShell.openDialog renderState
    let openPageSetupDialog () =
        let doc = state.Document.Value
        let ps = doc.PageSetup
        let orient = if ps.PageWidth > ps.PageHeight then "landscape" else "portrait"
        WordUi.pageSetupDialog orient ps.MarginLeft ps.MarginRight ps.MarginTop ps.MarginBottom
            (fun (orient, ml, mr, mt, mb) ->
                applyChange state (fun d ->
                    let setup = d.PageSetup
                    let w, h =
                        match orient with
                        | "landscape" -> max setup.PageWidth setup.PageHeight, min setup.PageWidth setup.PageHeight
                        | _           -> min setup.PageWidth setup.PageHeight, max setup.PageWidth setup.PageHeight
                    { d with PageSetup = { setup with PageWidth = w; PageHeight = h; MarginLeft = ml; MarginRight = mr; MarginTop = mt; MarginBottom = mb } })
                status.Value <- "Page setup applied")
        |> Composit.Ui.OfficeShell.openDialog renderState

    let openAnnotation = function
        | "hyperlink" -> openLinkDialog ()
        | other -> status.Value <- sprintf "%s annotation not wired" other

    let openFileAction () =
        WordUi.openFileDialog loadFromPath
        |> Composit.Ui.OfficeShell.openDialog renderState

    let writeTo path =
        try
            DocxWrite.save state.Document.Value path
            currentPath.Value <- Some path
            editsSinceSave.Value <- 0
            rememberRecent path
            status.Value <- sprintf "Saved %s" (Path.GetFileName path)
        with ex ->
            status.Value <- sprintf "Save failed: %s" ex.Message

    let saveFileAction () =
        let initial = currentPath.Value |> Option.defaultValue ""
        WordUi.saveFileDialog initial writeTo
        |> Composit.Ui.OfficeShell.openDialog renderState
    {
        OpenFile = openFileAction
        SaveFile = saveFileAction
        PrintPdf = say "Print PDF not wired in native shell"
        OpenFindReplace = openFindReplaceDialog
        OpenImagePicker = openImageDialog
        OpenImageSize = say "Image size dialog not wired"
        OpenTextBoxStyle = say "Text-box style dialog not wired"
        OpenPageSetup = openPageSetupDialog
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
        OpenAnnotation = openAnnotation
        OpenParagraphStyle = say "Paragraph style dialog not wired"
        OpenTableCellStyle = say "Cell style dialog not wired"
        OpenTableMargins = say "Cell margins dialog not wired"
        OpenInsertTable = openInsertTableDialog
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
    Actions: WordUi.ShellActions
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
                // Ctrl+S triggers save without going through the editor.
                elif state.CtrlDown && (key = "s" || key = "S") then
                    state.Actions.SaveFile ()
                    state.Window.RequestRedraw()
                // Ctrl+O triggers open.
                elif state.CtrlDown && (key = "o" || key = "O") then
                    state.Actions.OpenFile ()
                    state.Window.RequestRedraw()
                // Ctrl+F triggers find / replace.
                elif state.CtrlDown && (key = "f" || key = "F") then
                    state.Actions.OpenFindReplace ()
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
                // Surface the click position so chrome popups (combo
                // dropdowns, color pickers) can anchor next to the cursor.
                state.RenderState.LastClickPos.Value <- (x, y)
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
        | Some root when button = 2 && pressed ->
            // Right-click: fire the deepest ContextMenuRegion handler if any.
            match contextMenuHitTestAll root x y with
            | Some handler -> handler (x, y)
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
            | Some parsed ->
                rememberRecent path
                parsed
            | None ->
                eprintfn "Failed to parse %s; opening blank document" path
                emptyDoc ()
        | None ->
            emptyDoc ()

    let editor = Editor.createState doc
    let status = signal "Ready"
    let currentPath = signal (argv |> Array.tryHead)
    let renderState = Chrome.closed
    let loadFromPath path =
        match loadDocx path with
        | Some doc ->
            editor.Document.Value <- doc
            editor.UndoStack.Value <- []
            editor.RedoStack.Value <- []
            currentPath.Value <- Some path
            editsSinceSave.Value <- 0
            rememberRecent path
            status.Value <- sprintf "Opened %s" (Path.GetFileName path)
        | None -> status.Value <- sprintf "Failed to open %s" path

    let state = shellState editor loadFromPath
    let actions = shellActions editor status currentPath renderState loadFromPath

    // Any change to the document — typing, toolbar mutation, undo — bumps
    // the unsaved-edits counter. openFileAction / writeTo reset it to 0
    // after they finish mutating Document so the load itself doesn't count.
    editor.Document.Subscribe(fun () -> editsSinceSave.Value <- editsSinceSave.Value + 1)

    let countWords (s: string) =
        let mutable c = 0
        let mutable inWord = false
        for ch in s do
            if System.Char.IsWhiteSpace ch then inWord <- false
            elif not inWord then
                inWord <- true
                c <- c + 1
        c

    /// Status string. Falls back to live word/character counts when no
    /// transient message has been set (or the message is the default
    /// "Ready" sentinel).
    let statusText () =
        let doc = editor.Document.Value
        let counts () =
            match doc.Selection with
            | Some sel when not sel.IsCollapsed ->
                let text = Doc.getSelectionText doc
                sprintf "%d of %d words selected   %d of %d characters"
                    (countWords text) (countWords (Doc.getText doc))
                    text.Length (Doc.getText doc).Length
            | _ ->
                let text = Doc.getText doc
                sprintf "%d words   %d characters" (countWords text) text.Length
        match status.Value with
        | message when message <> "Ready" && message <> "" -> message + "   |   " + counts ()
        | _ -> counts ()

    let titleText () =
        let prefix = if editsSinceSave.Value > 0 then "• " else ""
        let file =
            match currentPath.Value with
            | Some path -> Path.GetFileName path
            | None -> "Untitled"
        sprintf "%sRedBat Word - %s" prefix file

    let chrome = WordUi.chrome
    let workspaceWidth = chrome.DocumentTrackWidth
    // Chrome bands: title 32 + menu 32 + toolbar (32 per row * 2) + ruler 22 + status 22, plus desk padding 48.
    let bodyOverhead = 32.0 + 32.0 + 64.0 + 22.0 + 22.0 + 48.0

    let initialWidth = int chrome.WindowWidth
    let initialHeight = 1000
    let x11win = X11Window(titleText (), initialWidth, initialHeight)
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
        Actions = actions
        ControlServer = controlServer
    }

    // Right-click context menu over the document workspace. Items are a
    // mix of clipboard ops, character formatting, and paragraph style — the
    // most useful subset of what's in the menus.
    let openDocContextMenu (x, y) =
        let item label isChecked action =
            RedBat.Office.Ui.menuItem (sprintf "ctx-%s" label) label isChecked action
        let items = [
            item "Cut" None (fun () ->
                let text = Doc.getSelectionText editor.Document.Value
                if text <> "" then
                    Platform.clipboard.SetText text
                    applyChange editor Doc.deleteSelection
                    status.Value <- sprintf "Cut %d chars" text.Length)
            item "Copy" None (fun () ->
                let text = Doc.getSelectionText editor.Document.Value
                if text <> "" then
                    Platform.clipboard.SetText text
                    status.Value <- sprintf "Copied %d chars" text.Length)
            item "Paste" None (fun () ->
                match Platform.clipboard.GetText() with
                | Some text when text <> "" ->
                    applyChange editor (Doc.insertText text)
                    status.Value <- sprintf "Pasted %d chars" text.Length
                | _ -> ())
            item "Bold" None (fun () -> applyChange editor Doc.toggleBold)
            item "Italic" None (fun () -> applyChange editor Doc.toggleItalic)
            item "Underline" None (fun () -> applyChange editor Doc.toggleUnderline)
            item "Insert link..." None (fun () -> actions.OpenAnnotation "hyperlink")
            item "Page setup..." None actions.OpenPageSetup
        ]
        Composit.Ui.OfficeShell.openContextMenu renderState x y items

    let buildWidget () =
        let bodyHeight = max 200.0 (float appState.Height - bodyOverhead)
        let editorWidget = Editor.scrollWidget workspaceWidth bodyHeight (fun _ -> None) editor
        let workspace =
            Chrome.documentWorkspace editorWidget
            |> onContextMenu openDocContextMenu
        Chrome.viewWith renderState (titleText ()) statusText state actions workspace

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
    controlServer.RightClickAtHandler <- Some (fun (x, y) ->
        handleInput appState (MouseButton(2, true, x, y))
        handleInput appState (MouseButton(2, false, x, y)))
    controlServer.MoveAtHandler <- Some (fun (x, y) ->
        handleInput appState (MouseMove(x, y)))
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
