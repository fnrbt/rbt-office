module WordWispApp

// wisp-word: the native Word app on Wisp. Reuses Rbt.Word (model/layout/paint + editing ops) and
// the Wisp OfficeShell chrome; the ../widgets-facing layer is gone. EditorHost maps Wisp input to the
// shared Document ops, so this is now actually editable (type, click, select, undo, …).

open System
open Rbt.Renderer
open Wisp.Reactive
open Wisp.Layout
open Rbt.Office.Ui
open Rbt.Word.Document

module Ui = Wisp.Atrium
module Shell = Wisp.OfficeShell
module Doc = Rbt.Word.Document
module WordUi = Rbt.Word.Ui
module WordTextBuffer = Rbt.Word.TextBuffer
module EH = Rbt.WordWisp.EditorHost

let private pageWidth viewportWidth = max 1.0 (min 760.0 (viewportWidth - 56.0))

let private rpath x y w h : Rbt.Renderer.Types.PathSeg list =
    [ MoveTo(vec2 x y); LineTo(vec2 (x + w) y); LineTo(vec2 (x + w) (y + h)); LineTo(vec2 x (y + h)); Close ]

let private mkPara (text: string) : Doc.ParagraphBlock =
    { Id = Guid.NewGuid().ToString("N"); Buffer = WordTextBuffer.ofString text; Metadata = Doc.defaultMetadata }

let private sampleDoc () : Doc.Document =
    { Blocks = [ mkPara "wisp-word"; mkPara "Type here — this document is editable. Click to place the cursor, drag to select, Ctrl+Z to undo." ] |> List.map Doc.Paragraph
      Selection = None; PageSetup = Doc.defaultPageSetup; Styles = []; Headers = []; Footers = []; HeaderParts = []; FooterParts = []; Package = Doc.emptyPackageData }

let private loadDocx (path: string) : Doc.Document option =
    try
        Rbt.Word.Docx.loadBytesNative (IO.File.ReadAllBytes path)
    with ex -> eprintfn "[wisp-word] failed to load %s: %s" path ex.Message; None

[<EntryPoint>]
let main argv =
    let docxPath = argv |> Array.tryFind (fun (a: string) -> a.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
    let loaded = docxPath |> Option.bind loadDocx
    let doc0 = loaded |> Option.defaultWith sampleDoc
    // put the cursor at the start of the first paragraph so the caret shows
    let doc0 =
        match doc0.Blocks |> List.tryPick (function Doc.Paragraph p | Doc.Heading(_, p) -> Some p.Id | _ -> None) with
        | Some bid -> setCursor bid 0 doc0
        | None -> doc0
    let st = EH.create doc0
    let scrollY = signal 0.0
    let mutable lastW = 1000.0
    let mutable lastH = 600.0
    let shell = Shell.create ()
    let currentPath = refSignal (if loaded.IsSome then docxPath else None)
    let status =
        refSignal
            (match docxPath, loaded with
             | Some path, Some _ -> "Opened " + IO.Path.GetFileName path
             | Some path, None -> "Open failed: " + path
             | None, _ -> "Ready")
    st.ReportError <- fun message -> status.Value <- message

    let writeTo path =
        let snapshot = st.Session.CaptureSave()
        try
            let package = Rbt.Word.Docx.exportTypedDocumentPackage snapshot.Document
            let bytes = DocumentWriter.writeTypedPackageBytes { package with Unchanged = snapshot.Unchanged }
            IO.File.WriteAllBytes(path, bytes)
            st.Session.MarkSaved snapshot
            if obj.ReferenceEquals(snapshot.Epoch, st.Session.Epoch) then
                currentPath.Value <- Some path
                status.Value <- "Saved " + IO.Path.GetFileName path
        with ex ->
            status.Value <- "Save failed: " + ex.Message

    let saveAs () =
        st.Session.StopTyping()
        WordUi.saveFileDialog (currentPath.Value |> Option.defaultValue "") writeTo
        |> Shell.openDialog shell

    let save () =
        match currentPath.Value with
        | Some path -> writeTo path
        | None -> saveAs ()

    let loadFromPath path =
        match loadDocx path with
        | Some doc ->
            let doc =
                match doc.Blocks |> List.tryPick (function Doc.Paragraph p | Doc.Heading(_, p) -> Some p.Id | _ -> None) with
                | Some bid -> setCursor bid 0 doc
                | None -> doc
            EH.replaceDocument st doc
            currentPath.Value <- Some path
            scrollY.Value <- 0.0
            status.Value <- "Opened " + IO.Path.GetFileName path
        | None -> status.Value <- "Open failed: " + path

    let openFile () =
        st.Session.StopTyping()
        let picker () = WordUi.openFileDialog loadFromPath |> Shell.openDialog shell
        if st.Session.IsModified then
            { TestId = "word-confirm-open"
              Title = "Discard unsaved changes?"
              Subtitle = Some "Opening another document will discard the current edits."
              Fields = []; Tabs = []
              Buttons =
                [ dialogCancel
                  dialogButton "word-discard-open" "Discard and open" true (fun _ ->
                      picker ()
                      // The picker replaces this dialog; do not close its replacement.
                      false) ] }
            |> Shell.openDialog shell
        else picker ()

    let mutable lastFormat : (Doc.Document * Rbt.Word.WordStyles.SelectionFormat) option = None
    let selectionFormat () =
        let doc = st.Doc.Value
        match lastFormat with
        | Some (cached, format) when Object.ReferenceEquals(cached, doc) -> format
        | _ ->
            let format = Rbt.Word.WordStyles.selectionFormat doc
            lastFormat <- Some (doc, format)
            format
    let characterActive pendingValue selectedValue =
        match st.Session.PendingFormat with
        | Some pending -> pendingValue pending.PendingStyle
        | None -> selectedValue (selectionFormat()) = Rbt.Word.WordStyles.Uniform true

    let command structuralCommand =
        match EH.executeCommand st structuralCommand None with
        | Ok _ -> Scheduler.requestFrame ()
        | Error message -> status.Value <- message
    let selectedParagraph () =
        st.Doc.Value.Selection
        |> Option.bind (fun selected -> Doc.findParagraph st.Doc.Value selected.Focus.BlockId)
    let alignment () =
        match selectedParagraph () |> Option.map (fun para -> para.Metadata.Alignment) with
        | Some Center -> "center"
        | Some Right -> "right"
        | Some Justify -> "justify"
        | _ -> "left"
    let listStyle () =
        match selectedParagraph () |> Option.map (fun para -> para.Metadata.ListStyle) with
        | Some Bullet -> "bullet"
        | Some Numbered -> "numbered"
        | _ -> "none"
    let alignmentCommand = function
        | "center" -> Rbt.Word.Session.AlignCenter
        | "right" -> Rbt.Word.Session.AlignRight
        | "justify" -> Rbt.Word.Session.AlignJustify
        | _ -> Rbt.Word.Session.AlignLeft

    // Keep the already-wired paragraph/list chrome in sync with the session document.
    let formatToolbar command =
        EH.formatCharacter st command
        Scheduler.requestFrame ()
    let wordState : WordUi.ShellState =
        { IsBoldActive = (fun () -> characterActive (fun style -> style.Bold) (fun format -> format.Bold))
          IsItalicActive = (fun () -> characterActive (fun style -> style.Italic) (fun format -> format.Italic))
          IsUnderlineActive = (fun () -> characterActive (fun style -> style.Underline) (fun format -> format.Underline))
          IsAllCapsActive = (fun () -> false); IsSmallCapsActive = (fun () -> false); IsStrikeActive = (fun () -> false)
          Alignment = alignment; BlockStyle = (fun () -> "Normal"); ListStyle = listStyle
          NumberFormat = (fun () -> "decimal"); VertAlign = (fun () -> "baseline")
          TextColor = (fun () -> "#000000"); HighlightColor = (fun () -> "transparent")
          FontFamily = (fun () -> "Calibri"); FontSize = (fun () -> "11"); IsTrackChangesActive = (fun () -> false)
          ParagraphStyles = (fun () -> []); CharacterStyles = (fun () -> [])
          IsInTable = (fun () -> false); CanEdit = (fun () -> false)
          RecentFiles = (fun () -> []); OpenRecentFile = ignore }
    let wordActions : WordUi.ShellActions =
        { OpenFile = openFile; SaveFile = save; PrintPdf = ignore; OpenFindReplace = ignore
          OpenImagePicker = ignore; OpenImageSize = ignore; OpenTextBoxStyle = ignore; OpenPageSetup = ignore
          OpenHeaderFooter = ignore; InsertDateTime = ignore; InsertPageNumber = ignore; InsertTextBox = ignore
          OpenCaption = ignore; OpenCrossReference = ignore; OpenBookmark = ignore; OpenContentControl = ignore
          OpenField = ignore; RefreshFields = ignore; OpenEquation = ignore; OpenRevision = ignore
          ToggleTrackChanges = ignore; SelectNextRevision = ignore; SelectPreviousRevision = ignore
          AcceptSelectedRevision = ignore; RejectSelectedRevision = ignore; AcceptAllRevisions = ignore
          RejectAllRevisions = ignore; InsertSectionBreak = ignore; OpenAnnotation = ignore
          OpenParagraphStyle = ignore; OpenTableCellStyle = ignore; OpenTableMargins = ignore
          OpenFontDialog = None; OpenParagraphBorders = None; OpenListsDialog = None; OpenVbaEditor = None
          OpenInsertTable = ignore; InsertDefaultTable = ignore; InsertTableRow = ignore; InsertTableColumn = ignore
          DeleteTableRow = ignore; DeleteTableColumn = ignore; ApplyHeading = ignore; ApplyNormalParagraph = ignore
          ApplyParagraphStyle = ignore; ApplyCharacterStyle = ignore
          ToggleBullet = (fun () -> command Rbt.Word.Session.ToggleBulletList)
          ToggleNumbered = (fun () -> command Rbt.Word.Session.ToggleNumberedList)
          ApplyNumberFormat = ignore
          IncreaseListIndent = (fun () -> command Rbt.Word.Session.IncreaseIndent)
          DecreaseListIndent = (fun () -> command Rbt.Word.Session.DecreaseIndent)
          ToggleBold = (fun () -> formatToolbar Rbt.Word.Session.CharacterCommand.Bold); ToggleItalic = (fun () -> formatToolbar Rbt.Word.Session.CharacterCommand.Italic)
          ToggleUnderline = (fun () -> formatToolbar Rbt.Word.Session.CharacterCommand.Underline); ToggleAllCaps = ignore
          ToggleSmallCaps = ignore; ToggleStrike = ignore; ClearFormatting = ignore; ApplyVertAlign = ignore
          ApplyAlignment = (alignmentCommand >> command); ApplyFontFamily = ignore; ApplyFontSize = ignore
          ApplyTextColor = ignore; ApplyHighlight = ignore }
    let menus =
        WordUi.menus wordState wordActions
        |> List.map (fun menu ->
            if menu.TestId = "word-file-menu" then
                { menu with
                    Items =
                        menu.Items |> List.collect (fun item ->
                            if item.TestId = "word-menu-save" then
                                [ item; menuItem "word-menu-save-as" "Save as..." None saveAs ]
                            else [ item ]) }
            else menu)
    let toolbarRows = WordUi.toolbarRows wordState wordActions

    let workspace () : Element =
        let el =
            canvasFill (fun x y w h ->
                let previousPageW = pageWidth lastW
                lastW <- w
                lastH <- h
                let pageW = pageWidth w
                let doc = st.Doc.Value
                let pgx = x + (w - pageW) / 2.0
                let docH = EH.documentHeight st pageW doc
                let scroll = max 0.0 (min (max 0.0 (docH + 56.0 - h)) (scrollY.Peek() * pageW / previousPageW))
                if scroll <> scrollY.Peek() then scrollY.Value <- scroll
                let topY = y + 28.0 - scroll
                [ FillRect(rect x y w h, Color.rgb 232uy 234uy 238uy)
                  Cmd.Clip(rpath x y w h, NonZero,
                      [ BoxShadow(rect pgx (topY + 3.0) pageW docH, 0.0, 2.0, 16, 0.0<px>, Color.rgba 0uy 0uy 0uy 46uy, false, None)
                        FillRect(rect pgx topY pageW docH, Color.white)
                        Group(Matrix3x2.translate pgx topY, 1.0, EH.renderDoc st pageW doc) ]) ])
        let area =
            pointerRegion (fun kind lx ly ->
                let pageW = pageWidth lastW
                EH.handlePointer st pageW kind (lx - (lastW - pageW) / 2.0) (ly - 28.0 + scrollY.Peek())) el
        area.OnScroll <- Some(fun dy ->
            let docH = EH.documentHeight st (pageWidth lastW) (st.Doc.Peek())
            scrollY.Value <- max 0.0 (min (max 0.0 (docH + 56.0 - lastH)) (scrollY.Peek() - dy)))
        area

    let title () =
        let name = currentPath.Value |> Option.map IO.Path.GetFileName |> Option.defaultValue "Document1"
        name + (if st.Session.IsModified then " *" else "") + " — wisp-word"
    let view () =
        let content =
            columnA MStart CStretch 0.0
                [ expanded 1 (workspace ())
                  heightBox 22.0 (padXY 8.0 3.0 (textf 11.5 Ui.P.dim false (fun () -> status.Value))) ]
        Shell.chrome shell (title ()) menus toolbarRows content

    match Array.tryFindIndex ((=) "--shot") argv with
    | Some i when i + 1 < argv.Length ->
        (argv |> Array.tryPick (fun (a: string) -> match Double.TryParse a with true, v -> Some v | _ -> None)) |> Option.iter (fun v -> scrollY.Value <- v)
        Wisp.Host.shot 1200 820 view argv.[i + 1]
        EH.dispose st
        printfn "wrote %s" argv.[i + 1]; 0
    | _ ->
        // Editor keys flow through here (no focused field on the page); chrome fields, when open, take focus first.
        Wisp.Host.keyHandler.Value <- fun kc shift ctrl alt ->
            if kc = 9 && shell.OpenMenu.Peek() >= 0 then shell.OpenMenu.Value <- -1
            elif shell.Dialog.Peek().IsSome then ()
            else
                match KeyMap.translate kc shift with
                | Some key when ctrl && (key = "o" || key = "O") -> openFile (); Scheduler.requestFrame ()
                | Some key when ctrl && (key = "s" || key = "S") ->
                    if shift then saveAs () else save ()
                    Scheduler.requestFrame ()
                | Some key -> EH.handleKey st (pageWidth lastW) { Key = key; Shift = shift; Ctrl = ctrl; Alt = alt } |> ignore; Scheduler.requestFrame ()
                | None -> ()
        Wisp.Host.run "wisp-word" 1200 820 view
        EH.dispose st
        0
