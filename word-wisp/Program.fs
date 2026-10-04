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
module WordImport = Rbt.Word.Import
module WordUi = Rbt.Word.Ui
module WordTextBuffer = Rbt.Word.TextBuffer
module WR = Rbt.WordWisp.WordRender
module EH = Rbt.WordWisp.EditorHost

let private pageW = 760.0

let private rpath x y w h : Rbt.Renderer.Types.PathSeg list =
    [ MoveTo(vec2 x y); LineTo(vec2 (x + w) y); LineTo(vec2 (x + w) (y + h)); LineTo(vec2 x (y + h)); Close ]

let private mkPara (text: string) : Doc.ParagraphBlock =
    { Id = Guid.NewGuid().ToString("N"); Buffer = WordTextBuffer.ofString text; Metadata = Doc.defaultMetadata }

let private sampleDoc () : Doc.Document =
    { Blocks = [ mkPara "wisp-word"; mkPara "Type here — this document is editable. Click to place the cursor, drag to select, Ctrl+Z to undo." ] |> List.map Doc.Paragraph
      Selection = None; PageSetup = Doc.defaultPageSetup; Styles = []; Headers = []; Footers = []; HeaderParts = []; FooterParts = [] }

let private loadDocx (path: string) : Doc.Document option =
    try
        match DocxReader.loadBytesNative (IO.File.ReadAllBytes path) with
        | Some parsed -> Some(WordImport.importDoc parsed)
        | None -> None
    with ex -> eprintfn "[wisp-word] failed to load %s: %s" path ex.Message; None

[<EntryPoint>]
let main argv =
    let docxPath = argv |> Array.tryFind (fun (a: string) -> a.EndsWith ".docx")
    let doc0 =
        match docxPath |> Option.bind loadDocx with
        | Some d -> eprintfn "[wisp-word] opened %s" (docxPath |> Option.defaultValue ""); d
        | None -> sampleDoc ()
    // put the cursor at the start of the first paragraph so the caret shows
    let doc0 =
        match doc0.Blocks |> List.tryPick (function Doc.Paragraph p | Doc.Heading(_, p) -> Some p.Id | _ -> None) with
        | Some bid -> setCursor bid 0 doc0
        | None -> doc0
    let st = EH.create doc0
    let scrollY = signal 0.0
    let mutable lastW = 1000.0

    // Word's real chrome: state getters as read-only defaults; actions wired to the editor where cheap.
    let wordState : WordUi.ShellState =
        { IsBoldActive = (fun () -> false); IsItalicActive = (fun () -> false); IsUnderlineActive = (fun () -> false)
          IsAllCapsActive = (fun () -> false); IsSmallCapsActive = (fun () -> false); IsStrikeActive = (fun () -> false)
          Alignment = (fun () -> "left"); BlockStyle = (fun () -> "Normal"); ListStyle = (fun () -> "none")
          NumberFormat = (fun () -> "decimal"); VertAlign = (fun () -> "baseline")
          TextColor = (fun () -> "#000000"); HighlightColor = (fun () -> "transparent")
          FontFamily = (fun () -> "Calibri"); FontSize = (fun () -> "11"); IsTrackChangesActive = (fun () -> false)
          ParagraphStyles = (fun () -> []); CharacterStyles = (fun () -> [])
          RecentFiles = (fun () -> []); OpenRecentFile = ignore }
    let alignFn = function "center" -> alignCenter | "right" -> alignRight | "justify" -> alignJustify | _ -> alignLeft
    let wordActions : WordUi.ShellActions =
        { OpenFile = ignore; SaveFile = ignore; PrintPdf = ignore; OpenFindReplace = ignore
          OpenImagePicker = ignore; OpenImageSize = ignore; OpenTextBoxStyle = ignore; OpenPageSetup = ignore
          OpenHeaderFooter = ignore; InsertDateTime = ignore; InsertPageNumber = ignore; InsertTextBox = ignore
          OpenCaption = ignore; OpenCrossReference = ignore; OpenBookmark = ignore; OpenContentControl = ignore
          OpenField = ignore; RefreshFields = ignore; OpenEquation = ignore; OpenRevision = ignore
          ToggleTrackChanges = ignore; SelectNextRevision = ignore; SelectPreviousRevision = ignore
          AcceptSelectedRevision = ignore; RejectSelectedRevision = ignore; AcceptAllRevisions = ignore
          RejectAllRevisions = ignore; InsertSectionBreak = ignore; OpenAnnotation = ignore
          OpenParagraphStyle = ignore; OpenTableCellStyle = ignore; OpenTableMargins = ignore
          OpenInsertTable = ignore; InsertDefaultTable = ignore; InsertTableRow = ignore; InsertTableColumn = ignore
          DeleteTableRow = ignore; DeleteTableColumn = ignore; ApplyHeading = ignore; ApplyNormalParagraph = ignore
          ApplyParagraphStyle = ignore; ApplyCharacterStyle = ignore
          ToggleBullet = (fun () -> EH.edit st toggleBulletList); ToggleNumbered = (fun () -> EH.edit st toggleNumberedList)
          ApplyNumberFormat = ignore
          IncreaseListIndent = (fun () -> EH.edit st increaseIndent); DecreaseListIndent = (fun () -> EH.edit st decreaseIndent)
          ToggleBold = (fun () -> EH.edit st Rbt.Word.WordStyles.toggleBold); ToggleItalic = (fun () -> EH.edit st Rbt.Word.WordStyles.toggleItalic)
          ToggleUnderline = (fun () -> EH.edit st Rbt.Word.WordStyles.toggleUnderline); ToggleAllCaps = ignore
          ToggleSmallCaps = ignore; ToggleStrike = ignore; ClearFormatting = ignore; ApplyVertAlign = ignore
          ApplyAlignment = (fun a -> EH.edit st (alignFn a)); ApplyFontFamily = ignore; ApplyFontSize = ignore
          ApplyTextColor = ignore; ApplyHighlight = ignore }
    let menus = WordUi.menus wordState wordActions
    let toolbarRows = WordUi.toolbarRows wordState wordActions

    let workspace () : Element =
        let el =
            canvasFill (fun x y w h ->
                lastW <- w
                let doc = st.Doc.Value
                let pgx = x + (w - pageW) / 2.0
                let docH = EH.documentHeight st pageW doc
                let topY = y + 28.0 - scrollY.Value
                [ FillRect(rect x y w h, Color.rgb 232uy 234uy 238uy)
                  Cmd.Clip(rpath x y w h, NonZero,
                      [ BoxShadow(rect pgx (topY + 3.0) pageW docH, 0.0, 2.0, 16, 0.0<px>, Color.rgba 0uy 0uy 0uy 46uy, false, None)
                        FillRect(rect pgx topY pageW docH, Color.white)
                        Group(Matrix3x2.translate pgx topY, 1.0, EH.renderDoc st pageW doc) ]) ])
        let area =
            pointerRegion (fun kind lx ly ->
                EH.handlePointer st pageW kind (lx - (lastW - pageW) / 2.0) (ly - 28.0 + scrollY.Peek())) el
        area.OnScroll <- Some(fun dy ->
            let docH = EH.documentHeight st pageW (st.Doc.Peek())
            scrollY.Value <- max 0.0 (min (max 0.0 (docH - 300.0)) (scrollY.Peek() - dy)))
        area

    let title = (match docxPath with Some p -> IO.Path.GetFileName p | None -> "Document1") + " — wisp-word"
    let shell = Shell.create ()
    let view () = Shell.chrome shell title menus toolbarRows (workspace ())

    match Array.tryFindIndex ((=) "--shot") argv with
    | Some i when i + 1 < argv.Length ->
        (argv |> Array.tryPick (fun (a: string) -> match Double.TryParse a with true, v -> Some v | _ -> None)) |> Option.iter (fun v -> scrollY.Value <- v)
        Wisp.Host.shot 1200 820 view argv.[i + 1]
        printfn "wrote %s" argv.[i + 1]; 0
    | _ ->
        // Editor keys flow through here (no focused field on the page); chrome fields, when open, take focus first.
        Wisp.Host.keyHandler.Value <- fun kc shift ctrl alt ->
            if kc = 9 && shell.OpenMenu.Peek() >= 0 then shell.OpenMenu.Value <- -1
            else
                match KeyMap.translate kc shift with
                | Some key -> EH.handleKey st pageW { Key = key; Shift = shift; Ctrl = ctrl; Alt = alt } |> ignore; Scheduler.requestFrame ()
                | None -> ()
        Wisp.Host.run "wisp-word" 1200 820 view
        0
