// The order a page is painted in: the headers and footers (their drawings
// too - a watermark) below the body; of the body, the anchored drawings that
// stand behind the text (wp:anchor behindDoc="1") in the order of their
// relativeHeight, then the text, then the drawings in front of it in the
// order of their relativeHeight - pictures, text boxes and shapes alike. The
// same from the commands of the whole document (the PDF export, the native
// painters) and from the chunks the browser paints.
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

open Rbt.Word.Document
open Rbt.Word.Layout
module TB = Rbt.Word.TextBuffer
module Paint = Rbt.Word.Paint

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-50s expected %A got %A" label expected got

let settings : LayoutSettings =
    { Width = 816.0; PageHeight = 1056.0; FontSize = 16.0; LineHeight = 24.0; PaddingLeft = 96.0; PaddingTop = 96.0
      MarginRight = 96.0; MarginBottom = 96.0; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5; DefaultTabStop = 48.0; CompatibilityMode = 12 }
let measureText (text: string) (style: TB.TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = rangeMeasurerFromTextMeasurer measureText
let para id (text: string) : ParagraphBlock = { Id = id; Buffer = TB.ofString text; Metadata = { defaultMetadata with SpacingBefore = Some 0.0; SpacingAfter = Some 0.0 } }
let words n = String.concat " " [ for i in 1 .. n -> sprintf "w%d" i ]

/// A paragraph with a floating picture `src` (no wrapping) of a relativeHeight, behind the text or in front of it.
let withPicture id (text: string) (pictures: (string * bool * string option) list) : ParagraphBlock =
    let buffer =
        pictures
        |> List.fold (fun buffer (src, _, _) ->
            buffer |> TB.insertImageWithMetadata 0 src false (Some 120.0) (Some 90.0) None None None true (Some "none") (Some 0.0) (Some 0.0) None None None None None None None None None None None None None None None None None None) (TB.ofString text)
    let runs =
        buffer.Runs
        |> List.map (function
            | TB.Image image ->
                let _, behind, height = pictures |> List.find (fun (src, _, _) -> src = image.Src)
                TB.Image { image with BehindDoc = (if behind then Some true else None); RelativeHeight = height; Wrap = Some "none" }
            | run -> run)
    { para id text with Buffer = { buffer with Runs = runs } }

let box id (color: string) (height: string) (behind: bool) (anchor: string) : Block =
    TextBox { defaultTextBox with Id = id; Content = [ para (id + "-text") ("in " + id) ]; X = Some 20.0; Y = Some 20.0; FillColor = Some color; BorderColor = None
                                  RelativeHeight = Some height; BehindDoc = (if behind then Some true else None); AnchorBlockId = Some anchor; Wrap = Some "none" }

// The body: a picture behind the text; two pictures in front, the higher one first in the text; text boxes
// between and above them and one behind the text; a picture in a table cell.
let cell (paragraph: ParagraphBlock) : TableCell =
    { Content = [ paragraph ]; NestedTables = []; NestedTableAnchors = []; EnclosingControls = []; ColSpan = 1; BackgroundColor = None; Shading = None; Width = None; PreferredWidth = None
      BorderColor = None; BorderWidth = None; Borders = None; VerticalAlignment = None; VerticalMerge = None; PaddingTop = None; PaddingRight = None; PaddingBottom = None; PaddingLeft = None; Origin = None }
let table : TableBlock =
    { Id = "t"; Rows = [ { Cells = [ cell (withPicture "cell" "cell text" [ "img:cell", false, Some "40" ]) ]; IsHeader = false; Height = None; HeightRule = None; CanSplit = None; CellSpacing = None
                           PaddingTop = None; PaddingRight = None; PaddingBottom = None; PaddingLeft = None; GridBefore = 0; GridAfter = 0; WidthBefore = None; WidthAfter = None; BordersOverride = None; Origin = None } ]
      PreferredWidth = None; Alignment = None; Indent = None; LayoutMode = None; GridColumns = []; Borders = None; BackgroundColor = None; Shading = None; CellSpacing = None
      PaddingTop = None; PaddingRight = None; PaddingBottom = None; PaddingLeft = None; Floating = false; FloatingXml = None; StyleId = None
      ContentControlId = None; ContentControlTag = None; ContentControlAlias = None; ContentControlXml = None
      ContentControlGroup = None; ContentControlAncestors = []; Look = None }
let headerPart : HeaderFooterPart =
    { Kind = "default"; RelationshipId = Some "h1"
      Blocks = [ Paragraph (withPicture "h" "HEADER TEXT" [ "img:watermark", true, Some "1" ; "img:logo", false, Some "2" ]) ] }
let doc : Document =
    { empty with
        Blocks =
            [ Paragraph (withPicture "a" ("alpha " + words 40) [ "img:behind9", true, Some "9"; "img:behind3", true, Some "3" ])
              Paragraph (withPicture "b" ("beta " + words 40) [ "img:front20", false, Some "20"; "img:front10", false, Some "10" ])
              Table table
              Paragraph (para "c" ("gamma " + words 40))
              box "box15" "#111111" "15" false "a"
              box "box30" "#222222" "30" false "c"
              box "boxBehind5" "#333333" "5" true "c" ]
        PageSetup = { defaultPageSetup with PageWidth = 816.0; PageHeight = 1056.0; MarginLeft = 96.0; MarginRight = 96.0; MarginTop = 96.0; MarginBottom = 96.0; HeaderDistance = 48.0; FooterDistance = 48.0; HeaderReferences = [ "default", "h1" ] }
        HeaderParts = [ headerPart ]
        Styles = []
        Selection = None }

/// What a command paints, for the things this test follows (None: something else).
let nameOf (command: Paint.PaintCommand) : string option =
    match command with
    | Paint.DrawImage (src, _, _, _, _, _, _, _, _) -> Some src
    | Paint.FillRect (color, _, _, _, _) when color.StartsWith "#" && color.Length = 7 && color.[1] = color.[2] && color <> "#ffffff" -> Some color
    | Paint.DrawText (text, _, _, _, _) when text.StartsWith "HEADER" -> Some "header-text"
    | Paint.DrawText (text, _, _, _, _) when text.StartsWith "alpha" -> Some "body-text-first"
    | Paint.DrawText (text, _, _, _, _) when text.StartsWith "gamma" -> Some "body-text-last"
    | Paint.DrawText (text, _, _, _, _) when text.StartsWith "in box" -> Some ("text " + text)
    | _ -> None
let expected =
    [ // the header: its picture behind the text, its text, its picture in front
      "img:watermark"; "header-text"; "img:logo"
      // the body's drawings behind the text, the lowest first
      "img:behind3"; "#333333"; "text in boxBehind5"; "img:behind9"
      // the text
      "body-text-first"; "body-text-last"
      // the drawings in front, the lowest first
      "img:front10"; "#111111"; "text in box15"; "img:front20"; "#222222"; "text in box30"; "img:cell" ]

let layouts, pages = layoutDocumentPages settings measureText measureRange doc
let _, commands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange doc layouts pages
// (A text box's text is drawn in a group, clipped to the box.)
let named (commands: Paint.PaintCommand list) = Paint.flattenedCommands commands |> List.choose nameOf
check "the commands of the whole document, in layers" (named commands) expected
// (Without the page geometry too: the path of the PDF export.)
let _, plainCommands, _ = Paint.collectDocumentCommandsWithLayouts settings "Calibri" measureText measureRange doc (Some layouts)
check "the same without the pages given" (named plainCommands) expected

// The chunks the browser paints, in the same order; and after an edit.
let scope = obj ()
let state = createIncrementalLayoutState ()
let cache = Paint.createPaintCache ()
let chunked (doc: Document) =
    let layouts, pages = withTablePlanScope scope (fun () -> layoutDocumentPagesIncremental state settings measureText measureRange doc)
    match state.LastSlices with
    | Some slices ->
        match withTablePlanScope scope (fun () -> Paint.collectDocumentChunks cache settings "Calibri" measureText measureRange doc slices pages) with
        | Some chunks -> Some (Paint.commandsOfChunks chunks, Paint.paintCacheFurnitureCount cache, chunks)
        | None -> None
    | None -> None
match chunked doc with
| Some (chunkCommands, furniture, chunks) ->
    check "the chunks, in layers" (named chunkCommands) expected
    check "the same commands as at once" (List.sort (chunkCommands |> List.map string) = List.sort (commands |> List.map string)) true
    check "the headers and footers are the first chunks" (furniture, Paint.commandsOfChunks chunks.[0 .. furniture - 1] |> List.choose nameOf) (pages.Length, [ "img:watermark"; "header-text"; "img:logo" ])
| None -> check "the document is painted in chunks" false true
let edited = { doc with Blocks = doc.Blocks |> mapParagraphsOfBlocks (fun p -> if p.Id = "c" then para "c" ("gamma delta " + words 30) else p) }
match chunked edited with
| Some (chunkCommands, _, _) ->
    check "the chunks after an edit, in layers" (named chunkCommands) expected
| None -> check "the edited document is painted in chunks" false true

// A document without drawings: the text in the order of the document, the headers before it.
let plain = { doc with Blocks = [ Paragraph (para "a" "alpha one"); Paragraph (para "c" "gamma two") ]; HeaderParts = [ { headerPart with Blocks = [ Paragraph (para "h" "HEADER TEXT") ] } ] }
let _, plainDocCommands, _ = Paint.collectDocumentCommands settings "Calibri" measureText measureRange plain
check "no drawings: header, then text" (plainDocCommands |> List.choose nameOf) [ "header-text"; "body-text-first"; "body-text-last" ]

printfn "word drawing-order: %d passed, %d failed" pass fail
if fail > 0 then exit 1
