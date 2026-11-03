// Footnotes and endnotes are laid out from their blocks (Package.Notes): their
// paragraphs with their formatting, several of them, their tables; the pages
// make room for them by their real heights, a footnote too long for its page
// goes on at the foot of the next, an endnote too long for a page on the next
// page. A note whose reference's text was set on its own (setNote) shows that
// text; a note without a body shows its reference's text.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System.IO
open Rbt.Word.TextBuffer
open Rbt.Word.Layout
module Doc = Rbt.Word.Document

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-58s expected %A got %A" label expected got

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = rangeMeasurerFromTextMeasurer measureText

let pageSetup =
    { Doc.defaultPageSetup with
        PageWidth = 500.0; PageHeight = 400.0; MarginLeft = 0.0; MarginRight = 0.0; MarginTop = 40.0; MarginBottom = 40.0 }
let settings : LayoutSettings =
    { Width = 500.0
      PageHeight = 400.0
      FontSize = 16.0
      LineHeight = 24.0
      PaddingLeft = 0.0
      PaddingTop = 40.0
      MarginRight = 0.0
      MarginBottom = 40.0
      InTableCell = false
      TableContinuationHeaderHeight = 0.0
      LongTrailingSpaceFitSafety = 0.25; DefaultTabStop = 48.0 }
let stride = settings.PageHeight + pageGap
let pageOf (y: float) = int (floor (y / stride))
let contentTopOf page = float page * stride + settings.PaddingTop
let contentBottomOf page = float page * stride + settings.PageHeight - settings.MarginBottom

let paragraph id text : Doc.ParagraphBlock =
    { Id = id; Buffer = ofString text; Metadata = { Doc.defaultMetadata with SpacingAfter = Some 0.0; SpacingBefore = Some 0.0 } }

let tableCell content : Doc.TableCell =
    { Content = content; NestedTables = []; NestedTableAnchors = []; ColSpan = 1; BackgroundColor = None; Shading = None; Width = None
      PreferredWidth = None; BorderColor = None; BorderWidth = None; Borders = None; VerticalAlignment = None; VerticalMerge = None
      PaddingTop = None; PaddingRight = None; PaddingBottom = None; PaddingLeft = None; Origin = None }
let tableRow cells : Doc.TableRow =
    { Cells = cells; IsHeader = false; Height = None; HeightRule = None; CanSplit = None; CellSpacing = None; PaddingTop = None
      PaddingRight = None; PaddingBottom = None; PaddingLeft = None; GridBefore = 0; GridAfter = 0; WidthBefore = None; WidthAfter = None
      BordersOverride = None; Origin = None }
let table id rows : Doc.TableBlock =
    { Id = id; Rows = rows; PreferredWidth = None; Alignment = None; LayoutMode = None; GridColumns = []; Borders = None; BackgroundColor = None
      Shading = None; CellSpacing = None; PaddingTop = None; PaddingRight = None; PaddingBottom = None; PaddingLeft = None; Floating = false
      FloatingXml = None; StyleId = None; Look = None }

/// A note's body as the reader gives it: its blocks, its plain text (that of its paragraphs, joined).
let noteBody noteType id (blocks: Doc.Block list) : Doc.NoteBody =
    { NoteType = noteType; Id = id; Kind = None; HasMark = true; MarkStyle = None; Text = Doc.storyPlainText noteType blocks; Blocks = blocks }

/// A paragraph whose last character refers to a note with that text.
let referring id (text: string) noteType noteId (noteText: string) : Doc.ParagraphBlock =
    { paragraph id text with Buffer = ofString text |> setNote (text.Length - 1) 1 (Some noteId) (Some noteType) (Some noteText) }

let filler index = paragraph (sprintf "fill-%d" index) (sprintf "filler paragraph number %d of the body" index)

let documentWith (blocks: Doc.Block list) (notes: Doc.NoteBody list) : Doc.Document =
    { Doc.empty with Blocks = blocks; PageSetup = pageSetup; Selection = None; Package = { Doc.emptyPackageData with Notes = notes } }

let layOut (doc: Doc.Document) = layoutDocumentPages settings measureText measureRange doc

let isFootnote (layout: BlockLayout) = layout.BlockId.StartsWith footnoteBlockIdPrefix
let isEndnote (layout: BlockLayout) = layout.BlockId.StartsWith endnoteBlockIdPrefix
let isBody layout = not (isFootnote layout) && not (isEndnote layout)
let textOf (layout: BlockLayout) =
    layout.Lines |> List.collect (fun line -> line.Segments) |> List.map (function TextSegment (text, _, _, _, _) -> text | _ -> "") |> String.concat ""
/// The note's own paragraph a layout of a note area is laid out from (the block id's last part; see noteParagraphIdOfLayoutId).
let ownId (blockId: string) =
    let parts = blockId.Split('#')
    if parts.Length >= 3 then Some (System.String.Join("#", parts.[2..])) else None
let ownIds (layouts: BlockLayout list) = layouts |> List.map (fun layout -> ownId layout.BlockId |> Option.defaultValue "?")
let bottomOf (layouts: BlockLayout list) = layouts |> List.map (fun layout -> layout.Y + layout.Height) |> List.fold max 0.0
let topOf (layouts: BlockLayout list) = layouts |> List.map (fun layout -> layout.Y) |> List.fold min System.Double.MaxValue
let bodyLinesOnPage page (layouts: BlockLayout list) =
    layouts |> List.filter isBody |> List.collect (fun layout -> layout.Lines) |> List.filter (fun line -> pageOf line.Y = page)

// ---- a footnote of several formatted paragraphs ----------------------------------

let boldFirst = { paragraph "n1-p1" " First paragraph of the note" with Buffer = ofString " First paragraph of the note" |> applyStyle 1 5 (fun style -> { style with Bold = true; Explicit = style.Explicit ||| ExplicitFormat.Bold }) }
let bigSecond = { paragraph "n1-p2" "Second paragraph" with Buffer = ofString "Second paragraph" |> applyStyle 0 16 (fun style -> { style with FontSize = 30.0; Explicit = style.Explicit ||| ExplicitFormat.FontSize }) }
let formattedNote = noteBody "footnote" "1" [ Doc.Paragraph boldFirst; Doc.Paragraph bigSecond; Doc.Paragraph (paragraph "n1-p3" "Third") ]
let formattedDoc =
    documentWith [ Doc.Paragraph (referring "ref-1" "A body paragraph with a note x" "footnote" "1" formattedNote.Text); Doc.Paragraph (filler 1) ] [ formattedNote ]
let formattedLayouts, _ = layOut formattedDoc
let formattedArea = formattedLayouts |> List.filter isFootnote
check "the footnote is its three paragraphs" (ownIds formattedArea) [ "n1-p1"; "n1-p2"; "n1-p3" ]
check "each is of the note" (formattedArea |> List.map (fun layout -> noteOfLayoutId layout.BlockId) |> List.distinct) [ Some ("footnote", "1") ]
check "its first paragraph has the mark before its text" (formattedArea.Head.Prefix, formattedArea.Head.PrefixStyle |> Option.bind (fun style -> style.VertAlign)) ("1", Some "superscript")
check "the mark is not part of the text" (textOf formattedArea.Head) " First paragraph of the note"
check "the text starts after the mark"
    (formattedArea.Head.Lines.Head.Segments |> List.forall (function TextSegment (_, _, x, _, offset) -> offset >= 0 && x >= formattedArea.Head.PrefixWidth - 0.01 | _ -> true))
    true
check "the bold word is bold"
    (formattedArea.Head.Lines |> List.collect (fun line -> line.Segments) |> List.exists (function TextSegment (text, style, _, _, _) -> text.Contains "First" && style.Bold | _ -> false))
    true
let lineHeightOf (layout: BlockLayout) = layout.Lines.Head.Height
match formattedArea with
| [ first; second; third ] ->
    check "the larger paragraph has the taller line" (lineHeightOf second > lineHeightOf third + 2.0) true
    check "the paragraphs are stacked" (second.Y >= first.Y + first.Height - 0.01 && third.Y >= second.Y + second.Height - 0.01) true
| _ -> check "the note's paragraphs are laid out one by one" formattedArea.Length 3
check "the note ends above the bottom margin" (bottomOf formattedArea <= contentBottomOf 0 + 0.01) true
check "the note is on its reference's page" (formattedArea |> List.forall (fun layout -> pageOf layout.Y = 0)) true

// ---- the pages make room for the note's real height -----------------------------

let fillers count = [ for index in 1 .. count -> Doc.Paragraph (filler index) ]
let tallNoteBlocks = [ for index in 1 .. 6 -> Doc.Paragraph (paragraph (sprintf "tall-%d" index) (sprintf "note paragraph %d" index)) ]
let tallNote = noteBody "footnote" "1" tallNoteBlocks
let shortNote = noteBody "footnote" "1" [ Doc.Paragraph (paragraph "short-1" (tallNote.Text)) ]
let roomDoc (note: Doc.NoteBody) = documentWith (Doc.Paragraph (referring "ref-room" "The reference x" "footnote" "1" note.Text) :: fillers 30) [ note ]
let tallLayouts, _ = layOut (roomDoc tallNote)
let shortLayouts, _ = layOut (roomDoc shortNote)
let tallArea = tallLayouts |> List.filter (fun layout -> isFootnote layout && pageOf layout.Y = 0)
check "the six paragraphs of the note are on the page" tallArea.Length 6
check "the page's text ends above its note" ((bodyLinesOnPage 0 tallLayouts |> List.map (fun line -> line.Y + line.Height) |> List.max) <= topOf tallArea + 0.01) true
check "a page with a taller note holds fewer lines of text" ((bodyLinesOnPage 0 tallLayouts).Length < (bodyLinesOnPage 0 shortLayouts).Length) true
// Old: the room was the plain text's line count, the same for both notes.
check "it holds as many fewer as the note has lines more"
    ((bodyLinesOnPage 0 shortLayouts).Length - (bodyLinesOnPage 0 tallLayouts).Length >= 3)
    true

// ---- a footnote with a table --------------------------------------------------------

let noteTable = table "nt" [ tableRow [ tableCell [ paragraph "nt-a" "cell left" ]; tableCell [ paragraph "nt-b" "cell right" ] ] ]
let tableNote = noteBody "footnote" "1" [ Doc.Paragraph (paragraph "tn-1" " A note with a table"); Doc.Table noteTable; Doc.Paragraph (paragraph "tn-2" "after the table") ]
let tableLayouts, _ = layOut (documentWith [ Doc.Paragraph (referring "ref-t" "Text x" "footnote" "1" tableNote.Text); Doc.Paragraph (filler 1) ] [ tableNote ])
let tableArea = tableLayouts |> List.filter isFootnote
check "the note's paragraphs and cells are laid out" (ownIds tableArea |> List.sort) [ "nt-a"; "nt-b"; "tn-1"; "tn-2" ]
let firstX (layout: BlockLayout) = layout.Lines.Head.Segments |> List.pick (function TextSegment (_, _, x, _, _) -> Some x | _ -> None)
match [ "nt-a"; "nt-b"; "tn-1"; "tn-2" ] |> List.map (fun id -> tableArea |> List.tryFind (fun layout -> ownId layout.BlockId = Some id)) with
| [ Some cellA; Some cellB; Some first; Some last ] ->
    check "the cells are side by side" (abs (cellA.Lines.Head.Y - cellB.Lines.Head.Y) < 0.01 && firstX cellB > firstX cellA + 100.0) true
    check "the table is below the first paragraph" (cellA.Y >= first.Y + first.Height - 0.01) true
    check "the last paragraph is below the table" (last.Y >= cellA.Y + cellA.Height - 0.01) true
| _ -> check "the note's table is laid out" false true
check "the note with its table ends above the bottom margin" (bottomOf tableArea <= contentBottomOf 0 + 0.01) true

// ---- a footnote too long for its page goes on on the next -------------------------

let longNote = noteBody "footnote" "1" [ for index in 1 .. 30 -> Doc.Paragraph (paragraph (sprintf "long-%d" index) (sprintf "long note line %d" index)) ]
let longLayouts, longPages = layOut (documentWith (Doc.Paragraph (referring "ref-long" "Ref x" "footnote" "1" longNote.Text) :: fillers 40) [ longNote ])
let longArea = longLayouts |> List.filter isFootnote
check "every paragraph of the long note is laid out" (ownIds longArea |> List.distinct |> List.length) 30
check "the long note starts on its reference's page" (pageOf (topOf longArea)) 0
check "and goes on on the next" (longArea |> List.exists (fun layout -> pageOf layout.Y = 1)) true
check "no part of it is outside a page's text area"
    (longArea |> List.forall (fun layout -> let page = pageOf layout.Y in layout.Y >= contentTopOf page - 0.01 && layout.Y + layout.Height <= contentBottomOf page + 0.01))
    true
check "its paragraphs are in order" (ownIds longArea) [ for index in 1 .. 30 -> sprintf "long-%d" index ]
let carriedTop = longArea |> List.filter (fun layout -> pageOf layout.Y = 1) |> topOf
check "the next page's text ends above what goes on there" ((bodyLinesOnPage 1 longLayouts |> List.map (fun line -> line.Y + line.Height) |> List.fold max 0.0) <= carriedTop + 0.01) true

// ---- endnotes -------------------------------------------------------------------------

let endnoteA = noteBody "endnote" "1" [ Doc.Paragraph boldFirst; Doc.Paragraph (paragraph "en-a2" "second paragraph of the endnote") ]
let endnoteB = noteBody "endnote" "2" [ for index in 1 .. 40 -> Doc.Paragraph (paragraph (sprintf "en-b%d" index) (sprintf "endnote two line %d" index)) ]
let endnoteDoc =
    documentWith
        [ Doc.Paragraph (referring "ref-e1" "First endnote x" "endnote" "1" endnoteA.Text); Doc.Paragraph (referring "ref-e2" "Second endnote x" "endnote" "2" endnoteB.Text) ]
        [ endnoteA; endnoteB ]
let endnoteLayouts, endnotePages = layOut endnoteDoc
let endnotes = endnoteLayouts |> List.filter isEndnote
check "the endnotes are their paragraphs" (ownIds endnotes) ([ "n1-p1"; "en-a2" ] @ [ for index in 1 .. 40 -> sprintf "en-b%d" index ])
check "the first endnote has its mark" (endnotes.Head.Prefix, textOf endnotes.Head) ("i", " First paragraph of the note")
check "the endnotes come after the text" (topOf endnotes >= bottomOf (endnoteLayouts |> List.filter isBody)) true
check "the long endnote goes on over pages" (endnotes |> List.filter (fun layout -> noteOfLayoutId layout.BlockId = Some ("endnote", "2")) |> List.map (fun layout -> pageOf layout.Y) |> List.distinct |> List.length >= 2) true
check "every endnote paragraph is within a page's text area"
    (endnotes |> List.forall (fun layout -> let page = pageOf layout.Y in layout.Y >= contentTopOf page - 0.01 && layout.Y + layout.Height <= contentBottomOf page + 0.01))
    true
check "the pages cover the endnotes" (bottomOf endnotes <= (List.last endnotePages).Top + (List.last endnotePages).Height) true

// ---- the text of the reference -------------------------------------------------------

// The reference's text set on its own (setNote): the note shows that text.
let setNoteLayouts, _ = layOut (documentWith [ Doc.Paragraph (referring "ref-s" "Text x" "footnote" "1" "a text set on the reference") ] [ formattedNote ])
let setNoteArea = setNoteLayouts |> List.filter isFootnote
check "a text set on the reference is shown" (setNoteArea |> List.map textOf) [ " a text set on the reference" ]
check "with the mark" setNoteArea.Head.Prefix "1"
// No body: the reference's text.
let bareLayouts, _ = layOut (documentWith [ Doc.Paragraph (referring "ref-b" "Text x" "footnote" "7" "a note without a body") ] [])
check "a note without a body shows its reference's text" (bareLayouts |> List.filter isFootnote |> List.map textOf) [ " a note without a body" ]

// ---- an edit of the note's body (incremental layout) -----------------------------------

let state = createIncrementalLayoutState ()
let incremental doc = layoutDocumentPagesIncremental state settings measureText measureRange doc |> fst
let before = incremental formattedDoc |> List.filter isFootnote
let edited = Doc.withStoryBody "footnote" "1" [ Doc.Paragraph boldFirst; Doc.Paragraph (paragraph "n1-p2" "Second paragraph, changed"); Doc.Paragraph (paragraph "n1-p3" "Third"); Doc.Paragraph (paragraph "n1-p4" "Fourth") ] formattedDoc
let after = incremental edited |> List.filter isFootnote
check "the edited note is laid out again" (ownIds after) [ "n1-p1"; "n1-p2"; "n1-p3"; "n1-p4" ]
check "with its new text" (after |> List.tryItem 1 |> Option.map textOf) (Some "Second paragraph, changed")
check "as the whole document laid out at once" (after |> List.map (fun layout -> layout.BlockId, layout.Y, textOf layout)) (layOut edited |> fst |> List.filter isFootnote |> List.map (fun layout -> layout.BlockId, layout.Y, textOf layout))
// Formatting only (the plain text is the same): the note is laid out again too.
let bolded = Doc.withStoryBody "footnote" "1" [ Doc.Paragraph boldFirst; Doc.Paragraph bigSecond; Doc.Paragraph { paragraph "n1-p3" "Third" with Buffer = ofString "Third" |> applyStyle 0 5 (fun style -> { style with FontSize = 40.0; Explicit = style.Explicit ||| ExplicitFormat.FontSize }) } ] formattedDoc
let _ = incremental formattedDoc
let bigThird = incremental bolded |> List.filter isFootnote
check "a note made larger takes more room" (match List.tryItem 2 bigThird, List.tryItem 2 before with Some big, Some small -> big.Lines.Head.Height > small.Lines.Head.Height + 5.0 | _ -> false) true
check "the larger note is as laid out at once" (bigThird |> List.map (fun layout -> layout.BlockId, layout.Y)) (layOut bolded |> fst |> List.filter isFootnote |> List.map (fun layout -> layout.BlockId, layout.Y))

// ---- a document read from a file -------------------------------------------------------

let richNotes = Path.Combine(__SOURCE_DIRECTORY__, "roundtrip-saved", "rich-notes.original.docx")
match File.ReadAllBytes richNotes |> DocxReader.loadBytesNative |> Option.map Rbt.Word.Import.importDoc with
| None -> check "rich-notes reads" false true
| Some doc ->
    let layouts, _ = layoutDocumentPages { settings with Width = doc.PageSetup.PageWidth; PageHeight = doc.PageSetup.PageHeight; PaddingLeft = doc.PageSetup.MarginLeft; MarginRight = doc.PageSetup.MarginRight; PaddingTop = doc.PageSetup.MarginTop; MarginBottom = doc.PageSetup.MarginBottom } measureText measureRange doc
    let note = layouts |> List.filter (fun layout -> noteOfLayoutId layout.BlockId = Some ("footnote", "2"))
    let texts = note |> List.map textOf
    check "rich-notes: the footnote's bold words are bold"
        (note |> List.collect (fun layout -> layout.Lines) |> List.collect (fun line -> line.Segments) |> List.exists (function TextSegment (text, style, _, _, _) -> text.Contains "bold source" && style.Bold | _ -> false))
        true
    check "rich-notes: the footnote's table cells are laid out" (texts |> List.filter (fun text -> text.StartsWith "note cell")) [ "note cell a"; "note cell b" ]
    let cell name = note |> List.tryFind (fun layout -> textOf layout = name) |> Option.map (fun layout -> layout.Y)
    check "rich-notes: side by side" (match cell "note cell a", cell "note cell b" with Some a, Some b -> abs (a - b) < 0.01 | _ -> false) true
    check "rich-notes: the paragraph after the table" (texts |> List.last) "after the note table"

printfn ""
printfn "Word note layout tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
