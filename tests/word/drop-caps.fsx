// Drop caps (w:framePr w:dropCap). The drop cap is a paragraph of its own -
// the large first letter - that Word places at the top of the paragraph
// after it: it takes no room of its own, it is as high as the number of
// lines it says (w:lines), and those lines of the paragraph go around it,
// w:hSpace from it ("drop"); or it stands in the margin to their left
// ("margin"). Seven documents of the corpus have one (the calibre demo and
// its copies, sample-formatted, wikipedia-sample, microsoft-word-sample,
// complex-tables); they were laid out as a paragraph holding the letter,
// with the text below it.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open Rbt.Word.TextBuffer
module Doc = Rbt.Word.Document
module Layout = Rbt.Word.Layout

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-60s expected %A got %A" label expected got

let w = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
let package (body: string) =
    use stream = new IO.MemoryStream()
    do
        use archive = new IO.Compression.ZipArchive(stream, IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new IO.StreamWriter(archive.CreateEntry(name).Open(), Text.UTF8Encoding(false))
            writer.Write text
        add "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"""
        add "word/document.xml" (sprintf "<w:document %s><w:body>%s<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>" w body)
    stream.ToArray()
let import body = package body |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let settingsOf (doc: Doc.Document) : Layout.LayoutSettings =
    let s = doc.PageSetup
    { Width = s.PageWidth; PageHeight = s.PageHeight; FontSize = 16.0; LineHeight = 24.0; PaddingLeft = s.MarginLeft; PaddingTop = s.MarginTop
      MarginRight = s.MarginRight; MarginBottom = s.MarginBottom; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5; DefaultTabStop = s.DefaultTabStop; CompatibilityMode = s.CompatibilityMode }
let layoutsOf (doc: Doc.Document) = Layout.layoutDocumentPages (settingsOf doc) measureText measureRange doc |> fst
let paragraphs (doc: Doc.Document) = doc.Blocks |> List.choose Doc.tryGetParagraph

let para (text: string) = sprintf "<w:p><w:r><w:t xml:space=\"preserve\">%s</w:t></w:r></w:p>" text
/// The drop cap as Word writes it: the letter at 58.5 pt, an exact line height of three lines.
let dropCap kind (hSpace: int) (letter: string) =
    sprintf "<w:p><w:pPr><w:keepNext/><w:framePr w:dropCap=\"%s\" w:lines=\"3\" w:hSpace=\"%d\" w:wrap=\"around\" w:vAnchor=\"text\" w:hAnchor=\"text\"/><w:spacing w:line=\"951\" w:lineRule=\"exact\"/><w:textAlignment w:val=\"baseline\"/><w:rPr><w:position w:val=\"-10\"/><w:sz w:val=\"117\"/></w:rPr></w:pPr><w:r><w:rPr><w:position w:val=\"-10\"/><w:sz w:val=\"117\"/></w:rPr><w:t>%s</w:t></w:r></w:p>" kind hSpace letter
let longText = String.replicate 120 "word "

let firstX (line: Layout.MeasuredLine) =
    match line.Segments with
    | Layout.TextSegment (_, _, x, _, _) :: _ -> Math.Round(x, 2)
    | _ -> nan
let layoutOf (layouts: Layout.BlockLayout list) (para: Doc.ParagraphBlock) = layouts |> List.find (fun layout -> layout.BlockId = para.Id)

// ---- read ---------------------------------------------------------------------
let dropped = import (para "Before." + dropCap "drop" 144 "D" + para ("rop " + longText) + para "After.")
let droppedParas = paragraphs dropped
check "read: the drop cap's frame (drop, three lines, hSpace 144 twips = 9.6 px)"
    droppedParas.[1].Metadata.DropCap (Some { Doc.InMargin = false; Doc.Lines = 3; Doc.Distance = 9.6 })
check "read: the paragraphs around it have none" (droppedParas.[0].Metadata.DropCap, droppedParas.[2].Metadata.DropCap) (None, None)

// ---- placed -------------------------------------------------------------------
let layouts = layoutsOf dropped
let cap = layoutOf layouts droppedParas.[1]
let text = layoutOf layouts droppedParas.[2]
let capWidth = cap.Lines |> List.map (fun line -> line.Width) |> List.max
let withoutCap = import (para "Before." + para ("rop " + longText) + para "After.")
let textWithoutCap = layoutOf (layoutsOf withoutCap) (paragraphs withoutCap).[1]
check "the drop cap takes no room: the paragraph after it starts where it would without it"
    (Math.Round(text.Lines.Head.Y, 2)) (Math.Round(textWithoutCap.Lines.Head.Y, 2))
check "the paragraph's first three lines stand beside the letter, w:hSpace from it"
    (text.Lines |> List.take 3 |> List.map firstX) (List.replicate 3 (Math.Round(capWidth + 9.6, 2)))
check "its fourth line is at the indent again" (firstX text.Lines.[3]) 0.0
check "the text goes on from line to line (nothing left out beside the letter)"
    (text.Lines |> List.pairwise |> List.forall (fun (a, b) -> a.EndOffset = b.StartOffset)) true
let capLast = List.last cap.Lines
check "the letter is as high as three of the lines: its line ends where the third ends"
    (Math.Round(capLast.Y + capLast.Height, 2)) (Math.Round(text.Lines.[2].Y + text.Lines.[2].Height, 2))
check "the letter is drawn before the paragraph it belongs to"
    (layouts |> List.findIndex (fun layout -> layout.BlockId = cap.BlockId) < (layouts |> List.findIndex (fun layout -> layout.BlockId = text.BlockId))) true

// ---- in the margin --------------------------------------------------------------
let inMargin = import (para "Before." + dropCap "margin" 0 "M" + para ("argin " + longText))
let marginLayouts = layoutsOf inMargin
let marginCap = layoutOf marginLayouts (paragraphs inMargin).[1]
let marginText = layoutOf marginLayouts (paragraphs inMargin).[2]
check "w:dropCap margin: the letter stands left of the text, in the margin"
    (marginCap.Lines |> List.forall (fun line -> line.Segments |> List.forall (function Layout.TextSegment (_, _, x, width, _) -> x + width <= 0.01 | _ -> true))) true
check "w:dropCap margin: the text's lines are all at the indent" (marginText.Lines |> List.map firstX |> List.distinct) [ 0.0 ]

// ---- not a drop cap where nothing follows it ---------------------------------------
let alone = import (para "Before." + dropCap "drop" 0 "Z")
let aloneLayouts = layoutsOf alone
check "a drop cap with no paragraph after it is laid out as a paragraph"
    ((layoutOf aloneLayouts (paragraphs alone).[1]).Y >= (layoutOf aloneLayouts (paragraphs alone).[0]).Y + 1.0) true

// ---- the incremental layout is the cold one -----------------------------------------
let describe (layouts: Layout.BlockLayout list) =
    layouts |> List.map (fun layout -> layout.BlockId, Math.Round(layout.Y, 2), layout.Lines |> List.map (fun line -> Math.Round(line.Y, 2), firstX line))
let state = Layout.createIncrementalLayoutState ()
let incremental (doc: Doc.Document) = Layout.layoutDocumentPagesIncremental state (settingsOf doc) measureText measureRange doc |> fst
check "incremental: as cold" (describe (incremental dropped)) (describe layouts)
let editedText (doc: Doc.Document) index (text: string) =
    { doc with
        Blocks =
            doc.Blocks |> List.mapi (fun i block ->
                match block with
                | Doc.Paragraph p when i = index -> Doc.Paragraph { p with Buffer = ofString text }
                | _ -> block) }
let editedAfter = editedText dropped 2 ("rop " + String.replicate 10 "word ")
check "incremental: the paragraph after the letter edited" (describe (incremental editedAfter)) (describe (layoutsOf editedAfter))
let editedBefore = editedText editedAfter 0 (String.replicate 30 "Before ")
check "incremental: a paragraph before the letter edited" (describe (incremental editedBefore)) (describe (layoutsOf editedBefore))

// ---- a letter taller than its lines ---------------------------------------------------
// A letter with no exact line height (field-notes.docx: 40 pt, the document's 1.15 lines)
// can be higher than the lines it says: Word's frame is then as high as the letter's line,
// from the top of the first line, and the lines that start beside it go around it - four
// of them there. (Word writes an exact line height three lines high, as above.)
let tallCap (letter: string) =
    sprintf "<w:p><w:pPr><w:framePr w:dropCap=\"drop\" w:lines=\"3\" w:wrap=\"around\" w:vAnchor=\"text\" w:hAnchor=\"text\"/><w:spacing w:line=\"276\" w:lineRule=\"auto\"/></w:pPr><w:r><w:rPr><w:sz w:val=\"80\"/></w:rPr><w:t>%s</w:t></w:r></w:p>" letter
let tall = import (para "Before." + tallCap "T" + para ("all " + longText))
let tallLayouts = layoutsOf tall
let tallCapLayout = layoutOf tallLayouts (paragraphs tall).[1]
let tallText = layoutOf tallLayouts (paragraphs tall).[2]
let tallFirst = tallCapLayout.Lines.Head
let tallLast = List.last tallCapLayout.Lines
let tallHeight = tallLast.Y + tallLast.Height - tallFirst.Y
let besideCount = tallText.Lines |> List.takeWhile (fun line -> firstX line > 0.0) |> List.length
check "a tall letter is higher than three lines" (tallHeight > 3.0 * tallText.Lines.Head.Height) true
check "a tall letter: its top at the first line's" (Math.Round(tallFirst.Y, 2)) (Math.Round(tallText.Lines.Head.Y, 2))
check "a tall letter: every line that starts beside it goes around it"
    besideCount (tallText.Lines |> List.filter (fun line -> line.Y < tallFirst.Y + tallHeight - 0.5) |> List.length)
check "a tall letter: more than three lines beside it" (besideCount > 3) true

// ---- in a document with sections --------------------------------------------------
// A document with a section break is laid out section by section, in columns
// (Layout.flowColumnsKept): there, too, in every section - the first, one after a
// continuous break, and one of two columns - the paragraph after the letter goes
// around it. (field-notes.docx: the letter was a paragraph of its own above the text.)
let sectionEnd (columns: int) =
    sprintf "<w:p><w:pPr><w:sectPr><w:type w:val=\"continuous\"/><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/><w:cols w:num=\"%d\" w:space=\"720\"/></w:sectPr></w:pPr></w:p>" columns
let sectioned =
    import (dropCap "drop" 144 "O" + para ("ne " + longText) + sectionEnd 1
            + dropCap "drop" 144 "T" + para ("wo " + longText) + sectionEnd 1
            + dropCap "drop" 144 "C" + para ("olumns " + longText) + sectionEnd 2
            + para "Last.")
let sectionedParas = paragraphs sectioned
let sectionedLayouts = layoutsOf sectioned
for capIndex, name in [ 0, "the first section"; 3, "after a continuous break"; 6, "in two columns" ] do
    let capLayouts = sectionedLayouts |> List.filter (fun layout -> layout.BlockId = sectionedParas.[capIndex].Id)
    let textLayout = layoutOf sectionedLayouts sectionedParas.[capIndex + 1]
    match capLayouts with
    | [ capLayout ] ->
        let width = capLayout.Lines |> List.map (fun line -> line.Width) |> List.max
        let capX = firstX capLayout.Lines.Head
        check (name + ": the first three lines stand beside the letter")
            (textLayout.Lines |> List.take 3 |> List.map firstX) (List.replicate 3 (Math.Round(capX + width + 9.6, 2)))
        check (name + ": the fourth is at the indent again") (firstX textLayout.Lines.[3]) capX
        let capLast = List.last capLayout.Lines
        check (name + ": the letter's line ends where the third line ends")
            (Math.Round(capLast.Y + capLast.Height, 2)) (Math.Round(textLayout.Lines.[2].Y + textLayout.Lines.[2].Height, 2))
        check (name + ": the letter takes no room: it starts below the text above, not above the paragraph's first line")
            (capLayout.Y >= textLayout.Lines.Head.Y - 0.01 || capLast.Y + capLast.Height > textLayout.Lines.Head.Y) true
    | found -> check (name + ": the letter is laid out once") found.Length 1
check "incremental (sections): as cold" (describe (incremental sectioned)) (describe sectionedLayouts)

printfn "word drop caps: %d passed, %d failed" pass fail
if fail > 0 then exit 1
