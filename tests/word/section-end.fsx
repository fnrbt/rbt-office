// The empty paragraph that holds a section break (its pPr has the w:sectPr).
// Checked against Word 2019, compatibility modes 14 and 15, with "Alpha" / that paragraph /
// "Bravo" (the first paragraph of the next section):
// - empty (nothing, a bookmark, a run without text), before a continuous break: it takes no
//   room - Bravo is one line below Alpha, whatever spacing before or after, line spacing
//   (double, exact 40pt) or mark size (28pt) the paragraph has;
// - its spacing after still collapses with Bravo's spacing before: Bravo stands as far lower
//   as its spacing before is larger (after 12pt, before 24pt: 12pt; after 24pt, before 12pt:
//   nothing); Alpha's spacing after and Bravo's spacing before both apply;
// - with text, a space or a tab in it: a line, and its spacing after collapses with Bravo's
//   spacing before to the larger of the two (as between paragraphs of one section);
// - in a balanced two-column section it is not part of what is balanced (four lines: two and
//   two);
// - before a next-page, next-column or odd-page break it takes no room either: with the page
//   full above it, it does not go onto a page of its own (with text in it, it does); and on a
//   bottom-aligned page it does not lift the text above it;
// - a continuous section that starts below a full page starts at the top of the next one.
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
let page = "<w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>"
let sectPr (kind: string) (columns: int) = sprintf "<w:sectPr><w:type w:val=\"%s\"/>%s<w:cols w:num=\"%d\" w:space=\"720\"/></w:sectPr>" kind page columns
let package (body: string) (finalSectPr: string) =
    use stream = new IO.MemoryStream()
    do
        use archive = new IO.Compression.ZipArchive(stream, IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new IO.StreamWriter(archive.CreateEntry(name).Open(), Text.UTF8Encoding(false))
            writer.Write text
        add "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"""
        add "word/document.xml" (sprintf "<w:document %s><w:body>%s%s</w:body></w:document>" w body finalSectPr)
    stream.ToArray()
let import body finalSectPr = package body finalSectPr |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let noSpacing = "<w:spacing w:before=\"0\" w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/>"
let para (text: string) = sprintf "<w:p><w:pPr>%s</w:pPr><w:r><w:t xml:space=\"preserve\">%s</w:t></w:r></w:p>" noSpacing text
/// The paragraph ending the first (one-column) section, with its pPr extras and its runs.
/// (With no spacing unless it says.)
let sectionEnd (pPr: string) (runs: string) columns =
    let pPr = if pPr.Contains "w:spacing" then pPr else noSpacing + pPr
    sprintf "<w:p><w:pPr>%s%s</w:pPr>%s</w:p>" pPr (sectPr "nextPage" columns) runs

let layoutOf (doc: Doc.Document) =
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    settings, layouts, pages
let paragraphNamed (doc: Doc.Document) (prefix: string) =
    doc.Blocks |> List.choose Doc.tryGetParagraph |> List.find (fun p -> (toString p.Buffer).StartsWith prefix)
let firstLineY (doc: Doc.Document) layouts prefix =
    let p = paragraphNamed doc prefix
    let layout = layouts |> List.find (fun (l: Layout.BlockLayout) -> l.BlockId = p.Id)
    layout.Lines.Head.Y

let paraWith (pPr: string) (text: string) = sprintf "<w:p><w:pPr>%s</w:pPr><w:r><w:t xml:space=\"preserve\">%s</w:t></w:r></w:p>" pPr text
/// How far below Alpha's first line Bravo's first line stands (with spacing of their own).
let gapWith (alphaPPr: string) (breakPPr: string) (breakRuns: string) (bravoPPr: string) (nextKind: string) =
    let doc = import (paraWith alphaPPr "Alpha line" + sectionEnd breakPPr breakRuns 1 + paraWith bravoPPr "Bravo line") (sectPr nextKind 1)
    let _, layouts, _ = layoutOf doc
    firstLineY doc layouts "Bravo" - firstLineY doc layouts "Alpha"
let gap breakPPr breakRuns nextKind = gapWith noSpacing breakPPr breakRuns noSpacing nextKind

let lineGap = gap "" "<w:r><w:t>XXX</w:t></w:r>" "continuous" / 2.0
let near (a: float) (b: float) = abs (a - b) < 0.5
check "a line is taller than nothing" (lineGap > 5.0) true
check "empty, continuous: no room" (near (gap "" "" "continuous") lineGap) true
check "empty, continuous, spacing before 24pt: no room" (near (gap "<w:spacing w:before=\"480\" w:after=\"0\"/>" "" "continuous") lineGap) true
check "empty, continuous, spacing after 24pt: no room" (near (gap "<w:spacing w:before=\"0\" w:after=\"480\"/>" "" "continuous") lineGap) true
check "empty, continuous, before 12pt and after 24pt: no room" (near (gap "<w:spacing w:before=\"240\" w:after=\"480\"/>" "" "continuous") lineGap) true
check "empty, continuous, exact 40pt lines: no room" (near (gap "<w:spacing w:after=\"0\" w:line=\"800\" w:lineRule=\"exact\"/>" "" "continuous") lineGap) true
check "empty, continuous, 28pt mark: no room" (near (gap "<w:rPr><w:sz w:val=\"56\"/></w:rPr>" "" "continuous") lineGap) true
check "with text, continuous: a line" (near (gap "" "<w:r><w:t>XXX</w:t></w:r>" "continuous") (2.0 * lineGap)) true
check "with text, continuous, spacing after 24pt: a line and the spacing" (near (gap "<w:spacing w:before=\"0\" w:after=\"480\"/>" "<w:r><w:t>XXX</w:t></w:r>" "continuous") (2.0 * lineGap + 32.0)) true
check "with a space, continuous: a line" (near (gap "" "<w:r><w:t xml:space=\"preserve\"> </w:t></w:r>" "continuous") (2.0 * lineGap)) true

// The paragraph's spacing after still collapses with the spacing before of the first paragraph
// of the next section (to the larger of the two): Bravo stands below Alpha by a line and by as
// much as its spacing before is larger. Between paragraphs with text across the break, the
// spacing collapses as within a section.
let spacing (before: int option) (after: int option) =
    sprintf "<w:spacing%s%s/>" (match before with Some b -> sprintf " w:before=\"%d\"" b | None -> " w:before=\"0\"") (match after with Some a -> sprintf " w:after=\"%d\"" a | None -> " w:after=\"0\"")
let pt (points: float) = points * 96.0 / 72.0
check "empty after 12pt, Bravo before 24pt: 12pt" (near (gapWith noSpacing (spacing None (Some 240)) "" (spacing (Some 480) None) "continuous") (lineGap + pt 12.0)) true
check "empty after 24pt, Bravo before 36pt: 12pt" (near (gapWith noSpacing (spacing None (Some 480)) "" (spacing (Some 720) None) "continuous") (lineGap + pt 12.0)) true
check "empty after 24pt, Bravo before 12pt: nothing" (near (gapWith noSpacing (spacing None (Some 480)) "" (spacing (Some 240) None) "continuous") lineGap) true
check "empty before 24pt, Bravo before 12pt: 12pt" (near (gapWith noSpacing (spacing (Some 480) None) "" (spacing (Some 240) None) "continuous") (lineGap + pt 12.0)) true
check "empty, Bravo before 12pt: 12pt" (near (gapWith noSpacing "" "" (spacing (Some 240) None) "continuous") (lineGap + pt 12.0)) true
check "Alpha after 12pt, empty, Bravo before 24pt: 36pt" (near (gapWith (spacing None (Some 240)) "" "" (spacing (Some 480) None) "continuous") (lineGap + pt 36.0)) true
check "text after 12pt, Bravo before 24pt: 24pt" (near (gapWith noSpacing (spacing None (Some 240)) "<w:r><w:t>XXX</w:t></w:r>" (spacing (Some 480) None) "continuous") (2.0 * lineGap + pt 24.0)) true

// A two-column section of four one-line paragraphs and the empty one, before a continuous
// break: two lines in each column (the empty paragraph is not part of what is balanced).
let balancedLines =
    let doc = import (String.concat "" [ for i in 1 .. 4 -> para (sprintf "Alpha %d" i) ] + sectionEnd "" "" 2 + para "Bravo line") (sectPr "continuous" 1)
    let _, layouts, _ = layoutOf doc
    (firstLineY doc layouts "Bravo" - firstLineY doc layouts "Alpha 1") / lineGap
check "balanced: two lines a column" (abs (balancedLines - 2.0) < 0.05) true

// After a full page: the empty paragraph does not go onto a page of its own (before a
// next-page break either), and the continuous section starts at the top of the next page.
let pageCount (fill: int) (nextKind: string) =
    let body = String.concat "" [ for i in 1 .. fill -> para (sprintf "Fill line %d" i) ]
    let doc = import (body + sectionEnd "" "" 1 + para "Bravo line") (sectPr nextKind 1)
    let settings, layouts, pages = layoutOf doc
    let bravo = firstLineY doc layouts "Bravo"
    let index = pages |> List.findIndexBack (fun page -> page.Top <= bravo)
    // (Where on its page Bravo stands, against where the first line of page 1 does.)
    let atTop = near (bravo - (List.item index pages).Top) (firstLineY doc layouts "Fill line 1" - pages.Head.Top)
    index, atTop
// (How many one-line paragraphs a page holds.)
let capacity =
    let doc = import (String.concat "" [ for i in 1 .. 120 -> para (sprintf "Fill line %d" i) ] + sectionEnd "" "" 1 + para "Bravo line") (sectPr "nextPage" 1)
    let _, layouts, pages = layoutOf doc
    let second = (List.item 1 pages).Top
    doc.Blocks |> List.choose Doc.tryGetParagraph
    |> List.filter (fun p -> (toString p.Buffer).StartsWith "Fill")
    |> List.filter (fun p -> (layouts |> List.find (fun (l: Layout.BlockLayout) -> l.BlockId = p.Id)).Y < second)
    |> List.length
check "next page, full page: the empty paragraph takes no new page" (pageCount capacity "nextPage") (1, true)
check "continuous, full page: the next section at the top of the next page" (pageCount capacity "continuous") (1, true)

printfn "word section end: %d passed, %d failed" pass fail
if fail > 0 then exit 1
