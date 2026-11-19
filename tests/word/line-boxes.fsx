// Line heights and what of a line has to fit on a page, as Word 2019 lays them out
// (checked on the Windows VM with tools/word-parity/page-parity.mjs --dir and Word's
// PDFs of small documents; the numbers in the comments are Word's):
// - a line at single spacing is its font's single line - Calibri 11pt 13.43pt, Arial
//   and Times New Roman 11pt 12.65pt - whether single spacing is stated (w:line="240")
//   or not, and in a table cell as in the body; at a multiple, that times the multiple
//   (Calibri 11pt at 1.15: 15.44pt);
// - a document without w:pPrDefault has Word's own paragraph defaults: 8pt after and
//   lines of 1.158 of single (Calibri 11pt 15.54pt, 22pt 31.10, Arial 11pt 14.70,
//   Times New Roman 14.64); an empty w:pPrDefault is single spacing and nothing after;
// - the paragraph mark's properties (w:pPr/w:rPr) are the mark's: a run without
//   properties of its own is in the style's font (Times New Roman 12pt under a
//   Calibri 11pt mark);
// - a paragraph of a space is one line;
// - a line stays on a page when its text at single spacing ends above the bottom
//   margin: what a multiple adds may run below it (1.15: 42 lines of 15.44pt on a
//   648pt body; 1.62: 30 lines, 5pt over) - but a single-spaced line 1.6pt over goes
//   on; a paragraph's spacing after need not fit (its text 4pt above the margin, its
//   spacing after 8pt below it: it stays).
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
let near label (got: float) (expected: float) =
    if abs (got - expected) < 0.05 then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-60s expected %.3f got %.3f" label expected got

let w = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
let sectPr = "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr>"
/// Document defaults: a font and size, and the w:pPrDefault (None: none at all).
let styles (font: string) (halfPoints: int) (pPrDefault: string option) =
    let pPr = match pPrDefault with Some p -> sprintf "<w:pPrDefault>%s</w:pPrDefault>" p | None -> ""
    sprintf "<w:styles %s><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"%s\" w:hAnsi=\"%s\"/><w:sz w:val=\"%d\"/></w:rPr></w:rPrDefault>%s</w:docDefaults><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style><w:style w:type=\"table\" w:default=\"1\" w:styleId=\"TableNormal\"><w:name w:val=\"Normal Table\"/><w:tblPr><w:tblCellMar><w:top w:w=\"0\" w:type=\"dxa\"/><w:left w:w=\"108\" w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr></w:style></w:styles>" w font font halfPoints pPr
let single = Some "<w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr>"
let package (stylesXml: string) (body: string) =
    use stream = new IO.MemoryStream()
    do
        use archive = new IO.Compression.ZipArchive(stream, IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new IO.StreamWriter(archive.CreateEntry(name).Open(), Text.UTF8Encoding(false))
            writer.Write text
        add "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"""
        add "word/document.xml" (sprintf "<w:document %s><w:body>%s%s</w:body></w:document>" w body sectPr)
        add "word/styles.xml" stylesXml
    stream.ToArray()
let import stylesXml body = package stylesXml body |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let layoutOf (doc: Doc.Document) =
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    Layout.layoutDocumentPages settings measureText measureRange doc
let pt (points: float) = points * 96.0 / 72.0
let toPt (px: float) = px * 72.0 / 96.0

let para (pPr: string) (text: string) = sprintf "<w:p><w:pPr>%s</w:pPr><w:r><w:t xml:space=\"preserve\">%s</w:t></w:r></w:p>" pPr text
let textOf (layout: Layout.BlockLayout) =
    layout.Lines |> List.collect (fun l -> l.Segments |> List.choose (function Layout.TextSegment (t, _, _, _, _) -> Some t | _ -> None)) |> String.concat ""
let blockNamed (layouts: Layout.BlockLayout list) (text: string) = layouts |> List.find (fun l -> textOf l = text)
let pageOf (pages: Layout.PageInfo list) (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
/// The pitch of one-line paragraphs: how far below Alpha's line Bravo's stands.
let pitch stylesXml (pPr: string) =
    let doc = import stylesXml (para pPr "Alpha" + para pPr "Bravo")
    let layouts, _ = layoutOf doc
    toPt ((blockNamed layouts "Bravo").Lines.Head.Y - (blockNamed layouts "Alpha").Lines.Head.Y)

// Single spacing, stated: the font's single line.
near "Calibri 11pt single (w:line=240): 13.43pt" (pitch (styles "Calibri" 22 single) "") 13.427
near "Arial 11pt single: 12.65pt" (pitch (styles "Arial" 22 single) "") 12.649
near "Times New Roman 12pt single: 13.80pt" (pitch (styles "Times New Roman" 24 single) "") 13.799
near "Calibri 11pt at 1.15 (w:line=276): 15.44pt" (pitch (styles "Calibri" 22 single) "<w:spacing w:line=\"276\" w:lineRule=\"auto\"/>") 15.441
// An empty w:pPrDefault: single spacing, nothing after.
near "an empty w:pPrDefault: single, nothing after" (pitch (styles "Calibri" 22 (Some "")) "") 13.427
// No w:pPrDefault: 8pt after and 1.158 of single.
near "no w:pPrDefault, Calibri 11pt: 15.55pt and 8pt after" (pitch (styles "Calibri" 22 None) "") (13.427 * 1.158 + 8.0)
near "no w:pPrDefault, Calibri 22pt: 31.10pt and 8pt after" (pitch (styles "Calibri" 44 None) "") (26.854 * 1.158 + 8.0)
near "no w:pPrDefault, Arial 11pt" (pitch (styles "Arial" 22 None) "") (12.649 * 1.158 + 8.0)
near "no w:pPrDefault, a line spacing of the paragraph's own (1.5): 20.14pt and 8pt" (pitch (styles "Calibri" 22 None) "<w:spacing w:line=\"360\" w:lineRule=\"auto\"/>") (13.427 * 1.5 + 8.0)

// A table cell's line is the body's.
let cellLine =
    let row = "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p></w:tc></w:tr>"
    let doc = import (styles "Calibri" 22 single) (sprintf "<w:tbl><w:tblPr><w:tblW w:w=\"3000\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"3000\"/></w:tblGrid>%s</w:tbl>" row)
    let layouts, _ = layoutOf doc
    toPt ((blockNamed layouts "Two").Lines.Head.Y - (blockNamed layouts "One").Lines.Head.Y)
near "a table cell's single line: 13.43pt, as in the body" cellLine 13.427

// A table style based on another has that one's paragraph properties too (docx4j-headers-footers:
// ElementTable based on Table Grid, whose single spacing and no space after Word's PDF shows in its
// cells - 13.44pt lines - under document defaults of 1.15 and 10pt after).
let basedOnLine =
    let tableStyles =
        "<w:style w:type=\"table\" w:styleId=\"Grid\"><w:name w:val=\"Grid\"/><w:pPr><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/></w:pPr></w:style>"
        + "<w:style w:type=\"table\" w:styleId=\"Element\"><w:name w:val=\"Element\"/><w:basedOn w:val=\"Grid\"/><w:tblPr><w:tblCellMar><w:left w:w=\"115\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr></w:style></w:styles>"
    let stylesXml = (styles "Calibri" 22 (Some "<w:pPr><w:spacing w:after=\"200\" w:line=\"276\" w:lineRule=\"auto\"/></w:pPr>")).Replace("</w:styles>", tableStyles)
    let row = "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p></w:tc></w:tr>"
    let doc = import stylesXml (sprintf "<w:tbl><w:tblPr><w:tblStyle w:val=\"Element\"/><w:tblW w:w=\"3000\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"3000\"/></w:tblGrid>%s</w:tbl>" row)
    let layouts, _ = layoutOf doc
    toPt ((blockNamed layouts "Two").Lines.Head.Y - (blockNamed layouts "One").Lines.Head.Y)
near "a table style based on Table Grid: single, nothing after (13.43pt)" basedOnLine 13.427

// The paragraph mark's font is not its runs'.
let markDoc = import (styles "Times New Roman" 24 single) (para "<w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/><w:sz w:val=\"22\"/></w:rPr>" "Article one")
let markRun =
    let layouts, _ = layoutOf markDoc
    (blockNamed layouts "Article one").Lines.Head.Segments |> List.pick (function Layout.TextSegment (_, style, _, _, _) -> Some (style.FontFamily, style.FontSize) | _ -> None)
check "a run under a Calibri 11pt mark is in the style's Times New Roman 12pt" markRun (Some "Times New Roman", pt 12.0)
near "and its line is Times New Roman 12pt's" (pitch (styles "Times New Roman" 24 single) "<w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/><w:sz w:val=\"22\"/></w:rPr>") 13.799

// A paragraph of a space is one line.
let spaceGap =
    let spaced = "<w:spacing w:before=\"280\" w:after=\"280\"/>"
    let doc = import (styles "Times New Roman" 24 (Some "")) (para spaced "Bravo" + para spaced " " + para spaced "Delta")
    let layouts, _ = layoutOf doc
    toPt ((blockNamed layouts "Delta").Lines.Head.Y - (blockNamed layouts "Bravo").Lines.Head.Y)
near "Bravo, a paragraph of a space, Delta (14pt between): two lines and two gaps" spaceGap (2.0 * (13.799 + 14.0))

// What of a line has to fit: one-line paragraphs on a 648pt body (Letter, 1in margins).
let fill (count: int) (pPr: string) = String.concat "" [ for i in 1 .. count -> para pPr (sprintf "L%d" i) ]
let pageOfText stylesXml body (text: string) =
    let doc = import stylesXml body
    let layouts, pages = layoutOf doc
    pageOf pages (blockNamed layouts text).Lines.Head.Y
let at115 = "<w:spacing w:after=\"0\" w:line=\"276\" w:lineRule=\"auto\"/>"
check "1.15: the 42nd line (0.6pt over, its text 1.6pt above) stays" (pageOfText (styles "Calibri" 22 single) (fill 42 at115 + para at115 "After") "L42") 1
check "1.15: the 43rd goes on" (pageOfText (styles "Calibri" 22 single) (fill 42 at115 + para at115 "After") "After") 2
let at162 = "<w:spacing w:after=\"0\" w:line=\"389\" w:lineRule=\"auto\"/>"
check "1.62: the 30th line, 5pt over, stays" (pageOfText (styles "Calibri" 22 single) (fill 30 at162 + para at162 "After") "L30") 1
let singleBody = fill 1 "<w:spacing w:before=\"100\"/>" + String.concat "" [ for i in 1 .. 48 -> para "" (sprintf "M%d" i) ]
check "single, 5pt down: the 47th line fits" (pageOfText (styles "Calibri" 22 single) singleBody "M46") 1
check "single, 5pt down: the 48th (1.6pt over) goes on" (pageOfText (styles "Calibri" 22 single) singleBody "M47") 2
let after12 = "<w:spacing w:after=\"240\"/>"
let afterBody = para "<w:spacing w:before=\"400\" w:after=\"240\"/>" "L1" + String.concat "" [ for i in 1 .. 25 -> para after12 (sprintf "M%d" i) ]
check "spacing after need not fit: the 25th paragraph stays" (pageOfText (styles "Calibri" 22 single) afterBody "M24") 1
check "the 26th goes on" (pageOfText (styles "Calibri" 22 single) afterBody "M25") 2

// The body starts below the header where that reaches below the top margin, and ends above the
// footer where that reaches above the bottom margin - as far as their text reaches (docx4j-sample-
// docx: a header of two single-spaced Calibri 11pt paragraphs 36pt from the top, a top margin of
// 36pt; Word's first line of the body at 62.9pt).
let withHeaderFooter (headerParas: int) (footerParas: int) (body: string) =
    let story tag count = sprintf "<w:%s %s>%s</w:%s>" tag w (String.concat "" [ for i in 1 .. count -> para "" (sprintf "%s line %d" tag i) ]) tag
    let section = "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rIdH\"/><w:footerReference w:type=\"default\" r:id=\"rIdF\"/><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"720\" w:right=\"1440\" w:bottom=\"720\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr>"
    use stream = new IO.MemoryStream()
    do
        use archive = new IO.Compression.ZipArchive(stream, IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new IO.StreamWriter(archive.CreateEntry(name).Open(), Text.UTF8Encoding(false))
            writer.Write text
        add "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"""
        add "word/document.xml" (sprintf "<w:document %s xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>%s%s</w:body></w:document>" w body section)
        add "word/_rels/document.xml.rels" """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>"""
        add "word/styles.xml" (styles "Calibri" 22 single)
        add "word/header1.xml" (story "hdr" headerParas)
        add "word/footer1.xml" (story "ftr" footerParas)
    stream.ToArray() |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
let bodyTop headerParas =
    let doc = withHeaderFooter headerParas 1 (para "" "Body")
    let layouts, pages = layoutOf doc
    toPt ((blockNamed layouts "Body").Lines.Head.Y - pages.Head.Top)
near "a header of two lines 36pt down: the body at 36 + 2 x 13.43pt" (bodyTop 2) (36.0 + 2.0 * 13.427)
near "a header of five lines: the body at 36 + 5 x 13.43pt" (bodyTop 5) (36.0 + 5.0 * 13.427)
near "a header of one line: 36 + 13.43pt" (bodyTop 1) (36.0 + 13.427)
// A page holds what fits between them: 720pt - 36pt - (36 + 4 x 13.43pt) of single lines below 36 + 2 lines.
let linesOnFirstPage footerParas =
    let doc = withHeaderFooter 2 footerParas (String.concat "" [ for i in 1 .. 60 -> para "" (sprintf "B%d" i) ])
    let layouts, pages = layoutOf doc
    layouts |> List.filter (fun l -> pageOf pages l.Y = 1 && (textOf l).StartsWith "B") |> List.length
check "a footer of four lines 36pt up: the body ends above it" (linesOnFirstPage 4) (int (floor ((792.0 - 36.0 - 4.0 * 13.427 - (36.0 + 2.0 * 13.427)) / 13.427)))

printfn "line-boxes: %d passed, %d failed" pass fail
if fail > 0 then exit 1
