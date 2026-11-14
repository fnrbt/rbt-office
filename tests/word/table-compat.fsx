// Where a table starts, by the document's compatibility mode (settings.xml
// w:compat/w:compatSetting w:name="compatibilityMode"). In a document of mode 14
// (Word 2010) or older - and one with no such setting, which Word opens in
// compatibility mode - Word puts the text of the first cell where the
// paragraphs' text is: the table's edge is that cell's left margin left of the
// margin (its w:tcMar, else the table's w:tblCellMar, else 108 twips). In mode
// 15 (Word 2013 on) the table's edge is at the margin. Checked against Word 2019
// on a table with tblCellMar left 200 and a first cell with tcMar left 400:
// mode 14 and no setting move it 400 twips left, mode 15 not at all.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open Rbt.Word.TextBuffer
module Doc = Rbt.Word.Document
module Layout = Rbt.Word.Layout
module Paint = Rbt.Word.Paint

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-60s expected %A got %A" label expected got

let w = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
let package (settings: string option) (body: string) =
    use stream = new IO.MemoryStream()
    do
        use archive = new IO.Compression.ZipArchive(stream, IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new IO.StreamWriter(archive.CreateEntry(name).Open(), Text.UTF8Encoding(false))
            writer.Write text
        add "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"""
        add "word/document.xml" (sprintf "<w:document %s><w:body>%s<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>" w body)
        // (The default table style as Word writes it: cell margins of 108 twips.)
        add "word/styles.xml" (sprintf "<w:styles %s><w:style w:type=\"table\" w:default=\"1\" w:styleId=\"TableNormal\"><w:name w:val=\"Normal Table\"/><w:tblPr><w:tblInd w:w=\"0\" w:type=\"dxa\"/><w:tblCellMar><w:top w:w=\"0\" w:type=\"dxa\"/><w:left w:w=\"108\" w:type=\"dxa\"/><w:bottom w:w=\"0\" w:type=\"dxa\"/><w:right w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar></w:tblPr></w:style></w:styles>" w)
        match settings with
        | Some text -> add "word/settings.xml" (sprintf "<w:settings %s>%s</w:settings>" w text)
        | None -> ()
    stream.ToArray()
let compat mode =
    sprintf "<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"%d\"/></w:compat>" mode
let import settings body = package settings body |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let settingsOf (doc: Doc.Document) : Layout.LayoutSettings =
    Layout.documentSettings 16.0 24.0 3.5 doc

let para (text: string) = sprintf "<w:p><w:r><w:t xml:space=\"preserve\">%s</w:t></w:r></w:p>" text
// (Every cell shaded: the fills show where the table is.)
let cell (tcPr: string) (text: string) = sprintf "<w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"DDDDDD\"/>%s</w:tcPr>%s</w:tc>" tcPr (para text)
let table (tblCellMar: string) (firstTcPr: string) =
    sprintf "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/>%s<w:tblBorders><w:top w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/><w:left w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/><w:bottom w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/><w:right w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"3000\"/></w:tblGrid><w:tr>%s%s</w:tr><w:tr>%s%s</w:tr></w:tbl>"
        tblCellMar (cell firstTcPr "A1") (cell "" "B1") (cell "<w:tcMar><w:left w:w=\"50\" w:type=\"dxa\"/></w:tcMar>" "A2") (cell "" "B2")
let tableA = para "Before." + table "<w:tblCellMar><w:left w:w=\"200\" w:type=\"dxa\"/><w:right w:w=\"200\" w:type=\"dxa\"/></w:tblCellMar>" "<w:tcMar><w:left w:w=\"400\" w:type=\"dxa\"/></w:tcMar>" + para "After."

let cellText (doc: Doc.Document) (text: string) =
    let layouts = Layout.layoutDocumentPages (settingsOf doc) measureText measureRange doc |> fst
    let para =
        doc.Blocks |> List.pick (function
            | Doc.Table t -> t.Rows |> List.collect (fun r -> r.Cells |> List.collect (fun c -> c.Content)) |> List.tryFind (fun p -> toString p.Buffer = text)
            | _ -> None)
    let layout = layouts |> List.find (fun l -> l.BlockId = para.Id)
    match layout.Lines.Head.Segments with
    | Layout.TextSegment (_, _, x, _, _) :: _ -> Math.Round(x, 2)
    | _ -> nan
/// The x of the table's left edge (where its first cell's fill starts), from the margin.
let leftEdge (doc: Doc.Document) =
    let settings = settingsOf doc
    let tableBlock = doc.Blocks |> List.pick (function Doc.Table t -> Some t | _ -> None)
    let commands, _ = Paint.tableBorderCommandsForDocumentWithRange settings measureText measureRange doc tableBlock settings.PaddingTop
    commands
    |> List.choose (function
        | Paint.FillRect (_, x, _, _, _) -> Some x
        | _ -> None)
    |> List.min
    |> fun x -> Math.Round(x - settings.PaddingLeft, 1)

// (400 twips = 26.67 px, 50 twips = 3.33 px.)
let mode14 = import (Some (compat 14)) tableA
let mode15 = import (Some (compat 15)) tableA
let noSetting = import (Some "") tableA
let noPart = import None tableA
check "read: mode 14" mode14.PageSetup.CompatibilityMode 14
check "read: mode 15" mode15.PageSetup.CompatibilityMode 15
check "read: no compatSetting is an older document" noSetting.PageSetup.CompatibilityMode 12
check "read: no settings part either" noPart.PageSetup.CompatibilityMode 12
check "a new document is saved without a settings part: Word takes it for an old one, and so do we" Doc.defaultPageSetup.CompatibilityMode 12

check "mode 14: the first cell's text is at the margin" (cellText mode14 "A1") 0.0
check "mode 14: the table's edge is the first cell's margin (400 twips) left of it" (Math.Round(leftEdge mode14, 0)) -27.0
check "mode 14: the second row's cell keeps its own margin (50 twips)" (cellText mode14 "A2") -23.33
check "no compatSetting: as mode 14" (cellText noSetting "A1", cellText noPart "A1") (0.0, 0.0)
check "mode 15: the table's edge is at the margin" (Math.Round(leftEdge mode15, 0)) 0.0
check "mode 15: the first cell's text is its margin in" (cellText mode15 "A1") 26.67

// A first cell without a margin of its own: the table's (w:tblCellMar, 300 twips = 20 px).
let tableMargin = para "Before." + table "<w:tblCellMar><w:left w:w=\"300\" w:type=\"dxa\"/></w:tblCellMar>" "" + para "After."
check "mode 14, the table's margin: the text at the margin" (cellText (import (Some (compat 14)) tableMargin) "A1") 0.0
check "mode 15, the table's margin: the text 300 twips in" (cellText (import (Some (compat 15)) tableMargin) "A1") 20.0

// A table without margins of its own and no style: the default table style's (108 twips = 7.2 px).
let plain = para "Before." + table "" "" + para "After."
check "mode 14, the default style's margins: the text at the margin" (cellText (import (Some (compat 14)) plain) "A1") 0.0
check "mode 15, the default style's margins: the text 108 twips in" (cellText (import (Some (compat 15)) plain) "A1") 7.2

// A centred table is placed by its edges in mode 14 too (annual-review.docx: the full-width
// banner table with w:jc center and cell margins of 420 twips stands at the margins in Word).
let centred = tableA.Replace("<w:tblW w:w=\"0\" w:type=\"auto\"/>", "<w:tblW w:w=\"0\" w:type=\"auto\"/><w:jc w:val=\"center\"/>")
check "a centred table: where mode 15 has it" (leftEdge (import (Some (compat 14)) centred)) (leftEdge (import (Some (compat 15)) centred))
check "a centred table: in the middle" (leftEdge (import (Some (compat 14)) centred)) (Math.Round((624.0 - 400.0) / 2.0, 1))

// Saved and opened again, a document keeps its mode (the settings part is kept).
let reopen (doc: Doc.Document) =
    let package = Rbt.Word.Docx.exportTypedDocumentPackage doc
    DocumentWriter.writeTypedPackageBytes { package with Unchanged = false } |> Rbt.Word.Docx.loadBytesNative |> Option.get
check "saved: mode 14 stays 14" (reopen mode14).PageSetup.CompatibilityMode 14
check "saved: mode 15 stays 15" (reopen mode15).PageSetup.CompatibilityMode 15
check "saved: a new document stays 12" (reopen { Doc.empty with Blocks = mode15.Blocks }).PageSetup.CompatibilityMode 12
printfn "word table compat: %d passed, %d failed" pass fail
if fail > 0 then exit 1
