#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open Rbt.Word.Document

let mutable pass = 0
let mutable fail = 0

let check label got expected =
    if got = expected then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-34s expected %A got %A" label expected got

let ok label cond detail =
    if cond then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-34s %s" label detail

let textOfPara (para: ParagraphBlock) =
    Rbt.Word.TextBuffer.toString para.Buffer

let firstParagraphText (doc: Document) =
    doc.Blocks
    |> List.tryPick (function
        | Paragraph para
        | Heading (_, para) -> Some (textOfPara para)
        | Table table ->
            table.Rows
            |> List.tryPick (fun row ->
                row.Cells
                |> List.tryPick (fun cell ->
                    cell.Content |> List.tryHead |> Option.map textOfPara))
        | TextBox textBox ->
            textBox.Content |> List.tryHead |> Option.map textOfPara)
    |> Option.defaultValue ""

let blockKind = function
    | Paragraph _ -> "Paragraph"
    | Heading _ -> "Heading"
    | Table _ -> "Table"
    | TextBox _ -> "TextBox"

let hasSectionPageSetup (doc: Document) =
    let paraHasSetup (para: ParagraphBlock) =
        para.Metadata.SectionPageSetup.IsSome

    doc.Blocks
    |> List.exists (function
        | Paragraph para
        | Heading (_, para) -> paraHasSetup para
        | Table table ->
            table.Rows
            |> List.exists (fun row ->
                row.Cells
                |> List.exists (fun cell -> cell.Content |> List.exists paraHasSetup))
        | TextBox textBox -> textBox.Content |> List.exists paraHasSetup)

type Expected = {
    Path: string
    Blocks: int
    FirstText: string
    Kinds: string list
    HasSectionSetup: bool
}

let fixtures =
    [
        {
            Path = "tests/repros/word/two-column-section.docx"
            Blocks = 11
            FirstText = "Two column section repro"
            Kinds = List.replicate 11 "Paragraph"
            HasSectionSetup = false
        }
        {
            Path = "tests/repros/word/landscape-section.docx"
            Blocks = 9
            FirstText = "Landscape section (page 1)"
            Kinds = List.replicate 9 "Paragraph"
            HasSectionSetup = true
        }
        {
            Path = "tests/repros/word/table-import-preserves-structure.docx"
            Blocks = 3
            FirstText = "Table import repro"
            Kinds = [ "Paragraph"; "Table"; "Paragraph" ]
            HasSectionSetup = false
        }
    ]

let importFixture path importer =
    let root = IO.Path.GetFullPath(IO.Path.Combine(__SOURCE_DIRECTORY__, "../.."))
    let bytes = IO.File.ReadAllBytes(IO.Path.Combine(root, path))
    match DocxReader.loadBytesNative bytes with
    | Some parsed -> importer parsed
    | None -> failwithf "Could not parse %s" path

let assertFixture (expected: Expected) (doc: Document) =
    ok (expected.Path + " non-empty") (not (List.isEmpty doc.Blocks)) "document has no blocks"
    check (expected.Path + " block count") (List.length doc.Blocks) expected.Blocks
    check (expected.Path + " first text") (firstParagraphText doc) expected.FirstText
    check (expected.Path + " page width") doc.PageSetup.PageWidth 816.0
    check (expected.Path + " page height") doc.PageSetup.PageHeight 1056.0
    check (expected.Path + " margins") (doc.PageSetup.MarginLeft, doc.PageSetup.MarginRight, doc.PageSetup.MarginTop, doc.PageSetup.MarginBottom) (96.0, 96.0, 96.0, 96.0)
    check (expected.Path + " block kinds") (doc.Blocks |> List.map blockKind) expected.Kinds
    check (expected.Path + " section setup") (hasSectionPageSetup doc) expected.HasSectionSetup

for expected in fixtures do
    let doc = importFixture expected.Path Rbt.Word.Import.importDoc
    assertFixture expected doc

let summary (doc: Document) =
    List.length doc.Blocks,
    firstParagraphText doc,
    doc.PageSetup,
    (doc.Blocks |> List.map blockKind),
    hasSectionPageSetup doc

let libraryDoc = importFixture fixtures.[1].Path Rbt.Word.Import.importDoc
let compatDoc = importFixture fixtures.[1].Path NativeRunner.importDoc
check "compat alias summary" (summary compatDoc) (summary libraryDoc)

// ---- units and coverage shared with the browser importer --------------------
// The native importer must fill the model the way DocxReader -> Docx.fs does:
// indents in px (twips / 15), exact line spacing as line px / font px with the
// font size in ParagraphMarkFontSize, and the run / paragraph properties the
// typed EditorRun cannot carry.
let wNsAll = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\""

let syntheticDocx (body: string) =
    use stream = new IO.MemoryStream()
    do
        use archive = new IO.Compression.ZipArchive(stream, IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new IO.StreamWriter(archive.CreateEntry(name).Open(), Text.UTF8Encoding(false))
            writer.Write text
        add "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"""
        add "word/document.xml" (sprintf """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>%s</w:body></w:document>""" body)
    stream.ToArray()

let unitsDoc =
    syntheticDocx
        ("""<w:p><w:pPr><w:ind w:left="720" w:right="360" w:hanging="360"/></w:pPr><w:r><w:t>indented</w:t></w:r></w:p>"""
         + """<w:p><w:pPr><w:spacing w:line="280" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>exact</w:t></w:r></w:p>"""
         + """<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs><w:bidi/></w:pPr><w:r><w:rPr><w:vanish/><w:caps/><w:dstrike/><w:u w:val="wave" w:color="FF0000"/><w:position w:val="6"/></w:rPr><w:t>props</w:t></w:r><w:ins w:id="3" w:author="Ann"><w:r><w:t>added</w:t></w:r></w:ins></w:p>""")
    |> DocxReader.loadBytesNative
    |> Option.get
    |> Rbt.Word.Import.importDoc
let unitParas = unitsDoc.Blocks |> List.choose (function Paragraph para -> Some para | _ -> None)
let runStyles (para: ParagraphBlock) =
    para.Buffer.Runs |> List.choose (function Rbt.Word.TextBuffer.Text run -> Some run.Style | _ -> None)
match unitParas with
| [ indented; exact; props ] ->
    check "native indent px" (indented.Metadata.IndentLeft, indented.Metadata.IndentRight, indented.Metadata.FirstLineIndent) (Some 48.0, Some 24.0, Some -24.0)
    // 280 twips = 18.667px over a 14pt (18.667px) font.
    check "native exact rule" exact.Metadata.LineSpacingRule (Some "exact")
    ok "native exact spacing"
       (exact.Metadata.LineSpacing |> Option.exists (fun value -> abs (value - 1.0) < 1e-6)
        && exact.Metadata.ParagraphMarkFontSize |> Option.exists (fun value -> abs (value - 14.0 * 96.0 / 72.0) < 1e-6))
       (sprintf "%A %A" exact.Metadata.LineSpacing exact.Metadata.ParagraphMarkFontSize)
    check "native tab stops" (props.Metadata.TabStops |> List.map (fun stop -> stop.Position, stop.Alignment, stop.Leader)) [ 600.0, "right", Some "dot" ]
    check "native rtl" props.Metadata.Rtl true
    match runStyles props with
    | [ first; second ] ->
        check "native run props"
            (first.Hidden, first.AllCaps, first.DoubleStrike, first.UnderlineStyle, first.UnderlineColor, first.Position)
            (true, true, true, Some "wave", Some "FF0000", 3.0)
        check "native revision" (second.RevisionType, second.RevisionAuthor, second.Hidden) (Some "insert", Some "Ann", false)
    | other -> ok "native run props" false (sprintf "%d runs" other.Length)
| other -> ok "native units paragraphs" false (sprintf "%d paragraphs" other.Length)

// ---- explicit vs inherited formatting ---------------------------------------
// w:b w:val="0", an explicit size and an explicit w:jc="left" must survive under
// a style that is bold, large and centred.
let syntheticDocxWithStyles (styles: string) (body: string) =
    use stream = new IO.MemoryStream()
    do
        use archive = new IO.Compression.ZipArchive(stream, IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new IO.StreamWriter(archive.CreateEntry(name).Open(), Text.UTF8Encoding(false))
            writer.Write text
        let w = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
        add "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"""
        add "word/styles.xml" (sprintf "<w:styles %s>%s</w:styles>" w styles)
        add "word/document.xml" (sprintf "<w:document %s><w:body>%s</w:body></w:document>" w body)
    stream.ToArray()
let loudStyleXml = """<w:style w:type="paragraph" w:styleId="Loud"><w:name w:val="Loud"/><w:pPr><w:keepNext/><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style>"""
let explicitDoc =
    syntheticDocxWithStyles loudStyleXml
        ("""<w:p><w:pPr><w:pStyle w:val="Loud"/></w:pPr><w:r><w:t>inherits</w:t></w:r><w:r><w:rPr><w:b w:val="0"/><w:sz w:val="22"/></w:rPr><w:t>overrides</w:t></w:r></w:p>"""
         + """<w:p><w:pPr><w:pStyle w:val="Loud"/><w:jc w:val="left"/></w:pPr><w:r><w:t>left</w:t></w:r></w:p>""")
    |> DocxReader.loadBytesNative
    |> Option.get
    |> Rbt.Word.Import.importDoc
let explicitParas = explicitDoc.Blocks |> List.choose (function Paragraph para | Heading (_, para) -> Some para | _ -> None)
match explicitParas with
| [ styled; left ] ->
    let flags = Rbt.Word.TextBuffer.ExplicitFormat.Bold ||| Rbt.Word.TextBuffer.ExplicitFormat.FontSize
    check "native explicit flags" (runStyles styled |> List.map (fun style -> style.Bold, style.Explicit)) [ true, 0; false, flags ]
    let effective (para: ParagraphBlock) =
        let style = Rbt.Word.WordStyles.forBlockInDocument explicitDoc (Paragraph para)
        (Rbt.Word.WordStyles.effectiveBufferInDocument explicitDoc style para.Buffer).Runs
        |> List.choose (function Rbt.Word.TextBuffer.Text run -> Some (run.Style.Bold, Math.Round(run.Style.FontSize, 2)) | _ -> None)
    check "native explicit off overrides bold style" (effective styled) [ true, 26.67; false, 14.67 ]
    let alignmentOf (para: ParagraphBlock) =
        (Rbt.Word.WordStyles.effectiveMetadata (Rbt.Word.WordStyles.forBlockInDocument explicitDoc (Paragraph para)) para.Metadata).Alignment
    check "native alignment flags" (styled.Metadata.ExplicitAlignment, left.Metadata.ExplicitAlignment) (false, true)
    check "native explicit left overrides centred style" (alignmentOf styled, alignmentOf left) (Center, Left)
    // keepNext comes from the Loud style here, not from the paragraphs.
    let styleOf (para: ParagraphBlock) = Rbt.Word.WordStyles.forBlockInDocument explicitDoc (Paragraph para)
    check "native style keep-next is not direct" (styled.Metadata.KeepNext, Rbt.Word.WordStyles.keepNext (styleOf styled) styled.Metadata) (None, true)
| other -> ok "native explicit paragraphs" false (sprintf "%d paragraphs" other.Length)

let directKeepDoc =
    syntheticDocx """<w:p><w:pPr><w:keepNext/><w:keepLines/></w:pPr><w:r><w:t>kept</w:t></w:r></w:p><w:p><w:pPr><w:keepNext w:val="0"/></w:pPr><w:r><w:t>not kept</w:t></w:r></w:p><w:p><w:r><w:t>plain</w:t></w:r></w:p>"""
    |> DocxReader.loadBytesNative
    |> Option.get
    |> Rbt.Word.Import.importDoc
check "native direct keepNext / keepLines"
    (directKeepDoc.Blocks |> List.choose (function Paragraph para | Heading (_, para) -> Some (para.Metadata.KeepNext, para.Metadata.KeepLines) | _ -> None))
    [ Some true, Some true; Some false, None; None, None ]

// ----- footnote / endnote / comment bodies ---------------------------------
/// A package made of the given parts (word/document.xml gets the body wrapper).
let syntheticPackage (body: string) (parts: (string * string) list) =
    use stream = new IO.MemoryStream()
    do
        use archive = new IO.Compression.ZipArchive(stream, IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new IO.StreamWriter(archive.CreateEntry(name).Open(), Text.UTF8Encoding(false))
            writer.Write text
        add "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"""
        add "word/document.xml" (sprintf "<w:document %s><w:body>%s</w:body></w:document>" wNsAll body)
        for name, text in parts do add name text
    stream.ToArray()
let importPackage body parts =
    syntheticPackage body parts |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc

let notesNativeDoc =
    importPackage
        """<w:p><w:r><w:t>body</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:commentReference w:id="5"/></w:r></w:p>"""
        [ "word/footnotes.xml",
          sprintf """<w:footnotes %s><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:hyperlink r:id="rId9"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:footnote></w:footnotes>""" wNsAll
          "word/_rels/footnotes.xml.rels",
          """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/n" TargetMode="External"/></Relationships>"""
          "word/comments.xml",
          sprintf """<w:comments %s><w:comment w:id="5" w:author="Ann" w:initials="A"><w:p><w:r><w:rPr><w:i/></w:rPr><w:t>remark</w:t></w:r></w:p></w:comment></w:comments>""" wNsAll ]
let nativeNote = notesNativeDoc.Package.Notes |> List.tryFind (fun note -> note.Id = "2")
check "native note bodies" (notesNativeDoc.Package.Notes |> List.map (fun note -> note.NoteType, note.Id, note.Kind)) [ "footnote", "-1", Some "separator"; "footnote", "2", None ]
check "native note body keeps the table" (nativeNote |> Option.map (fun note -> note.Blocks |> List.map blockKind)) (Some [ "Paragraph"; "Table" ])
let nativeNoteRuns =
    nativeNote
    |> Option.map (fun note -> paragraphsOfBlocks note.Blocks |> List.collect (fun para -> para.Buffer.Runs |> List.choose (function Rbt.Word.TextBuffer.Text run -> Some (Rbt.Word.TextBuffer.slice run.Start run.Length para.Buffer, run.Style.Bold, run.Style.Hyperlink) | _ -> None)))
    |> Option.defaultValue []
check "native note body formatting and hyperlink" nativeNoteRuns [ "bold", true, None; "link", false, Some "https://example.com/n"; "cell", false, None ]
let nativeNoteReference =
    paragraphsOfBlocks notesNativeDoc.Blocks
    |> List.collect (fun para -> para.Buffer.Runs)
    |> List.tryPick (function Rbt.Word.TextBuffer.Text run when run.Style.NoteId = Some "2" -> Some run.Style.NoteText | _ -> None)
check "native note text matches the reference run" (nativeNote |> Option.map (fun note -> Some note.Text)) nativeNoteReference
check "native comment body"
    (notesNativeDoc.Package.Comments |> List.map (fun comment -> comment.Id, comment.Author, comment.Initials, comment.Text, comment.Blocks |> List.map blockKind))
    [ "5", Some "Ann", Some "A", "remark", [ "Paragraph" ] ]

// ----- header / footer tables stay tables ----------------------------------
let headerTableDoc =
    importPackage
        """<w:p><w:r><w:t>body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:footerReference w:type="default" r:id="rId8"/></w:sectPr>"""
        [ "word/_rels/document.xml.rels",
          """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>"""
          "word/header1.xml",
          sprintf """<w:hdr %s><w:p><w:r><w:t>above</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>logo</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>address</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>below</w:t></w:r></w:p></w:hdr>""" wNsAll
          "word/footer1.xml",
          sprintf """<w:ftr %s><w:tbl><w:tr><w:tc><w:p><w:r><w:t>page</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:ftr>""" wNsAll ]
check "native header keeps its table" (headerTableDoc.HeaderParts |> List.map (fun part -> part.Kind, part.Blocks |> List.map blockKind)) [ "default", [ "Paragraph"; "Table"; "Paragraph" ] ]
check "native footer keeps its table" (headerTableDoc.FooterParts |> List.map (fun part -> part.Blocks |> List.map blockKind)) [ [ "Table"; "Paragraph" ] ]
check "native header table cells"
    (headerTableDoc.HeaderParts
     |> List.collect (fun part -> part.Blocks)
     |> List.collect (function Table table -> table.Rows |> List.collect (fun row -> row.Cells |> List.map (fun cell -> cell.Content |> List.map textOfPara)) | _ -> []))
    [ [ "logo" ]; [ "address" ] ]
// Document.Headers stays the paragraph-only view of the first part.
check "native header paragraph view" (headerTableDoc.Headers |> List.map textOfPara) [ "above"; "below" ]

// ----- three sections: every section keeps its own references ---------------
let sectionPara refs extra = sprintf """<w:p><w:pPr><w:sectPr>%s<w:pgSz w:w="12240" w:h="15840"/>%s</w:sectPr></w:pPr></w:p>""" refs extra
let threeSectionNative =
    importPackage
        ("""<w:p><w:r><w:t>one</w:t></w:r></w:p>""" + sectionPara """<w:headerReference w:type="default" r:id="rId11"/>""" ""
         + """<w:p><w:r><w:t>two</w:t></w:r></w:p>""" + sectionPara """<w:headerReference w:type="default" r:id="rId12"/>""" """<w:pgNumType w:start="7"/>"""
         + """<w:p><w:r><w:t>three</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId13"/><w:footerReference w:type="default" r:id="rId14"/><w:pgSz w:w="12240" w:h="15840"/><w:pgNumType w:start="3"/></w:sectPr>""")
        [ "word/_rels/document.xml.rels",
          """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"""
          + ([ "rId11", "header", "header1.xml"; "rId12", "header", "header2.xml"; "rId13", "header", "header3.xml"; "rId14", "footer", "footer1.xml" ]
             |> List.map (fun (id, kind, target) -> sprintf """<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/%s" Target="%s"/>""" id kind target)
             |> String.concat "")
          + "</Relationships>"
          "word/header1.xml", sprintf """<w:hdr %s><w:p><w:r><w:t>H1</w:t></w:r></w:p></w:hdr>""" wNsAll
          "word/header2.xml", sprintf """<w:hdr %s><w:p><w:r><w:t>H2</w:t></w:r></w:p></w:hdr>""" wNsAll
          "word/header3.xml", sprintf """<w:hdr %s><w:p><w:r><w:t>H3</w:t></w:r></w:p></w:hdr>""" wNsAll
          "word/footer1.xml", sprintf """<w:ftr %s><w:p><w:r><w:t>F1</w:t></w:r></w:p></w:ftr>""" wNsAll ]
let threeSectionParas = threeSectionNative.Blocks |> List.choose (function Paragraph para | Heading (_, para) -> Some para | _ -> None)
check "native sections: references stay on the paragraph that ends the section"
    (threeSectionParas |> List.map (fun para -> textOfPara para, para.Metadata.HeaderReferences))
    [ "one", []; "", [ "default", "rId11" ]; "two", []; "", [ "default", "rId12" ]; "three", [] ]
check "native sections: the final section's references are on the page setup"
    (threeSectionNative.PageSetup.HeaderReferences, threeSectionNative.PageSetup.FooterReferences, threeSectionNative.PageSetup.PageNumberStart)
    ([ "default", "rId13" ], [ "default", "rId14" ], Some 3)
// Numbering restarts on the first page of the section that asks for it: section
// two starts at "two" (not at the paragraph that ends it), the final one at "three".
check "native sections: page numbering restarts at the section start"
    (pageNumberRestarts threeSectionNative |> List.map (fun (_, paragraphId, start) -> (paragraphId |> Option.bind (fun id -> threeSectionParas |> List.tryFind (fun para -> para.Id = id)) |> Option.map textOfPara), start))
    [ Some "two", 7; Some "three", 3 ]
let threeSectionPageMap = buildPageMap threeSectionNative
check "native sections: display page numbers"
    (threeSectionNative.Blocks |> List.map (fun block -> threeSectionPageMap.BlockAnchors |> List.tryFind (fun anchor -> (match block with Paragraph para | Heading (_, para) -> para.Id | _ -> "") = anchor.PageBlockId) |> Option.map (fun anchor -> anchor.PageNumber)))
    [ Some 1; Some 1; Some 7; Some 7; Some 3 ]

// ----- shapes keep their anchor paragraph and their look ----------------------
let nativeShapeDoc = importPackage ("""<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires="wps" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="7" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>952500</wp:posOffset></wp:positionH><wp:positionV relativeFrom="margin"><wp:align>top</wp:align></wp:positionV><wp:extent cx="1905000" cy="952500"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="3" name="Oval 3"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm rot="1800000"><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="s1" type="#_x0000_t202" style="position:absolute;width:150pt;height:75pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p><w:p><w:r><w:t>three</w:t></w:r></w:p>""") []
let nativeShapes = nativeShapeDoc.Blocks |> List.choose (function TextBox textBox -> Some textBox | _ -> None)
let nativeAnchorText (textBox: TextBoxBlock) =
    textBox.AnchorBlockId
    |> Option.bind (fun id -> nativeShapeDoc.Blocks |> List.tryPick (function Paragraph para when para.Id = id -> Some (textOfPara para) | _ -> None))
check "native shape: one shape, anchored in its paragraph" (nativeShapes |> List.map nativeAnchorText) [ Some "two" ]
check "native shape: geometry, fill, line, wrap, position"
    (nativeShapes |> List.map (fun shape -> shape.Geometry, shape.FillColor, shape.BorderColor, shape.BorderDash, shape.Wrap, shape.BehindDoc, shape.RelativeHeight, shape.PositionHRelativeFrom, shape.PositionVAlign, shape.Rotation))
    [ GeoPreset "ellipse", None, Some "#FF0000", Some "sysDash", Some "square", Some true, Some "7", Some "page", Some "top", 30.0 ]

let nativeFloatingImage =
    importPackage
        """<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="9" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="457200"/><wp:wrapTight wrapText="left"/><wp:docPr id="1" name="p"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId1"/><a:srcRect l="10000" b="5000"/></pic:blipFill><pic:spPr><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>"""
        [ "word/_rels/document.xml.rels", """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.com/a.png" TargetMode="External"/></Relationships>""" ]
check "native floating image: alignment, wrap side, crop, z-order"
    (paragraphsOfBlocks nativeFloatingImage.Blocks
     |> List.collect (fun para -> para.Buffer.Runs)
     |> List.choose (function Rbt.Word.TextBuffer.Image image -> Some (image.Floating, image.PositionHAlign, image.PositionX, image.PositionVRelativeFrom, image.Wrap, image.WrapText, image.BehindDoc, image.RelativeHeight, image.CropLeft, image.CropBottom) | _ -> None))
    [ true, Some "right", None, Some "page", Some "tight", Some "left", Some true, Some "9", Some "10000", Some "5000" ]

// ----- a chart is kept as an opaque object ----------------------------------
let nativeChartDoc =
    importPackage
        """<w:p><w:r><w:t>before</w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="1905000" cy="952500"/><wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId5"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"""
        [ "word/_rels/document.xml.rels", """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>"""
          "word/charts/chart1.xml", """<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>""" ]
let nativeChartRuns =
    paragraphsOfBlocks nativeChartDoc.Blocks
    |> List.collect (fun para -> para.Buffer.Runs)
    |> List.choose (function Rbt.Word.TextBuffer.Image image -> Some (image.Src, image.Width, image.Height, image.Opaque |> Option.map (fun opaque -> opaque.Kind, opaque.Xml.Contains "<c:chart r:id=\"rId5\"/>")) | _ -> None)
check "native chart: an opaque image run with a placeholder source" nativeChartRuns [ "opaque:chart", Some 200.0, Some 100.0, Some ("chart", true) ]
check "native chart: text is not replaced by a placeholder word" (firstParagraphText nativeChartDoc) "before\uFFFC"
check "native chart: its part is kept"
    ((nativeChartDoc.Package.Preserved :?> DocxReader.PackagePreservation).Parts |> List.map (fun part -> part.Path))
    [ "word/charts/chart1.xml" ]

printfn "word import smoke: %d passed, %d failed" pass fail
if fail > 0 then exit 1
