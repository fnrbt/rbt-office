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

// ----- bookmarks: a run has every bookmark that starts before it and ends after it -----
do
    let doc =
        syntheticDocx
            ("""<w:p><w:bookmarkStart w:id="0" w:name="A"/><w:bookmarkStart w:id="1" w:name="B"/><w:bookmarkStart w:id="2" w:name="C"/><w:r><w:t>three</w:t></w:r><w:bookmarkEnd w:id="2"/><w:bookmarkEnd w:id="0"/><w:r><w:t xml:space="preserve"> tail</w:t></w:r><w:bookmarkStart w:id="5" w:name="Point"/><w:bookmarkEnd w:id="5"/><w:r><w:t>after</w:t></w:r><w:bookmarkStart w:id="7" w:name="Open"/></w:p>"""
             + """<w:p><w:r><w:t>next</w:t></w:r><w:bookmarkEnd w:id="1"/><w:bookmarkEnd w:id="7"/></w:p>"""
             + """<w:p><w:r><w:t xml:space="preserve">form </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:bookmarkStart w:id="9" w:name="Text1"/><w:r><w:instrText> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>value</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="9"/></w:p>""")
        |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let runsOf index = (doc.Blocks |> List.choose (function Paragraph para -> Some para | _ -> None)).[index].Buffer.Runs |> List.choose (function Rbt.Word.TextBuffer.Text run -> Some run | _ -> None)
    let first = runsOf 0
    let start id name : Rbt.Word.TextBuffer.BookmarkStart = { Id = id; Name = name; EndsHere = false }
    check "three bookmarks start before a run: all three, with their names" first.[0].Style.BookmarkStarts [ start "0" "A"; start "1" "B"; start "2" "C" ]
    check "two end after it: both, in the order of the file" first.[0].Style.BookmarkEnds [ Rbt.Word.TextBuffer.Ends "2"; Rbt.Word.TextBuffer.Ends "0" ]
    check "a bookmark that ends where it starts stands before the next run" first.[2].Style.BookmarkStarts [ { start "5" "Point" with EndsHere = true } ]
    check "one that is opened after the last run of its paragraph" first.[2].Style.BookmarkEnds [ Rbt.Word.TextBuffer.Opens (start "7" "Open") ]
    check "the second paragraph: what ends after its run" ((runsOf 1).[0].Style.BookmarkStarts, (runsOf 1).[0].Style.BookmarkEnds) ([], [ Rbt.Word.TextBuffer.Ends "1"; Rbt.Word.TextBuffer.Ends "7" ])
    // (The one opened after the last run of a paragraph too: it starts at the paragraph's end.)
    check "the bookmarks the document lists" (bookmarkInfos doc |> List.map (fun info -> info.Id, info.Name) |> List.sort) [ ("0", "A"); ("1", "B"); ("2", "C"); ("5", "Point"); ("7", "Open"); ("9", "Text1") ]
    check "the one opened after the last run: where it starts" (bookmarkInfos doc |> List.filter (fun info -> info.Name = "Open") |> List.map (fun info -> info.StartOffset, info.EndOffset)) [ (15, 15) ]
    let form = runsOf 2
    check "a bookmark that starts inside a field (a form field's name) is the field's run's" (form |> List.map (fun run -> run.Style.BookmarkStarts, run.Style.BookmarkEnds, run.Style.FieldCode.IsSome)) [ ([], [], false); ([ start "9" "Text1" ], [ Rbt.Word.TextBuffer.Ends "9" ], true) ]
    // One of the three removed: the other two stay, and the end of the one that is gone goes with it.
    let removed = removeBookmark "A" doc
    let after = (removed.Blocks |> List.choose (function Paragraph para -> Some para | _ -> None)).[0].Buffer.Runs |> List.choose (function Rbt.Word.TextBuffer.Text run -> Some run | _ -> None)
    check "one of three removed" (after.[0].Style.BookmarkStarts, after.[0].Style.BookmarkEnds) ([ start "1" "B"; start "2" "C" ], [ Rbt.Word.TextBuffer.Ends "2" ])
    // The one opened after the last run removed: its opening and its end in the next paragraph go.
    let removedOpen = removeBookmark "Open" doc
    let runsAfter index = (removedOpen.Blocks |> List.choose (function Paragraph para -> Some para | _ -> None)).[index].Buffer.Runs |> List.choose (function Rbt.Word.TextBuffer.Text run -> Some run | _ -> None)
    check "the one opened after the last run removed: its opening and its end" ((runsAfter 0).[2].Style.BookmarkEnds, (runsAfter 1).[0].Style.BookmarkEnds) ([], [ Rbt.Word.TextBuffer.Ends "1" ])
    let renamedOpen = renameBookmark "Open" "Opened" doc
    check "... renamed" (bookmarkInfos renamedOpen |> List.map (fun info -> info.Name) |> List.sort) [ "A"; "B"; "C"; "Opened"; "Point"; "Text1" ]

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

// ----- what a paragraph states itself vs. what it inherits -------------------
let explicitNativeDoc =
    syntheticDocxWithStyles
        """<w:style w:type="paragraph" w:styleId="Spaced"><w:name w:val="Spaced"/><w:pPr><w:spacing w:before="100" w:after="200"/></w:pPr></w:style>"""
        ("""<w:p><w:pPr><w:pStyle w:val="Spaced"/></w:pPr><w:r><w:t>inherits</w:t></w:r></w:p>"""
         + """<w:p><w:pPr><w:pStyle w:val="Spaced"/><w:spacing w:after="0"/><w:ind w:left="720"/><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:r><w:rPr><w:color w:val="auto"/></w:rPr><w:t>states</w:t></w:r></w:p>""")
    |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
check "native explicit paragraph formatting"
    (paragraphsOfBlocks explicitNativeDoc.Blocks
     |> List.map (fun para -> para.Metadata.SpacingAfter, para.Metadata.Explicit, para.Metadata.HasMarkProperties, para.Metadata.ExplicitMarkFontSize |> Option.map (fun size -> Math.Round(size, 2))))
    [ Some 10.0, 0, false, None
      Some 0.0, ParagraphExplicit.SpacingAfter ||| ParagraphExplicit.IndentLeft, true, Some 18.67 ]
check "native explicit automatic colour"
    (paragraphsOfBlocks explicitNativeDoc.Blocks |> List.map (fun para -> runStyles para |> List.map (fun style -> Rbt.Word.TextBuffer.ExplicitFormat.has Rbt.Word.TextBuffer.ExplicitFormat.ColorAuto style)))
    [ [ false ]; [ true ] ]

// ----- hostile hyperlink targets and linked pictures --------------------------
let hostileNative =
    importPackage
        """<w:p><w:hyperlink r:id="rId1"><w:r><w:t>bad</w:t></w:r></w:hyperlink><w:hyperlink r:id="rId2"><w:r><w:t>good</w:t></w:r></w:hyperlink></w:p><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="p"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId3"/></pic:blipFill><pic:spPr><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"""
        [ "word/_rels/document.xml.rels", """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://tracker.example.com/p.png" TargetMode="External"/></Relationships>""" ]
let hostileNativeRuns = paragraphsOfBlocks hostileNative.Blocks |> List.collect (fun para -> para.Buffer.Runs)
check "native hyperlink targets are allow-listed"
    (hostileNativeRuns |> List.choose (function Rbt.Word.TextBuffer.Text run -> Some (run.Style.Hyperlink, run.Style.InactiveHyperlink) | _ -> None))
    [ None, Some "javascript:alert(1)"; Some "https://example.com/", None ]
check "native external picture is linked and shows a placeholder"
    (hostileNativeRuns |> List.choose (function Rbt.Word.TextBuffer.Image image -> Some (image.Linked, image.Src, Rbt.Word.TextBuffer.displayedImageSource image) | _ -> None))
    [ true, "https://tracker.example.com/p.png", "opaque:linked" ]

// ---- the selection's format over whole documents ------------------------------
// WordStyles.selectionFormat works each property out on its own, stops looking
// at a property once it is mixed and makes no style record per run: for every
// test document, with everything selected and with ranges inside it, it says
// what resolving every run and every paragraph in full says.
module FormatReference =
    module WS = Rbt.Word.WordStyles
    module TB = Rbt.Word.TextBuffer
    let add value current = WS.FormatValue.add value current
    let format (doc: Document) : WS.SelectionFormat =
        let context = WS.styleContext doc
        let headingLevel (id: string) = doc.Blocks |> List.tryPick (function Heading (level, para) when para.Id = id -> Some level | _ -> None)
        let styleOf (para: ParagraphBlock) =
            match headingLevel para.Id with
            | Some level -> WS.forBlockInContext context (Heading (level, para))
            | None -> WS.forBlockInContext context (Paragraph para)
        let mutable f : WS.SelectionFormat =
            { Bold = WS.NoValue; Italic = WS.NoValue; Underline = WS.NoValue; Strike = WS.NoValue; FontFamily = WS.NoValue; FontSize = WS.NoValue
              Color = WS.NoValue; Highlight = WS.NoValue; VertAlign = WS.NoValue; StyleId = WS.NoValue; HeadingLevel = WS.NoValue; Alignment = WS.NoValue; ListKind = WS.NoValue
              ListLevel = WS.NoValue; IndentLeft = WS.NoValue; IndentRight = WS.NoValue; FirstLineIndent = WS.NoValue; ParagraphLineSpacing = WS.NoValue
              ParagraphLineSpacingRule = WS.NoValue; ParagraphSpacingBefore = WS.NoValue; ParagraphSpacingAfter = WS.NoValue }
        let character (paragraphStyle: WS.ParagraphStyle) (runStyle: TB.TextStyle) =
            let style = WS.effectiveTextStyleInContext context paragraphStyle runStyle
            f <- { f with
                    Bold = add style.Bold f.Bold; Italic = add style.Italic f.Italic; Underline = add style.Underline f.Underline
                    Strike = add (style.Strike || style.DoubleStrike) f.Strike; FontFamily = add style.FontFamily f.FontFamily
                    FontSize = add style.FontSize f.FontSize; Color = add style.Color f.Color; Highlight = add style.Highlight f.Highlight
                    VertAlign = add style.VertAlign f.VertAlign }
        let paragraph (para: ParagraphBlock) (paragraphStyle: WS.ParagraphStyle) =
            let metadata = WS.effectiveMetadata paragraphStyle para.Metadata
            let level =
                match headingLevel para.Id with
                | Some level -> level
                | None -> (match paragraphStyle.OutlineLevel with Some level when level >= 1 && level <= 9 -> level | _ -> 0)
            let visible =
                para.Metadata.Style
                |> Option.filter (fun styleId ->
                    not (WS.isInternalStyleId styleId)
                    && (match Map.tryFind ("paragraph", styleId) context.Styles with Some style -> style.StyleType <> "defaults" | None -> true))
            f <- { f with
                    ParagraphLineSpacing = add (metadata.LineSpacing |> Option.orElse paragraphStyle.LineSpacing) f.ParagraphLineSpacing
                    ParagraphLineSpacingRule = add (metadata.LineSpacingRule |> Option.orElse paragraphStyle.LineSpacingRule) f.ParagraphLineSpacingRule
                    ParagraphSpacingBefore = add (metadata.SpacingBefore |> Option.orElse paragraphStyle.SpacingBefore |> Option.defaultValue 0.0) f.ParagraphSpacingBefore
                    ParagraphSpacingAfter = add (metadata.SpacingAfter |> Option.orElse paragraphStyle.SpacingAfter |> Option.defaultValue 0.0) f.ParagraphSpacingAfter
                    StyleId = add visible f.StyleId; HeadingLevel = add level f.HeadingLevel; Alignment = add metadata.Alignment f.Alignment
                    ListKind = add metadata.ListStyle f.ListKind; ListLevel = add (if metadata.ListStyle = NoList then 0 else metadata.ListLevel) f.ListLevel
                    IndentLeft = add (metadata.IndentLeft |> Option.defaultValue 0.0) f.IndentLeft
                    IndentRight = add (metadata.IndentRight |> Option.defaultValue 0.0) f.IndentRight
                    FirstLineIndent = add (metadata.FirstLineIndent |> Option.defaultValue 0.0) f.FirstLineIndent }
        let ranges = selectedParagraphRanges doc
        for (para, rangeStart, rangeEnd) in ranges do
            let paragraphStyle = styleOf para
            paragraph para paragraphStyle
            for run in para.Buffer.Runs do
                match run with
                | TB.Text textRun when textRun.Length > 0 && textRun.Start < rangeEnd && textRun.Start + textRun.Length > rangeStart -> character paragraphStyle textRun.Style
                | _ -> ()
        if f.Bold = WS.NoValue then
            match ranges with
            | (para, rangeStart, _) :: _ -> character (styleOf para) (TB.getStyleAt (if rangeStart > 0 then rangeStart - 1 else 0) para.Buffer)
            | [] -> ()
        f

let root = IO.Path.GetFullPath(IO.Path.Combine(__SOURCE_DIRECTORY__, "../.."))
let formatDocuments =
    IO.Directory.GetDirectories(IO.Path.Combine(root, "tests"), "test-documents*")
    |> Array.collect (fun dir -> IO.Directory.GetFiles(dir, "*.docx", IO.SearchOption.AllDirectories))
    |> Array.sort
let mutable formatRanges = 0
for path in formatDocuments do
    match (try DocxReader.loadBytesNative (IO.File.ReadAllBytes path) with _ -> None) with
    | None -> ()
    | Some parsed ->
        let doc = Rbt.Word.Import.importDoc parsed
        let paragraphs = (selectedParagraphRanges (selectAll doc)) |> List.map (fun (para, _, _) -> para) |> List.toArray
        if paragraphs.Length >= 2 then
            let pick (first: int) (last: int) =
                let a, b = paragraphs.[first], paragraphs.[last]
                setSelection { BlockId = a.Id; Offset = 0 } { BlockId = b.Id; Offset = Rbt.Word.TextBuffer.length b.Buffer } doc
            let n = paragraphs.Length
            for selected in [ selectAll doc; pick 0 (n / 2); pick (n / 3) (n - 1); pick (n / 2) (min (n - 1) (n / 2 + 20)); pick (n - 1) 0 ] do
              // (A caret is another matter: what typing would continue with.)
              if not (selectedParagraphRanges selected).IsEmpty then
                formatRanges <- formatRanges + 1
                let got = Rbt.Word.WordStyles.selectionFormat selected
                let expected = FormatReference.format selected
                if got <> expected then
                    fail <- fail + 1
                    printfn "FAIL selection format of %s: expected %A got %A" (IO.Path.GetFileName path) expected got
ok "selection format over the test documents" (formatRanges > 250) (sprintf "%d ranges of %d documents" formatRanges formatDocuments.Length)
pass <- pass + formatRanges

// ----- the indents of a numbered paragraph ------------------------------------
// What a paragraph states itself (w:pPr/w:ind) counts first; then the level of the numbering the
// paragraph itself refers to (w:pPr/w:numPr: the level's w:pPr/w:ind, 17.9.23 - the numbering's
// paragraph properties are applied over the paragraph style's); then the paragraph style; a
// paragraph numbered through its style (the style has the w:numPr, the level names the style) has
// the style's indent where the style states one and the level's where it does not. Attribute by
// attribute: a paragraph that states w:left alone keeps the level's hanging indent.
// file-sample.docx: "Simple Tables" is <w:pStyle w:val="ListParagraph"/> (w:ind w:left="720") with
// w:numId 2, whose level 0 has <w:ind w:left="1080" w:hanging="360"/>: text at 1080 twips = 72 px
// (0.75 in), the bullet 360 twips = 24 px before it.
let fileSample = importFixture "tests/test-documents-online/file-sample.docx" Rbt.Word.Import.importDoc
let indentsOf (doc: Document) (text: string) =
    paragraphsOfBlocks doc.Blocks
    |> List.tryFind (fun para -> (textOfPara para).StartsWith text)
    |> Option.map (fun para -> para.Metadata.IndentLeft, para.Metadata.FirstLineIndent)
check "file-sample: the level's indent, not the style's" (indentsOf fileSample "Simple Tables", indentsOf fileSample "Complex Tables") (Some (Some 72.0, Some -24.0), Some (Some 72.0, Some -24.0))
let numberedDoc =
    let numbered style extra text = sprintf """<w:p><w:pPr><w:pStyle w:val="%s"/>%s</w:pPr><w:r><w:t>%s</w:t></w:r></w:p>""" style extra text
    let level left = sprintf """<w:ind w:left="%d" w:hanging="360"/>""" left
    importPackage
        (numbered "ListParagraph" """<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>""" "by the level"
         + numbered "ListParagraph" """<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="2010"/>""" "its own left"
         + numbered "ListParagraph" """<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="2010" w:hanging="150"/>""" "its own both"
         + numbered "ListParagraph" """<w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>""" "second level"
         + numbered "NumStyled" "" "style with an indent"
         + numbered "NumPlain" "" "style without"
         + numbered "ListParagraph" "" "not numbered")
        [ "word/styles.xml",
          sprintf """<w:styles %s><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="NumStyled"><w:name w:val="Num Styled"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr><w:ind w:left="300"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="NumPlain"><w:name w:val="Num Plain"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:style></w:styles>""" wNsAll
          "word/numbering.xml",
          sprintf """<w:numbering %s><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%%1."/><w:pPr>%s</w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%%2."/><w:pPr>%s</w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="NumStyled"/><w:lvlText w:val="%%1."/><w:pPr>%s</w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="NumPlain"/><w:lvlText w:val="%%1."/><w:pPr>%s</w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num></w:numbering>""" wNsAll (level 1080) (level 1800) (level 1440) (level 1440) ]
check "numbered by its own numPr: the level's indents over the style's" (indentsOf numberedDoc "by the level") (Some (Some 72.0, Some -24.0))
check "the paragraph's own left, the level's hanging" (indentsOf numberedDoc "its own left") (Some (Some 134.0, Some -24.0))
check "the paragraph's own left and hanging" (indentsOf numberedDoc "its own both") (Some (Some 134.0, Some -10.0))
check "the second level's" (indentsOf numberedDoc "second level") (Some (Some 120.0, Some -24.0))
check "numbered through a style that states an indent: the style's left, the level's hanging" (indentsOf numberedDoc "style with an indent") (Some (Some 20.0, Some -24.0))
check "numbered through a style that states none: the level's" (indentsOf numberedDoc "style without") (Some (Some 96.0, Some -24.0))
check "not numbered: the style's" (indentsOf numberedDoc "not numbered") (Some (Some 48.0, None))

printfn "word import smoke: %d passed, %d failed" pass fail
if fail > 0 then exit 1
