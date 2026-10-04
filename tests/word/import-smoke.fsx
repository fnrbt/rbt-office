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

printfn "word import smoke: %d passed, %d failed" pass fail
if fail > 0 then exit 1
