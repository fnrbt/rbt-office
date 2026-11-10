#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

open System
open System.Text

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)

let complexFieldParagraph runCount =
    let sb = StringBuilder()
    sb.Append("<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">") |> ignore
    sb.Append("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>") |> ignore
    sb.Append("<w:r><w:instrText> REF ManyRuns </w:instrText></w:r>") |> ignore
    sb.Append("<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>") |> ignore
    for i in 1 .. runCount do
        sb.Append("<w:r><w:t>") |> ignore
        sb.Append(i % 10) |> ignore
        sb.Append("</w:t></w:r>") |> ignore
    sb.Append("<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>") |> ignore
    sb.Append("</w:p>") |> ignore
    XmlDsl.parse (sb.ToString())

let fieldParagraph = DocxReader.Docx.paragraph (complexFieldParagraph 12)
let field =
    match fieldParagraph.Runs with
    | [ run ] ->
        match run.Content with
        | DocxReader.FieldRef field -> Some field
        | _ -> None
    | _ -> None

ok "complex-field-collapses"
   (Option.isSome field)
   (sprintf "%A" fieldParagraph.Runs)
ok "complex-field-order"
   (field |> Option.exists (fun value -> value.Code = "REF ManyRuns" && value.Result = Some "123456789012"))
   (sprintf "%A" field)

// A result of which parts are formatted differently is as many runs, each with the field's
// code, its whole result and its own part of the text; the number of a page stays one run.
let mixedFieldParagraph (code: string) =
    XmlDsl.parse (
        "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
        + "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText>" + code + "</w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
        + "<w:r><w:t>pla</w:t></w:r><w:r><w:t xml:space=\"preserve\">in </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t xml:space=\"preserve\"> tail</w:t></w:r>"
        + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>")
let fieldPartsOf (code: string) =
    (DocxReader.Docx.paragraph (mixedFieldParagraph code)).Runs
    |> List.map (fun run ->
        match run.Content with
        | DocxReader.FieldRef field -> field.Code, field.Result, field.Part, run.Properties.Bold
        | _ -> "", None, None, None)
ok "complex-field-parts"
   (fieldPartsOf " QUOTE x " = [ "QUOTE x", Some "plain bold tail", Some "plain ", None
                                 "QUOTE x", Some "plain bold tail", Some "bold", Some true
                                 "QUOTE x", Some "plain bold tail", Some " tail", None ])
   (sprintf "%A" (fieldPartsOf " QUOTE x "))
ok "complex-field-page-is-one-run"
   (fieldPartsOf " PAGE " = [ "PAGE", Some "plain bold tail", None, None ])
   (sprintf "%A" (fieldPartsOf " PAGE "))

// The same of a field shown by its own element (w:fldSimple: what the writer makes of a field of the model).
let simpleFieldParts =
    let node =
        XmlDsl.parse (
            "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:fldSimple w:instr=\" QUOTE x \">"
            + "<w:r><w:t xml:space=\"preserve\">plain </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:fldSimple></w:p>")
    (DocxReader.Docx.paragraph node).Runs
    |> List.map (fun run ->
        match run.Content with
        | DocxReader.FieldRef field -> field.Result, field.Part, run.Properties.Bold
        | _ -> None, None, None)
ok "simple-field-parts"
   (simpleFieldParts = [ Some "plain bold", Some "plain ", None; Some "plain bold", Some "bold", Some true ])
   (sprintf "%A" simpleFieldParts)

let perfNode = complexFieldParagraph 6000
DocxReader.Docx.paragraph perfNode |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let before = GC.GetAllocatedBytesForCurrentThread()
let sw = Diagnostics.Stopwatch.StartNew()
DocxReader.Docx.paragraph perfNode |> ignore
sw.Stop()
let allocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - before) / 1024.0 / 1024.0
ok "complex-field-linear-alloc"
   (allocatedMb < 8.5)
   (sprintf "%.1f MB in %.1f ms" allocatedMb sw.Elapsed.TotalMilliseconds)

let xmlStringNode count =
    let mutable children = []
    for i = count downto 1 do
        let child =
            XmlDsl.Element(
                "m:r",
                Map.ofList [ "m:id", string i; "m:val", "a&b<c>d\"e'f" ],
                [ XmlDsl.Element("m:t", Map.empty, [ XmlDsl.Text("value & < > \" ' " + string i) ], XmlDsl.Span.none) ], XmlDsl.Span.none)
        children <- child :: children
    XmlDsl.Element("m:oMath", Map.ofList [ "xmlns:m", "urn:test" ], children, XmlDsl.Span.none)

let escapedXml =
    DocxReader.Docx.xmlToString
        (XmlDsl.Element(
            "m:r",
            Map.ofList [ "z", "2"; "a", "1&\"" ],
            [ XmlDsl.Text "A&B<"; XmlDsl.Element("m:t", Map.empty, [], XmlDsl.Span.none) ], XmlDsl.Span.none))
ok "xml-to-string-escapes"
   (escapedXml = "<m:r a=\"1&amp;&quot;\" z=\"2\">A&amp;B&lt;<m:t/></m:r>")
   escapedXml

let perfXmlNode = xmlStringNode 6000
DocxReader.Docx.xmlToString perfXmlNode |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let beforeXml = GC.GetAllocatedBytesForCurrentThread()
let xmlSw = Diagnostics.Stopwatch.StartNew()
let xmlResult = DocxReader.Docx.xmlToString perfXmlNode
xmlSw.Stop()
let xmlAllocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - beforeXml) / 1024.0 / 1024.0
ok "xml-to-string-linear-alloc"
   (xmlResult.Length > 0 && xmlAllocatedMb < 16.0)
   (sprintf "%.1f MB in %.1f ms len=%d" xmlAllocatedMb xmlSw.Elapsed.TotalMilliseconds xmlResult.Length)

let calibreBytes = IO.File.ReadAllBytes(IO.Path.Combine(__SOURCE_DIRECTORY__, "..", "test-documents", "calibre-demo.docx"))
let calibreDoc = DocxReader.loadBytesNative calibreBytes |> Option.get
DocxReader.toEditorParagraphsWithStylesThemeAndNumbering
    calibreDoc.Theme
    calibreDoc.Content
    calibreDoc.Images
    calibreDoc.Styles
    calibreDoc.NumberingDefinitions
    calibreDoc.Footnotes
    calibreDoc.Endnotes
|> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let editorBefore = GC.GetAllocatedBytesForCurrentThread()
let editorSw = Diagnostics.Stopwatch.StartNew()
let editorItems =
    DocxReader.toEditorParagraphsWithStylesThemeAndNumbering
        calibreDoc.Theme
        calibreDoc.Content
        calibreDoc.Images
        calibreDoc.Styles
        calibreDoc.NumberingDefinitions
        calibreDoc.Footnotes
        calibreDoc.Endnotes
editorSw.Stop()
let editorAllocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - editorBefore) / 1024.0 / 1024.0
ok "editor-conversion-native-alloc"
   (editorItems.Length = calibreDoc.Content.Length && editorAllocatedMb < 24.0)
   (sprintf "items=%d allocated=%.1f MB in %.1f ms" editorItems.Length editorAllocatedMb editorSw.Elapsed.TotalMilliseconds)

// ----- run properties and revisions on a synthetic package -------------------

let docxWithBody (body: string) =
    use stream = new System.IO.MemoryStream()
    do
        use archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new System.IO.StreamWriter(archive.CreateEntry(name).Open(), System.Text.UTF8Encoding(false))
            writer.Write text
        add "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"""
        add "word/document.xml" (sprintf """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>%s</w:body></w:document>""" body)
    stream.ToArray()

let underlineDoc =
    docxWithBody """<w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>on</w:t></w:r><w:r><w:rPr><w:u w:val="none"/></w:rPr><w:t>off</w:t></w:r></w:p><w:p><w:ins w:id="7" w:author="Ann" w:date="2024-01-02T03:04:05Z"><w:r><w:t>added</w:t></w:r></w:ins><w:del w:id="8" w:author="Bob"><w:r><w:delText>removed</w:delText></w:r></w:del></w:p>"""
    |> DocxReader.loadBytesNative
    |> Option.get
let underlineRuns =
    underlineDoc.Content
    |> List.collect (function
        | DocxReader.Para para -> para.Runs
        | _ -> [])
    |> List.choose (fun run ->
        match run.Content with
        | DocxReader.Text text -> Some (text, run.Properties.Underline)
        | _ -> None)
ok "underline-none-is-off"
   (underlineRuns |> List.contains ("on", Some true) && underlineRuns |> List.contains ("off", Some false))
   (sprintf "%A" underlineRuns)
let revisionSummary =
    underlineDoc.Revisions |> List.map (fun revision -> revision.Id, revision.Author, revision.Text)
ok "revisions-read-prefixed-attributes"
   (revisionSummary = [ "7", Some "Ann", "added"; "8", Some "Bob", "removed" ])
   (sprintf "%A" revisionSummary)


// ---- content the reader used to drop -------------------------------------

let wNs = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\""

/// A package with arbitrary extra parts (text or binary) next to document.xml.
let docxWithParts (body: string) (textParts: (string * string) list) (binaryParts: (string * byte[]) list) =
    use stream = new System.IO.MemoryStream()
    do
        use archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new System.IO.StreamWriter(archive.CreateEntry(name).Open(), System.Text.UTF8Encoding(false))
            writer.Write text
        add "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"""
        add "word/document.xml" (sprintf "<w:document %s><w:body>%s</w:body></w:document>" wNs body)
        for name, text in textParts do add name text
        for name, (bytes: byte[]) in binaryParts do
            use entry = archive.CreateEntry(name).Open()
            entry.Write(bytes, 0, bytes.Length)
    stream.ToArray()

let rec paragraphText (para: DocxReader.Paragraph) =
    para.Runs
    |> List.map (fun run ->
        match run.Content with
        | DocxReader.Text text -> text
        | DocxReader.Link link -> link.Text
        | _ -> "")
    |> String.concat ""
let rec contentTexts (content: DocxReader.DocumentContent list) =
    content
    |> List.collect (function
        | DocxReader.Para para -> [ paragraphText para ]
        | DocxReader.Tbl table -> tableTexts table)
and tableTexts (table: DocxReader.Table) =
    table.Rows
    |> List.collect (fun row ->
        row.Cells
        |> List.collect (fun cell -> (cell.Paragraphs |> List.map paragraphText) @ (cell.NestedTables |> List.collect tableTexts)))
let loadBody body = docxWithBody body |> DocxReader.loadBytesNative |> Option.get

let specialRuns =
    loadBody """<w:p><w:r><w:t>e</w:t><w:noBreakHyphen/><w:t>mail co</w:t><w:softHyphen/><w:t>operate</w:t><w:cr/><w:t>next</w:t><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/><w:t>end</w:t></w:r><w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r></w:p>"""
ok "run-special-children-kept"
   (contentTexts specialRuns.Content = [ "e‑mail co­operate\nnext\tend" ])
   (sprintf "%A" (contentTexts specialRuns.Content))
let symFont =
    specialRuns.Content
    |> List.collect (function DocxReader.Para para -> para.Runs | _ -> [])
    |> List.tryPick (fun run -> match run.Content with DocxReader.Text "" -> run.Properties.FontFamily | _ -> None)
ok "sym-keeps-font" (symFont = Some "Wingdings") (sprintf "%A" symFont)

let wrappers =
    docxWithParts
        """<w:p><w:smartTag w:uri="u" w:element="e"><w:r><w:t>smart</w:t></w:r></w:smartTag><w:customXml w:element="x"><w:r><w:t xml:space="preserve"> custom</w:t></w:r></w:customXml><w:dir w:val="rtl"><w:r><w:t xml:space="preserve"> dir</w:t></w:r></w:dir><w:bdo w:val="ltr"><w:r><w:t xml:space="preserve"> bdo</w:t></w:r></w:bdo><mc:AlternateContent><mc:Choice Requires="wps"><w:r><w:t xml:space="preserve"> choice</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t xml:space="preserve"> fallback</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p><w:p><w:moveFrom w:id="1" w:author="Cy"><w:r><w:t>origin</w:t></w:r></w:moveFrom><w:moveTo w:id="2" w:author="Cy"><w:r><w:t>destination</w:t></w:r></w:moveTo></w:p>"""
        [] []
    |> DocxReader.loadBytesNative |> Option.get
ok "paragraph-wrappers-kept"
   (contentTexts wrappers.Content = [ "smart custom dir bdo fallback"; "origindestination" ])
   (sprintf "%A" (contentTexts wrappers.Content))
let moveRevisions =
    wrappers.Content
    |> List.collect (function DocxReader.Para para -> para.Runs | _ -> [])
    |> List.choose (fun run -> match run.Content, run.RevisionType with DocxReader.Text text, Some kind -> Some (text, kind) | _ -> None)
// The moved text is present at both ends, and stays a move (it used to be
// re-labelled as a deletion and an insertion).
ok "tracked-move-kept-as-move"
   (moveRevisions = [ "origin", "moveFrom"; "destination", "moveTo" ])
   (sprintf "%A" moveRevisions)

let sdtTable =
    loadBody """<w:tbl><w:tr><w:tc><w:sdt><w:sdtContent><w:p><w:r><w:t>sdt in cell</w:t></w:r></w:p></w:sdtContent></w:sdt></w:tc><w:sdt><w:sdtContent><w:tc><w:p><w:r><w:t>sdt cell</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt></w:tr><w:sdt><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>sdt row</w:t></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt></w:tbl>"""
ok "table-sdt-wrappers-kept"
   (contentTexts sdtTable.Content = [ "sdt in cell"; "sdt cell"; "sdt row" ])
   (sprintf "%A" (contentTexts sdtTable.Content))

let relsXml (rels: string) = sprintf "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">%s</Relationships>" rels
let relType name = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + name
let drawing rid = sprintf "<w:r><w:drawing><wp:inline><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"1\" name=\"p\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:blipFill><a:blip r:embed=\"%s\"/></pic:blipFill><pic:spPr><a:prstGeom prst=\"rect\"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>" rid
let partsDoc =
    docxWithParts
        (sprintf "<w:p>%s</w:p><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rId9\"/><w:cols w:num=\"2\" w:space=\"708\" w:sep=\"1\"/></w:sectPr>" (drawing "rId1"))
        [ "word/_rels/document.xml.rels", relsXml (sprintf "<Relationship Id=\"rId1\" Type=\"%s\" Target=\"media/body.png\"/><Relationship Id=\"rId9\" Type=\"%s\" Target=\"header1.xml\"/>" (relType "image") (relType "header"))
          "word/header1.xml", sprintf "<w:hdr %s><w:p>%s</w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>in header table</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:hdr>" wNs (drawing "rId1")
          "word/_rels/header1.xml.rels", relsXml (sprintf "<Relationship Id=\"rId1\" Type=\"%s\" Target=\"media/head.png\"/>" (relType "image"))
          "word/settings.xml", sprintf "<w:settings %s><w:evenAndOddHeaders w:val=\"0\"/></w:settings>" wNs ]
        [ "word/media/body.png", [| 1uy; 2uy; 3uy |]; "word/media/head.png", [| 9uy; 8uy; 7uy; 6uy |] ]
    |> DocxReader.loadBytesNative |> Option.get
let imageKeys (paras: DocxReader.Paragraph list) =
    paras |> List.collect (fun para -> para.Runs) |> List.choose (fun run -> match run.Content with DocxReader.Image image -> Some image.RelationshipId | _ -> None)
let bodyImageKey = partsDoc.Content |> List.choose (function DocxReader.Para para -> Some para | _ -> None) |> imageKeys
let headerImageKey = partsDoc.Headers |> List.collect (fun header -> header.Content) |> imageKeys
let imageFor key = key |> List.tryHead |> Option.bind (fun k -> Map.tryFind k partsDoc.Images)
ok "part-images-do-not-collide"
   (bodyImageKey <> headerImageKey && (imageFor bodyImageKey).IsSome && (imageFor headerImageKey).IsSome && imageFor bodyImageKey <> imageFor headerImageKey)
   (sprintf "body %A header %A map %A" bodyImageKey headerImageKey (partsDoc.Images |> Map.toList |> List.map fst))
ok "header-keeps-table-structure"
   (partsDoc.Headers |> List.exists (fun header -> header.Blocks |> List.exists (function DocxReader.Tbl _ -> true | _ -> false)))
   (sprintf "%A" (partsDoc.Headers |> List.map (fun header -> header.Blocks.Length)))
ok "even-and-odd-headers-val-0-is-off"
   (partsDoc.PageSetup |> Option.bind (fun setup -> setup.EvenAndOddHeaders) = Some false)
   (sprintf "%A" (partsDoc.PageSetup |> Option.map (fun setup -> setup.EvenAndOddHeaders)))
ok "column-separator-read"
   (partsDoc.PageSetup |> Option.bind (fun setup -> setup.Columns) |> Option.exists (fun columns -> columns.Count = 2 && columns.Separator))
   (sprintf "%A" (partsDoc.PageSetup |> Option.map (fun setup -> setup.Columns)))
ok "picture-is-not-a-shape" partsDoc.Shapes.IsEmpty (sprintf "%d shapes" partsDoc.Shapes.Length)
// Text whose size no style, no document default and no run states is 10 pt in Word (w:sz's
// default of 20 half-points), not the 11 pt of Normal in Word's own template.
ok "default-run-size-is-10pt-in-px"
   (abs (DocumentProcessors.StyleResolution.defaultRunProps.FontSize - 10.0 * 96.0 / 72.0) < 1e-9)
   (string DocumentProcessors.StyleResolution.defaultRunProps.FontSize)
do
    let unsized =
        docxWithParts
            "<w:p><w:r><w:t>no size anywhere</w:t></w:r></w:p>"
            [ "word/styles.xml", sprintf "<w:styles %s><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Arial\"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style></w:styles>" wNs ]
            []
        |> DocxReader.loadBytesNative |> Option.get
    let items =
        DocxReader.toEditorParagraphsWithStylesThemeAndNumbering
            unsized.Theme unsized.Content unsized.Images unsized.Styles unsized.NumberingDefinitions unsized.Footnotes unsized.Endnotes
        |> Array.map (fun item -> item :?> System.Collections.Generic.Dictionary<string, obj>)
    let sizes =
        items
        |> Array.collect (fun item ->
            match item.TryGetValue "Runs" with
            | true, (:? (obj array) as runs) -> runs |> Array.map (fun run -> Math.Round(unbox<float> ((run :?> System.Collections.Generic.Dictionary<string, obj>).["FontSize"]) * 72.0 / 96.0, 2))
            | _ -> [||])
        |> Array.toList
    ok "a run no style sizes is 10 pt" (sizes = [ 10.0 ]) (sprintf "%A" sizes)

// ----- explicit (direct) run and paragraph formatting for the editor ---------
// The editor objects carry effective formatting; ExplicitX says what was set on
// the run / paragraph itself (null = inherited), so "off" can override a style.
let explicitStylesDoc =
    docxWithParts
        "<w:p><w:pPr><w:pStyle w:val=\"Loud\"/></w:pPr><w:r><w:t>inherits</w:t></w:r><w:r><w:rPr><w:b w:val=\"0\"/><w:i/><w:sz w:val=\"22\"/></w:rPr><w:t>overrides</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val=\"Loud\"/><w:jc w:val=\"left\"/></w:pPr><w:r><w:t>left</w:t></w:r></w:p>"
        [ "word/styles.xml", sprintf "<w:styles %s><w:style w:type=\"paragraph\" w:styleId=\"Loud\"><w:name w:val=\"Loud\"/><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:rPr><w:b/><w:sz w:val=\"40\"/></w:rPr></w:style></w:styles>" wNs ]
        []
    |> DocxReader.loadBytesNative |> Option.get
let explicitEditorItems =
    DocxReader.toEditorParagraphsWithStylesThemeAndNumbering
        explicitStylesDoc.Theme explicitStylesDoc.Content explicitStylesDoc.Images explicitStylesDoc.Styles
        explicitStylesDoc.NumberingDefinitions explicitStylesDoc.Footnotes explicitStylesDoc.Endnotes
    |> Array.map (fun item -> item :?> System.Collections.Generic.Dictionary<string, obj>)
let fieldOf (name: string) (item: System.Collections.Generic.Dictionary<string, obj>) =
    match item.TryGetValue name with
    | true, value -> value
    | _ -> null
let editorRunsOf (item: System.Collections.Generic.Dictionary<string, obj>) =
    (fieldOf "Runs" item :?> obj array) |> Array.map (fun run -> run :?> System.Collections.Generic.Dictionary<string, obj>)
let explicitSummary =
    editorRunsOf explicitEditorItems.[0]
    |> Array.map (fun run ->
        unbox<bool> (fieldOf "Bold" run),
        (fieldOf "ExplicitBold" run :?> bool option),
        (fieldOf "ExplicitItalic" run :?> bool option),
        (fieldOf "ExplicitFontSize" run :?> float option) |> Option.map (fun size -> Math.Round(size, 2)))
    |> Array.toList
ok "editor-run-explicit-fields"
   (explicitSummary = [ (true, None, None, None); (false, Some false, Some true, Some 14.67) ])
   (sprintf "%A" explicitSummary)
let explicitAlignments =
    explicitEditorItems
    |> Array.map (fun item -> string (fieldOf "Alignment" item), (fieldOf "ExplicitAlignment" item :?> string option))
    |> Array.toList
ok "editor-paragraph-explicit-alignment"
   (explicitAlignments = [ ("center", None); ("left", Some "left") ])
   (sprintf "%A" explicitAlignments)

// ----- keepNext / keepLines: the paragraph's own vs. its style's ------------
let keepDoc =
    docxWithParts
        "<w:p><w:pPr><w:pStyle w:val=\"Kept\"/></w:pPr><w:r><w:t>style keeps</w:t></w:r></w:p><w:p><w:pPr><w:keepNext/><w:keepLines w:val=\"0\"/></w:pPr><w:r><w:t>direct</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val=\"Kept\"/><w:keepNext w:val=\"0\"/></w:pPr><w:r><w:t>switched off</w:t></w:r></w:p>"
        [ "word/styles.xml", sprintf "<w:styles %s><w:style w:type=\"paragraph\" w:styleId=\"Kept\"><w:name w:val=\"Kept\"/><w:pPr><w:keepNext/></w:pPr></w:style></w:styles>" wNs ]
        []
    |> DocxReader.loadBytesNative |> Option.get
let keepItems =
    DocxReader.toEditorParagraphsWithStylesThemeAndNumbering
        keepDoc.Theme keepDoc.Content keepDoc.Images keepDoc.Styles keepDoc.NumberingDefinitions keepDoc.Footnotes keepDoc.Endnotes
    |> Array.map (fun item -> item :?> System.Collections.Generic.Dictionary<string, obj>)
let keepSummary =
    keepItems
    |> Array.map (fun item ->
        (fieldOf "KeepNext" item :?> bool option), (fieldOf "KeepLines" item :?> bool option), (fieldOf "StyleKeepNext" item :?> bool option))
    |> Array.toList
ok "editor-paragraph-direct-keep"
   (keepSummary = [ (None, None, Some true); (Some true, Some false, None); (Some false, None, Some true) ])
   (sprintf "%A" keepSummary)

// ----- footnote / endnote / comment bodies are blocks -----------------------
let relsNs = "xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\""
let notesDoc =
    docxWithParts
        "<w:p><w:r><w:t>body</w:t></w:r><w:r><w:footnoteReference w:id=\"2\"/></w:r><w:r><w:endnoteReference w:id=\"1\"/></w:r><w:r><w:commentReference w:id=\"5\"/></w:r></w:p>"
        [ "word/_rels/document.xml.rels",
          sprintf "<Relationships %s><Relationship Id=\"rId1\" Type=\"%s\" Target=\"notes/foot.xml\"/><Relationship Id=\"rId2\" Type=\"%s\" Target=\"endnotes.xml\"/><Relationship Id=\"rId3\" Type=\"%s\" Target=\"comments.xml\"/></Relationships>"
              relsNs (relType "footnotes") (relType "endnotes") (relType "comments")
          "word/notes/foot.xml",
          sprintf "<w:footnotes %s><w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id=\"2\"><w:p><w:r><w:rPr><w:rStyle w:val=\"FootnoteReference\"/></w:rPr><w:footnoteRef/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:hyperlink r:id=\"rId9\"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>tail</w:t></w:r></w:p></w:footnote></w:footnotes>" wNs
          "word/notes/_rels/foot.xml.rels",
          sprintf "<Relationships %s><Relationship Id=\"rId9\" Type=\"%s\" Target=\"https://example.com/n\" TargetMode=\"External\"/></Relationships>" relsNs (relType "hyperlink")
          "word/endnotes.xml",
          sprintf "<w:endnotes %s><w:endnote w:id=\"1\"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>end</w:t></w:r></w:p></w:endnote></w:endnotes>" wNs
          "word/comments.xml",
          sprintf "<w:comments %s><w:comment w:id=\"5\" w:author=\"Ann\" w:initials=\"A\"><w:p><w:r><w:t xml:space=\"preserve\">first </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>remark</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p></w:comment></w:comments>" wNs ]
        []
    |> DocxReader.loadBytesNative |> Option.get
let blockShape (blocks: DocxReader.DocumentContent list) =
    blocks |> List.map (function DocxReader.Para _ -> "p" | DocxReader.Tbl _ -> "tbl")
let footnote2 = notesDoc.Footnotes |> List.tryFind (fun note -> note.Id = "2")
ok "footnote-part-resolved-from-relationships"
   (notesDoc.Footnotes |> List.map (fun note -> note.Id, note.Kind) = [ "-1", Some "separator"; "2", None ])
   (sprintf "%A" (notesDoc.Footnotes |> List.map (fun note -> note.Id, note.Kind)))
ok "footnote-body-keeps-table"
   (footnote2 |> Option.map (fun note -> blockShape note.Blocks) = Some [ "p"; "tbl"; "p" ])
   (sprintf "%A" (footnote2 |> Option.map (fun note -> blockShape note.Blocks)))
ok "footnote-mark-and-style"
   (footnote2 |> Option.map (fun note -> note.HasMark, note.MarkStyle) = Some (true, Some "FootnoteReference"))
   (sprintf "%A" footnote2)
let footnoteRuns =
    footnote2
    |> Option.map (fun note -> DocxReader.blockParagraphs note.Blocks |> List.collect (fun para -> para.Runs))
    |> Option.defaultValue []
ok "footnote-run-formatting"
   (footnoteRuns |> List.exists (fun run -> run.Content = DocxReader.Text "bold" && run.Properties.Bold = Some true))
   (sprintf "%A" footnoteRuns)
ok "footnote-hyperlink-from-own-relationships"
   (footnoteRuns |> List.exists (fun run -> match run.Content with DocxReader.Link link -> link.Text = "link" && link.Target = Some "https://example.com/n" | _ -> false))
   (sprintf "%A" footnoteRuns)
ok "footnote-plain-text"
   (footnote2 |> Option.map DocxReader.Notes.plainText = Some "boldlinkcelltail")
   (sprintf "%A" (footnote2 |> Option.map DocxReader.Notes.plainText))
ok "endnote-body"
   (notesDoc.Endnotes |> List.map (fun note -> note.Id, note.HasMark, blockShape note.Blocks) = [ "1", true, [ "p" ] ])
   (sprintf "%A" notesDoc.Endnotes)
ok "comment-body-blocks"
   ((notesDoc.Comments |> List.map (fun comment -> comment.Id, comment.Author, comment.Initials, comment.Text, blockShape comment.Blocks))
     = [ "5", Some "Ann", Some "A", "first remark second", [ "p"; "p" ] ])
   (sprintf "%A" notesDoc.Comments)
let editorNotes =
    DocxReader.toEditorNotes notesDoc.Theme notesDoc.Images notesDoc.Styles notesDoc.NumberingDefinitions notesDoc.Footnotes
    |> Array.map (fun item -> item :?> System.Collections.Generic.Dictionary<string, obj>)
ok "editor-notes-carry-blocks"
   (editorNotes.Length = 2
    && (editorNotes.[1].["Blocks"] :?> obj array).Length = 3
    && ((editorNotes.[1].["Blocks"] :?> obj array).[1] :?> System.Collections.Generic.Dictionary<string, obj>).["Type"] = box "Table")
   (sprintf "%d notes" editorNotes.Length)

// ----- section properties stay on their own section --------------------------
let sectionsDoc =
    (fun body -> docxWithParts body [] [] |> DocxReader.loadBytesNative |> Option.get)
        ("<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:pPr><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rId11\"/></w:sectPr></w:pPr></w:p>"
         + "<w:p><w:r><w:t>two</w:t></w:r></w:p><w:p><w:pPr><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rId12\"/><w:pgNumType w:fmt=\"lowerRoman\" w:start=\"7\"/></w:sectPr></w:pPr></w:p>"
         + "<w:p><w:r><w:t>three</w:t></w:r></w:p><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rId13\"/><w:footerReference w:type=\"first\" r:id=\"rId14\"/><w:pgNumType w:start=\"3\"/></w:sectPr>")
let sectionRefs =
    sectionsDoc.Content
    |> List.choose (function
        | DocxReader.Para para -> Some (para.Properties.HeaderReferences |> List.map (fun r -> r.RelationshipId), para.Properties.SectionPageNumberStart)
        | _ -> None)
ok "section-references-stay-on-ending-paragraph"
   (sectionRefs = [ [], None; [ "rId11" ], None; [], None; [ "rId12" ], Some 7; [], None ])
   (sprintf "%A" sectionRefs)
let sectionFormats =
    sectionsDoc.Content
    |> List.choose (function
        | DocxReader.Para para when para.Properties.SectionBreak.IsSome -> Some para.Properties.SectionPageNumberFormat
        | _ -> None)
ok "section-page-number-format-is-read"
   (sectionFormats = [ None; Some "lowerRoman" ] && (sectionsDoc.PageSetup |> Option.bind (fun setup -> setup.PageNumberFormat)) = None)
   (sprintf "%A" sectionFormats)
let finalSection = sectionsDoc.PageSetup |> Option.map (fun setup -> setup.HeaderReferences |> List.map (fun r -> r.Kind, r.RelationshipId), setup.FooterReferences |> List.map (fun r -> r.Kind, r.RelationshipId), setup.PageNumberStart)
ok "final-section-references-on-page-setup"
   (finalSection = Some ([ "default", "rId13" ], [ "first", "rId14" ], Some 3))
   (sprintf "%A" finalSection)

// ----- shapes: one per AlternateContent, anchored, with wrap / line / position ---
let shapesDoc =
    docxWithParts "<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires=\"wps\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><w:drawing><wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"7\" behindDoc=\"1\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\"><wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"page\"><wp:posOffset>952500</wp:posOffset></wp:positionH><wp:positionV relativeFrom=\"margin\"><wp:align>top</wp:align></wp:positionV><wp:extent cx=\"1905000\" cy=\"952500\"/><wp:wrapSquare wrapText=\"bothSides\"/><wp:docPr id=\"3\" name=\"Oval 3\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm rot=\"1800000\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1905000\" cy=\"952500\"/></a:xfrm><a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom><a:noFill/><a:ln w=\"19050\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill><a:prstDash val=\"sysDash\"/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" id=\"s1\" type=\"#_x0000_t202\" style=\"position:absolute;width:150pt;height:75pt\"><v:textbox><w:txbxContent><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p><w:p><w:r><w:t>three</w:t></w:r></w:p>" [] []
    |> DocxReader.loadBytesNative |> Option.get
ok "shape-read-once-not-also-from-fallback" (shapesDoc.Shapes.Length = 1) (sprintf "%d shapes" shapesDoc.Shapes.Length)
let shapeSummary =
    shapesDoc.Shapes
    |> List.map (fun shape ->
        shape.AnchorBlock, shape.Preset, shape.Wrap, shape.BehindDoc, shape.RelativeHeight, shape.NoFill, shape.NoLine, shape.LineDash,
        shape.BorderWidth |> Option.map (fun width -> Math.Round(width, 2)), shape.PositionHRelativeFrom, shape.PositionVAlign, shape.Rotation, shape.Name)
ok "shape-anchor-wrap-line-position"
   (shapeSummary = [ Some 1, Some "ellipse", Some "square", Some true, Some "7", true, false, Some "sysDash", Some 2.0, Some "page", Some "top", 30.0, Some "Oval 3" ])
   (sprintf "%A" shapeSummary)
ok "shape-text-is-not-body-text"
   (contentTexts shapesDoc.Content = [ "one"; "two"; "three" ])
   (sprintf "%A" (contentTexts shapesDoc.Content))

// ----- flips: a:xfrm flipH / flipV of a shape and of a picture, VML flip:x / flip:y ----
let flippedShape (flips: string) =
    "<w:p><w:r><w:t>anchor</w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires=\"wps\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><w:drawing><wp:anchor distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" simplePos=\"0\" relativeHeight=\"3\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\"><wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"column\"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx=\"952500\" cy=\"952500\"/><wp:wrapNone/><wp:docPr id=\"4\" name=\"Triangle 4\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:wsp><wps:cNvSpPr/><wps:spPr>"
    + sprintf "<a:xfrm%s><a:off x=\"0\" y=\"0\"/><a:ext cx=\"952500\" cy=\"952500\"/></a:xfrm>" flips
    + "<a:prstGeom prst=\"rtTriangle\"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>text</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback/></mc:AlternateContent></w:r></w:p>"
let flipsOfShapes (body: string) =
    (docxWithParts body [] [] |> DocxReader.loadBytesNative |> Option.get).Shapes |> List.map (fun shape -> shape.FlipH, shape.FlipV)
ok "shape-flip-h" (flipsOfShapes (flippedShape " flipH=\"1\"") = [ true, false ]) (sprintf "%A" (flipsOfShapes (flippedShape " flipH=\"1\"")))
ok "shape-flip-v-and-rot" (flipsOfShapes (flippedShape " rot=\"5400000\" flipV=\"1\"") = [ false, true ]) (sprintf "%A" (flipsOfShapes (flippedShape " rot=\"5400000\" flipV=\"1\"")))
ok "shape-not-flipped" (flipsOfShapes (flippedShape "") = [ false, false ]) (sprintf "%A" (flipsOfShapes (flippedShape "")))
let vmlFlipped (flip: string) =
    sprintf "<w:p><w:r><w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" id=\"s9\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:10pt;margin-top:10pt;width:50pt;height:30pt;%s\"><v:textbox><w:txbxContent><w:p><w:r><w:t>vml</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>" flip
ok "vml-shape-flip-x-y" (flipsOfShapes (vmlFlipped "flip:x y") = [ true, true ]) (sprintf "%A" (flipsOfShapes (vmlFlipped "flip:x y")))
ok "vml-shape-flip-y" (flipsOfShapes (vmlFlipped "flip:y") = [ false, true ]) (sprintf "%A" (flipsOfShapes (vmlFlipped "flip:y")))
let flippedPicture (flips: string) =
    "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"457200\" cy=\"457200\"/><wp:docPr id=\"5\" name=\"pic\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:nvPicPr><pic:cNvPr id=\"5\" name=\"pic\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"rId1\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>"
    + sprintf "<pic:spPr><a:xfrm%s><a:off x=\"0\" y=\"0\"/><a:ext cx=\"457200\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"rect\"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>" flips
    + "<w:r><w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" id=\"p2\" style=\"width:20pt;height:20pt;flip:x\"><v:imagedata r:id=\"rId1\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" o:title=\"\"/></v:shape></w:pict></w:r></w:p>"
let flipsOfPictures (body: string) =
    match (loadBody body).Content with
    | DocxReader.Para para :: _ -> para.Runs |> List.choose (fun run -> match run.Content with DocxReader.Image image -> Some (image.FlipH, image.FlipV) | _ -> None)
    | _ -> []
ok "picture-flips" (flipsOfPictures (flippedPicture " flipH=\"1\" flipV=\"1\"") = [ (true, true); (true, false) ]) (sprintf "%A" (flipsOfPictures (flippedPicture " flipH=\"1\" flipV=\"1\"")))
ok "picture-flip-v-only" (flipsOfPictures (flippedPicture " flipV=\"true\"") = [ (false, true); (true, false) ]) (sprintf "%A" (flipsOfPictures (flippedPicture " flipV=\"true\"")))

// ----- objects the reader does not model are kept, with their parts -----------
let docxPackage (parts: (string * string) list) (binaryParts: (string * byte[]) list) =
    use stream = new System.IO.MemoryStream()
    do
        use archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create, true)
        for name, (text: string) in parts do
            use writer = new System.IO.StreamWriter(archive.CreateEntry(name).Open(), System.Text.UTF8Encoding(false))
            writer.Write text
        for name, (bytes: byte[]) in binaryParts do
            use entry = archive.CreateEntry(name).Open()
            entry.Write(bytes, 0, bytes.Length)
    stream.ToArray()
let chartNs = wNs + " xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\""
let chartDoc =
    docxPackage
        [ "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xlsx\" ContentType=\"application/x-sheet\"/><Override PartName=\"/word/charts/chart1.xml\" ContentType=\"application/x-chart\"/><Override PartName=\"/main/doc.xml\" ContentType=\"application/x-main\"/></Types>"
          "_rels/.rels", relsXml (sprintf "<Relationship Id=\"rId7\" Type=\"%s\" Target=\"main/doc.xml\"/><Relationship Id=\"rId8\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>" (relType "officeDocument"))
          "docProps/core.xml", "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>Kept title</dc:title></cp:coreProperties>"
          // The main part is not word/document.xml: it is found through the package relationships.
          "main/doc.xml", sprintf "<w:document %s><w:body><w:p><w:r><w:t>before</w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx=\"1905000\" cy=\"952500\"/><wp:docPr id=\"1\" name=\"Chart 1\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"rId5\"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>" chartNs
          "main/_rels/doc.xml.rels", relsXml (sprintf "<Relationship Id=\"rId5\" Type=\"%s\" Target=\"/word/charts/chart1.xml\"/><Relationship Id=\"rId6\" Type=\"%s\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId9\" Type=\"%s\" Target=\"media/unused.png\"/>" (relType "chart") (relType "theme") (relType "image"))
          "main/theme/theme1.xml", "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Kept\"/>"
          "word/charts/chart1.xml", "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>"
          "word/charts/_rels/chart1.xml.rels", relsXml (sprintf "<Relationship Id=\"rId1\" Type=\"%s\" Target=\"../embeddings/book.xlsx\"/>" (relType "package")) ]
        [ "word/embeddings/book.xlsx", [| 1uy; 2uy; 3uy |]; "main/media/unused.png", [| 9uy |] ]
    |> DocxReader.loadBytesNative
let chartRuns =
    chartDoc
    |> Option.map (fun doc -> DocxReader.blockParagraphs doc.Content |> List.collect (fun para -> para.Runs))
    |> Option.defaultValue []
ok "main-part-resolved-from-package-relationships"
   (chartDoc |> Option.map (fun doc -> contentTexts doc.Content) = Some [ "before" ])
   (sprintf "%A" (chartDoc |> Option.map (fun doc -> contentTexts doc.Content)))
ok "core-properties-through-package-relationships"
   (chartDoc |> Option.bind (fun doc -> doc.Properties) |> Option.bind (fun props -> props.Title) = Some "Kept title")
   (sprintf "%A" (chartDoc |> Option.bind (fun doc -> doc.Properties)))
let chartOpaque = chartRuns |> List.tryPick (fun run -> match run.Content with DocxReader.Image image -> image.Opaque |> Option.map (fun opaque -> image, opaque) | _ -> None)
ok "chart-is-an-opaque-object"
   (chartOpaque |> Option.map (fun (image, opaque) -> opaque.Kind, opaque.ParagraphLevel, image.Width, image.Height, image.RelationshipId) = Some ("chart", false, Some 200.0, Some 100.0, ""))
   (sprintf "%A" chartOpaque)
ok "opaque-xml-is-self-contained"
   (chartOpaque |> Option.exists (fun (_, opaque) ->
        opaque.Xml.StartsWith "<w:drawing " && opaque.Xml.Contains "xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\""
        && opaque.Xml.Contains "xmlns:wp=" && opaque.Xml.Contains "<c:chart r:id=\"rId5\"/>" && not (opaque.Xml.Contains "xmlns:pic=")
        && (XmlDsl.parse opaque.Xml |> DocxReader.Docx.xmlToString) = opaque.Xml))
   (sprintf "%A" (chartOpaque |> Option.map (fun (_, opaque) -> opaque.Xml)))
let keptParts = chartDoc |> Option.map (fun doc -> doc.Preserved.Parts |> List.map (fun part -> part.Path)) |> Option.defaultValue []
ok "kept-parts-are-those-reachable"
   (keptParts = [ "docProps/core.xml"; "main/theme/theme1.xml"; "word/charts/_rels/chart1.xml.rels"; "word/charts/chart1.xml"; "word/embeddings/book.xlsx" ])
   (sprintf "%A" keptParts)
let keptDocRels =
    chartDoc
    |> Option.bind (fun doc -> Map.tryFind "document" doc.Preserved.PartRelationships)
    |> Option.defaultValue []
    |> List.map (fun rel -> rel.Id, rel.Target)
ok "kept-relationships-have-absolute-targets"
   (keptDocRels = [ "rId5", "/word/charts/chart1.xml"; "rId6", "/main/theme/theme1.xml" ])
   (sprintf "%A" keptDocRels)
ok "kept-package-metadata"
   ((chartDoc |> Option.map (fun doc -> doc.Preserved.DocumentRelationshipId, doc.Preserved.MainContentType, doc.Preserved.RootRelationships |> List.map (fun rel -> rel.Target), doc.Preserved.Defaults))
     = Some ("rId7", Some "application/x-main", [ "/docProps/core.xml" ], [ "xlsx", "application/x-sheet" ]))
   (sprintf "%A" (chartDoc |> Option.map (fun doc -> doc.Preserved.DocumentRelationshipId, doc.Preserved.MainContentType)))
ok "self-contained-xml-keeps-attribute-order"
   ((DocxReader.selfContainedXml (Map.ofList [ "xmlns:a", "urn:a"; "xmlns:b", "urn:b"; "xmlns:mc", "urn:mc"; "mc:Ignorable", "b z" ]) "<a:x z=\"1\" b:k=\"2\"><a:y/></a:x>")
     = "<a:x b:k=\"2\" mc:Ignorable=\"b\" xmlns:a=\"urn:a\" xmlns:b=\"urn:b\" xmlns:mc=\"urn:mc\" z=\"1\"><a:y/></a:x>")
   (DocxReader.selfContainedXml (Map.ofList [ "xmlns:a", "urn:a"; "xmlns:b", "urn:b"; "xmlns:mc", "urn:mc"; "mc:Ignorable", "b z" ]) "<a:x z=\"1\" b:k=\"2\"><a:y/></a:x>")

// ----- settings and styles are handed to the writer as they are --------------
let settingsDoc =
    docxPackage
        [ "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>"
          "word/document.xml", sprintf "<w:document %s><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>" wNs
          "word/_rels/document.xml.rels", relsXml (sprintf "<Relationship Id=\"rId1\" Type=\"%s\" Target=\"conf/settings.xml\"/><Relationship Id=\"rId2\" Type=\"%s\" Target=\"styles.xml\"/>" (relType "settings") (relType "styles"))
          "word/conf/settings.xml", sprintf "<w:settings %s><w:evenAndOddHeaders/><w:defaultTabStop w:val=\"708\"/></w:settings>" wNs
          "word/conf/_rels/settings.xml.rels", relsXml (sprintf "<Relationship Id=\"rId1\" Type=\"%s\" Target=\"file:///t.dotm\" TargetMode=\"External\"/>" (relType "attachedTemplate"))
          "word/styles.xml", sprintf "<w:styles %s><w:latentStyles w:count=\"1\"/><w:style w:type=\"table\" w:styleId=\"Grid\"><w:name w:val=\"Grid\"/></w:style></w:styles>" wNs ]
        []
    |> DocxReader.loadBytesNative |> Option.get
ok "original-settings-and-styles-kept-for-the-writer"
   ((settingsDoc.Preserved.SettingsXml |> Option.exists (fun xml -> xml.Contains "w:defaultTabStop"))
    && (settingsDoc.Preserved.StylesXml |> Option.exists (fun xml -> xml.Contains "latentStyles" && xml.Contains "Grid")))
   (sprintf "%A" (settingsDoc.Preserved.SettingsXml, settingsDoc.Preserved.StylesXml))
let settingsRels = Map.tryFind "settings" settingsDoc.Preserved.PartRelationships |> Option.defaultValue [] |> List.map (fun rel -> rel.Id, rel.Target, rel.External)
ok "settings-relationships-kept" (settingsRels = [ "rId1", "file:///t.dotm", true ]) (sprintf "%A" settingsRels)
ok "regenerated-parts-are-not-kept-as-parts"
   (settingsDoc.Preserved.Parts |> List.map (fun part -> part.Path) = [])
   (sprintf "%A" (settingsDoc.Preserved.Parts |> List.map (fun part -> part.Path)))

// ----- what a paragraph states itself vs. what it inherits -------------------
let explicitParaDoc =
    docxWithParts
        ("<w:p><w:r><w:t>nothing</w:t></w:r></w:p>"
         + "<w:p><w:pPr><w:pStyle w:val=\"Spaced\"/></w:pPr><w:r><w:t>inherits</w:t></w:r></w:p>"
         + "<w:p><w:pPr><w:pStyle w:val=\"Spaced\"/><w:spacing w:after=\"0\"/><w:ind w:left=\"720\"/></w:pPr><w:r><w:rPr><w:color w:val=\"auto\"/></w:rPr><w:t>states</w:t></w:r></w:p>")
        [ "word/styles.xml", sprintf "<w:styles %s><w:style w:type=\"paragraph\" w:styleId=\"Spaced\"><w:name w:val=\"Spaced\"/><w:pPr><w:spacing w:before=\"100\" w:after=\"200\"/><w:ind w:hanging=\"180\"/></w:pPr></w:style></w:styles>" wNs ]
        []
    |> DocxReader.loadBytesNative |> Option.get
let explicitParaItems =
    DocxReader.toEditorParagraphsWithStylesThemeAndNumbering
        explicitParaDoc.Theme explicitParaDoc.Content explicitParaDoc.Images explicitParaDoc.Styles explicitParaDoc.NumberingDefinitions explicitParaDoc.Footnotes explicitParaDoc.Endnotes
    |> Array.map (fun item -> item :?> System.Collections.Generic.Dictionary<string, obj>)
let floatField name (item: System.Collections.Generic.Dictionary<string, obj>) = fieldOf name item :?> float option
let explicitParaSummary =
    explicitParaItems
    |> Array.map (fun item -> floatField "SpacingAfter" item, floatField "ExplicitSpacingAfter" item, floatField "SpacingBefore" item, floatField "ExplicitSpacingBefore" item, floatField "ExplicitIndentLeft" item, floatField "FirstLineIndent" item, floatField "ExplicitFirstLineIndent" item)
    |> Array.toList
ok "editor-paragraph-explicit-vs-inherited"
   (explicitParaSummary =
        [ None, None, None, None, None, None, None
          Some 10.0, None, Some 5.0, None, None, Some -12.0, None
          Some 0.0, Some 0.0, Some 5.0, None, Some 48.0, Some -12.0, None ])
   (sprintf "%A" explicitParaSummary)
ok "editor-run-explicit-automatic-colour"
   ((explicitParaItems.[2].["Runs"] :?> obj array).[0] :?> System.Collections.Generic.Dictionary<string, obj> |> fun run -> run.["ExplicitColorAuto"] = box true)
   "ExplicitColorAuto"

// ----- tracked changes: nested, on the paragraph mark, moves -----------------
let trackedDoc =
    loadBody
        ("<w:p><w:pPr><w:rPr><w:ins w:id=\"12\" w:author=\"Ann\"/></w:rPr></w:pPr><w:ins w:id=\"10\" w:author=\"Ann\"><w:r><w:t>added</w:t></w:r><w:del w:id=\"11\" w:author=\"Bob\"><w:r><w:delText>cut</w:delText></w:r></w:del></w:ins></w:p>"
         + "<w:p><w:moveFromRangeStart w:id=\"14\" w:name=\"move1\"/><w:moveFrom w:id=\"15\" w:author=\"Cy\"><w:r><w:t>moved</w:t></w:r></w:moveFrom></w:p>"
         + "<w:p><w:moveFrom w:id=\"18\" w:author=\"Cy\"><w:r><w:t>too</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id=\"14\"/><w:moveToRangeStart w:id=\"16\" w:name=\"move1\"/><w:moveTo w:id=\"17\" w:author=\"Cy\"><w:r><w:t>moved</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id=\"16\"/></w:p>")
let trackedRuns =
    DocxReader.blockParagraphs trackedDoc.Content
    |> List.collect (fun para -> para.Runs)
    |> List.map (fun run -> (match run.Content with DocxReader.Text text -> text | _ -> ""), run.RevisionType, run.RevisionAuthor, run.InsertedRevision |> Option.bind (fun mark -> mark.Author), run.MoveName)
ok "deletion-inside-insertion-keeps-both"
   (trackedRuns |> List.truncate 2 = [ "added", Some "insert", Some "Ann", None, None; "cut", Some "delete", Some "Bob", Some "Ann", None ])
   (sprintf "%A" trackedRuns)
ok "moves-stay-moves-with-their-range-name"
   (trackedRuns |> List.skip 2 = [ "moved", Some "moveFrom", Some "Cy", None, Some "move1"; "too", Some "moveFrom", Some "Cy", None, Some "move1"; "moved", Some "moveTo", Some "Cy", None, Some "move1" ])
   (sprintf "%A" trackedRuns)
let trackedMarks = DocxReader.blockParagraphs trackedDoc.Content |> List.map (fun para -> para.MarkRevision |> Option.map (fun (revisionType, mark) -> revisionType, mark.Id, mark.Author))
ok "tracked-paragraph-mark"
   (trackedMarks = [ Some ("insert", Some "12", Some "Ann"); None; None ])
   (sprintf "%A" trackedMarks)

// ----- hostile hyperlink targets and linked pictures -------------------------
let hostileDoc =
    docxWithParts
        ("<w:p><w:hyperlink r:id=\"rId1\"><w:r><w:t>bad</w:t></w:r></w:hyperlink><w:hyperlink r:id=\"rId2\"><w:r><w:t>good</w:t></w:r></w:hyperlink><w:hyperlink w:anchor=\"bm\"><w:r><w:t>inside</w:t></w:r></w:hyperlink></w:p>"
         + sprintf "<w:p>%s</w:p>" (drawing "rId3"))
        [ "word/_rels/document.xml.rels",
          relsXml (sprintf "<Relationship Id=\"rId1\" Type=\"%s\" Target=\"javascript:alert(1)\" TargetMode=\"External\"/><Relationship Id=\"rId2\" Type=\"%s\" Target=\"https://example.com/\" TargetMode=\"External\"/><Relationship Id=\"rId3\" Type=\"%s\" Target=\"https://tracker.example.com/p.png\" TargetMode=\"External\"/>"
                       (relType "hyperlink") (relType "hyperlink") (relType "image")) ]
        []
    |> DocxReader.loadBytesNative |> Option.get
let hostileItems =
    DocxReader.toEditorParagraphsWithStylesThemeAndNumbering
        hostileDoc.Theme hostileDoc.Content hostileDoc.Images hostileDoc.Styles hostileDoc.NumberingDefinitions hostileDoc.Footnotes hostileDoc.Endnotes
    |> Array.map (fun item -> item :?> System.Collections.Generic.Dictionary<string, obj>)
let hostileRuns index = (hostileItems.[index].["Runs"] :?> obj array) |> Array.map (fun run -> run :?> System.Collections.Generic.Dictionary<string, obj>)
let linkSummary =
    hostileRuns 0 |> Array.map (fun run -> string run.["Text"], (run.["Hyperlink"] :?> string option), (run.["InactiveHyperlink"] :?> string option)) |> Array.toList
ok "hyperlink-targets-are-allow-listed"
   (linkSummary = [ "bad", None, Some "javascript:alert(1)"; "good", Some "https://example.com/", None; "inside", Some "#bm", None ])
   (sprintf "%A" linkSummary)
ok "external-image-is-marked-linked"
   ((hostileRuns 1).[0].["ImageLinked"] = box true && (hostileRuns 1).[0].["ImageSrc"] = box "https://tracker.example.com/p.png")
   (sprintf "%A %A" (hostileRuns 1).[0].["ImageLinked"] (hostileRuns 1).[0].["ImageSrc"])
ok "active-hyperlink-target-rule"
   ([ "https://a"; "http://a"; "mailto:a@b"; "tel:1"; "#x" ] |> List.forall DocxReader.isActiveHyperlinkTarget
    && not ([ "javascript:alert(1)"; " javascript:x"; "file:///c:/x"; "data:text/html,x"; "vbscript:x"; "x.html"; "#" ] |> List.exists DocxReader.isActiveHyperlinkTarget))
   "rule"

// `xml:space` said once on the root holds for the text of its descendants (Word keeps these spaces);
// whitespace between other elements stays nothing, and an element can say otherwise.
do
    let xml = "<w:document xmlns:w=\"urn:w\" xml:space=\"preserve\"><w:body>\n <w:p>\n <w:r><w:t> </w:t></w:r><w:r><w:t xml:space=\"default\"> no\u00A0</w:t></w:r><w:r><w:t> a </w:t></w:r></w:p></w:body></w:document>"
    let rec texts node =
        match node with
        | XmlDsl.Element ("w:t", _, children, _) -> [ children |> List.choose (function XmlDsl.Text value -> Some value | _ -> None) |> String.concat "" ]
        | XmlDsl.Element (_, _, children, _) -> children |> List.collect texts
        | _ -> []
    let rec strays node =
        match node with
        | XmlDsl.Element ("w:t", _, _, _) -> 0
        | XmlDsl.Element (_, _, children, _) -> (children |> List.sumBy strays)
        | XmlDsl.Text _ -> 1
        | _ -> 0
    let scanned = XmlDsl.tryScan xml (System.Collections.Generic.HashSet<string>())
    ok "xml:space on the root: the text of runs keeps its spaces, by the parser" (texts (XmlDsl.parse xml) = [ " "; "no\u00A0"; " a " ]) (sprintf "%A" (texts (XmlDsl.parse xml)))
    ok "... and by the scanner" (scanned |> Option.map texts = Some [ " "; "no\u00A0"; " a " ]) (sprintf "%A" scanned)
    ok "... whitespace between elements stays nothing" (scanned |> Option.map strays = Some 0) ""

// A paragraph that starts with a w:lastRenderedPageBreak is where Word started a page: which one
// is the number of those marks up to it - also the marks of the table rows before it.
do
    let body =
        "<w:body xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
        + "<w:p><w:r><w:t>first page</w:t></w:r></w:p>"
        + "<w:tbl><w:tr><w:tc><w:p><w:r><w:lastRenderedPageBreak/><w:t>row on page 1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        + "<w:p><w:r><w:lastRenderedPageBreak/><w:t>page 2</w:t></w:r></w:p>"
        + "<w:p><w:r><w:t>text, then </w:t></w:r><w:r><w:lastRenderedPageBreak/><w:t>page 3</w:t></w:r></w:p>"
        + "<w:p><w:r><w:lastRenderedPageBreak/><w:t>page 4</w:t></w:r></w:p>"
        + "</w:body>"
    let marks =
        DocxReader.Docx.bodyContent (XmlDsl.parse body)
        |> List.choose (function DocxReader.Para para -> Some para.Properties.RenderedPageBreakBefore | _ -> None)
    ok "rendered page marks are counted through tables" (marks = [ None; Some 2; None; Some 4 ]) (sprintf "%A" marks)

// A w:br w:type="page" a paragraph starts with (before its text; a bookmark or Word's mark of a page it
// started is not text) is a break before the paragraph: its text is on the next page
// (phpword-sample-11: "Images" went on a page of its own, and its text and picture on the next).
// Anywhere else it is a break after the paragraph's text - in a paragraph that is a break and nothing else
// too (Word keeps its mark on the page: its w:lastRenderedPageBreak is in the next paragraph).
do
    let breaks =
        docxWithParts
            ("<w:p><w:r><w:t>before</w:t></w:r></w:p>"
             + "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"b\"/><w:r><w:br w:type=\"page\"/></w:r><w:r><w:lastRenderedPageBreak/><w:t>starts with a break</w:t></w:r></w:p>"
             + "<w:p><w:r><w:t>text, then</w:t></w:r><w:r><w:br w:type=\"page\"/></w:r></w:p>"
             + "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>"
             + "<w:p><w:r><w:t>after</w:t></w:r></w:p>")
            []
            []
        |> DocxReader.loadBytesNative |> Option.get
    let items =
        DocxReader.toEditorParagraphsWithStylesThemeAndNumbering
            breaks.Theme breaks.Content breaks.Images breaks.Styles breaks.NumberingDefinitions breaks.Footnotes breaks.Endnotes
        |> Array.map (fun item -> item :?> System.Collections.Generic.Dictionary<string, obj>)
    let field (name: string) (item: System.Collections.Generic.Dictionary<string, obj>) =
        match item.TryGetValue name with
        | true, (:? bool as value) -> value
        | true, (:? (bool option) as value) -> value = Some true
        | _ -> false
    let summary = items |> Array.map (fun item -> field "PageBreakBefore" item, field "PageBreakAtStart" item, field "PageBreakAfter" item) |> Array.toList
    ok "a page break a paragraph starts with is a break before it"
        (summary = [ (false, false, false); (true, true, false); (false, false, true); (false, false, true); (false, false, false) ])
        (sprintf "%A" summary)

printfn ""
printfn "OfficeFormat DocxReader tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
