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
                [ XmlDsl.Element("m:t", Map.empty, [ XmlDsl.Text("value & < > \" ' " + string i) ]) ])
        children <- child :: children
    XmlDsl.Element("m:oMath", Map.ofList [ "xmlns:m", "urn:test" ], children)

let escapedXml =
    DocxReader.Docx.xmlToString
        (XmlDsl.Element(
            "m:r",
            Map.ofList [ "z", "2"; "a", "1&\"" ],
            [ XmlDsl.Text "A&B<"; XmlDsl.Element("m:t", Map.empty, []) ]))
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

printfn ""
printfn "OfficeFormat DocxReader tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
