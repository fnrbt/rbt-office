#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

open System
open System.IO
open System.IO.Compression
open System.Text
open System.Xml.Linq
open Rbt.Word.Document
module Docx = Rbt.Word.Docx

let mutable checks = 0
let check name expected actual =
    checks <- checks + 1
    if actual <> expected then failwithf "%s: expected %A, got %A" name expected actual
let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
let r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
let ns = XNamespace.Get w
let wrap tag inner = sprintf "<w:%s xmlns:w=\"%s\" xmlns:r=\"%s\" xmlns:ext=\"urn:kept\">%s</w:%s>" tag w r inner tag
let paragraph text = "<w:p><w:r><w:t>" + text + "</w:t></w:r></w:p>"
let body =
    "<w:p ext:keep=\"body\"><w:pPr><w:keepNext/></w:pPr><w:bookmarkStart w:id=\"1\" w:name=\"Marker\"/><w:r><w:t>body</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>"
    + "<w:p><w:r><w:footnoteReference w:id=\"2\"/></w:r><w:r><w:footnoteReference w:id=\"3\"/></w:r><w:r><w:commentReference w:id=\"5\"/></w:r></w:p>"
    + "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rHeader\"/><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
let header = wrap "hdr" (paragraph "header")
let notes = wrap "footnotes" ("<w:footnote w:id=\"2\" ext:keep=\"note\">" + paragraph "note" + "</w:footnote><w:footnote w:id=\"3\">" + paragraph "other note" + "</w:footnote>")
let comments = wrap "comments" ("<w:comment w:id=\"5\" w:author=\"Ann\" ext:keep=\"comment\">" + paragraph "comment" + "</w:comment>")
let rel id kind target = sprintf "<Relationship Id=\"%s\" Type=\"%s/%s\" Target=\"%s\"/>" id r kind target
let relationships inner = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + inner + "</Relationships>"
let originals =
    [ "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/><Override PartName=\"/word/footnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/><Override PartName=\"/word/comments.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml\"/></Types>"
      "_rels/.rels", relationships (rel "rDoc" "officeDocument" "word/document.xml")
      "word/document.xml", wrap "document" ("<w:body>" + body + "</w:body>")
      "word/_rels/document.xml.rels", relationships (rel "rHeader" "header" "header1.xml" + rel "rNotes" "footnotes" "footnotes.xml" + rel "rComments" "comments" "comments.xml" + rel "rCustom" "customXml" "../customXml/item1.xml")
      "word/header1.xml", header
      "word/footnotes.xml", notes
      "word/comments.xml", comments
      "customXml/item1.xml", "<kept xmlns=\"urn:opaque\" value=\"original\">untouched</kept>" ] |> Map.ofList
let packageBytes (entries: Map<string, string>) =
    use stream = new MemoryStream()
    do
        use zip = new ZipArchive(stream, ZipArchiveMode.Create, true)
        for KeyValue(path, content) in entries do
            use writer = new StreamWriter(zip.CreateEntry(path).Open(), UTF8Encoding(false))
            writer.Write content
    stream.ToArray()
let source = packageBytes originals
let load bytes = Docx.loadBytesNative bytes |> Option.defaultWith (fun () -> failwith "DOCX failed to load")
let markerSettings : Rbt.Word.Layout.LayoutSettings =
    { Width = 600.0; PageHeight = 1000.0; FontSize = 16.0; LineHeight = 24.0
      PaddingLeft = 0.0; PaddingTop = 0.0; MarginRight = 0.0; MarginBottom = 0.0
      InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 0.25; DefaultTabStop = 48.0 }
let controlFrameCount doc =
    let measureText (value: string) (_: Rbt.Word.TextBuffer.TextStyle) = float value.Length * 7.0
    Rbt.Word.Layout.layoutDocument markerSettings measureText doc
    |> Rbt.Word.Paint.markupMarkerCommands (contentControlMemberships doc) markerSettings doc
    |> List.sumBy (function Rbt.Word.Paint.StrokeRect ("#2563EB", _, _, _, _, _) -> 1 | _ -> 0)
let write unchanged doc =
    let package = Docx.exportTypedDocumentPackage doc
    DocumentWriter.writeTypedPackageBytes { package with Unchanged = unchanged }
let parts bytes =
    use stream = new MemoryStream(bytes: byte[])
    use zip = new ZipArchive(stream, ZipArchiveMode.Read)
    zip.Entries |> Seq.map (fun entry ->
        use reader = new StreamReader(entry.Open())
        entry.FullName, reader.ReadToEnd()) |> Map.ofSeq
let text xml = XDocument.Parse(xml).Descendants(ns + "t") |> Seq.map (fun el -> el.Value) |> String.concat ""
let doc = load source
check "unchanged save preserves the entire original container" source (write true doc)
let first = paragraphsOfBlocks doc.Blocks |> List.head
let edited = doc |> setCursor first.Id 4 |> insertText " edited"
let saved = write false edited
let savedParts = parts saved
check "body edit reaches real DOCX text" "body edited" (text savedParts.["word/document.xml"])
for path in [ "customXml/item1.xml"; "word/header1.xml"; "word/footnotes.xml"; "word/comments.xml" ] do
    check ("body edit preserves " + path) originals.[path] savedParts.[path]
let savedBody = XDocument.Parse savedParts.["word/document.xml"]
check "unknown body attribute retained" [ "body" ] (savedBody.Descendants(ns + "p") |> Seq.choose (fun el -> Option.ofObj (el.Attribute(XName.Get("keep", "urn:kept"))) |> Option.map (fun a -> a.Value)) |> Seq.toList)
check "bookmark kept exactly once" [ "Marker" ] (savedBody.Descendants(ns + "bookmarkStart") |> Seq.map (fun el -> el.Attribute(ns + "name").Value) |> Seq.toList)
check "reopened body retains edit" "body edited" (load saved |> fun value -> paragraphsOfBlocks value.Blocks |> List.head |> fun p -> Rbt.Word.TextBuffer.toString p.Buffer)

// New columns adjacent to bound cells must not become part of the control.
let boundCellTable =
    "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>"
    + "<w:tr><w:tc>" + paragraph "left" + "</w:tc>"
    + "<w:sdt><w:sdtPr><w:id w:val=\"29\"/><w:dataBinding w:xpath=\"/invoice/price\" w:storeItemID=\"{8B049945-9DFE-4726-9DE9-CF5691E53858}\"/><w:text/></w:sdtPr>"
    + "<w:sdtContent><w:tc>" + paragraph "price" + "</w:tc></w:sdtContent></w:sdt></w:tr></w:tbl>"
let boundCellDoc =
    originals |> Map.add "word/document.xml" (wrap "document" ("<w:body>" + boundCellTable + "</w:body>"))
    |> packageBytes |> load
let boundCellParagraph = boundCellDoc.Blocks |> List.pick (function Table table -> Some table.Rows.Head.Cells.[1].Content.Head | _ -> None)
for name, insert in [ "before", insertTableColumnBeforeSelection; "after", insertTableColumnAfterSelection ] do
    let result = boundCellDoc |> setCursor boundCellParagraph.Id 0 |> insert |> write false |> parts
    let tree = XDocument.Parse result.["word/document.xml"]
    let row = tree.Descendants(ns + "tr") |> Seq.head
    let control = row.Element(ns + "sdt")
    check ("bound cell " + name + ": one cell remains in control") 1 (control.Element(ns + "sdtContent").Elements(ns + "tc") |> Seq.length)
    check ("bound cell " + name + ": original binding preserved") "/invoice/price" (control.Descendants(ns + "dataBinding") |> Seq.head |> fun binding -> binding.Attribute(ns + "xpath").Value)
    check ("bound cell " + name + ": new cell outside control")
        (if name = "before" then [ "tc"; "tc"; "sdt" ] else [ "tc"; "sdt"; "tc" ])
        (row.Elements() |> Seq.map (fun element -> element.Name.LocalName) |> Seq.toList)
let boundCellCaret = boundCellDoc |> setCursor boundCellParagraph.Id 2
check "single-line cell control rejects a paragraph break without changing saved content"
    (boundCellCaret |> write false |> parts) (boundCellCaret |> splitBlock |> write false |> parts)
let multilineCellDoc =
    originals
    |> Map.add "word/document.xml" (wrap "document" ("<w:body>" + boundCellTable.Replace("<w:text/>", "<w:text w:multiLine=\"on\"/>") + "</w:body>"))
    |> packageBytes |> load
let multilineCellParagraph = multilineCellDoc.Blocks |> List.pick (function Table table -> Some table.Rows.Head.Cells.[1].Content.Head | _ -> None)
let multilineSplit = multilineCellDoc |> setCursor multilineCellParagraph.Id 2 |> splitBlock
check "multiline cell control allows a paragraph break" 2
    (multilineSplit.Blocks |> List.pick (function Table table -> Some table.Rows.Head.Cells.[1].Content.Length | _ -> None))
let inlineControlDoc =
    originals
    |> Map.add "word/document.xml" (wrap "document" ("<w:body><w:p><w:sdt><w:sdtPr><w:id w:val=\"31\"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>hello</w:t></w:r></w:sdtContent></w:sdt></w:p></w:body>"))
    |> packageBytes |> load
let inlineParagraph = paragraphsOfBlocks inlineControlDoc.Blocks |> List.head
let inlineCaret = inlineControlDoc |> setCursor inlineParagraph.Id 2
check "inline control is not duplicated by a paragraph break"
    (inlineCaret |> write false |> parts) (inlineCaret |> splitBlock |> write false |> parts)
let formattedInline =
    inlineControlDoc |> setCursor inlineParagraph.Id 1 |> selectRange 1 2 |> toggleBold |> write false |> parts
let formattedInlineXml = XDocument.Parse formattedInline.["word/document.xml"]
let formattedInlineRuns = formattedInlineXml.Descendants(ns + "sdtContent") |> Seq.head |> fun content -> content.Elements(ns + "r") |> Seq.toList
check "formatting plain control keeps one run" 1 formattedInlineRuns.Length
check "formatting part of plain control formats its whole text" "hello"
    (formattedInlineRuns.Head.Element(ns + "t").Value)
check "formatting plain control sets bold" true
    (not (isNull (formattedInlineRuns.Head.Element(ns + "rPr").Element(ns + "b"))))
let invoiceBytes = File.ReadAllBytes(Path.Combine(__SOURCE_DIRECTORY__, "../test-documents-business/docx4j-invoice-2013.docx"))
let invoiceDoc = load invoiceBytes
let pictureParagraph =
    paragraphsOfBlocks invoiceDoc.Blocks
    |> List.find (fun para -> para.Metadata.ContentControlXml |> Option.exists (fun xml -> xml.Contains "<w:picture"))
let invoiceTyped = invoiceDoc |> setCursor pictureParagraph.Id 0 |> insertText "Q"
let invoiceTypedXml = invoiceTyped |> write false |> parts |> fun values -> XDocument.Parse values.["word/document.xml"]
let pictureControl =
    invoiceTypedXml.Descendants(ns + "sdt")
    |> Seq.find (fun control -> not (isNull (control.Element(ns + "sdtPr").Element(ns + "picture"))))
check "typing beside picture leaves picture content unchanged" false
    (pictureControl.Element(ns + "sdtContent").Descendants(ns + "t") |> Seq.exists (fun value -> value.Value.Contains "Q"))
check "typing beside picture remains visible outside control" true
    (invoiceTypedXml.Descendants(ns + "t") |> Seq.exists (fun value -> value.Value = "Q" && not (value.Ancestors(ns + "sdt") |> Seq.contains pictureControl)))
let dateParagraph =
    paragraphsOfBlocks invoiceDoc.Blocks
    |> List.find (fun para -> para.Metadata.EnclosingCellControls |> List.exists (fun layer -> layer.Xml |> Option.exists (fun xml -> xml.Contains "<w:date")))
let dateCaret = invoiceDoc |> setCursor dateParagraph.Id 2
check "date control rejects paragraph break without changing saved content"
    (dateCaret |> write false |> parts) (dateCaret |> splitBlock |> write false |> parts)

let editStory kind id document =
    let blocks = storyBody kind id document |> Option.get
    let paragraph = paragraphsOfBlocks blocks |> List.head
    let story = { document with Blocks = blocks; Selection = None } |> setCursor paragraph.Id 0 |> insertText "changed "
    withStoryBody kind id story.Blocks document
let storyEdited = doc |> editStory "footnote" "2" |> editStory "comment" "5"
let storyParts = storyEdited |> write false |> parts
check "story edit does not rewrite body" originals.["word/document.xml"] storyParts.["word/document.xml"]
check "story edit preserves custom XML" originals.["customXml/item1.xml"] storyParts.["customXml/item1.xml"]
let notesXml = XDocument.Parse storyParts.["word/footnotes.xml"]
let noteTexts = notesXml.Descendants(ns + "footnote") |> Seq.map (fun el -> el.Attribute(ns + "id").Value, text (el.ToString())) |> Map.ofSeq
check "changed note written through shared writer" "changed note" noteTexts.["2"]
check "unrelated note retained" "other note" noteTexts.["3"]
check "changed comment written through shared writer" "changed comment" (text storyParts.["word/comments.xml"])
let reopened = storyEdited |> write false |> load
check "edited note reopens" (Some "changed note") (reopened.Package.Notes |> List.tryFind (fun note -> note.Id = "2") |> Option.map (fun note -> note.Text))

// Inline control properties inherit declarations from their enclosing paragraph,
// not the part root. Model-only saves must carry that scope with the properties.
let w14 = "http://schemas.microsoft.com/office/word/2010/wordml"
let checkboxParagraph =
    "<w:p xmlns:w14=\"" + w14 + "\" xmlns:custom=\"urn:checkbox-property\">"
    + "<w:sdt><w:sdtPr><w:id w:val=\"42\"/><w14:checkbox><w14:checked w14:val=\"1\"/>"
    + "<custom:property custom:val=\"kept\"/></w14:checkbox></w:sdtPr>"
    + "<w:sdtContent><w:r><w:t>☒</w:t></w:r></w:sdtContent></w:sdt></w:p>"
let checkboxSource =
    originals
    |> Map.add "word/document.xml" (wrap "document" ("<w:body>" + checkboxParagraph
        + "<w:p><w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p>"
        + "<w:sectPr><w:headerReference w:type=\"default\" r:id=\"rHeader\"/></w:sectPr></w:body>"))
    |> Map.add "word/header1.xml" (wrap "hdr" checkboxParagraph)
    |> Map.add "word/footnotes.xml" (wrap "footnotes" ("<w:footnote w:id=\"2\">" + checkboxParagraph + "</w:footnote>"))
    |> packageBytes
let checkboxDoc = load checkboxSource
let checkboxPackage = Docx.exportTypedDocumentPackage checkboxDoc
let checkboxParts = DocumentWriter.writeTypedPackageBytes { checkboxPackage with Unchanged = false; RegenerateAll = true } |> parts
for path in [ "word/document.xml"; "word/header1.xml"; "word/footnotes.xml" ] do
    let xml = XDocument.Parse checkboxParts.[path]
    let controls = xml.Descendants(XName.Get("checkbox", w14)) |> Seq.toList
    check (path + " checkbox survives with its namespace") 1 controls.Length
    check (path + " checkbox state survives") "1" (controls.Head.Element(XName.Get("checked", w14)).Attribute(XName.Get("val", w14)).Value)
    check (path + " custom control namespace survives") "kept" (controls.Head.Element(XName.Get("property", "urn:checkbox-property")).Attribute(XName.Get("val", "urn:checkbox-property")).Value)

let fresh = empty |> setCursor "block-0" 0 |> insertText "new document"
check "new document saves and reopens without a source" "new document" (fresh |> write false |> load |> fun value -> paragraphsOfBlocks value.Blocks |> List.head |> fun p -> Rbt.Word.TextBuffer.toString p.Buffer)
let newGroup =
    { empty with Blocks = [ Paragraph { Id = "first"; Buffer = Rbt.Word.TextBuffer.ofString "one "; Metadata = defaultMetadata }
                            Paragraph { Id = "second"; Buffer = Rbt.Word.TextBuffer.ofString " two"; Metadata = defaultMetadata } ] }
    |> setSelection { BlockId = "first"; Offset = 0 } { BlockId = "second"; Offset = 4 }
    |> setContentControl "new-control" "new-tag" "New alias"
let newGroupBytes = write false newGroup
let newGroupXml = XDocument.Parse ((parts newGroupBytes).["word/document.xml"])
let newGroupControls = newGroupXml.Descendants(ns + "sdt") |> Seq.toList
check "source-free grouped control count" 1 newGroupControls.Length
check "source-free grouped control text" "one  two" (text (newGroupControls.Head.ToString()))
check "source-free grouped control reimports" [ Some "new-control"; Some "new-control" ]
    (newGroupBytes |> load |> fun doc -> paragraphsOfBlocks doc.Blocks |> List.map (fun p -> p.Metadata.ContentControlId))
check "source-free adjacent run spaces survive reimport" [ "one "; " two" ]
    (newGroupBytes |> load |> fun doc -> paragraphsOfBlocks doc.Blocks |> List.map (fun p -> Rbt.Word.TextBuffer.toString p.Buffer))
let sourceFreeTable =
    empty |> setCursor "block-0" 0 |> insertTable 1 1
    |> fun doc ->
        let table = doc.Blocks |> List.pick (function Table table -> Some table | _ -> None)
        let cell = table.Rows.Head.Cells.Head.Content.Head
        let edited = doc |> setCursor cell.Id 0 |> insertText "table cell"
        { edited with
            Blocks = edited.Blocks |> List.map (function
                | Table table -> Table { table with ContentControlId = Some "table-control"; ContentControlTag = Some "table-tag" }
                | block -> block) }
let sourceFreeTableBytes = write false sourceFreeTable
let sourceFreeTableXml = XDocument.Parse ((parts sourceFreeTableBytes).["word/document.xml"])
check "source-free table enclosed by control" 1
    (sourceFreeTableXml.Descendants(ns + "sdt") |> Seq.filter (fun sdt -> sdt.Element(ns + "sdtContent").Element(ns + "tbl") <> null) |> Seq.length)
check "source-free table control reimports" (Some "table-control")
    (sourceFreeTableBytes |> load |> fun doc -> doc.Blocks |> List.pick (function Table table -> Some table.ContentControlId | _ -> None))
let tableOnly = load sourceFreeTableBytes
let tableCell = tableOnly.Blocks |> List.pick (function Table table -> Some table.Rows.Head.Cells.Head.Content.Head | _ -> None)
check "table-only control appears in shared query" [ "table-control", tableCell.Id ]
    (contentControlInfos tableOnly |> List.map (fun info -> info.Id, info.BlockId))
let selectedTable = selectContentControl "table-tag" tableOnly
check "table-only control selects its cell text" (Some (tableCell.Id, 0, tableCell.Id, 10))
    (selectedTable.Selection |> Option.map (fun selection ->
        selection.Anchor.BlockId, selection.Anchor.Offset, selection.Focus.BlockId, selection.Focus.Offset))
check "table-only control draws a frame around its cell" 1 (controlFrameCount selectedTable)
let boundProps id =
    "<w:sdtPr><w:id w:val=\"" + id + "\"/><w:tag w:val=\"invoice\"/><w:alias w:val=\"Invoice\"/>"
    + "<w:dataBinding w:storeItemID=\"{77777777-7777-7777-7777-777777777777}\" w:xpath=\"/invoice/value\"/>"
    + "<w:placeholder><w:docPart w:val=\"DefaultPlaceholder_1\"/></w:placeholder><w:appearance w:val=\"hidden\"/>"
    + "<w:color w:val=\"ABCDEF\" ext:opaque=\"kept\"/></w:sdtPr>"
let controlledBody =
    "<w:sdt>" + boundProps "17" + "<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr><w:sdtContent>"
    + paragraph "left"
    + "<w:tbl><w:tblPr><w:tblW w:w=\"3000\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"3000\"/></w:tblGrid>"
    + "<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
    + "<w:p><w:r><w:t xml:space=\"preserve\"> right</w:t></w:r><w:r><w:t xml:space=\"preserve\"> tail </w:t></w:r></w:p>"
    + "</w:sdtContent></w:sdt>"
    + "<w:p><w:r><w:t>outside</w:t></w:r></w:p>"
    + "<w:sdt>" + boundProps "21" + "<w:sdtContent>" + paragraph "independent" + "</w:sdtContent></w:sdt>"
    + "<w:sdt><w:sdtPr><w:dataBinding w:storeItemID=\"{88888888-8888-8888-8888-888888888888}\" w:xpath=\"/invoice/optional\"/>"
    + "</w:sdtPr><w:sdtContent>" + paragraph "optional" + paragraph "more optional" + "</w:sdtContent></w:sdt>"
let controlledSource =
    originals
    |> Map.add "word/document.xml" (wrap "document" ("<w:body>" + controlledBody + "</w:body>"))
    |> packageBytes
let controlled = load controlledSource
let groupTable = controlled.Blocks |> List.pick (function Table t -> Some t | _ -> None)
check "bound table imported" (Some "17") groupTable.ContentControlId
check "table's bound properties imported" true (groupTable.ContentControlXml |> Option.exists (fun xml -> xml.Contains "dataBinding"))
let firstControlled = controlled.Blocks |> List.pick (function Paragraph p -> Some p | _ -> None)
let cellParagraph = groupTable.Rows.Head.Cells.Head.Content.Head
check "mixed paragraph/table control appears once" [ "17"; "21" ]
    (contentControlInfos controlled |> List.map (fun info -> info.Id))
let groupedParagraphs = paragraphsOfBlocks controlled.Blocks
let selectedMixed = selectContentControl "17" controlled
check "mixed control selection spans paragraph and table" (Some (firstControlled.Id, groupedParagraphs.[2].Id, 12))
    (selectedMixed.Selection |> Option.map (fun selection ->
        selection.Anchor.BlockId, selection.Focus.BlockId, selection.Focus.Offset))
check "mixed control frame includes its table cell" 3 (controlFrameCount selectedMixed)
let controlledEdit =
    controlled |> setCursor firstControlled.Id 2 |> insertText "X"
    |> setCursor cellParagraph.Id 0 |> insertText "new "
let assertControlledSave label bytes =
    let xml = (parts bytes).["word/document.xml"]
    check (label + " custom XML part kept") originals.["customXml/item1.xml"] (parts bytes).["customXml/item1.xml"]
    let document = XDocument.Parse xml
    let controls = document.Descendants(ns + "sdt") |> Seq.toList
    check (label + " control count") 3 controls.Length
    let group = controls.Head
    let content = group.Element(ns + "sdtContent")
    check (label + " grouped block kinds") [ "p"; "tbl"; "p" ] (content.Elements() |> Seq.map (fun node -> node.Name.LocalName) |> Seq.toList)
    check (label + " bound grouped table text") "new cell" (text (content.Element(ns + "tbl").ToString()))
    check (label + " complete grouped text") "leXftnew cell right tail " (text (content.ToString()))
    let spaced = content.Descendants(ns + "t") |> Seq.filter (fun node -> node.Value.StartsWith(" ") || node.Value.EndsWith(" ")) |> Seq.toList
    check (label + " spaces between runs") 2 spaced.Length
    check (label + " spaced runs preserve XML whitespace") true (spaced |> List.forall (fun node -> node.Attribute(XNamespace.Xml + "space").Value = "preserve"))
    check (label + " binding survives") "/invoice/value" (group.Descendants(ns + "dataBinding") |> Seq.head |> fun e -> e.Attribute(ns + "xpath").Value)
    check (label + " unknown control property survives") "kept" (group.Descendants(ns + "color") |> Seq.head |> fun e -> e.Attribute(XName.Get("opaque", "urn:kept")).Value)
    check (label + " end properties survive") 1 (group.Elements(ns + "sdtEndPr") |> Seq.length)
    check (label + " id-less binding survives") "/invoice/optional" (controls.[2].Descendants(ns + "dataBinding") |> Seq.head |> fun e -> e.Attribute(ns + "xpath").Value)
    check (label + " id-less grouped paragraphs") [ "optional"; "more optional" ]
        (controls.[2].Element(ns + "sdtContent").Elements(ns + "p") |> Seq.map (fun p -> text (p.ToString())) |> Seq.toList)
    let reopened = load bytes
    check (label + " table binding reimports") (Some "17") (reopened.Blocks |> List.pick (function Table t -> Some t.ContentControlId | _ -> None))
    check (label + " edited grouped text reimports") "leXftnew cell right tail "
        (reopened.Blocks |> List.collect (function Paragraph p -> [Rbt.Word.TextBuffer.toString p.Buffer] | Table t -> [Rbt.Word.TextBuffer.toString t.Rows.Head.Cells.Head.Content.Head.Buffer] | _ -> [])
         |> List.take 3 |> String.concat "")
assertControlledSave "source-backed" (write false controlledEdit)
let regenerated = Docx.exportTypedDocumentPackage controlledEdit
assertControlledSave "model-written" (DocumentWriter.writeTypedPackageBytes { regenerated with Unchanged = false; RegenerateAll = true })
let changedGroup = updateContentControl "17" "19" "edited-invoice" "Edited invoice" controlledEdit
let changedGroupPackage = Docx.exportTypedDocumentPackage changedGroup
let changedGroupBytes = DocumentWriter.writeTypedPackageBytes { changedGroupPackage with Unchanged = false; RegenerateAll = true }
let changedGroupXml = XDocument.Parse ((parts changedGroupBytes).["word/document.xml"])
let changedGroupControl = changedGroupXml.Descendants(ns + "sdt") |> Seq.head
check "updated grouped control stays grouped" [ "p"; "tbl"; "p" ]
    (changedGroupControl.Element(ns + "sdtContent").Elements() |> Seq.map (fun node -> node.Name.LocalName) |> Seq.toList)
check "updated grouped control retains binding" "/invoice/value"
    (changedGroupControl.Descendants(ns + "dataBinding") |> Seq.head |> fun node -> node.Attribute(ns + "xpath").Value)
check "updated grouped table reimports ID" (Some "19")
    (changedGroupBytes |> load |> fun doc -> doc.Blocks |> List.pick (function Table table -> Some table.ContentControlId | _ -> None))
let changedSourceXml = XDocument.Parse ((parts (write false changedGroup)).["word/document.xml"])
let changedSourceControl = changedSourceXml.Descendants(ns + "sdt") |> Seq.head
check "source-backed grouped rename reaches XML" "19"
    (changedSourceControl.Descendants(ns + "id") |> Seq.head |> fun node -> node.Attribute(ns + "val").Value)
check "source-backed grouped rename keeps table" [ "p"; "tbl"; "p" ]
    (changedSourceControl.Element(ns + "sdtContent").Elements() |> Seq.map (fun node -> node.Name.LocalName) |> Seq.toList)
let removedGroup = removeContentControl "17" controlled
check "removing grouped control clears table binding" None
    (removedGroup.Blocks |> List.pick (function Table table -> Some table.ContentControlXml | _ -> None))
let removedGroupPackage = Docx.exportTypedDocumentPackage removedGroup
let removedGroupXml = XDocument.Parse ((parts (DocumentWriter.writeTypedPackageBytes { removedGroupPackage with Unchanged = false; RegenerateAll = true })).["word/document.xml"])
check "removing grouped control unwraps its table" 0
    (removedGroupXml.Descendants(ns + "sdt") |> Seq.filter (fun control -> control.Descendants(ns + "tbl") |> Seq.isEmpty |> not) |> Seq.length)
let removedSourceXml = XDocument.Parse ((parts (write false removedGroup)).["word/document.xml"])
check "source-backed control removal unwraps table" 0
    (removedSourceXml.Descendants(ns + "sdt") |> Seq.filter (fun control -> control.Descendants(ns + "tbl") |> Seq.isEmpty |> not) |> Seq.length)
let renamed = updateContentControl "21" "22" "new-tag" "New alias" controlled
let renamedPackage = Docx.exportTypedDocumentPackage renamed
let renamedBytes = DocumentWriter.writeTypedPackageBytes { renamedPackage with Unchanged = false; RegenerateAll = true }
let renamedXml = XDocument.Parse ((parts renamedBytes).["word/document.xml"])
let renamedControl = renamedXml.Descendants(ns + "sdt") |> Seq.find (fun sdt -> text (sdt.ToString()) = "independent")
check "renamed bound control ID" "22" (renamedControl.Descendants(ns + "id") |> Seq.head |> fun node -> node.Attribute(ns + "val").Value)
check "renamed bound control tag" "new-tag" (renamedControl.Descendants(ns + "tag") |> Seq.head |> fun node -> node.Attribute(ns + "val").Value)
check "renamed control keeps binding" "/invoice/value" (renamedControl.Descendants(ns + "dataBinding") |> Seq.head |> fun node -> node.Attribute(ns + "xpath").Value)
check "renamed control reimports binding" true (renamedBytes |> load |> fun doc -> doc.Blocks |> List.exists (function Paragraph p -> p.Metadata.ContentControlId = Some "22" && (p.Metadata.ContentControlXml |> Option.exists (fun xml -> xml.Contains "dataBinding")) | _ -> false))
let removed = removeContentControl "21" controlled
let removedPackage = Docx.exportTypedDocumentPackage removed
let removedXml = XDocument.Parse ((parts (DocumentWriter.writeTypedPackageBytes { removedPackage with Unchanged = false; RegenerateAll = true })).["word/document.xml"])
check "removed bound control does not leave binding" false (removedXml.Descendants(ns + "sdt") |> Seq.exists (fun sdt -> text (sdt.ToString()) = "independent"))
let modelSave doc =
    let package = Docx.exportTypedDocumentPackage doc
    DocumentWriter.writeTypedPackageBytes { package with Unchanged = false; RegenerateAll = true }
let pairedProperties =
    "<w:sdtPr><w:dataBinding w:storeItemID=\"{88888888-8888-8888-8888-888888888888}\""
    + " w:xpath=\"/invoice/shared\"/><w:color ext:opaque=\"paired\"/></w:sdtPr>"
let pairedControl content = "<w:sdt>" + pairedProperties + "<w:sdtContent>" + content + "</w:sdtContent></w:sdt>"
let pairedSource =
    originals
    |> Map.add "word/document.xml" (wrap "document" ("<w:body>" +
        pairedControl (paragraph "first" + paragraph "second") +
        pairedControl (paragraph "third") + "</w:body>"))
    |> packageBytes
let paired = load pairedSource
let pairedParagraphs = paragraphsOfBlocks paired.Blocks
check "id-less bound groups are not exposed as editable IDs" [] (contentControlInfos paired |> List.map (fun info -> info.Id))
check "distinct id-less groups imported" true
    (pairedParagraphs.[0].Metadata.ContentControlGroup = pairedParagraphs.[1].Metadata.ContentControlGroup &&
     pairedParagraphs.[1].Metadata.ContentControlGroup <> pairedParagraphs.[2].Metadata.ContentControlGroup)
let editedPair = paired |> setCursor pairedParagraphs.[1].Id 2 |> insertText "X"
let assertPaired label bytes =
    let xml = XDocument.Parse ((parts bytes).["word/document.xml"])
    let groups = xml.Descendants(ns + "sdt") |> Seq.toList
    check (label + " keeps distinct equal-property controls") 2 groups.Length
    check (label + " retains block grouping") [ [ "first"; "seXcond" ]; [ "third" ] ]
        (groups |> List.map (fun control ->
            control.Element(ns + "sdtContent").Elements(ns + "p") |> Seq.map (fun p -> text (p.ToString())) |> Seq.toList))
    check (label + " retains both bindings") [ "/invoice/shared"; "/invoice/shared" ]
        (groups |> List.map (fun control -> control.Descendants(ns + "dataBinding") |> Seq.head |> fun node -> node.Attribute(ns + "xpath").Value))
    let again = bytes |> load |> fun doc -> paragraphsOfBlocks doc.Blocks
    check (label + " distinct groups reimport") true
        (again.[0].Metadata.ContentControlGroup = again.[1].Metadata.ContentControlGroup &&
         again.[1].Metadata.ContentControlGroup <> again.[2].Metadata.ContentControlGroup)
assertPaired "source-backed" (write false editedPair)
assertPaired "model-written" (modelSave editedPair)

let mutable copySerial = 0
let copiedControls =
    paired.Blocks |> copyBlocksWithNewIds (fun kind ->
        copySerial <- copySerial + 1
        sprintf "copied-%s-%d" kind copySerial)
let copiedPair = { paired with Blocks = paired.Blocks @ copiedControls }
let copiedPairXml = XDocument.Parse ((parts (modelSave copiedPair)).["word/document.xml"])
check "copying adjacent id-less groups keeps each distinct" [ 2; 1; 2; 1 ]
    (copiedPairXml.Descendants(ns + "sdt")
     |> Seq.map (fun sdt -> sdt.Element(ns + "sdtContent").Elements(ns + "p") |> Seq.length)
     |> Seq.toList)

let nestedOuter =
    "<w:sdtPr><w:id w:val=\"31\"/><w:tag w:val=\"outer\"/>"
    + "<w:dataBinding w:xpath=\"/invoice/outer\"/><w:color ext:opaque=\"outer\"/></w:sdtPr>"
let nestedInner =
    "<w:sdtPr><w:id w:val=\"32\"/><w:tag w:val=\"inner\"/>"
    + "<w:dataBinding w:xpath=\"/invoice/inner\"/><w:color ext:opaque=\"inner\"/></w:sdtPr>"
let nestedSource =
    originals
    |> Map.add "word/document.xml" (wrap "document" ("<w:body><w:sdt>" + nestedOuter +
        "<w:sdtContent>" + paragraph "outside-left" +
        "<w:sdt>" + nestedInner + "<w:sdtContent>" + paragraph "inside-one" + paragraph "inside-two" +
        "</w:sdtContent></w:sdt>" + paragraph "outside-right" +
        "</w:sdtContent></w:sdt></w:body>"))
    |> packageBytes
let nested = load nestedSource
let nestedParagraphs = paragraphsOfBlocks nested.Blocks
check "nested outer metadata imported" (Some "31")
    (nestedParagraphs.[1].Metadata.ContentControlAncestors |> List.tryHead |> Option.bind (fun layer -> layer.Id))
let editedNested = nested |> setCursor nestedParagraphs.[1].Id 3 |> insertText "X"
let assertNested label bytes =
    let xml = XDocument.Parse ((parts bytes).["word/document.xml"])
    let outer = xml.Descendants(ns + "sdt") |> Seq.head
    let inner = outer.Element(ns + "sdtContent").Element(ns + "sdt")
    check (label + " outer blocks retain inner control") [ "p"; "sdt"; "p" ]
        (outer.Element(ns + "sdtContent").Elements() |> Seq.map (fun el -> el.Name.LocalName) |> Seq.toList)
    check (label + " inner group text") [ "insXide-one"; "inside-two" ]
        (inner.Element(ns + "sdtContent").Elements(ns + "p") |> Seq.map (fun p -> text (p.ToString())) |> Seq.toList)
    check (label + " both bound layers") [ "/invoice/outer"; "/invoice/inner" ]
        ([ outer; inner ] |> List.map (fun layer ->
            layer.Element(ns + "sdtPr").Element(ns + "dataBinding").Attribute(ns + "xpath").Value))
    check (label + " unknown properties in both layers") [ "outer"; "inner" ]
        ([ outer; inner ] |> List.map (fun layer ->
            layer.Element(ns + "sdtPr").Element(ns + "color").Attribute(XName.Get("opaque", "urn:kept")).Value))
    let reopened = load bytes
    let imported = paragraphsOfBlocks reopened.Blocks
    check (label + " reimports inner and outer IDs") (Some "32", Some "31")
        (imported.[1].Metadata.ContentControlId,
         imported.[1].Metadata.ContentControlAncestors |> List.tryHead |> Option.bind (fun layer -> layer.Id))
    check (label + " panel lists outer and inner once") [ "31"; "32" ]
        (contentControlInfos reopened |> List.map (fun info -> info.Id))
    let selectedOuter = selectContentControl "31" reopened
    check (label + " outer selection includes nested paragraphs") (Some (imported.Head.Id, imported.[3].Id, 13))
        (selectedOuter.Selection |> Option.map (fun selection ->
            selection.Anchor.BlockId, selection.Focus.BlockId, selection.Focus.Offset))
    check (label + " outer control frames nested paragraphs") 4 (controlFrameCount selectedOuter)
    let selectedInner = selectContentControl "inner" reopened
    check (label + " inner selection spans both paragraphs") (Some (imported.[1].Id, imported.[2].Id, 10))
        (selectedInner.Selection |> Option.map (fun selection ->
            selection.Anchor.BlockId, selection.Focus.BlockId, selection.Focus.Offset))
assertNested "source-backed" (write false editedNested)
assertNested "model-written" (modelSave editedNested)
let updatedOuter = updateContentControl "31" "41" "changed-outer" "Changed outer" editedNested
let updatedOuterXml = XDocument.Parse ((parts (modelSave updatedOuter)).["word/document.xml"])
check "outer rename retains inner and outer binding" [ "41"; "32" ]
    (updatedOuterXml.Descendants(ns + "sdt") |> Seq.map (fun sdt -> sdt.Element(ns + "sdtPr").Element(ns + "id").Attribute(ns + "val").Value) |> Seq.toList)
let updatedOuterSourceXml = XDocument.Parse ((parts (write false updatedOuter)).["word/document.xml"])
check "source-backed outer rename retains inner" [ "41"; "32" ]
    (updatedOuterSourceXml.Descendants(ns + "sdt") |> Seq.map (fun sdt -> sdt.Element(ns + "sdtPr").Element(ns + "id").Attribute(ns + "val").Value) |> Seq.toList)
let removedOuter = removeContentControl "31" editedNested
let removedOuterXml = XDocument.Parse ((parts (modelSave removedOuter)).["word/document.xml"])
check "outer removal retains inner group" [ "insXide-one"; "inside-two" ]
    (removedOuterXml.Descendants(ns + "sdt") |> Seq.collect (fun sdt -> sdt.Element(ns + "sdtContent").Elements(ns + "p")) |> Seq.map (fun p -> text (p.ToString())) |> Seq.toList)
let removedOuterSourceXml = XDocument.Parse ((parts (write false removedOuter)).["word/document.xml"])
check "source-backed outer removal retains inner" 1 (removedOuterSourceXml.Descendants(ns + "sdt") |> Seq.length)
let updatedInner = updateContentControl "32" "42" "changed-inner" "Changed inner" editedNested
let updatedInnerXml = XDocument.Parse ((parts (modelSave updatedInner)).["word/document.xml"])
check "inner rename retains outer control" [ "31"; "42" ]
    (updatedInnerXml.Descendants(ns + "sdt") |> Seq.map (fun sdt -> sdt.Element(ns + "sdtPr").Element(ns + "id").Attribute(ns + "val").Value) |> Seq.toList)
let updatedInnerSourceXml = XDocument.Parse ((parts (write false updatedInner)).["word/document.xml"])
check "source-backed inner rename retains outer control" [ "31"; "42" ]
    (updatedInnerSourceXml.Descendants(ns + "sdt") |> Seq.map (fun sdt -> sdt.Element(ns + "sdtPr").Element(ns + "id").Attribute(ns + "val").Value) |> Seq.toList)
let removedInner = removeContentControl "32" editedNested
let removedInnerXml = XDocument.Parse ((parts (modelSave removedInner)).["word/document.xml"])
check "inner removal retains outer bound group" 1 (removedInnerXml.Descendants(ns + "sdt") |> Seq.length)
let removedInnerSourceXml = XDocument.Parse ((parts (write false removedInner)).["word/document.xml"])
check "source-backed inner removal retains outer" 1 (removedInnerSourceXml.Descendants(ns + "sdt") |> Seq.length)
let repeatedInner i =
    let properties =
        sprintf "<w:sdtPr><w:id w:val=\"%d\"/><w:tag w:val=\"inner\"/><w:dataBinding w:xpath=\"/invoice/inner\"/></w:sdtPr>" (100 + i)
    "<w:sdt>" + properties + "<w:sdtContent>" + paragraph (sprintf "inside-%d" i) + "</w:sdtContent></w:sdt>"
let repeatedContent =
    [ for i in 0 .. 39 -> paragraph (sprintf "lead-%d" i) + repeatedInner i ] |> String.concat ""
let repeatedSource =
    originals
    |> Map.add "word/document.xml" (wrap "document" ("<w:body><w:sdt>" + nestedOuter +
        "<w:sdtContent>" + repeatedContent + "</w:sdtContent></w:sdt></w:body>"))
    |> packageBytes
let repeated = load repeatedSource
let lastRepeated = repeated.Blocks |> paragraphsOfBlocks |> List.last
let repeatedEdited = repeated |> setCursor lastRepeated.Id (Rbt.Word.TextBuffer.length lastRepeated.Buffer) |> insertText "!"
let assertRepeated label bytes =
    let xml = XDocument.Parse ((parts bytes).["word/document.xml"])
    let outer = xml.Descendants(ns + "sdt") |> Seq.head
    let children = outer.Element(ns + "sdtContent").Elements() |> Seq.toList
    check (label + " 40 bounded nested transitions") 80 children.Length
    check (label + " alternates outer paragraphs and inner controls") true
        (children |> List.mapi (fun i child ->
            let expected = if i % 2 = 0 then "p" else "sdt"
            child.Name.LocalName = expected) |> List.forall id)
    check (label + " retains last edited block") "inside-39!" (text (children.[79].ToString()))
    check (label + " reimports all nested paragraphs") 80 (bytes |> load |> fun doc -> paragraphsOfBlocks doc.Blocks |> List.length)
assertRepeated "source-backed" (write false repeatedEdited)
assertRepeated "model-written" (modelSave repeatedEdited)
// A producer may choose another prefix for the Word namespace. Property
// lookup and source-preserving edits must not depend on the spelling "w".
let canonicalAliasEntries =
    parts nestedSource
    |> Map.change "word/document.xml" (Option.map (fun xml ->
        xml.Replace("<w:p>", "<w:p><w:pPr><w:jc w:val=\"right\"/><w:spacing w:after=\"120\"/></w:pPr>")))
let canonicalAlias = canonicalAliasEntries |> packageBytes |> load
let aliasedEntries =
    canonicalAliasEntries
    |> Map.change "word/document.xml" (Option.map (fun xml ->
        xml.Replace("xmlns:w=", "xmlns:ns0=").Replace("w:", "ns0:")
           .Replace("<ns0:id ns0:val=\"31\"/>", "<ns0:id ns0:val=\"31\" ext:val=\"opaque\"/>")))
let aliased = aliasedEntries |> packageBytes |> load
let aliasedParagraphs = paragraphsOfBlocks aliased.Blocks
check "aliased control IDs import" [ "31"; "32" ] (contentControlInfos aliased |> List.map (fun info -> info.Id))
let paragraphProperties doc =
    paragraphsOfBlocks doc.Blocks |> List.map (fun para -> para.Metadata.Alignment, para.Metadata.SpacingAfter)
check "aliased paragraph properties match canonical properties" (paragraphProperties canonicalAlias) (paragraphProperties aliased)
let aliasedEdited = aliased |> setCursor aliasedParagraphs.[1].Id 3 |> insertText "X"
assertNested "aliased source-backed" (write false aliasedEdited)
assertNested "aliased model-written" (modelSave aliasedEdited)
let aliasedRenamed = updateContentControl "31" "41" "renamed" "Renamed outer" aliasedEdited
for label, save in [ "source-backed", write false; "model-written", modelSave ] do
    let bytes = save aliasedRenamed
    let xml = XDocument.Parse ((parts bytes).["word/document.xml"])
    let properties = xml.Descendants(ns + "sdtPr") |> Seq.head
    check (label + " aliased rename updates the existing ID") "41" (properties.Element(ns + "id").Attribute(ns + "val").Value)
    check (label + " aliased rename keeps foreign attributes") "opaque" (properties.Element(ns + "id").Attribute(XName.Get("val", "urn:kept")).Value)
    check (label + " aliased rename adds the alias") "Renamed outer" (properties.Element(ns + "alias").Attribute(ns + "val").Value)
    check (label + " aliased rename retains binding") "/invoice/outer" (properties.Element(ns + "dataBinding").Attribute(ns + "xpath").Value)
    check (label + " aliased rename reimports") [ "41"; "32" ] (bytes |> load |> contentControlInfos |> List.map (fun info -> info.Id))
// Capture packages once: concurrent saves must neither share source state nor mutate
// the paragraph objects that a later save of the same package will use.
let otherEntries =
    originals
    |> Map.add "word/document.xml" (wrap "document" ("<w:body>" + controlledBody.Replace("/invoice/value", "/other/value").Replace("left", "other") + "</w:body>"))
    |> Map.add "word/_rels/document.xml.rels" (relationships (rel "rOther" "customXml" "../customXml/item1.xml"))
    |> Map.add "customXml/item1.xml" "<kept xmlns=\"urn:opaque\" value=\"other\">different</kept>"
let other = otherEntries |> packageBytes |> load
let otherFirst = paragraphsOfBlocks other.Blocks |> List.head
let otherEdit = other |> setCursor otherFirst.Id 2 |> insertText "Y"
let captured = { (Docx.exportTypedDocumentPackage controlledEdit) with Unchanged = false }
let capturedOther = { (Docx.exportTypedDocumentPackage otherEdit) with Unchanged = false }
let reference = DocumentWriter.writeTypedPackageBytes captured |> parts
let referenceOther = DocumentWriter.writeTypedPackageBytes capturedOther |> parts
let saves =
    [| for i in 0 .. 11 ->
        let package = if i % 2 = 0 then captured else capturedOther
        System.Threading.Tasks.Task.Run(fun () -> DocumentWriter.writeTypedPackageBytes package) |]
System.Threading.Tasks.Task.WaitAll(saves |> Array.map (fun task -> task :> System.Threading.Tasks.Task))
for i in 0 .. saves.Length - 1 do
    let savedParts = parts saves.[i].Result
    let expected = if i % 2 = 0 then reference else referenceOther
    check (sprintf "concurrent save %d preserves all package parts" i) expected savedParts
    let document = XDocument.Parse savedParts.["word/document.xml"]
    let binding = document.Descendants(ns + "dataBinding") |> Seq.head
    check (sprintf "concurrent save %d keeps its binding" i)
        (if i % 2 = 0 then "/invoice/value" else "/other/value")
        (binding.Attribute(ns + "xpath").Value)
    check (sprintf "concurrent save %d keeps its relationship" i) true
        (savedParts.["word/_rels/document.xml.rels"].Contains(if i % 2 = 0 then "rCustom" else "rOther"))
    check (sprintf "concurrent save %d keeps its edit" i) true
        (text savedParts.["word/document.xml"] |> fun value -> value.Contains(if i % 2 = 0 then "leXft" else "otYher"))
check "repeated captured save preserves all parts" reference (DocumentWriter.writeTypedPackageBytes captured |> parts)
check "repeated other captured save preserves all parts" referenceOther (DocumentWriter.writeTypedPackageBytes capturedOther |> parts)

printfn "Shared Word package IO checks: %d passed" checks
