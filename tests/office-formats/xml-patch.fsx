// XmlPatch (docs/save-architecture.md): typed edits of XML that was read, made in its text - what no edit
// names stays the bytes it was; and the edits that make one element as the model wrote it into another.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

open XmlPatch

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (want: 'a) =
    if got = want then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %s\n  got  %A\n  want %A" label got want

let el text = (read text).Value
let order = Word.order
let edit text edits = apply order (el text) edits
let mergeOf empty source was now = merge order empty source was now

// ---- apply: what no edit names stays as it was ----
let pPr = """<w:pPr>
  <w:pStyle w:val="Body"/>
  <w:spacing w:before="240" w:beforeAutospacing="1"   w:after='120'/>
  <x99:unknown a="b"/>
  <w:jc w:val="left"/>
  <w:pPrChange w:id="1" w:author="A"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>
</w:pPr>"""
check "no edits: the bytes" (edit pPr []) pPr
check "an attribute set in place: its quotes, its neighbours, the spacing between them stay"
    (edit pPr [ EditChild ("w:spacing", [ SetAttribute ("w:before", "360") ], None) ])
    (pPr.Replace("w:before=\"240\"", "w:before=\"360\""))
check "an attribute in single quotes set in place" (edit pPr [ EditChild ("w:spacing", [ SetAttribute ("w:after", "0") ], None) ]) (pPr.Replace("w:after='120'", "w:after='0'"))
check "an attribute taken out with the space before it"
    (edit pPr [ EditChild ("w:spacing", [ RemoveAttribute "w:beforeAutospacing" ], None) ])
    (pPr.Replace(" w:beforeAutospacing=\"1\"", ""))
check "an attribute put in goes last, before the tag's end"
    (edit "<w:jc w:val=\"left\"/>" [ SetAttribute ("w:other", "a&b\"") ])
    "<w:jc w:val=\"left\" w:other=\"a&amp;b&quot;\"/>"
check "a child set in place" (edit pPr [ SetChild ("w:jc", "<w:jc w:val=\"center\"/>") ]) (pPr.Replace("<w:jc w:val=\"left\"/>", "<w:jc w:val=\"center\"/>"))
check "a child taken out with the white space before it" (edit pPr [ RemoveChild "w:pStyle" ]) (pPr.Replace("\n  <w:pStyle w:val=\"Body\"/>", ""))
check "a child put in at its schema place: after spacing (and the unknown element after it), before jc"
    (edit pPr [ SetChild ("w:ind", "<w:ind w:left=\"720\"/>") ])
    (pPr.Replace("<w:jc w:val=\"left\"/>", "<w:ind w:left=\"720\"/><w:jc w:val=\"left\"/>"))
check "a child put in before the tracked change, which stays"
    (edit "<w:pPr><w:pPrChange w:id=\"1\"><w:pPr/></w:pPrChange></w:pPr>" [ SetChild ("w:jc", "<w:jc w:val=\"center\"/>") ])
    "<w:pPr><w:jc w:val=\"center\"/><w:pPrChange w:id=\"1\"><w:pPr/></w:pPrChange></w:pPr>"
check "several put in at one place go in schema order"
    (edit "<w:rPr/>" [ SetChild ("w:sz", "<w:sz w:val=\"24\"/>"); SetChild ("w:b", "<w:b/>") ])
    "<w:rPr><w:b/><w:sz w:val=\"24\"/></w:rPr>"
check "an empty element stays closed in its tag when nothing is put in" (edit "<w:rPr/>" [ RemoveChild "w:b" ]) "<w:rPr/>"
check "an edit of a child there is none of does nothing" (edit pPr [ EditChild ("w:ind", [ SetAttribute ("w:left", "1") ], None) ]) pPr
check "an edit of a child there is none of puts in the whole given"
    (edit "<w:rPr/>" [ EditChild ("w:color", [ SetAttribute ("w:val", "C00000") ], Some "<w:color w:val=\"C00000\"/>") ])
    "<w:rPr><w:color w:val=\"C00000\"/></w:rPr>"
check "merge: a property the model wrote of the style before and sets now, the XML not stating it: put in"
    (mergeOf "<w:rPr/>" "<w:rPr><w:b/></w:rPr>" "<w:rPr><w:color w:val=\"17365D\"/></w:rPr>" "<w:rPr><w:color w:val=\"C00000\"/></w:rPr>")
    (Some "<w:rPr><w:b/><w:color w:val=\"C00000\"/></w:rPr>")
check "children of one name set: the first's place, the others taken out"
    (edit "<w:tabs><w:tab w:val=\"left\" w:pos=\"1\"/> <w:tab w:val=\"left\" w:pos=\"2\"/></w:tabs>" [ SetChildren ("w:tab", [ "<w:tab w:val=\"right\" w:pos=\"9\"/>" ]) ])
    "<w:tabs><w:tab w:val=\"right\" w:pos=\"9\"/></w:tabs>"
check "a start tag's attributes edited while a namespace declaration stays"
    (edit "<w:tr xmlns:w14=\"urn:x\" w:rsidR=\"00A\" w14:paraId=\"1A2B\" w14:textId=\"77\"><w:tc/></w:tr>" [ RemoveAttribute "w14:paraId"; RemoveAttribute "w14:textId" ])
    "<w:tr xmlns:w14=\"urn:x\" w:rsidR=\"00A\"><w:tc/></w:tr>"

let row = "<w:tr w:rsidR=\"00A\" w14:paraId=\"1A\"><w:tc/></w:tr>"
let rowNode = (el row).Node
check "startTag: an attribute taken out" (startTag row rowNode [ RemoveAttribute "w14:paraId" ]) "<w:tr w:rsidR=\"00A\">"
check "startTag: an attribute put in" (startTag row rowNode [ SetAttribute ("w:x", "1") ]) "<w:tr w:rsidR=\"00A\" w14:paraId=\"1A\" w:x=\"1\">"
check "startTag: no edits" (startTag row rowNode []) "<w:tr w:rsidR=\"00A\" w14:paraId=\"1A\">"

// ---- layout: items among their wrappers ----
let tableText = "<w:tbl>\n <w:tblPr/>\n <w:tr><w:tc/></w:tr>\n <w:bookmarkStart w:id=\"1\"/>\n <w:sdt><w:sdtPr><w:tag w:val=\"r\"/></w:sdtPr><w:sdtContent>\n  <w:tr><w:tc/></w:tr>\n </w:sdtContent></w:sdt>\n</w:tbl>"
let tableNode = (el tableText).Node
let propsEnd = match tableNode with XmlDsl.Element (_, _, kids, _) -> kids |> List.pick (function XmlDsl.Element (n, _, _, sp) when n = "w:tblPr" -> Some sp.End | _ -> None) | _ -> 0
let laid = (layout (fun name -> name = "tr") (fun _ -> false) (fun _ -> false) ignore tableText tableNode propsEnd).Value
let rowStarts = laid.Children |> List.choose (function LaidItem (at, _) -> Some at | _ -> None)
let content = tableText.Substring(propsEnd, tableText.LastIndexOf "</w:tbl>" - propsEnd)
check "layout: the items" rowStarts.Length 2
check "layout: every item as it was, the contents as they were" (writeLayout laid (rowStarts |> List.map (fun at -> AtItem at, "<w:tr><w:tc/></w:tr>") |> List.toArray)) content
check "layout: an item taken out, the rest as it was - its wrapper, around nothing then, with it (as Word takes out a control whose last row goes)"
    (writeLayout laid [| AtItem rowStarts.[0], "<w:tr><w:tc/></w:tr>" |])
    "\n <w:tr><w:tc/></w:tr>\n <w:bookmarkStart w:id=\"1\"/>\n"
check "layout: a new item after the second, in its wrapper"
    (writeLayout laid [| AtItem rowStarts.[0], "<w:tr><w:tc/></w:tr>"; AtItem rowStarts.[1], "<w:tr><w:tc/></w:tr>"; WithPrevious, "<w:tr>new</w:tr>" |])
    (content.Replace("<w:tr><w:tc/></w:tr>\n </w:sdtContent>", "<w:tr><w:tc/></w:tr><w:tr>new</w:tr>\n </w:sdtContent>"))

// A paragraph: runs, some in a link the model writes itself (pass-through), a bookmark it writes itself (skipped).
let paraText = "<w:p><w:pPr/><w:r><w:t>a</w:t></w:r><w:ins w:id=\"1\"><w:r><w:t>b</w:t></w:r></w:ins><w:bookmarkStart w:id=\"2\"/><w:proofErr/><w:r><w:t>c</w:t></w:r></w:p>"
let paraNode = (el paraText).Node
let pPrEnd = paraText.IndexOf "<w:r>"
let paraLaid = (layout (fun name -> name = "r") (fun name -> name = "ins") (fun node -> match node with XmlDsl.Element (name, _, _, _) -> name = "w:bookmarkStart" | _ -> false) ignore paraText paraNode pPrEnd).Value
let runStarts = paraLaid.Children |> List.choose (function LaidItem (at, _) -> Some at | _ -> None)
check "layout: the runs, also the one in the pass-through" runStarts.Length 3
check "layout: the pass-through's tags are not in what is written, the skipped mark is not, the kept one is"
    (writeLayout paraLaid (runStarts |> List.mapi (fun i at -> AtItem at, sprintf "<R%d/>" i) |> List.toArray))
    "<R0/><R1/><w:proofErr/><R2/>"
check "layout: an item outside the wrappers closes them"
    (writeLayout laid [| AtItem rowStarts.[1], "<R/>"; Outside, "<O/>" |])
    "\n <w:bookmarkStart w:id=\"1\"/>\n <w:sdt><w:sdtPr><w:tag w:val=\"r\"/></w:sdtPr><w:sdtContent>\n  <R/>\n </w:sdtContent></w:sdt><O/>\n"
check "layout: a part of the paragraph without what is kept before its first run"
    (writeLayoutWith { KeptBefore = false; KeptAfter = true; Covered = [] } paraLaid [| AtItem runStarts.[2], "<R2/>" |])
    "<R2/>"

check "a run's properties put in go first, before what it holds" (edit "<w:r w:x=\"1\"><w:t>a</w:t></w:r>" [ SetChild ("w:rPr", "<w:rPr><w:b/></w:rPr>") ]) "<w:r w:x=\"1\"><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r>"
check "SetText: the text set, the start tag as it was" (edit "<w:t>edit me</w:t>" [ SetText "Qedit me" ]) "<w:t>Qedit me</w:t>"
check "SetText: escaped, and xml:space put in where white space at an end needs it" (edit "<w:t>a</w:t>" [ SetText " a&b" ]) "<w:t xml:space=\"preserve\"> a&amp;b</w:t>"
check "diff of a run: its text as SetText in its w:t (the model's xml:space, the same before and after, not touched)"
    (diff (el "<w:r><w:t xml:space=\"preserve\">edit me</w:t></w:r>") (el "<w:r><w:t xml:space=\"preserve\">Qedit me</w:t></w:r>"))
    [ EditChild ("w:t", [ SetText "Qedit me" ], Some "<w:t xml:space=\"preserve\">Qedit me</w:t>") ]
check "a source run with the model's change of text made in it"
    (edit "<w:r w:rsidRPr=\"001\"><w:t>edit me</w:t></w:r>" (diff (el "<w:r><w:t xml:space=\"preserve\">edit me</w:t></w:r>") (el "<w:r><w:t xml:space=\"preserve\">Qedit me</w:t></w:r>")))
    "<w:r w:rsidRPr=\"001\"><w:t>Qedit me</w:t></w:r>"

// ---- diff: the model's change as edits ----
check "diff: an attribute changed, one gone, a child put in, one gone, one changed"
    (diff (el "<w:pPr><w:spacing w:before=\"1\" w:after=\"2\"/><w:jc w:val=\"left\"/></w:pPr>") (el "<w:pPr><w:spacing w:before=\"3\"/><w:ind w:left=\"5\"/></w:pPr>"))
    [ EditChild ("w:spacing", [ SetAttribute ("w:before", "3"); RemoveAttribute "w:after" ], Some "<w:spacing w:before=\"3\"/>"); RemoveChild "w:jc"; SetChild ("w:ind", "<w:ind w:left=\"5\"/>") ]
check "diff: nothing changed, no edits" (diff (el "<w:rPr><w:b/></w:rPr>") (el "<w:rPr><w:b/></w:rPr>")) []

// ---- merge: the model's change made in the element as it was read ----
check "merge: what the model does not write stays (an unknown element, a tracked change, an attribute it does not hold)"
    (mergeOf "<w:pPr/>" pPr "<w:pPr><w:spacing w:before=\"240\" w:after=\"120\"/><w:jc w:val=\"left\"/></w:pPr>" "<w:pPr><w:spacing w:before=\"240\" w:after=\"120\"/><w:jc w:val=\"center\"/></w:pPr>")
    (Some (pPr.Replace("<w:jc w:val=\"left\"/>", "<w:jc w:val=\"center\"/>")))
check "merge: none read - the change made in an empty element"
    (mergeOf "<w:rPr/>" "" "<w:rPr/>" "<w:rPr><w:i/></w:rPr>") (Some "<w:rPr><w:i/></w:rPr>")
check "merge: a text that is not one element" (mergeOf "<w:rPr/>" "<w:rPr>" "<w:rPr/>" "<w:rPr/>") None
check "isEmpty" [ isEmpty "<w:rPr/>"; isEmpty "<w:rPr> </w:rPr>"; isEmpty "<w:rPr><w:b/></w:rPr>"; isEmpty "<w:rPr w:x=\"1\"/>" ] [ true; true; false; false ]

// ---- edits of an element anywhere in the tree; insert, prepend, append, remove ----
let drawing = """<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor a="1">
  <wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>
  <wp:extent cx="1" cy="2"/></wp:anchor></w:drawing></mc:Choice>
  <mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent>"""
let source = el drawing
let positionH = (descendants "positionH" source.Node).Head
let extent = (descendants "extent" source.Node).Head
let fallback = (descendants "Fallback" source.Node).Head
check "editsAt: edits of descendants, each in place, the rest the bytes"
    (apply (fun _ -> []) source (editsAt source.Node [ (positionH, [ RemoveChild "wp:align"; AddChild ("wp:posOffset", "<wp:posOffset>9</wp:posOffset>") ]);
                                                     (extent, [ SetAttribute ("cx", "5") ]);
                                                     (fallback, [ Removed ]) ]))
    """<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor a="1">
  <wp:positionH relativeFrom="column"><wp:posOffset>9</wp:posOffset></wp:positionH>
  <wp:extent cx="5" cy="2"/></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent>"""
let para = el "<w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>"
check "Prepend: after the properties; Append: at the end"
    (apply order para [ Prepend "<w:bookmarkStart w:id=\"1\"/>"; Append "<w:bookmarkEnd w:id=\"1\"/>" ])
    "<w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr><w:bookmarkStart w:id=\"1\"/><w:r><w:t>a</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>"
check "Prepend into an empty paragraph opens it" (apply order (el "<w:p/>") [ Prepend "<w:r/>" ]) "<w:p><w:r/></w:p>"
let run = (descendants "r" para.Node).Head
check "InsertBefore / InsertAfter a child"
    (apply order para (editsAt para.Node [ run, [ InsertBefore "<x/>"; InsertAfter "<y/>" ] ]))
    "<w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr><x/><w:r><w:t>a</w:t></w:r><y/></w:p>"
let fragment = (readFragmentRoot "<w:bookmarkStart w:id=\"1\"/> <w:p/>").Value
check "applyFragment: a fragment's elements edited as one"
    (applyFragment order fragment (editsAt fragment.Node [ (descendants "bookmarkStart" fragment.Node).Head, [ InsertAfter "<w:bookmarkEnd w:id=\"1\"/>" ] ]))
    "<w:bookmarkStart w:id=\"1\"/><w:bookmarkEnd w:id=\"1\"/> <w:p/>"
check "Style: declarations set in place and after, taken out"
    (Style.parse "position:absolute; margin-left:1pt;width:2pt" |> Style.set "margin-left" "3pt" |> Style.remove "width" |> Style.set "height" "4pt" |> Style.write)
    "position:absolute;margin-left:3pt;height:4pt"

// ---- the package's plumbing parts ----
let rels = "<?xml version=\"1.0\"?>\n<Relationships xmlns=\"p\">\n  <Relationship Id=\"rId1\" Type=\"t/styles\" Target=\"styles.xml\"/>\n  <Relationship Id=\"rId2\" Type=\"t/image\" Target=\"media/a.png\"/>\n</Relationships>"
check "relationships read" (OpcXml.relationships rels |> Option.map (List.map (fun rel -> rel.Id))) (Some [ "rId1"; "rId2" ])
check "a relationship taken out with its line, the rest the bytes"
    (OpcXml.withoutRelationships (fun rel -> rel.Id = "rId1") rels)
    "<?xml version=\"1.0\"?>\n<Relationships xmlns=\"p\">\n  <Relationship Id=\"rId2\" Type=\"t/image\" Target=\"media/a.png\"/>\n</Relationships>"
check "relationship ids referred to: by the relationships namespace, whatever its prefix"
    (set (OpcXml.relationshipIdsIn "<w:drawing xmlns:x=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><a:blip x:embed=\"rId5\" w:id=\"rId9\"/></w:drawing><v:imagedata r:id=\"rId6\" o:relid=\"rId7\"/>"))
    (set [ "rId5"; "rId6"; "rId7" ])
check "content types: by part name, by extension"
    (let lookup, _ = OpcXml.contentTypeLookup "<Types><Default Extension=\"png\" ContentType=\"image/png\"/><Override PartName=\"/word/document.xml\" ContentType=\"doc\"/></Types>"
     [ lookup "word/document.xml"; lookup "word/media/a.PNG"; lookup "word/x.bin" ])
    [ Some "doc"; Some "image/png"; None ]
check "part paths" [ PartPath.ofRelationships "word/_rels/document.xml.rels"; PartPath.ofRelationships "_rels/.rels"; PartPath.ofRelationships "word/document.xml" ] [ Some "word/document.xml"; Some ""; None ]
check "part paths: relative, within, numbered"
    [ PartPath.relative "xl/worksheets" "xl/comments1.xml"; PartPath.relative "xl/worksheets" "xl/worksheets/sheet2.xml"; string (PartPath.isIn "xl" "xl/media/a.png"); string (PartPath.isIn "xl" "docProps/app.xml"); string (PartPath.number "table" "xl/tables/table12.xml") ]
    [ "../comments1.xml"; "sheet2.xml"; "True"; "False"; "Some(12)" ]
check "relationship kinds" (OpcXml.relationshipKind "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet") "worksheet"

// ---- elements by where they start; an element replaced ----
let sheetText = "<worksheet><sheetData>\n<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c></row>\n<row r=\"2\"/></sheetData></worksheet>"
let sheet = (readPart sheetText).Value
let rowsOf (part: SourceElement) = descendants "row" part.Node
let cellB1 = (descendants "c" sheet.Node).[1]
check "EditAt: the child that starts there, its edits; Replaced: the element is other XML"
    (applyPart Spreadsheet.order sheet
        [ EditChild ("sheetData", [ EditAt ((span (rowsOf sheet).[0]).Start, [ SetAttribute ("ht", "20"); EditAt ((span cellB1).Start, [ InsertBefore "<c r=\"B0\"/>"; Replaced "<c r=\"B1\"><v>3</v></c>" ]) ])
                                    EditAt ((span (rowsOf sheet).[1]).Start, [ Append "<c r=\"A2\"/>" ]) ], None) ])
    "<worksheet><sheetData>\n<row r=\"1\" ht=\"20\"><c r=\"A1\"><v>1</v></c><c r=\"B0\"/><c r=\"B1\"><v>3</v></c></row>\n<row r=\"2\"><c r=\"A2\"/></row></sheetData></worksheet>"

// ---- a large part read in outline ----
let outline = (readPartExcept (fun name -> name = "row") sheetText).Value
check "readPartExcept: the shallow elements without their children, their spans and attributes exact"
    [ for row in rowsOf outline -> elements row |> List.length, bytesOf outline.Text row, attribute row "r" ]
    [ 0, "<row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"B1\"><v>2</v></c></row>", Some "1"; 0, "<row r=\"2\"/>", Some "2" ]
check "readPartExcept: an element of the shallow name in one of them is passed over with it"
    (readPartExcept (fun name -> name = "g") "<a><g k=\"&amp;\"><g/><g><b/></g></g><c x=\"1\"/></a>" |> Option.map (fun part -> [ for kid in elements part.Node -> name kid, attribute kid "k", elements kid |> List.length ]))
    (Some [ "g", Some "&", 0; "c", None, 0 ])
check "readPartExcept: not well formed" (readPartExcept (fun name -> name = "row") "<a><row></a>") None
let deepened = withWhole outline [ (span (rowsOf outline).[0]).Start ]
check "withWhole: the elements read whole where they are" [ for row in rowsOf deepened -> elements row |> List.length ] [ 2; 0 ]
check "withWhole: edits of what they hold" (applyPart Spreadsheet.order deepened [ EditChild ("sheetData", [ EditAt ((span (rowsOf deepened).[0]).Start, [ EditNth ("c", 0, [ SetAttribute ("s", "1") ]) ]) ], None) ])
    "<worksheet><sheetData>\n<row r=\"1\"><c r=\"A1\" s=\"1\"><v>1</v></c><c r=\"B1\"><v>2</v></c></row>\n<row r=\"2\"/></sheetData></worksheet>"
check "whole: an element read whole" (whole outline (rowsOf outline).[0] |> Option.map (fun row -> elements row.Node |> List.length)) (Some 2)
check "attributeBytes: as written, in order"
    (attributeBytes "<row  r=\"1\" s='2'/>" (readPart "<row  r=\"1\" s='2'/>").Value.Node)
    [ "r", "  r=\"1\""; "s", " s='2'" ]
check "Spreadsheet.order: a list put in at its place"
    (apply Spreadsheet.order (el "<styleSheet><fonts/><cellXfs/></styleSheet>") [ SetChild ("numFmts", "<numFmts/>"); SetChild ("dxfs", "<dxfs/>") ])
    "<styleSheet><numFmts/><fonts/><cellXfs/><dxfs/></styleSheet>"

printfn "xml-patch: %d passed, %d failed" pass fail
if fail > 0 then exit 1
