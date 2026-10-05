// The XML scanner (XmlDsl.tryScan) against the platform parser (XmlDsl.parse):
// every XML part of every DOCX under tests/ scans to the nodes `parse` gives,
// and what the scanner does not read is handed back (None) rather than read
// wrongly. The DOCX reader scans a document's main part (see
// Ns.parseWithLocalNames); the browser build is checked the same way by
// tools/word-browser (check `xml-scan`).
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

open System
open System.IO
open System.IO.Compression
open System.Collections.Generic

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-40s %s" label detail

let scan (xml: string) = XmlDsl.tryScan xml (HashSet<string>())

// ---- hand-written cases ----
let same label (xml: string) =
    let expected = XmlDsl.parse xml
    match scan xml with
    | Some node -> ok label (node = expected) (sprintf "scanned %A, parsed %A" node expected)
    | None -> ok label false "the scanner gave up"

same "elements and attributes" "<a x=\"1\" y='2'><b/><c z=\"q\">text</c></a>"
same "declaration, comments, instructions" "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<!-- c --><a><?pi data?><b>t</b><!-- inside --></a>\n<!-- after -->"
same "references" "<a v=\"&lt;&gt;&amp;&quot;&apos;&#65;&#x42;&#x1F600;\">&lt;x&gt; &amp; &#169; &#x1F600;</a>"
same "whitespace is trimmed" "<a>  <b>  padded  </b>\n  <c> </c></a>"
same "xml:space keeps it" "<a><t xml:space=\"preserve\">  kept  </t><t xml:space=\"preserve\"> </t><t xml:space=\"default\">  not  </t></a>"
same "xml:space is not inherited" "<a xml:space=\"preserve\"><b><c>  x  </c></b></a>"
same "CDATA" "<a><![CDATA[ <raw> & ]]></a>"
same "CDATA next to text" "<a>before <![CDATA[mid]]> after</a>"
same "comment splits text" "<a xml:space=\"preserve\">one<!-- c -->two</a>"
same "line ends" "<a v=\"l1\r\nl2\tx\">r1\r\nr2\rr3</a>"
same "namespaces as written" "<w:document xmlns:w=\"urn:w\" xmlns=\"urn:d\"><w:body w:val=\"1\"><plain/></w:body></w:document>"
same "byte order mark" "﻿<a/>"
same "space before the end of a tag" "<a  x = \"1\" ><b ></b ></a >"
same "an empty element with attributes" "<a><b x=\"1\"/><b x=\"2\" /></a>"
same "gt in an attribute" "<a v=\"1 > 0\">x</a>"

// What it does not read is left to the parser.
let givesUp label (xml: string) = ok label (scan xml).IsNone "the scanner read it"
givesUp "a DTD" "<!DOCTYPE a [<!ENTITY e \"x\">]><a>&e;</a>"
givesUp "an unknown entity" "<a>&nbsp;</a>"
givesUp "tags that do not match" "<a><b></a></b>"
givesUp "an element left open" "<a><b>"
givesUp "two roots" "<a/><b/>"
givesUp "text outside the root" "<a/>x"
givesUp "an attribute twice" "<a x=\"1\" x=\"2\"/>"
givesUp "an unquoted attribute" "<a x=1/>"
givesUp "a character reference XML does not allow" "<a>&#1;</a>"
givesUp "nothing" ""

// The names of the elements, as Ns.parseWithLocalNames reports them.
let _, names = XmlDsl.Ns.parseWithLocalNames "<w:a xmlns:w=\"urn:w\"><w:b/><c><w:d/></c></w:a>"
ok "local names" (names.SetEquals [ "a"; "b"; "c"; "d" ]) (sprintf "%A" (Seq.toList names))
let _, fallbackNames = XmlDsl.Ns.parseWithLocalNames "<w:a xmlns:w=\"urn:w\"><w:b>&#65;</w:b><!-- - --></w:a>"
ok "local names (b)" (fallbackNames.SetEquals [ "a"; "b" ]) (sprintf "%A" (Seq.toList fallbackNames))

// ---- every XML part of the test documents ----
let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let documents =
    Directory.GetDirectories(root, "test-documents*")
    |> Array.collect (fun dir -> Directory.GetFiles(dir, "*.docx", SearchOption.AllDirectories))
    |> Array.sort
let mutable parts = 0
let mutable scanned = 0
for path in documents do
    try
        use archive = ZipFile.OpenRead path
        for entry in archive.Entries do
            if entry.FullName.EndsWith ".xml" || entry.FullName.EndsWith ".rels" then
                use reader = new StreamReader(entry.Open(), Text.Encoding.UTF8)
                let xml = reader.ReadToEnd()
                match (try Some (XmlDsl.parse xml) with _ -> None) with
                | None -> ok (Path.GetFileName path + " " + entry.FullName) (scan xml).IsNone "the parser rejects the part, the scanner reads it"
                | Some expected ->
                    parts <- parts + 1
                    match scan xml with
                    | Some node ->
                        scanned <- scanned + 1
                        ok (Path.GetFileName path + " " + entry.FullName) (node = expected) "the scanner and the parser differ"
                    | None -> ()
    with :? InvalidDataException -> ()   // (not a zip: a fixture of the reader's own tests)
ok "the test documents were found" (documents.Length > 50 && parts > 500) (sprintf "%d documents, %d parts" documents.Length parts)
// (A part the scanner hands back is still read - by the parser - but then nothing is gained.)
ok "the scanner reads nearly every part" (scanned * 100 >= parts * 99) (sprintf "%d of %d parts" scanned parts)

printfn "word xml scan: %d passed, %d failed (%d documents, %d of %d parts scanned)" pass fail documents.Length scanned parts
if fail > 0 then exit 1
