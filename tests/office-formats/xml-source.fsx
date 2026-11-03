// Spans (docs/save-architecture.md) over every XML part of every document and workbook of the
// corpus: the scanner reads the tree .NET's parser reads (spans aside), and the span of every
// element is its place in the part - its slice starts with its start tag and ends with its end
// tag, and the elements in it lie inside its content, one after another. The root's children read one at
// a time by the lexer are the scanner's, and an outline read has them where they are.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

open System
open System.IO
open System.IO.Compression
open XmlDsl

let mutable pass = 0
let mutable fail = 0
let failures = ResizeArray<string>()
let ok label cond detail =
    if cond then pass <- pass + 1
    else
        fail <- fail + 1
        if failures.Count < 40 then failures.Add (sprintf "FAIL %s  %s" label detail)

let root = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "..", ".."))
let files =
    Directory.GetFiles(Path.Combine(root, "tests"), "*.*", SearchOption.AllDirectories)
    |> Array.filter (fun f -> [ ".docx"; ".xlsx"; ".xlsm"; ".pptx"; ".dotx" ] |> List.contains (Path.GetExtension(f).ToLowerInvariant()))
    |> Array.sort

let rec withoutSpans (node: XmlNode) =
    match node with
    | Element (name, attrs, children, _) -> Element (name, attrs, List.map withoutSpans children, Span.none)
    | other -> other

let rec placed (text: string) (node: XmlNode) : string option =
    match node with
    | Element (name, _, children, span) ->
        if not (Span.isKnown span) then Some (name + ": no span")
        elif String.CompareOrdinal(text, span.Start, "<" + name, 0, name.Length + 1) <> 0 then Some (sprintf "%s does not start at %d" name span.Start)
        elif span.StartTagEnd = span.End then
            if text.[span.End - 2] <> '/' || text.[span.End - 1] <> '>' then Some (name + ": self-closed, no />") else None
        elif not (span.Start < span.StartTagEnd && span.StartTagEnd <= span.EndTagStart && span.EndTagStart < span.End) then Some (name + ": span out of order")
        elif text.Substring(span.EndTagStart, span.End - span.EndTagStart).TrimEnd('>').TrimEnd() <> "</" + name then Some (name + ": end tag " + text.Substring(span.EndTagStart, span.End - span.EndTagStart))
        else
            let mutable at = span.StartTagEnd
            let mutable problem = None
            for child in children do
                match problem, child with
                | None, Element (_, _, _, inner) ->
                    if inner.Start < at then problem <- Some (sprintf "%s: a child at %d before %d" name inner.Start at)
                    else
                        at <- inner.End
                        problem <- placed text child
                | _ -> ()
            match problem with
            | None when at > span.EndTagStart -> Some (name + ": a child past the end tag")
            | other -> other
    | _ -> None

let scanWatch = Diagnostics.Stopwatch()
let mutable parts = 0
let mutable bytes = 0L
let mutable notScanned = 0
for file in files do
    let name = Path.GetRelativePath(root, file)
    try
        use zip = ZipFile.OpenRead file
        for entry in zip.Entries do
            if entry.FullName.EndsWith ".xml" || entry.FullName.EndsWith ".rels" || entry.FullName.EndsWith ".vml" then
                use reader = new StreamReader(entry.Open(), Text.Encoding.UTF8, true)
                let text = reader.ReadToEnd()
                parts <- parts + 1
                bytes <- bytes + int64 text.Length
                let label = name + "!" + entry.FullName
                scanWatch.Start()
                let scanned = tryScan text (Collections.Generic.HashSet())
                scanWatch.Stop()
                match scanned with
                | None -> notScanned <- notScanned + 1
                | Some node ->
                    let reference = try Some (parse text) with _ -> None
                    match reference with
                    | Some other -> ok (label + ": the tree .NET reads") (withoutSpans node = withoutSpans other) ""
                    | None -> ok label false ".NET does not read what the scanner reads"
                    match placed text node with
                    | Some problem -> ok (label + ": spans") false problem
                    | None -> ok (label + ": spans") true ""
                    // The root's children read one at a time by the lexer (as a long document's body is
                    // read) are the scanner's; in outline (none of them read) they are where they are.
                    match node, rootStartTag text with
                    | Element (rootName, rootAttrs, kids, span), Some (Element (name, attrs, [], start), lexer) ->
                        ok (label + ": the root's start tag") (name = rootName && attrs = rootAttrs && start.Start = span.Start && start.StartTagEnd = span.StartTagEnd) ""
                        let read = ResizeArray<XmlNode>()
                        let mutable going = not lexer.SelfClosing
                        while going do
                            match lexer.Next () with
                            | XmlToken.StartTag ->
                                match tryReadElement lexer (Collections.Generic.HashSet()) with
                                | Some child -> read.Add child
                                | None -> going <- false
                            | XmlToken.EndTag | XmlToken.EndOfText | XmlToken.Malformed -> going <- false
                            | _ -> ()
                        let elements = kids |> List.filter (function Element _ -> true | _ -> false)
                        // (A root's xml:space holds for the whole part: not for its elements read one by one.)
                        if not (rootAttrs.ContainsKey "xml:space") then
                            ok (label + ": element by element") (List.ofSeq read = elements) ""
                        let names = elements |> List.choose Xml.name |> List.map Ns.stripNs |> Set.ofList
                        match tryScanOutline names.Contains text with
                        | Some (Element (_, _, outlineKids, _)) ->
                            let shallow = elements |> List.map (function Element (name, attrs, _, span) -> Element (name, attrs, [], span) | other -> other)
                            ok (label + ": in outline") (outlineKids |> List.filter (function Element _ -> true | _ -> false) = shallow) ""
                        | _ -> ok (label + ": in outline") false "not read"
                    | _ -> ok (label + ": the root's start tag") false "not read"
    with error -> ok name false ("not read: " + error.Message)

for line in failures do printfn "%s" line
printfn "%d parts (%d MB) of %d files; %d the scanner does not read; scanner %d ms" parts (bytes / 1000000L) files.Length notScanned scanWatch.ElapsedMilliseconds
printfn "XML span tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
