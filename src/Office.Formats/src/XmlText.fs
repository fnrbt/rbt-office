/// XML as text: a node serialised, and a fragment made self-contained. (Used by the DOCX and the
/// XLSX readers both; a module of its own so that what reads a workbook does not load DocxReader.)
module XmlText

open XmlDsl
open System

let private appendEscapedXml (builder: System.Text.StringBuilder) (text: string) =
    if not (String.IsNullOrEmpty text) then
        let mutable segmentStart = 0
        for i = 0 to text.Length - 1 do
            let replacement =
                match text.[i] with
                | '&' -> ValueSome "&amp;"
                | '<' -> ValueSome "&lt;"
                | '>' -> ValueSome "&gt;"
                | '"' -> ValueSome "&quot;"
                | '\'' -> ValueSome "&apos;"
                | _ -> ValueNone
            match replacement with
            | ValueSome escaped ->
                if i > segmentStart then
                    builder.Append(text, segmentStart, i - segmentStart) |> ignore
                builder.Append(escaped) |> ignore
                segmentStart <- i + 1
            | ValueNone -> ()
        if segmentStart < text.Length then
            builder.Append(text, segmentStart, text.Length - segmentStart) |> ignore

let rec private appendXmlTo (builder: System.Text.StringBuilder) node =
    match node with
    | XmlDsl.Text text ->
        appendEscapedXml builder text
    | XmlDsl.Empty -> ()
    | XmlDsl.Element(name, attrs, children, _) ->
        builder.Append('<').Append(name) |> ignore
        for KeyValue(key, value) in attrs do
            builder.Append(' ').Append(key).Append("=\"") |> ignore
            appendEscapedXml builder value
            builder.Append('"') |> ignore
        match children with
        | [] ->
            builder.Append("/>") |> ignore
        | _ ->
            builder.Append('>') |> ignore
            for child in children do
                appendXmlTo builder child
            builder.Append("</").Append(name).Append('>') |> ignore

let xmlToString node =
    let builder = System.Text.StringBuilder()
    appendXmlTo builder node
    builder.ToString()

let private xmlnsDeclarations (attrs: Map<string, string>) : Map<string, string> =
    attrs
    |> Map.toList
    |> List.choose (fun (name, value) -> Ns.declaredPrefix name |> Option.map (fun prefix -> prefix, value))
    |> Map.ofList

/// The words of an attribute's value that is a list (mc:Ignorable="w14 wp14", Requires="wps").
let private valueWords (value: string) = ValueText.words value

/// `xml` (one element, serialised) with the namespace declarations of the
/// part's root element it relies on added to its own root, so it stays valid
/// wherever it is written. `rootAttributes` are the attributes of the root
/// element of the part the fragment was taken from.
let selfContainedXml (rootAttributes: Map<string, string>) (xml: string) : string =
    let rootNamespaces = xmlnsDeclarations rootAttributes
    // The declarations one element needs: (its edits) - none when it declares what it uses.
    let edits (element: XmlNode) : XmlPatch.ElementEdit list =
        let nodes = XmlPatch.descendantsWhere (fun _ -> true) element
        let attributes = nodes |> List.collect (fun node -> XmlPatch.attributes node |> Map.toList)
        // (The prefixes it uses: of its elements and attributes, and those its alternate content requires or ignores.)
        let used =
            [ for node in nodes do yield XmlPatch.prefix node
              for (name, value) in attributes do
                  match ModelText.prefixOf name with
                  | "" | "xmlns" -> ()
                  | prefix -> yield prefix
                  if name = "Requires" || name = "mc:Ignorable" then yield! valueWords value ]
            |> List.filter (fun prefix -> prefix <> "")
            |> Set.ofList
        let own = XmlPatch.attributes element
        // Prefixes the part declares ignorable and the fragment uses.
        let ignorable =
            match Map.tryFind "mc:Ignorable" rootAttributes with
            | Some value when not (own.ContainsKey "mc:Ignorable") -> valueWords value |> List.filter (fun prefix -> Set.contains prefix used)
            | _ -> []
        let used = if ignorable.IsEmpty then used else Set.add "mc" used
        let additions =
            [ for KeyValue(prefix, uri) in rootNamespaces do
                if Set.contains prefix used && not (own.ContainsKey ("xmlns:" + prefix)) then yield "xmlns:" + prefix, uri
              if not ignorable.IsEmpty then yield "mc:Ignorable", String.concat " " ignorable ]
        // (Its attributes in name order, the order the serialiser writes them in: reading the result again and
        // serialising it gives the same text. The rest of it stays the text it was.)
        if additions.IsEmpty then [] else [ XmlPatch.SetAllAttributes (Map.toList own @ additions |> Map.ofList |> Map.toList) ]
    if rootNamespaces.IsEmpty then xml
    else
        // (One element, or several one after another - a content control's properties and end properties.)
        match XmlPatch.readFragmentRoot xml with
        | Some fragment ->
            let changes = XmlPatch.elements fragment.Node |> List.map (fun element -> element, edits element) |> List.filter (fun (_, found) -> not found.IsEmpty)
            if changes.IsEmpty then xml else XmlPatch.applyFragment (fun _ -> []) fragment (XmlPatch.editsAt fragment.Node changes)
        | None -> xml
