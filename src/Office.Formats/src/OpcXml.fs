/// The package's plumbing parts (Open Packaging Conventions): relationships parts and [Content_Types].xml,
/// read through the span tree and changed by XmlPatch edits (docs/save-architecture.md) - an entry that
/// is taken out goes with the white space before it, one that is added goes after the others; the rest
/// stays the bytes it was read as.
module OpcXml

open XmlDsl

/// One relationship of a relationships part, its target as written.
type Relationship = { Id: string; Type: string; Target: string; External: bool }

/// The root element of a part: None when the part is not well formed.
let private root (xml: string) = XmlPatch.readPart xml

let private ofLocal (localName: string) (node: XmlNode) = XmlPatch.localName node = localName

/// The relationships of a relationships part, in order. None when it is not well formed.
let relationships (xml: string) : Relationship list option =
    root xml
    |> Option.map (fun part ->
        XmlPatch.elements part.Node
        |> List.filter (ofLocal "Relationship")
        |> List.map (fun node ->
            let value name = XmlPatch.attribute node name |> Option.defaultValue ""
            { Id = value "Id"; Type = value "Type"; Target = value "Target"; External = value "TargetMode" = "External" }))

/// The relationships of a relationships part, or none when it is not well formed.
let relationshipsOrEmpty (xml: string) : Relationship list = relationships xml |> Option.defaultValue []

/// A relationship's element.
let relationshipXml (rel: Relationship) : string =
    sprintf "<Relationship Id=\"%s\" Type=\"%s\" Target=\"%s\"%s/>"
        (XmlPatch.escapeAttribute rel.Id) (XmlPatch.escapeAttribute rel.Type) (XmlPatch.escapeAttribute rel.Target)
        (if rel.External then " TargetMode=\"External\"" else "")

/// The name of the relationships' element in the part (with the prefix the part gives it).
let private relationshipName (part: XmlPatch.SourceElement) =
    XmlPatch.elements part.Node |> List.tryFind (ofLocal "Relationship") |> Option.map XmlPatch.name |> Option.defaultValue "Relationship"

/// A relationships part without the relationships of these ids, and with these put in after the others.
let editRelationships (removed: string list) (added: Relationship list) (xml: string) : string =
    if removed.IsEmpty && added.IsEmpty then xml
    else
        match root xml with
        | None -> xml
        | Some part ->
            let name = relationshipName part
            XmlPatch.applyPart (fun _ -> []) part
                [ for id in List.distinct removed do yield XmlPatch.RemoveChildWhere (name, "Id", id)
                  for rel in added do yield XmlPatch.AddChild (name, relationshipXml rel) ]

/// A relationships part without the relationships `unwanted` says.
let withoutRelationships (unwanted: Relationship -> bool) (xml: string) : string =
    match relationships xml with
    | Some rels -> editRelationships (rels |> List.filter unwanted |> List.map (fun rel -> rel.Id)) [] xml
    | None -> xml

/// The directory of the part a relationships part is of ("word/_rels/document.xml.rels" -> "word").
let ownerDirectory (relsPath: string) : string =
    let segments = relsPath.Split('/') |> Array.toList
    match List.rev segments with
    | _ :: "_rels" :: directory -> directory |> List.rev |> String.concat "/"
    | _ -> ""

/// A relationship's target as the part path it names (an external one as it is).
let targetPath (ownerDirectory: string) (rel: Relationship) : string =
    if rel.External then rel.Target else ZipInterop.resolveTarget ownerDirectory rel.Target

/// A relationships part of the file for what is written now (`made`): the file's text while it names the
/// same parts; else the entries that are gone taken out of it and the new ones put in after the others.
/// `key` says when two entries are the same (with the targets as part paths).
let mergeRelationships (key: Relationship -> string) (renumber: bool) (source: string option) (made: string) : string =
    match source |> Option.bind (fun text -> relationships text |> Option.map (fun rels -> text, rels)), relationships made with
    | Some (text, was), Some now ->
        let nowKeys = now |> List.map key |> Set.ofList
        let wasKeys = was |> List.map key |> Set.ofList
        let gone = was |> List.filter (fun rel -> not (nowKeys.Contains (key rel)))
        let added = now |> List.filter (fun rel -> not (wasKeys.Contains (key rel)))
        if gone.IsEmpty && added.IsEmpty then text
        else
            // (A new entry keeps its id while no entry that stays has it; else - where ids may change - it gets one.)
            let ids = System.Collections.Generic.HashSet<string>(was |> List.filter (fun rel -> nowKeys.Contains (key rel)) |> List.map (fun rel -> rel.Id))
            let added =
                added
                |> List.map (fun rel ->
                    if ids.Add rel.Id || not renumber then rel
                    else
                        let mutable number = 1
                        while not (ids.Add (sprintf "rId%d" number)) do number <- number + 1
                        { rel with Id = sprintf "rId%d" number })
            editRelationships (gone |> List.map (fun rel -> rel.Id)) added text
    | _ -> made

// ---- [Content_Types].xml ----

/// An entry of [Content_Types].xml: a default (by extension) or an override (by part name).
type ContentTypeEntry =
    | Default of extension: string * contentType: string
    | Override of partName: string * contentType: string

/// The entries of [Content_Types].xml, in order. None when it is not well formed.
let contentTypeEntries (xml: string) : ContentTypeEntry list option =
    root xml
    |> Option.map (fun part ->
        XmlPatch.elements part.Node
        |> List.choose (fun node ->
            let value name = XmlPatch.attribute node name |> Option.defaultValue ""
            match XmlPatch.localName node with
            | "Default" -> Some (Default (value "Extension", value "ContentType"))
            | "Override" -> Some (Override (value "PartName", value "ContentType"))
            | _ -> None))

/// The content type [Content_Types].xml gives each part (by its path in the package): by its name, or by
/// its extension; and the part names it has overrides for (lower case, with their "/").
let contentTypeLookup (xml: string) : (string -> string option) * string list =
    let entries = contentTypeEntries xml |> Option.defaultValue []
    let defaults = entries |> List.choose (function Default (extension, contentType) -> Some (extension.ToLowerInvariant(), contentType) | _ -> None) |> Map.ofList
    let overrides = entries |> List.choose (function Override (name, contentType) -> Some (name.ToLowerInvariant(), contentType) | _ -> None) |> Map.ofList
    let lookup (path: string) =
        match Map.tryFind (("/" + path).ToLowerInvariant()) overrides with
        | Some contentType -> Some contentType
        | None ->
            match path.Split('/') |> Array.last |> fun name -> name.Split('.') with
            | parts when parts.Length > 1 -> Map.tryFind (parts.[parts.Length - 1].ToLowerInvariant()) defaults
            | _ -> None
    lookup, overrides |> Map.toList |> List.map fst

let private entryXml (entry: ContentTypeEntry) =
    match entry with
    | Default (extension, contentType) -> sprintf "<Default Extension=\"%s\" ContentType=\"%s\"/>" (XmlPatch.escapeAttribute extension) (XmlPatch.escapeAttribute contentType)
    | Override (name, contentType) -> sprintf "<Override PartName=\"%s\" ContentType=\"%s\"/>" (XmlPatch.escapeAttribute name) (XmlPatch.escapeAttribute contentType)

/// [Content_Types].xml without these entries and with these put in after the others of their kind.
let editContentTypes (removed: ContentTypeEntry list) (added: ContentTypeEntry list) (xml: string) : string =
    if removed.IsEmpty && added.IsEmpty then xml
    else
        match root xml with
        | None -> xml
        | Some part ->
            let nameOf (kind: string) = XmlPatch.elements part.Node |> List.tryFind (ofLocal kind) |> Option.map XmlPatch.name |> Option.defaultValue kind
            XmlPatch.applyPart (fun _ -> [ "Default"; "Override" ]) part
                [ for entry in removed do
                      match entry with
                      | Default (extension, _) -> yield XmlPatch.RemoveChildWhere (nameOf "Default", "Extension", extension)
                      | Override (name, _) -> yield XmlPatch.RemoveChildWhere (nameOf "Override", "PartName", name)
                  for entry in added do
                      match entry with
                      | Default _ -> yield XmlPatch.AddChild (nameOf "Default", entryXml entry)
                      | Override _ -> yield XmlPatch.AddChild (nameOf "Override", entryXml entry) ]

/// [Content_Types].xml of the file for what is written now (`made`), likewise: the entries that are gone
/// taken out, the new ones put in - but a default for an extension `unused` names (no part written has it).
let mergeContentTypes (unused: string list) (source: string option) (made: string) : string =
    let key (entry: ContentTypeEntry) =
        match entry with
        | Default (extension, contentType) -> "Default\n" + extension.ToLowerInvariant() + "\n" + contentType
        | Override (name, contentType) -> "Override\n" + name.ToLowerInvariant() + "\n" + contentType
    match source |> Option.bind (fun text -> contentTypeEntries text |> Option.map (fun entries -> text, entries)), contentTypeEntries made with
    | Some (text, was), Some now ->
        let nowKeys = now |> List.map key |> Set.ofList
        let wasKeys = was |> List.map key |> Set.ofList
        let gone = was |> List.filter (fun entry -> not (nowKeys.Contains (key entry)))
        let added =
            now
            |> List.filter (fun entry ->
                not (wasKeys.Contains (key entry))
                && (match entry with Default (extension, _) -> not (unused |> List.contains (extension.ToLowerInvariant())) | Override _ -> true))
            |> List.distinctBy key
        if gone.IsEmpty && added.IsEmpty then text else editContentTypes gone added text
    | _ -> made

// ---- References to relationships ----

let private relationshipsNamespaces =
    set [ "http://schemas.openxmlformats.org/officeDocument/2006/relationships"; "http://purl.oclc.org/ooxml/officeDocument/relationships" ]

/// The relationship ids the elements refer to - by an attribute in the relationships namespace (r:id,
/// r:embed, r:link, r:dm ...: the prefix as the XML declares it; "r" in a fragment that declares none), or
/// VML's o:relid -, each once, in order.
let private referencesIn (nodes: XmlNode list) : string list =
    let found = ResizeArray<string>()
    let rec visit (prefixes: Map<string, string>) (node: XmlNode) =
        match node with
        | Element (_, attrs, kids, _) ->
            let prefixes =
                attrs |> Map.fold (fun (known: Map<string, string>) name value -> if name.StartsWith "xmlns:" then Map.add (name.Substring 6) value known else known) prefixes
            for KeyValue (name, value) in attrs do
                match name.Split(':') with
                | [| prefix; _ |] when prefix <> "xmlns" && (Map.tryFind prefix prefixes |> Option.exists relationshipsNamespaces.Contains) -> found.Add value
                | [| "o"; "relid" |] -> found.Add value
                | _ -> ()
            for kid in kids do visit prefixes kid
        | _ -> ()
    for node in nodes do visit (Map.ofList [ "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ]) node
    found |> Seq.distinct |> List.ofSeq

/// The relationship ids XML refers to (a fragment: elements one after another).
let relationshipIdsIn (xml: string) : string list =
    match XmlPatch.readFragment xml with
    | Some elements -> referencesIn (elements |> List.map (fun element -> element.Node))
    | None -> []

/// The relationship ids a part refers to.
let relationshipIdsOfPart (xml: string) : Set<string> =
    match root xml with
    | Some part -> referencesIn [ part.Node ] |> Set.ofList
    | None -> Set.empty

/// The kind of a relationship by its type: the last segment of the type (".../relationships/worksheet" ->
/// "worksheet").
let relationshipKind (relType: string) : string = PartPath.fileName relType

/// Whether a relationship target names a place in the part itself ("#Sheet1!A1"), not a part.
let isAnchorTarget (target: string) : bool = target.StartsWith "#"
