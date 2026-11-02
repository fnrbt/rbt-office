module OfficeDsl

open XmlDsl

// Common Office Open XML patterns
type Relationship = {
    Id: string
    Type: string
    Target: string
    TargetMode: string option
}

type OfficeDocument<'TPart> = {
    Parts: Map<string, 'TPart>
    Relationships: Relationship list
}

// DSL for Office document structure traversal
module Office =
    let relType (relType: string) (rels: Relationship list) =
        rels |> List.filter (fun r -> r.Type.EndsWith(relType))

    let relById id (rels: Relationship list) =
        rels |> List.tryFind (fun r -> r.Id = id)

    let partByPath path (doc: OfficeDocument<'T>) =
        Map.tryFind path doc.Parts

// Relationship parsing DSL
// Matched by local name: a package may bind the relationships namespace to a
// prefix (`<rel:Relationship>`), which is still a valid .rels part.
let parseRelationships (xmlNode: XmlNode) : Relationship list =
    Ns.descendantsNs "Relationship" xmlNode
    |> List.choose (fun node ->
        xml {
            let! id = Xml.attr "Id" node
            let! relType = Xml.attr "Type" node
            let! target = Xml.attr "Target" node
            return { Id = id; Type = relType; Target = target; TargetMode = Xml.attr "TargetMode" node }
        })

/// The relationship id an element carries (`r:id` in practice). The prefix is
/// whatever the part bound the relationships namespace to, so fall back to any
/// prefixed `id` attribute — the unprefixed `id` (e.g. on `<brk>`) is a
/// different attribute and never matches.
let relationshipId (node: XmlNode) : string option =
    match Xml.attr "r:id" node with
    | Some v -> Some v
    | None ->
        match node with
        | Element(_, attrs, _, _) ->
            attrs
            |> Map.tryPick (fun name value ->
                if name.EndsWith ":id" && not (name.StartsWith "xmlns:") then Some value else None)
        | _ -> None

// Common text extraction pattern
type TextRun = {
    Text: string
    Style: Map<string, string>
}

module Text =
    let run text style = { Text = text; Style = style }
    let plain text = run text Map.empty
    let join (runs: TextRun list) = runs |> List.map (fun r -> r.Text) |> String.concat ""

// Computation expression for building structured documents
type DocumentBuilder() =
    member _.Yield(x) = [x]
    member _.YieldFrom(xs) = xs
    member _.Combine(a, b) = a @ b
    member _.Delay(f) = f()
    member _.Zero() = []
    member _.For(xs, f) = List.collect f xs

let document = DocumentBuilder()

// Active patterns for common Office XML structures
let (|WithText|_|) node =
    Xml.text node

let (|WithAttr|_|) attrName node =
    Xml.attr attrName node

// Lens-like accessors for nested Office structures
type Lens<'a, 'b> = {
    Get: 'a -> 'b option
    GetAll: 'a -> 'b list
}

let lens get getAll = { Get = get; GetAll = getAll }

module Lens =
    let get l x = l.Get x
    let getAll l x = l.GetAll x
    let compose l1 l2 = {
        Get = fun x -> l1.Get x |> Option.bind l2.Get
        GetAll = fun x -> l1.GetAll x |> List.collect l2.GetAll
    }

    let (>->) = compose
