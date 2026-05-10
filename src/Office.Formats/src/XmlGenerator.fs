module XmlGenerator

/// Escape XML entities
let escapeXml (text: string) : string =
    text
        .Replace("&", "&amp;")
        .Replace("<", "&lt;")
        .Replace(">", "&gt;")
        .Replace("\"", "&quot;")
        .Replace("'", "&apos;")

/// Generate XML element with attributes and children
let element (name: string) (attrs: (string * string) list) (children: string list) : string =
    let attrStr =
        attrs
        |> List.map (fun (k, v) -> sprintf "%s=\"%s\"" k (escapeXml v))
        |> String.concat " "

    let attrPart = if attrStr = "" then "" else " " + attrStr

    match children with
    | [] -> sprintf "<%s%s/>" name attrPart
    | _ ->
        let childStr = String.concat "" children
        sprintf "<%s%s>%s</%s>" name attrPart childStr name

/// Generate XML element with text content (automatically escaped)
let elementWithText (name: string) (attrs: (string * string) list) (text: string) : string =
    let attrStr =
        attrs
        |> List.map (fun (k, v) -> sprintf "%s=\"%s\"" k (escapeXml v))
        |> String.concat " "

    let attrPart = if attrStr = "" then "" else " " + attrStr

    if text = "" then
        sprintf "<%s%s/>" name attrPart
    else
        sprintf "<%s%s>%s</%s>" name attrPart (escapeXml text) name

/// Generate self-closing XML element
let selfClosing (name: string) (attrs: (string * string) list) : string =
    element name attrs []

/// Generate XML declaration
let declaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"

/// Common Office namespaces
module Namespaces =
    let word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    let math = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    let spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    let relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    let contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types"
    let relationshipsPackage = "http://schemas.openxmlformats.org/package/2006/relationships"

/// Common relationship types
module RelationshipTypes =
    let officeDocument = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    let worksheet = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    let styles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    let sharedStrings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    let image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    let hyperlink = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

/// Build a complete XML document with declaration
let document (root: string) : string =
    declaration + "\n" + root
