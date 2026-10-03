module XmlGenerator

let private isForbiddenXmlChar (code: int) =
    (code < 0x20 && code <> 0x9 && code <> 0xA && code <> 0xD) || code = 0xFFFE || code = 0xFFFF

/// Drop characters XML 1.0 cannot represent at all: C0 controls other than
/// tab / LF / CR, U+FFFE / U+FFFF and unpaired surrogates. They cannot be
/// written as character references either (&#7; is just as ill-formed), so a
/// stray BEL or form feed pasted into a document used to produce a part that
/// neither Word nor this library's reader could open.
let stripInvalidXmlChars (text: string) : string =
    if isNull text then ""
    else
        let mutable suspect = false
        let mutable i = 0
        while not suspect && i < text.Length do
            let code = int text.[i]
            if isForbiddenXmlChar code || (code >= 0xD800 && code <= 0xDFFF) then suspect <- true
            i <- i + 1
        if not suspect then text
        else
            let sb = System.Text.StringBuilder(text.Length)
            let mutable j = 0
            while j < text.Length do
                let ch = text.[j]
                let code = int ch
                if code >= 0xD800 && code <= 0xDBFF then
                    if j + 1 < text.Length && int text.[j + 1] >= 0xDC00 && int text.[j + 1] <= 0xDFFF then
                        sb.Append(ch).Append(text.[j + 1]) |> ignore
                        j <- j + 1
                elif code >= 0xDC00 && code <= 0xDFFF then ()
                elif not (isForbiddenXmlChar code) then sb.Append(ch) |> ignore
                j <- j + 1
            sb.ToString()

/// Escape XML entities
let escapeXml (text: string) : string =
    (stripInvalidXmlChars text)
        .Replace("&", "&amp;")
        .Replace("<", "&lt;")
        .Replace(">", "&gt;")
        .Replace("\"", "&quot;")
        .Replace("'", "&apos;")

/// Escape an attribute value. Literal tab / LF / CR in an attribute are
/// normalised to spaces by every XML parser, so they go out as character
/// references to survive the round trip.
let escapeXmlAttribute (text: string) : string =
    (escapeXml text)
        .Replace("\t", "&#9;")
        .Replace("\n", "&#10;")
        .Replace("\r", "&#13;")

/// Generate XML element with attributes and children
let element (name: string) (attrs: (string * string) list) (children: string list) : string =
    let attrStr =
        attrs
        |> List.map (fun (k, v) -> sprintf "%s=\"%s\"" k (escapeXmlAttribute v))
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
        |> List.map (fun (k, v) -> sprintf "%s=\"%s\"" k (escapeXmlAttribute v))
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
