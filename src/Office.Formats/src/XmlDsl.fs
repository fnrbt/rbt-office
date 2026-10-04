module XmlDsl

open Fable.Core
open Fable.Core.JsInterop
open Browser.Types
#if !FABLE_COMPILER
open System.Xml.Linq
#endif

// XML DSL for succinct navigation and querying
type XmlNode =
    | Element of name: string * attrs: Map<string, string> * children: XmlNode list
    | Text of string
    | Empty

// Active patterns for XML matching
let (|ElementNamed|_|) name = function
    | Element(n, attrs, children) when n = name -> Some(attrs, children)
    | _ -> None

let (|HasAttr|_|) attrName node =
    match node with
    | Element(_, attrs, _) -> Map.tryFind attrName attrs
    | _ -> None

let (|TextValue|_|) = function
    | Text t -> Some t
    | _ -> None

// Computation expression for XML navigation
type XmlBuilder() =
    member _.Bind(node, f) =
        match node with
        | Some n -> f n
        | None -> None

    member _.Return(x) = Some x
    member _.ReturnFrom(x) = x
    member _.Zero() = None

let xml = XmlBuilder()

// Core XML operations with point-free style
module Xml =
    let rec children = function
        | Element(_, _, kids) -> kids
        | _ -> []

    let child name node =
        children node |> List.tryFind (function Element(n, _, _) when n = name -> true | _ -> false)

    let descendants name node =
        let rec descend acc = function
            | Element(n, a, kids) as elem ->
                let acc' = if n = name then elem :: acc else acc
                List.fold descend acc' kids
            | _ -> acc
        descend [] node |> List.rev

    let attr name = function
        | Element(_, attrs, _) -> Map.tryFind name attrs
        | _ -> None

    let text = function
        | Text t -> Some t
        | Element(_, _, kids) ->
            kids
            |> List.choose (function Text t -> Some t | _ -> None)
            |> String.concat ""
            |> Some
        | _ -> None

    let name = function
        | Element(n, _, _) -> Some n
        | _ -> None

    // Combinator-style queries
    let (|>>) node f = Option.bind f (Some node)
    let (=>>) nodes f = List.choose f nodes
    let (?>) node name = child name node

// Parse XML from string (uses browser DOMParser in Fable)
#if FABLE_COMPILER
[<Emit("new DOMParser().parseFromString($0, 'text/xml')")>]
let private domParse (xml: string) : obj = jsNative

// DOMParser never throws on malformed XML: it returns a document holding a
// <parsererror> element (the root in Firefox, nested in the partial tree in
// Chromium/WebKit). Returns the error text, or null when the parse was clean.
[<Emit("(function (d) { var e = d.getElementsByTagName('parsererror'); return e.length > 0 ? (e[0].textContent || 'parse error') : null; })($0)")>]
let private domParseError (doc: obj) : string = jsNative

[<Emit("($0.doctype != null)")>]
let private domHasDoctype (doc: obj) : bool = jsNative

/// A DOM node as an XmlNode. `preserve`: the parent element has
/// xml:space="preserve" (in Word XML: whitespace is text).
let rec private convertNode (node: obj) (preserve: bool) : XmlNode =
    let nodeType: int = node?nodeType

    match nodeType with
    | 1 -> // ELEMENT_NODE
        let nodeName: string = node?nodeName
        let attrList: obj = node?attributes
        let length: int = attrList?length
        let mutable attrs : Map<string, string> = Map.empty
        let mutable preserveChildren = false
        for i in 0 .. length - 1 do
            let attr: obj = attrList?(i)
            let name: string = attr?name
            let value: string = attr?value
            if name = "xml:space" then preserveChildren <- (value = "preserve")
            attrs <- Map.add name value attrs

        // (From the last child to the first: the list is built front to back.)
        let childNodes: obj = node?childNodes
        let childCount: int = childNodes?length
        let mutable children : XmlNode list = []
        for i = childCount - 1 downto 0 do
            match convertNode (childNodes?(i)) preserveChildren with
            | Empty -> ()
            | child -> children <- child :: children

        Element(nodeName, attrs, children)

    | 3 -> // TEXT_NODE
        let textContent: string = node?textContent
        if preserve then
            // Preserve all whitespace when xml:space="preserve"
            if textContent.Length > 0 then Text textContent else Empty
        else
            // Otherwise trim whitespace as usual
            let trimmed = textContent.Trim()
            if trimmed.Length > 0 then Text trimmed else Empty

    | 4 -> // CDATA_SECTION_NODE — literal text, kept verbatim like the .NET path
        let textContent: string = node?textContent
        if textContent.Length > 0 then Text textContent else Empty

    | _ -> Empty

let parse (xmlString: string) : XmlNode =
    let xmlString = xmlString.TrimStart(char 0xFEFF, char 0x200B)
    let doc = domParse xmlString
    let error = domParseError doc
    if not (isNull error) then
        failwith ("Malformed XML: " + error)
    // OPC parts must not carry a DTD; the .NET path prohibits them too.
    if domHasDoctype doc then
        failwith "Malformed XML: DTD is not allowed"
    let root: obj = doc?documentElement
    if isNull root then Empty else convertNode root false
#else
let private xName (element: XElement) (name: XName) =
    let prefix = element.GetPrefixOfNamespace(name.Namespace)
    if System.String.IsNullOrEmpty(prefix) then name.LocalName else prefix + ":" + name.LocalName

let rec private convertXNode (parentAttrs: Map<string, string> option) (node: XNode) : XmlNode =
    match node with
    | :? XElement as element ->
        let attrs =
            element.Attributes()
            |> Seq.map (fun attr -> xName element attr.Name, attr.Value)
            |> Map.ofSeq

        let children =
            element.Nodes()
            |> Seq.map (convertXNode (Some attrs))
            |> Seq.filter ((<>) Empty)
            |> Seq.toList

        Element(xName element element.Name, attrs, children)

    | :? XCData as cdata ->
        if cdata.Value.Length > 0 then Text cdata.Value else Empty

    | :? XText as text ->
        let shouldPreserve =
            parentAttrs
            |> Option.bind (Map.tryFind "xml:space")
            |> Option.map ((=) "preserve")
            |> Option.defaultValue false

        if shouldPreserve then
            if text.Value.Length > 0 then Text text.Value else Empty
        else
            let trimmed = text.Value.Trim()
            if trimmed.Length > 0 then Text trimmed else Empty

    | _ -> Empty

let parse (xmlString: string) : XmlNode =
    let xmlString = xmlString.TrimStart(char 0xFEFF, char 0x200B)
    // Parts come from untrusted packages: never process a DTD (entity
    // expansion / external subsets), matching the streaming XmlReader path.
    let settings = System.Xml.XmlReaderSettings()
    settings.DtdProcessing <- System.Xml.DtdProcessing.Prohibit
    settings.XmlResolver <- null
    use stringReader = new System.IO.StringReader(xmlString)
    use reader = System.Xml.XmlReader.Create(stringReader, settings)
    let doc = XDocument.Load(reader, LoadOptions.PreserveWhitespace)
    match doc.Root with
    | null -> Empty
    | root -> convertXNode None root
#endif

// Namespace-aware helpers for Office Open XML
module Ns =
    let stripNs (name: string) =
        match name.LastIndexOf(':') with
        | -1 -> name
        | idx -> name.Substring(idx + 1)

    /// `stripNs name = localName`, without making the stripped name.
    let private hasLocalName (localName: string) (name: string) =
        let colon = name.LastIndexOf(':')
        if colon < 0 then name = localName
        else name.Length - colon - 1 = localName.Length && name.EndsWith(localName, System.StringComparison.Ordinal)

    let childNs localName node =
        Xml.children node
        |> List.tryFind (fun n ->
            match n with
            | Element (name, _, _) -> hasLocalName localName name
            | _ -> false)

    /// The elements under (and including) a node that have a local name, in
    /// document order.
    let descendantsNs localName node =
        let found = ResizeArray<XmlNode>()
        let rec visit (current: XmlNode) =
            match current with
            | Element (name, _, kids) ->
                if hasLocalName localName name then found.Add current
                for kid in kids do
                    visit kid
            | _ -> ()
        visit node
        let mutable result = []
        for i = found.Count - 1 downto 0 do
            result <- found.[i] :: result
        result
