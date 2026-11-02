module XmlDsl

open Fable.Core
open Fable.Core.JsInterop
open Browser.Types
#if !FABLE_COMPILER
open System.Xml.Linq
#endif

/// Where an element is in the text of its part (docs/save-architecture.md): offsets of the `<` of
/// its start tag, after that tag's `>`, of the `<` of its end tag and after its `>`. A self-closed
/// element has StartTagEnd = EndTagStart = End. An element that was not read from a part (built, or
/// read by the platform's parser) has Start = -1.
type Span = {
    Start: int
    StartTagEnd: int
    EndTagStart: int
    End: int
}

module Span =
    let none = { Start = -1; StartTagEnd = -1; EndTagStart = -1; End = -1 }
    let isKnown (span: Span) = span.Start >= 0

// XML DSL for succinct navigation and querying
type XmlNode =
    | Element of name: string * attrs: Map<string, string> * children: XmlNode list * span: Span
    | Text of string
    | Empty

// Active patterns for XML matching
let (|ElementNamed|_|) name = function
    | Element(n, attrs, children, _) when n = name -> Some(attrs, children)
    | _ -> None

let (|HasAttr|_|) attrName node =
    match node with
    | Element(_, attrs, _, _) -> Map.tryFind attrName attrs
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
        | Element(_, _, kids, _) -> kids
        | _ -> []

    let child name node =
        children node |> List.tryFind (function Element(n, _, _, _) when n = name -> true | _ -> false)

    let descendants name node =
        let rec descend acc = function
            | Element(n, a, kids, _) as elem ->
                let acc' = if n = name then elem :: acc else acc
                List.fold descend acc' kids
            | _ -> acc
        descend [] node |> List.rev

    let attr name = function
        | Element(_, attrs, _, _) -> Map.tryFind name attrs
        | _ -> None

    let text = function
        | Text t -> Some t
        | Element(_, _, kids, _) ->
            kids
            |> List.choose (function Text t -> Some t | _ -> None)
            |> String.concat ""
            |> Some
        | _ -> None

    let name = function
        | Element(n, _, _, _) -> Some n
        | _ -> None

    // Combinator-style queries
    let (|>>) node f = Option.bind f (Some node)
    let (=>>) nodes f = List.choose f nodes
    let (?>) node name = child name node

/// Text without its XML whitespace: a no-break or an en space is a character of the text.
let private trimXml (text: string) = text.Trim (' ', '\t', '\r', '\n')

/// `xml:space` holds for an element's descendants too: a part can say "preserve" once, on its root
/// (Word keeps the spaces of `resume-template`'s runs by that). What is inherited counts only in the
/// elements that hold text: the whitespace between other elements of such a part is not text.
let private holdsText (name: string) =
    match name.Substring (name.IndexOf ':' + 1) with
    | "t" | "delText" | "instrText" | "delInstrText" -> true
    | _ -> false
let private keepsSpace (name: string) (own: string option) (inherited: bool) =
    match own with
    | Some value -> value = "preserve"
    | None -> inherited && holdsText name
let private spaceInherited (own: string option) (inherited: bool) =
    match own with
    | Some value -> value = "preserve"
    | None -> inherited

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

/// A DOM node as an XmlNode. `preserve`: its text keeps its whitespace in
/// the parent element; `inherited`: the `xml:space` of the parent's ancestors.
let rec private convertNode (node: obj) (preserve: bool) (inherited: bool) : XmlNode =
    let nodeType: int = node?nodeType

    match nodeType with
    | 1 -> // ELEMENT_NODE
        let nodeName: string = node?nodeName
        let attrList: obj = node?attributes
        let length: int = attrList?length
        let mutable attrs : Map<string, string> = Map.empty
        let mutable space : string option = None
        for i in 0 .. length - 1 do
            let attr: obj = attrList?(i)
            let name: string = attr?name
            let value: string = attr?value
            if name = "xml:space" then space <- Some value
            attrs <- Map.add name value attrs
        let preserveChildren = keepsSpace nodeName space inherited
        let inheritedByChildren = spaceInherited space inherited

        // (From the last child to the first: the list is built front to back.)
        let childNodes: obj = node?childNodes
        let childCount: int = childNodes?length
        let mutable children : XmlNode list = []
        for i = childCount - 1 downto 0 do
            match convertNode (childNodes?(i)) preserveChildren inheritedByChildren with
            | Empty -> ()
            | child -> children <- child :: children

        Element(nodeName, attrs, children, Span.none)

    | 3 -> // TEXT_NODE
        let textContent: string = node?textContent
        if preserve then
            // Preserve all whitespace when xml:space="preserve"
            if textContent.Length > 0 then Text textContent else Empty
        else
            // Otherwise trim whitespace as usual
            let trimmed = trimXml textContent
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
    if isNull root then Empty else convertNode root false false
#else
let private xName (element: XElement) (name: XName) =
    let prefix = element.GetPrefixOfNamespace(name.Namespace)
    if System.String.IsNullOrEmpty(prefix) then name.LocalName else prefix + ":" + name.LocalName

let rec private convertXNode (preserve: bool) (inherited: bool) (node: XNode) : XmlNode =
    match node with
    | :? XElement as element ->
        let attrs =
            element.Attributes()
            |> Seq.map (fun attr -> xName element attr.Name, attr.Value)
            |> Map.ofSeq

        let name = xName element element.Name
        let space = Map.tryFind "xml:space" attrs
        let children =
            element.Nodes()
            |> Seq.map (convertXNode (keepsSpace name space inherited) (spaceInherited space inherited))
            |> Seq.filter ((<>) Empty)
            |> Seq.toList

        Element(name, attrs, children, Span.none)

    | :? XCData as cdata ->
        if cdata.Value.Length > 0 then Text cdata.Value else Empty

    | :? XText as text ->
        if preserve then
            if text.Value.Length > 0 then Text text.Value else Empty
        else
            let trimmed = trimXml text.Value
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
    | root -> convertXNode false false root
#endif

// ---------------------------------------------------------------------------
// A scanner from the XML text straight to XmlNode, for the large parts of a
// package (a long document's main part). `parse` goes through the platform's
// parser and then walks its tree; on a document of several megabytes that is
// most of the time the reader takes. The scanner reads what Office writes -
// elements, attributes, text, CDATA, comments, processing instructions and
// the predefined / numeric character references - and gives up (None) on
// anything else (a DTD, an unknown entity, tags that do not match), which
// the caller then hands to `parse`.
// ---------------------------------------------------------------------------

exception private ScanGiveUp

let private scanGiveUp () : 'T = raise ScanGiveUp

let private indexOfOrdinal (s: string) (value: string) (from: int) : int =
#if FABLE_COMPILER
    s.IndexOf(value, from)
#else
    s.IndexOf(value, from, System.StringComparison.Ordinal)
#endif

let private startsAt (s: string) (at: int) (value: string) : bool =
    if at + value.Length > s.Length then false
    else
        let mutable same = true
        let mutable k = 0
        while same && k < value.Length do
            if s.[at + k] <> value.[k] then same <- false
            k <- k + 1
        same

let private decodeReferences (raw: string) : string =
    let first = raw.IndexOf '&'
    if first < 0 then raw
    else
        let sb = System.Text.StringBuilder()
        let mutable i = 0
        while i < raw.Length do
            let amp = raw.IndexOf('&', i)
            if amp < 0 then
                sb.Append(raw.Substring i) |> ignore
                i <- raw.Length
            else
                if amp > i then sb.Append(raw.Substring(i, amp - i)) |> ignore
                let semi = raw.IndexOf(';', amp)
                if semi < 0 || semi - amp > 10 || semi - amp < 2 then scanGiveUp ()
                let entity = raw.Substring(amp + 1, semi - amp - 1)
                match entity with
                | "lt" -> sb.Append "<" |> ignore
                | "gt" -> sb.Append ">" |> ignore
                | "amp" -> sb.Append "&" |> ignore
                | "quot" -> sb.Append "\"" |> ignore
                | "apos" -> sb.Append "'" |> ignore
                | _ when entity.[0] = '#' ->
                    let mutable code = 0
                    let hex = entity.Length > 1 && entity.[1] = 'x'
                    let start = if hex then 2 else 1
                    if entity.Length <= start then scanGiveUp ()
                    for k in start .. entity.Length - 1 do
                        let c = int entity.[k]
                        let digit =
                            if c >= 48 && c <= 57 then c - 48
                            elif hex && c >= 97 && c <= 102 then c - 87
                            elif hex && c >= 65 && c <= 70 then c - 55
                            else scanGiveUp ()
                        code <- code * (if hex then 16 else 10) + digit
                        if code > 0x10FFFF then scanGiveUp ()
                    // (Characters XML does not allow are an error of the part: the platform's parser reports it.)
                    if code = 0 || (code < 32 && code <> 9 && code <> 10 && code <> 13) || (code >= 0xD800 && code <= 0xDFFF) || code = 0xFFFE || code = 0xFFFF then scanGiveUp ()
                    if code < 0x10000 then sb.Append(string (char code)) |> ignore
                    else
                        let rest = code - 0x10000
                        sb.Append(string (char (0xD800 + (rest >>> 10)))) |> ignore
                        sb.Append(string (char (0xDC00 + (rest &&& 0x3FF)))) |> ignore
                | _ -> scanGiveUp ()
                i <- semi + 1
        sb.ToString()

let private normaliseLineEnds (raw: string) =
    if raw.IndexOf '\r' >= 0 then raw.Replace("\r\n", "\n").Replace("\r", "\n") else raw

/// The scanner works on the part's text as it is - line ends are normalised
/// in the values it reads (XML 1.0, 2.11), not in the text - so the span of
/// every element is its place in the part.
let private scanRange (s: string) (from: int) (until: int) (names: System.Collections.Generic.HashSet<string>) : XmlNode =
    let n = until
    let mutable pos = from
    // (A byte order mark or a zero-width space before the declaration is not text.)
    while pos < n && (s.[pos] = char 0xFEFF || s.[pos] = char 0x200B) do pos <- pos + 1
    let isSpace (c: char) = c = ' ' || c = '\n' || c = '\t' || c = '\r'
    // The open elements, innermost last: name, attributes, start, end of start tag, whether text
    // keeps its whitespace, whether descendants inherit that, children (last first).
    let openNames = ResizeArray<string>()
    let openAttrs = ResizeArray<Map<string, string>>()
    let openStarts = ResizeArray<int>()
    let openTagEnds = ResizeArray<int>()
    let openPreserve = ResizeArray<bool>()
    let openInherited = ResizeArray<bool>()
    let openChildren = ResizeArray<XmlNode list>()
    let mutable root = Empty
    let mutable rootSeen = false
    let addChild (node: XmlNode) =
        let top = openChildren.Count - 1
        openChildren.[top] <- node :: openChildren.[top]
    let closeElement (name: string) (attrs: Map<string, string>) (reversed: XmlNode list) (span: Span) =
        let mutable children = []
        let mutable rest = reversed
        while not rest.IsEmpty do
            children <- rest.Head :: children
            rest <- rest.Tail
        let node = Element (name, attrs, children, span)
        if openNames.Count = 0 then
            root <- node
            rootSeen <- true
        else addChild node
    let addText (start: int) (finish: int) =
        if finish > start then
            if openNames.Count = 0 then
                // Outside the root element only whitespace may stand.
                if s.Substring(start, finish - start).Trim().Length > 0 then scanGiveUp ()
            else
                let preserve = openPreserve.[openPreserve.Count - 1]
                // (Whitespace between elements of a part that does not keep it is not text: not even decoded.)
                let mutable blank = true
                let mutable i = start
                while blank && i < finish do
                    if not (isSpace s.[i]) then blank <- false
                    i <- i + 1
                if preserve || not blank then
                    let text = decodeReferences (normaliseLineEnds (s.Substring(start, finish - start)))
                    if preserve then addChild (Text text)
                    else
                        let trimmed = trimXml text
                        if trimmed.Length > 0 then addChild (Text trimmed)
    while pos < n do
        let lt = (let at = s.IndexOf('<', pos) in if at >= n then -1 else at)
        if lt < 0 then
            addText pos n
            pos <- n
        else
            addText pos lt
            if lt + 1 >= n then scanGiveUp ()
            let c = s.[lt + 1]
            if c = '/' then
                let gt = s.IndexOf('>', lt + 2)
                if gt < 0 || gt >= n || openNames.Count = 0 then scanGiveUp ()
                let mutable nameEnd = gt
                while nameEnd > lt + 2 && isSpace s.[nameEnd - 1] do nameEnd <- nameEnd - 1
                let top = openNames.Count - 1
                let name = openNames.[top]
                if nameEnd - (lt + 2) <> name.Length || not (startsAt s (lt + 2) name) then scanGiveUp ()
                let attrs = openAttrs.[top]
                let children = openChildren.[top]
                let span = { Start = openStarts.[top]; StartTagEnd = openTagEnds.[top]; EndTagStart = lt; End = gt + 1 }
                openNames.RemoveAt top
                openAttrs.RemoveAt top
                openStarts.RemoveAt top
                openTagEnds.RemoveAt top
                openPreserve.RemoveAt top
                openInherited.RemoveAt top
                openChildren.RemoveAt top
                closeElement name attrs children span
                pos <- gt + 1
            elif c = '?' then
                let close = indexOfOrdinal s "?>" (lt + 2)
                if close < 0 || close >= n then scanGiveUp ()
                pos <- close + 2
            elif c = '!' then
                if startsAt s lt "<!--" then
                    let close = indexOfOrdinal s "-->" (lt + 4)
                    if close < 0 || close >= n then scanGiveUp ()
                    pos <- close + 3
                elif startsAt s lt "<![CDATA[" then
                    let close = indexOfOrdinal s "]]>" (lt + 9)
                    if close < 0 || close >= n || openNames.Count = 0 then scanGiveUp ()
                    if close > lt + 9 then addChild (Text (normaliseLineEnds (s.Substring(lt + 9, close - lt - 9))))
                    pos <- close + 3
                else scanGiveUp ()
            else
                if rootSeen then scanGiveUp ()
                // <name attribute="value" ...> or .../>
                let mutable i = lt + 1
                while i < n && (let ch = s.[i] in not (isSpace ch) && ch <> '>' && ch <> '/') do i <- i + 1
                if i >= n || i = lt + 1 then scanGiveUp ()
                let name = s.Substring(lt + 1, i - lt - 1)
                names.Add name |> ignore
                let mutable attrs : Map<string, string> = Map.empty
                let mutable space : string option = None
                let mutable selfClosing = false
                let mutable inTag = true
                while inTag do
                    while i < n && isSpace s.[i] do i <- i + 1
                    if i >= n then scanGiveUp ()
                    let ch = s.[i]
                    if ch = '>' then
                        inTag <- false
                        i <- i + 1
                    elif ch = '/' then
                        if i + 1 >= n || s.[i + 1] <> '>' then scanGiveUp ()
                        selfClosing <- true
                        inTag <- false
                        i <- i + 2
                    else
                        let nameStart = i
                        while i < n && (let a = s.[i] in a <> '=' && not (isSpace a) && a <> '>' && a <> '/') do i <- i + 1
                        let attrName = s.Substring(nameStart, i - nameStart)
                        while i < n && isSpace s.[i] do i <- i + 1
                        if i >= n || s.[i] <> '=' || attrName.Length = 0 then scanGiveUp ()
                        i <- i + 1
                        while i < n && isSpace s.[i] do i <- i + 1
                        if i >= n then scanGiveUp ()
                        let quote = s.[i]
                        if quote <> '"' && quote <> '\'' then scanGiveUp ()
                        let close = s.IndexOf(quote, i + 1)
                        if close < 0 || close >= n then scanGiveUp ()
                        let raw = s.Substring(i + 1, close - i - 1)
                        if raw.IndexOf '<' >= 0 then scanGiveUp ()
                        // Literal tabs and line ends of an attribute value are spaces (XML 1.0, 3.3.3).
                        let raw = normaliseLineEnds raw
                        let raw = if raw.IndexOf '\n' >= 0 || raw.IndexOf '\t' >= 0 then raw.Replace("\n", " ").Replace("\t", " ") else raw
                        let value = decodeReferences raw
                        if attrs.ContainsKey attrName then scanGiveUp ()
                        if attrName = "xml:space" then space <- Some value
                        attrs <- Map.add attrName value attrs
                        i <- close + 1
                if selfClosing then closeElement name attrs [] { Start = lt; StartTagEnd = i; EndTagStart = i; End = i }
                else
                    openNames.Add name
                    openAttrs.Add attrs
                    openStarts.Add lt
                    openTagEnds.Add i
                    let inherited = openInherited.Count > 0 && openInherited.[openInherited.Count - 1]
                    openPreserve.Add (keepsSpace name space inherited)
                    openInherited.Add (spaceInherited space inherited)
                    openChildren.Add []
                pos <- i
    if openNames.Count > 0 || not rootSeen then scanGiveUp ()
    root

let private scanDocument (s: string) (names: System.Collections.Generic.HashSet<string>) : XmlNode =
    scanRange s 0 s.Length names

/// `parse` by the scanner above: Some node, with the element names of the
/// part (as written, prefix included) added to `names`; None when the part
/// holds something the scanner does not read - `parse` then says what.
/// `tryScan` of one element of a part, the text from `start` to `finish`, with the spans of its
/// elements in the part's text.
let tryScanAt (xmlString: string) (start: int) (finish: int) (names: System.Collections.Generic.HashSet<string>) : XmlNode option =
    try Some (scanRange xmlString start finish names)
    with _ -> None

let tryScan (xmlString: string) (names: System.Collections.Generic.HashSet<string>) : XmlNode option =
    try Some (scanDocument xmlString names)
    with _ ->
        names.Clear()
        None

// Namespace-aware helpers for Office Open XML
module Ns =
    let stripNs (name: string) =
        match name.LastIndexOf(':') with
        | -1 -> name
        | idx -> name.Substring(idx + 1)

    /// `stripNs name = localName`, without making the stripped name.
    /// (A local name holds no colon: a name that ends with ":" + localName has it after its last colon.)
    let private hasLocalName (localName: string) (name: string) =
        let extra = name.Length - localName.Length
        if extra = 0 then name = localName
        elif extra < 2 then false
        else name.[extra - 1] = ':' && name.EndsWith(localName, System.StringComparison.Ordinal)

    let childNs localName node =
        Xml.children node
        |> List.tryFind (fun n ->
            match n with
            | Element (name, _, _, _) -> hasLocalName localName name
            | _ -> false)

    /// The elements under (and including) a node that have a local name, in
    /// document order.
    let descendantsNs localName node =
        let found = ResizeArray<XmlNode>()
        let rec visit (current: XmlNode) =
            match current with
            | Element (name, _, kids, _) ->
                if hasLocalName localName name then found.Add current
                for kid in kids do
                    visit kid
            | _ -> ()
        visit node
        let mutable result = []
        for i = found.Count - 1 downto 0 do
            result <- found.[i] :: result
        result

    /// A large part parsed with the local names of its elements: asking whether
    /// the part holds an element at all costs nothing, where `descendantsNs`
    /// walks the whole tree to say there is none.
    let parseWithLocalNames (xmlString: string) : XmlNode * System.Collections.Generic.HashSet<string> =
        let names = System.Collections.Generic.HashSet<string>()
        let node =
            match tryScan xmlString names with
            | Some node -> node
            | None ->
                let node = parse xmlString
                let rec visit (current: XmlNode) =
                    match current with
                    | Element (name, _, kids, _) ->
                        names.Add name |> ignore
                        for kid in kids do visit kid
                    | _ -> ()
                visit node
                node
        let local = System.Collections.Generic.HashSet<string>()
        for name in names do
            local.Add (stripNs name) |> ignore
        node, local
