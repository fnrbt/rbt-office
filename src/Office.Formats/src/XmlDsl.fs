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

/// The tree without the places of its elements (to compare trees read by different parsers).
let rec withoutSpans (node: XmlNode) : XmlNode =
    match node with
    | Element (name, attrs, children, _) -> Element (name, attrs, List.map withoutSpans children, Span.none)
    | other -> other

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

/// The code of the character at `at` (in JavaScript a character of a string is a string: the scanner
/// compares codes).
#if FABLE_COMPILER
[<Emit("$0.charCodeAt($1)")>]
let private codeAt (text: string) (at: int) : int = jsNative
#else
let inline private codeAt (text: string) (at: int) : int = int text.[at]
#endif

/// XML white space: space, tab, line feed, carriage return.
let inline private isSpaceCode (c: int) = c = 32 || c = 10 || c = 9 || c = 13

/// Text without its XML whitespace: a no-break or an en space is a character of the text.
let private trimXml (text: string) =
    let mutable first = 0
    let mutable last = text.Length - 1
    while first <= last && isSpaceCode (codeAt text first) do first <- first + 1
    while last >= first && isSpaceCode (codeAt text last) do last <- last - 1
    if first = 0 && last = text.Length - 1 then text else text.Substring(first, last - first + 1)

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

// ---- The lexer ----
// The one place the text of XML is taken apart: the scanner below (`tryScan`, `tryScanAt`), the readers that
// go through a large part element by element (a long document's body, a sheet's rows) and the outline reads
// all stand on it. It reads the text as it is - line ends are normalised in the values it gives (XML 1.0,
// 2.11), not in the text -, so the span of every element is its place in the part.

/// What an XmlLexer stands on.
type XmlToken =
    /// A start tag, or an empty-element tag (`SelfClosing`).
    | StartTag = 0
    | EndTag = 1
    /// Character data between tags, as it is written (`TextValue` reads it).
    | Text = 2
    | CData = 3
    | Comment = 4
    /// A processing instruction, or the XML declaration.
    | Instruction = 5
    | EndOfText = 6
    /// What the lexer does not read - a document type declaration, a tag that does not end, an end tag that
    /// does not close the element that is open -: it reads nothing after it (but from where `Reset` puts it).
    | Malformed = 7

/// Reads the tokens of XML text from `from` to `until`, one at a time (`Next`). An end tag must close the
/// element that is open (`Nesting`): the lexer keeps the names of the open elements, and `Depth` is how many
/// are open. What a token holds is read off it while the lexer stands on it - its name, its attributes (as
/// written, or decoded), its text - without making strings of what is not asked for.
[<Sealed>]
type XmlLexer(text: string, from: int, until: int) =
    let mutable pos = from
    let mutable token = XmlToken.EndOfText
    let mutable tokenStart = from
    let mutable tokenEnd = from
    let mutable nameStart = 0
    let mutable nameEnd = 0
    let mutable contentStart = 0
    let mutable contentEnd = 0
    let mutable selfClosing = false
    // Per attribute: where its name starts and ends, where its value starts and ends (inside the quotes).
    let mutable attrCount = 0
    let mutable attrBounds : int[] = Array.zeroCreate 64
    // Per open element: where its name starts in the text, and its length.
    let mutable depth = 0
    let mutable openNames : int[] = Array.zeroCreate 64
    let mutable nesting = true
    // Where the next "<!" and "<?" are (SeekStartTag): -2 not looked for, -1 none.
    let mutable commentAt = -2
    let mutable instructionAt = -2

    let malformed () =
        token <- XmlToken.Malformed
        tokenStart <- pos
        tokenEnd <- pos
        token

    let sameText (at: int) (other: int) (length: int) =
        let mutable same = true
        let mutable k = 0
        while same && k < length do
            if codeAt text (at + k) <> codeAt text (other + k) then same <- false
            k <- k + 1
        same

    let nameEquals (start: int) (finish: int) (name: string) =
        if finish - start <> name.Length then false
        else
            let mutable same = true
            let mutable k = 0
            while same && k < name.Length do
                if codeAt text (start + k) <> codeAt name k then same <- false
                k <- k + 1
            same

    let localNameEquals (start: int) (finish: int) (localName: string) =
        let length = finish - start
        if length = localName.Length then nameEquals start finish localName
        elif length > localName.Length + 1 && codeAt text (finish - localName.Length - 1) = 58 then nameEquals (finish - localName.Length) finish localName
        else false

    let addAttribute (a: int) (b: int) (c: int) (d: int) =
        let at = attrCount * 4
        if at + 4 > attrBounds.Length then
            let bigger : int[] = Array.zeroCreate (attrBounds.Length * 2)
            Array.blit attrBounds 0 bigger 0 attrBounds.Length
            attrBounds <- bigger
        attrBounds.[at] <- a
        attrBounds.[at + 1] <- b
        attrBounds.[at + 2] <- c
        attrBounds.[at + 3] <- d
        attrCount <- attrCount + 1

    let push (start: int) (length: int) =
        let at = depth * 2
        if at + 2 > openNames.Length then
            let bigger : int[] = Array.zeroCreate (openNames.Length * 2)
            Array.blit openNames 0 bigger 0 openNames.Length
            openNames <- bigger
        openNames.[at] <- start
        openNames.[at + 1] <- length
        depth <- depth + 1

    // <name attribute="value" ...> or .../>
    let startTag (lt: int) =
        let mutable i = lt + 1
        while i < until && (let ch = codeAt text (i) in not (isSpaceCode ch) && ch <> 62 && ch <> 47) do i <- i + 1
        if i >= until || i = lt + 1 then malformed ()
        else
            nameStart <- lt + 1
            nameEnd <- i
            attrCount <- 0
            selfClosing <- false
            let mutable ok = true
            let mutable inTag = true
            // (The first '<' after the tag's: a value holds none.)
            let mutable nextLt = -2
            while ok && inTag do
                while i < until && isSpaceCode (codeAt text (i)) do i <- i + 1
                if i >= until then ok <- false
                else
                    let ch = codeAt text (i)
                    if ch = 62 then
                        inTag <- false
                        i <- i + 1
                    elif ch = 47 then
                        if i + 1 >= until || codeAt text (i + 1) <> 62 then ok <- false
                        else
                            selfClosing <- true
                            inTag <- false
                            i <- i + 2
                    else
                        let attrNameStart = i
                        while i < until && (let a = codeAt text (i) in a <> 61 && not (isSpaceCode a) && a <> 62 && a <> 47) do i <- i + 1
                        let attrNameEnd = i
                        while i < until && isSpaceCode (codeAt text (i)) do i <- i + 1
                        if i >= until || codeAt text (i) <> 61 || attrNameEnd = attrNameStart then ok <- false
                        else
                            i <- i + 1
                            while i < until && isSpaceCode (codeAt text (i)) do i <- i + 1
                            if i >= until then ok <- false
                            else
                                let quote = text.[i]
                                let quoteCode = int quote
                                if quoteCode <> 34 && quoteCode <> 39 then ok <- false
                                else
                                    let close = text.IndexOf(quote, i + 1)
                                    if close < 0 || close >= until then ok <- false
                                    else
                                        if nextLt = -2 then nextLt <- text.IndexOf('<', lt + 1)
                                        if nextLt >= 0 && nextLt < close then ok <- false
                                        else
                                            addAttribute attrNameStart attrNameEnd (i + 1) close
                                            i <- close + 1
            if not ok then malformed ()
            else
                token <- XmlToken.StartTag
                tokenStart <- lt
                tokenEnd <- i
                pos <- i
                if nesting && not selfClosing then push nameStart (nameEnd - nameStart)
                token

    let endTag (lt: int) =
        let gt = text.IndexOf('>', lt + 2)
        if gt < 0 || gt >= until then malformed ()
        else
            let mutable finish = gt
            while finish > lt + 2 && isSpaceCode (codeAt text (finish - 1)) do finish <- finish - 1
            let matches =
                if not nesting then true
                elif depth = 0 then false
                else
                    let top = (depth - 1) * 2
                    finish - (lt + 2) = openNames.[top + 1] && sameText (lt + 2) openNames.[top] openNames.[top + 1]
            if not matches then malformed ()
            else
                if nesting then depth <- depth - 1
                token <- XmlToken.EndTag
                nameStart <- lt + 2
                nameEnd <- finish
                attrCount <- 0
                selfClosing <- false
                tokenStart <- lt
                tokenEnd <- gt + 1
                pos <- gt + 1
                token

    let between (kind: XmlToken) (lt: int) (opening: int) (closing: string) =
        let close = indexOfOrdinal text closing (lt + opening)
        if close < 0 || close + closing.Length > until then malformed ()
        else
            token <- kind
            tokenStart <- lt
            tokenEnd <- close + closing.Length
            contentStart <- lt + opening
            contentEnd <- close
            pos <- tokenEnd
            token

    let markup (lt: int) =
        if lt + 1 >= until then malformed ()
        else
            let c = codeAt text (lt + 1)
            if c = 47 then endTag lt
            elif c = 63 then between XmlToken.Instruction lt 2 "?>"
            elif c = 33 then
                if startsAt text lt "<!--" then between XmlToken.Comment lt 4 "-->"
                elif startsAt text lt "<![CDATA[" then between XmlToken.CData lt 9 "]]>"
                else malformed ()
            else startTag lt

    /// The text read.
    member _.Source = text
    member _.Until = until
    /// The token the lexer stands on.
    member _.Token = token
    /// Where the token starts (its '<') and where it ends (after its '>').
    member _.Start = tokenStart
    member _.End = tokenEnd
    /// Where the next token is read from.
    member _.Position = pos
    /// How many elements are open (a start tag counts once it is read).
    member _.Depth = depth
    member _.SelfClosing = selfClosing
    /// Whether end tags are matched with the start tags (true): without, the lexer reads tags one after
    /// another, whatever they nest like.
    member _.Nesting with get () = nesting and set (value: bool) = nesting <- value

    /// The qualified name of the start or end tag.
    member _.Name = text.Substring(nameStart, nameEnd - nameStart)
    member _.NameIs (name: string) = nameEquals nameStart nameEnd name
    /// Whether the tag's name is the local name, with a prefix or without.
    member _.LocalNameIs (localName: string) = localNameEquals nameStart nameEnd localName

    member _.AttributeCount = attrCount
    member _.AttributeName (index: int) = text.Substring(attrBounds.[index * 4], attrBounds.[index * 4 + 1] - attrBounds.[index * 4])
    member _.AttributeNameIs (index: int, name: string) = nameEquals attrBounds.[index * 4] attrBounds.[index * 4 + 1] name
    /// The value as it is read: line ends, tabs and line feeds as spaces (XML 1.0, 3.3.3), references
    /// decoded. (A reference that is none: the scanner gives up.)
    member _.AttributeValue (index: int) =
        let start = attrBounds.[index * 4 + 2]
        let finish = attrBounds.[index * 4 + 3]
        // (Most values hold no line end, tab or reference: they are their text.)
        let mutable plain = true
        let mutable i = start
        while plain && i < finish do
            let c = codeAt text i
            if c = 38 || c = 13 || c = 10 || c = 9 then plain <- false
            i <- i + 1
        let raw = text.Substring(start, finish - start)
        if plain then raw
        else
            let raw = normaliseLineEnds raw
            let raw = if raw.IndexOf '\n' >= 0 || raw.IndexOf '\t' >= 0 then raw.Replace("\n", " ").Replace("\t", " ") else raw
            decodeReferences raw
    /// The value of the attribute of that qualified name, or null.
    member this.Attribute (name: string) : string =
        let mutable found : string = null
        let mutable k = 0
        while k < attrCount do
            if nameEquals attrBounds.[k * 4] attrBounds.[k * 4 + 1] name then
                found <- this.AttributeValue k
                k <- attrCount
            else k <- k + 1
        found
    /// The values of the attributes `names` into `values` (null for one the tag does not have), in one pass.
    member this.ReadAttributes (names: string[], values: string[]) =
        for k in 0 .. values.Length - 1 do
            values.[k] <- null
        for a in 0 .. attrCount - 1 do
            let start = attrBounds.[a * 4]
            let finish = attrBounds.[a * 4 + 1]
            let mutable k = 0
            while k < names.Length do
                if nameEquals start finish names.[k] then
                    values.[k] <- this.AttributeValue a
                    k <- names.Length
                else k <- k + 1
    /// The attributes of the start tag (an attribute twice: the scanner gives up).
    member this.Attributes () : Map<string, string> =
        let mutable attrs : Map<string, string> = Map.empty
        for a in 0 .. attrCount - 1 do
            let name = this.AttributeName a
            if attrs.ContainsKey name then scanGiveUp ()
            attrs <- Map.add name (this.AttributeValue a) attrs
        attrs

    /// Whether the text (or CDATA) is XML white space alone.
    member _.TextIsBlank =
        let mutable blank = true
        let mutable i = contentStart
        while blank && i < contentEnd do
            if not (isSpaceCode (codeAt text (i))) then blank <- false
            i <- i + 1
        blank
    /// Whether the text is white space alone, as a trim says (what may stand outside the root).
    member _.TextIsWhiteSpace = text.Substring(contentStart, contentEnd - contentStart).Trim().Length = 0
    member _.TextLength = contentEnd - contentStart
    /// The text as it is read: line ends normalised, references decoded (CDATA: as it is).
    member _.TextValue () : string =
        let raw = normaliseLineEnds (text.Substring(contentStart, contentEnd - contentStart))
        if token = XmlToken.CData then raw else decodeReferences raw

    /// The text as it is read, without the XML white space at its ends.
    member this.TextValueTrimmed () : string = trimXml (this.TextValue ())

    /// The next token.
    member _.Next () : XmlToken =
        if token = XmlToken.Malformed then token
        elif pos >= until then
            token <- XmlToken.EndOfText
            tokenStart <- until
            tokenEnd <- until
            token
        elif codeAt text (pos) <> 60 then
            let lt = text.IndexOf('<', pos)
            let finish = if lt < 0 || lt >= until then until else lt
            token <- XmlToken.Text
            tokenStart <- pos
            tokenEnd <- finish
            contentStart <- pos
            contentEnd <- finish
            pos <- finish
            token
        else markup pos

    /// From a start tag: on to the end tag of its element (the lexer stands on it); where the element ends,
    /// or -1 when the text is not well formed there.
    member this.SkipElement () : int =
        if token <> XmlToken.StartTag then -1
        elif selfClosing then tokenEnd
        else
            let outer = depth - 1
            let mutable result = -2
            while result = -2 do
                match this.Next () with
                | XmlToken.EndTag when depth = outer -> result <- tokenEnd
                | XmlToken.Malformed | XmlToken.EndOfText -> result <- -1
                | _ -> ()
            result

    /// From a start tag: on to the end tag of its element (the lexer stands on it), found by its name - past
    /// the elements of that name in it, each to its own - without reading what is between: for an element
    /// of many (a sheet's rows) that is passed over. (What it holds must not be a comment or CDATA with a tag
    /// of that name.) Where the element ends, or -1 when the text is not well formed there.
    member this.SkipContent () : int =
        if token <> XmlToken.StartTag then -1
        elif selfClosing then tokenEnd
        else
            let name = text.Substring(nameStart, nameEnd - nameStart)
            let opening = "<" + name
            let closing = "</" + name
            let isNameEnd (at: int) = at < until && (let c = codeAt text (at) in isSpaceCode c || c = 62 || c = 47)
            let mutable open' = 1
            let mutable at = tokenEnd
            let mutable close = -1
            // (Where the next start tag of the name is: looked for again once it is passed.)
            let mutable inner = -2
            while close = -1 do
                let ending = indexOfOrdinal text closing at
                if ending < 0 || ending >= until then close <- -2
                else
                    if inner <> -1 && inner < at then inner <- indexOfOrdinal text opening at
                    if inner >= 0 && inner < ending then
                        if isNameEnd (inner + opening.Length) then
                            let tag = XmlLexer(text, inner, until)
                            tag.Nesting <- false
                            if tag.Next () <> XmlToken.StartTag then close <- -2
                            else
                                if not tag.SelfClosing then open' <- open' + 1
                                at <- tag.End
                        else at <- inner + opening.Length
                    elif isNameEnd (ending + closing.Length) then
                        open' <- open' - 1
                        if open' = 0 then close <- ending
                        at <- ending + closing.Length
                    else at <- ending + closing.Length
            if close < 0 then
                pos <- tokenEnd
                malformed () |> ignore
                -1
            else
                pos <- close
                if this.Next () = XmlToken.EndTag then tokenEnd else -1

    /// On to the next start tag of that qualified name (the lexer stands on it), passing over what comes
    /// before without reading it - but for comments, CDATA and processing instructions, which are read past -
    /// in a lexer without `Nesting`; false at the end of the text.
    member this.SeekStartTag (name: string) : bool =
        let pattern = "<" + name
        let mutable found = false
        let mutable going = not nesting && token <> XmlToken.Malformed
        while going do
            let at = indexOfOrdinal text pattern pos
            if at < 0 || at >= until then
                pos <- until
                this.Next () |> ignore
                going <- false
            else
                if commentAt <> -1 && commentAt < pos then commentAt <- indexOfOrdinal text "<!" pos
                if instructionAt <> -1 && instructionAt < pos then instructionAt <- indexOfOrdinal text "<?" pos
                let other =
                    if commentAt >= 0 && commentAt < at then commentAt
                    elif instructionAt >= 0 && instructionAt < at then instructionAt
                    else -1
                if other >= 0 then
                    if markup other = XmlToken.Malformed then going <- false
                elif at + pattern.Length < until && (let c = codeAt text (at + pattern.Length) in isSpaceCode c || c = 62 || c = 47) then
                    if markup at = XmlToken.StartTag then found <- true
                    going <- false
                else pos <- at + 1
        found

    /// Read on from `at`, with no element open (and what was malformed forgotten).
    member _.Reset (at: int) =
        pos <- at
        depth <- 0
        token <- XmlToken.EndOfText
        commentAt <- -2
        instructionAt <- -2
    /// Read on from `at` - the start tag of an element read before - with the elements open that were open
    /// then (`depth` of them): to read an element again.
    member _.Rewind (at: int, openDepth: int) =
        pos <- at
        depth <- openDepth
        token <- XmlToken.EndOfText
        commentAt <- -2
        instructionAt <- -2

// ---- The scanner ----
// The text of a part straight to XmlNode. `parse` goes through the platform's parser and then walks its tree;
// on a part of several megabytes that is most of the time a reader takes. The scanner reads what Office
// writes - elements, attributes, text, CDATA, comments, processing instructions and the predefined / numeric
// character references - and gives up (None) on anything else (a DTD, an unknown entity, tags that do not
// match), which the caller then hands to `parse`.


let private localNameOf (name: string) =
    match name.LastIndexOf ':' with
    | -1 -> name
    | at -> name.Substring(at + 1)

/// The element whose start tag the lexer stands on, read to its end tag (the lexer then stands on that).
/// `inherited`: the `xml:space` of its ancestors. The elements whose local name `shallow` says are read
/// without what they hold (their names, attributes and spans are in the tree).
let private readElement (lexer: XmlLexer) (names: System.Collections.Generic.HashSet<string>) (shallow: (string -> bool) option) (keep: (string -> bool) option) (inherited: bool) : XmlNode =
    // The open elements, innermost last: name, attributes, start, end of start tag, whether text
    // keeps its whitespace, whether descendants inherit that, children (last first).
    let openNames = ResizeArray<string>()
    let openAttrs = ResizeArray<Map<string, string>>()
    let openStarts = ResizeArray<int>()
    let openTagEnds = ResizeArray<int>()
    let openPreserve = ResizeArray<bool>()
    let openInherited = ResizeArray<bool>()
    let openChildren = ResizeArray<XmlNode list>()
    let mutable result = Empty
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
        if openNames.Count = 0 then result <- node
        else addChild node
    let openElement () =
        let name = lexer.Name
        names.Add name |> ignore
        let attrs = lexer.Attributes ()
        let start = lexer.Start
        let tagEnd = lexer.End
        if lexer.SelfClosing then closeElement name attrs [] { Start = start; StartTagEnd = tagEnd; EndTagStart = tagEnd; End = tagEnd }
        elif (match shallow with Some isShallow -> isShallow (localNameOf name) | None -> false) then
            let finish = lexer.SkipContent ()
            if finish < 0 then scanGiveUp ()
            closeElement name attrs [] { Start = start; StartTagEnd = tagEnd; EndTagStart = lexer.Start; End = finish }
        else
            let inherited = if openInherited.Count > 0 then openInherited.[openInherited.Count - 1] else inherited
            let space = Map.tryFind "xml:space" attrs
            openNames.Add name
            openAttrs.Add attrs
            openStarts.Add start
            openTagEnds.Add tagEnd
            openPreserve.Add (keepsSpace name space inherited || (match keep with Some keeps -> keeps name | None -> false))
            openInherited.Add (spaceInherited space inherited)
            openChildren.Add []
    openElement ()
    while openNames.Count > 0 do
        match lexer.Next () with
        | XmlToken.Text ->
            let preserve = openPreserve.[openPreserve.Count - 1]
            // (Whitespace between elements of a part that does not keep it is not text: not even decoded.)
            if preserve || not lexer.TextIsBlank then
                let text = lexer.TextValue ()
                if preserve then addChild (Text text)
                else
                    let trimmed = trimXml text
                    if trimmed.Length > 0 then addChild (Text trimmed)
        | XmlToken.CData -> if lexer.TextLength > 0 then addChild (Text (lexer.TextValue ()))
        | XmlToken.Comment | XmlToken.Instruction -> ()
        | XmlToken.StartTag -> openElement ()
        | XmlToken.EndTag ->
            let top = openNames.Count - 1
            let name = openNames.[top]
            let attrs = openAttrs.[top]
            let children = openChildren.[top]
            let span = { Start = openStarts.[top]; StartTagEnd = openTagEnds.[top]; EndTagStart = lexer.Start; End = lexer.End }
            Portable.Stack.dropLast openNames
            Portable.Stack.dropLast openAttrs
            Portable.Stack.dropLast openStarts
            Portable.Stack.dropLast openTagEnds
            Portable.Stack.dropLast openPreserve
            Portable.Stack.dropLast openInherited
            Portable.Stack.dropLast openChildren
            closeElement name attrs children span
        | _ -> scanGiveUp ()
    result

/// The one element of the text from `from` to `until` (white space, comments and processing instructions
/// around it), read by `readElement`.
let private scanRange (s: string) (from: int) (until: int) (names: System.Collections.Generic.HashSet<string>) (shallow: (string -> bool) option) (keep: (string -> bool) option) : XmlNode =
    let mutable start = from
    // (A byte order mark or a zero-width space before the declaration is not text.)
    while start < until && (s.[start] = char 0xFEFF || s.[start] = char 0x200B) do start <- start + 1
    let lexer = XmlLexer(s, start, until)
    let mutable root = Empty
    let mutable rootSeen = false
    let mutable going = true
    while going do
        match lexer.Next () with
        // (Outside the root element only whitespace may stand.)
        | XmlToken.Text -> if not lexer.TextIsWhiteSpace then scanGiveUp ()
        | XmlToken.Comment | XmlToken.Instruction -> ()
        | XmlToken.StartTag ->
            if rootSeen then scanGiveUp ()
            root <- readElement lexer names shallow keep false
            rootSeen <- true
        | XmlToken.EndOfText -> going <- false
        | _ -> scanGiveUp ()
    if not rootSeen then scanGiveUp ()
    root

/// `parse` by the scanner: Some node, with the element names of the part (as written, prefix included) added
/// to `names`; None when the part holds something the scanner does not read - `parse` then says what.
let tryScan (xmlString: string) (names: System.Collections.Generic.HashSet<string>) : XmlNode option =
    try Some (scanRange xmlString 0 xmlString.Length names None None)
    with _ ->
        names.Clear()
        None

/// `tryScan` of one element of a part, the text from `start` to `finish`, with the spans of its elements in
/// the part's text.
let tryScanAt (xmlString: string) (start: int) (finish: int) (names: System.Collections.Generic.HashSet<string>) : XmlNode option =
    try Some (scanRange xmlString start finish names None None)
    with _ -> None

/// A part read in outline: the elements whose local name `shallow` says without what they hold (their names,
/// attributes and spans are in the tree), the rest as `tryScan` reads it.
let tryScanOutline (shallow: string -> bool) (xmlString: string) : XmlNode option =
    try Some (scanRange xmlString 0 xmlString.Length (System.Collections.Generic.HashSet<string>()) (Some shallow) None)
    with _ -> None

/// `tryScan` with the text of the elements whose qualified name `keeps` says kept as it is written, white
/// space and all, as `xml:space="preserve"` keeps it (a drawing's `a:t`, whose spaces at the ends are text).
let tryScanKeeping (keeps: string -> bool) (xmlString: string) : XmlNode option =
    try Some (scanRange xmlString 0 xmlString.Length (System.Collections.Generic.HashSet<string>()) None (Some keeps))
    with _ -> None

/// The element whose start tag the lexer stands on, read to its end tag (the lexer then stands on that), the
/// names of its elements added to `names`. None when the scanner does not read it: the lexer then stands on
/// the element's end tag - when only a value in it is what the scanner does not read (`parseElementOf` reads
/// it) - or is Malformed.
let tryReadElement (lexer: XmlLexer) (names: System.Collections.Generic.HashSet<string>) : XmlNode option =
    if lexer.Token <> XmlToken.StartTag then None
    else
        let start = lexer.Start
        let outer = if lexer.SelfClosing then lexer.Depth else lexer.Depth - 1
        try Some (readElement lexer names None None false)
        with _ ->
            // (On to its end tag.)
            let mutable going = not (lexer.Token = XmlToken.StartTag && lexer.Start = start && lexer.SelfClosing)
            while going do
                match lexer.Token with
                | XmlToken.EndTag when lexer.Depth = outer -> going <- false
                | XmlToken.Malformed | XmlToken.EndOfText -> going <- false
                | _ -> lexer.Next () |> ignore
            None

/// The element that starts at `at` (its start tag's '<'), read to its end tag; None as `tryReadElement`.
let tryReadElementAt (xmlString: string) (at: int) (names: System.Collections.Generic.HashSet<string>) : XmlNode option =
    let lexer = XmlLexer(xmlString, at, xmlString.Length)
    if lexer.Next () = XmlToken.StartTag then tryReadElement lexer names else None

/// Where the element that starts at `at` ends; -1 when the text is not well formed there.
let elementEndAt (xmlString: string) (at: int) : int =
    let lexer = XmlLexer(xmlString, at, xmlString.Length)
    if lexer.Next () = XmlToken.StartTag then lexer.SkipElement () else -1

/// An element of a part - the text from `start` to `finish` - read by the platform's parser in the
/// namespaces its root declares (`root`: the part's root element as the scanner read its start tag), for
/// what the scanner does not read.
let parseElementOf (xmlString: string) (root: XmlNode) (start: int) (finish: int) : XmlNode =
    match root with
    | Element (rootName, _, _, span) when Span.isKnown span ->
        let rootTag = xmlString.Substring(span.Start, span.StartTagEnd - span.Start)
        match parse (rootTag + xmlString.Substring(start, finish - start) + "</" + rootName + ">") with
        | Element (_, _, children, _) -> children |> List.tryFind (function Element _ -> true | _ -> false) |> Option.defaultValue Empty
        | _ -> Empty
    | _ -> Empty

/// The start tag of a part's root element (after the declaration, comments and processing instructions):
/// the root without its children - its span's start and the end of its start tag; the end of the element is
/// not looked for (-1) - and the lexer on it. None for a part that holds no element or a document type
/// declaration, or whose root's attributes the scanner does not read.
let rootStartTag (xmlString: string) : (XmlNode * XmlLexer) option =
    let lexer = XmlLexer(xmlString, 0, xmlString.Length)
    let mutable result = None
    let mutable going = true
    while going do
        match lexer.Next () with
        | XmlToken.Text | XmlToken.Comment | XmlToken.Instruction -> ()
        | XmlToken.StartTag ->
            going <- false
            try result <- Some (Element (lexer.Name, lexer.Attributes (), [], { Start = lexer.Start; StartTagEnd = lexer.End; EndTagStart = -1; End = -1 }), lexer)
            with _ -> ()
        | _ -> going <- false
    result

/// The start tags of a part whose local name is one of `localNames`, in the order of the text: each an
/// element without children - its attributes (none where the scanner does not read them), its span's start
/// and the end of its start tag (the end of the element is not looked for: -1). Tags in comments, CDATA
/// and processing instructions are not among them. The tags are read one after another, however they nest:
/// where the text is not well formed, from the next tag on.
let startTagsNamed (localNames: string list) (xmlString: string) : XmlNode list =
    let wanted = Array.ofList localNames
    let lexer = XmlLexer(xmlString, 0, xmlString.Length)
    lexer.Nesting <- false
    let found = ResizeArray<XmlNode>()
    let mutable going = true
    while going do
        match lexer.Next () with
        | XmlToken.StartTag ->
            let mutable k = 0
            while k < wanted.Length do
                if lexer.LocalNameIs wanted.[k] then
                    let attrs = try lexer.Attributes () with _ -> Map.empty
                    found.Add (Element (lexer.Name, attrs, [], { Start = lexer.Start; StartTagEnd = lexer.End; EndTagStart = -1; End = -1 }))
                    k <- wanted.Length
                else k <- k + 1
        | XmlToken.EndOfText -> going <- false
        | XmlToken.Malformed -> lexer.Reset (lexer.Start + 1)
        | _ -> ()
    List.ofSeq found

// ---- Read back from an end ----
// A reader that needs the last elements of a large part before the rest (the section properties that end a
// document's body) finds where they start from the end of the text; what it finds is read forward again by
// the scanner, which says whether it is an element.

let private isXmlSpace (c: char) = c = ' ' || c = '\n' || c = '\t' || c = '\r'

/// Where `value` last starts before `before` (-1: nowhere).
let private lastStartBefore (s: string) (value: string) (before: int) =
    let mutable at = before - value.Length
    while at >= 0 && not (startsAt s at value) do at <- at - 1
    at

/// Where the white space, comments and processing instructions that end at `finish` start.
let private backOverMisc (s: string) (finish: int) =
    let mutable fin = finish
    let mutable going = true
    while going do
        while fin > 0 && isXmlSpace s.[fin - 1] do fin <- fin - 1
        if fin >= 7 && startsAt s (fin - 3) "-->" then
            let at = lastStartBefore s "<!--" (fin - 3)
            if at < 0 then going <- false else fin <- at
        elif fin >= 4 && startsAt s (fin - 2) "?>" then
            let at = lastStartBefore s "<?" (fin - 2)
            if at < 0 then going <- false else fin <- at
        else going <- false
    fin

/// Where the last start tag of that qualified name before `before` (and at or after `limit`) starts; -1: none.
/// (What is found there is read forward again by a lexer, which says whether it is one.)
let startTagBefore (xmlString: string) (name: string) (before: int) (limit: int) : int =
    let pattern = "<" + name
    let mutable at = before - 1
    let mutable found = -1
    while found < 0 && at >= limit do
        if codeAt xmlString at = 60 && startsAt xmlString at pattern && at + pattern.Length < xmlString.Length
           && (let c = codeAt xmlString (at + pattern.Length) in isSpaceCode c || c = 62 || c = 47) then
            found <- at
        else at <- at - 1
    found

/// The end tag that ends at `finish`, but for white space, comments and processing instructions before
/// `finish`: where it starts, and its name. None when what ends there is not an end tag.
let endTagBefore (xmlString: string) (finish: int) : (int * string) option =
    let fin = backOverMisc xmlString finish
    if fin < 4 || xmlString.[fin - 1] <> '>' then None
    else
        let mutable lt = fin - 2
        while lt >= 0 && xmlString.[lt] <> '<' do lt <- lt - 1
        if lt < 0 || xmlString.[lt + 1] <> '/' then None
        else
            let name = xmlString.Substring(lt + 2, fin - 1 - (lt + 2)).Trim(' ', '\t', '\r', '\n')
            if name = "" then None else Some (lt, name)

/// The element that ends at `finish` (but for white space before it) and starts at or after `limit`, read by
/// the scanner (`tryReadElementAt`); None when no element the scanner reads ends there.
let elementBefore (xmlString: string) (finish: int) (limit: int) (names: System.Collections.Generic.HashSet<string>) : XmlNode option =
    let mutable fin = finish
    while fin > limit && isXmlSpace xmlString.[fin - 1] do fin <- fin - 1
    if fin - limit < 4 || xmlString.[fin - 1] <> '>' then None
    else
        // The '<' of the tag that ends there (an attribute value holds none).
        let mutable lt = fin - 2
        while lt >= limit && xmlString.[lt] <> '<' do lt <- lt - 1
        let start =
            if lt < limit then -1
            elif xmlString.[fin - 2] = '/' then (if xmlString.[lt + 1] <> '/' && xmlString.[lt + 1] <> '!' && xmlString.[lt + 1] <> '?' then lt else -1)
            elif xmlString.[lt + 1] <> '/' then -1
            else
                let name = xmlString.Substring(lt + 2, fin - 1 - (lt + 2)).Trim(' ', '\t', '\r', '\n')
                // Back to the start tag this end tag closes (an element may hold earlier ones of its name).
                let opening = "<" + name
                let closing = "</" + name
                let isNameEnd (at: int) = at < xmlString.Length && (isXmlSpace xmlString.[at] || xmlString.[at] = '>' || xmlString.[at] = '/')
                let mutable depth = 1
                let mutable at = lt - 1
                let mutable found = -1
                while found < 0 && at >= limit do
                    if xmlString.[at] = '<' then
                        if startsAt xmlString at closing && isNameEnd (at + closing.Length) then depth <- depth + 1
                        elif startsAt xmlString at opening && isNameEnd (at + opening.Length) then
                            let tag = XmlLexer(xmlString, at, fin)
                            tag.Nesting <- false
                            if tag.Next () = XmlToken.StartTag && not tag.SelfClosing then
                                depth <- depth - 1
                                if depth = 0 then found <- at
                    at <- at - 1
                found
        if start < 0 then None
        else
            match tryReadElementAt xmlString start names with
            | Some (Element (_, _, _, span) as node) when span.End = fin -> Some node
            | _ -> None

// Namespace-aware helpers for Office Open XML
module Ns =
    let stripNs (name: string) =
        match name.LastIndexOf(':') with
        | -1 -> name
        | idx -> name.Substring(idx + 1)

    /// Whether a qualified name has that prefix ("v:shape": "v").
    let hasPrefix (prefix: string) (name: string) : bool = name.StartsWith(prefix + ":")

    /// The prefix a namespace declaration declares ("xmlns:w" -> Some "w"); None for another attribute.
    let declaredPrefix (attributeName: string) : string option =
        if attributeName.StartsWith "xmlns:" then Some (attributeName.Substring 6) else None

    /// A local name with the prefix of a qualified name ("w:settings", "evenAndOddHeaders" -> "w:evenAndOddHeaders").
    let withPrefixOf (qualified: string) (localName: string) =
        match qualified.LastIndexOf(':') with
        | -1 -> localName
        | idx -> qualified.Substring(0, idx + 1) + localName

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
