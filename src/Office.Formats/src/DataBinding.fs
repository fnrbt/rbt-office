/// Content controls bound to custom XML (w:sdtPr/w:dataBinding): what Word shows in them is the value of
/// the node their XPath names in a custom XML part, not the text document.xml caches - and an edit of such
/// a control goes back into that part when the document is saved. The rules are Word's, from the fixtures
/// of tools/word-parity/fixtures/data-binding.py (docs/layout-parity-flow.md, the data-binding rows):
///
/// - The store is the custom XML part whose itemProps' ds:itemID is the w:storeItemID (case does not
///   matter). With no such part (or no w:storeItemID) it is the first part the XPath matches a node in;
///   a part of that ID the XPath matches nothing in is not given up for another.
/// - No node, or an element that holds elements: the control keeps the text it has.
/// - The value shown: the element's text or the attribute's value, as it is (white space kept); a tab is a
///   tab; a line break a w:br in a control of w:text w:multiLine, nothing in any other; a date control's
///   value (xsd:date / dateTime) in its w:dateFormat (the w:lid's month and day names), one that is no
///   date as it is; a drop-down list's or combo box's the w:displayText of the list item of that value,
///   else the value; a check box checked by "true" or "1".
/// - The shown text the same as the cached text: nothing changes (the runs keep their formatting).
///   Otherwise the control holds one run of the text, in w:sdtPr/w:rPr's properties (none without it);
///   a check box keeps its run and changes its glyph. An empty value: the control shows its placeholder -
///   what it has when it shows it already, else five spaces (w:showingPlcHdr).
/// - A repeating section (w15:repeatingSection) bound to a node set - its XPath without a position on
///   the last step: Word saves ".../item[1]" and repeats every item - has an item for each node: the
///   items there are kept, the last one is copied for more (its controls' XPaths to the n-th node), and
///   those past the set's end go.
/// - Saving (writeBack): a bound control whose text is not what its node's value shows (an edit) gives
///   the node its text - a check box "true" / "false", a list the value of its item of that text -, an
///   attribute its value; of several controls of one node the last. One showing its placeholder gives
///   nothing. The rest of the custom XML part stays as it was.
///
/// The XPaths read: absolute location paths of element steps - `name`, `prefix:name` (the prefixes of
/// w:prefixMappings, matched by namespace URI) or `*` - each with any of the predicates `[n]`,
/// `[@attr='value']` and `[name='value']`, optionally ending in `/@attr` or `/text()`. Any other XPath
/// matches nothing (the control keeps its text). Custom XML parts must be UTF-8. Repeating sections
/// nested in a repeating section are left as they are; w15:dataBinding of a rich text control (flat
/// WordprocessingML in the store) is not read, nor a picture control's image (base64 in the store): they
/// keep what they hold, as do building blocks, citations, equations and groups.
module DataBinding

open System
open XmlDsl

let private local (name: string) = Ns.stripNs name

let private elementKids (node: XmlNode) =
    match node with
    | Element (_, _, kids, _) -> kids |> List.filter (function Element _ -> true | _ -> false)
    | _ -> []

let private nameOf (node: XmlNode) = match node with Element (name, _, _, _) -> name | _ -> ""
let private attrsOf (node: XmlNode) = match node with Element (_, attrs, _, _) -> attrs | _ -> Map.empty
let private spanOf (node: XmlNode) = match node with Element (_, _, _, span) -> span | _ -> Span.none

let private child (name: string) (node: XmlNode) = elementKids node |> List.tryFind (fun kid -> local (nameOf kid) = name)

/// An attribute's value by its local name.
let private attr (name: string) (node: XmlNode) =
    attrsOf node |> Map.tryPick (fun key value -> if local key = name && not (key.StartsWith "xmlns") then Some value else None)

/// The qualified name of an element's attribute of that local name, or the name it would have with the
/// element's prefix.
let private attrName (name: string) (node: XmlNode) =
    attrsOf node |> Map.tryPick (fun key _ -> if local key = name && not (key.StartsWith "xmlns") then Some key else None)
    |> Option.defaultValue (Ns.withPrefixOf (nameOf node) name)

let private slice (text: string) (node: XmlNode) =
    let span = spanOf node
    text.Substring(span.Start, span.End - span.Start)

let private startTag (text: string) (node: XmlNode) =
    let span = spanOf node
    text.Substring(span.Start, span.StartTagEnd - span.Start)

/// The element's text as written: its start tag (made an open tag if it closed itself), `content`, its end tag.
let private withContent (text: string) (node: XmlNode) (content: string) =
    let span = spanOf node
    if span.StartTagEnd = span.End then
        let tag = startTag text node
        tag.Substring(0, tag.Length - 2).TrimEnd() + ">" + content + "</" + nameOf node + ">"
    else startTag text node + content + text.Substring(span.EndTagStart, span.End - span.EndTagStart)

let private escape (value: string) = value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")

/// What a text keeps all of: a custom XML part's values are its text as it is.
let private scanKeeping (text: string) : XmlNode option =
    match XmlDsl.tryScanKeeping (fun _ -> true) text with
    | Some (Element _ as node) -> Some node
    | _ -> None

// ---- The custom XML parts ----

/// A custom XML part: its ds:itemID (upper case, in braces), path, text and root element.
type Store = { Id: string; Path: string; Text: string; Root: XmlNode }

let private normalId (id: string) =
    let bare = id.Trim().Trim('{', '}').ToUpperInvariant()
    "{" + bare + "}"

let private relationshipsIn (text: string option) =
    match text |> Option.bind (fun xml -> XmlDsl.tryScan xml (Collections.Generic.HashSet<string>())) with
    | Some root ->
        elementKids root
        |> List.filter (fun rel -> local (nameOf rel) = "Relationship" && attr "TargetMode" rel <> Some "External")
        |> List.choose (fun rel -> match attr "Type" rel, attr "Target" rel with Some kind, Some target -> Some (kind, target) | _ -> None)
    | None -> []

/// The custom XML parts of a package, in the order of the main part's relationships to them (then any other
/// customXml/itemN.xml): `text` reads a part.
let storesOf (text: string -> string option) (mainPath: string) : Store list =
    let fromRelationships =
        relationshipsIn (text (PartPath.relationshipsOf mainPath))
        |> List.filter (fun (kind, _) -> kind.EndsWith "/customXml")
        |> List.map (fun (_, target) -> ZipInterop.resolveTarget (PartPath.directory mainPath) target)
    let numbered =
        let found = ResizeArray<string>()
        let mutable n = 1
        while n <= 1000 && (text (sprintf "customXml/item%d.xml" n)).IsSome do
            found.Add (sprintf "customXml/item%d.xml" n)
            n <- n + 1
        List.ofSeq found
    (fromRelationships @ numbered)
    |> List.distinctBy (fun path -> path.ToLowerInvariant())
    |> List.choose (fun path ->
        let props =
            relationshipsIn (text (PartPath.relationshipsOf path))
            |> List.tryFind (fun (kind, _) -> kind.EndsWith "/customXmlProps")
            |> Option.map (fun (_, target) -> ZipInterop.resolveTarget (PartPath.directory path) target)
        let id =
            props
            |> Option.bind text
            |> Option.bind (fun xml -> XmlDsl.tryScan xml (Collections.Generic.HashSet<string>()))
            |> Option.bind (attr "itemID")
        match id, text path |> Option.bind (fun xml -> scanKeeping xml |> Option.map (fun root -> xml, root)) with
        | Some id, Some (xml, root) -> Some { Id = normalId id; Path = path; Text = xml; Root = root }
        | _ -> None)

// ---- The XPaths ----

type private Name = { Prefix: string option; Local: string }

type private Predicate =
    | Position of int
    | AttributeIs of Name * string
    | ChildIs of Name * string

type private Step =
    | ElementStep of Name * Predicate list
    | AttributeStep of Name
    | TextStep

/// A node an XPath gives: an element (with the namespaces in scope at it) or an attribute (its element,
/// its qualified name and value).
type private Found =
    | FoundElement of XmlNode
    | FoundAttribute of owner: XmlNode * name: string * value: string

let private nameOfText (text: string) : Name option =
    let ok (part: string) = part <> "" && part |> Seq.forall (fun c -> Char.IsLetterOrDigit c || c = '_' || c = '-' || c = '.')
    match text.Split(':') with
    | [| one |] when ok one || one = "*" -> Some { Prefix = None; Local = one }
    | [| prefix; name |] when ok prefix && (ok name || name = "*") -> Some { Prefix = Some prefix; Local = name }
    | _ -> None

let private quoted (text: string) =
    let text = text.Trim()
    if text.Length >= 2 && (text.[0] = '\'' || text.[0] = '"') && text.[text.Length - 1] = text.[0] then Some (text.Substring(1, text.Length - 2))
    else None

let private predicateOf (text: string) : Predicate option =
    let text = text.Trim()
    match Int32.TryParse text with
    | true, n when n >= 1 -> Some (Position n)
    | _ ->
        match text.IndexOf '=' with
        | at when at > 0 ->
            let left, right = text.Substring(0, at).Trim(), text.Substring(at + 1)
            match quoted right with
            | Some value when left.StartsWith "@" -> nameOfText (left.Substring 1) |> Option.map (fun name -> AttributeIs (name, value))
            | Some value -> nameOfText left |> Option.map (fun name -> ChildIs (name, value))
            | None -> None
        | _ -> None

/// The steps of an absolute XPath, outside brackets and quotes split at "/".
let private stepsOf (xpath: string) : Step list option =
    let xpath = xpath.Trim()
    if not (xpath.StartsWith "/") || xpath.StartsWith "//" then None
    else
        let pieces = ResizeArray<string>()
        let current = Text.StringBuilder()
        let mutable depth = 0
        let mutable quote = ' '
        for c in xpath.Substring 1 do
            if quote <> ' ' then
                (if c = quote then quote <- ' ')
                current.Append c |> ignore
            elif c = '\'' || c = '"' then quote <- c; current.Append c |> ignore
            elif c = '[' then depth <- depth + 1; current.Append c |> ignore
            elif c = ']' then depth <- depth - 1; current.Append c |> ignore
            elif c = '/' && depth = 0 then pieces.Add (current.ToString()); current.Clear() |> ignore
            else current.Append c |> ignore
        pieces.Add (current.ToString())
        let step (piece: string) : Step option =
            let piece = piece.Trim()
            if piece = "text()" then Some TextStep
            elif piece.StartsWith "@" then nameOfText (piece.Substring 1) |> Option.map AttributeStep
            else
                let at = piece.IndexOf '['
                let nameText = if at < 0 then piece else piece.Substring(0, at)
                let predicates =
                    if at < 0 then Some []
                    else
                        // "[a][b]": each bracket's text.
                        let found = ResizeArray<Predicate option>()
                        let mutable rest = piece.Substring at
                        let mutable ok = true
                        while ok && rest <> "" do
                            if rest.[0] <> '[' then ok <- false
                            else
                                let close = rest.IndexOf ']'
                                if close < 0 then ok <- false
                                else
                                    found.Add (predicateOf (rest.Substring(1, close - 1)))
                                    rest <- rest.Substring(close + 1).Trim()
                        if ok && found |> Seq.forall Option.isSome then Some (found |> Seq.choose id |> List.ofSeq) else None
                match nameOfText nameText, predicates with
                | Some name, Some predicates -> Some (ElementStep (name, predicates))
                | _ -> None
        let steps = pieces |> Seq.map step |> List.ofSeq
        if steps |> List.forall Option.isSome then
            let steps = steps |> List.choose id
            // (Only the last step may be an attribute or text().)
            let inner = steps |> List.take (steps.Length - 1)
            if steps.IsEmpty || inner |> List.exists (function ElementStep _ -> false | _ -> true) then None else Some steps
        else None

/// The namespaces of w:prefixMappings ("xmlns:ns0='urn:a' xmlns:ns1='urn:b'"), by prefix.
let private prefixMappingsOf (text: string) : Map<string, string> =
    [ for m in Text.RegularExpressions.Regex.Matches(text, "xmlns:([^\\s=]+)\\s*=\\s*(['\"])(.*?)\\2") -> m.Groups.[1].Value, m.Groups.[3].Value ]
    |> Map.ofList

/// The namespaces declared on an element added to those in scope ("" the default one).
let private scopeOf (scope: Map<string, string>) (node: XmlNode) =
    attrsOf node |> Map.fold (fun scope key value ->
        if key = "xmlns" then Map.add "" value scope
        elif key.StartsWith "xmlns:" then Map.add (key.Substring 6) value scope
        else scope) scope

/// An element's or attribute's namespace URI and local name (an attribute without a prefix: none).
let private expanded (scope: Map<string, string>) (isAttribute: bool) (qualified: string) =
    match qualified.IndexOf ':' with
    | -1 -> (if isAttribute then "" else Map.tryFind "" scope |> Option.defaultValue ""), qualified
    | at -> (Map.tryFind (qualified.Substring(0, at)) scope |> Option.defaultValue ("?" + qualified.Substring(0, at))), qualified.Substring(at + 1)

/// Whether a qualified name read in `scope` is the XPath's name (its prefix's namespace by `mappings`; an
/// unprefixed name: no namespace).
let private matches (mappings: Map<string, string>) (scope: Map<string, string>) (isAttribute: bool) (name: Name) (qualified: string) =
    let uri, localName = expanded scope isAttribute qualified
    let wanted = match name.Prefix with None -> Some "" | Some prefix -> Map.tryFind prefix mappings
    match wanted with
    | Some wantedUri -> (name.Local = "*" || name.Local = localName) && wantedUri = uri
    | None -> false

let private textOfElement (node: XmlNode) =
    match node with
    | Element (_, _, kids, _) -> kids |> List.choose (function Text t -> Some t | _ -> None) |> String.concat ""
    | _ -> ""

/// The nodes an XPath gives in a store's tree.
let private evaluate (mappings: Map<string, string>) (steps: Step list) (root: XmlNode) : Found list =
    let scopeAtRoot = scopeOf Map.empty root
    let rec walk (context: (XmlNode * Map<string, string>) list) (steps: Step list) (atDocument: bool) : Found list =
        match steps with
        | [] -> context |> List.map (fun (node, _) -> FoundElement node)
        | TextStep :: _ -> context |> List.map (fun (node, _) -> FoundElement node)
        | AttributeStep name :: _ ->
            if atDocument then []
            else
                context |> List.collect (fun (node, scope) ->
                    attrsOf node
                    |> Map.toList
                    |> List.filter (fun (key, _) -> not (key = "xmlns" || key.StartsWith "xmlns:") && matches mappings scope true name key)
                    |> List.map (fun (key, value) -> FoundAttribute (node, key, value)))
        | ElementStep (name, predicates) :: rest ->
            let next =
                context |> List.collect (fun (node, scope) ->
                    let candidates =
                        if atDocument then [ root, scopeAtRoot ]
                        else elementKids node |> List.map (fun kid -> kid, scopeOf scope kid)
                    let named = candidates |> List.filter (fun (kid, kidScope) -> matches mappings kidScope false name (nameOf kid))
                    predicates |> List.fold (fun (set: (XmlNode * Map<string, string>) list) predicate ->
                        match predicate with
                        | Position n -> set |> List.tryItem (n - 1) |> Option.toList
                        | AttributeIs (attrName, value) ->
                            set |> List.filter (fun (kid, kidScope) ->
                                attrsOf kid |> Map.exists (fun key v -> not (key.StartsWith "xmlns") && v = value && matches mappings kidScope true attrName key))
                        | ChildIs (childName, value) ->
                            set |> List.filter (fun (kid, kidScope) ->
                                elementKids kid |> List.exists (fun grandchild ->
                                    matches mappings (scopeOf kidScope grandchild) false childName (nameOf grandchild) && textOfElement grandchild = value))) named)
            walk next rest false
    walk [ root, Map.empty ] steps true

/// The node a binding names, and its store: None when there is none (the control keeps its text).
let private lookup (stores: Store list) (binding: XmlNode) : (Store * Found) option =
    match attr "xpath" binding |> Option.bind stepsOf with
    | None -> None
    | Some steps ->
        let mappings = attr "prefixMappings" binding |> Option.map prefixMappingsOf |> Option.defaultValue Map.empty
        let first (store: Store) = evaluate mappings steps store.Root |> List.tryHead |> Option.map (fun found -> store, found)
        match attr "storeItemID" binding |> Option.map normalId |> Option.bind (fun id -> stores |> List.tryFind (fun store -> store.Id = id)) with
        | Some store -> first store
        | None -> stores |> List.tryPick first

/// The value of a node: an element's text (None for one that holds elements), an attribute's value.
let private valueOf (found: Found) : string option =
    match found with
    | FoundAttribute (_, _, value) -> Some value
    | FoundElement node -> if (elementKids node).IsEmpty then Some (textOfElement node) else None

// ---- Dates ----

let private monthNames =
    dict [
        "en", [| "January"; "February"; "March"; "April"; "May"; "June"; "July"; "August"; "September"; "October"; "November"; "December" |]
        "de", [| "Januar"; "Februar"; "März"; "April"; "Mai"; "Juni"; "Juli"; "August"; "September"; "Oktober"; "November"; "Dezember" |]
        "fr", [| "janvier"; "février"; "mars"; "avril"; "mai"; "juin"; "juillet"; "août"; "septembre"; "octobre"; "novembre"; "décembre" |]
        "es", [| "enero"; "febrero"; "marzo"; "abril"; "mayo"; "junio"; "julio"; "agosto"; "septiembre"; "octubre"; "noviembre"; "diciembre" |]
        "it", [| "gennaio"; "febbraio"; "marzo"; "aprile"; "maggio"; "giugno"; "luglio"; "agosto"; "settembre"; "ottobre"; "novembre"; "dicembre" |]
        "nl", [| "januari"; "februari"; "maart"; "april"; "mei"; "juni"; "juli"; "augustus"; "september"; "oktober"; "november"; "december" |]
    ]

let private dayNames =
    dict [
        "en", [| "Sunday"; "Monday"; "Tuesday"; "Wednesday"; "Thursday"; "Friday"; "Saturday" |]
        "de", [| "Sonntag"; "Montag"; "Dienstag"; "Mittwoch"; "Donnerstag"; "Freitag"; "Samstag" |]
        "fr", [| "dimanche"; "lundi"; "mardi"; "mercredi"; "jeudi"; "vendredi"; "samedi" |]
        "es", [| "domingo"; "lunes"; "martes"; "miércoles"; "jueves"; "viernes"; "sábado" |]
        "it", [| "domenica"; "lunedì"; "martedì"; "mercoledì"; "giovedì"; "venerdì"; "sabato" |]
        "nl", [| "zondag"; "maandag"; "dinsdag"; "woensdag"; "donderdag"; "vrijdag"; "zaterdag" |]
    ]

/// An xsd:date or xsd:dateTime value's date and time (a time zone is not applied).
let private dateOf (value: string) : (int * int * int * int * int * int) option =
    let m = Text.RegularExpressions.Regex.Match(value.Trim(), "^(\\d{4})-(\\d{2})-(\\d{2})(?:T(\\d{2}):(\\d{2})(?::(\\d{2})(?:\\.\\d+)?)?)?(?:Z|[+-]\\d{2}:\\d{2})?$")
    if not m.Success then None
    else
        let part (i: int) = if m.Groups.[i].Success then int m.Groups.[i].Value else 0
        let y, mo, d = part 1, part 2, part 3
        if mo < 1 || mo > 12 || d < 1 || d > DateTime.DaysInMonth(y, mo) || part 4 > 23 || part 5 > 59 || part 6 > 59 then None
        else Some (y, mo, d, part 4, part 5, part 6)

/// A date in a Word date picture (d dd ddd dddd, M MM MMM MMMM, yy yyyy, h hh H HH, m mm, s ss, am/pm,
/// 'quoted text'); `lid` gives the names (English for a language not known here).
let formatDate (picture: string) (lid: string) (y: int, mo: int, d: int, h: int, mi: int, s: int) : string =
    let language = (if lid.Length >= 2 then lid.Substring(0, 2) else "en").ToLowerInvariant()
    let months = if monthNames.ContainsKey language then monthNames.[language] else monthNames.["en"]
    let days = if dayNames.ContainsKey language then dayNames.[language] else dayNames.["en"]
    let weekday = int (DateTime(y, mo, d).DayOfWeek)
    let out = Text.StringBuilder()
    let mutable i = 0
    let two (n: int) = (if n < 10 then "0" else "") + string n
    let short (name: string) = if name.Length > 3 then name.Substring(0, 3) else name
    while i < picture.Length do
        let c = picture.[i]
        let mutable run = 1
        while i + run < picture.Length && picture.[i + run] = c do run <- run + 1
        let rest = picture.Substring i
        if c = '\'' then
            let close = picture.IndexOf('\'', i + 1)
            if close < 0 then out.Append(picture.Substring(i + 1)) |> ignore; i <- picture.Length
            else out.Append(picture.Substring(i + 1, close - i - 1)) |> ignore; i <- close + 1
        elif rest.StartsWith("am/pm", StringComparison.OrdinalIgnoreCase) then
            out.Append(if h < 12 then "AM" else "PM") |> ignore; i <- i + 5
        else
            match c with
            | 'd' | 'D' ->
                out.Append(match run with 1 -> string d | 2 -> two d | 3 -> short days.[weekday] | _ -> days.[weekday]) |> ignore
            | 'M' ->
                out.Append(match run with 1 -> string mo | 2 -> two mo | 3 -> short months.[mo - 1] | _ -> months.[mo - 1]) |> ignore
            | 'y' | 'Y' -> out.Append(if run <= 2 then two (y % 100) else string y) |> ignore
            | 'h' -> let h12 = (if h % 12 = 0 then 12 else h % 12) in out.Append(if run = 1 then string h12 else two h12) |> ignore
            | 'H' -> out.Append(if run = 1 then string h else two h) |> ignore
            | 'm' -> out.Append(if run = 1 then string mi else two mi) |> ignore
            | 's' | 'S' -> out.Append(if run = 1 then string s else two s) |> ignore
            | _ -> out.Append(String.replicate run (string c)) |> ignore
            i <- i + run
    out.ToString()

// ---- The controls ----

type private Kind =
    | PlainText of multiLine: bool
    | DateControl of picture: string * lid: string
    | ListControl of (string * string) list
    | CheckBox of checkedElement: XmlNode option * checkedGlyph: string * uncheckedGlyph: string

let private isOn (value: string option) =
    match value with
    | Some v -> let v = v.Trim().ToLowerInvariant() in v = "1" || v = "true" || v = "on"
    | None -> false

let private glyphOf (state: XmlNode option) (fallback: int) =
    let code =
        state |> Option.bind (attr "val")
        |> Option.bind (fun hex ->
            let hex = hex.Trim()
            if hex <> "" && hex.Length <= 4 && hex |> Seq.forall (fun c -> "0123456789abcdefABCDEF".IndexOf c >= 0) then Some (Convert.ToInt32(hex, 16)) else None)
        |> Option.defaultValue fallback
    string (char code)

let private kindOf (props: XmlNode) : Kind =
    match child "checkbox" props, child "date" props, (match child "dropDownList" props with Some list -> Some list | None -> child "comboBox" props) with
    | Some box, _, _ -> CheckBox (child "checked" box, glyphOf (child "checkedState" box) 0x2612, glyphOf (child "uncheckedState" box) 0x2610)
    | None, Some date, _ ->
        DateControl (child "dateFormat" date |> Option.bind (attr "val") |> Option.defaultValue "M/d/yyyy",
                     child "lid" date |> Option.bind (attr "val") |> Option.defaultValue "en-US")
    | None, None, Some list ->
        ListControl (elementKids list |> List.filter (fun item -> local (nameOf item) = "listItem")
                     |> List.choose (fun item -> attr "value" item |> Option.map (fun value -> value, attr "displayText" item |> Option.defaultValue value)))
    | None, None, None -> PlainText (child "text" props |> Option.bind (attr "multiLine") |> isOn)

/// The text a control shows for a value (a check box: its glyph).
let private shownText (kind: Kind) (value: string) : string =
    match kind with
    | PlainText multiLine -> if multiLine then value.Replace("\r\n", "\n").Replace("\r", "\n") else value.Replace("\r", "").Replace("\n", "")
    | DateControl (picture, lid) -> match dateOf value with Some date -> formatDate picture lid date | None -> value
    | ListControl items -> items |> List.tryPick (fun (v, shown) -> if v = value then Some shown else None) |> Option.defaultValue value
    | CheckBox (_, on, off) -> if isOn (Some value) then on else off

/// The text runs hold under an element: w:t, a run's w:tab and w:br / w:cr as tab and line break, a new
/// paragraph as a line break.
let private textUnder (node: XmlNode) : string =
    let out = Text.StringBuilder()
    let rec visit (inRun: bool) (node: XmlNode) =
        match node with
        | Element (name, _, kids, _) ->
            match local name with
            | "t" -> out.Append(textOfElement node) |> ignore
            | "tab" when inRun -> out.Append '\t' |> ignore
            | ("br" | "cr") when inRun -> out.Append '\n' |> ignore
            | "delText" | "instrText" | "rPr" | "pPr" | "sdtPr" | "sdtEndPr" -> ()
            | "p" ->
                if out.Length > 0 then out.Append '\n' |> ignore
                kids |> List.iter (visit false)
            | "r" -> kids |> List.iter (visit true)
            | _ -> kids |> List.iter (visit inRun)
        | _ -> ()
    visit false node
    out.ToString()

/// A run of a text in the w:rPr given: tabs as w:tab, line breaks as w:br.
let private runOf (prefix: string) (rPr: string) (text: string) =
    let q (name: string) = if prefix = "" then name else prefix + ":" + name
    let out = Text.StringBuilder()
    out.Append("<" + q "r" + ">" + rPr) |> ignore
    let piece = Text.StringBuilder()
    let flush () =
        if piece.Length > 0 then
            let value = piece.ToString()
            let space = if XmlPatch.needsPreserve value then " xml:space=\"preserve\"" else ""
            out.Append("<" + q "t" + space + ">" + escape value + "</" + q "t" + ">") |> ignore
            piece.Clear() |> ignore
    for c in text do
        match c with
        | '\t' -> flush (); out.Append("<" + q "tab" + "/>") |> ignore
        | '\n' -> flush (); out.Append("<" + q "br" + "/>") |> ignore
        | c -> piece.Append c |> ignore
    flush ()
    out.Append("</" + q "r" + ">").ToString()

let private prefixOf (qualified: string) = match qualified.IndexOf ':' with -1 -> "" | at -> qualified.Substring(0, at)

/// The plain data binding of a control's properties (not a rich text control's w15:dataBinding, nor a
/// repeating section's).
let private plainBinding (sdt: XmlNode) (props: XmlNode) =
    if (child "repeatingSection" props).IsSome then None
    else elementKids props |> List.tryFind (fun kid -> nameOf kid = Ns.withPrefixOf (nameOf sdt) "dataBinding")

/// Every element of a name under a node, outermost first, not looking inside one found.
let private outermost (name: string) (node: XmlNode) : XmlNode list =
    let found = ResizeArray<XmlNode>()
    let rec visit (node: XmlNode) =
        for kid in elementKids node do
            if local (nameOf kid) = name then found.Add kid else visit kid
    visit node
    List.ofSeq found

let private allNamed (name: string) (node: XmlNode) : XmlNode list =
    let found = ResizeArray<XmlNode>()
    let rec visit (node: XmlNode) =
        for kid in elementKids node do
            if local (nameOf kid) = name then found.Add kid
            visit kid
    visit node
    List.ofSeq found

let private noOrder : XmlPatch.SchemaOrder = fun _ -> []

/// A repeating section's items as many as its node set has nodes (see the module's comment).
let private repeatedSections (stores: Store list) (text: string) (root: XmlNode) : (XmlNode * XmlPatch.ElementEdit list) list =
    let sections =
        outermost "sdt" root
        |> List.collect (fun sdt -> sdt :: allNamed "sdt" sdt)
        |> List.filter (fun sdt -> match child "sdtPr" sdt with Some props -> (child "repeatingSection" props).IsSome | None -> false)
    // (One inside another is left as it is.)
    let inside (outer: XmlNode) (inner: XmlNode) = let a, b = spanOf outer, spanOf inner in a <> b && b.Start >= a.Start && b.End <= a.End
    let sections = sections |> List.filter (fun sdt -> not (sections |> List.exists (fun outer -> inside outer sdt)))
    sections |> List.choose (fun sdt ->
        let props = (child "sdtPr" sdt).Value
        match child "dataBinding" props, child "sdtContent" sdt with
        | Some binding, Some content ->
            // (The node set: the XPath without a position on its last step - Word saves "/a[1]/item[1]" and
            // repeats every item.)
            let xpath = Text.RegularExpressions.Regex.Replace((attr "xpath" binding |> Option.defaultValue "").Trim(), "\\[\\s*\\d+\\s*\\]$", "")
            let mappings = attr "prefixMappings" binding |> Option.map prefixMappingsOf |> Option.defaultValue Map.empty
            let count =
                match stepsOf xpath with
                | Some steps ->
                    let inStore (store: Store) = match evaluate mappings steps store.Root with [] -> None | found -> Some found.Length
                    match attr "storeItemID" binding |> Option.map normalId |> Option.bind (fun id -> stores |> List.tryFind (fun store -> store.Id = id)) with
                    | Some store -> inStore store |> Option.defaultValue 0
                    | None -> stores |> List.tryPick inStore |> Option.defaultValue 0
                | None -> 0
            let items =
                elementKids content
                |> List.filter (fun kid -> local (nameOf kid) = "sdt" && (child "sdtPr" kid |> Option.bind (child "repeatingSectionItem")).IsSome)
            if count = 0 || items.IsEmpty || count = items.Length then None
            else
                let last = List.last items
                let copy (k: int) =
                    let from = sprintf "%s[%d]" xpath items.Length
                    let targets =
                        [ for sdtInside in last :: allNamed "sdt" last do
                              match child "sdtPr" sdtInside with
                              | Some innerProps ->
                                  match child "id" innerProps with
                                  | Some idNode ->
                                      let id = attr "val" idNode |> Option.bind (fun v -> match Int32.TryParse v with | true, n -> Some n | _ -> None) |> Option.defaultValue 0
                                      let fresh = int ((int64 id * 31L + int64 k * 7919L) % 2147483647L)
                                      yield idNode, [ XmlPatch.SetAttribute (attrName "val" idNode, string fresh) ]
                                  | None -> ()
                                  for innerBinding in elementKids innerProps |> List.filter (fun kid -> local (nameOf kid) = "dataBinding") do
                                      match attr "xpath" innerBinding with
                                      | Some inner when inner.StartsWith from ->
                                          yield innerBinding, [ XmlPatch.SetAttribute (attrName "xpath" innerBinding, sprintf "%s[%d]%s" xpath k (inner.Substring from.Length)) ]
                                      | _ -> ()
                              | None -> () ]
                    XmlPatch.apply noOrder { Text = text; Node = last } (XmlPatch.editsAt last targets)
                let itemsText =
                    [ for k in 1 .. count -> if k <= items.Length then slice text items.[k - 1] else copy k ] |> String.concat ""
                let first = List.head items
                let span = spanOf content
                let before = text.Substring(span.StartTagEnd, (spanOf first).Start - span.StartTagEnd)
                let after = text.Substring((spanOf last).End, span.EndTagStart - (spanOf last).End)
                Some (content, [ XmlPatch.Replaced (withContent text content (before + itemsText + after)) ])
        | _ -> None)

/// The controls whose bound value is not text: a picture (the image a store's base64 holds is not read),
/// building blocks, a citation, an equation, a group.
let private unread = [ "picture"; "docPartObj"; "docPartList"; "citation"; "bibliography"; "equation"; "group" ]

/// The edits that make each bound control of a part show its node's value.
let private boundValues (stores: Store list) (text: string) (root: XmlNode) : (XmlNode * XmlPatch.ElementEdit list) list =
    let sdts = root :: allNamed "sdt" root |> List.filter (fun node -> local (nameOf node) = "sdt")
    let edits =
        sdts |> List.collect (fun sdt ->
            match child "sdtPr" sdt, child "sdtContent" sdt with
            | Some props, Some content ->
                match plainBinding sdt props |> Option.bind (lookup stores) |> Option.bind (fun (_, found) -> valueOf found) with
                | None -> []
                | Some _ when unread |> List.exists (fun name -> (child name props).IsSome) -> []
                | Some value ->
                    let kind = kindOf props
                    let prefix = prefixOf (nameOf sdt)
                    let q (name: string) = if prefix = "" then name else prefix + ":" + name
                    let rPr = child "rPr" props |> Option.map (slice text) |> Option.defaultValue ""
                    let placeholder = child "showingPlcHdr" props
                    let binding = plainBinding sdt props |> Option.get
                    let kids = elementKids content
                    let isBlock = kids |> List.exists (fun kid -> local (nameOf kid) = "p")
                    let holdsCells = kids |> List.exists (fun kid -> match local (nameOf kid) with "tbl" | "tr" | "tc" -> true | _ -> false)
                    // The control's content made one run of the text.
                    let contentOf (shown: string) =
                        let run = runOf prefix rPr shown
                        if isBlock then
                            match kids |> List.tryFind (fun kid -> local (nameOf kid) = "p") with
                            | Some p ->
                                let pPr = child "pPr" p |> Option.map (slice text) |> Option.defaultValue ""
                                withContent text content (withContent text p (pPr + run))
                            | None -> withContent text content run
                        else withContent text content run
                    if holdsCells then []
                    else
                        match kind with
                        | CheckBox (checkedNode, on, off) ->
                            let want = isOn (Some value)
                            let now = checkedNode |> Option.bind (attr "val") |> fun v -> (match v with None -> checkedNode.IsSome | Some _ -> isOn v)
                            if want = now then []
                            else
                                let glyph = if want then on else off
                                [ match checkedNode with
                                  | Some node -> yield node, [ XmlPatch.SetAttribute (attrName "val" node, if want then "1" else "0") ]
                                  | None -> ()
                                  match allNamed "t" content |> List.tryHead with
                                  | Some t -> yield t, [ XmlPatch.Replaced (withContent text t glyph) ]
                                  | None -> yield content, [ XmlPatch.Replaced (contentOf glyph) ] ]
                        | _ when value = "" ->
                            // (An empty value: the placeholder.)
                            if placeholder.IsSome then []
                            else [ content, [ XmlPatch.Replaced (contentOf "     ") ]
                                   binding, [ XmlPatch.InsertBefore ("<" + q "showingPlcHdr" + "/>") ] ]
                        | _ ->
                            let shown = shownText kind value
                            if placeholder.IsNone && shown = textUnder content then []
                            else
                                [ yield content, [ XmlPatch.Replaced (contentOf shown) ]
                                  match placeholder with
                                  | Some node -> yield node, [ XmlPatch.Removed ]
                                  | None -> () ]
            | _ -> [])
    // (A control inside one whose content is replaced goes with it.)
    let replaced = edits |> List.filter (fun (_, e) -> e |> List.exists (function XmlPatch.Replaced _ -> true | _ -> false)) |> List.map (fst >> spanOf)
    edits |> List.filter (fun (node, _) ->
        let s = spanOf node
        not (replaced |> List.exists (fun outer -> outer <> s && s.Start >= outer.Start && s.End <= outer.End)))

/// Whether a part has a data binding element (not the word in its text: a document about them is not
/// scanned whole for none - docx4j-getting-started, opened front first).
let private bindingTag = Text.RegularExpressions.Regex("<([A-Za-z0-9_]+:)?dataBinding[\\s/>]")
let private holdsBinding (text: string) = text.Contains "dataBinding" && bindingTag.IsMatch text

/// A part's XML with its bound controls showing their values (the part itself when none changes).
let resolvePart (stores: Store list) (text: string) : string =
    if stores.IsEmpty || not (holdsBinding text) then text
    else
        let pass (edit: string -> XmlNode -> (XmlNode * XmlPatch.ElementEdit list) list) (text: string) =
            match XmlPatch.readPart text with
            | Some part ->
                match edit text part.Node with
                | [] -> text
                | targets -> XmlPatch.applyPart noOrder part (XmlPatch.editsAt part.Node targets)
            | None -> text
        try text |> pass (repeatedSections stores) |> pass (boundValues stores)
        with _ -> text

// ---- Saving ----

/// The value a control's content gives its node, when it is not what the node's value shows (an edit).
let private editedValue (kind: Kind) (content: XmlNode) (props: XmlNode) (value: string) : string option =
    match kind with
    | CheckBox (checkedNode, _, _) ->
        let now = checkedNode |> Option.bind (attr "val") |> fun v -> (match v with None -> checkedNode.IsSome | Some _ -> isOn v)
        if now = isOn (Some value) then None else Some (if now then "true" else "false")
    | _ ->
        if (child "showingPlcHdr" props).IsSome then None
        else
            let current = textUnder content
            let shown = if value = "" then "     " else shownText kind value
            if current = shown then None
            else
                match kind with
                | ListControl items -> Some (items |> List.tryPick (fun (v, shownItem) -> if shownItem = current then Some v else None) |> Option.defaultValue current)
                | _ -> Some current

/// The custom XML parts the controls of `stories` (parts' text) were edited in, as their new text: `text`
/// reads the package's other parts; `mainPath` the main document part's path.
let writeBack (text: string -> string option) (mainPath: string) (stories: string list) : (string * string) list =
    let stories = stories |> List.filter holdsBinding
    if stories.IsEmpty then []
    else
        let stores = storesOf text mainPath
        if stores.IsEmpty then []
        else
            // Store path -> (node, value): the last edit of a node wins.
            let changes = Collections.Generic.Dictionary<string, ResizeArray<Found * string>>()
            for story in stories do
                match XmlPatch.readPart story with
                | Some part ->
                    for sdt in allNamed "sdt" part.Node do
                        match child "sdtPr" sdt, child "sdtContent" sdt with
                        | Some props, Some content ->
                            match plainBinding sdt props |> Option.filter (fun _ -> not (unread |> List.exists (fun name -> (child name props).IsSome))) |> Option.bind (lookup stores) with
                            | Some (store, found) ->
                                match valueOf found |> Option.bind (editedValue (kindOf props) content props) with
                                | Some value ->
                                    if not (changes.ContainsKey store.Path) then changes.[store.Path] <- ResizeArray()
                                    changes.[store.Path].Add ((found, value))
                                | None -> ()
                            | None -> ()
                        | _ -> ()
                | None -> ()
            [ for store in stores do
                  match Portable.Dict.tryFind store.Path changes with
                  | ValueSome edits ->
                      let byNode = Collections.Generic.Dictionary<string, XmlNode * XmlPatch.ElementEdit list>()
                      for (found, value) in edits do
                          match found with
                          | FoundElement node ->
                              byNode.[sprintf "%d" (spanOf node).Start] <- (node, [ XmlPatch.Replaced (withContent store.Text node (escape value)) ])
                          | FoundAttribute (owner, name, _) ->
                              byNode.[sprintf "%d@%s" (spanOf owner).Start name] <- (owner, [ XmlPatch.SetAttribute (name, value) ])
                      // (Several attributes of one element: one edit of it.)
                      let targets =
                          byNode.Values |> List.ofSeq |> List.groupBy (fun (node, _) -> (spanOf node).Start)
                          |> List.map (fun (_, group) -> fst group.Head, group |> List.collect snd)
                      let root = store.Root
                      let rootEdits = targets |> List.filter (fun (node, _) -> spanOf node = spanOf root) |> List.collect snd
                      let newText =
                          match rootEdits |> List.tryPick (function XmlPatch.Replaced xml -> Some xml | _ -> None) with
                          | Some xml ->
                              let span = spanOf root
                              store.Text.Substring(0, span.Start) + xml + store.Text.Substring span.End
                          | None -> XmlPatch.applyPart noOrder { Text = store.Text; Node = root } (XmlPatch.editsAt root targets)
                      if newText <> store.Text then yield store.Path, newText
                  | ValueNone -> () ]
