/// Changes made to XML that was read, made in its text (docs/save-architecture.md). A change is a typed
/// edit of an element of the span tree (XmlDsl.tryScan); what no edit names stays the bytes it was read
/// as - attributes, children, white space, elements and namespaces nothing here knows. The writers change
/// XML that was read through this alone: no searching, cutting or matching of its text elsewhere.
module XmlPatch

open XmlDsl

/// An edit of one element.
type ElementEdit =
    /// The attribute (qualified name) has this value: changed in place, or put after the others.
    | SetAttribute of name: string * value: string
    /// The start tag's attributes are these, in this order (what stands between them as written goes).
    | SetAllAttributes of attributes: (string * string) list
    | RemoveAttribute of name: string
    /// The child of that name (qualified; the first, where there are several) is this XML; where there is
    /// none it is put in at its place in the element's schema order.
    | SetChild of name: string * xml: string
    /// Every child of that name is taken out, and these put in where the first stood (or at the schema place).
    | SetChildren of name: string * xmls: string list
    | RemoveChild of name: string
    /// The child of that name changed by these edits; where there is none, `whole` is put in (at its schema
    /// place; nothing for None).
    | EditChild of name: string * edits: ElementEdit list * whole: string option
    /// The element holds this text and nothing else (escaped; with xml:space="preserve" put in where the
    /// text needs it - white space at an end, or two together - and the element does not say so).
    | SetText of text: string
    /// The child of that name at that place among those of its name (from 0) changed by these edits.
    | EditNth of name: string * index: int * edits: ElementEdit list
    /// A child put in, after those of its kind (at its schema place).
    | AddChild of name: string * xml: string
    /// The child of that name whose attribute has that value is this XML; put in (as AddChild) where there is none.
    | SetChildWhere of name: string * attribute: string * value: string * xml: string
    | RemoveChildWhere of name: string * attribute: string * value: string
    /// The child of that name whose attribute has that value changed by these edits; nothing where there is none.
    | EditChildWhere of name: string * attribute: string * value: string * edits: ElementEdit list
    /// The element taken out, with the white space before it (the edit of a child: `editsAt`).
    | Removed
    /// XML put in just before the element, or just after it (the edit of a child: `editsAt`).
    | InsertBefore of xml: string
    | InsertAfter of xml: string
    /// XML put in at the start of what the element holds - after the children its schema order puts before
    /// the others ("*": a paragraph's properties) -, or at its end.
    | Prepend of xml: string
    | Append of xml: string
    /// The child element that starts at that place of the text (its span's start: an element of the tree
    /// read) changed by these edits.
    | EditAt of start: int * edits: ElementEdit list
    /// The element is this XML instead (the edit of a child: `EditAt`, `EditNth` ...).
    | Replaced of xml: string

/// The order of the children of an element (local names, by the element's qualified name), for where a
/// child that is put in goes. An element with no order puts a new child after the others. "*" stands for
/// every name the order does not name (w:r: its properties, then what it holds); where there is no "*" a
/// child the order does not name has the place of the one before it.
type SchemaOrder = string -> string list

let private local (name: string) = Ns.stripNs name

/// An element read from a text, with that text (its span is an offset into it).
type SourceElement = { Text: string; Node: XmlNode }

/// The one element of a text: None when the text is not one well-formed element.
let read (text: string) : SourceElement option =
    if text = "" then None
    else
        match tryScan text (System.Collections.Generic.HashSet<string>()) with
        | Some (Element (_, _, _, span) as node) when Span.isKnown span && span.Start >= 0 && text.Substring(span.End).Trim() = "" && text.Substring(0, span.Start).Trim() = "" -> Some { Text = text; Node = node }
        | _ -> None

let private nameOf (node: XmlNode) = match node with Element (name, _, _, _) -> name | _ -> ""
let private attrsOf (node: XmlNode) = match node with Element (_, attrs, _, _) -> attrs | _ -> Map.empty
let private spanOf (node: XmlNode) = match node with Element (_, _, _, span) -> span | _ -> Span.none
let private elementChildren (node: XmlNode) = match node with Element (_, _, kids, _) -> kids |> List.filter (function Element _ -> true | _ -> false) | _ -> []
let private hasText (node: XmlNode) = match node with Element (_, _, kids, _) -> kids |> List.exists (function Text t -> t <> "" | _ -> false) | _ -> false

let escapeText (value: string) =
    value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")

/// The text an element holds (its text nodes, one after another).
let textOf (node: XmlNode) : string =
    match node with
    | Element (_, _, kids, _) -> kids |> List.choose (function Text t -> Some t | _ -> None) |> String.concat ""
    | _ -> ""

/// Whether text in an element needs xml:space="preserve" to keep its white space.
let needsPreserve (text: string) =
    text <> "" && (System.Char.IsWhiteSpace text.[0] || System.Char.IsWhiteSpace text.[text.Length - 1] || text.Contains "  " || text.Contains "\t" || text.Contains "\n")

let escapeAttribute (value: string) =
    value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;")

/// The attributes of the start tag of the element at `span` in `text`: qualified name, where the attribute
/// starts (with the white space before it) and ends, and where its value starts and ends (inside the quotes).
let private attributeRanges (text: string) (span: Span) : (string * int * int * int * int) list =
    let found = ResizeArray<string * int * int * int * int>()
    let tagEnd = span.StartTagEnd
    let mutable at = span.Start + 1
    while at < tagEnd && not (System.Char.IsWhiteSpace text.[at]) && text.[at] <> '>' && text.[at] <> '/' do at <- at + 1
    let mutable going = true
    while going do
        let wsStart = at
        while at < tagEnd && System.Char.IsWhiteSpace text.[at] do at <- at + 1
        if at >= tagEnd || text.[at] = '>' || text.[at] = '/' then going <- false
        else
            let nameStart = at
            while at < tagEnd && text.[at] <> '=' && not (System.Char.IsWhiteSpace text.[at]) do at <- at + 1
            let name = text.Substring(nameStart, at - nameStart)
            while at < tagEnd && text.[at] <> '"' && text.[at] <> '\'' do at <- at + 1
            if at >= tagEnd then going <- false
            else
                let quote = text.[at]
                let valueStart = at + 1
                let valueEnd = text.IndexOf(quote, valueStart)
                if valueEnd < 0 || valueEnd >= tagEnd then going <- false
                else
                    at <- valueEnd + 1
                    found.Add ((name, wsStart, at, valueStart, valueEnd))
    List.ofSeq found

/// The text of `node` (an element of `text`) with `edits` made in it.
let rec private applyAt (order: SchemaOrder) (text: string) (node: XmlNode) (edits: ElementEdit list) : string =
    let span = spanOf node
    if edits.IsEmpty then text.Substring(span.Start, span.End - span.Start)
    else
        let name = nameOf node
        // (Text that needs its white space kept says so.)
        let edits =
            match edits |> List.tryPick (function SetText value -> Some value | _ -> None) with
            | Some value when needsPreserve value && not ((attrsOf node).ContainsKey "xml:space") && not (edits |> List.exists (function SetAttribute ("xml:space", _) -> true | _ -> false)) ->
                edits @ [ SetAttribute ("xml:space", "preserve") ]
            | _ -> edits
        // The start tag, its attributes changed.
        let attributes = attributeRanges text span
        let tagClose = if text.[span.StartTagEnd - 2] = '/' then span.StartTagEnd - 2 else span.StartTagEnd - 1   // at "/>" or ">"
        let tagClose = (let mutable at = tagClose in (while at > span.Start && System.Char.IsWhiteSpace text.[at - 1] do at <- at - 1); at)
        let tagParts = ResizeArray<int * int * string>()
        let added = System.Text.StringBuilder()
        for edit in edits do
            match edit with
            | SetAttribute (attr, value) ->
                match attributes |> List.tryFind (fun (n, _, _, _, _) -> n = attr) with
                | Some (_, _, _, valueStart, valueEnd) -> tagParts.Add ((valueStart, valueEnd, escapeAttribute value))
                | None -> added.Append(" " + attr + "=\"" + escapeAttribute value + "\"") |> ignore
            | RemoveAttribute attr ->
                match attributes |> List.tryFind (fun (n, _, _, _, _) -> n = attr) with
                | Some (_, start, finish, _, _) -> tagParts.Add ((start, finish, ""))
                | None -> ()
            | _ -> ()
        if added.Length > 0 then tagParts.Add ((tagClose, tagClose, added.ToString()))
        match edits |> List.tryPick (function SetAllAttributes all -> Some all | _ -> None) with
        | Some all ->
            tagParts.Clear()
            let nameEnd = span.Start + 1 + name.Length
            tagParts.Add ((nameEnd, tagClose, all |> List.map (fun (attr, value) -> " " + attr + "=\"" + escapeAttribute value + "\"") |> String.concat ""))
        | None -> ()
        let spliced (from: int) (until: int) (parts: (int * int * string) seq) =
            let out = System.Text.StringBuilder()
            let mutable at = from
            for (start, finish, replacement) in parts |> Seq.sortBy (fun (start, finish, _) -> start, finish) do
                if start >= at then
                    out.Append(text, at, start - at).Append(replacement) |> ignore
                    at <- finish
            out.Append(text, at, until - at).ToString()
        let selfClosed = span.StartTagEnd = span.End
        let openTag =
            if selfClosed then spliced span.Start tagClose tagParts + ">"
            else spliced span.Start span.StartTagEnd tagParts
        // The children.
        let children = elementChildren node |> List.filter (fun kid -> Span.isKnown (spanOf kid))
        let ranks = order name
        let rankOf (child: string) =
            match ranks |> List.tryFindIndex (fun known -> known = local child) with
            | Some rank -> Some rank
            | None -> ranks |> List.tryFindIndex (fun known -> known = "*")
        let contentStart, contentEnd = if selfClosed then span.End, span.End else span.StartTagEnd, span.EndTagStart
        // Where a child of that name goes that is not there: before the first child of a later rank (one of
        // no rank has the rank of the one before it), else after the last.
        let placeFor (child: string) =
            match rankOf child with
            | None -> contentEnd
            | Some rank ->
                let mutable last = -1
                let later =
                    children |> List.tryFind (fun kid ->
                        let kidRank = match rankOf (nameOf kid) with Some r -> r | None -> last
                        last <- max last kidRank
                        kidRank > rank)
                match later with Some kid -> (spanOf kid).Start | None -> contentEnd
        // (The white space before a child that is taken out goes with it.)
        let withSpaceBefore (start: int) =
            let mutable at = start
            while at > contentStart && System.Char.IsWhiteSpace text.[at - 1] do at <- at - 1
            if at > contentStart && text.[at - 1] <> '>' then start else at
        let parts = ResizeArray<int * int * string>()
        let inserts = ResizeArray<int * int * string>()
        for edit in edits do
            match edit with
            | SetText value -> parts.Add ((contentStart, contentEnd, escapeText value))
            | _ -> ()
        let first (child: string) = children |> List.tryFind (fun kid -> nameOf kid = child)
        // (A child changed by edits: taken out, or its text with them made - with what goes before and after it.)
        let edited (kid: XmlNode) (inner: ElementEdit list) =
            if List.contains Removed inner then (withSpaceBefore (spanOf kid).Start, (spanOf kid).End, "")
            else
                let before = inner |> List.choose (function InsertBefore xml -> Some xml | _ -> None) |> String.concat ""
                let after = inner |> List.choose (function InsertAfter xml -> Some xml | _ -> None) |> String.concat ""
                let own = inner |> List.filter (function InsertBefore _ | InsertAfter _ -> false | _ -> true)
                match own |> List.tryPick (function Replaced xml -> Some xml | _ -> None) with
                | Some xml -> ((spanOf kid).Start, (spanOf kid).End, before + xml + after)
                | None -> ((spanOf kid).Start, (spanOf kid).End, before + applyAt order text kid own + after)
        // (Where what is put in at the start goes: before the first child the order does not put first.)
        let startPlace =
            match ranks |> List.tryFindIndex (fun known -> known = "*") with
            | None -> contentStart
            | Some star ->
                match children |> List.tryFind (fun kid -> match rankOf (nameOf kid) with Some rank -> rank >= star | None -> true) with
                | Some kid -> (spanOf kid).Start
                | None -> contentEnd
        let where (child: string) (attribute: string) (value: string) =
            children |> List.filter (fun kid -> nameOf kid = child && Map.tryFind attribute (attrsOf kid) = Some value)
        // (The children by where they start, for edits that name them so.)
        let byStart =
            lazy (let found = System.Collections.Generic.Dictionary<int, XmlNode>()
                  for kid in children do found.[(spanOf kid).Start] <- kid
                  found)
        for edit in edits do
            match edit with
            | SetChild (child, xml) ->
                match first child with
                | Some kid -> parts.Add ((spanOf kid).Start, (spanOf kid).End, xml)
                | None -> inserts.Add ((placeFor child, defaultArg (rankOf child) 9999, xml))
            | SetChildren (child, xmls) ->
                let all = children |> List.filter (fun kid -> nameOf kid = child)
                match all with
                | head :: rest ->
                    parts.Add ((spanOf head).Start, (spanOf head).End, String.concat "" xmls)
                    for kid in rest do parts.Add ((withSpaceBefore (spanOf kid).Start, (spanOf kid).End, ""))
                | [] -> if not xmls.IsEmpty then inserts.Add ((placeFor child, defaultArg (rankOf child) 9999, String.concat "" xmls))
            | RemoveChild child ->
                match first child with
                | Some kid -> parts.Add ((withSpaceBefore (spanOf kid).Start, (spanOf kid).End, ""))
                | None -> ()
            | EditChild (child, inner, whole) ->
                match first child, whole with
                | Some kid, _ -> parts.Add (edited kid inner)
                | None, Some xml -> inserts.Add ((placeFor child, defaultArg (rankOf child) 9999, xml))
                | None, None -> ()
            | AddChild (child, xml) -> inserts.Add ((placeFor child, defaultArg (rankOf child) 9999, xml))
            | Prepend xml -> inserts.Add ((startPlace, -1, xml))
            | Append xml -> inserts.Add ((contentEnd, System.Int32.MaxValue, xml))
            | EditNth (child, index, inner) ->
                match children |> List.filter (fun kid -> nameOf kid = child) |> List.tryItem index with
                | Some kid -> parts.Add (edited kid inner)
                | None -> ()
            | SetChildWhere (child, attribute, value, xml) ->
                match where child attribute value with
                | kid :: _ -> parts.Add ((spanOf kid).Start, (spanOf kid).End, xml)
                | [] -> inserts.Add ((placeFor child, defaultArg (rankOf child) 9999, xml))
            | RemoveChildWhere (child, attribute, value) ->
                for kid in where child attribute value do parts.Add ((withSpaceBefore (spanOf kid).Start, (spanOf kid).End, ""))
            | EditChildWhere (child, attribute, value, inner) ->
                match where child attribute value with
                | kid :: _ -> parts.Add (edited kid inner)
                | [] -> ()
            | EditAt (start, inner) ->
                match byStart.Value.TryGetValue start with
                | true, kid -> parts.Add (edited kid inner)
                | false, _ -> ()
            | _ -> ()
        // (Several put in at one place go in their schema order.)
        for (at, _, xml) in inserts |> Seq.sortBy (fun (at, rank, _) -> at, rank) do parts.Add ((at, at, xml))
        let content = spliced contentStart contentEnd parts
        if selfClosed then
            if content = "" then spliced span.Start span.End tagParts
            else openTag + content + "</" + name + ">"
        else openTag + content + text.Substring(span.EndTagStart, span.End - span.EndTagStart)

/// The text of an element read from a text with `edits` made in it.
let apply (order: SchemaOrder) (source: SourceElement) (edits: ElementEdit list) : string =
    applyAt order source.Text source.Node edits

/// The edits of `node` that make `targets` - elements of its tree, at any depth (itself too), each with its
/// edits - in it: each reached through the children it is in (EditNth), one chain of edits for those in the
/// same child.
let rec editsAt (node: XmlNode) (targets: (XmlNode * ElementEdit list) list) : ElementEdit list =
    let span = spanOf node
    let own = targets |> List.filter (fun (target, _) -> spanOf target = span) |> List.collect snd
    let within (outer: Span) (target: XmlNode) = let inner = spanOf target in inner.Start >= outer.Start && inner.End <= outer.End
    let kids = elementChildren node |> List.filter (fun kid -> Span.isKnown (spanOf kid))
    let childEdits =
        [ for kid in kids do
              let inside = targets |> List.filter (fun (target, _) -> spanOf target <> span && within (spanOf kid) target)
              if not inside.IsEmpty then
                  let index = kids |> List.filter (fun other -> nameOf other = nameOf kid) |> List.findIndex (fun other -> spanOf other = spanOf kid)
                  yield EditNth (nameOf kid, index, editsAt kid inside) ]
    own @ childEdits

/// The elements of a fragment of XML (one after another, no root), each with its text.
let readFragment (text: string) : SourceElement list option =
    let wrapped = "<fragment>" + text + "</fragment>"
    match tryScan wrapped (System.Collections.Generic.HashSet<string>()) with
    | Some (Element (_, _, kids, _)) ->
        Some [ for kid in kids do
                 match kid with
                 | Element _ -> yield { Text = wrapped; Node = kid }
                 | _ -> () ]
    | _ -> None

/// A fragment of XML (elements one after another) as the children of one element, to be changed as one
/// (`applyFragment`). None when it is not well formed.
let readFragmentRoot (text: string) : SourceElement option =
    let wrapped = "<fragment>" + text + "</fragment>"
    match tryScan wrapped (System.Collections.Generic.HashSet<string>()) with
    | Some (Element _ as node) -> Some { Text = wrapped; Node = node }
    | _ -> None

/// The fragment with these edits made in it (edits of the element `readFragmentRoot` makes of it).
let applyFragment (order: SchemaOrder) (fragment: SourceElement) (edits: ElementEdit list) : string =
    let edited = applyAt order fragment.Text fragment.Node edits
    match tryScan edited (System.Collections.Generic.HashSet<string>()) with
    | Some (Element (_, _, _, span)) when Span.isKnown span -> edited.Substring(span.StartTagEnd, span.EndTagStart - span.StartTagEnd)
    | _ -> edited

/// The root element of a part's text (what stands around it - the XML declaration - stays): None when the
/// text is not well formed.
let readPart (text: string) : SourceElement option =
    match tryScan text (System.Collections.Generic.HashSet<string>()) with
    | Some (Element (_, _, _, span) as node) when Span.isKnown span -> Some { Text = text; Node = node }
    | _ -> None

/// A part's text with these edits made in its root element.
let applyPart (order: SchemaOrder) (part: SourceElement) (edits: ElementEdit list) : string =
    let span = spanOf part.Node
    part.Text.Substring(0, span.Start) + applyAt order part.Text part.Node edits + part.Text.Substring span.End

/// The children of an element read whose name is `name` (qualified).
let childrenNamed (source: SourceElement) (name: string) : XmlNode list =
    elementChildren source.Node |> List.filter (fun kid -> nameOf kid = name)

/// The edits that make `was` (an element as the model wrote it) `now` (as it writes it now): attributes
/// set and taken out, children by name put in, taken out and changed - a child of a name there are several
/// of, or with text in it, is set whole.
let rec diff (was: SourceElement) (now: SourceElement) : ElementEdit list =
    diffNodes was.Text was.Node now.Text now.Node
and private diffNodes (wasText: string) (wasNode: XmlNode) (nowText: string) (nowNode: XmlNode) : ElementEdit list =
    let slice (text: string) (node: XmlNode) = let span = spanOf node in text.Substring(span.Start, span.End - span.Start)
    let wasAttrs, nowAttrs = attrsOf wasNode, attrsOf nowNode
    let attributeEdits =
        [ for KeyValue (name, value) in nowAttrs do
              if Map.tryFind name wasAttrs <> Some value then yield SetAttribute (name, value)
          // (A text's xml:space="preserve" is never wrong to keep: the model writes it only where its text needs it.)
          for KeyValue (name, _) in wasAttrs do
              if not (nowAttrs.ContainsKey name) && name <> "xml:space" then yield RemoveAttribute name ]
    let wasKids, nowKids = elementChildren wasNode, elementChildren nowNode
    // (Whether the children are of the same names in the same order: each can be changed in its place.)
    let sameOrder = (wasKids |> List.map nameOf) = (nowKids |> List.map nameOf)
    let names = (wasKids @ nowKids) |> List.map nameOf |> List.distinct
    let childEdits =
        [ for name in names do
              let before = wasKids |> List.filter (fun kid -> nameOf kid = name)
              let after = nowKids |> List.filter (fun kid -> nameOf kid = name)
              // (One child changed: its edits; `edit` makes the edit of it.)
              let changed (edit: ElementEdit list -> ElementEdit) (a: XmlNode) (b: XmlNode) =
                  let a', b' = slice wasText a, slice nowText b
                  if a' = b' then []
                  else
                      let onlyText (node: XmlNode) = elementChildren node |> List.isEmpty
                      if (hasText a || hasText b) && onlyText a && onlyText b then
                          // (Its attributes as the others', and its text.)
                          let attributes = diffNodes wasText (Element (nameOf a, attrsOf a, [], Span.none)) nowText (Element (nameOf b, attrsOf b, [], Span.none))
                          [ edit (attributes @ (if textOf a <> textOf b then [ SetText (textOf b) ] else [])) ]
                      elif hasText a || hasText b then [ SetChildren (name, [ b' ]) ]
                      else [ edit (diffNodes wasText a nowText b) ]
              match before, after with
              | [ a ], [ b ] -> yield! changed (fun edits -> EditChild (name, edits, Some (slice nowText b))) a b
              | [], [ b ] -> yield SetChild (name, slice nowText b)
              | [ _ ], [] -> yield RemoveChild name
              | _ :: _, [] -> yield SetChildren (name, [])
              // (As many as there were, in the same order among the others: each changed in its place.)
              | _ when before.Length = after.Length && sameOrder ->
                  for index, (a, b) in List.indexed (List.zip before after) do
                      yield! changed (fun edits -> EditNth (name, index, edits)) a b
              | _ ->
                  let a' = before |> List.map (slice wasText)
                  let b' = after |> List.map (slice nowText)
                  if a' <> b' then yield SetChildren (name, b') ]
    attributeEdits @ childEdits

/// `source` (an element as it was read; "" for one there was none of, `empty` then) with the changes the
/// model made between writing it `was` and `now` ("" for none) made in it: what the model does not write -
/// attributes, children - stays as it was read. None when one of the three is not one element.
let merge (order: SchemaOrder) (empty: string) (source: string) (was: string) (now: string) : string option =
    let orEmpty (text: string) = if text = "" then empty else text
    match read (orEmpty source), read (orEmpty was), read (orEmpty now) with
    | Some s, Some w, Some n -> Some (apply order s (diff w n))
    | _ -> None

/// The child an edit is of (qualified name); "" for an edit of an attribute.
let childOf (edit: ElementEdit) : string =
    match edit with
    | SetChild (name, _) | SetChildren (name, _) | RemoveChild name | EditChild (name, _, _) | AddChild (name, _)
    | SetChildWhere (name, _, _, _) | RemoveChildWhere (name, _, _) | EditChildWhere (name, _, _, _) | EditNth (name, _, _) -> name
    | SetAttribute _ | SetAllAttributes _ | RemoveAttribute _ | SetText _ | Removed | InsertBefore _ | InsertAfter _ | Prepend _ | Append _ | EditAt _ | Replaced _ -> ""

/// Whether the element has a child of that name (qualified).
let hasChild (source: SourceElement) (name: string) : bool =
    elementChildren source.Node |> List.exists (fun kid -> nameOf kid = name)

/// The attributes of the element's child of that name (qualified; the first, where there are several).
let childAttributes (source: SourceElement) (name: string) : Map<string, string> =
    elementChildren source.Node |> List.tryFind (fun kid -> nameOf kid = name) |> Option.map attrsOf |> Option.defaultValue Map.empty

/// An element's text with another name: its start tag's and end tag's name changed, nothing else.
let renamed (name: string) (text: string) : string =
    match read text with
    | Some { Node = Element (old, _, _, span) } ->
        let close = if span.StartTagEnd = span.End then "" else "</" + name + ">"
        "<" + name + text.Substring(span.Start + 1 + old.Length, span.StartTagEnd - span.Start - 1 - old.Length)
        + (if close = "" then "" else text.Substring(span.StartTagEnd, span.EndTagStart - span.StartTagEnd) + close)
    | _ -> text

/// The same edit of the child of another name.
let ofChild (name: string) (edit: ElementEdit) : ElementEdit =
    match edit with
    | SetChild (_, xml) -> SetChild (name, renamed name xml)
    | SetChildren (_, xmls) -> SetChildren (name, List.map (renamed name) xmls)
    | RemoveChild _ -> RemoveChild name
    | EditChild (_, inner, whole) -> EditChild (name, inner, whole |> Option.map (renamed name))
    | other -> other

/// The text of an element as it was read: its bytes.
let bytes (source: SourceElement) : string =
    let span = spanOf source.Node
    source.Text.Substring(span.Start, span.End - span.Start)

/// The text of a node of a text read (its span is into `text`).
let bytesOf (text: string) (node: XmlNode) : string =
    let span = spanOf node
    text.Substring(span.Start, span.End - span.Start)

/// The text of a text read from `start` to `finish` (places from the span tree).
let slice (text: string) (start: int) (finish: int) : string = text.Substring(start, finish - start)

/// The start tag of an element of a text read, with these attribute edits made in it.
let startTag (text: string) (node: XmlNode) (edits: ElementEdit list) : string =
    let span = spanOf node
    let tag = { Start = span.Start; StartTagEnd = span.StartTagEnd; EndTagStart = span.StartTagEnd; End = span.StartTagEnd }
    let attributeEdits = edits |> List.filter (function SetAttribute _ | SetAllAttributes _ | RemoveAttribute _ -> true | _ -> false)
    if attributeEdits.IsEmpty then slice text span.Start span.StartTagEnd
    else
        // (The start tag alone, as an element that ends with it.)
        applyAt (fun _ -> []) text (Element (nameOf node, attrsOf node, [], tag)) attributeEdits

// ---- Items among what stands around them ----
// The children of a table are rows, of a row cells, of a cell paragraphs and tables, of a paragraph runs:
// each in the wrappers it is in (a content control, custom XML, a smart tag ...), among elements the model
// does not hold (marks, unknown elements) and white space. A layout is that, as it was read; the writer
// writes the items it has (each its own way) and the layout puts them back among the rest: an item of the
// layout in the wrappers it was in, where it stood; one that is new in the wrappers of the one before it;
// what the model does not hold where it stood.

/// One child of a layout: an item (where its element starts) or an element kept as it is, each with the
/// wrappers around it (where each starts, outermost first).
type LayoutChild =
    | LaidItem of at: int * wrappers: int list
    | LaidKept of at: int * wrappers: int list

type Layout = {
    Text: string
    /// Wrapper (where it starts) -> its opening: from the text before it to the start of what it holds
    /// (with what says what it is: w:sdtPr ...); and its closing: from after what it holds to its end.
    Wrappers: System.Collections.Generic.Dictionary<int, (int * int) * (int * int)>
    Children: LayoutChild list
    /// Item or kept element (where it starts) -> the range of the text before it (white space).
    Leads: System.Collections.Generic.Dictionary<int, int * int>
    /// Kept element (where it starts) -> where it ends.
    KeptEnds: System.Collections.Generic.Dictionary<int, int>
    /// After the last child, to the end of the element's contents.
    Tail: int * int
}

/// The layout of the children of `node` (an element of `text`) from `from` on (after the properties):
/// the items (children `isItem` says are, handed to `take`), in the wrappers they are in - a child
/// that holds items, at any depth, and is not one `passThrough` says the model holds itself (then its own
/// children are taken as the element's) -; elements `skip` says the model writes itself (or are to go) left
/// out; the rest kept. None when an element of it has no place in the text.
let layout (isItem: XmlNode -> bool) (passThrough: string -> bool) (skip: XmlNode -> bool) (take: XmlNode -> unit) (text: string) (node: XmlNode) (from: int) : Layout option =
    let span = spanOf node
    let wrappers = System.Collections.Generic.Dictionary<int, (int * int) * (int * int)>()
    let leads = System.Collections.Generic.Dictionary<int, int * int>()
    let keptEnds = System.Collections.Generic.Dictionary<int, int>()
    let found = ResizeArray<LayoutChild>()
    let mutable sound = Span.isKnown span && span.StartTagEnd < span.End
    let rec holds (kid: XmlNode) = isItem kid || List.exists holds (elementChildren kid)
    // The children of one element from `start` on: where the last ends.
    let rec visitAll (chain: int list) (start: int) (nodes: XmlNode list) : int =
        let mutable at = start
        for kid in nodes do
            let kidSpan = spanOf kid
            match kid with
            | Element _ when not (Span.isKnown kidSpan) -> sound <- false
            | Element _ when kidSpan.End <= from -> ()
            | Element _ ->
                let name = local (nameOf kid)
                if isItem kid then
                    found.Add (LaidItem (kidSpan.Start, chain))
                    leads.[kidSpan.Start] <- (at, kidSpan.Start)
                    take kid
                    at <- kidSpan.End
                elif passThrough name then
                    // (Its start and end tags are the model's: what it holds begins after its start tag.)
                    visitAll chain kidSpan.StartTagEnd (match kid with Element (_, _, kids, _) -> kids | _ -> []) |> ignore
                    at <- kidSpan.End
                elif skip kid then
                    at <- kidSpan.End
                elif holds kid && kidSpan.StartTagEnd < kidSpan.End then
                    let kids = elementChildren kid
                    let content, headEnd =
                        match kids |> List.tryFind (fun inner -> local (nameOf inner) = "sdtContent") with
                        | Some (Element (_, _, inner, contentSpan)) when name = "sdt" && contentSpan.StartTagEnd < contentSpan.End ->
                            inner, contentSpan.StartTagEnd
                        | _ ->
                            let leading = kids |> List.takeWhile (fun inner -> (local (nameOf inner)).EndsWith "Pr")
                            let headEnd = leading |> List.fold (fun last inner -> let innerSpan = spanOf inner in if Span.isKnown innerSpan then innerSpan.End else last) kidSpan.StartTagEnd
                            List.skip leading.Length kids, headEnd
                    let last = visitAll (chain @ [ kidSpan.Start ]) headEnd content
                    // (Its closing: from after what it holds through its end tag - an sdtContent's end tag too.)
                    wrappers.[kidSpan.Start] <- ((at, headEnd), (last, kidSpan.End))
                    at <- kidSpan.End
                else
                    found.Add (LaidKept (kidSpan.Start, chain))
                    leads.[kidSpan.Start] <- (at, kidSpan.Start)
                    keptEnds.[kidSpan.Start] <- kidSpan.End
                    at <- kidSpan.End
            | _ -> ()
        at
    let last = visitAll [] from (match node with Element (_, _, kids, _) -> kids | _ -> [])
    if not sound then None
    else Some { Text = text; Wrappers = wrappers; Children = List.ofSeq found; Leads = leads; KeptEnds = keptEnds; Tail = (last, span.EndTagStart) }

/// Where an item that is written goes in a layout.
type Placed =
    /// It is the item of the layout that starts there: in the wrappers that one was in, where it stood.
    | AtItem of int
    /// New: in the wrappers of the one before it.
    | WithPrevious
    /// A further part of the item of the layout that starts there (a run split): in the wrappers that one
    /// was in, after what is written before it.
    | InItem of int
    /// New: outside the wrappers of the layout (what it is in, the XML written has).
    | Outside

/// How a layout is written: for a part of the element (a half of a paragraph split), whether the elements
/// kept before its first item and after its last are its (`KeptBefore`, `KeptAfter`); the ranges of the
/// text an item that is written stands for, what is in them with it (`Covered`: a field written whole).
type LayoutWrite = { KeptBefore: bool; KeptAfter: bool; Covered: (int * int) list }
let wholeElement = { KeptBefore = true; KeptAfter = true; Covered = [] }

/// The contents of the element of a layout with the items written: `written` is each item's XML in order,
/// with the item of the layout it is (where that starts), when it is one - an item of the layout that is not
/// written is taken out, with what stood just before it; one with no place in the layout is new.
let rec writeLayout (laid: Layout) (written: (Placed * string) array) : string = writeLayoutWith wholeElement laid written
and writeLayoutWith (how: LayoutWrite) (laid: Layout) (written: (Placed * string) array) : string =
    let out = System.Text.StringBuilder()
    let slice (range: int * int) = let (a, b) = range in laid.Text.Substring(a, b - a)
    let chainOf = System.Collections.Generic.Dictionary<int, int list>()
    for child in laid.Children do (match child with LaidItem (at, chain) -> chainOf.[at] <- chain | _ -> ())
    let mutable openChain : int list = []
    let moveTo (chain: int list) =
        let rec common (a: int list) (b: int list) = match a, b with x :: xs, y :: ys when x = y -> 1 + common xs ys | _ -> 0
        let keep = common openChain chain
        for wrapper in List.rev (List.skip keep openChain) do out.Append(slice (snd laid.Wrappers.[wrapper])) |> ignore
        for wrapper in List.skip keep chain do out.Append(slice (fst laid.Wrappers.[wrapper])) |> ignore
        openChain <- chain
    let covered (start: int) = how.Covered |> List.exists (fun (a, b) -> start > a && start < b)
    let mutable rest = laid.Children |> List.filter (function LaidItem (start, _) | LaidKept (start, _) -> not (covered start))
    let mutable started = how.KeptBefore
    let emitBefore (at: int option) =
        let mutable going = true
        while going && not rest.IsEmpty do
            match rest.Head with
            | LaidItem (other, _) when Some other = at -> going <- false
            | LaidItem _ -> rest <- rest.Tail
            | LaidKept (start, chain) ->
                if started && (at.IsSome || how.KeptAfter) then
                    moveTo chain
                    out.Append(slice laid.Leads.[start]).Append(laid.Text, start, laid.KeptEnds.[start] - start) |> ignore
                rest <- rest.Tail
    let mutable placedAny = false
    for placed, xml in written do
        // (New content of an element that had no items: after what it held - the marks of a paragraph without
        // runs are at its start still.)
        match placed with
        | AtItem at when chainOf.ContainsKey at -> ()
        | _ when not placedAny && not (laid.Children |> List.exists (function LaidItem _ -> true | _ -> false)) -> emitBefore None
        | _ -> ()
        placedAny <- true
        match placed with
        | AtItem at when chainOf.ContainsKey at ->
            if rest |> List.exists (function LaidItem (other, _) -> other = at | _ -> false) then emitBefore (Some at)
            started <- true
            moveTo chainOf.[at]
            out.Append(slice laid.Leads.[at]) |> ignore
        | InItem at when chainOf.ContainsKey at -> moveTo chainOf.[at]
        | Outside -> moveTo []
        | _ -> ()
        out.Append(xml) |> ignore
    emitBefore None
    moveTo []
    (if how.KeptAfter then out.Append(slice laid.Tail) else out).ToString()

/// The qualified names of an element's attributes.
let attributeNames (node: XmlNode) : string list = attrsOf node |> Map.toList |> List.map fst

/// The attributes of an element (qualified name -> value).
let attributes (node: XmlNode) : Map<string, string> = attrsOf node

/// The qualified name of an element ("" for text).
let name (node: XmlNode) : string = nameOf node

/// Where an element is in the text it was read from.
let span (node: XmlNode) : Span = spanOf node

/// The prefix of an element's name ("" for none).
let prefix (node: XmlNode) : string =
    let qualified = nameOf node
    match qualified.IndexOf ':' with
    | -1 -> ""
    | at -> qualified.Substring(0, at)

/// The local name of an element (without its prefix).
let localName (node: XmlNode) : string = local (nameOf node)

/// The elements an element holds, in order.
let elements (node: XmlNode) : XmlNode list = elementChildren node

/// Every attribute of an element and of the elements in it, at any depth (qualified name, value), in
/// document order.
let rec allAttributes (node: XmlNode) : (string * string) list =
    match node with
    | Element (_, attrs, kids, _) -> (attrs |> Map.toList) @ List.collect allAttributes kids
    | _ -> []

/// The text of a text read from `start` to `finish` without the elements in it that `drop` names (local names).
let sliceWithout (text: string) (node: XmlNode) (start: int) (finish: int) (drop: string -> bool) : string =
    let gone = ResizeArray<int * int>()
    let rec visit (kid: XmlNode) =
        match kid with
        | Element (name, _, kids, span) when Span.isKnown span && span.End > start && span.Start < finish ->
            if span.Start >= start && span.End <= finish && drop (local name) then gone.Add ((span.Start, span.End))
            else List.iter visit kids
        | _ -> ()
    visit node
    let out = System.Text.StringBuilder()
    let mutable at = start
    for (a, b) in gone |> Seq.sortBy fst do
        if a >= at then
            out.Append(text, at, a - at) |> ignore
            at <- b
    out.Append(text, at, finish - at).ToString()

/// The text of a text read from `start` to `finish` with the elements in it at the places given replaced
/// (where each starts and ends, and what goes there).
let sliceWith (text: string) (start: int) (finish: int) (replacements: (int * int * string) list) : string =
    let out = System.Text.StringBuilder()
    let mutable at = start
    for (a, b, xml) in replacements |> List.sortBy (fun (a, _, _) -> a) do
        if a >= at && b <= finish then
            out.Append(text, at, a - at).Append(xml) |> ignore
            at <- b
    out.Append(text, at, finish - at).ToString()

/// The value of an attribute of an element (qualified name).
let attribute (node: XmlNode) (name: string) : string option = Map.tryFind name (attrsOf node)

/// The elements of a tree (itself too) `wanted` says, in document order.
let rec descendantsWhere (wanted: XmlNode -> bool) (node: XmlNode) : XmlNode list =
    match node with
    | Element (_, _, kids, _) -> (if wanted node then [ node ] else []) @ List.collect (descendantsWhere wanted) kids
    | _ -> []

/// The elements of a tree whose local name is `name`, in document order.
let rec descendants (name: string) (node: XmlNode) : XmlNode list =
    match node with
    | Element (n, _, kids, _) -> (if local n = name then [ node ] else []) @ List.collect (descendants name) kids
    | _ -> []

/// The bytes of the children of an element read whose local name `keep` says, one after another.
let childrenBytes (source: SourceElement) (keep: string -> bool) : string =
    elementChildren source.Node |> List.filter (fun kid -> keep (local (nameOf kid))) |> List.map (bytesOf source.Text) |> String.concat ""

/// The child of an element whose local name is `name`.
let child (name: string) (node: XmlNode) : XmlNode option = elementChildren node |> List.tryFind (fun kid -> local (nameOf kid) = name)

/// Whether a layout's item (where it starts) is one of it.
let isItemOf (laid: Layout) (at: int) : bool = laid.Leads.ContainsKey at && not (laid.KeptEnds.ContainsKey at)

/// Whether an element's text says nothing: no attributes, no children.
let isEmpty (text: string) : bool =
    match read text with
    | Some { Node = Element (_, attrs, kids, _) } -> attrs.IsEmpty && kids |> List.forall (function Element _ -> false | Text t -> t.Trim() = "" | Empty -> true)
    | _ -> text = ""

/// The schema order of the children of the WordprocessingML elements the writer changes (local names).
module Word =
    let private runProperties =
        [ "rStyle"; "rFonts"; "b"; "bCs"; "i"; "iCs"; "caps"; "smallCaps"; "strike"; "dstrike"; "outline"; "shadow"; "emboss"; "imprint"
          "noProof"; "snapToGrid"; "vanish"; "webHidden"; "color"; "spacing"; "w"; "kern"; "position"; "sz"; "szCs"; "highlight"; "u"
          "effect"; "bdr"; "shd"; "fitText"; "vertAlign"; "rtl"; "cs"; "em"; "lang"; "eastAsianLayout"; "specVanish"; "oMath"; "ins"; "del"; "rPrChange" ]
    let private paragraphProperties =
        [ "pStyle"; "keepNext"; "keepLines"; "pageBreakBefore"; "framePr"; "widowControl"; "numPr"; "suppressLineNumbers"; "pBdr"; "shd"
          "tabs"; "suppressAutoHyphens"; "kinsoku"; "wordWrap"; "overflowPunct"; "topLinePunct"; "autoSpaceDE"; "autoSpaceDN"; "bidi"
          "adjustRightInd"; "snapToGrid"; "spacing"; "ind"; "contextualSpacing"; "mirrorIndents"; "suppressOverlap"; "jc"; "textDirection"
          "textAlignment"; "textboxTightWrap"; "outlineLvl"; "divId"; "cnfStyle"; "rPr"; "sectPr"; "pPrChange" ]
    let private cellProperties =
        [ "cnfStyle"; "tcW"; "gridSpan"; "hMerge"; "vMerge"; "tcBorders"; "shd"; "noWrap"; "tcMar"; "textDirection"; "tcFitText"; "vAlign"; "hideMark"; "headers"; "cellIns"; "cellDel"; "cellMerge"; "tcPrChange" ]
    let private rowProperties =
        [ "cnfStyle"; "divId"; "gridBefore"; "gridAfter"; "wBefore"; "wAfter"; "cantSplit"; "trHeight"; "tblHeader"; "tblCellSpacing"; "jc"; "hidden"; "ins"; "del"; "trPrChange" ]
    let private tableProperties =
        [ "tblStyle"; "tblpPr"; "tblOverlap"; "bidiVisual"; "tblStyleRowBandSize"; "tblStyleColBandSize"; "tblW"; "jc"; "tblCellSpacing"; "tblInd"; "tblBorders"; "shd"; "tblLayout"; "tblCellMar"; "tblLook"; "tblCaption"; "tblDescription"; "tblPrChange" ]
    let private borders = [ "top"; "start"; "left"; "bottom"; "end"; "right"; "insideH"; "insideV"; "tl2br"; "tr2bl"; "between"; "bar" ]
    let private margins = [ "top"; "start"; "left"; "bottom"; "end"; "right" ]
    let order : SchemaOrder = function
        | "w:rPr" -> runProperties
        | "w:pPr" -> paragraphProperties
        | "w:tcPr" -> cellProperties
        | "w:trPr" -> rowProperties
        | "w:tblPr" -> tableProperties
        | "w:pBdr" | "w:tcBorders" | "w:tblBorders" -> borders
        | "w:tcMar" | "w:tblCellMar" -> margins
        | "w:tc" -> [ "tcPr"; "*" ]
        | "w:tr" -> [ "tblPrEx"; "trPr"; "*" ]
        | "w:tbl" -> [ "tblPr"; "tblGrid"; "*" ]
        | "w:p" -> [ "pPr"; "*" ]
        | "w:r" -> [ "rPr"; "*" ]
        | "w:numbering" -> [ "numPicBullet"; "abstractNum"; "num"; "numIdMacAtCleanup" ]
        | "w:abstractNum" -> [ "nsid"; "multiLevelType"; "tmpl"; "name"; "styleLink"; "numStyleLink"; "lvl" ]
        | "w:num" -> [ "abstractNumId"; "lvlOverride" ]
        | _ -> []

/// A style attribute's value (VML's style="position:absolute;margin-left:12pt"): its declarations in
/// order, each a property and its value.
module Style =
    let parse (value: string) : (string * string) list =
        value.Split(';')
        |> Array.toList
        |> List.choose (fun declaration ->
            match declaration.Split([| ':' |], 2) with
            | [| property; value |] when property.Trim() <> "" -> Some (property.Trim(), value.Trim())
            | _ -> None)

    let write (declarations: (string * string) list) : string =
        declarations |> List.map (fun (property, value) -> property + ":" + value) |> String.concat ";"

    /// The declarations with the property's value set: in its place, or after the others.
    let set (property: string) (value: string) (declarations: (string * string) list) =
        if declarations |> List.exists (fun (known, _) -> known = property) then
            declarations |> List.map (fun (known, was) -> if known = property then known, value else known, was)
        else declarations @ [ property, value ]

    let remove (property: string) (declarations: (string * string) list) =
        declarations |> List.filter (fun (known, _) -> known <> property)

// ---- A large part read in outline ----
// A part of many elements (a worksheet of many rows) is read with the elements a change is made in whole,
// and the many others - each a range of the text - without what they hold (`readPartExcept`). An element
// of those that a change is made in is then read whole where it is (`withWhole`), and the edits are made
// as in any other tree.

exception private OutlineGiveUp

/// An attribute value or text of an element with its references decoded.
let private decoded (raw: string) : string =
    if raw.IndexOf '&' < 0 then raw
    else
        let sb = System.Text.StringBuilder()
        let mutable i = 0
        while i < raw.Length do
            let amp = raw.IndexOf('&', i)
            if amp < 0 then
                sb.Append(raw, i, raw.Length - i) |> ignore
                i <- raw.Length
            else
                sb.Append(raw, i, amp - i) |> ignore
                let semi = raw.IndexOf(';', amp)
                if semi < 0 || semi - amp < 2 || semi - amp > 10 then raise OutlineGiveUp
                match raw.Substring(amp + 1, semi - amp - 1) with
                | "lt" -> sb.Append '<' |> ignore
                | "gt" -> sb.Append '>' |> ignore
                | "amp" -> sb.Append '&' |> ignore
                | "quot" -> sb.Append '"' |> ignore
                | "apos" -> sb.Append '\'' |> ignore
                | entity when entity.[0] = '#' ->
                    let hex = entity.Length > 1 && entity.[1] = 'x'
                    let digits = entity.Substring(if hex then 2 else 1)
                    let mutable code = 0
                    if digits = "" then raise OutlineGiveUp
                    for c in digits do
                        let digit =
                            if c >= '0' && c <= '9' then int c - 48
                            elif hex && c >= 'a' && c <= 'f' then int c - 87
                            elif hex && c >= 'A' && c <= 'F' then int c - 55
                            else raise OutlineGiveUp
                        code <- code * (if hex then 16 else 10) + digit
                        if code > 0x10FFFF then raise OutlineGiveUp
                    if code < 0x10000 then sb.Append(char code) |> ignore
                    else
                        let rest = code - 0x10000
                        sb.Append(char (0xD800 + (rest >>> 10))).Append(char (0xDC00 + (rest &&& 0x3FF))) |> ignore
                | _ -> raise OutlineGiveUp
                i <- semi + 1
        sb.ToString()

/// A part's text read with the elements whose local name `shallow` says without what they hold - their
/// children are not in the tree; their names, attributes and spans are -, the rest as `readPart` reads it.
/// None when the text is not well formed.
let readPartExcept (shallow: string -> bool) (text: string) : SourceElement option =
    let n = text.Length
    let isSpace (c: char) = c = ' ' || c = '\n' || c = '\t' || c = '\r'
    let ordinalAt (at: int) (value: string) =
        if at + value.Length > n then false
        else
            let mutable same = true
            let mutable k = 0
            while same && k < value.Length do
                if text.[at + k] <> value.[k] then same <- false
                k <- k + 1
            same
    let indexFrom (value: string) (from: int) =
#if FABLE_COMPILER
        text.IndexOf(value, from)
#else
        text.IndexOf(value, from, System.StringComparison.Ordinal)
#endif
    // The start tag at `lt`: its name, attributes, where it ends, whether it closes itself.
    let startTag (lt: int) =
        let mutable i = lt + 1
        while i < n && (let c = text.[i] in not (isSpace c) && c <> '>' && c <> '/') do i <- i + 1
        if i = lt + 1 || i >= n then raise OutlineGiveUp
        let name = text.Substring(lt + 1, i - lt - 1)
        let mutable attrs : Map<string, string> = Map.empty
        let mutable selfClosing = false
        let mutable going = true
        while going do
            while i < n && isSpace text.[i] do i <- i + 1
            if i >= n then raise OutlineGiveUp
            let c = text.[i]
            if c = '>' then
                going <- false
                i <- i + 1
            elif c = '/' then
                if i + 1 >= n || text.[i + 1] <> '>' then raise OutlineGiveUp
                selfClosing <- true
                going <- false
                i <- i + 2
            else
                let nameStart = i
                while i < n && (let a = text.[i] in a <> '=' && not (isSpace a) && a <> '>' && a <> '/') do i <- i + 1
                let attrName = text.Substring(nameStart, i - nameStart)
                while i < n && isSpace text.[i] do i <- i + 1
                if i >= n || text.[i] <> '=' || attrName = "" then raise OutlineGiveUp
                i <- i + 1
                while i < n && isSpace text.[i] do i <- i + 1
                if i >= n || (text.[i] <> '"' && text.[i] <> '\'') then raise OutlineGiveUp
                let close = text.IndexOf(text.[i], i + 1)
                if close < 0 then raise OutlineGiveUp
                let raw = text.Substring(i + 1, close - i - 1)
                if raw.IndexOf '<' >= 0 || attrs.ContainsKey attrName then raise OutlineGiveUp
                let raw = raw.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ')
                attrs <- Map.add attrName (decoded raw) attrs
                i <- close + 1
        name, attrs, i, selfClosing
    // What is not an element at `lt` (a comment, a processing instruction, CDATA): where it ends.
    let skipOther (lt: int) =
        let finish (closing: string) (from: int) = (let at = indexFrom closing from in if at < 0 then raise OutlineGiveUp else at + closing.Length)
        if ordinalAt lt "<?" then finish "?>" (lt + 2)
        elif ordinalAt lt "<!--" then finish "-->" (lt + 4)
        elif ordinalAt lt "<![CDATA[" then finish "]]>" (lt + 9)
        else raise OutlineGiveUp
    // The end tag of `name` at `lt`: where it ends.
    let endTag (lt: int) (name: string) =
        let gt = text.IndexOf('>', lt + 2)
        if gt < 0 then raise OutlineGiveUp
        let mutable nameEnd = gt
        while nameEnd > lt + 2 && isSpace text.[nameEnd - 1] do nameEnd <- nameEnd - 1
        if nameEnd - lt - 2 <> name.Length || not (ordinalAt (lt + 2) name) then raise OutlineGiveUp
        gt + 1
    // The element whose start tag is at `lt`, and whether a shallow element is in it (or is it): one that
    // has none in it is read by the scanner, whole.
    let rec element (lt: int) : XmlNode * bool =
        let name, attrs, tagEnd, selfClosing = startTag lt
        let isShallow = shallow (Ns.stripNs name)
        if selfClosing then Element (name, attrs, [], { Start = lt; StartTagEnd = tagEnd; EndTagStart = tagEnd; End = tagEnd }), isShallow
        elif isShallow then
            // (What it holds is passed over: to its end tag - past the elements of its name in it, each
            // to its own.)
            let opening = "<" + name
            let closing = "</" + name
            let isNameEnd (at: int) = at < n && (let c = text.[at] in isSpace c || c = '>' || c = '/')
            let mutable depth = 1
            let mutable at = tagEnd
            let mutable close = -1
            while depth > 0 do
                let ending = indexFrom closing at
                if ending < 0 then raise OutlineGiveUp
                let inner = indexFrom opening at
                if inner >= 0 && inner < ending then
                    if isNameEnd (inner + opening.Length) then
                        let gt = text.IndexOf('>', inner)
                        if gt < 0 then raise OutlineGiveUp
                        if text.[gt - 1] <> '/' then depth <- depth + 1
                        at <- gt + 1
                    else at <- inner + opening.Length
                elif isNameEnd (ending + closing.Length) then
                    depth <- depth - 1
                    if depth = 0 then close <- ending
                    at <- ending + closing.Length
                else at <- ending + closing.Length
            Element (name, attrs, [], { Start = lt; StartTagEnd = tagEnd; EndTagStart = close; End = endTag close name }), true
        else
            let kids = ResizeArray<XmlNode>()
            let mutable anyShallow = false
            let mutable at = tagEnd
            let mutable close = -1
            while close < 0 do
                let next = text.IndexOf('<', at)
                if next < 0 || next + 1 >= n then raise OutlineGiveUp
                if next > at then
                    let raw = text.Substring(at, next - at).Trim()
                    if raw <> "" then kids.Add (Text (decoded raw))
                let c = text.[next + 1]
                if c = '/' then close <- next
                elif c = '?' || c = '!' then
                    if ordinalAt next "<![CDATA[" then
                        let finish = indexFrom "]]>" (next + 9)
                        if finish < 0 then raise OutlineGiveUp
                        kids.Add (Text (text.Substring(next + 9, finish - next - 9)))
                    at <- skipOther next
                else
                    let kid, holds = element next
                    kids.Add kid
                    if holds then anyShallow <- true
                    at <- (spanOf kid).End
            let span = { Start = lt; StartTagEnd = tagEnd; EndTagStart = close; End = endTag close name }
            if anyShallow then Element (name, attrs, List.ofSeq kids, span), true
            else
                match tryScanAt text span.Start span.End (System.Collections.Generic.HashSet<string>()) with
                | Some whole -> whole, false
                | None -> raise OutlineGiveUp
    try
        let mutable at = 0
        let mutable root = None
        while root.IsNone do
            let lt = text.IndexOf('<', at)
            if lt < 0 || lt + 1 >= n || text.[lt + 1] = '/' then raise OutlineGiveUp
            if text.[lt + 1] = '?' || text.[lt + 1] = '!' then at <- skipOther lt
            else root <- Some (fst (element lt))
        Some { Text = text; Node = root.Value }
    with _ -> None

/// The tree of a part read in outline with the elements that start at these places (elements of it) read
/// whole.
let withWhole (source: SourceElement) (starts: int seq) : SourceElement =
    let wanted = System.Collections.Generic.HashSet<int>(starts)
    if wanted.Count = 0 then source
    else
        let sorted = Seq.toArray wanted |> Array.sort
        // (Whether one of them starts within the span: the tree changes under it.)
        let within (span: Span) =
            let mutable low = 0
            let mutable high = sorted.Length
            while low < high do
                let mid = (low + high) / 2
                if sorted.[mid] < span.Start then low <- mid + 1 else high <- mid
            low < sorted.Length && sorted.[low] < span.End
        let rec visit (node: XmlNode) =
            match node with
            | Element (name, attrs, kids, span) when Span.isKnown span && within span ->
                if wanted.Contains span.Start then
                    match tryScanAt source.Text span.Start span.End (System.Collections.Generic.HashSet<string>()) with
                    | Some read -> read
                    | None -> node
                else Element (name, attrs, List.map visit kids, span)
            | _ -> node
        { source with Node = visit source.Node }

/// An element of a part read in outline, read whole (with what it holds).
let whole (source: SourceElement) (node: XmlNode) : SourceElement option =
    let span = spanOf node
    if not (Span.isKnown span) then None
    else tryScanAt source.Text span.Start span.End (System.Collections.Generic.HashSet<string>()) |> Option.map (fun read -> { Text = source.Text; Node = read })

/// An element of a text read, as an element of its own (`apply`, `childrenNamed` ... of it).
let at (source: SourceElement) (node: XmlNode) : SourceElement = { Text = source.Text; Node = node }

/// The attributes of an element of a text read, in the order they are written: each its qualified name and
/// its text as it is written (with the white space before it).
let attributeBytes (text: string) (node: XmlNode) : (string * string) list =
    attributeRanges text (spanOf node) |> List.map (fun (name, start, finish, _, _) -> name, text.Substring(start, finish - start))

/// The schema order of the children of the SpreadsheetML elements the writer changes (by local name: a part
/// may name them with a prefix).
module Spreadsheet =
    let private worksheet =
        [ "sheetPr"; "dimension"; "sheetViews"; "sheetFormatPr"; "cols"; "sheetData"; "sheetCalcPr"; "sheetProtection"; "protectedRanges"; "scenarios"
          "autoFilter"; "sortState"; "dataConsolidate"; "customSheetViews"; "mergeCells"; "phoneticPr"; "conditionalFormatting"; "dataValidations"
          "hyperlinks"; "printOptions"; "pageMargins"; "pageSetup"; "headerFooter"; "rowBreaks"; "colBreaks"; "customProperties"; "cellWatches"
          "ignoredErrors"; "smartTags"; "drawing"; "legacyDrawing"; "legacyDrawingHF"; "drawingHF"; "picture"; "oleObjects"; "controls"
          "webPublishItems"; "tableParts"; "extLst" ]
    let private workbook =
        [ "fileVersion"; "fileSharing"; "workbookPr"; "workbookProtection"; "bookViews"; "sheets"; "functionGroups"; "externalReferences"
          "definedNames"; "calcPr"; "oleSize"; "customWorkbookViews"; "pivotCaches"; "smartTagPr"; "smartTagTypes"; "webPublishing"
          "fileRecoveryPr"; "webPublishObjects"; "extLst" ]
    let private styleSheet =
        [ "numFmts"; "fonts"; "fills"; "borders"; "cellStyleXfs"; "cellXfs"; "cellStyles"; "dxfs"; "tableStyles"; "colors"; "extLst" ]
    let private metadata =
        [ "metadataTypes"; "metadataStrings"; "mdxMetadata"; "futureMetadata"; "cellMetadata"; "valueMetadata"; "extLst" ]
    let order : SchemaOrder = fun name ->
        match local name with
        | "worksheet" -> worksheet
        | "workbook" -> workbook
        | "styleSheet" -> styleSheet
        | "metadata" -> metadata
        | "sheetPr" -> [ "tabColor"; "outlinePr"; "pageSetUpPr" ]
        | "sheetViews" -> [ "sheetView"; "extLst" ]
        | "sheetView" -> [ "pane"; "selection"; "pivotSelection"; "extLst" ]
        | "row" -> [ "c"; "extLst" ]
        | "c" -> [ "f"; "v"; "is"; "extLst" ]
        | "xf" -> [ "alignment"; "protection"; "extLst" ]
        | "sst" -> [ "si"; "extLst" ]
        | "comments" -> [ "authors"; "commentList"; "extLst" ]
        | "comment" -> [ "text"; "commentPr" ]
        | "table" -> [ "autoFilter"; "sortState"; "tableColumns"; "tableStyleInfo"; "extLst" ]
        | "tableColumn" -> [ "calculatedColumnFormula"; "totalsRowFormula"; "xmlColumnPr"; "extLst" ]
        | "queryTable" -> [ "queryTableRefresh"; "extLst" ]
        | "queryTableRefresh" -> [ "sortState"; "queryTableFields"; "queryTableDeletedFields"; "extLst" ]
        | "bookViews" -> [ "workbookView" ]
        | _ -> []
