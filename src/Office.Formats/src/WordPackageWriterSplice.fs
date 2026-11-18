module WordPackageWriterSplice

open XmlGenerator
open WordPackageWriterModel

/// Generate document.xml from typed editor blocks.
let generateDocumentXmlWithFormatting (context: WriterContext) (paragraphs: WordEditorPackage.EditorBlock array) : string =
    let hyperlinks = hyperlinkMap context paragraphs
    let paraElements =
        paragraphs
        |> Array.map (generateDocumentItem context hyperlinks)
        |> String.concat ""

    let body = element "w:body" [] [paraElements]

    let doc =
        element "w:document"
            [("xmlns:w", Namespaces.word); ("xmlns:r", Namespaces.relationships); ("xmlns:m", Namespaces.math)]
            [body]

    document doc


// ---- the body as it was read ---------------------------------------------------
//
// A document that was opened is not written from the model alone: a child
// element of w:body whose blocks the editor did not change goes back into
// document.xml as the text it was read from, with everything in it the model
// does not hold (revision ids, proofing marks, complex fields, content
// control properties, attributes and elements of later formats). Only the
// blocks that were changed, added or moved next to a change are written from
// the model.

let rec storyParagraphCount (item: WordEditorPackage.EditorBlock) : int =
    match item with
    | WordEditorPackage.TextBox _ -> 0
    | WordEditorPackage.Table table ->
        table.Rows
        |> Array.sumBy (fun row ->
            row.Cells
            |> Array.sumBy (fun cell ->
                cell.Paragraphs.Length
                + (cell.NestedTables |> Array.sumBy (WordEditorPackage.Table >> storyParagraphCount))))
    | WordEditorPackage.Paragraph _ -> 1

// ---- a change of formatting made in the XML a paragraph was read from ---------
//
// The editor says of a paragraph whose text is what it was what it was
// before its formatting was changed (the paragraph as the model writes it,
// Before) and what it is now (After). Neither is written: what differs
// between the properties the model writes for the two - of the paragraph, and
// of the run each character is in - is set in the properties the paragraph's
// own XML has. An element the model writes differently replaces the one of
// that name (attribute by attribute where both are a single tag: an attribute
// the model does not write stays, and the order of the attributes), one it
// writes no more goes, a new one stands where the schema puts it. Where the
// formatting of only a part of a run changed the run is split, each part with
// the run's start tag (its revision ids) and properties, the run's children
// shared out by the characters they stand for. Everything else of the
// paragraph stays the text it was read from: bookmarks, fields, hyperlinks
// and content controls around the runs, comment ranges, proofing marks,
// properties and elements the model does not hold. Where the XML cannot be
// matched with the characters the editor shows (a run that stands for text in
// a way not known here, where the change does not cover all it may stand
// for), or holds a tracked change of the properties themselves, the paragraph
// is written from the model.

/// The children of a properties element as the model wrote it (`<w:pPr>...</w:pPr>`, `<w:pPr/>`), each its text.
let private propertyElementsOf (element: string) : string list =
    match XmlPatch.read element with
    | Some source -> XmlPatch.elements source.Node |> List.map (XmlPatch.bytesOf source.Text)
    | None -> []

/// The properties element `source` as it was read (w:rPr, w:pPr, w:tcPr ...; "": there is none yet) with the
/// change the model made between writing its children `was` and `now` made in it (XmlPatch): what the model
/// does not write - elements, attributes, a tracked change of the properties - stays as it was read. Where
/// the XML states the complex-script twin of bold or italic (w:bCs, w:iCs) and the model does not, the twin
/// changes with it. "" when nothing is left in it; None when `source` is not one element.
let mergedProperties (container: string) (source: string) (was: string list) (now: string list) : string option =
    let wrap (children: string list) = "<" + container + ">" + String.concat "" children + "</" + container + ">"
    match XmlPatch.read (if source = "" then "<" + container + "/>" else source), XmlPatch.read (wrap was), XmlPatch.read (wrap now) with
    | Some read, Some before, Some after ->
        let edits = XmlPatch.diff before after
        let named = edits |> List.map XmlPatch.childOf
        // An attribute the model does not write that would overrule one it sets: the same through the
        // document's theme (w:asciiTheme beside w:ascii, w:themeColor beside w:val), or in other units
        // (w:afterLines beside w:after, w:leftChars beside w:left).
        let overrules (element: string) (written: Map<string, string>) (key: string) =
            let lower = key.ToLowerInvariant()
            not (written.ContainsKey key)
            && (if element = "w:rFonts" then lower.Contains "theme" && written |> Map.exists (fun known _ -> lower = known.ToLowerInvariant() + "theme")
                else lower.Contains "theme" || ModelText.isInOtherUnits key)
        let edits =
            edits
            |> List.collect (fun edit ->
                let edit =
                    match edit with
                    // A property element the model changed says what the model writes now in every attribute the
                    // model writes (also where the XML said it another way than the model did before: a spacing
                    // that was automatic); one the model wrote and writes no more goes; one it never writes stays,
                    // but for those that would overrule what it sets.
                    | XmlPatch.EditChild (name, inner, whole) when inner |> List.exists (function XmlPatch.SetAttribute _ | XmlPatch.RemoveAttribute _ -> true | _ -> false) ->
                        let written = XmlPatch.childAttributes after name
                        let gone = XmlPatch.childAttributes read name |> Map.toList |> List.map fst |> List.filter (overrules name written)
                        let removed = inner |> List.filter (function XmlPatch.RemoveAttribute _ -> true | _ -> false)
                        let children = inner |> List.filter (function XmlPatch.SetAttribute _ | XmlPatch.RemoveAttribute _ -> false | _ -> true)
                        let set = written |> Map.toList |> List.map XmlPatch.SetAttribute
                        XmlPatch.EditChild (name, set @ removed @ (gone |> List.map XmlPatch.RemoveAttribute) @ children, whole)
                    | other -> other
                match XmlPatch.childOf edit with
                | "w:b" | "w:i" as name when XmlPatch.hasChild read (name + "Cs") && not (List.contains (name + "Cs") named) -> [ edit; XmlPatch.ofChild (name + "Cs") edit ]
                | _ -> [ edit ])
        if edits.IsEmpty then Some source
        else
            let merged = XmlPatch.apply XmlPatch.Word.order read edits
            Some (if XmlPatch.isEmpty merged then "" else merged)
    | _ -> None

/// A run of a paragraph's XML and the characters of the paragraph's text it stands for.
/// A paragraph as it was read, for one written from its source: the layout of its runs (XmlPatch.layout -
/// in the wrappers the model does not hold, among elements it does not hold), each run's element, its
/// fields, its properties.
type private SpliceSource = {
    Layout: XmlPatch.Layout
    /// Run start -> its element.
    Runs: System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode>
    /// A field's begin run (where it starts) -> where its end run ends: a field the model holds as one run
    /// is written from there, begin mark to end mark (its code, a form field's data), while it is as read.
    Fields: System.Collections.Generic.Dictionary<int, int>
    /// The paragraph's element, for what is in a field's range.
    Paragraph: XmlDsl.XmlNode
    /// Its start tag and what stands after it before its properties (or what it holds), as read; and for a copy.
    Head: string
    CopyHead: string
    /// The w:pPr as it was read.
    Properties: string option
}

/// The elements whose content the model writes itself (the runs in them are the model's) - a simple field;
/// the others it writes (links, tracked changes) are kept from the source while the model's are what they
/// were (see splicedRuns):
/// links, tracked changes, simple fields. The marks of ranges are carried (carriedMarksOf);
/// equations and paragraph-level objects are runs of the model.
let private modelledWrapper (local: string) =
    // (A simple field is a wrapper as it was read, as a link is: a field the model writes otherwise is written outside it.)
    match local with
    | _ -> false
let private modelledLeaf (local: string) =
    match local with
    | "pPr" | "moveFromRangeStart" | "moveFromRangeEnd" | "moveToRangeStart" | "moveToRangeEnd"
    | "oMath" | "oMathPara" | "AlternateContent" -> true
    | _ -> false

/// `bookmarks`: the ids of the bookmarks the model writes in the paragraph (with its runs): those are left out
/// of the source; the others - a bookmark where no run is there to have it - stay where they stood, as do
/// the ranges of comments (of comments the document still has) and of permissions, which the model does
/// not hold.
/// Whether a shape (a text box, a drawing the model holds as an item of its own) is in an element as read.
let rec holdsShape (node: XmlDsl.XmlNode) : bool =
    match node with
    | XmlDsl.Element (name, _, kids, _) -> name = "mc:AlternateContent" || name = "w:pict" || XmlPatch.prefix node = "wps" || XmlPatch.prefix node = "wpg" || List.exists holdsShape kids
    | _ -> false

/// Whether a run holds a shape the model holds as an item of its own (a text box, a drawing shape or group:
/// DrawingML wps / wpg, or VML with a text box) - not a picture, which is a run of the model's.
/// (A drawing canvas or a VML group the reader holds as a group is one too: see DocxReader.Docx.isModelledGroup.)
let holdsShapeItem (run: XmlDsl.XmlNode) : bool =
    not (XmlPatch.descendantsWhere (fun node -> XmlPatch.prefix node = "wps" || XmlPatch.prefix node = "wpg" || XmlPatch.name node = "v:textbox") run).IsEmpty
    || (DocxReader.Docx.modelledShapeRoots run |> List.exists DocxReader.Docx.isModelledGroup)

let private spliceSourceOf (context: WriterContext) (text: string) (start: int) (finish: int) (bookmarks: Set<string>) (keepBookmarks: bool) (keepBreaks: string option) (breaksKept: unit -> unit) : SpliceSource option =
    match XmlDsl.tryScanAt text start finish (System.Collections.Generic.HashSet<string>()) with
    | Some (XmlDsl.Element (_, _, _, span) as paragraph) when XmlDsl.Span.isKnown span && span.StartTagEnd < span.End ->
        let runs = System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode>()
        let order = ResizeArray<int>()
        let properties = XmlPatch.child "pPr" paragraph
        let from = match properties with Some (XmlDsl.Element (_, _, _, propsSpan)) -> propsSpan.End | _ -> span.StartTagEnd
        let take (run: XmlDsl.XmlNode) =
            match run with
            | XmlDsl.Element (_, _, _, runSpan) ->
                runs.[runSpan.Start] <- run
                order.Add runSpan.Start
            | _ -> ()
        let skip (node: XmlDsl.XmlNode) =
            match node with
            | XmlDsl.Element (name, _, _, _) ->
                match XmlDsl.Ns.stripNs name with
                // (The model's bookmarks: written by it - or, in the whole of the paragraph, the source's where they stand.
                // One the model does not hold stays: not every bookmark of a file is one the model has.)
                | "bookmarkStart" | "bookmarkEnd" -> not keepBookmarks && (XmlPatch.attribute node "w:id" |> Option.exists bookmarks.Contains)
                | "commentRangeStart" | "commentRangeEnd" -> XmlPatch.attribute node "w:id" |> Option.exists (fun id -> Set.contains id context.DeletedCommentIds)
                | local -> modelledLeaf local
            | _ -> false
        // (The runs that hold the marks of a field that begins or ends in another paragraph - a table of contents
        // - are the paragraph's as they were: the model does not hold that field.)
        let outerFieldRuns =
            let kept = System.Collections.Generic.HashSet<int>()
            let mutable opened : (int list) list = []   // per open field: the runs of its marks, last first
            for run in XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "w:r") paragraph do
                let start = (XmlPatch.span run).Start
                let marks = XmlPatch.descendants "fldChar" run |> List.choose (fun mark -> XmlPatch.attribute mark "w:fldCharType")
                let isCode = not (XmlPatch.descendants "instrText" run).IsEmpty
                for mark in marks do
                    match mark, opened with
                    | "begin", _ -> opened <- [ start ] :: opened
                    | "end", _ :: rest -> opened <- rest
                    | "end", [] -> kept.Add start |> ignore
                    | _, current :: rest -> opened <- (start :: current) :: rest
                    | _ -> ()
                if isCode && marks.IsEmpty then
                    match opened with
                    | current :: rest -> opened <- (start :: current) :: rest
                    | [] -> ()
            for runs in opened do for start in runs do kept.Add start |> ignore
            kept
        // (A run that holds a shape anchored in the paragraph is the paragraph's as it was too: the model holds the
        // shape as an item of its own, and a block whose shapes changed is not written from its source.)
        // (Nor are the runs in such a shape - a text box's - the paragraph's.)
        let shapeRuns =
            XmlPatch.elements paragraph
            |> List.collect (fun kid -> XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "w:r" && holdsShapeItem node) kid |> List.truncate 1)
            |> List.map XmlPatch.span
        let inShape (node: XmlDsl.XmlNode) =
            let span = XmlPatch.span node
            shapeRuns |> List.exists (fun outer -> span.Start >= outer.Start && span.End <= outer.End)
        // (Nor is a run that holds nothing but its properties: the model has no run of it.)
        let holdsContent (node: XmlDsl.XmlNode) = XmlPatch.elements node |> List.exists (fun kid -> XmlPatch.localName kid <> "rPr")
        // (A run that holds only a page or column break, where the model holds the paragraph's break - after it,
        // wherever it stands -: the break as it stands.)
        let breakOnly (node: XmlDsl.XmlNode) =
            match keepBreaks, XmlPatch.elements node |> List.filter (fun kid -> XmlPatch.localName kid <> "rPr" && XmlPatch.localName kid <> "lastRenderedPageBreak") with
            | Some kind, [ only ] when XmlPatch.name only = "w:br" && XmlPatch.attribute only "w:type" = Some kind -> breaksKept (); true
            | _ -> false
        let isItem (node: XmlDsl.XmlNode) =
            XmlPatch.localName node = "r" && holdsContent node && not (outerFieldRuns.Contains (XmlPatch.span node).Start) && not (inShape node) && not (breakOnly node)
        XmlPatch.layout isItem modelledWrapper skip take text paragraph from
        |> Option.map (fun laid ->
            // The fields: each begin mark with the end mark that closes it, through the runs in order.
            let fields = System.Collections.Generic.Dictionary<int, int>()
            let mutable opened : int list = []
            for at in order do
                let run = runs.[at]
                for mark in XmlPatch.descendants "fldChar" run do
                    match XmlPatch.attribute mark "w:fldCharType" with
                    | Some "begin" -> opened <- at :: opened
                    | Some "end" ->
                        match opened with
                        | opener :: rest ->
                            fields.[opener] <- (match run with XmlDsl.Element (_, _, _, runSpan) -> runSpan.End | _ -> at)
                            opened <- rest
                        | [] -> ()
                    | _ -> ()
            let headEnd = match properties with Some (XmlDsl.Element (_, _, _, propsSpan)) -> propsSpan.Start | _ -> span.StartTagEnd
            let between = XmlPatch.slice text span.StartTagEnd headEnd
            { Layout = laid; Runs = runs; Fields = fields; Paragraph = paragraph
              Properties = properties |> Option.map (XmlPatch.bytesOf text)
              Head = XmlPatch.slice text span.Start headEnd
              CopyHead = XmlPatch.startTag text paragraph [ XmlPatch.RemoveAttribute "w14:paraId"; XmlPatch.RemoveAttribute "w14:textId" ] + between })
    | _ -> None

/// The runs of a paragraph written from its source (see above): `pieces` are the model's runs
/// with their XML (the run itself, and the link or tracked change around it).

let private sameRunExceptText (ignoreMark: bool) (left: WordEditorPackage.EditorRun) (right: WordEditorPackage.EditorRun) =
    (match left.Content, right.Content with
     | WordEditorPackage.Text _, WordEditorPackage.Text _ -> true
     | WordEditorPackage.Equation (_, xml), WordEditorPackage.Equation (_, otherXml) -> xml = otherXml
     | a, b -> a = b)
    && left.Formatting = right.Formatting
    && left.Hyperlink = right.Hyperlink
    && left.InactiveHyperlink = right.InactiveHyperlink
    && left.CommentId = right.CommentId
    && left.CommentText = right.CommentText
    && left.CommentAuthor = right.CommentAuthor
    && left.CommentDate = right.CommentDate
    && left.NoteId = right.NoteId
    && left.NoteType = right.NoteType
    && left.NoteText = right.NoteText
    && left.BookmarkStarts = right.BookmarkStarts
    && left.BookmarkEnds = right.BookmarkEnds
    && left.FieldCode = right.FieldCode
    && left.FieldResult = right.FieldResult
    && left.Revision = right.Revision
    && (ignoreMark || left.RenderedPageBreakBefore = right.RenderedPageBreakBefore)
    && left.Checkbox = right.Checkbox
    && left.ContentControls = right.ContentControls
    && left.SourceRun = right.SourceRun
    && left.NoProps = right.NoProps

let private splicedRuns (context: WriterContext) (source: SpliceSource) (original: WordEditorPackage.EditorParagraph) (joined: WordEditorPackage.EditorParagraph array) (marksBefore: Map<int, string>) (ownBookmarks: Set<string>) (bookmarkEdits: Map<int, XmlPatch.ElementEdit list>) (leading: bool) (trailing: bool) (pieces: (WordEditorPackage.EditorRun * string * (string -> string)) list) (runXml: WordEditorPackage.EditorRun -> string) (wrapOf: WordEditorPackage.EditorRun -> string -> string) : string =
    let text = source.Layout.Text
    let originalRuns = System.Collections.Generic.Dictionary<int, ResizeArray<WordEditorPackage.EditorRun>>()
    for run in Array.append original.Runs (joined |> Array.collect (fun para -> para.Runs)) do
        match run.SourceRun with
        | Some at -> (match originalRuns.TryGetValue at with | true, list -> list.Add run | _ -> originalRuns.[at] <- ResizeArray [ run ])
        | None -> ()
    let currentCount = System.Collections.Generic.Dictionary<int, int>()
    for run, _, _ in pieces do
        match run.SourceRun with
        | Some at -> currentCount.[at] <- (match currentCount.TryGetValue at with | true, n -> n + 1 | _ -> 1)
        | None -> ()
    let asRead = System.Collections.Generic.HashSet<WordEditorPackage.EditorRun>(HashIdentity.Reference)
    for run, _, _ in pieces do
        match run.SourceRun with
        | Some at when currentCount.[at] > 1 ->
            match Portable.Dict.tryFind at originalRuns with
            | ValueSome list when list.Count = 1 && list.[0] = run -> asRead.Add run |> ignore
            | _ -> ()
        | _ -> ()
    // What the model writes of a run: the marks before it (a bookmark's start ...), the run, the marks after it.
    let partsOf (inner: string) =
        match XmlPatch.readFragment inner with
        | Some elements ->
            let isRun (element: XmlPatch.SourceElement) = match element.Node with XmlDsl.Element (name, _, _, _) -> name = "w:r" | _ -> false
            match elements |> List.filter isRun with
            | [ run ] ->
                let index = elements |> List.findIndex isRun
                // (Without the bookmarks the source has where they stand: see spliceSourceOf.)
                let theSource (element: XmlPatch.SourceElement) =
                    (XmlPatch.name element.Node = "w:bookmarkStart" || XmlPatch.name element.Node = "w:bookmarkEnd")
                    && (XmlPatch.attribute element.Node "w:id" |> Option.exists ownBookmarks.Contains)
                let bytes (list: XmlPatch.SourceElement list) = list |> List.filter (theSource >> not) |> List.map XmlPatch.bytes |> String.concat ""
                Some (bytes (List.take index elements), XmlPatch.bytes run, bytes (List.skip (index + 1) elements))
            | _ -> None
        | None -> None
    // (Whether what the model writes of a run is marks with it: a simple field alone - the model's form of a field
    // run - is not.)
    let hasMarks (whole: string) =
        match partsOf whole with
        | Some ("", _, "") -> false
        | Some _ -> true
        | None -> (match XmlPatch.read whole with Some field -> XmlPatch.name field.Node <> "w:fldSimple" | None -> true)
    // (A run that was read whose text and formatting are what they were - however the model holds it, as one
    // run or several (a tab, text, a break), merged where they are formatted alike: its bytes, for the first of
    // them, and nothing for the others.)
    let alike (ignoreMark: bool) (runs: WordEditorPackage.EditorRun seq) =
        runs
        |> Seq.fold (fun (merged: (WordEditorPackage.EditorRun * string) list) run ->
            let text = runText run
            match merged with
            | (previous, previousText) :: rest when sameRunExceptText ignoreMark previous run ->
                (previous, previousText + text) :: rest
            | _ -> (run, text) :: merged) []
    // (Whether a run that was read is in a simple field.)
    let simpleFields = XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "w:fldSimple") source.Paragraph |> List.map XmlPatch.span
    let inSimpleField (at: int) = simpleFields |> List.exists (fun span -> at > span.Start && at < span.End)
    let wholeAsRead = System.Collections.Generic.HashSet<int>()
    for KeyValue (at, list) in originalRuns do
        let now = pieces |> List.filter (fun (run, _, _) -> run.SourceRun = Some at) |> List.map (fun (run, _, _) -> run)
        // (With no marks of the model's around them: a bookmark on one of them is written with it.)
        let marked = pieces |> List.exists (fun (run, whole, _) -> run.SourceRun = Some at && hasMarks whole)
        // (Not the begin run of a field of several runs: that is written from its begin mark to its end mark. A run of
        // a field is as read only as the one run of a simple field, all of the field - none of its result in parts.)
        let fieldAsRead =
            now |> List.forall (fun run ->
                run.FieldCode.IsNone || (inSimpleField at && not (context.SpliceFieldParts.ContainsKey run)))
        if not now.IsEmpty && source.Runs.ContainsKey at && not (source.Fields.ContainsKey at) && fieldAsRead && not marked && alike false now = alike false list then wholeAsRead.Add at |> ignore
    // (A run that was read whose formatting is what it was and whose text changed - however the model holds it, the
    // mark of a page break as it was rendered on its first character, say -: its text changed in the w:t the
    // change is in, the rest of it as it was; written for the first of the model's runs of it.)
    let textOfWhole = System.Collections.Generic.Dictionary<int, XmlPatch.ElementEdit list>()
    let alikeIgnoringMark (runs: WordEditorPackage.EditorRun seq) = alike true runs
    for KeyValue (at, list) in originalRuns do
        if not (wholeAsRead.Contains at) && source.Runs.ContainsKey at && not (source.Fields.ContainsKey at) then
            let indexes = pieces |> List.indexed |> List.filter (fun (_, (run, _, _)) -> run.SourceRun = Some at) |> List.map fst
            let now = indexes |> List.map (fun index -> let runObj, _, _ = pieces.[index] in runObj)
            let marked = indexes |> List.exists (fun index -> let _, whole, _ = pieces.[index] in hasMarks whole)
            let together = not indexes.IsEmpty && List.last indexes - List.head indexes = indexes.Length - 1
            // (Not of a field: the model writes a field as a whole.)
            let ofField = now |> List.exists (fun run -> run.FieldCode.IsSome)
            let plain (value: string) = ModelText.textPieces value |> List.forall (function ModelText.Text _ -> true | _ -> false)
            match alikeIgnoringMark list, alikeIgnoringMark now with
            | [ (formatWas, was) ], [ (formatNow, nowText) ] when formatWas = formatNow && together && not marked && not ofField && was <> nowText && plain was && plain nowText ->
                let run = source.Runs.[at]
                let content = XmlPatch.elements run |> List.filter (fun kid -> XmlPatch.localName kid <> "rPr")
                if not content.IsEmpty && content |> List.forall (fun kid -> XmlPatch.name kid = "w:t" || XmlPatch.name kid = "w:lastRenderedPageBreak")
                   && content |> List.exists (fun kid -> XmlPatch.name kid = "w:t") then
                    let texts = content |> List.filter (fun kid -> XmlPatch.name kid = "w:t") |> List.map XmlPatch.textOf
                    match ModelText.changeInPieces texts was nowText with
                    | Some (index, value) -> textOfWhole.[at] <- [ XmlPatch.EditNth ("w:t", index, [ XmlPatch.SetText value ]) ]
                    | None -> ()
            | _ -> ()
    let firstOfWhole = System.Collections.Generic.HashSet<int>()
    let placedAt = System.Collections.Generic.HashSet<int>()
    let sourceOfRun (run: WordEditorPackage.EditorRun) =
        match run.SourceRun |> Option.filter (fun at -> source.Runs.ContainsKey at) with
        | Some at when currentCount.[at] > 1 && asRead.Count > 0 && (pieces |> List.exists (fun (other, _, _) -> asRead.Contains other && other.SourceRun = Some at)) ->
            if asRead.Contains run then Some (at, true) else None
        | Some at -> Some (at, false)
        | None -> None
    // A new run made like a run that was read (text typed next to it, with its style): that run's start tag and
    // properties - with what the model writes differently of them changed - and the model's content.
    let copyOf (run: WordEditorPackage.EditorRun) (inner: string) : string option =
        match run.SourceRun |> Option.filter (fun at -> source.Runs.ContainsKey at) with
        | Some at ->
            match originalRuns.TryGetValue at, XmlPatch.read inner with
            // (Of what the model writes as a run: a field it writes - a simple field around its run - is its own.)
            | (true, list), Some now when list.Count = 1 && XmlPatch.name now.Node = "w:r" && (match run.Content with WordEditorPackage.Picture _ | WordEditorPackage.Opaque _ -> false | _ -> true) ->
                let sourceRun = source.Runs.[at]
                let props = XmlPatch.child "rPr" sourceRun |> Option.map (XmlPatch.bytesOf text) |> Option.defaultValue ""
                match mergedProperties "w:rPr" props (modelRunProperties list.[0]) (modelRunProperties run), sourceRun with
                | Some merged, XmlDsl.Element (_, _, _, runSpan) when runSpan.StartTagEnd < runSpan.End ->
                    // (Without the ids of the editing session the run was made in: the text is new.)
                    let rsids = XmlPatch.attributeNames sourceRun |> List.filter ModelText.isRevisionId |> List.map XmlPatch.RemoveAttribute
                    Some (XmlPatch.startTag text sourceRun rsids + merged + XmlPatch.childrenBytes now (fun name -> name <> "rPr") + "</w:r>")
                | _ -> None
            | _ -> None
        | None -> None
    let covered = ResizeArray<int * int>()
    // What the model writes around a run (its link, its tracked change), as one text: a run whose wrapping is what
    // it was is written in the wrappers of the source (XmlPatch.layout), one whose wrapping changed outside them,
    // in the model's; one that is new in the wrappers of the run before it, when its wrapping is that one's.
    let wrapping (run: WordEditorPackage.EditorRun) = wrapOf run "\u0002"
    let mutable previousInSource = false
    let mutable previousWrapping = ""
    // (The runs of a field's result: after the run with its separate mark, to its end mark.)
    let fieldResultRunsOf (at: int) (finish: int) =
        let fieldRuns = source.Runs |> Seq.filter (fun (KeyValue (start, _)) -> start >= at && start < finish) |> Seq.sortBy (fun (KeyValue (start, _)) -> start) |> Seq.map (fun (KeyValue (_, node)) -> node) |> List.ofSeq
        let marks (node: XmlDsl.XmlNode) = XmlPatch.descendants "fldChar" node |> List.choose (fun mark -> XmlPatch.attribute mark "w:fldCharType")
        match fieldRuns |> List.tryFindIndex (fun node -> marks node |> List.contains "separate") with
        | Some index -> fieldRuns |> List.skip (index + 1) |> List.filter (fun node -> not (marks node |> List.contains "end"))
        | None -> []
    let marksWritten = System.Collections.Generic.HashSet<int>()
    let written =
        [| for runObj, whole, wrap in pieces do
           match runObj.SourceRun |> Option.bind (fun at -> Map.tryFind at marksBefore |> Option.map (fun marks -> at, marks)) with
           | Some (at, marks) when marksWritten.Add at -> yield XmlPatch.Outside, marks
           | _ -> ()
           match runObj.SourceRun with
           | Some at when wholeAsRead.Contains at ->
               if firstOfWhole.Add at then
                   previousInSource <- true
                   previousWrapping <- wrapping runObj
                   yield XmlPatch.AtItem at, XmlPatch.bytesOf text source.Runs.[at]
           | Some at when textOfWhole.ContainsKey at ->
               if firstOfWhole.Add at then
                   placedAt.Add at |> ignore
                   previousInSource <- true
                   previousWrapping <- wrapping runObj
                   yield XmlPatch.AtItem at, XmlPatch.apply XmlPatch.Word.order { XmlPatch.SourceElement.Text = text; XmlPatch.SourceElement.Node = source.Runs.[at] } textOfWhole.[at]
           | _ ->
             let now = wrapping runObj
             // (The marks the model writes around the run go around what is written of it.)
             let before, inner, after = match partsOf whole with Some parts -> parts | None -> "", whole, ""
             let wrap (xml: string) = if xml = whole then wrap xml else wrap (before + xml + after)
             match sourceOfRun runObj with
             | Some (at, alone) ->
                 let run = source.Runs.[at]
                 // The model's run then, when it was one run of the source and is one now (or is the one of several) -
                 // or when the runs it was are of one format and hold the text it holds (a tab, text, a break): its
                 // formatting alone is what can have changed.
                 let sameText =
                     match Portable.Dict.tryFind at originalRuns with
                     | ValueSome list when currentCount.[at] = 1 ->
                         (match alike false list, alike false [ runObj ] with
                          | [ (_, before) ], [ (_, after) ] -> before = after
                          | _ -> false)
                     | _ -> false
                 // (A run split - a part of it formatted -: each part is the run with its own change, and text.)
                 let was =
                     match Portable.Dict.tryFind at originalRuns with
                     | ValueSome list when list.Count = 1 -> Some list.[0]
                     | ValueSome list when sameText -> Some list.[0]
                     // (A run the model held as several of one format - a tab, text, a break - split: each part the
                     // run with the change of format made in it, and the model's content.)
                     | ValueSome list when (alike false list).Length = 1 -> Some list.[0]
                     | _ -> None
                 // (Whether the run is one of several the model held it as, and is not all of them now.)
                 let ofSeveral = not sameText && (match originalRuns.TryGetValue at with | true, list -> list.Count > 1 | _ -> false)
                 ignore alone
                 // (A run of a paragraph joined to this one is not an item of its layout: it is written in the model's wrapping.)
                 let wrappedAsBefore =
                     XmlPatch.isItemOf source.Layout at &&
                     match Portable.Dict.tryFind at originalRuns with
                     | ValueSome list when list.Count > 0 -> list |> Seq.forall (fun model -> wrapping model = now)
                     | _ -> false
                 // (The parts of a run that was split after the first are in its place after it.)
                 let place, wrap =
                     if wrappedAsBefore then (if placedAt.Add at then XmlPatch.AtItem at else XmlPatch.InItem at), (fun xml -> if xml = whole then xml else before + xml + after)
                     else XmlPatch.Outside, wrap
                 previousInSource <- wrappedAsBefore
                 previousWrapping <- now
                 match was with
                 // A field as it was read: its runs, begin mark to end mark, as they are in the source (the bookmarks
                 // in it are the model's, written around it); what stood among them is in them.
                 | Some model when source.Fields.ContainsKey at && model = runObj ->
                     let finish = source.Fields.[at]
                     covered.Add ((at, finish))
                     yield place, wrap (XmlPatch.sliceWithout text source.Paragraph at finish (fun name -> name = "bookmarkStart" || name = "bookmarkEnd"))
                 // A field as it was read but for its formatting: its runs, the result's with the change made in their
                 // properties (see mergedProperties); its code and marks as they were.
                 | Some model when source.Fields.ContainsKey at && runText model = runText runObj && model.FieldCode = runObj.FieldCode ->
                     let finish = source.Fields.[at]
                     covered.Add ((at, finish))
                     let was, now = modelRunProperties model, modelRunProperties runObj
                     // (The runs after the one with the field's separate mark, to its end mark.)
                     let fieldRuns = source.Runs |> Seq.filter (fun (KeyValue (start, _)) -> start >= at && start < finish) |> Seq.sortBy (fun (KeyValue (start, _)) -> start) |> Seq.map (fun (KeyValue (_, node)) -> node) |> List.ofSeq
                     let marks (node: XmlDsl.XmlNode) = XmlPatch.descendants "fldChar" node |> List.choose (fun mark -> XmlPatch.attribute mark "w:fldCharType")
                     let separateAt = fieldRuns |> List.tryFindIndex (fun node -> marks node |> List.contains "separate")
                     let resultRuns =
                         match separateAt with
                         | Some index -> fieldRuns |> List.skip (index + 1) |> List.filter (fun node -> not (marks node |> List.contains "end"))
                         | None -> []
                     let replacements =
                         resultRuns |> List.choose (fun node ->
                             let props = XmlPatch.child "rPr" node |> Option.map (XmlPatch.bytesOf text) |> Option.defaultValue ""
                             mergedProperties "w:rPr" props was now
                             |> Option.map (fun merged ->
                                 let span = match node with XmlDsl.Element (_, _, _, sp) -> sp | _ -> XmlDsl.Span.none
                                 let edit = if merged = "" then XmlPatch.RemoveChild "w:rPr" else XmlPatch.SetChild ("w:rPr", merged)
                                 span.Start, span.End, XmlPatch.apply XmlPatch.Word.order { XmlPatch.SourceElement.Text = text; XmlPatch.SourceElement.Node = node } [ edit ]))
                     if replacements.Length = resultRuns.Length && not resultRuns.IsEmpty then
                         let bytes = XmlPatch.sliceWith text at finish replacements
                         // (The bookmarks in it are the model's, written around it.)
                         let bytes = match XmlPatch.readFragment bytes with
                                     | Some elements -> elements |> List.filter (fun element -> match element.Node with XmlDsl.Element (name, _, _, _) -> name <> "w:bookmarkStart" && name <> "w:bookmarkEnd" | _ -> true) |> List.map XmlPatch.bytes |> String.concat ""
                                     | None -> bytes
                         yield place, wrap bytes
                     else
                         Portable.Stack.dropLast covered
                         yield place, wrap inner
                 // A field whose result was typed into (its code as it was): its runs as read, its one result run
                 // written as the model's parts of the result - each that run with the part's formatting and text.
                 | Some model when source.Fields.ContainsKey at && model.FieldCode = runObj.FieldCode
                                   && (fieldResultRunsOf at source.Fields.[at]).Length = 1 ->
                     let finish = source.Fields.[at]
                     let resultRun = (fieldResultRunsOf at finish).Head
                     let props = XmlPatch.child "rPr" resultRun |> Option.map (XmlPatch.bytesOf text) |> Option.defaultValue ""
                     let parts = runObj :: (match context.SpliceFieldParts.TryGetValue runObj with | true, parts -> parts | _ -> [])
                     let written =
                         parts |> List.map (fun part ->
                             mergedProperties "w:rPr" props (modelRunProperties model) (modelRunProperties part)
                             |> Option.map (fun merged ->
                                 XmlPatch.startTag text resultRun [] + merged
                                 + String.concat "" (runTextElements (isDeletedRun part) (runText part)) + "</w:r>"))
                     if written |> List.forall Option.isSome then
                         covered.Add ((at, finish))
                         let span = XmlPatch.span resultRun
                         let bytes = XmlPatch.sliceWith text at finish [ span.Start, span.End, written |> List.choose id |> String.concat "" ]
                         let bytes = match XmlPatch.readFragment bytes with
                                     | Some elements -> elements |> List.filter (fun element -> match element.Node with XmlDsl.Element (name, _, _, _) -> name <> "w:bookmarkStart" && name <> "w:bookmarkEnd" | _ -> true) |> List.map XmlPatch.bytes |> String.concat ""
                                     | None -> bytes
                         yield place, wrap bytes
                     else yield place, wrap inner
                 // (Written by the model: outside the wrappers it was in - a simple field the model writes is not in the one it was.)
                 | Some _ when runObj.FieldCode.IsSome -> yield XmlPatch.Outside, wrap inner
                 // As it was read in every respect: its bytes.
                 | Some model when model = runObj && not ofSeveral -> yield place, wrap (XmlPatch.bytesOf text run)
                 // Changed in its formatting, its text, or both: the run as it was read with what the model changed
                 // made in it (XmlPatch): its properties merged (see mergedProperties), its text set in its w:t.
                 // Where its children are not of the same names in the same order the model's are written in it.
                 | Some model when (match runObj.Content with WordEditorPackage.Picture _ | WordEditorPackage.Opaque _ -> false | _ -> true) ->
                     let props = XmlPatch.child "rPr" run
                     let modelRun = match partsOf (runXml model) with Some (_, run, _) -> run | None -> runXml model
                     match mergedProperties "w:rPr" (props |> Option.map (XmlPatch.bytesOf text) |> Option.defaultValue "") (modelRunProperties model) (modelRunProperties runObj), XmlPatch.read modelRun, XmlPatch.read inner with
                     | Some merged, Some before, Some after ->
                         let source = { XmlPatch.SourceElement.Text = text; XmlPatch.SourceElement.Node = run }
                         let propsEdit = if merged = "" then XmlPatch.RemoveChild "w:rPr" else XmlPatch.SetChild ("w:rPr", merged)
                         let content =
                             if sameText then []
                             else XmlPatch.diff before after |> List.filter (fun edit -> let child = XmlPatch.childOf edit in child <> "" && child <> "w:rPr")
                         // (Content of its own - not the run's as it was - where the run's text changed in a way edits do not
                         // say, or the run is one of several the model held it as: its part of the content is the model's.)
                         // (Or where what the run holds is not of the same kinds in the same order: text typed before a tab -
                         // an edit puts a child in at its schema place, and the order of a run's content is its text.)
                         let contentNames (element: XmlPatch.SourceElement) = XmlPatch.elements element.Node |> List.map XmlPatch.name |> List.filter (fun name -> name <> "w:rPr")
                         let reordered = not sameText && contentNames before <> contentNames after
                         // (The run as read holds its text otherwise than the model writes it - in two w:t, say: the edits of
                         // the model's run are not edits of it. Its text written in the w:t the change is in, where the
                         // change is in one; else the model's content.)
                         let sourceShape = contentNames source
                         let otherShape = not sameText && sourceShape <> contentNames before
                         let textOnly (names: string list) = not names.IsEmpty && names |> List.forall (fun name -> name = "w:t")
                         // (The model's text of the run as read: all of the runs it held it as, when they are of one format.)
                         let wasText = match originalRuns.TryGetValue at with | true, list -> (match alike false list with [ (_, whole) ] -> Some whole | _ -> None) | _ -> None
                         let textInPlace =
                             let nowText = runText runObj
                             let plain (value: string) = ModelText.textPieces value |> List.forall (function ModelText.Text _ -> true | _ -> false)
                             match wasText with
                             | Some was when not sameText && currentCount.[at] = 1 && textOnly sourceShape && sourceShape.Length > 1 && plain was && plain nowText ->
                                 let pieces = XmlPatch.elements source.Node |> List.filter (fun kid -> XmlPatch.name kid = "w:t") |> List.map XmlPatch.textOf
                                 ModelText.changeInPieces pieces was nowText
                                 |> Option.map (fun (index, value) -> [ XmlPatch.EditNth ("w:t", index, [ XmlPatch.SetText value ]) ])
                             | _ -> None
                         // (Text typed before a run whose content does not begin with its text - a tab, a break -, in its
                         // formatting: a run of its own before it, made like it; the run as it was.)
                         let typedBefore =
                             match wasText with
                             | Some was when currentCount.[at] = 1 && merged = (props |> Option.map (XmlPatch.bytesOf text) |> Option.defaultValue "")
                                             && (match sourceShape with first :: _ -> first <> "w:t" | [] -> false) ->
                                 ModelText.insertedBefore was (runText runObj)
                                 |> Option.filter (fun typed -> ModelText.textPieces typed |> List.forall (function ModelText.Text _ -> true | _ -> false))
                             | _ -> None
                         match textInPlace, typedBefore with
                         | _, Some typed ->
                             let rsids = XmlPatch.attributeNames run |> List.filter ModelText.isRevisionId |> List.map XmlPatch.RemoveAttribute
                             let copy = XmlPatch.startTag text run rsids + merged + String.concat "" (runTextElements false typed) + "</w:r>"
                             yield place, wrap (copy + XmlPatch.bytesOf text run)
                         | Some textEdits, None -> yield place, wrap (XmlPatch.apply XmlPatch.Word.order source (propsEdit :: textEdits))
                         | None, None ->
                         if ofSeveral || reordered || otherShape || content |> List.exists (function XmlPatch.SetChildren _ -> true | _ -> false) then
                             let runSpan = match run with XmlDsl.Element (_, _, _, sp) -> sp | _ -> XmlDsl.Span.none
                             let modelContent = XmlPatch.childrenBytes after (fun name -> name <> "rPr")
                             if runSpan.StartTagEnd < runSpan.End then
                                 yield place, wrap (XmlPatch.slice text runSpan.Start runSpan.StartTagEnd + merged + modelContent + XmlPatch.slice text runSpan.EndTagStart runSpan.End)
                             else yield place, wrap inner
                         else yield place, wrap (XmlPatch.apply XmlPatch.Word.order source (propsEdit :: content))
                     | _ -> yield place, wrap inner
                 | _ -> yield place, wrap inner
             | None ->
                 let run = copyOf runObj inner |> Option.defaultValue inner
                 if previousInSource && now = previousWrapping then yield XmlPatch.WithPrevious, before + run + after
                 else
                     previousInSource <- false
                     yield XmlPatch.Outside, wrap run |]
    XmlPatch.writeLayoutEditing { KeptBefore = leading; KeptAfter = trailing; Covered = List.ofSeq covered } XmlPatch.Word.order bookmarkEdits source.Layout written


/// The bookmarks of a paragraph written from its source that the editor renamed or removed
/// (renameBookmark / removeBookmark): the change made in the source's w:bookmarkStart and w:bookmarkEnd
/// where they stand - its w:name set, or the two taken out (an end in another paragraph is taken out there) -
/// as edits of the kept elements of the layout, by where each starts. The source's bookmarks stand where they
/// are in the whole of a paragraph (see spliceSourceOf), so this is of the whole of one only: the bookmarks
/// its runs had when it was read (SpliceOriginal) against those they have now. A bookmark the model never had
/// (one where no run is there to have it) is not one of them: it stays.
let private sourceBookmarkEdits (para: WordEditorPackage.EditorParagraph) (source: SpliceSource) (whole: bool) : Map<int, XmlPatch.ElementEdit list> =
    match para.SpliceOriginal with
    | None -> Map.empty
    | Some original when not whole -> Map.empty
    | Some original ->
        let startsIn (runs: WordEditorPackage.EditorRun array) : ModelText.BookmarkStart list =
            runs |> List.ofArray |> List.collect (fun run ->
                run.BookmarkStarts @ (run.BookmarkEnds |> List.choose (function ModelText.Opens start -> Some start | ModelText.Ends _ -> None)))
        let endsIn (runs: WordEditorPackage.EditorRun array) : Set<string> =
            runs |> List.ofArray |> List.collect (fun run ->
                (run.BookmarkStarts |> List.filter (fun start -> start.EndsHere) |> List.map (fun start -> start.Id))
                @ (run.BookmarkEnds |> List.choose (function ModelText.Ends id -> Some id | ModelText.Opens start when start.EndsHere -> Some start.Id | ModelText.Opens _ -> None)))
            |> Set.ofList
        let namesOf (starts: ModelText.BookmarkStart list) = starts |> List.map (fun start -> start.Id, start.Name) |> Map.ofList
        let wasStarts, nowStarts = namesOf (startsIn original.Runs), namesOf (startsIn para.Runs)
        let wasEnds, nowEnds = endsIn original.Runs, endsIn para.Runs
        XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "w:bookmarkStart" || XmlPatch.name node = "w:bookmarkEnd") source.Paragraph
        |> List.choose (fun node ->
            let at = (XmlPatch.span node).Start
            match XmlPatch.attribute node "w:id" with
            | Some id when XmlPatch.name node = "w:bookmarkStart" ->
                match Map.tryFind id wasStarts, Map.tryFind id nowStarts with
                | Some _, None -> Some (at, [ XmlPatch.Removed ])
                | Some was, Some now when now <> was && now <> "" && XmlPatch.attribute node "w:name" <> Some now -> Some (at, [ XmlPatch.SetAttribute ("w:name", now) ])
                | _ -> None
            | Some id when wasEnds.Contains id && not (nowEnds.Contains id) -> Some (at, [ XmlPatch.Removed ])
            | _ -> None)
        |> Map.ofList

let configure (context: WriterContext) =
    context.ParagraphSplicer <- Some
      { new ParagraphSplicer with
        member _.Splice para pieces runXml wrapOf =
            match para.SpliceFrom, para.SpliceOriginal with
            | Some (start, finish), Some original when context.SpliceSourceText <> "" ->
                // (The bookmarks the model writes: in what it writes of the runs.)
                let bookmarks =
                    [ for _, whole, _ in pieces do
                        match XmlPatch.readFragment whole with
                        | Some elements ->
                            for element in elements do
                                for name in [ "bookmarkStart"; "bookmarkEnd" ] do
                                    for mark in XmlPatch.descendants name element.Node do
                                        yield! Option.toList (XmlPatch.attribute mark "w:id")
                        | None -> () ] |> Set.ofList
                // (Only in the whole of the paragraph: a part of one, or one with others joined to it, has its break where
                // the model says - the source's may stand in another part.)
                let whole = para.SpliceLeading && para.SpliceTrailing && para.SpliceJoined.Length = 0
                let keepBreaks =
                    if whole && para.Formatting.PageBreakAfter then Some (if para.Formatting.ColumnBreakAfter then "column" else "page")
                    elif whole && para.Formatting.PageBreakAtStart then Some "page"
                    else None
                let mutable breakKept = false
                spliceSourceOf context context.SpliceSourceText start finish bookmarks whole keepBreaks (fun () -> breakKept <- true)
                |> Option.map (fun source ->
                    let joined = para.SpliceJoined
                    // (And the starts of the ranges that stood between it and the paragraph before it: before its first run.)
                    let mutable marksBefore = Map.empty
                    let mutable previousEnd = finish
                    for other in joined do
                        match context.ParagraphSource |> Option.bind (fun provider -> provider.SpanOf other) with
                        | Some (start, finish) ->
                            match XmlDsl.tryScanAt context.SpliceSourceText start finish (System.Collections.Generic.HashSet<string>()) with
                            | Some node ->
                                let runs = XmlPatch.descendantsWhere (fun inner -> XmlPatch.name inner = "w:r") node
                                for run in runs do source.Runs.[(XmlPatch.span run).Start] <- run
                                if previousEnd >= 0 && previousEnd <= start && not runs.IsEmpty then
                                    match XmlPatch.readFragmentRoot (XmlPatch.slice context.SpliceSourceText previousEnd start) with
                                    | Some between ->
                                        // (All of them, in their order: a bookmark that began and ended there is where the text begins.)
                                        let starts = rangeMarksIn carriedRangeKinds between.Node |> List.filter (fun mark -> mark.Kind <> "commentRange" || Set.contains mark.Id context.LiveCommentIds) |> List.map (fun mark -> XmlPatch.bytesOf between.Text mark.Node)
                                        if not starts.IsEmpty then marksBefore <- Map.add (XmlPatch.span runs.Head).Start (String.concat "" starts) marksBefore
                                    | None -> ()
                            | None -> ()
                            previousEnd <- finish
                        | None -> previousEnd <- -1
                    // (The bookmarks the model writes that the source has, in the whole of the paragraph: the source's.)
                    let ownBookmarks =
                        if not whole then Set.empty
                        else
                            XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "w:bookmarkStart" || XmlPatch.name node = "w:bookmarkEnd") source.Paragraph
                            |> List.choose (fun node -> XmlPatch.attribute node "w:id")
                            |> List.filter bookmarks.Contains
                            |> Set.ofList
                    [ splicedRuns context source original joined marksBefore ownBookmarks (sourceBookmarkEdits para source whole) para.SpliceLeading para.SpliceTrailing pieces runXml wrapOf ], source.Properties,
                    (if para.Source.Copy then source.CopyHead else source.Head), breakKept)
            | _ -> None
        member _.MergeProperties source was now =
            mergedProperties "w:pPr" source (propertyElementsOf was) (propertyElementsOf now) }

