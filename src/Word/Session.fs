module Rbt.Word.Session

open Rbt.Word.Document
open Rbt.Word.Commands
module Doc = Rbt.Word.Document


#if FABLE_COMPILER
open Fable.Core
open Fable.Core.JsInterop

[<Emit("({ weak: new WeakMap(), primitive: new Map() })")>]
let private newIndex () : obj = jsNative
[<Emit("""((index, doc) => {
    let node = index;
    for (const key of [doc.Blocks, doc.PageSetup, doc.Styles, doc.Headers, doc.Footers, doc.HeaderParts, doc.FooterParts, doc.Package]) {
        const map = key !== null && (typeof key === 'object' || typeof key === 'function') ? node.weak : node.primitive;
        node = map.get(key); if (node === undefined) return -1;
    }
    return node.revision === undefined ? -1 : node.revision;
})($0, $1)""")>]
let private findRevision (_index: obj) (_doc: Document) : int = jsNative
[<Emit("""((index, doc, revision) => {
    let node = index;
    for (const key of [doc.Blocks, doc.PageSetup, doc.Styles, doc.Headers, doc.Footers, doc.HeaderParts, doc.FooterParts, doc.Package]) {
        const map = key !== null && (typeof key === 'object' || typeof key === 'function') ? node.weak : node.primitive;
        let next = map.get(key);
        if (next === undefined) { next = { weak: new WeakMap(), primitive: new Map() }; map.set(key, next); }
        node = next;
    }
    node.revision = revision;
})($0, $1, $2)""")>]
let private rememberRevision (_index: obj) (_doc: Document) (_revision: int) : unit = jsNative
#else
open System.Runtime.CompilerServices

[<AllowNullLiteral>]
type private Node() =
    let children = ConditionalWeakTable<obj, Node>()
    let mutable nullChild: Node = null
    member val Revision = -1 with get, set
    member _.Child(key: obj, create: bool) =
        if isNull key then
            if isNull nullChild && create then nullChild <- Node()
            nullChild
        else
            let mutable child = Unchecked.defaultof<Node>
            if children.TryGetValue(key, &child) then child
            elif create then
                let child = Node()
                children.Add(key, child)
                child
            else null
let private newIndex () = Node()
let private nodeAt (index: Node) (doc: Document) create =
    let mutable node = index.Child(box doc.Blocks, create)
    if not (isNull node) then node <- node.Child(box doc.PageSetup, create)
    if not (isNull node) then node <- node.Child(box doc.Styles, create)
    if not (isNull node) then node <- node.Child(box doc.Headers, create)
    if not (isNull node) then node <- node.Child(box doc.Footers, create)
    if not (isNull node) then node <- node.Child(box doc.HeaderParts, create)
    if not (isNull node) then node <- node.Child(box doc.FooterParts, create)
    if not (isNull node) then node <- node.Child(box doc.Package, create)
    node
let private findRevision index doc =
    let node = nodeAt index doc false
    if isNull node then -1 else node.Revision
let private rememberRevision index doc revision =
    let node = nodeAt index doc true
    node.Revision <- revision
#endif

let sameContentObjects (left: Document) (right: Document) =
    obj.ReferenceEquals(left, right) ||
    (obj.ReferenceEquals(left.Blocks, right.Blocks) && obj.ReferenceEquals(left.PageSetup, right.PageSetup)
     && obj.ReferenceEquals(left.Styles, right.Styles) && obj.ReferenceEquals(left.Headers, right.Headers)
     && obj.ReferenceEquals(left.Footers, right.Footers) && obj.ReferenceEquals(left.HeaderParts, right.HeaderParts)
     && obj.ReferenceEquals(left.FooterParts, right.FooterParts) && obj.ReferenceEquals(left.Package, right.Package))

type SaveSnapshot = { Epoch: obj; Document: Document; Revision: int; Unchanged: bool }

/// A header or footer that is edited in place (see `enterStory`). While it
/// is, the document the handler works on is the story's own: its blocks are
/// the header's blocks and its selection is in them, so typing, selecting,
/// formatting, tables, pictures and fields work there as they do in the body.
/// The whole document - the body with the story's blocks in their part - is
/// what is painted (dimmed), saved and kept in the undo history.
type StorySession = {
    IsHeader: bool
    PageIndex: int
    /// "default" | "first" | "even".
    Kind: string
    /// The section the header is that of (0-based).
    Section: int
    /// The place of the part in the document's HeaderParts / FooterParts.
    PartIndex: int
    /// The name on the boundary.
    Label: string
    /// The whole document when the story was entered: its body, and the body's selection.
    Main: Doc.Document
    /// The whole document the story's blocks are put into once they are
    /// changed. Main itself for a header that exists; for one that is only
    /// made when something is typed into it (a page that shows none, a title
    /// or even page without a header of its own), Main with the new part and
    /// the section's reference to it. PartIndex is the part's place in this.
    Edited: Doc.Document
    /// The part's blocks when the story was entered, and the story's blocks
    /// then (the part's, with the page's number in the fields that show it).
    PartBlocks: Doc.Block list
    EntryBlocks: Doc.Block list
    /// The parts with the story's blocks in their part, by those blocks: a
    /// state of the story that comes back (undo, redo) is the same document again.
    Folded: ResizeArray<struct (Doc.Block list * Doc.HeaderFooterPart list)>
    /// The story is the text of this text box of the body (see
    /// `enterTextBox`), not a header or footer: its blocks are the box's
    /// paragraphs, and the whole document is the body with them in the box.
    Box: Doc.BlockId option
    /// The body's blocks with the story's blocks in the box, by those blocks (as Folded).
    BoxFolded: ResizeArray<struct (Doc.Block list * Doc.Block list)>
    /// The story is the body of this note or comment ("footnote" | "endnote" |
    /// "comment", id; see `Doc.withStoryBody`), not a header or footer: its
    /// blocks are the body's blocks in the package, and the whole document is
    /// Main with them there (and their text on the references).
    Note: (string * string) option
    /// The whole documents with the story's blocks as the note's body, by those blocks (as Folded).
    NoteFolded: ResizeArray<struct (Doc.Block list * Doc.Document)>
}

/// Character formatting chosen with nothing selected (Ctrl+B at a caret, a
/// font size picked in the toolbar): what is typed next at that caret has it.
/// It is not in the document - nothing has changed there until text is typed.
type PendingFormat = {
    /// The document it was chosen in. It holds for that text and that caret:
    /// an edit or a caret that moves drops it.
    PendingDocument: Doc.Document
    /// The style the typed text gets.
    PendingStyle: TextBuffer.TextStyle
}
/// The session's parts (the headers, or the footers) with `blocks` in the story's part.
let private storyParts (session: StorySession) (blocks: Doc.Block list) : Doc.HeaderFooterPart list =
    // (As entered: the document as it was - showing a header is not a change of the document.)
    if obj.ReferenceEquals(blocks, session.EntryBlocks) then (if session.IsHeader then session.Main.HeaderParts else session.Main.FooterParts)
    else
        let parts = if session.IsHeader then session.Edited.HeaderParts else session.Edited.FooterParts
        let mutable found = Unchecked.defaultof<Doc.HeaderFooterPart list>
        let mutable known = false
        for struct (knownBlocks, knownParts) in session.Folded do
            if not known && obj.ReferenceEquals(knownBlocks, blocks) then
                found <- knownParts
                known <- true
        if known then found
        else
            let next = parts |> List.mapi (fun index part -> if index = session.PartIndex then { part with Blocks = blocks } else part)
            if session.Folded.Count >= 256 then session.Folded.RemoveAt 0
            session.Folded.Add(struct (blocks, next))
            next

/// The indexes of the body's blocks that refer to a note or comment (see
/// Doc.storyReferenceIndexes), for the body a story was entered from: the
/// other blocks are not looked at again at each keystroke in the note.
let mutable private knownReferenceIndexes : (Doc.Document * string * string * int array) option = None
let private storyReferenceIndexesOf (main: Doc.Document) (kind: string) (id: string) : int array =
    match knownReferenceIndexes with
    | Some (known, knownKind, knownId, indexes) when obj.ReferenceEquals(known, main) && knownKind = kind && knownId = id -> indexes
    | _ ->
        let indexes = Doc.storyReferenceIndexes kind id main
        knownReferenceIndexes <- Some (main, kind, id, indexes)
        indexes

/// The whole document a story document stands for: the body, the story's
/// blocks in their part, the story's selection (and whatever else an edit of
/// the story changed - the styles).
let private foldStory (session: StorySession) (story: Doc.Document) : Doc.Document =
    match session.Note with
    | Some (kind, id) ->
        // The body of a note or comment: the document as it was, the story's blocks as that body.
        let main = session.Main
        if obj.ReferenceEquals(story.Blocks, session.EntryBlocks) then
            { story with Blocks = main.Blocks; HeaderParts = main.HeaderParts; FooterParts = main.FooterParts; Package = main.Package }
        else
            let mutable found = Unchecked.defaultof<Doc.Document>
            let mutable known = false
            for struct (knownBlocks, knownWhole) in session.NoteFolded do
                if not known && obj.ReferenceEquals(knownBlocks, story.Blocks) then
                    found <- knownWhole
                    known <- true
            if not known then
                found <- Doc.withStoryBodyAt kind id story.Blocks (Some (storyReferenceIndexesOf main kind id)) main
                if session.NoteFolded.Count >= 256 then session.NoteFolded.RemoveAt 0
                session.NoteFolded.Add(struct (story.Blocks, found))
            { story with Blocks = found.Blocks; HeaderParts = found.HeaderParts; FooterParts = found.FooterParts; Package = found.Package }
    | None ->
    match session.Box with
    | Some boxId when session.PartIndex >= 0 ->
        // A text box of a header or footer: the document as it was, the story's paragraphs in the box of its part.
        let main = session.Main
        if obj.ReferenceEquals(story.Blocks, session.EntryBlocks) then { story with Blocks = main.Blocks; HeaderParts = main.HeaderParts; FooterParts = main.FooterParts }
        else
            let mutable parts = Unchecked.defaultof<Doc.HeaderFooterPart list>
            let mutable known = false
            for struct (knownBlocks, knownParts) in session.Folded do
                if not known && obj.ReferenceEquals(knownBlocks, story.Blocks) then
                    parts <- knownParts
                    known <- true
            if not known then
                let content = Doc.paragraphsOfBlocks story.Blocks
                parts <-
                    (if session.IsHeader then main.HeaderParts else main.FooterParts)
                    |> List.mapi (fun index part ->
                        if index <> session.PartIndex then part
                        else { part with Blocks = part.Blocks |> List.map (function Doc.TextBox box when box.Id = boxId -> Doc.TextBox { box with Content = content } | block -> block) })
                if session.Folded.Count >= 256 then session.Folded.RemoveAt 0
                session.Folded.Add(struct (story.Blocks, parts))
            if session.IsHeader then { story with Blocks = main.Blocks; HeaderParts = parts; FooterParts = main.FooterParts }
            else { story with Blocks = main.Blocks; HeaderParts = main.HeaderParts; FooterParts = parts }
    | Some boxId ->
        // A text box: the body as it was, the story's paragraphs in the box.
        if obj.ReferenceEquals(story.Blocks, session.EntryBlocks) then { story with Blocks = session.Main.Blocks }
        else
            let mutable found = Unchecked.defaultof<Doc.Block list>
            let mutable known = false
            for struct (knownBlocks, knownBody) in session.BoxFolded do
                if not known && obj.ReferenceEquals(knownBlocks, story.Blocks) then
                    found <- knownBody
                    known <- true
            if known then { story with Blocks = found }
            else
                let content = Doc.paragraphsOfBlocks story.Blocks
                let body =
                    session.Main.Blocks
                    |> List.map (function
                        | Doc.TextBox box when box.Id = boxId -> Doc.TextBox { box with Content = content }
                        | block -> block)
                if session.BoxFolded.Count >= 256 then session.BoxFolded.RemoveAt 0
                session.BoxFolded.Add(struct (story.Blocks, body))
                { story with Blocks = body }
    | None ->
    let parts = storyParts session story.Blocks
    // (The body, the plain header paragraphs and the sections' references are those of the
    // document the blocks belong in: a header made by the first edit comes with its reference.)
    let basis = if obj.ReferenceEquals(story.Blocks, session.EntryBlocks) then session.Main else session.Edited
    let pageSetup =
        if obj.ReferenceEquals(basis.PageSetup.HeaderReferences, story.PageSetup.HeaderReferences) && obj.ReferenceEquals(basis.PageSetup.FooterReferences, story.PageSetup.FooterReferences) then story.PageSetup
        else
            let wanted = { story.PageSetup with HeaderReferences = basis.PageSetup.HeaderReferences; FooterReferences = basis.PageSetup.FooterReferences }
            // (The same object when nothing else differs: what is kept by reference stays kept.)
            if wanted = basis.PageSetup then basis.PageSetup else wanted
    if session.IsHeader then { story with Blocks = basis.Blocks; PageSetup = pageSetup; Headers = basis.Headers; HeaderParts = parts; FooterParts = basis.FooterParts }
    else { story with Blocks = basis.Blocks; PageSetup = pageSetup; Footers = basis.Footers; HeaderParts = basis.HeaderParts; FooterParts = parts }

/// The story document of a whole document: the blocks of the story's part.
let private unfoldStory (session: StorySession) (whole: Doc.Document) : Doc.Document =
    match session.Note with
    | Some (kind, id) ->
        let blocks =
            if obj.ReferenceEquals(whole.Package, session.Main.Package) then session.EntryBlocks
            else
                let mutable found = Unchecked.defaultof<Doc.Block list>
                let mutable known = false
                for struct (knownBlocks, knownWhole) in session.NoteFolded do
                    if not known && obj.ReferenceEquals(knownWhole.Package, whole.Package) then
                        found <- knownBlocks
                        known <- true
                if known then found
                else
                    match Doc.storyBody kind id whole with
                    | Some blocks when not (Doc.paragraphsOfBlocks blocks).IsEmpty -> blocks
                    | _ -> session.EntryBlocks
        { whole with Blocks = blocks }
    | None ->
    match session.Box with
    | Some boxId when session.PartIndex >= 0 ->
        let parts = if session.IsHeader then whole.HeaderParts else whole.FooterParts
        let entered = if session.IsHeader then session.Main.HeaderParts else session.Main.FooterParts
        let blocks =
            if obj.ReferenceEquals(parts, entered) then session.EntryBlocks
            else
                let mutable found = Unchecked.defaultof<Doc.Block list>
                let mutable known = false
                for struct (knownBlocks, knownParts) in session.Folded do
                    if not known && obj.ReferenceEquals(knownParts, parts) then
                        found <- knownBlocks
                        known <- true
                if known then found
                else
                    match List.tryItem session.PartIndex parts |> Option.bind (fun part -> part.Blocks |> List.tryPick (function Doc.TextBox box when box.Id = boxId -> Some box | _ -> None)) with
                    | Some box when not box.Content.IsEmpty -> box.Content |> List.map Doc.Paragraph
                    | _ -> session.EntryBlocks
        { whole with Blocks = blocks }
    | Some boxId ->
        let blocks =
            if obj.ReferenceEquals(whole.Blocks, session.Main.Blocks) then session.EntryBlocks
            else
                let mutable found = Unchecked.defaultof<Doc.Block list>
                let mutable known = false
                for struct (knownBlocks, knownBody) in session.BoxFolded do
                    if not known && obj.ReferenceEquals(knownBody, whole.Blocks) then
                        found <- knownBlocks
                        known <- true
                if known then found
                else
                    match whole.Blocks |> List.tryPick (function Doc.TextBox box when box.Id = boxId -> Some box | _ -> None) with
                    | Some box when not box.Content.IsEmpty -> box.Content |> List.map Doc.Paragraph
                    | _ -> session.EntryBlocks
        { whole with Blocks = blocks }
    | None ->
    let parts = if session.IsHeader then whole.HeaderParts else whole.FooterParts
    let blocks =
        match List.tryItem session.PartIndex parts with
        // (A part without a block is the story as entered too: its one paragraph is made by the first edit.)
        | Some part -> if obj.ReferenceEquals(part.Blocks, session.PartBlocks) || part.Blocks.IsEmpty then session.EntryBlocks else part.Blocks
        // (A header that the first edit makes, before that edit: as entered.)
        | None -> session.EntryBlocks
    // (The story's document has the page setup it was entered with; the references of a
    // header made by an edit are put back when it is folded.)
    let pageSetup =
        if obj.ReferenceEquals(whole.PageSetup, session.Edited.PageSetup) then session.Main.PageSetup else whole.PageSetup
    { whole with Blocks = blocks; PageSetup = pageSetup }
/// An undo step made in a story. The history holds whole documents (a step
/// can be undone after the story was left, or from another one): the inner
/// command, which works on the story's document, is given that of the whole
/// document and its result is put back in.
type private StoryCommand(inner: ICommand, session: StorySession, stepped: StorySession -> unit) =
    interface ICommand with
        member _.Execute whole =
            stepped session
            foldStory session (inner.Execute (unfoldStory session whole))
        member _.Undo whole =
            stepped session
            foldStory session (inner.Undo (unfoldStory session whole))
        member _.Description = inner.Description

/// A step that changes the whole document from within a header or footer
/// (the section's reference to it: "link to previous"). Undone or redone, the
/// caret is in that page's header again.
type private StoryPlaceCommand(inner: ICommand, pageIndex: int, isHeader: bool, stepped: int -> bool -> unit) =
    interface ICommand with
        member _.Execute whole =
            stepped pageIndex isHeader
            inner.Execute whole
        member _.Undo whole =
            stepped pageIndex isHeader
            inner.Undo whole
        member _.Description = inner.Description
let private styleOfCharacterAt (doc: Doc.Document) (position: Doc.Position) : TextBuffer.TextStyle option =
    Doc.findParagraph doc position.BlockId
    |> Option.bind (fun para ->
        para.Buffer.Runs
        |> List.tryPick (function
            | TextBuffer.Text run when run.Start <= position.Offset && position.Offset < run.Start + run.Length -> Some run.Style
            | _ -> None))

/// The document with one character typed at the caret - with the pending
/// format, if there is one - and selected: what a format command does to it
/// is what it does to the text typed next, and its format is the caret's.
/// None when nothing can be typed at the selection (no caret).
let private pendingProbe (pending: PendingFormat option) (doc: Doc.Document) : (Doc.Document * Doc.Position) option =
    match doc.Selection with
    | Some selection when selection.IsCollapsed ->
        let at = selection.Focus
        let inserted = Doc.insertText "x" doc
        match inserted.Selection with
        | Some after when after.Focus.BlockId = at.BlockId && after.Focus.Offset = at.Offset + 1 ->
            let styled =
                match pending, Doc.findParagraph inserted at.BlockId with
                | Some pending, Some para ->
                    Doc.replaceParagraph inserted para.Id { para with Buffer = TextBuffer.applyStyle at.Offset 1 (fun _ -> pending.PendingStyle) para.Buffer }
                | _ -> inserted
            Some (Doc.setSelection at { at with Offset = at.Offset + 1 } styled, at)
        | _ -> None
    | _ -> None
/// A character of a word (Ctrl+Left / Right and a double click go by them).
let isWordChar (ch: char) = System.Char.IsLetterOrDigit ch || ch = '_' || ch = '\'' || ch = '’'

/// The word a caret is inside of - a character of the word before it and
/// one after it -, from its start to its end. None at a word's start or end,
/// in white space, or with text selected.
let wordAroundCaret (doc: Doc.Document) : (Doc.Position * Doc.Position) option =
    match doc.Selection with
    | Some selection when selection.IsCollapsed ->
        let at = selection.Focus
        match Doc.findParagraph doc at.BlockId with
        | Some para ->
            let text = TextBuffer.toString para.Buffer
            if at.Offset > 0 && at.Offset < text.Length && isWordChar text.[at.Offset - 1] && isWordChar text.[at.Offset] then
                let mutable first = at.Offset - 1
                while first > 0 && isWordChar text.[first - 1] do first <- first - 1
                let mutable last = at.Offset + 1
                while last < text.Length && isWordChar text.[last] do last <- last + 1
                Some ({ at with Offset = first }, { at with Offset = last })
            else None
        | None -> None
    | _ -> None

/// The style of text typed with a pending format: the format that was chosen,
/// and of what the text is part of - a link, a comment's range, a tracked
/// insertion - what typing there gives without one. It is not the bookmarks,
/// the field or the note of the run the caret was in when the format was
/// chosen (the chosen style is that of a character of that run): two
/// characters typed one after the other are one run, and a bookmark around
/// the run they are typed into is not theirs.
let private pendingStyleOver (chosen: TextBuffer.TextStyle) (typed: TextBuffer.TextStyle) : TextBuffer.TextStyle =
    { typed with
        Style = chosen.Style; Bold = chosen.Bold; Italic = chosen.Italic; Underline = chosen.Underline
        UnderlineStyle = chosen.UnderlineStyle; UnderlineColor = chosen.UnderlineColor; AllCaps = chosen.AllCaps
        SmallCaps = chosen.SmallCaps; Strike = chosen.Strike; DoubleStrike = chosen.DoubleStrike; Hidden = chosen.Hidden
        Position = chosen.Position; EmphasisMark = chosen.EmphasisMark; FontSize = chosen.FontSize
        CharacterSpacing = chosen.CharacterSpacing; FontFamily = chosen.FontFamily; Color = chosen.Color
        Highlight = chosen.Highlight; VertAlign = chosen.VertAlign; Explicit = chosen.Explicit
        BookmarkStarts = []; BookmarkEnds = []; RenderedPageBreakBefore = false }

/// `after`: `before` with text typed at its caret. The text has the pending
/// format of that caret, if there is one. (What says that the text is a
/// tracked insertion stays as the insertion made it.)
let private withPendingFormat (pending: PendingFormat option) (before: Doc.Document) (after: Doc.Document) : Doc.Document =
    match pending with
    | Some pending when obj.ReferenceEquals(pending.PendingDocument.Blocks, before.Blocks) && pending.PendingDocument.Selection = before.Selection ->
        match before.Selection, after.Selection with
        | Some was, Some now when was.IsCollapsed && now.IsCollapsed && now.Focus.BlockId = was.Focus.BlockId && now.Focus.Offset > was.Focus.Offset ->
            match Doc.findParagraph after now.Focus.BlockId with
            | Some para ->
                let styleOf (typed: TextBuffer.TextStyle) = pendingStyleOver pending.PendingStyle typed
                let buffer = TextBuffer.applyStyle was.Focus.Offset (now.Focus.Offset - was.Focus.Offset) styleOf para.Buffer
                { Doc.replaceParagraph after para.Id { para with Buffer = buffer } with Selection = after.Selection }
            | None -> after
        | Some was, Some now when was.IsCollapsed && now.IsCollapsed && now.Focus.BlockId <> was.Focus.BlockId ->
            let styleOf (typed: TextBuffer.TextStyle) = pendingStyleOver pending.PendingStyle typed
            let ranged = Doc.setSelection was.Focus now.Focus after
            let styled = Doc.applyFormattingToSelection (fun start len buffer -> TextBuffer.applyStyle start len styleOf buffer) ranged
            { styled with Selection = after.Selection }
        | _ -> after
    | _ -> after
/// The paragraphs that end the sections before the last one, in order.
let private sectionEnds (doc: Doc.Document) : Doc.ParagraphBlock list =
    doc.Blocks
    |> List.choose (function
        | Doc.Paragraph para
        | Doc.Heading (_, para) when para.Metadata.SectionBreak.IsSome -> Some para
        | _ -> None)

/// The header (or footer) references a section states itself: (kind, relationship id).
let sectionReferences (doc: Doc.Document) (section: int) (isHeader: bool) : (string * string) list =
    match List.tryItem section (sectionEnds doc) with
    | Some para -> if isHeader then para.Metadata.HeaderReferences else para.Metadata.FooterReferences
    | None -> if isHeader then doc.PageSetup.HeaderReferences else doc.PageSetup.FooterReferences

let withSectionReferences (doc: Doc.Document) (section: int) (isHeader: bool) (references: (string * string) list) : Doc.Document =
    match List.tryItem section (sectionEnds doc) with
    | Some para ->
        let metadata = if isHeader then { para.Metadata with HeaderReferences = references } else { para.Metadata with FooterReferences = references }
        Doc.replaceParagraph doc para.Id { para with Metadata = metadata }
    | None ->
        { doc with PageSetup = (if isHeader then { doc.PageSetup with HeaderReferences = references } else { doc.PageSetup with FooterReferences = references }) }

let freshId (prefix: string) = prefix + System.Guid.NewGuid().ToString().Replace("-", "")

/// A relationship id no header or footer of the document has.
let private freshPartId (doc: Doc.Document) =
    let used = System.Collections.Generic.HashSet<string>()
    for part in doc.HeaderParts @ doc.FooterParts do
        match part.RelationshipId with
        | Some id -> used.Add id |> ignore
        | None -> ()
    let mutable number = 1
    while used.Contains (sprintf "rIdRbtHf%d" number) do
        number <- number + 1
    sprintf "rIdRbtHf%d" number

/// The document with a new header (or footer) of `kind` for a section,
/// holding `blocks` (none: one empty paragraph), and the part's place.
let withNewPart (doc: Doc.Document) (section: int) (isHeader: bool) (kind: string) (blocks: Doc.Block list) : Doc.Document * int =
    let blocks =
        if blocks.IsEmpty then
            let para : Doc.ParagraphBlock =
                { Id = freshId (if isHeader then "header-" else "footer-")
                  Buffer = TextBuffer.empty
                  Metadata = { Doc.defaultMetadata with Style = Some (if isHeader then "Header" else "Footer") } }
            [ Doc.Paragraph para ]
        else blocks
    let id = freshPartId doc
    let part : Doc.HeaderFooterPart = { Kind = kind; RelationshipId = Some id; Blocks = blocks }
    let references = (sectionReferences doc section isHeader |> List.filter (fun (referenceKind, _) -> referenceKind <> kind)) @ [ kind, id ]
    let doc = withSectionReferences doc section isHeader references
    if isHeader then { doc with HeaderParts = doc.HeaderParts @ [ part ]; Headers = [] }, doc.HeaderParts.Length
    else { doc with FooterParts = doc.FooterParts @ [ part ]; Footers = [] }, doc.FooterParts.Length

/// Construct the editable header/footer story; a missing part is materialized only on edit.
let headerFooterStory (doc: Document) (pageIndex: int) (isHeader: bool) (kind: string) (section: int) (targetPart: Doc.HeaderFooterPart option) (label: string) (fields: Doc.Block list -> Doc.Block list) =
    let parts = if isHeader then doc.HeaderParts else doc.FooterParts
    let existing = targetPart |> Option.bind (fun part -> parts |> List.tryFindIndex (fun candidate -> obj.ReferenceEquals(candidate, part)))
    let edited, partIndex =
        match existing with
        | Some index ->
            match List.item index parts with
            | part when part.Blocks.IsEmpty ->
                let para: Doc.ParagraphBlock =
                    { Id = freshId (if isHeader then "header-" else "footer-"); Buffer = TextBuffer.empty
                      Metadata = { Doc.defaultMetadata with Style = Some (if isHeader then "Header" else "Footer") } }
                let filled = parts |> List.mapi (fun at candidate -> if at = index then { candidate with Blocks = [ Doc.Paragraph para ] } else candidate)
                (if isHeader then { doc with HeaderParts = filled } else { doc with FooterParts = filled }), index
            | _ -> doc, index
        | None ->
            let seed =
                if kind <> "default" || not parts.IsEmpty then []
                else (if isHeader then doc.Headers else doc.Footers) |> List.map Doc.Paragraph
            withNewPart doc section isHeader kind seed
    List.tryItem partIndex (if isHeader then edited.HeaderParts else edited.FooterParts)
    |> Option.map (fun part ->
        { IsHeader = isHeader; PageIndex = pageIndex; Kind = kind; Section = section; PartIndex = partIndex; Label = label
          Main = doc; Edited = edited; PartBlocks = part.Blocks; EntryBlocks = fields part.Blocks
          Folded = ResizeArray(); Box = None; BoxFolded = ResizeArray(); Note = None; NoteFolded = ResizeArray() })
/// A character operation and its caret policy, independent of the host that
/// invoked it. Dialog and named-style changes intentionally target future
/// typing at a caret, even when the caret is inside a word.
[<RequireQualifiedAccess>]
type CharacterCommand =
    | Bold | Italic | Underline | AllCaps | SmallCaps | Strike
    | Clear | GrowFont of increase: bool * standardSizes: bool
    | Superscript | Subscript
    | FontSize of float | FontFamily of string
    | Color of string option | Highlight of string option | VertAlign of string option
    | NamedStyle of string
    | FontDialog of (TextBuffer.TextStyle -> TextBuffer.TextStyle)

let private characterOperation command =
    match command with
    | CharacterCommand.Bold -> WordStyles.toggleBold, false
    | CharacterCommand.Italic -> WordStyles.toggleItalic, false
    | CharacterCommand.Underline -> WordStyles.toggleUnderline, false
    | CharacterCommand.AllCaps -> WordStyles.toggleAllCaps, false
    | CharacterCommand.SmallCaps -> WordStyles.toggleSmallCaps, false
    | CharacterCommand.Strike -> WordStyles.toggleStrike, false
    | CharacterCommand.Clear -> Doc.clearFormatting, false
    | CharacterCommand.GrowFont (increase, standardSizes) -> WordStyles.growFont increase standardSizes, false
    | CharacterCommand.Superscript -> WordStyles.toggleSuperscript, false
    | CharacterCommand.Subscript -> WordStyles.toggleSubscript, false
    | CharacterCommand.FontSize size -> Doc.setFontSize size, false
    | CharacterCommand.FontFamily family -> Doc.setFontFamily family, false
    | CharacterCommand.Color color -> Doc.setColor color, false
    | CharacterCommand.Highlight color -> Doc.setHighlight color, false
    | CharacterCommand.VertAlign align -> Doc.setVertAlign align, false
    | CharacterCommand.NamedStyle style -> WordStyles.setCharacterStyle style, true
    | CharacterCommand.FontDialog change ->
        (fun doc -> Doc.applyFormattingToSelection (fun start length buffer -> TextBuffer.applyStyle start length change buffer) doc), true

/// An immutable transfer from one selection. A delayed cut can only remove
/// that same selection in that same session/history epoch.
type ClipboardCapture = {
    Text: string
    Html: string
    Clip: Doc.Clip
    Source: Document
    History: CommandHistory
    Story: StorySession option
    Epoch: obj
    SelectionVersion: int
}

[<RequireQualifiedAccess>]
type Navigation =
    | Character of forward: bool * extend: bool
    | Word of forward: bool * extend: bool
    | Paragraph of forward: bool * extend: bool
    | DocumentEdge of toEnd: bool * extend: bool
    | VisualTarget of position: Doc.Position option * toEnd: bool * extend: bool
    | SelectAll

let private focusAt (doc: Document) extend position =
    match doc.Selection with
    | Some selection when extend -> { doc with Selection = Some { selection with Focus = position } }
    | _ -> Doc.setSelection position position doc

let private wordBoundary (text: string) offset forward =
    let mutable index = max 0 (min text.Length offset)
    if forward then
        if index < text.Length then
            if isWordChar text.[index] then
                while index < text.Length && isWordChar text.[index] do index <- index + 1
            elif not (System.Char.IsWhiteSpace text.[index]) then
                while index < text.Length && not (isWordChar text.[index]) && not (System.Char.IsWhiteSpace text.[index]) do index <- index + 1
        while index < text.Length && System.Char.IsWhiteSpace text.[index] do index <- index + 1
    else
        while index > 0 && System.Char.IsWhiteSpace text.[index - 1] do index <- index - 1
        if index > 0 then
            if isWordChar text.[index - 1] then
                while index > 0 && isWordChar text.[index - 1] do index <- index - 1
            else
                while index > 0 && not (isWordChar text.[index - 1]) && not (System.Char.IsWhiteSpace text.[index - 1]) do index <- index - 1
    index

let navigate command (doc: Document) =
    let moveCharacter forward extend =
        if extend then (if forward then Doc.extendSelectionRight else Doc.extendSelectionLeft) doc
        else (if forward then Doc.moveCursorRight else Doc.moveCursorLeft) doc
    let documentEdge toEnd extend =
        match (Doc.selectAll doc).Selection with
        | Some whole -> focusAt doc extend (if toEnd then whole.Focus else whole.Anchor)
        | None -> doc
    match command with
    | Navigation.SelectAll -> Doc.selectAll doc
    | Navigation.Character (forward, extend) -> moveCharacter forward extend
    | Navigation.DocumentEdge (toEnd, extend) -> documentEdge toEnd extend
    | Navigation.VisualTarget (Some position, _, extend) -> focusAt doc extend position
    | Navigation.VisualTarget (None, toEnd, extend) -> documentEdge toEnd extend
    | Navigation.Word (forward, extend) ->
        match doc.Selection with
        | None -> doc
        | Some selection ->
            match Doc.findParagraph doc selection.Focus.BlockId with
            | None -> doc
            | Some para ->
                let text = TextBuffer.toString para.Buffer
                if (if forward then selection.Focus.Offset >= text.Length else selection.Focus.Offset <= 0) then
                    let collapsed = Doc.setSelection selection.Focus selection.Focus doc
                    let moved = (if forward then Doc.moveCursorRight else Doc.moveCursorLeft) collapsed
                    match moved.Selection with
                    | Some target -> focusAt doc extend target.Focus
                    | None -> doc
                else focusAt doc extend { selection.Focus with Offset = wordBoundary text selection.Focus.Offset forward }
    | Navigation.Paragraph (forward, extend) ->
        match doc.Selection with
        | None -> doc
        | Some selection ->
            if not forward && selection.Focus.Offset > 0 then focusAt doc extend { selection.Focus with Offset = 0 }
            else
                let collapsed =
                    match Doc.findParagraph doc selection.Focus.BlockId with
                    | Some para when forward ->
                        let endPosition = { selection.Focus with Offset = TextBuffer.length para.Buffer }
                        Doc.setSelection endPosition endPosition doc
                    | _ -> Doc.setSelection selection.Focus selection.Focus doc
                let moved = (if forward then Doc.moveCursorRight else Doc.moveCursorLeft) collapsed
                match moved.Selection with
                | Some target when target.Focus.BlockId <> selection.Focus.BlockId ->
                    focusAt doc extend { target.Focus with Offset = 0 }
                | Some target when forward -> focusAt doc extend target.Focus
                | _ -> doc

/// Portable editing operations. Dialogs pass partial transformations so fields
/// not exposed by a host remain intact.
type WordCommand =
    | Character of CharacterCommand
    | Enter | Tab of shift: bool * control: bool
    | PageBreak | ColumnBreak | LineBreak
    | AlignLeft | AlignCenter | AlignRight | AlignJustify
    | ParagraphStyle of string | HeadingLevel of int | ClearHeading
    | LineSpacing of float option | SpacingBefore of float option | SpacingAfter of float option
    | ToggleBulletList | ToggleNumberedList | ApplyBulletList | NormalStyle
    | NumberFormat of string
    | IncreaseIndent | DecreaseIndent | ListLevel of int | IndentParagraphs of bool
    | InsertTable of int * int
    | InsertRowBefore | InsertRowAfter | InsertColumnBefore | InsertColumnAfter
    | DeleteRow | DeleteColumn | MergeCells | SplitCell | MoveCell of bool
    | ParagraphMetadata of (Doc.BlockMetadata -> Doc.BlockMetadata)
    | ListDialog of (unit -> string) * Doc.ListDialogChange
    | TableCell of (Doc.TableCell -> Doc.TableCell)
    | TableProperties of (Doc.TableBlock -> int -> Doc.TableBlock)
    | ParagraphFormat of (Doc.Document -> Doc.Document)
    | TableFormat of (Doc.Document -> Doc.Document)

type CommandOutcome = NoChange | SelectionChanged | ContentChanged

let private pageBreakAllowed (story: StorySession option) (doc: Document) =
    story.IsNone
    && (match doc.Selection with
        | Some selection ->
            (Doc.textBoxOfParagraph doc selection.Focus.BlockId).IsNone
            && not (selection.Focus.BlockId.StartsWith "__footnote__")
            && not (selection.Focus.BlockId.StartsWith "__endnote__")
        | None -> false)

/// Hosts without HTML decoding pass this when the transfer is plain text.
let noExternalClip (_html: string) : Doc.Clip option = None

type private TypingRun = { Kind: string; Start: Document; After: Document; Step: ICommand }

type EditingSession(initialDocument: Document, ?maxHistory: int) =
    let mutable document = initialDocument
    let mutable observed = initialDocument
    let mutable history = { empty with MaxHistory = defaultArg maxHistory 100 }
    let mutable typing: TypingRun option = None
    let mutable story: StorySession option = None
    let mutable historyStory: StorySession option = None
    let mutable historyPlace: (int * bool) option = None
    let mutable pending: PendingFormat option = None
    let mutable selectionVersion = 0
    let mutable index = newIndex ()
    let mutable epoch = obj()
    let mutable revision = 0
    let mutable lastRevision = 0
    let mutable epochStart = 0
    let mutable openedRevision = 0
    let mutable savedRevision = 0
    let mutable wholeObservation = false
    do rememberRevision index initialDocument 0

    let observe doc =
        if not (sameContentObjects doc observed) then
            let known = findRevision index doc
            if known >= 0 then revision <- known
            else
                lastRevision <- lastRevision + 1
                revision <- lastRevision
                rememberRevision index doc revision
            observed <- doc
    let whole doc = match story with Some active -> foldStory active doc | None -> doc
    let dropPending (doc: Document) =
        match pending with
        | Some value when not (obj.ReferenceEquals(value.PendingDocument.Blocks, doc.Blocks) && value.PendingDocument.Selection = doc.Selection) -> pending <- None
        | _ -> ()
    let wrapStory command =
        match story with
        | Some active -> StoryCommand(command, active, fun stepped -> historyStory <- Some stepped) :> ICommand
        | None -> command
    let mutable richClip: (string * Doc.Clip) option = None
    let commitClipboard description after =
        let before = document
        let after = Doc.snapSelectionToMarks before.Selection after
        if sameContentObjects before after || Doc.sameContent before after then
            if before.Selection <> after.Selection then
                selectionVersion <- selectionVersion + 1
                document <- { before with Selection = after.Selection }
            NoChange
        else
            typing <- None
            pending <- None
            let cmd = wrapStory (DeleteTextCommand(before, after, description) :> ICommand)
            history <- { history with Undos = cmd :: List.truncate (history.MaxHistory - 1) history.Undos; Redos = [] }
            selectionVersion <- selectionVersion + 1
            document <- after
            if not wholeObservation then observe (whole after)
            ContentChanged

    member this.Navigate(command: Navigation) =
        let before = document
        let after = Doc.snapSelectionToMarks before.Selection (navigate command before)
        typing <- None
        if before.Selection = after.Selection then NoChange
        else
            this.Document <- { before with Selection = after.Selection }
            SelectionChanged
    member _.CaptureClipboard(htmlOf: Doc.Clip -> string) =
        typing <- None
        match Doc.copySelection document with
        | None -> None
        | Some clip ->
            let html = htmlOf clip
            richClip <- Some (html, clip)
            Some { Text = clip.Text; Html = html; Clip = clip; Source = document
                   History = history; Story = story; Epoch = epoch; SelectionVersion = selectionVersion }
    member _.CommitCut(capture: ClipboardCapture, tracking: (string * string) option) =
        typing <- None
        if not (obj.ReferenceEquals(epoch, capture.Epoch)
                && capture.SelectionVersion = selectionVersion
                && sameContentObjects capture.Source document
                && capture.Source.Selection = document.Selection
                && obj.ReferenceEquals(capture.History.Undos, history.Undos)
                && obj.ReferenceEquals(capture.History.Redos, history.Redos)
                && (match capture.Story, story with
                    | None, None -> true
                    | Some captured, Some current -> obj.ReferenceEquals(captured, current)
                    | _ -> false)) then None
        else
            let after =
                match tracking with
                | Some (author, date) -> Doc.trackDeleteSelection author date document
                | None -> Doc.deleteSelection document
            Some (commitClipboard "Cut" after)
    member _.PasteClipboard(text: string, html: string, decodeHtml: string -> Doc.Clip option, tracking: (string * string) option) =
        typing <- None
        let text = if isNull text then "" else text
        let html = if isNull html then "" else html
        let normalize (value: string) = value.Replace("\r\n", "\n").Replace("\r", "\n")
        let clip =
            if tracking.IsSome then None
            else
                match richClip with
                | Some (known, copied) when html <> "" && html = known && normalize text = normalize copied.Text -> Some copied
                | _ -> decodeHtml html
        match clip with
        | Some copied ->
            let copied =
                match story with
                | Some active when active.Box.IsSome -> Doc.clipWithoutTables copied
                | _ -> copied
            commitClipboard "Paste" (Doc.pasteClip copied document)
        | None when text = "" -> NoChange
        | None ->
            let before = document
            let mutable after = before
            match after.Selection with
            | Some selected when not selected.IsCollapsed ->
                after <- (match tracking with
                          | Some (author, date) -> Doc.trackDeleteSelection author date after
                          | None -> Doc.deleteSelection after)
            | _ -> ()
            let insert segment =
                after <- (match tracking with
                          | Some (author, date) -> Doc.insertTrackedText author date segment after
                          | None -> Doc.insertText segment after)
            let mutable start = 0
            let mutable index = 0
            while index < text.Length do
                if text.[index] = '\r' || text.[index] = '\n' then
                    insert (text.Substring(start, index - start))
                    after <- (match tracking with
                              | Some (author, date) -> Doc.trackSplitBlock author date after
                              | None -> Doc.splitBlock after)
                    if text.[index] = '\r' && index + 1 < text.Length && text.[index + 1] = '\n' then index <- index + 1
                    start <- index + 1
                index <- index + 1
            insert (text.Substring(start))
            after <- withPendingFormat pending before after
            commitClipboard "Paste" after


    member _.Document
        with get () = document
        and set value =
            if document.Selection <> value.Selection then selectionVersion <- selectionVersion + 1
            dropPending value
            document <- value
            if not wholeObservation then observe (whole value)
    member _.History
        with get () = history
        and set value = history <- value
    member _.UseWholeObservation with get () = wholeObservation and set value = wholeObservation <- value
    member _.ObserveWhole(doc: Document) = observe doc
    member _.Story with get () = story and set value = story <- value
    member _.HistoryPlace with get () = historyPlace and set value = historyPlace <- value
    member _.FoldStory(active: StorySession, doc: Document) = foldStory active doc
    member _.Whole(doc: Document) = whole doc
    member _.WholeDocument = whole document
    member _.WrapStory(command: ICommand) = wrapStory command
    member _.WrapStoryPlace(command: ICommand, pageIndex: int, isHeader: bool) =
        StoryPlaceCommand(command, pageIndex, isHeader, fun page header -> historyPlace <- Some (page, header)) :> ICommand
    member _.PendingFormat =
        match pending with
        | Some value when obj.ReferenceEquals(value.PendingDocument.Blocks, document.Blocks) && value.PendingDocument.Selection = document.Selection -> Some value
        | _ -> None
    member this.PendingProbe() = pendingProbe this.PendingFormat document
    member private this.FormatAtCaret(command: Document -> Document) =
        match this.PendingProbe() with
        | Some (probe, at) ->
            match styleOfCharacterAt probe at, styleOfCharacterAt (command probe) at with
            | Some before, Some style when before <> style ->
                typing <- None
                pending <- Some { PendingDocument = document; PendingStyle = style }
            | _ -> ()
            true
        | None -> false
    member _.WithPendingFormat(before: Document, after: Document) = withPendingFormat pending before after
    member _.DropPendingFormat(doc: Document) = dropPending doc
    member _.CommitPendingInsertion(after: Document) =
        match pending with
        | Some value when obj.ReferenceEquals(value.PendingDocument.Blocks, document.Blocks) && value.PendingDocument.Selection = document.Selection ->
            pending <- Some { value with PendingDocument = after }
        | _ -> pending <- None
    member _.EnterStory(active: StorySession, selection: Selection option) =
        let selection =
            match selection with
            | Some _ -> selection
            | None ->
                match Doc.paragraphsOfBlocks active.EntryBlocks with
                | first :: _ ->
                    let position: Position = { BlockId = first.Id; Offset = 0 }
                    Some { Anchor = position; Focus = position }
                | [] -> None
        typing <- None
        pending <- None
        selectionVersion <- selectionVersion + 1
        story <- Some active
        document <- { active.Main with Blocks = active.EntryBlocks; Selection = selection }
        if not wholeObservation then observe (whole document)
    member _.ExitStory() =
        match story with
        | None -> false
        | Some active ->
            selectionVersion <- selectionVersion + 1
            document <- { foldStory active document with Selection = active.Main.Selection }
            story <- None
            typing <- None
            pending <- None
            if not wholeObservation then observe document
            true
    member _.StopTyping() = typing <- None
    member _.RecordChange(description: string, before: Document, after: Document, wrap: ICommand -> ICommand) =
        if not (sameContentObjects before after) then
            typing <- None
            let cmd = wrapStory (wrap (DeleteTextCommand(before, after, description) :> ICommand))
            history <- { history with Undos = cmd :: List.truncate (history.MaxHistory - 1) history.Undos; Redos = [] }
    /// Execute one atomic command, preserving the old snapshot for a no-op.
    member this.Execute(command: WordCommand, tracking: (string * string) option) : Result<CommandOutcome, string> =
        let before = document
        let insert text =
            let after =
                match tracking with
                | Some (author, date) -> Doc.insertTrackedText author date text before
                | None -> Doc.insertText text before
            withPendingFormat pending before after
        let result =
            match command with
            | Enter ->
                match before.Selection |> Option.bind (fun selection -> Doc.findParagraph before selection.Focus.BlockId) with
                | Some para when TextBuffer.length para.Buffer = 0 && para.Metadata.ListStyle <> Doc.NoList -> Ok (Doc.decreaseIndent before)
                | Some _ ->
                    Ok (match tracking with Some (author, date) -> Doc.trackSplitBlock author date before | None -> Doc.splitBlock before)
                | None -> Ok before
            | Tab (shift, control) ->
                let context = Doc.caretContext before
                if control && not shift && context.InTable then Ok (insert "\t")
                else
                    match Doc.moveToAdjacentTableCell (not shift) before with
                    | Some moved -> Ok moved
                    | None when shift && not context.InTable -> Ok (Doc.decreaseIndent before)
                    | None when not shift && context.InList -> Ok (Doc.increaseIndent before)
                    | None when not shift && not context.InTable -> Ok (insert "\t")
                    | None -> Ok before
            | PageBreak -> Ok (if pageBreakAllowed story before then Doc.insertPageBreak before else before)
            | ColumnBreak -> Ok (if pageBreakAllowed story before then Doc.insertColumnBreak before else before)
            | LineBreak -> Ok (insert "\n")
            | AlignLeft -> Ok (Doc.alignLeft before)
            | AlignCenter -> Ok (Doc.alignCenter before)
            | AlignRight -> Ok (Doc.alignRight before)
            | AlignJustify -> Ok (Doc.alignJustify before)
            | ParagraphStyle style -> Ok (WordStyles.setParagraphStyle style before)
            | HeadingLevel level -> Ok (Doc.setHeadingLevel level before)
            | ClearHeading -> Ok (Doc.clearHeading before)
            | LineSpacing value -> Ok (Doc.setLineSpacing value before)
            | SpacingBefore value -> Ok (Doc.setSpacingBefore value before)
            | SpacingAfter value -> Ok (Doc.setSpacingAfter value before)
            | ToggleBulletList -> Ok (Doc.toggleBulletList before)
            | ToggleNumberedList -> Ok (Doc.toggleNumberedList before)
            | ApplyBulletList -> Ok (Doc.applyBulletList before)
            | NumberFormat format -> Ok (Doc.setNumberFormat format before)
            | NormalStyle -> Ok (Doc.setNormalStyle before)
            | IncreaseIndent -> Ok (Doc.increaseIndent before)
            | DecreaseIndent -> Ok (Doc.decreaseIndent before)
            | ListLevel level -> Ok (Doc.setListLevel level before)
            | IndentParagraphs increase -> Ok (Doc.indentParagraphs increase before)
            | InsertTable (rows, columns) -> Ok (Doc.insertTable rows columns before)
            | InsertRowBefore -> Ok (Doc.insertTableRowBeforeSelection before)
            | InsertRowAfter -> Ok (Doc.insertTableRowAfterSelection before)
            | InsertColumnBefore -> Ok (Doc.insertTableColumnBeforeSelection before)
            | InsertColumnAfter -> Ok (Doc.insertTableColumnAfterSelection before)
            | DeleteRow -> Ok (Doc.deleteSelectedTableRow before)
            | DeleteColumn -> Ok (Doc.deleteSelectedTableColumn before)
            | MergeCells -> Ok (Doc.mergeSelectedTableCells before)
            | SplitCell -> Ok (Doc.splitSelectedTableCell before)
            | MoveCell forward -> Ok (Doc.moveToAdjacentTableCell forward before |> Option.defaultValue before)
            | ParagraphMetadata change -> Ok (Doc.updateSelectedParagraphMetadata change before)
            | ListDialog (allocate, change) -> Doc.applyListDialog allocate change before
            | TableCell change -> Ok (Doc.updateSelectedTableCell change before)
            | TableProperties change -> Ok (Doc.updateSelectedTableProperties change before)
            | Character format ->
                let apply, caretOnly = characterOperation format
                if caretOnly then
                    Ok (if this.FormatAtCaret(apply) then before else apply before)
                else
                    match wordAroundCaret before with
                    | Some (first, last) ->
                        let formatted = apply (Doc.setSelection first last before)
                        Ok { formatted with Selection = before.Selection }
                    | None -> Ok (if this.FormatAtCaret(apply) then before else apply before)
            | ParagraphFormat change | TableFormat change -> Ok (change before)
        match result with
        | Error message -> Error message
        | Ok after ->
            let after = Doc.snapSelectionToMarks before.Selection after
            if sameContentObjects before after || Doc.sameContent before after then
                if before.Selection = after.Selection then Ok NoChange
                else
                    this.Document <- { before with Selection = after.Selection }
                    Ok SelectionChanged
            else
                this.RecordChange("Edit", before, after, id)
                this.Document <- after
                Ok ContentChanged
    member _.RecordTyping(kind: string, description: string, before: Document, after: Document, wrap: ICommand -> ICommand) =
        let start, below =
            match typing, history.Undos with
            | Some run, top :: rest when run.Kind = kind && obj.ReferenceEquals(top, run.Step) && obj.ReferenceEquals(before, run.After) -> run.Start, { history with Undos = rest }
            | _ -> before, history
        let cmd = wrapStory (wrap (DeleteTextCommand(start, after, description) :> ICommand))
        history <- { below with Undos = cmd :: List.truncate (below.MaxHistory - 1) below.Undos; Redos = [] }
        typing <- Some { Kind = kind; Start = start; After = after; Step = cmd }
    member _.HistoryStep(step: Document -> CommandHistory -> Document * CommandHistory) =
        typing <- None
        historyStory <- None
        historyPlace <- None
        let next, nextHistory = step (whole document) history
        if obj.ReferenceEquals(nextHistory, history) then false
        else
            pending <- None
            history <- nextHistory
            story <- historyStory
            selectionVersion <- selectionVersion + 1
            document <- match story with Some active -> unfoldStory active next | None -> next
            if not wholeObservation then observe next
            true
    member this.Undo() = this.HistoryStep(undo)
    member this.Redo() = this.HistoryStep(redo)
    member _.Replace(value: Document) =
        typing <- None
        story <- None
        historyStory <- None
        historyPlace <- None
        pending <- None
        richClip <- None
        history <- { empty with MaxHistory = history.MaxHistory }
        selectionVersion <- selectionVersion + 1
        epoch <- obj()
        index <- newIndex ()
        lastRevision <- lastRevision + 1
        revision <- lastRevision
        epochStart <- revision
        openedRevision <- revision
        savedRevision <- revision
        document <- value
        observed <- value
        rememberRevision index value revision
    member _.Epoch = epoch
    member _.LastRevision = lastRevision
    member _.CaptureSave() =
        typing <- None
        { Epoch = epoch; Document = observed; Revision = revision; Unchanged = revision = openedRevision }
    member _.MarkSaved(snapshot: SaveSnapshot) =
        if obj.ReferenceEquals(epoch, snapshot.Epoch) then savedRevision <- snapshot.Revision
    member _.MarkSavedRevision(value: int) =
        if value >= epochStart && value <= lastRevision then savedRevision <- value
    member _.TransferExtension(extend: Document -> Document, doc: Document) =
        let known = findRevision index doc
        let extended = extend doc
        if known >= 0 then rememberRevision index extended known
        extended
    member this.ExtendObserved(extend: Document -> Document) =
        observed <- this.TransferExtension(extend, observed)
    member _.IsModified = revision <> savedRevision
    member _.Revision = revision
