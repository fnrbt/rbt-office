module Rbt.Word.Input

let private sameListOrEqual left right =
    obj.ReferenceEquals(left, right) || left = right

/// The bodies of the notes and comments are the same (a note's body is edited in place, see Input.StorySession.Note).
let private sameNoteBodies (left: Rbt.Word.Document.Document) (right: Rbt.Word.Document.Document) =
    obj.ReferenceEquals(left.Package, right.Package)
    || (sameListOrEqual left.Package.Notes right.Package.Notes && sameListOrEqual left.Package.Comments right.Package.Comments)

let private documentsEqualForChange (left: Rbt.Word.Document.Document) (right: Rbt.Word.Document.Document) =
    obj.ReferenceEquals(left, right)
    || (left.Selection = right.Selection
        && left.PageSetup = right.PageSetup
        && sameListOrEqual left.Blocks right.Blocks
        && sameListOrEqual left.Styles right.Styles
        && sameListOrEqual left.Headers right.Headers
        && sameListOrEqual left.Footers right.Footers
        && sameListOrEqual left.HeaderParts right.HeaderParts
        && sameListOrEqual left.FooterParts right.FooterParts
        && sameNoteBodies left right)

/// True when two documents differ at most in their selection. Moving the
/// caret is not something Undo steps through.
/// Whether two documents are the same content by identity: every part but
/// the selection is the same object (what a caret move leaves, and what an
/// undo comes back to). Constant time; false for equal content in other objects.
let sameContentObjects (left: Rbt.Word.Document.Document) (right: Rbt.Word.Document.Document) =
    obj.ReferenceEquals(left, right)
    || (obj.ReferenceEquals(left.Blocks, right.Blocks)
        && obj.ReferenceEquals(left.PageSetup, right.PageSetup)
        && obj.ReferenceEquals(left.Styles, right.Styles)
        && obj.ReferenceEquals(left.Headers, right.Headers)
        && obj.ReferenceEquals(left.Footers, right.Footers)
        && obj.ReferenceEquals(left.HeaderParts, right.HeaderParts)
        && obj.ReferenceEquals(left.FooterParts, right.FooterParts)
        && obj.ReferenceEquals(left.Package, right.Package))

let sameContent (left: Rbt.Word.Document.Document) (right: Rbt.Word.Document.Document) =
    sameContentObjects left right
    || (left.PageSetup = right.PageSetup
        && sameListOrEqual left.Blocks right.Blocks
        && sameListOrEqual left.Styles right.Styles
        && sameListOrEqual left.Headers right.Headers
        && sameListOrEqual left.Footers right.Footers
        && sameListOrEqual left.HeaderParts right.HeaderParts
        && sameListOrEqual left.FooterParts right.FooterParts
        && sameNoteBodies left right)

#if FABLE_COMPILER
open Fable.Core
open Fable.Core.JsInterop
open Browser.Types
open Browser.Dom
open Rbt.Word.Renderer
open Rbt.Word.Commands
open Rbt.Word.Document

module Doc = Rbt.Word.Document
module Layout = Rbt.Word.Layout
module TextBuffer = Rbt.Word.TextBuffer
module Cmd = Rbt.Word.Commands
module Clipboard = Rbt.Word.Clipboard

[<Emit("$0.style.cursor = $1")>]
let setCursor (elem: HTMLCanvasElement) (cursor: string) : unit = jsNative

[<Emit("window.addEventListener($0, $1)")>]
let private addWindowListener (eventName: string) (listener: Event -> unit) : unit = jsNative

[<Emit("window.removeEventListener($0, $1)")>]
let private removeWindowListener (eventName: string) (listener: Event -> unit) : unit = jsNative

/// A `key` value that is one typed character: a single UTF-16 unit, or a
/// surrogate pair (emoji and other astral characters arrive as length 2).
/// Named keys ("Enter", "ArrowLeft", "Dead", ...) are longer or not a pair.
let private isSingleCodePoint (key: string) =
    if key.Length = 1 then
        true
    elif key.Length = 2 then
        let high = int key.[0]
        let low = int key.[1]
        high >= 0xD800 && high <= 0xDBFF && low >= 0xDC00 && low <= 0xDFFF
    else
        false

let private keyEquals (expected: char) (key: string) =
    if key.Length <> 1 then
        false
    else
        let actual = key.[0]
        actual = expected || System.Char.ToUpperInvariant actual = System.Char.ToUpperInvariant expected

/// Resize state for images
type ResizeState = {
    BlockId: string
    Offset: int
    InitialX: float
    InitialY: float
    InitialWidth: float
    InitialHeight: float
}

type DragImageState = {
    BlockId: string
    Offset: int
    InitialX: float
    InitialY: float
}

/// A shape that is being dragged with the pointer: all of it, or one of its handles.
type ShapeDragState = {
    ShapeId: string
    /// "move", or the name of the handle ("nw", "n", "ne", "e", "se", "s", "sw", "w").
    Part: string
    StartX: float
    StartY: float
    /// The shape when the drag began, and where it was painted then: its
    /// left edge, top, width and height, and the top of its page.
    Shape: Doc.TextBoxBlock
    RectX: float
    RectY: float
    RectWidth: float
    RectHeight: float
    PageTop: float
    /// The pointer went far enough for a drag (a click moves nothing).
    mutable Moved: bool
}

/// Text an IME (or a dead key) is composing. It is shown underlined at the
/// caret but is not in the document until the composition ends.
type CompositionState = {
    mutable Text: string
}

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

/// Keys typed one after the other make one undo step, as in Word: a run of
/// typed characters, of Backspaces, of Deletes. The run goes on while each
/// keystroke starts from the document the one before it left (no caret move,
/// no other edit in between) and the run's step is still the last one of the
/// history.
type TypingRun = {
    /// "type", "backspace" or "delete".
    Kind: string
    /// The document before the run's first keystroke: what its step undoes to.
    Start: Doc.Document
    /// The document after its last keystroke.
    After: Doc.Document
    /// The run's step in the history (the top of the undo list while the run goes on).
    Step: ICommand
}

/// Input handler state
type InputHandler = {
    Canvas: HTMLCanvasElement
    RenderContext: RenderContext
    mutable Document: Doc.Document
    OnChange: Doc.Document -> unit
    mutable IsMouseDown: bool
    mutable SelectionAnchor: Doc.Position option
    mutable History: CommandHistory
    mutable ResizingImage: ResizeState option
    mutable DraggingImage: DragImageState option
    mutable HoveredImage: (string * int) option  // (blockId, offset) of hovered image
    mutable TrackChanges: bool
    mutable TrackChangesAuthor: string
    mutable TrackChangesDate: string
    /// Document as it was when the current image resize/drag started, so the
    /// whole gesture becomes one undo step.
    mutable PointerStartDocument: Doc.Document option
    /// Hidden text input next to the canvas that holds the keyboard focus: a
    /// canvas gets no composition events, so IME, dead-key and virtual-keyboard
    /// text arrives through this element. Created and removed by the handler.
    mutable InputProxy: HTMLElement option
    /// The composition in progress, if any.
    mutable Composition: CompositionState option
    /// Moves the input proxy to the caret, so the IME candidate window opens
    /// there (set by `create`).
    mutable SyncInputProxy: Doc.Document -> unit
    /// Removes listeners registered outside the canvas (set by `create`).
    mutable Dispose: unit -> unit
    /// The column Up / Down / PageUp / PageDown keep to (layout x of the
    /// caret when the run of vertical moves started).
    mutable GoalX: float option
    /// The header or footer that is edited in place, if any.
    mutable Story: StorySession option
    /// The story the undo / redo step in hand was made in (set by the step).
    mutable HistoryStory: StorySession option
    /// The header or footer (page, header?) the undo / redo step in hand
    /// belongs to though it changed the whole document - "link to previous"
    /// (set by the step): the caret goes back into that page's header.
    mutable HistoryPlace: (int * bool) option
    /// Puts the caret into a page's header or footer after such a step (set by the editor).
    mutable OnHistoryPlace: int -> bool -> bool
    /// A double click at a point (layout coordinates): true when it was
    /// taken (into a header or footer, or back to the body) - set by the editor.
    mutable OnDoubleClick: float -> float -> bool
    /// Called before a key that goes by the whole document (Ctrl+A, Ctrl+End):
    /// the editor reads a document that is still being read to its end.
    mutable EnsureWhole: unit -> unit
    /// See PendingFormat (formatAtCaret sets it).
    mutable Pending: PendingFormat option
    /// Called when the pending format changed: what the selection's format
    /// is said to be has changed though the document has not (set by the editor).
    mutable OnPendingFormat: unit -> unit
    /// Put plain text in at the caret, or over the selection, as Ctrl+V does
    /// (its lines are paragraphs; one undo step). Set when the handler is made.
    mutable PasteText: string -> unit
    /// Paste what the clipboard has, its text and its HTML (either may be
    /// empty), as Ctrl+V does: what the last Copy kept, with its formatting,
    /// while the text is its text; else the HTML with its formatting; else the
    /// text. One undo step. Set when the handler is made.
    mutable PasteClipboard: string -> string -> unit
    /// Copy (true: Cut) the selection, as Ctrl+C / Ctrl+X do: kept for Paste;
    /// its text and its HTML for the system clipboard, None with nothing
    /// selected. Set when the handler is made.
    mutable CopyToClipboard: bool -> (string * string) option
    /// The shape that is selected (see selectShape): its frame and handles
    /// are painted, the pointer drags it by its frame and resizes it by its
    /// handles, Delete takes it out, the arrow keys move it, Escape lets go of it.
    mutable SelectedShape: Doc.BlockId option
    /// The drag of the selected shape that is under way.
    mutable ShapeDrag: ShapeDragState option
    /// Called when another shape, or none, is selected (set by the editor).
    mutable OnShapeSelection: unit -> unit
    /// The run of typing the last undo step is, while it can go on (see TypingRun).
    mutable Typing: TypingRun option
}

[<Emit("$0.focus({ preventScroll: true })")>]
let private focusWithoutScroll (_element: HTMLElement) : unit = jsNative

[<Emit("$0.isComposing === true || $0.keyCode === 229")>]
let private keyBelongsToComposition (_event: KeyboardEvent) : bool = jsNative

[<Emit("(typeof $0.data === 'string' ? $0.data : null)")>]
let private eventData (_event: Event) : string = jsNative

[<Emit("(typeof $0.inputType === 'string' ? $0.inputType : '')")>]
let private eventInputType (_event: Event) : string = jsNative

[<Emit("$0.value")>]
let private proxyValue (_element: HTMLElement) : string = jsNative

[<Emit("$0.value = $1")>]
let private setProxyValue (_element: HTMLElement) (_value: string) : unit = jsNative

[<Emit("$0.select()")>]
let private selectProxyText (_element: HTMLElement) : unit = jsNative

[<Emit("setTimeout($0, 0)")>]
let private soon (_action: unit -> unit) : unit = jsNative

// ---- the clipboard -----------------------------------------------------------
// Copy and Cut put the selection's text and its HTML on the system clipboard
// and keep what they copied as a clip (Doc.copySelection); Paste puts the
// clip back with its formatting while the clipboard's text is still the
// clip's, else the clipboard's HTML (from Word, a browser, a mail), else its
// text. The keys work through the browser's own copy, cut and paste events
// on the input proxy: no permission is asked, and the HTML reaches us.

/// What the last Copy or Cut of an editor on the page kept.
let mutable private lastClip : Doc.Clip option = None

[<Emit("(function (h) { try { return new DOMParser().parseFromString(h, 'text/html').body; } catch (e) { return null; } })($0)")>]
let private parseHtmlBody (_html: string) : obj = jsNative

[<Emit("Array.from($0 || [])")>]
let private arrayOf (_list: obj) : obj array = jsNative

let rec private htmlNodeOf (node: obj) : Doc.HtmlNode option =
    match unbox<int> node?nodeType with
    | 3 -> Some (Doc.HtmlText (unbox<string> node?nodeValue))
    | 1 ->
        let tag = (unbox<string> node?tagName).ToLowerInvariant()
        let attributes =
            arrayOf node?attributes
            |> Array.map (fun attribute -> (unbox<string> attribute?name).ToLowerInvariant(), unbox<string> attribute?value)
            |> List.ofArray
        let children = arrayOf node?childNodes |> Array.choose htmlNodeOf |> List.ofArray
        Some (Doc.HtmlElement (tag, attributes, children))
    | _ -> None

/// The clipboard's HTML as a clip (see Doc.clipOfHtml); None for no HTML, or HTML without text.
let htmlClip (html: string) : Doc.Clip option =
    if System.String.IsNullOrWhiteSpace html then None
    else
        let body = parseHtmlBody html
        if isNull body then None
        else arrayOf body?childNodes |> Array.choose htmlNodeOf |> List.ofArray |> Doc.clipOfHtml

// ---- a header or footer edited in place --------------------------------------

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

/// The whole document (body, headers and footers) of the document the
/// handler works on: that document, or - while a header or footer is edited
/// in place - the document with the story's blocks in their part. This is
/// what is saved, and what a revision is a state of.
let wholeDocument (handler: InputHandler) (doc: Doc.Document) : Doc.Document =
    match handler.Story with
    | Some session -> foldStory session doc
    | None -> doc

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

/// A history with the step just added to it made a story's step when a story is being edited.
let private storyHistory (handler: InputHandler) (before: CommandHistory) (after: CommandHistory) : CommandHistory =
    match handler.Story, after.Undos with
    | Some session, head :: rest when not (obj.ReferenceEquals(after.Undos, before.Undos)) ->
        { after with Undos = (StoryCommand(head, session, fun stepped -> handler.HistoryStory <- Some stepped) :> ICommand) :: rest }
    | _ -> after

let private showStory (handler: InputHandler) (session: StorySession) =
    let ctx = handler.RenderContext
    // (The whole document for the story document last asked about: a caret move asks again and again.)
    let last = ResizeArray<Doc.Document * Doc.Document>()
    let whole (story: Doc.Document) =
        if last.Count > 0 && sameContentObjects story (fst last.[0]) then snd last.[0]
        else
            let folded = { foldStory session story with Selection = None }
            last.Clear()
            last.Add((story, folded))
            folded
    handler.Story <- Some session
    ctx.Story <- Some { Whole = whole; PageIndex = session.PageIndex; IsHeader = session.IsHeader; Label = session.Label; Box = session.Box; Note = session.Note }
    ctx.StoryLayouts <- None
    ctx.StoryNotePlace <- None
    ctx.StoryNoteLayouts <- None

let private hideStory (handler: InputHandler) =
    handler.Story <- None
    handler.RenderContext.Story <- None
    handler.RenderContext.StoryLayouts <- None
    handler.RenderContext.StoryNotePlace <- None
    handler.RenderContext.StoryNoteLayouts <- None

/// Record a finished change in the undo history as one step. Selection-only
/// changes are not recorded.
let private recordChange (handler: InputHandler) (description: string) (beforeDoc: Doc.Document) (afterDoc: Doc.Document) =
    if not (sameContent beforeDoc afterDoc) then
        handler.Typing <- None
        let cmd = DeleteTextCommand(beforeDoc, afterDoc, description) :> ICommand
        let _, history = Cmd.execute cmd beforeDoc handler.History
        handler.History <- storyHistory handler handler.History history

/// The history with a keystroke of a run of typing (`kind`: see TypingRun)
/// that made `afterDoc` of `beforeDoc`: the run's step undoes it too when the
/// keystroke goes on from where the run's last one left the document,
/// otherwise the keystroke starts a run of its own, a new step.
let private typingHistory (handler: InputHandler) (kind: string) (description: string) (beforeDoc: Doc.Document) (afterDoc: Doc.Document) : CommandHistory =
    let continued =
        match handler.Typing, handler.History.Undos with
        | Some run, top :: rest when run.Kind = kind && obj.ReferenceEquals(top, run.Step) && obj.ReferenceEquals(beforeDoc, run.After) -> Some (run, rest)
        | _ -> None
    let start, below =
        match continued with
        | Some (run, rest) -> run.Start, { handler.History with Undos = rest }
        | None -> beforeDoc, handler.History
    let cmd = DeleteTextCommand(start, afterDoc, description) :> ICommand
    let _, history = Cmd.execute cmd start below
    let history = storyHistory handler below history
    handler.Typing <- Some { Kind = kind; Start = start; After = afterDoc; Step = List.head history.Undos }
    history

/// A keystroke of a run of typing made `afterDoc` of the handler's document:
/// recorded (see typingHistory) and shown.
let private commitTyping (handler: InputHandler) (kind: string) (description: string) (afterDoc: Doc.Document) : bool =
    let beforeDoc = handler.Document
    if sameContent beforeDoc afterDoc then false
    else
        handler.History <- typingHistory handler kind description beforeDoc afterDoc
        true

let private showDocument (handler: InputHandler) (doc: Doc.Document) =
    handler.Document <- doc
    renderWithHover handler.RenderContext doc handler.HoveredImage
    handler.SyncInputProxy doc
    handler.OnChange doc

/// One step back or forward in the history (`step`: Cmd.undo or Cmd.redo).
/// The step is taken on the whole document; what comes back is the document
/// to work on - the story's when the step was made in a header or footer
/// (which is then the story on display), the whole document otherwise.
let private historyStep (handler: InputHandler) (step: Doc.Document -> CommandHistory -> Doc.Document * CommandHistory) : Doc.Document * CommandHistory =
    // (What is typed after an undo or a redo is a step of its own.)
    handler.Typing <- None
    handler.HistoryStory <- None
    handler.HistoryPlace <- None
    let whole, history = step (wholeDocument handler handler.Document) handler.History
    // (Nothing to step to: everything stays as it is.)
    if obj.ReferenceEquals(history, handler.History) then handler.Document, history else
    match handler.HistoryStory with
    | Some session ->
        showStory handler session
        unfoldStory session whole, history
    | None ->
        if handler.Story.IsSome then hideStory handler
        whole, history

/// Show the document a step of the history came to. A step that belongs to a
/// page's header or footer (see HistoryPlace) is shown there: the editor
/// enters it, which shows the document.
let private showHistoryStep (handler: InputHandler) (doc: Doc.Document) =
    match handler.HistoryPlace with
    | Some (pageIndex, isHeader) ->
        handler.HistoryPlace <- None
        handler.Document <- doc
        handler.OnChange doc
        // (No such page any more: the document as it is, in the body.)
        if not (handler.OnHistoryPlace pageIndex isHeader) then showDocument handler doc
    | None -> showDocument handler doc

/// Undo the last recorded change. Returns false when there is nothing to undo.
let undo (handler: InputHandler) =
    if handler.History.Undos.IsEmpty then
        false
    else
        let doc, history = historyStep handler Cmd.undo
        handler.History <- history
        showHistoryStep handler doc
        true

/// Redo the last undone change. Returns false when there is nothing to redo.
let redo (handler: InputHandler) =
    if handler.History.Redos.IsEmpty then
        false
    else
        let doc, history = historyStep handler Cmd.redo
        handler.History <- history
        showHistoryStep handler doc
        true

/// Edit a header or footer in place: `session` says which (see StorySession;
/// its Main is the whole document as it is now), `selection` where the caret
/// goes (None: the start of the story).
let enterStory (handler: InputHandler) (session: StorySession) (selection: Doc.Selection option) =
    let selection =
        match selection with
        | Some selection -> Some selection
        | None ->
            match Doc.paragraphsOfBlocks session.EntryBlocks with
            | first :: _ ->
                let position : Doc.Position = { BlockId = first.Id; Offset = 0 }
                Some { Anchor = position; Focus = position }
            | [] -> None
    showStory handler session
    handler.IsMouseDown <- false
    handler.SelectionAnchor <- None
    handler.GoalX <- None
    showDocument handler { session.Main with Blocks = session.EntryBlocks; Selection = selection }

/// Back to the body, with the selection it had. False when no header or footer was being edited.
let exitStory (handler: InputHandler) : bool =
    match handler.Story with
    | None -> false
    | Some session ->
        let whole = { foldStory session handler.Document with Selection = session.Main.Selection }
        hideStory handler
        handler.IsMouseDown <- false
        handler.SelectionAnchor <- None
        handler.GoalX <- None
        showDocument handler whole
        true

/// The pending format of the caret, if it has one: chosen in this text with the caret where it is.
let pendingFormat (handler: InputHandler) : PendingFormat option =
    match handler.Pending with
    | Some pending when obj.ReferenceEquals(pending.PendingDocument.Blocks, handler.Document.Blocks) && pending.PendingDocument.Selection = handler.Document.Selection -> Some pending
    | _ -> None

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
let pendingProbe (handler: InputHandler) : (Doc.Document * Doc.Position) option =
    let doc = handler.Document
    match doc.Selection with
    | Some selection when selection.IsCollapsed ->
        let at = selection.Focus
        let inserted = Doc.insertText "x" doc
        match inserted.Selection with
        | Some after when after.Focus.BlockId = at.BlockId && after.Focus.Offset = at.Offset + 1 ->
            let styled =
                match pendingFormat handler, Doc.findParagraph inserted at.BlockId with
                | Some pending, Some para ->
                    Doc.replaceParagraph inserted para.Id { para with Buffer = TextBuffer.applyStyle at.Offset 1 (fun _ -> pending.PendingStyle) para.Buffer }
                | _ -> inserted
            Some (Doc.setSelection at { at with Offset = at.Offset + 1 } styled, at)
        | _ -> None
    | _ -> None

/// A character format command given with nothing selected: it is kept for
/// the text typed next at the caret (see PendingFormat). `command` is the
/// command as it is applied to a selection. False when there is no caret
/// (text is selected, or nothing is): the command is then the selection's.
let formatAtCaret (handler: InputHandler) (command: Doc.Document -> Doc.Document) : bool =
    match pendingProbe handler with
    | Some (probe, at) ->
        match styleOfCharacterAt probe at, styleOfCharacterAt (command probe) at with
        | Some before, Some style when before <> style ->
            // (What is typed in the format chosen is a step of its own.)
            handler.Typing <- None
            handler.Pending <- Some { PendingDocument = handler.Document; PendingStyle = style }
            handler.OnPendingFormat ()
        | _ -> ()
        true
    | None -> false

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

/// The document after a character format command, as Word gives it: on the
/// selection; with the caret inside a word on that word (the caret stays
/// where it was); with the caret elsewhere for the text typed next (see
/// formatAtCaret: the document is as it was).
let characterFormatted (handler: InputHandler) (command: Doc.Document -> Doc.Document) : Doc.Document =
    match wordAroundCaret handler.Document with
    | Some (first, last) ->
        let formatted = command (Doc.setSelection first last handler.Document)
        { formatted with Selection = handler.Document.Selection }
    | None -> if formatAtCaret handler command then handler.Document else command handler.Document

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
let withPendingFormat (handler: InputHandler) (before: Doc.Document) (after: Doc.Document) : Doc.Document =
    match handler.Pending with
    | Some pending when obj.ReferenceEquals(pending.PendingDocument.Blocks, before.Blocks) && pending.PendingDocument.Selection = before.Selection ->
        match before.Selection, after.Selection with
        | Some was, Some now when was.IsCollapsed && now.IsCollapsed && now.Focus.BlockId = was.Focus.BlockId && now.Focus.Offset > was.Focus.Offset ->
            match Doc.findParagraph after now.Focus.BlockId with
            | Some para ->
                let styleOf (typed: TextBuffer.TextStyle) = pendingStyleOver pending.PendingStyle typed
                let buffer = TextBuffer.applyStyle was.Focus.Offset (now.Focus.Offset - was.Focus.Offset) styleOf para.Buffer
                { Doc.replaceParagraph after para.Id { para with Buffer = buffer } with Selection = after.Selection }
            | None -> after
        | _ -> after
    | _ -> after

/// A pending format does not outlive a change of the document or of the selection.
let private dropPendingFormat (handler: InputHandler) (doc: Doc.Document) =
    match handler.Pending with
    | Some pending when not (obj.ReferenceEquals(pending.PendingDocument.Blocks, doc.Blocks) && pending.PendingDocument.Selection = doc.Selection) -> handler.Pending <- None
    | _ -> ()

/// Edit the text of a text box of the body in place (as a header is: see
/// `enterStory`): typing, selecting, formatting and undo work in it; Escape,
/// or a click outside the box, goes back to the body. `selection`: where the
/// caret goes (None: the start of the box's text). False when the document
/// has no such box, or its text is not laid out as text (a watermark), or it
/// is a group or a drawing canvas (the text of its shapes is not edited).
///
/// A text box of a header or footer is edited where `page` paints it when
/// that page shows it (the page the pointer is on), else on the first page
/// that does.
let enterTextBoxOnPage (handler: InputHandler) (boxId: Doc.BlockId) (selection: Doc.Selection option) (page: int option) : bool =
    // (Another story - a header, another box - is left first.)
    if handler.Story.IsSome then exitStory handler |> ignore
    let doc = handler.Document
    match doc.Blocks |> List.tryPick (function Doc.TextBox box when box.Id = boxId -> Some box | _ -> None) with
    | Some box when not box.FitText && box.Children.IsEmpty ->
        // (A box without a paragraph gets one with the first edit.)
        let entry = (if box.Content.IsEmpty then [ Doc.newEmptyParagraph () ] else box.Content) |> List.map Doc.Paragraph
        let pageIndex =
            match shapeRects handler.RenderContext doc |> List.tryFind (fun (textBox, _, _, _, _) -> textBox.Id = boxId) with
            | Some (_, x, y, _, height) ->
                match pageRegionAt handler.RenderContext doc x (y + height / 2.0) with
                | Some (index, _) -> index
                | None -> 0
            | None -> 0
        let session : StorySession =
            { IsHeader = false; PageIndex = pageIndex; Kind = "textbox"; Section = 0; PartIndex = -1; Label = "Text box"
              Main = doc; Edited = doc; PartBlocks = []; EntryBlocks = entry; Folded = ResizeArray()
              Box = Some boxId; BoxFolded = ResizeArray(); Note = None; NoteFolded = ResizeArray() }
        enterStory handler session selection
        true
    | Some _ -> false
    | None ->
        // A text box of a header or footer: edited where the first page that shows its part paints it.
        let inParts (isHeader: bool) (parts: Doc.HeaderFooterPart list) =
            parts
            |> List.mapi (fun index part -> index, part)
            |> List.tryPick (fun (index, part) ->
                part.Blocks |> List.tryPick (function Doc.TextBox box when box.Id = boxId -> Some (isHeader, index, box) | _ -> None))
        match inParts true doc.HeaderParts |> Option.orElseWith (fun () -> inParts false doc.FooterParts) with
        | Some (isHeader, partIndex, box) when not box.FitText && box.Children.IsEmpty ->
            let onPage =
                match page with
                | Some pageIndex when pageShapeRects handler.RenderContext doc pageIndex |> List.exists (fun (textBox, _, _, _, _, _) -> textBox.Id = boxId) -> Some pageIndex
                | _ ->
                    furnitureShapeRects handler.RenderContext doc
                    |> List.tryFind (fun (textBox, _, _, _, _, _, _) -> textBox.Id = boxId)
                    |> Option.map (fun (_, _, _, _, _, _, pageIndex) -> pageIndex)
            match onPage with
            | Some pageIndex ->
                let entry = (if box.Content.IsEmpty then [ Doc.newEmptyParagraph () ] else box.Content) |> List.map Doc.Paragraph
                let session : StorySession =
                    { IsHeader = isHeader; PageIndex = pageIndex; Kind = "textbox"; Section = 0; PartIndex = partIndex; Label = "Text box"
                      Main = doc; Edited = doc; PartBlocks = []; EntryBlocks = entry; Folded = ResizeArray()
                      Box = Some boxId; BoxFolded = ResizeArray(); Note = None; NoteFolded = ResizeArray() }
                enterStory handler session selection
                true
            // (No page that is laid out shows that header.)
            | None -> false
        | _ -> false

/// Whether a page (or a column) can break at the caret: in the body's own
/// text - not in a text box, a header or a footer (the story that is edited
/// is not the body), in a footnote or an endnote, or with nothing selected.
let pageBreakAllowed (handler: InputHandler) : bool =
    handler.Story.IsNone
    && (match handler.Document.Selection with
        | Some selection ->
            (Doc.textBoxOfParagraph handler.Document selection.Focus.BlockId).IsNone
            && not (selection.Focus.BlockId.StartsWith Layout.footnoteBlockIdPrefix)
            && not (selection.Focus.BlockId.StartsWith Layout.endnoteBlockIdPrefix)
        | None -> false)

let enterTextBox (handler: InputHandler) (boxId: Doc.BlockId) (selection: Doc.Selection option) : bool =
    enterTextBoxOnPage handler boxId selection None

/// Which shape is selected, without painting: the renderer paints its frame
/// and handles with the next frame it builds. True when it is another one.
let private setSelectedShape (handler: InputHandler) (id: Doc.BlockId option) : bool =
    if handler.SelectedShape = id then false
    else
        handler.SelectedShape <- id
        handler.ShapeDrag <- None
        handler.RenderContext.SelectedShape <- id
        handler.RenderContext.LastPaint <- None
        true

/// Select a shape (a text box) of the body, a header or a footer - None: no
/// shape. Its frame and its eight handles are painted; see
/// InputHandler.SelectedShape for what the pointer and the keys then do.
let selectShape (handler: InputHandler) (id: Doc.BlockId option) : unit =
    if setSelectedShape handler id then
        renderWithHover handler.RenderContext handler.Document handler.HoveredImage
        handler.OnShapeSelection ()

/// The document with a shape `dx` further right and `dy` further down and -
/// with `size` (width, height) - of another size, in one change. The shape
/// keeps the frame its place is measured from and its offset from it changes;
/// in a direction it was aligned in (or had no offset of its own) it is by
/// its offset from its page from then on: `rectX` and `rectY - pageTop` are
/// where it is painted on its page. None: the document has no such shape.
let shapePlacedDocument (doc: Doc.Document) (shape: Doc.TextBoxBlock) (rectX: float) (rectY: float) (pageTop: float) (dx: float) (dy: float) (size: (float * float) option) : Doc.Document option =
    let horizontal =
        if shape.Inline || abs dx <= 0.0001 then None
        else
            match shape.PositionHAlign, shape.X with
            | None, Some offset -> Some (None, offset + dx)
            | _ -> Some (Some "page", rectX + dx)
    let vertical =
        if shape.Inline || abs dy <= 0.0001 then None
        else
            match shape.PositionVAlign, shape.Y with
            | None, Some offset -> Some (None, offset + dy)
            | _ -> Some (Some "page", rectY - pageTop + dy)
    Doc.placeTextBox shape.Id horizontal vertical size doc

/// The document with a shape resized by one of its handles (see
/// Doc.resizedShapeRect): the handle dragged by (dx, dy), the side across
/// from it where it was. One change - the shape's place and its size.
let shapeResizedDocument (doc: Doc.Document) (shape: Doc.TextBoxBlock) (rectX: float) (rectY: float) (width: float) (height: float) (pageTop: float) (handle: string) (dx: float) (dy: float) : Doc.Document option =
    let left, top, newWidth, newHeight = Doc.resizedShapeRect shape.Rotation rectX rectY width height handle dx dy
    let size = if abs (newWidth - width) > 0.0001 || abs (newHeight - height) > 0.0001 then Some (newWidth, newHeight) else None
    shapePlacedDocument doc shape rectX rectY pageTop (left - rectX) (top - rectY) size

let canUndo (handler: InputHandler) = not handler.History.Undos.IsEmpty
let canRedo (handler: InputHandler) = not handler.History.Redos.IsEmpty

/// Create input handler
let create (canvas: HTMLCanvasElement) (initialDoc: Doc.Document) (onChange: Doc.Document -> unit) =
    let ctx = createContext canvas
    let handler = {
        Canvas = canvas
        RenderContext = ctx
        Document = initialDoc
        OnChange = onChange
        IsMouseDown = false
        SelectionAnchor = None
        History = Cmd.empty
        ResizingImage = None
        DraggingImage = None
        HoveredImage = None
        TrackChanges = false
        TrackChangesAuthor = "RedBat Word"
        TrackChangesDate = ""
        PointerStartDocument = None
        InputProxy = None
        Composition = None
        SyncInputProxy = ignore
        Dispose = ignore
        GoalX = None
        Story = None
        HistoryStory = None
        HistoryPlace = None
        OnHistoryPlace = fun _ _ -> false
        OnDoubleClick = fun _ _ -> false
        EnsureWhole = ignore
        Pending = None
        OnPendingFormat = ignore
        PasteText = ignore
        PasteClipboard = fun _ _ -> ()
        CopyToClipboard = fun _ -> None
        SelectedShape = None
        ShapeDrag = None
        OnShapeSelection = ignore
        Typing = None
    }

    // ---- caret movement over the layout that is on screen -------------------
    // Lines are visual lines: a paragraph wraps into several, and Up / Down /
    // Home / End / PageUp / PageDown move by them, keeping the column.

    let measureRange = measureRangeWithContext ctx

    /// The visual lines of the document's editable text, in layout order.
    // (Kept for the layout they were made from: a run of Up / Down keys, which
    // changes the selection only, enumerates the document's lines once.)
    let mutable visualLinesOf : obj = null
    let mutable visualLinesKnown : (Layout.BlockLayout * Layout.MeasuredLine) array = [||]
    let visualLineStarts = System.Collections.Generic.Dictionary<string, int>()
    let visualLines (doc: Doc.Document) : (Layout.BlockLayout * Layout.MeasuredLine) array =
        let layouts = layoutDocument ctx doc
        if not (obj.ReferenceEquals(visualLinesOf, layouts)) then
            let lines = ResizeArray<Layout.BlockLayout * Layout.MeasuredLine>()
            visualLineStarts.Clear()
            for layout in layouts do
                if not (layout.BlockId.StartsWith Layout.footnoteBlockIdPrefix) && not (layout.BlockId.StartsWith Layout.endnoteBlockIdPrefix) then
                    if not layout.Lines.IsEmpty && not (visualLineStarts.ContainsKey layout.BlockId) then visualLineStarts.[layout.BlockId] <- lines.Count
                    for line in layout.Lines do
                        lines.Add (layout, line)
            visualLinesKnown <- lines.ToArray()
            visualLinesOf <- box layouts
        visualLinesKnown

    /// Layout x of an offset on a line.
    let lineX (doc: Doc.Document) (layout: Layout.BlockLayout) (line: Layout.MeasuredLine) (offset: int) =
        (layoutSettings ctx doc).PaddingLeft + Layout.xForOffsetWithRange measureRange line offset + Layout.alignmentOffset layout line

    /// The visual line the position is on (index into `lines`).
    let lineIndexOf (lines: (Layout.BlockLayout * Layout.MeasuredLine) array) (position: Doc.Position) : int option =
        // A paragraph continued in another column or on another page has one layout per fragment.
        let mutable found = -1
        let mutable fallback = -1
        // From the block's first line on, when these are the lines of `visualLines`
        // (a block's lines follow one another: the search ends where they end).
        let indexed = obj.ReferenceEquals(lines, visualLinesKnown)
        let mutable i =
            if indexed then
                match visualLineStarts.TryGetValue position.BlockId with
                | true, start -> start
                | false, _ -> lines.Length
            else 0
        let mutable searching = true
        while searching && i < lines.Length do
            let layout, line = lines.[i]
            if layout.BlockId = position.BlockId then
                if fallback < 0 then fallback <- i
                if line.StartOffset <= position.Offset && position.Offset <= line.EndOffset
                   && obj.ReferenceEquals(Layout.lineForOffset layout position.Offset, line) then
                    found <- i
                    searching <- false
            elif indexed && fallback >= 0 then searching <- false
            i <- i + 1
        if found >= 0 then Some found elif fallback >= 0 then Some fallback else None

    /// The last offset the caret can have and still be drawn on the line: a
    /// wrapped line ends where the next one starts, and that offset belongs
    /// to the next line.
    let lastOffsetOnLine (layout: Layout.BlockLayout) (line: Layout.MeasuredLine) =
        let continues = layout.Lines |> List.exists (fun other -> other.StartOffset = line.EndOffset && not (obj.ReferenceEquals(other, line)))
        if continues && line.EndOffset > line.StartOffset then line.EndOffset - 1 else line.EndOffset

    /// The position on a line nearest to a layout x.
    let positionOnLine (doc: Doc.Document) (layout: Layout.BlockLayout) (line: Layout.MeasuredLine) (x: float) : Doc.Position =
        let offset =
            match Layout.hitTestLayoutsWithRange measureRange (layoutSettings ctx doc).PaddingLeft [ layout ] x (line.Y + line.Height / 2.0) with
            | Some (_, offset) -> offset
            | None -> line.StartOffset
        { BlockId = layout.BlockId; Offset = max line.StartOffset (min (lastOffsetOnLine layout line) offset) }

    /// Horizontal distance from x to a line (0 when x is over it).
    let distanceToLine (doc: Doc.Document) (layout: Layout.BlockLayout) (line: Layout.MeasuredLine) (x: float) =
        let left = lineX doc layout line line.StartOffset
        let right = lineX doc layout line line.EndOffset
        let left, right = min left right, max left right
        if x < left then left - x elif x > right then x - right else 0.0

    /// The line a vertical move from `current` lands on: the nearest line
    /// above / below it and, among lines side by side (table cells), the one
    /// under the goal column.
    let lineInDirection (doc: Doc.Document) (lines: (Layout.BlockLayout * Layout.MeasuredLine) array) (current: int) (down: bool) (goalX: float) : int option =
        let _, from = lines.[current]
        let mutable best = -1
        let mutable bestGap = System.Double.MaxValue
        let mutable bestDistance = System.Double.MaxValue
        for i in 0 .. lines.Length - 1 do
            if i <> current then
                let layout, line = lines.[i]
                let gap = if down then line.Y - (from.Y + from.Height) else from.Y - (line.Y + line.Height)
                // A line counts as above / below when it does not overlap the current one vertically.
                if gap > -0.5 then
                    let distance = distanceToLine doc layout line goalX
                    let nearer = gap < bestGap - 2.0 || (abs (gap - bestGap) <= 2.0 && distance < bestDistance)
                    if best < 0 || nearer then
                        best <- i
                        bestGap <- min gap bestGap
                        bestDistance <- distance
        if best >= 0 then Some best else None

    /// The line at (or nearest to) a layout y, under the goal column.
    let lineAtY (doc: Doc.Document) (lines: (Layout.BlockLayout * Layout.MeasuredLine) array) (y: float) (goalX: float) : int option =
        let mutable best = -1
        let mutable bestGap = System.Double.MaxValue
        let mutable bestDistance = System.Double.MaxValue
        for i in 0 .. lines.Length - 1 do
            let layout, line = lines.[i]
            let gap = if y < line.Y then line.Y - y elif y > line.Y + line.Height then y - (line.Y + line.Height) else 0.0
            let distance = distanceToLine doc layout line goalX
            if best < 0 || gap < bestGap - 0.5 || (abs (gap - bestGap) <= 0.5 && distance < bestDistance) then
                best <- i
                bestGap <- gap
                bestDistance <- distance
        if best >= 0 then Some best else None

    /// The caret (the selection's focus) and the layout x it is drawn at.
    let caretLine (doc: Doc.Document) =
        match doc.Selection with
        | None -> None
        | Some selection ->
            let lines = visualLines doc
            lineIndexOf lines selection.Focus
            |> Option.map (fun index ->
                let layout, line = lines.[index]
                lines, index, lineX doc layout line (max line.StartOffset (min line.EndOffset selection.Focus.Offset)))

    /// Put the focus at `position`: the whole selection with it, or with
    /// `extend` (Shift) only the focus, the anchor staying where it is.
    let withFocus (doc: Doc.Document) (extend: bool) (position: Doc.Position) : Doc.Document =
        match doc.Selection with
        | Some selection when extend -> { doc with Selection = Some { selection with Focus = position } }
        | _ -> { doc with Selection = Some { Anchor = position; Focus = position } }

    /// Up / Down by `count` visual lines, keeping the column.
    let moveVertically (doc: Doc.Document) (down: bool) (extend: bool) : Doc.Document =
        match caretLine doc with
        | None -> doc
        | Some (lines, index, x) ->
            let goalX = handler.GoalX |> Option.defaultValue x
            handler.GoalX <- Some goalX
            match lineInDirection doc lines index down goalX with
            | Some target ->
                let layout, line = lines.[target]
                withFocus doc extend (positionOnLine doc layout line goalX)
            | None ->
                // No line further on: the start / end of the text, as Word does.
                let layout, line = lines.[index]
                withFocus doc extend { BlockId = layout.BlockId; Offset = (if down then line.EndOffset else line.StartOffset) }

    /// PageUp / PageDown: a screenful (the visible height) up or down, keeping the column.
    let movePage (doc: Doc.Document) (down: bool) (extend: bool) : Doc.Document =
        match caretLine doc with
        | None -> doc
        | Some (lines, index, x) ->
            let goalX = handler.GoalX |> Option.defaultValue x
            handler.GoalX <- Some goalX
            let page =
                match ctx.Viewport with
                | Some vp -> max 40.0 (vp.Height / vp.Scale - 40.0)
                | None -> max 40.0 ((layoutSettings ctx doc).PageHeight - 80.0)
            let _, from = lines.[index]
            let targetY = from.Y + from.Height / 2.0 + (if down then page else -page)
            match lineAtY doc lines targetY goalX with
            | Some target when target <> index ->
                let layout, line = lines.[target]
                withFocus doc extend (positionOnLine doc layout line goalX)
            | _ ->
                let layout, line = lines.[index]
                withFocus doc extend { BlockId = layout.BlockId; Offset = (if down then line.EndOffset else line.StartOffset) }

    /// Home / End: the start / end of the visual line.
    let moveToLineEdge (doc: Doc.Document) (toEnd: bool) (extend: bool) : Doc.Document =
        match doc.Selection with
        | None -> doc
        | Some selection ->
            let lines = visualLines doc
            match lineIndexOf lines selection.Focus with
            | Some index ->
                let layout, line = lines.[index]
                withFocus doc extend { BlockId = layout.BlockId; Offset = (if toEnd then lastOffsetOnLine layout line else line.StartOffset) }
            | None ->
                // Text that has no layout (yet): the paragraph's ends.
                let moved = if toEnd then Doc.moveCursorToLineEnd doc else Doc.moveCursorToLineStart doc
                match moved.Selection with
                | Some target -> withFocus doc extend target.Focus
                | None -> doc

    /// Ctrl+Home / Ctrl+End: the start / end of the document.
    let moveToDocumentEdge (doc: Doc.Document) (toEnd: bool) (extend: bool) : Doc.Document =
        match (Doc.selectAll doc).Selection with
        | Some whole -> withFocus doc extend (if toEnd then whole.Focus else whole.Anchor)
        | None -> doc

    /// The start of the word at or after `offset` going right, or before it going left.
    let wordBoundary (text: string) (offset: int) (forward: bool) : int =
        let mutable i = max 0 (min text.Length offset)
        if forward then
            // Over the rest of this word (or run of punctuation), then the space after it.
            if i < text.Length then
                if isWordChar text.[i] then
                    while i < text.Length && isWordChar text.[i] do i <- i + 1
                elif not (System.Char.IsWhiteSpace text.[i]) then
                    while i < text.Length && not (isWordChar text.[i]) && not (System.Char.IsWhiteSpace text.[i]) do i <- i + 1
            while i < text.Length && System.Char.IsWhiteSpace text.[i] do i <- i + 1
        else
            while i > 0 && System.Char.IsWhiteSpace text.[i - 1] do i <- i - 1
            if i > 0 then
                if isWordChar text.[i - 1] then
                    while i > 0 && isWordChar text.[i - 1] do i <- i - 1
                else
                    while i > 0 && not (isWordChar text.[i - 1]) && not (System.Char.IsWhiteSpace text.[i - 1]) do i <- i - 1
        i

    /// Ctrl+Left / Ctrl+Right: by words, across paragraph boundaries.
    let moveByWord (doc: Doc.Document) (forward: bool) (extend: bool) : Doc.Document =
        match doc.Selection with
        | None -> doc
        | Some selection ->
            match Doc.findParagraph doc selection.Focus.BlockId with
            | None -> doc
            | Some para ->
                let text = TextBuffer.toString para.Buffer
                let atEdge = if forward then selection.Focus.Offset >= text.Length else selection.Focus.Offset <= 0
                if atEdge then
                    // Into the next / previous paragraph, as a plain arrow does.
                    let collapsed = { doc with Selection = Some { Anchor = selection.Focus; Focus = selection.Focus } }
                    let moved = if forward then Doc.moveCursorRight collapsed else Doc.moveCursorLeft collapsed
                    match moved.Selection with
                    | Some target -> withFocus doc extend target.Focus
                    | None -> doc
                else
                    withFocus doc extend { selection.Focus with Offset = wordBoundary text selection.Focus.Offset forward }

    /// Ctrl+Up / Ctrl+Down: the start of this paragraph (then the previous one) / of the next one.
    let moveByParagraph (doc: Doc.Document) (forward: bool) (extend: bool) : Doc.Document =
        match doc.Selection with
        | None -> doc
        | Some selection ->
            if not forward && selection.Focus.Offset > 0 then withFocus doc extend { selection.Focus with Offset = 0 }
            else
                let collapsed =
                    match Doc.findParagraph doc selection.Focus.BlockId with
                    | Some para when forward -> { doc with Selection = Some { Anchor = { selection.Focus with Offset = TextBuffer.length para.Buffer }; Focus = { selection.Focus with Offset = TextBuffer.length para.Buffer } } }
                    | _ -> { doc with Selection = Some { Anchor = selection.Focus; Focus = selection.Focus } }
                let moved = if forward then Doc.moveCursorRight collapsed else Doc.moveCursorLeft collapsed
                match moved.Selection with
                | Some target when target.Focus.BlockId <> selection.Focus.BlockId -> withFocus doc extend { target.Focus with Offset = 0 }
                | Some target when forward -> withFocus doc extend target.Focus
                | _ -> doc

    /// The word at a position (a double click selects it).
    let wordAt (doc: Doc.Document) (position: Doc.Position) : Doc.Document =
        match Doc.findParagraph doc position.BlockId with
        | None -> doc
        | Some para ->
            let text = TextBuffer.toString para.Buffer
            let at = max 0 (min text.Length position.Offset)
            // The character clicked on: the one after the caret position, or the last one.
            let probe = if at < text.Length then at elif at > 0 then at - 1 else 0
            if text.Length = 0 then Doc.setSelection position position doc
            else
                let sameKind (a: char) (b: char) =
                    (isWordChar a && isWordChar b)
                    || (System.Char.IsWhiteSpace a && System.Char.IsWhiteSpace b)
                    || (not (isWordChar a) && not (isWordChar b) && not (System.Char.IsWhiteSpace a) && not (System.Char.IsWhiteSpace b))
                let kind = text.[probe]
                let mutable first = probe
                while first > 0 && sameKind text.[first - 1] kind do first <- first - 1
                let mutable last = probe + 1
                while last < text.Length && sameKind text.[last] kind do last <- last + 1
                Doc.setSelection { position with Offset = first } { position with Offset = last } doc

    /// The whole paragraph at a position (a triple click selects it).
    let paragraphAt (doc: Doc.Document) (position: Doc.Position) : Doc.Document =
        match Doc.findParagraph doc position.BlockId with
        | None -> doc
        | Some para -> Doc.setSelection { position with Offset = 0 } { position with Offset = TextBuffer.length para.Buffer } doc

    let insertTextForMode text doc =
        withPendingFormat handler doc (if handler.TrackChanges then Doc.insertTrackedText handler.TrackChangesAuthor handler.TrackChangesDate text doc else Doc.insertText text doc)

    let splitBlockForMode doc =
        if handler.TrackChanges then Doc.trackSplitBlock handler.TrackChangesAuthor handler.TrackChangesDate doc else Doc.splitBlock doc

    let deleteSelectionForMode doc =
        if handler.TrackChanges then Doc.trackDeleteSelection handler.TrackChangesAuthor handler.TrackChangesDate doc else Doc.deleteSelection doc

    /// Plain text put in at the caret of `beforeDoc`, or over its selection:
    /// the lines of the text are paragraphs. With nothing selected the text
    /// has the format chosen at the caret for what is typed next, if one was -
    /// all of it, the lines after the first too.
    let pastedText (text: string) (beforeDoc: Doc.Document) : Doc.Document =
        let mutable doc' = beforeDoc
        // Delete selection first if any
        let hasSelection =
            match doc'.Selection with
            | Some sel when sel.Anchor <> sel.Focus -> true
            | _ -> false
        let pending = if hasSelection then None else pendingFormat handler
        if hasSelection then
            doc' <- deleteSelectionForMode doc'

        let mutable segmentStart = 0
        let mutable index = 0
        let mutable sawLineBreak = false
        let insertSegment segmentEnd =
            doc' <- insertTextForMode (text.Substring(segmentStart, segmentEnd - segmentStart)) doc'

        while index < text.Length do
            let ch = text.[index]
            if ch = '\r' || ch = '\n' then
                sawLineBreak <- true
                insertSegment index
                doc' <- splitBlockForMode doc'
                if ch = '\r' && index + 1 < text.Length && text.[index + 1] = '\n' then
                    index <- index + 1
                segmentStart <- index + 1
            index <- index + 1

        if sawLineBreak then
            insertSegment text.Length
        else
            doc' <- insertTextForMode text doc'
        // (The first line has the pending format from insertTextForMode; the paragraphs after it are given it here.)
        match pending, beforeDoc.Selection, doc'.Selection with
        | Some pending, Some was, Some now when was.IsCollapsed && now.IsCollapsed && now.Focus.BlockId <> was.Focus.BlockId ->
            let styleOf (typed: TextBuffer.TextStyle) = pendingStyleOver pending.PendingStyle typed
            let ranged = Doc.setSelection was.Focus now.Focus doc'
            let styled = Doc.applyFormattingToSelection (fun start len buffer -> TextBuffer.applyStyle start len styleOf buffer) ranged
            { styled with Selection = doc'.Selection }
        | _ -> doc'

    let pasteText (text: string) =
        if not (isNull text) && text <> "" then
            let beforeDoc = handler.Document
            let doc' = pastedText text beforeDoc
            if not (documentsEqualForChange doc' beforeDoc) then
                let cmd = DeleteTextCommand(beforeDoc, doc', "Paste") :> ICommand
                let doc'', hist' = Cmd.execute cmd beforeDoc handler.History
                handler.Document <- doc''
                handler.History <- storyHistory handler handler.History hist'
                renderWithHover handler.RenderContext doc'' handler.HoveredImage
                handler.OnChange doc''
    handler.PasteText <- pasteText

    /// An edit of the clipboard (a paste, a cut) as one undo step, shown.
    let commitClipboardEdit (description: string) (afterDoc: Doc.Document) =
        let beforeDoc = handler.Document
        if not (documentsEqualForChange afterDoc beforeDoc) then
            let cmd = DeleteTextCommand(beforeDoc, afterDoc, description) :> ICommand
            let doc', history = Cmd.execute cmd beforeDoc handler.History
            dropPendingFormat handler doc'
            handler.Document <- doc'
            handler.History <- storyHistory handler handler.History history
            renderWithHover handler.RenderContext doc' handler.HoveredImage
            handler.SyncInputProxy doc'
            handler.OnChange doc'

    /// Copy (cut: and Cut): the selection kept as the clip to paste, and its
    /// text and HTML for the system clipboard. None with nothing selected.
    let copyToClipboard (cut: bool) : (string * string) option =
        match Doc.copySelection handler.Document with
        | None -> None
        | Some clip ->
            lastClip <- Some clip
            if cut then commitClipboardEdit "Cut" (deleteSelectionForMode handler.Document)
            Some (clip.Text, Doc.clipToHtml clip)

    /// Paste the clipboard's `text` and `html` (either may be empty): the
    /// clip Copy kept while the text is still its text, with its formatting;
    /// else the HTML's text with its formatting; else the text (pasteText).
    /// With changes tracked: the text.
    let pasteClipboard (text: string) (html: string) =
        let text = if isNull text then "" else text
        let normalize (value: string) = value.Replace("\r\n", "\n").Replace("\r", "\n")
        let clip =
            if handler.TrackChanges then None
            else
                match lastClip with
                | Some clip when text <> "" && normalize text = normalize clip.Text -> Some clip
                | _ -> htmlClip html
        match clip with
        | Some clip ->
            // (A text box holds paragraphs only: a table's rows are pasted as paragraphs.)
            let clip = match handler.Story with Some session when session.Box.IsSome -> Doc.clipWithoutTables clip | _ -> clip
            commitClipboardEdit "Paste" (Doc.pasteClip clip handler.Document)
        | None -> pasteText text
    handler.PasteClipboard <- pasteClipboard
    handler.CopyToClipboard <- copyToClipboard

    let deleteBackwardForMode doc =
        if handler.TrackChanges then Doc.trackDeleteBackward handler.TrackChangesAuthor handler.TrackChangesDate doc else Doc.deleteBackward doc

    let deleteForwardForMode doc =
        if handler.TrackChanges then Doc.trackDeleteForward handler.TrackChangesAuthor handler.TrackChangesDate doc else Doc.deleteForward doc

    /// Ctrl+Backspace / Ctrl+Delete: the word before / after the caret, as far
    /// as Ctrl+Left / Ctrl+Right goes from it (at the edge of a paragraph: the
    /// paragraph mark, the paragraphs are joined); a selection is deleted as
    /// by Backspace / Delete.
    let deleteWordForMode (forward: bool) (doc: Doc.Document) =
        match doc.Selection with
        | Some selection when selection.IsCollapsed ->
            let ranged = moveByWord doc forward true
            match ranged.Selection with
            | Some word when not word.IsCollapsed -> deleteSelectionForMode ranged
            | _ -> doc
        | Some _ -> if forward then deleteForwardForMode doc else deleteBackwardForMode doc
        | None -> doc

    /// Layout coordinates of a pointer event: through the viewport transform
    /// when the renderer has one, one to one with canvas pixels otherwise.
    let pointerPosition (me: MouseEvent) =
        let rect = canvas.getBoundingClientRect()
        viewportToLayout handler.RenderContext (me.clientX - rect.left) (me.clientY - rect.top)

    let isResizeHandle x y imageX imageY width height =
        // Ten pixels on screen, whatever the zoom.
        let handle = 10.0 / viewportScale handler.RenderContext
        let nearLeft = abs (x - imageX) <= handle
        let nearRight = abs (x - (imageX + width)) <= handle
        let nearTop = abs (y - imageY) <= handle
        let nearBottom = abs (y - (imageY + height)) <= handle
        (nearLeft || nearRight) && (nearTop || nearBottom)

    let applyPointerDocumentChange newDoc =
        // (A click or a drag into the mark of a comment: to one of its ends.)
        let newDoc = Doc.snapSelectionToMarks handler.Document.Selection newDoc
        if not (documentsEqualForChange newDoc handler.Document) then
            dropPendingFormat handler newDoc
            handler.Document <- newDoc
            renderWithHover handler.RenderContext newDoc handler.HoveredImage
            handler.SyncInputProxy newDoc
            handler.OnChange newDoc

    // Make canvas focusable
    canvas.tabIndex <- 0

    /// The shape the keys act on - Delete, the arrows: the one that is
    /// selected, while no text of a box, a header or a footer is edited (the
    /// keys are that text's then) - with where it is painted: its left edge,
    /// its top, the top of its page.
    let selectedShapeForKeys () : (Doc.TextBoxBlock * float * float * float) option =
        match handler.SelectedShape, handler.Story with
        | Some id, None ->
            shapeRectOf handler.RenderContext handler.Document id
            |> Option.map (fun (shape, x, y, _, _, pageTop) -> shape, x, y, pageTop)
        | _ -> None

    /// A character format given by a key (Ctrl+B and the like): the
    /// selection's, or with nothing selected the caret's (see formatAtCaret).
    /// The document and the history with it, when the document changed.
    let characterFormatStep (description: string) (command: Doc.Document -> Doc.Document) : (Doc.Document * CommandHistory) option =
        let beforeDoc = handler.Document
        let afterDoc = characterFormatted handler command
        if documentsEqualForChange afterDoc beforeDoc then None
        else Some (Cmd.execute (FormatTextCommand(beforeDoc, afterDoc, description) :> ICommand) beforeDoc handler.History)

    // Handle keyboard input. The listener serves the canvas and the input
    // proxy (which normally has the focus).
    let onKeyDown (e: Event) =
        let ke = e :?> KeyboardEvent
        // (Everything, or the end of the document: of the whole document.)
        if (ke.ctrlKey || ke.metaKey) && (keyEquals 'a' ke.key || ke.key = "End") then handler.EnsureWhole ()
        let mutable newDoc = handler.Document
        let mutable newHistory = handler.History
        let mutable useCommand = false
        // An undo / redo step, or a keystroke of a run of typing: the history
        // in newHistory is the one to keep as it is (storyHistory is not applied).
        let mutable stepped = false
        // Up / Down / PageUp / PageDown keep their column from one to the next.
        let goalBefore = handler.GoalX
        let mutable keepGoal = false
        let shift = ke.shiftKey

        let commandKey = ke.ctrlKey || ke.metaKey
        // While an IME or a dead key is composing, keys belong to it (Enter
        // confirms a candidate, arrows move within it, ...); the text arrives
        // through the composition events.
        let composing = handler.Composition.IsSome || keyBelongsToComposition ke

        match (if composing then "" else ke.key) with
        | key when commandKey && keyEquals 'z' key ->
            // Undo/Redo
            if ke.shiftKey then
                // Redo (Ctrl+Shift+Z)
                let (doc', hist') = historyStep handler Cmd.redo
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
                stepped <- true
            else
                // Undo (Ctrl+Z)
                let (doc', hist') = historyStep handler Cmd.undo
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
                stepped <- true
            e.preventDefault()
        | key when commandKey && keyEquals 'y' key ->
            // Redo (Ctrl+Y)
            let (doc', hist') = historyStep handler Cmd.redo
            newDoc <- doc'
            newHistory <- hist'
            useCommand <- true
            stepped <- true
            e.preventDefault()
        | "Escape" when handler.Story.IsSome ->
            // Out of the header or footer, back to the body.
            exitStory handler |> ignore
            newDoc <- handler.Document
            e.preventDefault()
        | "Escape" when handler.SelectedShape.IsSome ->
            // (The shape that is selected is let go of.)
            selectShape handler None
            e.preventDefault()
        | "Delete"
        | "Backspace" when (selectedShapeForKeys ()).IsSome ->
            // The shape that is selected is taken out, its text with it.
            match selectedShapeForKeys () with
            | Some (shape, _, _, _) ->
                let beforeDoc = handler.Document
                match Doc.deleteTextBox shape.Id beforeDoc with
                | Some afterDoc ->
                    let cmd = DeleteTextCommand(beforeDoc, afterDoc, "Delete shape") :> ICommand
                    let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                    newDoc <- doc'
                    newHistory <- hist'
                    useCommand <- true
                    if setSelectedShape handler None then handler.OnShapeSelection ()
                | None -> ()
            | None -> ()
            e.preventDefault()
        | "ArrowLeft"
        | "ArrowRight"
        | "ArrowUp"
        | "ArrowDown" when not commandKey && not ke.altKey && (selectedShapeForKeys ()).IsSome ->
            // The shape that is selected moves by a pixel (with Shift: ten).
            match selectedShapeForKeys () with
            | Some (shape, rectX, rectY, pageTop) ->
                let step = if shift then 10.0 else 1.0
                let dx = (match ke.key with "ArrowLeft" -> -step | "ArrowRight" -> step | _ -> 0.0)
                let dy = (match ke.key with "ArrowUp" -> -step | "ArrowDown" -> step | _ -> 0.0)
                let beforeDoc = handler.Document
                match shapePlacedDocument beforeDoc shape rectX rectY pageTop dx dy None with
                | Some afterDoc when not (documentsEqualForChange afterDoc beforeDoc) ->
                    let cmd = DeleteTextCommand(beforeDoc, afterDoc, "Move shape") :> ICommand
                    let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                    newDoc <- doc'
                    newHistory <- hist'
                    useCommand <- true
                | _ -> ()
            | None -> ()
            e.preventDefault()
        | key when commandKey && keyEquals 'a' key ->
            newDoc <- Doc.selectAll handler.Document
            e.preventDefault()
        | key when commandKey && keyEquals 'b' key ->
            let beforeDoc = handler.Document
            // (With nothing selected: the word the caret is in, or what is typed next.)
            let afterDoc = characterFormatted handler WordStyles.toggleBold
            if not (documentsEqualForChange afterDoc beforeDoc) then
                let cmd = FormatTextCommand(beforeDoc, afterDoc, "bold") :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | key when commandKey && keyEquals 'i' key ->
            let beforeDoc = handler.Document
            // (With nothing selected: the word the caret is in, or what is typed next.)
            let afterDoc = characterFormatted handler WordStyles.toggleItalic
            if not (documentsEqualForChange afterDoc beforeDoc) then
                let cmd = FormatTextCommand(beforeDoc, afterDoc, "italic") :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | key when commandKey && keyEquals 'u' key ->
            let beforeDoc = handler.Document
            // (With nothing selected: the word the caret is in, or what is typed next.)
            let afterDoc = characterFormatted handler WordStyles.toggleUnderline
            if not (documentsEqualForChange afterDoc beforeDoc) then
                let cmd = FormatTextCommand(beforeDoc, afterDoc, "underline") :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | "1" when commandKey ->
            newDoc <- Doc.setHeadingLevel 1 handler.Document
            e.preventDefault()
        | "2" when commandKey ->
            newDoc <- Doc.setHeadingLevel 2 handler.Document
            e.preventDefault()
        | "3" when commandKey ->
            newDoc <- Doc.setHeadingLevel 3 handler.Document
            e.preventDefault()
        | "F3" when shift && not commandKey && not ke.altKey ->
            // Shift+F3: Change Case of the selection or the word at the caret.
            newDoc <- Doc.changeCase handler.Document
            e.preventDefault()
        | key when commandKey && shift && not ke.altKey && keyEquals 'l' key ->
            // Ctrl+Shift+L: a bulleted list (Word's List Bullet).
            newDoc <- Doc.applyBulletList handler.Document
            e.preventDefault()
        | key when commandKey && shift && not ke.altKey && keyEquals 'n' key ->
            // Ctrl+Shift+N: the Normal style.
            newDoc <- Doc.setNormalStyle handler.Document
            e.preventDefault()
        | key when commandKey && not ke.altKey && not shift && (keyEquals 'e' key || keyEquals 'l' key || keyEquals 'r' key || keyEquals 'j' key) ->
            // Word's alignment keys: Ctrl+E centre, Ctrl+L left, Ctrl+R right, Ctrl+J justify.
            let align =
                if keyEquals 'e' key then Doc.alignCenter
                elif keyEquals 'r' key then Doc.alignRight
                elif keyEquals 'j' key then Doc.alignJustify
                else Doc.alignLeft
            newDoc <- align handler.Document
            e.preventDefault()
        | key when commandKey && not ke.altKey && keyEquals 'm' key ->
            // Ctrl+M / Ctrl+Shift+M: Increase / Decrease Indent.
            newDoc <- Doc.indentParagraphs (not shift) handler.Document
            e.preventDefault()
        | key when commandKey && not ke.altKey && (key = "]" || key = "[" || key = ">" || key = "<" || key = "=" || key = "+" || key = " " || (shift && (key = "." || key = ","))) ->
            // Word's character keys: Ctrl+] / Ctrl+[ a point larger / smaller,
            // Ctrl+Shift+> / < the next size up / down, Ctrl+= subscript,
            // Ctrl+Shift+= superscript, Ctrl+Space the character formatting cleared.
            let description, command =
                match key with
                | "]" -> "Grow font", WordStyles.growFont true false
                | "[" -> "Shrink font", WordStyles.growFont false false
                | ">" | "." -> "Grow font", WordStyles.growFont true true
                | "<" | "," -> "Shrink font", WordStyles.growFont false true
                | "+" -> "Superscript", WordStyles.toggleSuperscript
                | "=" when shift -> "Superscript", WordStyles.toggleSuperscript
                | "=" -> "Subscript", WordStyles.toggleSubscript
                | _ -> "Clear formatting", Doc.clearFormatting
            match characterFormatStep description command with
            | Some (doc', hist') ->
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            | None -> ()
            e.preventDefault()
        | key when commandKey && not ke.altKey && (keyEquals 'c' key || keyEquals 'x' key || keyEquals 'v' key) ->
            // Copy, Cut, Paste. On the input proxy the browser's own copy / cut /
            // paste event follows the key (see the listeners below): the key is
            // left to it. For a browser that copies only what is selected, the
            // selection's text is put in the proxy, selected, until then.
            let fromProxy = match handler.InputProxy with Some element -> obj.ReferenceEquals(e.target, element) | None -> false
            if fromProxy then
                match handler.InputProxy with
                | Some element when not (keyEquals 'v' key) ->
                    let text = Doc.getSelectionText handler.Document
                    if text <> "" then
                        setProxyValue element text
                        selectProxyText element
                        soon (fun () -> if proxyValue element = text then setProxyValue element "")
                | _ -> ()
            else
                // (The key reached the canvas itself: through the asynchronous clipboard.)
                e.preventDefault()
                if keyEquals 'v' key then
                    Clipboard.pasteText().``then``(fun text -> pasteClipboard text "") |> ignore
                else
                    let copied = copyToClipboard (keyEquals 'x' key)
                    // (A cut has changed the document already.)
                    newDoc <- handler.Document
                    copied |> Option.iter (fun (text, _) -> Clipboard.copyText text |> ignore)
        | "Backspace" ->
            let beforeDoc = handler.Document
            // (Ctrl+Backspace, Option+Backspace on a Mac: the word before the caret.)
            let byWord = ke.ctrlKey || ke.altKey
            let afterDoc = if byWord then deleteWordForMode false handler.Document else deleteBackwardForMode handler.Document
            if not (documentsEqualForChange afterDoc beforeDoc) then
                if byWord then
                    let cmd = DeleteTextCommand(beforeDoc, afterDoc, "Delete backward") :> ICommand
                    let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                    newDoc <- doc'
                    newHistory <- hist'
                else
                    // (Backspaces one after the other are one undo step.)
                    newDoc <- afterDoc
                    newHistory <- typingHistory handler "backspace" "Delete backward" beforeDoc afterDoc
                    stepped <- true
                useCommand <- true
            e.preventDefault()
        | "Delete" ->
            let beforeDoc = handler.Document
            // (Ctrl+Delete: the word after the caret.)
            let byWord = ke.ctrlKey || ke.altKey
            let afterDoc = if byWord then deleteWordForMode true handler.Document else deleteForwardForMode handler.Document
            if not (documentsEqualForChange afterDoc beforeDoc) then
                if byWord then
                    let cmd = DeleteTextCommand(beforeDoc, afterDoc, "Delete forward") :> ICommand
                    let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                    newDoc <- doc'
                    newHistory <- hist'
                else
                    newDoc <- afterDoc
                    newHistory <- typingHistory handler "delete" "Delete forward" beforeDoc afterDoc
                    stepped <- true
                useCommand <- true
            e.preventDefault()
        | "ArrowLeft" ->
            if ke.metaKey then newDoc <- moveToLineEdge handler.Document false shift
            elif ke.ctrlKey || ke.altKey then newDoc <- moveByWord handler.Document false shift
            elif shift then newDoc <- Doc.extendSelectionLeft handler.Document
            else newDoc <- Doc.moveCursorLeft handler.Document
            e.preventDefault()
        | "ArrowRight" ->
            if ke.metaKey then newDoc <- moveToLineEdge handler.Document true shift
            elif ke.ctrlKey || ke.altKey then newDoc <- moveByWord handler.Document true shift
            elif shift then newDoc <- Doc.extendSelectionRight handler.Document
            else newDoc <- Doc.moveCursorRight handler.Document
            e.preventDefault()
        | "ArrowUp" ->
            if ke.metaKey then newDoc <- moveToDocumentEdge handler.Document false shift
            elif ke.ctrlKey || ke.altKey then newDoc <- moveByParagraph handler.Document false shift
            else
                newDoc <- moveVertically handler.Document false shift
                keepGoal <- true
            e.preventDefault()
        | "ArrowDown" ->
            if ke.metaKey then newDoc <- moveToDocumentEdge handler.Document true shift
            elif ke.ctrlKey || ke.altKey then newDoc <- moveByParagraph handler.Document true shift
            else
                newDoc <- moveVertically handler.Document true shift
                keepGoal <- true
            e.preventDefault()
        | "PageUp" ->
            newDoc <- movePage handler.Document false shift
            keepGoal <- true
            e.preventDefault()
        | "PageDown" ->
            newDoc <- movePage handler.Document true shift
            keepGoal <- true
            e.preventDefault()
        | "Home" ->
            newDoc <- if commandKey then moveToDocumentEdge handler.Document false shift else moveToLineEdge handler.Document false shift
            e.preventDefault()
        | "End" ->
            newDoc <- if commandKey then moveToDocumentEdge handler.Document true shift else moveToLineEdge handler.Document true shift
            e.preventDefault()
        | "Tab" ->
            // In a table Tab and Shift+Tab go from cell to cell (Tab in the
            // last cell adds a row), as in Word; in a list they move the item
            // a level down or up; elsewhere Tab is a tab character (in a
            // table: Ctrl+Tab).
            let typeTab () =
                let beforeDoc = handler.Document
                let afterDoc = insertTextForMode "\t" handler.Document
                if not (documentsEqualForChange afterDoc beforeDoc) then
                    let cmd = DeleteTextCommand(beforeDoc, afterDoc, "Insert tab") :> ICommand
                    let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                    newDoc <- doc'
                    newHistory <- hist'
                    useCommand <- true
            let context = Doc.caretContext handler.Document
            if ke.ctrlKey && not ke.shiftKey && context.InTable then typeTab ()
            else
            match Doc.moveToAdjacentTableCell (not ke.shiftKey) handler.Document with
            | Some moved -> newDoc <- moved
            | None ->
                if ke.shiftKey then
                    // Shift+Tab: Decrease indent (in the first cell of a table: nothing)
                    if not context.InTable then newDoc <- Doc.decreaseIndent handler.Document
                elif context.InList then
                    // Tab: Increase indent
                    newDoc <- Doc.increaseIndent handler.Document
                elif not context.InTable then typeTab ()
            e.preventDefault()
        | "Enter" when commandKey ->
            // Ctrl+Enter: a page break at the caret; Ctrl+Shift+Enter: a column
            // break. (Where a page cannot break - a text box, a header, a
            // note - nothing, as insertPageBreak.)
            if pageBreakAllowed handler then
                let beforeDoc = handler.Document
                let afterDoc = if shift then Doc.insertColumnBreak beforeDoc else Doc.insertPageBreak beforeDoc
                if not (documentsEqualForChange afterDoc beforeDoc) then
                    let cmd = SplitBlockCommand(beforeDoc, afterDoc) :> ICommand
                    let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                    newDoc <- doc'
                    newHistory <- hist'
                    useCommand <- true
            e.preventDefault()
        | "Enter" when shift ->
            // Shift+Enter: a line break in the paragraph (<w:br/>).
            let beforeDoc = handler.Document
            let afterDoc = insertTextForMode "\n" handler.Document
            if not (documentsEqualForChange afterDoc beforeDoc) then
                let cmd = DeleteTextCommand(beforeDoc, afterDoc, "Insert line break") :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | "Enter" ->
            // Check if we're on an empty list item - if so, exit list
            match handler.Document.Selection with
            | Some sel ->
                match Doc.findParagraph handler.Document sel.Focus.BlockId with
                | Some para ->
                    let isEmpty = TextBuffer.length para.Buffer = 0
                    let inList = para.Metadata.ListStyle <> NoList

                    if isEmpty && inList then
                        // Exit list instead of splitting
                        newDoc <- Doc.decreaseIndent handler.Document
                        e.preventDefault()
                    else
                        // Normal split
                        let beforeDoc = handler.Document
                        let afterDoc = splitBlockForMode handler.Document
                        if not (documentsEqualForChange afterDoc beforeDoc) then
                            let cmd = SplitBlockCommand(beforeDoc, afterDoc) :> ICommand
                            let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                            newDoc <- doc'
                            newHistory <- hist'
                            useCommand <- true
                        e.preventDefault()
                | None ->
                    e.preventDefault()
            | None ->
                e.preventDefault()
        | key when isSingleCodePoint key && not ke.ctrlKey && not ke.metaKey ->
            // Regular character - use command
            let beforeDoc = handler.Document
            let afterDoc = insertTextForMode key handler.Document
            if not (documentsEqualForChange afterDoc beforeDoc) then
                // (Characters typed one after the other are one undo step, as in Word.)
                newDoc <- afterDoc
                newHistory <- typingHistory handler "type" "Typing" beforeDoc afterDoc
                stepped <- true
                useCommand <- true
            e.preventDefault()
        | _ -> ()

        // Any other key ends a run of vertical moves: the next one starts from where the caret then is.
        if not keepGoal && not composing && ke.key <> "Shift" && ke.key <> "Control" && ke.key <> "Alt" && ke.key <> "Meta" then handler.GoalX <- None
        ignore goalBefore
        if useCommand then
            // Also when the document compares equal (undoing a no-op step
            // must still pop it).
            handler.History <- (if stepped then newHistory else storyHistory handler handler.History newHistory)
        // (A caret moved into the mark of a comment goes on to its other end: see Doc.snapSelectionToMarks.)
        if not useCommand then newDoc <- Doc.snapSelectionToMarks handler.Document.Selection newDoc
        if not (documentsEqualForChange newDoc handler.Document) then
            if not useCommand then
                // Shortcuts that changed the document without building a
                // command (heading level, Tab indent, leaving a list) are
                // undo steps too.
                recordChange handler "Edit" handler.Document newDoc
            dropPendingFormat handler newDoc
            handler.Document <- newDoc
            renderWithHover handler.RenderContext newDoc handler.HoveredImage
            handler.SyncInputProxy newDoc
            handler.OnChange newDoc

    canvas.addEventListener("keydown", onKeyDown)

    // ---- text input proxy: IME / dead keys / virtual keyboards ----------------
    // A canvas receives no composition or input events. A hidden textarea next
    // to it takes the keyboard focus instead; it sits at the caret so the IME
    // candidate window opens in the right place.
    let proxy = document.createElement "textarea"
    proxy.setAttribute("data-rbt-word-input", "true")
    proxy.setAttribute("aria-label", "Document text input")
    proxy.setAttribute("autocomplete", "off")
    proxy.setAttribute("autocorrect", "off")
    proxy.setAttribute("autocapitalize", "off")
    proxy.setAttribute("spellcheck", "false")
    proxy.setAttribute("tabindex", "-1")
    proxy.setAttribute(
        "style",
        "position:fixed;left:0;top:0;width:2px;height:16px;padding:0;margin:0;border:0;outline:none;resize:none;overflow:hidden;"
        + "opacity:0;pointer-events:none;white-space:pre;background:transparent;color:transparent;caret-color:transparent;")
    if isNull canvas.parentElement then document.body.appendChild proxy |> ignore
    else canvas.parentElement.insertBefore(proxy, canvas.nextSibling) |> ignore
    handler.InputProxy <- Some proxy
    let mutable proxyAttached = true

    let syncInputProxy (doc: Doc.Document) =
        if proxyAttached then
            match caretRect handler.RenderContext doc with
            | Some (x, top, height) ->
                // The same mapping as mouse handling, the other way round: one
                // to one in whole-canvas mode, through the viewport otherwise.
                let rect = canvas.getBoundingClientRect()
                let scale = viewportScale handler.RenderContext
                let viewX, viewY = layoutToViewport handler.RenderContext x top
                let height = height * scale
                // A caret scrolled out of view keeps its proxy inside the canvas
                // box, so focusing it never scrolls the page.
                let viewX, viewY =
                    match handler.RenderContext.Viewport with
                    | Some vp -> max 0.0 (min (vp.Width - 2.0) viewX), max 0.0 (min (max 0.0 (vp.Height - height)) viewY)
                    | None -> viewX, viewY
                proxy?style?left <- sprintf "%fpx" (rect.left + viewX)
                proxy?style?top <- sprintf "%fpx" (rect.top + viewY)
                proxy?style?height <- sprintf "%fpx" (max 8.0 height)
                proxy?style?fontSize <- sprintf "%fpx" (max 8.0 (height * 0.75))
            | None -> ()
    handler.SyncInputProxy <- syncInputProxy

    /// The document with the composing text shown underlined at the caret.
    let compositionPreview (text: string) =
        if text = "" then handler.Document
        else
            let inserted = withPendingFormat handler handler.Document (Doc.insertText text handler.Document)
            match inserted.Selection with
            | Some selection when selection.IsCollapsed && selection.Focus.Offset >= text.Length ->
                match Doc.findParagraph inserted selection.Focus.BlockId with
                | Some para ->
                    let buffer = TextBuffer.applyStyle (selection.Focus.Offset - text.Length) text.Length (TextBuffer.withExplicitUnderline true) para.Buffer
                    Doc.replaceParagraph inserted para.Id { para with Buffer = buffer }
                | None -> inserted
            | _ -> inserted

    let showComposition (text: string) =
        let preview = compositionPreview text
        renderWithHover handler.RenderContext preview handler.HoveredImage
        syncInputProxy preview

    /// Text that arrived through the proxy becomes one undo step; with a
    /// `typing` kind (see TypingRun) it goes on the run of typing before it.
    let commitInputAs (typing: string option) description (afterDoc: Doc.Document) =
        let beforeDoc = handler.Document
        if not (documentsEqualForChange afterDoc beforeDoc) then
            match typing with
            | Some kind -> commitTyping handler kind description afterDoc |> ignore
            | None -> recordChange handler description beforeDoc afterDoc
            dropPendingFormat handler afterDoc
            handler.Document <- afterDoc
            renderWithHover handler.RenderContext afterDoc handler.HoveredImage
            syncInputProxy afterDoc
            handler.OnChange afterDoc
        else
            // Take a composition preview off the canvas again.
            renderWithHover handler.RenderContext beforeDoc handler.HoveredImage
            syncInputProxy beforeDoc

    let commitInput description afterDoc = commitInputAs None description afterDoc

    proxy.addEventListener("keydown", onKeyDown)

    proxy.addEventListener("compositionstart", fun (_: Event) ->
        handler.Composition <- Some { Text = "" }
        syncInputProxy handler.Document)

    proxy.addEventListener("compositionupdate", fun (e: Event) ->
        let text = match eventData e with | null -> "" | data -> data
        match handler.Composition with
        | Some state -> state.Text <- text
        | None -> handler.Composition <- Some { Text = text }
        showComposition text)

    proxy.addEventListener("compositionend", fun (e: Event) ->
        let text =
            match eventData e with
            | null -> (match handler.Composition with Some state -> state.Text | None -> "")
            | data -> data
        handler.Composition <- None
        setProxyValue proxy ""
        // The whole composition is one insertion, and one undo step.
        commitInput "Typing" (if text = "" then handler.Document else insertTextForMode text handler.Document))

    // Text and deletions that do not come as handled keydowns (virtual
    // keyboards, dictation, the emoji picker, autocorrect).
    proxy.addEventListener("beforeinput", fun (e: Event) ->
        if handler.Composition.IsNone then
            match eventInputType e with
            | "insertText" | "insertReplacementText" ->
                match eventData e with
                | null | "" -> ()
                | data ->
                    e.preventDefault()
                    commitInputAs (Some "type") "Typing" (insertTextForMode data handler.Document)
            | "deleteContentBackward" ->
                e.preventDefault()
                commitInputAs (Some "backspace") "Delete backward" (deleteBackwardForMode handler.Document)
            | "deleteContentForward" ->
                e.preventDefault()
                commitInputAs (Some "delete") "Delete forward" (deleteForwardForMode handler.Document)
            | "deleteWordBackward" ->
                e.preventDefault()
                commitInput "Delete" (deleteWordForMode false handler.Document)
            | "deleteWordForward" ->
                e.preventDefault()
                commitInput "Delete" (deleteWordForMode true handler.Document)
            | "insertLineBreak" ->
                e.preventDefault()
                commitInput "Typing" (insertTextForMode "\n" handler.Document)
            | _ -> ())

    // The browser's copy / cut / paste on the proxy (Ctrl+C / X / V, the
    // browser's menu): the clipboard is the editor's selection, not the proxy's text.
    let onCopy (cut: bool) (e: Event) =
        let data = e?clipboardData
        if not (isNull data) && handler.Composition.IsNone then
            match copyToClipboard cut with
            | Some (text, html) ->
                data?setData("text/plain", text) |> ignore
                data?setData("text/html", html) |> ignore
                e.preventDefault()
            | None -> ()
            setProxyValue proxy ""
    proxy.addEventListener("copy", onCopy false)
    proxy.addEventListener("cut", onCopy true)
    proxy.addEventListener("paste", fun (e: Event) ->
        let data = e?clipboardData
        if not (isNull data) && handler.Composition.IsNone then
            e.preventDefault()
            let read (kind: string) = match unbox<string> (data?getData(kind)) with null -> "" | value -> value
            pasteClipboard (read "text/plain") (read "text/html"))

    proxy.addEventListener("input", fun (_: Event) ->
        // Whatever still reached the proxy (a paste from the context menu, an
        // input type not handled above) is taken over as typed text.
        if handler.Composition.IsNone then
            let text = proxyValue proxy
            if not (System.String.IsNullOrEmpty text) then
                setProxyValue proxy ""
                commitInput "Typing" (insertTextForMode text handler.Document))

    // The canvas stays focusable (clicks, Tab, hosts calling canvas.focus());
    // focus is handed on to the proxy so that text input reaches it.
    canvas.addEventListener("focus", fun (_: Event) ->
        if proxyAttached then
            syncInputProxy handler.Document
            focusWithoutScroll proxy)

    // Handle mouse down - start selection or image resize
    canvas.addEventListener("mousedown", fun (e: Event) ->
        let me = e :?> MouseEvent
        let x, y = pointerPosition me

        // A double click on a page's header or footer edits it in place; one on the body goes back.
        // The shape that is selected is taken by a handle (resized) or by its frame (moved); one
        // whose text a click does not reach (behind the text, a watermark) by any point of it.
        let grabbed =
            match handler.SelectedShape with
            | Some id when int me.detail < 2 ->
                match shapePartAt handler.RenderContext handler.Document id x y with
                | "" -> None
                | "inside" ->
                    let editingIt = (match handler.Story with Some session -> session.Box = Some id | None -> false)
                    let reached =
                        editingIt
                        || shapeAt handler.RenderContext handler.Document x y = Some id
                        || (handler.Story.IsSome && (match furnitureShapeAt handler.RenderContext handler.Document x y with Some (boxId, _, _) -> boxId = id | None -> false))
                    if reached then None else Some (id, "move")
                | "border" -> Some (id, "move")
                | handle -> Some (id, handle)
            | _ -> None
        // Another shape of the body is taken by its frame: selected, and moved if the pointer is
        // dragged - a shape behind the text and one without text are reached so. (Not while a
        // header or a footer is edited: the body's shapes are out of reach then.)
        let grabbed =
            match grabbed, handler.Story with
            | Some _, _ -> grabbed
            | None, story when int me.detail < 2 && (match story with None -> true | Some session -> session.Box.IsSome && session.PartIndex < 0) ->
                shapeEdgeAt handler.RenderContext handler.Document x y |> Option.map (fun id -> id, "move")
            | None, _ -> None
        match grabbed with
        | Some (id, part) ->
            // (Its text, or the header that is edited, is left first: the shape is the document's.)
            if handler.Story.IsSome then exitStory handler |> ignore
            if setSelectedShape handler (Some id) then
                renderWithHover handler.RenderContext handler.Document handler.HoveredImage
                handler.OnShapeSelection ()
            match shapeRectOf handler.RenderContext handler.Document id with
            | Some (shape, rectX, rectY, width, height, pageTop) ->
                handler.ShapeDrag <- Some { ShapeId = id; Part = part; StartX = x; StartY = y; Shape = shape; RectX = rectX; RectY = rectY; RectWidth = width; RectHeight = height; PageTop = pageTop; Moved = false }
                handler.PointerStartDocument <- Some handler.Document
                handler.IsMouseDown <- true
            | None -> ()
            e.preventDefault()
            canvas.focus()
        | None ->
        if int me.detail = 2 && handler.OnDoubleClick x y then
            e.preventDefault()
            canvas.focus()
        else
        // A click in a text box edits its text; one outside the box that is edited goes back to the body.
        let boxAt = shapeAt handler.RenderContext handler.Document x y
        (match handler.Story with
         // The text box of a header or footer is edited: a click in it on another page that shows
         // it, or in another box of a header or footer, edits that; one elsewhere in a header or
         // footer goes back to that header, one on the body to the body.
         | Some session when session.Box.IsSome && session.PartIndex >= 0 ->
            match furnitureShapeAt handler.RenderContext handler.Document x y with
            | Some (boxId, _, pageIndex) when Some boxId = session.Box && pageIndex = session.PageIndex -> ()
            | Some (boxId, _, pageIndex) -> enterTextBoxOnPage handler boxId None (Some pageIndex) |> ignore
            | None ->
                let region = pageRegionAt handler.RenderContext handler.Document x y
                exitStory handler |> ignore
                match region, boxAt with
                | Some (pageIndex, place), _ when place = "header" || place = "footer" -> handler.OnHistoryPlace pageIndex (place = "header") |> ignore
                | _, Some boxId -> enterTextBox handler boxId None |> ignore
                | _ -> ()
         | Some session when session.Box.IsSome ->
            if boxAt <> session.Box then
                exitStory handler |> ignore
                match boxAt with
                | Some boxId -> enterTextBox handler boxId None |> ignore
                | None -> ()
         | Some session when session.Note.IsSome ->
            // A note or comment is edited: a click off it goes back to the body (the caret goes where the click was).
            if not (onNoteStory handler.RenderContext handler.Document x y) then exitStory handler |> ignore
         | Some _ ->
            // A header or footer is edited: a click in one of its text boxes (or of another
            // header's or footer's) edits the box's text, on the page the click is on.
            match furnitureShapeAt handler.RenderContext handler.Document x y with
            | Some (boxId, _, pageIndex) -> enterTextBoxOnPage handler boxId None (Some pageIndex) |> ignore
            | None -> ()
         | None ->
            match boxAt with
            | Some boxId -> enterTextBox handler boxId None |> ignore
            | None -> ())
        // The box whose text the click went into is the shape that is selected; a click anywhere
        // else lets go of the one that was.
        let shapeNow = (match handler.Story with Some session -> session.Box | None -> None)
        if setSelectedShape handler shapeNow then
            renderWithHover handler.RenderContext handler.Document handler.HoveredImage
            handler.OnShapeSelection ()
        // Check if clicking on an image first
        match hitTestForImage handler.RenderContext handler.Document x y with
        | Some (blockId, offset, width, height, imageX, imageY) ->
            if isResizeHandle x y imageX imageY width height then
                handler.ResizingImage <- Some {
                    BlockId = blockId
                    Offset = offset
                    InitialX = x
                    InitialY = y
                    InitialWidth = width
                    InitialHeight = height
                }
            else
                handler.DraggingImage <- Some { BlockId = blockId; Offset = offset; InitialX = x; InitialY = y }
            handler.PointerStartDocument <- Some handler.Document
            handler.IsMouseDown <- true
            canvas.focus()
        | None ->
            // Regular text selection
            handler.GoalX <- None
            match hitTest handler.RenderContext handler.Document x y with
            | None -> ()
            | Some (blockId, offset) ->
                let position : Doc.Position = { BlockId = blockId; Offset = offset }
                let clicks = int me.detail
                if clicks >= 3 then
                    // Third click: the paragraph. (No drag selection follows from it.)
                    applyPointerDocumentChange (paragraphAt handler.Document position)
                elif clicks = 2 then
                    // Double click: the word.
                    applyPointerDocumentChange (wordAt handler.Document position)
                elif me.shiftKey && handler.Document.Selection.IsSome then
                    // Shift+click: from where the selection started to here.
                    let anchor = handler.Document.Selection.Value.Anchor
                    handler.IsMouseDown <- true
                    handler.SelectionAnchor <- Some anchor
                    applyPointerDocumentChange (Doc.setSelection anchor position handler.Document)
                else
                    handler.IsMouseDown <- true
                    handler.SelectionAnchor <- Some position
                    let newDoc = Doc.setCursor blockId offset handler.Document
                    applyPointerDocumentChange newDoc
                canvas.focus()
    )

    // Handle mouse move - extend selection, resize image, or update hover
    canvas.addEventListener("mousemove", fun (e: Event) ->
        let me = e :?> MouseEvent
        let x, y = pointerPosition me

        if handler.IsMouseDown && handler.ShapeDrag.IsSome then
            // The selected shape follows the pointer: from the document the drag began with, so
            // that the whole drag is one change (and one undo step, see the mouse up).
            match handler.ShapeDrag, handler.PointerStartDocument with
            | Some drag, Some startDoc ->
                let dx, dy = x - drag.StartX, y - drag.StartY
                let threshold = 3.0 / viewportScale handler.RenderContext
                if drag.Moved || abs dx > threshold || abs dy > threshold then
                    drag.Moved <- true
                    let moved =
                        if drag.Part = "move" then shapePlacedDocument startDoc drag.Shape drag.RectX drag.RectY drag.PageTop dx dy None
                        else shapeResizedDocument startDoc drag.Shape drag.RectX drag.RectY drag.RectWidth drag.RectHeight drag.PageTop drag.Part dx dy
                    match moved with
                    | Some newDoc -> applyPointerDocumentChange newDoc
                    | None -> ()
            | _ -> ()
        elif handler.IsMouseDown then
            match handler.ResizingImage with
            | Some resizeState ->
                // Resize image based on mouse delta
                let deltaX = x - resizeState.InitialX
                let deltaY = y - resizeState.InitialY

                // Maintain aspect ratio by using the larger delta
                let delta = if abs deltaX > abs deltaY then deltaX else deltaY
                let newWidth = max 20.0 (resizeState.InitialWidth + delta)
                let newHeight = max 20.0 (resizeState.InitialHeight + delta)

                let newDoc = Doc.resizeImage resizeState.BlockId resizeState.Offset newWidth newHeight handler.Document
                applyPointerDocumentChange newDoc
            | None ->
                match handler.DraggingImage with
                | Some _ -> setCursor canvas "move"
                | None ->
                    // Regular text selection
                    match hitTest handler.RenderContext handler.Document x y, handler.SelectionAnchor with
                    | Some (blockId, offset), Some anchor ->
                        let newDoc = Doc.setSelection anchor { BlockId = blockId; Offset = offset } handler.Document
                        applyPointerDocumentChange newDoc
                    | _ -> ()
        else
            // Over a handle or the frame of the shape that is selected: the cursor says what a drag does.
            let overShape =
                match handler.SelectedShape with
                | Some id ->
                    match shapePartAt handler.RenderContext handler.Document id x y with
                    | "" | "inside" -> None
                    | "border" -> Some "move"
                    | handle -> shapeHandleSpots |> List.tryPick (fun (name, _, _, cursor) -> if name = handle then Some cursor else None)
                | None -> None
            match overShape with
            | Some cursor -> setCursor canvas cursor
            | None ->
            // Update hover state and cursor
            match hitTestForImage handler.RenderContext handler.Document x y with
            | Some (blockId, offset, width, height, imageX, imageY) ->
                setCursor canvas (if isResizeHandle x y imageX imageY width height then "nwse-resize" else "move")
                let newHover = Some (blockId, offset)
                if handler.HoveredImage <> newHover then
                    handler.HoveredImage <- newHover
                    renderWithHover handler.RenderContext handler.Document handler.HoveredImage
            | None ->
                setCursor canvas "text"
                if handler.HoveredImage.IsSome then
                    handler.HoveredImage <- None
                    renderWithHover handler.RenderContext handler.Document None
    )

    // Handle mouse up - end selection, resize or drag. Listening on the
    // window rather than the canvas means a release outside the canvas still
    // ends the gesture; otherwise IsMouseDown stays set and the next mouse
    // move over the canvas keeps selecting or resizing with no button held.
    let onMouseUp (e: Event) =
        if handler.ShapeDrag.IsSome then
            // The whole drag of a shape is one undo step (a click on it: none).
            match handler.ShapeDrag, handler.PointerStartDocument with
            | Some drag, Some startDoc when drag.Moved ->
                recordChange handler (if drag.Part = "move" then "Move shape" else "Resize shape") startDoc handler.Document
                // (The listeners heard of the drag as it went; now there is a step to undo.)
                handler.OnChange handler.Document
            | _ -> ()
            handler.ShapeDrag <- None
            handler.PointerStartDocument <- None
            handler.IsMouseDown <- false
        elif handler.IsMouseDown || handler.ResizingImage.IsSome || handler.DraggingImage.IsSome then
            let me = e :?> MouseEvent
            let x, y = pointerPosition me

            // A release that is still on the picked-up object is a click, not
            // a move: the object stays where it is, exactly as it was.
            let releasedOnDraggedImage (dragState: DragImageState) =
                match hitTestForImage handler.RenderContext handler.Document x y with
                | Some (blockId, offset, _, _, _, _) -> blockId = dragState.BlockId && offset = dragState.Offset
                | None -> false

            let dragThreshold = 3.0 / viewportScale handler.RenderContext
            match handler.DraggingImage with
            | Some dragState when (abs (x - dragState.InitialX) > dragThreshold || abs (y - dragState.InitialY) > dragThreshold) && not (releasedOnDraggedImage dragState) ->
                let target =
                    match hitTest handler.RenderContext handler.Document x y with
                    | Some target -> target
                    | None ->
                        match Doc.findParagraph handler.Document dragState.BlockId with
                        | Some para -> (dragState.BlockId, TextBuffer.length para.Buffer)
                        | None -> (dragState.BlockId, dragState.Offset)
                let targetBlockId, targetOffset = target
                let effectiveOffset =
                    match Doc.findParagraph handler.Document targetBlockId with
                    | Some para when targetBlockId = dragState.BlockId && targetOffset <= dragState.Offset -> TextBuffer.length para.Buffer
                    | Some para when targetBlockId = dragState.BlockId && targetOffset = dragState.Offset -> TextBuffer.length para.Buffer
                    | _ -> targetOffset
                let newDoc = Doc.moveImage dragState.BlockId dragState.Offset targetBlockId effectiveOffset handler.Document
                applyPointerDocumentChange newDoc
            | Some dragState when handler.ResizingImage.IsNone ->
                // A click on the picture (no drag): the picture is selected, as it is.
                let first : Doc.Position = { BlockId = dragState.BlockId; Offset = dragState.Offset }
                applyPointerDocumentChange (Doc.setSelection first { first with Offset = dragState.Offset + 1 } handler.Document)
            | _ -> ()

            // The whole resize or move is one undo step.
            match handler.PointerStartDocument with
            | Some startDoc ->
                recordChange handler (if handler.ResizingImage.IsSome then "Resize image" else "Move image") startDoc handler.Document
            | None -> ()

            handler.PointerStartDocument <- None
            handler.IsMouseDown <- false
            handler.ResizingImage <- None
            handler.DraggingImage <- None

    addWindowListener "mouseup" onMouseUp
    handler.Dispose <- fun () ->
        removeWindowListener "mouseup" onMouseUp
        // The input proxy belongs to the handler: take it out of the page.
        proxyAttached <- false
        handler.Composition <- None
        handler.SyncInputProxy <- ignore
        handler.InputProxy <- None
        if not (isNull proxy.parentNode) then proxy.parentNode.removeChild proxy |> ignore
        handler.Dispose <- ignore

    // Initial render
    renderWithHover ctx initialDoc None

    handler

/// Show a document without touching the undo history (selection changes,
/// re-renders).
let updateDocument (handler: InputHandler) (doc: Doc.Document) =
    // (No end of the selection in the mark of a comment: see Doc.snapSelectionToMarks.)
    let doc = Doc.snapSelectionToMarks handler.Document.Selection doc
    dropPendingFormat handler doc
    handler.Document <- doc
    renderWithHover handler.RenderContext doc None
    handler.SyncInputProxy doc
    handler.OnChange doc

/// Paint the current document again (after the viewport or the canvas
/// changed) and put the input proxy back at the caret.
let repaint (handler: InputHandler) =
    renderWithHover handler.RenderContext handler.Document handler.HoveredImage
    handler.SyncInputProxy handler.Document

/// Apply a change made outside the keyboard (API call, toolbar, dialog) as
/// one undo step.
let applyEdit (handler: InputHandler) (description: string) (doc: Doc.Document) =
    recordChange handler description handler.Document doc
    updateDocument handler doc

/// Apply a change of the whole document that was made from within a page's
/// header or footer (no story is on display while it is applied) as one undo
/// step; undo and redo of it put the caret back into that page's header.
let applyEditAtStoryPlace (handler: InputHandler) (description: string) (pageIndex: int) (isHeader: bool) (doc: Doc.Document) =
    let before = handler.History
    recordChange handler description handler.Document doc
    match handler.History.Undos with
    | head :: rest when not (obj.ReferenceEquals(handler.History.Undos, before.Undos)) ->
        let stepped page header = handler.HistoryPlace <- Some (page, header)
        handler.History <- { handler.History with Undos = (StoryPlaceCommand(head, pageIndex, isHeader, stepped) :> ICommand) :: rest }
    | _ -> ()
    handler.Document <- doc
    handler.OnChange doc

/// Replace the document with a newly loaded one. Undo history belongs to the
/// previous document and is discarded.
let loadDocument (handler: InputHandler) (doc: Doc.Document) =
    hideStory handler
    handler.History <- Cmd.empty
    handler.PointerStartDocument <- None
    handler.IsMouseDown <- false
    handler.ResizingImage <- None
    handler.DraggingImage <- None
    handler.SelectionAnchor <- None
    setSelectedShape handler None |> ignore
    updateDocument handler doc

/// A step of the history whose documents were made before more of the
/// document was read: what was read since is given to what it comes back to.
type private ExtendedCommand(inner: ICommand, extend: Doc.Document -> Doc.Document) =
    interface ICommand with
        member _.Execute doc = extend (inner.Execute doc)
        member _.Undo doc = extend (inner.Undo doc)
        member _.Description = inner.Description

/// More of the document was read (a package opened front first): `extend`
/// gives a document what was read. The document on display gets it, and so
/// does every document the handler may come back to - the steps of the undo
/// history hold the documents before and after each edit, made when less was
/// read. Nothing is painted and nothing is recorded: the caller repaints when
/// it has to. Returns the document on display. (Not while a header or footer
/// is edited in place: the editor reads the document to its end before that.)
let extendDocuments (handler: InputHandler) (extend: Doc.Document -> Doc.Document) : Doc.Document =
    let doc = extend handler.Document
    handler.Document <- doc
    handler.PointerStartDocument <- handler.PointerStartDocument |> Option.map extend
    let history = handler.History
    if not history.Undos.IsEmpty || not history.Redos.IsEmpty then
        let extended (commands: ICommand list) = commands |> List.map (fun command -> ExtendedCommand(command, extend) :> ICommand)
        handler.History <- { history with Undos = extended history.Undos; Redos = extended history.Redos }
    doc

/// Detach listeners registered outside the canvas.
let dispose (handler: InputHandler) =
    handler.Dispose ()

#else
// Server-side stub
type InputHandler = unit
let create _ _ _ = ()
let updateDocument _ _ = ()
let repaint _ = ()
let applyEdit _ _ _ = ()
let pendingFormat _ = None
let pendingProbe _ = None
let formatAtCaret _ _ = false
let characterFormatted _ command doc = command doc
let withPendingFormat _ _ after = after
let applyEditAtStoryPlace _ _ _ _ _ = ()
let enterTextBox _ _ _ = false
let enterTextBoxOnPage _ _ _ _ = false
let selectShape _ _ = ()
let shapePlacedDocument _ _ _ _ _ _ _ _ = None
let shapeResizedDocument _ _ _ _ _ _ _ _ _ _ = None
let pageBreakAllowed _ = false
let loadDocument _ _ = ()
let extendDocuments _ _ = ()
let undo _ = false
let redo _ = false
let canUndo _ = false
let canRedo _ = false
let dispose _ = ()
#endif
