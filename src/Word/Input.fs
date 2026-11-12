module Rbt.Word.Input

let private sameListOrEqual left right =
    obj.ReferenceEquals(left, right) || left = right

/// The bodies of the notes and comments are the same (a note's body is edited in place).
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
let sameContentObjects = Rbt.Word.Session.sameContentObjects

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
open Rbt.Word.Session

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



/// Input handler state
type InputHandler = {
    Canvas: HTMLCanvasElement
    RenderContext: RenderContext
    Session: Rbt.Word.Session.EditingSession
    OnChange: Doc.Document -> unit
    mutable IsMouseDown: bool
    mutable SelectionAnchor: Doc.Position option
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
    /// Removes all listeners and invalidates pending callbacks (set by `create`).
    mutable Dispose: unit -> unit
    /// The column Up / Down / PageUp / PageDown keep to (layout x of the
    /// caret when the run of vertical moves started).
    mutable GoalX: float option
    /// Puts the caret into a page's header or footer after such a step (set by the editor).
    mutable OnHistoryPlace: int -> bool -> bool
    /// A double click at a point (layout coordinates): true when it was
    /// taken (into a header or footer, or back to the body) - set by the editor.
    mutable OnDoubleClick: float -> float -> bool
    /// Called before a key that goes by the whole document (Ctrl+A, Ctrl+End):
    /// the editor reads a document that is still being read to its end.
    mutable EnsureWhole: unit -> unit
    /// Called when the pending format changed: what the selection's format
    /// is said to be has changed though the document has not (set by the editor).
    mutable OnPendingFormat: unit -> unit
    /// Put plain text in at the caret, or over the selection, as Ctrl+V does
    /// (its lines are paragraphs; one undo step). Set when the handler is made.
    mutable PasteText: string -> unit
    /// Paste clipboard text and HTML. Only a marked copy from this handler
    /// reuses its original rich clip; other HTML is parsed independently, and
    /// plain text is always pasted as plain text. One undo step.
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
} with
    member this.Document
        with get () = this.Session.Document
        and set value = this.Session.Document <- value
    member this.History
        with get () = this.Session.History
        and set value = this.Session.History <- value
    member this.Story
        with get () = this.Session.Story
        and set value = this.Session.Story <- value
    member this.HistoryPlace
        with get () = this.Session.HistoryPlace
        and set value = this.Session.HistoryPlace <- value

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

// Internal rich paste requires provenance in clipboard data, not equality
// with text left by an earlier copy (possibly in another editor).
[<Emit("(globalThis.crypto && crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2) + Date.now().toString(36))")>]
let private clipboardId () : string = jsNative

[<Emit("console.error($0, $1)")>]
let private clipboardError (_message: string) (_error: obj) : unit = jsNative

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


/// The whole document (body, headers and footers) of the document the
/// handler works on: that document, or - while a header or footer is edited
/// in place - the document with the story's blocks in their part. This is
/// what is saved, and what a revision is a state of.
let wholeDocument (handler: InputHandler) (doc: Doc.Document) : Doc.Document =
    handler.Session.Whole(doc)


/// A history with the step just added to it made a story's step when a story is being edited.
let private storyHistory (handler: InputHandler) (before: CommandHistory) (after: CommandHistory) : CommandHistory =
    match after.Undos with
    | head :: rest when not (obj.ReferenceEquals(after.Undos, before.Undos)) ->
        { after with Undos = handler.Session.WrapStory(head) :: rest }
    | _ -> after

let private showStory (handler: InputHandler) (session: StorySession) =
    let ctx = handler.RenderContext
    // (The whole document for the story document last asked about: a caret move asks again and again.)
    let last = ResizeArray<Doc.Document * Doc.Document>()
    let whole (story: Doc.Document) =
        if last.Count > 0 && sameContentObjects story (fst last.[0]) then snd last.[0]
        else
            let folded = { handler.Session.FoldStory(session, story) with Selection = None }
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


/// Record a finished change; selection-only changes are not recorded.
let private recordChange (handler: InputHandler) (description: string) (beforeDoc: Doc.Document) (afterDoc: Doc.Document) =
    if not (sameContent beforeDoc afterDoc) then
        handler.Session.RecordChange(description, beforeDoc, afterDoc, id)

let private typingHistory (handler: InputHandler) (kind: string) (description: string) (beforeDoc: Doc.Document) (afterDoc: Doc.Document) : CommandHistory =
    handler.Session.RecordTyping(kind, description, beforeDoc, afterDoc, id)
    handler.History

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
    if handler.Session.HistoryStep(step) then
        match handler.Story with
        | Some session -> showStory handler session
        | None -> hideStory handler
    handler.Document, handler.History

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
    handler.Session.EnterStory(session, selection)
    showStory handler session
    handler.IsMouseDown <- false
    handler.SelectionAnchor <- None
    handler.GoalX <- None
    showDocument handler handler.Document

/// Back to the body, with the selection it had. False when no header or footer was being edited.
let exitStory (handler: InputHandler) : bool =
    if not (handler.Session.ExitStory()) then false
    else
        hideStory handler
        handler.IsMouseDown <- false
        handler.SelectionAnchor <- None
        handler.GoalX <- None
        showDocument handler handler.Document
        true

let pendingFormat (handler: InputHandler) = handler.Session.PendingFormat
let pendingProbe (handler: InputHandler) = handler.Session.PendingProbe()
let private executeSessionCommand (handler: InputHandler) command tracking =
    let previous = handler.Session.PendingFormat |> Option.map (fun pending -> pending.PendingStyle)
    let result = handler.Session.Execute(command, tracking)
    if previous <> (handler.Session.PendingFormat |> Option.map (fun pending -> pending.PendingStyle)) then handler.OnPendingFormat ()
    result
let withPendingFormat (handler: InputHandler) before after = handler.Session.WithPendingFormat(before, after)
let private dropPendingFormat (handler: InputHandler) doc = handler.Session.DropPendingFormat(doc)
let private isWordChar = Rbt.Word.Session.isWordChar

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
        Session = Rbt.Word.Session.EditingSession(initialDoc)
        OnChange = onChange
        IsMouseDown = false
        SelectionAnchor = None
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
        OnHistoryPlace = fun _ _ -> false
        OnDoubleClick = fun _ _ -> false
        EnsureWhole = ignore
        OnPendingFormat = ignore
        PasteText = ignore
        PasteClipboard = fun _ _ -> ()
        CopyToClipboard = fun _ -> None
        SelectedShape = None
        ShapeDrag = None
        OnShapeSelection = ignore
    }
    let mutable disposed = false
    let listeners = ResizeArray<HTMLElement * string * (Event -> unit)>()
    let listen (element: HTMLElement) (eventName: string) (callback: Event -> unit) =
        let guarded event = if not disposed then callback event
        element.addEventListener(eventName, guarded)
        listeners.Add(element, eventName, guarded)
    let mutable skipParagraphInput = false
    let mutable suppressPhysicalParagraphBeforeInput = false

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
        Rbt.Word.Session.navigate (Rbt.Word.Session.Navigation.VisualTarget(Some position, false, extend)) doc

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

    let tracking () =
        if handler.TrackChanges then Some (handler.TrackChangesAuthor, handler.TrackChangesDate)
        else None
    let publishClipboard outcome =
        if outcome <> Rbt.Word.Session.NoChange then
            let doc = handler.Document
            renderWithHover handler.RenderContext doc handler.HoveredImage
            handler.SyncInputProxy doc
            handler.OnChange doc
    let captureCopy () =
        handler.Session.CaptureClipboard(fun clip ->
            Doc.clipToHtml clip + sprintf "<!--rbt-word-clip:%s-->" (clipboardId ()))
    let commitCut capture =
        match handler.Session.CommitCut(capture, tracking ()) with
        | Some outcome -> publishClipboard outcome
        | None -> clipboardError "Word cut cancelled: selection or document changed while copying" null
    let pasteClipboard (text: string) (html: string) =
        if not disposed then
            publishClipboard (handler.Session.PasteClipboard(text, html, htmlClip, tracking ()))
    let pasteText text =
        if not disposed then
            publishClipboard (handler.Session.PasteClipboard(text, "", Rbt.Word.Session.noExternalClip, tracking ()))
    let copyToClipboard (cut: bool) =
        if disposed then None
        else
            captureCopy ()
            |> Option.map (fun capture ->
                if cut then commitCut capture
                capture.Text, capture.Html)
    handler.PasteText <- pasteText
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
            let ranged = Rbt.Word.Session.navigate (Rbt.Word.Session.Navigation.Word(forward, true)) doc
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
    let originalTabIndex = canvas.tabIndex
    let originalCursor : string = canvas?style?cursor
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

    /// Character shortcuts use the same atomic session command as the toolbar.
    let tracking () =
        if handler.TrackChanges then Some (handler.TrackChangesAuthor, handler.TrackChangesDate) else None

    let executeWord command =
        executeSessionCommand handler command (tracking ())

    let showWordCommand command =
        match executeWord command with
        | Ok Rbt.Word.Session.NoChange | Error _ -> ()
        | Ok _ ->
            let doc = handler.Document
            renderWithHover handler.RenderContext doc handler.HoveredImage
            handler.SyncInputProxy doc
            handler.OnChange doc


    // Handle keyboard input. The listener serves the canvas and the input
    // proxy (which normally has the focus).
    let onKeyDown (e: Event) =
        let ke = e :?> KeyboardEvent
        // (Everything, or the end of the document: of the whole document.)
        if (ke.ctrlKey || ke.metaKey) && (keyEquals 'a' ke.key || ke.key = "End") then handler.EnsureWhole ()
        let beforeKey = handler.Document
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
        let wordCommand command =
            match executeWord command with
            | Ok Rbt.Word.Session.NoChange | Error _ -> ()
            | Ok _ ->
                newDoc <- handler.Document
                newHistory <- handler.History
                useCommand <- true
                stepped <- true
        let navigation command =
            handler.Session.Navigate(command) |> ignore
            newDoc <- handler.Document
        let visualNavigation (after: Doc.Document) =
            match after.Selection with
            | Some selected -> navigation (Rbt.Word.Session.Navigation.VisualTarget(Some selected.Focus, false, shift))
            | None -> ()

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
            navigation Rbt.Word.Session.Navigation.SelectAll
            e.preventDefault()
        | key when commandKey && keyEquals 'b' key ->
            wordCommand (Rbt.Word.Session.WordCommand.Character Rbt.Word.Session.CharacterCommand.Bold)
            e.preventDefault()
        | key when commandKey && keyEquals 'i' key ->
            wordCommand (Rbt.Word.Session.WordCommand.Character Rbt.Word.Session.CharacterCommand.Italic)
            e.preventDefault()
        | key when commandKey && keyEquals 'u' key ->
            wordCommand (Rbt.Word.Session.WordCommand.Character Rbt.Word.Session.CharacterCommand.Underline)
            e.preventDefault()
        | "1" when commandKey ->
            wordCommand (Rbt.Word.Session.HeadingLevel 1)
            e.preventDefault()
        | "2" when commandKey ->
            wordCommand (Rbt.Word.Session.HeadingLevel 2)
            e.preventDefault()
        | "3" when commandKey ->
            wordCommand (Rbt.Word.Session.HeadingLevel 3)
            e.preventDefault()
        | "F3" when shift && not commandKey && not ke.altKey ->
            // Shift+F3: Change Case of the selection or the word at the caret.
            newDoc <- Doc.changeCase handler.Document
            e.preventDefault()
        | key when commandKey && shift && not ke.altKey && keyEquals 'l' key ->
            // Ctrl+Shift+L: a bulleted list (Word's List Bullet).
            wordCommand Rbt.Word.Session.ApplyBulletList
            e.preventDefault()
        | key when commandKey && shift && not ke.altKey && keyEquals 'n' key ->
            // Ctrl+Shift+N: the Normal style.
            wordCommand Rbt.Word.Session.NormalStyle
            e.preventDefault()
        | key when commandKey && not ke.altKey && not shift && (keyEquals 'e' key || keyEquals 'l' key || keyEquals 'r' key || keyEquals 'j' key) ->
            // Word's alignment keys: Ctrl+E centre, Ctrl+L left, Ctrl+R right, Ctrl+J justify.
            let align =
                if keyEquals 'e' key then Rbt.Word.Session.AlignCenter
                elif keyEquals 'r' key then Rbt.Word.Session.AlignRight
                elif keyEquals 'j' key then Rbt.Word.Session.AlignJustify
                else Rbt.Word.Session.AlignLeft
            wordCommand align
            e.preventDefault()
        | key when commandKey && not ke.altKey && keyEquals 'm' key ->
            // Ctrl+M / Ctrl+Shift+M: Increase / Decrease Indent.
            wordCommand (Rbt.Word.Session.IndentParagraphs (not shift))
            e.preventDefault()
        | key when commandKey && not ke.altKey && (key = "]" || key = "[" || key = ">" || key = "<" || key = "=" || key = "+" || key = " " || (shift && (key = "." || key = ","))) ->
            // Word's character keys: Ctrl+] / Ctrl+[ a point larger / smaller,
            // Ctrl+Shift+> / < the next size up / down, Ctrl+= subscript,
            // Ctrl+Shift+= superscript, Ctrl+Space the character formatting cleared.
            let format =
                match key with
                | "]" -> Rbt.Word.Session.CharacterCommand.GrowFont(true, false)
                | "[" -> Rbt.Word.Session.CharacterCommand.GrowFont(false, false)
                | ">" | "." -> Rbt.Word.Session.CharacterCommand.GrowFont(true, true)
                | "<" | "," -> Rbt.Word.Session.CharacterCommand.GrowFont(false, true)
                | "+" -> Rbt.Word.Session.CharacterCommand.Superscript
                | "=" when shift -> Rbt.Word.Session.CharacterCommand.Superscript
                | "=" -> Rbt.Word.Session.CharacterCommand.Subscript
                | _ -> Rbt.Word.Session.CharacterCommand.Clear
            wordCommand (Rbt.Word.Session.WordCommand.Character format)
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
                        soon (fun () -> if not disposed && proxyValue element = text then setProxyValue element "")
                | _ -> ()
            else
                // (The key reached the canvas itself: through the asynchronous clipboard.)
                e.preventDefault()
                if keyEquals 'v' key then
                    let epoch = handler.Session.Epoch
                    try
                        Clipboard.pasteText()
                            .``then``(fun text ->
                                if not disposed && obj.ReferenceEquals(epoch, handler.Session.Epoch) then
                                    pasteClipboard text "")
                            .``catch``(fun error -> if not disposed then clipboardError "Word paste failed" (box error))
                        |> ignore
                    with error ->
                        clipboardError "Word paste failed" (box error)
                else
                    match captureCopy () with
                    | None -> ()
                    | Some capture ->
                        try
                            (Clipboard.copyText capture.Text)
                                .``then``(fun () -> if not disposed && keyEquals 'x' key then commitCut capture)
                                .``catch``(fun error -> if not disposed then clipboardError "Word clipboard write failed" (box error))
                            |> ignore
                        with error ->
                            clipboardError "Word clipboard write failed" (box error)
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
            if ke.metaKey then visualNavigation (moveToLineEdge handler.Document false shift)
            else navigation (if ke.ctrlKey || ke.altKey then Rbt.Word.Session.Navigation.Word(false, shift)
                             else Rbt.Word.Session.Navigation.Character(false, shift))
            e.preventDefault()
        | "ArrowRight" ->
            if ke.metaKey then visualNavigation (moveToLineEdge handler.Document true shift)
            else navigation (if ke.ctrlKey || ke.altKey then Rbt.Word.Session.Navigation.Word(true, shift)
                             else Rbt.Word.Session.Navigation.Character(true, shift))
            e.preventDefault()
        | "ArrowUp" ->
            if ke.metaKey then navigation (Rbt.Word.Session.Navigation.DocumentEdge(false, shift))
            elif ke.ctrlKey || ke.altKey then navigation (Rbt.Word.Session.Navigation.Paragraph(false, shift))
            else
                visualNavigation (moveVertically handler.Document false shift)
                keepGoal <- true
            e.preventDefault()
        | "ArrowDown" ->
            if ke.metaKey then navigation (Rbt.Word.Session.Navigation.DocumentEdge(true, shift))
            elif ke.ctrlKey || ke.altKey then navigation (Rbt.Word.Session.Navigation.Paragraph(true, shift))
            else
                visualNavigation (moveVertically handler.Document true shift)
                keepGoal <- true
            e.preventDefault()
        | "PageUp" ->
            visualNavigation (movePage handler.Document false shift)
            keepGoal <- true
            e.preventDefault()
        | "PageDown" ->
            visualNavigation (movePage handler.Document true shift)
            keepGoal <- true
            e.preventDefault()
        | "Home" ->
            if commandKey then navigation (Rbt.Word.Session.Navigation.DocumentEdge(false, shift))
            else visualNavigation (moveToLineEdge handler.Document false shift)
            e.preventDefault()
        | "End" ->
            if commandKey then navigation (Rbt.Word.Session.Navigation.DocumentEdge(true, shift))
            else visualNavigation (moveToLineEdge handler.Document true shift)
            e.preventDefault()
        | "Tab" ->
            wordCommand (Rbt.Word.Session.Tab (shift, ke.ctrlKey))
            e.preventDefault()
        | "Enter" when commandKey ->
            wordCommand (if shift then Rbt.Word.Session.ColumnBreak else Rbt.Word.Session.PageBreak)
            e.preventDefault()
        | "Enter" when shift ->
            wordCommand Rbt.Word.Session.LineBreak
            e.preventDefault()
        | "Enter" ->
            wordCommand Rbt.Word.Session.Enter
            skipParagraphInput <- true
            suppressPhysicalParagraphBeforeInput <- true
            soon (fun () ->
                if not disposed then
                    skipParagraphInput <- false
                    suppressPhysicalParagraphBeforeInput <- false)
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
        if not (documentsEqualForChange newDoc beforeKey) then
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

    listen canvas "keydown" onKeyDown

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
    let commitParagraphInput () = showWordCommand Rbt.Word.Session.Enter

    listen proxy "keydown" onKeyDown

    listen proxy "compositionstart" (fun (_: Event) ->
        handler.Composition <- Some { Text = "" }
        syncInputProxy handler.Document)

    listen proxy "compositionupdate" (fun (e: Event) ->
        let text = match eventData e with | null -> "" | data -> data
        match handler.Composition with
        | Some state -> state.Text <- text
        | None -> handler.Composition <- Some { Text = text }
        showComposition text)

    listen proxy "compositionend" (fun (e: Event) ->
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
    listen proxy "beforeinput" (fun (e: Event) ->
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
            | "insertParagraph" ->
                e.preventDefault()
                setProxyValue proxy ""
                if suppressPhysicalParagraphBeforeInput then
                    suppressPhysicalParagraphBeforeInput <- false
                else
                    skipParagraphInput <- true
                    soon (fun () -> if not disposed then skipParagraphInput <- false)
                    commitParagraphInput ()
            | "insertLineBreak" ->
                e.preventDefault()
                showWordCommand Rbt.Word.Session.LineBreak
            | _ -> ())

    // The browser's copy / cut / paste on the proxy (Ctrl+C / X / V, the
    // browser's menu): the clipboard is the editor's selection, not the proxy's text.
    let onCopy (cut: bool) (e: Event) =
        let data = e?clipboardData
        if not (isNull data) && handler.Composition.IsNone then
            match captureCopy () with
            | Some capture ->
                try
                    data?setData("text/plain", capture.Text) |> ignore
                    data?setData("text/html", capture.Html) |> ignore
                    e.preventDefault()
                    if cut then commitCut capture
                with error ->
                    clipboardError "Word clipboard transfer failed" (box error)
            | None -> ()
            setProxyValue proxy ""
    listen proxy "copy" (onCopy false)
    listen proxy "cut" (onCopy true)
    listen proxy "paste" (fun (e: Event) ->
        let data = e?clipboardData
        if not (isNull data) && handler.Composition.IsNone then
            e.preventDefault()
            let read (kind: string) = match unbox<string> (data?getData(kind)) with null -> "" | value -> value
            pasteClipboard (read "text/plain") (read "text/html"))

    listen proxy "input" (fun (e: Event) ->
        // Browser fallback for virtual paragraph input without beforeinput.
        // A prevented beforeinput does not normally emit input; a synthetic
        // duplicate must not insert a second paragraph.
        if handler.Composition.IsNone then
            let text = proxyValue proxy
            setProxyValue proxy ""
            if eventInputType e = "insertParagraph" then
                if not skipParagraphInput then commitParagraphInput ()
                skipParagraphInput <- false
            elif not (System.String.IsNullOrEmpty text) then
                commitInput "Typing" (insertTextForMode text handler.Document))

    // The canvas stays focusable (clicks, Tab, hosts calling canvas.focus());
    // focus is handed on to the proxy so that text input reaches it.
    listen canvas "focus" (fun (_: Event) ->
        if proxyAttached then
            syncInputProxy handler.Document
            focusWithoutScroll proxy)

    // Handle mouse down - start selection or image resize
    let onMouseDown (e: Event) =
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
    listen canvas "mousedown" (fun e ->
        // Secondary and auxiliary clicks are not selection or drag gestures.
        // In particular, leave the browser's context menu alone in a bare editor.
        if (e :?> MouseEvent).button = 0 then onMouseDown e)

    // Handle mouse move - extend selection, resize image, or update hover
    listen canvas "mousemove" (fun (e: Event) ->
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

    let guardedMouseUp event = if not disposed then onMouseUp event
    addWindowListener "mouseup" guardedMouseUp
    handler.Dispose <- fun () ->
        disposed <- true
        removeWindowListener "mouseup" guardedMouseUp
        for element, eventName, listener in listeners do
            element.removeEventListener(eventName, listener)
        listeners.Clear()
        canvas.tabIndex <- originalTabIndex
        canvas?style?cursor <- originalCursor
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

/// Run a shared command and notify the browser only when it changed something.
let executeWordCommand (handler: InputHandler) command =
    let tracking = if handler.TrackChanges then Some (handler.TrackChangesAuthor, handler.TrackChangesDate) else None
    let result = executeSessionCommand handler command tracking
    match result with
    | Ok Rbt.Word.Session.SelectionChanged | Ok Rbt.Word.Session.ContentChanged ->
        updateDocument handler handler.Document
    | _ -> ()
    result

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
        handler.History <- { handler.History with Undos = handler.Session.WrapStoryPlace(head, pageIndex, isHeader) :: rest }
    | _ -> ()
    handler.Document <- doc
    handler.OnChange doc

/// Replace the document with a newly loaded one. Undo history belongs to the
/// previous document and is discarded.
let loadDocument (handler: InputHandler) (doc: Doc.Document) =
    hideStory handler
    handler.Session.Replace doc
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
let withPendingFormat _ _ after = after
let applyEditAtStoryPlace _ _ _ _ _ = ()
let enterTextBox _ _ _ = false
let enterTextBoxOnPage _ _ _ _ = false
let selectShape _ _ = ()
let shapePlacedDocument _ _ _ _ _ _ _ _ = None
let shapeResizedDocument _ _ _ _ _ _ _ _ _ _ = None
let loadDocument _ _ = ()
let extendDocuments _ _ = ()
let undo _ = false
let redo _ = false
let canUndo _ = false
let canRedo _ = false
let dispose _ = ()
#endif
