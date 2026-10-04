module Rbt.Word.Input

let private sameListOrEqual left right =
    obj.ReferenceEquals(left, right) || left = right

let private documentsEqualForChange (left: Rbt.Word.Document.Document) (right: Rbt.Word.Document.Document) =
    obj.ReferenceEquals(left, right)
    || (left.Selection = right.Selection
        && left.PageSetup = right.PageSetup
        && sameListOrEqual left.Blocks right.Blocks
        && sameListOrEqual left.Styles right.Styles
        && sameListOrEqual left.Headers right.Headers
        && sameListOrEqual left.Footers right.Footers
        && sameListOrEqual left.HeaderParts right.HeaderParts
        && sameListOrEqual left.FooterParts right.FooterParts)

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
        && obj.ReferenceEquals(left.FooterParts, right.FooterParts))

let sameContent (left: Rbt.Word.Document.Document) (right: Rbt.Word.Document.Document) =
    sameContentObjects left right
    || (left.PageSetup = right.PageSetup
        && sameListOrEqual left.Blocks right.Blocks
        && sameListOrEqual left.Styles right.Styles
        && sameListOrEqual left.Headers right.Headers
        && sameListOrEqual left.Footers right.Footers
        && sameListOrEqual left.HeaderParts right.HeaderParts
        && sameListOrEqual left.FooterParts right.FooterParts)

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

/// Text an IME (or a dead key) is composing. It is shown underlined at the
/// caret but is not in the document until the composition ends.
type CompositionState = {
    mutable Text: string
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

/// Record a finished change in the undo history as one step. Selection-only
/// changes are not recorded.
let private recordChange (handler: InputHandler) (description: string) (beforeDoc: Doc.Document) (afterDoc: Doc.Document) =
    if not (sameContent beforeDoc afterDoc) then
        let cmd = DeleteTextCommand(beforeDoc, afterDoc, description) :> ICommand
        let _, history = Cmd.execute cmd beforeDoc handler.History
        handler.History <- history

let private showDocument (handler: InputHandler) (doc: Doc.Document) =
    handler.Document <- doc
    renderWithHover handler.RenderContext doc handler.HoveredImage
    handler.SyncInputProxy doc
    handler.OnChange doc

/// Undo the last recorded change. Returns false when there is nothing to undo.
let undo (handler: InputHandler) =
    if handler.History.Undos.IsEmpty then
        false
    else
        let doc, history = Cmd.undo handler.Document handler.History
        handler.History <- history
        showDocument handler doc
        true

/// Redo the last undone change. Returns false when there is nothing to redo.
let redo (handler: InputHandler) =
    if handler.History.Redos.IsEmpty then
        false
    else
        let doc, history = Cmd.redo handler.Document handler.History
        handler.History <- history
        showDocument handler doc
        true

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

    let isWordChar (ch: char) = System.Char.IsLetterOrDigit ch || ch = '_' || ch = '\'' || ch = '\u2019'

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
        if handler.TrackChanges then Doc.insertTrackedText handler.TrackChangesAuthor handler.TrackChangesDate text doc else Doc.insertText text doc

    let splitBlockForMode doc =
        if handler.TrackChanges then Doc.trackSplitBlock handler.TrackChangesAuthor handler.TrackChangesDate doc else Doc.splitBlock doc

    let deleteSelectionForMode doc =
        if handler.TrackChanges then Doc.trackDeleteSelection handler.TrackChangesAuthor handler.TrackChangesDate doc else Doc.deleteSelection doc

    let deleteBackwardForMode doc =
        if handler.TrackChanges then Doc.trackDeleteBackward handler.TrackChangesAuthor handler.TrackChangesDate doc else Doc.deleteBackward doc

    let deleteForwardForMode doc =
        if handler.TrackChanges then Doc.trackDeleteForward handler.TrackChangesAuthor handler.TrackChangesDate doc else Doc.deleteForward doc

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
        if not (documentsEqualForChange newDoc handler.Document) then
            handler.Document <- newDoc
            renderWithHover handler.RenderContext newDoc handler.HoveredImage
            handler.SyncInputProxy newDoc
            handler.OnChange newDoc

    // Make canvas focusable
    canvas.tabIndex <- 0

    // Handle keyboard input. The listener serves the canvas and the input
    // proxy (which normally has the focus).
    let onKeyDown (e: Event) =
        let ke = e :?> KeyboardEvent
        let mutable newDoc = handler.Document
        let mutable newHistory = handler.History
        let mutable useCommand = false
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
                let (doc', hist') = Cmd.redo handler.Document handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            else
                // Undo (Ctrl+Z)
                let (doc', hist') = Cmd.undo handler.Document handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | key when commandKey && keyEquals 'y' key ->
            // Redo (Ctrl+Y)
            let (doc', hist') = Cmd.redo handler.Document handler.History
            newDoc <- doc'
            newHistory <- hist'
            useCommand <- true
            e.preventDefault()
        | key when commandKey && keyEquals 'a' key ->
            newDoc <- Doc.selectAll handler.Document
            e.preventDefault()
        | key when commandKey && keyEquals 'b' key ->
            let beforeDoc = handler.Document
            let afterDoc = WordStyles.toggleBold handler.Document
            if not (documentsEqualForChange afterDoc beforeDoc) then
                let cmd = FormatTextCommand(beforeDoc, afterDoc, "bold") :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | key when commandKey && keyEquals 'i' key ->
            let beforeDoc = handler.Document
            let afterDoc = WordStyles.toggleItalic handler.Document
            if not (documentsEqualForChange afterDoc beforeDoc) then
                let cmd = FormatTextCommand(beforeDoc, afterDoc, "italic") :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | key when commandKey && keyEquals 'u' key ->
            let beforeDoc = handler.Document
            let afterDoc = WordStyles.toggleUnderline handler.Document
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
        | key when commandKey && keyEquals 'c' key ->
            // Copy (Ctrl+C)
            let text = Doc.getSelectionText handler.Document
            if text <> "" then
                Clipboard.copyText text |> ignore
            e.preventDefault()
        | key when commandKey && keyEquals 'x' key ->
            // Cut (Ctrl+X)
            let text = Doc.getSelectionText handler.Document
            if text <> "" then
                Clipboard.copyText text |> ignore
                // Delete selection
                let beforeDoc = handler.Document
                let afterDoc = deleteSelectionForMode handler.Document
                if not (documentsEqualForChange afterDoc beforeDoc) then
                    let cmd = DeleteTextCommand(beforeDoc, afterDoc, "Cut") :> ICommand
                    let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                    newDoc <- doc'
                    newHistory <- hist'
                    useCommand <- true
            e.preventDefault()
        | key when commandKey && keyEquals 'v' key ->
            // Paste (Ctrl+V)
            e.preventDefault()
            Clipboard.pasteText().``then``(fun text ->
                if text <> "" then
                    let beforeDoc = handler.Document
                    let mutable doc' = beforeDoc

                    // Delete selection first if any
                    let hasSelection =
                        match doc'.Selection with
                        | Some sel when sel.Anchor <> sel.Focus -> true
                        | _ -> false

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

                    if not (documentsEqualForChange doc' beforeDoc) then
                        let cmd = DeleteTextCommand(beforeDoc, doc', "Paste") :> ICommand
                        let doc'', hist' = Cmd.execute cmd beforeDoc handler.History
                        handler.Document <- doc''
                        handler.History <- hist'
                        renderWithHover handler.RenderContext doc'' handler.HoveredImage
                        handler.OnChange doc''
            ) |> ignore
        | "Backspace" ->
            let beforeDoc = handler.Document
            let afterDoc = deleteBackwardForMode handler.Document
            if not (documentsEqualForChange afterDoc beforeDoc) then
                let cmd = DeleteTextCommand(beforeDoc, afterDoc, "Delete backward") :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | "Delete" ->
            let beforeDoc = handler.Document
            let afterDoc = deleteForwardForMode handler.Document
            if not (documentsEqualForChange afterDoc beforeDoc) then
                let cmd = DeleteTextCommand(beforeDoc, afterDoc, "Delete forward") :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
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
            // last cell adds a row), as in Word; elsewhere they indent.
            match Doc.moveToAdjacentTableCell (not ke.shiftKey) handler.Document with
            | Some moved -> newDoc <- moved
            | None ->
                if ke.shiftKey then
                    // Shift+Tab: Decrease indent (in the first cell of a table: nothing)
                    if not (Doc.caretContext handler.Document).InTable then newDoc <- Doc.decreaseIndent handler.Document
                else
                    // Tab: Increase indent
                    newDoc <- Doc.increaseIndent handler.Document
            e.preventDefault()
        | "Enter" ->
            // Check if we're on an empty list item - if so, exit list
            match handler.Document.Selection with
            | Some sel ->
                match Doc.findBlock handler.Document sel.Focus.BlockId with
                | Some block ->
                    let para = Doc.getParagraph block
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
                let cmd = DeleteTextCommand(beforeDoc, afterDoc, sprintf "Insert '%s'" key) :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | _ -> ()

        // Any other key ends a run of vertical moves: the next one starts from where the caret then is.
        if not keepGoal && not composing && ke.key <> "Shift" && ke.key <> "Control" && ke.key <> "Alt" && ke.key <> "Meta" then handler.GoalX <- None
        ignore goalBefore
        if useCommand then
            // Also when the document compares equal (undoing a no-op step
            // must still pop it).
            handler.History <- newHistory
        if not (documentsEqualForChange newDoc handler.Document) then
            if not useCommand then
                // Shortcuts that changed the document without building a
                // command (heading level, Tab indent, leaving a list) are
                // undo steps too.
                recordChange handler "Edit" handler.Document newDoc
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
            let inserted = Doc.insertText text handler.Document
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

    /// Text that arrived through the proxy becomes one undo step.
    let commitInput description (afterDoc: Doc.Document) =
        let beforeDoc = handler.Document
        if not (documentsEqualForChange afterDoc beforeDoc) then
            recordChange handler description beforeDoc afterDoc
            handler.Document <- afterDoc
            renderWithHover handler.RenderContext afterDoc handler.HoveredImage
            syncInputProxy afterDoc
            handler.OnChange afterDoc
        else
            // Take a composition preview off the canvas again.
            renderWithHover handler.RenderContext beforeDoc handler.HoveredImage
            syncInputProxy beforeDoc

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
                    commitInput "Typing" (insertTextForMode data handler.Document)
            | "deleteContentBackward" ->
                e.preventDefault()
                commitInput "Delete" (deleteBackwardForMode handler.Document)
            | "deleteContentForward" ->
                e.preventDefault()
                commitInput "Delete" (deleteForwardForMode handler.Document)
            | _ -> ())

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

        if handler.IsMouseDown then
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
        if handler.IsMouseDown || handler.ResizingImage.IsSome || handler.DraggingImage.IsSome then
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

/// Replace the document with a newly loaded one. Undo history belongs to the
/// previous document and is discarded.
let loadDocument (handler: InputHandler) (doc: Doc.Document) =
    handler.History <- Cmd.empty
    handler.PointerStartDocument <- None
    handler.IsMouseDown <- false
    handler.ResizingImage <- None
    handler.DraggingImage <- None
    handler.SelectionAnchor <- None
    updateDocument handler doc

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
let loadDocument _ _ = ()
let undo _ = false
let redo _ = false
let canUndo _ = false
let canRedo _ = false
let dispose _ = ()
#endif
