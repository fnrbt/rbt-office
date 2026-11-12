module Rbt.Word.Renderer

#if FABLE_COMPILER
open System
open Fable.Core
open Browser.Types
open Rbt.Word.Document
open Rbt.Word.TextBuffer
open Rbt.Word.Layout
open Rbt.Word.Paint

/// What the canvas shows of the document when the host scrolls and zooms it
/// (see `setViewport`). All lengths are CSS pixels; the document itself is
/// laid out in layout pixels (CSS pixels at 100% zoom).
type Viewport = {
    /// Zoom: CSS pixels per layout pixel.
    Scale: float
    /// Backing-store pixels per CSS pixel.
    DevicePixelRatio: float
    /// Scroll position of the scrollable content the document sits in.
    ScrollLeft: float
    ScrollTop: float
    /// Size of the visible area; the canvas is made exactly this large.
    Width: float
    Height: float
    /// Where the document's top-left corner sits inside the scrollable
    /// content (a host margin, or the offset that centres the pages).
    OriginX: float
    OriginY: float
}

type PaintFrame = {
    Document: Document
    Layouts: BlockLayout list
    /// Geometry of every page frame, as the layout pass produced it (sections
    /// can differ in size and orientation), extended over the note areas.
    Pages: PageInfo list
    /// The body's paint commands in chunks with their vertical extents (a
    /// block, the note areas, a page's header and footer; see
    /// `Paint.collectDocumentChunks`), in paint order. The chunks of what an
    /// edit did not change are the chunks of the frame before.
    Body: PaintChunk array
    /// Overlay that follows the document's content only (comment anchors):
    /// kept with the body commands when just the selection changes.
    StaticOverlay: PaintCommand list
    /// Overlay that follows the selection and the pointer: selection
    /// highlight, caret, the frame of the caret's content control, the hover
    /// outline. Painted after the static overlay.
    OverlayCommands: PaintCommand list
    /// The band of the document the overlay's selection highlight was made
    /// for (None: the whole document).
    OverlayBand: (float * float) option
    CanvasWidth: float
    CanvasHeight: float
    FontFamily: string
    FontSize: float
    LineHeight: float
    PaddingLeft: float
    PaddingTop: float
    /// The text path painter and the text measurers the frame was built
    /// with, by the context's version counters (see RenderContext).
    PainterVersion: int
    MeasurerVersion: int
    /// The pass of the layout state the frame's layouts are from (see
    /// IncrementalLayoutState.LayoutSerial): a frame made while the start of a
    /// long document was laid out is not the frame of the layout that has more.
    LayoutSerial: int
    /// How many of the chunks of Body - the first ones - are the pages'
    /// headers and footers: while a header is edited in place the body is
    /// dimmed below them.
    FurnitureCount: int
    /// The hover outline the overlay commands were built for.
    Hovered: (string * int) option
    /// The viewport the frame was painted through (None: whole-canvas mode).
    Viewport: Viewport option
    mutable StaticOverlayIndex: PaintChunk option
    /// The first layout of each block, for the overlay of the next selection
    /// (made when a frame is first reused for another selection).
    mutable LayoutById: System.Collections.Generic.Dictionary<string, BlockLayout> option
    /// The band the body's pictures are in (kept: found by walking every command).
    mutable ImageBand: Choice<(float * float) option, unit> option
}

/// A long document is laid out from its start, in slices (see
/// `beginProgressiveLayout`, `scheduleLayout`): what a context remembers of it.
type LayoutProgress = {
    /// A slice is waiting for an idle moment.
    mutable Scheduled: bool
    /// Raised to call off the slices that are waiting (the context is disposed of).
    mutable Epoch: int
    /// Top-level blocks the next slice adds: what the last ones say fits a slice's time.
    mutable Chunk: int
    /// When the hosts' listeners were last told (they are not told with every slice).
    mutable Notified: float
    /// Called once the whole document is laid out (whenLayoutComplete).
    Waiters: ResizeArray<unit -> unit>
    /// The document on display is still being read (a package opened front
    /// first): what is laid out, even all of what is there, is not the whole
    /// document yet.
    mutable LoadPending: bool
    /// How the slices went (see `renderStats`): how many there were, how long the longest took
    /// with what followed it (the host's repaint and its listeners), which it was, and what it was
    /// made of: its passes, the layout's part of its time, the block limit it started from.
    mutable Slices: int
    mutable LongestSliceMs: float
    mutable LongestSlice: int
    mutable LongestSlicePasses: int
    mutable LongestSliceLayoutMs: float
    mutable LongestSliceFrom: int
    /// Passes of the slice in hand.
    mutable SlicePasses: int
    /// Most of a long document was changed at once (everything made bold):
    /// it is laid out again from its start as a newly loaded one is - what is
    /// in view first, the rest in slices (see `deferMassRelayout`).
    mutable MassChange: bool
    /// What the view is held by while the layout above it changes: a block
    /// (its place among the blocks), which of its layouts, and the y that
    /// layout had. Once the block is laid out again the viewport is moved by
    /// as much as the layout moved (see `captureAnchor`, `applyAnchor`).
    mutable Anchor: (int * int * float) option
    /// The viewport was moved to stay on the text it shows (its ScrollTop is
    /// not the one the host set last): the host is told, see OnScrollAdjusted.
    mutable ScrollAdjusted: bool
    /// Called when the viewport was moved (the host scrolls its content to it).
    mutable OnScrollAdjusted: unit -> unit
    /// The layout of the document itself took the place of the one that was
    /// shown meanwhile (see StaleLayout): what is on the canvas is painted again.
    mutable Adopted: bool
}

/// Most of a long document changed at once while the view is far down it
/// (everything made bold on page 300 of 600). Laying it out again from its
/// start as far as the view takes as long as half the document takes; the
/// blocks above the view are not seen. So the layout that is shown is of
/// another list of blocks: the first `Keep` blocks as they were before the
/// change (`StaleBlocks` - their layouts stand, not one is made again), and
/// the document's own blocks from there on, laid out from the place the pass
/// had before block `Keep`. Meanwhile the document itself is laid out from
/// its start in slices, in a layout state of its own (`Ahead`); when that
/// reaches what the view shows it takes the place of the context's layout
/// state, and the viewport is moved by as much as the first block in view
/// moved, so that the text in view stays where it is on screen. Until then
/// the layout is not complete, and the pages above the view are those the
/// document had before the change.
type StaleLayout = {
    mutable StaleBlocks: Block list
    mutable Keep: int
    /// The blocks of the document the view is of.
    mutable Target: Block list
    /// The blocks that are laid out for a list of the document's blocks (kept: the same list for the same document).
    mutable HybridFor: Block list
    mutable HybridKeep: int
    mutable Hybrid: Block list
    Ahead: IncrementalLayoutState
}

/// A header or footer that is edited in place. The document the editor
/// works on is then the story's own (its blocks are the header's blocks, its
/// selection is in them); the renderer shows the whole document, dimmed,
/// with the story's selection and caret where the page paints that header.
type StoryView = {
    /// The whole document (body, headers and footers - with the story's
    /// blocks where they belong - and no selection) for a story document.
    Whole: Document -> Document
    PageIndex: int
    IsHeader: bool
    /// The name on the boundary ("Header", "First Page Footer", ...).
    Label: string
    /// The story is the text of this text box of the body, not a header or
    /// footer: the body is not dimmed, and the story's layouts are those of
    /// the box's text where the page paints the box.
    Box: BlockId option
    /// The story is the body of this note or comment ("footnote" | "endnote" |
    /// "comment", id): the body is not dimmed; the story is laid out where the
    /// page shows the note (a comment: in a balloon below its anchor) and
    /// painted there over what the page shows of it.
    Note: (string * string) option
}

/// Rendering context
type RenderContext = {
    Canvas: HTMLCanvasElement
    Ctx: CanvasRenderingContext2D
    FontFamily: string
    FontSize: float
    LineHeight: float
    PaddingLeft: float
    PaddingTop: float
    /// Map of image source -> decoded HTMLImageElement. Images are drawn
    /// synchronously, so a host decodes them first and puts them here.
    mutable ImageCache: obj option
    mutable TextMeasurer: TextMeasurer option
    mutable TextRangeMeasurer: TextRangeMeasurer option
    mutable TextPathPainter: (string -> TextPaintStyle -> float -> float -> float -> bool) option
    /// Makes a text path painter for a canvas context: how text is painted on
    /// canvases other than the context's own (`renderPage` targets). A
    /// painter draws on one context; this makes one for any.
    /// (Delegates, not curried functions: the factory and the painter it
    /// returns come from JavaScript.)
    mutable TextPathPainterFactory: System.Func<CanvasRenderingContext2D, System.Func<string, TextPaintStyle, float, float, float, bool>> option
    /// Count the times a text path painter / the text measurers were
    /// installed or cleared (`setTextPathPainter`, `setTextMeasurers`). What
    /// was painted or laid out with an earlier one is stale; nothing else
    /// about a painter or measurer is compared - a function handed over from
    /// JavaScript is a different object every time it is looked at.
    mutable TextPathPainterVersion: int
    mutable TextMeasurerVersion: int
    /// Only the current block snapshot is retained, not an index per undo entry.
    mutable ControlMemberships: (Block list * System.Collections.Generic.Dictionary<BlockId, System.Collections.Generic.HashSet<string>>) option
    /// How many paint frames were built (the commands of the whole document
    /// collected) - for hosts and checks that want to see frames reused.
    mutable FrameBuilds: int
    /// Of the last frame built: the top-level blocks that were laid out, and
    /// those whose paint commands were collected again (-1: all of them, the
    /// document was collected as a whole).
    mutable LastLaidOutBlocks: int
    mutable LastCollectedBlocks: int
    /// Paragraphs whose paint commands were collected when a band that
    /// reaches them was first painted (not when their frame was built).
    mutable CollectedOnPaint: int
    /// Stands for the text measurers the context lays out with (a new object
    /// whenever they change): table plans are kept per scope, see
    /// `Layout.withTablePlanScope`.
    mutable PlanScope: obj
    mutable LayoutState: IncrementalLayoutState
    /// The blocks above the view are shown as they were laid out before a
    /// change of most of the document (see StaleLayout).
    mutable Stale: StaleLayout option
    /// The paint commands of the last frame, block by block (see `Paint.PaintCache`).
    PaintCache: PaintCache
    /// Which blocks of the last frame's document hold comments (parallel arrays).
    mutable CommentBlocks: Block array
    mutable CommentFlags: bool array
    mutable LastPaint: PaintFrame option
    /// None: the canvas holds the whole document at layout scale and the host
    /// sizes it (from pageFrames). Some: the canvas is the visible viewport,
    /// sized and painted by the renderer.
    mutable Viewport: Viewport option
    /// Frame kept for renderPage when the document is not the one on the canvas.
    mutable PageFrame: PaintFrame option
    Progress: LayoutProgress
    /// The header or footer being edited in place, if any (see StoryView).
    mutable Story: StoryView option
    /// The last paint showed a story (the next one repaints what is in view).
    mutable StoryPainted: bool
    /// The boundary, selection and caret of the story on display.
    mutable StoryOverlay: PaintCommand list
    /// The story's layouts, for the blocks and the place they were made for.
    mutable StoryLayouts: (Block list * float * int * BlockLayout list) option
    /// The turn of the text box that is edited in place, when it is turned:
    /// degrees clockwise and the centre it is turned about (layout
    /// coordinates). Its text is laid out upright where the box is and
    /// painted turned; its caret, selection and frame are drawn turned too,
    /// and a point of the canvas is turned back before it is looked up in the
    /// layouts (set by storyLayouts).
    mutable StoryTurn: (float * float * float) option
    /// Where the note or comment that is edited in place was last shown: the
    /// left edge, top and width of its text (it is shown there while the page
    /// shows it nowhere - a note whose text was all deleted).
    mutable StoryNotePlace: (float * float * float) option
    /// The layouts of the footnote or endnote that is edited in place, and
    /// the layouts of the whole document they were taken from.
    mutable StoryNoteLayouts: (BlockLayout list * BlockLayout list) option
    /// The page frames of a layout pass (by its serial): a host asks for them with every change.
    mutable KnownPageFrames: (int * PageInfo list) option
    /// The shape that is selected (see Input.selectShape): its frame and its
    /// eight handles are painted over the page.
    mutable SelectedShape: BlockId option
}

type CommentAnchor = {
    Id: string
    Text: string
    Author: string option
    Date: string option
    Excerpt: string
    BlockId: string
    StartOffset: int
    EndOffset: int
    X: float
    Y: float
    Width: float
    Height: float
    CalloutX: float
    CalloutY: float
    IsSelected: bool
}

[<Emit("$0.setLineDash($1)")>]
let private setLineDash (_ctx: CanvasRenderingContext2D) (_dash: float array) : unit = jsNative

type FootnoteAnchor = {
    Id: string
    Text: string
    PageIndex: int
}

type EndnoteAnchor = {
    Id: string
    Text: string
}

/// Create render context
let createContext (canvas: HTMLCanvasElement) = {
    Canvas = canvas
    Ctx = canvas.getContext_2d()
    FontFamily = "Calibri"
    FontSize = 16.0
    LineHeight = 24.0
    PaddingLeft = 20.0
    PaddingTop = 20.0
    ImageCache = None
    TextMeasurer = None
    TextRangeMeasurer = None
    TextPathPainter = None
    TextPathPainterFactory = None
    TextPathPainterVersion = 0
    TextMeasurerVersion = 0
    ControlMemberships = None
    FrameBuilds = 0
    LastLaidOutBlocks = 0
    LastCollectedBlocks = 0
    CollectedOnPaint = 0
    PlanScope = obj ()
    LayoutState = Rbt.Word.Layout.createIncrementalLayoutState ()
    Stale = None
    PaintCache = createPaintCache ()
    CommentBlocks = [||]
    CommentFlags = [||]
    LastPaint = None
    Viewport = None
    PageFrame = None
    Progress = { Anchor = None; ScrollAdjusted = false; OnScrollAdjusted = ignore; Adopted = false; MassChange = false; Scheduled = false; Epoch = 0; Chunk = 48; Notified = 0.0; Waiters = ResizeArray(); LoadPending = false
                 Slices = 0; LongestSliceMs = 0.0; LongestSlice = 0; LongestSlicePasses = 0; LongestSliceLayoutMs = 0.0; LongestSliceFrom = 0; SlicePasses = 0 }
    Story = None
    StoryPainted = false
    StoryOverlay = []
    StoryLayouts = None
    StoryTurn = None
    StoryNotePlace = None
    StoryNoteLayouts = None
    KnownPageFrames = None
    SelectedShape = None
}

let private cssFontFamily (family: string) =
    if family.Contains(" ") then sprintf "\"%s\"" (family.Replace("\"", "")) else family

let private browserMetricScale (family: string) (style: TextBuffer.TextStyle) =
    // Chromium canvas metrics for loaded Windows faces diverge from Word/PDF
    // shaping enough to change centering and wrap decisions in real DOCX bodies.
    match family.Trim().ToLowerInvariant() with
    | "cambria" when not style.Bold && not style.Italic -> 0.982
    | "times new roman" when style.Bold && style.Italic -> 0.945
    | "times new roman" when not style.Bold && not style.Italic -> 1.0185
    | _ -> 1.0

/// Get image from cache (JS Map.get())
[<Emit("$0.get($1)")>]
let private mapGet (map: obj) (key: string) : obj = jsNative

/// The canvas font of a text style and the scale its measured widths are
/// corrected by: made once per style object (a run's style is asked for with
/// every word of the run).
type private MeasureFont = { Font: string; Scale: float }

[<Emit("new WeakMap()")>]
let private newStyleMap () : obj = jsNative
[<Emit("$0.get($1)")>]
let private styleMapGet (_map: obj) (_style: TextBuffer.TextStyle) : MeasureFont = jsNative
[<Emit("$0.set($1, $2)")>]
let private styleMapSet (_map: obj) (_style: TextBuffer.TextStyle) (_font: MeasureFont) : unit = jsNative
let private measureFonts = newStyleMap ()

/// Widths the canvas gave, by font and text. A document's words repeat, and
/// measuring every word of every run was most of the time of a first layout.
/// (A width depends on the font string and the text alone.)
let private measuredWidths = System.Collections.Generic.Dictionary<string, System.Collections.Generic.Dictionary<string, float>>()
let mutable private measuredWidthCount = 0

/// Forget the measured widths: the fonts the canvas measures with changed.
let clearMeasuredWidths () =
    measuredWidths.Clear()
    measuredWidthCount <- 0

[<Emit("((typeof document !== 'undefined' && document.fonts && document.fonts.addEventListener) ? document.fonts.addEventListener('loadingdone', $0) : undefined)")>]
let private onFontsLoaded (_callback: obj -> unit) : unit = jsNative
// A font face that finished loading measures differently from the fallback
// that stood in for it.
onFontsLoaded (fun _ -> clearMeasuredWidths ())

/// Measure text width with style
let measureTextStyled (ctx: RenderContext) (text: string) (style: TextBuffer.TextStyle) =
    let font =
        let known = styleMapGet measureFonts style
        if not (isNull (box known)) && (style.FontFamily.IsSome || known.Font.EndsWith (cssFontFamily ctx.FontFamily)) then known
        else
            let fontStyle = if style.Bold && style.Italic then "bold italic"
                            elif style.Bold then "bold"
                            elif style.Italic then "italic"
                            else ""
            let fontFamily = style.FontFamily |> Option.defaultValue ctx.FontFamily
            let fontSize = if style.VertAlign.IsSome then style.FontSize * 0.7 else style.FontSize
            let made = { Font = sprintf "%s %fpx %s" fontStyle fontSize (cssFontFamily fontFamily); Scale = browserMetricScale fontFamily style }
            styleMapSet measureFonts style made
            made
    let widths =
        match measuredWidths.TryGetValue font.Font with
        | true, widths -> widths
        | false, _ ->
            let widths = System.Collections.Generic.Dictionary<string, float>()
            measuredWidths.[font.Font] <- widths
            widths
    let width =
        match widths.TryGetValue text with
        | true, width -> width
        | false, _ ->
            ctx.Ctx.font <- font.Font
            let width = ctx.Ctx.measureText(text).width
            // (Bounded: long texts are measured as they come, and the lot is dropped when it grows large.)
            if text.Length <= 48 then
                if measuredWidthCount > 200000 then
                    measuredWidths.Clear()
                    measuredWidthCount <- 0
                    measuredWidths.[font.Font] <- widths
                    widths.Clear()
                widths.[text] <- width
                measuredWidthCount <- measuredWidthCount + 1
            width
    (width * font.Scale) + TextBuffer.characterSpacingAdvance text style

let layoutSettings (ctx: RenderContext) (doc: Document) =
    documentSettings ctx.FontSize ctx.LineHeight 3.5 doc

let measureWithContext ctx text style =
    match ctx.TextMeasurer with
    | Some measure -> measure text style
    | None -> measureTextStyled ctx text style

let measureRangeWithContext ctx text style start length =
    match ctx.TextRangeMeasurer with
    | Some measureRange -> measureRange text style start length
    | None -> rangeMeasurerFromTextMeasurer (measureWithContext ctx) text style start length

// ---- a long document is laid out from its start ------------------------------
//
// The layout state lays out the first BlockLimit top-level blocks (see
// Layout.IncrementalLayoutState). A newly loaded document starts with a few
// (`beginProgressiveLayout`), a render raises the limit until what is in view
// is laid out (`coverViewport`, `coverSelection`), and idle slices raise it to
// the end (`scheduleLayout`). Whatever needs the layout of the whole document
// asks for it (`completeLayout`): the result is then the layout a pass over
// the whole document makes.

[<Emit("performance.now()")>]
let private nowMs () : float = jsNative

/// Whether the layout the context has (or makes next) is of all the blocks the document has.
let private layoutIsWhole (ctx: RenderContext) =
    ctx.Stale.IsNone && Rbt.Word.Layout.incrementalLayoutIsWhole ctx.LayoutState

/// Whether the layout state on display has all the blocks it is given (while
/// blocks above the view are stale that is not the layout of the document).
let private stateIsWhole (ctx: RenderContext) =
    Rbt.Word.Layout.incrementalLayoutIsWhole ctx.LayoutState

/// Whether the layout the context has (or makes next) is of the whole
/// document: all its blocks, and none still to be read.
let layoutIsComplete (ctx: RenderContext) =
    layoutIsWhole ctx && not ctx.Progress.LoadPending

/// More of the document on display was read: `blocksBefore` blocks were
/// there. A layout that covered them all does not cover the new ones at
/// once - they are laid out as the rest of a long document is, in slices or
/// when something shows them. True when what was laid out was all there was
/// (the view may reach further than that: the host repaints).
let documentExtended (ctx: RenderContext) (blocksBefore: int) : bool =
    if ctx.LayoutState.BlockLimit = System.Int32.MaxValue || ctx.LayoutState.BlockLimit >= blocksBefore then
        ctx.LayoutState.BlockLimit <- max 1 blocksBefore
        true
    else false

let private resolveLayoutWaiters (ctx: RenderContext) =
    if ctx.Progress.Waiters.Count > 0 && layoutIsComplete ctx then
        let waiters = ctx.Progress.Waiters.ToArray()
        ctx.Progress.Waiters.Clear()
        for waiter in waiters do waiter ()

/// The layout that holds the view where it is: the block at the top of the
/// view, the layout of it that is there, and that layout's y.
let private captureAnchor (ctx: RenderContext) =
    match ctx.Viewport, ctx.LayoutState.LastSlices with
    | Some vp, Some slices when slices.SliceLayouts.Length > 0 && ctx.Progress.Anchor.IsNone ->
        let top = (vp.ScrollTop - vp.OriginY) / vp.Scale
        let layouts = slices.SliceLayouts
        // The last layout that starts at or above the top of the view (the next one when that ends above it).
        let mutable lo = 0
        let mutable hi = layouts.Length
        while lo < hi do
            let mid = lo + (hi - lo) / 2
            if layouts.[mid].Y <= top then lo <- mid + 1 else hi <- mid
        let mutable at = max 0 (lo - 1)
        if layouts.[at].Y + layouts.[at].Height <= top && at + 1 < layouts.Length then at <- at + 1
        let starts = slices.SliceStarts
        let mutable first = 0
        let mutable last = slices.SliceBlocks.Length - 1
        while first < last do
            let mid = first + (last - first + 1) / 2
            if starts.[mid] <= at then first <- mid else last <- mid - 1
        if first >= 0 && first < slices.SliceBlocks.Length then
            ctx.Progress.Anchor <- Some (first, at - starts.[first], layouts.[at].Y)
    | _ -> ()

/// The view is moved by as much as the layout that holds it moved (once that
/// block is laid out again).
let private applyAnchor (ctx: RenderContext) =
    match ctx.Progress.Anchor with
    | None -> ()
    | Some (block, ordinal, y) ->
        match ctx.Viewport, ctx.LayoutState.LastSlices with
        | Some vp, Some slices when block < slices.SliceBlocks.Length ->
            ctx.Progress.Anchor <- None
            let first = slices.SliceStarts.[block]
            let last = slices.SliceStarts.[block + 1]
            if last > first then
                let delta = slices.SliceLayouts.[min (first + ordinal) (last - 1)].Y - y
                if abs delta >= 0.5 then
                    ctx.Viewport <- Some { vp with ScrollTop = max 0.0 (vp.ScrollTop + delta * vp.Scale) }
                    ctx.Progress.ScrollAdjusted <- true
                    ctx.Progress.OnScrollAdjusted ()
        | Some _, Some _ -> ()
        | _ -> ctx.Progress.Anchor <- None

/// The layout of the document itself, as far as it is, takes the place of
/// the layout with stale blocks above the view (see StaleLayout).
let private adoptAhead (ctx: RenderContext) =
    match ctx.Stale with
    | Some stale ->
        captureAnchor ctx
        stale.Ahead.LayoutSerial <- (max stale.Ahead.LayoutSerial ctx.LayoutState.LayoutSerial) + 1
        ctx.LayoutState <- stale.Ahead
        ctx.Stale <- None
        ctx.KnownPageFrames <- None
        ctx.Progress.Adopted <- true
    | None -> ()

/// The blocks that are laid out for the document's blocks while the blocks
/// above the view are stale: the first Keep stale ones, the document's own
/// after them. A change of the document meanwhile: one above the view lowers
/// Keep to where it is; another change of most of the document (an undo)
/// makes what is shown the stale layout, and the layout of the document
/// itself starts again.
let private hybridBlocks (stale: StaleLayout) (blocks: Block list) : Block list =
    if not (obj.ReferenceEquals(stale.Target, blocks)) then
        let mutable before = stale.Target
        let mutable after = blocks
        let mutable prefix = 0
        while prefix < stale.Keep && not before.IsEmpty && not after.IsEmpty && obj.ReferenceEquals(before.Head, after.Head) do
            before <- before.Tail
            after <- after.Tail
            prefix <- prefix + 1
        if prefix < stale.Keep then
            let mutable probed = 0
            let mutable differing = 0
            while probed < 16 && not before.IsEmpty && not after.IsEmpty do
                if not (obj.ReferenceEquals(before.Head, after.Head)) then differing <- differing + 1
                before <- before.Tail
                after <- after.Tail
                probed <- probed + 1
            if probed = 16 && differing = 16 && not stale.Hybrid.IsEmpty then
                stale.StaleBlocks <- stale.Hybrid
                // (From its start again, a few blocks at first: none of them is laid out as it is now.)
                stale.Ahead.BlockLimit <- 6
            else stale.Keep <- prefix
        stale.Target <- blocks
    if obj.ReferenceEquals(stale.HybridFor, blocks) && stale.HybridKeep = stale.Keep && not stale.Hybrid.IsEmpty then stale.Hybrid
    else
        let head = ResizeArray<Block>()
        let mutable old = stale.StaleBlocks
        let mutable rest = blocks
        while head.Count < stale.Keep && not old.IsEmpty && not rest.IsEmpty do
            head.Add old.Head
            old <- old.Tail
            rest <- rest.Tail
        stale.Keep <- head.Count
        let mutable result = rest
        for index in head.Count - 1 .. -1 .. 0 do
            result <- head.[index] :: result
        stale.HybridFor <- blocks
        stale.HybridKeep <- stale.Keep
        stale.Hybrid <- result
        result

/// The document the context's layout state lays out for a document on
/// display: the document, or - while the blocks above the view are stale -
/// the one with those blocks as they were (see StaleLayout).
let private layoutTarget (ctx: RenderContext) (doc: Document) : Document =
    match ctx.Stale with
    | Some _ when ctx.Story.IsSome ->
        // (A header or footer edited in place is found by its page: the layout of the document itself.)
        adoptAhead ctx
        doc
    | Some stale ->
        let blocks = hybridBlocks stale doc.Blocks
        if stale.Keep = 0 then
            // (Nothing stale is left above the view: the layout is the document's.)
            ctx.Stale <- None
            doc
        elif obj.ReferenceEquals(blocks, doc.Blocks) then doc
        else { doc with Blocks = blocks }
    | None -> doc

/// Layouts plus the per-page geometry of the same (incremental) layout pass:
/// of the blocks the layout state's limit covers.
let layoutDocumentWithPages (ctx: RenderContext) (doc: Document) =
    let target = layoutTarget ctx doc
    let result =
        Rbt.Word.Layout.withTablePlanScope ctx.PlanScope (fun () ->
            Rbt.Word.Layout.layoutDocumentPagesIncremental ctx.LayoutState (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) target)
    applyAnchor ctx
    resolveLayoutWaiters ctx
    result

/// The document whose body is laid out and painted for a document on display:
/// the document itself, or the whole document a story (a header or footer
/// edited in place) belongs to.
let bodyDocument (ctx: RenderContext) (doc: Document) =
    match ctx.Story with
    | Some story -> story.Whole doc
    | None -> doc

/// The pages a document shows for the pages of a layout pass: those, and the
/// pages of the notes after them (`pageFramesForLayouts`, defined further down).
let private framePages : (RenderContext -> Document -> BlockLayout list -> PageInfo list -> PageInfo list) ref = ref (fun _ _ _ pages -> pages)

/// The shapes of a page's header or footer where the page paints them (see Paint.storyShapeRects).
let private pageFurnitureShapeRects (ctx: RenderContext) (whole: Document) (bodyLayouts: BlockLayout list) (pages: PageInfo list) (pageIndex: int) (isHeader: bool) : (TextBoxBlock * float * float * float * float) list =
    let settings = layoutSettings ctx whole
    match List.tryItem pageIndex pages with
    | None -> []
    | Some page ->
        let sectionEnds = bodyLayouts |> List.filter (fun layout -> endsSection layout.Metadata)
        match headerFooterTargetForPage settings whole sectionEnds pages pageIndex isHeader with
        | Some target when target.TargetPart |> Option.exists (fun part -> part.Blocks |> List.exists (function TextBox _ -> true | _ -> false)) ->
            storyShapeRects settings (measureWithContext ctx) (measureRangeWithContext ctx) whole page target.TargetY target.TargetPart.Value.Blocks
        | _ -> []

/// Where the body of a note or comment is shown while it is edited in place
/// (see StoryView.Note), by the layouts of the whole document: the left edge,
/// top and width of its text (defined further down).
let private notePlace : (RenderContext -> Document -> BlockLayout list -> string * string -> (float * float * float) option) ref = ref (fun _ _ _ _ -> None)

/// The layouts of a footnote's or endnote's paragraphs where the note area
/// shows the note (the first place, for a note shown twice), with the note's
/// own block ids: a paragraph the area breaks between two pages is one layout
/// with the lines of both.
let private noteLayoutsShown (layouts: BlockLayout list) (kind: string) (id: string) : BlockLayout list =
    let found = ResizeArray<BlockLayout>()
    let mutable slot = None
    for layout in layouts do
        match noteOfLayoutId layout.BlockId with
        | Some (layoutKind, layoutId) when layoutKind = kind && layoutId = id ->
            let layoutSlot = noteAreaSlotOfLayoutId layout.BlockId
            if slot.IsNone then slot <- Some layoutSlot
            if slot = Some layoutSlot then
                match noteParagraphIdOfLayoutId layout.BlockId with
                | Some own ->
                    if found.Count > 0 && found.[found.Count - 1].BlockId = own then
                        let before = found.[found.Count - 1]
                        found.[found.Count - 1] <- { before with Lines = before.Lines @ layout.Lines; Height = layout.Y + layout.Height - before.Y }
                    else found.Add { layout with BlockId = own }
                | None -> ()
        | _ -> ()
    List.ofSeq found

/// The layouts of a story's blocks where its page paints them.
let private storyLayouts (ctx: RenderContext) (story: StoryView) (doc: Document) : BlockLayout list =
    let whole = story.Whole doc
    let settings = layoutSettings ctx whole
    let bodyLayouts, layoutPages = layoutDocumentWithPages ctx whole
    match story.Note with
    | Some (kind, id) when kind <> "comment" ->
        // A footnote or an endnote: its paragraphs where the note area shows them (the layout lays
        // the note out from the story's blocks, see Doc.withStoryBody), with their own block ids.
        ctx.StoryTurn <- None
        match ctx.StoryNoteLayouts with
        | Some (source, layouts) when obj.ReferenceEquals(source, bodyLayouts) -> layouts
        | _ ->
            let layouts = noteLayoutsShown bodyLayouts kind id
            ctx.StoryNoteLayouts <- Some (bodyLayouts, layouts)
            layouts
    | Some note ->
        // A comment: its paragraphs in a balloon below its anchor (as it was last shown, while it shows it nowhere).
        ctx.StoryTurn <- None
        match notePlace.Value ctx whole bodyLayouts note with
        | Some place -> ctx.StoryNotePlace <- Some place
        | None -> ()
        match ctx.StoryNotePlace with
        | None -> []
        | Some (left, top, width) ->
            let key = top + left / 1000000.0 + width / 1000000000.0
            match ctx.StoryLayouts with
            | Some (blocks, known, version, layouts) when obj.ReferenceEquals(blocks, doc.Blocks) && known = key && version = ctx.TextMeasurerVersion -> layouts
            | _ ->
                let layouts = noteStoryLayouts settings (measureWithContext ctx) (measureRangeWithContext ctx) whole left top width doc.Blocks
                ctx.StoryLayouts <- Some (doc.Blocks, key, ctx.TextMeasurerVersion, layouts)
                layouts
    | None ->
    let pages = framePages.Value ctx whole bodyLayouts layoutPages
    match story.Box with
    | Some boxId ->
        // A text box: its text where the box is painted.
        match bodyShapeRects settings (measureWithContext ctx) (measureRangeWithContext ctx) whole bodyLayouts pages |> List.tryFind (fun (textBox, _, _, _, _) -> textBox.Id = boxId)
              // (A text box of a header or footer: where the story's page paints it.)
              |> Option.orElseWith (fun () -> pageFurnitureShapeRects ctx whole bodyLayouts pages story.PageIndex story.IsHeader |> List.tryFind (fun (textBox, _, _, _, _) -> textBox.Id = boxId)) with
        | None ->
            ctx.StoryTurn <- None
            []
        | Some (textBox, x, y, width, height) ->
            ctx.StoryTurn <- (if abs textBox.Rotation >= 0.01 then Some (textBox.Rotation, x + width / 2.0, y + height / 2.0) else None)
            let key = y + x / 1000000.0 + height / 1000000000.0
            match ctx.StoryLayouts with
            | Some (blocks, known, version, layouts) when obj.ReferenceEquals(blocks, doc.Blocks) && known = key && version = ctx.TextMeasurerVersion -> layouts
            | _ ->
                let shown = { textBox with Content = Rbt.Word.Document.paragraphsOfBlocks doc.Blocks }
                let layouts = shapeTextLayouts settings (measureWithContext ctx) (measureRangeWithContext ctx) whole shown x y height
                ctx.StoryLayouts <- Some (doc.Blocks, key, ctx.TextMeasurerVersion, layouts)
                layouts
    | None ->
    ctx.StoryTurn <- None
    match List.tryItem story.PageIndex pages with
    | None -> []
    | Some page ->
        let sectionEnds = bodyLayouts |> List.filter (fun layout -> endsSection layout.Metadata)
        let startY =
            headerFooterTargetForPage settings whole sectionEnds pages story.PageIndex story.IsHeader
            |> Option.map (fun target -> target.TargetY)
            |> Option.defaultWith (fun () -> headerFooterStartY settings whole.PageSetup page story.IsHeader)
        match ctx.StoryLayouts with
        | Some (blocks, y, version, layouts) when obj.ReferenceEquals(blocks, doc.Blocks) && y = startY && version = ctx.TextMeasurerVersion -> layouts
        | _ ->
            let layouts =
                Rbt.Word.Layout.withTablePlanScope ctx.PlanScope (fun () ->
                    storyBlockLayouts settings (measureWithContext ctx) (measureRangeWithContext ctx) whole startY doc.Blocks)
            ctx.StoryLayouts <- Some (doc.Blocks, startY, ctx.TextMeasurerVersion, layouts)
            layouts

/// A point turned about a centre by `degrees` clockwise (as the canvas turns a shape).
let turnedPoint (degrees: float) (centerX: float) (centerY: float) (x: float) (y: float) : float * float =
    let angle = degrees * System.Math.PI / 180.0
    let cos, sin = System.Math.Cos angle, System.Math.Sin angle
    let dx, dy = x - centerX, y - centerY
    centerX + dx * cos - dy * sin, centerY + dx * sin + dy * cos

/// A point of the canvas as a point of the upright layout of the story on
/// display: turned back when the story is the text of a turned text box.
/// (Call after storyLayouts, which says how the box is turned.)
let private storyPoint (ctx: RenderContext) (x: float) (y: float) : float * float =
    match ctx.StoryTurn with
    | Some (degrees, centerX, centerY) when ctx.Story.IsSome -> turnedPoint -degrees centerX centerY x y
    | _ -> x, y

/// The layouts the caret and the pointer go by: the body's, or - while a
/// header or footer is edited in place - those of its blocks on its page.
let layoutDocument (ctx: RenderContext) (doc: Document) =
    match ctx.Story with
    | Some story -> storyLayouts ctx story doc
    | None -> layoutDocumentWithPages ctx doc |> fst

/// From here on the whole document is laid out (by the next layout asked for).
let completeLayout (ctx: RenderContext) =
    // (Stale blocks above the view: the layout of the document itself, as far as it is, is taken on.)
    adoptAhead ctx
    ctx.LayoutState.BlockLimit <- System.Int32.MaxValue

/// The layout of the whole document, with its pages.
let layoutWholeDocumentWithPages (ctx: RenderContext) (doc: Document) =
    completeLayout ctx
    layoutDocumentWithPages ctx doc

/// The layout of the whole document.
let layoutWholeDocument (ctx: RenderContext) (doc: Document) =
    layoutWholeDocumentWithPages ctx doc |> fst

/// Call `callback` once the whole document is laid out (at once when it is).
let whenLayoutComplete (ctx: RenderContext) (callback: unit -> unit) =
    if layoutIsComplete ctx then callback ()
    else ctx.Progress.Waiters.Add callback

/// A document was loaded: its start is laid out first. (Only through a
/// viewport, where a part of the document is what is shown; a whole-canvas
/// host sizes its canvas from all the pages.)
let beginProgressiveLayout (ctx: RenderContext) =
    ctx.Progress.Slices <- 0
    ctx.Progress.LongestSliceMs <- 0.0
    ctx.Progress.LongestSlice <- 0
    ctx.Progress.LongestSlicePasses <- 0
    ctx.Progress.LongestSliceLayoutMs <- 0.0
    ctx.Progress.LongestSliceFrom <- 0
    ctx.Stale <- None
    ctx.Progress.Anchor <- None
    match ctx.Viewport with
    | Some _ ->
        // (A few blocks: the render raises the limit until what is in view is laid out. Forty
        // to start with were twenty to eighty pages of a document of long paragraphs, tables
        // or page breaks - laid out before anything was shown.)
        ctx.LayoutState.BlockLimit <- 6
        // (Few blocks in the first pass of the slices: what a block costs is not known yet, and
        // forty-eight of them took 140 ms with a host's shaped text.)
        ctx.Progress.Chunk <- 8
    | None -> completeLayout ctx

/// The bottom of what is laid out of the body.
let private laidOutBottom (ctx: RenderContext) (layouts: BlockLayout list) =
    match ctx.LayoutState.LastSlices with
    | Some slices when slices.SliceLayouts.Length > 0 ->
        let last = slices.SliceLayouts.[slices.SliceLayouts.Length - 1]
        last.Y + last.Height
    | _ ->
        let mutable bottom = 0.0
        for layout in layouts do
            bottom <- max bottom (layout.Y + layout.Height)
        bottom

/// Lay out as much of the document as reaches down to `bottom` (layout y).
let private coverDown (ctx: RenderContext) (doc: Document) (bottom: float) =
    let mutable more = not (stateIsWhole ctx)
    while more do
        let layouts, _ = layoutDocumentWithPages ctx doc
        if stateIsWhole ctx || laidOutBottom ctx layouts >= bottom then more <- false
        else
            // (In steps of a third: a pass over a few more blocks costs little, the blocks themselves cost their measuring.)
            let limit = ctx.LayoutState.BlockLimit
            // (Of what is laid out anew: the stale blocks above the view cost nothing.)
            let fresh = match ctx.Stale with Some stale -> max 0 (limit - stale.Keep) | None -> limit
            // (A few at first: the blocks of the first screens may be pages long each.)
            ctx.LayoutState.BlockLimit <- (if limit > 1000000 then System.Int32.MaxValue else limit + max 4 (fresh / 3))

/// What a viewport shows, and a screen below it, is laid out.
let private coverViewport (ctx: RenderContext) (doc: Document) (vp: Viewport) =
    if not (layoutIsWhole ctx) then
        let top = (vp.ScrollTop - vp.OriginY) / vp.Scale
        let height = vp.Height / vp.Scale
        // The view reaches into the stale blocks above it: the layout of the document itself
        // takes over when it has come as far as the view, else fewer blocks are kept stale.
        match ctx.Stale, ctx.LayoutState.LastSlices with
        | Some stale, Some slices when stale.Keep > 0 && stale.Keep < slices.SliceBlocks.Length && slices.SliceStarts.[stale.Keep] < slices.SliceLayouts.Length
                                       && top < slices.SliceLayouts.[slices.SliceStarts.[stale.Keep]].Y ->
            let blockAt (y: float) =
                let layouts = slices.SliceLayouts
                let mutable lo = 0
                let mutable hi = layouts.Length
                while lo < hi do
                    let mid = lo + (hi - lo) / 2
                    if layouts.[mid].Y <= y then lo <- mid + 1 else hi <- mid
                let at = max 0 (lo - 1)
                let starts = slices.SliceStarts
                let mutable first = 0
                let mutable last = slices.SliceBlocks.Length - 1
                while first < last do
                    let mid = first + (last - first + 1) / 2
                    if starts.[mid] <= at then first <- mid else last <- mid - 1
                first
            let ahead = stale.Ahead
            if ahead.LastDocument.IsSome && (ahead.LastWasWhole || ahead.LastLimit > blockAt (top + 2.0 * height) + 8) then adoptAhead ctx
            else
                captureAnchor ctx
                stale.Keep <- max 0 (blockAt (top - height))
                ctx.LayoutState.BlockLimit <- stale.Keep + 6
        | _ -> ()
        coverDown ctx doc (top + 2.0 * height)

/// The blocks the selection is in are laid out, and some after them (the
/// caret moves down by lines that must be there).
let private coverSelection (ctx: RenderContext) (doc: Document) =
    if not (layoutIsWhole ctx) then
        match doc.Selection with
        | None -> ()
        // (Everything selected and changed: the end of the selection is the end of the document,
        // which is laid out with the rest - see deferMassRelayout. A caret is laid out as ever.)
        | Some selection when ctx.Progress.MassChange && selection.Anchor <> selection.Focus -> ()
        | Some selection ->
            let indexOf (blockId: string) =
                match locateBlock doc blockId with
                | ValueSome (struct (index, _)) -> index
                | ValueNone -> -1
            let last = max (indexOf selection.Focus.BlockId) (indexOf selection.Anchor.BlockId)
            // (Near the end of what is laid out: on by a screen or so of blocks. The caret at the
            // start of a document that was just loaded asks for nothing more than is shown.)
            let limit = ctx.LayoutState.BlockLimit
            if last + 4 >= limit && limit >= 24 then
                ctx.LayoutState.BlockLimit <- last + 96
            elif last + 4 >= limit then
                ctx.LayoutState.BlockLimit <- last + 24

/// Blocks a change must touch at once for the document to be laid out again as a newly loaded one is.
let private massChangeBlocks = 400

/// Blocks there must be above the view for them to be kept as they were laid out (see StaleLayout).
let private staleKeepBlocks = 200

/// A change of most of a long document at once - everything selected and made
/// bold, a font or a size given to all of it - has every paragraph measured
/// and broken into lines again: a second and more for six hundred pages,
/// before the keystroke returns. Through a viewport such a change is laid out
/// as a newly loaded document is: from the first block that changed as far as
/// is in view (the render sees to that), and the rest in slices - the layout
/// is not complete meanwhile (`layoutIsComplete`), and what needs all of it
/// lays the rest out at once. (What is in view half way down a long document
/// far down a long document keeps the blocks above it as they were laid out
/// until the layout from the start has come as far - see StaleLayout.)
let private deferMassRelayout (ctx: RenderContext) (doc: Document) =
    if layoutIsWhole ctx && ctx.Progress.MassChange then ctx.Progress.MassChange <- false
    match ctx.Viewport, ctx.Story, ctx.Stale, ctx.LayoutState.LastDocument with
    | Some _, None, None, Some last when not ctx.Progress.LoadPending && not (obj.ReferenceEquals(last.Blocks, doc.Blocks)) ->
        // (A layout of the start of the document only: what is laid out of it can be kept stale above the view.)
        let whole = ctx.LayoutState.LastWasWhole && layoutIsWhole ctx
        // The first block that is another object, and whether the ones after it are too (an
        // edit of one paragraph leaves its neighbours the objects they were).
        let mutable before = last.Blocks
        let mutable after = doc.Blocks
        let mutable prefix = 0
        while not before.IsEmpty && not after.IsEmpty && obj.ReferenceEquals(before.Head, after.Head) do
            before <- before.Tail
            after <- after.Tail
            prefix <- prefix + 1
        let mutable probed = 0
        let mutable differing = 0
        let mutable probeBefore = before
        let mutable probeAfter = after
        while probed < 16 && not probeBefore.IsEmpty && not probeAfter.IsEmpty do
            if not (obj.ReferenceEquals(probeBefore.Head, probeAfter.Head)) then differing <- differing + 1
            probeBefore <- probeBefore.Tail
            probeAfter <- probeAfter.Tail
            probed <- probed + 1
        if probed = 16 && differing = 16 then
            let struct (_, suffix) = Rbt.Word.Layout.commonBlockEnds (List.toArray before) (List.toArray after)
            let changed = List.length after - suffix
            if changed >= massChangeBlocks then
                // The block at the top of the view and the first one below it and a screen more, as
                // they were laid out before the change.
                let blockOfLayout (slices: Rbt.Word.Layout.LayoutSlices) (at: int) =
                    let starts = slices.SliceStarts
                    let mutable first = 0
                    let mutable last = slices.SliceBlocks.Length - 1
                    while first < last do
                        let mid = first + (last - first + 1) / 2
                        if starts.[mid] <= at then first <- mid else last <- mid - 1
                    first
                let firstLayoutBelow (slices: Rbt.Word.Layout.LayoutSlices) (y: float) =
                    let layouts = slices.SliceLayouts
                    let mutable lo = 0
                    let mutable hi = layouts.Length
                    while lo < hi do
                        let mid = lo + (hi - lo) / 2
                        if layouts.[mid].Y <= y then lo <- mid + 1 else hi <- mid
                    lo
                let atTop, inView =
                    match ctx.Viewport, ctx.LayoutState.LastSlices with
                    | Some vp, Some slices when slices.SliceLayouts.Length > 0 ->
                        let top = (vp.ScrollTop - vp.OriginY) / vp.Scale
                        let bottom = top + 2.0 * vp.Height / vp.Scale
                        blockOfLayout slices (max 0 (firstLayoutBelow slices top - 1)), blockOfLayout slices (firstLayoutBelow slices bottom) + 8
                    | _ -> 0, 0
                if atTop - prefix >= staleKeepBlocks then
                    // Far down the document: the blocks above the view stay as they were laid out, the
                    // document is laid out from the block at the top of the view, and from its start
                    // in slices (see StaleLayout).
                    captureAnchor ctx
                    let ahead = Rbt.Word.Layout.createIncrementalLayoutState ()
                    ahead.BlockLimit <- 6
                    ctx.Stale <- Some { StaleBlocks = last.Blocks; Keep = atTop; Target = doc.Blocks; HybridFor = []; HybridKeep = -1; Hybrid = []; Ahead = ahead }
                    ctx.LayoutState.BlockLimit <- max (atTop + 6) inView
                    ctx.Progress.Chunk <- 8
                    ctx.Progress.MassChange <- true
                elif whole then
                    // As far as was in view before the change, in one pass (a pass has a cost of its own
                    // that grows with what is laid out: raising the limit step by step to a view half way
                    // down six hundred pages took longer than laying all of them out at once). The render
                    // takes it further when the text got shorter.
                    ctx.LayoutState.BlockLimit <- max (prefix + 6) inView
                    ctx.Progress.Chunk <- 8
                    ctx.Progress.MassChange <- true
    | _ -> ()

/// One slice of the layout of the rest of the document: blocks are added for
/// about `budget` milliseconds. True when the whole document is laid out.
let advanceLayout (ctx: RenderContext) (doc: Document) (budget: float) : bool =
    let progress = ctx.Progress
    let started = nowMs ()
    // (A pass aims at a third of the slice: the blocks differ in what they cost - a page of a
    // table against a heading - and a pass that turns out three times as long as the last ones
    // said still ends near the slice's time.)
    let target = max 3.0 (min budget 30.0 / 3.0)
    let mutable more = not (layoutIsWhole ctx)
    while more do
        let passStart = nowMs ()
        // (Stale blocks above the view: the layout that is advanced is that of the document itself,
        // of the blocks the document has now.)
        if ctx.Stale.IsSome && ctx.Story.IsSome then adoptAhead ctx
        match ctx.Stale with
        | Some stale -> hybridBlocks stale doc.Blocks |> ignore
        | None -> ()
        let state = match ctx.Stale with Some stale -> stale.Ahead | None -> ctx.LayoutState
        let limit = state.BlockLimit
        // (Never a sliver of what is laid out already: a pass has a cost of its own that grows with the document.)
        let chunk = max progress.Chunk (limit / 512)
        state.BlockLimit <- limit + chunk
        match ctx.Stale with
        | Some stale ->
            Rbt.Word.Layout.withTablePlanScope ctx.PlanScope (fun () ->
                Rbt.Word.Layout.layoutDocumentPagesIncremental stale.Ahead (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) doc) |> ignore
            // It has come as far as the layout on display: that is the layout from here on.
            if Rbt.Word.Layout.incrementalLayoutIsWhole stale.Ahead
               || (ctx.LayoutState.BlockLimit < System.Int32.MaxValue && stale.Ahead.BlockLimit >= ctx.LayoutState.BlockLimit) then
                adoptAhead ctx
                layoutDocumentWithPages ctx doc |> ignore
        | None -> layoutDocumentWithPages ctx doc |> ignore
        progress.SlicePasses <- progress.SlicePasses + 1
        let took = max 0.25 (nowMs () - passStart)
        progress.Chunk <- max 8 (min (chunk * 2) (int (float chunk * target / took)))
        if layoutIsWhole ctx || nowMs () - started + took * 0.6 > budget then more <- false
    layoutIsComplete ctx

[<Emit("(typeof document !== 'undefined' && document.hidden === true)")>]
let private pageIsHidden () : bool = jsNative

// The next slice runs as a task of its own right after the current one:
// input events and frames the browser has waiting come first, and no time is
// spent waiting for an idle period (which comes once per frame at best - a
// long document would take a second longer to lay out). In two hops: a timer
// that came due during a slice is made ready when the slice's task ends,
// behind a next slice that was posted from inside it - the page's timers
// then saw two slices as one. The next slice is posted from a task of nothing.
[<Emit("""(function (run) {
    var hop = function (next) {
        if (typeof scheduler !== 'undefined' && scheduler && typeof scheduler.postTask === 'function') { scheduler.postTask(next, { priority: 'user-visible' }); return; }
        if (typeof MessageChannel === 'function') { var channel = new MessageChannel(); channel.port1.onmessage = function () { channel.port1.close(); next(); }; channel.port2.postMessage(0); return; }
        setTimeout(next, 0);
    };
    hop(function () { hop(run); });
})($0)""")>]
let private afterThisTask (_callback: unit -> unit) : unit = jsNative

/// Lay out the rest of the document in slices (tasks of about 24 ms that let
/// input and frames in between), until it is whole.
/// `currentDocument` gives the document on display (it changes while the
/// user types), `afterSlice complete tell` is called after every slice: the
/// host repaints when it has to and, when `tell` is set (not with every
/// slice), tells its listeners that there are more pages.
let rec scheduleLayout (ctx: RenderContext) (currentDocument: unit -> Document) (afterSlice: bool -> bool -> unit) : unit =
    let progress = ctx.Progress
    // (All of what is there laid out while more is being read: the slices start again when it is.)
    if not progress.Scheduled && not (layoutIsWhole ctx) then
        progress.Scheduled <- true
        let epoch = progress.Epoch
        afterThisTask (fun () ->
            progress.Scheduled <- false
            if epoch = progress.Epoch then
                if not (layoutIsWhole ctx) then
                    // A page that is not shown gets its tasks rarely: longer slices then.
                    let budget = if pageIsHidden () then 250.0 else 24.0
                    let sliceStart = nowMs ()
                    let from = (match ctx.Stale with Some stale -> stale.Ahead.BlockLimit | None -> ctx.LayoutState.BlockLimit)
                    progress.SlicePasses <- 0
                    let complete = advanceLayout ctx (currentDocument ()) budget
                    let now = nowMs ()
                    let tell = complete || progress.Adopted || now - progress.Notified > 150.0
                    if tell then progress.Notified <- now
                    afterSlice complete tell
                    let took = nowMs () - sliceStart
                    progress.Slices <- progress.Slices + 1
                    if took > progress.LongestSliceMs then
                        progress.LongestSliceMs <- took
                        progress.LongestSlice <- progress.Slices
                        progress.LongestSlicePasses <- progress.SlicePasses
                        progress.LongestSliceLayoutMs <- now - sliceStart
                        progress.LongestSliceFrom <- from
                scheduleLayout ctx currentDocument afterSlice)

/// Call off the slices that are waiting (the context is not used any more).
let cancelScheduledLayout (ctx: RenderContext) =
    ctx.Progress.Epoch <- ctx.Progress.Epoch + 1
    ctx.Progress.Scheduled <- false
    ctx.Progress.Waiters.Clear()

/// The layouts the state has of this very document, at the limit it is at.
let cachedLayouts (ctx: RenderContext) (doc: Document) =
    match ctx.Story with
    | Some story -> Some (storyLayouts ctx story doc)
    | None ->
    match ctx.LayoutState.LastDocument, ctx.LayoutState.LastLayouts with
    | Some cachedDoc, Some layouts when obj.ReferenceEquals(cachedDoc, doc) && Rbt.Word.Layout.incrementalLayoutIsCurrent ctx.LayoutState -> Some layouts
    | _ -> None

let clearLayoutCache (ctx: RenderContext) =
    clearMeasuredWidths ()
    ctx.Stale <- None
    ctx.Progress.Anchor <- None
    ctx.PlanScope <- obj ()
    Rbt.Word.Layout.clearIncrementalLayoutState ctx.LayoutState
    clearPaintCache ctx.PaintCache
    ctx.LastPaint <- None
    ctx.PageFrame <- None

/// The first `Keep` blocks of a layout with stale blocks above the view (see StaleLayout).
let private staleHead (stale: StaleLayout) : Block list =
    (if stale.Hybrid.IsEmpty then stale.StaleBlocks else stale.Hybrid) |> List.truncate stale.Keep

/// What was measured is measured otherwise from now on (other measurers,
/// another face) while the blocks above the view are shown as they were laid
/// out before a change (see StaleLayout): those blocks stay as they are - the
/// document's height and its pages above the view with them, and the view on
/// its text - and everything else `uses` names is laid out again: what is in
/// view at once, the document itself from its start in slices, as before.
/// (Taking the layout of the document itself on here, as far as it had come -
/// three pages of six hundred - made the document that short and threw the
/// view to its top.) Returns how many paragraphs and blocks are laid out
/// again, or None when no blocks are stale.
let private invalidateBelowStale (ctx: RenderContext) (uses: (TextBuffer.TextStyle -> bool) option) : int option =
    match ctx.Stale with
    | Some stale when ctx.Story.IsNone && stale.Keep > 0 ->
        let count =
            match uses with
            | Some uses ->
                let shown = Rbt.Word.Layout.invalidateFacesInIncrementalLayoutBelow ctx.LayoutState uses (staleHead stale)
                shown + Rbt.Word.Layout.invalidateFacesInIncrementalLayout stale.Ahead uses
            | None ->
                Rbt.Word.Layout.invalidateFacesInIncrementalLayoutBelow ctx.LayoutState (fun _ -> true) (staleHead stale) |> ignore
                // (The layout of the document itself: from its start again.)
                Rbt.Word.Layout.clearIncrementalLayoutState stale.Ahead
                stale.Ahead.BlockLimit <- 6
                ctx.Progress.Chunk <- 8
                1
        Some count
    | _ -> None

/// Install the text path painter of the context's own canvas and the factory
/// that makes one for any other canvas (both None: text is painted with
/// canvas text). What was painted with the previous painter is stale from
/// here on; nothing is invalidated at any other time.
let setTextPathPainter (ctx: RenderContext)
                       (painter: (string -> TextPaintStyle -> float -> float -> float -> bool) option)
                       (factory: System.Func<CanvasRenderingContext2D, System.Func<string, TextPaintStyle, float, float, float, bool>> option) =
    ctx.TextPathPainter <- painter
    ctx.TextPathPainterFactory <- factory
    ctx.TextPathPainterVersion <- ctx.TextPathPainterVersion + 1
    ctx.PageFrame <- None

/// As `setTextPathPainter`, from a factory: the painter of the context's own
/// canvas is the one the factory makes for it.
let setTextPathPainterFactory (ctx: RenderContext) (factory: System.Func<CanvasRenderingContext2D, System.Func<string, TextPaintStyle, float, float, float, bool>>) =
    if isNull (box factory) then setTextPathPainter ctx None None
    else
        let paint = factory.Invoke ctx.Ctx
        let painter =
            if isNull (box paint) then None
            else Some (fun (text: string) (style: TextPaintStyle) (x: float) (baseline: float) (width: float) -> paint.Invoke(text, style, x, baseline, width))
        setTextPathPainter ctx painter (Some factory)

/// Install the text measurers (None: measure with canvas text). Lines and
/// pages measured with the previous ones are laid out again.
let setTextMeasurers (ctx: RenderContext) (measureText: TextMeasurer option) (measureRange: TextRangeMeasurer option) =
    ctx.TextMeasurer <- measureText
    ctx.TextRangeMeasurer <- measureRange
    ctx.TextMeasurerVersion <- ctx.TextMeasurerVersion + 1
    match invalidateBelowStale ctx None with
    | Some _ ->
        clearMeasuredWidths ()
        ctx.PlanScope <- obj ()
        clearPaintCache ctx.PaintCache
        ctx.LastPaint <- None
        ctx.PageFrame <- None
    | None -> clearLayoutCache ctx

/// As `setTextMeasurers`, for measurers that differ from the ones installed
/// only in how they measure text that names one of `families` (the host got
/// faces of those families and set such text in a fallback until now): the
/// paragraphs set in them are laid out again, as paragraphs an edit changed
/// are; the lines of all the others stand. Returns how many paragraphs and
/// blocks were set in the families (0: the layout stands as it is).
let setTextMeasurersForFamilies (ctx: RenderContext) (measureText: TextMeasurer option) (measureRange: TextRangeMeasurer option) (families: string seq) : int =
    ctx.TextMeasurer <- measureText
    ctx.TextRangeMeasurer <- measureRange
    ctx.TextMeasurerVersion <- ctx.TextMeasurerVersion + 1
    let normalize (name: string) = name.Trim().ToLowerInvariant()
    let names = System.Collections.Generic.HashSet<string>()
    for family in families do
        if not (isNull family) then names.Add (normalize family) |> ignore
    let uses (style: TextBuffer.TextStyle) =
        names.Contains (normalize (style.FontFamily |> Option.filter (fun name -> name.Trim() <> "") |> Option.defaultValue ctx.FontFamily))
    let count =
        if names.Count = 0 then 0
        else
            // (Blocks above the view that are shown as they were laid out before a change stay so.)
            match invalidateBelowStale ctx (Some uses) with
            | Some count -> count
            | None ->
                adoptAhead ctx
                Rbt.Word.Layout.invalidateFacesInIncrementalLayout ctx.LayoutState uses
    if count > 0 then
        // (The tables that are laid out again are planned again; what is painted is collected again.)
        clearMeasuredWidths ()
        ctx.PlanScope <- obj ()
        clearPaintCache ctx.PaintCache
        ctx.LastPaint <- None
        ctx.PageFrame <- None
    count

let layoutLookupFor (ctx: RenderContext) (doc: Document) =
    // (Asked for by commands that may name any block: the whole document.)
    completeLayout ctx
    let layouts =
        cachedLayouts ctx doc
        |> Option.defaultWith (fun () -> layoutDocument ctx doc)
    fun blockId ->
        layouts |> List.tryFind (fun layout -> layout.BlockId = blockId)

let private blockHasComments (block: Block) =
    let inParagraph (para: ParagraphBlock) =
        let mutable found = false
        for run in para.Buffer.Runs do
            match run with
            | Text text when text.Style.CommentId.IsSome -> found <- true
            | _ -> ()
        found
    match block with
    | Paragraph para
    | Heading (_, para) -> inParagraph para
    | Table table ->
        let mutable found = false
        for row in table.Rows do
            for cell in row.Cells do
                for para in cell.Content do
                    if not found && inParagraph para then found <- true
        found
    | TextBox textBox -> textBox.Content |> List.exists inParagraph

/// The indexes of the blocks that hold comments. Which blocks do is kept
/// from frame to frame (a block that is the same object has not changed), so
/// an edit looks at the runs of the blocks it changed.
let private commentBlockIndexes (ctx: RenderContext) (blocks: Block array) =
    let struct (prefix, suffix) = Rbt.Word.Layout.commonBlockEnds ctx.CommentBlocks blocks
    let shift = ctx.CommentBlocks.Length - blocks.Length
    let flags = Array.zeroCreate<bool> blocks.Length
    let indexes = ResizeArray<int>()
    for index in 0 .. blocks.Length - 1 do
        let flag =
            if index < prefix then ctx.CommentFlags.[index]
            elif index >= blocks.Length - suffix then ctx.CommentFlags.[index + shift]
            else blockHasComments blocks.[index]
        flags.[index] <- flag
        if flag then indexes.Add index
    ctx.CommentBlocks <- blocks
    ctx.CommentFlags <- flags
    indexes

let private commentAnchorsForLayouts (ctx: RenderContext) (doc: Document) (layouts: BlockLayout list) =
    // The layouts block by block, when they are the layout state's own.
    let slices =
        match ctx.LayoutState.LastSlices, ctx.LayoutState.LastLayouts with
        | Some slices, Some known when obj.ReferenceEquals(known, layouts) -> Some slices
        | _ -> None
    let blocks =
        match slices with
        | Some slices -> slices.SliceBlocks
        | None -> List.toArray doc.Blocks
    let commented = commentBlockIndexes ctx blocks
    if commented.Count = 0 then [] else
    let settings = layoutSettings ctx doc
    // (The last layout of a paragraph that has several.)
    let layoutById = System.Collections.Generic.Dictionary<string, BlockLayout>()
    match slices with
    | Some slices ->
        for index in commented do
            for at in slices.SliceStarts.[index] .. slices.SliceStarts.[index + 1] - 1 do
                layoutById.[slices.SliceLayouts.[at].BlockId] <- slices.SliceLayouts.[at]
    | None ->
        for layout in layouts do
            layoutById.[layout.BlockId] <- layout
    let tryLayout (blockId: string) =
        match layoutById.TryGetValue blockId with
        | true, layout -> Some layout
        | false, _ -> None
    // (Blocks without comments have no part in the list.)
    let commentedDocument = { doc with Blocks = [ for index in commented -> blocks.[index] ] }
    let calloutX = max (settings.Width - settings.MarginRight + 12.0) (settings.PaddingLeft + 120.0)

    let isSelected blockId startOffset endOffset =
        match doc.Selection with
        | Some selection when selection.Anchor.BlockId = blockId && selection.Focus.BlockId = blockId ->
            let selectionStart = min selection.Anchor.Offset selection.Focus.Offset
            let selectionEnd = max selection.Anchor.Offset selection.Focus.Offset
            selectionStart = startOffset && selectionEnd = endOffset
        | _ -> false

    let anchors =
        commentInfos commentedDocument
        |> List.choose (fun info ->
            match tryLayout info.BlockId with
            | None -> None
            | Some layout ->
                let line = lineForOffset layout info.StartOffset
                let lineEnd = min info.EndOffset line.EndOffset
                if lineEnd <= info.StartOffset then None
                else
                    let lineAlignmentOffset = alignmentOffset layout line
                    let startX = xForOffsetWithRange (measureRangeWithContext ctx) line info.StartOffset
                    let endX = xForOffsetWithRange (measureRangeWithContext ctx) line lineEnd
                    Some {
                        Id = info.Id
                        Text = info.Text
                        Author = info.Author
                        Date = info.Date
                        Excerpt = info.Excerpt
                        BlockId = info.BlockId
                        StartOffset = info.StartOffset
                        EndOffset = info.EndOffset
                        X = settings.PaddingLeft + startX + lineAlignmentOffset
                        Y = line.Y
                        Width = max 1.0 (endX - startX)
                        Height = line.Height
                        CalloutX = calloutX
                        CalloutY = line.Y
                        IsSelected = isSelected info.BlockId info.StartOffset info.EndOffset
                    })
        |> List.distinctBy (fun anchor -> anchor.Id, anchor.Text)
        |> List.sortBy (fun anchor -> anchor.Y)

    let minGap = 8.0
    let defaultCalloutHeight = 58.0
    (([], -100000.0), anchors)
    ||> List.fold (fun (acc, previousBottom) anchor ->
        let calloutY = max anchor.Y (previousBottom + minGap)
        ({ anchor with CalloutY = calloutY } :: acc, calloutY + defaultCalloutHeight))
    |> fst
    |> List.rev

do notePlace.Value <- fun ctx whole layouts (kind, id) ->
    let settings = layoutSettings ctx whole
    if kind = "comment" then
        // A comment: in a balloon below the line its anchor starts on.
        commentAnchorsForLayouts ctx whole layouts
        |> List.tryFind (fun anchor -> anchor.Id = id)
        |> Option.map (fun anchor ->
            let contentWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
            settings.PaddingLeft + 24.0, anchor.Y + anchor.Height + 10.0, min 360.0 (max 120.0 (contentWidth - 48.0)))
    else
        // A footnote or an endnote: where the note area shows it, after its mark.
        layouts
        |> List.tryFind (fun layout -> noteOfLayoutId layout.BlockId = Some (kind, id))
        |> Option.bind (fun layout ->
            match layout.Lines with
            | [] -> None
            | first :: _ ->
                let markX = xForOffsetWithRange (measureRangeWithContext ctx) first (noteMarkText kind id).Length
                Some (settings.PaddingLeft + markX + alignmentOffset layout first, layout.Y, max 40.0 (layout.AvailableWidth - markX)))

/// The parts of a note's layouts page by page: the top and bottom of each.
let private noteFrames (settings: LayoutSettings) (layouts: BlockLayout list) : (float * float) list =
    let stride = settings.PageHeight + pageGap
    let parts = ResizeArray<int * float * float>()
    for layout in layouts do
        for line in layout.Lines do
            let page = int (floor (line.Y / stride))
            match parts |> Seq.tryFindIndex (fun (known, _, _) -> known = page) with
            | Some index ->
                let _, top, bottom = parts.[index]
                parts.[index] <- (page, min top line.Y, max bottom (line.Y + line.Height))
            | None -> parts.Add (page, line.Y, line.Y + line.Height)
        if layout.Lines.IsEmpty then
            let page = int (floor (layout.Y / stride))
            if not (parts |> Seq.exists (fun (known, _, _) -> known = page)) then parts.Add (page, layout.Y, layout.Y + layout.Height)
    parts |> Seq.map (fun (_, top, bottom) -> top, bottom) |> List.ofSeq

/// What is painted over the page for a note or comment that is edited in place
/// (see StoryView.Note): what the page shows of it covered, its paragraphs as
/// they are laid out to be edited, and a frame around them (a comment: a balloon).
let private noteStoryCommands (ctx: RenderContext) (whole: Document) (layouts: BlockLayout list) (kind: string) (id: string) : PaintCommand list =
    if kind <> "comment" then
        // A footnote or an endnote is edited where the note area shows it, as it is laid out
        // there: a frame around it (each page's part of it) says that it is the note that is edited.
        let settings = layoutSettings ctx whole
        let frames = noteFrames settings layouts
        frames
        |> List.collect (fun (top, bottom) ->
            let x0, y0, x1, y1 = settings.PaddingLeft - 4.0, top - 2.5, settings.Width - settings.MarginRight + 4.0, bottom + 2.5
            [ StrokeLine ("#5b9bd5", 1.0, 2, 0.0, x0, y0, x1, y0)
              StrokeLine ("#5b9bd5", 1.0, 2, 0.0, x1, y0, x1, y1)
              StrokeLine ("#5b9bd5", 1.0, 2, 0.0, x1, y1, x0, y1)
              StrokeLine ("#5b9bd5", 1.0, 2, 0.0, x0, y1, x0, y0) ])
    else
    match ctx.StoryNotePlace with
    | None -> []
    | Some (left, top, width) ->
        let settings = layoutSettings ctx whole
        let bottom =
            match layouts with
            | [] -> top + ctx.LineHeight
            | _ -> layouts |> List.map (fun layout -> layout.Y + layout.Height) |> List.max
        let text = layouts |> List.collect (paragraphCommands settings ctx.FontFamily (measureWithContext ctx) (measureRangeWithContext ctx))
        if kind = "comment" then
            [ FillRect ("#fffdf3", left - 10.0, top - 6.0, width + 20.0, bottom - top + 12.0)
              StrokeRect ("#C79B00", 1.0, left - 9.5, top - 5.5, width + 19.0, bottom - top + 11.0) ]
            @ text
        else []

/// Whether a point (layout coordinates) is on the note or comment that is
/// edited in place (true while another story, or none, is edited).
let onNoteStory (ctx: RenderContext) (doc: Document) (x: float) (y: float) : bool =
    match ctx.Story with
    | Some story when (match story.Note with Some (kind, _) -> kind <> "comment" | None -> false) ->
        // A footnote or an endnote: on the lines of the note where the note area shows them.
        let whole = story.Whole doc
        let settings = layoutSettings ctx whole
        noteFrames settings (storyLayouts ctx story doc)
        |> List.exists (fun (top, bottom) -> x >= settings.PaddingLeft - 12.0 && x <= settings.Width - settings.MarginRight + 12.0 && y >= top - 6.0 && y <= bottom + 6.0)
    | Some story when story.Note.IsSome ->
        match ctx.StoryNotePlace with
        | None -> false
        | Some (left, top, width) ->
            let layouts = storyLayouts ctx story doc
            let bottom =
                match layouts with
                | [] -> top + ctx.LineHeight
                | _ -> layouts |> List.map (fun layout -> layout.Y + layout.Height) |> List.max
            x >= left - 12.0 && x <= left + width + 12.0 && y >= top - 8.0 && y <= bottom + 8.0
    | _ -> true

/// The footnote or endnote (type, id) the page shows at a point (layout coordinates), if any.
let noteAt (ctx: RenderContext) (doc: Document) (x: float) (y: float) : (string * string) option =
    let whole = bodyDocument ctx doc
    let settings = layoutSettings ctx whole
    layoutDocumentWithPages ctx whole |> fst
    |> List.tryPick (fun layout ->
        match noteOfLayoutId layout.BlockId with
        | Some note when y >= layout.Y && y <= layout.Y + layout.Height
                         && x >= settings.PaddingLeft - 4.0 && x <= settings.PaddingLeft + layout.AvailableWidth + 4.0 -> Some note
        | _ -> None)

let commentAnchors (ctx: RenderContext) (doc: Document) =
    let doc = bodyDocument ctx doc
    commentAnchorsForLayouts ctx doc (layoutWholeDocument ctx doc)

let private layoutBottom (settings: LayoutSettings) (layouts: BlockLayout list) =
    let mutable bottom = settings.PaddingTop
    for layout in layouts do
        bottom <- max bottom (layout.Y + layout.Height)
    bottom

/// Shared layout already places footnotes and endnotes. Cover their actual bounds;
/// estimating them again would append phantom pages and inflate NUMPAGES.
let private pageFramesOfLayouts (settings: LayoutSettings) (layouts: BlockLayout list) (pages: PageInfo list) =
    pageInfosCoveringBottom (layoutBottom settings layouts) pages

/// `pageFramesOfLayouts` for the layouts and pages the layout state made
/// last: kept for that pass (a host asks for the document's size with every change).
let private pageFramesForLayouts (ctx: RenderContext) (settings: LayoutSettings) (doc: Document) (layouts: BlockLayout list) (pages: PageInfo list) =
    let current =
        match ctx.LayoutState.LastLayouts with
        | Some known -> obj.ReferenceEquals(known, layouts)
        | None -> false
    match ctx.KnownPageFrames with
    | Some (serial, frames) when current && serial = ctx.LayoutState.LayoutSerial -> frames
    | _ ->
        let frames = pageFramesOfLayouts settings layouts pages
        if current then ctx.KnownPageFrames <- Some (ctx.LayoutState.LayoutSerial, frames)
        frames

do framePages.Value <- fun ctx doc layouts pages -> pageFramesForLayouts ctx (layoutSettings ctx doc) doc layouts pages

let private renderPages (ctx: RenderContext) (pages: PageInfo list) =
    ctx.Ctx.fillStyle <- U3.Case1 "#eef0f3"
    ctx.Ctx.fillRect(0.0, 0.0, ctx.Canvas.width, ctx.Canvas.height)

    for page in pages do
        ctx.Ctx.fillStyle <- U3.Case1 "#ffffff"
        ctx.Ctx.fillRect(0.0, page.Top, page.Width, page.Height)
        ctx.Ctx.strokeStyle <- U3.Case1 "#d0d5dd"
        ctx.Ctx.lineWidth <- 1.0
        ctx.Ctx.strokeRect(0.5, page.Top + 0.5, page.Width - 1.0, page.Height - 1.0)

let private fontString (style: TextPaintStyle) defaultFontFamily =
    let fontStyle = if style.Bold && style.Italic then "bold italic"
                    elif style.Bold then "bold"
                    elif style.Italic then "italic"
                    else ""
    let fontFamily = style.FontFamily |> Option.defaultValue defaultFontFamily
    sprintf "%s %fpx %s" fontStyle style.FontSize (cssFontFamily fontFamily)

let private imageElement (ctx: RenderContext) (src: string) =
    match ctx.ImageCache with
    | Some cache ->
        let cachedImg = mapGet cache src
        if not (isNull cachedImg) then
            cachedImg :?> Browser.Types.HTMLImageElement
        else
            let newImg = Browser.Dom.document.createElement("img") :?> Browser.Types.HTMLImageElement
            newImg.src <- src
            newImg
    | None ->
        let newImg = Browser.Dom.document.createElement("img") :?> Browser.Types.HTMLImageElement
        newImg.src <- src
        newImg

[<Emit("$0.drawImage($1, $2, $3, $4, $5, $6, $7, $8, $9)")>]
let private drawImageCropped (_ctx: CanvasRenderingContext2D) (_img: HTMLImageElement) (_sx: float) (_sy: float) (_sw: float) (_sh: float) (_dx: float) (_dy: float) (_dw: float) (_dh: float) : unit = jsNative

let private cropValue value =
    match value with
    | Some text ->
        match System.Double.TryParse(text: string) with
        | true, parsed -> max 0.0 (min 100000.0 parsed)
        | false, _ -> 0.0
    | None -> 0.0

let private drawImageWithCrop (ctx: RenderContext) (src: string) (x: float) (y: float) (width: float) (height: float) cropLeft cropTop cropRight cropBottom =
    let img = imageElement ctx src
    let sourceWidth = float img.naturalWidth
    let sourceHeight = float img.naturalHeight
    let left = cropValue cropLeft
    let top = cropValue cropTop
    let right = cropValue cropRight
    let bottom = cropValue cropBottom
    if sourceWidth <= 0.0 || sourceHeight <= 0.0 || left + right >= 100000.0 || top + bottom >= 100000.0 then
        ctx.Ctx.drawImage(U3.Case1 img, x, y, width, height)
    else
        let sx = sourceWidth * left / 100000.0
        let sy = sourceHeight * top / 100000.0
        let sw = sourceWidth * (100000.0 - left - right) / 100000.0
        let sh = sourceHeight * (100000.0 - top - bottom) / 100000.0
        if left = 0.0 && top = 0.0 && right = 0.0 && bottom = 0.0 then
            ctx.Ctx.drawImage(U3.Case1 img, x, y, width, height)
        else
            drawImageCropped ctx.Ctx img sx sy sw sh x y width height

let private paintCharacterSpacedText (ctx: RenderContext) (text: string) characterSpacing =
    if abs characterSpacing < 0.001 || text.Length <= 1 then
        ctx.Ctx.fillText(text, 0.0, 0.0)
    else
        let mutable cursor = 0.0
        for i in 0 .. text.Length - 1 do
            let ch = text.Substring(i, 1)
            ctx.Ctx.fillText(ch, cursor, 0.0)
            cursor <- cursor + ctx.Ctx.measureText(ch).width
            if i < text.Length - 1 || System.Char.IsWhiteSpace text.[i] then
                cursor <- cursor + characterSpacing

let private fillTextFitted (ctx: RenderContext) (style: TextPaintStyle) text x baseline targetWidth =
    if targetWidth <= 0.0 || System.String.IsNullOrEmpty text then
        ctx.Ctx.save()
        ctx.Ctx.translate(x, baseline)
        paintCharacterSpacedText ctx text style.CharacterSpacing
        ctx.Ctx.restore()
    else
        let actualWidth = ctx.Ctx.measureText(text).width + style.CharacterSpacing * TextBuffer.characterSpacingSlots text
        if actualWidth <= 0.0 || abs (actualWidth - targetWidth) < 0.01 then
            ctx.Ctx.save()
            ctx.Ctx.translate(x, baseline)
            paintCharacterSpacedText ctx text style.CharacterSpacing
            ctx.Ctx.restore()
        else
            ctx.Ctx.save()
            ctx.Ctx.translate(x, baseline)
            ctx.Ctx.scale(targetWidth / actualWidth, 1.0)
            paintCharacterSpacedText ctx text style.CharacterSpacing
            ctx.Ctx.restore()

let private tracePath (ctx: RenderContext) (segs: PathSeg list) =
    ctx.Ctx.beginPath()
    segs
    |> List.iter (function
        | PMove(x, y) -> ctx.Ctx.moveTo(x, y)
        | PLine(x, y) -> ctx.Ctx.lineTo(x, y)
        | PQuad(cx, cy, x, y) -> ctx.Ctx.quadraticCurveTo(cx, cy, x, y)
        | PCubic(c1x, c1y, c2x, c2y, x, y) -> ctx.Ctx.bezierCurveTo(c1x, c1y, c2x, c2y, x, y)
        | PClose -> ctx.Ctx.closePath())

let private shapeFillStyle (ctx: RenderContext) (fill: ShapeFill) (bx, by, bw, bh) =
    match fill with
    | FillNone -> None
    | FillSolid color -> Some (U3.Case1 color)
    | FillGradient g ->
        let grad =
            if g.Radial then
                let cx, cy = bx + bw / 2.0, by + bh / 2.0
                ctx.Ctx.createRadialGradient(cx, cy, 0.0, cx, cy, (max bw bh) / 2.0)
            else
                let rad = g.Angle * System.Math.PI / 180.0
                let cx, cy = bx + bw / 2.0, by + bh / 2.0
                let dx, dy = cos rad * bw / 2.0, sin rad * bh / 2.0
                ctx.Ctx.createLinearGradient(cx - dx, cy - dy, cx + dx, cy + dy)
        g.Stops |> List.iter (fun s -> grad.addColorStop(max 0.0 (min 1.0 s.Offset), s.Color))
        Some (U3.Case2 grad)

let rec private renderPaintCommand (ctx: RenderContext) = function
    | DrawText (text, style, x, baseline, width) ->
        match ctx.TextPathPainter with
        | Some paint when paint text style x baseline width -> ()
        | _ ->
            ctx.Ctx.font <- fontString style ctx.FontFamily
            ctx.Ctx.fillStyle <- U3.Case1 style.Color
            fillTextFitted ctx style text x baseline width
    | DrawUnderline (color, lineStyle, x, y, width) ->
        ctx.Ctx.strokeStyle <- U3.Case1 color
        let ls = lineStyle.ToLowerInvariant()
        let heavy = ls.Contains("heavy") || ls = "thick"
        ctx.Ctx.lineWidth <- if heavy then 2.0 else 1.0
        let flat dy =
            ctx.Ctx.beginPath()
            ctx.Ctx.moveTo(x, y + dy)
            ctx.Ctx.lineTo(x + width, y + dy)
            ctx.Ctx.stroke()
        let dashed pattern =
            setLineDash ctx.Ctx pattern
            flat 0.0
            setLineDash ctx.Ctx [||]
        let wavy dy =
            // triangular (zigzag) wave — only needs lineTo, looks wavy at text scale
            ctx.Ctx.beginPath()
            ctx.Ctx.moveTo(x, y + dy)
            let mutable cx = x
            let mutable up = true
            while cx < x + width do
                let nx = min (x + width) (cx + 2.0)
                ctx.Ctx.lineTo(nx, y + dy + (if up then -1.3 else 1.3))
                cx <- nx
                up <- not up
            ctx.Ctx.stroke()
        if ls.Contains("wav") then
            wavy 0.0
            if ls.Contains("double") then wavy 2.6
        elif ls.Contains("dotdotdash") || ls.Contains("dashdotdot") then dashed [|4.0;2.0;1.0;2.0;1.0;2.0|]
        elif ls.Contains("dotdash") || ls.Contains("dashdot") then dashed [|4.0;2.0;1.0;2.0|]
        elif ls.Contains("dash") then dashed [|4.0;2.0|]
        elif ls.Contains("dot") then dashed [|1.0;2.0|]
        elif ls = "double" then
            flat 0.0
            flat 2.0
        else
            flat 0.0
    | DrawImage (src, x, y, width, height, cropLeft, cropTop, cropRight, cropBottom) ->
        drawImageWithCrop ctx src x y width height cropLeft cropTop cropRight cropBottom
    | FillRect (color, x, y, width, height) ->
        ctx.Ctx.fillStyle <- U3.Case1 color
        ctx.Ctx.fillRect(x, y, width, height)
    | StrokeRect (color, lineWidth, x, y, width, height) ->
        ctx.Ctx.strokeStyle <- U3.Case1 color
        ctx.Ctx.lineWidth <- lineWidth
        setLineDash ctx.Ctx [||]
        ctx.Ctx.strokeRect(x, y, width, height)
    | StrokeLine (color, lineWidth, dashStyle, dashOffset, x1, y1, x2, y2) ->
        let dash =
            match dashStyle with
            | 1 -> [| lineWidth; lineWidth |]
            | 2 -> [| lineWidth * 3.0; lineWidth * 1.5 |]
            | 3 -> [| lineWidth * 3.0; lineWidth * 1.5; lineWidth; lineWidth * 1.5 |]
            | _ -> [||]
        ctx.Ctx.strokeStyle <- U3.Case1 color
        ctx.Ctx.lineWidth <- lineWidth
        ctx.Ctx.lineDashOffset <- dashOffset
        setLineDash ctx.Ctx dash
        ctx.Ctx.beginPath()
        ctx.Ctx.moveTo(x1, y1)
        ctx.Ctx.lineTo(x2, y2)
        ctx.Ctx.stroke()
        setLineDash ctx.Ctx [||]
    | DrawShape (fill, stroke, strokeWidth, bbox, segs) ->
        tracePath ctx segs
        match shapeFillStyle ctx fill bbox with
        | Some style -> ctx.Ctx.fillStyle <- style; ctx.Ctx.fill()
        | None -> ()
        match stroke with
        | Some color ->
            ctx.Ctx.strokeStyle <- U3.Case1 color
            ctx.Ctx.lineWidth <- strokeWidth
            setLineDash ctx.Ctx [||]
            ctx.Ctx.stroke()
        | None -> ()
    | DrawGroup (clip, rotation, centerX, centerY, opacity, _, commands) ->
        ctx.Ctx.save()
        if opacity < 0.999 then ctx.Ctx.globalAlpha <- ctx.Ctx.globalAlpha * (max 0.0 opacity)
        if abs rotation >= 0.01 then
            ctx.Ctx.translate(centerX, centerY)
            ctx.Ctx.rotate(rotation * System.Math.PI / 180.0)
            ctx.Ctx.translate(-centerX, -centerY)
        match clip with
        | Some (x, y, width, height) ->
            ctx.Ctx.beginPath()
            ctx.Ctx.rect(x, y, width, height)
            ctx.Ctx.clip()
        | None -> ()
        for command in commands do
            renderPaintCommand ctx command
        ctx.Ctx.restore()
    | DrawMirrored (flipH, flipV, centerX, centerY, _, commands) ->
        ctx.Ctx.save()
        ctx.Ctx.translate(centerX, centerY)
        ctx.Ctx.scale((if flipH then -1.0 else 1.0), (if flipV then -1.0 else 1.0))
        ctx.Ctx.translate(-centerX, -centerY)
        for command in commands do
            renderPaintCommand ctx command
        ctx.Ctx.restore()

let private commentAnchorCommands (ctx: RenderContext) (doc: Document) (layouts: BlockLayout list) =
    commentAnchorsForLayouts ctx doc layouts
    |> List.collect (fun anchor ->
        [
            FillRect ("rgba(255, 230, 128, 0.42)", anchor.X, anchor.Y, anchor.Width, anchor.Height)
            StrokeRect ("#C79B00", 1.0, anchor.X + 0.5, anchor.Y + 0.5, max 1.0 (anchor.Width - 1.0), max 1.0 (anchor.Height - 1.0))
        ])

type private DirtyRange = {
    Top: float
    Bottom: float
}

type private DirtyDecision =
    | FullRepaint
    | NoRepaint
    | PartialRepaint of DirtyRange

let private canvasWidth (ctx: RenderContext) = float ctx.Canvas.width

let private canvasHeight (ctx: RenderContext) = float ctx.Canvas.height

let private frameInvalidated (ctx: RenderContext) (previous: PaintFrame) =
    previous.CanvasWidth <> canvasWidth ctx
    || previous.CanvasHeight <> canvasHeight ctx
    || previous.FontFamily <> ctx.FontFamily
    || previous.FontSize <> ctx.FontSize
    || previous.LineHeight <> ctx.LineHeight
    || previous.PaddingLeft <> ctx.PaddingLeft
    || previous.PaddingTop <> ctx.PaddingTop
    || previous.PainterVersion <> ctx.TextPathPainterVersion
    || previous.MeasurerVersion <> ctx.TextMeasurerVersion

let private unionRange next current =
    match current, next with
    | None, None -> None
    | Some range, None -> Some range
    | None, Some range -> Some range
    | Some left, Some right ->
        Some { Top = min left.Top right.Top; Bottom = max left.Bottom right.Bottom }

let private rangeForLayout (layout: BlockLayout) =
    { Top = layout.Y; Bottom = layout.Y + layout.Height }

let private rangeForCommand command =
    commandYExtent command
    |> Option.map (fun (top, bottom) -> { Top = top; Bottom = bottom })

let private layoutSame (oldLayout: BlockLayout) (newLayout: BlockLayout) =
    obj.ReferenceEquals(oldLayout, newLayout)

let private allRemainingLayoutsSame (oldTail: BlockLayout list) (newTail: BlockLayout list) =
    let rec loop (left: BlockLayout list) (right: BlockLayout list) =
        match left, right with
        | [], [] -> true
        | oldLayout :: oldRest, newLayout :: newRest when layoutSame oldLayout newLayout -> loop oldRest newRest
        | _ -> false
    loop oldTail newTail

let private firstLayoutRange (oldLayouts: BlockLayout list) (newLayouts: BlockLayout list) canvasBottom =
    let rec loop (oldRemaining: BlockLayout list) (newRemaining: BlockLayout list) =
        match oldRemaining, newRemaining with
        | [], [] -> None
        | oldLayout :: _, [] ->
            Some { Top = oldLayout.Y; Bottom = canvasBottom }
        | [], newLayout :: _ ->
            Some { Top = newLayout.Y; Bottom = canvasBottom }
        | oldLayout :: oldTail, newLayout :: newTail ->
            if layoutSame oldLayout newLayout then
                loop oldTail newTail
            elif oldLayout.BlockId <> newLayout.BlockId then
                Some { Top = min oldLayout.Y newLayout.Y; Bottom = canvasBottom }
            else
                let changedBand =
                    unionRange (Some (rangeForLayout newLayout)) (Some (rangeForLayout oldLayout))
                    |> Option.defaultValue { Top = min oldLayout.Y newLayout.Y; Bottom = max (oldLayout.Y + oldLayout.Height) (newLayout.Y + newLayout.Height) }
                if allRemainingLayoutsSame oldTail newTail then Some changedBand
                else Some { Top = changedBand.Top; Bottom = canvasBottom }
    loop oldLayouts newLayouts

let private commandListRange commands =
    let rec loop remaining current =
        match remaining with
        | [] -> Choice1Of2 current
        | command :: rest ->
            match rangeForCommand command with
            | Some range -> loop rest (unionRange (Some range) current)
            | None -> Choice2Of2 ()
    loop commands None

let private imageBandOf (frame: PaintFrame) =
    let band =
        match frame.ImageBand with
        | Some band -> band
        | None ->
            let mutable top = System.Double.PositiveInfinity
            let mutable bottom = System.Double.NegativeInfinity
            for chunk in frame.Body do
                if chunk.ImageTop <= chunk.ImageBottom then
                    top <- min top chunk.ImageTop
                    bottom <- max bottom chunk.ImageBottom
            let band : Choice<(float * float) option, unit> = Choice1Of2 (if top <= bottom then Some (top, bottom) else None)
            frame.ImageBand <- Some band
            band
    match band with
    | Choice1Of2 range -> Choice1Of2 (range |> Option.map (fun (top, bottom) -> { Top = top; Bottom = bottom }))
    | Choice2Of2 () -> Choice2Of2 ()

let private computeDirtyRange (previous: PaintFrame) (nextFrame: PaintFrame) (canvasBottom: float) =
    // A frame made from the previous one for another selection has the
    // previous one's body commands and layouts - the same objects: nothing of
    // the body is compared, and its pictures are painted already.
    let sameBody = obj.ReferenceEquals(previous.Body, nextFrame.Body)
    let sameDocument = obj.ReferenceEquals(previous.Document, nextFrame.Document)
    // The chunks of what an edit did not change are the same objects: only the others are compared.
    let bodyDiff =
        if sameBody then Choice1Of2 None
        else
            match chunkDifferenceBand previous.Body nextFrame.Body with
            | Choice1Of2 band -> Choice1Of2 (band |> Option.map (fun (top, bottom) -> { Top = top; Bottom = bottom }))
            | Choice2Of2 () -> Choice2Of2 ()
    let staticDiff =
        if obj.ReferenceEquals(previous.StaticOverlay, nextFrame.StaticOverlay) then Choice1Of2 None
        else
            match commandListRange previous.StaticOverlay, commandListRange nextFrame.StaticOverlay with
            | Choice1Of2 oldRange, Choice1Of2 newRange -> Choice1Of2 (unionRange oldRange newRange)
            | _ -> Choice2Of2 ()
    match bodyDiff, staticDiff with
    | Choice2Of2 (), _
    | _, Choice2Of2 () -> FullRepaint
    | Choice1Of2 bodyRange, Choice1Of2 staticRange ->
        match commandListRange previous.OverlayCommands, commandListRange nextFrame.OverlayCommands with
        | Choice2Of2 (), _
        | _, Choice2Of2 () -> FullRepaint
        | Choice1Of2 oldOverlayRange, Choice1Of2 newOverlayRange ->
            let images =
                if sameBody && not sameDocument then Choice1Of2 None, Choice1Of2 None
                else imageBandOf previous, imageBandOf nextFrame
            match images with
            | Choice2Of2 (), _
            | _, Choice2Of2 () -> FullRepaint
            | Choice1Of2 oldImageRange, Choice1Of2 newImageRange ->
                let previousLayouts: BlockLayout list = previous.Layouts
                let nextLayouts: BlockLayout list = nextFrame.Layouts
                let range =
                    (if obj.ReferenceEquals(previousLayouts, nextLayouts) then None else firstLayoutRange previousLayouts nextLayouts canvasBottom)
                    |> unionRange bodyRange
                    |> unionRange staticRange
                    |> unionRange oldOverlayRange
                    |> unionRange newOverlayRange
                    |> unionRange oldImageRange
                    |> unionRange newImageRange
                match range with
                | None -> NoRepaint
                | Some dirty when dirty.Bottom <= dirty.Top -> NoRepaint
                | Some dirty ->
                    let pad = 6.0
                    let top = max 0.0 (floor (dirty.Top - pad))
                    let bottom = min canvasBottom (ceil (dirty.Bottom + pad))
                    if bottom <= top then NoRepaint
                    elif top <= 0.0 && bottom >= canvasBottom then FullRepaint
                    else PartialRepaint { Top = top; Bottom = bottom }

let private commandIntersectsBand band command =
    match commandYExtent command with
    | None -> true
    | Some (top, bottom) -> bottom >= band.Top && top <= band.Bottom

let private renderCommandList ctx commands =
    commands |> List.iter (renderPaintCommand ctx)

let private renderChunks ctx (chunks: PaintChunk array) =
    for chunk in chunks do
        for command in (forceChunk chunk).Commands do
            renderPaintCommand ctx command

/// The commands of the chunks first .. last - 1 that reach into a band (top .. bottom).
let private renderChunkRangeInBand ctx (chunks: PaintChunk array) (first: int) (last: int) (top: float) (bottom: float) =
    for index in first .. last - 1 do
        let chunk = chunks.[index]
        if chunk.Bottom >= top && chunk.Top <= bottom then
            if chunk.Pending.IsSome then ctx.CollectedOnPaint <- ctx.CollectedOnPaint + 1
            let chunk = forceChunk chunk
            let commands = chunk.Commands
            for i = 0 to commands.Length - 1 do
                if chunk.Bottoms.[i] >= top && chunk.Tops.[i] <= bottom then
                    renderPaintCommand ctx commands.[i]

/// What is laid over the body while a header or footer is edited in place.
let private storyDimColor = "rgba(255, 255, 255, 0.6)"

/// The commands of the chunks that reach into a band (top .. bottom).
let private renderChunksInBand ctx (chunks: PaintChunk array) (top: float) (bottom: float) =
    for chunk in chunks do
        if chunk.Bottom >= top && chunk.Top <= bottom then
            // (The commands of a paragraph that was not painted yet are collected now.)
            if chunk.Pending.IsSome then ctx.CollectedOnPaint <- ctx.CollectedOnPaint + 1
            let chunk = forceChunk chunk
            let commands = chunk.Commands
            for i = 0 to commands.Length - 1 do
                if chunk.Bottoms.[i] >= top && chunk.Tops.[i] <= bottom then
                    renderPaintCommand ctx commands.[i]

let private renderCommandListInBand ctx band commands =
    commands
    |> List.iter (fun command ->
        if commandIntersectsBand band command then renderPaintCommand ctx command)

let private renderFullFrame ctx doc frame =
    ctx.Ctx.clearRect(0.0, 0.0, frame.CanvasWidth, frame.CanvasHeight)

    ctx.Ctx.font <- sprintf "%fpx %s" ctx.FontSize ctx.FontFamily
    ctx.Ctx.fillStyle <- U3.Case1 "#000000"

    renderPages ctx frame.Pages
    match ctx.Story with
    | Some story when story.Box.IsSome || story.Note.IsSome ->
        // A text box (a note) edited in place: the page as it is, and over it the box's frame, selection and caret.
        renderChunks ctx frame.Body
        renderCommandList ctx frame.StaticOverlay
        renderCommandList ctx ctx.StoryOverlay
    | Some _ ->
        // The body, dimmed; over it the headers and footers and the story's boundary, selection and caret.
        let furniture = min frame.FurnitureCount frame.Body.Length
        renderChunkRangeInBand ctx frame.Body furniture frame.Body.Length System.Double.NegativeInfinity System.Double.PositiveInfinity
        renderCommandList ctx frame.StaticOverlay
        for page in frame.Pages do
            ctx.Ctx.fillStyle <- U3.Case1 storyDimColor
            ctx.Ctx.fillRect(0.0, page.Top, page.Width, page.Height)
        renderChunkRangeInBand ctx frame.Body 0 furniture System.Double.NegativeInfinity System.Double.PositiveInfinity
        renderCommandList ctx ctx.StoryOverlay
    | None ->
        renderChunks ctx frame.Body
        renderCommandList ctx frame.StaticOverlay
        renderCommandList ctx frame.OverlayCommands

    ctx.Ctx.fillStyle <- U3.Case1 "#000000"

let private renderPartialFrame ctx doc frame band =
    let height = band.Bottom - band.Top
    ctx.Ctx.save()
    ctx.Ctx.beginPath()
    ctx.Ctx.rect(0.0, band.Top, frame.CanvasWidth, height)
    ctx.Ctx.clip()
    ctx.Ctx.clearRect(0.0, band.Top, frame.CanvasWidth, height)

    ctx.Ctx.font <- sprintf "%fpx %s" ctx.FontSize ctx.FontFamily
    ctx.Ctx.fillStyle <- U3.Case1 "#000000"

    renderPages ctx frame.Pages
    renderChunksInBand ctx frame.Body band.Top band.Bottom
    renderCommandListInBand ctx band frame.StaticOverlay
    renderCommandListInBand ctx band frame.OverlayCommands

    ctx.Ctx.restore()
    ctx.Ctx.fillStyle <- U3.Case1 "#000000"

/// The slices of the layout the context holds, when they are the layouts given.
let private slicesFor (ctx: RenderContext) (layouts: BlockLayout list) =
    match ctx.LayoutState.LastSlices, ctx.LayoutState.LastLayouts with
    | Some slices, Some known when obj.ReferenceEquals(known, layouts) -> Some slices
    | _ -> None

/// The band of the document a selection highlight is made for through a
/// viewport: what is visible and a screen above and below it.
let private overlayBandOf (ctx: RenderContext) : (float * float) option =
    match ctx.Viewport with
    | Some vp ->
        let top = (vp.ScrollTop - vp.OriginY) / vp.Scale
        let height = vp.Height / vp.Scale
        Some (top - height, top + 2.0 * height)
    | None -> None

/// The handles of a shape: the name, where on the shape's rectangle (a part
/// of its width and of its height), and the CSS cursor.
let shapeHandleSpots : (string * float * float * string) list =
    [ "nw", 0.0, 0.0, "nwse-resize"; "n", 0.5, 0.0, "ns-resize"; "ne", 1.0, 0.0, "nesw-resize"; "e", 1.0, 0.5, "ew-resize"
      "se", 1.0, 1.0, "nwse-resize"; "s", 0.5, 1.0, "ns-resize"; "sw", 0.0, 1.0, "nesw-resize"; "w", 0.0, 0.5, "ew-resize" ]

/// A shape (text box) where it is painted, for the layouts and pages of a
/// pass: the box, its left edge, top, width and height (before its turn),
/// and the top of its page. A shape of the body; else one of a header or
/// footer on the page whose box is being edited, or the first page that shows it.
let private shapeRectIn (ctx: RenderContext) (settings: LayoutSettings) (doc: Document) (layouts: BlockLayout list) (pages: PageInfo list) (id: BlockId) : (TextBoxBlock * float * float * float * float * float) option =
    let isIt (textBox: TextBoxBlock, _: float, _: float, _: float, _: float) = textBox.Id = id
    let inBody = doc.Blocks |> List.exists (function TextBox box -> box.Id = id | _ -> false)
    if inBody then
        bodyShapeRects settings (measureWithContext ctx) (measureRangeWithContext ctx) doc layouts pages
        |> List.tryFind isIt
        |> Option.map (fun (textBox, x, y, width, height) ->
            let top = pages |> List.tryFindBack (fun page -> page.Top <= y + 0.001) |> Option.map (fun page -> page.Top) |> Option.defaultValue 0.0
            textBox, x, y, width, height, top)
    else
        let inParts (parts: HeaderFooterPart list) =
            parts |> List.exists (fun part -> part.Blocks |> List.exists (function TextBox box -> box.Id = id | _ -> false))
        if not (inParts doc.HeaderParts || inParts doc.FooterParts) then None
        else
            let frames = List.toArray pages
            let onPage (pageIndex: int) =
                if pageIndex < 0 || pageIndex >= frames.Length then None
                else
                    [ true; false ]
                    |> List.tryPick (fun isHeader -> pageFurnitureShapeRects ctx doc layouts pages pageIndex isHeader |> List.tryFind isIt)
                    |> Option.map (fun (textBox, x, y, width, height) -> textBox, x, y, width, height, frames.[pageIndex].Top)
            let edited =
                match ctx.Story with
                | Some story when story.Box = Some id -> onPage story.PageIndex
                | _ -> None
            match edited with
            | Some found -> Some found
            | None ->
                let mutable found = None
                let mutable pageIndex = 0
                while found.IsNone && pageIndex < frames.Length do
                    found <- onPage pageIndex
                    pageIndex <- pageIndex + 1
                found

/// The frame and the eight handles of the shape that is selected (turned as
/// the shape is): the handles are eight pixels on screen whatever the zoom.
let private selectedShapeCommands (ctx: RenderContext) (settings: LayoutSettings) (doc: Document) (layouts: BlockLayout list) (pages: PageInfo list) : PaintCommand list =
    match ctx.SelectedShape with
    | None -> []
    | Some id ->
        match shapeRectIn ctx settings doc layouts pages id with
        | None -> []
        | Some (textBox, x, y, width, height, _) ->
            let scale = match ctx.Viewport with Some vp -> max 0.01 vp.Scale | None -> 1.0
            let size = 8.0 / scale
            let line = 1.0 / scale
            let frame = [ StrokeRect ("#5b9bd5", line, x, y, width, height) ]
            let handles =
                shapeHandleSpots
                |> List.collect (fun (_, fx, fy, _) ->
                    let left, top = x + width * fx - size / 2.0, y + height * fy - size / 2.0
                    [ FillRect ("#ffffff", left, top, size, size)
                      StrokeRect ("#5b9bd5", line, left, top, size, size) ])
            let commands = frame @ handles
            if abs textBox.Rotation >= 0.01 then
                let centerX, centerY = x + width / 2.0, y + height / 2.0
                let reach = sqrt (width * width + height * height) / 2.0 + size + 2.0
                [ DrawGroup (None, textBox.Rotation, centerX, centerY, 1.0, (centerX - reach, centerY - reach, reach * 2.0, reach * 2.0), commands) ]
            else commands

/// The overlay that follows the selection and the pointer, and the band of
/// the document it was made for (None: all of it). A block's layout is found
/// by the block's place when the layout pass handed its layouts over block
/// by block; a selection over several blocks is highlighted in the band only.
let private selectionOverlay (ctx: RenderContext) (settings: LayoutSettings) (doc: Document) (layouts: BlockLayout list) (pages: PageInfo list) (tryFindLayout: (string -> BlockLayout option) option) hoveredImage : PaintCommand list * (float * float) option =
    let slices = slicesFor ctx layouts
    let tryFind =
        match slices, tryFindLayout with
        | Some slices, _ -> Some (sliceLayoutLookup doc slices)
        | None, known -> known
    let banded =
        match slices, overlayBandOf ctx with
        | Some slices, Some (top, bottom) ->
            selectionCommandsInBand settings (measureRangeWithContext ctx) doc slices top bottom |> Option.map (fun commands -> commands, (top, bottom))
        | _ -> None
    let selection, band =
        match banded, tryFind with
        | Some (commands, band), _ -> commands, Some band
        | None, Some tryFind -> selectionCommandsWithRangeLookup settings (measureRangeWithContext ctx) doc tryFind, None
        | None, None -> selectionCommandsWithRange settings (measureRangeWithContext ctx) doc layouts, None
    let hover =
        match tryFind with
        | Some tryFind -> hoverCommandsLookup settings tryFind hoveredImage
        | None -> hoverCommands settings layouts hoveredImage
    let controls =
        if doc.Selection.IsNone then
            ctx.ControlMemberships <- None
            []
        else
            let members =
                match ctx.ControlMemberships with
                | Some (blocks, members) when obj.ReferenceEquals(blocks, doc.Blocks) -> members
                | _ ->
                    let members = contentControlMemberships doc
                    ctx.ControlMemberships <- Some (doc.Blocks, members)
                    members
            markupMarkerCommands members settings doc layouts
    List.concat [ selection; controls; hover; selectedShapeCommands ctx settings doc layouts pages ], band

let private buildFrame (ctx: RenderContext) (doc: Document) hoveredImage =
    ctx.FrameBuilds <- ctx.FrameBuilds + 1
    let settings = layoutSettings ctx doc
    let layouts, layoutPages = layoutDocumentWithPages ctx doc
    // (The pages of the notes after the body are pages too: they have the headers and footers,
    // and the page count the fields show is theirs.)
    let pages = pageFramesForLayouts ctx settings doc layouts layoutPages
    // (The chunks of the pages' headers and footers are the first ones: they lie below the body.)
    let furnitureCount = ref 0
    let body =
        Rbt.Word.Layout.withTablePlanScope ctx.PlanScope (fun () ->
            let chunked =
                match ctx.LayoutState.LastSlices with
                | Some slices -> collectDocumentChunks ctx.PaintCache settings ctx.FontFamily (measureWithContext ctx) (measureRangeWithContext ctx) doc slices pages
                | None -> None
            ctx.LastLaidOutBlocks <- Rbt.Word.Layout.incrementalLayoutBlocksLaidOut ctx.LayoutState
            ctx.LastCollectedBlocks <- (match chunked with Some _ -> paintCacheCollectedBlocks ctx.PaintCache | None -> -1)
            match chunked with
            | Some chunks ->
                furnitureCount.Value <- paintCacheFurnitureCount ctx.PaintCache
                chunks
            | None ->
                // The whole document's commands at once, as one chunk.
                let _, commands, _ = collectDocumentCommandsWithLayoutsAndPages settings ctx.FontFamily (measureWithContext ctx) (measureRangeWithContext ctx) doc layouts pages
                [| chunkOfCommands (List.toArray commands) |])
    let overlay, overlayBand = selectionOverlay ctx settings doc layouts pages None hoveredImage
    {
        Document = doc
        Layouts = layouts
        Pages = pages
        Body = body
        StaticOverlay = commentAnchorCommands ctx doc layouts
        OverlayCommands = overlay
        OverlayBand = overlayBand
        CanvasWidth = canvasWidth ctx
        CanvasHeight = canvasHeight ctx
        FontFamily = ctx.FontFamily
        FontSize = ctx.FontSize
        LineHeight = ctx.LineHeight
        PaddingLeft = ctx.PaddingLeft
        PaddingTop = ctx.PaddingTop
        PainterVersion = ctx.TextPathPainterVersion
        MeasurerVersion = ctx.TextMeasurerVersion
        LayoutSerial = ctx.LayoutState.LayoutSerial
        FurnitureCount = furnitureCount.Value
        Hovered = hoveredImage
        Viewport = ctx.Viewport
        StaticOverlayIndex = None
        LayoutById = None
        ImageBand = None
    }

/// The first layout of a block of the frame, by id.
let private frameLayoutLookup (frame: PaintFrame) =
    let byId =
        match frame.LayoutById with
        | Some byId -> byId
        | None ->
            let byId = System.Collections.Generic.Dictionary<string, BlockLayout>()
            for layout in frame.Layouts do
                if not (byId.ContainsKey layout.BlockId) then byId.[layout.BlockId] <- layout
            frame.LayoutById <- Some byId
            byId
    fun (blockId: string) ->
        match byId.TryGetValue blockId with
        | true, layout -> Some layout
        | false, _ -> None

/// The frame of a document: the previous frame when it shows this document,
/// the previous frame with another overlay when only the selection (or the
/// hover outline) differs - the body commands, the layouts, the pages and
/// their index are the same objects - and a newly built frame otherwise.
/// (The canvas size and the viewport are the caller's to compare.)
let private frameFor (ctx: RenderContext) (doc: Document) hoveredImage =
    // (Most of a long document changed at once: laid out again as a newly loaded one is.)
    deferMassRelayout ctx doc
    // While a long document is being laid out from its start: what is in view
    // and what the selection is in are laid out before they are shown.
    if not (ctx.LayoutState.LastWasWhole && layoutIsWhole ctx) then
        coverSelection ctx doc
        // (The view may be moved to stay on its text when the layout above it changes: then what it shows now is laid out.)
        if ctx.Progress.Anchor.IsSome then layoutDocumentWithPages ctx doc |> ignore
        match ctx.Viewport with
        | Some vp ->
            coverViewport ctx doc vp
            match ctx.Viewport with
            | Some moved when moved.ScrollTop <> vp.ScrollTop -> coverViewport ctx doc moved
            | _ -> ()
        | None -> ()
        layoutDocumentWithPages ctx doc |> ignore
    match ctx.LastPaint with
    | Some previous when not (frameInvalidated ctx { previous with CanvasWidth = canvasWidth ctx; CanvasHeight = canvasHeight ctx })
                         && previous.LayoutSerial = ctx.LayoutState.LayoutSerial
                         && Rbt.Word.Layout.sameDocumentApartFromSelection previous.Document doc ->
        if obj.ReferenceEquals(previous.Document, doc) && previous.Hovered = hoveredImage then previous
        else
            // The layout state follows the document on display (hit tests, the caret rectangle).
            layoutDocumentWithPages ctx doc |> ignore
            let overlay, overlayBand = selectionOverlay ctx (layoutSettings ctx doc) doc previous.Layouts previous.Pages (Some (fun id -> frameLayoutLookup previous id)) hoveredImage
            { previous with Document = doc; Hovered = hoveredImage; OverlayCommands = overlay; OverlayBand = overlayBand }
    | _ -> buildFrame ctx doc hoveredImage

// ---- viewport ---------------------------------------------------------------

[<Emit("$0.setTransform($1, $2, $3, $4, $5, $6)")>]
let private setTransform (_ctx: CanvasRenderingContext2D) (_a: float) (_b: float) (_c: float) (_d: float) (_e: float) (_f: float) : unit = jsNative

[<Emit("$0.style.width = $1 + 'px'; $0.style.height = $2 + 'px'")>]
let private setCanvasCssSize (_canvas: HTMLCanvasElement) (_width: float) (_height: float) : unit = jsNative

let private staticOverlayIndex (frame: PaintFrame) =
    match frame.StaticOverlayIndex with
    | Some index -> index
    | None ->
        let index = chunkOfCommands (List.toArray frame.StaticOverlay)
        frame.StaticOverlayIndex <- Some index
        index

/// Pages are laid out with their left edge at x = 0. Shown through a
/// viewport, a page narrower than the widest one (a portrait section among
/// landscape ones) is centred in the document's width: this is how far it is
/// moved right. Whole-canvas mode leaves pages where the layout put them.
let private pageOffsetX (documentWidth: float) (page: PageInfo) =
    max 0.0 (floor ((documentWidth - page.Width) / 2.0))

let private documentWidthOf (pages: PageInfo list) =
    let mutable width = 0.0
    for page in pages do
        width <- max width page.Width
    width

/// The pages of the frame on display, when there is one.
let private displayedPages (ctx: RenderContext) : PageInfo list =
    match ctx.LastPaint with
    | Some frame -> frame.Pages
    | None -> []

/// How far the page at layout y is moved right on screen (0 in whole-canvas
/// mode, or when all pages are as wide as the document).
let private offsetXAtY (ctx: RenderContext) (y: float) : float =
    match ctx.Viewport with
    | None -> 0.0
    | Some _ ->
        let pages = displayedPages ctx
        let width = documentWidthOf pages
        // The page the point is on, or the one it is nearest to (between pages).
        let mutable best = 0.0
        let mutable bestDistance = System.Double.MaxValue
        for page in pages do
            let distance = if y < page.Top then page.Top - y elif y > page.Top + page.Height then y - (page.Top + page.Height) else 0.0
            if distance < bestDistance then
                bestDistance <- distance
                best <- pageOffsetX width page
        best

/// The x of each page on screen in layout pixels (see `pageOffsetX`),
/// parallel to `pages`: all 0 in whole-canvas mode.
let pageOffsets (ctx: RenderContext) (pages: PageInfo list) : float list =
    match ctx.Viewport with
    | None -> pages |> List.map (fun _ -> 0.0)
    | Some _ ->
        let width = documentWidthOf pages
        pages |> List.map (pageOffsetX width)

/// Layout coordinates of a point given in CSS pixels relative to the canvas.
/// Whole-canvas mode maps one to one (the host scales the canvas itself).
let viewportToLayout (ctx: RenderContext) (x: float) (y: float) =
    match ctx.Viewport with
    | None -> x, y
    | Some vp ->
        let layoutY = (y + vp.ScrollTop - vp.OriginY) / vp.Scale
        (x + vp.ScrollLeft - vp.OriginX) / vp.Scale - offsetXAtY ctx layoutY, layoutY

/// CSS pixels relative to the canvas of a point in layout coordinates.
let layoutToViewport (ctx: RenderContext) (x: float) (y: float) =
    match ctx.Viewport with
    | None -> x, y
    | Some vp -> (x + offsetXAtY ctx y) * vp.Scale + vp.OriginX - vp.ScrollLeft, y * vp.Scale + vp.OriginY - vp.ScrollTop

/// CSS pixels per layout pixel on the canvas (1 in whole-canvas mode).
let viewportScale (ctx: RenderContext) =
    match ctx.Viewport with
    | None -> 1.0
    | Some vp -> vp.Scale

/// The band of the document (layout coordinates) a viewport shows.
let private visibleBand (vp: Viewport) =
    { Top = (vp.ScrollTop - vp.OriginY) / vp.Scale; Bottom = (vp.ScrollTop - vp.OriginY + vp.Height) / vp.Scale }

/// Make the canvas exactly as large as the viewport. True when the backing
/// store was resized (which clears it).
let private fitCanvasToViewport (ctx: RenderContext) (vp: Viewport) =
    let width = max 1.0 (round (vp.Width * vp.DevicePixelRatio))
    let height = max 1.0 (round (vp.Height * vp.DevicePixelRatio))
    setCanvasCssSize ctx.Canvas vp.Width vp.Height
    if ctx.Canvas.width <> width || ctx.Canvas.height <> height then
        ctx.Canvas.width <- width
        ctx.Canvas.height <- height
        true
    else
        false

let private deskColor = "#eef0f3"

let private renderPageSheet (ctx: RenderContext) (page: PageInfo) =
    ctx.Ctx.fillStyle <- U3.Case1 "#ffffff"
    ctx.Ctx.fillRect(0.0, page.Top, page.Width, page.Height)
    ctx.Ctx.strokeStyle <- U3.Case1 "#d0d5dd"
    ctx.Ctx.lineWidth <- 1.0
    setLineDash ctx.Ctx [||]
    ctx.Ctx.strokeRect(0.5, page.Top + 0.5, page.Width - 1.0, page.Height - 1.0)

let private renderPagesInBand (ctx: RenderContext) (pages: PageInfo list) top bottom =
    for page in pages do
        if page.Top + page.Height >= top && page.Top <= bottom then
            renderPageSheet ctx page

/// Paint a band of the document through the viewport transform: the desk,
/// the pages and the commands that intersect it, and nothing else.
let private renderViewportBand (ctx: RenderContext) (frame: PaintFrame) (vp: Viewport) (band: DirtyRange) =
    let deviceScale = vp.Scale * vp.DevicePixelRatio
    // Whole device pixels, so text and hairlines stay sharp while scrolling.
    let offsetX = round ((vp.OriginX - vp.ScrollLeft) * vp.DevicePixelRatio)
    let offsetY = round ((vp.OriginY - vp.ScrollTop) * vp.DevicePixelRatio)
    let left = -offsetX / deviceScale
    let width = float ctx.Canvas.width / deviceScale
    let top = band.Top
    let height = band.Bottom - band.Top
    ctx.Ctx.save()
    setTransform ctx.Ctx deviceScale 0.0 0.0 deviceScale offsetX offsetY
    ctx.Ctx.beginPath()
    ctx.Ctx.rect(left, top, width, height)
    ctx.Ctx.clip()
    ctx.Ctx.fillStyle <- U3.Case1 deskColor
    ctx.Ctx.fillRect(left, top, width, height)
    ctx.Ctx.font <- sprintf "%fpx %s" ctx.FontSize ctx.FontFamily
    // Page by page: each is painted at its own x (a narrower page is centred
    // in the document's width) and clipped to its sheet, so nothing of a page
    // - a table border running on to the next one - shows in the gap between
    // pages.
    let documentWidth = documentWidthOf frame.Pages
    for page in frame.Pages do
        let pageTop = max top page.Top
        let pageBottom = min band.Bottom (page.Top + page.Height)
        if pageBottom > pageTop then
            ctx.Ctx.save()
            ctx.Ctx.translate(pageOffsetX documentWidth page, 0.0)
            ctx.Ctx.beginPath()
            ctx.Ctx.rect(0.0, pageTop, page.Width, pageBottom - pageTop)
            ctx.Ctx.clip()
            renderPageSheet ctx page
            match ctx.Story with
            | Some story when story.Box.IsSome || story.Note.IsSome ->
                // A text box (a note) edited in place: the page as it is, and over it the box's frame, selection and caret.
                renderChunksInBand ctx frame.Body pageTop pageBottom
                renderChunksInBand ctx [| staticOverlayIndex frame |] pageTop pageBottom
                renderCommandListInBand ctx { Top = pageTop; Bottom = pageBottom } ctx.StoryOverlay
            | Some _ ->
                // The body, dimmed; over it the headers and footers and the story's boundary, selection and caret.
                let furniture = min frame.FurnitureCount frame.Body.Length
                renderChunkRangeInBand ctx frame.Body furniture frame.Body.Length pageTop pageBottom
                renderChunksInBand ctx [| staticOverlayIndex frame |] pageTop pageBottom
                ctx.Ctx.fillStyle <- U3.Case1 storyDimColor
                ctx.Ctx.fillRect(0.0, pageTop, page.Width, pageBottom - pageTop)
                renderChunkRangeInBand ctx frame.Body 0 furniture pageTop pageBottom
                renderCommandListInBand ctx { Top = pageTop; Bottom = pageBottom } ctx.StoryOverlay
            | None ->
                renderChunksInBand ctx frame.Body pageTop pageBottom
                renderChunksInBand ctx [| staticOverlayIndex frame |] pageTop pageBottom
                renderCommandListInBand ctx { Top = pageTop; Bottom = pageBottom } frame.OverlayCommands
            ctx.Ctx.restore()
    ctx.Ctx.restore()
    ctx.Ctx.fillStyle <- U3.Case1 "#000000"

let private renderThroughViewport (ctx: RenderContext) (doc: Document) hoveredImage (vp: Viewport) =
    let resized = fitCanvasToViewport ctx vp
    // Scrolling and zooming show the same document: its commands are reused.
    let frame =
        let frame = frameFor ctx doc hoveredImage
        // A selection highlight made for a band of the document: made again
        // for the band in view once the view leaves it.
        let frame =
            match frame.OverlayBand with
            | Some (top, bottom) ->
                let viewTop = (vp.ScrollTop - vp.OriginY) / vp.Scale
                let viewBottom = viewTop + vp.Height / vp.Scale
                if viewTop >= top && viewBottom <= bottom then frame
                else
                    let overlay, overlayBand = selectionOverlay ctx (layoutSettings ctx doc) doc frame.Layouts frame.Pages None hoveredImage
                    { frame with OverlayCommands = overlay; OverlayBand = overlayBand }
            | None -> frame
        if frame.Viewport = Some vp && not resized then frame
        else { frame with Viewport = Some vp; CanvasWidth = canvasWidth ctx; CanvasHeight = canvasHeight ctx }
    let visible = visibleBand vp
    // One layout pixel of slack either side: rounding the offset to device pixels can show a sliver more.
    let visible = { Top = visible.Top - 1.0; Bottom = visible.Bottom + 1.0 }
    match ctx.LastPaint with
    // (With a story on display, or just gone: what is in view is painted again - the dimming is not part of a frame.)
    | _ when ctx.Story.IsSome || ctx.StoryPainted -> renderViewportBand ctx frame vp visible
    | Some previous when obj.ReferenceEquals(previous, frame) -> ()
    | Some previous when not resized && previous.Viewport = Some vp && not (frameInvalidated ctx previous) && previous.Pages = frame.Pages ->
        match computeDirtyRange previous frame System.Double.MaxValue with
        | NoRepaint -> ()
        | FullRepaint -> renderViewportBand ctx frame vp visible
        | PartialRepaint band ->
            let top = max band.Top visible.Top
            let bottom = min band.Bottom visible.Bottom
            if bottom > top then renderViewportBand ctx frame vp { Top = top; Bottom = bottom }
    | _ -> renderViewportBand ctx frame vp visible
    ctx.LastPaint <- Some frame

let private renderBody (ctx: RenderContext) (doc: Document) hoveredImage =
    match ctx.Viewport with
    | Some vp -> renderThroughViewport ctx doc hoveredImage vp
    | None ->
    let frame =
        let frame = frameFor ctx doc hoveredImage
        if frame.Viewport.IsNone && frame.CanvasWidth = canvasWidth ctx && frame.CanvasHeight = canvasHeight ctx then frame
        else { frame with Viewport = None; CanvasWidth = canvasWidth ctx; CanvasHeight = canvasHeight ctx }
    match ctx.LastPaint with
    | None ->
        renderFullFrame ctx doc frame
    | _ when ctx.Story.IsSome || ctx.StoryPainted ->
        renderFullFrame ctx doc frame
    | Some previous when previous.Viewport.IsSome || frameInvalidated ctx previous || previous.Pages <> frame.Pages ->
        renderFullFrame ctx doc frame
    | Some previous ->
        match computeDirtyRange previous frame frame.CanvasHeight with
        | FullRepaint -> renderFullFrame ctx doc frame
        | NoRepaint -> ()
        | PartialRepaint band -> renderPartialFrame ctx doc frame band
    ctx.LastPaint <- Some frame

let private renderCore (ctx: RenderContext) (doc: Document) hoveredImage =
    match ctx.Story with
    | None ->
        renderBody ctx doc hoveredImage
        ctx.StoryPainted <- false
    | Some story ->
        // A header or footer edited in place: the whole document without a selection, and over
        // it the story's boundary, its selection and its caret (see renderViewportBand).
        let whole = story.Whole doc
        let settings = layoutSettings ctx whole
        let layouts = storyLayouts ctx story doc
        let tryFind (blockId: string) = layouts |> List.tryFind (fun layout -> layout.BlockId = blockId)
        let boundary =
            match story.Note with
            | Some (kind, id) -> noteStoryCommands ctx whole layouts kind id
            | None ->
            let bodyLayouts, layoutPages = layoutDocumentWithPages ctx whole
            let pages = pageFramesForLayouts ctx settings whole bodyLayouts layoutPages
            match story.Box with
            | Some boxId ->
                // A text box: a frame around it.
                match bodyShapeRects settings (measureWithContext ctx) (measureRangeWithContext ctx) whole bodyLayouts pages |> List.tryFind (fun (textBox, _, _, _, _) -> textBox.Id = boxId) with
                | Some (_, x, y, width, height) ->
                    [ StrokeLine ("#5b9bd5", 1.0, 2, 0.0, x - 2.5, y - 2.5, x + width + 2.5, y - 2.5)
                      StrokeLine ("#5b9bd5", 1.0, 2, 0.0, x + width + 2.5, y - 2.5, x + width + 2.5, y + height + 2.5)
                      StrokeLine ("#5b9bd5", 1.0, 2, 0.0, x + width + 2.5, y + height + 2.5, x - 2.5, y + height + 2.5)
                      StrokeLine ("#5b9bd5", 1.0, 2, 0.0, x - 2.5, y + height + 2.5, x - 2.5, y - 2.5) ]
                | None -> []
            | None ->
            match List.tryItem story.PageIndex pages with
            | Some page -> storyBoundaryCommands settings ctx.FontFamily page story.IsHeader story.Label
            | None -> []
        let overlay =
            List.concat [ boundary
                          selectionCommandsWithRangeLookup settings (measureRangeWithContext ctx) doc tryFind
                          hoverCommandsLookup settings tryFind hoveredImage ]
        ctx.StoryOverlay <-
            // (The frame, the selection and the caret of a turned box are turned as its text is.)
            match ctx.StoryTurn, story.Box with
            | Some (degrees, centerX, centerY), Some _ when not overlay.IsEmpty ->
                let mutable radius = 0.0
                for layout in layouts do
                    for line in layout.Lines do
                        radius <- max radius (abs (line.Y - centerY) + line.Height)
                // (Everything of the overlay is within the box or its text: as far from the centre as the document is wide at most.)
                let reach = max radius (settings.Width + settings.PageHeight)
                [ DrawGroup (None, degrees, centerX, centerY, 1.0, (centerX - reach, centerY - reach, reach * 2.0, reach * 2.0), overlay) ]
            | _ -> overlay
        renderBody ctx whole None
        ctx.StoryPainted <- true

/// The page a point (layout coordinates) is on and the part of it: "header"
/// (above the body text), "footer" (below it) or "body". None between pages.
let pageRegionAt (ctx: RenderContext) (doc: Document) (x: float) (y: float) : (int * string) option =
    let doc = bodyDocument ctx doc
    let layouts, layoutPages = layoutDocumentWithPages ctx doc
    let pages = pageFramesForLayouts ctx (layoutSettings ctx doc) doc layouts layoutPages
    let mutable found = None
    let mutable index = 0
    for page in pages do
        if found.IsNone && y >= page.Top && y <= page.Top + page.Height then
            let region =
                if y < page.Top + page.MarginTop then "header"
                elif y > page.Top + page.Height - page.MarginBottom then "footer"
                else "body"
            found <- Some (index, region)
        index <- index + 1
    found

/// The header (or footer) a page shows, for editing it in place (see
/// Paint.HeaderFooterTarget). The whole document is laid out for it.
let headerFooterTarget (ctx: RenderContext) (doc: Document) (pageIndex: int) (isHeader: bool) : HeaderFooterTarget option =
    let doc = bodyDocument ctx doc
    let layouts, layoutPages = layoutWholeDocumentWithPages ctx doc
    let pages = pageFramesForLayouts ctx (layoutSettings ctx doc) doc layouts layoutPages
    let sectionEnds = layouts |> List.filter (fun layout -> endsSection layout.Metadata)
    headerFooterTargetForPage (layoutSettings ctx doc) doc sectionEnds pages pageIndex isHeader

/// What the running fields (PAGE, NUMPAGES, SECTIONPAGES) show on a page,
/// from the layout there is (not forced to the end of the document: see
/// `layoutIsComplete`). None when there is no such page.
let runningFieldsOfPage (ctx: RenderContext) (doc: Document) (pageIndex: int) : RunningFieldValues option =
    let doc = bodyDocument ctx doc
    let layouts, layoutPages = layoutDocumentWithPages ctx doc
    let pages = pageFramesForLayouts ctx (layoutSettings ctx doc) doc layouts layoutPages
    let sectionEnds = layouts |> List.filter (fun layout -> endsSection layout.Metadata)
    headerFooterTargetForPage (layoutSettings ctx doc) doc sectionEnds pages pageIndex true
    |> Option.map (fun target -> target.TargetFields)

/// Show the document through a viewport (or, with None, go back to the whole
/// document on a host-sized canvas). The caller repaints afterwards.
let setViewport (ctx: RenderContext) (viewport: Viewport option) =
    match ctx.Viewport, viewport with
    | Some _, None ->
        // The canvas is the host's again; whatever is on it is stale.
        setTransform ctx.Ctx 1.0 0.0 0.0 1.0 0.0 0.0
        ctx.LastPaint <- None
    | _ -> ()
    // (The host says where the view is: it is not moved to stay on a block any more.)
    ctx.Progress.Anchor <- None
    ctx.Progress.ScrollAdjusted <- false
    ctx.Viewport <- viewport

/// Size of the document in layout pixels: the widest page and the bottom of
/// the last one. A host multiplies by its zoom to size the scrollable area.
let documentSize (ctx: RenderContext) (doc: Document) =
    let doc = bodyDocument ctx doc
    let frames =
        match ctx.LastPaint with
        | Some frame when obj.ReferenceEquals(frame.Document, doc) && frame.LayoutSerial = ctx.LayoutState.LayoutSerial && Rbt.Word.Layout.incrementalLayoutIsCurrent ctx.LayoutState -> frame.Pages
        | _ ->
            let layouts, pages = layoutDocumentWithPages ctx doc
            pageFramesForLayouts ctx (layoutSettings ctx doc) doc layouts pages
    let mutable width = 0.0
    let mutable height = 0.0
    for page in frames do
        width <- max width page.Width
        height <- max height (page.Top + page.Height)
    width, height, List.length frames

/// Paint one page of the document onto a canvas of the host's, which is
/// resized to the page at `scale` (backing-store pixels per layout pixel).
/// Read-only: no selection, caret or hover is drawn. The document's paint
/// commands are collected once and reused for every page, so painting all
/// pages costs one pass over the document, not one per page.
/// Text is measured with the context's measurers, so the pages are the pages
/// of the editor / viewer the context belongs to. A text path painter draws
/// on one canvas: the page is painted with the painter the context's factory
/// (`TextPathPainterFactory`) makes for the target canvas, or with canvas
/// text when there is none.
let renderPage (ctx: RenderContext) (doc: Document) (canvas: HTMLCanvasElement) (pageIndex: int) (scale: float) =
    let doc = bodyDocument ctx doc
    let usable (frame: PaintFrame) =
        obj.ReferenceEquals(frame.Document, doc)
        && frame.LayoutSerial = ctx.LayoutState.LayoutSerial
        && Rbt.Word.Layout.incrementalLayoutIsCurrent ctx.LayoutState
        && not (frameInvalidated ctx { frame with CanvasWidth = canvasWidth ctx; CanvasHeight = canvasHeight ctx })
    let frameNow () =
        match ctx.LastPaint, ctx.PageFrame with
        | Some frame, _ when usable frame -> frame
        | _, Some frame when usable frame -> frame
        | _ ->
            let frame = buildFrame ctx doc None
            ctx.PageFrame <- Some frame
            frame
    let frame =
        let frame = frameNow ()
        // A page past what is laid out of a long document so far: the whole document is.
        if pageIndex >= List.length frame.Pages && pageIndex >= 0 && not (layoutIsWhole ctx && ctx.LayoutState.LastWasWhole) then
            completeLayout ctx
            frameNow ()
        else frame
    match List.tryItem pageIndex frame.Pages with
    | None -> false
    | Some page ->
        canvas.width <- max 1.0 (round (page.Width * scale))
        canvas.height <- max 1.0 (round (page.Height * scale))
        let targetContext = canvas.getContext_2d()
        let painter =
            match ctx.TextPathPainterFactory with
            | Some factory ->
                let paint = factory.Invoke targetContext
                if isNull (box paint) then None
                else Some (fun (text: string) (style: TextPaintStyle) (x: float) (baseline: float) (width: float) -> paint.Invoke(text, style, x, baseline, width))
            | None -> None
        let target = { ctx with Canvas = canvas; Ctx = targetContext; TextPathPainter = painter }
        let bottom = page.Top + page.Height
        target.Ctx.save()
        setTransform target.Ctx scale 0.0 0.0 scale 0.0 (-page.Top * scale)
        target.Ctx.beginPath()
        target.Ctx.rect(0.0, page.Top, page.Width, page.Height)
        target.Ctx.clip()
        target.Ctx.font <- sprintf "%fpx %s" ctx.FontSize ctx.FontFamily
        renderPagesInBand target [ page ] page.Top bottom
        renderChunksInBand target frame.Body page.Top bottom
        target.Ctx.restore()
        // (The target is a copy of the context: what it counted is the context's.)
        ctx.CollectedOnPaint <- target.CollectedOnPaint
        true

let private fontFacesToJs (faces: (string * bool * bool) list) : obj array =
    faces
    |> List.map (fun (family, bold, italic) ->
        Fable.Core.JsInterop.createObj [ "family", box family; "bold", box bold; "italic", box italic ])
    |> List.toArray

/// The font faces the document on a renderer is set in (see
/// Paint.requiredFontFaces), as [{ family, bold, italic }], from the layout
/// the renderer has.
let requiredFonts (ctx: RenderContext) (doc: Document) : obj array =
    let doc = bodyDocument ctx doc
    let layouts, pages = layoutWholeDocumentWithPages ctx doc
    Rbt.Word.Layout.withTablePlanScope ctx.PlanScope (fun () ->
        requiredFontFaces (layoutSettings ctx doc) ctx.FontFamily (measureWithContext ctx) (measureRangeWithContext ctx) doc (Some (layouts, pages)))
    |> fontFacesToJs

/// The font faces a document is set in, without a renderer: before a host
/// has fonts to measure with. (Which faces are used does not depend on the
/// measurements.)
let requiredFontsOfDocument (doc: Document) : obj array =
    let measureText (text: string) (style: TextBuffer.TextStyle) = float text.Length * style.FontSize * 0.5
    requiredFontFaces (documentSettings 16.0 24.0 3.5 doc) "Calibri" measureText (rangeMeasurerFromTextMeasurer measureText) doc None
    |> fontFacesToJs

/// Counters of a renderer, for hosts and checks (see API.getRenderStats).
let renderStats (ctx: RenderContext) (doc: Document) : obj =
    Fable.Core.JsInterop.createObj [
        "FrameBuilds", box ctx.FrameBuilds
        "LaidOutBlocks", box ctx.LastLaidOutBlocks
        "CollectedBlocks", box ctx.LastCollectedBlocks
        "CollectedOnPaint", box ctx.CollectedOnPaint
        "Blocks", box (List.length doc.Blocks)
        "PainterVersion", box ctx.TextPathPainterVersion
        "MeasurerVersion", box ctx.TextMeasurerVersion
        "LayoutComplete", box (layoutIsComplete ctx)
        "LayoutPasses", box ctx.LayoutState.LayoutSerial
        "LaidOutLimit", box (if layoutIsWhole ctx then List.length doc.Blocks else min ctx.LayoutState.BlockLimit (List.length doc.Blocks))
        // Blocks above the view that are shown as they were laid out before a change of most of the document (0: none).
        "StaleBlocks", box (match ctx.Stale with Some stale -> stale.Keep | None -> 0)
        // The slices a long document was laid out in since it was loaded: how many, and the longest -
        // what it took with the repaint and the listeners that followed it, which it was (1: the
        // first), its passes, the layout's part of its time, the blocks that were laid out before it.
        "LayoutSlices", box ctx.Progress.Slices
        "LongestLayoutSliceMs", box ctx.Progress.LongestSliceMs
        "LongestLayoutSlice", box ctx.Progress.LongestSlice
        "LongestLayoutSlicePasses", box ctx.Progress.LongestSlicePasses
        "LongestLayoutSliceLayoutMs", box ctx.Progress.LongestSliceLayoutMs
        "LongestLayoutSliceFrom", box ctx.Progress.LongestSliceFrom
    ]

/// Render document to canvas
let render (ctx: RenderContext) (doc: Document) =
    renderCore ctx doc None

/// The frames of the pages laid out, including footnote/endnote areas. Hosts
/// use them to size the canvas (pages are stacked with pageGap between them).
/// While a long document is being laid out from its start these are the
/// pages there are so far (see `layoutIsComplete`).
let pageFrames (ctx: RenderContext) (doc: Document) =
    let doc = bodyDocument ctx doc
    let layouts, pages = layoutDocumentWithPages ctx doc
    pageFramesForLayouts ctx (layoutSettings ctx doc) doc layouts pages

/// Where the shapes (text boxes) of the body are painted: each with its left
/// edge from the left edge of its page, its top down the canvas (layout
/// pixels), its width and its height.
let shapeRects (ctx: RenderContext) (doc: Document) : (TextBoxBlock * float * float * float * float) list =
    let doc = bodyDocument ctx doc
    if not (doc.Blocks |> List.exists (function TextBox _ -> true | _ -> false)) then []
    else
        let layouts, pages = layoutDocumentWithPages ctx doc
        let settings = layoutSettings ctx doc
        bodyShapeRects settings (measureWithContext ctx) (measureRangeWithContext ctx) doc layouts (pageFramesForLayouts ctx settings doc layouts pages)

/// The text boxes of the headers and footers, each on the first page (of
/// those laid out) that shows its part: the box, its left edge, top, width
/// and height as `shapeRects` gives them, whether it is a header's, and the
/// page.
let furnitureShapeRects (ctx: RenderContext) (doc: Document) : (TextBoxBlock * float * float * float * float * bool * int) list =
    let doc = bodyDocument ctx doc
    let boxesOf (parts: HeaderFooterPart list) =
        parts |> List.sumBy (fun part -> part.Blocks |> List.sumBy (function TextBox _ -> 1 | _ -> 0))
    let total = boxesOf doc.HeaderParts + boxesOf doc.FooterParts
    if total = 0 then []
    else
        let layouts, layoutPages = layoutDocumentWithPages ctx doc
        let pages = pageFramesForLayouts ctx (layoutSettings ctx doc) doc layouts layoutPages
        let seen = System.Collections.Generic.HashSet<BlockId>()
        let found = ResizeArray<TextBoxBlock * float * float * float * float * bool * int>()
        let mutable pageIndex = 0
        let pageCount = List.length pages
        while pageIndex < pageCount && seen.Count < total do
            for isHeader in [ true; false ] do
                for (textBox, x, y, width, height) in pageFurnitureShapeRects ctx doc layouts pages pageIndex isHeader do
                    if seen.Add textBox.Id then found.Add ((textBox, x, y, width, height, isHeader, pageIndex))
            pageIndex <- pageIndex + 1
        List.ofSeq found

/// The text boxes of the header and of the footer a page shows, where that
/// page paints them: the box, its left edge, top, width and height as
/// `shapeRects` gives them, and whether it is the header's.
let pageShapeRects (ctx: RenderContext) (doc: Document) (pageIndex: int) : (TextBoxBlock * float * float * float * float * bool) list =
    let doc = bodyDocument ctx doc
    let hasBoxes (parts: HeaderFooterPart list) =
        parts |> List.exists (fun part -> part.Blocks |> List.exists (function TextBox _ -> true | _ -> false))
    if pageIndex < 0 || not (hasBoxes doc.HeaderParts || hasBoxes doc.FooterParts) then []
    else
        let layouts, layoutPages = layoutDocumentWithPages ctx doc
        let pages = pageFramesForLayouts ctx (layoutSettings ctx doc) doc layouts layoutPages
        [ for isHeader in [ true; false ] do
            for (textBox, x, y, width, height) in pageFurnitureShapeRects ctx doc layouts pages pageIndex isHeader do
                yield textBox, x, y, width, height, isHeader ]

/// The text box of a header or footer that a point (layout coordinates) is
/// in, on the page the point is on: its id, whether it is the header's, and
/// the page. (Not a watermark, whose text is drawn as large as its box; a
/// box behind the text is found - the text over it is the body's, and the
/// point is looked up for a header that is edited or about to be.)
let furnitureShapeAt (ctx: RenderContext) (doc: Document) (x: float) (y: float) : (BlockId * bool * int) option =
    match pageRegionAt ctx doc x y with
    | None -> None
    | Some (pageIndex, _) ->
        pageShapeRects ctx doc pageIndex
        |> List.filter (fun (textBox, left, top, width, height, _) ->
            let x, y = if abs textBox.Rotation >= 0.01 then turnedPoint -textBox.Rotation (left + width / 2.0) (top + height / 2.0) x y else x, y
            not textBox.FitText && x >= left && x <= left + width && y >= top && y <= top + height)
        |> List.tryLast
        |> Option.map (fun (textBox, _, _, _, _, isHeader) -> textBox.Id, isHeader, pageIndex)

/// A shape (text box) of the body, a header or a footer where it is painted:
/// the box, its left edge, top, width and height (before its turn), and the
/// top of its page. (A header's or footer's: on the first page that shows it,
/// or - while its text is edited - on the page it is edited on.)
let shapeRectOf (ctx: RenderContext) (doc: Document) (id: BlockId) : (TextBoxBlock * float * float * float * float * float) option =
    let doc = bodyDocument ctx doc
    let layouts, layoutPages = layoutDocumentWithPages ctx doc
    let settings = layoutSettings ctx doc
    shapeRectIn ctx settings doc layouts (pageFramesForLayouts ctx settings doc layouts layoutPages) id

/// What of a shape is at a point (layout coordinates; for a turned shape the
/// point is turned back first): the name of the handle there ("nw" ... "w"),
/// "border" on its frame (five pixels on screen to either side), "inside",
/// or "" outside it (and when there is no such shape).
let shapePartAt (ctx: RenderContext) (doc: Document) (id: BlockId) (pointX: float) (pointY: float) : string =
    match shapeRectOf ctx doc id with
    | None -> ""
    | Some (textBox, x, y, width, height, _) ->
        let pointX, pointY =
            if abs textBox.Rotation >= 0.01 then turnedPoint -textBox.Rotation (x + width / 2.0) (y + height / 2.0) pointX pointY else pointX, pointY
        let reach = 5.0 / max 0.01 (viewportScale ctx)
        match shapeHandleSpots |> List.tryFind (fun (_, fx, fy, _) -> abs (pointX - (x + width * fx)) <= reach && abs (pointY - (y + height * fy)) <= reach) with
        | Some (name, _, _, _) -> name
        | None ->
            let within (margin: float) = pointX >= x + margin && pointX <= x + width - margin && pointY >= y + margin && pointY <= y + height - margin
            if not (within -reach) then "" elif within reach then "inside" else "border"

/// Where a shape is in the order shapes are painted in (wp:anchor relativeHeight): the larger, the more in front.
let private paintOrderOf (textBox: TextBoxBlock) =
    match textBox.RelativeHeight with
    | Some text ->
        match System.Double.TryParse text with
        | true, number -> number
        | false, _ -> 0.0
    | None -> 0.0

/// The text box in front of the text that a point (layout coordinates) is in
/// - the one painted last of those that hold it. A box behind the text, and
/// one whose text is drawn as large as the box (a watermark), is not found by
/// a point: the text over it is.
let shapeAt (ctx: RenderContext) (doc: Document) (x: float) (y: float) : BlockId option =
    shapeRects ctx doc
    |> List.filter (fun (textBox, left, top, width, height) ->
        // (A turned box: the point as it is in the box before the turn.)
        let x, y = if abs textBox.Rotation >= 0.01 then turnedPoint -textBox.Rotation (left + width / 2.0) (top + height / 2.0) x y else x, y
        not (textBox.BehindDoc = Some true) && not textBox.FitText && x >= left && x <= left + width && y >= top && y <= top + height)
    |> List.sortBy (fun (textBox, _, _, _, _) -> paintOrderOf textBox)
    |> List.tryLast
    |> Option.map (fun (textBox, _, _, _, _) -> textBox.Id)

/// The shape of the body whose frame is at a point (layout coordinates):
/// five pixels on screen to either side of its edge (shapePartAt "border",
/// or a handle's place), the one in front first. Any shape: one behind the
/// text, and one without text, too. (Of a box in front of the text that has
/// text, the outer side only: within it a click goes to its text.)
let shapeEdgeAt (ctx: RenderContext) (doc: Document) (x: float) (y: float) : BlockId option =
    let hasText (textBox: TextBoxBlock) = textBox.Content |> List.exists (fun para -> TextBuffer.length para.Buffer > 0)
    shapeRects ctx doc
    |> List.sortBy (fun (textBox, _, _, _, _) -> paintOrderOf textBox)
    |> List.rev
    |> List.tryPick (fun (textBox, left, top, width, height) ->
        match shapePartAt ctx doc textBox.Id x y with
        | "" | "inside" -> None
        | _ ->
            let px, py = if abs textBox.Rotation >= 0.01 then turnedPoint -textBox.Rotation (left + width / 2.0) (top + height / 2.0) x y else x, y
            let within = px >= left && px <= left + width && py >= top && py <= top + height
            let textReached = not (textBox.BehindDoc = Some true) && not textBox.FitText && hasText textBox
            if within && textReached then None else Some textBox.Id)

/// Number of pages the document lays out to, including footnote/endnote areas.
/// Pages are stacked with pageGap between them, but need not be the same size
/// (sections); hosts size the canvas from pageFrames. The whole document is
/// laid out for it.
let pageCount (ctx: RenderContext) (doc: Document) =
    completeLayout ctx
    List.length (pageFrames ctx doc)

/// Convert mouse position to document position
let hitTest (ctx: RenderContext) (doc: Document) (mouseX: float) (mouseY: float) =
    match ctx.Story with
    | Some story ->
        // In the story: the position nearest to the point, wherever on the page it is.
        let layouts = storyLayouts ctx story doc |> List.toArray
        if layouts.Length = 0 then None
        else
            let mutable top = System.Double.PositiveInfinity
            let mutable bottom = System.Double.NegativeInfinity
            for layout in layouts do
                top <- min top layout.Y
                bottom <- max bottom (layout.Y + layout.Height)
            // (In a turned text box: the point as it is in the upright layout.)
            let mouseX, mouseY = storyPoint ctx mouseX mouseY
            let y = max (top + 0.5) (min (bottom - 0.5) mouseY)
            hitTestLayoutArrayWithRange (measureRangeWithContext ctx) (layoutSettings ctx doc).PaddingLeft layouts mouseX y
    | None ->
    // Same layout the last render produced (reused while the document is
    // unchanged) rather than a fresh full layout per mouse event.
    Rbt.Word.Layout.hitTestIncrementalWithRange ctx.LayoutState (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) (layoutTarget ctx doc) mouseX mouseY

/// How the caret is turned (degrees clockwise, about its middle): 0 but in
/// the text of a turned text box that is edited in place.
let caretTurn (ctx: RenderContext) : float =
    match ctx.Story, ctx.StoryTurn with
    | Some story, Some (degrees, _, _) when story.Box.IsSome -> degrees
    | _ -> 0.0

/// Where the caret (the selection focus) is in the layout that is on
/// screen: x, top and height of its line - upright, also in a turned text
/// box (see caretRect).
let private uprightCaretRect (ctx: RenderContext) (doc: Document) : (float * float * float) option =
    match doc.Selection with
    | None -> None
    | Some selection ->
        let settings = layoutSettings ctx doc
        let offset = selection.Focus.Offset
        if ctx.Story.IsNone then coverSelection ctx doc
        // A paragraph continued in another column has one layout per fragment.
        let mutable found: BlockLayout option = None
        let mutable fallback: BlockLayout option = None
        let consider (layout: BlockLayout) =
            if layout.BlockId = selection.Focus.BlockId then
                fallback <- Some layout
                if found.IsNone && layout.Lines |> List.exists (fun line -> line.StartOffset <= offset && offset <= line.EndOffset) then
                    found <- Some layout
        let layouts = layoutDocument ctx doc
        // The layouts of the caret's block, by the block's place, when the pass handed its layouts
        // over block by block: asked with every keystroke, and a walk of all the layouts of a long
        // document was a fifth of a keystroke's time.
        let ofBlock =
            if ctx.Story.IsSome then None
            else
                match ctx.LayoutState.LastSlices, ctx.LayoutState.LastLayouts with
                | Some slices, Some known when obj.ReferenceEquals(known, layouts) ->
                    match locateBlock (layoutTarget ctx doc) selection.Focus.BlockId with
                    | ValueSome (struct (index, _)) when index < slices.SliceBlocks.Length -> Some (slices, index)
                    | _ -> None
                | _ -> None
        match ofBlock with
        | Some (slices, index) ->
            for at in slices.SliceStarts.[index] .. slices.SliceStarts.[index + 1] - 1 do consider slices.SliceLayouts.[at]
            // (Not among them - a note's paragraph: all the layouts.)
            if fallback.IsNone then (for layout in layouts do consider layout)
        | None -> for layout in layouts do consider layout
        match found |> Option.orElse fallback with
        | None -> None
        | Some layout when layout.Lines.IsEmpty ->
            // An empty paragraph: where its text would start (its indent - for
            // a paragraph of a table cell that is the cell's place - and its alignment).
            let aligned =
                match layout.Metadata.Alignment with
                | Center -> layout.AvailableWidth / 2.0
                | Right -> layout.AvailableWidth
                | _ -> 0.0
            Some (settings.PaddingLeft + layout.Indent + aligned, layout.Y, (if layout.Height > 0.0 then min layout.Height (max ctx.LineHeight 1.0) else ctx.LineHeight))
        | Some layout ->
            let line = lineForOffset layout offset
            let x = xForOffsetWithRange (measureRangeWithContext ctx) line offset
            Some (settings.PaddingLeft + x + alignmentOffset layout line, line.Y, line.Height)

/// Where the caret (the selection focus) is on the canvas: x, top and height
/// of its line, from the layout that is on screen. None without a selection.
/// In a turned text box the caret is turned too (caretTurn): this is the
/// upright line of its height through its middle.
let caretRect (ctx: RenderContext) (doc: Document) : (float * float * float) option =
    match uprightCaretRect ctx doc with
    | Some (x, top, height) ->
        match ctx.Story, ctx.StoryTurn with
        | Some story, Some (degrees, centerX, centerY) when story.Box.IsSome ->
            let middleX, middleY = turnedPoint degrees centerX centerY x (top + height / 2.0)
            Some (middleX, middleY - height / 2.0, height)
        | _ -> Some (x, top, height)
    | None -> None

/// Hit test for images - returns blockId, offset, and image dimensions if clicking on an image
let hitTestForImage (ctx: RenderContext) (doc: Document) (mouseX: float) (mouseY: float) : (string * int * float * float * float * float) option =
    let found =
        match ctx.Story with
        | Some story ->
            let layouts = storyLayouts ctx story doc
            let mouseX, mouseY = storyPoint ctx mouseX mouseY
            hitTestForImageLayouts (layoutSettings ctx doc).PaddingLeft layouts mouseX mouseY
        | None ->
        Rbt.Word.Layout.hitTestForImageIncremental ctx.LayoutState (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) (layoutTarget ctx doc) mouseX mouseY
    // (Text wraps along the outline of a "tight" picture: a point of its box outside the outline is on that text.)
    match found with
    | Some (blockId, offset, width, height, imageX, imageY) ->
        let onPicture =
            match Rbt.Word.Document.findParagraph doc blockId with
            | Some para ->
                para.Buffer.Runs
                |> List.tryPick (function
                    | TextBuffer.Image image when image.Start = offset -> Some (WordLayout.Floating.pointOnPicture image imageX imageY width height mouseX mouseY)
                    | _ -> None)
                |> Option.defaultValue true
            | None -> true
        if onPicture then found else None
    | None -> None

/// Render document with hover indicators for images
let renderWithHover (ctx: RenderContext) (doc: Document) (hoveredImage: (string * int) option) =
    renderCore ctx doc hoveredImage

#else
// Server-side stub
type RenderContext = unit
type BlockLayout = unit
let createContext _ = ()
let render _ _ = ()
let pageCount _ _ = 1
let pageFrames _ _ = []
let shapeRects _ _ = []
let furnitureShapeRects _ _ = []
let pageShapeRects _ _ _ = []
let shapeRectOf _ _ _ = None
let shapePartAt _ _ _ _ _ = ""
let furnitureShapeAt _ _ _ _ = None
let shapeAt _ _ _ _ = None
let shapeEdgeAt _ _ _ _ = None
let renderWithHover _ _ _ = ()
let setViewport _ _ = ()
let documentSize _ _ = 0.0, 0.0, 0
let renderPage _ _ _ _ _ = false
let viewportToLayout _ x y = x, y
let layoutToViewport _ x y = x, y
let viewportScale _ = 1.0
let hitTest _ _ _ _ = None
let hitTestForImage _ _ _ _ = None
let caretRect _ _ = None
let caretTurn _ = 0.0
#endif
