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
    /// Where the chunks of the pages' headers and footers start in Body (its
    /// length when there are none): while a header is edited in place the
    /// body is dimmed below them.
    FurnitureStart: int
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
    LayoutState: IncrementalLayoutState
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
    /// The page frames of a layout pass (by its serial): a host asks for them with every change.
    mutable KnownPageFrames: (int * PageInfo list) option
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
    FrameBuilds = 0
    LastLaidOutBlocks = 0
    LastCollectedBlocks = 0
    CollectedOnPaint = 0
    PlanScope = obj ()
    LayoutState = Rbt.Word.Layout.createIncrementalLayoutState ()
    PaintCache = createPaintCache ()
    CommentBlocks = [||]
    CommentFlags = [||]
    LastPaint = None
    Viewport = None
    PageFrame = None
    Progress = { Scheduled = false; Epoch = 0; Chunk = 48; Notified = 0.0; Waiters = ResizeArray(); LoadPending = false
                 Slices = 0; LongestSliceMs = 0.0; LongestSlice = 0; LongestSlicePasses = 0; LongestSliceLayoutMs = 0.0; LongestSliceFrom = 0; SlicePasses = 0 }
    Story = None
    StoryPainted = false
    StoryOverlay = []
    StoryLayouts = None
    KnownPageFrames = None
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

/// The layout settings of a document for a base font size and line height.
let layoutSettingsFor (fontSize: float) (lineHeight: float) (doc: Document) =
    let hasFooter = not doc.Footers.IsEmpty || not doc.FooterParts.IsEmpty
    let footerReserve =
        if hasFooter then doc.PageSetup.FooterDistance + lineHeight * 1.5
        else doc.PageSetup.MarginBottom
    {
    Width = doc.PageSetup.PageWidth
    PageHeight = doc.PageSetup.PageHeight
    FontSize = fontSize
    LineHeight = lineHeight
    PaddingLeft = doc.PageSetup.MarginLeft
    PaddingTop = Rbt.Word.WordLayout.Page.bodyTopWithRunningHeader doc
    MarginRight = doc.PageSetup.MarginRight
    MarginBottom = max doc.PageSetup.MarginBottom footerReserve
    InTableCell = false
    TableContinuationHeaderHeight = 0.0
    LongTrailingSpaceFitSafety = 3.5
    }

let layoutSettings (ctx: RenderContext) (doc: Document) =
    layoutSettingsFor ctx.FontSize ctx.LineHeight doc

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

/// Layouts plus the per-page geometry of the same (incremental) layout pass:
/// of the blocks the layout state's limit covers.
let layoutDocumentWithPages (ctx: RenderContext) (doc: Document) =
    let result =
        Rbt.Word.Layout.withTablePlanScope ctx.PlanScope (fun () ->
            Rbt.Word.Layout.layoutDocumentPagesIncremental ctx.LayoutState (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) doc)
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

/// The layouts of a story's blocks where its page paints them.
let private storyLayouts (ctx: RenderContext) (story: StoryView) (doc: Document) : BlockLayout list =
    let whole = story.Whole doc
    let settings = layoutSettings ctx whole
    let bodyLayouts, layoutPages = layoutDocumentWithPages ctx whole
    let pages = framePages.Value ctx whole bodyLayouts layoutPages
    match List.tryItem story.PageIndex pages with
    | None -> []
    | Some page ->
        let startY = headerFooterStartY settings whole.PageSetup page story.IsHeader
        match ctx.StoryLayouts with
        | Some (blocks, y, version, layouts) when obj.ReferenceEquals(blocks, doc.Blocks) && y = startY && version = ctx.TextMeasurerVersion -> layouts
        | _ ->
            let layouts =
                Rbt.Word.Layout.withTablePlanScope ctx.PlanScope (fun () ->
                    storyBlockLayouts settings (measureWithContext ctx) (measureRangeWithContext ctx) whole startY doc.Blocks)
            ctx.StoryLayouts <- Some (doc.Blocks, startY, ctx.TextMeasurerVersion, layouts)
            layouts

/// The layouts the caret and the pointer go by: the body's, or - while a
/// header or footer is edited in place - those of its blocks on its page.
let layoutDocument (ctx: RenderContext) (doc: Document) =
    match ctx.Story with
    | Some story -> storyLayouts ctx story doc
    | None -> layoutDocumentWithPages ctx doc |> fst

/// From here on the whole document is laid out (by the next layout asked for).
let completeLayout (ctx: RenderContext) =
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
    let mutable more = not (layoutIsWhole ctx)
    while more do
        let layouts, _ = layoutDocumentWithPages ctx doc
        if layoutIsWhole ctx || laidOutBottom ctx layouts >= bottom then more <- false
        else
            // (In steps of a third: a pass over a few more blocks costs little, the blocks themselves cost their measuring.)
            let limit = ctx.LayoutState.BlockLimit
            // (A few at first: the blocks of the first screens may be pages long each.)
            ctx.LayoutState.BlockLimit <- (if limit > 1000000 then System.Int32.MaxValue else limit + max 4 (limit / 3))

/// What a viewport shows, and a screen below it, is laid out.
let private coverViewport (ctx: RenderContext) (doc: Document) (vp: Viewport) =
    if not (layoutIsWhole ctx) then
        let top = (vp.ScrollTop - vp.OriginY) / vp.Scale
        let height = vp.Height / vp.Scale
        coverDown ctx doc (top + 2.0 * height)

/// The blocks the selection is in are laid out, and some after them (the
/// caret moves down by lines that must be there).
let private coverSelection (ctx: RenderContext) (doc: Document) =
    if not (layoutIsWhole ctx) then
        match doc.Selection with
        | None -> ()
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
        let limit = ctx.LayoutState.BlockLimit
        // (Never a sliver of what is laid out already: a pass has a cost of its own that grows with the document.)
        let chunk = max progress.Chunk (limit / 512)
        ctx.LayoutState.BlockLimit <- limit + chunk
        layoutDocumentWithPages ctx doc |> ignore
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
                    let from = ctx.LayoutState.BlockLimit
                    progress.SlicePasses <- 0
                    let complete = advanceLayout ctx (currentDocument ()) budget
                    let now = nowMs ()
                    let tell = complete || now - progress.Notified > 150.0
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
    ctx.PlanScope <- obj ()
    Rbt.Word.Layout.clearIncrementalLayoutState ctx.LayoutState
    clearPaintCache ctx.PaintCache
    ctx.LastPaint <- None
    ctx.PageFrame <- None

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
    clearLayoutCache ctx

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
    let count = if names.Count = 0 then 0 else Rbt.Word.Layout.invalidateFacesInIncrementalLayout ctx.LayoutState uses
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

let commentAnchors (ctx: RenderContext) (doc: Document) =
    let doc = bodyDocument ctx doc
    commentAnchorsForLayouts ctx doc (layoutWholeDocument ctx doc)

let private pageIndexForY settings y =
    let stride = pageHeight settings + pageGap
    int (floor (max 0.0 y / stride))

let private endnoteAnchorsForLayouts (layouts: BlockLayout list) =
    layouts
    |> List.collect (fun layout ->
        layout.Lines
        |> List.collect (fun line ->
            line.Segments
            |> List.choose (function
                | TextSegment (_, style, _, _, _) ->
                    match style.NoteId, style.NoteType, style.NoteText with
                    | Some id, Some "endnote", Some text when text <> "" -> Some { Id = id; Text = text }
                    | _ -> None
                | ImageSegment _ -> None
                | MathSegment _ -> None)))
    |> List.distinctBy _.Id

let private footnoteAnchorsForLayouts (settings: LayoutSettings) (layouts: BlockLayout list) =
    layouts
    |> List.collect (fun layout ->
        layout.Lines
        |> List.collect (fun line ->
            line.Segments
            |> List.choose (function
                | TextSegment (_, style, _, _, _) ->
                    match style.NoteId, style.NoteType, style.NoteText with
                    | Some id, noteType, Some text when noteType |> Option.defaultValue "footnote" = "footnote" && text <> "" ->
                        Some { Id = id; Text = text; PageIndex = pageIndexForY settings line.Y }
                    | _ -> None
                | ImageSegment _ -> None
                | MathSegment _ -> None)))
    |> List.distinctBy (fun anchor -> anchor.PageIndex, anchor.Id)

let private bodyBottom (settings: LayoutSettings) (layouts: BlockLayout list) =
    let mutable bottom = settings.PaddingTop
    for layout in layouts do
        bottom <- max bottom (layout.Y + layout.Height)
    bottom

let private endnoteStartY (settings: LayoutSettings) (layouts: BlockLayout list) =
    let bottom = bodyBottom settings layouts
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) bottom
    let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
    if bottom + 48.0 > contentBottom then WordLayout.Page.advanceToNextPage (pageHeight settings) settings.PaddingTop bottom else bottom + 24.0

let private documentHeightWithNotes (settings: LayoutSettings) (layouts: BlockLayout list) =
    let bottom = bodyBottom settings layouts
    let endnotes = endnoteAnchorsForLayouts layouts
    let endnoteBottom =
        if endnotes.IsEmpty then bottom
        else
            let lineCount = endnotes |> List.sumBy (fun anchor -> noteLineCount settings (sprintf "%s. %s" anchor.Id anchor.Text))
            endnoteStartY settings layouts + 18.0 + float lineCount * 13.0
    let footnoteBottom =
        footnoteAnchorsForLayouts settings layouts
        |> List.groupBy _.PageIndex
        |> List.map (fun (pageIndex, anchors) ->
            let pageTop = float pageIndex * (pageHeight settings + pageGap)
            let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
            let lineCount = anchors |> List.sumBy (fun anchor -> noteLineCount settings (sprintf "%s. %s" anchor.Id anchor.Text))
            contentBottom + 20.0 + float lineCount * 13.0)
        |> List.fold max bottom
    max endnoteBottom footnoteBottom

/// The page frames to show: the pages of the layout pass, each with its own
/// section's size, plus pages for note areas that reach below the last one.
let private pageFramesOfLayouts (settings: LayoutSettings) (doc: Document) (layouts: BlockLayout list) (pages: PageInfo list) =
    let bottom =
        if documentHasSections doc then bodyBottom settings layouts
        else documentHeightWithNotes settings layouts
    pageInfosCoveringBottom bottom pages

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
        let frames = pageFramesOfLayouts settings doc layouts pages
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

let private renderPaintCommand (ctx: RenderContext) = function
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
    | Some _ ->
        // The body, dimmed; over it the headers and footers and the story's boundary, selection and caret.
        let furniture = min frame.FurnitureStart frame.Body.Length
        renderChunkRangeInBand ctx frame.Body 0 furniture System.Double.NegativeInfinity System.Double.PositiveInfinity
        renderCommandList ctx frame.StaticOverlay
        for page in frame.Pages do
            ctx.Ctx.fillStyle <- U3.Case1 storyDimColor
            ctx.Ctx.fillRect(0.0, page.Top, page.Width, page.Height)
        renderChunkRangeInBand ctx frame.Body furniture frame.Body.Length System.Double.NegativeInfinity System.Double.PositiveInfinity
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

/// The overlay that follows the selection and the pointer, and the band of
/// the document it was made for (None: all of it). A block's layout is found
/// by the block's place when the layout pass handed its layouts over block
/// by block; a selection over several blocks is highlighted in the band only.
let private selectionOverlay (ctx: RenderContext) (settings: LayoutSettings) (doc: Document) (layouts: BlockLayout list) (tryFindLayout: (string -> BlockLayout option) option) hoveredImage : PaintCommand list * (float * float) option =
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
    List.concat [ selection; markupMarkerCommands settings doc layouts; hover ], band

let private buildFrame (ctx: RenderContext) (doc: Document) hoveredImage =
    ctx.FrameBuilds <- ctx.FrameBuilds + 1
    let settings = layoutSettings ctx doc
    let layouts, layoutPages = layoutDocumentWithPages ctx doc
    // (The pages of the notes after the body are pages too: they have the headers and footers,
    // and the page count the fields show is theirs.)
    let pages = pageFramesForLayouts ctx settings doc layouts layoutPages
    let body =
        Rbt.Word.Layout.withTablePlanScope ctx.PlanScope (fun () ->
            let chunked =
                match ctx.LayoutState.LastSlices with
                | Some slices -> collectDocumentChunks ctx.PaintCache settings ctx.FontFamily (measureWithContext ctx) (measureRangeWithContext ctx) doc slices pages
                | None -> None
            ctx.LastLaidOutBlocks <- Rbt.Word.Layout.incrementalLayoutBlocksLaidOut ctx.LayoutState
            ctx.LastCollectedBlocks <- (match chunked with Some _ -> paintCacheCollectedBlocks ctx.PaintCache | None -> -1)
            match chunked with
            | Some chunks -> chunks
            | None ->
                // The whole document's commands at once, as one chunk.
                let _, commands, _ = collectDocumentCommandsWithLayoutsAndPages settings ctx.FontFamily (measureWithContext ctx) (measureRangeWithContext ctx) doc layouts pages
                [| chunkOfCommands (List.toArray commands) |])
    let overlay, overlayBand = selectionOverlay ctx settings doc layouts None hoveredImage
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
        FurnitureStart =
            // (The chunks of the pages' headers and footers are the last ones, one per page of the layout.)
            if ctx.LayoutState.LastSlices.IsSome && body.Length > 1
               && not (doc.Headers.IsEmpty && doc.Footers.IsEmpty && doc.HeaderParts.IsEmpty && doc.FooterParts.IsEmpty) then
                max 0 (body.Length - List.length pages)
            else body.Length
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
    // While a long document is being laid out from its start: what is in view
    // and what the selection is in are laid out before they are shown.
    if not (ctx.LayoutState.LastWasWhole && layoutIsWhole ctx) then
        coverSelection ctx doc
        match ctx.Viewport with
        | Some vp -> coverViewport ctx doc vp
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
            let overlay, overlayBand = selectionOverlay ctx (layoutSettings ctx doc) doc previous.Layouts (Some (fun id -> frameLayoutLookup previous id)) hoveredImage
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
            | Some _ ->
                // The body, dimmed; over it the headers and footers and the story's boundary, selection and caret.
                let furniture = min frame.FurnitureStart frame.Body.Length
                renderChunkRangeInBand ctx frame.Body 0 furniture pageTop pageBottom
                renderChunksInBand ctx [| staticOverlayIndex frame |] pageTop pageBottom
                ctx.Ctx.fillStyle <- U3.Case1 storyDimColor
                ctx.Ctx.fillRect(0.0, pageTop, page.Width, pageBottom - pageTop)
                renderChunkRangeInBand ctx frame.Body furniture frame.Body.Length pageTop pageBottom
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
                    let overlay, overlayBand = selectionOverlay ctx (layoutSettings ctx doc) doc frame.Layouts None hoveredImage
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
            let bodyLayouts, layoutPages = layoutDocumentWithPages ctx whole
            match List.tryItem story.PageIndex (pageFramesForLayouts ctx settings whole bodyLayouts layoutPages) with
            | Some page -> storyBoundaryCommands settings ctx.FontFamily page story.IsHeader story.Label
            | None -> []
        ctx.StoryOverlay <-
            List.concat [ boundary
                          selectionCommandsWithRangeLookup settings (measureRangeWithContext ctx) doc tryFind
                          hoverCommandsLookup settings tryFind hoveredImage ]
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
    requiredFontFaces (layoutSettingsFor 16.0 24.0 doc) "Calibri" measureText (rangeMeasurerFromTextMeasurer measureText) doc None
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
            let y = max (top + 0.5) (min (bottom - 0.5) mouseY)
            hitTestLayoutArrayWithRange (measureRangeWithContext ctx) (layoutSettings ctx doc).PaddingLeft layouts mouseX y
    | None ->
    // Same layout the last render produced (reused while the document is
    // unchanged) rather than a fresh full layout per mouse event.
    Rbt.Word.Layout.hitTestIncrementalWithRange ctx.LayoutState (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) doc mouseX mouseY

/// Where the caret (the selection focus) is on the canvas: x, top and height
/// of its line, from the layout that is on screen. None without a selection.
let caretRect (ctx: RenderContext) (doc: Document) : (float * float * float) option =
    match doc.Selection with
    | None -> None
    | Some selection ->
        let settings = layoutSettings ctx doc
        let offset = selection.Focus.Offset
        if ctx.Story.IsNone then coverSelection ctx doc
        // A paragraph continued in another column has one layout per fragment.
        let mutable found: BlockLayout option = None
        let mutable fallback: BlockLayout option = None
        for layout in layoutDocument ctx doc do
            if layout.BlockId = selection.Focus.BlockId then
                fallback <- Some layout
                if found.IsNone && layout.Lines |> List.exists (fun line -> line.StartOffset <= offset && offset <= line.EndOffset) then
                    found <- Some layout
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

/// Hit test for images - returns blockId, offset, and image dimensions if clicking on an image
let hitTestForImage (ctx: RenderContext) (doc: Document) (mouseX: float) (mouseY: float) : (string * int * float * float * float * float) option =
    let found =
        match ctx.Story with
        | Some story -> hitTestForImageLayouts (layoutSettings ctx doc).PaddingLeft (storyLayouts ctx story doc) mouseX mouseY
        | None ->
        Rbt.Word.Layout.hitTestForImageIncremental ctx.LayoutState (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) doc mouseX mouseY
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
#endif
