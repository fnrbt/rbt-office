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
    PlanScope = obj ()
    LayoutState = Rbt.Word.Layout.createIncrementalLayoutState ()
    PaintCache = createPaintCache ()
    CommentBlocks = [||]
    CommentFlags = [||]
    LastPaint = None
    Viewport = None
    PageFrame = None
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

/// Measure text width with style
let measureTextStyled (ctx: RenderContext) (text: string) (style: TextBuffer.TextStyle) =
    let fontStyle = if style.Bold && style.Italic then "bold italic"
                    elif style.Bold then "bold"
                    elif style.Italic then "italic"
                    else ""
    let fontFamily = style.FontFamily |> Option.defaultValue ctx.FontFamily
    let fontSize = if style.VertAlign.IsSome then style.FontSize * 0.7 else style.FontSize
    ctx.Ctx.font <- sprintf "%s %fpx %s" fontStyle fontSize (cssFontFamily fontFamily)
    (ctx.Ctx.measureText(text).width * browserMetricScale fontFamily style) + TextBuffer.characterSpacingAdvance text style

let layoutSettings (ctx: RenderContext) (doc: Document) =
    let hasFooter = not doc.Footers.IsEmpty || not doc.FooterParts.IsEmpty
    let footerReserve =
        if hasFooter then doc.PageSetup.FooterDistance + ctx.LineHeight * 1.5
        else doc.PageSetup.MarginBottom
    {
    Width = doc.PageSetup.PageWidth
    PageHeight = doc.PageSetup.PageHeight
    FontSize = ctx.FontSize
    LineHeight = ctx.LineHeight
    PaddingLeft = doc.PageSetup.MarginLeft
    PaddingTop = Rbt.Word.WordLayout.Page.bodyTopWithRunningHeader doc
    MarginRight = doc.PageSetup.MarginRight
    MarginBottom = max doc.PageSetup.MarginBottom footerReserve
    InTableCell = false
    TableContinuationHeaderHeight = 0.0
    LongTrailingSpaceFitSafety = 3.5
    }

let measureWithContext ctx text style =
    match ctx.TextMeasurer with
    | Some measure -> measure text style
    | None -> measureTextStyled ctx text style

let measureRangeWithContext ctx text style start length =
    match ctx.TextRangeMeasurer with
    | Some measureRange -> measureRange text style start length
    | None -> rangeMeasurerFromTextMeasurer (measureWithContext ctx) text style start length

let layoutDocument (ctx: RenderContext) (doc: Document) =
    Rbt.Word.Layout.withTablePlanScope ctx.PlanScope (fun () ->
        Rbt.Word.Layout.layoutDocumentIncrementalWithRange ctx.LayoutState (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) doc)

/// Layouts plus the per-page geometry of the same (incremental) layout pass.
let layoutDocumentWithPages (ctx: RenderContext) (doc: Document) =
    Rbt.Word.Layout.withTablePlanScope ctx.PlanScope (fun () ->
        Rbt.Word.Layout.layoutDocumentPagesIncremental ctx.LayoutState (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) doc)

let cachedLayouts (ctx: RenderContext) (doc: Document) =
    match ctx.LayoutState.LastDocument, ctx.LayoutState.LastLayouts with
    | Some cachedDoc, Some layouts when obj.ReferenceEquals(cachedDoc, doc) -> Some layouts
    | _ -> None

let clearLayoutCache (ctx: RenderContext) =
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

let layoutLookupFor (ctx: RenderContext) (doc: Document) =
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
    commentAnchorsForLayouts ctx doc (layoutDocument ctx doc)

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
let private pageFramesForLayouts (settings: LayoutSettings) (doc: Document) (layouts: BlockLayout list) (pages: PageInfo list) =
    let bottom =
        if documentHasSections doc then bodyBottom settings layouts
        else documentHeightWithNotes settings layouts
    pageInfosCoveringBottom bottom pages

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
        for command in chunk.Commands do
            renderPaintCommand ctx command

/// The commands of the chunks that reach into a band (top .. bottom).
let private renderChunksInBand ctx (chunks: PaintChunk array) (top: float) (bottom: float) =
    for chunk in chunks do
        if chunk.Bottom >= top && chunk.Top <= bottom then
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

/// The overlay that follows the selection and the pointer.
let private selectionOverlay (ctx: RenderContext) (settings: LayoutSettings) (doc: Document) (layouts: BlockLayout list) (tryFindLayout: (string -> BlockLayout option) option) hoveredImage =
    match tryFindLayout with
    | Some tryFind ->
        List.concat [
            selectionCommandsWithRangeLookup settings (measureRangeWithContext ctx) doc tryFind
            markupMarkerCommands settings doc layouts
            hoverCommandsLookup settings tryFind hoveredImage
        ]
    | None ->
        List.concat [
            selectionCommandsWithRange settings (measureRangeWithContext ctx) doc layouts
            markupMarkerCommands settings doc layouts
            hoverCommands settings layouts hoveredImage
        ]

let private buildFrame (ctx: RenderContext) (doc: Document) hoveredImage =
    ctx.FrameBuilds <- ctx.FrameBuilds + 1
    let settings = layoutSettings ctx doc
    let layouts, pages = layoutDocumentWithPages ctx doc
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
    {
        Document = doc
        Layouts = layouts
        Pages = pageFramesForLayouts settings doc layouts pages
        Body = body
        StaticOverlay = commentAnchorCommands ctx doc layouts
        OverlayCommands = selectionOverlay ctx settings doc layouts None hoveredImage
        CanvasWidth = canvasWidth ctx
        CanvasHeight = canvasHeight ctx
        FontFamily = ctx.FontFamily
        FontSize = ctx.FontSize
        LineHeight = ctx.LineHeight
        PaddingLeft = ctx.PaddingLeft
        PaddingTop = ctx.PaddingTop
        PainterVersion = ctx.TextPathPainterVersion
        MeasurerVersion = ctx.TextMeasurerVersion
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
    match ctx.LastPaint with
    | Some previous when not (frameInvalidated ctx { previous with CanvasWidth = canvasWidth ctx; CanvasHeight = canvasHeight ctx })
                         && Rbt.Word.Layout.sameDocumentApartFromSelection previous.Document doc ->
        if obj.ReferenceEquals(previous.Document, doc) && previous.Hovered = hoveredImage then previous
        else
            // The layout state follows the document on display (hit tests, the caret rectangle).
            layoutDocumentWithPages ctx doc |> ignore
            let overlay = selectionOverlay ctx (layoutSettings ctx doc) doc previous.Layouts (Some (frameLayoutLookup previous)) hoveredImage
            { previous with Document = doc; Hovered = hoveredImage; OverlayCommands = overlay }
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
        if frame.Viewport = Some vp && not resized then frame
        else { frame with Viewport = Some vp; CanvasWidth = canvasWidth ctx; CanvasHeight = canvasHeight ctx }
    let visible = visibleBand vp
    // One layout pixel of slack either side: rounding the offset to device pixels can show a sliver more.
    let visible = { Top = visible.Top - 1.0; Bottom = visible.Bottom + 1.0 }
    match ctx.LastPaint with
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

let private renderCore (ctx: RenderContext) (doc: Document) hoveredImage =
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
    | Some previous when previous.Viewport.IsSome || frameInvalidated ctx previous || previous.Pages <> frame.Pages ->
        renderFullFrame ctx doc frame
    | Some previous ->
        match computeDirtyRange previous frame frame.CanvasHeight with
        | FullRepaint -> renderFullFrame ctx doc frame
        | NoRepaint -> ()
        | PartialRepaint band -> renderPartialFrame ctx doc frame band
    ctx.LastPaint <- Some frame

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
    let frames =
        match ctx.LastPaint with
        | Some frame when obj.ReferenceEquals(frame.Document, doc) -> frame.Pages
        | _ ->
            let layouts, pages = layoutDocumentWithPages ctx doc
            pageFramesForLayouts (layoutSettings ctx doc) doc layouts pages
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
    let usable (frame: PaintFrame) =
        obj.ReferenceEquals(frame.Document, doc)
        && not (frameInvalidated ctx { frame with CanvasWidth = canvasWidth ctx; CanvasHeight = canvasHeight ctx })
    let frame =
        match ctx.LastPaint, ctx.PageFrame with
        | Some frame, _ when usable frame -> frame
        | _, Some frame when usable frame -> frame
        | _ ->
            let frame = buildFrame ctx doc None
            ctx.PageFrame <- Some frame
            frame
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
        true

/// Render document to canvas
let render (ctx: RenderContext) (doc: Document) =
    renderCore ctx doc None

/// Number of pages the document lays out to, including footnote/endnote areas.
/// Hosts use it to size the canvas (pages are stacked at pageHeight + pageGap).
let pageFrames (ctx: RenderContext) (doc: Document) =
    let layouts, pages = layoutDocumentWithPages ctx doc
    pageFramesForLayouts (layoutSettings ctx doc) doc layouts pages

/// Number of pages the document lays out to, including footnote/endnote areas.
/// Pages are stacked with pageGap between them, but need not be the same size
/// (sections); hosts size the canvas from pageFrames.
let pageCount (ctx: RenderContext) (doc: Document) =
    List.length (pageFrames ctx doc)

/// Convert mouse position to document position
let hitTest (ctx: RenderContext) (doc: Document) (mouseX: float) (mouseY: float) =
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
    Rbt.Word.Layout.hitTestForImageIncremental ctx.LayoutState (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) doc mouseX mouseY

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
