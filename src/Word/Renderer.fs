module Rbt.Word.Renderer

#if FABLE_COMPILER
open System
open Fable.Core
open Browser.Types
open Rbt.Word.Document
open Rbt.Word.TextBuffer
open Rbt.Word.Layout
open Rbt.Word.Paint

type PaintFrame = {
    Document: Document
    Layouts: BlockLayout list
    /// Geometry of every page frame, as the layout pass produced it (sections
    /// can differ in size and orientation), extended over the note areas.
    Pages: PageInfo list
    BodyCommands: PaintCommand list
    OverlayCommands: PaintCommand list
    CanvasWidth: float
    CanvasHeight: float
    FontFamily: string
    FontSize: float
    LineHeight: float
    PaddingLeft: float
    PaddingTop: float
    TextPathPainter: obj option
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
    ImageCache: obj option  // Map of URL -> HTMLImageElement
    mutable TextMeasurer: TextMeasurer option
    mutable TextRangeMeasurer: TextRangeMeasurer option
    mutable TextPathPainter: (string -> TextPaintStyle -> float -> float -> float -> bool) option
    LayoutState: IncrementalLayoutState
    mutable LastPaint: PaintFrame option
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
    LayoutState = Rbt.Word.Layout.createIncrementalLayoutState ()
    LastPaint = None
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
    Rbt.Word.Layout.layoutDocumentIncrementalWithRange ctx.LayoutState (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) doc

/// Layouts plus the per-page geometry of the same (incremental) layout pass.
let layoutDocumentWithPages (ctx: RenderContext) (doc: Document) =
    Rbt.Word.Layout.layoutDocumentPagesIncremental ctx.LayoutState (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) doc

let cachedLayouts (ctx: RenderContext) (doc: Document) =
    match ctx.LayoutState.LastDocument, ctx.LayoutState.LastLayouts with
    | Some cachedDoc, Some layouts when obj.ReferenceEquals(cachedDoc, doc) -> Some layouts
    | _ -> None

let clearLayoutCache (ctx: RenderContext) =
    Rbt.Word.Layout.clearIncrementalLayoutState ctx.LayoutState
    ctx.LastPaint <- None

let layoutLookupFor (ctx: RenderContext) (doc: Document) =
    let layouts =
        cachedLayouts ctx doc
        |> Option.defaultWith (fun () -> layoutDocument ctx doc)
    fun blockId ->
        layouts |> List.tryFind (fun layout -> layout.BlockId = blockId)

let private commentAnchorsForLayouts (ctx: RenderContext) (doc: Document) (layouts: BlockLayout list) =
    let settings = layoutSettings ctx doc
    let layoutById = layouts |> List.map (fun layout -> layout.BlockId, layout) |> Map.ofList
    let calloutX = max (settings.Width - settings.MarginRight + 12.0) (settings.PaddingLeft + 120.0)

    let isSelected blockId startOffset endOffset =
        match doc.Selection with
        | Some selection when selection.Anchor.BlockId = blockId && selection.Focus.BlockId = blockId ->
            let selectionStart = min selection.Anchor.Offset selection.Focus.Offset
            let selectionEnd = max selection.Anchor.Offset selection.Focus.Offset
            selectionStart = startOffset && selectionEnd = endOffset
        | _ -> false

    let anchors =
        commentInfos doc
        |> List.choose (fun info ->
            match Map.tryFind info.BlockId layoutById with
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
    layouts
    |> List.map (fun layout -> layout.Y + layout.Height)
    |> List.fold max settings.PaddingTop

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

let private currentTextPathPainter (ctx: RenderContext) =
    ctx.TextPathPainter |> Option.map box

let private frameInvalidated (ctx: RenderContext) (previous: PaintFrame) =
    previous.CanvasWidth <> canvasWidth ctx
    || previous.CanvasHeight <> canvasHeight ctx
    || previous.FontFamily <> ctx.FontFamily
    || previous.FontSize <> ctx.FontSize
    || previous.LineHeight <> ctx.LineHeight
    || previous.PaddingLeft <> ctx.PaddingLeft
    || previous.PaddingTop <> ctx.PaddingTop
    || (match previous.TextPathPainter, currentTextPathPainter ctx with
        | None, None -> false
        | Some left, Some right -> not (obj.ReferenceEquals(left, right))
        | _ -> true)

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
    oldLayout.BlockId = newLayout.BlockId
    && obj.ReferenceEquals(oldLayout, newLayout)
    && abs (oldLayout.Y - newLayout.Y) < 0.001
    && abs (oldLayout.Height - newLayout.Height) < 0.001

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

let private commandDiffRange oldCommands newCommands =
    let rec loop oldRemaining newRemaining current =
        match oldRemaining, newRemaining with
        | [], [] -> Choice1Of2 current
        | oldCommand :: oldTail, newCommand :: newTail when oldCommand = newCommand ->
            loop oldTail newTail current
        | oldCommand :: oldTail, newCommand :: newTail ->
            match rangeForCommand oldCommand, rangeForCommand newCommand with
            | Some oldRange, Some newRange ->
                loop oldTail newTail (current |> unionRange (Some oldRange) |> unionRange (Some newRange))
            | _ -> Choice2Of2 ()
        | oldCommand :: oldTail, [] ->
            match rangeForCommand oldCommand with
            | Some oldRange -> loop oldTail [] (unionRange (Some oldRange) current)
            | None -> Choice2Of2 ()
        | [], newCommand :: newTail ->
            match rangeForCommand newCommand with
            | Some newRange -> loop [] newTail (unionRange (Some newRange) current)
            | None -> Choice2Of2 ()
    loop oldCommands newCommands None

let private commandListRange commands =
    let rec loop remaining current =
        match remaining with
        | [] -> Choice1Of2 current
        | command :: rest ->
            match rangeForCommand command with
            | Some range -> loop rest (unionRange (Some range) current)
            | None -> Choice2Of2 ()
    loop commands None

let private imageCommandRange commands =
    let rec loop remaining current =
        match remaining with
        | [] -> Choice1Of2 current
        | (DrawImage _ as command) :: rest ->
            match rangeForCommand command with
            | Some range -> loop rest (unionRange (Some range) current)
            | None -> Choice2Of2 ()
        | _ :: rest -> loop rest current
    loop commands None

let private computeDirtyRange (previous: PaintFrame) (nextFrame: PaintFrame) =
    match commandDiffRange previous.BodyCommands nextFrame.BodyCommands with
    | Choice2Of2 () -> FullRepaint
    | Choice1Of2 bodyRange ->
        match commandListRange previous.OverlayCommands, commandListRange nextFrame.OverlayCommands with
        | Choice2Of2 (), _
        | _, Choice2Of2 () -> FullRepaint
        | Choice1Of2 oldOverlayRange, Choice1Of2 newOverlayRange ->
            match imageCommandRange previous.BodyCommands, imageCommandRange nextFrame.BodyCommands with
            | Choice2Of2 (), _
            | _, Choice2Of2 () -> FullRepaint
            | Choice1Of2 oldImageRange, Choice1Of2 newImageRange ->
                let canvasBottom = nextFrame.CanvasHeight
                let previousLayouts: BlockLayout list = previous.Layouts
                let nextLayouts: BlockLayout list = nextFrame.Layouts
                let range =
                    firstLayoutRange previousLayouts nextLayouts canvasBottom
                    |> unionRange bodyRange
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

let private renderCommandListInBand ctx band commands =
    commands
    |> List.iter (fun command ->
        if commandIntersectsBand band command then renderPaintCommand ctx command)

let private renderFullFrame ctx doc frame =
    ctx.Ctx.clearRect(0.0, 0.0, frame.CanvasWidth, frame.CanvasHeight)

    ctx.Ctx.font <- sprintf "%fpx %s" ctx.FontSize ctx.FontFamily
    ctx.Ctx.fillStyle <- U3.Case1 "#000000"

    renderPages ctx frame.Pages
    renderCommandList ctx frame.BodyCommands
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
    renderCommandListInBand ctx band frame.BodyCommands
    renderCommandListInBand ctx band frame.OverlayCommands

    ctx.Ctx.restore()
    ctx.Ctx.fillStyle <- U3.Case1 "#000000"

let private buildFrame (ctx: RenderContext) (doc: Document) hoveredImage =
    let settings = layoutSettings ctx doc
    let layouts, pages = layoutDocumentWithPages ctx doc
    let layouts, bodyCommands, _ = collectDocumentCommandsWithLayoutsAndPages settings ctx.FontFamily (measureWithContext ctx) (measureRangeWithContext ctx) doc layouts pages
    let overlayCommands =
        List.concat [
            commentAnchorCommands ctx doc layouts
            selectionCommandsWithRange settings (measureRangeWithContext ctx) doc layouts
            markupMarkerCommands settings layouts
            hoverCommands settings layouts hoveredImage
        ]
    {
        Document = doc
        Layouts = layouts
        Pages = pageFramesForLayouts settings doc layouts pages
        BodyCommands = bodyCommands
        OverlayCommands = overlayCommands
        CanvasWidth = canvasWidth ctx
        CanvasHeight = canvasHeight ctx
        FontFamily = ctx.FontFamily
        FontSize = ctx.FontSize
        LineHeight = ctx.LineHeight
        PaddingLeft = ctx.PaddingLeft
        PaddingTop = ctx.PaddingTop
        TextPathPainter = currentTextPathPainter ctx
    }

let private renderCore (ctx: RenderContext) (doc: Document) hoveredImage =
    let frame = buildFrame ctx doc hoveredImage
    match ctx.LastPaint with
    | None ->
        renderFullFrame ctx doc frame
    | Some previous when frameInvalidated ctx previous || previous.Pages <> frame.Pages ->
        renderFullFrame ctx doc frame
    | Some previous ->
        match computeDirtyRange previous frame with
        | FullRepaint -> renderFullFrame ctx doc frame
        | NoRepaint -> ()
        | PartialRepaint band -> renderPartialFrame ctx doc frame band
    ctx.LastPaint <- Some frame

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
        | Some layout when layout.Lines.IsEmpty -> Some (settings.PaddingLeft, layout.Y, ctx.LineHeight)
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
let hitTest _ _ _ _ = None
let hitTestForImage _ _ _ _ = None
let caretRect _ _ = None
#endif
