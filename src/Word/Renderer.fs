module RedBat.Word.Renderer

#if FABLE_COMPILER
open System
open Fable.Core
open Browser.Types
open RedBat.Word.Document
open RedBat.Word.TextBuffer
open RedBat.Word.Layout
open RedBat.Word.Paint

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
    mutable LastDocument: Document option
    mutable LastLayouts: BlockLayout list option
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
    LastDocument = None
    LastLayouts = None
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
    PaddingTop = RedBat.Word.WordLayout.Page.bodyTopWithRunningHeader doc
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
    let layouts = RedBat.Word.Layout.layoutDocumentWithRange (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) doc
    ctx.LastDocument <- Some doc
    ctx.LastLayouts <- Some layouts
    layouts

let cachedLayouts (ctx: RenderContext) (doc: Document) =
    match ctx.LastDocument, ctx.LastLayouts with
    | Some cachedDoc, Some layouts when obj.ReferenceEquals(cachedDoc, doc) -> Some layouts
    | _ -> None

let private editableParagraphs doc =
    doc.Blocks
    |> List.collect (function
        | Paragraph para
        | Heading (_, para) -> [para]
        | Table table -> table.Rows |> List.collect (fun row -> row.Cells |> List.collect (fun cell -> cell.Content))
        | TextBox textBox -> textBox.Content)

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
        editableParagraphs doc
        |> List.collect (fun para ->
            match Map.tryFind para.Id layoutById with
            | None -> []
            | Some layout ->
                para.Buffer.Runs
                |> List.choose (function
                    | Text tr ->
                        tr.Style.CommentId
                        |> Option.bind (fun id ->
                            let line = lineForOffset layout tr.Start
                            let lineEnd = min (tr.Start + tr.Length) line.EndOffset
                            if lineEnd <= tr.Start then None
                            else
                                let lineAlignmentOffset = alignmentOffset layout line
                                let startX = xForOffsetWithRange (measureRangeWithContext ctx) line tr.Start
                                let endX = xForOffsetWithRange (measureRangeWithContext ctx) line lineEnd
                                Some {
                                    Id = id
                                    Text = tr.Style.CommentText |> Option.defaultValue ""
                                    Author = tr.Style.CommentAuthor
                                    Date = tr.Style.CommentDate
                                    Excerpt = TextBuffer.slice tr.Start tr.Length para.Buffer
                                    BlockId = para.Id
                                    StartOffset = tr.Start
                                    EndOffset = tr.Start + tr.Length
                                    X = settings.PaddingLeft + startX + lineAlignmentOffset
                                    Y = line.Y
                                    Width = max 1.0 (endX - startX)
                                    Height = line.Height
                                    CalloutX = calloutX
                                    CalloutY = line.Y
                                    IsSelected = isSelected para.Id tr.Start (tr.Start + tr.Length)
                                })
                    | Image _ -> None))
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
                | ImageSegment _ -> None)))
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
                | ImageSegment _ -> None)))
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

let private renderPages (ctx: RenderContext) (doc: Document) (layouts: BlockLayout list) =
    let settings = layoutSettings ctx doc
    let docHeight = documentHeightWithNotes settings layouts
    let pageHeight = RedBat.Word.Layout.pageHeight settings
    let stride = pageHeight + RedBat.Word.Layout.pageGap
    let pageCount = max 1 (int (ceil (docHeight / stride)))

    ctx.Ctx.fillStyle <- U3.Case1 "#eef0f3"
    ctx.Ctx.fillRect(0.0, 0.0, ctx.Canvas.width, ctx.Canvas.height)

    for pageIndex in 0 .. pageCount - 1 do
        let y = float pageIndex * stride
        ctx.Ctx.fillStyle <- U3.Case1 "#ffffff"
        ctx.Ctx.fillRect(0.0, y, settings.Width, pageHeight)
        ctx.Ctx.strokeStyle <- U3.Case1 "#d0d5dd"
        ctx.Ctx.lineWidth <- 1.0
        ctx.Ctx.strokeRect(0.5, y + 0.5, settings.Width - 1.0, pageHeight - 1.0)

let private pageCountForLayouts (settings: LayoutSettings) (layouts: BlockLayout list) =
    let docHeight = documentHeightWithNotes settings layouts
    let stride = pageHeight settings + pageGap
    max 1 (int (ceil (docHeight / stride)))

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

let private renderPaintCommand (ctx: RenderContext) = function
    | DrawText (text, style, x, baseline, width) ->
        match ctx.TextPathPainter with
        | Some paint when paint text style x baseline width -> ()
        | _ ->
            ctx.Ctx.font <- fontString style ctx.FontFamily
            ctx.Ctx.fillStyle <- U3.Case1 style.Color
            fillTextFitted ctx style text x baseline width
    | DrawUnderline (color, x, y, width) ->
        ctx.Ctx.strokeStyle <- U3.Case1 color
        ctx.Ctx.lineWidth <- 1.0
        ctx.Ctx.beginPath()
        ctx.Ctx.moveTo(x, y)
        ctx.Ctx.lineTo(x + width, y)
        ctx.Ctx.stroke()
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

let private renderCommentAnchors (ctx: RenderContext) (doc: Document) (layouts: BlockLayout list) =
    commentAnchorsForLayouts ctx doc layouts
    |> List.iter (fun anchor ->
        renderPaintCommand ctx (FillRect ("rgba(255, 230, 128, 0.42)", anchor.X, anchor.Y, anchor.Width, anchor.Height))
        renderPaintCommand ctx (StrokeRect ("#C79B00", 1.0, anchor.X + 0.5, anchor.Y + 0.5, max 1.0 (anchor.Width - 1.0), max 1.0 (anchor.Height - 1.0))))

/// Render document to canvas
let render (ctx: RenderContext) (doc: Document) =
    // Clear canvas
    ctx.Ctx.clearRect(0.0, 0.0, ctx.Canvas.width, ctx.Canvas.height)

    // Set font
    ctx.Ctx.font <- sprintf "%fpx %s" ctx.FontSize ctx.FontFamily
    ctx.Ctx.fillStyle <- U3.Case1 "#000000"

    let settings = layoutSettings ctx doc
    let layouts, commands, _ = collectDocumentCommands settings ctx.FontFamily (measureWithContext ctx) (measureRangeWithContext ctx) doc
    renderPages ctx doc layouts
    commands |> List.iter (renderPaintCommand ctx)
    // Footnotes and endnotes are laid out and painted via the shared Layout/Paint
    // pass (so the browser canvas and the PDF/native paths match).
    renderCommentAnchors ctx doc layouts

    // Reset style
    ctx.Ctx.fillStyle <- U3.Case1 "#000000"

    selectionCommandsWithRange settings (measureRangeWithContext ctx) doc layouts
    |> List.iter (renderPaintCommand ctx)

    markupMarkerCommands settings layouts
    |> List.iter (renderPaintCommand ctx)

/// Convert mouse position to document position
let hitTest (ctx: RenderContext) (doc: Document) (mouseX: float) (mouseY: float) =
    RedBat.Word.Layout.hitTestWithRange (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) doc mouseX mouseY

/// Hit test for images - returns blockId, offset, and image dimensions if clicking on an image
let hitTestForImage (ctx: RenderContext) (doc: Document) (mouseX: float) (mouseY: float) : (string * int * float * float * float * float) option =
    RedBat.Word.Layout.hitTestForImage (layoutSettings ctx doc) (measureWithContext ctx) doc mouseX mouseY

/// Render document with hover indicators for images
let renderWithHover (ctx: RenderContext) (doc: Document) (hoveredImage: (string * int) option) =
    // Render the document normally
    render ctx doc

    let layouts = layoutDocument ctx doc
    hoverCommands (layoutSettings ctx doc) layouts hoveredImage
    |> List.iter (renderPaintCommand ctx)

#else
// Server-side stub
type RenderContext = unit
type BlockLayout = unit
let createContext _ = ()
let render _ _ = ()
let renderWithHover _ _ _ = ()
let hitTest _ _ _ _ = None
let hitTestForImage _ _ _ _ = None
#endif
