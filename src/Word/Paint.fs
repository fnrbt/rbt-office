module Rbt.Word.Paint

open System.Collections.Generic
open Rbt.Word.Layout
open Rbt.Word.Document
open Rbt.Word.TextBuffer
open Rbt.Word.WordStyles

module Eq = Rbt.Word.Math

let private equalsIgnoreCase (left: string) (right: string) =
    System.String.Equals(left, right, System.StringComparison.OrdinalIgnoreCase)

let private optionEqualsIgnoreCase expected value =
    match value with
    | Some actual -> equalsIgnoreCase actual expected
    | None -> false

let private trimmedEqualsIgnoreCase (value: string) (expected: string) =
    if isNull value then false
    else
        let mutable start = 0
        let mutable finish = value.Length - 1
        while start <= finish && System.Char.IsWhiteSpace value.[start] do
            start <- start + 1
        while finish >= start && System.Char.IsWhiteSpace value.[finish] do
            finish <- finish - 1
        let length = finish - start + 1
        if length <> expected.Length then false
        else
            let mutable index = 0
            let mutable equal = true
            while equal && index < length do
                let left = value.[start + index]
                let right = expected.[index]
                if left <> right && System.Char.ToUpperInvariant left <> System.Char.ToUpperInvariant right then
                    equal <- false
                index <- index + 1
            equal

let private optionTrimmedEqualsIgnoreCase expected value =
    match value with
    | Some actual -> trimmedEqualsIgnoreCase actual expected
    | None -> false

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private columnWidthPrefix (columnWidths: float list) =
    let widths = columnWidths |> List.toArray
    let prefix = Array.zeroCreate<float> (widths.Length + 1)
    for i = 0 to widths.Length - 1 do
        prefix.[i + 1] <- prefix.[i] + widths.[i]
    prefix

let private startOffsetFromPrefix (prefix: float[]) columnIndex =
    if columnIndex <= 0 then 0.0
    else prefix.[min columnIndex (prefix.Length - 1)]

let private spannedWidthFromPrefix (prefix: float[]) start span fallback =
    let widthCount = prefix.Length - 1
    let mutable total = 0.0
    let mutable remaining = span
    let mutable index = start
    while remaining > 0 && index < 0 do
        total <- total + fallback
        index <- index + 1
        remaining <- remaining - 1
    if remaining > 0 && index < widthCount then
        let finish = min widthCount (index + remaining)
        total <- total + prefix.[finish] - prefix.[index]
        remaining <- remaining - (finish - index)
        index <- finish
    if remaining > 0 then
        total <- total + fallback * float remaining
    total

let private spacedStartOffsetFromPrefix prefix spacing columnIndex =
    startOffsetFromPrefix prefix columnIndex + spacing

let private spacedSpannedWidthFromPrefix prefix spacing start span fallback =
    max 0.0 (spannedWidthFromPrefix prefix start span fallback - spacing * 2.0)

let private mapListWithChange (mapper: 'T -> 'T * bool) (items: 'T list) =
    let mutable mapped = Unchecked.defaultof<ResizeArray<'T>>
    let mutable changed = false
    let mutable index = 0
    let mutable remaining = items
    while not remaining.IsEmpty do
        match remaining with
        | item :: tail ->
            let next, itemChanged = mapper item
            if changed then
                mapped.Add next
            elif itemChanged then
                changed <- true
                mapped <- ResizeArray<'T>()
                let mutable prefix = items
                let mutable prefixIndex = 0
                while prefixIndex < index do
                    match prefix with
                    | original :: prefixTail ->
                        mapped.Add original
                        prefix <- prefixTail
                        prefixIndex <- prefixIndex + 1
                    | [] -> prefixIndex <- index
                mapped.Add next
            remaining <- tail
            index <- index + 1
        | [] -> ()
    if changed then resizeArrayToList mapped, true else items, false

let private lastTableRowFragment (rows: TableLayoutPlanRow list) =
    let mutable result = None
    for row in rows do
        for fragment in row.Fragments do
            result <- Some fragment
    result

[<Struct>]
type TextPaintStyle = {
    Bold: bool
    Italic: bool
    Underline: bool
    AllCaps: bool
    SmallCaps: bool
    Strike: bool
    FontSize: float
    CharacterSpacing: float
    FontFamily: string option
    Color: string
    Highlight: string option
    VertAlign: string option
}

/// Fill of a vector shape: solid colour, gradient, or none.
type ShapeFill =
    | FillNone
    | FillSolid of color: string
    | FillGradient of GradientSpec

type PaintCommand =
    | DrawText of text: string * style: TextPaintStyle * x: float * baseline: float * width: float
    | DrawUnderline of color: string * lineStyle: string * x: float * y: float * width: float
    | DrawImage of src: string * x: float * y: float * width: float * height: float * cropLeft: string option * cropTop: string option * cropRight: string option * cropBottom: string option
    | FillRect of color: string * x: float * y: float * width: float * height: float
    | StrokeRect of color: string * lineWidth: float * x: float * y: float * width: float * height: float
    | StrokeLine of color: string * lineWidth: float * dashStyle: int * dashOffset: float * x1: float * y1: float * x2: float * y2: float
    /// A vector shape (preset/custom geometry) with absolute page-coordinate path
    /// segments, an optional stroke, and a gradient bounding box (for gradient fills).
    | DrawShape of fill: ShapeFill * stroke: string option * strokeWidth: float * bbox: (float * float * float * float) * segs: PathSeg list
    /// Commands drawn together (a text box, a shape with its text, WordArt):
    /// clipped to a rectangle (x, y, width, height) when one is given, turned
    /// by `rotation` degrees clockwise about (`centerX`, `centerY`), at an
    /// `opacity` (1: opaque). `bbox`: a rectangle (x, y, width, height) on the
    /// page that holds all of it as it is painted - turned.
    | DrawGroup of clip: (float * float * float * float) option * rotation: float * centerX: float * centerY: float * opacity: float * bbox: (float * float * float * float) * commands: PaintCommand list
    /// Commands drawn mirrored across the vertical line through `centerX`
    /// (`flipH`) and / or the horizontal one through `centerY` (`flipV`): a
    /// flipped picture (a:xfrm flipH / flipV). `bbox`: the rectangle (x, y,
    /// width, height) that holds all of it.
    | DrawMirrored of flipH: bool * flipV: bool * centerX: float * centerY: float * bbox: (float * float * float * float) * commands: PaintCommand list

let commandYExtent = function
    | DrawText (_, style, _, baseline, _) ->
        let pad = max 4.0 style.FontSize
        Some (baseline - style.FontSize - pad, baseline + pad)
    | DrawUnderline (_, lineStyle, _, y, _) ->
        let extra = if lineStyle.ToLowerInvariant().Contains("double") then 4.0 else 2.0
        Some (y - extra, y + extra)
    | DrawImage (_, _, y, _, height, _, _, _, _) ->
        Some (y, y + height)
    | FillRect (_, _, y, _, height) ->
        Some (y, y + height)
    | StrokeRect (_, lineWidth, _, y, _, height) ->
        let pad = (max 1.0 lineWidth) + 1.0
        Some (y - pad, y + height + pad)
    | StrokeLine (_, lineWidth, _, _, _, y1, _, y2) ->
        let pad = (max 1.0 lineWidth) + 2.0
        Some (min y1 y2 - pad, max y1 y2 + pad)
    | DrawShape (_, _, strokeWidth, (_, y, _, height), segs) ->
        let pad = (max 1.0 strokeWidth) + 1.0
        let mutable top = y
        let mutable bottom = y + height
        for seg in segs do
            match seg with
            | PMove (_, sy)
            | PLine (_, sy) ->
                top <- min top sy
                bottom <- max bottom sy
            | PQuad (_, cy, _, sy) ->
                top <- min top (min cy sy)
                bottom <- max bottom (max cy sy)
            | PCubic (_, c1y, _, c2y, _, sy) ->
                top <- min top (min c1y (min c2y sy))
                bottom <- max bottom (max c1y (max c2y sy))
            | PClose -> ()
        Some (top - pad, bottom + pad)
    | DrawGroup (_, _, _, _, _, (_, y, _, height), _)
    | DrawMirrored (_, _, _, _, (_, y, _, height), _) ->
        Some (y - 1.0, y + height + 1.0)

// ---- anchored drawings and the order they are painted in --------------------
// A page is painted in layers: the anchored drawings that stand behind the
// text (wp:anchor behindDoc="1"), the text, the anchored drawings in front of
// it - the drawings of a layer in the order of their relativeHeight, the
// lowest first, and those of one height in the order of the document. The
// headers and footers of a page - their drawings too, a watermark - lie below
// all of the body. (Every drawing was painted with its paragraph, a text box
// after all the text: what covered what followed from the order of the
// paragraphs.)

/// The paint commands of an anchored drawing (a floating picture, a text box,
/// a shape) and its place among the drawings of its page.
type Drawing = {
    DrawingBehind: bool
    DrawingZ: float
    DrawingCommands: PaintCommand array
}

/// The number of a relativeHeight (0 when the file gives none).
let private drawingHeight (value: string option) : float =
    match value with
    | Some text ->
        // (The one-argument form: Fable's parse with a number style and a culture found no number.)
        match System.Double.TryParse(text: string) with
        | true, number -> number
        | false, _ -> 0.0
    | None -> 0.0

/// Where the anchored drawings of what is being collected go (null: with the
/// commands of their paragraph, in the order of the text).
let mutable private drawingSink : ResizeArray<Drawing> = null

/// `collect`, the anchored drawings of what it collects given to `sink`.
let private withDrawingSink (sink: ResizeArray<Drawing>) (collect: unit -> 'T) : 'T =
    let outer = drawingSink
    drawingSink <- sink
    try collect ()
    finally drawingSink <- outer

/// The drawings of a layer in the order they are painted in.
let private drawingsInOrder (behind: bool) (drawings: ResizeArray<Drawing>) : Drawing list =
    drawings
    |> Seq.filter (fun drawing -> drawing.DrawingBehind = behind)
    |> Seq.toList
    // (List.sortBy keeps the order of the document among equal heights.)
    |> List.sortBy (fun drawing -> drawing.DrawingZ)

/// The commands of text and of its anchored drawings in the order of the layers.
let private commandsInLayers (own: ResizeArray<PaintCommand>) (drawings: ResizeArray<Drawing>) : ResizeArray<PaintCommand> =
    if drawings.Count = 0 then own
    else
        let result = ResizeArray<PaintCommand>(own.Count + drawings.Count * 2)
        for drawing in drawingsInOrder true drawings do result.AddRange drawing.DrawingCommands
        result.AddRange own
        for drawing in drawingsInOrder false drawings do result.AddRange drawing.DrawingCommands
        result

type private LayoutIndex = {
    ById: Dictionary<BlockId, BlockLayout>
    ByBlock: Dictionary<BlockId, ResizeArray<BlockLayout>>
    Footnotes: ResizeArray<BlockLayout>
    Endnotes: ResizeArray<BlockLayout>
    BodyLinesByPage: Dictionary<int, ResizeArray<MeasuredLine>>
}

let private pageOfY settings y =
    int (floor (y / (pageHeight settings + pageGap)))

let private buildLayoutIndex settings includeBodyLines (estimatedLayouts: int) (layouts: BlockLayout list) =
    let byId = Dictionary<BlockId, BlockLayout>(estimatedLayouts)
    let byBlock = Dictionary<BlockId, ResizeArray<BlockLayout>>()
    let footnotes = ResizeArray<BlockLayout>()
    let endnotes = ResizeArray<BlockLayout>()
    let bodyLinesByPage = Dictionary<int, ResizeArray<MeasuredLine>>()

    let addLine page line =
        match bodyLinesByPage.TryGetValue page with
        | true, lines -> lines.Add line
        | false, _ ->
            let lines = ResizeArray<MeasuredLine>()
            lines.Add line
            bodyLinesByPage.[page] <- lines

    for layout in layouts do
        match byBlock.TryGetValue layout.BlockId with
        | true, blockLayouts -> blockLayouts.Add layout
        | false, _ ->
            match byId.TryGetValue layout.BlockId with
            | true, previousLayout ->
                let blockLayouts = ResizeArray<BlockLayout>(2)
                blockLayouts.Add previousLayout
                blockLayouts.Add layout
                byBlock.[layout.BlockId] <- blockLayouts
            | false, _ -> ()
        byId.[layout.BlockId] <- layout

        if layout.BlockId.StartsWith footnoteBlockIdPrefix then
            footnotes.Add layout
        elif layout.BlockId.StartsWith endnoteBlockIdPrefix then
            endnotes.Add layout
        elif includeBodyLines then
            for line in layout.Lines do
                addLine (pageOfY settings line.Y) line

    { ById = byId; ByBlock = byBlock; Footnotes = footnotes; Endnotes = endnotes; BodyLinesByPage = bodyLinesByPage }

let private tryLayout id (index: LayoutIndex) =
    match index.ById.TryGetValue id with
    | true, layout -> Some layout
    | false, _ -> None

let private tryFirstBlockLayout id (index: LayoutIndex) =
    match index.ByBlock.TryGetValue id with
    | true, layouts when layouts.Count > 0 -> Some layouts.[0]
    | _ ->
        match index.ById.TryGetValue id with
        | true, layout -> Some layout
        | false, _ -> None

let private minResizeArray (values: ResizeArray<float>) =
    if values.Count = 0 then None
    else
        let mutable minimum = values.[0]
        for i = 1 to values.Count - 1 do
            if values.[i] < minimum then minimum <- values.[i]
        Some minimum

let private lineVerticalBounds defaultTop defaultBottom (lines: MeasuredLine list) =
    match lines with
    | [] -> defaultTop, defaultBottom
    | first :: rest ->
        let mutable bottom = first.Y + first.Height
        for line in rest do
            bottom <- line.Y + line.Height
        first.Y, bottom

let private styleFromTextStyle (style: TextStyle) = {
    Bold = style.Bold
    Italic = style.Italic
    // (A link is underlined unless its text says that it is not: the underline of a link can be switched off.)
    Underline = TextBuffer.underlineShown style || WordLayout.Text.isInsertedRevision style
    AllCaps = style.AllCaps
    SmallCaps = style.SmallCaps
    Strike = style.Strike || WordLayout.Text.isDeletedRevision style
    FontSize = WordLayout.Text.printFontSize style
    CharacterSpacing = style.CharacterSpacing
    FontFamily = style.FontFamily
    Color = WordLayout.Text.fallbackColor style
    Highlight = style.Highlight
    VertAlign = style.VertAlign
}

let private defaultPaintStyle fontSize fontFamily = {
    Bold = false
    Italic = false
    Underline = false
    AllCaps = false
    SmallCaps = false
    Strike = false
    FontSize = fontSize
    CharacterSpacing = 0.0
    FontFamily = Some fontFamily
    Color = "#000000"
    Highlight = None
    VertAlign = None
}

let private paginateBlock (settings: LayoutSettings) y blockHeight =
    WordLayout.Page.paginateBlock (pageHeight settings) settings.PaddingTop settings.MarginBottom y blockHeight

let private availableContentHeight settings y =
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
    max 0.0 (contentBottom - y)

let private borderDashStyle (style: string option) =
    if optionEqualsIgnoreCase "dashed" style then 2
    elif optionEqualsIgnoreCase "dotted" style then 1
    elif optionEqualsIgnoreCase "dotdash" style then 3
    else 0

let private lineBaselineWithOffset baselineOffset (fallbackFontSize: float) (line: MeasuredLine) =
    let mutable fontSize = fallbackFontSize
    for segment in line.Segments do
        match segment with
        | TextSegment (_, style, _, _, _) ->
            fontSize <- max fontSize (WordLayout.Text.printFontSize style)
        | ImageSegment _ -> ()
        // Use the equation's box ascent so its top fits within the line and
        // its baseline aligns with surrounding text (math extends upward).
        | MathSegment (_, _, _, _, ascent, _, _, _) ->
            fontSize <- max fontSize ascent
    line.Y + fontSize + baselineOffset

let private lineBaseline fallbackFontSize line =
    lineBaselineWithOffset 0.0 fallbackFontSize line

/// The baseline of a line of a paragraph. A line of an exact height
/// (w:lineRule="exact") that is lower than its text keeps the text's descent
/// above its bottom, as Word does - what does not fit is cut off at the top:
/// the baseline is a fifth of the text's size above the line's bottom. (It
/// stood the text's size below the line's top, which for 55 pt text in a line
/// of 54 pt is below the line: the text was drawn into what follows, or cut
/// off by its text box.)
let private paragraphLineBaseline baselineOffset (fallbackFontSize: float) (metadata: BlockMetadata) (line: MeasuredLine) =
    let baseline = lineBaselineWithOffset baselineOffset fallbackFontSize line
    if metadata.LineSpacingRule <> Some "exact" then baseline
    else
        let mutable size = 0.0
        for segment in line.Segments do
            match segment with
            | TextSegment (_, style, _, _, _) -> size <- max size (WordLayout.Text.printFontSize style)
            | _ -> ()
        if size > 0.0 && size * 1.2 > line.Height then min baseline (line.Y + line.Height - size * 0.2 + baselineOffset) else baseline

let private tabLeaderText (measureText: TextMeasurer) (style: TextStyle) (width: float) (leader: string option) =
    if optionTrimmedEqualsIgnoreCase "dot" leader && width > 0.0 then
        let dotAdvance = measureText "." style
        let leaderPitch = max 1.0 (WordLayout.Text.printFontSize style * 0.25)
        if dotAdvance <= 0.0 then None
        else
            let wideMeasuredDot = dotAdvance > leaderPitch * 1.05
            let inset =
                if wideMeasuredDot then min leaderPitch (width * 0.5)
                else min (dotAdvance * 0.35) (width * 0.5)
            let count =
                if wideMeasuredDot then
                    int (floor ((max 0.0 (width - inset)) / leaderPitch + 0.5))
                else
                    int (floor (width / dotAdvance + 0.5))
            if count <= 0 then None
            else
                let text = String.replicate count "."
                let targetWidth =
                    if wideMeasuredDot then max 0.0 (width - inset)
                    else max 0.0 (measureText text style - inset)
                Some (text, targetWidth, inset)
    else None

let private isTabOnlySegment (text: string) =
    if text.IndexOf('\t') < 0 then
        false
    else
        let mutable allWhitespace = true
        let mutable i = 0
        while allWhitespace && i < text.Length do
            if not (System.Char.IsWhiteSpace text.[i]) then
                allWhitespace <- false
            i <- i + 1
        allWhitespace

let private dashStyleForTableBorderStyle style =
    match Rbt.BorderGrid.styleKey style with
    | "dotted" -> 1
    | "dashed"
    | "dashsmallgap" -> 2
    | "dotdash"
    | "dotdotdash"
    | "dashdotstroked" -> 3
    | _ -> 0

let private usableColor fallback value =
    match value with
    | Some color when color <> "" && color <> "auto" -> color
    | _ -> fallback

let private startsWithIgnoreCase (prefix: string) (value: string) =
    value.StartsWith(prefix, System.StringComparison.OrdinalIgnoreCase)

let private percentPatternCoverage (pattern: string option) =
    let parsePercent (value: string) =
        let mutable index = 3
        let mutable finish = value.Length - 1
        while index <= finish && System.Char.IsWhiteSpace value.[index] do
            index <- index + 1
        while finish >= index && System.Char.IsWhiteSpace value.[finish] do
            finish <- finish - 1
        let mutable whole = 0.0
        let mutable fraction = 0.0
        let mutable scale = 1.0
        let mutable afterDot = false
        let mutable sawDigit = false
        let mutable valid = index <= finish
        while valid && index <= finish do
            let ch = value.[index]
            if ch >= '0' && ch <= '9' then
                sawDigit <- true
                let digit = float (int ch - int '0')
                if afterDot then
                    scale <- scale * 10.0
                    fraction <- fraction + digit / scale
                else
                    whole <- whole * 10.0 + digit
                index <- index + 1
            elif ch = '.' && not afterDot then
                afterDot <- true
                index <- index + 1
            else
                valid <- false
        if valid && sawDigit then Some (whole + fraction) else None

    match pattern with
    | Some value when startsWithIgnoreCase "pct" value ->
        match parsePercent value with
        | Some percent -> Some (max 0.0 (min 1.0 (percent / 100.0)))
        | None -> None
    | _ -> None

let private shadingBackgroundColor (shading: Shading) =
    if optionEqualsIgnoreCase "nil" shading.Pattern then None
    elif optionEqualsIgnoreCase "solid" shading.Pattern then Some (usableColor "#000000" shading.ForegroundColor)
    else Some (usableColor "#FFFFFF" shading.BackgroundColor)

let private tryParseHexColor (value: string) =
    let mutable start = 0
    let mutable finish = value.Length - 1
    while start <= finish && System.Char.IsWhiteSpace value.[start] do
        start <- start + 1
    while finish >= start && System.Char.IsWhiteSpace value.[finish] do
        finish <- finish - 1

    let hexStart = if start <= finish && value.[start] = '#' then start + 1 else start
    if finish - hexStart + 1 = 6 then
        let hexValue ch =
            if ch >= '0' && ch <= '9' then Some (int ch - int '0')
            elif ch >= 'a' && ch <= 'f' then Some (10 + int ch - int 'a')
            elif ch >= 'A' && ch <= 'F' then Some (10 + int ch - int 'A')
            else None
        let parse offset =
            let index = hexStart + offset
            match hexValue value.[index], hexValue value.[index + 1] with
            | Some hi, Some lo -> Some (hi * 16 + lo)
            | _ -> None
        match parse 0, parse 2, parse 4 with
        | Some r, Some g, Some b -> Some (r, g, b)
        | _ -> None
    else
        None

let private blendedColor background foreground coverage =
    match tryParseHexColor background, tryParseHexColor foreground with
    | Some (br, bg, bb), Some (fr, fg, fb) ->
        let mix b f = int (round (float b + (float f - float b) * coverage)) |> max 0 |> min 255
        sprintf "#%02X%02X%02X" (mix br fr) (mix bg fg) (mix bb fb)
    | _ -> foreground

let private fillRectClipped color x y width height =
    if width <= 0.0 || height <= 0.0 then []
    else [FillRect (color, x, y, width, height)]

let private horizontalStripeCommands color x y width height =
    let commands = ResizeArray<PaintCommand>()
    let spacing = 2.0
    let mutable yy = y + 0.5
    while yy <= y + height do
        commands.Add(StrokeLine (color, 1.0, 0, 0.0, x, yy, x + width, yy))
        yy <- yy + spacing
    resizeArrayToList commands

let private verticalStripeCommands color x y width height =
    let commands = ResizeArray<PaintCommand>()
    let spacing = 2.0
    let mutable xx = x + 0.5
    while xx <= x + width do
        commands.Add(StrokeLine (color, 1.0, 0, 0.0, xx, y, xx, y + height))
        xx <- xx + spacing
    resizeArrayToList commands

let private diagonalStripeCommands reverse color x y width height =
    let commands = ResizeArray<PaintCommand>()
    let spacing = 3.3
    if reverse then
        let mutable c = 0.0
        while c <= width + height do
            let u0 = max 0.0 (c - height)
            let u1 = min width c
            if u1 >= u0 then
                commands.Add(StrokeLine (color, 1.0, 0, 0.0, x + u0, y + c - u0, x + u1, y + c - u1))
            c <- c + spacing
    else
        let mutable c = -width
        while c <= height do
            let u0 = max 0.0 (-c)
            let u1 = min width (height - c)
            if u1 >= u0 then
                commands.Add(StrokeLine (color, 1.0, 0, 0.0, x + u0, y + u0 + c, x + u1, y + u1 + c))
            c <- c + spacing
    resizeArrayToList commands

let private shadingRectCommands (shading: Shading) x y width height =
    if width <= 0.0 || height <= 0.0 then []
    else
        let patternKey = shading.Pattern
        let foreground = usableColor "#000000" shading.ForegroundColor
        let backgroundColor = shadingBackgroundColor shading
        let background = backgroundColor |> Option.map (fun color -> FillRect (color, x, y, width, height) |> List.singleton) |> Option.defaultValue []
        let percentCoverage = percentPatternCoverage patternKey
        let pattern =
            if patternKey.IsNone
               || optionEqualsIgnoreCase "nil" patternKey
               || optionEqualsIgnoreCase "clear" patternKey
               || optionEqualsIgnoreCase "solid" patternKey then []
            elif optionEqualsIgnoreCase "horzstripe" patternKey then horizontalStripeCommands foreground x y width height
            elif optionEqualsIgnoreCase "vertstripe" patternKey then verticalStripeCommands foreground x y width height
            elif optionEqualsIgnoreCase "diagstripe" patternKey then diagonalStripeCommands false foreground x y width height
            elif optionEqualsIgnoreCase "reversediagstripe" patternKey then diagonalStripeCommands true foreground x y width height
            elif optionEqualsIgnoreCase "horzcross" patternKey then horizontalStripeCommands foreground x y width height @ verticalStripeCommands foreground x y width height
            elif optionEqualsIgnoreCase "diagcross" patternKey then diagonalStripeCommands false foreground x y width height @ diagonalStripeCommands true foreground x y width height
            else
                match percentCoverage with
                | Some coverage when coverage >= 1.0 -> fillRectClipped foreground x y width height
                | Some coverage when coverage > 0.0 ->
                    match backgroundColor with
                    | Some bg -> fillRectClipped (blendedColor bg foreground coverage) x y width height
                    | None -> fillRectClipped foreground x y width height
                | _ -> []
        if percentCoverage.IsSome then pattern else background @ pattern

let private tableFillCommands backgroundColor shading x y width height =
    match shading with
    | Some value -> shadingRectCommands value x y width height
    | None ->
        match backgroundColor with
        | Some color -> [FillRect (color, x, y, width, height)]
        | None -> []

/// How near two border segments' coordinates are for the segments to meet.
/// (Half a pixel: the rows of a table do not all start to the hundredth of a
/// pixel where the row above ends, and borders that miss each other by that
/// much were drawn as if nothing crossed them.)
let private borderJoinEpsilon = 0.5

let private segmentIsHorizontal (segment: Rbt.BorderGrid.Segment) =
    abs (segment.Y1 - segment.Y2) < borderJoinEpsilon

let private segmentIsVertical (segment: Rbt.BorderGrid.Segment) =
    abs (segment.X1 - segment.X2) < borderJoinEpsilon

let private coversCoordinate value first second =
    value >= min first second - borderJoinEpsilon
    && value <= max first second + borderJoinEpsilon

type private TableBorderJoinIndex = {
    Horizontal: Dictionary<int, ResizeArray<Rbt.BorderGrid.Segment>>
    Vertical: Dictionary<int, ResizeArray<Rbt.BorderGrid.Segment>>
}

let private borderJoinCoordinateKey value =
    int (floor (value / borderJoinEpsilon + 0.5))

let private addBorderJoinSegment (index: Dictionary<int, ResizeArray<Rbt.BorderGrid.Segment>>) coordinate segment =
    let key = borderJoinCoordinateKey coordinate
    match index.TryGetValue key with
    | true, values -> values.Add segment
    | false, _ ->
        let values = ResizeArray()
        values.Add segment
        index.[key] <- values

let private tableBorderJoinIndex (segments: Rbt.BorderGrid.Segment list) =
    let horizontal = Dictionary<int, ResizeArray<Rbt.BorderGrid.Segment>>()
    let vertical = Dictionary<int, ResizeArray<Rbt.BorderGrid.Segment>>()
    for segment in segments do
        if segmentIsHorizontal segment then
            addBorderJoinSegment horizontal segment.Y1 segment
        elif segmentIsVertical segment then
            addBorderJoinSegment vertical segment.X1 segment
    { Horizontal = horizontal; Vertical = vertical }

let private indexedPerpendicularHalfWidth (segmentsByCoordinate: Dictionary<int, ResizeArray<Rbt.BorderGrid.Segment>>) coordinate covers =
    let key = borderJoinCoordinateKey coordinate
    let mutable maxWidth = 0.0
    for delta in -1 .. 1 do
        match segmentsByCoordinate.TryGetValue(key + delta) with
        | false, _ -> ()
        | true, candidates ->
            for candidate in candidates do
                if covers candidate then
                    maxWidth <- max maxWidth (candidate.Stroke.Width / 2.0)
    maxWidth

/// What a border segment meets at one of its ends: the border that crosses
/// its line there on the side of the lower coordinates (above a horizontal
/// segment, left of a vertical one) and on the other side, and whether its
/// own line goes on past the end in a segment of its own.
[<Struct>]
type private BorderJoin = {
    Before: Rbt.BorderGrid.Stroke voption
    After: Rbt.BorderGrid.Stroke voption
    Continues: bool
}

let private noBorderJoin : BorderJoin = { Before = ValueNone; After = ValueNone; Continues = false }

let private widerStroke (current: Rbt.BorderGrid.Stroke voption) (candidate: Rbt.BorderGrid.Stroke) =
    match current with
    | ValueSome known when known.Width >= candidate.Width -> current
    | _ -> ValueSome candidate

let private tableBorderJoinCaps (index: TableBorderJoinIndex) (segment: Rbt.BorderGrid.Segment) : BorderJoin * BorderJoin =
    let horizontal = segmentIsHorizontal segment
    let vertical = segmentIsVertical segment
    // `along`: the coordinate along the segment's line, `across`: the line's own.
    let joinAt (x: float) (y: float) (atStart: bool) : BorderJoin =
        if not horizontal && not vertical then noBorderJoin
        else
            let along, across = if horizontal then x, y else y, x
            let crossing = if horizontal then index.Vertical else index.Horizontal
            let own = if horizontal then index.Horizontal else index.Vertical
            let mutable before = ValueNone
            let mutable after = ValueNone
            let mutable continues = false
            let alongKey = borderJoinCoordinateKey along
            for delta in -1 .. 1 do
                match crossing.TryGetValue(alongKey + delta) with
                | false, _ -> ()
                | true, candidates ->
                    for candidate in candidates do
                        let fixedAt, first, second = if horizontal then candidate.X1, candidate.Y1, candidate.Y2 else candidate.Y1, candidate.X1, candidate.X2
                        if abs (fixedAt - along) < borderJoinEpsilon && coversCoordinate across first second then
                            if min first second < across - borderJoinEpsilon then before <- widerStroke before candidate.Stroke
                            if max first second > across + borderJoinEpsilon then after <- widerStroke after candidate.Stroke
            let acrossKey = borderJoinCoordinateKey across
            for delta in -1 .. 1 do
                match own.TryGetValue(acrossKey + delta) with
                | false, _ -> ()
                | true, candidates ->
                    for candidate in candidates do
                        if not (obj.ReferenceEquals(candidate, segment)) then
                            let fixedAt, first, second = if horizontal then candidate.Y1, candidate.X1, candidate.X2 else candidate.X1, candidate.Y1, candidate.Y2
                            if abs (fixedAt - across) < borderJoinEpsilon then
                                if atStart then
                                    if min first second < along - borderJoinEpsilon && max first second >= along - borderJoinEpsilon then continues <- true
                                elif max first second > along + borderJoinEpsilon && min first second <= along + borderJoinEpsilon then continues <- true
            { Before = before; After = after; Continues = continues }
    joinAt segment.X1 segment.Y1 true, joinAt segment.X2 segment.Y2 false

/// How far one line of a border segment runs past the segment's end (less
/// than nothing: it stops short of it), so that the lines of compound
/// borders meet as Word draws them. `partOffset`: where the line lies across
/// the border (below zero: on the side of the lower coordinates).
///
/// A border that crosses is itself one or more lines. A line on a side where
/// a border crosses turns into that border's nearest line (the inner lines
/// of a double border meet at a corner, and both stop at the first line of a
/// border they run into); a line on a side where none crosses goes straight
/// on when its border continues, and otherwise runs round the outside of
/// the corner, out to the far line of the border on the other side. With
/// single lines all of this is the old rule: half the crossing border's width.
let private borderLineCap (join: BorderJoin) (partOffset: float) (atStart: bool) : float =
    let lines (stroke: Rbt.BorderGrid.Stroke) =
        let width = max 0.25 stroke.Width
        let parts = Rbt.BorderGrid.strokePattern stroke.Style
        let firstOffset, firstWidth = List.head parts
        let lastOffset, lastWidth = List.last parts
        struct (firstOffset * width, max 0.25 (firstWidth * width), lastOffset * width, max 0.25 (lastWidth * width))
    // (At its start a segment lies after the junction: the crossing border's last line is the nearest.)
    let near (stroke: Rbt.BorderGrid.Stroke) =
        let struct (firstOffset, firstWidth, lastOffset, lastWidth) = lines stroke
        if atStart then lastWidth / 2.0 - lastOffset else firstOffset + firstWidth / 2.0
    let far (stroke: Rbt.BorderGrid.Stroke) =
        let struct (firstOffset, firstWidth, lastOffset, lastWidth) = lines stroke
        if atStart then firstWidth / 2.0 - firstOffset else lastOffset + lastWidth / 2.0
    let side (own: Rbt.BorderGrid.Stroke voption) (other: Rbt.BorderGrid.Stroke voption) =
        match own with
        | ValueSome stroke -> near stroke
        | ValueNone ->
            if join.Continues then 0.0
            else
                match other with
                | ValueSome stroke -> far stroke
                | ValueNone -> 0.0
    if partOffset < -0.001 then side join.Before join.After
    elif partOffset > 0.001 then side join.After join.Before
    else
        // A line in the middle of its border (a single line): it ends at the first line of a border that
        // runs through, and covers the corner of one it turns into.
        match join.Before, join.After with
        | ValueSome before, ValueSome after ->
            let stroke = if before.Width >= after.Width then before else after
            if join.Continues then far stroke else near stroke
        | ValueSome stroke, ValueNone
        | ValueNone, ValueSome stroke -> far stroke
        | ValueNone, ValueNone -> 0.0

let private addStrokeTableBorderCommands (commands: ResizeArray<PaintCommand>) color width style x1 y1 x2 y2 (startJoin: BorderJoin) (endJoin: BorderJoin) =
    let width = max 0.25 width
    let dashStyle = dashStyleForTableBorderStyle style
    for patternOffset, partWidth in Rbt.BorderGrid.strokePattern style do
        let offset = patternOffset * width
        let partWidth = max 0.25 (partWidth * width)
        let startCap = borderLineCap startJoin patternOffset true
        let endCap = borderLineCap endJoin patternOffset false
        if y1 = y2 then
            commands.Add(StrokeLine (color, partWidth, dashStyle, 0.0, x1 - startCap, y1 + offset, x2 + endCap, y2 + offset))
        else
            commands.Add(StrokeLine (color, partWidth, dashStyle, 0.0, x1 + offset, y1 - startCap, x2 + offset, y2 + endCap))

let private tableFragmentBorderGridStroke priority (stroke: WordLayout.Table.PrintStroke option) =
    stroke
    |> Option.map (fun stroke -> ({
        Color = stroke.Color
        Width = stroke.Width
        Style = stroke.Style
        Priority = priority
    }: Rbt.BorderGrid.Stroke))

/// The border cells of a row fragment. `topOf` / `bottomOf`: the border a
/// cell has at its top / bottom edge (its own, or - where a page break parts
/// the row from its neighbour - the one the neighbour holds, see
/// tableFragmentPaintCommands).
let private tableFragmentBorderCellsWith (topOf: TableCellFragmentPlan -> WordLayout.Table.PrintStroke option) (bottomOf: TableCellFragmentPlan -> WordLayout.Table.PrintStroke option) (rowFragments: ResizeArray<TableCellFragmentPlan>) =
    let result = ResizeArray<Rbt.BorderGrid.Cell>(rowFragments.Count)
    for cell in rowFragments do
        let fragmentStartsAfterRow = cell.FragmentY > cell.RowY + 0.01
        let fragmentEndsBeforeRow = cell.FragmentY + cell.FragmentHeight < cell.RowY + cell.RowHeight - 0.01
        result.Add({
            Row = cell.RowIndex
            Column = cell.StartColumn
            RowSpan = cell.RowSpan
            ColumnSpan = cell.ColumnSpan
            X = cell.CellX - cell.CellSpacing
            Y = cell.CellY
            Width = cell.CellWidth
            Height = cell.CellHeight
            Edges = {
                Top = if fragmentStartsAfterRow then None else tableFragmentBorderGridStroke 1 (topOf cell)
                Right = tableFragmentBorderGridStroke 1 cell.BorderRight
                Bottom = if fragmentEndsBeforeRow then None else tableFragmentBorderGridStroke 1 (bottomOf cell)
                Left = tableFragmentBorderGridStroke 1 cell.BorderLeft
            }
        }: Rbt.BorderGrid.Cell)
    result

/// A row fragment that is painted: where it is, its cells, and whether it is
/// a header row painted again at the top of a later page.
[<Struct>]
type private PaintedRowFragment = {
    PaintPage: int
    PaintY: float
    PaintHeight: float
    PaintRow: int
    PaintCells: ResizeArray<TableCellFragmentPlan>
    PaintFills: ResizeArray<PaintCommand>
    RepeatedHeader: bool
}

let tableFragmentPaintCommands (plan: TableLayoutPlan) =
    let commands = ResizeArray<PaintCommand>()
    let tableSpacing = WordLayout.Table.cellSpacing WordLayout.canvasScale plan.Table
    let tablePaintX = plan.TableX - tableSpacing
    let tableBackgroundX = plan.TableX - tableSpacing * 2.0
    let cellFragments = tableCellPaintFragmentPlans plan
    let mutable remainingCellFragments = cellFragments
    ignore tablePaintX
    // The row fragments that are painted, in order.
    let painted = ResizeArray<PaintedRowFragment>()
    // The cells of each row (of its first fragment): what a neighbour's border is taken from.
    let rowCells = Dictionary<int, ResizeArray<TableCellFragmentPlan>>()
    // The header rows as they stand at the top of the table: (y, height, cells).
    let headerRows = ResizeArray<struct (float * float * ResizeArray<TableCellFragmentPlan>)>()
    let headerHeight = plan.Rows |> List.sumBy (fun row -> if row.IsHeader && row.RowIndex < plan.HeaderRowCount then row.Height else 0.0)
    let fillsOf (y: float) (height: float) (cells: ResizeArray<TableCellFragmentPlan>) =
        let fills = ResizeArray<PaintCommand>()
        tableFillCommands plan.Table.BackgroundColor plan.Table.Shading tableBackgroundX y plan.TableWidth height
        |> List.iter fills.Add
        for cellFragment in cells do
            tableFillCommands cellFragment.CellBackgroundColor cellFragment.CellShading (cellFragment.CellX - cellFragment.CellSpacing) cellFragment.CellY cellFragment.CellWidth cellFragment.CellHeight
            |> List.iter fills.Add
        fills

    for row in plan.Rows do
        for fragment in row.Fragments do
            let rowFragments = ResizeArray<TableCellFragmentPlan>()
            let mutable consume = true
            while consume do
                match remainingCellFragments with
                | cell :: tail when cell.RowIndex = row.RowIndex
                                    && cell.Page = fragment.Page
                                    && cell.FragmentY = fragment.Y
                                    && cell.FragmentHeight = fragment.Height ->
                    rowFragments.Add cell
                    remainingCellFragments <- tail
                | _ ->
                    consume <- false
            if not (rowCells.ContainsKey row.RowIndex) then rowCells.[row.RowIndex] <- rowFragments
            if row.IsHeader && row.RowIndex < plan.HeaderRowCount && row.Fragments.Length = 1 then
                headerRows.Add(struct (fragment.Y, fragment.Height, rowFragments))

            let alreadyConsumedEmptyContinuation =
                fragment.LineCount = 0
                && fragment.ConsumedLineStart = fragment.ConsumedLineEnd
                && fragment.ConsumedLineStart > 0
            if not alreadyConsumedEmptyContinuation then
                // The header rows are repeated above the first row of a later page (the
                // layout leaves their height free there and lays their text out again):
                // their fills and borders are painted there too.
                let startsPage = painted.Count > 0 && painted.[painted.Count - 1].PaintPage < fragment.Page
                if startsPage && row.HeadersAbove && headerRows.Count = plan.HeaderRowCount && headerHeight > 0.0
                   && abs (fragment.Y - row.Y) < 0.01 then
                    let mutable top = fragment.Y - headerHeight
                    for struct (headerY, height, cells) in headerRows do
                        let dy = top - headerY
                        let moved = ResizeArray<TableCellFragmentPlan>(cells.Count)
                        for cell in cells do
                            moved.Add { cell with Page = fragment.Page; RowY = cell.RowY + dy; FragmentY = cell.FragmentY + dy; CellY = cell.CellY + dy; ContentY = cell.ContentY + dy }
                        let headerRowIndex = if cells.Count > 0 then cells.[0].RowIndex else 0
                        painted.Add { PaintPage = fragment.Page; PaintY = top; PaintHeight = height; PaintRow = headerRowIndex; PaintCells = moved; PaintFills = fillsOf top height moved; RepeatedHeader = true }
                        top <- top + height
                let fills =
                    if rowFragments.Count > 0 then fillsOf fragment.Y fragment.Height rowFragments
                    else fillsOf fragment.Y fragment.Height (ResizeArray<TableCellFragmentPlan>())
                painted.Add { PaintPage = fragment.Page; PaintY = fragment.Y; PaintHeight = fragment.Height; PaintRow = row.RowIndex; PaintCells = rowFragments; PaintFills = fills; RepeatedHeader = false }

    // The border between two rows belongs to one of them: the upper row's
    // bottom or the lower row's top, whichever won where they meet. Where the
    // two are not painted one under the other - a page break between them, a
    // header row repeated above a row that is not the one after it - each has
    // the line: the first row of a page had no top border when the row before
    // it held the border, the last row of a page no bottom border when the
    // row after it did.
    let follows (upper: int) (lower: int) =
        upper >= 0 && lower < painted.Count
        && painted.[upper].PaintPage = painted.[lower].PaintPage
        && abs (painted.[upper].PaintY + painted.[upper].PaintHeight - painted.[lower].PaintY) < 1.0
    let covering (rowIndex: int) (column: int) =
        match rowCells.TryGetValue rowIndex with
        | true, cells -> cells |> Seq.tryFind (fun cell -> cell.StartColumn <= column && column < cell.StartColumn + cell.ColumnSpan)
        | false, _ -> None
    let segments = Array.zeroCreate<Rbt.BorderGrid.Segment list> painted.Count
    for index in 0 .. painted.Count - 1 do
        let fragment = painted.[index]
        if fragment.PaintCells.Count = 0 then segments.[index] <- []
        else
            // (The first row of the table has its own top border, the last its own bottom.)
            let naturalAbove = index > 0 && follows (index - 1) index && painted.[index - 1].PaintRow = fragment.PaintRow - 1 && not painted.[index - 1].RepeatedHeader
            let naturalBelow = index + 1 < painted.Count && follows index (index + 1) && painted.[index + 1].PaintRow = fragment.PaintRow + 1 && not fragment.RepeatedHeader
            let underSomething = index > 0 && follows (index - 1) index
            let topOf (cell: TableCellFragmentPlan) =
                match cell.BorderTop with
                | Some stroke -> Some stroke
                // (Below a repeated header row the header's bottom border is the line.)
                | None when not naturalAbove && not underSomething -> covering (cell.RowIndex - 1) cell.StartColumn |> Option.bind (fun above -> above.BorderBottom)
                | None -> None
            let bottomOf (cell: TableCellFragmentPlan) =
                match cell.BorderBottom with
                | Some stroke -> Some stroke
                | None when not naturalBelow && cell.RowSpan <= 1 -> covering (cell.RowIndex + 1) cell.StartColumn |> Option.bind (fun below -> below.BorderTop)
                | None -> None
            segments.[index] <- tableFragmentBorderCellsWith topOf bottomOf fragment.PaintCells |> Rbt.BorderGrid.resolveResizeArray

    // A row's borders are drawn with the row, and meet the borders of the rows next to it on
    // the page: where a line of a compound border ends depends on what crosses it there, and
    // at the bottom of a row that is the row below.
    for index in 0 .. painted.Count - 1 do
        commands.AddRange painted.[index].PaintFills
        let borderSegments = segments.[index]
        if not borderSegments.IsEmpty then
            let neighbours =
                [ if follows (index - 1) index then yield! segments.[index - 1]
                  yield! borderSegments
                  if index + 1 < painted.Count && follows index (index + 1) then yield! segments.[index + 1] ]
            let joinIndex = tableBorderJoinIndex neighbours
            for segment in borderSegments do
                let startCap, endCap = tableBorderJoinCaps joinIndex segment
                addStrokeTableBorderCommands commands segment.Stroke.Color segment.Stroke.Width segment.Stroke.Style segment.X1 segment.Y1 segment.X2 segment.Y2 startCap endCap

    let bottomSpacing = WordLayout.Table.cellSpacing WordLayout.canvasScale plan.Table
    if bottomSpacing > 0.0 then
        match lastTableRowFragment plan.Rows with
        | Some fragment ->
            tableFillCommands plan.Table.BackgroundColor plan.Table.Shading tableBackgroundX (fragment.Y + fragment.Height) plan.TableWidth bottomSpacing
            |> List.iter commands.Add
        | None -> ()

    resizeArrayToList commands

let private paragraphBorderCommands (settings: LayoutSettings) (layout: BlockLayout) =
    let borderLine (border: ParagraphBorder option) x1 y1 x2 y2 =
        border
        |> Option.bind (fun value ->
            if optionEqualsIgnoreCase "none" value.Style || optionEqualsIgnoreCase "nil" value.Style then None
            else
                let color = value.Color |> Option.filter (fun color -> color <> "" && color <> "auto") |> Option.defaultValue "#000000"
                let width = value.Width |> Option.defaultValue 0.5 |> max 0.25
                Some (StrokeLine (color, width, borderDashStyle value.Style, 0.0, x1, y1, x2, y2)))

    match layout.Metadata.ParagraphBorders with
    | None -> []
    | Some borders ->
        let visualInset = 5.0
        let borderSpace border = border |> Option.bind (fun value -> value.Space) |> Option.defaultValue 0.0
        let leftSpace = borderSpace borders.Left + visualInset
        let rightSpace = borderSpace borders.Right + visualInset
        let topSpace = borderSpace borders.Top + visualInset
        let bottomSpace = borderSpace borders.Bottom + visualInset
        let left = settings.PaddingLeft + layout.Indent - leftSpace
        let right = settings.PaddingLeft + layout.Indent + max 1.0 layout.AvailableWidth + rightSpace
        let contentTop, contentBottom = lineVerticalBounds layout.Y (layout.Y + layout.Height) layout.Lines
        let top = contentTop - topSpace
        let bottom = contentBottom + bottomSpace
        let commands = ResizeArray<PaintCommand>()
        match borderLine borders.Top left top right top with
        | Some command -> commands.Add command
        | None -> ()
        match borderLine borders.Right right top right bottom with
        | Some command -> commands.Add command
        | None -> ()
        match borderLine borders.Bottom left bottom right bottom with
        | Some command -> commands.Add command
        | None -> ()
        match borderLine borders.Left left top left bottom with
        | Some command -> commands.Add command
        | None -> ()
        resizeArrayToList commands

let private paginateTableRow settings (row: TableRow) y rowHeight =
    let canSplit = row.CanSplit |> Option.defaultValue true
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
    let remaining = max 0.0 (contentBottom - y)
    if canSplit && y < contentBottom && rowHeight > remaining then y else paginateBlock settings y rowHeight

let private rowFragments settings y height =
    let rowBottom = y + height
    let fragments = ResizeArray<float * float>()
    let mutable pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    while pageTop < rowBottom do
        let top = max y (WordLayout.Page.contentTop settings.PaddingTop pageTop)
        let bottom = min rowBottom (WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop)
        if bottom > top then fragments.Add(top, bottom - top)
        pageTop <- pageTop + pageHeight settings + pageGap
    resizeArrayToList fragments

let private addParagraphCommandsWithBaselineOffset baselineOffset (settings: LayoutSettings) (fontFamily: string) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (commands: ResizeArray<PaintCommand>) (layout: BlockLayout) =
    let blockX = settings.PaddingLeft + layout.Indent
    let blockWidth = max 1.0 layout.AvailableWidth
    let shading =
        layout.Metadata.ParagraphShadingPattern
        |> Option.map (fun pattern -> { Pattern = Some pattern; ForegroundColor = layout.Metadata.ParagraphShadingForeground; BackgroundColor = layout.Metadata.BackgroundColor })
    commands.AddRange(tableFillCommands layout.Metadata.BackgroundColor shading blockX layout.Y blockWidth layout.Height)

    match layout.Metadata.ParagraphBorders with
    | Some _ -> commands.AddRange(paragraphBorderCommands settings layout)
    | None -> ()

    // (A content control's frame is not part of the page: see markupMarkerCommands.)

    if layout.Prefix <> "" && not layout.Lines.IsEmpty then
        let firstLine = layout.Lines.Head
        let baseline = paragraphLineBaseline baselineOffset settings.FontSize layout.Metadata firstLine
        let firstLineIndent = WordLayout.Paragraph.firstLineIndent layout.Metadata
        // (The number stands at the first-line indent, hanging or not.)
        let prefixX = settings.PaddingLeft + layout.Indent + firstLineIndent
        let prefixText = WordLayout.Numbering.displayMarkerText (layout.PrefixStyle |> Option.bind (fun style -> style.FontFamily)) layout.Prefix
        let prefixStyle =
            if prefixText <> layout.Prefix then
                // (In the document's face - the host has no Wingdings or Symbol -, in the marker's colour.)
                let plain = defaultPaintStyle settings.FontSize fontFamily
                match layout.PrefixStyle with
                | Some style -> { plain with Color = WordLayout.Text.fallbackColor style }
                | None -> plain
            else
                layout.PrefixStyle
                |> Option.map styleFromTextStyle
                |> Option.defaultValue (defaultPaintStyle settings.FontSize fontFamily)
        // (w:lvlJc: a number that is right-aligned ends where a left-aligned one starts - roman
        // numbers in a column; a centred one has its middle there.)
        let prefixX =
            match WordLayout.Numbering.markerJustification layout.Metadata, layout.PrefixStyle with
            | ("right" | "end"), Some style -> prefixX - measureText (prefixText.TrimEnd()) style
            | "center", Some style -> prefixX - measureText (prefixText.TrimEnd()) style / 2.0
            | _ -> prefixX
        let prefixDrawWidth =
            if prefixText <> layout.Prefix then 0.0
            else
                layout.PrefixStyle
                |> Option.map (measureText prefixText)
                |> Option.defaultValue 0.0
        // (A note's mark stands raised, as a superscript run of the text does.)
        let prefixBaseline =
            match prefixStyle.VertAlign with
            | Some "superscript" -> baseline - prefixStyle.FontSize * 0.5
            | Some "subscript" -> baseline + prefixStyle.FontSize * 0.25
            | _ -> baseline
        commands.Add(DrawText (prefixText, prefixStyle, prefixX, prefixBaseline, prefixDrawWidth))

    for line in layout.Lines do
        let baseline = paragraphLineBaseline baselineOffset settings.FontSize layout.Metadata line
        let lineAlignmentOffset = alignmentOffset layout line

        for segment in line.Segments do
            match segment with
            | TextSegment (text, style, x, width, _) ->
                let paintStyle = styleFromTextStyle style
                let textX = settings.PaddingLeft + x + lineAlignmentOffset
                match paintStyle.Highlight with
                | Some color -> commands.Add(FillRect (color, textX, line.Y, width, line.Height))
                | None -> ()
                let textBaseline =
                    let b =
                        match paintStyle.VertAlign with
                        | Some "superscript" -> baseline - paintStyle.FontSize * 0.5
                        | Some "subscript" -> baseline + paintStyle.FontSize * 0.25
                        | _ -> baseline
                    // w:position raises(+)/lowers(-) the baseline by N points.
                    b - style.Position * 96.0 / 72.0
                let displayText = WordLayout.Text.displayText style text
                if isTabOnlySegment displayText then
                    match tabLeaderText measureText style width layout.Metadata.RightTabLeader with
                    | Some (leaderText, leaderWidth, leaderInset) ->
                        commands.Add(DrawText (leaderText, paintStyle, textX + leaderInset, textBaseline, leaderWidth))
                    | None -> ()
                else
                    commands.Add(DrawText (displayText, paintStyle, textX, textBaseline, width))

                if TextBuffer.underlineShown style || style.Strike || style.DoubleStrike then
                    // Skip underline/strike on segments that are entirely whitespace —
                    // tab segments (eg. the gap between a TOC entry and its page
                    // number) inherit the surrounding run's hyperlink style and would
                    // otherwise draw a long underline across the tab gap.  Word treats
                    // the tab leader (dots/lines/underscores in w:leader) as the
                    // ink for that gap and doesn't extend the run's underline through
                    // it.
                    // Compute the visible-ink width: for segments that contain a
                    // trailing tab (eg. "1\t" in a TOC entry), the segment's `width`
                    // is bloated by the tab-gap advance, but the underline should
                    // only span the visible-ink prefix.  Word draws tab leaders
                    // (dots/lines/underscores per w:leader) in the gap and does
                    // *not* extend the run's underline through it.
                    let inkWidth =
                        let lastInkIndex =
                            let mutable idx = -1
                            for i in 0 .. displayText.Length - 1 do
                                if not (System.Char.IsWhiteSpace displayText.[i]) then idx <- i
                            idx
                        if lastInkIndex < 0 then 0.0
                        elif lastInkIndex = displayText.Length - 1 then width
                        else
                            // Measure the prefix up to and including the last non-
                            // whitespace character; fall back to width if measurement
                            // isn't sensible (eg. zero-width content).
                            let measured = measureRange displayText style 0 (lastInkIndex + 1)
                            if measured > 0.0 then measured else width
                    if TextBuffer.underlineShown style && inkWidth > 0.0 then
                        let ulColor =
                            match style.UnderlineColor with
                            | Some c when c <> "" && not (equalsIgnoreCase c "auto") -> if c.StartsWith("#") then c else "#" + c
                            | _ -> paintStyle.Color
                        let ulStyle = style.UnderlineStyle |> Option.defaultValue "single"
                        commands.Add(DrawUnderline (ulColor, ulStyle, textX, textBaseline + WordLayout.Text.canvasUnderlineOffset, inkWidth))
                    if style.Strike && inkWidth > 0.0 then
                        commands.Add(DrawUnderline (paintStyle.Color, "single", textX, textBaseline - WordLayout.Text.canvasStrikethroughOffset style, inkWidth))
                    if style.DoubleStrike && inkWidth > 0.0 then
                        let sy = textBaseline - WordLayout.Text.canvasStrikethroughOffset style
                        commands.Add(DrawUnderline (paintStyle.Color, "single", textX, sy - 1.0, inkWidth))
                        commands.Add(DrawUnderline (paintStyle.Color, "single", textX, sy + 1.5, inkWidth))
                match style.EmphasisMark with
                | Some mark when not (System.String.IsNullOrWhiteSpace displayText) ->
                    // A small mark above (or below for underDot) each non-space char.
                    let below = equalsIgnoreCase mark "underdot"
                    let markY = if below then textBaseline + 3.0 else textBaseline - paintStyle.FontSize * 0.95
                    let mutable acc = 0.0
                    for i in 0 .. displayText.Length - 1 do
                        let ch = displayText.[i]
                        let cw = measureRange displayText style i 1
                        if not (System.Char.IsWhiteSpace ch) then
                            commands.Add(FillRect (paintStyle.Color, textX + acc + cw / 2.0 - 1.0, markY, 2.0, 2.0))
                        acc <- acc + cw
                | _ -> ()

            | ImageSegment (src, x, yOffset, width, height, _, floating, _, relativeHeight, cropLeft, cropTop, cropRight, cropBottom, flip) ->
                // (An anchored picture is a drawing of its page: see Drawing.)
                let target = if floating && not (isNull drawingSink) then ResizeArray<PaintCommand>() else commands
                let imageX = settings.PaddingLeft + x + lineAlignmentOffset
                let imageY = line.Y + yOffset
                if src.StartsWith TextBuffer.opaqueSourcePrefix then
                    // An object kept from the file that has no picture to show: a frame of
                    // its size with its kind, for the ones a reader expects to see something
                    // (a chart, SmartArt, an embedded object). Other drawings leave their space empty.
                    let label =
                        match src.Substring TextBuffer.opaqueSourcePrefix.Length with
                        | "chart" -> Some "Chart"
                        | "diagram" -> Some "SmartArt"
                        | "ole" -> Some "Embedded object"
                        // A linked picture the host has not allowed to be loaded.
                        | "linked" -> Some "Linked picture"
                        | _ -> None
                    match label with
                    | Some text when width > 0.0 && height > 0.0 ->
                        target.Add(FillRect ("#F3F4F6", imageX, imageY, width, height))
                        target.Add(StrokeRect ("#9CA3AF", 1.0, imageX, imageY, width, height))
                        // The label at its natural size, in the middle of the frame. (The
                        // width given to DrawText is the width the text is fitted to: it is
                        // the label's own width, not the frame's.) A frame too narrow for
                        // the label shows as much of it as fits, with an ellipsis.
                        let labelTextStyle = { TextBuffer.defaultStyle with FontSize = 13.0; Color = Some "#4B5563" }
                        let labelStyle = styleFromTextStyle labelTextStyle
                        let room = width - 12.0
                        let fitted =
                            if measureText text labelTextStyle <= room then text
                            else
                                let mutable shown = text.Length - 1
                                while shown > 0 && measureText (text.Substring(0, shown) + "\u2026") labelTextStyle > room do
                                    shown <- shown - 1
                                if shown > 0 then text.Substring(0, shown) + "\u2026" else ""
                        if height >= 20.0 && fitted <> "" then
                            let labelWidth = measureText fitted labelTextStyle
                            target.Add(DrawText (fitted, labelStyle, imageX + (width - labelWidth) / 2.0, imageY + height / 2.0 + 13.0 * 0.35, labelWidth))
                    | _ -> ()
                else
                    let picture = DrawImage (src, imageX, imageY, width, height, cropLeft, cropTop, cropRight, cropBottom)
                    match flip with
                    | false, false -> target.Add picture
                    // (A flipped picture is drawn mirrored about its middle.)
                    | flipH, flipV -> target.Add(DrawMirrored (flipH, flipV, imageX + width / 2.0, imageY + height / 2.0, (imageX, imageY, width, height), [ picture ]))
                if not (obj.ReferenceEquals(target, commands)) && target.Count > 0 then
                    let behind, height = splitSegmentHeight relativeHeight
                    drawingSink.Add { DrawingBehind = behind; DrawingZ = drawingHeight height; DrawingCommands = target.ToArray() }

            | MathSegment (prims, style, x, _, _, _, _, _) ->
                // Emit the pre-laid-out equation primitives relative to the math
                // box origin (mathX) and the line's text baseline.
                let mathBase = styleFromTextStyle style
                let mathX = settings.PaddingLeft + x + lineAlignmentOffset
                for prim in prims do
                    match prim with
                    | Eq.MGlyph (text, gx, gb, fontPx, italic) ->
                        let glyphStyle = { mathBase with FontSize = fontPx; Italic = italic; FontFamily = mathBase.FontFamily |> Option.orElse (Some "Cambria Math") }
                        let gw = measureText text { style with FontSize = fontPx; Italic = italic }
                        commands.Add(DrawText (text, glyphStyle, mathX + gx, baseline + gb, gw))
                    | Eq.MRule (rx, top, w, h) ->
                        commands.Add(FillRect (mathBase.Color, mathX + rx, baseline + top, w, h))

    ()

let private paragraphCommandsWithBaselineOffset baselineOffset (settings: LayoutSettings) (fontFamily: string) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (layout: BlockLayout) =
    let commands = ResizeArray<PaintCommand>()
    addParagraphCommandsWithBaselineOffset baselineOffset settings fontFamily measureText measureRange commands layout
    resizeArrayToList commands

let private addParagraphCommands settings fontFamily measureText measureRange commands layout =
    addParagraphCommandsWithBaselineOffset 0.0 settings fontFamily measureText measureRange commands layout

let paragraphCommands settings fontFamily measureText measureRange layout =
    paragraphCommandsWithBaselineOffset 0.0 settings fontFamily measureText measureRange layout

// ---- the commands of a paragraph, kept with its layout ---------------------------
// The incremental layout hands back the same BlockLayout object for a
// paragraph that an edit did not touch or move. Its paint commands depend on
// that layout alone (and on the settings and the font it was laid out with;
// new measurers give new layouts), so they are kept with it: an edit collects
// the commands of the paragraphs that changed, not of the whole document -
// which was a third of the cost of a keystroke in a 600-page document
// (measuring every run again for its underline, its tab leaders ...).

/// What was collected for a layout, and with what.
type private ParagraphCommandsEntry = {
    Settings: LayoutSettings
    FontFamily: string
    Commands: PaintCommand array
    /// The paragraph's anchored drawings (not among Commands).
    Drawings: Drawing array
}

#if FABLE_COMPILER
[<Fable.Core.Emit("new WeakMap()")>]
let private newLayoutMap () : obj = Fable.Core.Util.jsNative
[<Fable.Core.Emit("$0.get($1)")>]
let private layoutMapGet (map: obj) (key: BlockLayout) : ParagraphCommandsEntry = Fable.Core.Util.jsNative
[<Fable.Core.Emit("$0.set($1, $2)")>]
let private layoutMapSet (map: obj) (key: BlockLayout) (value: ParagraphCommandsEntry) : unit = Fable.Core.Util.jsNative
let private paragraphCommandsByLayout = newLayoutMap ()
let private tryKnownParagraphCommands (layout: BlockLayout) : ParagraphCommandsEntry option =
    let entry = layoutMapGet paragraphCommandsByLayout layout
    if isNull (box entry) then None else Some entry
let private rememberParagraphCommands (layout: BlockLayout) (entry: ParagraphCommandsEntry) =
    layoutMapSet paragraphCommandsByLayout layout entry
#else
let private paragraphCommandsByLayout = System.Runtime.CompilerServices.ConditionalWeakTable<BlockLayout, ParagraphCommandsEntry>()
let private tryKnownParagraphCommands (layout: BlockLayout) : ParagraphCommandsEntry option =
    match paragraphCommandsByLayout.TryGetValue layout with
    | true, entry -> Some entry
    | false, _ -> None
let private rememberParagraphCommands (layout: BlockLayout) (entry: ParagraphCommandsEntry) =
    lock paragraphCommandsByLayout (fun () ->
        paragraphCommandsByLayout.Remove layout |> ignore
        paragraphCommandsByLayout.Add(layout, entry))
#endif

/// `addParagraphCommands`, with the commands kept for the next time the same
/// layout is painted.
let private addParagraphCommandsKept (settings: LayoutSettings) (fontFamily: string) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (commands: ResizeArray<PaintCommand>) (layout: BlockLayout) =
    let entry =
        match tryKnownParagraphCommands layout with
        | Some entry when entry.FontFamily = fontFamily && (obj.ReferenceEquals(entry.Settings, settings) || entry.Settings = settings) -> entry
        | _ ->
            let own = ResizeArray<PaintCommand>()
            let drawings = ResizeArray<Drawing>()
            withDrawingSink drawings (fun () -> addParagraphCommands settings fontFamily measureText measureRange own layout)
            let entry = { Settings = settings; FontFamily = fontFamily; Commands = own.ToArray(); Drawings = drawings.ToArray() }
            rememberParagraphCommands layout entry
            entry
    commands.AddRange entry.Commands
    if entry.Drawings.Length > 0 then
        if isNull drawingSink then
            for drawing in entry.Drawings do commands.AddRange drawing.DrawingCommands
        else drawingSink.AddRange entry.Drawings

// Keep line layout stable and correct only the drawn glyph baseline for Word's
// spaced table cells, whose visual text boxes sit higher than our default paint.
let private spacedTableCellBaselineOffset = -4.0

let private defaultParagraphFirstLineOffsetInDocument doc =
    let para = { Id = ""; Buffer = TextBuffer.empty; Metadata = defaultMetadata }
    let paragraphStyle = forBlockInDocument doc (Paragraph para)
    paragraphStyle.FirstLineIndent |> Option.defaultValue 0.0 |> max 0.0

let private defaultParagraphFirstLineOffsetInContext context =
    let para = { Id = ""; Buffer = TextBuffer.empty; Metadata = defaultMetadata }
    let paragraphStyle = forBlockInContext context (Paragraph para)
    paragraphStyle.FirstLineIndent |> Option.defaultValue 0.0 |> max 0.0

let private tableLeftOffsetForStyleContext = function
    | Some context -> defaultParagraphFirstLineOffsetInContext context
    | None -> 0.0

let private tableAlignmentOffset (alignment: string option) availableWidth tableWidth =
    if optionEqualsIgnoreCase "center" alignment then max 0.0 ((availableWidth - tableWidth) / 2.0)
    elif optionEqualsIgnoreCase "right" alignment || optionEqualsIgnoreCase "end" alignment then max 0.0 (availableWidth - tableWidth)
    else 0.0

/// The frame of a content control, as an editing aid: drawn around the
/// paragraphs of the control the caret or the selection is in, as Word does,
/// and nowhere else. It is an overlay: a document without a selection (a
/// viewer, a page rendered for print or preview) has none.
let markupMarkerCommands (members: Dictionary<BlockId, HashSet<string>>) (settings: LayoutSettings) (doc: Document) (layouts: BlockLayout list) =
    match doc.Selection with
    | None -> []
    | Some selection ->
        let selected = HashSet<string>()
        for blockId in [ selection.Anchor.BlockId; selection.Focus.BlockId ] do
            match members.TryGetValue blockId with
            | true, groups -> for group in groups do selected.Add group |> ignore
            | _ -> ()
        if selected.Count = 0 then []
        else
            let commands = ResizeArray<PaintCommand>()
            for layout in layouts do
                match members.TryGetValue layout.BlockId with
                | true, groups when groups |> Seq.exists selected.Contains ->
                    let blockX = settings.PaddingLeft + layout.Indent
                    let blockWidth = max 1.0 layout.AvailableWidth
                    commands.Add(StrokeRect ("#2563EB", 1.0, blockX, layout.Y, blockWidth, layout.Height))
                    commands.Add(FillRect ("#2563EB", blockX, layout.Y, min 32.0 blockWidth, min 3.0 layout.Height))
                | _ -> ()
            resizeArrayToList commands

let private maxNestedTablePaintDepth = 512

let rec private tableBorderCommandsCoreAt nestedTablePaintDepth (styleContext: DocumentStyleContext option) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (y: float) (availableWidth: float) (baseHorizontalOffset: float) (rowTops: float array option) =
    // `rowTops`: where the layout put each body row (None: rows follow each
    // other from `y`, moved to the next page when they do not fit). A nested
    // table is drawn from its rows' own layouts: the row of the outer table
    // that holds it may be split across a page break, and then its rows are
    // not where a table laid out on its own would have them.
    let mutable currentY = y
    let commands = ResizeArray<PaintCommand>()

    let numColumns = WordLayout.Table.columnCount tableBlock

    if nestedTablePaintDepth >= maxNestedTablePaintDepth then
        ([], currentY + WordLayout.Table.emptyTableAdvance settings.LineHeight)
    elif numColumns = 0 then
        ([], currentY + WordLayout.Table.emptyTableAdvance settings.LineHeight)
    else
        let outerLeft, outerRight = WordLayout.Table.outerCellInsets WordLayout.canvasScale settings.CompatibilityMode tableBlock
        let tableAvailableWidth = max 0.0 (availableWidth - WordLayout.Table.tableIndent tableBlock)
        let tableMaxWidth = tableAvailableWidth + outerLeft + outerRight
        let resolvedTable = WordLayout.Table.resolveTable WordLayout.canvasScale (Some tableMaxWidth) tableBlock
        let columnWidths = resolvedTable.ColumnWidths
        let columnWidthPrefix = columnWidthPrefix columnWidths
        let columnMetrics = tableColumnMetrics tableMaxWidth columnWidths
        let fallbackColumnWidth = tableMaxWidth / float numColumns
        let tableWidth = WordLayout.Table.tableWidthWithSpacing columnWidths (WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock)
        let tableHorizontalOffset = baseHorizontalOffset + WordLayout.Table.tableIndent tableBlock + tableAlignmentOffset tableBlock.Alignment tableAvailableWidth tableWidth
        let tableX = settings.PaddingLeft + tableHorizontalOffset - outerLeft - WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock

        let columnX spacing startColumn =
            tableX + spacedStartOffsetFromPrefix columnWidthPrefix spacing startColumn - WordLayout.Table.spacedCellPaintXOffset spacing startColumn

        let columnWidth spacing startColumn span =
            spacedSpannedWidthFromPrefix columnWidthPrefix spacing startColumn span fallbackColumnWidth
            - WordLayout.Table.spacedCellPaintWidthAdjustment spacing numColumns startColumn span
            |> max 0.0

        let borderStroke priority (stroke: WordLayout.Table.PrintStroke option) =
            stroke
            |> Option.map (fun stroke -> ({
                Color = stroke.Color
                Width = stroke.Width
                Style = stroke.Style
                Priority = priority
            }: Rbt.BorderGrid.Stroke))

        // The table's list numbering, built at most once for all its rows.
        let measureScope = lazy (tableMeasureScope None tableBlock)
        let rowHeightAt (row: WordLayout.Table.ResolvedRow) =
            let paragraphStyleForParagraph, effectiveBufferForParagraph, effectiveTextStyleForParagraph =
                match styleContext with
                | Some context -> forTableCellParagraphInContext context, effectiveBufferInContext context, effectiveTextStyleInContext context
                | None -> (fun _ -> normal), effectiveBuffer, effectiveTextStyle
            let rowLocalSettings = {
                settings with
                    PageHeight = 1000000.0
                    MarginBottom = 0.0
                    TableContinuationHeaderHeight = 0.0
            }
            match tryMeasureTableRowHeightWithMetrics rowLocalSettings measureText tableBlock columnMetrics row.Source paragraphStyleForParagraph effectiveTextStyleForParagraph with
            | Some height -> height
            | None ->
                let measurement = measureTableRowWithStyleUsingMetrics measureScope.Value rowLocalSettings measureText measureRange tableBlock columnMetrics outerLeft tableHorizontalOffset row.Source 0.0 0.0 paragraphStyleForParagraph effectiveBufferForParagraph
                measurement.Height

        let rowHeights = ResizeArray<float>()
        for row in resolvedTable.BodyRows do
            rowHeights.Add(rowHeightAt row)

        let rowHeightPrefix = Array.zeroCreate<float> (rowHeights.Count + 1)
        for index = 0 to rowHeights.Count - 1 do
            rowHeightPrefix.[index + 1] <- rowHeightPrefix.[index] + rowHeights.[index]

        // Where each body row starts. With `rowTops`, a row the layout put
        // further down than it would follow here (on the next page) starts there.
        let givenTop rowIndex =
            match rowTops with
            | Some tops when rowIndex < tops.Length && not (System.Double.IsNaN tops.[rowIndex]) -> Some tops.[rowIndex]
            | _ -> None
        let plannedTops =
            let tops = Array.zeroCreate<float> rowHeights.Count
            let mutable at = y
            let mutable index = 0
            for row in resolvedTable.BodyRows do
                let paged = paginateTableRow settings row.Source at rowHeights.[index]
                let top =
                    match givenTop index with
                    | Some given when given > paged + 0.5 -> given
                    | _ -> paged
                tops.[index] <- top
                at <- top + rowHeights.[index]
                index <- index + 1
            tops
        // The rows from `rowIndex` on, within a span, that follow each other
        // without a break (a break: the next row is on another page).
        let contiguousRows rowIndex rowSpan =
            let last = min rowHeights.Count (rowIndex + max 1 rowSpan)
            let mutable count = 1
            if rowTops.IsSome then
                let mutable index = rowIndex + 1
                let mutable together = true
                while together && index < last do
                    if abs (plannedTops.[index] - (plannedTops.[index - 1] + rowHeights.[index - 1])) < 1.0 then
                        count <- count + 1
                        index <- index + 1
                    else together <- false
            else count <- max 1 (last - rowIndex)
            count

        let rowSpanHeightAt rowIndex rowSpan fallback =
            let start = max 0 (min rowIndex rowHeights.Count)
            let finish = min rowHeights.Count (start + contiguousRows start rowSpan)
            max (rowHeightPrefix.[finish] - rowHeightPrefix.[start]) fallback

        // Cells merged over rows that a page break separates: the part below
        // the break (cell, x, width, the row it continues at, rows left).
        let continuations = ResizeArray<WordLayout.Table.ResolvedCell * float * float * int * int>()

        // The border between two rows belongs to one of them (the upper row's
        // bottom or the lower row's top, whichever won). Where a page break
        // parts the two, each page gets the line: the first row of a page had
        // no top border when the row before it held the border, the last row
        // of a page no bottom border when the row after it did. The same for a
        // header row repeated above a row that is not the one after it.
        let bodyRowArray = List.toArray resolvedTable.BodyRows
        let cellCovering (neighbour: int) (column: int) =
            if neighbour < 0 || neighbour >= bodyRowArray.Length then None
            else bodyRowArray.[neighbour].Cells |> List.tryFind (fun cell -> cell.StartColumn <= column && column < cell.StartColumn + cell.ColSpan)
        let topAtBreak openAbove rowIndex (resolved: WordLayout.Table.ResolvedCell) =
            match resolved.Borders.Top with
            | Some stroke -> Some stroke
            | None when openAbove -> cellCovering (rowIndex - 1) resolved.StartColumn |> Option.bind (fun above -> above.Borders.Bottom)
            | None -> None
        let bottomAtBreak openBelow rowIndex (resolved: WordLayout.Table.ResolvedCell) =
            match resolved.Borders.Bottom with
            | Some stroke -> Some stroke
            | None when openBelow && resolved.RowSpan <= 1 -> cellCovering (rowIndex + 1) resolved.StartColumn |> Option.bind (fun below -> below.Borders.Top)
            | None -> None

        let addRowCommands (openAbove: bool) (openBelow: bool) rowIndex (row: WordLayout.Table.ResolvedRow) currentY rowHeight =
            let cellSpacing = WordLayout.Table.rowCellSpacing WordLayout.canvasScale tableBlock row.Source
            let outerStartInset = outerLeft + cellSpacing
            for fragmentY, fragmentHeight in rowFragments settings currentY rowHeight do
                commands.AddRange(tableFillCommands tableBlock.BackgroundColor tableBlock.Shading tableX fragmentY tableWidth fragmentHeight)

            let cellBuffer = ResizeArray<Rbt.BorderGrid.Cell>(row.Cells.Length)
            let mutable borderFragmentHeight = rowHeight
            // The lower part of a cell merged from a row above the page break.
            for (resolved, x, width, atRow, remaining) in continuations.ToArray() do
                if atRow = rowIndex then
                    continuations.Remove ((resolved, x, width, atRow, remaining)) |> ignore
                    let together = contiguousRows rowIndex remaining
                    let finish = min rowHeights.Count (rowIndex + together)
                    let height = max rowHeight (rowHeightPrefix.[finish] - rowHeightPrefix.[rowIndex])
                    if height > borderFragmentHeight then borderFragmentHeight <- height
                    for fragmentY, fragmentHeight in rowFragments settings currentY height do
                        commands.AddRange(tableFillCommands resolved.Cell.BackgroundColor resolved.Cell.Shading x fragmentY width fragmentHeight)
                    if together < remaining then continuations.Add ((resolved, x, width, rowIndex + together, remaining - together))
                    cellBuffer.Add({
                        Row = 0
                        Column = resolved.StartColumn
                        RowSpan = remaining
                        ColumnSpan = resolved.ColSpan
                        X = x
                        Y = currentY
                        Width = width
                        Height = height
                        Edges = {
                            Top = None
                            Right = borderStroke 1 resolved.Borders.Right
                            Bottom = (if together < remaining then None else borderStroke 1 resolved.Borders.Bottom)
                            Left = borderStroke 1 resolved.Borders.Left
                        }
                    }: Rbt.BorderGrid.Cell)
            for resolved in row.Cells do
                let cell = resolved.Cell
                let x = columnX cellSpacing resolved.StartColumn
                let width = columnWidth cellSpacing resolved.StartColumn resolved.ColSpan
                let cellPadding = WordLayout.Table.cellPaddingForRow WordLayout.canvasScale tableBlock row.Source cell
                let textPadding =
                    if cellSpacing > 0.0 then
                        let rowPadding = WordLayout.Table.rowCellPaddingDefaults WordLayout.canvasScale tableBlock row.Source
                        { cellPadding with Top = rowPadding.Top; Bottom = rowPadding.Bottom }
                    else
                        cellPadding
                let paintTopAdjustment = WordLayout.Table.spacedCellPaintTopAdjustment cellSpacing textPadding cellPadding
                let cellY = currentY + WordLayout.Table.spacedCellPaintTopOffset cellSpacing - paintTopAdjustment
                let spannedRowHeight = rowSpanHeightAt rowIndex resolved.RowSpan rowHeight
                let cellHeight = WordLayout.Table.spacedCellPaintHeight spannedRowHeight cellSpacing + paintTopAdjustment
                if cellHeight > borderFragmentHeight then
                    borderFragmentHeight <- cellHeight
                for fragmentY, fragmentHeight in rowFragments settings cellY cellHeight do
                    commands.AddRange(tableFillCommands cell.BackgroundColor cell.Shading x fragmentY width fragmentHeight)

                if not cell.NestedTables.IsEmpty then
                    let contentColWidth = spacedSpannedWidthFromPrefix columnWidthPrefix cellSpacing resolved.StartColumn resolved.ColSpan fallbackColumnWidth
                    let nestedAvailableWidth = WordLayout.Table.cellContentWidth contentColWidth textPadding
                    let nestedX =
                        tableHorizontalOffset
                        + spacedStartOffsetFromPrefix columnWidthPrefix cellSpacing resolved.StartColumn
                        + textPadding.Left
                        - outerStartInset
                        - WordLayout.Table.internalSpacingTextOffset cellSpacing resolved.StartColumn
                    let nestedCellY = currentY + cellSpacing + WordLayout.Table.spacedCellTextTopOffset cellSpacing textPadding.Top
                    let mutable nestedY = WordLayout.Table.cellContentStartY nestedCellY textPadding
                    for nestedTable in cell.NestedTables do
                        let nestedCommands, nextNestedY = tableBorderCommandsCoreAt (nestedTablePaintDepth + 1) styleContext settings measureText measureRange nestedTable nestedY nestedAvailableWidth nestedX None
                        commands.AddRange nestedCommands
                        nestedY <- nextNestedY
                cellBuffer.Add({
                    Row = 0
                    Column = resolved.StartColumn
                    RowSpan = resolved.RowSpan
                    ColumnSpan = resolved.ColSpan
                    X = x
                    Y = cellY
                    Width = width
                    Height = cellHeight
                    Edges = {
                        Top = borderStroke 1 (topAtBreak openAbove rowIndex resolved)
                        Right = borderStroke 1 resolved.Borders.Right
                        // A merged cell cut by a page break is open at the break.
                        Bottom = (if contiguousRows rowIndex resolved.RowSpan < resolved.RowSpan then None else borderStroke 1 (bottomAtBreak openBelow rowIndex resolved))
                        Left = borderStroke 1 resolved.Borders.Left
                    }
                }: Rbt.BorderGrid.Cell)
                let together = contiguousRows rowIndex resolved.RowSpan
                if together < resolved.RowSpan then
                    continuations.Add ((resolved, x, width, rowIndex + together, resolved.RowSpan - together))

            let borderSegments = Rbt.BorderGrid.resolveResizeArray cellBuffer
            let joinIndex = tableBorderJoinIndex borderSegments
            for segment in borderSegments do
                let startCap, endCap = tableBorderJoinCaps joinIndex segment
                for fragmentY, fragmentHeight in rowFragments settings currentY borderFragmentHeight do
                    let y1 = max segment.Y1 fragmentY
                    let y2 = min segment.Y2 (fragmentY + fragmentHeight)
                    if segment.Y1 = segment.Y2 then
                        if segment.Y1 >= fragmentY && segment.Y1 <= fragmentY + fragmentHeight then
                            addStrokeTableBorderCommands commands segment.Stroke.Color segment.Stroke.Width segment.Stroke.Style segment.X1 segment.Y1 segment.X2 segment.Y2 startCap endCap
                    elif y2 > y1 then
                        let clippedStartCap = if abs (y1 - segment.Y1) < borderJoinEpsilon then startCap else noBorderJoin
                        let clippedEndCap = if abs (y2 - segment.Y2) < borderJoinEpsilon then endCap else noBorderJoin
                        addStrokeTableBorderCommands commands segment.Stroke.Color segment.Stroke.Width segment.Stroke.Style segment.X1 y1 segment.X2 y2 clippedStartCap clippedEndCap

            // Row-level tblPrEx override (BordersOverride): paint the row's
            // table-outer borders at the row's outer bounding box, OUTSIDE
            // the cells.  Coexists with cell tcBorders inside, matching Word
            // when a row's tblPrEx declares an outer frame distinct from
            // per-cell borders.  Only triggers when cellSpacing > 0 so the
            // outer position is geometrically distinct from cell edges.  The
            // outer box wraps the *cells* of this row, not the full grid:
            // gridBefore/gridAfter phantoms are outside this row's frame.
            match row.Source.BordersOverride with
            | None -> ()
            | Some overrideBorders ->
                if cellSpacing > 0.0 && cellBuffer.Count > 0 then
                    let firstCell = cellBuffer.[0]
                    let lastCell = cellBuffer.[cellBuffer.Count - 1]
                    let outerLeft = firstCell.X - cellSpacing
                    let outerRight = lastCell.X + lastCell.Width + cellSpacing
                    let outerTop = currentY
                    let outerBottom = currentY + rowHeight
                    let paintOuterBorder (border: TableBorder option) x1 y1 x2 y2 =
                        match border with
                        | None -> ()
                        | Some b ->
                            if optionEqualsIgnoreCase "nil" b.Style || optionEqualsIgnoreCase "none" b.Style then ()
                            else
                                let width = b.Width |> Option.defaultValue 0.5
                                let color = b.Color |> Option.defaultValue "#000000"
                                addStrokeTableBorderCommands commands color width b.Style x1 y1 x2 y2 noBorderJoin noBorderJoin
                    paintOuterBorder overrideBorders.Top outerLeft outerTop outerRight outerTop
                    paintOuterBorder overrideBorders.Bottom outerLeft outerBottom outerRight outerBottom
                    paintOuterBorder overrideBorders.Left outerLeft outerTop outerLeft outerBottom
                    paintOuterBorder overrideBorders.Right outerRight outerTop outerRight outerBottom

        let headerRows = resolvedTable.HeaderRows
        let rows = resolvedTable.BodyRows

        // Where the rows go, before anything is drawn: a row knows then whether
        // a page break follows it.
        let headerCount = List.length headerRows
        let headerHeightAt headerIndex (headerRow: WordLayout.Table.ResolvedRow) =
            if headerIndex < rowHeights.Count then rowHeights.[headerIndex] else rowHeightAt headerRow
        let rowTop = Array.zeroCreate<float> rowHeights.Count
        let movedDown = Array.zeroCreate<bool> rowHeights.Count
        let headersAbove = Array.zeroCreate<bool> rowHeights.Count
        let mutable placeIndex = 0
        let mutable placeY = currentY
        for row in rows do
            let rowHeight = rowHeights.[placeIndex]
            let pagedY = paginateTableRow settings row.Source placeY rowHeight
            let pagedY =
                match givenTop placeIndex with
                | Some given when given > pagedY + 0.5 -> given
                | _ -> pagedY
            movedDown.[placeIndex] <- placeIndex > 0 && pagedY <> placeY
            if pagedY <> placeY && not row.IsHeader && not headerRows.IsEmpty then
                headersAbove.[placeIndex] <- true
                placeY <- pagedY
                headerRows |> List.iteri (fun headerIndex headerRow -> placeY <- placeY + headerHeightAt headerIndex headerRow)
            else
                placeY <- pagedY
            rowTop.[placeIndex] <- placeY
            placeY <- placeY + rowHeight
            placeIndex <- placeIndex + 1

        let mutable rowIndex = 0
        for row in rows do
            let rowHeight = rowHeights.[rowIndex]
            if headersAbove.[rowIndex] then
                currentY <- rowTop.[rowIndex] - (headerRows |> List.mapi headerHeightAt |> List.sum)
                let mutable headerIndex = 0
                for headerRow in headerRows do
                    let headerHeight = headerHeightAt headerIndex headerRow
                    // (The last header row stands above a row that is not its neighbour.)
                    addRowCommands false (headerIndex = headerCount - 1) headerIndex headerRow currentY headerHeight
                    currentY <- currentY + headerHeight
                    headerIndex <- headerIndex + 1
            currentY <- rowTop.[rowIndex]
            let breakAfter = rowIndex + 1 < rowHeights.Count && movedDown.[rowIndex + 1]
            addRowCommands movedDown.[rowIndex] breakAfter rowIndex row currentY rowHeight
            currentY <- currentY + rowHeight
            rowIndex <- rowIndex + 1

        let bottomSpacing = WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock
        if bottomSpacing > 0.0 then
            commands.AddRange(tableFillCommands tableBlock.BackgroundColor tableBlock.Shading tableX currentY tableWidth bottomSpacing)

        (resizeArrayToList commands, currentY + WordLayout.Table.tableAfterSpacing WordLayout.canvasScale tableBlock)

let private tableBorderCommandsCore (styleContext: DocumentStyleContext option) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (y: float) =
    let maxWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
    let tableLeftOffset = tableLeftOffsetForStyleContext styleContext
    let tableAvailableWidth = max 0.0 (maxWidth - tableLeftOffset)
    tableBorderCommandsCoreAt 0 styleContext settings measureText measureRange tableBlock y tableAvailableWidth tableLeftOffset None

let tableBorderCommandsForDocumentWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (tableBlock: TableBlock) (y: float) =
    tableBorderCommandsCore (Some (WordStyles.styleContext doc)) settings measureText measureRange tableBlock y

let tableBorderCommandsWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (y: float) =
    tableBorderCommandsCore None settings measureText measureRange tableBlock y

let tableBorderCommands (settings: LayoutSettings) (measureText: TextMeasurer) (tableBlock: TableBlock) (y: float) =
    tableBorderCommandsWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) tableBlock y

let private fieldCodeName (code: string) =
    let mutable i = 0
    while i < code.Length && System.Char.IsWhiteSpace code.[i] do
        i <- i + 1
    let start = i
    while i < code.Length && not (System.Char.IsWhiteSpace code.[i]) do
        i <- i + 1
    if i = start then "" else code.Substring(start, i - start)

let private fieldAtFormat (code: string) =
    // Read a "\@ \"fmt\"" date/time picture switch if present.
    let i = code.IndexOf("\\@")
    if i < 0 then None
    else
        let q1 = code.IndexOf('"', i)
        if q1 < 0 then None
        else
            let q2 = code.IndexOf('"', q1 + 1)
            if q2 <= q1 then None else Some (code.Substring(q1 + 1, q2 - q1 - 1))

/// What the running fields of a page show: its number as its section counts
/// it (w:pgNumType w:start) in the section's number format (w:fmt), the
/// number of pages of the document and of the page's section, and the
/// section's number (1-based).
type RunningFieldValues = {
    FieldPage: int
    FieldPageFormat: string
    FieldPageCount: int
    FieldSectionPages: int
    FieldSection: int
}

/// The values of a page that is the `pageNumber`-th of `pageCount` in a document of one section.
let runningFieldValues (pageNumber: int) (pageCount: int) : RunningFieldValues =
    { FieldPage = pageNumber; FieldPageFormat = "decimal"; FieldPageCount = pageCount; FieldSectionPages = pageCount; FieldSection = 1 }

/// A page number in a number format of w:pgNumType w:fmt (see Document.formatPageNumber).
let formatPageNumber (format: string) (number: int) : string =
    Rbt.Word.Document.formatPageNumber format number

let private fieldNumberFormat (code: string) : string option =
    Rbt.Word.Document.fieldNumberFormat code

let private formatCodeOf (format: string) =
    match format with
    | "upperRoman" -> 1
    | "lowerRoman" -> 2
    | "upperLetter" -> 3
    | "lowerLetter" -> 4
    | "numberInDash" -> 5
    | _ -> 0

/// What the values come to as one number (what a page's commands were made with, see the caches).
let private fieldValuesKey (values: RunningFieldValues) =
    (values.FieldSectionPages * 4096 + values.FieldSection) * 8 + formatCodeOf values.FieldPageFormat

let private runningFieldResult (values: RunningFieldValues) code currentText =
    let own = fieldNumberFormat code
    match fieldCodeName code with
    | "PAGE" -> formatPageNumber (own |> Option.defaultValue values.FieldPageFormat) values.FieldPage
    | "NUMPAGES" -> formatPageNumber (own |> Option.defaultValue "decimal") values.FieldPageCount
    | "SECTIONPAGES" -> formatPageNumber (own |> Option.defaultValue "decimal") values.FieldSectionPages
    | "SECTION" -> formatPageNumber (own |> Option.defaultValue "decimal") values.FieldSection
    | "DATE" -> System.DateTime.Now.ToString(fieldAtFormat code |> Option.defaultValue "yyyy-MM-dd")
    | "TIME" -> System.DateTime.Now.ToString(fieldAtFormat code |> Option.defaultValue "h:mm tt")
    | _ -> currentText

let private paragraphWithRunningFields (values: RunningFieldValues) (para: ParagraphBlock) =
    let mutable buffer = para.Buffer
    let mutable shift = 0
    let mutable changed = false
    for run in para.Buffer.Runs do
        match run with
        | Image _ -> ()
        | Text textRun ->
            match textRun.Style.FieldCode with
            | None -> ()
            | Some code ->
                let start = textRun.Start + shift
                let oldText = TextBuffer.slice start textRun.Length buffer
                let newText = runningFieldResult values code oldText
                if newText <> oldText then
                    let style = { textRun.Style with FieldResult = Some newText }
                    buffer <- buffer |> TextBuffer.delete start textRun.Length |> TextBuffer.insert start newText |> TextBuffer.applyStyle start newText.Length (fun _ -> style)
                    shift <- shift + newText.Length - textRun.Length
                    changed <- true
    if changed then { para with Buffer = buffer }, true else para, false

let private blockWithRunningFields (values: RunningFieldValues) = function
    | Paragraph para ->
        let para, changed = paragraphWithRunningFields values para
        if changed then Paragraph para, true else Paragraph para, false
    | Heading (level, para) ->
        let para, changed = paragraphWithRunningFields values para
        if changed then Heading (level, para), true else Heading (level, para), false
    | Table table ->
        let rec tableWithRunningFields depth (table: TableBlock) =
            if depth >= maxNestedTablePaintDepth then
                table, false
            else
                let rows, rowsChanged =
                    table.Rows
                    |> mapListWithChange (fun row ->
                        let cells, cellsChanged =
                            row.Cells
                            |> mapListWithChange (fun cell ->
                                let content, contentChanged =
                                    cell.Content
                                    |> mapListWithChange (paragraphWithRunningFields values)
                                let nestedTables, nestedChanged =
                                    cell.NestedTables
                                    |> mapListWithChange (tableWithRunningFields (depth + 1))
                                if contentChanged || nestedChanged then
                                    { cell with Content = content; NestedTables = nestedTables }, true
                                else
                                    cell, false)
                        if cellsChanged then { row with Cells = cells }, true else row, false)
                if rowsChanged then { table with Rows = rows }, true else table, false
        let table, changed = tableWithRunningFields 0 table
        if changed then Table table, true else Table table, false
    | TextBox textBox ->
        let content, changed =
            textBox.Content
            |> mapListWithChange (paragraphWithRunningFields values)
        if changed then TextBox { textBox with Content = content }, true else TextBox textBox, false

let private paragraphHasRunningFields (para: ParagraphBlock) =
    let mutable found = false
    for run in para.Buffer.Runs do
        if not found then
            match run with
            | Text textRun when textRun.Style.FieldCode.IsSome -> found <- true
            | _ -> ()
    found

let private blockHasRunningFields block =
    let blocks = Stack<Block>()
    blocks.Push block
    let mutable found = false
    while blocks.Count > 0 && not found do
        match blocks.Pop() with
        | Paragraph para
        | Heading (_, para) ->
            found <- paragraphHasRunningFields para
        | Table table ->
            for row in table.Rows do
                if not found then
                    for cell in row.Cells do
                        if not found then
                            for para in cell.Content do
                                if not found && paragraphHasRunningFields para then
                                    found <- true
                            for nestedTable in cell.NestedTables do
                                blocks.Push(Table nestedTable)
        | TextBox textBox ->
            for para in textBox.Content do
                if not found && paragraphHasRunningFields para then
                    found <- true
    found

let private blocksHaveRunningFields blocks =
    let mutable found = false
    for block in blocks do
        if not found && blockHasRunningFields block then
            found <- true
    found

type private HeaderFooterPartIndex = {
    Parts: HeaderFooterPart list
    ByRelationshipId: Dictionary<string, HeaderFooterPart>
    ByKind: Dictionary<string, HeaderFooterPart>
    First: HeaderFooterPart option
}

/// `referenced`: the relationship ids the document's sections refer to. A
/// part a section refers to is that section's (and of the sections that
/// follow it without one of their own); only a part no section refers to is
/// found by its kind, for every page.
let private buildHeaderFooterPartIndexFor (referenced: HashSet<string>) (parts: HeaderFooterPart list) =
    let byRelationshipId = Dictionary<string, HeaderFooterPart>()
    let byKind = Dictionary<string, HeaderFooterPart>()
    let mutable first = None
    for part in parts do
        let ofASection =
            match part.RelationshipId with
            | Some relationshipId -> referenced.Contains relationshipId
            | None -> false
        if first.IsNone && not ofASection then
            first <- Some part
        match part.RelationshipId with
        | Some relationshipId when not (byRelationshipId.ContainsKey relationshipId) ->
            byRelationshipId.[relationshipId] <- part
        | _ -> ()
        if not ofASection && not (byKind.ContainsKey part.Kind) then
            byKind.[part.Kind] <- part
    { Parts = parts; ByRelationshipId = byRelationshipId; ByKind = byKind; First = first }

let private tryFindPartByRelationshipId relationshipId (index: HeaderFooterPartIndex) =
    let mutable part = Unchecked.defaultof<HeaderFooterPart>
    if index.ByRelationshipId.TryGetValue(relationshipId, &part) then Some part else None

let private tryFindPartByKind kind (index: HeaderFooterPartIndex) =
    let mutable part = Unchecked.defaultof<HeaderFooterPart>
    if index.ByKind.TryGetValue(kind, &part) then Some part else None

let private partForReferences preferredKind (refs: (string * string) list) (index: HeaderFooterPartIndex) =
    let mutable found = None
    let mutable remaining = refs
    while found.IsNone && not remaining.IsEmpty do
        match remaining with
        | (kind, relationshipId) :: tail ->
            if kind = preferredKind then
                found <- tryFindPartByRelationshipId relationshipId index
            remaining <- tail
        | [] -> ()
    found

let private pageIndexForY (pages: PageInfo array) y =
    let target = y + 0.01
    let mutable lo = 0
    let mutable hi = pages.Length - 1
    let mutable idx = 0
    while lo <= hi do
        let mid = lo + (hi - lo) / 2
        if pages.[mid].Top <= target then
            idx <- mid
            lo <- mid + 1
        else
            hi <- mid - 1
    idx

type private SectionReferenceIndex = {
    EndPages: int array
    References: (string * string) list array
    /// The final section's references: for the pages after the last section end.
    Final: (string * string) list
}

/// Every section's header (or footer) references with the page its section
/// ends on. A section's sectPr sits on its LAST paragraph, so the references
/// there apply to the pages up to that paragraph; the final section's come
/// from the page setup. A section that names no reference of a kind uses the
/// previous section's (Word's "link to previous").
let private buildSectionReferenceIndex (pages: PageInfo array) (layouts: BlockLayout seq) (finalReferences: (string * string) list) isHeader =
    let endPages = ResizeArray<int>()
    let references = ResizeArray<(string * string) list>()
    let mutable effective : (string * string) list = []
    for layout in layouts do
        if endsSection layout.Metadata then
            let refs = if isHeader then layout.Metadata.HeaderReferences else layout.Metadata.FooterReferences
            effective <- inheritSectionReferences effective refs
            // (Also a section that shows no header: its pages are its own, not those of the section after it.)
            let endPage = pageIndexForY pages (layout.Y + layout.Height)
            let lastIndex = endPages.Count - 1
            if lastIndex < 0 || endPages.[lastIndex] <> endPage then
                endPages.Add endPage
                references.Add effective
    { EndPages = endPages.ToArray(); References = references.ToArray(); Final = inheritSectionReferences effective finalReferences }

let private referencedPartIds (index: SectionReferenceIndex) : HashSet<string> =
    let ids = HashSet<string>()
    for references in index.References do
        for (_, id) in references do ids.Add id |> ignore
    for (_, id) in index.Final do ids.Add id |> ignore
    ids

// For a target page we want the section with the smallest end-page that still
// contains it; a page after every section end belongs to the final section.
let private sectionRefsForPage (index: SectionReferenceIndex) pageIndex =
    let pages = index.EndPages
    let mutable lo = 0
    let mutable hi = pages.Length
    while lo < hi do
        let mid = lo + (hi - lo) / 2
        if pages.[mid] < pageIndex then
            lo <- mid + 1
        else
            hi <- mid
    if lo < pages.Length then index.References.[lo] else index.Final

/// The sections of a document by the pages they are on: where each one
/// ends and starts, and whether its first page has a header and footer of
/// its own (w:titlePg). The last entry is the final section.
type private SectionPageIndex = {
    SectionEndPages: int array
    SectionStartPages: int array
    SectionTitlePage: bool array
    SectionEvenOdd: bool array
    SectionHeaderDistance: float array
    SectionFooterDistance: float array
    /// The number a section's first page has when it states one (w:pgNumType w:start), and its number format.
    SectionNumberStart: int option array
    SectionNumberFormat: string array
}

/// `sectionEnds`: the layouts of the paragraphs that end a section, in order.
let private buildSectionPageIndex (pages: PageInfo array) (sectionEnds: BlockLayout seq) (pageSetup: PageSetup) : SectionPageIndex =
    let ends = ResizeArray<int>()
    let titles = ResizeArray<bool>()
    let evenOdd = ResizeArray<bool>()
    let headerDistances = ResizeArray<float>()
    let footerDistances = ResizeArray<float>()
    // (How a section starts is said by its own properties: a continuous one starts on the page the one before it ends on.)
    let continuous = ResizeArray<bool>()
    let kinds = ResizeArray<string option>()
    let numberStarts = ResizeArray<int option>()
    let numberFormats = ResizeArray<string>()
    let mutable lastId = ""
    for layout in sectionEnds do
        if layout.Metadata.SectionBreak.IsSome && layout.BlockId = lastId then
            // (A paragraph that runs on in another column or on another page: the section ends where its last part does.)
            ends.[ends.Count - 1] <- pageIndexForY pages (layout.Y + layout.Height)
        elif layout.Metadata.SectionBreak.IsSome then
            lastId <- layout.BlockId
            ends.Add (pageIndexForY pages (layout.Y + layout.Height))
            titles.Add (match layout.Metadata.SectionPageSetup with Some setup -> setup.DifferentFirstPage | None -> false)
            evenOdd.Add (match layout.Metadata.SectionPageSetup with Some setup -> setup.EvenAndOddHeaders | None -> false)
            headerDistances.Add (match layout.Metadata.SectionPageSetup with Some setup -> setup.HeaderDistance | None -> pageSetup.HeaderDistance)
            footerDistances.Add (match layout.Metadata.SectionPageSetup with Some setup -> setup.FooterDistance | None -> pageSetup.FooterDistance)
            continuous.Add (layout.Metadata.SectionBreak = Some "continuous")
            kinds.Add layout.Metadata.SectionBreak
            numberStarts.Add layout.Metadata.SectionPageNumberStart
            numberFormats.Add (layout.Metadata.SectionPageNumberFormat |> Option.defaultValue "decimal")
    ends.Add (max 0 (pages.Length - 1))
    numberStarts.Add pageSetup.PageNumberStart
    numberFormats.Add (pageSetup.PageNumberFormat |> Option.defaultValue "decimal")
    titles.Add pageSetup.DifferentFirstPage
    evenOdd.Add pageSetup.EvenAndOddHeaders
    headerDistances.Add pageSetup.HeaderDistance
    footerDistances.Add pageSetup.FooterDistance
    continuous.Add (pageSetup.SectionStart = Some "continuous")
    kinds.Add pageSetup.SectionStart
    let starts = Array.zeroCreate<int> ends.Count
    // (A section that starts on an even / odd page has a blank page of its own before its
    // first page where the next page is of the other kind: by the pages' numbers, counted as
    // the layout counts them - see Layout.layoutSections - and as the page fields show them.)
    let mutable number = (numberStarts.[0] |> Option.defaultValue 1) + ends.[0]
    let mutable owned = ends.[0]
    for index in 1 .. ends.Count - 1 do
        if continuous.[index] then
            starts.[index] <- ends.[index - 1]
            if ends.[index] > owned then
                number <- (numberStarts.[index] |> Option.defaultValue (number + 1)) + (ends.[index] - owned - 1)
                owned <- ends.[index]
        else
            let first = numberStarts.[index] |> Option.defaultValue (number + 1)
            let blank = Rbt.Word.WordLayout.Page.sectionStartsAfterBlankPage kinds.[index] first && ends.[index] > ends.[index - 1] + 1
            starts.[index] <- ends.[index - 1] + 1 + (if blank then 1 else 0)
            if ends.[index] > owned then
                number <- first + (ends.[index] - owned - 1)
                owned <- ends.[index]
    // (The last section reaches to the last page - the pages of the notes after it too.)
    if ends.Count > 1 && ends.[ends.Count - 1] < ends.[ends.Count - 2] then ends.[ends.Count - 1] <- ends.[ends.Count - 2]
    { SectionEndPages = ends.ToArray(); SectionStartPages = starts; SectionTitlePage = titles.ToArray(); SectionEvenOdd = evenOdd.ToArray()
      SectionHeaderDistance = headerDistances.ToArray(); SectionFooterDistance = footerDistances.ToArray()
      SectionNumberStart = numberStarts.ToArray(); SectionNumberFormat = numberFormats.ToArray() }

/// The section a page's header and footer are those of: the first that ends on or after the page.
let private sectionOfPage (index: SectionPageIndex) pageIndex =
    let ends = index.SectionEndPages
    let mutable lo = 0
    let mutable hi = ends.Length - 1
    while lo < hi do
        let mid = lo + (hi - lo) / 2
        if ends.[mid] < pageIndex then lo <- mid + 1 else hi <- mid
    lo

/// What the running fields show page by page: the pages are numbered on from
/// 1, a section that states a start (w:pgNumType w:start) from that on its
/// first page; a section's pages are those from its first to its last.
let private pageFieldValues (index: SectionPageIndex) (pageTotal: int) : RunningFieldValues array =
    let values = Array.zeroCreate<RunningFieldValues> pageTotal
    let mutable number = 0
    let mutable previous = -1
    for pageIndex in 0 .. pageTotal - 1 do
        let section = sectionOfPage index pageIndex
        number <-
            match index.SectionNumberStart.[section] with
            | Some start when section <> previous -> start
            | _ -> number + 1
        previous <- section
        // (A continuous section starts on the page the one before it ends on: that page is the earlier one's.)
        // (A blank page before a section that starts on an odd / even page is that section's.)
        let first = if section > 0 then index.SectionEndPages.[section - 1] + 1 else 0
        values.[pageIndex] <-
            { FieldPage = number
              FieldPageFormat = index.SectionNumberFormat.[section]
              FieldPageCount = pageTotal
              FieldSectionPages = max 1 (index.SectionEndPages.[section] - first + 1)
              FieldSection = section + 1 }
    values

/// Which header / footer a page shows: "first" on the first page of a
/// section that has a title page, "even" on even pages when the document
/// tells even and odd pages apart, else "default".
let private headerFooterKindOfPage (index: SectionPageIndex) (pageSetup: PageSetup) pageIndex =
    let section = sectionOfPage index pageIndex
    if index.SectionTitlePage.[section] && pageIndex = index.SectionStartPages.[section] then "first"
    elif (pageIndex + 1) % 2 = 0 && (pageSetup.EvenAndOddHeaders || index.SectionEvenOdd.[section]) then "even"
    else "default"

let private partForPage (preferredKind: string) pageIndex (parts: HeaderFooterPartIndex) sectionRefs =
    let refs = sectionRefsForPage sectionRefs pageIndex
    // A title page or an even page without a header of that kind - its own or one of a section
    // before it - shows none (ECMA-376 17.10.2: "a new blank header"), not the default header.
    // (A part no section refers to - one made without a reference - is found by its kind.)
    partForReferences preferredKind refs parts
    |> Option.orElseWith (fun () -> tryFindPartByKind preferredKind parts)
    |> Option.orElseWith (fun () -> if preferredKind = "default" then parts.First else None)

let private blocksForPage (preferredKind: string) pageIndex fallbackBlocks (parts: HeaderFooterPartIndex) sectionRefs =
    if parts.Parts.IsEmpty then
        fallbackBlocks ()
    else
        match partForPage preferredKind pageIndex parts sectionRefs with
        | Some part -> part.Blocks
        | None -> []

let rec private shiftCommandY dy command =
    match command with
    | DrawText (t, s, x, b, w) -> DrawText (t, s, x, b + dy, w)
    | DrawUnderline (c, s, x, y, w) -> DrawUnderline (c, s, x, y + dy, w)
    | DrawImage (s, x, y, w, h, cl, ct, cr, cb) -> DrawImage (s, x, y + dy, w, h, cl, ct, cr, cb)
    | FillRect (c, x, y, w, h) -> FillRect (c, x, y + dy, w, h)
    | StrokeRect (c, lw, x, y, w, h) -> StrokeRect (c, lw, x, y + dy, w, h)
    | StrokeLine (c, lw, ds, doff, x1, y1, x2, y2) -> StrokeLine (c, lw, ds, doff, x1, y1 + dy, x2, y2 + dy)
    | DrawShape (fill, stroke, sw, (bx, by, bw, bh), segs) ->
        DrawShape (fill, stroke, sw, (bx, by + dy, bw, bh), Shapes.translatePath 0.0 dy segs)
    | DrawGroup (clip, rotation, cx, cy, opacity, (bx, by, bw, bh), commands) ->
        DrawGroup (clip |> Option.map (fun (x, y, w, h) -> (x, y + dy, w, h)), rotation, cx, cy + dy, opacity, (bx, by + dy, bw, bh), commands |> List.map (shiftCommandY dy))
    | DrawMirrored (flipH, flipV, cx, cy, (bx, by, bw, bh), commands) ->
        DrawMirrored (flipH, flipV, cx, cy + dy, (bx, by + dy, bw, bh), commands |> List.map (shiftCommandY dy))

let private shiftCommandsY dy commands =
    if abs dy < 0.001 then
        commands
    else
        let shifted = ResizeArray<PaintCommand>()
        for command in commands do
            shifted.Add(shiftCommandY dy command)
        resizeArrayToList shifted

/// The document a header's or footer's table is laid out in: the table alone, with the document's styles.
let private storyTableDocument (doc: Document) (table: TableBlock) : Document =
    { Rbt.Word.Document.empty with Blocks = [ Table table ]; Selection = None; Styles = doc.Styles }

/// A shape's fill and outline, its box at `x`, `y` and `height` high (not
/// turned) - mirrored in its box when the shape is flipped (its gradient too).
let private shapeFrameCommands (textBox: TextBoxBlock) (x: float) (y: float) (height: float) : PaintCommand list =
    let commands = ResizeArray<PaintCommand>()
    let borderWidth = textBox.BorderWidth |> Option.defaultValue 1.0
    let useShape = (match textBox.Geometry with GeoRect -> false | _ -> true) || textBox.Gradient.IsSome || textBox.BorderDash.IsSome
    if useShape then
        let segs =
            Shapes.geometryPath textBox.Geometry textBox.Width height
            |> Shapes.flipPath textBox.FlipH textBox.FlipV textBox.Width height
            |> Shapes.translatePath x y
        let fill =
            match textBox.Gradient with
            | Some g when (textBox.FlipH || textBox.FlipV) && not g.Radial ->
                // (The direction of a linear gradient, mirrored as the shape is.)
                let angle = if textBox.FlipH then 180.0 - g.Angle else g.Angle
                let angle = if textBox.FlipV then -angle else angle
                FillGradient { g with Angle = ((angle % 360.0) + 360.0) % 360.0 }
            | Some g -> FillGradient g
            | None -> match textBox.FillColor with Some c -> FillSolid c | None -> FillNone
        if fill <> FillNone || textBox.BorderColor.IsSome then
            commands.Add(DrawShape(fill, textBox.BorderColor, borderWidth, (x, y, textBox.Width, height), segs))
    else
        textBox.FillColor |> Option.iter (fun color -> commands.Add(FillRect (color, x, y, textBox.Width, height)))
        textBox.BorderColor |> Option.iter (fun color -> commands.Add(StrokeRect (color, borderWidth, x, y, textBox.Width, height)))
    resizeArrayToList commands

/// Text drawn as large as its shape (WordArt, a watermark): one line in the
/// face and colour of its first run, as wide as the shape or as high.
let private fittedTextCommands (fontFamily: string) (measureText: TextMeasurer) (textBox: TextBoxBlock) (x: float) (y: float) (height: float) : PaintCommand list =
    let text = textBox.Content |> List.map (fun para -> TextBuffer.toString para.Buffer) |> String.concat " "
    if System.String.IsNullOrWhiteSpace text || textBox.Width <= 0.0 || height <= 0.0 then []
    else
        let style =
            textBox.Content
            |> List.tryPick (fun para -> para.Buffer.Runs |> List.tryPick (function TextBuffer.Text run -> Some run.Style | _ -> None))
            |> Option.defaultValue TextBuffer.defaultStyle
        let probe = { style with FontSize = 100.0 }
        let natural = measureText text probe
        if natural <= 0.0 then []
        else
            // (A line of a face is about 1.2 times its size high.)
            let size = min (100.0 * textBox.Width / natural) (height / 1.2)
            let fitted = { style with FontSize = size }
            let width = measureText text fitted
            let paintStyle = { styleFromTextStyle fitted with FontFamily = (styleFromTextStyle fitted).FontFamily |> Option.orElse (Some fontFamily) }
            [ DrawText (text, paintStyle, x + (textBox.Width - width) / 2.0, y + height / 2.0 + size * 0.35, width) ]

/// The shapes of a group or a drawing canvas as they are drawn in the group's
/// box (`height` high): each with its left edge and top from the group's
/// top-left corner and its height - scaled as the group is from the size its
/// shapes' places are given in (ChildSpace; a canvas's are not), and mirrored
/// when the group is flipped: a shape's place across (upside down), its own
/// flip the other way, its turn the other way round (a shape mirrored after
/// it is turned is the mirrored shape turned back). Its text stays upright,
/// as a flipped shape's does.
let groupChildBoxes (group: TextBoxBlock) (height: float) : (TextBoxBlock * (float * float * float)) list =
    let scaleX, scaleY =
        match group.ChildSpace with
        | Some (width, spaceHeight) when width > 0.0 && spaceHeight > 0.0 -> group.Width / width, height / spaceHeight
        | _ -> 1.0, 1.0
    let mirrored = group.FlipH <> group.FlipV
    group.Children
    |> List.map (fun child ->
        let width, childHeight = child.Width * scaleX, child.Height * scaleY
        let left = (child.X |> Option.defaultValue 0.0) * scaleX
        let top = (child.Y |> Option.defaultValue 0.0) * scaleY
        let left = if group.FlipH then group.Width - left - width else left
        let top = if group.FlipV then height - top - childHeight else top
        let placed =
            if not group.FlipH && not group.FlipV && scaleX = 1.0 && scaleY = 1.0 then child
            else
                { child with
                    Width = width
                    Height = childHeight
                    FlipH = (child.FlipH <> group.FlipH)
                    FlipV = (child.FlipV <> group.FlipV)
                    Rotation = (if mirrored then -child.Rotation else child.Rotation)
                    // (A group in it scales its own shapes from the size it is drawn at.)
                    ChildSpace = child.ChildSpace }
        placed, (left, top, childHeight))

/// A text box or shape, its box at `x`, `y` (across and down the canvas) and
/// `height` high: its fill and outline, and its text inside its insets -
/// from the top, the middle or the bottom of the shape -, clipped to the
/// shape when the shape has a height of its own; all of it turned when the
/// shape is. (The text stood 6 pixels inside the box whatever the shape said,
/// ran out of the box below, and stayed upright in a turned shape.)
///
/// A group or a drawing canvas (Children): its shapes - each drawn as a shape
/// is, with its text - where the group's child coordinates put them (see
/// groupChildBoxes), over the canvas's background; all of it turned when the
/// group is. A picture of a group (ImageSource): the picture in its box,
/// mirrored and turned as it says.
let rec private textBoxCommandsAt (settings: LayoutSettings) fontFamily measureText measureRange (styleContext: DocumentStyleContext) (textBox: TextBoxBlock) (x: float) (y: float) (height: float) : PaintCommand list =
    let frame =
        match textBox.ImageSource with
        | Some src ->
            let picture = DrawImage (src, x, y, textBox.Width, height, None, None, None, None)
            if textBox.FlipH || textBox.FlipV then [ DrawMirrored (textBox.FlipH, textBox.FlipV, x + textBox.Width / 2.0, y + height / 2.0, (x, y, textBox.Width, height), [ picture ]) ]
            else [ picture ]
        | None -> shapeFrameCommands textBox x y height
    let text = ResizeArray<PaintCommand>()
    if not textBox.Children.IsEmpty then
        for child, (childX, childY, childHeight) in groupChildBoxes textBox height do
            text.AddRange (textBoxCommandsAt settings fontFamily measureText measureRange styleContext child (x + childX) (y + childY) childHeight)
    elif textBox.FitText then text.AddRange (fittedTextCommands fontFamily measureText textBox x y height)
    else
        // (A picture anchored in the box's text is part of the box.)
        let outerSink = drawingSink
        drawingSink <- null
        try
            for layout in textBoxContentLayouts settings measureText measureRange styleContext textBox x y height do
                addParagraphCommands settings fontFamily measureText measureRange text layout
        finally drawingSink <- outerSink
    let turned = abs textBox.Rotation >= 0.01
    let clipped = text.Count > 0 && not textBox.AutoFit && not textBox.NoTextWrap && not textBox.FitText && textBox.Children.IsEmpty
    let translucent = textBox.Opacity < 0.999
    if not turned && not clipped && not translucent then frame @ resizeArrayToList text
    else
        let pad = (textBox.BorderWidth |> Option.defaultValue 1.0) + 1.0
        let centerX, centerY = x + textBox.Width / 2.0, y + height / 2.0
        let bbox =
            if turned then
                let radius = sqrt (textBox.Width * textBox.Width + height * height) / 2.0 + pad
                (centerX - radius, centerY - radius, radius * 2.0, radius * 2.0)
            else (x - pad, y - pad, textBox.Width + pad * 2.0, height + pad * 2.0)
        let inner =
            if clipped then frame @ [ DrawGroup (Some (x, y, textBox.Width, height), 0.0, centerX, centerY, 1.0, (x, y, textBox.Width, height), resizeArrayToList text) ]
            else frame @ resizeArrayToList text
        if not turned && not translucent then inner
        else [ DrawGroup (None, textBox.Rotation, centerX, centerY, textBox.Opacity, bbox, inner) ]

/// A shape where its anchor puts it (see `textBoxPlace`), as a box: its left
/// edge from the left edge of the text, the left margin of its page, and its height.
let private anchoredShapeBox (settings: LayoutSettings) measureText measureRange (styleContext: DocumentStyleContext) (pages: PageInfo array) (top: float) (textBox: TextBoxBlock) : struct (WordLayout.Floating.Box * float * float) =
    let height = textBoxFrameHeight settings measureText measureRange styleContext textBox
    let info = pages.[pageIndexForY pages top]
    let page : WordLayout.Floating.AnchorPage =
        { PageWidth = info.Width; PageHeight = info.Height; MarginLeft = info.MarginLeft; MarginRight = info.MarginRight; MarginTop = info.MarginTop; MarginBottom = info.MarginBottom }
    struct (WordLayout.Floating.anchoredTextBoxBox page info.Top top height textBox, info.MarginLeft, height)

/// Where a shape is: its left edge and top (across and down the canvas) and
/// its height. `anchorTop`: the top of the block it is anchored in, as laid
/// out (None for a shape without an anchor, or whose anchor has no layout:
/// by its offsets from the first page's margins, as before). `shift`: how far
/// down it is moved to keep it clear of another shape (see `overlapShifts`).
let private textBoxPlace (settings: LayoutSettings) measureText measureRange (styleContext: DocumentStyleContext) (pages: PageInfo array) (anchorTop: float option) (shift: float) (textBox: TextBoxBlock) : struct (float * float * float) =
    match anchorTop with
    | Some top when pages.Length > 0 ->
        let struct (box, marginLeft, height) = anchoredShapeBox settings measureText measureRange styleContext pages top textBox
        struct (marginLeft + box.X, box.Y + shift, height)
    | _ ->
        let height = textBoxFrameHeight settings measureText measureRange styleContext textBox
        let box = WordLayout.Floating.textBoxCanvasBox settings.PaddingLeft settings.PaddingTop textBox
        struct (box.X, box.Y, height)

/// How far down the shapes that may not overlap (allowOverlap="0") are moved
/// to keep clear of those before them, as the layout moves them (see
/// WordLayout.Floating.keptClearOfShapes): by id, for the shapes that are
/// moved. `shapes`: in the order they are placed (that of the document);
/// `anchorTopOf`: where the block a shape is anchored in starts.
let private overlapShifts (settings: LayoutSettings) measureText measureRange (styleContext: DocumentStyleContext) (pages: PageInfo array) (anchorTopOf: TextBoxBlock -> float option) (shapes: TextBoxBlock seq) : Dictionary<BlockId, float> =
    let shifts = Dictionary<BlockId, float>()
    let keptClear = shapes |> Seq.filter (fun textBox -> textBox.AllowOverlap = Some false && not textBox.Inline) |> Seq.toList
    if keptClear.Length > 1 && pages.Length > 0 then
        let mutable placed : WordLayout.Floating.Box list = []
        for textBox in keptClear do
            match anchorTopOf textBox with
            | Some top ->
                let struct (box, _, _) = anchoredShapeBox settings measureText measureRange styleContext pages top textBox
                let kept = WordLayout.Floating.keptClearOfShapes placed box
                if abs (kept.Y - box.Y) > 0.001 then shifts.[textBox.Id] <- kept.Y - box.Y
                placed <- kept :: placed
            | None -> ()
    shifts

let private shiftOf (shifts: Dictionary<BlockId, float>) (textBox: TextBoxBlock) =
    match shifts.TryGetValue textBox.Id with
    | true, shift -> shift
    | false, _ -> 0.0

/// A text box as a drawing of its page.
let private textBoxDrawing (settings: LayoutSettings) fontFamily measureText measureRange (styleContext: DocumentStyleContext) (pages: PageInfo array) (anchorTop: float option) (shift: float) (textBox: TextBoxBlock) : Drawing =
    let struct (x, y, height) = textBoxPlace settings measureText measureRange styleContext pages anchorTop shift textBox
    { DrawingBehind = (textBox.BehindDoc = Some true)
      DrawingZ = drawingHeight textBox.RelativeHeight
      DrawingCommands = textBoxCommandsAt settings fontFamily measureText measureRange styleContext textBox x y height |> List.toArray }

/// The paragraph whose layout says where a block starts.
let rec private firstParagraphIdOf (block: Block) : BlockId option =
    match block with
    | Paragraph para
    | Heading (_, para) -> Some para.Id
    | Table table ->
        table.Rows |> List.tryPick (fun row -> row.Cells |> List.tryPick (fun cell -> cell.Content |> List.tryHead |> Option.map (fun para -> para.Id)))
    | TextBox _ -> None

let private repeatedBlockCommands (settings: LayoutSettings) fontFamily measureText measureRange (doc: Document) (values: RunningFieldValues) (page: PageInfo option) startY blocks =
    // (A header's paragraphs have the document's styles - "Header", the default paragraph
    // style - as the body's have: laid out without them they came out in the built-in 11 pt.)
    let styleContext = WordStyles.styleContext doc
    let repeatedSettings = { settings with MarginBottom = 0.0 }
    let maxWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
    let mutable y = startY
    // (The header's own drawings below and above its text: see Drawing.)
    let commands = ResizeArray<PaintCommand>()
    let drawings = ResizeArray<Drawing>()
    let outerSink = drawingSink
    drawingSink <- drawings
    // (Where the blocks start, for the shapes anchored in them.)
    let tops = Dictionary<BlockId, float>()
    try
        for rawBlock in blocks do
            let block, _ = blockWithRunningFields values rawBlock
            (match block with
             | Paragraph para
             | Heading (_, para) -> tops.[para.Id] <- y
             | Table table -> tops.[table.Id] <- y
             | TextBox _ -> ())
            match block with
            | Paragraph para
            | Heading (_, para) ->
                let safePara = { para with Metadata = { para.Metadata with PageBreakAfter = false } }
                let layout, _ = layoutParagraphAtWithStyleWithRange repeatedSettings measureText measureRange maxWidth y 0.0 (forBlockInContext styleContext block) (effectiveBufferInContext styleContext) safePara Map.empty
                addParagraphCommands settings fontFamily measureText measureRange commands layout
                y <- layout.Y + layout.Height
            | Table table ->
                let tableCommands, nextY = tableBorderCommandsCore (Some styleContext) repeatedSettings measureText measureRange table y
                commands.AddRange(tableCommands)
                // The borders alone are not the table: lay its cells out as well
                // and paint their text. The table is laid out at the body origin
                // on an unpaged measuring page (a header/footer table never
                // splits) and the resulting commands are moved to y.
                let measureSettings = { repeatedSettings with PageHeight = max repeatedSettings.PageHeight 1000000.0 }
                let tableDoc = storyTableDocument doc table
                let cellCommands = ResizeArray<PaintCommand>()
                for layout in layoutDocumentWithRange measureSettings measureText measureRange tableDoc do
                    if not (layout.BlockId.StartsWith footnoteBlockIdPrefix) && not (layout.BlockId.StartsWith endnoteBlockIdPrefix) then
                        addParagraphCommands settings fontFamily measureText measureRange cellCommands layout
                commands.AddRange(shiftCommandsY (y - measureSettings.PaddingTop) (resizeArrayToList cellCommands))
                y <- nextY
            | TextBox textBox ->
                // A shape of the header: where its anchor puts it on the page (a watermark in the middle of it).
                let anchorTop =
                    match textBox.AnchorBlockId with
                    | Some anchorId ->
                        match tops.TryGetValue anchorId with
                        | true, top -> Some top
                        | false, _ -> None
                    | None -> None
                match page, anchorTop with
                | Some info, Some _ -> drawings.Add (textBoxDrawing settings fontFamily measureText measureRange styleContext [| info |] anchorTop 0.0 textBox)
                | _ -> ()
    finally drawingSink <- outerSink
    resizeArrayToList (commandsInLayers commands drawings)

/// The header and footer commands of one page, and what they were made from.
type private HeaderFooterPageEntry = {
    PageHeaders: Block list
    PageFooters: Block list
    PageHeaderY: float
    PageFooterY: float
    /// The page number and count the commands were made with; 0 when the
    /// blocks hold no running field (PAGE, NUMPAGES), so neither matters.
    PageNumber: int
    PageCount: int
    /// What else the running fields show (see fieldValuesKey); 0 without running fields.
    PageFieldsKey: int
    PageCommands: PaintCommand array
}

/// What is kept of the headers and footers from one paint to the next (see
/// `PaintCache`): the commands of each page, the blocks made of a document's
/// plain header / footer paragraphs, and which block lists hold running fields.
type private HeaderFooterCache = {
    HeaderFooterPages: Dictionary<int, HeaderFooterPageEntry>
    mutable FallbackHeaderSource: ParagraphBlock list
    mutable FallbackHeaderBlocks: Block list
    mutable FallbackFooterSource: ParagraphBlock list
    mutable FallbackFooterBlocks: Block list
    /// Block lists (the few a document's pages use, by reference) and whether they hold running fields.
    RunningFields: ResizeArray<struct (Block list * bool)>
}

let private createHeaderFooterCache () = {
    HeaderFooterPages = Dictionary<int, HeaderFooterPageEntry>()
    FallbackHeaderSource = []
    FallbackHeaderBlocks = []
    FallbackFooterSource = []
    FallbackFooterBlocks = []
    RunningFields = ResizeArray<struct (Block list * bool)>()
}

/// What a page's header and footer are made from: the blocks, where they
/// stand, and whether they show the page number or the page count.
[<Struct>]
type private HeaderFooterPageSource = {
    SourceHeaders: Block list
    SourceFooters: Block list
    SourceHeaderY: float
    SourceFooterY: float
    SourceRunning: bool
}

/// The headers and footers of a document's pages: how many pages, what each
/// page's are made from, and the commands of a page (made when asked for).
let private headerFooterPageSources (cache: HeaderFooterCache option) (settings: LayoutSettings) fontFamily measureText measureRange (doc: Document) (layouts: BlockLayout seq) (pages: PageInfo list)
        : int * (int -> HeaderFooterPageSource) * (int -> HeaderFooterPageSource -> PaintCommand array) * (int -> RunningFieldValues) =
    let pageCount = pages.Length
    if doc.Headers.IsEmpty && doc.Footers.IsEmpty && doc.HeaderParts.IsEmpty && doc.FooterParts.IsEmpty then
        0, (fun _ -> { SourceHeaders = []; SourceFooters = []; SourceHeaderY = 0.0; SourceFooterY = 0.0; SourceRunning = false }), (fun _ _ -> [||]), (fun pageIndex -> runningFieldValues (pageIndex + 1) pageCount)
    else
        let mutable fallbackHeaders = ValueNone
        let mutable fallbackFooters = ValueNone
        let fallbackHeaderBlocks () =
            match fallbackHeaders with
            | ValueSome blocks -> blocks
            | ValueNone ->
                let blocks =
                    match cache with
                    | Some cache when obj.ReferenceEquals(cache.FallbackHeaderSource, doc.Headers) && not doc.Headers.IsEmpty -> cache.FallbackHeaderBlocks
                    | _ ->
                        let blocks = doc.Headers |> List.map Paragraph
                        match cache with
                        | Some cache ->
                            cache.FallbackHeaderSource <- doc.Headers
                            cache.FallbackHeaderBlocks <- blocks
                        | None -> ()
                        blocks
                fallbackHeaders <- ValueSome blocks
                blocks
        let fallbackFooterBlocks () =
            match fallbackFooters with
            | ValueSome blocks -> blocks
            | ValueNone ->
                let blocks =
                    match cache with
                    | Some cache when obj.ReferenceEquals(cache.FallbackFooterSource, doc.Footers) && not doc.Footers.IsEmpty -> cache.FallbackFooterBlocks
                    | _ ->
                        let blocks = doc.Footers |> List.map Paragraph
                        match cache with
                        | Some cache ->
                            cache.FallbackFooterSource <- doc.Footers
                            cache.FallbackFooterBlocks <- blocks
                        | None -> ()
                        blocks
                fallbackFooters <- ValueSome blocks
                blocks
        let pagesArray = List.toArray pages
        let headerRefs =
            if doc.Headers.IsEmpty && doc.HeaderParts.IsEmpty then { EndPages = [||]; References = [||]; Final = [] }
            else buildSectionReferenceIndex pagesArray layouts doc.PageSetup.HeaderReferences true
        let footerRefs =
            if doc.Footers.IsEmpty && doc.FooterParts.IsEmpty then { EndPages = [||]; References = [||]; Final = [] }
            else buildSectionReferenceIndex pagesArray layouts doc.PageSetup.FooterReferences false
        let headerParts = buildHeaderFooterPartIndexFor (referencedPartIds headerRefs) doc.HeaderParts
        let footerParts = buildHeaderFooterPartIndexFor (referencedPartIds footerRefs) doc.FooterParts
        // (Looked up by reference in short lists: a document's pages share a handful of block lists.)
        let runningFields =
            match cache with
            | Some cache -> cache.RunningFields
            | None -> ResizeArray<struct (Block list * bool)>()
        let hasRunningFields (blocks: Block list) =
            if blocks.IsEmpty then false
            else
                let mutable at = -1
                let mutable i = 0
                while at < 0 && i < runningFields.Count do
                    let struct (known, _) = runningFields.[i]
                    if obj.ReferenceEquals(known, blocks) then at <- i
                    i <- i + 1
                if at >= 0 then
                    let struct (_, found) = runningFields.[at]
                    found
                else
                    if runningFields.Count > 64 then runningFields.Clear()
                    let found = blocksHaveRunningFields blocks
                    runningFields.Add(struct (blocks, found))
                    found
        let staticRepeatedCommands = ResizeArray<struct (Block list * PaintCommand list)>()
        let repeatedCommands (values: RunningFieldValues) (info: PageInfo) startY (blocks: Block list) =
            // (A shape is placed on its page: nothing of it is kept from page to page.)
            if hasRunningFields blocks || (blocks |> List.exists (function TextBox _ -> true | _ -> false)) then
                repeatedBlockCommands settings fontFamily measureText measureRange doc values (Some info) startY blocks
            else
                let mutable originCommands = Unchecked.defaultof<PaintCommand list>
                let mutable found = false
                let mutable i = 0
                while not found && i < staticRepeatedCommands.Count do
                    let struct (known, commands) = staticRepeatedCommands.[i]
                    if obj.ReferenceEquals(known, blocks) then
                        originCommands <- commands
                        found <- true
                    i <- i + 1
                if not found then
                    originCommands <- repeatedBlockCommands settings fontFamily measureText measureRange doc values None 0.0 blocks
                    staticRepeatedCommands.Add(struct (blocks, originCommands))
                shiftCommandsY startY originCommands
        let sectionPages = buildSectionPageIndex pagesArray layouts doc.PageSetup
        let mutable fieldValues : RunningFieldValues array = null
        let valuesOf (pageIndex: int) : RunningFieldValues =
            if isNull fieldValues then fieldValues <- pageFieldValues sectionPages pagesArray.Length
            fieldValues.[pageIndex]
        let sourceOf (pageIndex: int) : HeaderFooterPageSource =
            let info = pagesArray.[pageIndex]
            let kind = headerFooterKindOfPage sectionPages doc.PageSetup pageIndex
            let headers = blocksForPage kind pageIndex fallbackHeaderBlocks headerParts headerRefs
            let footers = blocksForPage kind pageIndex fallbackFooterBlocks footerParts footerRefs
            let section = sectionOfPage sectionPages pageIndex
            { SourceHeaders = headers; SourceFooters = footers
              SourceHeaderY = info.Top + sectionPages.SectionHeaderDistance.[section]
              SourceFooterY = info.Top + info.Height - sectionPages.SectionFooterDistance.[section] - settings.LineHeight * 0.78
              SourceRunning = hasRunningFields headers || hasRunningFields footers }
        let commandsOf (pageIndex: int) (source: HeaderFooterPageSource) : PaintCommand array =
            let values = valuesOf pageIndex
            let commands = ResizeArray<PaintCommand>()
            let info = pagesArray.[pageIndex]
            if not source.SourceHeaders.IsEmpty then commands.AddRange(repeatedCommands values info source.SourceHeaderY source.SourceHeaders)
            if not source.SourceFooters.IsEmpty then commands.AddRange(repeatedCommands values info source.SourceFooterY source.SourceFooters)
            commands.ToArray()
        pagesArray.Length, sourceOf, commandsOf, valuesOf

/// The header and footer commands page by page. With a cache, a page whose
/// header and footer blocks, place and (where they show it) page number and
/// page count are what they were keeps its commands - the same objects.
let private headerFooterCommandsByPage (cache: HeaderFooterCache option) (settings: LayoutSettings) fontFamily measureText measureRange (doc: Document) (layouts: BlockLayout seq) (pages: PageInfo list) : PaintCommand array array =
    let pageTotal, sourceOf, commandsOf, valuesOf = headerFooterPageSources cache settings fontFamily measureText measureRange doc layouts pages
    if pageTotal = 0 then [||]
    else
        let pageCount = pageTotal
        let result = Array.zeroCreate<PaintCommand array> pageTotal
        for pageIndex = 0 to pageTotal - 1 do
            let source = sourceOf pageIndex
            let pageNumber = if source.SourceRunning then (valuesOf pageIndex).FieldPage else 0
            let fieldsKey = if source.SourceRunning then fieldValuesKey (valuesOf pageIndex) else 0
            let headers, footers, headerY, footerY, running = source.SourceHeaders, source.SourceFooters, source.SourceHeaderY, source.SourceFooterY, source.SourceRunning
            let known =
                match cache with
                | Some cache ->
                    match cache.HeaderFooterPages.TryGetValue pageIndex with
                    | true, entry
                        when obj.ReferenceEquals(entry.PageHeaders, headers) && obj.ReferenceEquals(entry.PageFooters, footers)
                             && entry.PageHeaderY = headerY && entry.PageFooterY = footerY
                             && entry.PageNumber = pageNumber && entry.PageCount = (if running then pageCount else 0) && entry.PageFieldsKey = fieldsKey -> Some entry.PageCommands
                    | _ -> None
                | None -> None
            match known with
            | Some commands -> result.[pageIndex] <- commands
            | None ->
                let commands = commandsOf pageIndex source
                result.[pageIndex] <- commands
                match cache with
                | Some cache ->
                    cache.HeaderFooterPages.[pageIndex] <- {
                        PageHeaders = headers; PageFooters = footers; PageHeaderY = headerY; PageFooterY = footerY
                        PageNumber = pageNumber; PageCount = (if running then pageCount else 0); PageFieldsKey = fieldsKey; PageCommands = commands }
                | None -> ()
        let pagesArray = result
        match cache with
        | Some cache ->
            // Pages the document no longer has.
            if cache.HeaderFooterPages.Count > pagesArray.Length then
                let stale = ResizeArray<int>()
                for KeyValue (pageIndex, _) in cache.HeaderFooterPages do
                    if pageIndex >= pagesArray.Length then stale.Add pageIndex
                for pageIndex in stale do
                    cache.HeaderFooterPages.Remove pageIndex |> ignore
        | None -> ()
        result

let private headerFooterCommands (settings: LayoutSettings) fontFamily measureText measureRange (doc: Document) layouts (pages: PageInfo list) =
    let commands = ResizeArray<PaintCommand>()
    for pageCommands in headerFooterCommandsByPage None settings fontFamily measureText measureRange doc layouts pages do
        commands.AddRange pageCommands
    resizeArrayToList commands

// ---- a page's header or footer as a story of its own (edited in place) ------

/// Where a page's header and footer blocks start.
let headerFooterStartY (settings: LayoutSettings) (pageSetup: PageSetup) (page: PageInfo) (isHeader: bool) =
    if isHeader then page.Top + pageSetup.HeaderDistance
    else page.Top + page.Height - pageSetup.FooterDistance - settings.LineHeight * 0.78

/// The header (or footer) a page shows, for editing it in place.
type HeaderFooterTarget = {
    TargetIsHeader: bool
    TargetPageIndex: int
    /// "default" | "first" | "even": which of its section's headers the page shows.
    TargetKind: string
    /// The section the page's header is that of (0-based), and how many sections there are.
    TargetSection: int
    TargetSectionCount: int
    /// The first page of that section.
    TargetSectionStartPage: int
    /// The part whose blocks the page shows; None when it shows none (or the
    /// document's plain header paragraphs, which are not a part).
    TargetPart: HeaderFooterPart option
    /// Whether the section names a part of this kind itself (false: it shows
    /// the one of the section before it - "link to previous" - or none).
    TargetOwn: bool
    /// The references of the section, as it states them (kind, relationship id).
    TargetReferences: (string * string) list
    /// Where the blocks start on the page.
    TargetY: float
    /// What the running fields show on the page.
    TargetFields: RunningFieldValues
}

/// `sectionEnds`: the layouts of the paragraphs that end a section, in order.
let headerFooterTargetForPage (settings: LayoutSettings) (doc: Document) (sectionEnds: BlockLayout seq) (pages: PageInfo list) (pageIndex: int) (isHeader: bool) : HeaderFooterTarget option =
    let pagesArray = List.toArray pages
    if pageIndex < 0 || pageIndex >= pagesArray.Length then None
    else
        let sectionPages = buildSectionPageIndex pagesArray sectionEnds doc.PageSetup
        let kind = headerFooterKindOfPage sectionPages doc.PageSetup pageIndex
        let section = sectionOfPage sectionPages pageIndex
        let refs = buildSectionReferenceIndex pagesArray sectionEnds (if isHeader then doc.PageSetup.HeaderReferences else doc.PageSetup.FooterReferences) isHeader
        let parts = buildHeaderFooterPartIndexFor (referencedPartIds refs) (if isHeader then doc.HeaderParts else doc.FooterParts)
        // What the section itself states (the n-th paragraph that ends a section, or the page setup for the last).
        let own =
            let mutable found : (string * string) list option = None
            let mutable ordinal = 0
            let mutable lastId = ""
            for layout in sectionEnds do
                if layout.Metadata.SectionBreak.IsSome && layout.BlockId <> lastId then
                    lastId <- layout.BlockId
                    if ordinal = section && found.IsNone then
                        found <- Some (if isHeader then layout.Metadata.HeaderReferences else layout.Metadata.FooterReferences)
                    ordinal <- ordinal + 1
            match found with
            | Some references -> references
            | None -> if isHeader then doc.PageSetup.HeaderReferences else doc.PageSetup.FooterReferences
        Some {
            TargetIsHeader = isHeader
            TargetPageIndex = pageIndex
            TargetKind = kind
            TargetSection = section
            TargetSectionCount = sectionPages.SectionEndPages.Length
            TargetSectionStartPage = sectionPages.SectionStartPages.[section]
            TargetPart = if parts.Parts.IsEmpty then None else partForPage kind pageIndex parts refs
            TargetOwn = own |> List.exists (fun (referenceKind, _) -> referenceKind = kind)
            TargetReferences = own
            TargetY =
                if isHeader then pagesArray.[pageIndex].Top + sectionPages.SectionHeaderDistance.[section]
                else pagesArray.[pageIndex].Top + pagesArray.[pageIndex].Height - sectionPages.SectionFooterDistance.[section] - settings.LineHeight * 0.78
            TargetFields = (pageFieldValues sectionPages pagesArray.Length).[pageIndex]
        }

/// The blocks of a header or footer with the page number and the page count
/// in the fields that show them (what a page paints).
let blocksWithRunningFields (values: RunningFieldValues) (blocks: Block list) : Block list =
    let mutable changed = false
    let mapped =
        blocks |> List.map (fun block ->
            let next, blockChanged = blockWithRunningFields values block
            if blockChanged then changed <- true
            next)
    if changed then mapped else blocks

/// The layouts of a header's or footer's blocks where a page paints them
/// (from `startY` down, as `repeatedBlockCommands` places them): what the
/// caret, the selection and the pointer go by while it is edited in place.
let storyBlockLayouts (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (startY: float) (blocks: Block list) : BlockLayout list =
    let styleContext = WordStyles.styleContext doc
    let repeatedSettings = { settings with MarginBottom = 0.0 }
    let maxWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
    let layouts = ResizeArray<BlockLayout>()
    let mutable y = startY
    for block in blocks do
        match block with
        | Paragraph para
        | Heading (_, para) ->
            let safePara = { para with Metadata = { para.Metadata with PageBreakAfter = false } }
            let layout, _ = layoutParagraphAtWithStyleWithRange repeatedSettings measureText measureRange maxWidth y 0.0 (forBlockInContext styleContext block) (effectiveBufferInContext styleContext) safePara Map.empty
            layouts.Add layout
            y <- layout.Y + layout.Height
        | Table table ->
            let _, nextY = tableBorderCommandsCore (Some styleContext) repeatedSettings measureText measureRange table y
            let measureSettings = { repeatedSettings with PageHeight = max repeatedSettings.PageHeight 1000000.0 }
            let tableDoc = storyTableDocument doc table
            let dy = y - measureSettings.PaddingTop
            for layout in layoutDocumentWithRange measureSettings measureText measureRange tableDoc do
                if not (layout.BlockId.StartsWith footnoteBlockIdPrefix) && not (layout.BlockId.StartsWith endnoteBlockIdPrefix) then
                    layouts.Add { layout with Y = layout.Y + dy; Lines = layout.Lines |> List.map (fun line -> { line with Y = line.Y + dy }) }
            y <- nextY
        | TextBox _ -> ()
    resizeArrayToList layouts

/// Where the shapes of a header or footer are on a page that shows it: each
/// with its left edge (from the left edge of the page), its top (down the
/// canvas), its width and its height - where the page paints them. `startY`:
/// where the header's (footer's) blocks start on that page (headerFooterStartY).
let storyShapeRects (settings: LayoutSettings) measureText measureRange (doc: Document) (page: PageInfo) (startY: float) (blocks: Block list) : (TextBoxBlock * float * float * float * float) list =
    let shapes = blocks |> List.choose (function TextBox textBox -> Some textBox | _ -> None)
    if shapes.IsEmpty then []
    else
        let styleContext = WordStyles.styleContext doc
        let tops = Dictionary<BlockId, float>()
        for layout in storyBlockLayouts settings measureText measureRange doc startY blocks do
            if not (tops.ContainsKey layout.BlockId) then tops.[layout.BlockId] <- layout.Y
        shapes
        |> List.choose (fun textBox ->
            let anchorTop =
                textBox.AnchorBlockId
                |> Option.bind (fun anchorId ->
                    blocks
                    |> List.tryFind (function
                        | Paragraph para
                        | Heading (_, para) -> para.Id = anchorId
                        | Table table -> table.Id = anchorId
                        | TextBox _ -> false))
                |> Option.bind firstParagraphIdOf
                |> Option.bind (fun id ->
                    match tops.TryGetValue id with
                    | true, top -> Some top
                    | false, _ -> None)
            match anchorTop with
            | Some _ ->
                let struct (x, y, height) = textBoxPlace settings measureText measureRange styleContext [| page |] anchorTop 0.0 textBox
                Some (textBox, x, y, textBox.Width, height)
            // (A shape of a header that is anchored in nothing is not painted.)
            | None -> None)

/// The boundary of the header or footer that is being edited: a dashed line
/// across the page where the body text begins (ends), and a tab with its name.
let storyBoundaryCommands (settings: LayoutSettings) fontFamily (page: PageInfo) (isHeader: bool) (label: string) : PaintCommand list =
    let y = if isHeader then page.Top + page.MarginTop else page.Top + page.Height - page.MarginBottom
    let tabWidth = 14.0 + float label.Length * 6.6
    let tabHeight = 20.0
    let tabX = max 8.0 (page.MarginLeft - 4.0)
    let tabY = if isHeader then y else y - tabHeight
    let style = { defaultPaintStyle 12.0 fontFamily with Color = "#1f4e99" }
    [ StrokeLine ("#5b9bd5", 1.0, 2, 0.0, 0.0, y, page.Width, y)
      FillRect ("#ffffff", tabX, tabY, tabWidth, tabHeight)
      StrokeRect ("#5b9bd5", 1.0, tabX + 0.5, tabY + 0.5, tabWidth - 1.0, tabHeight - 1.0)
      DrawText (label, style, tabX + 7.0, tabY + 14.0, 0.0) ]

/// Where the shapes of the body are: each with its left edge (from the left
/// edge of its page), its top (down the canvas), its width and its height -
/// where the painter draws them, from the layouts and pages of a layout pass.
let bodyShapeRects (settings: LayoutSettings) measureText measureRange (doc: Document) (layouts: BlockLayout list) (pages: PageInfo list) : (TextBoxBlock * float * float * float * float) list =
    let shapes = doc.Blocks |> List.choose (function TextBox textBox -> Some textBox | _ -> None)
    if shapes.IsEmpty then []
    else
        let styleContext = WordStyles.styleContext doc
        let pageArray = List.toArray pages
        let tops = Dictionary<BlockId, float>()
        for layout in layouts do
            if not (tops.ContainsKey layout.BlockId) then tops.[layout.BlockId] <- layout.Y
        let anchorTopOf (textBox: TextBoxBlock) =
            textBox.AnchorBlockId
            |> Option.bind (fun anchorId ->
                doc.Blocks
                |> List.tryFind (function
                    | Paragraph para
                    | Heading (_, para) -> para.Id = anchorId
                    | Table table -> table.Id = anchorId
                    | TextBox _ -> false))
            |> Option.bind firstParagraphIdOf
            |> Option.bind (fun id ->
                match tops.TryGetValue id with
                | true, top -> Some top
                | false, _ -> None)
        let shifts = overlapShifts settings measureText measureRange styleContext pageArray anchorTopOf shapes
        shapes
        |> List.map (fun textBox ->
            let struct (x, y, height) = textBoxPlace settings measureText measureRange styleContext pageArray (anchorTopOf textBox) (shiftOf shifts textBox) textBox
            textBox, x, y, textBox.Width, height)

/// The layouts of a shape's text where the painter draws it (see `bodyShapeRects` for the shape's place).
let shapeTextLayouts (settings: LayoutSettings) measureText measureRange (doc: Document) (textBox: TextBoxBlock) (x: float) (y: float) (height: float) : BlockLayout list =
    if textBox.FitText then [] else textBoxContentLayouts settings measureText measureRange (WordStyles.styleContext doc) textBox x y height

/// The layouts of the paragraphs of a note's or comment's body that is edited
/// in place, their left edge at `left` and the first one's top at `top`
/// (across and down the canvas), each `width` wide at most, with the
/// formatting they have (not broken over pages).
let noteStoryLayouts (settings: LayoutSettings) measureText measureRange (doc: Document) (left: float) (top: float) (width: float) (blocks: Block list) : BlockLayout list =
    let styleContext = WordStyles.styleContext doc
    let storySettings = { settings with MarginBottom = 0.0; InTableCell = false }
    let result = ResizeArray<BlockLayout>()
    let mutable y = top
    let mutable counters : Map<string * int, int> = Map.empty
    for para in paragraphsOfBlocks blocks do
        let safePara = { para with Metadata = { para.Metadata with PageBreakAfter = false } }
        let layout, next =
            layoutParagraphAtWithStyleWithRange storySettings measureText measureRange (max 1.0 width) y (left - settings.PaddingLeft) (forBlockInContext styleContext (Paragraph para)) (effectiveBufferInContext styleContext) safePara counters
        counters <- next
        result.Add layout
        y <- layout.Y + layout.Height
    resizeArrayToList result

let private tableHasRowSpanningVerticalAlignment (tableBlock: TableBlock) =
    let mutable found = false
    for row in tableBlock.Rows do
        if not found then
            for cell in row.Cells do
                if not found then
                    match cell.VerticalMerge, cell.VerticalAlignment with
                    | Some RestartMerge, Some alignment when alignment = "center" || alignment = "bottom" -> found <- true
                    | _ -> ()
    found

let private firstActualTableContentLayout (index: LayoutIndex) tableBlock =
    let mutable firstId = ""
    let mutable firstY = System.Double.PositiveInfinity
    for row in tableBlock.Rows do
        for cell in row.Cells do
            for para in cell.Content do
                // (Where the paragraph stands first: a header row's is laid out again
                // on every page the table goes on to, and by its last place the table
                // seemed to start at its first body row - the borders were drawn that
                // far down, and the header rows had none.)
                match tryFirstBlockLayout para.Id index with
                | Some layout when layout.Y < firstY ->
                    firstId <- para.Id
                    firstY <- layout.Y
                | None -> ()
                | _ -> ()

    if System.Double.IsPositiveInfinity firstY then None else Some (firstId, firstY)

let private rowContainsBlock blockId (row: TableRow) =
    let mutable found = false
    for cell in row.Cells do
        if not found then
            for para in cell.Content do
                if para.Id = blockId then found <- true
    found

let private plannedContentOffsetForBlock (settings: LayoutSettings) measureText measureRange doc tableBlock blockId =
    if tableHasRowSpanningVerticalAlignment tableBlock then
        None
    else
        let styleContext = WordStyles.styleContext doc
        let tableLeftOffset = tableLeftOffsetForStyleContext (Some styleContext) + WordLayout.Table.tableIndent tableBlock
        let maxWidth = max 0.0 (settings.Width - settings.PaddingLeft - settings.MarginRight - tableLeftOffset)
        let outerLeft, outerRight = WordLayout.Table.outerCellInsets WordLayout.canvasScale settings.CompatibilityMode tableBlock
        let tableMaxWidth = maxWidth + outerLeft + outerRight
        let columnWidths = tableColumnWidths tableMaxWidth tableBlock
        let columnMetrics = tableColumnMetrics tableMaxWidth columnWidths
        let tableWidth = WordLayout.Table.tableWidthWithSpacing columnWidths (WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock)
        let tableHorizontalOffset = tableLeftOffset + tableAlignmentOffset tableBlock.Alignment maxWidth tableWidth
        let mutable result = None

        for row in tableBlock.Rows do
            if result.IsNone && rowContainsBlock blockId row then
                let measurement =
                    measureTableRowForContextRowLocalWithColumnMetrics (tableMeasureScopeFrom None (listCountersBeforeTable doc tableBlock) tableBlock) settings measureText measureRange styleContext tableBlock columnMetrics outerLeft tableHorizontalOffset row
                for group in measurement.CellGroups do
                    if result.IsNone then
                        let offset =
                            WordLayout.Table.verticalContentOffset group.VerticalAlignment measurement.Height group.PaddingTop group.PaddingBottom group.ContentHeight
                        for layout in group.Layouts do
                            if result.IsNone && layout.BlockId = blockId then
                                result <- Some (layout.Y + offset)
        result

/// What `plannedContentOffsetForBlock` gave for a table, kept like the table
/// plans (per scope, see `Layout.withTablePlanScope`): it measures a row of
/// the table, and was asked for every table of the document at every paint.
type private KnownContentOffset = {
    Scope: obj
    Settings: LayoutSettings
    Styles: NamedStyle list
    BlockId: BlockId
    Offset: float option
}

#if FABLE_COMPILER
[<Fable.Core.Emit("new WeakMap()")>]
let private newContentOffsetMap () : obj = Fable.Core.Util.jsNative
[<Fable.Core.Emit("$0.get($1)")>]
let private contentOffsetMapGet (map: obj) (key: TableBlock) : KnownContentOffset = Fable.Core.Util.jsNative
[<Fable.Core.Emit("$0.set($1, $2)")>]
let private contentOffsetMapSet (map: obj) (key: TableBlock) (value: KnownContentOffset) : unit = Fable.Core.Util.jsNative
let private knownContentOffsets = newContentOffsetMap ()
let private tryKnownContentOffset (table: TableBlock) : KnownContentOffset option =
    let entry = contentOffsetMapGet knownContentOffsets table
    if isNull (box entry) then None else Some entry
let private rememberContentOffset (table: TableBlock) (entry: KnownContentOffset) = contentOffsetMapSet knownContentOffsets table entry
#else
let private knownContentOffsets = System.Runtime.CompilerServices.ConditionalWeakTable<TableBlock, KnownContentOffset>()
let private tryKnownContentOffset (table: TableBlock) : KnownContentOffset option =
    match knownContentOffsets.TryGetValue table with
    | true, entry -> Some entry
    | false, _ -> None
let private rememberContentOffset (table: TableBlock) (entry: KnownContentOffset) =
    lock knownContentOffsets (fun () ->
        knownContentOffsets.Remove table |> ignore
        knownContentOffsets.Add(table, entry))
#endif

let private plannedContentOffsetForBlockKept (settings: LayoutSettings) measureText measureRange (doc: Document) (tableBlock: TableBlock) (blockId: BlockId) =
    let scope = currentTablePlanScope ()
    if isNull scope then plannedContentOffsetForBlock settings measureText measureRange doc tableBlock blockId
    else
        match tryKnownContentOffset tableBlock with
        | Some entry when obj.ReferenceEquals(entry.Scope, scope) && entry.BlockId = blockId && obj.ReferenceEquals(entry.Styles, doc.Styles)
                          && (obj.ReferenceEquals(entry.Settings, settings) || entry.Settings = settings) ->
            entry.Offset
        | _ ->
            let offset = plannedContentOffsetForBlock settings measureText measureRange doc tableBlock blockId
            rememberContentOffset tableBlock { Scope = scope; Settings = settings; Styles = doc.Styles; BlockId = blockId; Offset = offset }
            offset

let private plannedContentOffsetFromFullPlan (settings: LayoutSettings) measureText measureRange doc tableBlock =
    let mutable firstPlannedContentOffset = System.Double.PositiveInfinity
    let planAtZero = buildTableLayoutPlanForDocumentAt settings measureText measureRange doc tableBlock 0.0
    let resolveAlignmentHeight = tableCellVerticalAlignmentHeightResolverForPlan planAtZero
    for row in planAtZero.Rows do
        let mutable groupIndex = 0
        for group in row.Measurement.CellGroups do
            let alignmentHeight = resolveAlignmentHeight row.RowIndex groupIndex row.Height
            let offset = WordLayout.Table.verticalContentOffset group.VerticalAlignment alignmentHeight group.PaddingTop group.PaddingBottom group.ContentHeight
            for layout in group.Layouts do
                firstPlannedContentOffset <- min firstPlannedContentOffset (layout.Y + offset)
            groupIndex <- groupIndex + 1
    if System.Double.IsPositiveInfinity firstPlannedContentOffset then None else Some firstPlannedContentOffset

let private canonicalTableYFromLayouts (settings: LayoutSettings) measureText measureRange doc (index: LayoutIndex) tableBlock =
    match firstActualTableContentLayout index tableBlock with
    | None -> None
    | Some (blockId, firstActualContentY) ->
        match plannedContentOffsetForBlockKept settings measureText measureRange doc tableBlock blockId
              |> Option.orElseWith (fun () -> plannedContentOffsetFromFullPlan settings measureText measureRange doc tableBlock) with
        | Some firstPlannedContentOffset -> Some (firstActualContentY - firstPlannedContentOffset)
        | None -> Some firstActualContentY

let rec private shiftCommandX dx command =
    match command with
    | DrawText (t, s, x, b, w) -> DrawText (t, s, x + dx, b, w)
    | DrawUnderline (c, s, x, y, w) -> DrawUnderline (c, s, x + dx, y, w)
    | DrawImage (s, x, y, w, h, cl, ct, cr, cb) -> DrawImage (s, x + dx, y, w, h, cl, ct, cr, cb)
    | FillRect (c, x, y, w, h) -> FillRect (c, x + dx, y, w, h)
    | StrokeRect (c, lw, x, y, w, h) -> StrokeRect (c, lw, x + dx, y, w, h)
    | StrokeLine (c, lw, ds, doff, x1, y1, x2, y2) -> StrokeLine (c, lw, ds, doff, x1 + dx, y1, x2 + dx, y2)
    | DrawShape (fill, stroke, sw, (bx, by, bw, bh), segs) ->
        DrawShape (fill, stroke, sw, (bx + dx, by, bw, bh), Shapes.translatePath dx 0.0 segs)
    | DrawGroup (clip, rotation, cx, cy, opacity, (bx, by, bw, bh), commands) ->
        DrawGroup (clip |> Option.map (fun (x, y, w, h) -> (x + dx, y, w, h)), rotation, cx + dx, cy, opacity, (bx + dx, by, bw, bh), commands |> List.map (shiftCommandX dx))
    | DrawMirrored (flipH, flipV, cx, cy, (bx, by, bw, bh), commands) ->
        DrawMirrored (flipH, flipV, cx + dx, cy, (bx + dx, by, bw, bh), commands |> List.map (shiftCommandX dx))

let private shiftCommandsX dx commands =
    let shifted = ResizeArray<PaintCommand>()
    for command in commands do
        shifted.Add(shiftCommandX dx command)
    resizeArrayToList shifted

let private sectionColumnsAreMulti = function
    | Some columns when columns.Count >= 2 -> true
    | _ -> false

/// Whether layouts of the document can stand to the right of where a block
/// laid out on its own stands: in a column of a multi-column section, or in
/// a section whose left margin is not the document's (the final section's).
/// A table's borders then follow its cells' layouts (see tableColumnOffset).
let private documentHasMultiColumnLayout (doc: Document) =
    let mutable found = sectionColumnsAreMulti doc.PageSetup.Columns
    for block in doc.Blocks do
        if not found then
            match Rbt.Word.Document.tryGetParagraph block with
            | Some para ->
                found <-
                    sectionColumnsAreMulti para.Metadata.SectionColumns
                    || (match para.Metadata.SectionPageSetup with
                        | Some setup -> setup.MarginLeft <> doc.PageSetup.MarginLeft
                        | None -> false)
            | None -> ()
    found

/// When a table sits in a multi-column section its cell layouts were shifted by
/// the column offset (their Indent gained columnDx); the freshly built plan is
/// at full width. The difference is the X offset to shift the borders by.
let private tableColumnOffset (index: LayoutIndex) (plan: TableLayoutPlan) (tableBlock: TableBlock) =
    let cellIdSet = HashSet<BlockId>()
    let mutable actualMin = System.Double.PositiveInfinity
    for row in tableBlock.Rows do
        for cell in row.Cells do
            for para in cell.Content do
                cellIdSet.Add para.Id |> ignore
                match tryLayout para.Id index with
                | Some layout when layout.Indent < actualMin -> actualMin <- layout.Indent
                | _ -> ()

    let mutable plannedMin = System.Double.PositiveInfinity
    for row in plan.Rows do
        for group in row.Measurement.CellGroups do
            for layout in group.Layouts do
                if cellIdSet.Contains layout.BlockId && layout.Indent < plannedMin then
                    plannedMin <- layout.Indent

    if System.Double.IsPositiveInfinity actualMin || System.Double.IsPositiveInfinity plannedMin then 0.0
    else actualMin - plannedMin

/// What of a command lies in the band `top` .. `bottom` (the rows of a table
/// that stand together): lines along the band's edges, the part of an
/// upright line and of a filled rectangle that is inside it.
let private clipCommandToBand (top: float) (bottom: float) (command: PaintCommand) : PaintCommand option =
    let slack = 0.75
    match command with
    | StrokeLine (c, lw, ds, doff, x1, y1, x2, y2) ->
        if abs (y1 - y2) < 0.01 then
            if y1 >= top - slack && y1 <= bottom + slack then Some command else None
        else
            let low, high = max top (min y1 y2), min bottom (max y1 y2)
            // (An upright line across several rows keeps its direction.)
            if high - low < 1.0 then None
            elif y1 <= y2 then Some (StrokeLine (c, lw, ds, doff, x1, low, x2, high))
            else Some (StrokeLine (c, lw, ds, doff, x1, high, x2, low))
    | FillRect (c, x, y, w, h) ->
        let low, high = max top y, min bottom (y + h)
        if high - low < 0.5 then None else Some (FillRect (c, x, low, w, high - low))
    | StrokeRect (_, _, _, y, _, h) -> if y >= top - slack && y + h <= bottom + slack then Some command else None
    | DrawText (_, _, _, baseline, _) -> if baseline >= top && baseline <= bottom + slack then Some command else None
    | DrawUnderline (_, _, _, y, _) -> if y >= top && y <= bottom + slack then Some command else None
    | DrawImage (_, _, y, _, h, _, _, _, _) -> if y >= top - slack && y + h <= bottom + slack then Some command else None
    | DrawShape (_, _, _, (_, y, _, h), _) -> if y >= top - slack && y + h <= bottom + slack then Some command else None
    | DrawGroup (_, _, _, _, _, (_, y, _, h), _) -> if y >= top - slack && y + h <= bottom + slack then Some command else None
    | DrawMirrored (_, _, _, _, (_, y, _, h), _) -> if y >= top - slack && y + h <= bottom + slack then Some command else None

/// The borders and fills of a table in a multi-column section. The layout
/// pass lays such a table out on a strip as wide as a column and places it row
/// by row: the rows that fit go under one another, the next ones start the
/// next column (or page) below the header rows, which are repeated there. The
/// table's commands are made on the same strip and moved, stretch of rows by
/// stretch of rows, to where the rows' paragraphs are. None when the table is
/// not in such a section (or its paragraphs have no layout).
let private columnTableFragmentCommands (settings: LayoutSettings) measureText measureRange (doc: Document) (index: LayoutIndex) (tableBlock: TableBlock) : (PaintCommand list * float) option =
    let actualLayouts (blockId: BlockId) : BlockLayout list =
        match index.ByBlock.TryGetValue blockId with
        | true, layouts -> List.ofSeq layouts
        | false, _ ->
            match index.ById.TryGetValue blockId with
            | true, layout -> [ layout ]
            | false, _ -> []
    // The strip the table was laid out on: as wide as a column of its section. (Held against
    // the layouts - the cells are as wide there as they were laid out - so that a document of
    // which only the start is laid out, whose sections are not all known, is painted the same.)
    let stripOf (candidate: LayoutSettings) =
        let plan, bands = tableStripRows None ValueNone candidate measureText measureRange doc tableBlock
        let fits =
            plan.Rows |> List.forall (fun row ->
                row.Measurement.CellGroups |> List.forall (fun group ->
                    group.Layouts |> List.forall (fun planned ->
                        match actualLayouts planned.BlockId with
                        | actual :: _ -> abs (actual.AvailableWidth - planned.AvailableWidth) < 0.5
                        | [] -> true)))
        if fits then Some (candidate, plan, bands) else None
    let strip =
        match tableColumnStripSettings settings doc tableBlock with
        | None -> None
        | Some stripSettings ->
            stripOf stripSettings
            |> Option.orElseWith (fun () -> stripOf { settings with PageHeight = stripSettings.PageHeight })
    match strip with
    | None -> None
    | Some (stripSettings, plan, bands) ->
        let stripCommands = tableFragmentPaintCommands plan
        // Where a row is: how far its paragraphs' layouts are from the strip's (every occurrence, for a repeated header row).
        let offsetsOfRow (row: TableLayoutPlanRow) : (float * float) list =
            row.Measurement.CellGroups
            |> List.tryPick (fun group ->
                group.Layouts
                |> List.tryPick (fun planned ->
                    match actualLayouts planned.BlockId with
                    | [] -> None
                    | found -> Some (found |> List.map (fun layout -> (settings.PaddingLeft + layout.Indent) - (stripSettings.PaddingLeft + planned.Indent), layout.Y - planned.Y))))
            |> Option.defaultValue []
        let rows = List.toArray plan.Rows
        let offsets = rows |> Array.map offsetsOfRow
        if rows.Length = 0 || bands.Length <> rows.Length || offsets |> Array.forall List.isEmpty then None
        else
            let commands = ResizeArray<PaintCommand>()
            let addStretch (first: int) (last: int) (dx: float, dy: float) =
                let top, bottom = fst bands.[first], snd bands.[last]
                for command in stripCommands do
                    match clipCommandToBand top bottom command with
                    | Some clipped -> commands.Add (clipped |> shiftCommandX dx |> shiftCommandY dy)
                    | None -> ()
            // The rows in order, each where its first occurrence is (a row without a paragraph of its own goes with the row above it).
            let mutable current = offsets |> Array.tryPick List.tryHead |> Option.defaultValue (0.0, 0.0)
            let mutable stretchStart = 0
            let mutable lastOffset = current
            for rowIndex in 0 .. rows.Length - 1 do
                let here = match offsets.[rowIndex] with first :: _ -> first | [] -> current
                if rowIndex > 0 && (abs (fst here - fst current) > 0.5 || abs (snd here - snd current) > 0.5) then
                    addStretch stretchStart (rowIndex - 1) current
                    stretchStart <- rowIndex
                current <- here
                lastOffset <- here
            addStretch stretchStart (rows.Length - 1) current
            // The header rows where they are repeated.
            let headerCount = min plan.HeaderRowCount rows.Length
            if headerCount > 0 then
                for occurrence in List.tail (match offsets.[0] with [] -> [ (0.0, 0.0) ] | found -> found) do
                    addStretch 0 (headerCount - 1) occurrence
            Some (resizeArrayToList commands, plan.NextY + snd lastOffset)

let private tableFragmentPaintCommandsForDocument settings measureText measureRange doc index hasMultiColumnLayout tableBlock =
    match (if hasMultiColumnLayout then columnTableFragmentCommands settings measureText measureRange doc index tableBlock else None) with
    | Some placed -> Some placed
    | None ->
    canonicalTableYFromLayouts settings measureText measureRange doc index tableBlock
    |> Option.map (fun tableY ->
        let plan = buildTableLayoutPlanForDocumentAt settings measureText measureRange doc tableBlock tableY
        let dx = if hasMultiColumnLayout then tableColumnOffset index plan tableBlock else 0.0
        let commands = tableFragmentPaintCommands plan
        let commands = if abs dx > 0.5 then shiftCommandsX dx commands else commands
        commands, plan.NextY)

type private TableTextCommandFrame =
    | TableTextTable of TableBlock
    | TableTextCell of TableBlock * TableRow * TableCell

let private addTableTextCommands (settings: LayoutSettings) fontFamily measureText measureRange (index: LayoutIndex) (commands: ResizeArray<PaintCommand>) (tableBlock: TableBlock) =
    let frames = Stack<TableTextCommandFrame>()
    frames.Push(TableTextTable tableBlock)
    while frames.Count > 0 do
        match frames.Pop() with
        | TableTextTable table ->
            for row in List.rev table.Rows do
                for cell in List.rev row.Cells do
                    frames.Push(TableTextCell (table, row, cell))
        | TableTextCell (table, row, cell) ->
            let tableCellBaselineOffset =
                if WordLayout.Table.rowCellSpacing WordLayout.canvasScale table row > 0.0 then spacedTableCellBaselineOffset else 0.0
            for para in cell.Content do
                match index.ByBlock.TryGetValue para.Id with
                | true, paraLayouts ->
                    for paraLayout in paraLayouts do
                        addParagraphCommandsWithBaselineOffset tableCellBaselineOffset settings fontFamily measureText measureRange commands paraLayout
                | false, _ ->
                    match index.ById.TryGetValue para.Id with
                    | true, paraLayout ->
                        addParagraphCommandsWithBaselineOffset tableCellBaselineOffset settings fontFamily measureText measureRange commands paraLayout
                    | false, _ -> ()
            for nestedTable in List.rev cell.NestedTables do
                frames.Push(TableTextTable nestedTable)

// The fragment paint path draws outer-table borders but not nested-table
// borders. Locate each nested table from its content paragraph layouts and draw
// its grid with the shared border painter at that position.
type private NestedTableBorderFrame =
    | ScanNestedTable of TableBlock
    | PaintNestedTable of TableBlock

let private addNestedTableBorders (styleContext: DocumentStyleContext option) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (index: LayoutIndex) (commands: ResizeArray<PaintCommand>) (tableBlock: TableBlock) =
    let frames = Stack<NestedTableBorderFrame>()
    frames.Push(ScanNestedTable tableBlock)
    while frames.Count > 0 do
        match frames.Pop() with
        | ScanNestedTable table ->
            for row in List.rev table.Rows do
                for cell in List.rev row.Cells do
                    for nestedTable in List.rev cell.NestedTables do
                        frames.Push(PaintNestedTable nestedTable)
        | PaintNestedTable nestedTable ->
            let nestedLayouts = ResizeArray<BlockLayout>()
            let nestedTopYs = ResizeArray<float>()
            let segXs = ResizeArray<float>()
            for nestedRow in nestedTable.Rows do
                for nestedCell in nestedRow.Cells do
                    for para in nestedCell.Content do
                        let addNestedLayout layout =
                            nestedLayouts.Add layout
                            nestedTopYs.Add layout.Y
                            for line in layout.Lines do
                                for segment in line.Segments do
                                    match segment with
                                    | TextSegment (_, _, x, _, _) -> segXs.Add x
                                    | _ -> ()
                        match index.ByBlock.TryGetValue para.Id with
                        | true, paraLayouts ->
                            for layout in paraLayouts do
                                addNestedLayout layout
                        | false, _ ->
                            match index.ById.TryGetValue para.Id with
                            | true, layout -> addNestedLayout layout
                            | false, _ -> ()
            // Where each row of the nested table really is: the first line of
            // its cells. A cell's layout starts at its row's top, and its first
            // line a fixed distance below that - unless the row was moved to
            // the next page with the rest of the outer row, where the layout
            // stays behind and the lines are on the next page.
            let firstLineOffsets = ResizeArray<float>()
            let rowFirstLines =
                nestedTable.Rows
                |> List.map (fun nestedRow ->
                    let firstLines = ResizeArray<float>()
                    for nestedCell in nestedRow.Cells do
                        match nestedCell.Content with
                        | para :: _ ->
                            let consider (layout: BlockLayout) =
                                match layout.Lines with
                                | line :: _ ->
                                    firstLines.Add line.Y
                                    firstLineOffsets.Add (line.Y - layout.Y)
                                | [] -> ()
                            match index.ByBlock.TryGetValue para.Id with
                            | true, paraLayouts -> (match paraLayouts |> Seq.tryHead with Some layout -> consider layout | None -> ())
                            | false, _ ->
                                match index.ById.TryGetValue para.Id with
                                | true, layout -> consider layout
                                | false, _ -> ()
                        | [] -> ()
                    minResizeArray firstLines)
            let rowTops =
                match minResizeArray firstLineOffsets with
                | Some offset ->
                    rowFirstLines
                    |> List.map (fun firstLine ->
                        match firstLine with
                        | Some lineY ->
                            // (At the top of a page there is no space above the first line.)
                            let contentTop = WordLayout.Page.contentTop settings.PaddingTop (WordLayout.Page.pageTopForY (pageHeight settings) lineY)
                            max (lineY - offset) (min lineY contentTop)
                        | None -> System.Double.NaN)
                    |> List.toArray
                | None -> [||]
            match minResizeArray nestedTopYs, minResizeArray segXs with
            | Some topY, Some minX ->
                let firstRow = nestedTable.Rows |> List.head
                let firstCell = firstRow.Cells |> List.head
                let padLeft = (WordLayout.Table.cellPaddingForRow WordLayout.canvasScale nestedTable firstRow firstCell).Left
                let baseHorizontalOffset = minX - padLeft
                let availableWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
                // (Header rows repeat on their own; the given tops are those of the body rows.)
                let bodyTops = if nestedTable.Rows |> List.exists (fun row -> row.IsHeader) then None else Some rowTops
                let topY = match bodyTops with Some tops when tops.Length > 0 && not (System.Double.IsNaN tops.[0]) && tops.[0] > topY + 0.5 -> tops.[0] | _ -> topY
                let cmds, _ = tableBorderCommandsCoreAt 0 styleContext settings measureText measureRange nestedTable topY availableWidth baseHorizontalOffset bodyTops
                commands.AddRange cmds
            | _ -> ()
            frames.Push(ScanNestedTable nestedTable)

/// The footnote and endnote areas: the synthetic note paragraphs the layout
/// appended, with their separator rules.
/// The borders of the tables of the notes the note areas show, where their
/// cells are laid out (a note's paragraphs and cells are laid out by the
/// layout pass with the note's own block ids in theirs, see Layout.noteParagraphIdOfLayoutId).
let private addNoteTableCommands (settings: LayoutSettings) measureText measureRange (doc: Document) (commands: ResizeArray<PaintCommand>) (noteLayouts: seq<BlockLayout>) =
    // (Note by note, where a page shows it: a note shown twice has its tables twice.)
    let shown = Dictionary<string, ResizeArray<BlockLayout>>()
    let order = ResizeArray<string * string * string>()
    for layout in noteLayouts do
        match noteOfLayoutId layout.BlockId, noteParagraphIdOfLayoutId layout.BlockId with
        | Some (kind, id), Some own ->
            let key = noteAreaSlotOfLayoutId layout.BlockId + "\u0001" + kind + "\u0001" + id
            match shown.TryGetValue key with
            | true, list -> list.Add { layout with BlockId = own }
            | false, _ ->
                let list = ResizeArray<BlockLayout>()
                list.Add { layout with BlockId = own }
                shown.[key] <- list
                order.Add (key, kind, id)
        | _ -> ()
    for key, kind, id in order do
        match storyBody kind id doc with
        | Some blocks when blocks |> List.exists (function Table _ -> true | _ -> false) ->
            let layouts = List.ofSeq shown.[key]
            let index = buildLayoutIndex settings false layouts.Length layouts
            // (Laid out on an unpaged page, as the note was: a table of a note is not broken by the page it is on.)
            let noteSettings = { settings with PageHeight = 1.0e12; MarginBottom = 0.0; TableContinuationHeaderHeight = 0.0 }
            let noteDoc = { doc with Blocks = blocks; PageSetup = { doc.PageSetup with Columns = None } }
            for block in blocks do
                match block with
                | Table table ->
                    match tableFragmentPaintCommandsForDocument noteSettings measureText measureRange noteDoc index true table with
                    | Some (tableCommands, _) -> commands.AddRange tableCommands
                    | None -> ()
                | _ -> ()
        | _ -> ()

let private addNoteAreaCommands (settings: LayoutSettings) fontFamily measureText measureRange (doc: Document) (commands: ResizeArray<PaintCommand>) (footnotes: ResizeArray<BlockLayout>) (endnotes: ResizeArray<BlockLayout>) =
    // The tables of the notes, under their text.
    if footnotes.Count > 0 || endnotes.Count > 0 then
        addNoteTableCommands settings measureText measureRange doc commands (Seq.append footnotes endnotes)
    // Footnote area: paint the synthetic footnote paragraphs appended by layout
    // at the bottom of each page, and a short separator rule above the topmost
    // note on each page (Word draws ~2in).
    for layout in footnotes do
        addParagraphCommands settings fontFamily measureText measureRange commands layout
    if footnotes.Count > 0 then
        let footnoteTopByPage = Dictionary<int, float>()
        for layout in footnotes do
            let page = pageOfY settings layout.Y
            match footnoteTopByPage.TryGetValue page with
            | true, topY when topY <= layout.Y -> ()
            | _ -> footnoteTopByPage.[page] <- layout.Y
        for KeyValue(_, topY) in footnoteTopByPage do
            let separatorX = settings.PaddingLeft
            commands.Add(StrokeLine ("#000000", 0.75, 0, 0.0, separatorX, topY - 6.0, separatorX + 192.0, topY - 6.0))

    // Endnote area: paint the synthetic endnote paragraphs appended after the
    // body content, with a single separator rule above the first endnote.
    for layout in endnotes do
        addParagraphCommands settings fontFamily measureText measureRange commands layout
    if endnotes.Count > 0 then
        // Short (~2in) separator above the first endnote; Word draws a full-width
        // continuation separator on pages the endnotes spill onto.
        let endnoteTopByPage = Dictionary<int, float>()
        for layout in endnotes do
            let page = pageOfY settings layout.Y
            match endnoteTopByPage.TryGetValue page with
            | true, topY when topY <= layout.Y -> ()
            | _ -> endnoteTopByPage.[page] <- layout.Y
        let mutable startPage = System.Int32.MaxValue
        for KeyValue(page, _) in endnoteTopByPage do
            if page < startPage then startPage <- page
        let fullWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
        for KeyValue(page, topY) in endnoteTopByPage do
            let separatorX = settings.PaddingLeft
            let width = if page = startPage then 192.0 else fullWidth
            commands.Add(StrokeLine ("#000000", 0.75, 0, 0.0, separatorX, topY - 8.0, separatorX + width, topY - 8.0))

let private collectDocumentCommandsWithLayoutsUncached (settings: LayoutSettings) fontFamily (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (cachedLayouts: BlockLayout list option) (cachedPages: PageInfo list option) =
    let layouts, pages, layoutCount =
        match cachedLayouts, cachedPages with
        | Some ls, Some pages ->
            // The layout pass that produced the layouts also produced the page
            // geometry (per section); use it rather than guess.
            ls, pages, List.length ls
        | Some ls, None when documentHasSections doc ->
            // Sections may differ in page size and orientation, which cannot be
            // read back off the layouts: ask the layout pass for the pages.
            let _, pages = layoutDocumentPages settings measureText measureRange doc
            ls, pages, List.length ls
        | Some ls, None ->
            let pages, layoutCount = uniformPageInfosWithLayoutCount settings ls
            ls, pages, layoutCount
        | None, _ ->
            let layouts, pages = layoutDocumentPages settings measureText measureRange doc
            layouts, pages, List.length layouts
    let needsBodyLineIndex =
        match doc.PageSetup.Columns with
        | Some columns when columns.Count >= 2 && columns.Separator -> true
        | _ -> false
    let hasMultiColumnLayout = documentHasMultiColumnLayout doc
    let index = buildLayoutIndex settings needsBodyLineIndex layoutCount layouts
    let commands = ResizeArray<PaintCommand>(max 16 (layoutCount * 2))
    let mutable currentY = doc.PageSetup.MarginTop
    let sectionPageBreaks = lazy (sectionPageBreakParagraphs doc)
    let styleContext = WordStyles.styleContext doc
    // (The anchored drawings of the body are painted below and above its text: see Drawing.)
    let drawings = ResizeArray<Drawing>()
    let pageArray = lazy (List.toArray pages)
    // (Where the block a shape is anchored in starts.)
    let anchorTopOf (textBox: TextBoxBlock) =
        textBox.AnchorBlockId
        |> Option.bind (fun anchorId ->
            doc.Blocks
            |> List.tryFind (function
                | Paragraph para
                | Heading (_, para) -> para.Id = anchorId
                | Table table -> table.Id = anchorId
                | TextBox _ -> false))
        |> Option.bind firstParagraphIdOf
        |> Option.bind (fun id -> tryFirstBlockLayout id index)
        |> Option.map (fun layout -> layout.Y)
    let shifts = lazy (overlapShifts settings measureText measureRange styleContext pageArray.Value anchorTopOf (doc.Blocks |> Seq.choose (function TextBox textBox -> Some textBox | _ -> None)))
    let outerSink = drawingSink
    drawingSink <- drawings

    try
     for block in doc.Blocks do
        match block with
        | Table tableBlock ->
            let tableCommands, nextY =
                tableFragmentPaintCommandsForDocument settings measureText measureRange doc index hasMultiColumnLayout tableBlock
                |> Option.defaultWith (fun () -> tableBorderCommandsCore (Some styleContext) settings measureText measureRange tableBlock currentY)
            commands.AddRange(tableCommands)
            if not tableBlock.Floating then currentY <- nextY
            addTableTextCommands settings fontFamily measureText measureRange index commands tableBlock
            addNestedTableBorders (Some styleContext) settings measureText measureRange index commands tableBlock
        | TextBox textBox ->
            drawings.Add (textBoxDrawing settings fontFamily measureText measureRange styleContext pageArray.Value (anchorTopOf textBox) (shiftOf shifts.Value textBox) textBox)
        | _ ->
            match Rbt.Word.Document.tryGetParagraph block with
            | None -> ()
            | Some para ->
                match tryFirstBlockLayout para.Id index with
                | None -> ()
                | Some layout ->
                    match index.ByBlock.TryGetValue para.Id with
                    | true, paraLayouts ->
                        for paraLayout in paraLayouts do
                            addParagraphCommandsKept settings fontFamily measureText measureRange commands paraLayout
                    | false, _ ->
                        addParagraphCommandsKept settings fontFamily measureText measureRange commands layout
                    currentY <- layout.Y + layout.Height
                    if para.Metadata.PageBreakAfter || (para.Metadata.SectionBreak.IsSome && sectionPageBreaks.Value.Contains para.Id) then
                        currentY <- WordLayout.Page.advanceToNextPage settings.PageHeight settings.PaddingTop currentY

     addNoteAreaCommands settings fontFamily measureText measureRange doc commands index.Footnotes index.Endnotes
    finally drawingSink <- outerSink

    // Column separator rules (w:sep): a vertical line centred in each gap,
    // spanning the body content on each page.
    match doc.PageSetup.Columns with
    | Some columns when columns.Count >= 2 && columns.Separator ->
        let pageArr = Array.ofList pages
        let pageCount = max 1 pageArr.Length
        let n = columns.Count
        let full = settings.Width - settings.PaddingLeft - settings.MarginRight
        let colWidth = (full - float (n - 1) * columns.Gap) / float n
        for page in 0 .. pageCount - 1 do
            let info = pageArr.[page]
            let contentTop = info.Top + info.MarginTop
            let contentBottom = info.Top + info.Height - info.MarginBottom
            let pageBottom =
                match index.BodyLinesByPage.TryGetValue page with
                | false, _ -> contentTop
                | true, lines ->
                    let mutable bottom = contentTop
                    for line in lines do
                        if line.Y >= info.Top - 0.5 && line.Y < info.Top + info.Height then
                            bottom <- max bottom (line.Y + line.Height)
                    min contentBottom bottom
            if pageBottom > contentTop then
                for c in 0 .. n - 2 do
                    let x = settings.PaddingLeft + float (c + 1) * colWidth + float c * columns.Gap + columns.Gap / 2.0
                    commands.Add(StrokeLine ("#000000", 0.5, 0, 0.0, x, contentTop, x, pageBottom))
    | _ -> ()

    // The headers and footers lie below the body: a watermark is under the text.
    let furniture =
        if not doc.Headers.IsEmpty || not doc.Footers.IsEmpty || not doc.HeaderParts.IsEmpty || not doc.FooterParts.IsEmpty then
            headerFooterCommands settings fontFamily measureText measureRange doc layouts pages
        else []
    let commands = commandsInLayers commands drawings
    let commandCount = commands.Count + List.length furniture
    layouts, (if furniture.IsEmpty then resizeArrayToList commands else furniture @ resizeArrayToList commands), pages, commandCount

let collectDocumentCommandsWithLayouts (settings: LayoutSettings) fontFamily (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (cachedLayouts: BlockLayout list option) =
    let layouts, commands, pages, _ =
        collectDocumentCommandsWithLayoutsUncached settings fontFamily measureText measureRange doc cachedLayouts None
    layouts, commands, pages

/// Paint from an existing layout pass: its layouts and the page geometry it
/// produced (layoutDocumentPages / layoutDocumentPagesIncremental).
let collectDocumentCommandsWithLayoutsAndPages (settings: LayoutSettings) fontFamily (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (layouts: BlockLayout list) (pages: PageInfo list) =
    let layouts, commands, pages, _ =
        collectDocumentCommandsWithLayoutsUncached settings fontFamily measureText measureRange doc (Some layouts) (Some pages)
    layouts, commands, pages

let collectDocumentCommands settings fontFamily measureText measureRange doc =
    collectDocumentCommandsWithLayouts settings fontFamily measureText measureRange doc None

// ---------------------------------------------------------------------------
// Paint commands in chunks, kept from one paint to the next
// ---------------------------------------------------------------------------
// The layout pass hands over the body layouts block by block (LayoutSlices),
// and the layouts of blocks an edit did not touch or move are the objects of
// the pass before. A block's commands depend on the block, on its layouts and
// on what the whole collection is made with (settings, font, styles, the
// table plan scope); so a block with the same layouts keeps its commands, and
// their vertical extents with them - an edit collects and indexes the commands
// of what changed, not of the document.

/// The paint commands of one part of the document (a block, the note areas, a
/// page's header and footer) with their vertical extents.
type PaintChunk = {
    mutable Commands: PaintCommand array
    mutable Tops: float array
    mutable Bottoms: float array
    /// The extent of the whole chunk (infinite when a command's is unknown).
    /// Of a chunk whose commands are not collected yet: a band they are
    /// certain to lie in.
    mutable Top: float
    mutable Bottom: float
    /// The extent of the chunk's pictures (Top > Bottom when it has none).
    mutable ImageTop: float
    mutable ImageBottom: float
    /// Collects the commands of a chunk that has not been painted yet (see
    /// `forceChunk`): a paragraph's commands are collected when a band that
    /// reaches its place is painted, not when the frame is built - opening a
    /// long document, or painting its first page, does not collect the
    /// commands of every page.
    mutable Pending: (unit -> PaintCommand array) option
}

/// Commands with those of their groups in their place (for what looks at
/// every command: the pictures a document shows, its text).
let rec flattenedCommands (commands: PaintCommand list) : PaintCommand list =
    if commands |> List.exists (function DrawGroup _ | DrawMirrored _ -> true | _ -> false) then
        commands |> List.collect (function
            | DrawGroup (_, _, _, _, _, _, inner) -> flattenedCommands inner
            | DrawMirrored (_, _, _, _, _, inner) -> flattenedCommands inner
            | command -> [ command ])
    else commands

/// Whether commands (those in their groups too) draw a picture.
let rec commandsHoldPicture (commands: PaintCommand list) : bool =
    commands |> List.exists (function
        | DrawImage _ -> true
        | DrawGroup (_, _, _, _, _, _, inner) -> commandsHoldPicture inner
        | DrawMirrored (_, _, _, _, _, inner) -> commandsHoldPicture inner
        | _ -> false)

let private fillChunk (chunk: PaintChunk) (commands: PaintCommand array) =
    let tops = Array.zeroCreate<float> commands.Length
    let bottoms = Array.zeroCreate<float> commands.Length
    let mutable top = System.Double.PositiveInfinity
    let mutable bottom = System.Double.NegativeInfinity
    let mutable imageTop = System.Double.PositiveInfinity
    let mutable imageBottom = System.Double.NegativeInfinity
    for i = 0 to commands.Length - 1 do
        match commandYExtent commands.[i] with
        | Some (commandTop, commandBottom) ->
            tops.[i] <- commandTop
            bottoms.[i] <- commandBottom
            match commands.[i] with
            | DrawImage _ ->
                imageTop <- min imageTop commandTop
                imageBottom <- max imageBottom commandBottom
            | DrawGroup (_, _, _, _, _, _, inner)
            | DrawMirrored (_, _, _, _, _, inner) when commandsHoldPicture inner ->
                imageTop <- min imageTop commandTop
                imageBottom <- max imageBottom commandBottom
            | _ -> ()
        | None ->
            // No extent known: part of every band.
            tops.[i] <- System.Double.NegativeInfinity
            bottoms.[i] <- System.Double.PositiveInfinity
        top <- min top tops.[i]
        bottom <- max bottom bottoms.[i]
    chunk.Commands <- commands
    chunk.Tops <- tops
    chunk.Bottoms <- bottoms
    chunk.Top <- top
    chunk.Bottom <- bottom
    chunk.ImageTop <- imageTop
    chunk.ImageBottom <- imageBottom
    chunk.Pending <- None

let private blankChunk () : PaintChunk =
    { Commands = [||]; Tops = [||]; Bottoms = [||]; Top = System.Double.PositiveInfinity; Bottom = System.Double.NegativeInfinity
      ImageTop = System.Double.PositiveInfinity; ImageBottom = System.Double.NegativeInfinity; Pending = None }

let chunkOfCommands (commands: PaintCommand array) : PaintChunk =
    let chunk = blankChunk ()
    fillChunk chunk commands
    chunk

/// A chunk whose commands are collected when they are first wanted; until
/// then it stands for the band top .. bottom, which they lie in.
let private deferredChunk (top: float) (bottom: float) (collect: unit -> PaintCommand array) : PaintChunk =
    let chunk = blankChunk ()
    chunk.Top <- top
    chunk.Bottom <- bottom
    chunk.Pending <- Some collect
    chunk

/// The chunk with its commands (collected now when they were not yet).
let forceChunk (chunk: PaintChunk) : PaintChunk =
    match chunk.Pending with
    | Some collect ->
        chunk.Pending <- None
        fillChunk chunk (collect ())
        chunk
    | None -> chunk

let private emptyChunk = chunkOfCommands [||]

/// The commands of chunks, in order.
let commandsOfChunks (chunks: PaintChunk array) : PaintCommand list =
    let mutable result = []
    for chunkIndex = chunks.Length - 1 downto 0 do
        let commands = (forceChunk chunks.[chunkIndex]).Commands
        for i = commands.Length - 1 downto 0 do
            result <- commands.[i] :: result
    result

/// The band (top, bottom) in which what two chunk arrays paint differs: None
/// when they paint the same, Choice2Of2 when the band is not known (a command
/// without an extent). Chunks that are the same objects are not looked into.
let chunkDifferenceBand (before: PaintChunk array) (after: PaintChunk array) : Choice<(float * float) option, unit> =
    let limit = min before.Length after.Length
    let mutable prefix = 0
    while prefix < limit && obj.ReferenceEquals(before.[prefix], after.[prefix]) do
        prefix <- prefix + 1
    let mutable suffix = 0
    while suffix < limit - prefix && obj.ReferenceEquals(before.[before.Length - 1 - suffix], after.[after.Length - 1 - suffix]) do
        suffix <- suffix + 1
    let mutable top = System.Double.PositiveInfinity
    let mutable bottom = System.Double.NegativeInfinity
    let mutable unknown = false
    let add (commandTop: float) (commandBottom: float) =
        if System.Double.IsInfinity commandTop || System.Double.IsInfinity commandBottom then unknown <- true
        else
            top <- min top commandTop
            bottom <- max bottom commandBottom
    let addChunk (chunk: PaintChunk) =
        if chunk.Commands.Length > 0 || chunk.Pending.IsSome then add chunk.Top chunk.Bottom
    let beforeCount = before.Length - prefix - suffix
    let afterCount = after.Length - prefix - suffix
    if beforeCount = afterCount && beforeCount <= 64 then
        // The same parts, some with other commands: command by command.
        for offset in 0 .. beforeCount - 1 do
            let left = before.[prefix + offset]
            let right = after.[prefix + offset]
            if obj.ReferenceEquals(left, right) then ()
            elif left.Pending.IsSome then
                // What was there was never painted: the band both stand for.
                addChunk left
                addChunk right
            else
                let right = forceChunk right
                let shared = min left.Commands.Length right.Commands.Length
                for i in 0 .. shared - 1 do
                    if not (obj.ReferenceEquals(left.Commands.[i], right.Commands.[i])) && left.Commands.[i] <> right.Commands.[i] then
                        add left.Tops.[i] left.Bottoms.[i]
                        add right.Tops.[i] right.Bottoms.[i]
                for i in shared .. left.Commands.Length - 1 do
                    add left.Tops.[i] left.Bottoms.[i]
                for i in shared .. right.Commands.Length - 1 do
                    add right.Tops.[i] right.Bottoms.[i]
    else
        for offset in 0 .. beforeCount - 1 do
            addChunk before.[prefix + offset]
        for offset in 0 .. afterCount - 1 do
            addChunk after.[prefix + offset]
    if unknown then Choice2Of2 ()
    elif top > bottom then Choice1Of2 None
    else Choice1Of2 (Some (top, bottom))

/// An anchored drawing in a chunk, with its place among the drawings (see Drawing).
type private DrawingChunk = {
    ChunkBehind: bool
    ChunkZ: float
    ChunkOfDrawing: PaintChunk
}

let private noDrawingChunks : DrawingChunk array = [||]

type private BlockPaintEntry = {
    EntryBlock: Block
    /// The layouts the commands were collected from: EntryCount of EntryLayouts from EntryStart.
    EntryLayouts: BlockLayout array
    EntryStart: int
    EntryCount: int
    /// The running y a table without layouts was drawn at (NaN for every
    /// other block: its commands follow from its layouts alone), and the
    /// running y a table hands on when that is known (NaN: not known, or the
    /// block hands on what it was given - see `yAfter` in the collection).
    EntryYIn: float
    EntryYOut: float
    EntryChunk: PaintChunk
    /// The block's anchored drawings, each a chunk of its own (not in EntryChunk).
    EntryDrawings: DrawingChunk array
    /// Of a shape: the top of the block it is anchored in when it was drawn
    /// (NaN: none - the shape stands where its offsets say).
    EntryAnchorTop: float
}

/// A page's header and footer in a chunk, and what they were made from.
type private FurniturePage = {
    FurnitureSource: HeaderFooterPageSource
    FurnitureNumber: int
    FurnitureCount: int
    FurnitureFieldsKey: int
    FurnitureChunk: PaintChunk
}

/// What a paint keeps for the next one (one per render context).
type PaintCache = private {
    mutable CacheSettings: LayoutSettings option
    mutable CacheFontFamily: string
    mutable CacheStyles: NamedStyle list
    mutable CachePageSetup: PageSetup option
    mutable CacheScope: obj
    mutable CacheMultiColumn: bool
    mutable CacheBlocks: Block array
    mutable CacheEntries: BlockPaintEntry array
    mutable CacheNotes: BlockLayout list
    mutable CacheNoteChunk: PaintChunk
    mutable CacheHeaderFooter: HeaderFooterCache
    mutable CacheHeaderFooterChunks: PaintChunk array
    mutable CacheHeaderFooterCommands: PaintCommand array array
    /// The headers and footers page by page, their commands collected when a page is painted.
    mutable CacheFurniture: FurniturePage array
    /// Block by block (as CacheBlocks), whether a paragraph of the block ends a section.
    /// Empty when it is not known for the blocks of the cache.
    mutable CacheSectionEnds: bool array
    /// How many blocks' commands the last collection made anew.
    mutable CollectedBlocks: int
    /// The layout pass (LayoutSlices.SliceGeneration) the entries were collected from.
    mutable CacheGeneration: int
    /// How many of the chunks last handed out - the first ones - are the pages' headers and footers.
    mutable CacheFurnitureCount: int
}

let createPaintCache () : PaintCache = {
    CacheSettings = None
    CacheFontFamily = ""
    CacheStyles = []
    CachePageSetup = None
    CacheScope = null
    CacheMultiColumn = false
    CacheBlocks = [||]
    CacheEntries = [||]
    CacheNotes = []
    CacheNoteChunk = emptyChunk
    CacheHeaderFooter = createHeaderFooterCache ()
    CacheHeaderFooterChunks = [||]
    CacheHeaderFooterCommands = [||]
    CacheFurniture = [||]
    CacheSectionEnds = [||]
    CollectedBlocks = 0
    CacheGeneration = 0
    CacheFurnitureCount = 0
}

let clearPaintCache (cache: PaintCache) =
    cache.CacheGeneration <- 0
    cache.CacheSettings <- None
    cache.CacheBlocks <- [||]
    cache.CacheSectionEnds <- [||]
    cache.CacheEntries <- [||]
    cache.CacheNotes <- []
    cache.CacheNoteChunk <- emptyChunk
    cache.CacheHeaderFooter <- createHeaderFooterCache ()
    cache.CacheHeaderFooterChunks <- [||]
    cache.CacheHeaderFooterCommands <- [||]
    cache.CacheFurniture <- [||]

/// How many of the chunks the last `collectDocumentChunks` gave - the first
/// ones - are the pages' headers and footers (they lie below the body).
let paintCacheFurnitureCount (cache: PaintCache) = cache.CacheFurnitureCount

/// Blocks whose commands the last `collectDocumentChunks` made anew (the rest were kept).
let paintCacheCollectedBlocks (cache: PaintCache) = cache.CollectedBlocks

let private sameLayoutList (left: BlockLayout list) (right: BlockLayout list) =
    let mutable a = left
    let mutable b = right
    let mutable same = true
    while same && not a.IsEmpty && not b.IsEmpty do
        if obj.ReferenceEquals(a.Head, b.Head) then
            a <- a.Tail
            b <- b.Tail
        else same <- false
    same && a.IsEmpty && b.IsEmpty

/// The paint commands of a document in chunks - one per page for its header
/// and footer, one per anchored drawing behind the text, one per block, one
/// for the note areas, one per anchored drawing in front of the text - in
/// paint order (see Drawing), from
/// the layouts of a pass that knows them block by block. The commands are
/// those of `collectDocumentCommandsWithLayoutsAndPages`; with the cache of
/// the paint before, the chunks of what did not change are the same objects.
/// None when the document needs the whole-document collection (column
/// separator rules, which follow every line of a page).
let collectDocumentChunksWith (deferCollection: bool) (cache: PaintCache) (settings: LayoutSettings) fontFamily (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (slices: LayoutSlices) (pages: PageInfo list) : PaintChunk array option =
    let separators =
        match doc.PageSetup.Columns with
        | Some columns when columns.Count >= 2 && columns.Separator -> true
        | _ -> false
    if separators then None
    else
        let scope = currentTablePlanScope ()
        let hasMultiColumnLayout = documentHasMultiColumnLayout doc
        let usable =
            (match cache.CacheSettings with
             | Some known -> obj.ReferenceEquals(known, settings) || known = settings
             | None -> false)
            && cache.CacheFontFamily = fontFamily
            && obj.ReferenceEquals(cache.CacheStyles, doc.Styles)
            && (match cache.CachePageSetup with
                | Some known -> obj.ReferenceEquals(known, doc.PageSetup) || known = doc.PageSetup
                | None -> false)
            && obj.ReferenceEquals(cache.CacheScope, scope)
            && not (isNull scope)
            && cache.CacheMultiColumn = hasMultiColumnLayout
        if not usable then
            clearPaintCache cache
            cache.CacheSettings <- Some settings
            cache.CacheFontFamily <- fontFamily
            cache.CacheStyles <- doc.Styles
            cache.CachePageSetup <- Some doc.PageSetup
            cache.CacheScope <- scope
            cache.CacheMultiColumn <- hasMultiColumnLayout
        let blocks = slices.SliceBlocks
        let layouts = slices.SliceLayouts
        let struct (prefix, suffix) = commonBlockEnds cache.CacheBlocks blocks
        let shift = cache.CacheBlocks.Length - blocks.Length
        // (Which blocks end a section, as known for the blocks of the paint before.)
        let knownSectionEnds = if cache.CacheSectionEnds.Length = cache.CacheBlocks.Length then cache.CacheSectionEnds else [||]
        let styleContext = WordStyles.styleContext doc
        let entries = Array.zeroCreate<BlockPaintEntry> blocks.Length
        let chunks = ResizeArray<PaintChunk>(blocks.Length + 8)
        let furnitureChunks = ResizeArray<PaintChunk>()
        let drawingChunks = ResizeArray<DrawingChunk>()
        let mutable collected = 0
        // The blocks the layout pass says kept their layouts from the pass the
        // cache was collected from keep their entries without a look: the
        // collection is the work of the blocks in between.
        let trustedPrefix, trustedSuffix =
            if cache.CacheGeneration = 0 || cache.CacheEntries.Length <> cache.CacheBlocks.Length then 0, 0
            elif slices.SliceGeneration = cache.CacheGeneration && cache.CacheBlocks.Length = blocks.Length then blocks.Length, 0
            elif slices.SliceBase = cache.CacheGeneration then
                min prefix slices.SliceKeptPrefix, min suffix slices.SliceKeptSuffix
            else 0, 0
        let suffixStart = blocks.Length - trustedSuffix
        let sliceOf (blockIndex: int) =
            let start = slices.SliceStarts.[blockIndex]
            let count = slices.SliceStarts.[blockIndex + 1] - start
            struct (start, count)
        // The running y after a block: where a table that has no layouts of its
        // own is drawn (below what stands before it). Blocks with layouts are
        // drawn where their layouts are, so it is worked out only when such a
        // table asks - from the block before it back to the nearest block that
        // says where it ends.
        let yAfterKnown = Array.create blocks.Length System.Double.NaN
        let sectionPageBreaks = lazy (sectionPageBreakParagraphs doc)
        let rec yAfter (blockIndex: int) : float =
            if blockIndex < 0 then doc.PageSetup.MarginTop
            elif not (System.Double.IsNaN yAfterKnown.[blockIndex]) then yAfterKnown.[blockIndex]
            else
                let struct (start, count) = sliceOf blockIndex
                let y =
                    match blocks.[blockIndex] with
                    | Table tableBlock when tableBlock.Floating -> yAfter (blockIndex - 1)
                    | Table tableBlock ->
                        let sliceLayouts = [ for offset in 0 .. count - 1 -> layouts.[start + offset] ]
                        let index = buildLayoutIndex settings false count sliceLayouts
                        match tableFragmentPaintCommandsForDocument settings measureText measureRange doc index hasMultiColumnLayout tableBlock with
                        | Some (_, nextY) -> nextY
                        | None -> tableBorderCommandsCore (Some styleContext) settings measureText measureRange tableBlock (yAfter (blockIndex - 1)) |> snd
                    | TextBox _ -> yAfter (blockIndex - 1)
                    | block ->
                        match Rbt.Word.Document.tryGetParagraph block with
                        | Some para when count > 0 ->
                            let layout = layouts.[start]
                            let bottom = layout.Y + layout.Height
                            if para.Metadata.PageBreakAfter || (para.Metadata.SectionBreak.IsSome && sectionPageBreaks.Value.Contains para.Id) then
                                WordLayout.Page.advanceToNextPage settings.PageHeight settings.PaddingTop bottom
                            else bottom
                        | _ -> yAfter (blockIndex - 1)
                yAfterKnown.[blockIndex] <- y
                y
        // Where the block a shape is anchored in starts (NaN: it has no such block, or that block no layout).
        let blockIndexById =
            lazy (
                let byId = Dictionary<BlockId, int>(blocks.Length)
                for index in 0 .. blocks.Length - 1 do
                    match blocks.[index] with
                    | Paragraph para
                    | Heading (_, para) -> byId.[para.Id] <- index
                    | Table table -> byId.[table.Id] <- index
                    | TextBox _ -> ()
                byId)
        let anchorTopOf (textBox: TextBoxBlock) : float =
            match textBox.AnchorBlockId with
            | None -> System.Double.NaN
            | Some anchorId ->
                match blockIndexById.Value.TryGetValue anchorId with
                | true, index when slices.SliceStarts.[index + 1] > slices.SliceStarts.[index] -> layouts.[slices.SliceStarts.[index]].Y
                | _ -> System.Double.NaN
        let pageArray = lazy (List.toArray pages)
        // (How far the shapes that may not overlap are moved: see overlapShifts.)
        let shifts =
            lazy (
                let anchorTopOrNone (textBox: TextBoxBlock) = let top = anchorTopOf textBox in if System.Double.IsNaN top then None else Some top
                overlapShifts settings measureText measureRange styleContext pageArray.Value anchorTopOrNone (blocks |> Seq.choose (function TextBox textBox -> Some textBox | _ -> None)))
        let keepEntry (blockIndex: int) (entry: BlockPaintEntry) =
            entries.[blockIndex] <- entry
            chunks.Add entry.EntryChunk
            if entry.EntryDrawings.Length > 0 then drawingChunks.AddRange entry.EntryDrawings
            if not (System.Double.IsNaN entry.EntryYOut) then yAfterKnown.[blockIndex] <- entry.EntryYOut
        // (An entry drawn at the running y stands while that y is what it was.)
        let entryStands (blockIndex: int) (entry: BlockPaintEntry) =
            System.Double.IsNaN entry.EntryYIn || entry.EntryYIn = yAfter (blockIndex - 1)
        for blockIndex in 0 .. blocks.Length - 1 do
            // (A shape is drawn from where its anchor is, not from layouts of its own: it is looked at every time.)
            // (And how far it is moved to keep clear of another shape: a shape that is moved is drawn again.)
            let shapeShift = match blocks.[blockIndex] with TextBox textBox when textBox.AllowOverlap = Some false -> shiftOf shifts.Value textBox | _ -> 0.0
            let anchorTop = match blocks.[blockIndex] with TextBox textBox -> anchorTopOf textBox + shapeShift | _ -> System.Double.NaN
            let isShape = match blocks.[blockIndex] with TextBox _ -> true | _ -> false
            let anchorStands (entry: BlockPaintEntry) =
                not isShape || entry.EntryAnchorTop = anchorTop || (System.Double.IsNaN entry.EntryAnchorTop && System.Double.IsNaN anchorTop)
            if blockIndex < trustedPrefix && not isShape then keepEntry blockIndex cache.CacheEntries.[blockIndex]
            elif blockIndex >= suffixStart && not isShape && entryStands blockIndex cache.CacheEntries.[blockIndex + shift] then keepEntry blockIndex cache.CacheEntries.[blockIndex + shift]
            else
            let block = blocks.[blockIndex]
            let struct (start, count) = sliceOf blockIndex
            let known =
                let candidate =
                    if blockIndex < prefix then ValueSome cache.CacheEntries.[blockIndex]
                    elif blockIndex >= blocks.Length - suffix then ValueSome cache.CacheEntries.[blockIndex + shift]
                    else ValueNone
                match candidate with
                | ValueSome entry when entry.EntryCount = count && entryStands blockIndex entry && anchorStands entry ->
                    let mutable same = true
                    let mutable offset = 0
                    while same && offset < count do
                        if not (obj.ReferenceEquals(entry.EntryLayouts.[entry.EntryStart + offset], layouts.[start + offset])) then same <- false
                        offset <- offset + 1
                    if same then ValueSome entry else ValueNone
                | _ -> ValueNone
            match known with
            | ValueSome entry ->
                keepEntry blockIndex (if obj.ReferenceEquals(entry.EntryLayouts, layouts) && entry.EntryStart = start then entry else { entry with EntryLayouts = layouts; EntryStart = start })
            | ValueNone ->
                collected <- collected + 1
                // (The commands of a block that does not wait; most blocks do.)
                let mutable commands : ResizeArray<PaintCommand> = null
                let mutable deferred : PaintChunk option = None
                let mutable yIn = System.Double.NaN
                let mutable yOut = System.Double.NaN
                // (The block's anchored drawings, which are not painted with it.)
                let drawings = ResizeArray<Drawing>()
                // The band the block's layouts lie in, the tallest of their lines, and whether they hold text alone.
                let mutable plain = true
                let mutable top = System.Double.PositiveInfinity
                let mutable bottom = System.Double.NegativeInfinity
                let mutable tallest = settings.LineHeight
                for offset in 0 .. count - 1 do
                    let layout = layouts.[start + offset]
                    top <- min top layout.Y
                    bottom <- max bottom (layout.Y + layout.Height)
                    for line in layout.Lines do
                        top <- min top line.Y
                        bottom <- max bottom (line.Y + line.Height)
                        tallest <- max tallest line.Height
                        for segment in line.Segments do
                            match segment with
                            | TextSegment _ -> ()
                            | _ -> plain <- false
                match block with
                | Table tableBlock ->
                    // (The commands, the running y the table hands on, and the one it was drawn at when it has no layouts.)
                    let collectTable () =
                        let own = ResizeArray<PaintCommand>()
                        let sliceLayouts = [ for offset in 0 .. count - 1 -> layouts.[start + offset] ]
                        let index = buildLayoutIndex settings false count sliceLayouts
                        let struct (tableCommands, nextY, drawnAt) =
                            match tableFragmentPaintCommandsForDocument settings measureText measureRange doc index hasMultiColumnLayout tableBlock with
                            | Some (drawn, nextY) -> struct (drawn, nextY, System.Double.NaN)
                            | None ->
                                let given = yAfter (blockIndex - 1)
                                let drawn, nextY = tableBorderCommandsCore (Some styleContext) settings measureText measureRange tableBlock given
                                struct (drawn, nextY, given)
                        own.AddRange(tableCommands)
                        addTableTextCommands settings fontFamily measureText measureRange index own tableBlock
                        addNestedTableBorders (Some styleContext) settings measureText measureRange index own tableBlock
                        struct (own, nextY, drawnAt)
                    if deferCollection && count > 0 && plain && not tableBlock.Floating then
                        // A table of text in the flow: it is drawn within a page's height of its
                        // cells' paragraphs (a row can be taller than its text, and the plan of
                        // its rows on the pages can break a row the layout kept whole - never by
                        // more than a page). Its commands - that plan, the borders, the text -
                        // wait until a band that reaches it is painted.
                        let reach = max settings.PageHeight 400.0 + tallest * 2.0 + 64.0
                        let bandTop = top - reach
                        let bandBottom = bottom + reach
                        deferred <- Some (deferredChunk bandTop bandBottom (fun () -> withTablePlanScope scope (fun () -> (let struct (own, _, _) = collectTable () in own.ToArray()))))
                    else
                        let struct (own, nextY, drawnAt) = withDrawingSink drawings collectTable
                        commands <- own
                        yIn <- drawnAt
                        if not tableBlock.Floating then yOut <- nextY
                | TextBox textBox ->
                    drawings.Add (textBoxDrawing settings fontFamily measureText measureRange styleContext pageArray.Value (if System.Double.IsNaN anchorTop then None else Some (anchorTop - shapeShift)) shapeShift textBox)
                | _ ->
                    match Rbt.Word.Document.tryGetParagraph block with
                    | None -> ()
                    | Some _ ->
                        if count > 0 then
                            // A paragraph of text draws within its lines (and a little around
                            // them): its commands wait until a band that reaches it is painted.
                            // One with a picture or an equation, which can sit anywhere, does not wait.
                            if plain && deferCollection then
                                let margin = tallest * 2.0 + 64.0
                                deferred <-
                                    Some (deferredChunk (top - margin) (bottom + margin) (fun () ->
                                        let own = ResizeArray<PaintCommand>()
                                        for offset in 0 .. count - 1 do
                                            addParagraphCommandsKept settings fontFamily measureText measureRange own layouts.[start + offset]
                                        own.ToArray()))
                            else
                                let own = ResizeArray<PaintCommand>()
                                withDrawingSink drawings (fun () ->
                                    for offset in 0 .. count - 1 do
                                        addParagraphCommandsKept settings fontFamily measureText measureRange own layouts.[start + offset])
                                commands <- own
                let chunk =
                    match deferred with
                    | Some chunk -> chunk
                    | None -> if isNull commands || commands.Count = 0 then emptyChunk else chunkOfCommands (commands.ToArray())
                let ownDrawings =
                    if drawings.Count = 0 then noDrawingChunks
                    else
                        drawings.ToArray()
                        |> Array.map (fun drawing -> { ChunkBehind = drawing.DrawingBehind; ChunkZ = drawing.DrawingZ; ChunkOfDrawing = chunkOfCommands drawing.DrawingCommands })
                keepEntry blockIndex { EntryBlock = block; EntryLayouts = layouts; EntryStart = start; EntryCount = count; EntryYIn = yIn; EntryYOut = yOut; EntryChunk = chunk; EntryDrawings = ownDrawings; EntryAnchorTop = anchorTop }
        cache.CacheBlocks <- blocks
        cache.CacheEntries <- entries
        cache.CollectedBlocks <- collected
        cache.CacheGeneration <- slices.SliceGeneration

        // The note areas.
        if not (sameLayoutList cache.CacheNotes slices.SliceNotes) || (slices.SliceNotes.IsEmpty && cache.CacheNoteChunk.Commands.Length > 0) then
            let footnotes = ResizeArray<BlockLayout>()
            let endnotes = ResizeArray<BlockLayout>()
            for layout in slices.SliceNotes do
                if layout.BlockId.StartsWith footnoteBlockIdPrefix then footnotes.Add layout
                elif layout.BlockId.StartsWith endnoteBlockIdPrefix then endnotes.Add layout
            let commands = ResizeArray<PaintCommand>()
            addNoteAreaCommands settings fontFamily measureText measureRange doc commands footnotes endnotes
            cache.CacheNotes <- slices.SliceNotes
            cache.CacheNoteChunk <- if commands.Count = 0 then emptyChunk else chunkOfCommands (commands.ToArray())
        chunks.Add cache.CacheNoteChunk

        // Headers and footers, page by page.
        if not doc.Headers.IsEmpty || not doc.Footers.IsEmpty || not doc.HeaderParts.IsEmpty || not doc.FooterParts.IsEmpty then
            // (Of the layouts, the headers and footers follow those that end a section.)
            // (Body paragraphs: looked for block by block, not layout by layout.)
            // (Whether a block ends a section is a matter of the block: what the paint before found
            // stands for the blocks that are the same objects. Asking every paragraph and every
            // cell of every table again was a millisecond of each keystroke in a long document.)
            let sectionEnds = ResizeArray<BlockLayout>()
            let sectionEndBlocks = Array.zeroCreate<bool> blocks.Length
            for blockIndex in 0 .. blocks.Length - 1 do
                let known =
                    if knownSectionEnds.Length = 0 then ValueNone
                    elif blockIndex < prefix then ValueSome knownSectionEnds.[blockIndex]
                    elif blockIndex >= blocks.Length - suffix then ValueSome knownSectionEnds.[blockIndex + shift]
                    else ValueNone
                let ends =
                    match known with
                    | ValueSome ends -> ends
                    | ValueNone ->
                        match blocks.[blockIndex] with
                        | Paragraph para
                        | Heading (_, para) -> endsSection para.Metadata
                        | _ ->
                            // (A table cell's or a text box's paragraph could carry section properties in a broken file: as before.)
                            let mutable found = false
                            for at in slices.SliceStarts.[blockIndex] .. slices.SliceStarts.[blockIndex + 1] - 1 do
                                if not found && endsSection layouts.[at].Metadata then found <- true
                            found
                sectionEndBlocks.[blockIndex] <- ends
                if ends then
                    for at in slices.SliceStarts.[blockIndex] .. slices.SliceStarts.[blockIndex + 1] - 1 do
                        if endsSection layouts.[at].Metadata then sectionEnds.Add layouts.[at]
            cache.CacheSectionEnds <- sectionEndBlocks
            if deferCollection then
                // What a page's header and footer are made from is looked at for every page;
                // their commands (the layout of their blocks, with the page's number) are
                // made when the page is painted.
                let pageTotal, sourceOf, commandsOf, valuesOf = headerFooterPageSources (Some cache.CacheHeaderFooter) settings fontFamily measureText measureRange doc sectionEnds pages
                let pagesArray = List.toArray pages
                let furniture = Array.zeroCreate<FurniturePage> pageTotal
                for pageIndex in 0 .. pageTotal - 1 do
                    let source = sourceOf pageIndex
                    let number = if source.SourceRunning then (valuesOf pageIndex).FieldPage else 0
                    let count = if source.SourceRunning then pageTotal else 0
                    let fieldsKey = if source.SourceRunning then fieldValuesKey (valuesOf pageIndex) else 0
                    let known =
                        if pageIndex < cache.CacheFurniture.Length then
                            let entry = cache.CacheFurniture.[pageIndex]
                            if obj.ReferenceEquals(entry.FurnitureSource.SourceHeaders, source.SourceHeaders) && obj.ReferenceEquals(entry.FurnitureSource.SourceFooters, source.SourceFooters)
                               && entry.FurnitureSource.SourceHeaderY = source.SourceHeaderY && entry.FurnitureSource.SourceFooterY = source.SourceFooterY
                               && entry.FurnitureNumber = number && entry.FurnitureCount = count && entry.FurnitureFieldsKey = fieldsKey then ValueSome entry
                            else ValueNone
                        else ValueNone
                    furniture.[pageIndex] <-
                        match known with
                        | ValueSome entry -> entry
                        | ValueNone ->
                            let chunk =
                                if source.SourceHeaders.IsEmpty && source.SourceFooters.IsEmpty then emptyChunk
                                else
                                    let page = pagesArray.[pageIndex]
                                    deferredChunk page.Top (page.Top + page.Height) (fun () -> withTablePlanScope scope (fun () -> commandsOf pageIndex source))
                            { FurnitureSource = source; FurnitureNumber = number; FurnitureCount = count; FurnitureFieldsKey = fieldsKey; FurnitureChunk = chunk }
                    furnitureChunks.Add furniture.[pageIndex].FurnitureChunk
                cache.CacheFurniture <- furniture
            else
            let byPage = headerFooterCommandsByPage (Some cache.CacheHeaderFooter) settings fontFamily measureText measureRange doc sectionEnds pages
            let pageChunks = Array.zeroCreate<PaintChunk> byPage.Length
            for pageIndex in 0 .. byPage.Length - 1 do
                pageChunks.[pageIndex] <-
                    if pageIndex < cache.CacheHeaderFooterCommands.Length && obj.ReferenceEquals(cache.CacheHeaderFooterCommands.[pageIndex], byPage.[pageIndex]) then cache.CacheHeaderFooterChunks.[pageIndex]
                    elif byPage.[pageIndex].Length = 0 then emptyChunk
                    else chunkOfCommands byPage.[pageIndex]
            cache.CacheHeaderFooterCommands <- byPage
            cache.CacheHeaderFooterChunks <- pageChunks
            furnitureChunks.AddRange pageChunks
        else
            // (Not looked at: nothing is known for these blocks.)
            cache.CacheSectionEnds <- [||]
        // In paint order: the headers and footers, the drawings behind the text, the text, the drawings in front of it.
        cache.CacheFurnitureCount <- furnitureChunks.Count
        if furnitureChunks.Count = 0 && drawingChunks.Count = 0 then Some (chunks.ToArray())
        else
            let layer (behind: bool) =
                drawingChunks
                |> Seq.filter (fun drawing -> drawing.ChunkBehind = behind)
                |> Seq.toList
                |> List.sortBy (fun drawing -> drawing.ChunkZ)
            let ordered = ResizeArray<PaintChunk>(furnitureChunks.Count + chunks.Count + drawingChunks.Count)
            ordered.AddRange furnitureChunks
            for drawing in layer true do ordered.Add drawing.ChunkOfDrawing
            ordered.AddRange chunks
            for drawing in layer false do ordered.Add drawing.ChunkOfDrawing
            Some (ordered.ToArray())

/// `collectDocumentChunksWith`, the commands of plain paragraphs collected
/// when they are first painted (see `PaintChunk.Pending`).
let collectDocumentChunks (cache: PaintCache) (settings: LayoutSettings) fontFamily (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (slices: LayoutSlices) (pages: PageInfo list) : PaintChunk array option =
    collectDocumentChunksWith true cache settings fontFamily measureText measureRange doc slices pages

/// What of a paragraph decides the faces its text and its list marker are
/// set in: its style, its list and the formatting of its runs. (The text
/// itself, and everything about spacing and position, has no part in it.)
let private fontFaceSignature (kind: string) (para: ParagraphBlock) : string option =
    let metadata = para.Metadata
    // A paragraph that ends a section stays (the section's headers and footers hang on it);
    // so does one with a picture, an equation or a field that is not a page number.
    if metadata.SectionBreak.IsSome then None
    else
        let sb = System.Text.StringBuilder()
        let text (value: string option) = match value with Some v -> v | None -> ""
        sb.Append(kind).Append("\u0001").Append(text metadata.Style).Append("\u0001")
          .Append(string metadata.ListStyle).Append("\u0001").Append(string metadata.ListLevel).Append("\u0001")
          .Append(text metadata.ListId).Append("\u0001").Append(text metadata.ListNumberFormat).Append("\u0001")
          .Append(text metadata.ListLevelText).Append("\u0001").Append(if metadata.HasMarkProperties then "m" else "") |> ignore
        let mutable plain = true
        let mutable last = ""
        for run in para.Buffer.Runs do
            match run with
            | Image _ -> plain <- false
            | Text run ->
                let style = run.Style
                if style.EquationXml.IsSome then plain <- false
                else
                    let field =
                        match style.FieldCode with
                        | Some code ->
                            let code = code.Trim()
                            let space = code.IndexOf ' '
                            if space < 0 then code else code.Substring(0, space)
                        | None -> ""
                    let own =
                        text style.Style + "\u0003" + text style.FontFamily + "\u0003" + (if style.Bold then "b" else "") + (if style.Italic then "i" else "")
                        + "\u0003" + string style.Explicit + "\u0003" + text style.VertAlign + "\u0003" + text style.NoteType + (if style.NoteId.IsSome then "n" else "")
                        + "\u0003" + field + (if style.Hidden then "h" else "") + (if style.AllCaps then "c" else "") + (if style.SmallCaps then "s" else "")
                    if own <> last then
                        sb.Append("\u0002").Append(own) |> ignore
                        last <- own
        if plain then Some (sb.ToString()) else None

/// A document with the blocks of `doc` that can bring a font face of their
/// own: of the body's paragraphs and tables that are alike in what decides
/// their faces (see fontFaceSignature) the first stays. The blocks of every
/// header and footer are added to the body, so their faces are found whatever
/// pages the sample has. A long document's faces are those of this sample,
/// which is laid out in a fraction of the time.
let private fontFaceSample (doc: Document) : Document =
    let seen = HashSet<string>()
    let rec tableSignature (sb: System.Text.StringBuilder) (table: TableBlock) : bool =
        let mutable plain = true
        sb.Append("T").Append(match table.StyleId with Some id -> id | None -> "") |> ignore
        for row in table.Rows do
            sb.Append("\u0004") |> ignore
            for cell in row.Cells do
                sb.Append("\u0005") |> ignore
                for para in cell.Content do
                    match fontFaceSignature "c" para with
                    | Some signature -> sb.Append(signature).Append("\u0006") |> ignore
                    | None -> plain <- false
                for nested in cell.NestedTables do
                    if not (tableSignature sb nested) then plain <- false
        plain
    let keep (block: Block) =
        match block with
        | Paragraph para ->
            match fontFaceSignature "p" para with
            | Some signature -> seen.Add signature
            | None -> true
        | Heading (level, para) ->
            match fontFaceSignature ("h" + string level) para with
            | Some signature -> seen.Add signature
            | None -> true
        | Table table ->
            let sb = System.Text.StringBuilder()
            if tableSignature sb table then seen.Add (sb.ToString()) else true
        | TextBox _ -> true
    let body = doc.Blocks |> List.filter keep
    let furniture =
        [ for part in doc.HeaderParts @ doc.FooterParts do
            for block in part.Blocks do
                // (Part of the body here: without what makes a paragraph end a section.)
                match block with
                | Paragraph para when para.Metadata.SectionBreak.IsSome -> ()
                | Heading (_, para) when para.Metadata.SectionBreak.IsSome -> ()
                | _ -> if keep block then yield block ]
    // (Before the body's last block: the body's final paragraph stays the final one.)
    let blocks =
        match List.rev body with
        | last :: before -> List.rev before @ furniture @ [ last ]
        | [] -> furniture
    { doc with Blocks = blocks; Selection = None }

/// The font faces a document's text is set in, after styles are resolved:
/// (family, bold, italic) of every run of the body, of table cells, text
/// boxes, footnotes and endnotes, of headers and footers, and of what the
/// engine itself sets in a face of its own - a list marker (the "o" bullet in
/// Courier New, symbol bullets in Symbol / Wingdings), a note's text. Text
/// without a family of its own is in `defaultFamily`. A host fetches exactly
/// these faces before it measures and paints the document.
/// `known`: the layouts and pages of a pass over the document, when there is
/// one; otherwise a sample of the document is laid out here (the faces do not
/// depend on what the measurers say, nor on where a block stands).
let requiredFontFaces (settings: LayoutSettings) (defaultFamily: string) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (known: (BlockLayout list * PageInfo list) option) : (string * bool * bool) list =
    let faces = HashSet<string * bool * bool>()
    let add (family: string option) (bold: bool) (italic: bool) =
        let family = family |> Option.filter (fun name -> name.Trim() <> "") |> Option.defaultValue defaultFamily
        faces.Add ((family, bold, italic)) |> ignore
    let doc, layouts, pages =
        match known with
        | Some (layouts, pages) -> doc, layouts, pages
        | None ->
            // No layout to read the faces from: one is made, of the blocks that can bring a face of their own.
            let reduced = fontFaceSample doc
            let layouts, pages = layoutDocumentPages settings measureText measureRange reduced
            reduced, layouts, pages
    // The lines of the body, of table cells and of the note areas: every run, whitespace too
    // (it is measured in its face), and the list markers.
    for layout in layouts do
        match layout.PrefixStyle with
        | Some style when layout.Prefix <> "" -> add style.FontFamily style.Bold style.Italic
        | _ -> ()
        for line in layout.Lines do
            for segment in line.Segments do
                match segment with
                | TextSegment (_, style, _, _, _) -> add style.FontFamily style.Bold style.Italic
                | MathSegment (_, style, _, _, _, _, _, _) -> add style.FontFamily style.Bold style.Italic
                | ImageSegment _ -> ()
    // What is painted besides: headers and footers, text boxes, markers.
    let _, commands, _ = collectDocumentCommandsWithLayoutsAndPages settings defaultFamily measureText measureRange doc layouts pages
    for command in commands do
        match command with
        | DrawText (_, style, _, _, _) -> add style.FontFamily style.Bold style.Italic
        | _ -> ()
    faces |> Seq.toList |> List.sort

let private layoutListLookup (layouts: BlockLayout list) =
    fun blockId ->
        layouts |> List.tryFind (fun l -> l.BlockId = blockId)

/// The first layout of a block, by id, from the layouts a pass handed over
/// block by block: the block is found by its place, the layout in its slice.
let sliceLayoutLookup (doc: Document) (slices: LayoutSlices) : BlockId -> BlockLayout option =
    fun blockId ->
        match locateBlock doc blockId with
        | ValueSome (struct (index, _)) when index < slices.SliceBlocks.Length ->
            let mutable found = None
            let mutable at = slices.SliceStarts.[index]
            let last = slices.SliceStarts.[index + 1]
            while found.IsNone && at < last do
                if slices.SliceLayouts.[at].BlockId = blockId then found <- Some slices.SliceLayouts.[at]
                at <- at + 1
            found
        | _ -> None

let private selectionCommandsForNormalizedRange
    (settings: LayoutSettings)
    (measureRange: TextRangeMeasurer)
    (sel: Selection)
    (startPos: Position)
    (endPos: Position)
    (blockIds: BlockId list)
    (tryFindLayout: BlockId -> BlockLayout option) =
    let effectiveBlockIds =
        if blockIds.IsEmpty && startPos.BlockId = endPos.BlockId then [startPos.BlockId]
        else blockIds

    let commands = ResizeArray<PaintCommand>()
    let blockCount = effectiveBlockIds.Length
    let mutable i = 0
    let mutable focusLayout = None
    let lineFullSelectionBounds (line: MeasuredLine) =
        let mutable hasSegment = false
        let mutable left = 0.0
        let mutable right = 0.0
        for segment in line.Segments do
            let x = segmentX segment
            let width = segmentWidth segment
            if not hasSegment then
                left <- x
                right <- x + width
                hasSegment <- true
            else
                left <- min left x
                right <- max right (x + width)
        if hasSegment then struct (left, right) else struct (0.0, line.Width)

    for blockId in effectiveBlockIds do
        match tryFindLayout blockId with
        | None -> ()
        | Some layout ->
            if blockId = sel.Focus.BlockId then
                focusLayout <- Some layout
            let bufferLen = maxOffsetInLayout layout

            let (highlightStart, highlightEnd) =
                if blockCount = 1 then
                    (startPos.Offset, endPos.Offset)
                elif i = 0 then
                    (startPos.Offset, bufferLen)
                elif i = blockCount - 1 then
                    (0, endPos.Offset)
                else
                    (0, bufferLen)

            for line in layout.Lines do
                let lineStart = max highlightStart line.StartOffset
                let lineEnd = min highlightEnd line.EndOffset
                if lineEnd > lineStart then
                    let mutable startX = 0.0
                    let mutable endX = 0.0
                    if lineStart <= line.StartOffset && lineEnd >= line.EndOffset then
                        let struct (left, right) = lineFullSelectionBounds line
                        startX <- left
                        endX <- right
                    else
                        startX <- xForOffsetWithRange measureRange line lineStart
                        endX <- xForOffsetWithRange measureRange line lineEnd
                    let lineAlignmentOffset = alignmentOffset layout line
                    commands.Add(FillRect ("rgba(0, 120, 215, 0.3)", settings.PaddingLeft + startX + lineAlignmentOffset, line.Y, endX - startX, line.Height))
        i <- i + 1

    match focusLayout |> Option.orElseWith (fun () -> tryFindLayout sel.Focus.BlockId) with
    | None -> resizeArrayToList commands
    | Some layout ->
        let line = lineForOffset layout sel.Focus.Offset
        let x = xForOffsetWithRange measureRange line sel.Focus.Offset
        let lineAlignmentOffset = alignmentOffset layout line
        commands.Add(FillRect ("#0066ff", settings.PaddingLeft + x + lineAlignmentOffset, line.Y, 2.0, line.Height))
        resizeArrayToList commands

let selectionCommandsWithRangeLookup (settings: LayoutSettings) (measureRange: TextRangeMeasurer) (doc: Document) (tryFindLayout: BlockId -> BlockLayout option) =
    match doc.Selection with
    | None -> []
    | Some sel when sel.IsCollapsed ->
        match tryFindLayout sel.Focus.BlockId with
        | None -> []
        | Some layout ->
            let line = lineForOffset layout sel.Focus.Offset
            let x = xForOffsetWithRange measureRange line sel.Focus.Offset
            let lineAlignmentOffset = alignmentOffset layout line
            [FillRect ("#0066ff", settings.PaddingLeft + x + lineAlignmentOffset, line.Y, 2.0, line.Height)]
    | Some sel ->
        let (startPos, endPos, blockIds) = getNormalizedRange sel doc
        selectionCommandsForNormalizedRange settings measureRange sel startPos endPos blockIds tryFindLayout

let selectionCommandsWithRange (settings: LayoutSettings) (measureRange: TextRangeMeasurer) (doc: Document) (layouts: BlockLayout list) =
    match doc.Selection with
    | None -> []
    | Some sel when sel.IsCollapsed ->
        let tryFindList = layoutListLookup layouts
        selectionCommandsWithRangeLookup settings measureRange doc tryFindList
    | Some sel ->
        let tryFindList = layoutListLookup layouts
        let startPos, endPos, blockIds = getNormalizedRange sel doc
        selectionCommandsForNormalizedRange settings measureRange sel startPos endPos blockIds tryFindList

let selectionCommands (settings: LayoutSettings) (measureText: TextMeasurer) (doc: Document) (layouts: BlockLayout list) =
    selectionCommandsWithRange settings (rangeMeasurerFromTextMeasurer measureText) doc layouts

/// The highlight of a selection that runs over several top-level blocks, for
/// the band top .. bottom of the document only: the blocks between the two
/// ends are found by their place, and of their layouts those that reach into
/// the band get their lines highlighted - a selection of the whole of a long
/// document costs what is on screen. With the caret. None when the two ends
/// are in one block (the caller's general path is cheap then).
let selectionCommandsInBand (settings: LayoutSettings) (measureRange: TextRangeMeasurer) (doc: Document) (slices: LayoutSlices) (bandTop: float) (bandBottom: float) : PaintCommand list option =
    match doc.Selection with
    | Some sel when not sel.IsCollapsed ->
        match locateBlock doc sel.Anchor.BlockId, locateBlock doc sel.Focus.BlockId with
        | ValueSome (struct (anchorIndex, _)), ValueSome (struct (focusIndex, _))
            // (An end beyond what is laid out - everything selected in a document that is being
            // laid out: the selection is shown as far as the layout goes. It was looked up block
            // by block in the whole document then, each block not found in a walk of all of them.)
            when anchorIndex <> focusIndex && min anchorIndex focusIndex < slices.SliceBlocks.Length ->
            let startPos, endPos, first, last =
                if anchorIndex < focusIndex then sel.Anchor, sel.Focus, anchorIndex, focusIndex
                else sel.Focus, sel.Anchor, focusIndex, anchorIndex
            let commands = ResizeArray<PaintCommand>()
            // The editable paragraphs of a block, in reading order.
            let idsOf (block: Block) =
                match block with
                | Paragraph p
                | Heading (_, p) -> [ p.Id ]
                | Table t -> [ for row in t.Rows do for cell in row.Cells do for para in cell.Content -> para.Id ]
                | TextBox t -> t.Content |> List.map (fun para -> para.Id)
            let highlight (layout: BlockLayout) (fromOffset: int) (toOffset: int) =
                for line in layout.Lines do
                    if line.Y + line.Height >= bandTop && line.Y <= bandBottom then
                        let lineStart = max fromOffset line.StartOffset
                        let lineEnd = min toOffset line.EndOffset
                        if lineEnd > lineStart then
                            let mutable startX = 0.0
                            let mutable endX = 0.0
                            if lineStart <= line.StartOffset && lineEnd >= line.EndOffset then
                                let mutable hasSegment = false
                                for segment in line.Segments do
                                    let x = segmentX segment
                                    let width = segmentWidth segment
                                    if not hasSegment then
                                        startX <- x
                                        endX <- x + width
                                        hasSegment <- true
                                    else
                                        startX <- min startX x
                                        endX <- max endX (x + width)
                                if not hasSegment then endX <- line.Width
                            else
                                startX <- xForOffsetWithRange measureRange line lineStart
                                endX <- xForOffsetWithRange measureRange line lineEnd
                            commands.Add(FillRect ("rgba(0, 120, 215, 0.3)", settings.PaddingLeft + startX + alignmentOffset layout line, line.Y, endX - startX, line.Height))
            for index in first .. min last (slices.SliceBlocks.Length - 1) do
                let from = slices.SliceStarts.[index]
                let until = slices.SliceStarts.[index + 1]
                // (Nothing of the block in the band: nothing to look at.)
                let mutable inBand = false
                let mutable at = from
                while not inBand && at < until do
                    let layout = slices.SliceLayouts.[at]
                    if layout.Y + layout.Height >= bandTop && layout.Y <= bandBottom then inBand <- true
                    at <- at + 1
                if inBand then
                    if index <> first && index <> last then
                        for at in from .. until - 1 do
                            let layout = slices.SliceLayouts.[at]
                            highlight layout 0 (maxOffsetInLayout layout)
                    else
                        // The block of an end: its paragraphs from the start position on / up to the end position.
                        let ids = idsOf slices.SliceBlocks.[index] |> List.toArray
                        let place (id: BlockId) = System.Array.IndexOf(ids, id)
                        let startAt = if index = first then place startPos.BlockId else -1
                        let endAt = if index = last then place endPos.BlockId else System.Int32.MaxValue
                        for at in from .. until - 1 do
                            let layout = slices.SliceLayouts.[at]
                            let own = place layout.BlockId
                            if own >= startAt && own <= endAt then
                                let fromOffset = if index = first && own = startAt then startPos.Offset else 0
                                let toOffset = if index = last && own = endAt then endPos.Offset else maxOffsetInLayout layout
                                highlight layout fromOffset toOffset
            match sliceLayoutLookup doc slices sel.Focus.BlockId with
            | Some layout ->
                let line = lineForOffset layout sel.Focus.Offset
                if line.Y + line.Height >= bandTop && line.Y <= bandBottom then
                    let x = xForOffsetWithRange measureRange line sel.Focus.Offset
                    commands.Add(FillRect ("#0066ff", settings.PaddingLeft + x + alignmentOffset layout line, line.Y, 2.0, line.Height))
            | None -> ()
            Some (resizeArrayToList commands)
        | _ -> None
    | _ -> None

let hoverCommands (settings: LayoutSettings) (layouts: BlockLayout list) (hoveredImage: (BlockId * int) option) =
    match hoveredImage with
    | None -> []
    | Some (blockId, offset) ->
        match layoutListLookup layouts blockId |> Option.bind (fun layout -> imageBoxInLayout settings.PaddingLeft layout offset) with
        | None -> []
        | Some (imageX, imageY, imageWidth, imageHeight) ->
            let handleSize = 8.0
            [
                StrokeRect ("#0078d4", 2.0, imageX, imageY, imageWidth, imageHeight)
                FillRect ("#0078d4", imageX - handleSize / 2.0, imageY - handleSize / 2.0, handleSize, handleSize)
                FillRect ("#0078d4", imageX + imageWidth - handleSize / 2.0, imageY - handleSize / 2.0, handleSize, handleSize)
                FillRect ("#0078d4", imageX - handleSize / 2.0, imageY + imageHeight - handleSize / 2.0, handleSize, handleSize)
                FillRect ("#0078d4", imageX + imageWidth - handleSize / 2.0, imageY + imageHeight - handleSize / 2.0, handleSize, handleSize)
            ]

let hoverCommandsLookup (settings: LayoutSettings) (tryFindLayout: BlockId -> BlockLayout option) (hoveredImage: (BlockId * int) option) =
    match hoveredImage with
    | None -> []
    | Some (blockId, offset) ->
        match tryFindLayout blockId |> Option.bind (fun layout -> imageBoxInLayout settings.PaddingLeft layout offset) with
        | None -> []
        | Some (imageX, imageY, imageWidth, imageHeight) ->
            let handleSize = 8.0
            [
                StrokeRect ("#0078d4", 2.0, imageX, imageY, imageWidth, imageHeight)
                FillRect ("#0078d4", imageX - handleSize / 2.0, imageY - handleSize / 2.0, handleSize, handleSize)
                FillRect ("#0078d4", imageX + imageWidth - handleSize / 2.0, imageY - handleSize / 2.0, handleSize, handleSize)
                FillRect ("#0078d4", imageX - handleSize / 2.0, imageY + imageHeight - handleSize / 2.0, handleSize, handleSize)
                FillRect ("#0078d4", imageX + imageWidth - handleSize / 2.0, imageY + imageHeight - handleSize / 2.0, handleSize, handleSize)
            ]
