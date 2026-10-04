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
    Underline = style.Underline || style.Hyperlink.IsSome || WordLayout.Text.isInsertedRevision style
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
    match RedBat.BorderGrid.styleKey style with
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

let private borderJoinEpsilon = 0.01

let private segmentIsHorizontal (segment: RedBat.BorderGrid.Segment) =
    abs (segment.Y1 - segment.Y2) < borderJoinEpsilon

let private segmentIsVertical (segment: RedBat.BorderGrid.Segment) =
    abs (segment.X1 - segment.X2) < borderJoinEpsilon

let private coversCoordinate value first second =
    value >= min first second - borderJoinEpsilon
    && value <= max first second + borderJoinEpsilon

type private TableBorderJoinIndex = {
    Horizontal: Dictionary<int, ResizeArray<RedBat.BorderGrid.Segment>>
    Vertical: Dictionary<int, ResizeArray<RedBat.BorderGrid.Segment>>
}

let private borderJoinCoordinateKey value =
    int (floor (value / borderJoinEpsilon + 0.5))

let private addBorderJoinSegment (index: Dictionary<int, ResizeArray<RedBat.BorderGrid.Segment>>) coordinate segment =
    let key = borderJoinCoordinateKey coordinate
    match index.TryGetValue key with
    | true, values -> values.Add segment
    | false, _ ->
        let values = ResizeArray()
        values.Add segment
        index.[key] <- values

let private tableBorderJoinIndex (segments: RedBat.BorderGrid.Segment list) =
    let horizontal = Dictionary<int, ResizeArray<RedBat.BorderGrid.Segment>>()
    let vertical = Dictionary<int, ResizeArray<RedBat.BorderGrid.Segment>>()
    for segment in segments do
        if segmentIsHorizontal segment then
            addBorderJoinSegment horizontal segment.Y1 segment
        elif segmentIsVertical segment then
            addBorderJoinSegment vertical segment.X1 segment
    { Horizontal = horizontal; Vertical = vertical }

let private indexedPerpendicularHalfWidth (segmentsByCoordinate: Dictionary<int, ResizeArray<RedBat.BorderGrid.Segment>>) coordinate covers =
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

let private tableBorderJoinCaps (index: TableBorderJoinIndex) (segment: RedBat.BorderGrid.Segment) =
    let perpendicularHalfWidth x y =
        if segmentIsHorizontal segment then
            indexedPerpendicularHalfWidth index.Vertical x (fun candidate ->
                abs (candidate.X1 - x) < borderJoinEpsilon
                && coversCoordinate y candidate.Y1 candidate.Y2)
        elif segmentIsVertical segment then
            indexedPerpendicularHalfWidth index.Horizontal y (fun candidate ->
                abs (candidate.Y1 - y) < borderJoinEpsilon
                && coversCoordinate x candidate.X1 candidate.X2)
        else
            0.0

    perpendicularHalfWidth segment.X1 segment.Y1, perpendicularHalfWidth segment.X2 segment.Y2

let private addStrokeTableBorderCommands (commands: ResizeArray<PaintCommand>) color width style x1 y1 x2 y2 startCap endCap =
    let width = max 0.25 width
    let dashStyle = dashStyleForTableBorderStyle style
    for offset, partWidth in RedBat.BorderGrid.strokePattern style do
        let offset = offset * width
        let partWidth = max 0.25 (partWidth * width)
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
    }: RedBat.BorderGrid.Stroke))

let private tableFragmentBorderCells (rowFragments: ResizeArray<TableCellFragmentPlan>) =
    let result = ResizeArray<RedBat.BorderGrid.Cell>(rowFragments.Count)
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
                Top = if fragmentStartsAfterRow then None else tableFragmentBorderGridStroke 1 cell.BorderTop
                Right = tableFragmentBorderGridStroke 1 cell.BorderRight
                Bottom = if fragmentEndsBeforeRow then None else tableFragmentBorderGridStroke 1 cell.BorderBottom
                Left = tableFragmentBorderGridStroke 1 cell.BorderLeft
            }
        }: RedBat.BorderGrid.Cell)
    result

let tableFragmentPaintCommands (plan: TableLayoutPlan) =
    let commands = ResizeArray<PaintCommand>()
    let tableSpacing = WordLayout.Table.cellSpacing WordLayout.canvasScale plan.Table
    let tablePaintX = plan.TableX - tableSpacing
    let tableBackgroundX = plan.TableX - tableSpacing * 2.0
    let cellFragments = tableCellPaintFragmentPlans plan
    let mutable remainingCellFragments = cellFragments

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

            let alreadyConsumedEmptyContinuation =
                fragment.LineCount = 0
                && fragment.ConsumedLineStart = fragment.ConsumedLineEnd
                && fragment.ConsumedLineStart > 0
            if not alreadyConsumedEmptyContinuation then
                tableFillCommands plan.Table.BackgroundColor plan.Table.Shading tableBackgroundX fragment.Y plan.TableWidth fragment.Height
                |> List.iter commands.Add

                if rowFragments.Count > 0 then
                    for cellFragment in rowFragments do
                        tableFillCommands cellFragment.CellBackgroundColor cellFragment.CellShading (cellFragment.CellX - cellFragment.CellSpacing) cellFragment.CellY cellFragment.CellWidth cellFragment.CellHeight
                        |> List.iter commands.Add

                    let borderSegments =
                        tableFragmentBorderCells rowFragments
                        |> RedBat.BorderGrid.resolveResizeArray
                    let joinIndex = tableBorderJoinIndex borderSegments

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
    match layout.Metadata.BackgroundColor with
    | Some color when color <> "" ->
        let blockX = settings.PaddingLeft + layout.Indent
        let blockWidth = max 1.0 layout.AvailableWidth
        commands.Add(FillRect (color, blockX, layout.Y, blockWidth, layout.Height))
    | _ -> ()

    match layout.Metadata.ParagraphBorders with
    | Some _ -> commands.AddRange(paragraphBorderCommands settings layout)
    | None -> ()

    // (A content control's frame is not part of the page: see markupMarkerCommands.)

    if layout.Prefix <> "" && not layout.Lines.IsEmpty then
        let firstLine = layout.Lines.Head
        let baseline = lineBaselineWithOffset baselineOffset settings.FontSize firstLine
        let firstLineIndent = WordLayout.Paragraph.firstLineIndent layout.Metadata
        let prefixX = settings.PaddingLeft + layout.Indent + min 0.0 firstLineIndent
        let prefixText = WordLayout.Numbering.displayMarkerText layout.Prefix
        let prefixStyle =
            if prefixText <> layout.Prefix then defaultPaintStyle settings.FontSize fontFamily
            else
                layout.PrefixStyle
                |> Option.map styleFromTextStyle
                |> Option.defaultValue (defaultPaintStyle settings.FontSize fontFamily)
        let prefixDrawWidth =
            if prefixText <> layout.Prefix then 0.0
            else
                layout.PrefixStyle
                |> Option.map (measureText prefixText)
                |> Option.defaultValue 0.0
        commands.Add(DrawText (prefixText, prefixStyle, prefixX, baseline, prefixDrawWidth))

    for line in layout.Lines do
        let baseline = lineBaselineWithOffset baselineOffset settings.FontSize line
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

                if style.Underline || style.Strike || style.DoubleStrike then
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
                    if style.Underline && inkWidth > 0.0 then
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

            | ImageSegment (src, x, yOffset, width, height, _, _, _, _, cropLeft, cropTop, cropRight, cropBottom) ->
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
                        commands.Add(FillRect ("#F3F4F6", imageX, imageY, width, height))
                        commands.Add(StrokeRect ("#9CA3AF", 1.0, imageX, imageY, width, height))
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
                            commands.Add(DrawText (fitted, labelStyle, imageX + (width - labelWidth) / 2.0, imageY + height / 2.0 + 13.0 * 0.35, labelWidth))
                    | _ -> ()
                else
                    commands.Add(DrawImage (src, imageX, imageY, width, height, cropLeft, cropTop, cropRight, cropBottom))

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
    elif optionEqualsIgnoreCase "right" alignment then max 0.0 (availableWidth - tableWidth)
    else 0.0

/// The content controls the selection is in: the controls of the paragraphs
/// its two ends are in.
let private selectedContentControls (doc: Document) : string list =
    match doc.Selection with
    | None -> []
    | Some selection ->
        [ selection.Anchor.BlockId; selection.Focus.BlockId ]
        |> List.distinct
        |> List.choose (fun blockId -> findParagraph doc blockId |> Option.bind (fun para -> para.Metadata.ContentControlId))
        |> List.distinct

/// The frame of a content control, as an editing aid: drawn around the
/// paragraphs of the control the caret or the selection is in, as Word does,
/// and nowhere else. It is an overlay: a document without a selection (a
/// viewer, a page rendered for print or preview) has none.
let markupMarkerCommands (settings: LayoutSettings) (doc: Document) (layouts: BlockLayout list) =
    match selectedContentControls doc with
    | [] -> []
    | controls ->
        let commands = ResizeArray<PaintCommand>()
        for layout in layouts do
            match layout.Metadata.ContentControlId with
            | Some controlId when List.contains controlId controls ->
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
        let outerLeft, outerRight = WordLayout.Table.outerCellInsets WordLayout.canvasScale tableBlock
        let tableAvailableWidth = max 0.0 availableWidth
        let tableMaxWidth = tableAvailableWidth + outerLeft + outerRight
        let resolvedTable = WordLayout.Table.resolveTable WordLayout.canvasScale (Some tableMaxWidth) tableBlock
        let columnWidths = resolvedTable.ColumnWidths
        let columnWidthPrefix = columnWidthPrefix columnWidths
        let columnMetrics = tableColumnMetrics tableMaxWidth columnWidths
        let fallbackColumnWidth = tableMaxWidth / float numColumns
        let tableWidth = WordLayout.Table.tableWidthWithSpacing columnWidths (WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock)
        let tableHorizontalOffset = baseHorizontalOffset + tableAlignmentOffset tableBlock.Alignment tableAvailableWidth tableWidth
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
            }: RedBat.BorderGrid.Stroke))

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

        let addRowCommands rowIndex (row: WordLayout.Table.ResolvedRow) currentY rowHeight =
            let cellSpacing = WordLayout.Table.rowCellSpacing WordLayout.canvasScale tableBlock row.Source
            let outerStartInset = outerLeft + cellSpacing
            for fragmentY, fragmentHeight in rowFragments settings currentY rowHeight do
                commands.AddRange(tableFillCommands tableBlock.BackgroundColor tableBlock.Shading tableX fragmentY tableWidth fragmentHeight)

            let cellBuffer = ResizeArray<RedBat.BorderGrid.Cell>(row.Cells.Length)
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
                    }: RedBat.BorderGrid.Cell)
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
                        Top = borderStroke 1 resolved.Borders.Top
                        Right = borderStroke 1 resolved.Borders.Right
                        // A merged cell cut by a page break is open at the break.
                        Bottom = (if contiguousRows rowIndex resolved.RowSpan < resolved.RowSpan then None else borderStroke 1 resolved.Borders.Bottom)
                        Left = borderStroke 1 resolved.Borders.Left
                    }
                }: RedBat.BorderGrid.Cell)
                let together = contiguousRows rowIndex resolved.RowSpan
                if together < resolved.RowSpan then
                    continuations.Add ((resolved, x, width, rowIndex + together, resolved.RowSpan - together))

            let borderSegments = RedBat.BorderGrid.resolveResizeArray cellBuffer
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
                        let clippedStartCap = if abs (y1 - segment.Y1) < borderJoinEpsilon then startCap else 0.0
                        let clippedEndCap = if abs (y2 - segment.Y2) < borderJoinEpsilon then endCap else 0.0
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
                                addStrokeTableBorderCommands commands color width b.Style x1 y1 x2 y2 0.0 0.0
                    paintOuterBorder overrideBorders.Top outerLeft outerTop outerRight outerTop
                    paintOuterBorder overrideBorders.Bottom outerLeft outerBottom outerRight outerBottom
                    paintOuterBorder overrideBorders.Left outerLeft outerTop outerLeft outerBottom
                    paintOuterBorder overrideBorders.Right outerRight outerTop outerRight outerBottom

        let headerRows = resolvedTable.HeaderRows
        let rows = resolvedTable.BodyRows

        let mutable rowIndex = 0
        for row in rows do
            let rowHeight = rowHeights.[rowIndex]
            let pagedY = paginateTableRow settings row.Source currentY rowHeight
            let pagedY =
                match givenTop rowIndex with
                | Some given when given > pagedY + 0.5 -> given
                | _ -> pagedY

            if pagedY <> currentY && not row.IsHeader && not headerRows.IsEmpty then
                currentY <- pagedY
                let mutable headerIndex = 0
                for headerRow in headerRows do
                    let headerHeight =
                        if headerIndex < rowHeights.Count then rowHeights.[headerIndex]
                        else rowHeightAt headerRow
                    addRowCommands headerIndex headerRow currentY headerHeight
                    currentY <- currentY + headerHeight
                    headerIndex <- headerIndex + 1
            else
                currentY <- pagedY

            addRowCommands rowIndex row currentY rowHeight
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

let private runningFieldResult pageNumber pageCount code currentText =
    match fieldCodeName code with
    | "PAGE" -> string pageNumber
    | "NUMPAGES" -> string pageCount
    | "DATE" -> System.DateTime.Now.ToString(fieldAtFormat code |> Option.defaultValue "yyyy-MM-dd")
    | "TIME" -> System.DateTime.Now.ToString(fieldAtFormat code |> Option.defaultValue "h:mm tt")
    | _ -> currentText

let private paragraphWithRunningFields pageNumber pageCount (para: ParagraphBlock) =
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
                let newText = runningFieldResult pageNumber pageCount code oldText
                if newText <> oldText then
                    let style = { textRun.Style with FieldResult = Some newText }
                    buffer <- buffer |> TextBuffer.delete start textRun.Length |> TextBuffer.insert start newText |> TextBuffer.applyStyle start newText.Length (fun _ -> style)
                    shift <- shift + newText.Length - textRun.Length
                    changed <- true
    if changed then { para with Buffer = buffer }, true else para, false

let private blockWithRunningFields pageNumber pageCount = function
    | Paragraph para ->
        let para, changed = paragraphWithRunningFields pageNumber pageCount para
        if changed then Paragraph para, true else Paragraph para, false
    | Heading (level, para) ->
        let para, changed = paragraphWithRunningFields pageNumber pageCount para
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
                                    |> mapListWithChange (paragraphWithRunningFields pageNumber pageCount)
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
            |> mapListWithChange (paragraphWithRunningFields pageNumber pageCount)
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

let private buildHeaderFooterPartIndex (parts: HeaderFooterPart list) =
    let byRelationshipId = Dictionary<string, HeaderFooterPart>()
    let byKind = Dictionary<string, HeaderFooterPart>()
    let mutable first = None
    for part in parts do
        if first.IsNone then
            first <- Some part
        match part.RelationshipId with
        | Some relationshipId when not (byRelationshipId.ContainsKey relationshipId) ->
            byRelationshipId.[relationshipId] <- part
        | _ -> ()
        if not (byKind.ContainsKey part.Kind) then
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
    remaining <- refs
    while found.IsNone && not remaining.IsEmpty do
        match remaining with
        | (kind, relationshipId) :: tail ->
            if kind = "default" then
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
let private buildSectionReferenceIndex (pages: PageInfo array) (layouts: BlockLayout list) (finalReferences: (string * string) list) isHeader =
    let endPages = ResizeArray<int>()
    let references = ResizeArray<(string * string) list>()
    let mutable effective : (string * string) list = []
    for layout in layouts do
        if endsSection layout.Metadata then
            let refs = if isHeader then layout.Metadata.HeaderReferences else layout.Metadata.FooterReferences
            effective <- inheritSectionReferences effective refs
            if not effective.IsEmpty then
                let endPage = pageIndexForY pages (layout.Y + layout.Height)
                let lastIndex = endPages.Count - 1
                if lastIndex < 0 || endPages.[lastIndex] <> endPage then
                    endPages.Add endPage
                    references.Add effective
    { EndPages = endPages.ToArray(); References = references.ToArray(); Final = inheritSectionReferences effective finalReferences }

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

let private blocksForPage (settings: LayoutSettings) (pageSetup: PageSetup) pageIndex fallbackBlocks (parts: HeaderFooterPartIndex) sectionRefs =
    if parts.Parts.IsEmpty then
        fallbackBlocks ()
    else
        let pageNumber = pageIndex + 1
        let preferredKind =
            if pageNumber = 1 && pageSetup.DifferentFirstPage then "first"
            elif pageNumber % 2 = 0 && pageSetup.EvenAndOddHeaders then "even"
            else "default"
        let refs = sectionRefsForPage sectionRefs pageIndex
        let part =
            partForReferences preferredKind refs parts
            |> Option.orElseWith (fun () -> tryFindPartByKind preferredKind parts)
            |> Option.orElseWith (fun () -> tryFindPartByKind "default" parts)
            |> Option.orElseWith (fun () -> parts.First)
        match part with
        | Some part -> part.Blocks
        | None -> fallbackBlocks ()

let private shiftCommandY dy command =
    match command with
    | DrawText (t, s, x, b, w) -> DrawText (t, s, x, b + dy, w)
    | DrawUnderline (c, s, x, y, w) -> DrawUnderline (c, s, x, y + dy, w)
    | DrawImage (s, x, y, w, h, cl, ct, cr, cb) -> DrawImage (s, x, y + dy, w, h, cl, ct, cr, cb)
    | FillRect (c, x, y, w, h) -> FillRect (c, x, y + dy, w, h)
    | StrokeRect (c, lw, x, y, w, h) -> StrokeRect (c, lw, x, y + dy, w, h)
    | StrokeLine (c, lw, ds, doff, x1, y1, x2, y2) -> StrokeLine (c, lw, ds, doff, x1, y1 + dy, x2, y2 + dy)
    | DrawShape (fill, stroke, sw, (bx, by, bw, bh), segs) ->
        DrawShape (fill, stroke, sw, (bx, by + dy, bw, bh), Shapes.translatePath 0.0 dy segs)

let private shiftCommandsY dy commands =
    if abs dy < 0.001 then
        commands
    else
        let shifted = ResizeArray<PaintCommand>()
        for command in commands do
            shifted.Add(shiftCommandY dy command)
        resizeArrayToList shifted

let private repeatedBlockCommands (settings: LayoutSettings) fontFamily measureText measureRange pageNumber pageCount startY blocks =
    let repeatedSettings = { settings with MarginBottom = 0.0 }
    let maxWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
    let commands = ResizeArray<PaintCommand>()
    let mutable y = startY
    for rawBlock in blocks do
        let block, _ = blockWithRunningFields pageNumber pageCount rawBlock
        match block with
        | Paragraph para
        | Heading (_, para) ->
            let safePara = { para with Metadata = { para.Metadata with PageBreakAfter = false } }
            let layout, _ = layoutParagraphAtWithRange repeatedSettings measureText measureRange maxWidth y 0.0 safePara Map.empty
            addParagraphCommands settings fontFamily measureText measureRange commands layout
            y <- layout.Y + layout.Height
        | Table table ->
            let tableCommands, nextY = tableBorderCommandsWithRange repeatedSettings measureText measureRange table y
            commands.AddRange(tableCommands)
            // The borders alone are not the table: lay its cells out as well
            // and paint their text. The table is laid out at the body origin
            // on an unpaged measuring page (a header/footer table never
            // splits) and the resulting commands are moved to y.
            let measureSettings = { repeatedSettings with PageHeight = max repeatedSettings.PageHeight 1000000.0 }
            let tableDoc = { Rbt.Word.Document.empty with Blocks = [ Table table ]; Selection = None }
            let cellCommands = ResizeArray<PaintCommand>()
            for layout in layoutDocumentWithRange measureSettings measureText measureRange tableDoc do
                if not (layout.BlockId.StartsWith footnoteBlockIdPrefix) && not (layout.BlockId.StartsWith endnoteBlockIdPrefix) then
                    addParagraphCommands settings fontFamily measureText measureRange cellCommands layout
            commands.AddRange(shiftCommandsY (y - measureSettings.PaddingTop) (resizeArrayToList cellCommands))
            y <- nextY
        | TextBox _ -> ()
    resizeArrayToList commands

let private headerFooterCommands (settings: LayoutSettings) fontFamily measureText measureRange (doc: Document) layouts (pages: PageInfo list) =
    let pageCount = pages.Length
    if doc.Headers.IsEmpty && doc.Footers.IsEmpty && doc.HeaderParts.IsEmpty && doc.FooterParts.IsEmpty then
        []
    else
        let mutable fallbackHeaders = ValueNone
        let mutable fallbackFooters = ValueNone
        let fallbackHeaderBlocks () =
            match fallbackHeaders with
            | ValueSome blocks -> blocks
            | ValueNone ->
                let blocks = doc.Headers |> List.map Paragraph
                fallbackHeaders <- ValueSome blocks
                blocks
        let fallbackFooterBlocks () =
            match fallbackFooters with
            | ValueSome blocks -> blocks
            | ValueNone ->
                let blocks = doc.Footers |> List.map Paragraph
                fallbackFooters <- ValueSome blocks
                blocks
        let pagesArray = List.toArray pages
        let headerParts = buildHeaderFooterPartIndex doc.HeaderParts
        let footerParts = buildHeaderFooterPartIndex doc.FooterParts
        let headerRefs =
            if doc.Headers.IsEmpty && doc.HeaderParts.IsEmpty then { EndPages = [||]; References = [||]; Final = [] }
            else buildSectionReferenceIndex pagesArray layouts doc.PageSetup.HeaderReferences true
        let footerRefs =
            if doc.Footers.IsEmpty && doc.FooterParts.IsEmpty then { EndPages = [||]; References = [||]; Final = [] }
            else buildSectionReferenceIndex pagesArray layouts doc.PageSetup.FooterReferences false
        let referenceListComparer =
            { new IEqualityComparer<Block list> with
                member _.Equals(left, right) = obj.ReferenceEquals(left, right)
                member _.GetHashCode(value) = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value) }
        let staticRepeatedCommands = Dictionary<Block list, PaintCommand list>(referenceListComparer)
        let repeatedCommands pageNumber startY blocks =
            if blocksHaveRunningFields blocks then
                repeatedBlockCommands settings fontFamily measureText measureRange pageNumber pageCount startY blocks
            else
                let mutable cached = Unchecked.defaultof<PaintCommand list>
                let originCommands =
                    if staticRepeatedCommands.TryGetValue(blocks, &cached) then
                        cached
                    else
                        let computed = repeatedBlockCommands settings fontFamily measureText measureRange pageNumber pageCount 0.0 blocks
                        staticRepeatedCommands.[blocks] <- computed
                        computed
                shiftCommandsY startY originCommands
        let commands = ResizeArray<PaintCommand>()
        for pageIndex = 0 to pagesArray.Length - 1 do
            let info = pagesArray.[pageIndex]
            let pageNumber = pageIndex + 1
            let headers = blocksForPage settings doc.PageSetup pageIndex fallbackHeaderBlocks headerParts headerRefs
            let footers = blocksForPage settings doc.PageSetup pageIndex fallbackFooterBlocks footerParts footerRefs
            if not headers.IsEmpty then
                let headerY = info.Top + max 8.0 doc.PageSetup.HeaderDistance
                commands.AddRange(repeatedCommands pageNumber headerY headers)
            if not footers.IsEmpty then
                let footerY = info.Top + info.Height - max 16.0 doc.PageSetup.FooterDistance - settings.LineHeight * 0.78
                commands.AddRange(repeatedCommands pageNumber footerY footers)
        resizeArrayToList commands

let private textBoxCommands (settings: LayoutSettings) fontFamily measureText measureRange (textBox: TextBoxBlock) =
    let floatingBox = WordLayout.Floating.textBoxCanvasBox settings.PaddingLeft settings.PaddingTop textBox
    let padding = 6.0
    let maxWidth = max 1.0 (textBox.Width - padding * 2.0)
    let mutable paraY = floatingBox.Y + padding
    let commands = ResizeArray<PaintCommand>()

    let useShape = (match textBox.Geometry with GeoRect -> false | _ -> true) || textBox.Gradient.IsSome
    if useShape then
        let placed =
            Shapes.geometryPath textBox.Geometry textBox.Width textBox.Height
            |> Shapes.translatePath floatingBox.X floatingBox.Y
        let segs =
            if abs textBox.Rotation < 0.01 then placed
            else placed |> Shapes.rotatePath (floatingBox.X + textBox.Width / 2.0) (floatingBox.Y + textBox.Height / 2.0) textBox.Rotation
        let fill =
            match textBox.Gradient with
            | Some g -> FillGradient g
            | None -> match textBox.FillColor with Some c -> FillSolid c | None -> FillNone
        let stroke = textBox.BorderColor
        commands.Add(DrawShape(fill, stroke, textBox.BorderWidth |> Option.defaultValue 1.0,
                               (floatingBox.X, floatingBox.Y, textBox.Width, textBox.Height), segs))
    else
        textBox.FillColor |> Option.iter (fun color -> commands.Add(FillRect (color, floatingBox.X, floatingBox.Y, textBox.Width, textBox.Height)))
        textBox.BorderColor |> Option.iter (fun color -> commands.Add(StrokeRect (color, textBox.BorderWidth |> Option.defaultValue 1.0, floatingBox.X, floatingBox.Y, textBox.Width, textBox.Height)))

    for para in textBox.Content do
        let layout, _ = layoutParagraphAtWithRange settings measureText measureRange maxWidth paraY (floatingBox.X + padding - settings.PaddingLeft) para Map.empty
        addParagraphCommands settings fontFamily measureText measureRange commands layout
        paraY <- layout.Y + layout.Height

    resizeArrayToList commands

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
                match tryLayout para.Id index with
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
        let tableLeftOffset = tableLeftOffsetForStyleContext (Some styleContext)
        let maxWidth = max 0.0 (settings.Width - settings.PaddingLeft - settings.MarginRight - tableLeftOffset)
        let outerLeft, outerRight = WordLayout.Table.outerCellInsets WordLayout.canvasScale tableBlock
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
        match plannedContentOffsetForBlock settings measureText measureRange doc tableBlock blockId
              |> Option.orElseWith (fun () -> plannedContentOffsetFromFullPlan settings measureText measureRange doc tableBlock) with
        | Some firstPlannedContentOffset -> Some (firstActualContentY - firstPlannedContentOffset)
        | None -> Some firstActualContentY

let private shiftCommandX dx command =
    match command with
    | DrawText (t, s, x, b, w) -> DrawText (t, s, x + dx, b, w)
    | DrawUnderline (c, s, x, y, w) -> DrawUnderline (c, s, x + dx, y, w)
    | DrawImage (s, x, y, w, h, cl, ct, cr, cb) -> DrawImage (s, x + dx, y, w, h, cl, ct, cr, cb)
    | FillRect (c, x, y, w, h) -> FillRect (c, x + dx, y, w, h)
    | StrokeRect (c, lw, x, y, w, h) -> StrokeRect (c, lw, x + dx, y, w, h)
    | StrokeLine (c, lw, ds, doff, x1, y1, x2, y2) -> StrokeLine (c, lw, ds, doff, x1 + dx, y1, x2 + dx, y2)
    | DrawShape (fill, stroke, sw, (bx, by, bw, bh), segs) ->
        DrawShape (fill, stroke, sw, (bx + dx, by, bw, bh), Shapes.translatePath dx 0.0 segs)

let private shiftCommandsX dx commands =
    let shifted = ResizeArray<PaintCommand>()
    for command in commands do
        shifted.Add(shiftCommandX dx command)
    resizeArrayToList shifted

let private sectionColumnsAreMulti = function
    | Some columns when columns.Count >= 2 -> true
    | _ -> false

let private documentHasMultiColumnLayout (doc: Document) =
    let mutable found = sectionColumnsAreMulti doc.PageSetup.Columns
    for block in doc.Blocks do
        if not found then
            match Rbt.Word.Document.tryGetParagraph block with
            | Some para -> found <- sectionColumnsAreMulti para.Metadata.SectionColumns
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

let private tableFragmentPaintCommandsForDocument settings measureText measureRange doc index hasMultiColumnLayout tableBlock =
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
    let styleContext = WordStyles.styleContext doc

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
        | TextBox textBox -> commands.AddRange(textBoxCommands settings fontFamily measureText measureRange textBox)
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
                            addParagraphCommands settings fontFamily measureText measureRange commands paraLayout
                    | false, _ ->
                        addParagraphCommands settings fontFamily measureText measureRange commands layout
                    currentY <- layout.Y + layout.Height
                    if para.Metadata.PageBreakAfter || para.Metadata.SectionBreak = Some "nextPage" then
                        currentY <- WordLayout.Page.advanceToNextPage settings.PageHeight settings.PaddingTop currentY

    // Footnote area: paint the synthetic footnote paragraphs appended by layout
    // at the bottom of each page, and a short separator rule above the topmost
    // note on each page (Word draws ~2in).
    for layout in index.Footnotes do
        addParagraphCommands settings fontFamily measureText measureRange commands layout
    if index.Footnotes.Count > 0 then
        let footnoteTopByPage = Dictionary<int, float>()
        for layout in index.Footnotes do
            let page = pageOfY settings layout.Y
            match footnoteTopByPage.TryGetValue page with
            | true, topY when topY <= layout.Y -> ()
            | _ -> footnoteTopByPage.[page] <- layout.Y
        for KeyValue(_, topY) in footnoteTopByPage do
            let separatorX = settings.PaddingLeft
            commands.Add(StrokeLine ("#000000", 0.75, 0, 0.0, separatorX, topY - 6.0, separatorX + 192.0, topY - 6.0))

    // Endnote area: paint the synthetic endnote paragraphs appended after the
    // body content, with a single separator rule above the first endnote.
    for layout in index.Endnotes do
        addParagraphCommands settings fontFamily measureText measureRange commands layout
    if index.Endnotes.Count > 0 then
        // Short (~2in) separator above the first endnote; Word draws a full-width
        // continuation separator on pages the endnotes spill onto.
        let endnoteTopByPage = Dictionary<int, float>()
        for layout in index.Endnotes do
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

    if not doc.Headers.IsEmpty || not doc.Footers.IsEmpty || not doc.HeaderParts.IsEmpty || not doc.FooterParts.IsEmpty then
        commands.AddRange(headerFooterCommands settings fontFamily measureText measureRange doc layouts pages)
    let commandCount = commands.Count
    layouts, resizeArrayToList commands, pages, commandCount

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

let private layoutListLookup (layouts: BlockLayout list) =
    fun blockId ->
        layouts |> List.tryFind (fun l -> l.BlockId = blockId)

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
