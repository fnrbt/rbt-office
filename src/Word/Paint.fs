module RedBat.Word.Paint

open RedBat.Word.Layout
open RedBat.Word.Document
open RedBat.Word.TextBuffer
open RedBat.Word.WordStyles

module Eq = RedBat.Word.Math

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
    match style |> Option.map (fun value -> value.ToLowerInvariant()) with
    | Some "dashed" -> 2
    | Some "dotted" -> 1
    | Some "dotdash" -> 3
    | _ -> 0

let private lineBaselineWithOffset baselineOffset (fallbackFontSize: float) (line: MeasuredLine) =
    let fontSize =
        line.Segments
        |> List.choose (function
            | TextSegment (_, style, _, _, _) -> Some (WordLayout.Text.printFontSize style)
            | ImageSegment _ -> None
            // Use the equation's box ascent so its top fits within the line and
            // its baseline aligns with surrounding text (math extends upward).
            | MathSegment (_, _, _, _, ascent, _, _, _) -> Some ascent)
        |> function
            | [] -> fallbackFontSize
            | sizes -> List.max sizes
    line.Y + fontSize + baselineOffset

let private lineBaseline fallbackFontSize line =
    lineBaselineWithOffset 0.0 fallbackFontSize line

let private tabLeaderText (measureText: TextMeasurer) (style: TextStyle) (width: float) (leader: string option) =
    match leader |> Option.map (fun value -> value.Trim().ToLowerInvariant()) with
    | Some "dot" when width > 0.0 ->
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
                    int (System.Math.Round((max 0.0 (width - inset)) / leaderPitch, System.MidpointRounding.AwayFromZero))
                else
                    int (System.Math.Round(width / dotAdvance, System.MidpointRounding.AwayFromZero))
            if count <= 0 then None
            else
                let text = String.replicate count "."
                let targetWidth =
                    if wideMeasuredDot then max 0.0 (width - inset)
                    else max 0.0 (measureText text style - inset)
                Some (text, targetWidth, inset)
    | _ -> None

let private isTabOnlySegment (text: string) =
    text.IndexOf('\t') >= 0 && text |> Seq.forall System.Char.IsWhiteSpace

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

let private patternKey (shading: Shading) =
    shading.Pattern |> Option.map (fun value -> value.ToLowerInvariant())

let private percentPatternCoverage (key: string option) =
    match key with
    | Some value when value.StartsWith("pct") ->
        match System.Double.TryParse(value.Substring(3)) with
        | true, percent -> Some (max 0.0 (min 1.0 (percent / 100.0)))
        | false, _ -> None
    | _ -> None

let private shadingBackgroundColor (shading: Shading) =
    match patternKey shading with
    | Some "nil" -> None
    | Some "solid" -> Some (usableColor "#000000" shading.ForegroundColor)
    | _ -> Some (usableColor "#FFFFFF" shading.BackgroundColor)

let private tryParseHexColor (value: string) =
    let value = value.Trim()
    let hex =
        if value.StartsWith("#") then value.Substring(1)
        else value
    if hex.Length = 6 then
        let hexValue ch =
            if ch >= '0' && ch <= '9' then Some (int ch - int '0')
            elif ch >= 'a' && ch <= 'f' then Some (10 + int ch - int 'a')
            elif ch >= 'A' && ch <= 'F' then Some (10 + int ch - int 'A')
            else None
        let parse start =
            match hexValue hex.[start], hexValue hex.[start + 1] with
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
    List.ofSeq commands

let private verticalStripeCommands color x y width height =
    let commands = ResizeArray<PaintCommand>()
    let spacing = 2.0
    let mutable xx = x + 0.5
    while xx <= x + width do
        commands.Add(StrokeLine (color, 1.0, 0, 0.0, xx, y, xx, y + height))
        xx <- xx + spacing
    List.ofSeq commands

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
    List.ofSeq commands

let private shadingRectCommands (shading: Shading) x y width height =
    if width <= 0.0 || height <= 0.0 then []
    else
        let key = patternKey shading
        let foreground = usableColor "#000000" shading.ForegroundColor
        let backgroundColor = shadingBackgroundColor shading
        let background = backgroundColor |> Option.map (fun color -> FillRect (color, x, y, width, height) |> List.singleton) |> Option.defaultValue []
        let percentCoverage = percentPatternCoverage key
        let pattern =
            match key with
            | Some "nil"
            | None
            | Some "clear"
            | Some "solid" -> []
            | Some "horzstripe" -> horizontalStripeCommands foreground x y width height
            | Some "vertstripe" -> verticalStripeCommands foreground x y width height
            | Some "diagstripe" -> diagonalStripeCommands false foreground x y width height
            | Some "reversediagstripe" -> diagonalStripeCommands true foreground x y width height
            | Some "horzcross" -> horizontalStripeCommands foreground x y width height @ verticalStripeCommands foreground x y width height
            | Some "diagcross" -> diagonalStripeCommands false foreground x y width height @ diagonalStripeCommands true foreground x y width height
            | key ->
                match percentCoverage with
                | Some coverage when coverage >= 1.0 -> fillRectClipped foreground x y width height
                | Some coverage when coverage > 0.0 ->
                    backgroundColor
                    |> Option.map (fun bg -> fillRectClipped (blendedColor bg foreground coverage) x y width height)
                    |> Option.defaultValue (fillRectClipped foreground x y width height)
                | _ -> []
        if percentCoverage.IsSome then pattern else background @ pattern

let private tableFillCommands backgroundColor shading x y width height =
    match shading with
    | Some value -> shadingRectCommands value x y width height
    | None ->
        match backgroundColor with
        | Some color -> [FillRect (color, x, y, width, height)]
        | None -> []

let private tableBorderStrokeParts width style =
    let width = max 0.25 width
    RedBat.BorderGrid.strokePattern style
    |> List.map (fun (offset, partWidth) -> offset * width, max 0.25 (partWidth * width))

let private borderJoinEpsilon = 0.01

let private segmentIsHorizontal (segment: RedBat.BorderGrid.Segment) =
    abs (segment.Y1 - segment.Y2) < borderJoinEpsilon

let private segmentIsVertical (segment: RedBat.BorderGrid.Segment) =
    abs (segment.X1 - segment.X2) < borderJoinEpsilon

let private coversCoordinate value first second =
    value >= min first second - borderJoinEpsilon
    && value <= max first second + borderJoinEpsilon

let private tableBorderJoinCaps (segments: RedBat.BorderGrid.Segment list) (segment: RedBat.BorderGrid.Segment) =
    let perpendicularHalfWidth x y =
        segments
        |> List.choose (fun candidate ->
            if segmentIsHorizontal segment
               && segmentIsVertical candidate
               && abs (candidate.X1 - x) < borderJoinEpsilon
               && coversCoordinate y candidate.Y1 candidate.Y2 then
                Some (candidate.Stroke.Width / 2.0)
            elif segmentIsVertical segment
                 && segmentIsHorizontal candidate
                 && abs (candidate.Y1 - y) < borderJoinEpsilon
                 && coversCoordinate x candidate.X1 candidate.X2 then
                Some (candidate.Stroke.Width / 2.0)
            else
                None)
        |> function
            | [] -> 0.0
            | values -> values |> List.max

    perpendicularHalfWidth segment.X1 segment.Y1, perpendicularHalfWidth segment.X2 segment.Y2

let private strokeTableBorderCommands color width style x1 y1 x2 y2 startCap endCap =
    let dashStyle = dashStyleForTableBorderStyle style
    tableBorderStrokeParts width style
    |> List.map (fun (offset, partWidth) ->
        if y1 = y2 then StrokeLine (color, partWidth, dashStyle, 0.0, x1 - startCap, y1 + offset, x2 + endCap, y2 + offset)
        else StrokeLine (color, partWidth, dashStyle, 0.0, x1 + offset, y1 - startCap, x2 + offset, y2 + endCap))

let tableFragmentPaintCommands (plan: TableLayoutPlan) =
    let commands = ResizeArray<PaintCommand>()
    let tableSpacing = WordLayout.Table.cellSpacing WordLayout.canvasScale plan.Table
    let tablePaintX = plan.TableX - tableSpacing
    let tableBackgroundX = plan.TableX - tableSpacing * 2.0
    let cellFragments = tableCellFragmentPlans plan
    let cellFragmentsByRowFragment =
        cellFragments
        |> List.groupBy (fun cell -> cell.RowIndex, cell.Page, cell.FragmentY, cell.FragmentHeight)
        |> Map.ofList

    for row in plan.Rows do
        for fragment in row.Fragments do
            let alreadyConsumedEmptyContinuation =
                fragment.LineCount = 0
                && fragment.ConsumedLineStart = fragment.ConsumedLineEnd
                && fragment.ConsumedLineStart > 0
            if not alreadyConsumedEmptyContinuation then
                tableFillCommands plan.Table.BackgroundColor plan.Table.Shading tableBackgroundX fragment.Y plan.TableWidth fragment.Height
                |> List.iter commands.Add

                let rowFragments =
                    cellFragmentsByRowFragment
                    |> Map.tryFind (row.RowIndex, fragment.Page, fragment.Y, fragment.Height)
                    |> Option.defaultValue []

                for cellFragment in rowFragments do
                    let sourceRow = plan.Table.Rows.[cellFragment.RowIndex]
                    sourceRow.Cells
                    |> List.tryItem cellFragment.CellIndex
                    |> Option.iter (fun cell ->
                        tableFillCommands cell.BackgroundColor cell.Shading (cellFragment.CellX - cellFragment.CellSpacing) cellFragment.CellY cellFragment.CellWidth cellFragment.CellHeight
                        |> List.iter commands.Add)

                let paintCells =
                    rowFragments
                    |> List.map (fun cell ->
                        let fragmentStartsAfterRow = cell.FragmentY > cell.RowY + 0.01
                        let fragmentEndsBeforeRow = cell.FragmentY + cell.FragmentHeight < cell.RowY + cell.RowHeight - 0.01
                        {
                            cell with
                                CellX = cell.CellX - cell.CellSpacing
                                BorderTop = if fragmentStartsAfterRow then None else cell.BorderTop
                                BorderBottom = if fragmentEndsBeforeRow then None else cell.BorderBottom
                        })
                let borderSegments =
                    tableCellFragmentBorderCells paintCells
                    |> RedBat.BorderGrid.resolve

                borderSegments
                |> List.iter (fun segment ->
                    let startCap, endCap = tableBorderJoinCaps borderSegments segment
                    strokeTableBorderCommands segment.Stroke.Color segment.Stroke.Width segment.Stroke.Style segment.X1 segment.Y1 segment.X2 segment.Y2 startCap endCap
                    |> List.iter commands.Add)

    let bottomSpacing = WordLayout.Table.cellSpacing WordLayout.canvasScale plan.Table
    if bottomSpacing > 0.0 then
        plan.Rows
        |> List.tryLast
        |> Option.bind (fun row -> row.Fragments |> List.tryLast)
        |> Option.iter (fun fragment ->
            tableFillCommands plan.Table.BackgroundColor plan.Table.Shading tableBackgroundX (fragment.Y + fragment.Height) plan.TableWidth bottomSpacing
            |> List.iter commands.Add)

    List.ofSeq commands

let private paragraphBorderCommands (settings: LayoutSettings) (layout: BlockLayout) =
    let borderLine (border: ParagraphBorder option) x1 y1 x2 y2 =
        border
        |> Option.bind (fun value ->
            match value.Style |> Option.map (fun style -> style.ToLowerInvariant()) with
            | Some "none"
            | Some "nil" -> None
            | _ ->
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
        let contentTop = layout.Lines |> List.tryHead |> Option.map (fun line -> line.Y) |> Option.defaultValue layout.Y
        let contentBottom = layout.Lines |> List.tryLast |> Option.map (fun line -> line.Y + line.Height) |> Option.defaultValue (layout.Y + layout.Height)
        let top = contentTop - topSpace
        let bottom = contentBottom + bottomSpace
        [ borderLine borders.Top left top right top
          borderLine borders.Right right top right bottom
          borderLine borders.Bottom left bottom right bottom
          borderLine borders.Left left top left bottom ]
        |> List.choose id

let private paginateTableRow settings (row: TableRow) y rowHeight =
    let canSplit = row.CanSplit |> Option.defaultValue true
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
    let remaining = max 0.0 (contentBottom - y)
    if canSplit && y < contentBottom && rowHeight > remaining then y else paginateBlock settings y rowHeight

let private rowFragments settings y height =
    let rowBottom = y + height
    let mutable fragments = []
    let mutable pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    while pageTop < rowBottom do
        let top = max y (WordLayout.Page.contentTop settings.PaddingTop pageTop)
        let bottom = min rowBottom (WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop)
        if bottom > top then fragments <- fragments @ [(top, bottom - top)]
        pageTop <- pageTop + pageHeight settings + pageGap
    fragments

let private paragraphCommandsWithBaselineOffset baselineOffset (settings: LayoutSettings) (fontFamily: string) (measureText: TextMeasurer) (layout: BlockLayout) =
    let mutable commands = []

    match layout.Metadata.BackgroundColor with
    | Some color when color <> "" ->
        let blockX = settings.PaddingLeft + layout.Indent
        let blockWidth = max 1.0 layout.AvailableWidth
        commands <- commands @ [FillRect (color, blockX, layout.Y, blockWidth, layout.Height)]
    | _ -> ()

    commands <- commands @ paragraphBorderCommands settings layout

    match layout.Metadata.ContentControlId with
    | Some _ ->
        let blockX = settings.PaddingLeft + layout.Indent
        let blockWidth = max 1.0 layout.AvailableWidth
        commands <- commands @ [
            StrokeRect ("#2563EB", 1.0, blockX, layout.Y, blockWidth, layout.Height)
            FillRect ("#2563EB", blockX, layout.Y, min 32.0 blockWidth, min 3.0 layout.Height)
        ]
    | None -> ()

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
        commands <- commands @ [DrawText (prefixText, prefixStyle, prefixX, baseline, prefixDrawWidth)]

    for line in layout.Lines do
        let baseline = lineBaselineWithOffset baselineOffset settings.FontSize line
        let lineAlignmentOffset = alignmentOffset layout line

        for segment in line.Segments do
            match segment with
            | TextSegment (text, style, x, width, _) ->
                let paintStyle = styleFromTextStyle style
                let textX = settings.PaddingLeft + x + lineAlignmentOffset
                match paintStyle.Highlight with
                | Some color -> commands <- commands @ [FillRect (color, textX, line.Y, width, line.Height)]
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
                        commands <- commands @ [DrawText (leaderText, paintStyle, textX + leaderInset, textBaseline, leaderWidth)]
                    | None -> ()
                else
                    commands <- commands @ [DrawText (displayText, paintStyle, textX, textBaseline, width)]

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
                        let prefix = displayText.Substring(0, lastInkIndex + 1)
                        let measured = measureText prefix style
                        if measured > 0.0 then measured else width
                if style.Underline && inkWidth > 0.0 then
                    let ulColor =
                        match style.UnderlineColor with
                        | Some c when c <> "" && c.ToLowerInvariant() <> "auto" -> if c.StartsWith("#") then c else "#" + c
                        | _ -> paintStyle.Color
                    let ulStyle = style.UnderlineStyle |> Option.defaultValue "single"
                    commands <- commands @ [DrawUnderline (ulColor, ulStyle, textX, textBaseline + WordLayout.Text.canvasUnderlineOffset, inkWidth)]
                if style.Strike && inkWidth > 0.0 then
                    commands <- commands @ [DrawUnderline (paintStyle.Color, "single", textX, textBaseline - WordLayout.Text.canvasStrikethroughOffset style, inkWidth)]
                if style.DoubleStrike && inkWidth > 0.0 then
                    let sy = textBaseline - WordLayout.Text.canvasStrikethroughOffset style
                    commands <- commands @ [
                        DrawUnderline (paintStyle.Color, "single", textX, sy - 1.0, inkWidth)
                        DrawUnderline (paintStyle.Color, "single", textX, sy + 1.5, inkWidth) ]
                match style.EmphasisMark with
                | Some mark when not (System.String.IsNullOrWhiteSpace displayText) ->
                    // A small mark above (or below for underDot) each non-space char.
                    let below = mark.ToLowerInvariant() = "underdot"
                    let markY = if below then textBaseline + 3.0 else textBaseline - paintStyle.FontSize * 0.95
                    let mutable acc = 0.0
                    for i in 0 .. displayText.Length - 1 do
                        let ch = displayText.[i]
                        let cw = measureText (string ch) style
                        if not (System.Char.IsWhiteSpace ch) then
                            commands <- commands @ [FillRect (paintStyle.Color, textX + acc + cw / 2.0 - 1.0, markY, 2.0, 2.0)]
                        acc <- acc + cw
                | _ -> ()

            | ImageSegment (src, x, yOffset, width, height, _, _, _, _, cropLeft, cropTop, cropRight, cropBottom) ->
                let imageX = settings.PaddingLeft + x + lineAlignmentOffset
                let imageY = line.Y + yOffset
                commands <- commands @ [DrawImage (src, imageX, imageY, width, height, cropLeft, cropTop, cropRight, cropBottom)]

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
                        commands <- commands @ [DrawText (text, glyphStyle, mathX + gx, baseline + gb, gw)]
                    | Eq.MRule (rx, top, w, h) ->
                        commands <- commands @ [FillRect (mathBase.Color, mathX + rx, baseline + top, w, h)]

    commands

let paragraphCommands settings fontFamily measureText layout =
    paragraphCommandsWithBaselineOffset 0.0 settings fontFamily measureText layout

// Keep line layout stable and correct only the drawn glyph baseline for Word's
// spaced table cells, whose visual text boxes sit higher than our default paint.
let private spacedTableCellBaselineOffset = -4.0

let private defaultParagraphFirstLineOffsetInDocument doc =
    let para = { Id = ""; Buffer = TextBuffer.empty; Metadata = defaultMetadata }
    let paragraphStyle = forBlockInDocument doc (Paragraph para)
    paragraphStyle.FirstLineIndent |> Option.defaultValue 0.0 |> max 0.0

let private tableLeftOffsetForStyleContext = function
    | Some doc -> defaultParagraphFirstLineOffsetInDocument doc
    | None -> 0.0

let private tableAlignmentOffset (alignment: string option) availableWidth tableWidth =
    match alignment |> Option.map (fun value -> value.ToLowerInvariant()) with
    | Some "center" -> max 0.0 ((availableWidth - tableWidth) / 2.0)
    | Some "right" -> max 0.0 (availableWidth - tableWidth)
    | _ -> 0.0

let markupMarkerCommands (settings: LayoutSettings) (layouts: BlockLayout list) =
    let mutable commands = []
    for layout in layouts do
        match layout.Metadata.ContentControlId with
        | Some _ ->
            let blockX = settings.PaddingLeft + layout.Indent
            let blockWidth = max 1.0 layout.AvailableWidth
            commands <- commands @ [
                StrokeRect ("#2563EB", 1.0, blockX, layout.Y, blockWidth, layout.Height)
                FillRect ("#2563EB", blockX, layout.Y, min 32.0 blockWidth, min 3.0 layout.Height)
            ]
        | None -> ()

    commands

let rec private tableBorderCommandsCoreAt (styleContext: Document option) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (y: float) (availableWidth: float) (baseHorizontalOffset: float) =
    let mutable currentY = y
    let mutable commands = []

    let numColumns = WordLayout.Table.columnCount tableBlock

    if numColumns = 0 then
        ([], currentY + WordLayout.Table.emptyTableAdvance settings.LineHeight)
    else
        let outerLeft, outerRight = WordLayout.Table.outerCellInsets WordLayout.canvasScale tableBlock
        let tableAvailableWidth = max 0.0 availableWidth
        let tableMaxWidth = tableAvailableWidth + outerLeft + outerRight
        let resolvedTable = WordLayout.Table.resolveTable WordLayout.canvasScale (Some tableMaxWidth) tableBlock
        let columnWidths = resolvedTable.ColumnWidths
        let tableWidth = WordLayout.Table.tableWidthWithSpacing columnWidths (WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock)
        let tableHorizontalOffset = baseHorizontalOffset + tableAlignmentOffset tableBlock.Alignment tableAvailableWidth tableWidth
        let tableX = settings.PaddingLeft + tableHorizontalOffset - outerLeft - WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock

        let columnX spacing startColumn =
            tableX + WordLayout.Table.spacedStartOffset columnWidths spacing startColumn - WordLayout.Table.spacedCellPaintXOffset spacing startColumn

        let columnWidth spacing startColumn span =
            WordLayout.Table.spacedSpannedWidth columnWidths spacing startColumn span (tableMaxWidth / float numColumns)
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

        let rowHeightAt (row: WordLayout.Table.ResolvedRow) =
            let paragraphStyleForParagraph, effectiveBufferForParagraph =
                match styleContext with
                | Some doc -> forTableCellParagraphInDocument doc, effectiveBufferInDocument doc
                | None -> (fun _ -> normal), effectiveBuffer
            let rowLocalSettings = {
                settings with
                    PageHeight = 1000000.0
                    MarginBottom = 0.0
                    TableContinuationHeaderHeight = 0.0
            }
            let measurement = measureTableRowWithStyle rowLocalSettings measureText measureRange tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset row.Source 0.0 0.0 paragraphStyleForParagraph effectiveBufferForParagraph
            measurement.Height

        let rowHeights = resolvedTable.BodyRows |> List.map rowHeightAt |> List.toArray

        let rowSpanHeightAt rowIndex rowSpan fallback =
            [rowIndex .. rowIndex + max 1 rowSpan - 1]
            |> List.sumBy (fun index -> rowHeights |> Array.tryItem index |> Option.defaultValue 0.0)
            |> max fallback

        let addRowCommands rowIndex (row: WordLayout.Table.ResolvedRow) currentY rowHeight =
            let cellSpacing = WordLayout.Table.rowCellSpacing WordLayout.canvasScale tableBlock row.Source
            let outerStartInset = outerLeft + cellSpacing
            for fragmentY, fragmentHeight in rowFragments settings currentY rowHeight do
                commands <- commands @ tableFillCommands tableBlock.BackgroundColor tableBlock.Shading tableX fragmentY tableWidth fragmentHeight

            let cells =
                row.Cells
                |> List.map (fun resolved ->
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
                    for fragmentY, fragmentHeight in rowFragments settings cellY cellHeight do
                        commands <- commands @ tableFillCommands cell.BackgroundColor cell.Shading x fragmentY width fragmentHeight

                    if not cell.NestedTables.IsEmpty then
                        let contentColWidth = WordLayout.Table.spacedSpannedWidth columnWidths cellSpacing resolved.StartColumn resolved.ColSpan (tableMaxWidth / float numColumns)
                        let nestedAvailableWidth = WordLayout.Table.cellContentWidth contentColWidth textPadding
                        let nestedX =
                            tableHorizontalOffset
                            + WordLayout.Table.spacedStartOffset columnWidths cellSpacing resolved.StartColumn
                            + textPadding.Left
                            - outerStartInset
                            - WordLayout.Table.internalSpacingTextOffset cellSpacing resolved.StartColumn
                        let nestedCellY = currentY + cellSpacing + WordLayout.Table.spacedCellTextTopOffset cellSpacing textPadding.Top
                        let mutable nestedY = WordLayout.Table.cellContentStartY nestedCellY textPadding
                        for nestedTable in cell.NestedTables do
                            let nestedCommands, nextNestedY = tableBorderCommandsCoreAt styleContext settings measureText measureRange nestedTable nestedY nestedAvailableWidth nestedX
                            commands <- commands @ nestedCommands
                            nestedY <- nextNestedY
                    ({
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
                            Bottom = borderStroke 1 resolved.Borders.Bottom
                            Left = borderStroke 1 resolved.Borders.Left
                        }
                    }: RedBat.BorderGrid.Cell))

            let borderFragmentHeight =
                cells
                |> List.map _.Height
                |> List.fold max rowHeight

            let borderSegments = RedBat.BorderGrid.resolve cells
            borderSegments
            |> List.iter (fun segment ->
                let startCap, endCap = tableBorderJoinCaps borderSegments segment
                for fragmentY, fragmentHeight in rowFragments settings currentY borderFragmentHeight do
                    let y1 = max segment.Y1 fragmentY
                    let y2 = min segment.Y2 (fragmentY + fragmentHeight)
                    if segment.Y1 = segment.Y2 then
                        if segment.Y1 >= fragmentY && segment.Y1 <= fragmentY + fragmentHeight then
                            commands <- commands @ strokeTableBorderCommands segment.Stroke.Color segment.Stroke.Width segment.Stroke.Style segment.X1 segment.Y1 segment.X2 segment.Y2 startCap endCap
                    elif y2 > y1 then
                        let clippedStartCap = if abs (y1 - segment.Y1) < borderJoinEpsilon then startCap else 0.0
                        let clippedEndCap = if abs (y2 - segment.Y2) < borderJoinEpsilon then endCap else 0.0
                        commands <- commands @ strokeTableBorderCommands segment.Stroke.Color segment.Stroke.Width segment.Stroke.Style segment.X1 y1 segment.X2 y2 clippedStartCap clippedEndCap)

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
                if cellSpacing > 0.0 && not cells.IsEmpty then
                    let firstCell = cells |> List.head
                    let lastCell = cells |> List.last
                    let outerLeft = firstCell.X - cellSpacing
                    let outerRight = lastCell.X + lastCell.Width + cellSpacing
                    let outerTop = currentY
                    let outerBottom = currentY + rowHeight
                    let paintOuterBorder (border: TableBorder option) x1 y1 x2 y2 =
                        match border with
                        | None -> ()
                        | Some b ->
                            match b.Style |> Option.map (fun s -> s.ToLowerInvariant()) with
                            | Some "nil" | Some "none" -> ()
                            | _ ->
                                let width = b.Width |> Option.defaultValue 0.5
                                let color = b.Color |> Option.defaultValue "#000000"
                                commands <- commands @ strokeTableBorderCommands color width b.Style x1 y1 x2 y2 0.0 0.0
                    paintOuterBorder overrideBorders.Top outerLeft outerTop outerRight outerTop
                    paintOuterBorder overrideBorders.Bottom outerLeft outerBottom outerRight outerBottom
                    paintOuterBorder overrideBorders.Left outerLeft outerTop outerLeft outerBottom
                    paintOuterBorder overrideBorders.Right outerRight outerTop outerRight outerBottom

        let headerRows = resolvedTable.HeaderRows
        let rows = resolvedTable.BodyRows

        let headerHeightAt rowY =
            headerRows
            |> List.indexed
            |> List.fold (fun (currentY, totalHeight) (headerIndex, headerRow) ->
                let headerHeight = rowHeights |> Array.tryItem headerIndex |> Option.defaultWith (fun () -> rowHeightAt headerRow)
                currentY + headerHeight, totalHeight + headerHeight) (rowY, 0.0)
            |> snd

        for rowIndex in 0 .. rows.Length - 1 do
            let row = rows.[rowIndex]
            let rowHeight = rowHeights.[rowIndex]
            let pagedY = paginateTableRow settings row.Source currentY rowHeight

            if pagedY <> currentY && not row.IsHeader && not headerRows.IsEmpty then
                currentY <- pagedY
                for headerIndex in 0 .. headerRows.Length - 1 do
                    let headerRow = headerRows.[headerIndex]
                    let headerHeight = rowHeights |> Array.tryItem headerIndex |> Option.defaultWith (fun () -> rowHeightAt headerRow)
                    addRowCommands headerIndex headerRow currentY headerHeight
                    currentY <- currentY + headerHeight
            else
                currentY <- pagedY

            addRowCommands rowIndex row currentY rowHeight
            currentY <- currentY + rowHeight

        let bottomSpacing = WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock
        if bottomSpacing > 0.0 then
            commands <- commands @ tableFillCommands tableBlock.BackgroundColor tableBlock.Shading tableX currentY tableWidth bottomSpacing

        (commands, currentY + WordLayout.Table.tableAfterSpacing WordLayout.canvasScale tableBlock)

let private tableBorderCommandsCore (styleContext: Document option) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (y: float) =
    let maxWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
    let tableLeftOffset = tableLeftOffsetForStyleContext styleContext
    let tableAvailableWidth = max 0.0 (maxWidth - tableLeftOffset)
    tableBorderCommandsCoreAt styleContext settings measureText measureRange tableBlock y tableAvailableWidth tableLeftOffset

let tableBorderCommandsForDocumentWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (tableBlock: TableBlock) (y: float) =
    tableBorderCommandsCore (Some doc) settings measureText measureRange tableBlock y

let tableBorderCommandsWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (y: float) =
    tableBorderCommandsCore None settings measureText measureRange tableBlock y

let tableBorderCommands (settings: LayoutSettings) (measureText: TextMeasurer) (tableBlock: TableBlock) (y: float) =
    tableBorderCommandsWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) tableBlock y

let private fieldCodeParts (code: string) =
    code.Trim().Split([| ' '; '\t'; '\r'; '\n' |], System.StringSplitOptions.RemoveEmptyEntries)
    |> Array.toList

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
    match fieldCodeParts code with
    | "PAGE" :: _ -> string pageNumber
    | "NUMPAGES" :: _ -> string pageCount
    | "DATE" :: _ -> System.DateTime.Now.ToString(fieldAtFormat code |> Option.defaultValue "yyyy-MM-dd")
    | "TIME" :: _ -> System.DateTime.Now.ToString(fieldAtFormat code |> Option.defaultValue "h:mm tt")
    | _ -> currentText

let private paragraphWithRunningFields pageNumber pageCount (para: ParagraphBlock) =
    let mutable buffer = para.Buffer
    let mutable shift = 0
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
    { para with Buffer = buffer }

let private blockWithRunningFields pageNumber pageCount = function
    | Paragraph para -> Paragraph (paragraphWithRunningFields pageNumber pageCount para)
    | Heading (level, para) -> Heading (level, paragraphWithRunningFields pageNumber pageCount para)
    | Table table ->
        let rec tableWithRunningFields (table: TableBlock) =
            { table with
                Rows =
                    table.Rows
                    |> List.map (fun row ->
                        { row with
                            Cells =
                                row.Cells
                                |> List.map (fun cell ->
                                    { cell with
                                        Content = cell.Content |> List.map (paragraphWithRunningFields pageNumber pageCount)
                                        NestedTables = cell.NestedTables |> List.map tableWithRunningFields }) }) }
        Table (tableWithRunningFields table)
    | TextBox textBox -> TextBox { textBox with Content = textBox.Content |> List.map (paragraphWithRunningFields pageNumber pageCount) }

let private partForReferences preferredKind refs (parts: HeaderFooterPart list) =
    refs
    |> List.tryPick (fun (kind, relationshipId) ->
        if kind = preferredKind then parts |> List.tryFind (fun part -> part.RelationshipId = Some relationshipId)
        else None)
    |> Option.orElseWith (fun () ->
        refs
        |> List.tryPick (fun (kind, relationshipId) ->
            if kind = "default" then parts |> List.tryFind (fun part -> part.RelationshipId = Some relationshipId)
            else None))

let private pageIndexForY (pages: PageInfo list) y =
    let mutable idx = 0
    pages |> List.iteri (fun i info -> if y >= info.Top - 0.01 then idx <- i)
    idx

// A header/footer reference lives in a section's sectPr, which (for a section
// closed by a break paragraph) sits on that section's LAST paragraph. So the
// reference applies to the section ending on that page; for a target page we
// want the section with the smallest end-page that still contains it.
let private sectionRefsForPage (pages: PageInfo list) pageIndex isHeader (layouts: BlockLayout list) =
    layouts
    |> List.choose (fun layout ->
        let refs = if isHeader then layout.Metadata.HeaderReferences else layout.Metadata.FooterReferences
        if refs.IsEmpty then None
        else Some (pageIndexForY pages (layout.Y + layout.Height), refs))
    |> List.filter (fun (endPage, _) -> endPage >= pageIndex)
    |> List.sortBy fst
    |> List.tryHead
    |> Option.map snd
    |> Option.defaultValue []

let private blocksForPage (settings: LayoutSettings) (pageSetup: PageSetup) (pages: PageInfo list) pageIndex isHeader fallback (parts: HeaderFooterPart list) layouts =
    if parts.IsEmpty then
        fallback
    else
        let pageNumber = pageIndex + 1
        let preferredKind =
            if pageNumber = 1 && pageSetup.DifferentFirstPage then "first"
            elif pageNumber % 2 = 0 && pageSetup.EvenAndOddHeaders then "even"
            else "default"
        let refs = sectionRefsForPage pages pageIndex isHeader layouts
        partForReferences preferredKind refs parts
        |> Option.orElseWith (fun () -> parts |> List.tryFind (fun part -> part.Kind = preferredKind))
        |> Option.orElseWith (fun () -> parts |> List.tryFind (fun part -> part.Kind = "default"))
        |> Option.orElseWith (fun () -> parts |> List.tryHead)
        |> Option.map (fun part -> part.Blocks)
        |> Option.defaultValue fallback

let private repeatedBlockCommands (settings: LayoutSettings) fontFamily measureText measureRange pageNumber pageCount startY blocks =
    let repeatedSettings = { settings with MarginBottom = 0.0 }
    let maxWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
    let commands = ResizeArray<PaintCommand>()
    let mutable y = startY
    for rawBlock in blocks do
        let block = blockWithRunningFields pageNumber pageCount rawBlock
        match block with
        | Paragraph para
        | Heading (_, para) ->
            let safePara = { para with Metadata = { para.Metadata with PageBreakAfter = false } }
            let layout, _ = layoutParagraphAtWithRange repeatedSettings measureText measureRange maxWidth y 0.0 safePara Map.empty
            commands.AddRange(paragraphCommands settings fontFamily measureText layout)
            y <- layout.Y + layout.Height
        | Table table ->
            let tableCommands, nextY = tableBorderCommandsWithRange repeatedSettings measureText measureRange table y
            commands.AddRange(tableCommands)
            y <- nextY
        | TextBox _ -> ()
    List.ofSeq commands

let private headerFooterCommands (settings: LayoutSettings) fontFamily measureText measureRange (doc: Document) layouts (pages: PageInfo list) =
    let pageCount = pages.Length
    [ for pageIndex, info in List.indexed pages do
        let pageNumber = pageIndex + 1
        let headers = blocksForPage settings doc.PageSetup pages pageIndex true (doc.Headers |> List.map Paragraph) doc.HeaderParts layouts
        let footers = blocksForPage settings doc.PageSetup pages pageIndex false (doc.Footers |> List.map Paragraph) doc.FooterParts layouts
        if not headers.IsEmpty then
            let headerY = info.Top + max 8.0 doc.PageSetup.HeaderDistance
            yield! repeatedBlockCommands settings fontFamily measureText measureRange pageNumber pageCount headerY headers
        if not footers.IsEmpty then
            let footerY = info.Top + info.Height - max 16.0 doc.PageSetup.FooterDistance - settings.LineHeight * 0.78
            yield! repeatedBlockCommands settings fontFamily measureText measureRange pageNumber pageCount footerY footers ]

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
        commands.AddRange(paragraphCommands settings fontFamily measureText layout)
        paraY <- layout.Y + layout.Height

    List.ofSeq commands

let private canonicalTableYFromLayouts settings measureText measureRange doc (layoutById: Map<BlockId, BlockLayout>) tableBlock =
    tableBlock.Rows
    |> List.collect (fun row -> row.Cells |> List.collect (fun cell -> cell.Content |> List.map _.Id))
    |> List.choose (fun id -> layoutById |> Map.tryFind id |> Option.map _.Y)
    |> function
        | [] -> None
        | values ->
            let firstActualContentY = List.min values
            let planAtZero = buildTableLayoutPlanForDocumentAt settings measureText measureRange doc tableBlock 0.0
            let resolveAlignmentHeight = tableCellVerticalAlignmentHeightResolverForPlan planAtZero
            let firstPlannedContentOffset =
                planAtZero.Rows
                |> List.collect (fun row ->
                    row.Measurement.CellGroups
                    |> List.mapi (fun groupIndex group ->
                        let alignmentHeight = resolveAlignmentHeight row.RowIndex groupIndex row.Height
                        let offset = WordLayout.Table.verticalContentOffset group.VerticalAlignment alignmentHeight group.PaddingTop group.PaddingBottom group.ContentHeight
                        group.Layouts |> List.map (fun layout -> layout.Y + offset))
                    |> List.collect id)
                |> function
                    | [] -> 0.0
                    | plannedValues -> List.min plannedValues
            Some (firstActualContentY - firstPlannedContentOffset)

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

/// When a table sits in a multi-column section its cell layouts were shifted by
/// the column offset (their Indent gained columnDx); the freshly built plan is
/// at full width. The difference is the X offset to shift the borders by.
let private tableColumnOffset (layoutById: Map<BlockId, BlockLayout>) (plan: TableLayoutPlan) (tableBlock: TableBlock) =
    let cellIds = tableBlock.Rows |> List.collect (fun row -> row.Cells |> List.collect (fun cell -> cell.Content |> List.map (fun p -> p.Id)))
    let cellIdSet = Set.ofList cellIds
    let actual = cellIds |> List.choose (fun id -> layoutById |> Map.tryFind id |> Option.map (fun l -> l.Indent))
    let planned =
        plan.Rows
        |> List.collect (fun row -> row.Measurement.CellGroups |> List.collect (fun g -> g.Layouts))
        |> List.filter (fun (l: BlockLayout) -> cellIdSet.Contains l.BlockId)
        |> List.map (fun l -> l.Indent)
    match actual, planned with
    | (_ :: _), (_ :: _) -> List.min actual - List.min planned
    | _ -> 0.0

let private tableFragmentPaintCommandsForDocument settings measureText measureRange doc layoutById tableBlock =
    canonicalTableYFromLayouts settings measureText measureRange doc layoutById tableBlock
    |> Option.map (fun tableY ->
        let plan = buildTableLayoutPlanForDocumentAt settings measureText measureRange doc tableBlock tableY
        let dx = tableColumnOffset layoutById plan tableBlock
        let commands = tableFragmentPaintCommands plan
        let commands = if abs dx > 0.5 then commands |> List.map (shiftCommandX dx) else commands
        commands, plan.NextY)

let rec private addTableTextCommands (settings: LayoutSettings) fontFamily measureText layoutsById (commands: ResizeArray<PaintCommand>) (tableBlock: TableBlock) =
    for row in tableBlock.Rows do
        let tableCellBaselineOffset =
            if WordLayout.Table.rowCellSpacing WordLayout.canvasScale tableBlock row > 0.0 then spacedTableCellBaselineOffset else 0.0
        for cell in row.Cells do
            for para in cell.Content do
                match Map.tryFind para.Id layoutsById with
                | Some paraLayouts -> paraLayouts |> List.iter (paragraphCommandsWithBaselineOffset tableCellBaselineOffset settings fontFamily measureText >> commands.AddRange)
                | None -> ()
            for nestedTable in cell.NestedTables do
                addTableTextCommands settings fontFamily measureText layoutsById commands nestedTable

// The fragment paint path draws outer-table borders but not nested-table
// borders. Locate each nested table from its content paragraph layouts and draw
// its grid with the shared border painter at that position.
let rec private addNestedTableBorders (styleContext: Document option) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (layoutsById: Map<BlockId, BlockLayout list>) (commands: ResizeArray<PaintCommand>) (tableBlock: TableBlock) =
    for row in tableBlock.Rows do
        for cell in row.Cells do
            for nestedTable in cell.NestedTables do
                let nestedLayouts =
                    nestedTable.Rows
                    |> List.collect (fun r -> r.Cells)
                    |> List.collect (fun c -> c.Content)
                    |> List.collect (fun p -> Map.tryFind p.Id layoutsById |> Option.defaultValue [])
                let segXs =
                    nestedLayouts
                    |> List.collect (fun l -> l.Lines)
                    |> List.collect (fun ln -> ln.Segments)
                    |> List.choose (function | TextSegment (_, _, x, _, _) -> Some x | _ -> None)
                match nestedLayouts, segXs with
                | (_ :: _), (_ :: _) ->
                    let topY = nestedLayouts |> List.map (fun l -> l.Y) |> List.min
                    let minX = List.min segXs
                    let firstRow = nestedTable.Rows |> List.head
                    let firstCell = firstRow.Cells |> List.head
                    let padLeft = (WordLayout.Table.cellPaddingForRow WordLayout.canvasScale nestedTable firstRow firstCell).Left
                    let baseHorizontalOffset = minX - padLeft
                    let availableWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
                    let cmds, _ = tableBorderCommandsCoreAt styleContext settings measureText measureRange nestedTable topY availableWidth baseHorizontalOffset
                    commands.AddRange cmds
                | _ -> ()
                addNestedTableBorders styleContext settings measureText measureRange layoutsById commands nestedTable

let collectDocumentCommandsWithLayouts (settings: LayoutSettings) fontFamily (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (cachedLayouts: BlockLayout list option) =
    let layouts, pages =
        match cachedLayouts with
        | Some ls -> ls, uniformPageInfos settings ls
        | None -> layoutDocumentPages settings measureText measureRange doc
    let layoutById = layouts |> List.map (fun layout -> layout.BlockId, layout) |> Map.ofList
    let layoutsById = layouts |> List.groupBy _.BlockId |> Map.ofList
    let commands = ResizeArray<PaintCommand>()
    let mutable currentY = doc.PageSetup.MarginTop

    for block in doc.Blocks do
        match block with
        | Table tableBlock ->
            let tableCommands, nextY =
                tableFragmentPaintCommandsForDocument settings measureText measureRange doc layoutById tableBlock
                |> Option.defaultWith (fun () -> tableBorderCommandsForDocumentWithRange settings measureText measureRange doc tableBlock currentY)
            commands.AddRange(tableCommands)
            if not tableBlock.Floating then currentY <- nextY
            addTableTextCommands settings fontFamily measureText layoutsById commands tableBlock
            addNestedTableBorders (Some doc) settings measureText measureRange layoutsById commands tableBlock
        | TextBox textBox -> commands.AddRange(textBoxCommands settings fontFamily measureText measureRange textBox)
        | _ ->
            match RedBat.Word.Document.tryGetParagraph block with
            | None -> ()
            | Some para ->
                match Map.tryFind para.Id layoutsById with
                | None -> ()
                | Some (layout :: _ as paraLayouts) ->
                    paraLayouts |> List.iter (paragraphCommands settings fontFamily measureText >> commands.AddRange)
                    currentY <- layout.Y + layout.Height
                    if para.Metadata.PageBreakAfter || para.Metadata.SectionBreak = Some "nextPage" then
                        currentY <- WordLayout.Page.advanceToNextPage settings.PageHeight settings.PaddingTop currentY
                | Some [] -> ()

    // Footnote area: paint the synthetic footnote paragraphs appended by layout
    // at the bottom of each page, and a short separator rule above the topmost
    // note on each page (Word draws ~2in).
    let footnoteLayouts = layouts |> List.filter (fun layout -> layout.BlockId.StartsWith footnoteBlockIdPrefix)
    for layout in footnoteLayouts do
        commands.AddRange(paragraphCommands settings fontFamily measureText layout)
    footnoteLayouts
    |> List.groupBy (fun layout -> int (floor (layout.Y / (settings.PageHeight + pageGap))))
    |> List.iter (fun (_, pageLayouts) ->
        let topY = pageLayouts |> List.map (fun layout -> layout.Y) |> List.min
        let separatorX = settings.PaddingLeft
        commands.Add(StrokeLine ("#000000", 0.75, 0, 0.0, separatorX, topY - 6.0, separatorX + 192.0, topY - 6.0)))

    // Endnote area: paint the synthetic endnote paragraphs appended after the
    // body content, with a single separator rule above the first endnote.
    let endnoteLayouts = layouts |> List.filter (fun layout -> layout.BlockId.StartsWith endnoteBlockIdPrefix)
    for layout in endnoteLayouts do
        commands.AddRange(paragraphCommands settings fontFamily measureText layout)
    match endnoteLayouts with
    | [] -> ()
    | _ ->
        // Short (~2in) separator above the first endnote; Word draws a full-width
        // continuation separator on pages the endnotes spill onto.
        let pageOf (layout: BlockLayout) = int (floor (layout.Y / (settings.PageHeight + pageGap)))
        let startPage = endnoteLayouts |> List.map pageOf |> List.min
        let fullWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
        endnoteLayouts
        |> List.groupBy pageOf
        |> List.iter (fun (page, pageLayouts) ->
            let topY = pageLayouts |> List.map (fun layout -> layout.Y) |> List.min
            let separatorX = settings.PaddingLeft
            let width = if page = startPage then 192.0 else fullWidth
            commands.Add(StrokeLine ("#000000", 0.75, 0, 0.0, separatorX, topY - 8.0, separatorX + width, topY - 8.0)))

    let pageArr = Array.ofList pages
    let pageCount = max 1 pageArr.Length

    // Column separator rules (w:sep): a vertical line centred in each gap,
    // spanning the body content on each page.
    match doc.PageSetup.Columns with
    | Some columns when columns.Count >= 2 && columns.Separator ->
        let n = columns.Count
        let full = settings.Width - settings.PaddingLeft - settings.MarginRight
        let colWidth = (full - float (n - 1) * columns.Gap) / float n
        let bodyLines =
            layouts
            |> List.filter (fun layout -> not (layout.BlockId.StartsWith footnoteBlockIdPrefix || layout.BlockId.StartsWith endnoteBlockIdPrefix))
            |> List.collect (fun layout -> layout.Lines)
        for page in 0 .. pageCount - 1 do
            let info = pageArr.[page]
            let contentTop = info.Top + info.MarginTop
            let contentBottom = info.Top + info.Height - info.MarginBottom
            let pageBottom =
                bodyLines
                |> List.filter (fun line -> line.Y >= info.Top - 0.5 && line.Y < info.Top + info.Height)
                |> List.map (fun line -> line.Y + line.Height)
                |> function [] -> contentTop | ys -> min contentBottom (List.max ys)
            if pageBottom > contentTop then
                for c in 0 .. n - 2 do
                    let x = settings.PaddingLeft + float (c + 1) * colWidth + float c * columns.Gap + columns.Gap / 2.0
                    commands.Add(StrokeLine ("#000000", 0.5, 0, 0.0, x, contentTop, x, pageBottom))
    | _ -> ()

    commands.AddRange(headerFooterCommands settings fontFamily measureText measureRange doc layouts pages)
    layouts, List.ofSeq commands, pages

let collectDocumentCommands settings fontFamily measureText measureRange doc =
    collectDocumentCommandsWithLayouts settings fontFamily measureText measureRange doc None

let selectionCommandsWithRange (settings: LayoutSettings) (measureRange: TextRangeMeasurer) (doc: Document) (layouts: BlockLayout list) =
    match doc.Selection with
    | None -> []
    | Some sel when sel.IsCollapsed ->
        match layouts |> List.tryFind (fun l -> l.BlockId = sel.Focus.BlockId) with
        | None -> []
        | Some layout ->
            let line = lineForOffset layout sel.Focus.Offset
            let x = xForOffsetWithRange measureRange line sel.Focus.Offset
            let lineAlignmentOffset = alignmentOffset layout line
            [FillRect ("#0066ff", settings.PaddingLeft + x + lineAlignmentOffset, line.Y, 2.0, line.Height)]
    | Some sel ->
        let (startPos, endPos, blockIds) = getNormalizedRange sel doc
        let effectiveBlockIds =
            if blockIds.IsEmpty && startPos.BlockId = endPos.BlockId then [startPos.BlockId]
            else blockIds

        let mutable commands = []

        for (i, blockId) in (effectiveBlockIds |> List.mapi (fun i id -> (i, id))) do
            match layouts |> List.tryFind (fun l -> l.BlockId = blockId) with
            | None -> ()
            | Some layout ->
                let bufferLen = maxOffsetInLayout layout

                let (highlightStart, highlightEnd) =
                    if effectiveBlockIds.Length = 1 then
                        (startPos.Offset, endPos.Offset)
                    elif i = 0 then
                        (startPos.Offset, bufferLen)
                    elif i = effectiveBlockIds.Length - 1 then
                        (0, endPos.Offset)
                    else
                        (0, bufferLen)

                for line in layout.Lines do
                    let lineStart = max highlightStart line.StartOffset
                    let lineEnd = min highlightEnd line.EndOffset
                    if lineEnd > lineStart then
                        let startX = xForOffsetWithRange measureRange line lineStart
                        let endX = xForOffsetWithRange measureRange line lineEnd
                        let lineAlignmentOffset = alignmentOffset layout line
                        commands <- commands @ [FillRect ("rgba(0, 120, 215, 0.3)", settings.PaddingLeft + startX + lineAlignmentOffset, line.Y, endX - startX, line.Height)]

        match layouts |> List.tryFind (fun l -> l.BlockId = sel.Focus.BlockId) with
        | None -> commands
        | Some layout ->
            let line = lineForOffset layout sel.Focus.Offset
            let x = xForOffsetWithRange measureRange line sel.Focus.Offset
            let lineAlignmentOffset = alignmentOffset layout line
            commands @ [FillRect ("#0066ff", settings.PaddingLeft + x + lineAlignmentOffset, line.Y, 2.0, line.Height)]

let selectionCommands (settings: LayoutSettings) (measureText: TextMeasurer) (doc: Document) (layouts: BlockLayout list) =
    selectionCommandsWithRange settings (rangeMeasurerFromTextMeasurer measureText) doc layouts

let hoverCommands (settings: LayoutSettings) (layouts: BlockLayout list) (hoveredImage: (BlockId * int) option) =
    match hoveredImage with
    | None -> []
    | Some (blockId, offset) ->
        match imageBoxForOffset settings.PaddingLeft layouts blockId offset with
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
