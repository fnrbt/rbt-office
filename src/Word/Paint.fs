module RedBat.Word.Paint

open RedBat.Word.Layout
open RedBat.Word.Document
open RedBat.Word.TextBuffer
open RedBat.Word.WordStyles

type TextPaintStyle = {
    Bold: bool
    Italic: bool
    Underline: bool
    AllCaps: bool
    SmallCaps: bool
    Strike: bool
    FontSize: float
    FontFamily: string option
    Color: string
    Highlight: string option
    VertAlign: string option
}

type PaintCommand =
    | DrawText of text: string * style: TextPaintStyle * x: float * baseline: float * width: float
    | DrawUnderline of color: string * x: float * y: float * width: float
    | DrawImage of src: string * x: float * y: float * width: float * height: float * cropLeft: string option * cropTop: string option * cropRight: string option * cropBottom: string option
    | FillRect of color: string * x: float * y: float * width: float * height: float
    | StrokeRect of color: string * lineWidth: float * x: float * y: float * width: float * height: float
    | StrokeLine of color: string * lineWidth: float * dashStyle: int * dashOffset: float * x1: float * y1: float * x2: float * y2: float

let private styleFromTextStyle (style: TextStyle) = {
    Bold = style.Bold
    Italic = style.Italic
    Underline = style.Underline || style.Hyperlink.IsSome || WordLayout.Text.isInsertedRevision style
    AllCaps = style.AllCaps
    SmallCaps = style.SmallCaps
    Strike = style.Strike || WordLayout.Text.isDeletedRevision style
    FontSize = WordLayout.Text.printFontSize style
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

let private lineBaseline (fallbackFontSize: float) (line: MeasuredLine) =
    let fontSize =
        line.Segments
        |> List.choose (function
            | TextSegment (_, style, _, _, _) -> Some (WordLayout.Text.printFontSize style)
            | ImageSegment _ -> None)
        |> function
            | [] -> fallbackFontSize
            | sizes -> List.max sizes
    line.Y + fontSize

let private dashStyleForTableBorderStyle style =
    match RedBat.BorderGrid.styleKey style with
    | "dotted" -> 1
    | "dashed"
    | "dashsmallgap" -> 2
    | "dotdash"
    | "dotdotdash"
    | "dashdotstroked" -> 3
    | _ -> 0

let private tableBorderStrokeParts width style =
    let width = max 0.25 width
    RedBat.BorderGrid.strokePattern style
    |> List.map (fun (offset, partWidth) -> offset * width, max 0.25 (partWidth * width))

let private strokeTableBorderCommands color width style x1 y1 x2 y2 =
    let dashStyle = dashStyleForTableBorderStyle style
    tableBorderStrokeParts width style
    |> List.map (fun (offset, partWidth) ->
        if y1 = y2 then StrokeLine (color, partWidth, dashStyle, 0.0, x1, y1 + offset, x2, y2 + offset)
        else StrokeLine (color, partWidth, dashStyle, 0.0, x1 + offset, y1, x2 + offset, y2))

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
                match plan.Table.BackgroundColor with
                | Some color -> commands.Add(FillRect (color, tableBackgroundX, fragment.Y, plan.TableWidth, fragment.Height))
                | None -> ()

                let rowFragments =
                    cellFragmentsByRowFragment
                    |> Map.tryFind (row.RowIndex, fragment.Page, fragment.Y, fragment.Height)
                    |> Option.defaultValue []

                for cellFragment in rowFragments do
                    let sourceRow = plan.Table.Rows.[cellFragment.RowIndex]
                    match sourceRow.Cells |> List.tryItem cellFragment.CellIndex |> Option.bind _.BackgroundColor with
                    | Some color ->
                        let fillHeight = if cellFragment.CellSpacing > 0.0 then cellFragment.FragmentHeight else cellFragment.CellHeight
                        commands.Add(FillRect (color, cellFragment.CellX - cellFragment.CellSpacing, cellFragment.CellY, cellFragment.CellWidth, fillHeight))
                    | None -> ()

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
                tableCellFragmentBorderCells paintCells
                |> RedBat.BorderGrid.resolve
                |> List.iter (fun segment ->
                    strokeTableBorderCommands segment.Stroke.Color segment.Stroke.Width segment.Stroke.Style segment.X1 segment.Y1 segment.X2 segment.Y2
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

let paragraphCommands (settings: LayoutSettings) (fontFamily: string) (measureText: TextMeasurer) (layout: BlockLayout) =
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
        let baseline = lineBaseline settings.FontSize firstLine
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
        let baseline = lineBaseline settings.FontSize line
        let lineAlignmentOffset = alignmentOffset layout line

        for segment in line.Segments do
            match segment with
            | TextSegment (text, style, x, width, _) ->
                let paintStyle = styleFromTextStyle style
                let textX = settings.PaddingLeft + x + lineAlignmentOffset
                match paintStyle.Highlight with
                | Some color -> commands <- commands @ [FillRect (color, textX, line.Y, width, line.Height)]
                | None -> ()
                if style.BookmarkStartId.IsSome || style.BookmarkName.IsSome || style.BookmarkEndId.IsSome then
                    commands <- commands @ [
                        StrokeRect ("#7C3AED", 1.0, textX, line.Y, width, line.Height)
                        FillRect ("#A855F7", textX, line.Y, min 24.0 width, min 3.0 line.Height)
                    ]
                let textBaseline =
                    match paintStyle.VertAlign with
                    | Some "superscript" -> baseline - paintStyle.FontSize * 0.5
                    | Some "subscript" -> baseline + paintStyle.FontSize * 0.25
                    | _ -> baseline
                let displayText = WordLayout.Text.displayText style text
                commands <- commands @ [DrawText (displayText, paintStyle, textX, textBaseline, width)]

                if style.Underline then
                    commands <- commands @ [DrawUnderline (paintStyle.Color, textX, textBaseline + WordLayout.Text.canvasUnderlineOffset, width)]
                if style.Strike then
                    commands <- commands @ [DrawUnderline (paintStyle.Color, textX, textBaseline - WordLayout.Text.canvasStrikethroughOffset style, width)]

            | ImageSegment (src, x, yOffset, width, height, _, floating, _, _, cropLeft, cropTop, cropRight, cropBottom) ->
                let imageX = settings.PaddingLeft + x + lineAlignmentOffset
                let imageY = line.Y + yOffset
                commands <- commands @ [DrawImage (src, imageX, imageY, width, height, cropLeft, cropTop, cropRight, cropBottom)]
                if floating then
                    commands <- commands @ [
                        StrokeRect ("#B45309", 1.5, imageX, imageY, width, height)
                        FillRect ("#F59E0B", imageX, imageY, min 18.0 width, min 4.0 height)
                    ]

    commands

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

        for line in layout.Lines do
            let lineAlignmentOffset = alignmentOffset layout line
            for segment in line.Segments do
                match segment with
                | TextSegment (_, style, x, width, _) when style.BookmarkStartId.IsSome || style.BookmarkName.IsSome || style.BookmarkEndId.IsSome ->
                    let textX = settings.PaddingLeft + x + lineAlignmentOffset
                    commands <- commands @ [
                        StrokeRect ("#7C3AED", 1.0, textX, line.Y, width, line.Height)
                        FillRect ("#A855F7", textX, line.Y, min 24.0 width, min 3.0 line.Height)
                    ]
                | _ -> ()
    commands

let private tableBorderCommandsCore (styleContext: Document option) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (y: float) =
    let maxWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
    let mutable currentY = y
    let mutable commands = []

    let numColumns = WordLayout.Table.columnCount tableBlock

    if numColumns = 0 then
        ([], currentY + WordLayout.Table.emptyTableAdvance settings.LineHeight)
    else
        let outerLeft, outerRight = WordLayout.Table.outerCellInsets WordLayout.canvasScale tableBlock
        let tableMaxWidth = maxWidth + outerLeft + outerRight
        let tableX = settings.PaddingLeft - outerLeft - WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock
        let resolvedTable = WordLayout.Table.resolveTable WordLayout.printScale (Some tableMaxWidth) tableBlock
        let columnWidths = resolvedTable.ColumnWidths
        let tableWidth = WordLayout.Table.tableWidthWithSpacing columnWidths (WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock)

        let columnX spacing startColumn =
            tableX + WordLayout.Table.spacedStartOffset columnWidths spacing startColumn

        let columnWidth spacing startColumn span =
            WordLayout.Table.spacedSpannedWidth columnWidths spacing startColumn span (tableMaxWidth / float numColumns)

        let borderStroke priority (stroke: WordLayout.Table.PrintStroke option) =
            stroke
            |> Option.map (fun stroke -> ({
                Color = stroke.Color
                Width = stroke.Width
                Style = stroke.Style
                Priority = priority
            }: RedBat.BorderGrid.Stroke))

        let borderStyleKey = RedBat.BorderGrid.styleKey

        let dashStyleForBorderStyle style =
            match borderStyleKey style with
            | "dotted" -> 1
            | "dashed"
            | "dashsmallgap" -> 2
            | "dotdash"
            | "dotdotdash"
            | "dashdotstroked" -> 3
            | _ -> 0

        let styleStrokeParts width style =
            let width = max 0.25 width
            RedBat.BorderGrid.strokePattern style
            |> List.map (fun (offset, partWidth) -> offset * width, max 0.25 (partWidth * width))

        let strokeBorderCommands color width style x1 y1 x2 y2 =
            let dashStyle = dashStyleForBorderStyle style
            styleStrokeParts width style
            |> List.map (fun (offset, partWidth) ->
                if y1 = y2 then StrokeLine (color, partWidth, dashStyle, 0.0, x1, y1 + offset, x2, y2 + offset)
                else StrokeLine (color, partWidth, dashStyle, 0.0, x1 + offset, y1, x2 + offset, y2))

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
            let measurement = measureTableRowWithStyle rowLocalSettings measureText measureRange tableBlock tableMaxWidth columnWidths outerLeft row.Source 0.0 0.0 paragraphStyleForParagraph effectiveBufferForParagraph
            measurement.Height

        let addRowCommands (row: WordLayout.Table.ResolvedRow) currentY rowHeight =
            let cellSpacing = WordLayout.Table.rowCellSpacing WordLayout.canvasScale tableBlock row.Source
            for fragmentY, fragmentHeight in rowFragments settings currentY rowHeight do
                match tableBlock.BackgroundColor with
                | Some color -> commands <- commands @ [FillRect (color, tableX, fragmentY, tableWidth, fragmentHeight)]
                | None -> ()

            let cells =
                row.Cells
                |> List.map (fun resolved ->
                    let cell = resolved.Cell
                    let x = columnX cellSpacing resolved.StartColumn
                    let width = columnWidth cellSpacing resolved.StartColumn resolved.ColSpan
                    let cellY = currentY + cellSpacing
                    let cellHeight = max 0.0 (rowHeight - cellSpacing)
                    for fragmentY, fragmentHeight in rowFragments settings cellY cellHeight do
                        match cell.BackgroundColor with
                        | Some color -> commands <- commands @ [FillRect (color, x, fragmentY, width, fragmentHeight)]
                        | None -> ()
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

            RedBat.BorderGrid.resolve cells
            |> List.iter (fun segment ->
                for fragmentY, fragmentHeight in rowFragments settings currentY rowHeight do
                    let y1 = max segment.Y1 fragmentY
                    let y2 = min segment.Y2 (fragmentY + fragmentHeight)
                    if segment.Y1 = segment.Y2 then
                        if segment.Y1 >= fragmentY && segment.Y1 <= fragmentY + fragmentHeight then
                            commands <- commands @ strokeBorderCommands segment.Stroke.Color segment.Stroke.Width segment.Stroke.Style segment.X1 segment.Y1 segment.X2 segment.Y2
                    elif y2 > y1 then
                        commands <- commands @ strokeBorderCommands segment.Stroke.Color segment.Stroke.Width segment.Stroke.Style segment.X1 y1 segment.X2 y2)

        let headerRows = resolvedTable.HeaderRows
        let rows = resolvedTable.BodyRows

        let headerHeightAt rowY =
            headerRows
            |> List.fold (fun (currentY, totalHeight) headerRow ->
                let headerHeight = rowHeightAt headerRow
                currentY + headerHeight, totalHeight + headerHeight) (rowY, 0.0)
            |> snd

        for row in rows do
            let rowHeight = rowHeightAt row
            let pagedY = paginateTableRow settings row.Source currentY rowHeight

            if pagedY <> currentY && not row.IsHeader && not headerRows.IsEmpty then
                currentY <- pagedY
                for headerRow in headerRows do
                    let headerHeight = rowHeightAt headerRow
                    addRowCommands headerRow currentY headerHeight
                    currentY <- currentY + headerHeight
            else
                currentY <- pagedY

            addRowCommands row currentY rowHeight
            currentY <- currentY + rowHeight

        (commands, currentY + WordLayout.Table.afterSpacing)

let tableBorderCommandsForDocumentWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (tableBlock: TableBlock) (y: float) =
    tableBorderCommandsCore (Some doc) settings measureText measureRange tableBlock y

let tableBorderCommandsWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (y: float) =
    tableBorderCommandsCore None settings measureText measureRange tableBlock y

let tableBorderCommands (settings: LayoutSettings) (measureText: TextMeasurer) (tableBlock: TableBlock) (y: float) =
    tableBorderCommandsWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) tableBlock y

let private fieldCodeParts (code: string) =
    code.Trim().Split([| ' '; '\t'; '\r'; '\n' |], System.StringSplitOptions.RemoveEmptyEntries)
    |> Array.toList

let private runningFieldResult pageNumber pageCount code currentText =
    match fieldCodeParts code with
    | "PAGE" :: _ -> string pageNumber
    | "NUMPAGES" :: _ -> string pageCount
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
        Table {
            table with
                Rows =
                    table.Rows
                    |> List.map (fun row ->
                        { row with
                            Cells =
                                row.Cells
                                |> List.map (fun cell ->
                                    { cell with Content = cell.Content |> List.map (paragraphWithRunningFields pageNumber pageCount) }) })
        }
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

let private pageIndexForY (settings: LayoutSettings) y =
    let stride = settings.PageHeight + pageGap
    int (floor (max 0.0 y / stride))

let private sectionRefsForPage (settings: LayoutSettings) pageIndex isHeader (layouts: BlockLayout list) =
    layouts
    |> List.choose (fun layout ->
        let refs = if isHeader then layout.Metadata.HeaderReferences else layout.Metadata.FooterReferences
        if refs.IsEmpty then None
        else
            let startPage =
                if layout.Metadata.SectionBreak = Some "nextPage" || layout.Metadata.PageBreakAfter then pageIndexForY settings layout.Y + 1
                else pageIndexForY settings (layout.Y + layout.Height)
            Some (startPage, refs))
    |> List.filter (fun (startPage, _) -> startPage <= pageIndex)
    |> List.sortBy fst
    |> List.tryLast
    |> Option.map snd
    |> Option.defaultValue []

let private blocksForPage (settings: LayoutSettings) (pageSetup: PageSetup) pageIndex isHeader fallback (parts: HeaderFooterPart list) layouts =
    if parts.IsEmpty then
        fallback
    else
        let pageNumber = pageIndex + 1
        let preferredKind =
            if pageNumber = 1 && pageSetup.DifferentFirstPage then "first"
            elif pageNumber % 2 = 0 && pageSetup.EvenAndOddHeaders then "even"
            else "default"
        let refs = sectionRefsForPage settings pageIndex isHeader layouts
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

let private headerFooterCommands (settings: LayoutSettings) fontFamily measureText measureRange (doc: Document) layouts pageCount =
    let stride = settings.PageHeight + pageGap
    [ for pageIndex in 0 .. pageCount - 1 do
        let pageNumber = pageIndex + 1
        let pageY = float pageIndex * stride
        let headers = blocksForPage settings doc.PageSetup pageIndex true (doc.Headers |> List.map Paragraph) doc.HeaderParts layouts
        let footers = blocksForPage settings doc.PageSetup pageIndex false (doc.Footers |> List.map Paragraph) doc.FooterParts layouts
        if not headers.IsEmpty then
            let headerY = pageY + max 8.0 doc.PageSetup.HeaderDistance
            yield! repeatedBlockCommands settings fontFamily measureText measureRange pageNumber pageCount headerY headers
        if not footers.IsEmpty then
            let footerY = pageY + settings.PageHeight - max 16.0 doc.PageSetup.FooterDistance
            yield! repeatedBlockCommands settings fontFamily measureText measureRange pageNumber pageCount footerY footers ]

let private textBoxCommands (settings: LayoutSettings) fontFamily measureText measureRange (textBox: TextBoxBlock) =
    let floatingBox = WordLayout.Floating.textBoxCanvasBox settings.PaddingLeft settings.PaddingTop textBox
    let padding = 6.0
    let maxWidth = max 1.0 (textBox.Width - padding * 2.0)
    let mutable paraY = floatingBox.Y + padding
    let commands = ResizeArray<PaintCommand>()

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
            let firstPlannedContentOffset =
                planAtZero.Rows
                |> List.collect (fun row ->
                    row.Measurement.CellGroups
                    |> List.collect (fun group ->
                        let offset = WordLayout.Table.verticalContentOffset group.VerticalAlignment row.Height group.PaddingTop group.PaddingBottom group.ContentHeight
                        group.Layouts |> List.map (fun layout -> layout.Y + offset)))
                |> function
                    | [] -> 0.0
                    | plannedValues -> List.min plannedValues
            Some (firstActualContentY - firstPlannedContentOffset)

let private tableFragmentPaintCommandsForDocument settings measureText measureRange doc layoutById tableBlock =
    canonicalTableYFromLayouts settings measureText measureRange doc layoutById tableBlock
    |> Option.map (fun tableY ->
        let plan = buildTableLayoutPlanForDocumentAt settings measureText measureRange doc tableBlock tableY
        tableFragmentPaintCommands plan, plan.NextY)

let collectDocumentCommandsWithLayouts (settings: LayoutSettings) fontFamily (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (cachedLayouts: BlockLayout list option) =
    let layouts = cachedLayouts |> Option.defaultWith (fun () -> layoutDocumentWithRange settings measureText measureRange doc)
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
            for row in tableBlock.Rows do
                for cell in row.Cells do
                    for para in cell.Content do
                        match Map.tryFind para.Id layoutsById with
                        | Some paraLayouts -> paraLayouts |> List.iter (paragraphCommands settings fontFamily measureText >> commands.AddRange)
                        | None -> ()
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

    let bottom =
        layouts
        |> List.map (fun layout -> layout.Y + layout.Height)
        |> List.fold max settings.PaddingTop
    let stride = settings.PageHeight + pageGap
    let pageCount = max 1 (int (floor (max 0.0 bottom / stride)) + 1)
    commands.AddRange(headerFooterCommands settings fontFamily measureText measureRange doc layouts pageCount)
    layouts, List.ofSeq commands, pageCount

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
