module RedBat.Word.Renderer

#if FABLE_COMPILER
open System
open System.Collections.Generic
open Fable.Core
open Browser.Types
open RedBat.Word.Document
open RedBat.Word.TextBuffer
open RedBat.Word.Layout
open RedBat.Word.Paint

module Math = RedBat.Word.Math

type RenderLayoutCacheEntry = {
    mutable Document: Document
    LayoutDocument: Document
    Layouts: BlockLayout list
    mutable LayoutArray: BlockLayout array option
    mutable LayoutById: Dictionary<BlockId, BlockLayout> option
    LayoutCount: int
    SourceParagraphCount: int
    SourceTextCharCount: int
    SourceStyleCharCount: int
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
    mutable LayoutCache: RenderLayoutCacheEntry list
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
    LayoutCache = []
}

let private layoutDictionary (layouts: BlockLayout list) =
    let dict = Dictionary<BlockId, BlockLayout>()
    for layout in layouts do
        if not (dict.ContainsKey layout.BlockId) then
            dict.[layout.BlockId] <- layout
    dict

let private cssFontFamily (family: string) =
    if family.Contains(" ") then sprintf "\"%s\"" (family.Replace("\"", "")) else family

let private asciiCharStrings = Array.init 128 (fun code -> string (char code))

let private singleCharString (ch: char) =
    let code = int ch
    if code < asciiCharStrings.Length then asciiCharStrings.[code]
    else string ch

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private charEqualsIgnoreCase a b =
    a = b || Char.ToUpperInvariant a = Char.ToUpperInvariant b

let private equalsOrdinalIgnoreCase (left: string) (right: string) =
    if left.Length <> right.Length then
        false
    else
        let mutable matches = true
        let mutable i = 0
        while matches && i < left.Length do
            if not (charEqualsIgnoreCase left.[i] right.[i]) then
                matches <- false
            i <- i + 1
        matches

let private trimmedEqualsOrdinalIgnoreCase (left: string) (right: string) =
    let trimBounds (value: string) =
        let mutable startIndex = 0
        let mutable endIndex = value.Length - 1
        while startIndex <= endIndex && Char.IsWhiteSpace value.[startIndex] do
            startIndex <- startIndex + 1
        while endIndex >= startIndex && Char.IsWhiteSpace value.[endIndex] do
            endIndex <- endIndex - 1
        startIndex, endIndex

    let leftStart, leftEnd = trimBounds left
    let rightStart, rightEnd = trimBounds right
    let leftLength = leftEnd - leftStart + 1
    let rightLength = rightEnd - rightStart + 1
    if leftLength <> rightLength then
        false
    else
        let mutable matches = true
        let mutable i = 0
        while matches && i < leftLength do
            if not (charEqualsIgnoreCase left.[leftStart + i] right.[rightStart + i]) then
                matches <- false
            i <- i + 1
        matches

let private containsOrdinalIgnoreCase (needle: string) (text: string) =
    if needle.Length = 0 then
        true
    elif needle.Length > text.Length then
        false
    else
        let mutable found = false
        let mutable startIndex = 0
        while not found && startIndex <= text.Length - needle.Length do
            let mutable matches = true
            let mutable offset = 0
            while matches && offset < needle.Length do
                if not (charEqualsIgnoreCase text.[startIndex + offset] needle.[offset]) then
                    matches <- false
                offset <- offset + 1
            if matches then
                found <- true
            startIndex <- startIndex + 1
        found

let private browserMetricScale (family: string) (style: TextBuffer.TextStyle) =
    // Chromium canvas metrics for loaded Windows faces diverge from Word/PDF
    // shaping enough to change centering and wrap decisions in real DOCX bodies.
    if trimmedEqualsOrdinalIgnoreCase family "cambria" && not style.Bold && not style.Italic then 0.982
    elif trimmedEqualsOrdinalIgnoreCase family "times new roman" && style.Bold && style.Italic then 0.945
    elif trimmedEqualsOrdinalIgnoreCase family "times new roman" && not style.Bold && not style.Italic then 1.0185
    else 1.0

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

let private renderLayoutCacheMaxEntries = 2
let private renderLayoutCacheMaxTotalLayouts = 200000
let private renderLayoutCacheMaxTotalSourceParagraphs = 100000
let private renderLayoutCacheMaxTotalSourceTextChars = 5000000
let private renderLayoutCacheMaxTotalStyleChars = 500000

let private layoutInputsEqual (left: Document) (right: Document) =
    obj.ReferenceEquals(left.Blocks, right.Blocks)
    && left.PageSetup = right.PageSetup
    && obj.ReferenceEquals(left.Styles, right.Styles)
    && obj.ReferenceEquals(left.Headers, right.Headers)
    && obj.ReferenceEquals(left.Footers, right.Footers)
    && obj.ReferenceEquals(left.HeaderParts, right.HeaderParts)
    && obj.ReferenceEquals(left.FooterParts, right.FooterParts)

let private renderLayoutAddCapped limit value total =
    if total > limit then
        total
    else
        let value = max 0 value
        let remaining = limit - total
        if value > remaining then limit + 1
        else total + value

let private renderLayoutOptionStringCharCount (value: string option) =
    match value with
    | Some value -> value.Length
    | None -> 0

let private renderLayoutSectionColumnsStringCharCount (columns: SectionColumns option) =
    match columns with
    | Some columns ->
        let mutable total = 0
        let mutable seen = 0
        let mutable widths = columns.Widths
        while total <= renderLayoutCacheMaxTotalSourceTextChars
              && seen <= renderLayoutCacheMaxTotalSourceParagraphs
              && not widths.IsEmpty do
            total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars 1 total
            seen <- seen + 1
            widths <- widths.Tail
        if seen > renderLayoutCacheMaxTotalSourceParagraphs then
            renderLayoutCacheMaxTotalSourceTextChars + 1
        else
            total
    | None -> 0

let private renderLayoutTextStyleStringCharCount (style: TextBuffer.TextStyle) =
    let mutable total = 0
    let add value =
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount value) total
    add style.Style
    add style.UnderlineStyle
    add style.UnderlineColor
    add style.EmphasisMark
    add style.FontFamily
    add style.Color
    add style.Highlight
    add style.Hyperlink
    add style.CommentId
    add style.CommentText
    add style.CommentAuthor
    add style.CommentDate
    add style.NoteId
    add style.NoteType
    add style.NoteText
    add style.BookmarkStartId
    add style.BookmarkName
    add style.BookmarkEndId
    add style.FieldCode
    add style.FieldResult
    add style.EquationXml
    add style.RevisionId
    add style.RevisionType
    add style.RevisionAuthor
    add style.RevisionDate
    add style.VertAlign
    total

let private renderLayoutImageStringCharCount (image: TextBuffer.ImageRun) =
    let mutable total = image.Src.Length
    let add value =
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount value) total
    add image.Alt
    add image.Name
    add image.Title
    add image.Wrap
    add image.PositionHRelativeFrom
    add image.PositionVRelativeFrom
    add image.RelativeHeight
    add image.CropLeft
    add image.CropTop
    add image.CropRight
    add image.CropBottom
    total

let private renderLayoutRunStringCharCount = function
    | TextBuffer.Text textRun -> renderLayoutTextStyleStringCharCount textRun.Style
    | TextBuffer.Image image -> renderLayoutImageStringCharCount image

let private renderLayoutParagraphBorderStringCharCount (border: ParagraphBorder option) =
    match border with
    | Some border ->
        let mutable total = 0
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount border.Style) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount border.Color) total
        total
    | None -> 0

let private renderLayoutParagraphBordersStringCharCount (borders: ParagraphBorders option) =
    match borders with
    | Some borders ->
        let mutable total = 0
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutParagraphBorderStringCharCount borders.Top) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutParagraphBorderStringCharCount borders.Right) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutParagraphBorderStringCharCount borders.Bottom) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutParagraphBorderStringCharCount borders.Left) total
        total
    | None -> 0

let private renderLayoutTabStopStringCharCount (tabStop: TabStop) =
    renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount tabStop.Leader) tabStop.Alignment.Length

let private renderLayoutParagraphMetadataStringCharCount (metadata: BlockMetadata) =
    let mutable total = 0
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount metadata.Style) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount metadata.ContentControlId) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount metadata.ContentControlTag) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount metadata.ContentControlAlias) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount metadata.ListId) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount metadata.ListNumberFormat) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount metadata.ListLevelText) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount metadata.RightTabLeader) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount metadata.LineSpacingRule) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount metadata.BackgroundColor) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutParagraphBordersStringCharCount metadata.ParagraphBorders) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount metadata.SectionBreak) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutSectionColumnsStringCharCount metadata.SectionColumns) total
    let mutable seen = 0
    let mutable tabStops = metadata.TabStops
    while total <= renderLayoutCacheMaxTotalSourceTextChars
          && seen <= renderLayoutCacheMaxTotalSourceParagraphs
          && not tabStops.IsEmpty do
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTabStopStringCharCount tabStops.Head) total
        seen <- seen + 1
        tabStops <- tabStops.Tail
    let mutable headerReferences = metadata.HeaderReferences
    while total <= renderLayoutCacheMaxTotalSourceTextChars
          && seen <= renderLayoutCacheMaxTotalSourceParagraphs
          && not headerReferences.IsEmpty do
        let kind, relationshipId = headerReferences.Head
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars kind.Length total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars relationshipId.Length total
        seen <- seen + 1
        headerReferences <- headerReferences.Tail
    let mutable footerReferences = metadata.FooterReferences
    while total <= renderLayoutCacheMaxTotalSourceTextChars
          && seen <= renderLayoutCacheMaxTotalSourceParagraphs
          && not footerReferences.IsEmpty do
        let kind, relationshipId = footerReferences.Head
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars kind.Length total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars relationshipId.Length total
        seen <- seen + 1
        footerReferences <- footerReferences.Tail
    if seen > renderLayoutCacheMaxTotalSourceParagraphs then
        renderLayoutCacheMaxTotalSourceTextChars + 1
    else
        total

let private renderLayoutParagraphSourceTextCharCount (paragraph: ParagraphBlock) =
    let limit = renderLayoutCacheMaxTotalSourceTextChars
    let mutable total = 0
    total <- renderLayoutAddCapped limit paragraph.Id.Length total
    total <- renderLayoutAddCapped limit (TextBuffer.length paragraph.Buffer) total
    let mutable seen = 0
    let mutable runs = paragraph.Buffer.Runs
    while total <= limit
          && seen <= renderLayoutCacheMaxTotalSourceParagraphs
          && not runs.IsEmpty do
        total <- renderLayoutAddCapped limit (renderLayoutRunStringCharCount runs.Head) total
        seen <- seen + 1
        runs <- runs.Tail
    if seen > renderLayoutCacheMaxTotalSourceParagraphs then
        limit + 1
    else
        if total <= limit then
            total <- renderLayoutAddCapped limit (renderLayoutParagraphMetadataStringCharCount paragraph.Metadata) total
        total

let private renderLayoutTableBorderStringCharCount (border: TableBorder option) =
    match border with
    | Some border ->
        let mutable total = 0
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount border.Style) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount border.Color) total
        total
    | None -> 0

let private renderLayoutTableBordersStringCharCount (borders: TableBorders option) =
    match borders with
    | Some borders ->
        let mutable total = 0
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTableBorderStringCharCount borders.Top) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTableBorderStringCharCount borders.Right) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTableBorderStringCharCount borders.Bottom) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTableBorderStringCharCount borders.Left) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTableBorderStringCharCount borders.InsideH) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTableBorderStringCharCount borders.InsideV) total
        total
    | None -> 0

let private renderLayoutShadingStringCharCount (shading: Shading option) =
    match shading with
    | Some shading ->
        let mutable total = 0
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount shading.Pattern) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount shading.ForegroundColor) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount shading.BackgroundColor) total
        total
    | None -> 0

let private renderLayoutTableStringCharCount (table: TableBlock) =
    let mutable total = table.Id.Length
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount table.Alignment) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount table.BackgroundColor) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutShadingStringCharCount table.Shading) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTableBordersStringCharCount table.Borders) total
    total

let private renderLayoutTableRowStringCharCount (row: TableRow) =
    let mutable total = 0
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount row.HeightRule) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTableBordersStringCharCount row.BordersOverride) total
    total

let private renderLayoutTableCellStringCharCount (cell: TableCell) =
    let mutable total = 0
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount cell.BackgroundColor) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutShadingStringCharCount cell.Shading) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount cell.BorderColor) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTableBordersStringCharCount cell.Borders) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount cell.VerticalAlignment) total
    total

let private renderLayoutPathSegmentsStringCharCount (segments: PathSeg list) =
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = segments
    while total <= renderLayoutCacheMaxTotalSourceTextChars
          && seen <= renderLayoutCacheMaxTotalSourceParagraphs
          && not remaining.IsEmpty do
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars 1 total
        seen <- seen + 1
        remaining <- remaining.Tail
    if seen > renderLayoutCacheMaxTotalSourceParagraphs then
        renderLayoutCacheMaxTotalSourceTextChars + 1
    else
        total

let private renderLayoutShapeGeometryStringCharCount = function
    | GeoRect -> 0
    | GeoPreset preset -> preset.Length
    | GeoCustom (segments, _, _) -> renderLayoutPathSegmentsStringCharCount segments

let private renderLayoutGradientStringCharCount (gradient: GradientSpec option) =
    match gradient with
    | Some gradient ->
        let mutable total = 0
        let mutable seen = 0
        let mutable stops = gradient.Stops
        while total <= renderLayoutCacheMaxTotalSourceTextChars
              && seen <= renderLayoutCacheMaxTotalSourceParagraphs
              && not stops.IsEmpty do
            total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars stops.Head.Color.Length total
            seen <- seen + 1
            stops <- stops.Tail
        if seen > renderLayoutCacheMaxTotalSourceParagraphs then
            renderLayoutCacheMaxTotalSourceTextChars + 1
        else
            total
    | None -> 0

let private renderLayoutTextBoxStringCharCount (textBox: TextBoxBlock) =
    let mutable total = textBox.Id.Length
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount textBox.PositionHRelativeFrom) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount textBox.PositionVRelativeFrom) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount textBox.FillColor) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount textBox.BorderColor) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount textBox.RelativeHeight) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutShapeGeometryStringCharCount textBox.Geometry) total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutGradientStringCharCount textBox.Gradient) total
    total

let private renderLayoutSelectionStringCharCount (selection: Selection option) =
    match selection with
    | Some selection ->
        renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars selection.Focus.BlockId.Length selection.Anchor.BlockId.Length
    | None -> 0

let private renderLayoutPageSetupStringCharCount (setup: PageSetup) =
    renderLayoutAddCapped
        renderLayoutCacheMaxTotalSourceTextChars
        (renderLayoutSectionColumnsStringCharCount setup.Columns)
        (renderLayoutOptionStringCharCount setup.SectionStart)

let private renderLayoutHeaderFooterPartStringCharCount (part: HeaderFooterPart) =
    renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount part.RelationshipId) part.Kind.Length

let private renderLayoutHeaderFooterPartsMetadataStringCharCount (parts: HeaderFooterPart list) =
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = parts
    while total <= renderLayoutCacheMaxTotalSourceTextChars
          && seen <= renderLayoutCacheMaxTotalSourceParagraphs
          && not remaining.IsEmpty do
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutHeaderFooterPartStringCharCount remaining.Head) total
        seen <- seen + 1
        remaining <- remaining.Tail
    if seen > renderLayoutCacheMaxTotalSourceParagraphs then
        renderLayoutCacheMaxTotalSourceTextChars + 1
    else
        total

type private RenderLayoutSourceMetrics = {
    ParagraphCount: int
    TextCharCount: int
}

let private renderLayoutDocumentSourceMetrics (doc: Document) =
    let mutable paragraphCount = 0
    let mutable textCharCount = 0
    let mutable visitedItems = 0
    let hasRoom () =
        paragraphCount <= renderLayoutCacheMaxTotalSourceParagraphs
        && textCharCount <= renderLayoutCacheMaxTotalSourceTextChars
        && visitedItems <= renderLayoutCacheMaxTotalSourceParagraphs
    let visitItem () =
        visitedItems <- visitedItems + 1
    let addParagraph (paragraph: ParagraphBlock) =
        if hasRoom () then
            visitItem ()
            paragraphCount <- paragraphCount + 1
            textCharCount <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutParagraphSourceTextCharCount paragraph) textCharCount
    let addTable (table: TableBlock) =
        let tables = Stack<TableBlock>()
        tables.Push table
        while tables.Count > 0 && hasRoom () do
            let table = tables.Pop()
            visitItem ()
            textCharCount <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTableStringCharCount table) textCharCount
            let mutable rows = table.Rows
            while hasRoom () && not rows.IsEmpty do
                let row = rows.Head
                visitItem ()
                textCharCount <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTableRowStringCharCount row) textCharCount
                let mutable cells = row.Cells
                while hasRoom () && not cells.IsEmpty do
                    let cell = cells.Head
                    visitItem ()
                    textCharCount <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTableCellStringCharCount cell) textCharCount
                    let mutable paragraphs = cell.Content
                    while hasRoom () && not paragraphs.IsEmpty do
                        addParagraph paragraphs.Head
                        paragraphs <- paragraphs.Tail
                    let mutable nestedTables = cell.NestedTables
                    while hasRoom () && not nestedTables.IsEmpty do
                        tables.Push nestedTables.Head
                        nestedTables <- nestedTables.Tail
                    cells <- cells.Tail
                rows <- rows.Tail
    let addBlock block =
        if hasRoom () then
            match block with
            | Paragraph paragraph
            | Heading (_, paragraph) -> addParagraph paragraph
            | Table table -> addTable table
            | TextBox textBox ->
                visitItem ()
                textCharCount <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTextBoxStringCharCount textBox) textCharCount
                let mutable paragraphs = textBox.Content
                while hasRoom () && not paragraphs.IsEmpty do
                    addParagraph paragraphs.Head
                    paragraphs <- paragraphs.Tail
    if textCharCount <= renderLayoutCacheMaxTotalSourceTextChars then
        textCharCount <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutSelectionStringCharCount doc.Selection) textCharCount
        textCharCount <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutPageSetupStringCharCount doc.PageSetup) textCharCount
        textCharCount <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutHeaderFooterPartsMetadataStringCharCount doc.HeaderParts) textCharCount
        textCharCount <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutHeaderFooterPartsMetadataStringCharCount doc.FooterParts) textCharCount
    let mutable blocks = doc.Blocks
    while hasRoom () && not blocks.IsEmpty do
        addBlock blocks.Head
        blocks <- blocks.Tail
    let mutable headers = doc.Headers
    while hasRoom () && not headers.IsEmpty do
        addParagraph headers.Head
        headers <- headers.Tail
    let mutable footers = doc.Footers
    while hasRoom () && not footers.IsEmpty do
        addParagraph footers.Head
        footers <- footers.Tail
    let mutable headerParts = doc.HeaderParts
    while hasRoom () && not headerParts.IsEmpty do
        let part = headerParts.Head
        let mutable partBlocks = part.Blocks
        while hasRoom () && not partBlocks.IsEmpty do
            addBlock partBlocks.Head
            partBlocks <- partBlocks.Tail
        headerParts <- headerParts.Tail
    let mutable footerParts = doc.FooterParts
    while hasRoom () && not footerParts.IsEmpty do
        let part = footerParts.Head
        let mutable partBlocks = part.Blocks
        while hasRoom () && not partBlocks.IsEmpty do
            addBlock partBlocks.Head
            partBlocks <- partBlocks.Tail
        footerParts <- footerParts.Tail
    if visitedItems > renderLayoutCacheMaxTotalSourceParagraphs then
        {
            ParagraphCount = renderLayoutCacheMaxTotalSourceParagraphs + 1
            TextCharCount = renderLayoutCacheMaxTotalSourceTextChars + 1
        }
    else
        {
            ParagraphCount = min paragraphCount (renderLayoutCacheMaxTotalSourceParagraphs + 1)
            TextCharCount = min textCharCount (renderLayoutCacheMaxTotalSourceTextChars + 1)
        }

let private renderLayoutStyleStringCharCount (style: NamedStyle) =
    let mutable total = style.Id.Length + style.StyleType.Length
    let add value =
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalStyleChars (renderLayoutOptionStringCharCount value) total
    add style.Name
    add style.BasedOn
    add style.Next
    add style.Linked
    add style.Aliases
    add style.LineSpacingRule
    add style.BackgroundColor
    match style.Text with
    | Some textStyle ->
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalStyleChars (renderLayoutTextStyleStringCharCount textStyle) total
    | None -> ()
    total

let private renderLayoutStylesStringCharCount (styles: NamedStyle list) =
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = styles
    while total <= renderLayoutCacheMaxTotalStyleChars
          && seen <= renderLayoutCacheMaxTotalSourceParagraphs
          && not remaining.IsEmpty do
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalStyleChars (renderLayoutStyleStringCharCount remaining.Head) total
        seen <- seen + 1
        remaining <- remaining.Tail
    if seen > renderLayoutCacheMaxTotalSourceParagraphs then
        renderLayoutCacheMaxTotalStyleChars + 1
    else
        total

let private renderLayoutMathPrimsStringCharCount (prims: Math.MathPrim list) =
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = prims
    while total <= renderLayoutCacheMaxTotalSourceTextChars
          && seen <= renderLayoutCacheMaxTotalLayouts
          && not remaining.IsEmpty do
        match remaining.Head with
        | Math.MGlyph (text, _, _, _, _) ->
            total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars text.Length total
        | Math.MRule _ -> ()
        seen <- seen + 1
        remaining <- remaining.Tail
    if seen > renderLayoutCacheMaxTotalLayouts then
        renderLayoutCacheMaxTotalSourceTextChars + 1
    else
        total

let private renderLayoutSegmentStringCharCount = function
    | TextSegment (text, style, _, _, _) ->
        let mutable total = text.Length
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTextStyleStringCharCount style) total
        total
    | ImageSegment (src, _, _, _, _, _, _, wrap, relativeHeight, cropLeft, cropTop, cropRight, cropBottom) ->
        let mutable total = src.Length
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount wrap) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount relativeHeight) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount cropLeft) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount cropTop) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount cropRight) total
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutOptionStringCharCount cropBottom) total
        total
    | MathSegment (prims, style, _, _, _, _, _, _) ->
        let mutable total = renderLayoutTextStyleStringCharCount style
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutMathPrimsStringCharCount prims) total
        total

let private renderLayoutLineStringCharCount total (line: MeasuredLine) =
    let mutable total = total
    let mutable seen = 0
    let mutable remaining = line.Segments
    while total <= renderLayoutCacheMaxTotalSourceTextChars
          && seen <= renderLayoutCacheMaxTotalLayouts
          && not remaining.IsEmpty do
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutSegmentStringCharCount remaining.Head) total
        seen <- seen + 1
        remaining <- remaining.Tail
    if seen > renderLayoutCacheMaxTotalLayouts then
        renderLayoutCacheMaxTotalSourceTextChars + 1
    else
        total

let private renderLayoutBlockLayoutStringCharCount total (layout: BlockLayout) =
    let mutable total = total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars layout.BlockId.Length total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars layout.Prefix.Length total
    total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutParagraphMetadataStringCharCount layout.Metadata) total
    match layout.PrefixStyle with
    | Some style ->
        total <- renderLayoutAddCapped renderLayoutCacheMaxTotalSourceTextChars (renderLayoutTextStyleStringCharCount style) total
    | None -> ()
    let mutable seen = 0
    let mutable remaining = layout.Lines
    while total <= renderLayoutCacheMaxTotalSourceTextChars
          && seen <= renderLayoutCacheMaxTotalLayouts
          && not remaining.IsEmpty do
        total <- renderLayoutLineStringCharCount total remaining.Head
        seen <- seen + 1
        remaining <- remaining.Tail
    if seen > renderLayoutCacheMaxTotalLayouts then
        renderLayoutCacheMaxTotalSourceTextChars + 1
    else
        total

let private renderLayoutLayoutsStringCharCount (layouts: BlockLayout list) =
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = layouts
    while total <= renderLayoutCacheMaxTotalSourceTextChars
          && seen <= renderLayoutCacheMaxTotalLayouts
          && not remaining.IsEmpty do
        total <- renderLayoutBlockLayoutStringCharCount total remaining.Head
        seen <- seen + 1
        remaining <- remaining.Tail
    if seen > renderLayoutCacheMaxTotalLayouts then
        renderLayoutCacheMaxTotalSourceTextChars + 1
    else
        total

let private renderLayoutItemCountAddCapped value total =
    renderLayoutAddCapped renderLayoutCacheMaxTotalLayouts value total

let private renderLayoutMathPrimsRetainedItemCount total (prims: Math.MathPrim list) =
    let mutable total = total
    let mutable remaining = prims
    while total <= renderLayoutCacheMaxTotalLayouts && not remaining.IsEmpty do
        total <- renderLayoutItemCountAddCapped 1 total
        remaining <- remaining.Tail
    total

let private renderLayoutSegmentRetainedItemCount total segment =
    let total = renderLayoutItemCountAddCapped 1 total
    match segment with
    | MathSegment (prims, _, _, _, _, _, _, _) when total <= renderLayoutCacheMaxTotalLayouts ->
        renderLayoutMathPrimsRetainedItemCount total prims
    | _ -> total

let private renderLayoutLineRetainedItemCount total (line: MeasuredLine) =
    let mutable total = renderLayoutItemCountAddCapped 1 total
    let mutable remaining = line.Segments
    while total <= renderLayoutCacheMaxTotalLayouts && not remaining.IsEmpty do
        total <- renderLayoutSegmentRetainedItemCount total remaining.Head
        remaining <- remaining.Tail
    total

let private renderLayoutBlockLayoutRetainedItemCount total (layout: BlockLayout) =
    let mutable total = renderLayoutItemCountAddCapped 1 total
    let mutable remaining = layout.Lines
    while total <= renderLayoutCacheMaxTotalLayouts && not remaining.IsEmpty do
        total <- renderLayoutLineRetainedItemCount total remaining.Head
        remaining <- remaining.Tail
    total

let private renderLayoutLayoutsRetainedItemCount (layouts: BlockLayout list) =
    let mutable total = 0
    let mutable remaining = layouts
    while total <= renderLayoutCacheMaxTotalLayouts && not remaining.IsEmpty do
        total <- renderLayoutBlockLayoutRetainedItemCount total remaining.Head
        remaining <- remaining.Tail
    total

let private renderLayoutDocumentSourceMetricsFor retainedLayoutItemCount doc =
    if retainedLayoutItemCount > renderLayoutCacheMaxTotalLayouts then
        {
            ParagraphCount = renderLayoutCacheMaxTotalSourceParagraphs + 1
            TextCharCount = renderLayoutCacheMaxTotalSourceTextChars + 1
        }
    else
        renderLayoutDocumentSourceMetrics doc

let private renderLayoutSourceMetricsWithLayouts sourceMetrics layouts =
    if sourceMetrics.TextCharCount > renderLayoutCacheMaxTotalSourceTextChars then
        sourceMetrics
    else
        {
            sourceMetrics with
                TextCharCount =
                    renderLayoutAddCapped
                        renderLayoutCacheMaxTotalSourceTextChars
                        (renderLayoutLayoutsStringCharCount layouts)
                        sourceMetrics.TextCharCount
        }

let private renderLayoutStylesStringCharCountFor retainedLayoutItemCount sourceMetrics styles =
    if retainedLayoutItemCount > renderLayoutCacheMaxTotalLayouts
       || sourceMetrics.ParagraphCount > renderLayoutCacheMaxTotalSourceParagraphs
       || sourceMetrics.TextCharCount > renderLayoutCacheMaxTotalSourceTextChars then
        renderLayoutCacheMaxTotalStyleChars + 1
    else
        renderLayoutStylesStringCharCount styles

let private renderLayoutCacheEntryRetainable entry =
    entry.LayoutCount <= renderLayoutCacheMaxTotalLayouts
    && entry.SourceParagraphCount <= renderLayoutCacheMaxTotalSourceParagraphs
    && entry.SourceTextCharCount <= renderLayoutCacheMaxTotalSourceTextChars
    && entry.SourceStyleCharCount <= renderLayoutCacheMaxTotalStyleChars

let private rememberLayoutCacheEntry (ctx: RenderContext) entry =
    let trim entries =
        let rec loop remaining kept count totalLayouts totalSourceParagraphs totalSourceTextChars totalStyleChars =
            match remaining with
            | [] -> List.rev kept
            | candidate :: rest ->
                let layoutCount = max 1 candidate.LayoutCount
                let sourceParagraphCount = max 1 candidate.SourceParagraphCount
                let sourceTextCharCount = max 1 candidate.SourceTextCharCount
                let styleCharCount = max 1 candidate.SourceStyleCharCount
                let nextTotalLayouts = totalLayouts + layoutCount
                let nextTotalSourceParagraphs = totalSourceParagraphs + sourceParagraphCount
                let nextTotalSourceTextChars = totalSourceTextChars + sourceTextCharCount
                let nextTotalStyleChars = totalStyleChars + styleCharCount
                if renderLayoutCacheEntryRetainable candidate
                   && count < renderLayoutCacheMaxEntries
                   && nextTotalLayouts <= renderLayoutCacheMaxTotalLayouts
                   && nextTotalSourceParagraphs <= renderLayoutCacheMaxTotalSourceParagraphs
                   && nextTotalSourceTextChars <= renderLayoutCacheMaxTotalSourceTextChars
                   && nextTotalStyleChars <= renderLayoutCacheMaxTotalStyleChars then
                    loop
                        rest
                        (candidate :: kept)
                        (count + 1)
                        nextTotalLayouts
                        nextTotalSourceParagraphs
                        nextTotalSourceTextChars
                        nextTotalStyleChars
                else
                    loop rest kept count totalLayouts totalSourceParagraphs totalSourceTextChars totalStyleChars
        loop entries [] 0 0 0 0 0
    let retained =
        ctx.LayoutCache
        |> List.filter (fun cached -> not (layoutInputsEqual cached.LayoutDocument entry.LayoutDocument))
    let entries =
        if renderLayoutCacheEntryRetainable entry then entry :: retained
        else retained
    ctx.LayoutCache <- trim entries

let private buildLayoutCacheEntry (ctx: RenderContext) (doc: Document) =
    let layouts = RedBat.Word.Layout.layoutDocumentWithRange (layoutSettings ctx doc) (measureWithContext ctx) (measureRangeWithContext ctx) doc
    let retainedLayoutItemCount = renderLayoutLayoutsRetainedItemCount layouts
    let sourceMetrics =
        renderLayoutDocumentSourceMetricsFor retainedLayoutItemCount doc
        |> fun metrics -> renderLayoutSourceMetricsWithLayouts metrics layouts
    {
        Document = doc
        LayoutDocument = doc
        Layouts = layouts
        LayoutArray = Some (Array.ofList layouts)
        LayoutById = Some (layoutDictionary layouts)
        LayoutCount = retainedLayoutItemCount
        SourceParagraphCount = sourceMetrics.ParagraphCount
        SourceTextCharCount = sourceMetrics.TextCharCount
        SourceStyleCharCount = renderLayoutStylesStringCharCountFor retainedLayoutItemCount sourceMetrics doc.Styles
    }

let private tryLayoutCacheEntry (ctx: RenderContext) (doc: Document) =
    ctx.LayoutCache
    |> List.tryFind (fun cached ->
        obj.ReferenceEquals(cached.Document, doc)
        || layoutInputsEqual cached.LayoutDocument doc)
    |> Option.map (fun cached ->
        if not (obj.ReferenceEquals(cached.Document, doc)) then
            cached.Document <- doc
            rememberLayoutCacheEntry ctx cached
        cached)

let private layoutCacheEntryFor (ctx: RenderContext) (doc: Document) =
    match tryLayoutCacheEntry ctx doc with
    | Some cached -> cached
    | None ->
        let cached = buildLayoutCacheEntry ctx doc
        rememberLayoutCacheEntry ctx cached
        cached

let layoutDocument (ctx: RenderContext) (doc: Document) =
    let cached = buildLayoutCacheEntry ctx doc
    rememberLayoutCacheEntry ctx cached
    cached.Layouts

let clearLayoutCache (ctx: RenderContext) =
    ctx.LayoutCache <- []

let cachedLayouts (ctx: RenderContext) (doc: Document) =
    tryLayoutCacheEntry ctx doc |> Option.map _.Layouts

let private layoutsFor (ctx: RenderContext) (doc: Document) =
    (layoutCacheEntryFor ctx doc).Layouts

let private layoutArrayFor (ctx: RenderContext) (doc: Document) =
    let cached = layoutCacheEntryFor ctx doc
    match cached.LayoutArray with
    | Some layoutArray when layoutArray.Length = cached.Layouts.Length -> layoutArray
    | _ ->
        let layoutArray = Array.ofList cached.Layouts
        cached.LayoutArray <- Some layoutArray
        layoutArray

let private layoutLookupForEntry cached =
    let dict =
        match cached.LayoutById with
        | Some dict when dict.Count = cached.Layouts.Length -> dict
        | _ ->
            let dict = layoutDictionary cached.Layouts
            cached.LayoutById <- Some dict
            dict
    fun blockId ->
        let mutable layout = Unchecked.defaultof<BlockLayout>
        if dict.TryGetValue(blockId, &layout) then Some layout else None

let layoutLookupFor (ctx: RenderContext) (doc: Document) =
    layoutLookupForEntry (layoutCacheEntryFor ctx doc)

let private commentAnchorsForLayouts (ctx: RenderContext) (doc: Document) (layouts: BlockLayout list) =
    match commentInfos doc with
    | [] -> []
    | infos ->
        let settings = layoutSettings ctx doc
        let layoutById = Dictionary<string, BlockLayout>()
        for layout in layouts do
            if not (layoutById.ContainsKey layout.BlockId) then
                layoutById.[layout.BlockId] <- layout
        let calloutX = max (settings.Width - settings.MarginRight + 12.0) (settings.PaddingLeft + 120.0)

        let isSelected blockId startOffset endOffset =
            match doc.Selection with
            | Some selection when selection.Anchor.BlockId = blockId && selection.Focus.BlockId = blockId ->
                let selectionStart = min selection.Anchor.Offset selection.Focus.Offset
                let selectionEnd = max selection.Anchor.Offset selection.Focus.Offset
                selectionStart = startOffset && selectionEnd = endOffset
            | _ -> false

        let anchors = ResizeArray<CommentAnchor>()
        let measureRange = measureRangeWithContext ctx

        for info in infos do
            match layoutById.TryGetValue info.BlockId with
            | false, _ -> ()
            | true, layout ->
                let line = lineForOffset layout info.StartOffset
                let lineEnd = min info.EndOffset line.EndOffset
                if lineEnd > info.StartOffset then
                    let lineAlignmentOffset = alignmentOffset layout line
                    let startX = xForOffsetWithRange measureRange line info.StartOffset
                    let endX = xForOffsetWithRange measureRange line lineEnd
                    anchors.Add {
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
                    }

        anchors.Sort(Comparison<CommentAnchor>(fun left right -> compare left.Y right.Y))

        let minGap = 8.0
        let defaultCalloutHeight = 58.0
        let stacked = ResizeArray<CommentAnchor>(anchors.Count)
        let mutable previousBottom = -100000.0
        for anchor in anchors do
            let calloutY = max anchor.Y (previousBottom + minGap)
            stacked.Add { anchor with CalloutY = calloutY }
            previousBottom <- calloutY + defaultCalloutHeight
        resizeArrayToList stacked

let commentAnchors (ctx: RenderContext) (doc: Document) =
    commentAnchorsForLayouts ctx doc (layoutsFor ctx doc)

let private pageIndexForY settings y =
    let stride = pageHeight settings + pageGap
    int (floor (max 0.0 y / stride))

let private endnoteStartYFromBottom (settings: LayoutSettings) bottom =
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) bottom
    let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
    if bottom + 48.0 > contentBottom then WordLayout.Page.advanceToNextPage (pageHeight settings) settings.PaddingTop bottom else bottom + 24.0

let private documentHeightWithNotes (settings: LayoutSettings) (layouts: BlockLayout list) =
    let mutable bottom = settings.PaddingTop
    let seenEndnotes = HashSet<string>()
    let seenFootnotes = HashSet<string>()
    let mutable endnoteLineCount = 0
    let footnoteLineCountsByPage = Dictionary<int, int>()

    let addFootnote pageIndex id text =
        let key = string pageIndex + "\u0000" + id
        if seenFootnotes.Add key then
            let added = noteLineCount settings (sprintf "%s. %s" id text)
            match footnoteLineCountsByPage.TryGetValue pageIndex with
            | true, existing -> footnoteLineCountsByPage.[pageIndex] <- existing + added
            | false, _ -> footnoteLineCountsByPage.[pageIndex] <- added

    for layout in layouts do
        bottom <- max bottom (layout.Y + layout.Height)
        for line in layout.Lines do
            let pageIndex = pageIndexForY settings line.Y
            for segment in line.Segments do
                match segment with
                | TextSegment (_, style, _, _, _) ->
                    match style.NoteId, style.NoteType, style.NoteText with
                    | Some id, Some "endnote", Some text when text <> "" ->
                        if seenEndnotes.Add id then
                            endnoteLineCount <- endnoteLineCount + noteLineCount settings (sprintf "%s. %s" id text)
                    | Some id, noteType, Some text when noteType |> Option.defaultValue "footnote" = "footnote" && text <> "" ->
                        addFootnote pageIndex id text
                    | _ -> ()
                | ImageSegment _ -> ()
                | MathSegment _ -> ()

    let endnoteBottom =
        if endnoteLineCount = 0 then bottom
        else
            endnoteStartYFromBottom settings bottom + 18.0 + float endnoteLineCount * 13.0
    let mutable footnoteBottom = bottom
    for KeyValue(pageIndex, lineCount) in footnoteLineCountsByPage do
        let pageTop = float pageIndex * (pageHeight settings + pageGap)
        let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
        footnoteBottom <- max footnoteBottom (contentBottom + 20.0 + float lineCount * 13.0)
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
            let ch = text.[i]
            let chText = singleCharString ch
            ctx.Ctx.fillText(chText, cursor, 0.0)
            cursor <- cursor + ctx.Ctx.measureText(chText).width
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
        let hasStylePart part = containsOrdinalIgnoreCase part lineStyle
        let heavy = hasStylePart "heavy" || equalsOrdinalIgnoreCase lineStyle "thick"
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
        if hasStylePart "wav" then
            wavy 0.0
            if hasStylePart "double" then wavy 2.6
        elif hasStylePart "dotdotdash" || hasStylePart "dashdotdot" then dashed [|4.0;2.0;1.0;2.0;1.0;2.0|]
        elif hasStylePart "dotdash" || hasStylePart "dashdot" then dashed [|4.0;2.0;1.0;2.0|]
        elif hasStylePart "dash" then dashed [|4.0;2.0|]
        elif hasStylePart "dot" then dashed [|1.0;2.0|]
        elif equalsOrdinalIgnoreCase lineStyle "double" then
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

let private renderCommentAnchors (ctx: RenderContext) (doc: Document) (layouts: BlockLayout list) =
    commentAnchorsForLayouts ctx doc layouts
    |> List.iter (fun anchor ->
        renderPaintCommand ctx (FillRect ("rgba(255, 230, 128, 0.42)", anchor.X, anchor.Y, anchor.Width, anchor.Height))
        renderPaintCommand ctx (StrokeRect ("#C79B00", 1.0, anchor.X + 0.5, anchor.Y + 0.5, max 1.0 (anchor.Width - 1.0), max 1.0 (anchor.Height - 1.0))))

let private renderCore (ctx: RenderContext) (doc: Document) =
    // Clear canvas
    ctx.Ctx.clearRect(0.0, 0.0, ctx.Canvas.width, ctx.Canvas.height)

    // Set font
    ctx.Ctx.font <- sprintf "%fpx %s" ctx.FontSize ctx.FontFamily
    ctx.Ctx.fillStyle <- U3.Case1 "#000000"

    let settings = layoutSettings ctx doc
    let cachedEntry = layoutCacheEntryFor ctx doc
    let cached = cachedEntry.Layouts
    let layouts, commands, _ =
        collectDocumentCommandsWithLayouts settings ctx.FontFamily (measureWithContext ctx) (measureRangeWithContext ctx) doc (Some cached)
    renderPages ctx doc layouts
    commands |> List.iter (renderPaintCommand ctx)
    // Footnotes and endnotes are laid out and painted via the shared Layout/Paint
    // pass (so the browser canvas and the PDF/native paths match).
    renderCommentAnchors ctx doc layouts

    // Reset style
    ctx.Ctx.fillStyle <- U3.Case1 "#000000"

    selectionCommandsWithRangeLookup settings (measureRangeWithContext ctx) doc (layoutLookupForEntry cachedEntry)
    |> List.iter (renderPaintCommand ctx)

    markupMarkerCommands settings layouts
    |> List.iter (renderPaintCommand ctx)

    cachedEntry

/// Render document to canvas
let render (ctx: RenderContext) (doc: Document) =
    renderCore ctx doc |> ignore

/// Convert mouse position to document position
let hitTest (ctx: RenderContext) (doc: Document) (mouseX: float) (mouseY: float) =
    RedBat.Word.Layout.hitTestLayoutArrayWithRange (measureRangeWithContext ctx) (layoutSettings ctx doc).PaddingLeft (layoutArrayFor ctx doc) mouseX mouseY

/// Hit test for images - returns blockId, offset, and image dimensions if clicking on an image
let hitTestForImage (ctx: RenderContext) (doc: Document) (mouseX: float) (mouseY: float) : (string * int * float * float * float * float) option =
    RedBat.Word.Layout.hitTestForImageLayoutArray (layoutSettings ctx doc).PaddingLeft (layoutArrayFor ctx doc) mouseX mouseY

/// Render document with hover indicators for images
let renderWithHover (ctx: RenderContext) (doc: Document) (hoveredImage: (string * int) option) =
    // Render the document normally
    let cachedEntry = renderCore ctx doc

    hoverCommandsLookup (layoutSettings ctx doc) (layoutLookupForEntry cachedEntry) hoveredImage
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
