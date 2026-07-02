module RedBat.Word.Document

open System.Collections.Generic
open RedBat.Word.TextBuffer

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private clearArray (items: 'T array) =
    for i = 0 to items.Length - 1 do
        items.[i] <- Unchecked.defaultof<'T>

let private prependResizeArray (items: ResizeArray<'T>) tail =
    let mutable result = tail
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private tryRemoveListItemAtV index items =
    let rec removeSmall remaining items =
        match remaining, items with
        | 0, _ :: tail -> ValueSome tail
        | remaining, item :: tail when remaining > 0 ->
            match removeSmall (remaining - 1) tail with
            | ValueSome tail -> ValueSome (item :: tail)
            | ValueNone -> ValueNone
        | _ -> ValueNone

    if index < 0 then
        ValueNone
    elif index <= 8 then
        removeSmall index items
    else
        let prefix = ResizeArray<'T>(index)
        let rec loop position items =
            match items with
            | [] -> ValueNone
            | _ :: tail when position = index ->
                ValueSome (prependResizeArray prefix tail)
            | item :: tail ->
                prefix.Add item
                loop (position + 1) tail
        loop 0 items

let private tryInsertListItemAfterAtV index createItem items =
    let rec insertSmall remaining items =
        match remaining, items with
        | 0, item :: tail -> ValueSome (item :: createItem () :: tail)
        | remaining, item :: tail when remaining > 0 ->
            match insertSmall (remaining - 1) tail with
            | ValueSome tail -> ValueSome (item :: tail)
            | ValueNone -> ValueNone
        | _ -> ValueNone

    if index < 0 then
        ValueNone
    elif index <= 8 then
        insertSmall index items
    else
        let prefix = ResizeArray<'T>(index)
        let rec loop position items =
            match items with
            | [] -> ValueNone
            | item :: tail when position = index ->
                ValueSome (prependResizeArray prefix (item :: createItem () :: tail))
            | item :: tail ->
                prefix.Add item
                loop (position + 1) tail
        loop 0 items

let private stringResizeArrayConcatNewline (items: ResizeArray<string>) =
    if items.Count = 0 then
        ""
    else
        let mutable capacity = items.Count - 1
        for item in items do
            capacity <- capacity + item.Length
        let builder = System.Text.StringBuilder(capacity)
        builder.Append(items.[0]) |> ignore
        for i = 1 to items.Count - 1 do
            builder.Append('\n').Append(items.[i]) |> ignore
        builder.ToString()

/// Block ID for tracking blocks
type BlockId = string

/// Counter for generating unique block IDs
let mutable private blockIdCounter = 0

/// Generate next block ID
let private nextBlockId () =
    blockIdCounter <- blockIdCounter + 1
    sprintf "block-%d" blockIdCounter

/// Text style for runs
type TextStyle = {
    Bold: bool
    Italic: bool
    Underline: bool
    FontSize: float option
    Color: string option
}

/// Default text style
let defaultStyle = {
    Bold = false
    Italic = false
    Underline = false
    FontSize = None
    Color = None
}

/// Text run - contiguous text with same formatting
type TextRun = {
    Text: string
    Style: TextStyle
}

/// List style
type ListStyle =
    | NoList
    | Bullet
    | Numbered

/// Text alignment
type Alignment =
    | Left
    | Center
    | Right
    | Justify

type SectionPageSetup = {
    PageWidth: float
    PageHeight: float
    MarginLeft: float
    MarginRight: float
    MarginTop: float
    MarginBottom: float
    HeaderDistance: float
    FooterDistance: float
    DifferentFirstPage: bool
    EvenAndOddHeaders: bool
}

type ParagraphBorder = {
    Style: string option
    Color: string option
    Width: float option
    Space: float option
}

type ParagraphBorders = {
    Top: ParagraphBorder option
    Right: ParagraphBorder option
    Bottom: ParagraphBorder option
    Left: ParagraphBorder option
}

/// Block metadata
/// Newspaper columns for a section (w:cols). Per-column widths in Widths
/// (empty = equal width).
type SectionColumns = {
    Count: int
    Gap: float          // gap between columns, px (equal-width case)
    Separator: bool     // draw a rule between columns (w:sep)
    Widths: (float * float) list  // per-column (width, spaceAfter) px; empty = equal width
}

/// A w:tab stop. Position is px from the left text margin; Alignment is
/// left|center|right|decimal|bar; Leader is the optional fill (dot|hyphen|…).
type TabStop = {
    Position: float
    Alignment: string
    Leader: string option
}

type BlockMetadata = {
    Style: string option
    ContentControlId: string option
    ContentControlTag: string option
    ContentControlAlias: string option
    ListStyle: ListStyle
    ListLevel: int  // 0 = not in list, 1-9 = nesting level
    ListId: string option
    ListNumberFormat: string option
    ListLevelText: string option
    ListStart: int option
    Alignment: Alignment
    IndentLeft: float option
    IndentRight: float option
    FirstLineIndent: float option
    RightTabPosition: float option
    RightTabLeader: string option
    TabStops: TabStop list
    Rtl: bool   // w:bidi — right-to-left paragraph
    SpacingBefore: float option
    SpacingAfter: float option
    ParagraphMarkFontSize: float option
    ContextualSpacing: bool option
    LineSpacing: float option
    LineSpacingRule: string option
    WidowControl: bool option
    BackgroundColor: string option
    ParagraphBorders: ParagraphBorders option
    ParagraphBorderBefore: float option
    ParagraphBorderAfter: float option
    PageBreakBefore: bool
    PageBreakAfter: bool
    SectionBreak: string option
    SectionColumns: SectionColumns option
    SectionPageNumberStart: int option
    RenderedPageBreakBefore: int option
    SectionPageSetup: SectionPageSetup option
    HeaderReferences: (string * string) list
    FooterReferences: (string * string) list
}

/// Default block metadata
let defaultMetadata = {
    Style = None
    ContentControlId = None
    ContentControlTag = None
    ContentControlAlias = None
    ListStyle = NoList
    ListLevel = 0
    ListId = None
    ListNumberFormat = None
    ListLevelText = None
    ListStart = None
    Alignment = Left
    IndentLeft = None
    IndentRight = None
    FirstLineIndent = None
    RightTabPosition = None
    RightTabLeader = None
    TabStops = []
    Rtl = false
    SpacingBefore = None
    SpacingAfter = None
    ParagraphMarkFontSize = None
    ContextualSpacing = None
    LineSpacing = None
    LineSpacingRule = None
    WidowControl = None
    BackgroundColor = None
    ParagraphBorders = None
    ParagraphBorderBefore = None
    ParagraphBorderAfter = None
    PageBreakBefore = false
    PageBreakAfter = false
    SectionBreak = None
    SectionColumns = None
    SectionPageNumberStart = None
    RenderedPageBreakBefore = None
    SectionPageSetup = None
    HeaderReferences = []
    FooterReferences = []
}

/// One segment of a vector path. Coordinates are in the path's own coordinate
/// space (see ShapeGeometry); the painter scales them to the shape's box.
type PathSeg =
    | PMove of x: float * y: float
    | PLine of x: float * y: float
    | PCubic of c1x: float * c1y: float * c2x: float * c2y: float * x: float * y: float
    | PQuad of cx: float * cy: float * x: float * y: float
    | PClose

/// A gradient colour stop (offset 0..1).
type GradientStop = { Offset: float; Color: string }

/// Linear (angle in degrees) or radial gradient fill.
type GradientSpec = {
    Stops: GradientStop list
    Angle: float
    Radial: bool
}

/// Vector geometry of a drawing shape / text box.
type ShapeGeometry =
    /// Plain rectangle — the legacy text-box shape.
    | GeoRect
    /// An OOXML preset (a:prstGeom@prst, e.g. "ellipse", "roundRect"). Resolved
    /// to a path at paint time by Shapes.presetPath against the shape's w×h.
    | GeoPreset of preset: string
    /// A freeform a:custGeom path in its own coordinate space (coordW × coordH).
    | GeoCustom of segs: PathSeg list * coordW: float * coordH: float

/// Block types
type Block =
    | Paragraph of ParagraphBlock
    | Heading of level: int * ParagraphBlock
    | Table of TableBlock
    | TextBox of TextBoxBlock

and ParagraphBlock = {
    Id: BlockId
    Buffer: TextBuffer
    Metadata: BlockMetadata
}

and TableBlock = {
    Id: BlockId
    Rows: TableRow list
    PreferredWidth: TableWidth option
    Alignment: string option
    LayoutMode: TableLayoutMode option
    GridColumns: TableWidth list
    Borders: TableBorders option
    BackgroundColor: string option
    Shading: Shading option
    CellSpacing: float option
    PaddingTop: float option
    PaddingRight: float option
    PaddingBottom: float option
    PaddingLeft: float option
    Floating: bool
}

and TextBoxBlock = {
    Id: BlockId
    Content: ParagraphBlock list
    X: float option
    Y: float option
    PositionHRelativeFrom: string option
    PositionVRelativeFrom: string option
    Width: float
    Height: float
    FillColor: string option
    BorderColor: string option
    BorderWidth: float option
    RelativeHeight: string option
    Geometry: ShapeGeometry
    Gradient: GradientSpec option
    Rotation: float
}

and TableRow = {
    Cells: TableCell list
    IsHeader: bool
    Height: float option
    HeightRule: string option
    CanSplit: bool option
    CellSpacing: float option
    PaddingTop: float option
    PaddingRight: float option
    PaddingBottom: float option
    PaddingLeft: float option
    GridBefore: int
    GridAfter: int
    WidthBefore: float option
    WidthAfter: float option
    BordersOverride: TableBorders option
}

and TableCell = {
    Content: ParagraphBlock list  // Cells can contain multiple paragraphs
    NestedTables: TableBlock list
    // For each entry in NestedTables, the number of cell paragraphs that
    // precede it in document order. Lets layout/paint interleave paragraphs
    // and nested tables instead of hoisting all tables to the cell top.
    NestedTableAnchors: int list
    ColSpan: int
    BackgroundColor: string option
    Shading: Shading option
    Width: float option
    PreferredWidth: TableWidth option
    BorderColor: string option
    BorderWidth: float option
    Borders: TableBorders option
    VerticalAlignment: string option
    VerticalMerge: TableVerticalMerge option
    PaddingTop: float option
    PaddingRight: float option
    PaddingBottom: float option
    PaddingLeft: float option
}

and Shading = {
    Pattern: string option
    ForegroundColor: string option
    BackgroundColor: string option
}

and TableWidthKind =
    | AutoWidth
    | DxaWidth
    | PercentWidth
    | PointWidth

and TableWidth = {
    Kind: TableWidthKind
    Value: float option
}

and TableLayoutMode =
    | FixedTableLayout
    | AutoFitTableLayout

and TableBorder = {
    Style: string option
    Color: string option
    Width: float option
}

and TableBorders = {
    Top: TableBorder option
    Right: TableBorder option
    Bottom: TableBorder option
    Left: TableBorder option
    InsideH: TableBorder option
    InsideV: TableBorder option
}

and TableVerticalMerge =
    | RestartMerge
    | ContinueMerge

/// Cursor position
[<Struct>]
type Position = {
    BlockId: BlockId
    Offset: int  // Character offset in block
}

/// Selection (anchor = start, focus = end/cursor)
[<Struct>]
type Selection = {
    Anchor: Position
    Focus: Position
}
with
    /// Check if selection is collapsed (cursor only)
    member this.IsCollapsed =
        this.Anchor.BlockId = this.Focus.BlockId &&
        this.Anchor.Offset = this.Focus.Offset

    /// Get normalized range (start, end) accounting for direction - single block only
    member this.Range =
        if this.Anchor.BlockId = this.Focus.BlockId then
            if this.Anchor.Offset <= this.Focus.Offset then
                (this.Anchor, this.Focus)
            else
                (this.Focus, this.Anchor)
        else
            // For now, assume single block
            (this.Anchor, this.Focus)

type BookmarkInfo = {
    Id: string
    Name: string
    BlockId: BlockId
    StartOffset: int
    EndOffset: int
}

type ContentControlInfo = {
    Id: string
    Tag: string
    Alias: string
    BlockId: BlockId
}

type CommentInfo = {
    Id: string
    Text: string
    Author: string option
    Date: string option
    BlockId: BlockId
    StartOffset: int
    EndOffset: int
    Excerpt: string
}

let private documentSourceTextCacheMaxTotalChars = 5000000
let private sourceParagraphCountCacheLimit = 100000

let private sourceStructureVisitedItemLimit maxSourceParagraphs =
    maxSourceParagraphs * 4

let private sourceStructureOverwhelmsParagraphs maxSourceParagraphs paragraphCount visitedItems =
    visitedItems > maxSourceParagraphs
    && paragraphCount * 4 < visitedItems

let private documentAddCapped total value =
    if total > documentSourceTextCacheMaxTotalChars then
        total
    else
        let value = max 0 value
        let remaining = documentSourceTextCacheMaxTotalChars - total
        if value > remaining then documentSourceTextCacheMaxTotalChars + 1
        else total + value

let private documentOptionStringCharCount (value: string option) =
    match value with
    | Some value -> value.Length
    | None -> 0

let private documentSectionColumnsStringCharCount (columns: SectionColumns option) =
    match columns with
    | Some columns ->
        let mutable total = 0
        let mutable seen = 0
        let mutable widths = columns.Widths
        while total <= documentSourceTextCacheMaxTotalChars
              && seen <= sourceParagraphCountCacheLimit
              && not widths.IsEmpty do
            total <- documentAddCapped total 1
            seen <- seen + 1
            widths <- widths.Tail
        if seen > sourceParagraphCountCacheLimit then
            documentSourceTextCacheMaxTotalChars + 1
        else
            total
    | None -> 0

let private documentTextStyleStringCharCount (style: RedBat.Word.TextBuffer.TextStyle) =
    documentOptionStringCharCount style.Style
    + documentOptionStringCharCount style.UnderlineStyle
    + documentOptionStringCharCount style.UnderlineColor
    + documentOptionStringCharCount style.EmphasisMark
    + documentOptionStringCharCount style.FontFamily
    + documentOptionStringCharCount style.Color
    + documentOptionStringCharCount style.Highlight
    + documentOptionStringCharCount style.Hyperlink
    + documentOptionStringCharCount style.CommentId
    + documentOptionStringCharCount style.CommentText
    + documentOptionStringCharCount style.CommentAuthor
    + documentOptionStringCharCount style.CommentDate
    + documentOptionStringCharCount style.NoteId
    + documentOptionStringCharCount style.NoteType
    + documentOptionStringCharCount style.NoteText
    + documentOptionStringCharCount style.BookmarkStartId
    + documentOptionStringCharCount style.BookmarkName
    + documentOptionStringCharCount style.BookmarkEndId
    + documentOptionStringCharCount style.FieldCode
    + documentOptionStringCharCount style.FieldResult
    + documentOptionStringCharCount style.EquationXml
    + documentOptionStringCharCount style.RevisionId
    + documentOptionStringCharCount style.RevisionType
    + documentOptionStringCharCount style.RevisionAuthor
    + documentOptionStringCharCount style.RevisionDate
    + documentOptionStringCharCount style.VertAlign

let private documentImageStringCharCount (image: RedBat.Word.TextBuffer.ImageRun) =
    image.Src.Length
    + documentOptionStringCharCount image.Alt
    + documentOptionStringCharCount image.Name
    + documentOptionStringCharCount image.Title
    + documentOptionStringCharCount image.Wrap
    + documentOptionStringCharCount image.PositionHRelativeFrom
    + documentOptionStringCharCount image.PositionVRelativeFrom
    + documentOptionStringCharCount image.RelativeHeight
    + documentOptionStringCharCount image.CropLeft
    + documentOptionStringCharCount image.CropTop
    + documentOptionStringCharCount image.CropRight
    + documentOptionStringCharCount image.CropBottom

let private documentRunStringCharCount = function
    | Text textRun -> documentTextStyleStringCharCount textRun.Style
    | Image image -> documentImageStringCharCount image

let private documentBufferSourceTextCharCount (buffer: RedBat.Word.TextBuffer.TextBuffer) =
    let limit = documentSourceTextCacheMaxTotalChars + 1
    let mutable total = TextBuffer.length buffer
    let mutable seen = 0
    let mutable remaining = buffer.Runs
    while total <= documentSourceTextCacheMaxTotalChars
          && seen <= sourceParagraphCountCacheLimit
          && not remaining.IsEmpty do
        total <- documentAddCapped total (documentRunStringCharCount remaining.Head)
        seen <- seen + 1
        remaining <- remaining.Tail
    if seen > sourceParagraphCountCacheLimit then limit else min total limit

let private documentParagraphBorderStringCharCount (border: ParagraphBorder option) =
    match border with
    | Some border -> documentOptionStringCharCount border.Style + documentOptionStringCharCount border.Color
    | None -> 0

let private documentParagraphBordersStringCharCount (borders: ParagraphBorders option) =
    match borders with
    | Some borders ->
        documentParagraphBorderStringCharCount borders.Top
        + documentParagraphBorderStringCharCount borders.Right
        + documentParagraphBorderStringCharCount borders.Bottom
        + documentParagraphBorderStringCharCount borders.Left
    | None -> 0

let private documentTabStopStringCharCount (tabStop: TabStop) =
    tabStop.Alignment.Length + documentOptionStringCharCount tabStop.Leader

let private documentParagraphMetadataStringCharCount (metadata: BlockMetadata) =
    let mutable total =
        documentOptionStringCharCount metadata.Style
        + documentOptionStringCharCount metadata.ContentControlId
        + documentOptionStringCharCount metadata.ContentControlTag
        + documentOptionStringCharCount metadata.ContentControlAlias
        + documentOptionStringCharCount metadata.ListId
        + documentOptionStringCharCount metadata.ListNumberFormat
        + documentOptionStringCharCount metadata.ListLevelText
        + documentOptionStringCharCount metadata.RightTabLeader
        + documentOptionStringCharCount metadata.LineSpacingRule
        + documentOptionStringCharCount metadata.BackgroundColor
        + documentParagraphBordersStringCharCount metadata.ParagraphBorders
        + documentOptionStringCharCount metadata.SectionBreak
        + documentSectionColumnsStringCharCount metadata.SectionColumns
    let mutable seen = 0
    let mutable tabStops = metadata.TabStops
    while total <= documentSourceTextCacheMaxTotalChars
          && seen <= sourceParagraphCountCacheLimit
          && not tabStops.IsEmpty do
        total <- documentAddCapped total (documentTabStopStringCharCount tabStops.Head)
        seen <- seen + 1
        tabStops <- tabStops.Tail
    let mutable headerReferences = metadata.HeaderReferences
    while total <= documentSourceTextCacheMaxTotalChars
          && seen <= sourceParagraphCountCacheLimit
          && not headerReferences.IsEmpty do
        let kind, relationshipId = headerReferences.Head
        total <- documentAddCapped total (kind.Length + relationshipId.Length)
        seen <- seen + 1
        headerReferences <- headerReferences.Tail
    let mutable footerReferences = metadata.FooterReferences
    while total <= documentSourceTextCacheMaxTotalChars
          && seen <= sourceParagraphCountCacheLimit
          && not footerReferences.IsEmpty do
        let kind, relationshipId = footerReferences.Head
        total <- documentAddCapped total (kind.Length + relationshipId.Length)
        seen <- seen + 1
        footerReferences <- footerReferences.Tail
    if seen > sourceParagraphCountCacheLimit then
        documentSourceTextCacheMaxTotalChars + 1
    else
        total

let private documentParagraphSourceTextCharCount (paragraph: ParagraphBlock) =
    let mutable total = paragraph.Id.Length
    total <- documentAddCapped total (documentBufferSourceTextCharCount paragraph.Buffer)
    total <- documentAddCapped total (documentParagraphMetadataStringCharCount paragraph.Metadata)
    total

let private documentTableBorderStringCharCount (border: TableBorder option) =
    match border with
    | Some border ->
        let mutable total = 0
        total <- documentAddCapped total (documentOptionStringCharCount border.Style)
        total <- documentAddCapped total (documentOptionStringCharCount border.Color)
        total
    | None -> 0

let private documentTableBordersStringCharCount (borders: TableBorders option) =
    match borders with
    | Some borders ->
        let mutable total = 0
        total <- documentAddCapped total (documentTableBorderStringCharCount borders.Top)
        total <- documentAddCapped total (documentTableBorderStringCharCount borders.Right)
        total <- documentAddCapped total (documentTableBorderStringCharCount borders.Bottom)
        total <- documentAddCapped total (documentTableBorderStringCharCount borders.Left)
        total <- documentAddCapped total (documentTableBorderStringCharCount borders.InsideH)
        total <- documentAddCapped total (documentTableBorderStringCharCount borders.InsideV)
        total
    | None -> 0

let private documentShadingStringCharCount (shading: Shading option) =
    match shading with
    | Some shading ->
        let mutable total = 0
        total <- documentAddCapped total (documentOptionStringCharCount shading.Pattern)
        total <- documentAddCapped total (documentOptionStringCharCount shading.ForegroundColor)
        total <- documentAddCapped total (documentOptionStringCharCount shading.BackgroundColor)
        total
    | None -> 0

let private documentTableStringCharCount (table: TableBlock) =
    let mutable total = table.Id.Length
    total <- documentAddCapped total (documentOptionStringCharCount table.Alignment)
    total <- documentAddCapped total (documentOptionStringCharCount table.BackgroundColor)
    total <- documentAddCapped total (documentShadingStringCharCount table.Shading)
    total <- documentAddCapped total (documentTableBordersStringCharCount table.Borders)
    total

let private documentTableRowStringCharCount (row: TableRow) =
    let mutable total = 0
    total <- documentAddCapped total (documentOptionStringCharCount row.HeightRule)
    total <- documentAddCapped total (documentTableBordersStringCharCount row.BordersOverride)
    total

let private documentTableCellStringCharCount (cell: TableCell) =
    let mutable total = 0
    total <- documentAddCapped total (documentOptionStringCharCount cell.BackgroundColor)
    total <- documentAddCapped total (documentShadingStringCharCount cell.Shading)
    total <- documentAddCapped total (documentOptionStringCharCount cell.BorderColor)
    total <- documentAddCapped total (documentTableBordersStringCharCount cell.Borders)
    total <- documentAddCapped total (documentOptionStringCharCount cell.VerticalAlignment)
    total

let private documentPathSegmentsStringCharCount (segments: PathSeg list) =
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = segments
    while total <= documentSourceTextCacheMaxTotalChars
          && seen <= sourceParagraphCountCacheLimit
          && not remaining.IsEmpty do
        total <- documentAddCapped total 1
        seen <- seen + 1
        remaining <- remaining.Tail
    if seen > sourceParagraphCountCacheLimit then
        documentSourceTextCacheMaxTotalChars + 1
    else
        total

let private documentShapeGeometryStringCharCount = function
    | GeoRect -> 0
    | GeoPreset preset -> preset.Length
    | GeoCustom (segments, _, _) -> documentPathSegmentsStringCharCount segments

let private documentGradientStringCharCount (gradient: GradientSpec option) =
    match gradient with
    | Some gradient ->
        let mutable total = 0
        let mutable seen = 0
        let mutable remaining = gradient.Stops
        while total <= documentSourceTextCacheMaxTotalChars
              && seen <= sourceParagraphCountCacheLimit
              && not remaining.IsEmpty do
            total <- documentAddCapped total remaining.Head.Color.Length
            seen <- seen + 1
            remaining <- remaining.Tail
        if seen > sourceParagraphCountCacheLimit then
            documentSourceTextCacheMaxTotalChars + 1
        else
            total
    | None -> 0

let private documentTextBoxStringCharCount (textBox: TextBoxBlock) =
    let mutable total = textBox.Id.Length
    total <- documentAddCapped total (documentOptionStringCharCount textBox.PositionHRelativeFrom)
    total <- documentAddCapped total (documentOptionStringCharCount textBox.PositionVRelativeFrom)
    total <- documentAddCapped total (documentOptionStringCharCount textBox.FillColor)
    total <- documentAddCapped total (documentOptionStringCharCount textBox.BorderColor)
    total <- documentAddCapped total (documentOptionStringCharCount textBox.RelativeHeight)
    total <- documentAddCapped total (documentShapeGeometryStringCharCount textBox.Geometry)
    total <- documentAddCapped total (documentGradientStringCharCount textBox.Gradient)
    total

let private documentBlocksSourceTextCharCountWithLimit maxVisitedItems (blocks: Block list) =
    let limit = documentSourceTextCacheMaxTotalChars + 1
    let maxVisitedItems = sourceStructureVisitedItemLimit maxVisitedItems
    let mutable total = 0
    let mutable visitedItems = 0
    let canContinue () =
        total <= documentSourceTextCacheMaxTotalChars
        && visitedItems <= maxVisitedItems
    let visitItem () =
        visitedItems <- visitedItems + 1
    let addParagraph paragraph =
        if canContinue () then
            visitItem ()
            total <- documentAddCapped total (documentParagraphSourceTextCharCount paragraph)
    let addTable table =
        let stack = Stack<TableBlock>()
        stack.Push table
        while stack.Count > 0 && canContinue () do
            let table = stack.Pop()
            visitItem ()
            total <- documentAddCapped total (documentTableStringCharCount table)
            let mutable rows = table.Rows
            while canContinue () && not rows.IsEmpty do
                let row = rows.Head
                visitItem ()
                total <- documentAddCapped total (documentTableRowStringCharCount row)
                let mutable cells = row.Cells
                while canContinue () && not cells.IsEmpty do
                    let cell = cells.Head
                    visitItem ()
                    total <- documentAddCapped total (documentTableCellStringCharCount cell)
                    let mutable paragraphs = cell.Content
                    while canContinue () && not paragraphs.IsEmpty do
                        addParagraph paragraphs.Head
                        paragraphs <- paragraphs.Tail
                    let mutable nestedTables = cell.NestedTables
                    while canContinue () && not nestedTables.IsEmpty do
                        stack.Push nestedTables.Head
                        nestedTables <- nestedTables.Tail
                    cells <- cells.Tail
                rows <- rows.Tail
    let mutable remaining = blocks
    while canContinue () && not remaining.IsEmpty do
        match remaining.Head with
        | Paragraph paragraph
        | Heading (_, paragraph) ->
            addParagraph paragraph
        | Table table ->
            addTable table
        | TextBox textBox ->
            visitItem ()
            total <- documentAddCapped total (documentTextBoxStringCharCount textBox)
            let mutable paragraphs = textBox.Content
            while canContinue () && not paragraphs.IsEmpty do
                addParagraph paragraphs.Head
                paragraphs <- paragraphs.Tail
        remaining <- remaining.Tail
    if visitedItems > maxVisitedItems then limit else min total limit

let private documentBlocksSourceTextCharCount (blocks: Block list) =
    documentBlocksSourceTextCharCountWithLimit sourceParagraphCountCacheLimit blocks

let private documentBlocksSourceTextCharCountFor maxTotalSourceParagraphs sourceParagraphCount blocks =
    if sourceParagraphCount > maxTotalSourceParagraphs then
        documentSourceTextCacheMaxTotalChars + 1
    else
        documentBlocksSourceTextCharCountWithLimit maxTotalSourceParagraphs blocks

type private BookmarkInfoCache = {
    BookmarkSourceBlocks: Block list
    SourceParagraphCount: int
    SourceTextCharCount: int
    BookmarkInfoCount: int
    BookmarkInfos: BookmarkInfo list
}

type private ContentControlInfoCache = {
    ContentControlSourceBlocks: Block list
    SourceParagraphCount: int
    SourceTextCharCount: int
    ContentControlInfoCount: int
    ContentControlInfos: ContentControlInfo list
}

type private CommentInfoCache = {
    CommentSourceBlocks: Block list
    SourceParagraphCount: int
    SourceTextCharCount: int
    CommentInfoCount: int
    CommentInfos: CommentInfo list
}

let mutable private bookmarkInfoCache: BookmarkInfoCache list = []
let mutable private contentControlInfoCache: ContentControlInfoCache list = []
let mutable private commentInfoCache: CommentInfoCache list = []
let private documentInfoBlocksComparer =
    { new IEqualityComparer<Block list> with
        member _.Equals(left, right) = obj.ReferenceEquals(left, right)
        member _.GetHashCode(value) = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value) }
let private bookmarkInfoCacheByBlocks = Dictionary<Block list, BookmarkInfoCache>(documentInfoBlocksComparer)
let private bookmarkInfoFastSlots = Array.zeroCreate<BookmarkInfoCache> 256
let private contentControlInfoCacheByBlocks = Dictionary<Block list, ContentControlInfoCache>(documentInfoBlocksComparer)
let private contentControlInfoFastSlots = Array.zeroCreate<ContentControlInfoCache> 256
let private commentInfoCacheByBlocks = Dictionary<Block list, CommentInfoCache>(documentInfoBlocksComparer)
let private commentInfoFastSlots = Array.zeroCreate<CommentInfoCache> 256
let private documentInfoCacheMaxEntries = 256
let private documentInfoCacheMaxTotalInfos = 200000
let private documentInfoCacheMaxTotalBlocks = 100000
let private documentInfoCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars

let private bookmarkInfoCacheMatches blocks (cached: BookmarkInfoCache) =
    obj.ReferenceEquals(cached.BookmarkSourceBlocks, blocks)

let private contentControlInfoCacheMatches blocks (cached: ContentControlInfoCache) =
    obj.ReferenceEquals(cached.ContentControlSourceBlocks, blocks)

let private commentInfoCacheMatches blocks (cached: CommentInfoCache) =
    obj.ReferenceEquals(cached.CommentSourceBlocks, blocks)

let private setBookmarkInfoCache entries =
    bookmarkInfoCache <- entries
    bookmarkInfoCacheByBlocks.Clear()
    clearArray bookmarkInfoFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<BookmarkInfoCache>
        if not (bookmarkInfoCacheByBlocks.TryGetValue(entry.BookmarkSourceBlocks, &existing)) then
            bookmarkInfoCacheByBlocks.[entry.BookmarkSourceBlocks] <- entry
            let slot = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(entry.BookmarkSourceBlocks) &&& (bookmarkInfoFastSlots.Length - 1)
            bookmarkInfoFastSlots.[slot] <- entry

let private setContentControlInfoCache entries =
    contentControlInfoCache <- entries
    contentControlInfoCacheByBlocks.Clear()
    clearArray contentControlInfoFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<ContentControlInfoCache>
        if not (contentControlInfoCacheByBlocks.TryGetValue(entry.ContentControlSourceBlocks, &existing)) then
            contentControlInfoCacheByBlocks.[entry.ContentControlSourceBlocks] <- entry
            let slot = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(entry.ContentControlSourceBlocks) &&& (contentControlInfoFastSlots.Length - 1)
            contentControlInfoFastSlots.[slot] <- entry

let private setCommentInfoCache entries =
    commentInfoCache <- entries
    commentInfoCacheByBlocks.Clear()
    clearArray commentInfoFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<CommentInfoCache>
        if not (commentInfoCacheByBlocks.TryGetValue(entry.CommentSourceBlocks, &existing)) then
            commentInfoCacheByBlocks.[entry.CommentSourceBlocks] <- entry
            let slot = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(entry.CommentSourceBlocks) &&& (commentInfoFastSlots.Length - 1)
            commentInfoFastSlots.[slot] <- entry

let private documentInfoSlot blocks slotCount =
    System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(blocks) &&& (slotCount - 1)

let private tryGetBookmarkInfoCache blocks =
    let slot = documentInfoSlot blocks bookmarkInfoFastSlots.Length
    let fast = bookmarkInfoFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && bookmarkInfoCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<BookmarkInfoCache>
        if bookmarkInfoCacheByBlocks.TryGetValue(blocks, &cached) then
            bookmarkInfoFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private tryGetContentControlInfoCache blocks =
    let slot = documentInfoSlot blocks contentControlInfoFastSlots.Length
    let fast = contentControlInfoFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && contentControlInfoCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<ContentControlInfoCache>
        if contentControlInfoCacheByBlocks.TryGetValue(blocks, &cached) then
            contentControlInfoFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private tryGetCommentInfoCache blocks =
    let slot = documentInfoSlot blocks commentInfoFastSlots.Length
    let fast = commentInfoFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && commentInfoCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<CommentInfoCache>
        if commentInfoCacheByBlocks.TryGetValue(blocks, &cached) then
            commentInfoFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private bookmarkInfoCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<BookmarkInfoCache>
    if bookmarkInfoCacheByBlocks.TryGetValue(blocks, &cached) then
        bookmarkInfoCache
        |> List.filter (bookmarkInfoCacheMatches blocks >> not)
    else
        bookmarkInfoCache

let private removeBookmarkInfoCacheForBlocks blocks =
    let entries = bookmarkInfoCacheWithoutBlocks blocks
    if not (obj.ReferenceEquals(entries, bookmarkInfoCache)) then
        setBookmarkInfoCache entries

let private contentControlInfoCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<ContentControlInfoCache>
    if contentControlInfoCacheByBlocks.TryGetValue(blocks, &cached) then
        contentControlInfoCache
        |> List.filter (contentControlInfoCacheMatches blocks >> not)
    else
        contentControlInfoCache

let private removeContentControlInfoCacheForBlocks blocks =
    let entries = contentControlInfoCacheWithoutBlocks blocks
    if not (obj.ReferenceEquals(entries, contentControlInfoCache)) then
        setContentControlInfoCache entries

let private commentInfoCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<CommentInfoCache>
    if commentInfoCacheByBlocks.TryGetValue(blocks, &cached) then
        commentInfoCache
        |> List.filter (commentInfoCacheMatches blocks >> not)
    else
        commentInfoCache

let private removeCommentInfoCacheForBlocks blocks =
    let entries = commentInfoCacheWithoutBlocks blocks
    if not (obj.ReferenceEquals(entries, commentInfoCache)) then
        setCommentInfoCache entries

let private documentInfoCacheRetainable infoCount blockCount sourceTextCharCount entry =
    infoCount entry <= documentInfoCacheMaxTotalInfos
    && blockCount entry <= documentInfoCacheMaxTotalBlocks
    && sourceTextCharCount entry <= documentInfoCacheMaxTotalSourceTextChars

let private bookmarkInfoCacheRetainable entry =
    documentInfoCacheRetainable
        (fun entry -> entry.BookmarkInfoCount)
        (fun entry -> entry.SourceParagraphCount)
        (fun entry -> entry.SourceTextCharCount)
        entry

let private contentControlInfoCacheRetainable entry =
    documentInfoCacheRetainable
        (fun entry -> entry.ContentControlInfoCount)
        (fun entry -> entry.SourceParagraphCount)
        (fun entry -> entry.SourceTextCharCount)
        entry

let private commentInfoCacheRetainable entry =
    documentInfoCacheRetainable
        (fun entry -> entry.CommentInfoCount)
        (fun entry -> entry.SourceParagraphCount)
        (fun entry -> entry.SourceTextCharCount)
        entry

let private documentInfoCountTableParagraphs (table: TableBlock) =
    let mutable count = 0
    let mutable visitedItems = 0
    let canContinue () =
        count <= documentInfoCacheMaxTotalBlocks
        && visitedItems <= documentInfoCacheMaxTotalBlocks
    let visitItem () =
        visitedItems <- visitedItems + 1
    let addParagraphs (paragraphs: ParagraphBlock list) =
        let mutable remaining = paragraphs
        while canContinue () && not remaining.IsEmpty do
            count <- count + 1
            remaining <- remaining.Tail
    let stack = Stack<TableBlock>()
    stack.Push table
    while canContinue () && stack.Count > 0 do
        let table = stack.Pop()
        visitItem ()
        let mutable rows = table.Rows
        while canContinue () && not rows.IsEmpty do
            let row = rows.Head
            visitItem ()
            let mutable cells = row.Cells
            while canContinue () && not cells.IsEmpty do
                let cell = cells.Head
                visitItem ()
                addParagraphs cell.Content
                let mutable nestedTables = cell.NestedTables
                while canContinue () && not nestedTables.IsEmpty do
                    stack.Push nestedTables.Head
                    nestedTables <- nestedTables.Tail
                cells <- cells.Tail
            rows <- rows.Tail
    if visitedItems > documentInfoCacheMaxTotalBlocks then
        documentInfoCacheMaxTotalBlocks + 1
    else
        min count (documentInfoCacheMaxTotalBlocks + 1)

let private documentInfoCountSourceParagraphs (blocks: Block list) =
    let mutable count = 0
    let mutable visitedItems = 0
    let canContinue () =
        count <= documentInfoCacheMaxTotalBlocks
        && visitedItems <= documentInfoCacheMaxTotalBlocks
    let visitItem () =
        visitedItems <- visitedItems + 1
    let addParagraphs (paragraphs: ParagraphBlock list) =
        let mutable remainingParagraphs = paragraphs
        while canContinue () && not remainingParagraphs.IsEmpty do
            count <- count + 1
            remainingParagraphs <- remainingParagraphs.Tail
    let mutable remaining = blocks
    while canContinue () && not remaining.IsEmpty do
        visitItem ()
        match remaining.Head with
        | Paragraph _
        | Heading _ ->
            count <- count + 1
        | Table table ->
            count <- count + documentInfoCountTableParagraphs table
        | TextBox textBox ->
            addParagraphs textBox.Content
        remaining <- remaining.Tail
    if visitedItems > documentInfoCacheMaxTotalBlocks then
        documentInfoCacheMaxTotalBlocks + 1
    else
        min count (documentInfoCacheMaxTotalBlocks + 1)

let private documentInfoSourceParagraphCountFor infoCount blocks =
    if infoCount > documentInfoCacheMaxTotalInfos then
        documentInfoCacheMaxTotalBlocks + 1
    else
        documentInfoCountSourceParagraphs blocks

let private documentInfoSourceTextCharCountFor infoCount sourceParagraphCount blocks =
    if infoCount > documentInfoCacheMaxTotalInfos
       || sourceParagraphCount > documentInfoCacheMaxTotalBlocks then
        documentInfoCacheMaxTotalSourceTextChars + 1
    else
        documentBlocksSourceTextCharCountFor
            documentInfoCacheMaxTotalBlocks
            sourceParagraphCount
            blocks

let private tryDocumentInfoCachedSourceMetrics blocks =
    match tryGetBookmarkInfoCache blocks with
    | ValueSome cached -> ValueSome (struct (cached.SourceParagraphCount, cached.SourceTextCharCount))
    | ValueNone ->
        match tryGetContentControlInfoCache blocks with
        | ValueSome cached -> ValueSome (struct (cached.SourceParagraphCount, cached.SourceTextCharCount))
        | ValueNone ->
            match tryGetCommentInfoCache blocks with
            | ValueSome cached -> ValueSome (struct (cached.SourceParagraphCount, cached.SourceTextCharCount))
            | ValueNone -> ValueNone

let private documentInfoSourceMetricsFor infoCount blocks =
    if infoCount > documentInfoCacheMaxTotalInfos then
        struct (documentInfoCacheMaxTotalBlocks + 1, documentInfoCacheMaxTotalSourceTextChars + 1)
    else
        match tryDocumentInfoCachedSourceMetrics blocks with
        | ValueSome metrics -> metrics
        | ValueNone ->
            let sourceParagraphCount = documentInfoSourceParagraphCountFor infoCount blocks
            let sourceTextCharCount = documentInfoSourceTextCharCountFor infoCount sourceParagraphCount blocks
            struct (sourceParagraphCount, sourceTextCharCount)

let private trimDocumentInfoCache infoCount blockCount sourceTextCharCount entries =
    let rec loop keptCount keptInfos keptBlocks keptSourceTextChars acc remaining =
        match remaining with
        | [] -> List.rev acc
        | candidate :: tail ->
            let candidateInfos = max 1 (infoCount candidate)
            let candidateBlocks = max 1 (blockCount candidate)
            let candidateSourceTextChars = max 1 (sourceTextCharCount candidate)
            if documentInfoCacheRetainable infoCount blockCount sourceTextCharCount candidate
               && keptCount < documentInfoCacheMaxEntries
               && keptInfos + candidateInfos <= documentInfoCacheMaxTotalInfos
               && keptBlocks + candidateBlocks <= documentInfoCacheMaxTotalBlocks
               && keptSourceTextChars + candidateSourceTextChars <= documentInfoCacheMaxTotalSourceTextChars then
                loop
                    (keptCount + 1)
                    (keptInfos + candidateInfos)
                    (keptBlocks + candidateBlocks)
                    (keptSourceTextChars + candidateSourceTextChars)
                    (candidate :: acc)
                    tail
            else
                loop keptCount keptInfos keptBlocks keptSourceTextChars acc tail
    loop 0 0 0 0 [] entries

let private rememberBookmarkInfos cached =
    if not (bookmarkInfoCacheRetainable cached) then
        removeBookmarkInfoCacheForBlocks cached.BookmarkSourceBlocks
    else
        (cached
         :: bookmarkInfoCacheWithoutBlocks cached.BookmarkSourceBlocks)
        |> trimDocumentInfoCache (fun entry -> entry.BookmarkInfoCount) (fun entry -> entry.SourceParagraphCount) (fun entry -> entry.SourceTextCharCount)
        |> setBookmarkInfoCache

let private rememberContentControlInfos cached =
    if not (contentControlInfoCacheRetainable cached) then
        removeContentControlInfoCacheForBlocks cached.ContentControlSourceBlocks
    else
        (cached
         :: contentControlInfoCacheWithoutBlocks cached.ContentControlSourceBlocks)
        |> trimDocumentInfoCache (fun entry -> entry.ContentControlInfoCount) (fun entry -> entry.SourceParagraphCount) (fun entry -> entry.SourceTextCharCount)
        |> setContentControlInfoCache

let private rememberCommentInfos cached =
    if not (commentInfoCacheRetainable cached) then
        removeCommentInfoCacheForBlocks cached.CommentSourceBlocks
    else
        (cached
         :: commentInfoCacheWithoutBlocks cached.CommentSourceBlocks)
        |> trimDocumentInfoCache (fun entry -> entry.CommentInfoCount) (fun entry -> entry.SourceParagraphCount) (fun entry -> entry.SourceTextCharCount)
        |> setCommentInfoCache

/// Document
type HeaderFooterPart = {
    Kind: string
    RelationshipId: string option
    Blocks: Block list
}


type PageSetup = {
    PageWidth: float
    PageHeight: float
    MarginLeft: float
    MarginRight: float
    MarginTop: float
    MarginBottom: float
    HeaderDistance: float
    FooterDistance: float
    DifferentFirstPage: bool
    EvenAndOddHeaders: bool
    Columns: SectionColumns option
    SectionStart: string option
}

type NamedStyle = {
    Id: string
    Name: string option
    StyleType: string
    IsDefault: bool
    BasedOn: string option
    Next: string option
    Linked: string option
    IsQuickStyle: bool
    Aliases: string option
    UiPriority: int option
    SemiHidden: bool
    UnhideWhenUsed: bool
    Text: TextBuffer.TextStyle option
    TextBold: bool option
    TextItalic: bool option
    TextUnderline: bool option
    TextAllCaps: bool option
    TextSmallCaps: bool option
    TextStrike: bool option
    SpacingBefore: float option
    SpacingAfter: float option
    ContextualSpacing: bool option
    LineSpacing: float option
    LineSpacingRule: string option
    Alignment: Alignment option
    IndentLeft: float option
    IndentRight: float option
    FirstLineIndent: float option
    BackgroundColor: string option
    OutlineLevel: int option
    KeepNext: bool option
    KeepLines: bool option
    PageBreakBefore: bool option
    WidowControl: bool option
}

let defaultPageSetup = {
    PageWidth = 816.0
    PageHeight = 1056.0
    MarginLeft = 96.0
    MarginRight = 96.0
    MarginTop = 96.0
    MarginBottom = 96.0
    HeaderDistance = 48.0
    FooterDistance = 48.0
    DifferentFirstPage = false
    EvenAndOddHeaders = false
    Columns = None
    SectionStart = None
}

type Document = {
    Blocks: Block list
    Selection: Selection option
    PageSetup: PageSetup
    Styles: NamedStyle list
    Headers: ParagraphBlock list
    Footers: ParagraphBlock list
    HeaderParts: HeaderFooterPart list
    FooterParts: HeaderFooterPart list
}

let private documentReferenceComparer =
    { new IEqualityComparer<Document> with
        member _.Equals(left, right) = obj.ReferenceEquals(left, right)
        member _.GetHashCode(value) = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value) }

let setPageSetup (setup: PageSetup) (doc: Document) =
    if doc.PageSetup = setup then doc else { doc with PageSetup = setup }

type FindMatch = {
    BlockId: BlockId
    Start: int
    End: int
    Text: string
}

let private countTableParagraphsWithLimit maxSourceParagraphs (table: TableBlock) =
    let mutable count = 0
    let mutable visitedItems = 0
    let maxVisitedItems = sourceStructureVisitedItemLimit maxSourceParagraphs
    let canContinue () =
        count <= maxSourceParagraphs
        && visitedItems <= maxVisitedItems
    let visitItem () =
        visitedItems <- visitedItems + 1
    let addParagraphs (paragraphs: ParagraphBlock list) =
        let mutable remaining = paragraphs
        while canContinue () && not remaining.IsEmpty do
            count <- count + 1
            remaining <- remaining.Tail
    let stack = Stack<TableBlock>()
    stack.Push table
    while canContinue () && stack.Count > 0 do
        let table = stack.Pop()
        visitItem ()
        let mutable rows = table.Rows
        while canContinue () && not rows.IsEmpty do
            let row = rows.Head
            visitItem ()
            let mutable cells = row.Cells
            while canContinue () && not cells.IsEmpty do
                let cell = cells.Head
                visitItem ()
                addParagraphs cell.Content
                let mutable nestedTables = cell.NestedTables
                while canContinue () && not nestedTables.IsEmpty do
                    stack.Push nestedTables.Head
                    nestedTables <- nestedTables.Tail
                cells <- cells.Tail
            rows <- rows.Tail
    if visitedItems > maxVisitedItems
       || sourceStructureOverwhelmsParagraphs maxSourceParagraphs count visitedItems then
        maxSourceParagraphs + 1
    else
        min count (maxSourceParagraphs + 1)

let private countTableParagraphs (table: TableBlock) =
    countTableParagraphsWithLimit sourceParagraphCountCacheLimit table

let private countSourceParagraphsWithLimit maxSourceParagraphs (blocks: Block list) =
    let mutable count = 0
    let mutable visitedItems = 0
    let maxVisitedItems = sourceStructureVisitedItemLimit maxSourceParagraphs
    let canContinue () =
        count <= maxSourceParagraphs
        && visitedItems <= maxVisitedItems
    let visitItem () =
        visitedItems <- visitedItems + 1
    let addParagraphs (paragraphs: ParagraphBlock list) =
        let mutable remainingParagraphs = paragraphs
        while canContinue () && not remainingParagraphs.IsEmpty do
            count <- count + 1
            remainingParagraphs <- remainingParagraphs.Tail
    let mutable remaining = blocks
    while canContinue () && not remaining.IsEmpty do
        visitItem ()
        match remaining.Head with
        | Paragraph _
        | Heading _ ->
            count <- count + 1
        | Table table ->
            count <- count + countTableParagraphsWithLimit maxSourceParagraphs table
        | TextBox textBox ->
            addParagraphs textBox.Content
        remaining <- remaining.Tail
    if visitedItems > maxVisitedItems
       || sourceStructureOverwhelmsParagraphs maxSourceParagraphs count visitedItems then
        maxSourceParagraphs + 1
    else
        min count (maxSourceParagraphs + 1)

let private countSourceParagraphs (blocks: Block list) =
    countSourceParagraphsWithLimit sourceParagraphCountCacheLimit blocks

type RevisionInfo = {
    BlockId: BlockId
    Start: int
    End: int
    RevisionType: string
    Author: string option
    Date: string option
    Text: string
}

type private RevisionInfoCache = {
    SourceBlocks: Block list
    SourceParagraphCount: int
    SourceTextCharCount: int
    InfoCount: int
    Infos: RevisionInfo list
}

let mutable private revisionInfoCache: RevisionInfoCache list = []
let private revisionInfoCacheByBlocks = Dictionary<Block list, RevisionInfoCache>(documentInfoBlocksComparer)
let private revisionInfoFastSlots = Array.zeroCreate<RevisionInfoCache> 256
let private revisionInfoCacheMaxEntries = 256
let private revisionInfoCacheMaxTotalInfos = 200000
let private revisionInfoCacheMaxTotalSourceParagraphs = sourceParagraphCountCacheLimit
let private revisionInfoCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars

let private revisionInfoCacheMatches blocks (cached: RevisionInfoCache) =
    obj.ReferenceEquals(cached.SourceBlocks, blocks)

let private revisionInfoCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<RevisionInfoCache>
    if revisionInfoCacheByBlocks.TryGetValue(blocks, &cached)
       && revisionInfoCacheMatches blocks cached then
        revisionInfoCache
        |> List.filter (revisionInfoCacheMatches blocks >> not)
    else
        revisionInfoCache

let private setRevisionInfoCache entries =
    revisionInfoCache <- entries
    revisionInfoCacheByBlocks.Clear()
    clearArray revisionInfoFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<RevisionInfoCache>
        if not (revisionInfoCacheByBlocks.TryGetValue(entry.SourceBlocks, &existing)) then
            revisionInfoCacheByBlocks.[entry.SourceBlocks] <- entry
            let slot = documentInfoSlot entry.SourceBlocks revisionInfoFastSlots.Length
            revisionInfoFastSlots.[slot] <- entry

let private tryGetRevisionInfoCache blocks =
    let slot = documentInfoSlot blocks revisionInfoFastSlots.Length
    let fast = revisionInfoFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && revisionInfoCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<RevisionInfoCache>
        if revisionInfoCacheByBlocks.TryGetValue(blocks, &cached) then
            revisionInfoFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private removeRevisionInfoCacheForBlocks blocks =
    let entries = revisionInfoCacheWithoutBlocks blocks
    if not (obj.ReferenceEquals(entries, revisionInfoCache)) then
        setRevisionInfoCache entries

let private revisionInfoCacheRetainable (cached: RevisionInfoCache) =
    cached.InfoCount <= revisionInfoCacheMaxTotalInfos
    && cached.SourceParagraphCount <= revisionInfoCacheMaxTotalSourceParagraphs
    && cached.SourceTextCharCount <= revisionInfoCacheMaxTotalSourceTextChars

let private revisionInfoSourceParagraphCountFor infoCount blocks =
    if infoCount > revisionInfoCacheMaxTotalInfos then
        revisionInfoCacheMaxTotalSourceParagraphs + 1
    else
        countSourceParagraphs blocks

let private revisionInfoSourceTextCharCountFor infoCount sourceParagraphCount blocks =
    if infoCount > revisionInfoCacheMaxTotalInfos
       || sourceParagraphCount > revisionInfoCacheMaxTotalSourceParagraphs then
        revisionInfoCacheMaxTotalSourceTextChars + 1
    else
        documentBlocksSourceTextCharCountFor
            revisionInfoCacheMaxTotalSourceParagraphs
            sourceParagraphCount
            blocks

let private rememberRevisionInfos cached =
    if not (revisionInfoCacheRetainable cached) then
        removeRevisionInfoCacheForBlocks cached.SourceBlocks
    else
        let trim entries =
            let rec loop keptCount keptInfos keptBlocks keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateInfos = max 1 candidate.InfoCount
                    let candidateBlocks = max 1 candidate.SourceParagraphCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if revisionInfoCacheRetainable candidate
                       && keptCount < revisionInfoCacheMaxEntries
                       && keptInfos + candidateInfos <= revisionInfoCacheMaxTotalInfos
                       && keptBlocks + candidateBlocks <= revisionInfoCacheMaxTotalSourceParagraphs
                       && keptSourceTextChars + candidateSourceTextChars <= revisionInfoCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptInfos + candidateInfos)
                            (keptBlocks + candidateBlocks)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptInfos keptBlocks keptSourceTextChars acc tail
            loop 0 0 0 0 [] entries
        (cached
         :: revisionInfoCacheWithoutBlocks cached.SourceBlocks)
        |> trim
        |> setRevisionInfoCache

type private DocumentTextCache = {
    SourceBlocks: Block list
    SourceParagraphCount: int
    SourceTextCharCount: int
    Text: string
    TextLength: int
}

let mutable private documentTextCache: DocumentTextCache list = []
let private documentTextCacheByBlocks = Dictionary<Block list, DocumentTextCache>(documentInfoBlocksComparer)
let private documentTextFastSlots = Array.zeroCreate<DocumentTextCache> 256
let private documentTextCacheMaxEntries = 256
let private documentTextCacheMaxTotalChars = 5000000
let private documentTextCacheMaxTotalSourceParagraphs = sourceParagraphCountCacheLimit
let private documentTextCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars

let private documentTextCacheMatches blocks (cached: DocumentTextCache) =
    obj.ReferenceEquals(cached.SourceBlocks, blocks)

let private documentTextCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<DocumentTextCache>
    if documentTextCacheByBlocks.TryGetValue(blocks, &cached)
       && documentTextCacheMatches blocks cached then
        documentTextCache
        |> List.filter (documentTextCacheMatches blocks >> not)
    else
        documentTextCache

let private setDocumentTextCache entries =
    documentTextCache <- entries
    documentTextCacheByBlocks.Clear()
    clearArray documentTextFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<DocumentTextCache>
        if not (documentTextCacheByBlocks.TryGetValue(entry.SourceBlocks, &existing)) then
            documentTextCacheByBlocks.[entry.SourceBlocks] <- entry
            let slot = documentInfoSlot entry.SourceBlocks documentTextFastSlots.Length
            documentTextFastSlots.[slot] <- entry

let private tryGetDocumentTextCache blocks =
    let slot = documentInfoSlot blocks documentTextFastSlots.Length
    let fast = documentTextFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && documentTextCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<DocumentTextCache>
        if documentTextCacheByBlocks.TryGetValue(blocks, &cached) then
            documentTextFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private documentTextSourceParagraphCountFor textLength blocks =
    if textLength > documentTextCacheMaxTotalChars then
        documentTextCacheMaxTotalSourceParagraphs + 1
    else
        countSourceParagraphs blocks

let private documentTextSourceTextCharCountFor textLength sourceParagraphCount blocks =
    if textLength > documentTextCacheMaxTotalChars
       || sourceParagraphCount > documentTextCacheMaxTotalSourceParagraphs then
        documentTextCacheMaxTotalSourceTextChars + 1
    else
        documentBlocksSourceTextCharCountFor
            documentTextCacheMaxTotalSourceParagraphs
            sourceParagraphCount
            blocks

let private documentTextCanPreSizeBuilder (blocks: Block list) =
    let mutable count = 0
    let mutable canPreSize = true
    let mutable remaining = blocks
    while canPreSize
          && count <= 2048
          && not remaining.IsEmpty do
        match remaining.Head with
        | Paragraph _
        | Heading _ -> ()
        | Table _
        | TextBox _ -> canPreSize <- false
        count <- count + 1
        remaining <- remaining.Tail
    canPreSize && count <= 2048

let private rememberDocumentText cached =
    let trim entries =
        let rec loop keptCount keptChars keptBlocks keptSourceTextChars acc remaining =
            match remaining with
            | [] -> List.rev acc
            | candidate :: tail ->
                let candidateChars = max 1 candidate.TextLength
                let candidateBlocks = max 1 candidate.SourceParagraphCount
                let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                if keptCount < documentTextCacheMaxEntries
                   && keptChars + candidateChars <= documentTextCacheMaxTotalChars
                   && keptBlocks + candidateBlocks <= documentTextCacheMaxTotalSourceParagraphs
                   && keptSourceTextChars + candidateSourceTextChars <= documentTextCacheMaxTotalSourceTextChars then
                    loop
                        (keptCount + 1)
                        (keptChars + candidateChars)
                        (keptBlocks + candidateBlocks)
                        (keptSourceTextChars + candidateSourceTextChars)
                        (candidate :: acc)
                        tail
                else
                    loop keptCount keptChars keptBlocks keptSourceTextChars acc tail
        loop 0 0 0 0 [] entries
    cached
    :: documentTextCacheWithoutBlocks cached.SourceBlocks
    |> trim
    |> setDocumentTextCache

let private blockIdKeyEquals (left: BlockId) (right: BlockId) =
    obj.ReferenceEquals(left, right)
    || System.String.Equals(left, right, System.StringComparison.Ordinal)

let private blockIdKeyHash (blockId: BlockId) =
    if isNull blockId then
        0
    else
        let length = blockId.Length
        let mutable hashCode = length
        if length > 0 then
            hashCode <- (hashCode * 397) ^^^ int blockId.[0]
        if length > 1 then
            hashCode <- (hashCode * 397) ^^^ int blockId.[length / 4]
            hashCode <- (hashCode * 397) ^^^ int blockId.[length / 2]
            hashCode <- (hashCode * 397) ^^^ int blockId.[(length * 3) / 4]
            hashCode <- (hashCode * 397) ^^^ int blockId.[length - 1]
        if length > 2 then
            hashCode <- (hashCode * 397) ^^^ int blockId.[length - 2]
        if length > 3 then
            hashCode <- (hashCode * 397) ^^^ int blockId.[length - 3]
        if length > 4 then
            hashCode <- (hashCode * 397) ^^^ int blockId.[length - 4]
        hashCode ^^^ (hashCode >>> 16)

let private blockIdLookupComparer =
    { new IEqualityComparer<BlockId> with
        member _.Equals(left, right) = blockIdKeyEquals left right
        member _.GetHashCode(value) = blockIdKeyHash value }

[<Struct>]
type private SelectionTextCacheKey = {
    AnchorBlockId: BlockId
    AnchorOffset: int
    FocusBlockId: BlockId
    FocusOffset: int
}

[<Struct>]
type private SelectionTextCacheEntry = {
    Text: string
    TextLength: int
}

type private SelectionTextFastEntry = {
    Key: SelectionTextCacheKey
    Entry: SelectionTextCacheEntry
}

type private SelectionTextCache = {
    SourceBlocks: Block list
    SourceParagraphCount: int
    SourceTextCharCount: int
    Entries: Dictionary<SelectionTextCacheKey, SelectionTextCacheEntry>
    Order: Queue<SelectionTextCacheKey>
    EntryFastSlots: SelectionTextFastEntry array
    mutable TotalChars: int
}

let mutable private selectionTextCache: SelectionTextCache list = []
let private selectionTextCacheByBlocks = Dictionary<Block list, SelectionTextCache>(documentInfoBlocksComparer)
let private selectionTextFastSlots = Array.zeroCreate<SelectionTextCache> 256
let private selectionTextCacheMaxEntries = 256
let private selectionTextCacheMaxSources = 256
let private selectionTextCacheMaxTotalChars = 5000000
let private selectionTextCacheMaxTotalSourceParagraphs = sourceParagraphCountCacheLimit
let private selectionTextCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars

let private selectionTextKeyHash anchorBlockId anchorOffset focusBlockId focusOffset =
    let mutable hashCode = blockIdKeyHash anchorBlockId
    hashCode <- (hashCode * 397) ^^^ anchorOffset
    hashCode <- (hashCode * 397) ^^^ blockIdKeyHash focusBlockId
    hashCode <- (hashCode * 397) ^^^ focusOffset
    hashCode ^^^ (hashCode >>> 16)

let private selectionTextCacheKeyComparer =
    { new IEqualityComparer<SelectionTextCacheKey> with
        member _.Equals(left, right) =
            blockIdKeyEquals left.AnchorBlockId right.AnchorBlockId
            && left.AnchorOffset = right.AnchorOffset
            && blockIdKeyEquals left.FocusBlockId right.FocusBlockId
            && left.FocusOffset = right.FocusOffset

        member _.GetHashCode(key) =
            selectionTextKeyHash
                key.AnchorBlockId
                key.AnchorOffset
                key.FocusBlockId
                key.FocusOffset }

let private selectionTextEntryFastSlot
    (slots: SelectionTextFastEntry array)
    anchorBlockId
    anchorOffset
    focusBlockId
    focusOffset =
    selectionTextKeyHash anchorBlockId anchorOffset focusBlockId focusOffset &&& (slots.Length - 1)

let private clearSelectionTextEntryFastSlots (slots: SelectionTextFastEntry array) =
    for index = 0 to slots.Length - 1 do
        slots.[index] <- Unchecked.defaultof<SelectionTextFastEntry>

let private selectionTextKeyMatchesSelection (selection: Selection) (key: SelectionTextCacheKey) =
    blockIdKeyEquals key.AnchorBlockId selection.Anchor.BlockId
    && key.AnchorOffset = selection.Anchor.Offset
    && blockIdKeyEquals key.FocusBlockId selection.Focus.BlockId
    && key.FocusOffset = selection.Focus.Offset

let private tryGetSelectionTextEntryFastSlot (cache: SelectionTextCache) (selection: Selection) =
    let slot =
        selectionTextEntryFastSlot
            cache.EntryFastSlots
            selection.Anchor.BlockId
            selection.Anchor.Offset
            selection.Focus.BlockId
            selection.Focus.Offset
    let fast = cache.EntryFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && selectionTextKeyMatchesSelection selection fast.Key then
        ValueSome fast.Entry
    else
        ValueNone

let private setSelectionTextEntryFastSlot (cache: SelectionTextCache) key entry =
    let slot =
        selectionTextEntryFastSlot
            cache.EntryFastSlots
            key.AnchorBlockId
            key.AnchorOffset
            key.FocusBlockId
            key.FocusOffset
    cache.EntryFastSlots.[slot] <- {
        Key = key
        Entry = entry
    }

let private selectionTextCacheMatches blocks (cached: SelectionTextCache) =
    obj.ReferenceEquals(cached.SourceBlocks, blocks)

let private selectionTextCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<SelectionTextCache>
    if selectionTextCacheByBlocks.TryGetValue(blocks, &cached)
       && selectionTextCacheMatches blocks cached then
        selectionTextCache
        |> List.filter (selectionTextCacheMatches blocks >> not)
    else
        selectionTextCache

let private setSelectionTextCache entries =
    selectionTextCache <- entries
    selectionTextCacheByBlocks.Clear()
    clearArray selectionTextFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<SelectionTextCache>
        if not (selectionTextCacheByBlocks.TryGetValue(entry.SourceBlocks, &existing)) then
            selectionTextCacheByBlocks.[entry.SourceBlocks] <- entry
            let slot = documentInfoSlot entry.SourceBlocks selectionTextFastSlots.Length
            selectionTextFastSlots.[slot] <- entry

let private tryGetSelectionTextCache blocks =
    let slot = documentInfoSlot blocks selectionTextFastSlots.Length
    let fast = selectionTextFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && selectionTextCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<SelectionTextCache>
        if selectionTextCacheByBlocks.TryGetValue(blocks, &cached) then
            selectionTextFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private removeSelectionTextCacheForBlocks blocks =
    let entries = selectionTextCacheWithoutBlocks blocks
    if not (obj.ReferenceEquals(entries, selectionTextCache)) then
        setSelectionTextCache entries

let private selectionTextSourceParagraphCountFor textLength blocks =
    if textLength > selectionTextCacheMaxTotalChars then
        selectionTextCacheMaxTotalSourceParagraphs + 1
    else
        countSourceParagraphs blocks

let private selectionTextSourceTextCharCountFor textLength sourceParagraphCount blocks =
    if textLength > selectionTextCacheMaxTotalChars
       || sourceParagraphCount > selectionTextCacheMaxTotalSourceParagraphs then
        selectionTextCacheMaxTotalSourceTextChars + 1
    else
        documentBlocksSourceTextCharCountFor
            selectionTextCacheMaxTotalSourceParagraphs
            sourceParagraphCount
            blocks

let private selectionTextCacheWithinBudget (cached: SelectionTextCache) =
    cached.Entries.Count > 0
    && cached.Entries.Count <= selectionTextCacheMaxEntries
    && cached.TotalChars <= selectionTextCacheMaxTotalChars
    && cached.SourceParagraphCount <= selectionTextCacheMaxTotalSourceParagraphs
    && cached.SourceTextCharCount <= selectionTextCacheMaxTotalSourceTextChars

let private rememberSelectionTextCache cached =
    if not (selectionTextCacheWithinBudget cached) then
        removeSelectionTextCacheForBlocks cached.SourceBlocks
    else
        let mutable keptCount = 1
        let mutable keptChars = max 1 cached.TotalChars
        let mutable keptBlocks = max 1 cached.SourceParagraphCount
        let mutable keptSourceTextChars = max 1 cached.SourceTextCharCount
        let retained =
            selectionTextCacheWithoutBlocks cached.SourceBlocks
            |> List.choose (fun candidate ->
                let candidateChars = max 1 candidate.TotalChars
                let candidateBlocks = max 1 candidate.SourceParagraphCount
                let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                if selectionTextCacheWithinBudget candidate
                   && keptCount < selectionTextCacheMaxSources
                   && keptChars + candidateChars <= selectionTextCacheMaxTotalChars
                   && keptBlocks + candidateBlocks <= selectionTextCacheMaxTotalSourceParagraphs
                   && keptSourceTextChars + candidateSourceTextChars <= selectionTextCacheMaxTotalSourceTextChars then
                    keptCount <- keptCount + 1
                    keptChars <- keptChars + candidateChars
                    keptBlocks <- keptBlocks + candidateBlocks
                    keptSourceTextChars <- keptSourceTextChars + candidateSourceTextChars
                    Some candidate
                else
                    None)
        cached :: retained
        |> setSelectionTextCache

let private selectionTextKey (selection: Selection) = {
    AnchorBlockId = selection.Anchor.BlockId
    AnchorOffset = selection.Anchor.Offset
    FocusBlockId = selection.Focus.BlockId
    FocusOffset = selection.Focus.Offset
}

[<Struct>]
type private SelectionTextExactCacheKey = {
    SourceBlocks: Block list
    AnchorBlockId: BlockId
    AnchorOffset: int
    FocusBlockId: BlockId
    FocusOffset: int
}

[<Struct>]
type private SelectionTextExactCacheEntry = {
    Key: SelectionTextExactCacheKey
    Text: string
    TextLength: int
}

let private selectionTextExactCacheMaxEntries = selectionTextCacheMaxEntries
let private selectionTextExactCacheMaxTotalChars = selectionTextCacheMaxTotalChars

let private selectionTextExactKeyHash (key: SelectionTextExactCacheKey) =
    let mutable hashCode = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.SourceBlocks)
    hashCode <- (hashCode * 397) ^^^ System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.AnchorBlockId)
    hashCode <- (hashCode * 397) ^^^ key.AnchorOffset
    hashCode <- (hashCode * 397) ^^^ System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.FocusBlockId)
    hashCode <- (hashCode * 397) ^^^ key.FocusOffset
    hashCode <- hashCode ^^^ (hashCode >>> 16)
    hashCode

let private selectionTextExactKeyComparer =
    { new IEqualityComparer<SelectionTextExactCacheKey> with
        member _.Equals(left, right) =
            obj.ReferenceEquals(left.SourceBlocks, right.SourceBlocks)
            && obj.ReferenceEquals(left.AnchorBlockId, right.AnchorBlockId)
            && left.AnchorOffset = right.AnchorOffset
            && obj.ReferenceEquals(left.FocusBlockId, right.FocusBlockId)
            && left.FocusOffset = right.FocusOffset

        member _.GetHashCode(key) = selectionTextExactKeyHash key }

let private selectionTextExactCache =
    Dictionary<SelectionTextExactCacheKey, SelectionTextExactCacheEntry>(
        selectionTextExactCacheMaxEntries,
        selectionTextExactKeyComparer)
let private selectionTextExactOrder = Queue<SelectionTextExactCacheKey>(selectionTextExactCacheMaxEntries)
let mutable private selectionTextExactTotalChars = 0

let private selectionTextExactKey (doc: Document) (selection: Selection) = {
    SourceBlocks = doc.Blocks
    AnchorBlockId = selection.Anchor.BlockId
    AnchorOffset = selection.Anchor.Offset
    FocusBlockId = selection.Focus.BlockId
    FocusOffset = selection.Focus.Offset
}

let private removeSelectionTextExactEntry key =
    let mutable existing = Unchecked.defaultof<SelectionTextExactCacheEntry>
    if selectionTextExactCache.TryGetValue(key, &existing) then
        selectionTextExactCache.Remove key |> ignore
        selectionTextExactTotalChars <- max 0 (selectionTextExactTotalChars - existing.TextLength)

let private rememberSelectionTextExact (doc: Document) (text: string) =
    match doc.Selection with
    | None -> ()
    | Some selection ->
        let key = selectionTextExactKey doc selection
        let textLength = max 1 text.Length
        if textLength > selectionTextExactCacheMaxTotalChars then
            removeSelectionTextExactEntry key
        else
            let mutable existing = Unchecked.defaultof<SelectionTextExactCacheEntry>
            if selectionTextExactCache.TryGetValue(key, &existing) then
                selectionTextExactTotalChars <- max 0 (selectionTextExactTotalChars - existing.TextLength)
            else
                selectionTextExactOrder.Enqueue key
            let entry = {
                Key = key
                Text = text
                TextLength = textLength
            }
            selectionTextExactCache.[key] <- entry
            selectionTextExactTotalChars <- selectionTextExactTotalChars + textLength
            while (selectionTextExactCache.Count > selectionTextExactCacheMaxEntries
                   || selectionTextExactTotalChars > selectionTextExactCacheMaxTotalChars)
                  && selectionTextExactOrder.Count > 0 do
                let oldest = selectionTextExactOrder.Dequeue()
                if not (selectionTextExactKeyComparer.Equals(oldest, key)) then
                    removeSelectionTextExactEntry oldest
            if selectionTextExactCache.Count > selectionTextExactCacheMaxEntries
               || selectionTextExactTotalChars > selectionTextExactCacheMaxTotalChars then
                selectionTextExactCache.Clear()
                selectionTextExactOrder.Clear()
                selectionTextExactOrder.Enqueue key
                selectionTextExactCache.[key] <- entry
                selectionTextExactTotalChars <- textLength

let private tryGetSelectionTextExact (doc: Document) =
    match doc.Selection with
    | None -> ValueNone
    | Some selection ->
        let key = selectionTextExactKey doc selection
        let mutable entry = Unchecked.defaultof<SelectionTextExactCacheEntry>
        if selectionTextExactCache.TryGetValue(key, &entry) then
            ValueSome entry.Text
        else
            ValueNone

let private removeSelectionTextCacheEntry key (cache: SelectionTextCache) =
    let mutable existing = Unchecked.defaultof<SelectionTextCacheEntry>
    if cache.Entries.TryGetValue(key, &existing) then
        clearSelectionTextEntryFastSlots cache.EntryFastSlots
        cache.TotalChars <- max 0 (cache.TotalChars - existing.TextLength)
        cache.Entries.Remove key |> ignore

let private addSelectionTextCacheEntry key (text: string) (cache: SelectionTextCache) =
    let textLength = max 1 text.Length
    let mutable existing = Unchecked.defaultof<SelectionTextCacheEntry>
    if textLength > selectionTextCacheMaxTotalChars then
        removeSelectionTextCacheEntry key cache
    elif cache.Entries.TryGetValue(key, &existing) then
        cache.TotalChars <- max 0 (cache.TotalChars - existing.TextLength)
        let entry = { Text = text; TextLength = textLength }
        cache.Entries.[key] <- entry
        cache.TotalChars <- cache.TotalChars + textLength
        setSelectionTextEntryFastSlot cache key entry
    else
        cache.Order.Enqueue key
        let entry = { Text = text; TextLength = textLength }
        cache.Entries.[key] <- entry
        cache.TotalChars <- cache.TotalChars + textLength
        setSelectionTextEntryFastSlot cache key entry
    if textLength <= selectionTextCacheMaxTotalChars then
        let mutable removedEntry = false
        while (cache.Entries.Count > selectionTextCacheMaxEntries
               || cache.TotalChars > selectionTextCacheMaxTotalChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> key then
                removedEntry <- true
                removeSelectionTextCacheEntry oldest cache
        if cache.Entries.Count > selectionTextCacheMaxEntries
           || cache.TotalChars > selectionTextCacheMaxTotalChars then
            let current = cache.Entries.[key]
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue key
            cache.Entries.[key] <- current
            cache.TotalChars <- current.TextLength
            clearSelectionTextEntryFastSlots cache.EntryFastSlots
            setSelectionTextEntryFastSlot cache key current
        elif removedEntry then
            clearSelectionTextEntryFastSlots cache.EntryFastSlots
            let mutable current = Unchecked.defaultof<SelectionTextCacheEntry>
            if cache.Entries.TryGetValue(key, &current) then
                setSelectionTextEntryFastSlot cache key current

let private cachedSelectionText (doc: Document) =
    match tryGetSelectionTextExact doc with
    | ValueSome text -> ValueSome text
    | ValueNone ->
        match doc.Selection with
        | None -> ValueNone
        | Some selection ->
            match tryGetSelectionTextCache doc.Blocks with
            | ValueSome cached when selectionTextCacheWithinBudget cached ->
                match tryGetSelectionTextEntryFastSlot cached selection with
                | ValueSome entry ->
                    rememberSelectionTextExact doc entry.Text
                    ValueSome entry.Text
                | ValueNone ->
                    let key = selectionTextKey selection
                    let mutable entry = Unchecked.defaultof<SelectionTextCacheEntry>
                    if cached.Entries.TryGetValue(key, &entry) then
                        if selectionTextKeyMatchesSelection selection key then
                            setSelectionTextEntryFastSlot cached key entry
                        rememberSelectionTextExact doc entry.Text
                        ValueSome entry.Text
                    else
                        ValueNone
            | ValueSome cached ->
                rememberSelectionTextCache cached
                ValueNone
            | ValueNone -> ValueNone

let private rememberSelectionText (doc: Document) (text: string) =
    match doc.Selection with
    | None -> ()
    | Some selection ->
        let key = selectionTextKey selection
        match tryGetSelectionTextCache doc.Blocks with
        | ValueSome cached when selectionTextCacheWithinBudget cached ->
            addSelectionTextCacheEntry key text cached
            if selectionTextCacheWithinBudget cached then
                rememberSelectionTextExact doc text
            else
                removeSelectionTextCacheForBlocks cached.SourceBlocks
        | ValueSome cached ->
            removeSelectionTextCacheForBlocks cached.SourceBlocks
        | ValueNone ->
            let entries = Dictionary<SelectionTextCacheKey, SelectionTextCacheEntry>(selectionTextCacheKeyComparer)
            let order = Queue<SelectionTextCacheKey>()
            let sourceParagraphCount = selectionTextSourceParagraphCountFor text.Length doc.Blocks
            let cache = {
                SourceBlocks = doc.Blocks
                SourceParagraphCount = sourceParagraphCount
                SourceTextCharCount =
                    selectionTextSourceTextCharCountFor
                        text.Length
                        sourceParagraphCount
                        doc.Blocks
                Entries = entries
                Order = order
                EntryFastSlots = Array.zeroCreate<SelectionTextFastEntry> 2048
                TotalChars = 0
            }
            addSelectionTextCacheEntry key text cache
            rememberSelectionTextCache cache
            if selectionTextCacheWithinBudget cache then
                rememberSelectionTextExact doc text

type PageBlockAnchor = {
    PageBlockId: BlockId
    PageNumber: int
}

type PageFieldAnchor = {
    PageFieldBlockId: BlockId
    PageFieldStartOffset: int
    PageFieldCode: string
    PageNumber: int
}

type PageBookmarkAnchor = {
    PageBookmarkName: string
    PageBookmarkBlockId: BlockId
    PageBookmarkStartOffset: int
    PageNumber: int
}

type PageMap = {
    PageCount: int
    BlockAnchors: PageBlockAnchor list
    Fields: PageFieldAnchor list
    Bookmarks: PageBookmarkAnchor list
}

/// Create empty document
let empty = {
    Blocks = [
        Paragraph {
            Id = "block-0"
            Buffer = TextBuffer.empty
            Metadata = defaultMetadata
        }
    ]
    Selection = Some {
        Anchor = { BlockId = "block-0"; Offset = 0 }
        Focus = { BlockId = "block-0"; Offset = 0 }
    }
    PageSetup = defaultPageSetup
    Styles = []
    Headers = []
    Footers = []
    HeaderParts = []
    FooterParts = []
}

let private plainTextDocument (text: string) =
    {
        Blocks = [
            Paragraph {
                Id = "block-0"
                Buffer = TextBuffer.ofString text
                Metadata = defaultMetadata
            }
        ]
        Selection = Some {
            Anchor = { BlockId = "block-0"; Offset = text.Length }
            Focus = { BlockId = "block-0"; Offset = text.Length }
        }
        PageSetup = defaultPageSetup
        Styles = []
        Headers = []
        Footers = []
        HeaderParts = []
        FooterParts = []
    }

let private stringSegmentEquals (text: string) offset (segment: string) =
    if offset = 0 && segment.Length = text.Length then
        System.String.Equals(segment, text, System.StringComparison.Ordinal)
    elif offset < 0 || offset + segment.Length > text.Length then
        false
    else
        let mutable index = 0
        let mutable same = true
        while same && index < segment.Length do
            if text.[offset + index] <> segment.[index] then
                same <- false
            index <- index + 1
        same

let private ropeEqualsString (text: string) rope =
    match rope with
    | Rope.Empty -> text.Length = 0
    | Rope.Leaf segment -> System.String.Equals(segment, text, System.StringComparison.Ordinal)
    | Rope.Node _ ->
        if Rope.length rope <> text.Length then
            false
        else
            let rec loop offset current =
                match current with
                | Rope.Empty -> true
                | Rope.Leaf segment -> stringSegmentEquals text offset segment
                | Rope.Node (left, right, _) ->
                    loop offset left && loop (offset + Rope.length left) right
            loop 0 rope

let private isDefaultTextStyle style =
    obj.ReferenceEquals(style, TextBuffer.defaultStyle) || style = TextBuffer.defaultStyle

let private isDefaultBlockMetadata metadata =
    obj.ReferenceEquals(metadata, defaultMetadata) || metadata = defaultMetadata

let private isDefaultPageSetup setup =
    obj.ReferenceEquals(setup, defaultPageSetup) || setup = defaultPageSetup

let private isPlainTextBuffer (text: string) buffer =
    ropeEqualsString text buffer.Rope
        && match buffer.Runs with
           | [] -> text.Length = 0
           | [ Text run ] ->
               run.Start = 0
                   && run.Length = text.Length
                   && isDefaultTextStyle run.Style
           | _ -> false

let setPlainTextContent (text: string) doc =
    let textLength = text.Length
    let selectionAtEnd =
        match doc.Selection with
        | Some selection ->
            selection.Anchor.BlockId = "block-0"
                && selection.Anchor.Offset = textLength
                && selection.Focus.BlockId = "block-0"
                && selection.Focus.Offset = textLength
        | None -> false
    match doc.Blocks with
    | [ Paragraph para ]
        when para.Id = "block-0"
            && isDefaultBlockMetadata para.Metadata
            && isPlainTextBuffer text para.Buffer
            && selectionAtEnd
            && isDefaultPageSetup doc.PageSetup
            && doc.Styles.IsEmpty
            && doc.Headers.IsEmpty
            && doc.Footers.IsEmpty
            && doc.HeaderParts.IsEmpty
            && doc.FooterParts.IsEmpty ->
        doc
    | _ ->
        plainTextDocument text

type private BlockLookupEntry = {
    BlockId: BlockId
    TopBlock: Block
    TopIndex: int
    TopBlockResult: Block option
    TopIndexResult: int option
    TopParagraphResult: ParagraphBlock option
    mutable RecentLookupIdSet: bool
    mutable RecentLookupId: BlockId
}

let private getBlockId = function
    | Paragraph p -> p.Id
    | Heading (_, p) -> p.Id
    | Table t -> t.Id
    | TextBox t -> t.Id

let private blockLookupEntry blockId block index =
    let paragraph =
        match block with
        | Paragraph paragraph
        | Heading (_, paragraph) -> Some paragraph
        | _ -> None
    { BlockId = blockId
      TopBlock = block
      TopIndex = index
      TopBlockResult = Some block
      TopIndexResult = Some index
      TopParagraphResult = paragraph
      RecentLookupIdSet = false
      RecentLookupId = "" }

/// Find block by ID
type private BlockLookupCache = {
    CachedBlocks: Block list
    EntryCount: int
    SourceParagraphCount: int
    SourceTextCharCount: int
    EntriesById: Dictionary<BlockId, BlockLookupEntry>
    EntryFastSlots: BlockLookupEntry array
    TopBlocks: Block array
    mutable LastEntrySet: bool
    mutable LastEntryId: BlockId
    mutable LastEntry: BlockLookupEntry
}

let mutable private blockLookupCache: BlockLookupCache list = []
let private blockLookupBlocksComparer =
    { new IEqualityComparer<Block list> with
        member _.Equals(left, right) = obj.ReferenceEquals(left, right)
        member _.GetHashCode(value) = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value) }
let private blockLookupCacheByBlocks = Dictionary<Block list, BlockLookupCache>(blockLookupBlocksComparer)
let private blockLookupFastSlots = Array.zeroCreate<BlockLookupCache> 256
let private blockLookupCacheMaxEntries = 256
let private blockLookupCacheMaxTotalEntries = 200000
let private blockLookupCacheMaxTotalSourceParagraphs = sourceParagraphCountCacheLimit
let private blockLookupCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars
let private lookupEntryFastSlotMin = 2048
let private lookupEntryFastSlotMax = 8192

let private lookupEntryFastSlotSize entryCount =
    let target = max lookupEntryFastSlotMin (entryCount * 16)
    let mutable size = lookupEntryFastSlotMin
    while size < target && size < lookupEntryFastSlotMax do
        size <- size <<< 1
    size

let private blockLookupEntryFastSlot (slots: BlockLookupEntry array) (blockId: BlockId) =
    blockIdKeyHash blockId &&& (slots.Length - 1)

let private setBlockLookupEntryFastSlot (slots: BlockLookupEntry array) (entry: BlockLookupEntry) =
    slots.[blockLookupEntryFastSlot slots entry.BlockId] <- entry

let private blockLookupEntryMatches (entry: BlockLookupEntry) blockId =
    if obj.ReferenceEquals(entry.BlockId, blockId)
       || (entry.RecentLookupIdSet && obj.ReferenceEquals(entry.RecentLookupId, blockId)) then
        true
    elif blockIdKeyEquals entry.BlockId blockId then
        entry.RecentLookupIdSet <- true
        entry.RecentLookupId <- blockId
        true
    else
        false

let private blockLookupCacheMatches blocks (cached: BlockLookupCache) =
    obj.ReferenceEquals(cached.CachedBlocks, blocks)

let private blockLookupCacheRetainable (cached: BlockLookupCache) =
    cached.EntryCount <= blockLookupCacheMaxTotalEntries
    && cached.SourceParagraphCount <= blockLookupCacheMaxTotalSourceParagraphs
    && cached.SourceTextCharCount <= blockLookupCacheMaxTotalSourceTextChars

let private blockLookupCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<BlockLookupCache>
    if blockLookupCacheByBlocks.TryGetValue(blocks, &cached)
       && blockLookupCacheMatches blocks cached then
        blockLookupCache
        |> List.filter (blockLookupCacheMatches blocks >> not)
    else
        blockLookupCache

let private setBlockLookupCache entries =
    blockLookupCache <- entries
    blockLookupCacheByBlocks.Clear()
    clearArray blockLookupFastSlots
    for entry in entries do
        blockLookupCacheByBlocks.[entry.CachedBlocks] <- entry
        let slot = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(entry.CachedBlocks) &&& (blockLookupFastSlots.Length - 1)
        blockLookupFastSlots.[slot] <- entry

let private tryGetBlockLookup blocks =
    let slot = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(blocks) &&& (blockLookupFastSlots.Length - 1)
    let fast = blockLookupFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && obj.ReferenceEquals(fast.CachedBlocks, blocks)
       && blockLookupCacheRetainable fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<BlockLookupCache>
        if blockLookupCacheByBlocks.TryGetValue(blocks, &cached)
           && blockLookupCacheRetainable cached then
            blockLookupFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private tryFindBlockLookupEntry (lookup: BlockLookupCache) blockId =
    if lookup.LastEntrySet && obj.ReferenceEquals(lookup.LastEntryId, blockId) then
        ValueSome lookup.LastEntry
    else
        let fast = lookup.EntryFastSlots.[blockLookupEntryFastSlot lookup.EntryFastSlots blockId]
        if not (obj.ReferenceEquals(fast, null))
           && blockLookupEntryMatches fast blockId then
            lookup.LastEntrySet <- true
            lookup.LastEntryId <- blockId
            lookup.LastEntry <- fast
            ValueSome fast
        else
            let mutable entry = Unchecked.defaultof<BlockLookupEntry>
            if lookup.EntriesById.TryGetValue(blockId, &entry) then
                entry.RecentLookupIdSet <- true
                entry.RecentLookupId <- blockId
                setBlockLookupEntryFastSlot lookup.EntryFastSlots entry
                lookup.LastEntrySet <- true
                lookup.LastEntryId <- blockId
                lookup.LastEntry <- entry
                ValueSome entry
            else
                ValueNone

let private blockLookupSourceParagraphCountFor entryCount blocks =
    if entryCount > blockLookupCacheMaxTotalEntries then
        blockLookupCacheMaxTotalSourceParagraphs + 1
    else
        countSourceParagraphs blocks

let private blockLookupSourceTextCharCountFor entryCount sourceParagraphCount blocks =
    if entryCount > blockLookupCacheMaxTotalEntries
       || sourceParagraphCount > blockLookupCacheMaxTotalSourceParagraphs then
        blockLookupCacheMaxTotalSourceTextChars + 1
    else
        documentBlocksSourceTextCharCountFor
            blockLookupCacheMaxTotalSourceParagraphs
            sourceParagraphCount
            blocks

let private rememberBlockLookup cached =
    if not (blockLookupCacheRetainable cached) then
        let entries = blockLookupCacheWithoutBlocks cached.CachedBlocks
        if not (obj.ReferenceEquals(entries, blockLookupCache)) then
            setBlockLookupCache entries
    else
        let trim entries =
            let rec loop keptCount keptEntries keptSourceParagraphs keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateEntries = max 1 candidate.EntryCount
                    let candidateSourceParagraphs = max 1 candidate.SourceParagraphCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if blockLookupCacheRetainable candidate
                       && keptCount < blockLookupCacheMaxEntries
                       && keptEntries + candidateEntries <= blockLookupCacheMaxTotalEntries
                       && keptSourceParagraphs + candidateSourceParagraphs <= blockLookupCacheMaxTotalSourceParagraphs
                       && keptSourceTextChars + candidateSourceTextChars <= blockLookupCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptEntries + candidateEntries)
                            (keptSourceParagraphs + candidateSourceParagraphs)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptEntries keptSourceParagraphs keptSourceTextChars acc tail
            loop 0 0 0 0 [] entries
        let entries =
            (cached
             :: blockLookupCacheWithoutBlocks cached.CachedBlocks)
            |> trim
        setBlockLookupCache entries

let private buildBlockLookup blocks =
    let entriesById = Dictionary<BlockId, BlockLookupEntry>(blockIdLookupComparer)
    let topBlocks = ResizeArray<Block>()
    let mutable index = 0
    for block in blocks do
        topBlocks.Add block
        let blockId = getBlockId block
        if not (entriesById.ContainsKey blockId) then
            let entry = blockLookupEntry blockId block index
            entriesById.[blockId] <- entry
        index <- index + 1
    let entryFastSlots = Array.zeroCreate<BlockLookupEntry> (lookupEntryFastSlotSize entriesById.Count)
    for entry in entriesById.Values do
        setBlockLookupEntryFastSlot entryFastSlots entry
    let sourceParagraphCount = blockLookupSourceParagraphCountFor entriesById.Count blocks
    let cached = {
        CachedBlocks = blocks
        EntryCount = entriesById.Count
        SourceParagraphCount = sourceParagraphCount
        SourceTextCharCount = blockLookupSourceTextCharCountFor entriesById.Count sourceParagraphCount blocks
        EntriesById = entriesById
        EntryFastSlots = entryFastSlots
        TopBlocks = topBlocks.ToArray()
        LastEntrySet = false
        LastEntryId = ""
        LastEntry = Unchecked.defaultof<BlockLookupEntry>
    }
    rememberBlockLookup cached
    cached

let private blockLookup blocks =
    match tryGetBlockLookup blocks with
    | ValueSome cached -> cached
    | ValueNone -> buildBlockLookup blocks

let findBlock doc blockId =
    let lookup = blockLookup doc.Blocks
    match tryFindBlockLookupEntry lookup blockId with
    | ValueSome entry -> entry.TopBlockResult
    | ValueNone -> None

let private scanTopBlockUntil blocks blockId =
    let rec loop remaining =
        match remaining with
        | block :: _ when blockIdKeyEquals (getBlockId block) blockId -> ValueSome block
        | _ :: tail -> loop tail
        | [] -> ValueNone
    loop blocks

let private scanTopBlockWithIndexUntil blocks blockId =
    let rec loop index remaining =
        match remaining with
        | block :: _ when blockIdKeyEquals (getBlockId block) blockId -> ValueSome (struct (block, index))
        | _ :: tail -> loop (index + 1) tail
        | [] -> ValueNone
    loop 0 blocks

let private scanAdjacentTopBlock blocks blockId offset =
    if offset = 1 then
        let rec loop remaining =
            match remaining with
            | block :: next :: _ when blockIdKeyEquals (getBlockId block) blockId -> Some next
            | block :: [] when blockIdKeyEquals (getBlockId block) blockId -> None
            | _ :: tail -> loop tail
            | [] -> None
        loop blocks
    elif offset = -1 then
        let rec loop hasPrevious previous remaining =
            match remaining with
            | block :: _ when blockIdKeyEquals (getBlockId block) blockId ->
                if hasPrevious then Some previous else None
            | block :: tail -> loop true block tail
            | [] -> None
        loop false Unchecked.defaultof<Block> blocks
    else
        None

let private tryFindTopBlockUntil doc blockId =
    match tryGetBlockLookup doc.Blocks with
    | ValueSome lookup ->
        match tryFindBlockLookupEntry lookup blockId with
        | ValueSome entry ->
            match entry.TopBlockResult with
            | Some block -> ValueSome block
            | None -> ValueNone
        | ValueNone -> ValueNone
    | ValueNone -> scanTopBlockUntil doc.Blocks blockId

let private tryFindTopBlockWithIndexUntil doc blockId =
    match tryGetBlockLookup doc.Blocks with
    | ValueSome lookup ->
        match tryFindBlockLookupEntry lookup blockId with
        | ValueSome entry -> ValueSome (struct (entry.TopBlock, entry.TopIndex))
        | ValueNone -> ValueNone
    | ValueNone -> scanTopBlockWithIndexUntil doc.Blocks blockId

let private tryFindAdjacentTopBlock doc blockId offset =
    match tryGetBlockLookup doc.Blocks with
    | ValueSome lookup ->
        match tryFindBlockLookupEntry lookup blockId with
        | ValueSome entry ->
            let index = entry.TopIndex + offset
            if index >= 0 && index < lookup.TopBlocks.Length then
                Some lookup.TopBlocks.[index]
            else
                None
        | ValueNone -> None
    | ValueNone -> scanAdjacentTopBlock doc.Blocks blockId offset

/// Try to get the editable paragraph payload from a block.
let tryGetParagraph = function
    | Paragraph p -> Some p
    | Heading (_, p) -> Some p
    | Table _
    | TextBox _ -> None

/// Get paragraph from block. Callers that can encounter tables should use tryGetParagraph.
let getParagraph = function
    | Paragraph p -> p
    | Heading (_, p) -> p
    | Table _
    | TextBox _ -> invalidOp "This block does not have a paragraph payload"

let private appendBlock block blocks =
    let values = ResizeArray<Block>()
    for existing in blocks do
        values.Add existing
    values.Add block
    resizeArrayToList values

let private topLevelBlockId = function
    | Paragraph p
    | Heading (_, p) -> p.Id
    | Table t -> t.Id
    | TextBox t -> t.Id

let private insertBlockAfterTopLevelId blockId block blocks =
    let prefix = ResizeArray<Block>()
    let rec loop remaining =
        match remaining with
        | [] -> prependResizeArray prefix [ block ]
        | existing :: tail when topLevelBlockId existing = blockId ->
            prependResizeArray prefix (existing :: block :: tail)
        | existing :: tail ->
            prefix.Add existing
            loop tail
    loop blocks

let private replaceBlockAndInsertAfterIndex index replacement inserted blocks =
    let prefix = ResizeArray<Block>()
    let rec loop currentIndex remaining =
        match remaining with
        | [] -> prependResizeArray prefix []
        | _ :: tail when currentIndex = index ->
            prependResizeArray prefix (replacement :: inserted :: tail)
        | block :: tail ->
            prefix.Add block
            loop (currentIndex + 1) tail
    loop 0 blocks

let private replaceTopLevelBlockAtIndex index replacement blocks =
    let prefix = ResizeArray<Block>(max 0 index)
    let rec loop currentIndex remaining =
        match remaining with
        | [] -> prependResizeArray prefix []
        | _ :: tail when currentIndex = index ->
            prependResizeArray prefix (replacement :: tail)
        | block :: tail ->
            prefix.Add block
            loop (currentIndex + 1) tail
    loop 0 blocks

let private paragraphReplacementMatches (current: ParagraphBlock) (replacement: ParagraphBlock) =
    obj.ReferenceEquals(current, replacement)
    || (current.Id = replacement.Id
        && obj.ReferenceEquals(current.Buffer, replacement.Buffer)
        && current.Metadata = replacement.Metadata)

type private TopLevelParagraphEditCacheEntry = {
    SourceBlocks: Block list
    TargetParagraphId: BlockId
    TargetIndex: int
    PrefixBlocks: Block array
    TargetBlock: Block
    TailAfterTarget: Block list
}

let mutable private topLevelParagraphEditLastSet = false
let mutable private topLevelParagraphEditLast = Unchecked.defaultof<TopLevelParagraphEditCacheEntry>

let private prependBlockArray (prefix: Block array) tail =
    let mutable result = tail
    for i = prefix.Length - 1 downto 0 do
        result <- prefix.[i] :: result
    result

let private tryGetTopLevelParagraphEditEntry blocks paragraphId =
    if topLevelParagraphEditLastSet
       && obj.ReferenceEquals(topLevelParagraphEditLast.SourceBlocks, blocks)
       && blockIdKeyEquals topLevelParagraphEditLast.TargetParagraphId paragraphId then
        ValueSome topLevelParagraphEditLast
    else
        ValueNone

let private rememberTopLevelParagraphEditEntry blocks paragraphId index prefixBlocks targetBlock tailAfterTarget =
    topLevelParagraphEditLastSet <- true
    topLevelParagraphEditLast <-
        { SourceBlocks = blocks
          TargetParagraphId = paragraphId
          TargetIndex = index
          PrefixBlocks = prefixBlocks
          TargetBlock = targetBlock
          TailAfterTarget = tailAfterTarget }

let private rememberTopLevelParagraphEditFromReplacement entry nextBlocks replacement =
    rememberTopLevelParagraphEditEntry
        nextBlocks
        entry.TargetParagraphId
        entry.TargetIndex
        entry.PrefixBlocks
        replacement
        entry.TailAfterTarget

let private replaceTopLevelBlockFromCachedEntry entry replacement =
    prependBlockArray entry.PrefixBlocks (replacement :: entry.TailAfterTarget)

let private replaceTopLevelBlockAtIndexAndRemember index replacement paragraphId blocks =
    let prefix = ResizeArray<Block>(max 0 index)
    let mutable remaining = blocks
    let mutable currentIndex = 0
    let mutable result = blocks
    let mutable scanning = true
    while scanning do
        match remaining with
        | [] ->
            result <- prependResizeArray prefix []
            scanning <- false
        | _ :: tail when currentIndex = index ->
            let prefixBlocks = prefix.ToArray()
            result <- prependResizeArray prefix (replacement :: tail)
            rememberTopLevelParagraphEditEntry result paragraphId index prefixBlocks replacement tail
            scanning <- false
        | block :: tail ->
            prefix.Add block
            remaining <- tail
            currentIndex <- currentIndex + 1
    result

let private tryReplaceTopLevelParagraphFromCachedEntry
    (doc: Document)
    paragraphId
    (entry: TopLevelParagraphEditCacheEntry)
    newPara
    : Document voption =
    match entry.TargetBlock with
    | Paragraph para when blockIdKeyEquals para.Id paragraphId ->
        if paragraphReplacementMatches para newPara then
            ValueSome doc
        else
            let replacement = Paragraph newPara
            let nextBlocks = replaceTopLevelBlockFromCachedEntry entry replacement
            rememberTopLevelParagraphEditFromReplacement entry nextBlocks replacement
            ValueSome { doc with Blocks = nextBlocks }
    | Heading (level, para) when blockIdKeyEquals para.Id paragraphId ->
        if paragraphReplacementMatches para newPara then
            ValueSome doc
        else
            let replacement = Heading (level, newPara)
            let nextBlocks = replaceTopLevelBlockFromCachedEntry entry replacement
            rememberTopLevelParagraphEditFromReplacement entry nextBlocks replacement
            ValueSome { doc with Blocks = nextBlocks }
    | _ ->
        ValueNone

let private tryUpdateTopLevelParagraphFromCachedEntry
    (doc: Document)
    paragraphId
    (entry: TopLevelParagraphEditCacheEntry)
    updateParagraph
    : Document voption =
    match entry.TargetBlock with
    | Paragraph para when blockIdKeyEquals para.Id paragraphId ->
        let nextPara = updateParagraph para
        if paragraphReplacementMatches para nextPara then
            ValueSome doc
        else
            let replacement = Paragraph nextPara
            let nextBlocks = replaceTopLevelBlockFromCachedEntry entry replacement
            rememberTopLevelParagraphEditFromReplacement entry nextBlocks replacement
            ValueSome { doc with Blocks = nextBlocks }
    | Heading (level, para) when blockIdKeyEquals para.Id paragraphId ->
        let nextPara = updateParagraph para
        if paragraphReplacementMatches para nextPara then
            ValueSome doc
        else
            let replacement = Heading (level, nextPara)
            let nextBlocks = replaceTopLevelBlockFromCachedEntry entry replacement
            rememberTopLevelParagraphEditFromReplacement entry nextBlocks replacement
            ValueSome { doc with Blocks = nextBlocks }
    | _ ->
        ValueNone

let private withParagraphBuffer block para =
    match block with
    | Paragraph _ -> Some (Paragraph para)
    | Heading (level, _) -> Some (Heading (level, para))
    | Table _
    | TextBox _ -> None

let private blockReplacementMatches current replacement =
    obj.ReferenceEquals(current, replacement)
    || match current, replacement with
       | Paragraph currentPara, Paragraph replacementPara
       | Heading (_, currentPara), Heading (_, replacementPara) ->
           paragraphReplacementMatches currentPara replacementPara
       | Table currentTable, Table replacementTable ->
           obj.ReferenceEquals(currentTable, replacementTable)
       | TextBox currentTextBox, TextBox replacementTextBox ->
           obj.ReferenceEquals(currentTextBox, replacementTextBox)
       | _ ->
           false

/// Replace block in document
let replaceBlock (doc: Document) blockId newBlock =
    let rec revAppend reversed tail =
        match reversed with
        | [] -> tail
        | head :: rest -> revAppend rest (head :: tail)
    let rec containsBlockId remaining =
        match remaining with
        | [] -> false
        | block :: tail ->
            getBlockId block = blockId || containsBlockId tail
    let rec replaceMatching remaining =
        match remaining with
        | [] -> []
        | block :: tail ->
            let block = if getBlockId block = blockId then newBlock else block
            block :: replaceMatching tail
    let rec loop prefix remaining =
        match remaining with
        | [] ->
            doc
        | block :: tail ->
            if getBlockId block = blockId then
                if blockReplacementMatches block newBlock then
                    doc
                elif containsBlockId tail then
                    { doc with Blocks = revAppend prefix (newBlock :: replaceMatching tail) }
                else
                    { doc with Blocks = revAppend prefix (newBlock :: tail) }
            else
                loop (block :: prefix) tail
    loop [] doc.Blocks

let private replaceFirstTopLevelBlock (doc: Document) blockId newBlock =
    let rec revAppend reversed tail =
        match reversed with
        | [] -> tail
        | head :: rest -> revAppend rest (head :: tail)
    let rec loop prefix remaining =
        match remaining with
        | [] ->
            doc
        | block :: tail ->
            if getBlockId block = blockId then
                if blockReplacementMatches block newBlock then
                    doc
                else
                    { doc with Blocks = revAppend prefix (newBlock :: tail) }
            else
                loop (block :: prefix) tail
    loop [] doc.Blocks

let private paragraphNeedsReplacement (paragraphId: BlockId) (newPara: ParagraphBlock) (paragraph: ParagraphBlock) =
    paragraph.Id = paragraphId
    && not (paragraphReplacementMatches paragraph newPara)

let private contentNeedsParagraphReplacement (paragraphId: BlockId) (newPara: ParagraphBlock) (paragraphs: ParagraphBlock list) =
    let mutable found = false
    let mutable remaining = paragraphs
    while not found && not remaining.IsEmpty do
        match remaining with
        | paragraph :: tail ->
            found <- paragraphNeedsReplacement paragraphId newPara paragraph
            remaining <- tail
        | [] -> ()
    found

let private cellNeedsParagraphReplacement (paragraphId: BlockId) (newPara: ParagraphBlock) (cell: TableCell) =
    contentNeedsParagraphReplacement paragraphId newPara cell.Content

let private cellsNeedParagraphReplacement (paragraphId: BlockId) (newPara: ParagraphBlock) (cells: TableCell list) =
    let mutable found = false
    let mutable remaining = cells
    while not found && not remaining.IsEmpty do
        match remaining with
        | cell :: tail ->
            found <- cellNeedsParagraphReplacement paragraphId newPara cell
            remaining <- tail
        | [] -> ()
    found

let private rowNeedsParagraphReplacement (paragraphId: BlockId) (newPara: ParagraphBlock) (row: TableRow) =
    cellsNeedParagraphReplacement paragraphId newPara row.Cells

let private rowsNeedParagraphReplacement (paragraphId: BlockId) (newPara: ParagraphBlock) (rows: TableRow list) =
    let mutable found = false
    let mutable remaining = rows
    while not found && not remaining.IsEmpty do
        match remaining with
        | row :: tail ->
            found <- rowNeedsParagraphReplacement paragraphId newPara row
            remaining <- tail
        | [] -> ()
    found

let private tableNeedsParagraphReplacement (paragraphId: BlockId) (newPara: ParagraphBlock) (table: TableBlock) =
    rowsNeedParagraphReplacement paragraphId newPara table.Rows

let private replaceParagraphInContent (paragraphId: BlockId) (newPara: ParagraphBlock) (paragraphs: ParagraphBlock list) =
    if not (contentNeedsParagraphReplacement paragraphId newPara paragraphs) then
        paragraphs, false
    else
        let rec revAppend (reversed: ParagraphBlock list) (tail: ParagraphBlock list) =
            match reversed with
            | [] -> tail
            | head :: rest -> revAppend rest (head :: tail)
        let rec loop (prefix: ParagraphBlock list) changed (remaining: ParagraphBlock list) =
            match remaining with
            | [] ->
                if changed then Some (revAppend prefix []) else None
            | paragraph :: tail when paragraph.Id = paragraphId ->
                if paragraphReplacementMatches paragraph newPara then
                    loop (paragraph :: prefix) changed tail
                elif contentNeedsParagraphReplacement paragraphId newPara tail then
                    loop (newPara :: prefix) true tail
                else
                    Some (revAppend prefix (newPara :: tail))
            | paragraph :: tail ->
                loop (paragraph :: prefix) changed tail
        match loop [] false paragraphs with
        | Some paragraphs -> paragraphs, true
        | None -> paragraphs, false

let private replaceParagraphInCell paragraphId newPara (cell: TableCell) =
    let content, changed = replaceParagraphInContent paragraphId newPara cell.Content
    if changed then { cell with Content = content }, true else cell, false

let private replaceParagraphInRow paragraphId newPara (row: TableRow) =
    if not (rowNeedsParagraphReplacement paragraphId newPara row) then
        row, false
    else
        let rec revAppend reversed tail =
            match reversed with
            | [] -> tail
            | head :: rest -> revAppend rest (head :: tail)
        let rec loop prefix changed cells =
            match cells with
            | [] ->
                if changed then Some (revAppend prefix []) else None
            | cell :: tail ->
                let cell, cellChanged = replaceParagraphInCell paragraphId newPara cell
                if cellChanged && not (cellsNeedParagraphReplacement paragraphId newPara tail) then
                    Some (revAppend prefix (cell :: tail))
                else
                    loop (cell :: prefix) (changed || cellChanged) tail
        match loop [] false row.Cells with
        | Some cells -> { row with Cells = cells }, true
        | None -> row, false

let private replaceParagraphInBlock paragraphId newPara block =
    match block with
    | Paragraph p when p.Id = paragraphId ->
        if paragraphReplacementMatches p newPara then block, false else Paragraph newPara, true
    | Heading (level, p) when p.Id = paragraphId ->
        if paragraphReplacementMatches p newPara then block, false else Heading (level, newPara), true
    | Table table ->
        if not (tableNeedsParagraphReplacement paragraphId newPara table) then
            block, false
        else
            let rec revAppend reversed tail =
                match reversed with
                | [] -> tail
                | head :: rest -> revAppend rest (head :: tail)
            let rec loop prefix changed rows =
                match rows with
                | [] ->
                    if changed then Some (revAppend prefix []) else None
                | row :: tail ->
                    let row, rowChanged = replaceParagraphInRow paragraphId newPara row
                    if rowChanged && not (rowsNeedParagraphReplacement paragraphId newPara tail) then
                        Some (revAppend prefix (row :: tail))
                    else
                        loop (row :: prefix) (changed || rowChanged) tail
            match loop [] false table.Rows with
            | Some rows -> Table { table with Rows = rows }, true
            | None -> block, false
    | TextBox textBox ->
        if not (contentNeedsParagraphReplacement paragraphId newPara textBox.Content) then
            block, false
        else
            let content, changed = replaceParagraphInContent paragraphId newPara textBox.Content
            if changed then TextBox { textBox with Content = content }, true else block, false
    | _ -> block, false

let private updateParagraphInContent paragraphId (update: ParagraphBlock -> ParagraphBlock) (paragraphs: ParagraphBlock list) =
    let rec revAppend reversed tail =
        match reversed with
        | [] -> tail
        | head :: rest -> revAppend rest (head :: tail)
    let rec loop (prefix: ParagraphBlock list) (remaining: ParagraphBlock list) =
        match remaining with
        | [] -> None
        | paragraph :: tail when paragraph.Id = paragraphId ->
            let updated = update paragraph
            if obj.ReferenceEquals(updated, paragraph) || updated = paragraph then
                None
            else
                Some (revAppend prefix (updated :: tail))
        | paragraph :: tail ->
            loop (paragraph :: prefix) tail
    match loop [] paragraphs with
    | Some paragraphs -> paragraphs, true
    | None -> paragraphs, false

let private updateParagraphInCell paragraphId update (cell: TableCell) =
    let content, changed = updateParagraphInContent paragraphId update cell.Content
    if changed then { cell with Content = content }, true else cell, false

let private updateParagraphInRow paragraphId update (row: TableRow) =
    let rec revAppend reversed tail =
        match reversed with
        | [] -> tail
        | head :: rest -> revAppend rest (head :: tail)
    let rec loop prefix remaining =
        match remaining with
        | [] -> row, false
        | cell :: tail ->
            let cell, cellChanged = updateParagraphInCell paragraphId update cell
            if cellChanged then
                { row with Cells = revAppend prefix (cell :: tail) }, true
            else
                loop (cell :: prefix) tail
    loop [] row.Cells

let private updateParagraphInBlock paragraphId update block =
    match block with
    | Paragraph p when p.Id = paragraphId ->
        let updated = update p
        if obj.ReferenceEquals(updated, p) || updated = p then block, false else Paragraph updated, true
    | Heading (level, p) when p.Id = paragraphId ->
        let updated = update p
        if obj.ReferenceEquals(updated, p) || updated = p then block, false else Heading (level, updated), true
    | Table table ->
        let rec revAppend reversed tail =
            match reversed with
            | [] -> tail
            | head :: rest -> revAppend rest (head :: tail)
        let rec loop prefix remaining =
            match remaining with
            | [] -> block, false
            | row :: tail ->
                let row, rowChanged = updateParagraphInRow paragraphId update row
                if rowChanged then
                    Table { table with Rows = revAppend prefix (row :: tail) }, true
                else
                    loop (row :: prefix) tail
        loop [] table.Rows
    | TextBox textBox ->
        let content, changed = updateParagraphInContent paragraphId update textBox.Content
        if changed then TextBox { textBox with Content = content }, true else block, false
    | _ -> block, false

let private contentNeedsParagraphReplacements (replacements: Dictionary<BlockId, ParagraphBlock>) (paragraphs: ParagraphBlock list) =
    let mutable found = false
    let mutable remaining = paragraphs
    let mutable replacement = Unchecked.defaultof<ParagraphBlock>
    while not found && not remaining.IsEmpty do
        match remaining with
        | paragraph :: tail ->
            found <-
                replacements.TryGetValue(paragraph.Id, &replacement)
                && not (paragraphReplacementMatches paragraph replacement)
            remaining <- tail
        | [] -> ()
    found

let private cellNeedsParagraphReplacements replacements (cell: TableCell) =
    contentNeedsParagraphReplacements replacements cell.Content

let private rowNeedsParagraphReplacements replacements (row: TableRow) =
    let mutable found = false
    let mutable remaining = row.Cells
    while not found && not remaining.IsEmpty do
        match remaining with
        | cell :: tail ->
            found <- cellNeedsParagraphReplacements replacements cell
            remaining <- tail
        | [] -> ()
    found

let private tableNeedsParagraphReplacements replacements (table: TableBlock) =
    let mutable found = false
    let mutable remaining = table.Rows
    while not found && not remaining.IsEmpty do
        match remaining with
        | row :: tail ->
            found <- rowNeedsParagraphReplacements replacements row
            remaining <- tail
        | [] -> ()
    found

let private blockNeedsParagraphReplacements (replacements: Dictionary<BlockId, ParagraphBlock>) block =
    match block with
    | Paragraph paragraph
    | Heading (_, paragraph) ->
        let mutable replacement = Unchecked.defaultof<ParagraphBlock>
        replacements.TryGetValue(paragraph.Id, &replacement)
        && not (paragraphReplacementMatches paragraph replacement)
    | Table table ->
        tableNeedsParagraphReplacements replacements table
    | TextBox textBox ->
        contentNeedsParagraphReplacements replacements textBox.Content

let private blocksNeedParagraphReplacements (replacements: Dictionary<BlockId, ParagraphBlock>) (blocks: Block list) =
    let mutable found = false
    let mutable remaining = blocks
    while not found && not remaining.IsEmpty do
        match remaining with
        | block :: tail ->
            found <- blockNeedsParagraphReplacements replacements block
            remaining <- tail
        | [] -> ()
    found

let private replaceParagraphsInContent (replacements: Dictionary<BlockId, ParagraphBlock>) (paragraphs: ParagraphBlock list) =
    if not (contentNeedsParagraphReplacements replacements paragraphs) then
        paragraphs, false
    else
        let values = ResizeArray<ParagraphBlock>()
        let mutable changed = false
        for paragraph in paragraphs do
            let mutable replacement = Unchecked.defaultof<ParagraphBlock>
            if replacements.TryGetValue(paragraph.Id, &replacement) then
                if paragraphReplacementMatches paragraph replacement then
                    values.Add paragraph
                else
                    values.Add replacement
                    changed <- true
            else
                values.Add paragraph
        if changed then resizeArrayToList values, true else paragraphs, false

let private replaceParagraphsInCell replacements (cell: TableCell) =
    let content, changed = replaceParagraphsInContent replacements cell.Content
    if changed then { cell with Content = content }, true else cell, false

let private replaceParagraphsInRow replacements (row: TableRow) =
    if not (rowNeedsParagraphReplacements replacements row) then
        row, false
    else
        let cells = ResizeArray<TableCell>()
        let mutable changed = false
        for cell in row.Cells do
            let cell, cellChanged = replaceParagraphsInCell replacements cell
            cells.Add cell
            changed <- changed || cellChanged
        if changed then { row with Cells = resizeArrayToList cells }, true else row, false

let private replaceParagraphsInBlock (replacements: Dictionary<BlockId, ParagraphBlock>) block =
    match block with
    | Paragraph p ->
        let mutable replacement = Unchecked.defaultof<ParagraphBlock>
        if replacements.TryGetValue(p.Id, &replacement)
           && not (paragraphReplacementMatches p replacement) then
            Paragraph replacement, true
        else
            block, false
    | Heading (level, p) ->
        let mutable replacement = Unchecked.defaultof<ParagraphBlock>
        if replacements.TryGetValue(p.Id, &replacement)
           && not (paragraphReplacementMatches p replacement) then
            Heading (level, replacement), true
        else
            block, false
    | Table table ->
        if not (tableNeedsParagraphReplacements replacements table) then
            block, false
        else
            let rows = ResizeArray<TableRow>()
            let mutable changed = false
            for row in table.Rows do
                let row, rowChanged = replaceParagraphsInRow replacements row
                rows.Add row
                changed <- changed || rowChanged
            if changed then Table { table with Rows = resizeArrayToList rows }, true else block, false
    | TextBox textBox ->
        if not (contentNeedsParagraphReplacements replacements textBox.Content) then
            block, false
        else
            let content, changed = replaceParagraphsInContent replacements textBox.Content
            if changed then TextBox { textBox with Content = content }, true else block, false

let private replaceParagraphs (doc: Document) (replacements: Dictionary<BlockId, ParagraphBlock>) =
    if replacements.Count = 0 then
        doc
    elif replacements.Count = 1 then
        let prefix = ResizeArray<Block>()
        let mutable remaining = doc.Blocks
        let mutable result: Document voption = ValueNone
        while result.IsNone && not remaining.IsEmpty do
            match remaining with
            | block :: tail ->
                let block, blockChanged = replaceParagraphsInBlock replacements block
                if blockChanged then
                    result <- ValueSome { doc with Blocks = prependResizeArray prefix (block :: tail) }
                else
                    prefix.Add block
                    remaining <- tail
            | [] -> ()
        match result with
        | ValueSome doc -> doc
        | ValueNone -> doc
    else
        let blocks = ResizeArray<Block>()
        let mutable pendingUnchangedBlocks = Unchecked.defaultof<ResizeArray<Block>>
        let mutable changed = false
        let mutable tailAfterLastChange: Block list = []

        let flushPendingUnchangedBlocks () =
            if not (obj.ReferenceEquals(pendingUnchangedBlocks, null)) && pendingUnchangedBlocks.Count > 0 then
                blocks.AddRange pendingUnchangedBlocks
                pendingUnchangedBlocks.Clear()

        let addChangedBlock block tail =
            if changed then
                flushPendingUnchangedBlocks ()
            changed <- true
            blocks.Add block
            tailAfterLastChange <- tail

        let addUnchangedBlock block =
            if changed then
                if obj.ReferenceEquals(pendingUnchangedBlocks, null) then
                    pendingUnchangedBlocks <- ResizeArray<Block>()
                pendingUnchangedBlocks.Add block
            else
                blocks.Add block

        let mutable remaining = doc.Blocks
        while not remaining.IsEmpty do
            match remaining with
            | block :: tail ->
                let block, blockChanged = replaceParagraphsInBlock replacements block
                if blockChanged then
                    addChangedBlock block tail
                else
                    addUnchangedBlock block
                remaining <- tail
            | [] -> ()

        if changed then
            { doc with Blocks = prependResizeArray blocks tailAfterLastChange }
        else
            doc

type private ParagraphLookupCache = {
    CachedBlocks: Block list
    EntryCount: int
    SourceParagraphCount: int
    SourceTextCharCount: int
    Paragraphs: Dictionary<BlockId, ParagraphLookupEntry>
    ParagraphFastSlots: ParagraphLookupEntry array
}

and private ParagraphLookupEntry = {
    Paragraph: ParagraphBlock
    Result: ParagraphBlock option
    mutable RecentLookupIdSet: bool
    mutable RecentLookupId: BlockId
}

let mutable private paragraphLookupCache: ParagraphLookupCache list = []
let private paragraphLookupCacheByBlocks = Dictionary<Block list, ParagraphLookupCache>(blockLookupBlocksComparer)
let private paragraphLookupFastSlots = Array.zeroCreate<ParagraphLookupCache> 256
let private paragraphLookupCacheMaxEntries = 256
let private paragraphLookupCacheMaxTotalEntries = 200000
let private paragraphLookupCacheMaxTotalSourceParagraphs = sourceParagraphCountCacheLimit
let private paragraphLookupCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars

let private paragraphLookupEntryFastSlot (slots: ParagraphLookupEntry array) (paragraphId: BlockId) =
    blockIdKeyHash paragraphId &&& (slots.Length - 1)

let private setParagraphLookupEntryFastSlot (slots: ParagraphLookupEntry array) (entry: ParagraphLookupEntry) =
    slots.[paragraphLookupEntryFastSlot slots entry.Paragraph.Id] <- entry

let private paragraphLookupEntryMatches (entry: ParagraphLookupEntry) paragraphId =
    if obj.ReferenceEquals(entry.Paragraph.Id, paragraphId)
       || (entry.RecentLookupIdSet && obj.ReferenceEquals(entry.RecentLookupId, paragraphId)) then
        true
    elif blockIdKeyEquals entry.Paragraph.Id paragraphId then
        entry.RecentLookupIdSet <- true
        entry.RecentLookupId <- paragraphId
        true
    else
        false

let private paragraphLookupCacheMatches blocks (cached: ParagraphLookupCache) =
    obj.ReferenceEquals(cached.CachedBlocks, blocks)

let private paragraphLookupCacheRetainable (cached: ParagraphLookupCache) =
    cached.EntryCount <= paragraphLookupCacheMaxTotalEntries
    && cached.SourceParagraphCount <= paragraphLookupCacheMaxTotalSourceParagraphs
    && cached.SourceTextCharCount <= paragraphLookupCacheMaxTotalSourceTextChars

let private paragraphLookupCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<ParagraphLookupCache>
    if paragraphLookupCacheByBlocks.TryGetValue(blocks, &cached)
       && paragraphLookupCacheMatches blocks cached then
        paragraphLookupCache
        |> List.filter (paragraphLookupCacheMatches blocks >> not)
    else
        paragraphLookupCache

let private setParagraphLookupCache entries =
    paragraphLookupCache <- entries
    paragraphLookupCacheByBlocks.Clear()
    clearArray paragraphLookupFastSlots
    for entry in entries do
        paragraphLookupCacheByBlocks.[entry.CachedBlocks] <- entry
        let slot = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(entry.CachedBlocks) &&& (paragraphLookupFastSlots.Length - 1)
        paragraphLookupFastSlots.[slot] <- entry

let private tryGetParagraphLookup blocks =
    let slot = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(blocks) &&& (paragraphLookupFastSlots.Length - 1)
    let fast = paragraphLookupFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && obj.ReferenceEquals(fast.CachedBlocks, blocks)
       && paragraphLookupCacheRetainable fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<ParagraphLookupCache>
        if paragraphLookupCacheByBlocks.TryGetValue(blocks, &cached)
           && paragraphLookupCacheRetainable cached then
            paragraphLookupFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private paragraphLookupSourceTextCharCountFor entryCount sourceParagraphCount blocks =
    if entryCount > paragraphLookupCacheMaxTotalEntries
       || sourceParagraphCount > paragraphLookupCacheMaxTotalSourceParagraphs then
        paragraphLookupCacheMaxTotalSourceTextChars + 1
    else
        documentBlocksSourceTextCharCountFor
            paragraphLookupCacheMaxTotalSourceParagraphs
            sourceParagraphCount
            blocks

let private tryParagraphLookupSourceMetricsFromBlockLookup blocks =
    match tryGetBlockLookup blocks with
    | ValueSome cached -> ValueSome (struct (cached.SourceParagraphCount, cached.SourceTextCharCount))
    | ValueNone -> ValueNone

let private rememberParagraphLookup cached =
    if not (paragraphLookupCacheRetainable cached) then
        let entries = paragraphLookupCacheWithoutBlocks cached.CachedBlocks
        if not (obj.ReferenceEquals(entries, paragraphLookupCache)) then
            setParagraphLookupCache entries
    else
        let trim entries =
            let rec loop keptCount keptEntries keptSourceParagraphs keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateEntries = max 1 candidate.EntryCount
                    let candidateSourceParagraphs = max 1 candidate.SourceParagraphCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if paragraphLookupCacheRetainable candidate
                       && keptCount < paragraphLookupCacheMaxEntries
                       && keptEntries + candidateEntries <= paragraphLookupCacheMaxTotalEntries
                       && keptSourceParagraphs + candidateSourceParagraphs <= paragraphLookupCacheMaxTotalSourceParagraphs
                       && keptSourceTextChars + candidateSourceTextChars <= paragraphLookupCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptEntries + candidateEntries)
                            (keptSourceParagraphs + candidateSourceParagraphs)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptEntries keptSourceParagraphs keptSourceTextChars acc tail
            loop 0 0 0 0 [] entries
        let entries =
            (cached
             :: paragraphLookupCacheWithoutBlocks cached.CachedBlocks)
            |> trim
        setParagraphLookupCache entries

let private addParagraphLookup (paragraphs: Dictionary<BlockId, ParagraphLookupEntry>) (paragraph: ParagraphBlock) =
    if not (paragraphs.ContainsKey paragraph.Id) then
        paragraphs.[paragraph.Id] <- {
            Paragraph = paragraph
            Result = Some paragraph
            RecentLookupIdSet = false
            RecentLookupId = ""
        }

type private ParagraphLookupFrame =
    | ParagraphLookupTable of TableBlock
    | ParagraphLookupCell of TableCell
    | ParagraphLookupParagraph of ParagraphBlock

let private addTableParagraphLookups
    (paragraphs: Dictionary<BlockId, ParagraphLookupEntry>)
    visitSourceItem
    addSourceParagraph
    (table: TableBlock) =
    let stack = Stack<ParagraphLookupFrame>()
    let pushCellFrames (cells: TableCell list) =
        let cellArray = List.toArray cells
        for index = cellArray.Length - 1 downto 0 do
            visitSourceItem ()
            stack.Push(ParagraphLookupCell cellArray.[index])
    let pushCellContentFrames (cell: TableCell) =
        let frames = ResizeArray<ParagraphLookupFrame>()
        let mutable content = cell.Content
        let mutable contentIndex = 0
        let mutable tables = cell.NestedTables
        let mutable anchors = cell.NestedTableAnchors
        while not tables.IsEmpty do
            match tables with
            | table :: tableTail ->
                let anchor =
                    match anchors with
                    | anchor :: anchorTail ->
                        anchors <- anchorTail
                        anchor
                    | [] -> System.Int32.MaxValue
                let target = max contentIndex anchor
                while contentIndex < target && not content.IsEmpty do
                    match content with
                    | paragraph :: paragraphTail ->
                        frames.Add(ParagraphLookupParagraph paragraph)
                        content <- paragraphTail
                        contentIndex <- contentIndex + 1
                    | [] -> ()
                frames.Add(ParagraphLookupTable table)
                tables <- tableTail
            | [] -> ()
        while not content.IsEmpty do
            match content with
            | paragraph :: paragraphTail ->
                frames.Add(ParagraphLookupParagraph paragraph)
                content <- paragraphTail
            | [] -> ()
        for index = frames.Count - 1 downto 0 do
            stack.Push frames.[index]
    stack.Push(ParagraphLookupTable table)
    while stack.Count > 0 do
        match stack.Pop() with
        | ParagraphLookupParagraph paragraph ->
            addSourceParagraph ()
            addParagraphLookup paragraphs paragraph
        | ParagraphLookupCell cell ->
            pushCellContentFrames cell
        | ParagraphLookupTable table ->
            visitSourceItem ()
            let rowArray = List.toArray table.Rows
            for rowIndex = rowArray.Length - 1 downto 0 do
                visitSourceItem ()
                pushCellFrames rowArray.[rowIndex].Cells

let private addBlockParagraphLookups
    (paragraphs: Dictionary<BlockId, ParagraphLookupEntry>)
    visitSourceItem
    addSourceParagraph
    block =
    match block with
    | Paragraph paragraph
    | Heading (_, paragraph) ->
        addSourceParagraph ()
        addParagraphLookup paragraphs paragraph
    | Table table ->
        addTableParagraphLookups paragraphs visitSourceItem addSourceParagraph table
    | TextBox textBox ->
        for paragraph in textBox.Content do
            addSourceParagraph ()
            addParagraphLookup paragraphs paragraph

let private buildParagraphLookup (doc: Document) =
    let paragraphs = Dictionary<BlockId, ParagraphLookupEntry>(blockIdLookupComparer)
    let mutable sourceParagraphCount = 0
    let mutable visitedSourceItems = 0
    let maxVisitedSourceItems = sourceStructureVisitedItemLimit paragraphLookupCacheMaxTotalSourceParagraphs
    let canCountSource () =
        sourceParagraphCount <= paragraphLookupCacheMaxTotalSourceParagraphs
        && visitedSourceItems <= maxVisitedSourceItems
    let visitSourceItem () =
        if canCountSource () then
            visitedSourceItems <- visitedSourceItems + 1
    let addSourceParagraph () =
        if canCountSource () then
            sourceParagraphCount <- sourceParagraphCount + 1
    let sourceParagraphCountResult () =
        if visitedSourceItems > maxVisitedSourceItems
           || sourceStructureOverwhelmsParagraphs paragraphLookupCacheMaxTotalSourceParagraphs sourceParagraphCount visitedSourceItems then
            paragraphLookupCacheMaxTotalSourceParagraphs + 1
        else
            min sourceParagraphCount (paragraphLookupCacheMaxTotalSourceParagraphs + 1)
    for block in doc.Blocks do
        visitSourceItem ()
        addBlockParagraphLookups paragraphs visitSourceItem addSourceParagraph block
    let struct (sourceParagraphCount, sourceTextCharCount) =
        match tryParagraphLookupSourceMetricsFromBlockLookup doc.Blocks with
        | ValueSome metrics -> metrics
        | ValueNone ->
            let sourceParagraphCount =
                if paragraphs.Count > paragraphLookupCacheMaxTotalEntries then
                    paragraphLookupCacheMaxTotalSourceParagraphs + 1
                else
                    sourceParagraphCountResult ()
            let sourceTextCharCount =
                paragraphLookupSourceTextCharCountFor paragraphs.Count sourceParagraphCount doc.Blocks
            struct (sourceParagraphCount, sourceTextCharCount)
    let paragraphFastSlots = Array.zeroCreate<ParagraphLookupEntry> (lookupEntryFastSlotSize paragraphs.Count)
    for entry in paragraphs.Values do
        setParagraphLookupEntryFastSlot paragraphFastSlots entry
    let cached = {
        CachedBlocks = doc.Blocks
        EntryCount = paragraphs.Count
        SourceParagraphCount = sourceParagraphCount
        SourceTextCharCount = sourceTextCharCount
        Paragraphs = paragraphs
        ParagraphFastSlots = paragraphFastSlots
    }
    rememberParagraphLookup cached
    cached

let private paragraphLookup (doc: Document) =
    match tryGetParagraphLookup doc.Blocks with
    | ValueSome cached -> cached
    | ValueNone -> buildParagraphLookup doc

/// Find an editable paragraph anywhere in the document, including table cells.
let private findParagraphFromParagraphLookup (doc: Document) paragraphId =
    let lookup = paragraphLookup doc
    let fast = lookup.ParagraphFastSlots.[paragraphLookupEntryFastSlot lookup.ParagraphFastSlots paragraphId]
    if not (obj.ReferenceEquals(fast, null))
       && paragraphLookupEntryMatches fast paragraphId then
        fast.Result
    else
        let mutable entry = Unchecked.defaultof<ParagraphLookupEntry>
        if lookup.Paragraphs.TryGetValue(paragraphId, &entry) then
            entry.RecentLookupIdSet <- true
            entry.RecentLookupId <- paragraphId
            setParagraphLookupEntryFastSlot lookup.ParagraphFastSlots entry
            entry.Result
        else
            None

let findParagraph (doc: Document) paragraphId =
    let topLevel =
        match tryGetBlockLookup doc.Blocks with
        | ValueSome lookup ->
            match tryFindBlockLookupEntry lookup paragraphId with
            | ValueSome entry ->
                entry.TopParagraphResult
            | ValueNone -> None
        | ValueNone -> None
    match topLevel with
    | Some _ -> topLevel
    | None -> findParagraphFromParagraphLookup doc paragraphId

let private tryReplaceTopLevelParagraphBeforeNestedContent doc paragraphId newPara =
    let scan () =
        let mutable remaining = doc.Blocks
        let mutable result: Document voption = ValueNone
        let mutable blockedByNestedContent = false
        let mutable index = 0

        while result.IsNone && not blockedByNestedContent do
            match remaining with
            | [] ->
                result <- ValueSome doc
            | Paragraph para :: tail when para.Id = paragraphId ->
                result <-
                    if paragraphReplacementMatches para newPara then
                        ValueSome doc
                    else
                        let nextBlocks =
                            replaceTopLevelBlockAtIndexAndRemember index (Paragraph newPara) paragraphId doc.Blocks
                        ValueSome { doc with Blocks = nextBlocks }
            | Heading (level, para) :: tail when para.Id = paragraphId ->
                result <-
                    if paragraphReplacementMatches para newPara then
                        ValueSome doc
                    else
                        let nextBlocks =
                            replaceTopLevelBlockAtIndexAndRemember index (Heading (level, newPara)) paragraphId doc.Blocks
                        ValueSome { doc with Blocks = nextBlocks }
            | (Table _ | TextBox _) :: _ ->
                blockedByNestedContent <- true
            | block :: tail ->
                remaining <- tail
                index <- index + 1

        if blockedByNestedContent then ValueNone else result

    match tryGetTopLevelParagraphEditEntry doc.Blocks paragraphId with
    | ValueSome entry ->
        match tryReplaceTopLevelParagraphFromCachedEntry doc paragraphId entry newPara with
        | ValueSome updatedDoc -> ValueSome updatedDoc
        | ValueNone -> scan ()
    | ValueNone ->
        scan ()

let private tryUpdateTopLevelParagraphBeforeNestedContent doc paragraphId updateParagraph =
    let scan () =
        let mutable remaining = doc.Blocks
        let mutable result: Document voption = ValueNone
        let mutable blockedByNestedContent = false
        let mutable index = 0

        while result.IsNone && not blockedByNestedContent do
            match remaining with
            | [] ->
                result <- ValueSome doc
            | Paragraph para :: tail when para.Id = paragraphId ->
                let nextPara = updateParagraph para
                result <-
                    if paragraphReplacementMatches para nextPara then
                        ValueSome doc
                    else
                        let nextBlocks =
                            replaceTopLevelBlockAtIndexAndRemember index (Paragraph nextPara) paragraphId doc.Blocks
                        ValueSome { doc with Blocks = nextBlocks }
            | Heading (level, para) :: tail when para.Id = paragraphId ->
                let nextPara = updateParagraph para
                result <-
                    if paragraphReplacementMatches para nextPara then
                        ValueSome doc
                    else
                        let nextBlocks =
                            replaceTopLevelBlockAtIndexAndRemember index (Heading (level, nextPara)) paragraphId doc.Blocks
                        ValueSome { doc with Blocks = nextBlocks }
            | (Table _ | TextBox _) :: _ ->
                blockedByNestedContent <- true
            | block :: tail ->
                remaining <- tail
                index <- index + 1

        if blockedByNestedContent then ValueNone else result

    match tryGetTopLevelParagraphEditEntry doc.Blocks paragraphId with
    | ValueSome entry ->
        match tryUpdateTopLevelParagraphFromCachedEntry doc paragraphId entry updateParagraph with
        | ValueSome updatedDoc -> ValueSome updatedDoc
        | ValueNone -> scan ()
    | ValueNone ->
        scan ()

/// Replace an editable paragraph anywhere in the document, preserving its container.
let replaceParagraph (doc: Document) paragraphId newPara =
    match tryReplaceTopLevelParagraphBeforeNestedContent doc paragraphId newPara with
    | ValueSome updatedDoc -> updatedDoc
    | ValueNone ->
        let mutable shouldReplace = false
        let mutable scanning = true
        let mutable remaining = doc.Blocks
        while scanning do
            match remaining with
            | [] ->
                scanning <- false
            | Paragraph para :: tail ->
                if para.Id = paragraphId then
                    shouldReplace <- not (paragraphReplacementMatches para newPara)
                    scanning <- false
                else
                    remaining <- tail
            | Heading (_, para) :: tail ->
                if para.Id = paragraphId then
                    shouldReplace <- not (paragraphReplacementMatches para newPara)
                    scanning <- false
                else
                    remaining <- tail
            | (Table _ as block) :: tail
            | (TextBox _ as block) :: tail ->
                let _, blockChanged = replaceParagraphInBlock paragraphId newPara block
                if blockChanged then
                    shouldReplace <- true
                    scanning <- false
                else
                    remaining <- tail
        if shouldReplace then
            let rec revAppend reversed tail =
                match reversed with
                | [] -> tail
                | head :: rest -> revAppend rest (head :: tail)
            let rec loop prefix remaining =
                match remaining with
                | [] -> None
                | block :: tail ->
                    let block, blockChanged = replaceParagraphInBlock paragraphId newPara block
                    if blockChanged then
                        Some (revAppend prefix (block :: tail))
                    else
                        loop (block :: prefix) tail
            match loop [] doc.Blocks with
            | Some blocks -> { doc with Blocks = blocks }
            | None -> doc
        else
            doc

let private updateParagraph (doc: Document) paragraphId update =
    let rec revAppend reversed tail =
        match reversed with
        | [] -> tail
        | head :: rest -> revAppend rest (head :: tail)
    let rec loop prefix remaining =
        match remaining with
        | [] -> None
        | block :: tail ->
            let block, blockChanged = updateParagraphInBlock paragraphId update block
            if blockChanged then
                Some (revAppend prefix (block :: tail))
            else
                loop (block :: prefix) tail
    match loop [] doc.Blocks with
    | Some blocks -> { doc with Blocks = blocks }
    | None -> doc

let private editableParagraphsInOrder doc =
    let paragraphs = ResizeArray<ParagraphBlock>()
    for block in doc.Blocks do
        match block with
        | Paragraph p
        | Heading (_, p) -> paragraphs.Add p
        | Table t ->
            for row in t.Rows do
                for cell in row.Cells do
                    for para in cell.Content do
                        paragraphs.Add para
        | TextBox t ->
            for para in t.Content do
                paragraphs.Add para
    resizeArrayToList paragraphs

let private blockParagraphs (block: Block) =
    match block with
    | Paragraph p
    | Heading (_, p) -> [p]
    | Table table ->
        table.Rows
        |> List.collect (fun row -> row.Cells)
        |> List.collect (fun cell -> cell.Content)
    | TextBox textBox -> textBox.Content

let private iterBlockParagraphs (block: Block) f =
    match block with
    | Paragraph p
    | Heading (_, p) -> f p
    | Table table ->
        for row in table.Rows do
            for cell in row.Cells do
                for para in cell.Content do
                    f para
    | TextBox textBox ->
        for para in textBox.Content do
            f para

let private iterBlockParagraphsUntil (block: Block) f =
    let mutable keepGoing = true
    match block with
    | Paragraph p
    | Heading (_, p) ->
        keepGoing <- f p
    | Table table ->
        let mutable rows = table.Rows
        while keepGoing && not rows.IsEmpty do
            match rows with
            | row :: rowTail ->
                let mutable cells = row.Cells
                while keepGoing && not cells.IsEmpty do
                    match cells with
                    | cell :: cellTail ->
                        let mutable content = cell.Content
                        while keepGoing && not content.IsEmpty do
                            match content with
                            | para :: paraTail ->
                                keepGoing <- f para
                                content <- paraTail
                            | [] -> ()
                        cells <- cellTail
                    | [] -> ()
                rows <- rowTail
            | [] -> ()
    | TextBox textBox ->
        let mutable content = textBox.Content
        while keepGoing && not content.IsEmpty do
            match content with
            | para :: tail ->
                keepGoing <- f para
                content <- tail
            | [] -> ()
    keepGoing

let private iterEditableParagraphs doc f =
    for block in doc.Blocks do
        iterBlockParagraphs block f

type private EditableParagraphEndpointsCache = {
    EndpointSourceBlocks: Block list
    EndpointFirst: ParagraphBlock voption
    EndpointLast: ParagraphBlock voption
    EndpointParagraphCount: int
    EndpointSourceTextCharCount: int
}

let mutable private editableParagraphEndpointsCache: EditableParagraphEndpointsCache list = []
let private editableParagraphEndpointsCacheByBlocks = Dictionary<Block list, EditableParagraphEndpointsCache>(documentInfoBlocksComparer)
let private editableParagraphEndpointsFastSlots = Array.zeroCreate<EditableParagraphEndpointsCache> 256
let private editableParagraphEndpointsCacheMaxEntries = 256
let private editableParagraphEndpointsCacheMaxTotalParagraphs = 200000
let private editableParagraphEndpointsCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars

let private editableParagraphEndpointsCacheMatches blocks (cached: EditableParagraphEndpointsCache) =
    obj.ReferenceEquals(cached.EndpointSourceBlocks, blocks)

let private editableParagraphEndpointsCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<EditableParagraphEndpointsCache>
    if editableParagraphEndpointsCacheByBlocks.TryGetValue(blocks, &cached)
       && editableParagraphEndpointsCacheMatches blocks cached then
        editableParagraphEndpointsCache
        |> List.filter (editableParagraphEndpointsCacheMatches blocks >> not)
    else
        editableParagraphEndpointsCache

let private setEditableParagraphEndpointsCache entries =
    editableParagraphEndpointsCache <- entries
    editableParagraphEndpointsCacheByBlocks.Clear()
    clearArray editableParagraphEndpointsFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<EditableParagraphEndpointsCache>
        if not (editableParagraphEndpointsCacheByBlocks.TryGetValue(entry.EndpointSourceBlocks, &existing)) then
            editableParagraphEndpointsCacheByBlocks.[entry.EndpointSourceBlocks] <- entry
            let slot = documentInfoSlot entry.EndpointSourceBlocks editableParagraphEndpointsFastSlots.Length
            editableParagraphEndpointsFastSlots.[slot] <- entry

let private tryGetEditableParagraphEndpointsCache blocks =
    let slot = documentInfoSlot blocks editableParagraphEndpointsFastSlots.Length
    let fast = editableParagraphEndpointsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && editableParagraphEndpointsCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<EditableParagraphEndpointsCache>
        if editableParagraphEndpointsCacheByBlocks.TryGetValue(blocks, &cached) then
            editableParagraphEndpointsFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private trimEditableParagraphEndpointsCache entries =
    let rec loop remaining kept count totalParagraphs totalSourceTextChars =
        match remaining with
        | [] -> List.rev kept
        | entry :: rest ->
            let paragraphCount = max 1 entry.EndpointParagraphCount
            let sourceTextCharCount = max 1 entry.EndpointSourceTextCharCount
            let nextTotalParagraphs = totalParagraphs + paragraphCount
            let nextTotalSourceTextChars = totalSourceTextChars + sourceTextCharCount
            if count < editableParagraphEndpointsCacheMaxEntries
               && nextTotalParagraphs <= editableParagraphEndpointsCacheMaxTotalParagraphs
               && nextTotalSourceTextChars <= editableParagraphEndpointsCacheMaxTotalSourceTextChars then
                loop rest (entry :: kept) (count + 1) nextTotalParagraphs nextTotalSourceTextChars
            else
                loop rest kept count totalParagraphs totalSourceTextChars
    loop entries [] 0 0 0

let private rememberEditableParagraphEndpoints cached =
    cached
    :: editableParagraphEndpointsCacheWithoutBlocks cached.EndpointSourceBlocks
    |> trimEditableParagraphEndpointsCache
    |> setEditableParagraphEndpointsCache

let private editableParagraphEndpoints doc =
    match tryGetEditableParagraphEndpointsCache doc.Blocks with
    | ValueSome cached ->
        cached.EndpointFirst, cached.EndpointLast
    | ValueNone ->
        let mutable firstPara: ParagraphBlock voption = ValueNone
        let mutable lastPara: ParagraphBlock voption = ValueNone
        let mutable paragraphCount = 0
        iterEditableParagraphs doc (fun para ->
            paragraphCount <- paragraphCount + 1
            match firstPara with
            | ValueNone -> firstPara <- ValueSome para
            | ValueSome _ -> ()
            lastPara <- ValueSome para)
        rememberEditableParagraphEndpoints {
            EndpointSourceBlocks = doc.Blocks
            EndpointFirst = firstPara
            EndpointLast = lastPara
            EndpointParagraphCount = paragraphCount
            EndpointSourceTextCharCount =
                documentBlocksSourceTextCharCountFor
                    editableParagraphEndpointsCacheMaxTotalParagraphs
                    paragraphCount
                    doc.Blocks
        }
        firstPara, lastPara

type private SelectAllDocumentCacheEntry = {
    Source: Document
    Result: Document
    SourceParagraphCount: int
    SourceTextCharCount: int
}

let mutable private selectAllDocumentCache: SelectAllDocumentCacheEntry list = []
let private selectAllDocumentCacheBySource =
    Dictionary<Document, SelectAllDocumentCacheEntry>(documentReferenceComparer)
let private selectAllDocumentFastSlots = Array.zeroCreate<SelectAllDocumentCacheEntry> 256
let private selectAllDocumentCacheMaxEntries = 256
let private selectAllDocumentCacheMaxTotalSourceParagraphs = editableParagraphEndpointsCacheMaxTotalParagraphs
let private selectAllDocumentCacheMaxTotalSourceTextChars = editableParagraphEndpointsCacheMaxTotalSourceTextChars

let private selectAllDocumentCacheMatches doc (cached: SelectAllDocumentCacheEntry) =
    obj.ReferenceEquals(cached.Source, doc)

let private setSelectAllDocumentCache entries =
    selectAllDocumentCache <- entries
    selectAllDocumentCacheBySource.Clear()
    clearArray selectAllDocumentFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<SelectAllDocumentCacheEntry>
        if not (selectAllDocumentCacheBySource.TryGetValue(entry.Source, &existing)) then
            selectAllDocumentCacheBySource.[entry.Source] <- entry
            let slot = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(entry.Source) &&& (selectAllDocumentFastSlots.Length - 1)
            selectAllDocumentFastSlots.[slot] <- entry

let private tryGetSelectAllDocument doc =
    let slot = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(doc) &&& (selectAllDocumentFastSlots.Length - 1)
    let fast = selectAllDocumentFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && selectAllDocumentCacheMatches doc fast then
        ValueSome fast.Result
    else
        let mutable cached = Unchecked.defaultof<SelectAllDocumentCacheEntry>
        if selectAllDocumentCacheBySource.TryGetValue(doc, &cached) then
            selectAllDocumentFastSlots.[slot] <- cached
            ValueSome cached.Result
        else
            ValueNone

let private selectAllDocumentCacheWithoutSource doc =
    let mutable cached = Unchecked.defaultof<SelectAllDocumentCacheEntry>
    if selectAllDocumentCacheBySource.TryGetValue(doc, &cached)
       && selectAllDocumentCacheMatches doc cached then
        selectAllDocumentCache
        |> List.filter (selectAllDocumentCacheMatches doc >> not)
    else
        selectAllDocumentCache

let private selectAllDocumentCacheRetainable (cached: SelectAllDocumentCacheEntry) =
    cached.SourceParagraphCount <= selectAllDocumentCacheMaxTotalSourceParagraphs
    && cached.SourceTextCharCount <= selectAllDocumentCacheMaxTotalSourceTextChars

let private rememberSelectAllDocument source result =
    if not (obj.ReferenceEquals(source, result)) then
        let sourceParagraphCount, sourceTextCharCount =
            match tryGetEditableParagraphEndpointsCache source.Blocks with
            | ValueSome cached -> cached.EndpointParagraphCount, cached.EndpointSourceTextCharCount
            | ValueNone ->
                let paragraphCount = countSourceParagraphs source.Blocks
                paragraphCount,
                documentBlocksSourceTextCharCountFor
                    selectAllDocumentCacheMaxTotalSourceParagraphs
                    paragraphCount
                    source.Blocks
        let entry = {
            Source = source
            Result = result
            SourceParagraphCount = sourceParagraphCount
            SourceTextCharCount = sourceTextCharCount
        }
        if selectAllDocumentCacheRetainable entry then
            let rec trim remaining kept count totalParagraphs totalSourceTextChars =
                match remaining with
                | [] -> List.rev kept
                | candidate :: rest ->
                    let candidateParagraphs = max 1 candidate.SourceParagraphCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    let nextTotalParagraphs = totalParagraphs + candidateParagraphs
                    let nextTotalSourceTextChars = totalSourceTextChars + candidateSourceTextChars
                    if selectAllDocumentCacheRetainable candidate
                       && count < selectAllDocumentCacheMaxEntries
                       && nextTotalParagraphs <= selectAllDocumentCacheMaxTotalSourceParagraphs
                       && nextTotalSourceTextChars <= selectAllDocumentCacheMaxTotalSourceTextChars then
                        trim rest (candidate :: kept) (count + 1) nextTotalParagraphs nextTotalSourceTextChars
                    else
                        trim rest kept count totalParagraphs totalSourceTextChars
            let entries = entry :: selectAllDocumentCacheWithoutSource source
            trim entries [] 0 0 0
            |> setSelectAllDocumentCache

let private iterEditableParagraphsUntil doc f =
    let mutable keepGoing = true
    let mutable blocks = doc.Blocks
    while keepGoing && not blocks.IsEmpty do
        match blocks with
        | block :: tail ->
            keepGoing <- iterBlockParagraphsUntil block f
            blocks <- tail
        | [] -> ()

let private tryFindEditableParagraphUntil doc paragraphId =
    let mutable found: ParagraphBlock voption = ValueNone
    iterEditableParagraphsUntil doc (fun para ->
        if para.Id = paragraphId then
            found <- ValueSome para
            false
        else
            true)
    match found with
    | ValueSome para -> Some para
    | ValueNone -> None

let private tryFindTopLevelParagraphUntil doc paragraphId =
    let mutable found: ParagraphBlock voption = ValueNone
    let mutable remaining = doc.Blocks
    while found.IsNone && not remaining.IsEmpty do
        match remaining with
        | Paragraph para :: tail
        | Heading (_, para) :: tail ->
            if blockIdKeyEquals para.Id paragraphId then
                found <- ValueSome para
            remaining <- tail
        | _ :: tail ->
            remaining <- tail
        | [] -> ()
    match found with
    | ValueSome para -> Some para
    | ValueNone -> None

type private OrderedEditableParagraphsCache = {
    OrderedSourceBlocks: Block list
    EntryCount: int
    SourceTextCharCount: int
    OrderedParagraphs: ParagraphBlock array
    OrderedIndexById: Dictionary<BlockId, int>
}

let mutable private orderedEditableParagraphsCache: OrderedEditableParagraphsCache list = []
let private orderedEditableParagraphsCacheByBlocks = Dictionary<Block list, OrderedEditableParagraphsCache>(documentInfoBlocksComparer)
let private orderedEditableParagraphsFastSlots = Array.zeroCreate<OrderedEditableParagraphsCache> 256
let private orderedEditableParagraphsCacheMaxEntries = 256
let private orderedEditableParagraphsCacheMaxTotalEntries = 200000
let private orderedEditableParagraphsCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars

let private orderedEditableParagraphsCacheMatches blocks (cached: OrderedEditableParagraphsCache) =
    obj.ReferenceEquals(cached.OrderedSourceBlocks, blocks)

let private orderedEditableParagraphsCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<OrderedEditableParagraphsCache>
    if orderedEditableParagraphsCacheByBlocks.TryGetValue(blocks, &cached)
       && orderedEditableParagraphsCacheMatches blocks cached then
        orderedEditableParagraphsCache
        |> List.filter (orderedEditableParagraphsCacheMatches blocks >> not)
    else
        orderedEditableParagraphsCache

let private setOrderedEditableParagraphsCache entries =
    orderedEditableParagraphsCache <- entries
    orderedEditableParagraphsCacheByBlocks.Clear()
    clearArray orderedEditableParagraphsFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<OrderedEditableParagraphsCache>
        if not (orderedEditableParagraphsCacheByBlocks.TryGetValue(entry.OrderedSourceBlocks, &existing)) then
            orderedEditableParagraphsCacheByBlocks.[entry.OrderedSourceBlocks] <- entry
            let slot = documentInfoSlot entry.OrderedSourceBlocks orderedEditableParagraphsFastSlots.Length
            orderedEditableParagraphsFastSlots.[slot] <- entry

let private tryGetOrderedEditableParagraphsCache blocks =
    let slot = documentInfoSlot blocks orderedEditableParagraphsFastSlots.Length
    let fast = orderedEditableParagraphsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && orderedEditableParagraphsCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<OrderedEditableParagraphsCache>
        if orderedEditableParagraphsCacheByBlocks.TryGetValue(blocks, &cached) then
            orderedEditableParagraphsFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private removeOrderedEditableParagraphsCacheForBlocks blocks =
    let entries = orderedEditableParagraphsCacheWithoutBlocks blocks
    if not (obj.ReferenceEquals(entries, orderedEditableParagraphsCache)) then
        setOrderedEditableParagraphsCache entries

let private orderedEditableParagraphsCacheRetainable (cached: OrderedEditableParagraphsCache) =
    cached.EntryCount <= orderedEditableParagraphsCacheMaxTotalEntries
    && cached.SourceTextCharCount <= orderedEditableParagraphsCacheMaxTotalSourceTextChars

let private rememberOrderedEditableParagraphs cached =
    if not (orderedEditableParagraphsCacheRetainable cached) then
        removeOrderedEditableParagraphsCacheForBlocks cached.OrderedSourceBlocks
    else
        let trim entries =
            let rec loop keptCount keptEntries keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateEntries = max 1 candidate.EntryCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if orderedEditableParagraphsCacheRetainable candidate
                       && keptCount < orderedEditableParagraphsCacheMaxEntries
                       && keptEntries + candidateEntries <= orderedEditableParagraphsCacheMaxTotalEntries
                       && keptSourceTextChars + candidateSourceTextChars <= orderedEditableParagraphsCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptEntries + candidateEntries)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptEntries keptSourceTextChars acc tail
            loop 0 0 0 [] entries
        (cached
         :: orderedEditableParagraphsCacheWithoutBlocks cached.OrderedSourceBlocks)
        |> trim
        |> setOrderedEditableParagraphsCache

let private orderedEditableParagraphs doc =
    let build () =
        let paragraphs = ResizeArray<ParagraphBlock>()
        let indexById = Dictionary<BlockId, int>()
        iterEditableParagraphs doc (fun para ->
            if not (indexById.ContainsKey para.Id) then
                indexById.[para.Id] <- paragraphs.Count
            paragraphs.Add para)
        let cached = {
            OrderedSourceBlocks = doc.Blocks
            EntryCount = paragraphs.Count
            SourceTextCharCount =
                documentBlocksSourceTextCharCountFor
                    orderedEditableParagraphsCacheMaxTotalEntries
                    paragraphs.Count
                    doc.Blocks
            OrderedParagraphs = paragraphs.ToArray()
            OrderedIndexById = indexById
        }
        rememberOrderedEditableParagraphs cached
        cached
    match tryGetOrderedEditableParagraphsCache doc.Blocks with
    | ValueSome cached when orderedEditableParagraphsCacheRetainable cached ->
        cached
    | ValueSome cached ->
        rememberOrderedEditableParagraphs cached
        build ()
    | ValueNone ->
        build ()

let private paragraphIndexById doc =
    (orderedEditableParagraphs doc).OrderedIndexById

let private paragraphIndexOf (indexes: Dictionary<BlockId, int>) blockId =
    let mutable index = 0
    if indexes.TryGetValue(blockId, &index) then index else System.Int32.MaxValue

let private blockId (block: Block) =
    match block with
    | Paragraph p
    | Heading (_, p) -> p.Id
    | Table table -> table.Id
    | TextBox textBox -> textBox.Id

type private PageBreakBeforeLookupCache = {
    Styles: NamedStyle list
    ParagraphStyles: Dictionary<string, bool>
    DefaultBreak: bool
    StyleCount: int
    StyleTextCharCount: int
}

let mutable private pageBreakBeforeLookupCache: PageBreakBeforeLookupCache list = []
let private pageBreakBeforeLookupStylesComparer =
    { new IEqualityComparer<NamedStyle list> with
        member _.Equals(left, right) = obj.ReferenceEquals(left, right)
        member _.GetHashCode(value) = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value) }
let private pageBreakBeforeLookupCacheByStyles =
    Dictionary<NamedStyle list, PageBreakBeforeLookupCache>(pageBreakBeforeLookupStylesComparer)
let private pageBreakBeforeLookupFastSlots = Array.zeroCreate<PageBreakBeforeLookupCache> 256
let private pageBreakBeforeLookupCacheMaxEntries = 256
let private pageBreakBeforeLookupCacheMaxTotalStyles = 200000
let private pageBreakBeforeLookupCacheMaxTotalStyleTextChars = documentSourceTextCacheMaxTotalChars

let private pageBreakBeforeLookupCacheMatches styles (cached: PageBreakBeforeLookupCache) =
    obj.ReferenceEquals(cached.Styles, styles)

let private pageBreakBeforeLookupCacheRetainable cached =
    cached.StyleCount <= pageBreakBeforeLookupCacheMaxTotalStyles
    && cached.StyleTextCharCount <= pageBreakBeforeLookupCacheMaxTotalStyleTextChars

let private setPageBreakBeforeLookupCache entries =
    pageBreakBeforeLookupCache <- entries
    pageBreakBeforeLookupCacheByStyles.Clear()
    clearArray pageBreakBeforeLookupFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<PageBreakBeforeLookupCache>
        if not (pageBreakBeforeLookupCacheByStyles.TryGetValue(entry.Styles, &existing)) then
            pageBreakBeforeLookupCacheByStyles.[entry.Styles] <- entry
            let slot = documentInfoSlot entry.Styles pageBreakBeforeLookupFastSlots.Length
            pageBreakBeforeLookupFastSlots.[slot] <- entry

let private tryGetPageBreakBeforeLookupCache styles =
    let slot = documentInfoSlot styles pageBreakBeforeLookupFastSlots.Length
    let fast = pageBreakBeforeLookupFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && pageBreakBeforeLookupCacheMatches styles fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<PageBreakBeforeLookupCache>
        if pageBreakBeforeLookupCacheByStyles.TryGetValue(styles, &cached) then
            pageBreakBeforeLookupFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private pageBreakBeforeLookupCacheWithoutStyles styles =
    let mutable cached = Unchecked.defaultof<PageBreakBeforeLookupCache>
    if pageBreakBeforeLookupCacheByStyles.TryGetValue(styles, &cached)
       && pageBreakBeforeLookupCacheMatches styles cached then
        pageBreakBeforeLookupCache
        |> List.filter (pageBreakBeforeLookupCacheMatches styles >> not)
    else
        pageBreakBeforeLookupCache

let private optionalStringCharCount (value: string option) =
    match value with
    | Some value -> value.Length
    | None -> 0

let private textStyleStringCharCount (style: TextBuffer.TextStyle) =
    optionalStringCharCount style.Style
    + optionalStringCharCount style.UnderlineStyle
    + optionalStringCharCount style.UnderlineColor
    + optionalStringCharCount style.FontFamily
    + optionalStringCharCount style.Color
    + optionalStringCharCount style.Highlight
    + optionalStringCharCount style.Hyperlink
    + optionalStringCharCount style.CommentId
    + optionalStringCharCount style.CommentText
    + optionalStringCharCount style.CommentAuthor
    + optionalStringCharCount style.CommentDate
    + optionalStringCharCount style.NoteId
    + optionalStringCharCount style.NoteType
    + optionalStringCharCount style.NoteText
    + optionalStringCharCount style.BookmarkStartId
    + optionalStringCharCount style.BookmarkName
    + optionalStringCharCount style.BookmarkEndId
    + optionalStringCharCount style.FieldCode
    + optionalStringCharCount style.FieldResult
    + optionalStringCharCount style.EquationXml
    + optionalStringCharCount style.RevisionId
    + optionalStringCharCount style.RevisionType
    + optionalStringCharCount style.RevisionAuthor
    + optionalStringCharCount style.RevisionDate
    + optionalStringCharCount style.VertAlign

let private namedStyleStringCharCount (style: NamedStyle) =
    style.Id.Length
    + optionalStringCharCount style.Name
    + style.StyleType.Length
    + optionalStringCharCount style.BasedOn
    + optionalStringCharCount style.Next
    + optionalStringCharCount style.Linked
    + optionalStringCharCount style.Aliases
    + optionalStringCharCount style.LineSpacingRule
    + optionalStringCharCount style.BackgroundColor
    + (match style.Text with
       | Some textStyle -> textStyleStringCharCount textStyle
       | None -> 0)

let private pageBreakBeforeStylesTextCharCount (styles: NamedStyle list) =
    let limit = pageBreakBeforeLookupCacheMaxTotalStyleTextChars + 1
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = styles
    while total <= pageBreakBeforeLookupCacheMaxTotalStyleTextChars
          && seen <= pageBreakBeforeLookupCacheMaxTotalStyles
          && not remaining.IsEmpty do
        total <- min limit (total + namedStyleStringCharCount remaining.Head)
        seen <- seen + 1
        remaining <- remaining.Tail
    min total limit

let private trimPageBreakBeforeLookupCache caches =
    let rec loop remaining kept count totalStyles totalStyleTextChars =
        match remaining with
        | [] -> List.rev kept
        | cached :: rest ->
            let styleCount = max 1 cached.StyleCount
            let styleTextCharCount = max 1 cached.StyleTextCharCount
            let nextTotalStyles = totalStyles + styleCount
            let nextTotalStyleTextChars = totalStyleTextChars + styleTextCharCount
            if pageBreakBeforeLookupCacheRetainable cached
               && count < pageBreakBeforeLookupCacheMaxEntries
               && nextTotalStyles <= pageBreakBeforeLookupCacheMaxTotalStyles
               && nextTotalStyleTextChars <= pageBreakBeforeLookupCacheMaxTotalStyleTextChars then
                loop rest (cached :: kept) (count + 1) nextTotalStyles nextTotalStyleTextChars
            else
                loop rest kept count totalStyles totalStyleTextChars
    loop caches [] 0 0 0

let private rememberPageBreakBeforeLookup cached =
    if not (pageBreakBeforeLookupCacheRetainable cached) then
        let entries = pageBreakBeforeLookupCacheWithoutStyles cached.Styles
        if not (obj.ReferenceEquals(entries, pageBreakBeforeLookupCache)) then
            setPageBreakBeforeLookupCache entries
    else
        cached
        :: pageBreakBeforeLookupCacheWithoutStyles cached.Styles
        |> trimPageBreakBeforeLookupCache
        |> setPageBreakBeforeLookupCache

let private pageBreakBeforeLookup (doc: Document) =
    let cache =
        match tryGetPageBreakBeforeLookupCache doc.Styles with
        | ValueSome cached ->
            cached
        | ValueNone ->
            let paragraphStyles = Dictionary<string, bool>()
            let mutable defaultBreak = false
            for style in doc.Styles do
                match style.StyleType, style.PageBreakBefore with
                | "paragraph", Some pageBreakBefore ->
                    paragraphStyles.[style.Id] <- pageBreakBefore
                | "defaults", Some pageBreakBefore ->
                    defaultBreak <- pageBreakBefore
                | _ -> ()
            let cached = {
                Styles = doc.Styles
                ParagraphStyles = paragraphStyles
                DefaultBreak = defaultBreak
                StyleCount = doc.Styles.Length
                StyleTextCharCount = pageBreakBeforeStylesTextCharCount doc.Styles
            }
            rememberPageBreakBeforeLookup cached
            cached
    let paragraphPageBreakBefore (para: ParagraphBlock) =
        match para.Metadata.Style with
        | Some styleId ->
            let mutable pageBreakBefore = false
            if cache.ParagraphStyles.TryGetValue(styleId, &pageBreakBefore) then pageBreakBefore
            else cache.DefaultBreak
        | None ->
            cache.DefaultBreak
    paragraphPageBreakBefore

[<Struct>]
type private PageMapCacheKey = {
    SourceBlocks: Block list
    SourceStyles: NamedStyle list
}

type private PageMapCacheEntry = {
    Key: PageMapCacheKey
    PageMap: PageMap
    SourceParagraphCount: int
    SourceTextCharCount: int
    StyleCount: int
    StyleTextCharCount: int
    AnchorCount: int
}

let private pageMapCacheKeyHash key =
    let blockHash = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.SourceBlocks)
    let styleHash = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.SourceStyles)
    blockHash ^^^ (styleHash <<< 1)

let private pageMapCacheKeyMatches left right =
    obj.ReferenceEquals(left.SourceBlocks, right.SourceBlocks)
    && obj.ReferenceEquals(left.SourceStyles, right.SourceStyles)

let private pageMapCacheKeyComparer =
    { new IEqualityComparer<PageMapCacheKey> with
        member _.Equals(left, right) = pageMapCacheKeyMatches left right
        member _.GetHashCode(key) = pageMapCacheKeyHash key }

let mutable private pageMapCache: PageMapCacheEntry list = []
let private pageMapCacheByKey = Dictionary<PageMapCacheKey, PageMapCacheEntry>(pageMapCacheKeyComparer)
let private pageMapFastSlots = Array.zeroCreate<PageMapCacheEntry> 256
let private pageMapCacheMaxEntries = 256
let private pageMapCacheMaxTotalAnchors = 200000
let private pageMapCacheMaxTotalSourceParagraphs = sourceParagraphCountCacheLimit
let private pageMapCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars
let private pageMapCacheMaxTotalStyles = pageBreakBeforeLookupCacheMaxTotalStyles
let private pageMapCacheMaxTotalStyleTextChars = pageBreakBeforeLookupCacheMaxTotalStyleTextChars

let private pageMapCacheKey blocks styles = {
    SourceBlocks = blocks
    SourceStyles = styles
}

let private pageMapCacheRetainable cached =
    cached.AnchorCount <= pageMapCacheMaxTotalAnchors
    && cached.SourceParagraphCount <= pageMapCacheMaxTotalSourceParagraphs
    && cached.SourceTextCharCount <= pageMapCacheMaxTotalSourceTextChars
    && cached.StyleCount <= pageMapCacheMaxTotalStyles
    && cached.StyleTextCharCount <= pageMapCacheMaxTotalStyleTextChars

let private pageMapCacheEntryMatches key cached =
    pageMapCacheKeyMatches key cached.Key

let private setPageMapCache entries =
    pageMapCache <- entries
    pageMapCacheByKey.Clear()
    clearArray pageMapFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<PageMapCacheEntry>
        if not (pageMapCacheByKey.TryGetValue(entry.Key, &existing)) then
            pageMapCacheByKey.[entry.Key] <- entry
            pageMapFastSlots.[pageMapCacheKeyHash entry.Key &&& (pageMapFastSlots.Length - 1)] <- entry

let private tryGetPageMapCache key =
    let slot = pageMapCacheKeyHash key &&& (pageMapFastSlots.Length - 1)
    let fast = pageMapFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && pageMapCacheEntryMatches key fast
       && pageMapCacheRetainable fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<PageMapCacheEntry>
        if pageMapCacheByKey.TryGetValue(key, &cached)
           && pageMapCacheRetainable cached then
            pageMapFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private pageMapCacheWithoutKey key =
    let mutable cached = Unchecked.defaultof<PageMapCacheEntry>
    if pageMapCacheByKey.TryGetValue(key, &cached)
       && pageMapCacheEntryMatches key cached then
        pageMapCache
        |> List.filter (pageMapCacheEntryMatches key >> not)
    else
        pageMapCache

let private trimPageMapCache caches =
    let rec loop remaining kept count totalAnchors totalSourceParagraphs totalSourceTextChars totalStyles totalStyleTextChars =
        match remaining with
        | [] -> List.rev kept
        | cached :: rest ->
            let anchors = max 1 cached.AnchorCount
            let sourceParagraphs = max 1 cached.SourceParagraphCount
            let sourceTextChars = max 1 cached.SourceTextCharCount
            let styles = max 1 cached.StyleCount
            let styleTextChars = max 1 cached.StyleTextCharCount
            let nextAnchors = totalAnchors + anchors
            let nextSourceParagraphs = totalSourceParagraphs + sourceParagraphs
            let nextSourceTextChars = totalSourceTextChars + sourceTextChars
            let nextStyles = totalStyles + styles
            let nextStyleTextChars = totalStyleTextChars + styleTextChars
            if pageMapCacheRetainable cached
               && count < pageMapCacheMaxEntries
               && nextAnchors <= pageMapCacheMaxTotalAnchors
               && nextSourceParagraphs <= pageMapCacheMaxTotalSourceParagraphs
               && nextSourceTextChars <= pageMapCacheMaxTotalSourceTextChars
               && nextStyles <= pageMapCacheMaxTotalStyles
               && nextStyleTextChars <= pageMapCacheMaxTotalStyleTextChars then
                loop
                    rest
                    (cached :: kept)
                    (count + 1)
                    nextAnchors
                    nextSourceParagraphs
                    nextSourceTextChars
                    nextStyles
                    nextStyleTextChars
            else
                loop rest kept count totalAnchors totalSourceParagraphs totalSourceTextChars totalStyles totalStyleTextChars
    loop caches [] 0 0 0 0 0 0

let private rememberPageMap cached =
    if not (pageMapCacheRetainable cached) then
        let entries = pageMapCacheWithoutKey cached.Key
        if not (obj.ReferenceEquals(entries, pageMapCache)) then
            setPageMapCache entries
    else
        cached
        :: pageMapCacheWithoutKey cached.Key
        |> trimPageMapCache
        |> setPageMapCache

let private pageMapAnchorCount pageMap =
    pageMap.BlockAnchors.Length + pageMap.Fields.Length + pageMap.Bookmarks.Length

let private buildPageMapUncached (doc: Document) =
    let mutable physicalPageNumber = 1
    let mutable displayPageStart = 1
    let mutable displayPageStartPhysical = 1
    let blockAnchors = ResizeArray<PageBlockAnchor>()
    let fieldAnchors = ResizeArray<PageFieldAnchor>()
    let bookmarkAnchors = ResizeArray<PageBookmarkAnchor>()
    let seenBookmarkAnchors = HashSet<string>()
    let paragraphPageBreakBefore = pageBreakBeforeLookup doc

    let displayPageNumber () = displayPageStart + physicalPageNumber - displayPageStartPhysical

    let addParagraphAnchors para =
        let pageNumber = displayPageNumber ()
        for runValue in para.Buffer.Runs do
            match runValue with
            | Text run ->
                run.Style.FieldCode
                |> Option.iter (fun code ->
                    fieldAnchors.Add { PageFieldBlockId = para.Id; PageFieldStartOffset = run.Start; PageFieldCode = code; PageNumber = pageNumber })
                match run.Style.BookmarkName with
                | Some name when seenBookmarkAnchors.Add name ->
                    bookmarkAnchors.Add { PageBookmarkName = name; PageBookmarkBlockId = para.Id; PageBookmarkStartOffset = run.Start; PageNumber = pageNumber }
                | _ -> ()
            | Image _ -> ()

    for block in doc.Blocks do
        let mutable hasPageBreakBefore = false
        let mutable sectionPageStart = None
        let mutable hasPageBreakAfter = false
        iterBlockParagraphs block (fun para ->
            if paragraphPageBreakBefore para then
                hasPageBreakBefore <- true
            match para.Metadata.SectionPageNumberStart with
            | Some start -> sectionPageStart <- Some start
            | None -> ()
            if para.Metadata.PageBreakAfter || para.Metadata.SectionBreak = Some "nextPage" then
                hasPageBreakAfter <- true)
        if blockAnchors.Count > 0 && hasPageBreakBefore then
            physicalPageNumber <- physicalPageNumber + 1
        sectionPageStart
        |> Option.iter (fun start ->
            displayPageStart <- max 1 start
            displayPageStartPhysical <- physicalPageNumber)
        let pageNumber = displayPageNumber ()
        blockAnchors.Add { PageBlockId = blockId block; PageNumber = pageNumber }
        iterBlockParagraphs block addParagraphAnchors
        if hasPageBreakAfter then
            physicalPageNumber <- physicalPageNumber + 1

    {
        PageCount = max 1 physicalPageNumber
        BlockAnchors = resizeArrayToList blockAnchors
        Fields = resizeArrayToList fieldAnchors
        Bookmarks = resizeArrayToList bookmarkAnchors
    }

let buildPageMap (doc: Document) =
    let key = pageMapCacheKey doc.Blocks doc.Styles
    match tryGetPageMapCache key with
    | ValueSome cached -> cached.PageMap
    | ValueNone ->
        let pageMap = buildPageMapUncached doc
        let sourceParagraphCount = countSourceParagraphs doc.Blocks
        let cached = {
            Key = key
            PageMap = pageMap
            SourceParagraphCount = sourceParagraphCount
            SourceTextCharCount =
                documentBlocksSourceTextCharCountFor
                    pageMapCacheMaxTotalSourceParagraphs
                    sourceParagraphCount
                    doc.Blocks
            StyleCount = doc.Styles.Length
            StyleTextCharCount = pageBreakBeforeStylesTextCharCount doc.Styles
            AnchorCount = pageMapAnchorCount pageMap
        }
        rememberPageMap cached
        pageMap

let bookmarkInfos (doc: Document) =
    match tryGetBookmarkInfoCache doc.Blocks with
    | ValueSome cached ->
        cached.BookmarkInfos
    | ValueNone ->
        let infosByName = Dictionary<string, BookmarkInfo>()
        let namesInOrder = ResizeArray<string>()
        iterEditableParagraphs doc (fun para ->
            for runValue in para.Buffer.Runs do
                match runValue with
                | Text run ->
                    match run.Style.BookmarkStartId with
                    | Some bookmarkId ->
                        match run.Style.BookmarkName with
                        | Some bookmarkName ->
                            let mutable info = Unchecked.defaultof<BookmarkInfo>
                            if infosByName.TryGetValue(bookmarkName, &info) then
                                if info.BlockId = para.Id then
                                    infosByName.[bookmarkName] <- {
                                        info with
                                            StartOffset = min info.StartOffset run.Start
                                            EndOffset = max info.EndOffset (run.Start + run.Length)
                                    }
                                else
                                    ()
                            else
                                namesInOrder.Add bookmarkName
                                infosByName.[bookmarkName] <- {
                                    Id = bookmarkId
                                    Name = bookmarkName
                                    BlockId = para.Id
                                    StartOffset = run.Start
                                    EndOffset = run.Start + run.Length
                                }
                                ()
                        | None -> ()
                    | None -> ()
                | Image _ -> ())
        let infos = [
            for name in namesInOrder do
                let mutable info = Unchecked.defaultof<BookmarkInfo>
                if infosByName.TryGetValue(name, &info) then
                    yield info
        ]
        let struct (sourceParagraphCount, sourceTextCharCount) = documentInfoSourceMetricsFor infos.Length doc.Blocks
        rememberBookmarkInfos {
            BookmarkSourceBlocks = doc.Blocks
            SourceParagraphCount = sourceParagraphCount
            SourceTextCharCount = sourceTextCharCount
            BookmarkInfoCount = infos.Length
            BookmarkInfos = infos
        }
        infos

let contentControlInfos (doc: Document) =
    match tryGetContentControlInfoCache doc.Blocks with
    | ValueSome cached ->
        cached.ContentControlInfos
    | ValueNone ->
        let infos = ResizeArray<ContentControlInfo>()
        iterEditableParagraphs doc (fun para ->
            para.Metadata.ContentControlId
            |> Option.iter (fun controlId ->
                infos.Add {
                    Id = controlId
                    Tag = para.Metadata.ContentControlTag |> Option.defaultValue "content-control"
                    Alias = para.Metadata.ContentControlAlias |> Option.defaultValue (para.Metadata.ContentControlTag |> Option.defaultValue controlId)
                    BlockId = para.Id
                }))
        let infos = resizeArrayToList infos
        let struct (sourceParagraphCount, sourceTextCharCount) = documentInfoSourceMetricsFor infos.Length doc.Blocks
        rememberContentControlInfos {
            ContentControlSourceBlocks = doc.Blocks
            SourceParagraphCount = sourceParagraphCount
            SourceTextCharCount = sourceTextCharCount
            ContentControlInfoCount = infos.Length
            ContentControlInfos = infos
        }
        infos

let commentInfos (doc: Document) =
    match tryGetCommentInfoCache doc.Blocks with
    | ValueSome cached ->
        cached.CommentInfos
    | ValueNone ->
        let infos = ResizeArray<CommentInfo>()
        let seenAnchors = HashSet<string>()
        iterEditableParagraphs doc (fun para ->
            for runValue in para.Buffer.Runs do
                match runValue with
                | Text run ->
                    match run.Style.CommentId with
                    | Some id ->
                        let text = run.Style.CommentText |> Option.defaultValue ""
                        let key = id + "\u0000" + text
                        if seenAnchors.Add key then
                            infos.Add {
                                Id = id
                                Text = text
                                Author = run.Style.CommentAuthor
                                Date = run.Style.CommentDate
                                BlockId = para.Id
                                StartOffset = run.Start
                                EndOffset = run.Start + run.Length
                                Excerpt = TextBuffer.slice run.Start run.Length para.Buffer
                            }
                    | None -> ()
                | Image _ -> ())
        let infos = resizeArrayToList infos
        let struct (sourceParagraphCount, sourceTextCharCount) = documentInfoSourceMetricsFor infos.Length doc.Blocks
        rememberCommentInfos {
            CommentSourceBlocks = doc.Blocks
            SourceParagraphCount = sourceParagraphCount
            SourceTextCharCount = sourceTextCharCount
            CommentInfoCount = infos.Length
            CommentInfos = infos
        }
        infos

let private paragraphContentControlMatches controlId (para: ParagraphBlock) =
    match para.Metadata.ContentControlId with
    | None -> false
    | Some id ->
        id = controlId
        || (para.Metadata.ContentControlTag |> Option.defaultValue "content-control") = controlId

let private tryFindBookmarkInfo bookmarkId (doc: Document) =
    let mutable found = false
    let mutable result = Unchecked.defaultof<BookmarkInfo>
    iterEditableParagraphs doc (fun para ->
        if not found then
            let mutable paragraphHasMatch = false
            let mutable info = Unchecked.defaultof<BookmarkInfo>
            for runValue in para.Buffer.Runs do
                match runValue with
                | Text run ->
                    match run.Style.BookmarkStartId with
                    | Some runBookmarkId ->
                        match run.Style.BookmarkName with
                        | Some bookmarkName when runBookmarkId = bookmarkId || bookmarkName = bookmarkId ->
                            let runEnd = run.Start + run.Length
                            if paragraphHasMatch then
                                info <- {
                                    info with
                                        StartOffset = min info.StartOffset run.Start
                                        EndOffset = max info.EndOffset runEnd
                                }
                            else
                                paragraphHasMatch <- true
                                info <- {
                                    Id = runBookmarkId
                                    Name = bookmarkName
                                    BlockId = para.Id
                                    StartOffset = run.Start
                                    EndOffset = runEnd
                                }
                        | _ -> ()
                    | None -> ()
                | Image _ -> ()
            if paragraphHasMatch then
                result <- info
                found <- true)
    if found then Some result else None

let private selectionMatchesBookmark bookmarkId doc =
    match doc.Selection with
    | Some selection when selection.Anchor.BlockId = selection.Focus.BlockId ->
        match tryFindEditableParagraphUntil doc selection.Anchor.BlockId with
        | Some para ->
            let mutable found = false
            let mutable startOffset = 0
            let mutable endOffset = 0
            for runValue in para.Buffer.Runs do
                match runValue with
                | Text run when run.Style.BookmarkStartId = Some bookmarkId || run.Style.BookmarkName = Some bookmarkId ->
                    let runEnd = run.Start + run.Length
                    if found then
                        startOffset <- min startOffset run.Start
                        endOffset <- max endOffset runEnd
                    else
                        found <- true
                        startOffset <- run.Start
                        endOffset <- runEnd
                | _ -> ()
            found
                && selection.Anchor.Offset = startOffset
                && selection.Focus.Offset = endOffset
        | None -> false
    | _ -> false

let private selectionMatchesContentControl controlId doc =
    match doc.Selection with
    | Some selection when selection.Anchor.BlockId = selection.Focus.BlockId ->
        match tryFindEditableParagraphUntil doc selection.Anchor.BlockId with
        | Some para ->
            paragraphContentControlMatches controlId para
                && selection.Anchor.Offset = 0
                && selection.Focus.Offset = TextBuffer.length para.Buffer
        | None -> false
    | _ -> false

let selectBookmark bookmarkId doc =
    if selectionMatchesBookmark bookmarkId doc then
        doc
    else
        match tryFindBookmarkInfo bookmarkId doc with
        | None -> doc
        | Some info ->
            let anchor = { BlockId = info.BlockId; Offset = info.StartOffset }
            let focus = { BlockId = info.BlockId; Offset = info.EndOffset }
            match doc.Selection with
            | Some current when current.Anchor = anchor && current.Focus = focus ->
                doc
            | _ ->
                { doc with Selection = Some { Anchor = anchor; Focus = focus } }

let selectContentControl controlId doc =
    if selectionMatchesContentControl controlId doc then
        doc
    else
        let mutable selected = None
        iterEditableParagraphs doc (fun para ->
            if selected.IsNone && paragraphContentControlMatches controlId para then
                selected <- Some para)
        match selected with
        | None -> doc
        | Some para ->
            let anchor = { BlockId = para.Id; Offset = 0 }
            let focus = { BlockId = para.Id; Offset = TextBuffer.length para.Buffer }
            match doc.Selection with
            | Some current when current.Anchor = anchor && current.Focus = focus ->
                doc
            | _ ->
                { doc with Selection = Some { Anchor = anchor; Focus = focus } }

let private tryRemoveBookmarkFromParagraph bookmarkId para =
    let mutable result = ValueNone
    let mutable runs: ResizeArray<Run> = null
    let mutable prefixCount = 0
    let mutable changed = false
    let ensureChanged () =
        if not changed then
            changed <- true
            runs <- ResizeArray<Run>()
            let mutable copied = 0
            let mutable prefix = para.Buffer.Runs
            while copied < prefixCount do
                match prefix with
                | prefixRun :: tail ->
                    runs.Add prefixRun
                    prefix <- tail
                    copied <- copied + 1
                | [] -> copied <- prefixCount
    for runValue in para.Buffer.Runs do
        match runValue with
        | Text run ->
            if run.Style.BookmarkStartId = Some bookmarkId || run.Style.BookmarkName = Some bookmarkId then
                ensureChanged ()
                runs.Add(Text { run with Style = { run.Style with BookmarkStartId = None; BookmarkName = None; BookmarkEndId = None } })
            elif changed then
                runs.Add runValue
            else
                prefixCount <- prefixCount + 1
        | Image _ ->
            if changed then
                runs.Add runValue
            else
                prefixCount <- prefixCount + 1
    if changed then
        result <- ValueSome { para with Buffer = { para.Buffer with Runs = resizeArrayToList runs } }
    result

let private tryRemoveBookmarkFromTopLevelParagraphs bookmarkId doc =
    let blocks = ResizeArray<Block>(doc.Blocks.Length)
    let mutable pendingUnchangedBlocks = Unchecked.defaultof<ResizeArray<Block>>
    let mutable changed = false
    let mutable blockedByNestedContent = false
    let mutable tailAfterLastChange: Block list = []

    let flushPendingUnchangedBlocks () =
        if not (obj.ReferenceEquals(pendingUnchangedBlocks, null)) && pendingUnchangedBlocks.Count > 0 then
            blocks.AddRange pendingUnchangedBlocks
            pendingUnchangedBlocks.Clear()

    let addChangedBlock block tail =
        if changed then
            flushPendingUnchangedBlocks ()
        changed <- true
        blocks.Add block
        tailAfterLastChange <- tail

    let addUnchangedBlock block =
        if changed then
            if obj.ReferenceEquals(pendingUnchangedBlocks, null) then
                pendingUnchangedBlocks <- ResizeArray<Block>()
            pendingUnchangedBlocks.Add block
        else
            blocks.Add block

    let mutable remaining = doc.Blocks
    while not blockedByNestedContent && not remaining.IsEmpty do
        match remaining with
        | Paragraph para :: tail ->
            match tryRemoveBookmarkFromParagraph bookmarkId para with
            | ValueSome para -> addChangedBlock (Paragraph para) tail
            | ValueNone -> addUnchangedBlock remaining.Head
            remaining <- tail
        | Heading (level, para) :: tail ->
            match tryRemoveBookmarkFromParagraph bookmarkId para with
            | ValueSome para -> addChangedBlock (Heading (level, para)) tail
            | ValueNone -> addUnchangedBlock remaining.Head
            remaining <- tail
        | (Table _ | TextBox _) :: _ ->
            blockedByNestedContent <- true
        | [] -> ()

    if blockedByNestedContent then
        ValueNone
    elif changed then
        ValueSome { doc with Blocks = prependResizeArray blocks tailAfterLastChange }
    else
        ValueSome doc

let removeBookmark bookmarkId doc =
    match tryRemoveBookmarkFromTopLevelParagraphs bookmarkId doc with
    | ValueSome doc -> doc
    | ValueNone ->
        let mutable replacements: Dictionary<BlockId, ParagraphBlock> = null
        iterEditableParagraphs doc (fun para ->
            match tryRemoveBookmarkFromParagraph bookmarkId para with
            | ValueSome para ->
                if isNull replacements then
                    replacements <- Dictionary<BlockId, ParagraphBlock>()
                replacements.[para.Id] <- para
            | ValueNone -> ())
        if isNull replacements then doc else replaceParagraphs doc replacements

let renameBookmark bookmarkId bookmarkName doc =
    let normalizedName =
        if System.String.IsNullOrWhiteSpace bookmarkName then "Bookmark"
        else bookmarkName.Trim().Replace(" ", "_")

    let mutable replacements: Dictionary<BlockId, ParagraphBlock> = null
    iterEditableParagraphs doc (fun para ->
        let mutable runs: ResizeArray<Run> = null
        let mutable prefixCount = 0
        let mutable changed = false
        let ensureChanged () =
            if not changed then
                changed <- true
                runs <- ResizeArray<Run>()
                let mutable copied = 0
                let mutable prefix = para.Buffer.Runs
                while copied < prefixCount do
                    match prefix with
                    | prefixRun :: tail ->
                        runs.Add prefixRun
                        prefix <- tail
                        copied <- copied + 1
                    | [] -> copied <- prefixCount
        for runValue in para.Buffer.Runs do
            match runValue with
            | Text run ->
                if (run.Style.BookmarkStartId = Some bookmarkId || run.Style.BookmarkName = Some bookmarkId)
                    && run.Style.BookmarkName <> Some normalizedName then
                    ensureChanged ()
                    runs.Add(Text { run with Style = { run.Style with BookmarkName = Some normalizedName } })
                elif changed then
                    runs.Add runValue
                else
                    prefixCount <- prefixCount + 1
            | Image _ ->
                if changed then
                    runs.Add runValue
                else
                    prefixCount <- prefixCount + 1
        if changed then
            if isNull replacements then
                replacements <- Dictionary<BlockId, ParagraphBlock>()
            replacements.[para.Id] <- { para with Buffer = { para.Buffer with Runs = resizeArrayToList runs } })
    if isNull replacements then doc else replaceParagraphs doc replacements

let removeContentControl controlId doc =
    let mutable replacements: Dictionary<BlockId, ParagraphBlock> = null
    iterEditableParagraphs doc (fun para ->
        if paragraphContentControlMatches controlId para then
            if isNull replacements then
                replacements <- Dictionary<BlockId, ParagraphBlock>()
            replacements.[para.Id] <- {
                para with
                    Metadata = {
                        para.Metadata with
                            ContentControlId = None
                            ContentControlTag = None
                            ContentControlAlias = None
                    }
            })
    if isNull replacements then doc else replaceParagraphs doc replacements

let updateContentControl controlId newControlId tag alias doc =
    let normalize fallback (value: string) =
        if System.String.IsNullOrWhiteSpace value then fallback else value.Trim()
    let newControlId = normalize controlId newControlId
    let tag = normalize "content-control" tag
    let alias = normalize tag alias

    let mutable replacements: Dictionary<BlockId, ParagraphBlock> = null
    iterEditableParagraphs doc (fun para ->
        if paragraphContentControlMatches controlId para then
            if para.Metadata.ContentControlId <> Some newControlId
               || para.Metadata.ContentControlTag <> Some tag
               || para.Metadata.ContentControlAlias <> Some alias then
                let nextMetadata = {
                    para.Metadata with
                        ContentControlId = Some newControlId
                        ContentControlTag = Some tag
                        ContentControlAlias = Some alias
                }
                if isNull replacements then
                    replacements <- Dictionary<BlockId, ParagraphBlock>()
                replacements.[para.Id] <- { para with Metadata = nextMetadata })
    if isNull replacements then doc else replaceParagraphs doc replacements

let private emptyTableCell () =
    {
        Content = [
            {
                Id = nextBlockId ()
                Buffer = TextBuffer.empty
                Metadata = defaultMetadata
            }
        ]
        NestedTables = []
        NestedTableAnchors = []
        ColSpan = 1
        BackgroundColor = None
        Shading = None
        Width = None
        PreferredWidth = None
        BorderColor = None
        BorderWidth = None
        Borders = None
        VerticalAlignment = None
        VerticalMerge = None
        PaddingTop = None
        PaddingRight = None
        PaddingBottom = None
        PaddingLeft = None
    }

let private tryFindTableCellByParagraph doc paragraphId =
    let mutable result = None
    let mutable blocks = doc.Blocks
    while result.IsNone && not blocks.IsEmpty do
        match blocks with
        | Table table :: blockTail ->
            let mutable rowIndex = 0
            let mutable rows = table.Rows
            while result.IsNone && not rows.IsEmpty do
                match rows with
                | row :: rowTail ->
                    let mutable colIndex = 0
                    let mutable cells = row.Cells
                    while result.IsNone && not cells.IsEmpty do
                        match cells with
                        | cell :: cellTail ->
                            let mutable content = cell.Content
                            while result.IsNone && not content.IsEmpty do
                                match content with
                                | para :: paraTail ->
                                    if para.Id = paragraphId then
                                        result <- Some (table.Id, rowIndex, colIndex)
                                    content <- paraTail
                                | [] -> ()
                            colIndex <- colIndex + 1
                            cells <- cellTail
                        | [] -> ()
                    rowIndex <- rowIndex + 1
                    rows <- rowTail
                | [] -> ()
            blocks <- blockTail
        | _ :: blockTail ->
            blocks <- blockTail
        | [] -> ()
    result

type private CurrentTableCellCache = {
    SourceBlocks: Block list
    Entries: Dictionary<BlockId, CurrentTableCellCacheEntry>
    EntryFastSlots: CurrentTableCellCacheEntry array
    SourceTableParagraphCount: int
    SourceParagraphCount: int
    SourceTextCharCount: int
}

and private CurrentTableCellCacheEntry = {
    ParagraphId: BlockId
    Result: struct (int * int * int)
    mutable RecentLookupIdSet: bool
    mutable RecentLookupId: BlockId
}

let mutable private currentTableCellCache: CurrentTableCellCache list = []
let private currentTableCellCacheByBlocks = Dictionary<Block list, CurrentTableCellCache>(documentInfoBlocksComparer)
let private currentTableCellFastSlots = Array.zeroCreate<CurrentTableCellCache> 256
let private currentTableCellCacheMaxSources = 256
let private currentTableCellCacheMaxTotalTableParagraphs = 200000
let private currentTableCellCacheMaxTotalSourceParagraphs = sourceParagraphCountCacheLimit
let private currentTableCellCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars
let private currentTableCellCacheBuildEntryLimit = 32768

let private currentTableCellEntryFastSlot (slots: CurrentTableCellCacheEntry array) (paragraphId: BlockId) =
    blockIdKeyHash paragraphId &&& (slots.Length - 1)

let private setCurrentTableCellEntryFastSlot (slots: CurrentTableCellCacheEntry array) (entry: CurrentTableCellCacheEntry) =
    slots.[currentTableCellEntryFastSlot slots entry.ParagraphId] <- entry

let private currentTableCellEntryMatches (entry: CurrentTableCellCacheEntry) paragraphId =
    if obj.ReferenceEquals(entry.ParagraphId, paragraphId)
       || (entry.RecentLookupIdSet && obj.ReferenceEquals(entry.RecentLookupId, paragraphId)) then
        true
    elif blockIdKeyEquals entry.ParagraphId paragraphId then
        entry.RecentLookupIdSet <- true
        entry.RecentLookupId <- paragraphId
        true
    else
        false

let private currentTableCellCacheMatches blocks (cached: CurrentTableCellCache) =
    obj.ReferenceEquals(cached.SourceBlocks, blocks)

let private setCurrentTableCellCache entries =
    currentTableCellCache <- entries
    currentTableCellCacheByBlocks.Clear()
    clearArray currentTableCellFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<CurrentTableCellCache>
        if not (currentTableCellCacheByBlocks.TryGetValue(entry.SourceBlocks, &existing)) then
            currentTableCellCacheByBlocks.[entry.SourceBlocks] <- entry
            let slot = documentInfoSlot entry.SourceBlocks currentTableCellFastSlots.Length
            currentTableCellFastSlots.[slot] <- entry

let private tryGetCurrentTableCellCache blocks =
    let slot = documentInfoSlot blocks currentTableCellFastSlots.Length
    let fast = currentTableCellFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && currentTableCellCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<CurrentTableCellCache>
        if currentTableCellCacheByBlocks.TryGetValue(blocks, &cached) then
            currentTableCellFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private currentTableCellCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<CurrentTableCellCache>
    if currentTableCellCacheByBlocks.TryGetValue(blocks, &cached)
       && currentTableCellCacheMatches blocks cached then
        currentTableCellCache
        |> List.filter (currentTableCellCacheMatches blocks >> not)
    else
        currentTableCellCache

let private trimCurrentTableCellCache caches =
    let rec loop remaining kept count totalTableParagraphs totalSourceParagraphs totalSourceTextChars =
        match remaining with
        | [] -> List.rev kept
        | cached :: rest ->
            let tableParagraphCount = max 1 cached.SourceTableParagraphCount
            let sourceParagraphCount = max 1 cached.SourceParagraphCount
            let sourceTextCharCount = max 1 cached.SourceTextCharCount
            let nextTotalTableParagraphs = totalTableParagraphs + tableParagraphCount
            let nextTotalSourceParagraphs = totalSourceParagraphs + sourceParagraphCount
            let nextTotalSourceTextChars = totalSourceTextChars + sourceTextCharCount
            if count < currentTableCellCacheMaxSources
               && nextTotalTableParagraphs <= currentTableCellCacheMaxTotalTableParagraphs
               && nextTotalSourceParagraphs <= currentTableCellCacheMaxTotalSourceParagraphs
               && nextTotalSourceTextChars <= currentTableCellCacheMaxTotalSourceTextChars then
                loop
                    rest
                    (cached :: kept)
                    (count + 1)
                    nextTotalTableParagraphs
                    nextTotalSourceParagraphs
                    nextTotalSourceTextChars
            else
                loop rest kept count totalTableParagraphs totalSourceParagraphs totalSourceTextChars
    loop caches [] 0 0 0 0

let private rememberCurrentTableCellCache cached =
    cached
    :: currentTableCellCacheWithoutBlocks cached.SourceBlocks
    |> trimCurrentTableCellCache
    |> setCurrentTableCellCache

let private countTableCellParagraphs blocks =
    let mutable count = 0
    for block in blocks do
        match block with
        | Table table ->
            count <- count + countTableParagraphs table
        | _ -> ()
    count

let private currentTableCellSourceTableParagraphCountFor entryCount blocks =
    if entryCount > currentTableCellCacheMaxTotalTableParagraphs then
        currentTableCellCacheMaxTotalTableParagraphs + 1
    else
        countTableCellParagraphs blocks

let private currentTableCellSourceParagraphCountFor sourceTableParagraphCount blocks =
    if sourceTableParagraphCount > currentTableCellCacheMaxTotalTableParagraphs then
        currentTableCellCacheMaxTotalSourceParagraphs + 1
    else
        countSourceParagraphs blocks

let private currentTableCellSourceTextCharCountFor sourceTableParagraphCount sourceParagraphCount blocks =
    if sourceTableParagraphCount > currentTableCellCacheMaxTotalTableParagraphs
       || sourceParagraphCount > currentTableCellCacheMaxTotalSourceParagraphs then
        currentTableCellCacheMaxTotalSourceTextChars + 1
    else
        documentBlocksSourceTextCharCountFor
            currentTableCellCacheMaxTotalSourceParagraphs
            sourceParagraphCount
            blocks

let private addCurrentTableCellEntry (entries: Dictionary<BlockId, CurrentTableCellCacheEntry>) tableIndex rowIndex colIndex cell =
    for paragraph in cell.Content do
        if not (entries.ContainsKey paragraph.Id) then
            entries.[paragraph.Id] <- {
                ParagraphId = paragraph.Id
                Result = struct (tableIndex, rowIndex, colIndex)
                RecentLookupIdSet = false
                RecentLookupId = ""
            }

let private buildCurrentTableCellCache (doc: Document) =
    let entries = Dictionary<BlockId, CurrentTableCellCacheEntry>(blockIdLookupComparer)
    let mutable countedTableParagraphs = 0
    let mutable countedTableParagraphsExact = true
    let sourceTableParagraphLimit = sourceParagraphCountCacheLimit
    let sourceTableVisitedItemLimit = sourceStructureVisitedItemLimit sourceTableParagraphLimit
    let mutable tableIndex = 0
    for block in doc.Blocks do
        match block with
        | Table table ->
            let mutable tableParagraphCount = 0
            let mutable tableVisitedItems = 1
            let mutable rowIndex = 0
            for row in table.Rows do
                tableVisitedItems <- tableVisitedItems + 1
                let mutable colIndex = 0
                for cell in row.Cells do
                    tableVisitedItems <- tableVisitedItems + 1
                    if not cell.NestedTables.IsEmpty then
                        countedTableParagraphsExact <- false
                    tableParagraphCount <- tableParagraphCount + cell.Content.Length
                    addCurrentTableCellEntry entries tableIndex rowIndex colIndex cell
                    colIndex <- colIndex + 1
                rowIndex <- rowIndex + 1
            let tableCount =
                if tableVisitedItems > sourceTableVisitedItemLimit
                   || sourceStructureOverwhelmsParagraphs sourceTableParagraphLimit tableParagraphCount tableVisitedItems then
                    sourceTableParagraphLimit + 1
                else
                    min tableParagraphCount (sourceTableParagraphLimit + 1)
            countedTableParagraphs <- countedTableParagraphs + tableCount
        | _ -> ()
        tableIndex <- tableIndex + 1
    let sourceTableParagraphCount =
        if entries.Count > currentTableCellCacheMaxTotalTableParagraphs then
            currentTableCellCacheMaxTotalTableParagraphs + 1
        elif countedTableParagraphsExact then
            countedTableParagraphs
        else
            currentTableCellSourceTableParagraphCountFor entries.Count doc.Blocks
    let sourceParagraphCount = currentTableCellSourceParagraphCountFor sourceTableParagraphCount doc.Blocks
    let entryFastSlots = Array.zeroCreate<CurrentTableCellCacheEntry> (lookupEntryFastSlotSize entries.Count)
    for entry in entries.Values do
        setCurrentTableCellEntryFastSlot entryFastSlots entry
    {
        SourceBlocks = doc.Blocks
        Entries = entries
        EntryFastSlots = entryFastSlots
        SourceTableParagraphCount = sourceTableParagraphCount
        SourceParagraphCount = sourceParagraphCount
        SourceTextCharCount = currentTableCellSourceTextCharCountFor sourceTableParagraphCount sourceParagraphCount doc.Blocks
    }

let private currentTableCellLookup (doc: Document) =
    match tryGetCurrentTableCellCache doc.Blocks with
    | ValueSome cached ->
        cached
    | ValueNone ->
        let cached = buildCurrentTableCellCache doc
        rememberCurrentTableCellCache cached
        cached

let private tryFindCurrentTableCellEntry (cached: CurrentTableCellCache) blockId =
    let fast = cached.EntryFastSlots.[currentTableCellEntryFastSlot cached.EntryFastSlots blockId]
    if not (obj.ReferenceEquals(fast, null))
       && currentTableCellEntryMatches fast blockId then
        ValueSome fast.Result
    else
        let mutable entry = Unchecked.defaultof<CurrentTableCellCacheEntry>
        if cached.Entries.TryGetValue(blockId, &entry) then
            entry.RecentLookupIdSet <- true
            entry.RecentLookupId <- blockId
            setCurrentTableCellEntryFastSlot cached.EntryFastSlots entry
            ValueSome entry.Result
        else
            ValueNone

let private tryFindCachedCurrentTableCell blocks blockId =
    match tryGetCurrentTableCellCache blocks with
    | ValueSome cached -> tryFindCurrentTableCellEntry cached blockId
    | ValueNone -> ValueNone

let private countCurrentTableCellEntriesWithLimit limit (table: TableBlock) =
    let mutable count = 0
    let mutable rows = table.Rows
    while count <= limit && not rows.IsEmpty do
        let mutable cells = rows.Head.Cells
        while count <= limit && not cells.IsEmpty do
            let mutable paragraphs = cells.Head.Content
            while count <= limit && not paragraphs.IsEmpty do
                count <- count + 1
                paragraphs <- paragraphs.Tail
            cells <- cells.Tail
        rows <- rows.Tail
    count

let private shouldBuildCurrentTableCellCacheForSingleTable table =
    let count = countCurrentTableCellEntriesWithLimit (currentTableCellCacheBuildEntryLimit + 1) table
    count > 0 && count <= currentTableCellCacheBuildEntryLimit

let private countCurrentTableCellEntriesInBlocksWithLimit limit (blocks: Block list) =
    let mutable count = 0
    let mutable remaining = blocks
    while count <= limit && not remaining.IsEmpty do
        match remaining.Head with
        | Table table ->
            count <- count + countCurrentTableCellEntriesWithLimit (limit - count) table
        | _ -> ()
        remaining <- remaining.Tail
    count

let private shouldBuildCurrentTableCellCacheForBlocks blocks =
    let count = countCurrentTableCellEntriesInBlocksWithLimit (currentTableCellCacheBuildEntryLimit + 1) blocks
    count > 0 && count <= currentTableCellCacheBuildEntryLimit

let private tryFindCurrentTableCellUntil doc blockId =
    match tryFindCachedCurrentTableCell doc.Blocks blockId with
    | ValueSome _ as cached -> cached
    | ValueNone ->
        let mutable result = ValueNone
        let mutable blocks = doc.Blocks
        let mutable tableIndex = 0
        while result.IsNone && not blocks.IsEmpty do
            match blocks with
            | Table table :: tail ->
                let mutable rowIndex = 0
                let mutable rows = table.Rows
                while result.IsNone && not rows.IsEmpty do
                    match rows with
                    | row :: rowTail ->
                        let mutable colIndex = 0
                        let mutable cells = row.Cells
                        while result.IsNone && not cells.IsEmpty do
                            match cells with
                            | cell :: cellTail ->
                                if cell.Content |> List.exists (fun para -> para.Id = blockId) then
                                    result <- ValueSome (struct (tableIndex, rowIndex, colIndex))
                                colIndex <- colIndex + 1
                                cells <- cellTail
                            | [] -> ()
                        rowIndex <- rowIndex + 1
                        rows <- rowTail
                    | [] -> ()
                blocks <- tail
                tableIndex <- tableIndex + 1
            | _ :: tail ->
                blocks <- tail
                tableIndex <- tableIndex + 1
            | [] -> ()
        result

let private tryFindTableAtTopLevelIndex tableIndex (blocks: Block list) =
    let mutable result = ValueNone
    if tableIndex >= 0 then
        let mutable index = 0
        let mutable remaining = blocks
        while result.IsNone && index <= tableIndex && not remaining.IsEmpty do
            match remaining with
            | Table table :: _ when index = tableIndex ->
                result <- ValueSome table
            | _ :: tail ->
                index <- index + 1
                remaining <- tail
            | [] -> ()
    result

let private updateTable (doc: Document) tableId update =
    match tryFindTopBlockWithIndexUntil doc tableId with
    | ValueSome (struct (Table table, tableIndex)) ->
        let nextTable = update table
        if obj.ReferenceEquals(nextTable, table) || nextTable = table then
            doc
        else
            { doc with Blocks = replaceTopLevelBlockAtIndex tableIndex (Table nextTable) doc.Blocks }
    | _ ->
        doc

let private updateTableAtIndex tableIndex update (doc: Document) =
    match tryFindTableAtTopLevelIndex tableIndex doc.Blocks with
    | ValueSome table ->
        let nextTable = update table
        if obj.ReferenceEquals(nextTable, table) || nextTable = table then
            doc
        else
            { doc with Blocks = replaceTopLevelBlockAtIndex tableIndex (Table nextTable) doc.Blocks }
    | ValueNone ->
        doc

let private updateTableCell rowIndex colIndex updateCell (table: TableBlock) =
    let rowPrefix = ResizeArray<TableRow>(max 0 rowIndex)
    let rec updateRowAt rowPosition rows =
        match rows with
        | [] -> None
        | row :: tail when rowPosition = rowIndex ->
            let cellPrefix = ResizeArray<TableCell>(max 0 colIndex)
            let rec updateCellAt cellPosition cells =
                match cells with
                | [] -> None
                | cell :: cellTail when cellPosition = colIndex ->
                    let nextCell = updateCell cell
                    if nextCell = cell then
                        None
                    else
                        Some (
                            prependResizeArray
                                rowPrefix
                                ({ row with Cells = prependResizeArray cellPrefix (nextCell :: cellTail) } :: tail))
                | cell :: cellTail ->
                    cellPrefix.Add cell
                    updateCellAt (cellPosition + 1) cellTail
            updateCellAt 0 row.Cells
        | row :: tail ->
            rowPrefix.Add row
            updateRowAt (rowPosition + 1) tail
    match updateRowAt 0 table.Rows with
    | Some rows -> { table with Rows = rows }
    | None -> table

let private currentTableCell doc =
    match doc.Selection with
    | Some sel ->
        let cachedResult = tryFindCachedCurrentTableCell doc.Blocks sel.Focus.BlockId
        match cachedResult with
        | ValueSome _ -> cachedResult
        | ValueNone ->
            let shouldBuild =
                match doc.Blocks with
                | [ Table table ] -> shouldBuildCurrentTableCellCacheForSingleTable table
                | _ -> shouldBuildCurrentTableCellCacheForBlocks doc.Blocks
            if shouldBuild then
                let cached = currentTableCellLookup doc
                tryFindCurrentTableCellEntry cached sel.Focus.BlockId
            else
                tryFindCurrentTableCellUntil doc sel.Focus.BlockId
    | None -> ValueNone

let private tryGetTableCellParagraphInRow colIndex (row: TableRow) =
    row.Cells
    |> List.tryItem colIndex
    |> Option.bind (fun cell -> cell.Content |> List.tryHead)

/// Find block index by ID
let findBlockIndex doc blockId =
    let lookup = blockLookup doc.Blocks
    match tryFindBlockLookupEntry lookup blockId with
    | ValueSome entry -> entry.TopIndexResult
    | ValueNone -> None

[<Struct>]
type private NormalizedRangeCacheKey = {
    AnchorBlockId: BlockId
    AnchorOffset: int
    FocusBlockId: BlockId
    FocusOffset: int
}

[<Struct>]
type private NormalizedRangeCacheEntry = {
    BlockIdCount: int
    Result: Position * Position * BlockId list
}

type private NormalizedRangeFastEntry = {
    Key: NormalizedRangeCacheKey
    Entry: NormalizedRangeCacheEntry
}

type private NormalizedRangeCache = {
    SourceBlocks: Block list
    SourceParagraphCount: int
    SourceTextCharCount: int
    Entries: Dictionary<NormalizedRangeCacheKey, NormalizedRangeCacheEntry>
    Order: Queue<NormalizedRangeCacheKey>
    EntryFastSlots: NormalizedRangeFastEntry array
    mutable TotalBlockIds: int
}

let mutable private normalizedRangeCache: NormalizedRangeCache list = []
let private normalizedRangeCacheByBlocks = Dictionary<Block list, NormalizedRangeCache>(documentInfoBlocksComparer)
let private normalizedRangeFastSlots = Array.zeroCreate<NormalizedRangeCache> 256
let private normalizedRangeCacheMaxEntries = 1024
let private normalizedRangeCacheMaxSources = 256
let private normalizedRangeCacheMaxTotalBlockIds = 200000
let private normalizedRangeCacheMaxTotalSourceParagraphs = sourceParagraphCountCacheLimit
let private normalizedRangeCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars

let private normalizedRangeKeyHash anchorBlockId anchorOffset focusBlockId focusOffset =
    let mutable hashCode = blockIdKeyHash anchorBlockId
    hashCode <- (hashCode * 397) ^^^ anchorOffset
    hashCode <- (hashCode * 397) ^^^ blockIdKeyHash focusBlockId
    hashCode <- (hashCode * 397) ^^^ focusOffset
    hashCode ^^^ (hashCode >>> 16)

let private normalizedRangeCacheKeyComparer =
    { new IEqualityComparer<NormalizedRangeCacheKey> with
        member _.Equals(left, right) =
            blockIdKeyEquals left.AnchorBlockId right.AnchorBlockId
            && left.AnchorOffset = right.AnchorOffset
            && blockIdKeyEquals left.FocusBlockId right.FocusBlockId
            && left.FocusOffset = right.FocusOffset

        member _.GetHashCode(key) =
            normalizedRangeKeyHash
                key.AnchorBlockId
                key.AnchorOffset
                key.FocusBlockId
                key.FocusOffset }

let private normalizedRangeEntryFastSlot
    (slots: NormalizedRangeFastEntry array)
    anchorBlockId
    anchorOffset
    focusBlockId
    focusOffset =
    normalizedRangeKeyHash anchorBlockId anchorOffset focusBlockId focusOffset &&& (slots.Length - 1)

let private clearNormalizedRangeEntryFastSlots (slots: NormalizedRangeFastEntry array) =
    for index = 0 to slots.Length - 1 do
        slots.[index] <- Unchecked.defaultof<NormalizedRangeFastEntry>

let private normalizedRangeKeyMatchesSelection (sel: Selection) (key: NormalizedRangeCacheKey) =
    blockIdKeyEquals key.AnchorBlockId sel.Anchor.BlockId
    && key.AnchorOffset = sel.Anchor.Offset
    && blockIdKeyEquals key.FocusBlockId sel.Focus.BlockId
    && key.FocusOffset = sel.Focus.Offset

let private tryGetNormalizedRangeEntryFastSlot (cache: NormalizedRangeCache) (sel: Selection) =
    let slot =
        normalizedRangeEntryFastSlot
            cache.EntryFastSlots
            sel.Anchor.BlockId
            sel.Anchor.Offset
            sel.Focus.BlockId
            sel.Focus.Offset
    let fast = cache.EntryFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && normalizedRangeKeyMatchesSelection sel fast.Key then
        ValueSome fast.Entry
    else
        ValueNone

let private setNormalizedRangeEntryFastSlot (cache: NormalizedRangeCache) key entry =
    let slot =
        normalizedRangeEntryFastSlot
            cache.EntryFastSlots
            key.AnchorBlockId
            key.AnchorOffset
            key.FocusBlockId
            key.FocusOffset
    cache.EntryFastSlots.[slot] <- {
        Key = key
        Entry = entry
    }

let private normalizedRangeCacheMatches blocks (cached: NormalizedRangeCache) =
    obj.ReferenceEquals(cached.SourceBlocks, blocks)

let private normalizedRangeCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<NormalizedRangeCache>
    if normalizedRangeCacheByBlocks.TryGetValue(blocks, &cached)
       && normalizedRangeCacheMatches blocks cached then
        normalizedRangeCache
        |> List.filter (normalizedRangeCacheMatches blocks >> not)
    else
        normalizedRangeCache

let private setNormalizedRangeCache entries =
    normalizedRangeCache <- entries
    normalizedRangeCacheByBlocks.Clear()
    clearArray normalizedRangeFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<NormalizedRangeCache>
        if not (normalizedRangeCacheByBlocks.TryGetValue(entry.SourceBlocks, &existing)) then
            normalizedRangeCacheByBlocks.[entry.SourceBlocks] <- entry
            let slot = documentInfoSlot entry.SourceBlocks normalizedRangeFastSlots.Length
            normalizedRangeFastSlots.[slot] <- entry

let private tryGetNormalizedRangeCache blocks =
    let slot = documentInfoSlot blocks normalizedRangeFastSlots.Length
    let fast = normalizedRangeFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && normalizedRangeCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<NormalizedRangeCache>
        if normalizedRangeCacheByBlocks.TryGetValue(blocks, &cached) then
            normalizedRangeFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private removeNormalizedRangeCacheForBlocks blocks =
    let entries = normalizedRangeCacheWithoutBlocks blocks
    if not (obj.ReferenceEquals(entries, normalizedRangeCache)) then
        setNormalizedRangeCache entries

let private normalizedRangeCacheRetainable (cached: NormalizedRangeCache) =
    cached.Entries.Count > 0
    && cached.Entries.Count <= normalizedRangeCacheMaxEntries
    && cached.TotalBlockIds <= normalizedRangeCacheMaxTotalBlockIds
    && cached.SourceParagraphCount <= normalizedRangeCacheMaxTotalSourceParagraphs
    && cached.SourceTextCharCount <= normalizedRangeCacheMaxTotalSourceTextChars

let private rememberNormalizedRangeCache cached =
    if not (normalizedRangeCacheRetainable cached) then
        removeNormalizedRangeCacheForBlocks cached.SourceBlocks
    else
        let mutable keptCount = 1
        let mutable keptBlockIds = max 1 cached.TotalBlockIds
        let mutable keptBlocks = max 1 cached.SourceParagraphCount
        let mutable keptSourceTextChars = max 1 cached.SourceTextCharCount
        let retained =
            normalizedRangeCacheWithoutBlocks cached.SourceBlocks
            |> List.choose (fun candidate ->
                let candidateBlockIds = max 1 candidate.TotalBlockIds
                let candidateBlocks = max 1 candidate.SourceParagraphCount
                let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                if normalizedRangeCacheRetainable candidate
                   && keptCount < normalizedRangeCacheMaxSources
                   && keptBlockIds + candidateBlockIds <= normalizedRangeCacheMaxTotalBlockIds
                   && keptBlocks + candidateBlocks <= normalizedRangeCacheMaxTotalSourceParagraphs
                   && keptSourceTextChars + candidateSourceTextChars <= normalizedRangeCacheMaxTotalSourceTextChars then
                    keptCount <- keptCount + 1
                    keptBlockIds <- keptBlockIds + candidateBlockIds
                    keptBlocks <- keptBlocks + candidateBlocks
                    keptSourceTextChars <- keptSourceTextChars + candidateSourceTextChars
                    Some candidate
                else
                    None)
        cached :: retained
        |> setNormalizedRangeCache

let private normalizedRangeKey (sel: Selection) = {
    AnchorBlockId = sel.Anchor.BlockId
    AnchorOffset = sel.Anchor.Offset
    FocusBlockId = sel.Focus.BlockId
    FocusOffset = sel.Focus.Offset
}

let private normalizedRangeSourceParagraphCountFor blockIdCount blocks =
    if blockIdCount > normalizedRangeCacheMaxTotalBlockIds then
        normalizedRangeCacheMaxTotalSourceParagraphs + 1
    else
        countSourceParagraphs blocks

let private normalizedRangeSourceTextCharCountFor blockIdCount sourceParagraphCount blocks =
    if blockIdCount > normalizedRangeCacheMaxTotalBlockIds
       || sourceParagraphCount > normalizedRangeCacheMaxTotalSourceParagraphs then
        normalizedRangeCacheMaxTotalSourceTextChars + 1
    else
        documentBlocksSourceTextCharCountFor
            normalizedRangeCacheMaxTotalSourceParagraphs
            sourceParagraphCount
            blocks

let private removeNormalizedRangeEntry key (cache: NormalizedRangeCache) =
    let mutable entry = Unchecked.defaultof<NormalizedRangeCacheEntry>
    if cache.Entries.TryGetValue(key, &entry) then
        clearNormalizedRangeEntryFastSlots cache.EntryFastSlots
        cache.TotalBlockIds <- cache.TotalBlockIds - entry.BlockIdCount
        cache.Entries.Remove key |> ignore

let private addNormalizedRangeEntry key entry (cache: NormalizedRangeCache) =
    if entry.BlockIdCount > normalizedRangeCacheMaxTotalBlockIds then
        removeNormalizedRangeEntry key cache
    else
        let mutable removedEntry = false
        if not (cache.Entries.ContainsKey key) then
            cache.Order.Enqueue key
            cache.TotalBlockIds <- cache.TotalBlockIds + entry.BlockIdCount
        else
            let mutable oldEntry = Unchecked.defaultof<NormalizedRangeCacheEntry>
            if cache.Entries.TryGetValue(key, &oldEntry) then
                cache.TotalBlockIds <- cache.TotalBlockIds - oldEntry.BlockIdCount + entry.BlockIdCount
        cache.Entries.[key] <- entry
        while (cache.Entries.Count > normalizedRangeCacheMaxEntries
               || cache.TotalBlockIds > normalizedRangeCacheMaxTotalBlockIds)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> key then
                removedEntry <- true
                removeNormalizedRangeEntry oldest cache
        if cache.Entries.Count > normalizedRangeCacheMaxEntries
           || cache.TotalBlockIds > normalizedRangeCacheMaxTotalBlockIds then
            let current = cache.Entries.[key]
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue key
            cache.Entries.[key] <- current
            cache.TotalBlockIds <- current.BlockIdCount
            clearNormalizedRangeEntryFastSlots cache.EntryFastSlots
        elif removedEntry then
            clearNormalizedRangeEntryFastSlots cache.EntryFastSlots
        setNormalizedRangeEntryFastSlot cache key entry

let private cachedNormalizedRange doc sel =
    match tryGetNormalizedRangeCache doc.Blocks with
    | ValueSome cached when normalizedRangeCacheRetainable cached ->
        match tryGetNormalizedRangeEntryFastSlot cached sel with
        | ValueSome entry ->
            ValueSome entry.Result
        | ValueNone ->
            let key = normalizedRangeKey sel
            let mutable entry = Unchecked.defaultof<NormalizedRangeCacheEntry>
            if cached.Entries.TryGetValue(key, &entry) then
                if normalizedRangeKeyMatchesSelection sel key then
                    setNormalizedRangeEntryFastSlot cached key entry
                ValueSome entry.Result
            else
                ValueNone
    | ValueSome cached ->
        rememberNormalizedRangeCache cached
        ValueNone
    | ValueNone -> ValueNone

let private rememberNormalizedRange
    (doc: Document)
    sel
    (startPos: Position)
    (endPos: Position)
    (blockIds: BlockId list)
    =
    let key = normalizedRangeKey sel
    let result: Position * Position * BlockId list = (startPos, endPos, blockIds)
    let entry: NormalizedRangeCacheEntry = {
        BlockIdCount = blockIds.Length
        Result = result
    }
    match tryGetNormalizedRangeCache doc.Blocks with
    | ValueSome cached when normalizedRangeCacheRetainable cached ->
        addNormalizedRangeEntry key entry cached
        if not (normalizedRangeCacheRetainable cached) then
            removeNormalizedRangeCacheForBlocks cached.SourceBlocks
    | ValueSome cached ->
        removeNormalizedRangeCacheForBlocks cached.SourceBlocks
    | ValueNone ->
        let sourceParagraphCount = normalizedRangeSourceParagraphCountFor entry.BlockIdCount doc.Blocks
        let cache = {
            SourceBlocks = doc.Blocks
            SourceParagraphCount = sourceParagraphCount
            SourceTextCharCount = normalizedRangeSourceTextCharCountFor entry.BlockIdCount sourceParagraphCount doc.Blocks
            Entries = Dictionary<NormalizedRangeCacheKey, NormalizedRangeCacheEntry>(normalizedRangeCacheKeyComparer)
            Order = Queue<NormalizedRangeCacheKey>()
            EntryFastSlots = Array.zeroCreate<NormalizedRangeFastEntry> 2048
            TotalBlockIds = 0
        }
        addNormalizedRangeEntry key entry cache
        rememberNormalizedRangeCache cache

[<Struct>]
type private NormalizedRangeExactCacheKey = {
    SourceBlocks: Block list
    AnchorBlockId: BlockId
    AnchorOffset: int
    FocusBlockId: BlockId
    FocusOffset: int
}

type private NormalizedRangeExactCacheEntry = {
    Key: NormalizedRangeExactCacheKey
    Result: Position * Position * BlockId list
}

let private normalizedRangeExactCacheMaxEntries = 4096

let private normalizedRangeExactKeyHash (key: NormalizedRangeExactCacheKey) =
    let mutable hashCode = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.SourceBlocks)
    hashCode <- (hashCode * 397) ^^^ System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.AnchorBlockId)
    hashCode <- (hashCode * 397) ^^^ key.AnchorOffset
    hashCode <- (hashCode * 397) ^^^ System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.FocusBlockId)
    hashCode <- (hashCode * 397) ^^^ key.FocusOffset
    hashCode <- hashCode ^^^ (hashCode >>> 16)
    hashCode

let private normalizedRangeExactKeyComparer =
    { new IEqualityComparer<NormalizedRangeExactCacheKey> with
        member _.Equals(left, right) =
            obj.ReferenceEquals(left.SourceBlocks, right.SourceBlocks)
            && obj.ReferenceEquals(left.AnchorBlockId, right.AnchorBlockId)
            && left.AnchorOffset = right.AnchorOffset
            && obj.ReferenceEquals(left.FocusBlockId, right.FocusBlockId)
            && left.FocusOffset = right.FocusOffset

        member _.GetHashCode(key) = normalizedRangeExactKeyHash key }

let private normalizedRangeExactCache =
    Dictionary<NormalizedRangeExactCacheKey, NormalizedRangeExactCacheEntry>(
        normalizedRangeExactCacheMaxEntries,
        normalizedRangeExactKeyComparer)
let private normalizedRangeExactOrder = Queue<NormalizedRangeExactCacheKey>(normalizedRangeExactCacheMaxEntries)

let private normalizedRangeExactKey (doc: Document) (sel: Selection) = {
    SourceBlocks = doc.Blocks
    AnchorBlockId = sel.Anchor.BlockId
    AnchorOffset = sel.Anchor.Offset
    FocusBlockId = sel.Focus.BlockId
    FocusOffset = sel.Focus.Offset
}

let private rememberNormalizedRangeExact doc sel result =
    let key = normalizedRangeExactKey doc sel
    let mutable existing = Unchecked.defaultof<NormalizedRangeExactCacheEntry>
    if not (normalizedRangeExactCache.TryGetValue(key, &existing)) then
        normalizedRangeExactOrder.Enqueue key
    normalizedRangeExactCache.[key] <- {
        Key = key
        Result = result
    }
    while normalizedRangeExactCache.Count > normalizedRangeExactCacheMaxEntries
          && normalizedRangeExactOrder.Count > 0 do
        let oldest = normalizedRangeExactOrder.Dequeue()
        normalizedRangeExactCache.Remove oldest |> ignore

let private tryGetNormalizedRangeExact doc sel =
    let key = normalizedRangeExactKey doc sel
    let mutable entry = Unchecked.defaultof<NormalizedRangeExactCacheEntry>
    if normalizedRangeExactCache.TryGetValue(key, &entry) then
        ValueSome entry.Result
    else
        ValueNone

/// Get normalized range with block IDs for multi-block selection
let getNormalizedRange (sel: Selection) (doc: Document) =
    match tryGetNormalizedRangeExact doc sel with
    | ValueSome cached -> cached
    | ValueNone ->
        match cachedNormalizedRange doc sel with
        | ValueSome cached ->
            rememberNormalizedRangeExact doc sel cached
            cached
        | ValueNone ->
            let result =
                let ordered = orderedEditableParagraphs doc
                let mutable anchorIndex = 0
                let mutable focusIndex = 0
                if ordered.OrderedIndexById.TryGetValue(sel.Anchor.BlockId, &anchorIndex)
                   && ordered.OrderedIndexById.TryGetValue(sel.Focus.BlockId, &focusIndex) then
                    if sel.Anchor.BlockId = sel.Focus.BlockId then
                        if sel.Anchor.Offset <= sel.Focus.Offset then
                            sel.Anchor, sel.Focus, [ sel.Anchor.BlockId ]
                        else
                            sel.Focus, sel.Anchor, [ sel.Anchor.BlockId ]
                    else
                        let anchorFirst = anchorIndex <= focusIndex
                        let startPos, endPos, startIndex, endIndex =
                            if anchorFirst then
                                sel.Anchor, sel.Focus, anchorIndex, focusIndex
                            else
                                sel.Focus, sel.Anchor, focusIndex, anchorIndex
                        let blockIds = ResizeArray<BlockId>(endIndex - startIndex + 1)
                        for i = startIndex to endIndex do
                            blockIds.Add ordered.OrderedParagraphs.[i].Id
                        startPos, endPos, resizeArrayToList blockIds
                else
                    let blockIds = ResizeArray<BlockId>()
                    let mutable collecting = false
                    let mutable foundBoth = false
                    let mutable foundAnchor = false
                    let mutable foundFocus = false
                    let mutable anchorFirst = false

                    iterEditableParagraphsUntil doc (fun para ->
                        let isAnchor = para.Id = sel.Anchor.BlockId
                        let isFocus = para.Id = sel.Focus.BlockId
                        if isAnchor || isFocus then
                            if not collecting then
                                collecting <- true
                                anchorFirst <- isAnchor
                            blockIds.Add para.Id
                            foundAnchor <- foundAnchor || isAnchor
                            foundFocus <- foundFocus || isFocus
                            foundBoth <- foundAnchor && foundFocus
                        elif collecting then
                            blockIds.Add para.Id
                        not foundBoth)

                    if foundAnchor && foundFocus then
                        let startPos, endPos =
                            if sel.Anchor.BlockId = sel.Focus.BlockId then
                                if sel.Anchor.Offset <= sel.Focus.Offset then sel.Anchor, sel.Focus
                                else sel.Focus, sel.Anchor
                            elif anchorFirst then
                                sel.Anchor, sel.Focus
                            else
                                sel.Focus, sel.Anchor
                        (startPos, endPos, resizeArrayToList blockIds)
                    else
                        // Fallback if blocks not found
                        (sel.Anchor, sel.Focus, [])
            let startPos, endPos, blockIds = result
            rememberNormalizedRange doc sel startPos endPos blockIds
            rememberNormalizedRangeExact doc sel result
            result

let private selectedParagraphIds doc includeCollapsed =
    match doc.Selection with
    | None -> []
    | Some sel when sel.IsCollapsed ->
        if includeCollapsed then [sel.Focus.BlockId] else []
    | Some sel ->
        let _, _, ids = getNormalizedRange sel doc
        ids

let private tryUpdateTopLevelSelectedParagraphs
    (doc: Document)
    (sel: Selection)
    (updateParagraph: ParagraphBlock -> ParagraphBlock) =
    if sel.Anchor.BlockId = sel.Focus.BlockId then
        ValueNone
    else
        let blocks = ResizeArray<Block>()
        let mutable changed = false
        let mutable remaining = doc.Blocks
        let mutable collecting = false
        let mutable valid = true
        let mutable foundEnd = false
        let mutable endTarget = ""

        let addUpdatedBlock
            (block: Block)
            (para: ParagraphBlock)
            (wrap: ParagraphBlock -> Block) =
            let nextPara = updateParagraph para
            if nextPara = para then
                blocks.Add block
            else
                changed <- true
                blocks.Add(wrap nextPara)

        let startCollecting (block: Block) (para: ParagraphBlock) (wrap: ParagraphBlock -> Block) =
            let anchorMatch = para.Id = sel.Anchor.BlockId
            let focusMatch = para.Id = sel.Focus.BlockId
            if anchorMatch || focusMatch then
                collecting <- true
                endTarget <- if anchorMatch then sel.Focus.BlockId else sel.Anchor.BlockId
                addUpdatedBlock block para wrap
            else
                blocks.Add block

        let continueCollecting (block: Block) (para: ParagraphBlock) (wrap: ParagraphBlock -> Block) =
            let isLast = para.Id = endTarget
            addUpdatedBlock block para wrap
            if isLast then
                foundEnd <- true

        while valid && not foundEnd && not remaining.IsEmpty do
            match remaining with
            | block :: tail ->
                if collecting then
                    match block with
                    | Paragraph para ->
                        continueCollecting block para Paragraph
                    | Heading (level, para) ->
                        continueCollecting block para (fun para -> Heading (level, para))
                    | Table _
                    | TextBox _ ->
                        valid <- false
                else
                    match block with
                    | Paragraph para ->
                        startCollecting block para Paragraph
                    | Heading (level, para) ->
                        startCollecting block para (fun para -> Heading (level, para))
                    | Table _
                    | TextBox _ ->
                        blocks.Add block
                remaining <- tail
            | [] -> ()

        if valid && foundEnd then
            if changed then
                ValueSome { doc with Blocks = prependResizeArray blocks remaining }
            else
                ValueSome doc
        else
            ValueNone

let private tryUpdateTopLevelParagraphById
    (doc: Document)
    paragraphId
    (updateParagraph: ParagraphBlock -> ParagraphBlock) =
    let prefix = ResizeArray<Block>()
    let mutable remaining = doc.Blocks
    let mutable result: Document voption = ValueNone

    while result.IsNone && not remaining.IsEmpty do
        match remaining with
        | block :: tail ->
            match block with
            | Paragraph para when para.Id = paragraphId ->
                let nextPara = updateParagraph para
                result <-
                    if nextPara = para then
                        ValueSome doc
                    else
                        ValueSome { doc with Blocks = prependResizeArray prefix (Paragraph nextPara :: tail) }
            | Heading (level, para) when para.Id = paragraphId ->
                let nextPara = updateParagraph para
                result <-
                    if nextPara = para then
                        ValueSome doc
                    else
                        ValueSome { doc with Blocks = prependResizeArray prefix (Heading (level, nextPara) :: tail) }
            | _ ->
                prefix.Add block
                remaining <- tail
        | [] -> ()

    result

let private updateSelectedParagraphs includeCollapsed updateParagraph doc =
    let genericUpdate () =
        let blockIds = selectedParagraphIds doc includeCollapsed
        if blockIds.IsEmpty then
            doc
        else
            let selected = HashSet<BlockId>()
            for blockId in blockIds do
                selected.Add blockId |> ignore
            let replacements = Dictionary<BlockId, ParagraphBlock>()
            let processParagraph (para: ParagraphBlock) =
                if selected.Contains para.Id then
                    let nextPara = updateParagraph para
                    if nextPara <> para then
                        replacements.[para.Id] <- nextPara
                    true
                else
                    false
            if selected.Count = blockIds.Length then
                let mutable seen = 0
                iterEditableParagraphsUntil doc (fun para ->
                    if processParagraph para then
                        seen <- seen + 1
                    seen < selected.Count)
            else
                iterEditableParagraphs doc (fun para -> processParagraph para |> ignore)
            replaceParagraphs doc replacements

    match doc.Selection with
    | Some sel when sel.Anchor.BlockId = sel.Focus.BlockId && (includeCollapsed || not sel.IsCollapsed) ->
        match tryUpdateTopLevelParagraphById doc sel.Focus.BlockId updateParagraph with
        | ValueSome doc -> doc
        | ValueNone -> genericUpdate ()
    | Some sel when not sel.IsCollapsed && sel.Anchor.BlockId <> sel.Focus.BlockId ->
        match tryUpdateTopLevelSelectedParagraphs doc sel updateParagraph with
        | ValueSome doc -> doc
        | ValueNone -> genericUpdate ()
    | _ -> genericUpdate ()

let private updateSelectedParagraphMetadata includeCollapsed updateMetadata doc =
    updateSelectedParagraphs includeCollapsed (fun para ->
        let nextMetadata = updateMetadata para.Metadata
        if nextMetadata = para.Metadata then
            para
        else
            { para with Metadata = nextMetadata }) doc

/// Split block at cursor, creating two blocks
let splitBlock doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match tryFindTopBlockWithIndexUntil doc sel.Focus.BlockId with
        | ValueNone -> doc
        | ValueSome (struct (block, idx)) ->
            let para = getParagraph block
            let offset = sel.Focus.Offset

            // Split buffer at offset while preserving run metadata.
            let beforeBuffer = TextBuffer.sliceBuffer 0 offset para.Buffer
            let afterBuffer = TextBuffer.sliceBuffer offset (TextBuffer.length para.Buffer - offset) para.Buffer

            // Create new block for the second part (carry over metadata)
            let newBlockId = nextBlockId()
            let newPara = { Id = newBlockId; Buffer = afterBuffer; Metadata = para.Metadata }
            let newBlock =
                match block with
                | Paragraph _ -> Paragraph newPara
                | Heading (level, _) -> Heading (level, newPara)
                | Table t -> Table t
                | TextBox t -> TextBox t

            // Update first block with before text (keep existing metadata)
            let updatedPara = { para with Buffer = beforeBuffer }
            let updatedBlock =
                match block with
                | Paragraph _ -> Paragraph updatedPara
                | Heading (level, _) -> Heading (level, updatedPara)
                | Table t -> Table t
                | TextBox t -> TextBox t

            // Move cursor to start of new block
            let newSel = {
                Anchor = { BlockId = newBlockId; Offset = 0 }
                Focus = { BlockId = newBlockId; Offset = 0 }
            }
            { doc with Blocks = replaceBlockAndInsertAfterIndex idx updatedBlock newBlock doc.Blocks; Selection = Some newSel }

/// Merge block with previous block (for Backspace at start)
let mergeWithPrevious (doc: Document) blockId =
    let prefix = ResizeArray<Block>()
    let rec loop hasPrevious previous remaining =
        match remaining with
        | [] -> doc
        | block :: tail when getBlockId block = blockId ->
            if not hasPrevious then
                doc
            else
                match tryGetParagraph previous, tryGetParagraph block with
                | Some prevPara, Some currentPara ->
                    let prevLen = TextBuffer.length prevPara.Buffer
                    let mergedBuffer = TextBuffer.append prevPara.Buffer currentPara.Buffer
                    let mergedPara = { prevPara with Buffer = mergedBuffer }
                    let mergedBlock =
                        match previous with
                        | Paragraph _ -> Paragraph mergedPara
                        | Heading (level, _) -> Heading (level, mergedPara)
                        | Table t -> Table t
                        | TextBox t -> TextBox t
                    let selection =
                        Some {
                            Anchor = { BlockId = prevPara.Id; Offset = prevLen }
                            Focus = { BlockId = prevPara.Id; Offset = prevLen }
                        }
                    { doc with Blocks = prependResizeArray prefix (mergedBlock :: tail); Selection = selection }
                | _ -> doc
        | block :: tail ->
            if hasPrevious then
                prefix.Add previous
            loop true block tail
    loop false Unchecked.defaultof<Block> doc.Blocks

/// Merge block with next block (for Delete at end)
let mergeWithNext (doc: Document) blockId =
    let prefix = ResizeArray<Block>()
    let rec loop remaining =
        match remaining with
        | block :: nextBlock :: tail when getBlockId block = blockId ->
            match tryGetParagraph block, tryGetParagraph nextBlock with
            | Some currentPara, Some nextPara ->
                let mergedBuffer = TextBuffer.append currentPara.Buffer nextPara.Buffer
                let mergedPara = { currentPara with Buffer = mergedBuffer }
                let mergedBlock =
                    match block with
                    | Paragraph _ -> Paragraph mergedPara
                    | Heading (level, _) -> Heading (level, mergedPara)
                    | Table t -> Table t
                    | TextBox t -> TextBox t
                { doc with Blocks = prependResizeArray prefix (mergedBlock :: tail) }
            | _ -> doc
        | block :: tail ->
            prefix.Add block
            loop tail
        | [] -> doc
    loop doc.Blocks

let private tryNormalizeTopLevelSelectionRangeUntil (doc: Document) (sel: Selection) =
    if sel.Anchor.BlockId = sel.Focus.BlockId then
        None
    else
        let topLevelParagraphId = function
            | Paragraph p
            | Heading (_, p) -> Some p.Id
            | Table _
            | TextBox _ -> None
        let ids = ResizeArray<BlockId>()
        let mutable collecting = false
        let mutable validTopLevelRange = true
        let mutable foundAnchor = false
        let mutable foundFocus = false
        let mutable anchorFirst = false
        let mutable firstBlock = Unchecked.defaultof<Block>
        let mutable lastBlock = Unchecked.defaultof<Block>
        let mutable hasFirstBlock = false
        let mutable remaining = doc.Blocks
        while validTopLevelRange && not (foundAnchor && foundFocus) && not remaining.IsEmpty do
            match remaining with
            | block :: tail ->
                match topLevelParagraphId block with
                | Some blockId ->
                    let isAnchor = blockId = sel.Anchor.BlockId
                    let isFocus = blockId = sel.Focus.BlockId
                    if isAnchor || isFocus then
                        if not collecting then
                            collecting <- true
                            anchorFirst <- isAnchor
                            firstBlock <- block
                            hasFirstBlock <- true
                        lastBlock <- block
                        ids.Add blockId
                        foundAnchor <- foundAnchor || isAnchor
                        foundFocus <- foundFocus || isFocus
                    elif collecting then
                        ids.Add blockId
                | None ->
                    if collecting then
                        validTopLevelRange <- false
                remaining <- tail
            | [] -> ()
        if validTopLevelRange && hasFirstBlock && foundAnchor && foundFocus then
            let startPos, endPos =
                if anchorFirst then sel.Anchor, sel.Focus else sel.Focus, sel.Anchor
            Some (startPos, endPos, resizeArrayToList ids, firstBlock, lastBlock)
        else
            None

let private tryDeleteTopLevelSelectionRangeUntil (doc: Document) (sel: Selection) =
    if sel.Anchor.BlockId = sel.Focus.BlockId then
        ValueNone
    else
        let topLevelParagraphId = function
            | Paragraph p
            | Heading (_, p) -> ValueSome p.Id
            | Table _
            | TextBox _ -> ValueNone

        let prefix = ResizeArray<Block>()

        let buildResult startPos endPos firstBlock lastBlock afterLast =
            let firstPara = getParagraph firstBlock
            let lastPara = getParagraph lastBlock
            let beforeBuffer = TextBuffer.sliceBuffer 0 startPos.Offset firstPara.Buffer
            let lastBufferLen = TextBuffer.length lastPara.Buffer
            let afterBuffer = TextBuffer.sliceBuffer endPos.Offset (lastBufferLen - endPos.Offset) lastPara.Buffer
            let mergedBuffer = TextBuffer.append beforeBuffer afterBuffer
            let mergedPara = { firstPara with Buffer = mergedBuffer }
            let mergedBlock =
                match firstBlock with
                | Paragraph _ -> Paragraph mergedPara
                | Heading (level, _) -> Heading (level, mergedPara)
                | Table t -> Table t
                | TextBox t -> TextBox t
            let offset = TextBuffer.length beforeBuffer
            let selection =
                Some {
                    Anchor = { BlockId = firstPara.Id; Offset = offset }
                    Focus = { BlockId = firstPara.Id; Offset = offset }
                }
            { doc with Blocks = prependResizeArray prefix (mergedBlock :: afterLast); Selection = selection }

        let rec findEnd anchorFirst firstBlock remaining =
            let endTarget = if anchorFirst then sel.Focus.BlockId else sel.Anchor.BlockId
            match remaining with
            | [] -> ValueNone
            | block :: tail ->
                match topLevelParagraphId block with
                | ValueSome blockId when blockId = endTarget ->
                    let startPos, endPos =
                        if anchorFirst then sel.Anchor, sel.Focus else sel.Focus, sel.Anchor
                    ValueSome (buildResult startPos endPos firstBlock block tail)
                | ValueSome _ ->
                    findEnd anchorFirst firstBlock tail
                | ValueNone ->
                    ValueNone

        let rec findStart remaining =
            match remaining with
            | [] -> ValueNone
            | block :: tail ->
                match topLevelParagraphId block with
                | ValueSome blockId when blockId = sel.Anchor.BlockId || blockId = sel.Focus.BlockId ->
                    let anchorFirst = blockId = sel.Anchor.BlockId
                    findEnd anchorFirst block tail
                | _ ->
                    prefix.Add block
                    findStart tail

        findStart doc.Blocks

/// Delete selected text across blocks
let deleteSelection doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc  // No selection to delete
    | Some sel when sel.Anchor.BlockId = sel.Focus.BlockId ->
        let startOffset = min sel.Anchor.Offset sel.Focus.Offset
        let endOffset = max sel.Anchor.Offset sel.Focus.Offset
        let deleteLen = endOffset - startOffset
        if deleteLen <= 0 then
            doc
        else
            let mutable changed = false
            let updated =
                updateParagraph doc sel.Anchor.BlockId (fun para ->
                    changed <- true
                    { para with Buffer = TextBuffer.delete startOffset deleteLen para.Buffer })
            if changed then
                let pos = { BlockId = sel.Anchor.BlockId; Offset = startOffset }
                { updated with Selection = Some { Anchor = pos; Focus = pos } }
            else
                doc
    | Some sel ->
        match tryDeleteTopLevelSelectionRangeUntil doc sel with
        | ValueSome doc ->
            doc
        | ValueNone ->
            let topLevelRange = tryNormalizeTopLevelSelectionRangeUntil doc sel
            let (startPos, endPos, blockIds) =
                match topLevelRange with
                | Some (startPos, endPos, blockIds, _, _) -> startPos, endPos, blockIds
                | None -> getNormalizedRange sel doc

            match blockIds with
            | [] -> doc
            | [singleBlock] when singleBlock = startPos.BlockId && singleBlock = endPos.BlockId ->
                // Single block selection
                match tryFindEditableParagraphUntil doc singleBlock with
                | None -> doc
                | Some para ->
                    let len = endPos.Offset - startPos.Offset
                    let newBuffer = TextBuffer.delete startPos.Offset len para.Buffer
                    let newPara = { para with Buffer = newBuffer }
                    let newDoc = replaceParagraph doc para.Id newPara

                    // Move cursor to start of selection
                    let newSel = {
                        Anchor = startPos
                        Focus = startPos
                    }
                    { newDoc with Selection = Some newSel }
            | _ ->
                // Multi-block selection
                // Keep text before selection start in first block
                // Keep text after selection end in last block
                // Delete middle blocks

                if blockIds.IsEmpty then doc
                else
                    let firstBlockId, lastBlockId =
                        match topLevelRange with
                        | Some (_, _, _, firstBlock, lastBlock) -> getBlockId firstBlock, getBlockId lastBlock
                        | None -> List.head blockIds, List.last blockIds

                    let endpointBlocks =
                        match topLevelRange with
                        | Some (_, _, _, firstBlock, lastBlock) -> Some (firstBlock, lastBlock)
                        | None ->
                            match findBlock doc firstBlockId, findBlock doc lastBlockId with
                            | Some firstBlock, Some lastBlock -> Some (firstBlock, lastBlock)
                            | _ -> None

                    match endpointBlocks with
                    | Some (firstBlock, lastBlock) ->
                        let firstPara = getParagraph firstBlock
                        let lastPara = getParagraph lastBlock

                        // Keep content before the selection start and after the selection end.
                        let beforeBuffer = TextBuffer.sliceBuffer 0 startPos.Offset firstPara.Buffer
                        let lastBufferLen = TextBuffer.length lastPara.Buffer
                        let afterBuffer = TextBuffer.sliceBuffer endPos.Offset (lastBufferLen - endPos.Offset) lastPara.Buffer

                        // Merge into single buffer
                        let mergedBuffer = TextBuffer.append beforeBuffer afterBuffer

                        // Update first block with merged content
                        let mergedPara = { firstPara with Buffer = mergedBuffer }
                        let mergedBlock =
                            match firstBlock with
                            | Paragraph _ -> Paragraph mergedPara
                            | Heading (level, _) -> Heading (level, mergedPara)
                            | Table t -> Table t
                            | TextBox t -> TextBox t

                        let tryTopLevelRangeDelete () =
                            match topLevelRange with
                            | None -> None
                            | Some _ ->
                                let prefix = ResizeArray<Block>()
                                let rec skipToLast remaining =
                                    match remaining with
                                    | [] -> None
                                    | block :: tail when getBlockId block = lastBlockId -> Some tail
                                    | _ :: tail -> skipToLast tail
                                let rec loop remaining =
                                    match remaining with
                                    | [] -> None
                                    | block :: tail when getBlockId block = firstBlockId ->
                                        skipToLast tail
                                        |> Option.map (fun afterLast -> prependResizeArray prefix (mergedBlock :: afterLast))
                                    | block :: tail ->
                                        prefix.Add block
                                        loop tail
                                loop doc.Blocks

                        let deleteNestedSelectionFallback () =
                            // Remove all selected editable paragraphs and replace the first with merged content.
                            // Table-cell paragraphs are not top-level blocks, so filter them inside table cells.
                            let removeIds = HashSet<BlockId>()
                            let mutable remainingIds =
                                match blockIds with
                                | _ :: tail -> tail
                                | [] -> []
                            while not remainingIds.IsEmpty do
                                match remainingIds with
                                | blockId :: tail ->
                                    removeIds.Add blockId |> ignore
                                    remainingIds <- tail
                                | [] -> ()

                            let filteredCellContent (content: ParagraphBlock list) =
                                let values = ResizeArray<ParagraphBlock>()
                                let mutable changed = false
                                for para in content do
                                    if removeIds.Contains para.Id then
                                        changed <- true
                                    else
                                        values.Add para
                                if changed then resizeArrayToList values else content

                            let filteredRows (rows: TableRow list) =
                                let rowValues = ResizeArray<TableRow>()
                                let mutable changed = false
                                for row in rows do
                                    let cellValues = ResizeArray<TableCell>()
                                    let mutable rowChanged = false
                                    for cell in row.Cells do
                                        let content = filteredCellContent cell.Content
                                        rowChanged <- rowChanged || not (obj.ReferenceEquals(content, cell.Content))
                                        cellValues.Add { cell with Content = content }
                                    if rowChanged then
                                        rowValues.Add { row with Cells = resizeArrayToList cellValues }
                                        changed <- true
                                    else
                                        rowValues.Add row
                                if changed then resizeArrayToList rowValues else rows

                            let rec skipToLast remaining =
                                match remaining with
                                | [] -> None
                                | block :: tail when getBlockId block = lastBlockId -> Some tail
                                | block :: tail when removeIds.Contains(getBlockId block) -> skipToLast tail
                                | _ -> None
                            let rec loop prefix remaining =
                                match remaining with
                                | [] -> None
                                | block :: tail when getBlockId block = firstBlockId ->
                                    skipToLast tail
                                    |> Option.map (fun afterLast -> prependResizeArray prefix (mergedBlock :: afterLast))
                                | block :: tail ->
                                    prefix.Add block
                                    loop prefix tail
                            match loop (ResizeArray<Block>()) doc.Blocks with
                            | Some blocks -> blocks
                            | None ->
                                let newBlockValues = ResizeArray<Block>()
                                for block in doc.Blocks do
                                    match block with
                                    | Paragraph p when removeIds.Contains p.Id -> ()
                                    | Heading (_, p) when removeIds.Contains p.Id -> ()
                                    | Paragraph p when p.Id = firstBlockId -> newBlockValues.Add mergedBlock
                                    | Heading (_, p) when p.Id = firstBlockId -> newBlockValues.Add mergedBlock
                                    | Table t ->
                                        let rows = filteredRows t.Rows
                                        if obj.ReferenceEquals(rows, t.Rows) then
                                            newBlockValues.Add block
                                        else
                                            newBlockValues.Add(Table { t with Rows = rows })
                                    | block -> newBlockValues.Add block
                                resizeArrayToList newBlockValues

                        let newBlocks =
                            match tryTopLevelRangeDelete () with
                            | Some blocks -> blocks
                            | None -> deleteNestedSelectionFallback ()

                        // Move cursor to merge point
                        let newSel = {
                            Anchor = { BlockId = firstBlockId; Offset = TextBuffer.length beforeBuffer }
                            Focus = { BlockId = firstBlockId; Offset = TextBuffer.length beforeBuffer }
                        }

                        { doc with Blocks = newBlocks; Selection = Some newSel }
                    | _ -> doc

/// Delete forward (Delete key)
let private tryDeleteForwardTopLevelParagraph (doc: Document) (blockId: BlockId) offset =
    let prefix = ResizeArray<Block>()
    let finish tail (para: ParagraphBlock) wrap =
        let maxOffset = TextBuffer.length para.Buffer
        if offset >= maxOffset then
            match tail with
            | nextBlock :: afterNext ->
                match tryGetParagraph nextBlock with
                | Some nextPara ->
                    let mergedBuffer = TextBuffer.append para.Buffer nextPara.Buffer
                    let mergedPara = { para with Buffer = mergedBuffer }
                    ValueSome { doc with Blocks = prependResizeArray prefix (wrap mergedPara :: afterNext) }
                | None ->
                    ValueSome doc
            | [] ->
                ValueSome doc
        else
            let newBuffer = TextBuffer.delete offset 1 para.Buffer
            let updatedPara = { para with Buffer = newBuffer }
            ValueSome { doc with Blocks = prependResizeArray prefix (wrap updatedPara :: tail) }

    let rec loop remaining =
        match remaining with
        | [] -> ValueNone
        | block :: tail ->
            match block with
            | Paragraph para when para.Id = blockId ->
                finish tail para Paragraph
            | Heading (level, para) when para.Id = blockId ->
                finish tail para (fun para -> Heading (level, para))
            | _ ->
                prefix.Add block
                loop tail

    loop doc.Blocks

let deleteForward doc =
    match doc.Selection with
    | None -> doc
    | Some sel when not sel.IsCollapsed ->
        // Delete selection if active
        deleteSelection doc
    | Some sel ->
        match tryDeleteForwardTopLevelParagraph doc sel.Focus.BlockId sel.Focus.Offset with
        | ValueSome doc -> doc
        | ValueNone ->
            let mutable found = false
            let mutable atEnd = false
            let newDoc =
                updateParagraph doc sel.Focus.BlockId (fun para ->
                    found <- true
                    let maxOffset = TextBuffer.length para.Buffer

                    if sel.Focus.Offset >= maxOffset then
                        atEnd <- true
                        para
                    else
                        let newBuffer = TextBuffer.delete sel.Focus.Offset 1 para.Buffer
                        { para with Buffer = newBuffer })
            if not found then doc
            elif atEnd then mergeWithNext doc sel.Focus.BlockId
            else newDoc

let private setSelectionIfChanged selection doc =
    match doc.Selection with
    | Some current when current = selection -> doc
    | _ -> { doc with Selection = Some selection }

let mutable private moveCursorUpNoopCacheSet = false
let mutable private moveCursorUpNoopCacheDoc = Unchecked.defaultof<Document>

let private rememberMoveCursorUpNoop doc =
    moveCursorUpNoopCacheSet <- true
    moveCursorUpNoopCacheDoc <- doc
    doc

/// Move cursor up to previous block
let moveCursorUp doc =
    if moveCursorUpNoopCacheSet && obj.ReferenceEquals(moveCursorUpNoopCacheDoc, doc) then
        doc
    else
        let mutable resultDoc = doc
        match doc.Selection with
        | None -> ()
        | Some sel ->
            match tryFindAdjacentTopBlock doc sel.Focus.BlockId -1 with
            | None -> ()
            | Some prevBlock ->
                let prevPara = getParagraph prevBlock
                let prevLen = TextBuffer.length prevPara.Buffer

                // Move to end of previous block
                let newSel = {
                    Anchor = { BlockId = prevPara.Id; Offset = prevLen }
                    Focus = { BlockId = prevPara.Id; Offset = prevLen }
                }
                resultDoc <- setSelectionIfChanged newSel doc
        if obj.ReferenceEquals(resultDoc, doc) then
            rememberMoveCursorUpNoop doc
        else
            resultDoc

let mutable private moveCursorDownNoopCacheSet = false
let mutable private moveCursorDownNoopCacheDoc = Unchecked.defaultof<Document>

let private rememberMoveCursorDownNoop doc =
    moveCursorDownNoopCacheSet <- true
    moveCursorDownNoopCacheDoc <- doc
    doc

/// Move cursor down to next block
let moveCursorDown doc =
    if moveCursorDownNoopCacheSet && obj.ReferenceEquals(moveCursorDownNoopCacheDoc, doc) then
        doc
    else
        let mutable resultDoc = doc
        match doc.Selection with
        | None -> ()
        | Some sel ->
            match tryFindAdjacentTopBlock doc sel.Focus.BlockId 1 with
            | None -> ()
            | Some nextBlock ->
                let nextPara = getParagraph nextBlock

                // Move to start of next block
                let newSel = {
                    Anchor = { BlockId = nextPara.Id; Offset = 0 }
                    Focus = { BlockId = nextPara.Id; Offset = 0 }
                }
                resultDoc <- setSelectionIfChanged newSel doc
        if obj.ReferenceEquals(resultDoc, doc) then
            rememberMoveCursorDownNoop doc
        else
            resultDoc

/// Move cursor to start of current block (Home key)
let moveCursorToLineStart doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.Focus.Offset = 0 && sel.Anchor = sel.Focus ->
        doc
    | Some sel ->
        let newSel = {
            Anchor = { sel.Focus with Offset = 0 }
            Focus = { sel.Focus with Offset = 0 }
        }
        setSelectionIfChanged newSel doc

let private findCursorParagraph doc blockId =
    let cachedBlock =
        match tryGetBlockLookup doc.Blocks with
        | ValueSome cached ->
            match tryFindBlockLookupEntry cached blockId with
            | ValueSome entry -> entry.TopParagraphResult
            | ValueNone -> None
        | ValueNone -> None
    match cachedBlock with
    | Some _ -> cachedBlock
    | None ->
        let cachedParagraph =
            match tryGetParagraphLookup doc.Blocks with
            | ValueSome cached ->
                let fast = cached.ParagraphFastSlots.[paragraphLookupEntryFastSlot cached.ParagraphFastSlots blockId]
                if not (obj.ReferenceEquals(fast, null))
                   && paragraphLookupEntryMatches fast blockId then
                    fast.Result
                else
                    let mutable entry = Unchecked.defaultof<ParagraphLookupEntry>
                    if cached.Paragraphs.TryGetValue(blockId, &entry) then
                        entry.RecentLookupIdSet <- true
                        entry.RecentLookupId <- blockId
                        setParagraphLookupEntryFastSlot cached.ParagraphFastSlots entry
                        entry.Result
                    else
                        None
            | ValueNone -> None
        match cachedParagraph with
        | Some _ -> cachedParagraph
        | None ->
            match tryFindTopLevelParagraphUntil doc blockId with
            | Some _ as topLevel -> topLevel
            | None ->
                let shouldBuild =
                    match doc.Blocks with
                    | [ Table table ] -> shouldBuildCurrentTableCellCacheForSingleTable table
                    | _ -> shouldBuildCurrentTableCellCacheForBlocks doc.Blocks
                if shouldBuild then
                    findParagraphFromParagraphLookup doc blockId
                else
                    tryFindEditableParagraphUntil doc blockId

let private tryFindCursorParagraphLength doc blockId =
    match findCursorParagraph doc blockId with
    | Some para -> ValueSome (TextBuffer.length para.Buffer)
    | None -> ValueNone

/// Move cursor to end of current block (End key)
let moveCursorToLineEnd doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match tryFindCursorParagraphLength doc sel.Focus.BlockId with
        | ValueNone -> doc
        | ValueSome maxOffset ->
            if sel.Focus.Offset = maxOffset && sel.Anchor = sel.Focus then
                doc
            else
                let newSel = {
                    Anchor = { sel.Focus with Offset = maxOffset }
                    Focus = { sel.Focus with Offset = maxOffset }
                }
                setSelectionIfChanged newSel doc


/// Insert text at cursor
let insertText text doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        // Delete selection if not collapsed
        let doc' = if sel.IsCollapsed then doc else deleteSelection doc

        // Insert text at cursor (which is now at the selection start after deletion)
        match doc'.Selection with
        | None -> doc'
        | Some sel' ->
            let mutable changed = false
            let newDoc =
                updateParagraph doc' sel'.Focus.BlockId (fun para ->
                    changed <- true
                    let newBuffer = TextBuffer.insert sel'.Focus.Offset text para.Buffer
                    { para with Buffer = newBuffer })
            if changed then
                // Move cursor forward
                let newSel = {
                    Anchor = { sel'.Focus with Offset = sel'.Focus.Offset + text.Length }
                    Focus = { sel'.Focus with Offset = sel'.Focus.Offset + text.Length }
                }
                { newDoc with Selection = Some newSel }
            else doc'

let insertField fieldCode result doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        let doc' = if sel.IsCollapsed then doc else deleteSelection doc
        match doc'.Selection with
        | None -> doc'
        | Some sel' ->
            let startOffset = sel'.Focus.Offset
            let mutable changed = false
            let newDoc =
                updateParagraph doc' sel'.Focus.BlockId (fun para ->
                    changed <- true
                    let inserted = TextBuffer.insert startOffset result para.Buffer
                    let styled =
                        TextBuffer.applyStyle startOffset result.Length (fun style ->
                            { style with FieldCode = Some fieldCode; FieldResult = Some result }) inserted
                    { para with Buffer = styled })
            if changed then
                let nextPos = { sel'.Focus with Offset = startOffset + result.Length }
                { newDoc with Selection = Some { Anchor = nextPos; Focus = nextPos } }
            else
                doc'

let private escapeXml (text: string) =
    text
        .Replace("&", "&amp;")
        .Replace("<", "&lt;")
        .Replace(">", "&gt;")
        .Replace("\"", "&quot;")
        .Replace("'", "&apos;")

let insertEquation equationText doc =
    let text = "[Equation]"
    let equationXml = sprintf "<m:oMath><m:r><m:t>%s</m:t></m:r></m:oMath>" (escapeXml equationText)
    match doc.Selection with
    | None -> doc
    | Some sel ->
        let doc' = insertText text doc
        match doc'.Selection with
        | None -> doc'
        | Some sel' ->
            let startOffset = max 0 (sel'.Focus.Offset - text.Length)
            match findCursorParagraph doc' sel'.Focus.BlockId with
            | None -> doc'
            | Some para ->
                let newBuffer =
                    TextBuffer.applyStyle startOffset text.Length (fun style ->
                        { style with EquationXml = Some equationXml; FontFamily = Some "Cambria Math" }) para.Buffer
                let newPara = { para with Buffer = newBuffer }
                replaceParagraph doc' para.Id newPara

/// Insert image at cursor position
let insertImageAtCursor src width height alt doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        // Delete selection if not collapsed
        let doc' = if sel.IsCollapsed then doc else deleteSelection doc

        // Insert image at cursor
        match doc'.Selection with
        | None -> doc'
        | Some sel' ->
            match findCursorParagraph doc' sel'.Focus.BlockId with
            | None -> doc'
            | Some para ->
                let newBuffer = TextBuffer.insertImage sel'.Focus.Offset src width height alt para.Buffer
                let newPara = { para with Buffer = newBuffer }

                let newDoc = replaceParagraph doc' para.Id newPara

                // Move cursor forward by 1 (image occupies 1 character position)
                let newSel = {
                    Anchor = { sel'.Focus with Offset = sel'.Focus.Offset + 1 }
                    Focus = { sel'.Focus with Offset = sel'.Focus.Offset + 1 }
                }
                { newDoc with Selection = Some newSel }

/// Resize image at specified position
let resizeImage blockId offset newWidth newHeight doc =
    let resizeInParagraph para =
        let newBuffer = TextBuffer.resizeImage offset newWidth newHeight para.Buffer
        if obj.ReferenceEquals(newBuffer, para.Buffer) then
            para
        else
            { para with Buffer = newBuffer }

    match tryUpdateTopLevelParagraphBeforeNestedContent doc blockId resizeInParagraph with
    | ValueSome updatedDoc -> updatedDoc
    | ValueNone ->
        match findCursorParagraph doc blockId with
        | None -> doc
        | Some para ->
            let newBuffer = TextBuffer.resizeImage offset newWidth newHeight para.Buffer
            if obj.ReferenceEquals(newBuffer, para.Buffer) then
                doc
            else
                let newPara = { para with Buffer = newBuffer }
                replaceParagraph doc para.Id newPara

let resizeImageNearCursor newWidth newHeight doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        let resizeNearFocus para =
            let mutable imageOffset = None
            for run in para.Buffer.Runs do
                match run with
                | TextBuffer.Image img when img.Start <= sel.Focus.Offset ->
                    match imageOffset with
                    | Some offset when offset >= img.Start -> ()
                    | _ -> imageOffset <- Some img.Start
                | _ -> ()
            match imageOffset with
            | None -> para
            | Some offset ->
                let newBuffer = TextBuffer.resizeImage offset newWidth newHeight para.Buffer
                if obj.ReferenceEquals(newBuffer, para.Buffer) then
                    para
                else
                    { para with Buffer = newBuffer }

        match tryUpdateTopLevelParagraphBeforeNestedContent doc sel.Focus.BlockId resizeNearFocus with
        | ValueSome updatedDoc -> updatedDoc
        | ValueNone ->
            match findCursorParagraph doc sel.Focus.BlockId with
            | None -> doc
            | Some para ->
                let newPara = resizeNearFocus para
                if obj.ReferenceEquals(newPara, para) || paragraphReplacementMatches para newPara then
                    doc
                else
                    replaceParagraph doc para.Id newPara

/// Move an image run to a new document position.
let moveImage blockId offset targetBlockId targetOffset doc =
    let insertImageRun boundedOffset (img: TextBuffer.ImageRun) buffer =
        TextBuffer.insertImageWithMetadata boundedOffset img.Src img.Linked img.Width img.Height img.Alt img.Name img.Title img.Floating img.Wrap img.PositionX img.PositionY img.PositionHRelativeFrom img.PositionVRelativeFrom img.BehindDoc img.AllowOverlap img.LayoutInCell img.DistanceTop img.DistanceBottom img.DistanceLeft img.DistanceRight img.RelativeHeight img.Locked img.SimplePosition img.SimplePositionX img.SimplePositionY img.CropLeft img.CropTop img.CropRight img.CropBottom buffer
    if blockId = targetBlockId && (targetOffset = offset || targetOffset = offset + 1) then
        doc
    else
        match findCursorParagraph doc blockId with
        | None -> doc
        | Some sourcePara ->
            if blockId = targetBlockId then
                let insertOffset =
                    if targetOffset > offset then targetOffset - 1
                    else targetOffset
                let boundedOffset = max 0 (min (max 0 (TextBuffer.length sourcePara.Buffer - 1)) insertOffset)
                if boundedOffset = offset then
                    doc
                else
                    match sourcePara.Buffer.Runs |> List.tryPick (function TextBuffer.Image img when img.Start = offset -> Some img | _ -> None) with
                    | None -> doc
                    | Some img ->
                        let targetBuffer =
                            match TextBuffer.tryMoveImageWithinBuffer offset boundedOffset img sourcePara.Buffer with
                            | Some buffer -> buffer
                            | None ->
                                let sourceBuffer = TextBuffer.delete offset 1 sourcePara.Buffer
                                insertImageRun boundedOffset img sourceBuffer
                        if obj.ReferenceEquals(targetBuffer, sourcePara.Buffer) then
                            doc
                        else
                            let nextDoc = replaceParagraph doc sourcePara.Id { sourcePara with Buffer = targetBuffer }
                            { nextDoc with Selection = Some { Anchor = { BlockId = targetBlockId; Offset = boundedOffset + 1 }; Focus = { BlockId = targetBlockId; Offset = boundedOffset + 1 } } }
            else
                match sourcePara.Buffer.Runs |> List.tryPick (function TextBuffer.Image img when img.Start = offset -> Some img | _ -> None) with
                | None -> doc
                | Some img ->
                    let sourceBuffer = TextBuffer.delete offset 1 sourcePara.Buffer
                    let docWithoutImage = replaceParagraph doc sourcePara.Id { sourcePara with Buffer = sourceBuffer }
                    match findCursorParagraph docWithoutImage targetBlockId with
                    | None -> docWithoutImage
                    | Some targetPara ->
                        let boundedOffset = max 0 (min (TextBuffer.length targetPara.Buffer) targetOffset)
                        let targetBuffer = insertImageRun boundedOffset img targetPara.Buffer
                        let nextDoc = replaceParagraph docWithoutImage targetPara.Id { targetPara with Buffer = targetBuffer }
                        { nextDoc with Selection = Some { Anchor = { BlockId = targetBlockId; Offset = boundedOffset + 1 }; Focus = { BlockId = targetBlockId; Offset = boundedOffset + 1 } } }

/// Delete backward (backspace)
let deleteBackward doc =
    match doc.Selection with
    | None -> doc
    | Some sel when not sel.IsCollapsed ->
        // Delete selection if active
        deleteSelection doc
    | Some sel when sel.Focus.Offset = 0 ->
        // At start of block, merge with previous
        mergeWithPrevious doc sel.Focus.BlockId
    | Some sel ->
        match findCursorParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let newBuffer = TextBuffer.delete (sel.Focus.Offset - 1) 1 para.Buffer
            let newPara = { para with Buffer = newBuffer }

            let newDoc = replaceParagraph doc para.Id newPara

            // Move cursor back
            let newSel = {
                Anchor = { sel.Focus with Offset = sel.Focus.Offset - 1 }
                Focus = { sel.Focus with Offset = sel.Focus.Offset - 1 }
            }
            { newDoc with Selection = Some newSel }

/// Set cursor position
let setCursor blockId offset doc =
    match doc.Selection with
    | Some current
        when current.Anchor.BlockId = blockId
            && current.Anchor.Offset = offset
            && current.Focus.BlockId = blockId
            && current.Focus.Offset = offset ->
        doc
    | _ ->
        { doc with
            Selection = Some {
                Anchor = { BlockId = blockId; Offset = offset }
                Focus = { BlockId = blockId; Offset = offset }
            } }

let insertTableRowAfterSelection doc =
    match currentTableCell doc with
    | ValueNone -> doc
    | ValueSome (struct (tableIndex, rowIndex, _)) ->
        updateTableAtIndex tableIndex (fun table ->
            let newRow columnCount = { Cells = [ for _ in 1 .. columnCount -> emptyTableCell () ]; IsHeader = false; Height = None; HeightRule = None; CanSplit = None; CellSpacing = None; PaddingTop = None; PaddingRight = None; PaddingBottom = None; PaddingLeft = None; GridBefore = 0; GridAfter = 0; WidthBefore = None; WidthAfter = None; BordersOverride = None }
            let prefix = ResizeArray<TableRow>(max 0 rowIndex)
            let rec insertAfter rowPosition rows =
                match rows with
                | [] ->
                    prependResizeArray prefix [ newRow 1 ]
                | row :: tail when rowPosition = rowIndex ->
                    prependResizeArray prefix (row :: newRow (max 1 row.Cells.Length) :: tail)
                | row :: tail ->
                    prefix.Add row
                    insertAfter (rowPosition + 1) tail
            { table with Rows = insertAfter 0 table.Rows }) doc

let insertTableColumnAfterSelection doc =
    match currentTableCell doc with
    | ValueNone -> doc
    | ValueSome (struct (tableIndex, _, colIndex)) ->
        updateTableAtIndex tableIndex (fun table ->
            { table with
                Rows =
                    table.Rows
                    |> List.map (fun row ->
                        match tryInsertListItemAfterAtV colIndex emptyTableCell row.Cells with
                        | ValueSome cells -> { row with Cells = cells }
                        | ValueNone -> row) }) doc

let deleteSelectedTableRow doc =
    match currentTableCell doc with
    | ValueNone -> doc
    | ValueSome (struct (tableIndex, rowIndex, colIndex)) ->
        let mutable deletedRow = false
        let mutable targetPara = None
        let updatedDoc = updateTableAtIndex tableIndex (fun table ->
            match table.Rows with
            | [] | [_] -> table
            | rows ->
                let prefix = ResizeArray<TableRow>(max 0 rowIndex)
                let rec deleteAt rowPosition rows =
                    match rows with
                    | [] -> table
                    | _ :: tail when rowPosition = rowIndex ->
                        deletedRow <- true
                        targetPara <-
                            match tail with
                            | nextRow :: _ ->
                                tryGetTableCellParagraphInRow colIndex nextRow
                            | [] when prefix.Count > 0 ->
                                tryGetTableCellParagraphInRow colIndex prefix.[prefix.Count - 1]
                            | [] ->
                                None
                        let nextTable = { table with Rows = prependResizeArray prefix tail }
                        nextTable
                    | row :: tail ->
                        prefix.Add row
                        deleteAt (rowPosition + 1) tail
                deleteAt 0 rows) doc

        if deletedRow then
            match targetPara with
            | Some para -> setCursor para.Id 0 updatedDoc
            | None -> updatedDoc
        else
            updatedDoc

let deleteSelectedTableColumn doc =
    match currentTableCell doc with
    | ValueNone -> doc
    | ValueSome (struct (tableIndex, rowIndex, colIndex)) ->
        let mutable deletedColumn = false
        let mutable targetPara = None
        let updatedDoc = updateTableAtIndex tableIndex (fun table ->
            let mutable hasMultipleColumns = false
            let mutable rows = table.Rows
            while not hasMultipleColumns && not rows.IsEmpty do
                match rows.Head.Cells with
                | _ :: _ :: _ -> hasMultipleColumns <- true
                | _ -> ()
                rows <- rows.Tail
            if not hasMultipleColumns then table
            else
                deletedColumn <- true
                let targetCol = max 0 (colIndex - 1)
                let mutable rowPosition = 0
                let nextTable =
                    { table with
                        Rows =
                            table.Rows
                            |> List.map (fun row ->
                                let nextRow =
                                    match tryRemoveListItemAtV colIndex row.Cells with
                                    | ValueSome cells -> { row with Cells = cells }
                                    | ValueNone -> row
                                if rowPosition = rowIndex then
                                    targetPara <- tryGetTableCellParagraphInRow targetCol nextRow
                                rowPosition <- rowPosition + 1
                                nextRow) }
                nextTable) doc

        if deletedColumn then
            match targetPara with
            | Some para -> setCursor para.Id 0 updatedDoc
            | None -> updatedDoc
        else
            updatedDoc

let insertTable rows columns doc =
    let rowCount = max 1 rows
    let columnCount = max 1 columns
    let table = {
        Id = nextBlockId ()
        Rows = [ for _ in 1 .. rowCount -> { Cells = [ for _ in 1 .. columnCount -> emptyTableCell () ]; IsHeader = false; Height = None; HeightRule = None; CanSplit = None; CellSpacing = None; PaddingTop = None; PaddingRight = None; PaddingBottom = None; PaddingLeft = None; GridBefore = 0; GridAfter = 0; WidthBefore = None; WidthAfter = None; BordersOverride = None } ]
        PreferredWidth = None
        Alignment = None
        LayoutMode = None
        GridColumns = []
        Borders = None
        BackgroundColor = None
        Shading = None
        CellSpacing = None
        PaddingTop = None
        PaddingRight = None
        PaddingBottom = None
        PaddingLeft = None
        Floating = false
    }
    let firstCellPara = table.Rows.Head.Cells.Head.Content.Head
    let tableBlock = Table table
    let blocks =
        match doc.Selection with
        | Some sel ->
            insertBlockAfterTopLevelId sel.Focus.BlockId tableBlock doc.Blocks
        | None -> appendBlock tableBlock doc.Blocks
    { doc with Blocks = blocks } |> setCursor firstCellPara.Id 0

let private insertTopLevelBlockAfterSelection block doc =
    let blocks =
        match doc.Selection with
        | Some sel ->
            insertBlockAfterTopLevelId sel.Focus.BlockId block doc.Blocks
        | None -> appendBlock block doc.Blocks
    { doc with Blocks = blocks }

let insertTextBox text doc =
    let para = {
        Id = nextBlockId ()
        Buffer = TextBuffer.ofString text
        Metadata = defaultMetadata
    }
    let textBox = TextBox {
        Id = nextBlockId ()
        Content = [para]
        X = Some 96.0
        Y = Some 96.0
        PositionHRelativeFrom = None
        PositionVRelativeFrom = None
        Width = 240.0
        Height = 96.0
        FillColor = Some "#FFFFFF"
        BorderColor = Some "#808080"
        BorderWidth = Some 1.0
        RelativeHeight = None
        Geometry = GeoRect
        Gradient = None
        Rotation = 0.0
    }
    insertTopLevelBlockAfterSelection textBox doc

let private iterTextRunsInParagraph para f =
    for runValue in para.Buffer.Runs do
        match runValue with
        | TextBuffer.Text run -> f run
        | TextBuffer.Image _ -> ()

let private iterTextRunsInBlock block f =
    match block with
    | Paragraph para
    | Heading (_, para) -> iterTextRunsInParagraph para f
    | Table table ->
        for row in table.Rows do
            for cell in row.Cells do
                for para in cell.Content do
                    iterTextRunsInParagraph para f
    | TextBox textBox ->
        for para in textBox.Content do
            iterTextRunsInParagraph para f

type private CaptionIdMetricsCache = {
    SourceBlocks: Block list
    SourceLabel: string
    MaxSequenceResult: int
    MaxBookmarkId: int
    LatestCaption: (int * string * string) option
    SourceParagraphCount: int
    SourceTextCharCount: int
}

let mutable private captionIdMetricsCache: CaptionIdMetricsCache list = []
[<Struct>]
type private CaptionIdMetricsCacheKey = {
    KeyBlocks: Block list
    KeyLabel: string
}

let private captionIdMetricsCacheKeyComparer =
    { new IEqualityComparer<CaptionIdMetricsCacheKey> with
        member _.Equals(left, right) =
            obj.ReferenceEquals(left.KeyBlocks, right.KeyBlocks)
            && left.KeyLabel = right.KeyLabel
        member _.GetHashCode(value) =
            (System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value.KeyBlocks) * 397)
            ^^^ hash value.KeyLabel }
let private captionIdMetricsCacheByKey =
    Dictionary<CaptionIdMetricsCacheKey, CaptionIdMetricsCache>(captionIdMetricsCacheKeyComparer)
let private captionIdMetricsFastSlots = Array.zeroCreate<CaptionIdMetricsCache> 256
let private captionIdMetricsCacheMaxEntries = 256
let private captionIdMetricsCacheMaxTotalSourceParagraphs = sourceParagraphCountCacheLimit + 20000
let private captionIdMetricsCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars

let private captionIdMetricsCacheMatches blocks label cached =
    obj.ReferenceEquals(cached.SourceBlocks, blocks)
    && cached.SourceLabel = label

let private captionIdMetricsCacheKey blocks label = {
    KeyBlocks = blocks
    KeyLabel = label
}

let private captionIdMetricsCacheKeyMatches key cached =
    obj.ReferenceEquals(cached.SourceBlocks, key.KeyBlocks)
    && cached.SourceLabel = key.KeyLabel

let private captionIdMetricsCacheSlot key slotCount =
    ((System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.KeyBlocks) * 397)
     ^^^ hash key.KeyLabel)
    &&& (slotCount - 1)

let private setCaptionIdMetricsCache entries =
    captionIdMetricsCache <- entries
    captionIdMetricsCacheByKey.Clear()
    clearArray captionIdMetricsFastSlots
    for entry in entries do
        let key = captionIdMetricsCacheKey entry.SourceBlocks entry.SourceLabel
        let mutable existing = Unchecked.defaultof<CaptionIdMetricsCache>
        if not (captionIdMetricsCacheByKey.TryGetValue(key, &existing)) then
            captionIdMetricsCacheByKey.[key] <- entry
            captionIdMetricsFastSlots.[captionIdMetricsCacheSlot key captionIdMetricsFastSlots.Length] <- entry

let private tryGetCaptionIdMetricsCache blocks label =
    let key = captionIdMetricsCacheKey blocks label
    let slot = captionIdMetricsCacheSlot key captionIdMetricsFastSlots.Length
    let fast = captionIdMetricsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && captionIdMetricsCacheKeyMatches key fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<CaptionIdMetricsCache>
        if captionIdMetricsCacheByKey.TryGetValue(key, &cached) then
            captionIdMetricsFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private captionIdMetricsCacheRetainable cached =
    cached.SourceParagraphCount <= captionIdMetricsCacheMaxTotalSourceParagraphs
    && cached.SourceTextCharCount <= captionIdMetricsCacheMaxTotalSourceTextChars

let private trimCaptionIdMetricsCache entries =
    let rec loop keptCount keptBlocks keptSourceTextChars acc remaining =
        match remaining with
        | [] -> List.rev acc
        | candidate :: tail ->
            let candidateBlocks = max 1 candidate.SourceParagraphCount
            let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
            if captionIdMetricsCacheRetainable candidate
               && keptCount < captionIdMetricsCacheMaxEntries
               && keptBlocks + candidateBlocks <= captionIdMetricsCacheMaxTotalSourceParagraphs
               && keptSourceTextChars + candidateSourceTextChars <= captionIdMetricsCacheMaxTotalSourceTextChars then
                loop
                    (keptCount + 1)
                    (keptBlocks + candidateBlocks)
                    (keptSourceTextChars + candidateSourceTextChars)
                    (candidate :: acc)
                    tail
            else
                loop keptCount keptBlocks keptSourceTextChars acc tail
    loop 0 0 0 [] entries

let private captionIdMetricsCacheWithoutMatching blocks label =
    let key = captionIdMetricsCacheKey blocks label
    let mutable cached = Unchecked.defaultof<CaptionIdMetricsCache>
    if captionIdMetricsCacheByKey.TryGetValue(key, &cached)
       && captionIdMetricsCacheKeyMatches key cached then
        captionIdMetricsCache
        |> List.filter (captionIdMetricsCacheMatches blocks label >> not)
    else
        captionIdMetricsCache

let private rememberCaptionIdMetrics cached =
    let withoutMatching =
        captionIdMetricsCacheWithoutMatching cached.SourceBlocks cached.SourceLabel
    if not (captionIdMetricsCacheRetainable cached) then
        if not (obj.ReferenceEquals(withoutMatching, captionIdMetricsCache)) then
            setCaptionIdMetricsCache withoutMatching
    else
        (cached :: withoutMatching)
        |> trimCaptionIdMetricsCache
        |> setCaptionIdMetricsCache

let private captionIdMetricsSourceTextCharCountFor sourceParagraphCount blocks =
    documentBlocksSourceTextCharCountFor
        captionIdMetricsCacheMaxTotalSourceParagraphs
        sourceParagraphCount
        blocks

let private captionIdMetrics label doc =
    match tryGetCaptionIdMetricsCache doc.Blocks label with
    | ValueSome cached -> cached
    | ValueNone ->
        let prefix = sprintf " SEQ %s " label
        let mutable maxSequenceResult = 0
        let mutable maxBookmarkId = 0
        let mutable latestCaption = None
        let addBookmarkId value =
            match value with
            | Some raw ->
                match System.Int32.TryParse(raw: string) with
                | true, parsed -> maxBookmarkId <- max maxBookmarkId parsed
                | false, _ -> ()
            | None -> ()
        for block in doc.Blocks do
            iterTextRunsInBlock block (fun run ->
                match run.Style.FieldCode, run.Style.FieldResult with
                | Some code, Some result when code.StartsWith(prefix) ->
                    match System.Int32.TryParse(result: string) with
                    | true, value ->
                        maxSequenceResult <- max maxSequenceResult value
                        match run.Style.BookmarkName with
                        | Some bookmarkName ->
                            match latestCaption with
                            | Some (latestNumber, _, _) when latestNumber >= value -> ()
                            | _ -> latestCaption <- Some (value, result, bookmarkName)
                        | None -> ()
                    | false, _ -> ()
                | _ -> ()
                addBookmarkId run.Style.BookmarkStartId
                addBookmarkId run.Style.BookmarkEndId)
        let sourceParagraphCount = countSourceParagraphsWithLimit captionIdMetricsCacheMaxTotalSourceParagraphs doc.Blocks
        let cached = {
            SourceBlocks = doc.Blocks
            SourceLabel = label
            MaxSequenceResult = maxSequenceResult
            MaxBookmarkId = maxBookmarkId
            LatestCaption = latestCaption
            SourceParagraphCount = sourceParagraphCount
            SourceTextCharCount = captionIdMetricsSourceTextCharCountFor sourceParagraphCount doc.Blocks
        }
        rememberCaptionIdMetrics cached
        cached

let private nextSequenceResult label doc =
    let metrics = captionIdMetrics label doc
    string (metrics.MaxSequenceResult + 1)

type private BookmarkIdMetricsCache = {
    SourceBlocks: Block list
    MaxBookmarkId: int
    SourceParagraphCount: int
    SourceTextCharCount: int
}

let mutable private bookmarkIdMetricsCache: BookmarkIdMetricsCache list = []
let private bookmarkIdMetricsCacheByBlocks =
    Dictionary<Block list, BookmarkIdMetricsCache>(documentInfoBlocksComparer)
let private bookmarkIdMetricsFastSlots = Array.zeroCreate<BookmarkIdMetricsCache> 256
let private bookmarkIdMetricsCacheMaxEntries = 256
let private bookmarkIdMetricsCacheMaxTotalSourceParagraphs = sourceParagraphCountCacheLimit
let private bookmarkIdMetricsCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars

let private bookmarkIdMetricsCacheMatches blocks cached =
    obj.ReferenceEquals(cached.SourceBlocks, blocks)

let private setBookmarkIdMetricsCache entries =
    bookmarkIdMetricsCache <- entries
    bookmarkIdMetricsCacheByBlocks.Clear()
    clearArray bookmarkIdMetricsFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<BookmarkIdMetricsCache>
        if not (bookmarkIdMetricsCacheByBlocks.TryGetValue(entry.SourceBlocks, &existing)) then
            bookmarkIdMetricsCacheByBlocks.[entry.SourceBlocks] <- entry
            bookmarkIdMetricsFastSlots.[documentInfoSlot entry.SourceBlocks bookmarkIdMetricsFastSlots.Length] <- entry

let private tryGetBookmarkIdMetricsCache blocks =
    let slot = documentInfoSlot blocks bookmarkIdMetricsFastSlots.Length
    let fast = bookmarkIdMetricsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && bookmarkIdMetricsCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<BookmarkIdMetricsCache>
        if bookmarkIdMetricsCacheByBlocks.TryGetValue(blocks, &cached) then
            bookmarkIdMetricsFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private bookmarkIdMetricsCacheRetainable cached =
    cached.SourceParagraphCount <= bookmarkIdMetricsCacheMaxTotalSourceParagraphs
    && cached.SourceTextCharCount <= bookmarkIdMetricsCacheMaxTotalSourceTextChars

let private bookmarkIdMetricsSourceTextCharCountFor sourceParagraphCount blocks =
    documentBlocksSourceTextCharCountFor
        bookmarkIdMetricsCacheMaxTotalSourceParagraphs
        sourceParagraphCount
        blocks

let private tryParseBookmarkId value =
    match value with
    | Some raw ->
        match System.Int32.TryParse(raw: string) with
        | true, parsed -> ValueSome parsed
        | false, _ -> ValueNone
    | None -> ValueNone

let private bookmarkIdMetricsCacheWithoutMatching blocks =
    let mutable cached = Unchecked.defaultof<BookmarkIdMetricsCache>
    if bookmarkIdMetricsCacheByBlocks.TryGetValue(blocks, &cached)
       && bookmarkIdMetricsCacheMatches blocks cached then
        bookmarkIdMetricsCache
        |> List.filter (bookmarkIdMetricsCacheMatches blocks >> not)
    else
        bookmarkIdMetricsCache

let private rememberBookmarkIdMetrics cached =
    if not (bookmarkIdMetricsCacheRetainable cached) then
        let entries = bookmarkIdMetricsCacheWithoutMatching cached.SourceBlocks
        if not (obj.ReferenceEquals(entries, bookmarkIdMetricsCache)) then
            setBookmarkIdMetricsCache entries
    else
        let trim entries =
            let rec loop keptCount keptBlocks keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateBlocks = max 1 candidate.SourceParagraphCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if bookmarkIdMetricsCacheRetainable candidate
                       && keptCount < bookmarkIdMetricsCacheMaxEntries
                       && keptBlocks + candidateBlocks <= bookmarkIdMetricsCacheMaxTotalSourceParagraphs
                       && keptSourceTextChars + candidateSourceTextChars <= bookmarkIdMetricsCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptBlocks + candidateBlocks)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptBlocks keptSourceTextChars acc tail
            loop 0 0 0 [] entries
        (cached :: bookmarkIdMetricsCacheWithoutMatching cached.SourceBlocks)
        |> trim
        |> setBookmarkIdMetricsCache

let private bookmarkIdMetrics doc =
    match tryGetBookmarkIdMetricsCache doc.Blocks with
    | ValueSome cached -> cached
    | ValueNone ->
        let mutable maxExisting = 0
        for block in doc.Blocks do
            iterTextRunsInBlock block (fun run ->
                match tryParseBookmarkId run.Style.BookmarkStartId with
                | ValueSome parsed -> maxExisting <- max maxExisting parsed
                | ValueNone -> ()
                match tryParseBookmarkId run.Style.BookmarkEndId with
                | ValueSome parsed -> maxExisting <- max maxExisting parsed
                | ValueNone -> ())
        let sourceParagraphCount =
            countSourceParagraphsWithLimit bookmarkIdMetricsCacheMaxTotalSourceParagraphs doc.Blocks
        let cached = {
            SourceBlocks = doc.Blocks
            MaxBookmarkId = maxExisting
            SourceParagraphCount = sourceParagraphCount
            SourceTextCharCount = bookmarkIdMetricsSourceTextCharCountFor sourceParagraphCount doc.Blocks
        }
        rememberBookmarkIdMetrics cached
        cached

let private nextBookmarkId doc =
    string ((bookmarkIdMetrics doc).MaxBookmarkId + 1)

let private bookmarkStyleRetainedTextCharCount (bookmarkId: string) (bookmarkName: string) =
    bookmarkId.Length + bookmarkName.Length + bookmarkId.Length

let private rememberBookmarkIdAfterChange
    (sourceDoc: Document)
    (updatedDoc: Document)
    (bookmarkId: string)
    (bookmarkName: string) =
    match tryParseBookmarkId (Some bookmarkId) with
    | ValueNone -> ()
    | ValueSome parsed ->
        let sourceMetrics = bookmarkIdMetrics sourceDoc
        rememberBookmarkIdMetrics {
            SourceBlocks = updatedDoc.Blocks
            MaxBookmarkId = max sourceMetrics.MaxBookmarkId parsed
            SourceParagraphCount = sourceMetrics.SourceParagraphCount
            SourceTextCharCount =
                documentAddCapped
                    sourceMetrics.SourceTextCharCount
                    (bookmarkStyleRetainedTextCharCount bookmarkId bookmarkName)
        }

type private RevisionIdMetricsCache = {
    SourceBlocks: Block list
    MaxRevisionId: int
    SourceParagraphCount: int
    SourceTextCharCount: int
}

let mutable private revisionIdMetricsCache: RevisionIdMetricsCache list = []
let private revisionIdMetricsCacheByBlocks =
    Dictionary<Block list, RevisionIdMetricsCache>(documentInfoBlocksComparer)
let private revisionIdMetricsFastSlots = Array.zeroCreate<RevisionIdMetricsCache> 256
let private revisionIdMetricsCacheMaxEntries = 256
let private revisionIdMetricsCacheMaxTotalSourceParagraphs = sourceParagraphCountCacheLimit
let private revisionIdMetricsCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars

let private revisionIdMetricsCacheMatches blocks cached =
    obj.ReferenceEquals(cached.SourceBlocks, blocks)

let private setRevisionIdMetricsCache entries =
    revisionIdMetricsCache <- entries
    revisionIdMetricsCacheByBlocks.Clear()
    clearArray revisionIdMetricsFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<RevisionIdMetricsCache>
        if not (revisionIdMetricsCacheByBlocks.TryGetValue(entry.SourceBlocks, &existing)) then
            revisionIdMetricsCacheByBlocks.[entry.SourceBlocks] <- entry
            revisionIdMetricsFastSlots.[documentInfoSlot entry.SourceBlocks revisionIdMetricsFastSlots.Length] <- entry

let private tryGetRevisionIdMetricsCache blocks =
    let slot = documentInfoSlot blocks revisionIdMetricsFastSlots.Length
    let fast = revisionIdMetricsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && revisionIdMetricsCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<RevisionIdMetricsCache>
        if revisionIdMetricsCacheByBlocks.TryGetValue(blocks, &cached) then
            revisionIdMetricsFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private revisionIdMetricsCacheRetainable cached =
    cached.SourceParagraphCount <= revisionIdMetricsCacheMaxTotalSourceParagraphs
    && cached.SourceTextCharCount <= revisionIdMetricsCacheMaxTotalSourceTextChars

let private revisionIdMetricsSourceTextCharCountFor sourceParagraphCount blocks =
    documentBlocksSourceTextCharCountFor
        revisionIdMetricsCacheMaxTotalSourceParagraphs
        sourceParagraphCount
        blocks

let private tryParseRevisionId value =
    match value with
    | Some raw ->
        match System.Int32.TryParse(raw: string) with
        | true, parsed -> ValueSome parsed
        | false, _ -> ValueNone
    | None -> ValueNone

let private revisionIdMetricsCacheWithoutMatching blocks =
    let mutable cached = Unchecked.defaultof<RevisionIdMetricsCache>
    if revisionIdMetricsCacheByBlocks.TryGetValue(blocks, &cached)
       && revisionIdMetricsCacheMatches blocks cached then
        revisionIdMetricsCache
        |> List.filter (revisionIdMetricsCacheMatches blocks >> not)
    else
        revisionIdMetricsCache

let private rememberRevisionIdMetrics cached =
    if not (revisionIdMetricsCacheRetainable cached) then
        let entries = revisionIdMetricsCacheWithoutMatching cached.SourceBlocks
        if not (obj.ReferenceEquals(entries, revisionIdMetricsCache)) then
            setRevisionIdMetricsCache entries
    else
        let trim entries =
            let rec loop keptCount keptBlocks keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateBlocks = max 1 candidate.SourceParagraphCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if revisionIdMetricsCacheRetainable candidate
                       && keptCount < revisionIdMetricsCacheMaxEntries
                       && keptBlocks + candidateBlocks <= revisionIdMetricsCacheMaxTotalSourceParagraphs
                       && keptSourceTextChars + candidateSourceTextChars <= revisionIdMetricsCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptBlocks + candidateBlocks)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptBlocks keptSourceTextChars acc tail
            loop 0 0 0 [] entries
        (cached :: revisionIdMetricsCacheWithoutMatching cached.SourceBlocks)
        |> trim
        |> setRevisionIdMetricsCache

let private revisionIdMetrics doc =
    match tryGetRevisionIdMetricsCache doc.Blocks with
    | ValueSome cached -> cached
    | ValueNone ->
        let mutable maxExisting = 0
        for block in doc.Blocks do
            iterTextRunsInBlock block (fun run ->
                match tryParseRevisionId run.Style.RevisionId with
                | ValueSome parsed -> maxExisting <- max maxExisting parsed
                | ValueNone -> ())
        let sourceParagraphCount =
            countSourceParagraphsWithLimit revisionIdMetricsCacheMaxTotalSourceParagraphs doc.Blocks
        let cached = {
            SourceBlocks = doc.Blocks
            MaxRevisionId = maxExisting
            SourceParagraphCount = sourceParagraphCount
            SourceTextCharCount = revisionIdMetricsSourceTextCharCountFor sourceParagraphCount doc.Blocks
        }
        rememberRevisionIdMetrics cached
        cached

let private nextRevisionId doc =
    string ((revisionIdMetrics doc).MaxRevisionId + 1)

let private revisionStyleRetainedTextCharCount
    (revisionType: string)
    (revisionId: string)
    (author: string option)
    (date: string option) =
    revisionType.Length
    + revisionId.Length
    + documentOptionStringCharCount author
    + documentOptionStringCharCount date

let private rememberRevisionIdAfterChange
    (sourceDoc: Document)
    (updatedDoc: Document)
    (revisionType: string)
    (revisionId: string)
    (author: string option)
    (date: string option)
    addedTextChars =
    match tryParseRevisionId (Some revisionId) with
    | ValueNone -> ()
    | ValueSome parsed ->
        let sourceMetrics = revisionIdMetrics sourceDoc
        let delta =
            max 0 addedTextChars
            + revisionStyleRetainedTextCharCount revisionType revisionId author date
        rememberRevisionIdMetrics {
            SourceBlocks = updatedDoc.Blocks
            MaxRevisionId = max sourceMetrics.MaxRevisionId parsed
            SourceParagraphCount = sourceMetrics.SourceParagraphCount
            SourceTextCharCount = documentAddCapped sourceMetrics.SourceTextCharCount delta
        }

let insertCaption label captionText doc =
    let metrics = captionIdMetrics label doc
    let resultNumber = metrics.MaxSequenceResult + 1
    let bookmarkIdNumber = metrics.MaxBookmarkId + 1
    let result = string resultNumber
    let bookmarkId = string bookmarkIdNumber
    let bookmarkName = sprintf "_Ref%s" bookmarkId
    let prefix = label + " "
    let suffix = ": " + captionText
    let fullText = prefix + result + suffix
    let buffer =
        TextBuffer.ofString fullText
        |> TextBuffer.applyStyle prefix.Length result.Length (fun style ->
            { style with
                BookmarkStartId = Some bookmarkId
                BookmarkName = Some bookmarkName
                BookmarkEndId = Some bookmarkId
                FieldCode = Some (sprintf " SEQ %s \\* ARABIC " label)
                FieldResult = Some result })
    let paragraph = {
        Id = nextBlockId ()
        Buffer = buffer
        Metadata = { defaultMetadata with Style = Some "Caption" }
    }
    let inserted = insertTopLevelBlockAfterSelection (Paragraph paragraph) doc
    let sourceParagraphCount =
        if metrics.SourceParagraphCount > captionIdMetricsCacheMaxTotalSourceParagraphs then
            metrics.SourceParagraphCount
        else
            min (captionIdMetricsCacheMaxTotalSourceParagraphs + 1) (metrics.SourceParagraphCount + 1)
    rememberCaptionIdMetrics {
        SourceBlocks = inserted.Blocks
        SourceLabel = label
        MaxSequenceResult = resultNumber
        MaxBookmarkId = bookmarkIdNumber
        LatestCaption = Some (resultNumber, result, bookmarkName)
        SourceParagraphCount = sourceParagraphCount
        SourceTextCharCount =
            documentAddCapped
                metrics.SourceTextCharCount
                (documentParagraphSourceTextCharCount paragraph)
    }
    inserted

let insertCrossReferenceToLatestCaption label doc =
    match (captionIdMetrics label doc).LatestCaption with
    | Some (_, result, bookmarkName) ->
        insertField (sprintf " REF %s \\h " bookmarkName) result doc
    | None -> doc

let private isFieldCodeSeparator c =
    c = ' ' || c = '\t' || c = '\r' || c = '\n'

let private fieldCodeParts (code: string) =
    let parts = ResizeArray<string>()
    let mutable tokenStart = -1
    for i = 0 to code.Length - 1 do
        if isFieldCodeSeparator code.[i] then
            if tokenStart >= 0 then
                parts.Add(code.Substring(tokenStart, i - tokenStart))
                tokenStart <- -1
        elif tokenStart < 0 then
            tokenStart <- i
    if tokenStart >= 0 then
        parts.Add(code.Substring(tokenStart, code.Length - tokenStart))
    resizeArrayToList parts

let private dateFieldFormat (code: string) =
    let formatSwitch = code.IndexOf("\\@")
    if formatSwitch < 0 then "yyyy-MM-dd"
    else
        let firstQuote = code.IndexOf('"', formatSwitch)
        if firstQuote < 0 then "yyyy-MM-dd"
        else
            let secondQuote = code.IndexOf('"', firstQuote + 1)
            if secondQuote <= firstQuote then "yyyy-MM-dd"
            else code.Substring(firstQuote + 1, secondQuote - firstQuote - 1)

let private fieldSwitchValue (switchName: string) (code: string) =
    let switchIndex = code.IndexOf(switchName)
    if switchIndex < 0 then None
    else
        let firstQuote = code.IndexOf('"', switchIndex)
        if firstQuote < 0 then None
        else
            let secondQuote = code.IndexOf('"', firstQuote + 1)
            if secondQuote <= firstQuote then None
            else Some (code.Substring(firstQuote + 1, secondQuote - firstQuote - 1))

let private timeFieldFormat (code: string) =
    // TIME shares DATE's "\@ \"fmt\"" switch; default to a clock format when absent.
    if code.IndexOf("\\@") < 0 then "h:mm tt" else dateFieldFormat code

let private normalizeStyleKey (s: string) =
    let mutable needsNormalize = false
    let mutable i = 0
    while not needsNormalize && i < s.Length do
        let c = s.[i]
        if c = ' ' || c = '-' || System.Char.IsUpper c then
            needsNormalize <- true
        i <- i + 1
    if not needsNormalize then
        s
    else
        let builder = System.Text.StringBuilder(s.Length)
        for c in s do
            if c <> ' ' && c <> '-' then
                builder.Append(System.Char.ToLowerInvariant c) |> ignore
        builder.ToString()

/// Extract the style name a STYLEREF field points at: the first quoted string,
/// else the first non-switch token after the field name (e.g. STYLEREF Heading1,
/// STYLEREF "Heading 1", STYLEREF 1).
let private styleRefStyleName (code: string) =
    let q1 = code.IndexOf('"')
    if q1 >= 0 then
        let q2 = code.IndexOf('"', q1 + 1)
        if q2 > q1 then Some (code.Substring(q1 + 1, q2 - q1 - 1)) else None
    else
        fieldCodeParts code
        |> function
           | _ :: rest -> rest |> List.tryFind (fun t -> not (t.StartsWith("\\")))
           | [] -> None

let private paragraphText para = Rope.toString para.Buffer.Rope

let private tocHeadingRange (code: string) =
    let parseBound fallback (text: string) =
        match System.Int32.TryParse(text: string) with
        | true, value -> max 1 (min 9 value)
        | false, _ -> fallback

    match fieldSwitchValue "\\o" code with
    | None -> 1, 3
    | Some value ->
        let dashIndex = value.IndexOf('-')
        if dashIndex >= 0 && dashIndex = value.LastIndexOf('-') then
            let first = parseBound 1 (value.Substring(0, dashIndex))
            let last = parseBound 3 (value.Substring(dashIndex + 1))
            min first last, max first last
        else
            1, 3

let private tocHeadingLines code (blockPage: BlockId -> int) doc =
    let minLevel, maxLevel = tocHeadingRange code
    let builder = System.Text.StringBuilder()
    let mutable firstLine = true
    let appendLine (text: string) =
        if firstLine then
            firstLine <- false
        else
            builder.Append('\n') |> ignore
        builder.Append text |> ignore
    for block in doc.Blocks do
        match block with
        | Heading (level, para) when level >= minLevel && level <= maxLevel ->
            let pageNumber = blockPage para.Id
            appendLine (String.replicate (max 0 (level - 1)) "  " + paragraphText para + "\t" + string pageNumber)
        | _ -> ()
    builder.ToString()

let private tocCaptionLines label doc =
    let hasSequenceField para =
        para.Buffer.Runs
        |> List.exists (function
            | Text run ->
                match run.Style.FieldCode with
                | Some code -> (fieldCodeParts code) |> List.exists ((=) label)
                | None -> false
            | Image _ -> false)

    doc.Blocks
    |> List.choose (function
        | Paragraph para
        | Heading (_, para) when hasSequenceField para -> Some (paragraphText para)
        | _ -> None)
    |> String.concat "\n"

let private collectSequenceBookmarkResults fieldCodePartsFor doc =
    let counters = Dictionary<string, int>()
    let bookmarks = Dictionary<string, string>()

    let addSequence label bookmarkName =
        let mutable previous = 0
        let next =
            if counters.TryGetValue(label, &previous) then previous + 1
            else 1
        counters.[label] <- next
        bookmarkName |> Option.iter (fun name -> bookmarks.[name] <- string next)

    iterEditableParagraphs doc (fun para ->
        for runItem in para.Buffer.Runs do
            match runItem with
            | Text run ->
                match run.Style.FieldCode |> Option.map fieldCodePartsFor with
                | Some ("SEQ" :: label :: _) -> addSequence label run.Style.BookmarkName
                | _ -> ()
            | Image _ -> ())
    bookmarks

let private collectSequencePageResults fieldCodePartsFor doc fieldPage bookmarkPage =
    let counters = Dictionary<string, int>()
    let results = Dictionary<string, ResizeArray<string>>()
    iterEditableParagraphs doc (fun para ->
        for runItem in para.Buffer.Runs do
            match runItem with
            | Text run ->
                match run.Style.FieldCode |> Option.map fieldCodePartsFor with
                | Some ("SEQ" :: label :: _) ->
                    let mutable previous = 0
                    let next =
                        if counters.TryGetValue(label, &previous) then previous + 1
                        else 1
                    counters.[label] <- next
                    let pageNumber = run.Style.BookmarkName |> Option.map bookmarkPage |> Option.defaultValue (fieldPage para.Id run.Start)
                    let result = string next + "\t" + string pageNumber
                    let mutable labelResults = Unchecked.defaultof<ResizeArray<string>>
                    if results.TryGetValue(label, &labelResults) then
                        labelResults.Add result
                    else
                        let labelResults = ResizeArray<string>()
                        labelResults.Add result
                        results.[label] <- labelResults
                | _ -> ()
            | Image _ -> ())
    results

let refreshFieldsWithPageMap pageMap doc =
    let sequenceCounters = Dictionary<string, int>()
    let sequenceResults = Dictionary<string, ResizeArray<string>>()
    let fieldCodePartsCache = Dictionary<string, string list>()
    let fieldCodePartsFor code =
        let mutable parts = Unchecked.defaultof<string list>
        if fieldCodePartsCache.TryGetValue(code, &parts) then
            parts
        else
            let parts = fieldCodeParts code
            fieldCodePartsCache.[code] <- parts
            parts
    let fieldPageByPosition = Dictionary<struct (BlockId * int), int>()
    for anchor in pageMap.Fields do
        fieldPageByPosition.[struct (anchor.PageFieldBlockId, anchor.PageFieldStartOffset)] <- anchor.PageNumber
    let blockPageById = Dictionary<BlockId, int>()
    for anchor in pageMap.BlockAnchors do
        if not (blockPageById.ContainsKey anchor.PageBlockId) then
            blockPageById.[anchor.PageBlockId] <- anchor.PageNumber
    let bookmarkPageByName = Dictionary<string, int>()
    for anchor in pageMap.Bookmarks do
        if not (bookmarkPageByName.ContainsKey anchor.PageBookmarkName) then
            bookmarkPageByName.[anchor.PageBookmarkName] <- anchor.PageNumber

    let nextSeqResult label =
        let mutable previous = 0
        let next =
            if sequenceCounters.TryGetValue(label, &previous) then previous + 1
            else 1
        sequenceCounters.[label] <- next
        let result = string next
        let mutable results = Unchecked.defaultof<ResizeArray<string>>
        if sequenceResults.TryGetValue(label, &results) then
            results.Add result
        else
            let results = ResizeArray<string>()
            results.Add result
            sequenceResults.[label] <- results
        result

    let fieldPage blockId startOffset =
        let mutable pageNumber = 0
        if fieldPageByPosition.TryGetValue(struct (blockId, startOffset), &pageNumber) then pageNumber else 1

    let blockPage blockId =
        let mutable pageNumber = 0
        if blockPageById.TryGetValue(blockId, &pageNumber) then pageNumber else 1

    let bookmarkPage bookmarkName =
        let mutable pageNumber = 0
        if bookmarkPageByName.TryGetValue(bookmarkName, &pageNumber) then pageNumber else 1

    let mutable sequenceBookmarkResults = ValueNone
    let allBookmarkResults () =
        match sequenceBookmarkResults with
        | ValueSome results -> results
        | ValueNone ->
            let results = collectSequenceBookmarkResults fieldCodePartsFor doc
            sequenceBookmarkResults <- ValueSome results
            results

    let mutable sequencePageResults = ValueNone
    let allSequencePageResults () =
        match sequencePageResults with
        | ValueSome results -> results
        | ValueNone ->
            let results = collectSequencePageResults fieldCodePartsFor doc fieldPage bookmarkPage
            sequencePageResults <- ValueSome results
            results

    let sequencePageLines label =
        let mutable results = Unchecked.defaultof<ResizeArray<string>>
        if (allSequencePageResults ()).TryGetValue(label, &results) then
            stringResizeArrayConcatNewline results
        else
            ""

    let mutable styleRefCache = ValueNone
    let styleRefIndex () =
        match styleRefCache with
        | ValueSome cache -> cache
        | ValueNone ->
            // Ordered (blockId, styleKey, text) for STYLEREF: it shows the text of the
            // nearest paragraph of the referenced style — search backward from the field,
            // then forward (matches Word for inline/body STYLEREF use).
            let orderedBodyParas = ResizeArray<BlockId * string * string>()
            let orderedBodyIndex = Dictionary<BlockId, int>()
            iterEditableParagraphs doc (fun p ->
                if not (orderedBodyIndex.ContainsKey p.Id) then
                    orderedBodyIndex.[p.Id] <- orderedBodyParas.Count
                orderedBodyParas.Add(p.Id, normalizeStyleKey (p.Metadata.Style |> Option.defaultValue ""), paragraphText p))
            let cache = struct (orderedBodyParas, orderedBodyIndex)
            styleRefCache <- ValueSome cache
            cache

    let styleRefResult blockId styleName =
        let struct (orderedBodyParas, orderedBodyIndex) = styleRefIndex ()
        let wanted =
            match System.Int32.TryParse(styleName: string) with
            | true, n -> "heading" + string n
            | false, _ -> normalizeStyleKey styleName
        let mutable result: string option = None
        let tryPick index =
            if result.IsNone && index >= 0 && index < orderedBodyParas.Count then
                let _, key, text = orderedBodyParas.[index]
                if key = wanted && text <> "" then
                    result <- Some text
        let mutable index = 0
        if orderedBodyIndex.TryGetValue(blockId, &index) then
            let mutable i = index - 1
            while result.IsNone && i >= 0 do
                tryPick i
                i <- i - 1
            i <- index
            while result.IsNone && i < orderedBodyParas.Count do
                tryPick i
                i <- i + 1
        else
            let mutable i = 0
            while result.IsNone && i < orderedBodyParas.Count do
                tryPick i
                i <- i + 1
        result

    let resultForField blockId startOffset style text =
        match style.FieldCode |> Option.map (fun code -> code, fieldCodePartsFor code) with
        | Some (code, "DATE" :: _) -> System.DateTime.Now.ToString(dateFieldFormat code)
        | Some (code, "TIME" :: _) -> System.DateTime.Now.ToString(timeFieldFormat code)
        | Some (code, "STYLEREF" :: _) ->
            match styleRefStyleName code with
            | Some name -> styleRefResult blockId name |> Option.defaultValue text
            | None -> text
        | Some (code, "TOC" :: _) ->
            match fieldSwitchValue "\\c" code with
            | Some label -> sequencePageLines label
            | None -> tocHeadingLines code blockPage doc
        | Some (_, "SEQ" :: label :: _) ->
            nextSeqResult label
        | Some (_, "REF" :: bookmarkName :: _) ->
            let mutable result = Unchecked.defaultof<string>
            if (allBookmarkResults ()).TryGetValue(bookmarkName, &result) then result else text
        | Some (_, "PAGE" :: _) -> string (fieldPage blockId startOffset)
        | Some (_, "PAGEREF" :: bookmarkName :: _) -> string (bookmarkPage bookmarkName)
        | Some (_, "NUMPAGES" :: _) -> string pageMap.PageCount
        | _ -> text

    let replacements = Dictionary<BlockId, ParagraphBlock>()
    iterEditableParagraphs doc (fun para ->
        let mutable buffer = para.Buffer
        let mutable shift = 0
        let mutable changed = false
        for run in para.Buffer.Runs do
            match run with
            | Image _ -> ()
            | Text textRun ->
                match textRun.Style.FieldCode with
                | None -> ()
                | Some _ ->
                    let start = textRun.Start + shift
                    let oldText = TextBuffer.slice start textRun.Length buffer
                    let newText = resultForField para.Id textRun.Start textRun.Style oldText
                    if newText <> oldText then
                        let newStyle = { textRun.Style with FieldResult = Some newText }
                        buffer <- buffer |> TextBuffer.delete start textRun.Length |> TextBuffer.insert start newText |> TextBuffer.applyStyle start newText.Length (fun _ -> newStyle)
                        shift <- shift + newText.Length - textRun.Length
                        changed <- true
                    elif textRun.Style.FieldResult <> Some oldText then
                        buffer <- buffer |> TextBuffer.applyStyle start textRun.Length (fun style -> { style with FieldResult = Some oldText })
                        changed <- true
        if changed then
            replacements.[para.Id] <- { para with Buffer = buffer })
    replaceParagraphs doc replacements

type private FieldPresenceCache = {
    SourceBlocks: Block list
    SourceParagraphCount: int
    SourceTextCharCount: int
    HasFields: bool
    RequiresPageMap: bool
}

type private FieldPresenceScan = {
    ScanSourceParagraphCount: int
    ScanSourceTextCharCount: int
    ScanHasFields: bool
    ScanRequiresPageMap: bool
}

let mutable private fieldPresenceCache: FieldPresenceCache list = []
let private fieldPresenceCacheByBlocks = Dictionary<Block list, FieldPresenceCache>(documentInfoBlocksComparer)
let private fieldPresenceFastSlots = Array.zeroCreate<FieldPresenceCache> 256
let private fieldPresenceCacheMaxEntries = 256
let private fieldPresenceCacheMaxTotalSourceParagraphs = sourceParagraphCountCacheLimit
let private fieldPresenceCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars

let private fieldPresenceCacheMatches blocks (cached: FieldPresenceCache) =
    obj.ReferenceEquals(cached.SourceBlocks, blocks)

let private setFieldPresenceCache entries =
    fieldPresenceCache <- entries
    fieldPresenceCacheByBlocks.Clear()
    clearArray fieldPresenceFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<FieldPresenceCache>
        if not (fieldPresenceCacheByBlocks.TryGetValue(entry.SourceBlocks, &existing)) then
            fieldPresenceCacheByBlocks.[entry.SourceBlocks] <- entry
            let slot = documentInfoSlot entry.SourceBlocks fieldPresenceFastSlots.Length
            fieldPresenceFastSlots.[slot] <- entry

let private tryGetFieldPresenceCache blocks =
    let slot = documentInfoSlot blocks fieldPresenceFastSlots.Length
    let fast = fieldPresenceFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && fieldPresenceCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<FieldPresenceCache>
        if fieldPresenceCacheByBlocks.TryGetValue(blocks, &cached) then
            fieldPresenceFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private fieldPresenceCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<FieldPresenceCache>
    if fieldPresenceCacheByBlocks.TryGetValue(blocks, &cached) then
        fieldPresenceCache
        |> List.filter (fieldPresenceCacheMatches blocks >> not)
    else
        fieldPresenceCache

let private removeFieldPresenceCacheForBlocks blocks =
    let entries = fieldPresenceCacheWithoutBlocks blocks
    if not (obj.ReferenceEquals(entries, fieldPresenceCache)) then
        setFieldPresenceCache entries

let private fieldPresenceCacheRetainable (cached: FieldPresenceCache) =
    cached.SourceParagraphCount <= fieldPresenceCacheMaxTotalSourceParagraphs
    && cached.SourceTextCharCount <= fieldPresenceCacheMaxTotalSourceTextChars

let private rememberFieldPresence cached =
    if not (fieldPresenceCacheRetainable cached) then
        removeFieldPresenceCacheForBlocks cached.SourceBlocks
    else
        let trim (entries: FieldPresenceCache list) =
            let rec loop keptCount keptSourceParagraphs keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateSourceParagraphs = max 1 candidate.SourceParagraphCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if fieldPresenceCacheRetainable candidate
                       && keptCount < fieldPresenceCacheMaxEntries
                       && keptSourceParagraphs + candidateSourceParagraphs <= fieldPresenceCacheMaxTotalSourceParagraphs
                       && keptSourceTextChars + candidateSourceTextChars <= fieldPresenceCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptSourceParagraphs + candidateSourceParagraphs)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptSourceParagraphs keptSourceTextChars acc tail
            loop 0 0 0 [] entries
        (cached
         :: fieldPresenceCacheWithoutBlocks cached.SourceBlocks)
        |> trim
        |> setFieldPresenceCache

let private fieldCodeRequiresPageMap code =
    match fieldCodeParts code with
    | "PAGE" :: _
    | "PAGEREF" :: _
    | "NUMPAGES" :: _
    | "TOC" :: _ -> true
    | _ -> false

let private scanFieldPresence doc =
    let maxSourceParagraphs = fieldPresenceCacheMaxTotalSourceParagraphs
    let maxSourceTextChars = fieldPresenceCacheMaxTotalSourceTextChars
    let maxVisitedItems = sourceStructureVisitedItemLimit maxSourceParagraphs
    let sourceTextLimit = documentSourceTextCacheMaxTotalChars + 1
    let mutable sourceParagraphCount = 0
    let mutable sourceTextCharCount = 0
    let mutable visitedItems = 0
    let mutable hasFields = false
    let mutable requiresPageMap = false

    let canCountSource () =
        sourceParagraphCount <= maxSourceParagraphs
        && sourceTextCharCount <= maxSourceTextChars
        && visitedItems <= maxVisitedItems

    let visitSourceItem () =
        visitedItems <- visitedItems + 1

    let addSourceText value =
        sourceTextCharCount <- documentAddCapped sourceTextCharCount value

    let scanParagraphFields para =
        if not requiresPageMap then
            let mutable remaining = para.Buffer.Runs
            while not requiresPageMap && not remaining.IsEmpty do
                match remaining with
                | Text run :: tail ->
                    match run.Style.FieldCode with
                    | Some code ->
                        hasFields <- true
                        if fieldCodeRequiresPageMap code then
                            requiresPageMap <- true
                    | None -> ()
                    remaining <- tail
                | _ :: tail ->
                    remaining <- tail
                | [] -> ()

    let addParagraph scanFields para =
        if canCountSource () then
            visitSourceItem ()
            sourceParagraphCount <- sourceParagraphCount + 1
            addSourceText (documentParagraphSourceTextCharCount para)
        if scanFields then
            scanParagraphFields para

    let addParagraphs scanFields (paragraphs: ParagraphBlock list) =
        let mutable remaining = paragraphs
        while not remaining.IsEmpty do
            match remaining with
            | paragraph :: tail ->
                addParagraph scanFields paragraph
                remaining <- tail
            | [] -> ()

    let addTable table =
        let stack = Stack<TableBlock * bool>()
        stack.Push((table, true))
        while stack.Count > 0 do
            let table, scanTableFields = stack.Pop()
            if scanTableFields || canCountSource () then
                if canCountSource () then
                    visitSourceItem ()
                    addSourceText (documentTableStringCharCount table)
                let mutable rows = table.Rows
                while not rows.IsEmpty do
                    match rows with
                    | row :: rowTail ->
                        if scanTableFields || canCountSource () then
                            if canCountSource () then
                                visitSourceItem ()
                                addSourceText (documentTableRowStringCharCount row)
                            let mutable cells = row.Cells
                            while not cells.IsEmpty do
                                match cells with
                                | cell :: cellTail ->
                                    if scanTableFields || canCountSource () then
                                        if canCountSource () then
                                            visitSourceItem ()
                                            addSourceText (documentTableCellStringCharCount cell)
                                        addParagraphs scanTableFields cell.Content
                                        let mutable nestedTables = cell.NestedTables
                                        while not nestedTables.IsEmpty do
                                            match nestedTables with
                                            | nestedTable :: nestedTail ->
                                                stack.Push((nestedTable, false))
                                                nestedTables <- nestedTail
                                            | [] -> ()
                                    cells <- cellTail
                                | [] -> ()
                        rows <- rowTail
                    | [] -> ()

    let mutable remainingBlocks = doc.Blocks
    while not remainingBlocks.IsEmpty do
        match remainingBlocks with
        | block :: tail ->
            match block with
            | Paragraph paragraph
            | Heading (_, paragraph) ->
                addParagraph true paragraph
            | Table table ->
                addTable table
            | TextBox textBox ->
                if canCountSource () then
                    visitSourceItem ()
                    addSourceText (documentTextBoxStringCharCount textBox)
                addParagraphs true textBox.Content
            remainingBlocks <- tail
        | [] -> ()

    let sourceParagraphCount =
        if visitedItems > maxVisitedItems
           || sourceStructureOverwhelmsParagraphs maxSourceParagraphs sourceParagraphCount visitedItems then
            maxSourceParagraphs + 1
        else
            min sourceParagraphCount (maxSourceParagraphs + 1)

    let sourceTextCharCount =
        if sourceParagraphCount > maxSourceParagraphs
           || visitedItems > maxVisitedItems then
            sourceTextLimit
        else
            min sourceTextCharCount sourceTextLimit

    {
        ScanSourceParagraphCount = sourceParagraphCount
        ScanSourceTextCharCount = sourceTextCharCount
        ScanHasFields = hasFields
        ScanRequiresPageMap = requiresPageMap
    }

let private documentFieldRefreshRequirements doc =
    match tryGetFieldPresenceCache doc.Blocks with
    | ValueSome cached ->
        struct (cached.HasFields, cached.RequiresPageMap)
    | ValueNone ->
        let scan = scanFieldPresence doc
        rememberFieldPresence {
            SourceBlocks = doc.Blocks
            SourceParagraphCount = scan.ScanSourceParagraphCount
            SourceTextCharCount = scan.ScanSourceTextCharCount
            HasFields = scan.ScanHasFields
            RequiresPageMap = scan.ScanRequiresPageMap
        }
        struct (scan.ScanHasFields, scan.ScanRequiresPageMap)

let private emptyPageMap = {
    PageCount = 1
    BlockAnchors = []
    Fields = []
    Bookmarks = []
}

let refreshFields doc =
    match documentFieldRefreshRequirements doc with
    | struct (false, _) -> doc
    | struct (true, true) -> refreshFieldsWithPageMap (buildPageMap doc) doc
    | struct (true, false) -> refreshFieldsWithPageMap emptyPageMap doc

let setCurrentTextBoxStyle x y width height fillColor borderColor borderWidth doc =
    let fillColor = if System.String.IsNullOrWhiteSpace fillColor then None else Some fillColor
    let borderColor = if System.String.IsNullOrWhiteSpace borderColor then None else Some borderColor
    let borderWidth = if borderWidth > 0.0 then Some borderWidth else None
    let selectedBlockId =
        doc.Selection
        |> Option.map (fun sel -> sel.Focus.BlockId)
    let mutable selectedTextBox = None
    let mutable lastTextBox = None
    let mutable remaining = doc.Blocks
    let mutable scannedBlockCount = 0
    while selectedTextBox.IsNone && not remaining.IsEmpty do
        match remaining with
        | TextBox textBox :: tail ->
            scannedBlockCount <- scannedBlockCount + 1
            lastTextBox <- Some textBox
            match selectedBlockId with
            | Some blockId when blockId = textBox.Id -> selectedTextBox <- Some textBox
            | _ -> ()
            remaining <- tail
        | _ :: tail ->
            scannedBlockCount <- scannedBlockCount + 1
            remaining <- tail
        | [] -> ()
    if selectedTextBox.IsNone then
        for block in remaining do
            match block with
            | TextBox textBox -> lastTextBox <- Some textBox
            | _ -> ()
    match selectedTextBox |> Option.orElse lastTextBox with
    | None -> doc
    | Some textBox ->
        let nextTextBox =
            { textBox with
                X = Some x
                Y = Some y
                Width = width
                Height = height
                FillColor = fillColor
                BorderColor = borderColor
                BorderWidth = borderWidth }
        if nextTextBox = textBox then
            doc
        else
            let prefix = ResizeArray<Block>(max 0 scannedBlockCount)
            let rec updateBlocks remaining =
                match remaining with
                | [] -> doc
                | TextBox existing :: tail when existing.Id = textBox.Id ->
                    { doc with Blocks = prependResizeArray prefix (TextBox nextTextBox :: tail) }
                | block :: tail ->
                    prefix.Add block
                    updateBlocks tail
            updateBlocks doc.Blocks

let setSelectedTableCellStyleWithShading backgroundColor shading width borderColor borderWidth verticalAlignment paddingTop paddingRight paddingBottom paddingLeft doc =
    match currentTableCell doc with
    | ValueNone -> doc
    | ValueSome (struct (tableIndex, rowIndex, colIndex)) ->
        updateTableAtIndex tableIndex (fun table ->
            updateTableCell rowIndex colIndex (fun cell ->
                { cell with
                    BackgroundColor = backgroundColor
                    Shading = shading
                    Width = width
                    BorderColor = borderColor
                    BorderWidth = borderWidth
                    VerticalAlignment = verticalAlignment
                    PaddingTop = paddingTop
                    PaddingRight = paddingRight
                    PaddingBottom = paddingBottom
                    PaddingLeft = paddingLeft }) table) doc

let setSelectedTableCellStyle backgroundColor width borderColor borderWidth verticalAlignment paddingTop paddingRight paddingBottom paddingLeft doc =
    setSelectedTableCellStyleWithShading backgroundColor None width borderColor borderWidth verticalAlignment paddingTop paddingRight paddingBottom paddingLeft doc

let private setTableMarginsAtIndex paddingTop paddingRight paddingBottom paddingLeft doc tableIndex =
    let mutable resultDoc = doc
    if tableIndex >= 0 then
        let mutable doneScanning = false
        let mutable index = 0
        let mutable remaining = doc.Blocks
        while not doneScanning && index <= tableIndex && not remaining.IsEmpty do
            match remaining with
            | Table table :: _ when index = tableIndex ->
                if table.PaddingTop = paddingTop
                   && table.PaddingRight = paddingRight
                   && table.PaddingBottom = paddingBottom
                   && table.PaddingLeft = paddingLeft then
                    doneScanning <- true
                else
                    let nextTable: TableBlock = {
                        table with
                            PaddingTop = paddingTop
                            PaddingRight = paddingRight
                            PaddingBottom = paddingBottom
                            PaddingLeft = paddingLeft
                    }
                    resultDoc <- { doc with Blocks = replaceTopLevelBlockAtIndex tableIndex (Table nextTable) doc.Blocks }
                    doneScanning <- true
            | _ :: tail ->
                index <- index + 1
                remaining <- tail
            | [] -> ()
    resultDoc

let private floatOptionEquals (left: float option) (right: float option) =
    match left with
    | Some leftValue ->
        match right with
        | Some rightValue -> leftValue = rightValue
        | None -> false
    | None ->
        match right with
        | Some _ -> false
        | None -> true

let mutable private setCurrentTableCellMarginsNoopCacheSet = false
let mutable private setCurrentTableCellMarginsNoopCacheDoc = Unchecked.defaultof<Document>
let mutable private setCurrentTableCellMarginsNoopCacheTop: float option = None
let mutable private setCurrentTableCellMarginsNoopCacheRight: float option = None
let mutable private setCurrentTableCellMarginsNoopCacheBottom: float option = None
let mutable private setCurrentTableCellMarginsNoopCacheLeft: float option = None

let private setCurrentTableCellMarginsNoopCacheMatches paddingTop paddingRight paddingBottom paddingLeft doc =
    setCurrentTableCellMarginsNoopCacheSet
    && obj.ReferenceEquals(setCurrentTableCellMarginsNoopCacheDoc, doc)
    && floatOptionEquals setCurrentTableCellMarginsNoopCacheTop paddingTop
    && floatOptionEquals setCurrentTableCellMarginsNoopCacheRight paddingRight
    && floatOptionEquals setCurrentTableCellMarginsNoopCacheBottom paddingBottom
    && floatOptionEquals setCurrentTableCellMarginsNoopCacheLeft paddingLeft

let private rememberSetCurrentTableCellMarginsNoop paddingTop paddingRight paddingBottom paddingLeft doc =
    setCurrentTableCellMarginsNoopCacheSet <- true
    setCurrentTableCellMarginsNoopCacheDoc <- doc
    setCurrentTableCellMarginsNoopCacheTop <- paddingTop
    setCurrentTableCellMarginsNoopCacheRight <- paddingRight
    setCurrentTableCellMarginsNoopCacheBottom <- paddingBottom
    setCurrentTableCellMarginsNoopCacheLeft <- paddingLeft
    doc

let setCurrentTableCellMargins paddingTop paddingRight paddingBottom paddingLeft doc =
    if setCurrentTableCellMarginsNoopCacheMatches paddingTop paddingRight paddingBottom paddingLeft doc then
        doc
    else
        let mutable resultDoc = doc
        match doc.Selection with
        | Some sel ->
            let mutable handled = false
            match tryGetCurrentTableCellCache doc.Blocks with
            | ValueSome cached ->
                match tryFindCurrentTableCellEntry cached sel.Focus.BlockId with
                | ValueSome (struct (tableIndex, _, _)) ->
                    resultDoc <- setTableMarginsAtIndex paddingTop paddingRight paddingBottom paddingLeft doc tableIndex
                    handled <- true
                | ValueNone -> ()
            | ValueNone -> ()
            if handled then
                ()
            else
                match currentTableCell doc with
                | ValueNone -> ()
                | ValueSome (struct (tableIndex, _, _)) ->
                    resultDoc <- setTableMarginsAtIndex paddingTop paddingRight paddingBottom paddingLeft doc tableIndex
        | None -> ()
        if obj.ReferenceEquals(resultDoc, doc) then
            rememberSetCurrentTableCellMarginsNoop paddingTop paddingRight paddingBottom paddingLeft doc
        else
            resultDoc

let private setTableMarginsAndSpacingAtIndex paddingTop paddingRight paddingBottom paddingLeft cellSpacing doc tableIndex =
    let mutable resultDoc = doc
    if tableIndex >= 0 then
        let mutable doneScanning = false
        let mutable index = 0
        let mutable remaining = doc.Blocks
        while not doneScanning && index <= tableIndex && not remaining.IsEmpty do
            match remaining with
            | Table table :: _ when index = tableIndex ->
                if table.PaddingTop = paddingTop
                   && table.PaddingRight = paddingRight
                   && table.PaddingBottom = paddingBottom
                   && table.PaddingLeft = paddingLeft
                   && table.CellSpacing = cellSpacing then
                    doneScanning <- true
                else
                    let nextTable: TableBlock = {
                        table with
                            PaddingTop = paddingTop
                            PaddingRight = paddingRight
                            PaddingBottom = paddingBottom
                            PaddingLeft = paddingLeft
                            CellSpacing = cellSpacing
                    }
                    resultDoc <- { doc with Blocks = replaceTopLevelBlockAtIndex tableIndex (Table nextTable) doc.Blocks }
                    doneScanning <- true
            | _ :: tail ->
                index <- index + 1
                remaining <- tail
            | [] -> ()
    resultDoc

let mutable private setCurrentTableCellMarginsAndSpacingNoopCacheSet = false
let mutable private setCurrentTableCellMarginsAndSpacingNoopCacheDoc = Unchecked.defaultof<Document>
let mutable private setCurrentTableCellMarginsAndSpacingNoopCacheTop: float option = None
let mutable private setCurrentTableCellMarginsAndSpacingNoopCacheRight: float option = None
let mutable private setCurrentTableCellMarginsAndSpacingNoopCacheBottom: float option = None
let mutable private setCurrentTableCellMarginsAndSpacingNoopCacheLeft: float option = None
let mutable private setCurrentTableCellMarginsAndSpacingNoopCacheSpacing: float option = None

let private setCurrentTableCellMarginsAndSpacingNoopCacheMatches paddingTop paddingRight paddingBottom paddingLeft cellSpacing doc =
    setCurrentTableCellMarginsAndSpacingNoopCacheSet
    && obj.ReferenceEquals(setCurrentTableCellMarginsAndSpacingNoopCacheDoc, doc)
    && floatOptionEquals setCurrentTableCellMarginsAndSpacingNoopCacheTop paddingTop
    && floatOptionEquals setCurrentTableCellMarginsAndSpacingNoopCacheRight paddingRight
    && floatOptionEquals setCurrentTableCellMarginsAndSpacingNoopCacheBottom paddingBottom
    && floatOptionEquals setCurrentTableCellMarginsAndSpacingNoopCacheLeft paddingLeft
    && floatOptionEquals setCurrentTableCellMarginsAndSpacingNoopCacheSpacing cellSpacing

let private rememberSetCurrentTableCellMarginsAndSpacingNoop paddingTop paddingRight paddingBottom paddingLeft cellSpacing doc =
    setCurrentTableCellMarginsAndSpacingNoopCacheSet <- true
    setCurrentTableCellMarginsAndSpacingNoopCacheDoc <- doc
    setCurrentTableCellMarginsAndSpacingNoopCacheTop <- paddingTop
    setCurrentTableCellMarginsAndSpacingNoopCacheRight <- paddingRight
    setCurrentTableCellMarginsAndSpacingNoopCacheBottom <- paddingBottom
    setCurrentTableCellMarginsAndSpacingNoopCacheLeft <- paddingLeft
    setCurrentTableCellMarginsAndSpacingNoopCacheSpacing <- cellSpacing
    doc

let setCurrentTableCellMarginsAndSpacing paddingTop paddingRight paddingBottom paddingLeft cellSpacing doc =
    if setCurrentTableCellMarginsAndSpacingNoopCacheMatches paddingTop paddingRight paddingBottom paddingLeft cellSpacing doc then
        doc
    else
        let mutable resultDoc = doc
        match doc.Selection with
        | Some sel ->
            let mutable handled = false
            match tryGetCurrentTableCellCache doc.Blocks with
            | ValueSome cached ->
                match tryFindCurrentTableCellEntry cached sel.Focus.BlockId with
                | ValueSome (struct (tableIndex, _, _)) ->
                    resultDoc <- setTableMarginsAndSpacingAtIndex paddingTop paddingRight paddingBottom paddingLeft cellSpacing doc tableIndex
                    handled <- true
                | ValueNone -> ()
            | ValueNone -> ()
            if handled then
                ()
            else
                match currentTableCell doc with
                | ValueNone -> ()
                | ValueSome (struct (tableIndex, _, _)) ->
                    resultDoc <- setTableMarginsAndSpacingAtIndex paddingTop paddingRight paddingBottom paddingLeft cellSpacing doc tableIndex
        | None -> ()
        if obj.ReferenceEquals(resultDoc, doc) then
            rememberSetCurrentTableCellMarginsAndSpacingNoop paddingTop paddingRight paddingBottom paddingLeft cellSpacing doc
        else
            resultDoc

let private headingFontSize level =
    match level with
    | 1 -> 24.0
    | 2 -> 18.0
    | 3 -> 14.0
    | _ -> 11.0

let private hasHeadingStyle level buffer =
    let bold = level >= 1 && level <= 3
    let fontSize = headingFontSize level
    let mutable matches = true
    let mutable runs = buffer.Runs
    while matches && not runs.IsEmpty do
        match runs with
        | Text tr :: tail ->
            matches <- tr.Style.Bold = bold && tr.Style.FontSize = fontSize
            runs <- tail
        | Image _ :: tail ->
            runs <- tail
        | [] -> ()
    matches

let private styleHeadingBuffer level buffer =
    if hasHeadingStyle level buffer then buffer
    else
        let len = TextBuffer.length buffer
        if len = 0 then buffer
        else
            TextBuffer.applyStyle 0 len (fun style ->
                { style with
                    Bold = level >= 1 && level <= 3
                    FontSize = headingFontSize level }) buffer

let mutable private setHeadingLevelNoopCacheSet = false
let mutable private setHeadingLevelNoopCacheDoc = Unchecked.defaultof<Document>
let mutable private setHeadingLevelNoopCacheLevel = 0

let private rememberSetHeadingLevelNoop boundedLevel doc =
    setHeadingLevelNoopCacheSet <- true
    setHeadingLevelNoopCacheDoc <- doc
    setHeadingLevelNoopCacheLevel <- boundedLevel
    doc

let setHeadingLevel level doc =
    let boundedLevel = max 1 (min 3 level)
    if setHeadingLevelNoopCacheSet
       && setHeadingLevelNoopCacheLevel = boundedLevel
       && obj.ReferenceEquals(setHeadingLevelNoopCacheDoc, doc) then
        doc
    else
        match doc.Selection with
        | None -> rememberSetHeadingLevelNoop boundedLevel doc
        | Some sel ->
            match tryFindTopBlockUntil doc sel.Focus.BlockId with
            | ValueSome block ->
                match block with
                | Paragraph para
                | Heading (_, para) ->
                    match block with
                    | Heading (currentLevel, _) when currentLevel = boundedLevel && hasHeadingStyle boundedLevel para.Buffer ->
                        rememberSetHeadingLevelNoop boundedLevel doc
                    | _ ->
                        let styledBuffer = styleHeadingBuffer boundedLevel para.Buffer
                        let styledPara =
                            if obj.ReferenceEquals(styledBuffer, para.Buffer) then para
                            else { para with Buffer = styledBuffer }
                        replaceFirstTopLevelBlock doc para.Id (Heading (boundedLevel, styledPara))
                | _ -> rememberSetHeadingLevelNoop boundedLevel doc
            | ValueNone -> rememberSetHeadingLevelNoop boundedLevel doc

let clearHeading doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match tryFindTopBlockUntil doc sel.Focus.BlockId with
        | ValueSome (Heading (_, para)) -> replaceFirstTopLevelBlock doc para.Id (Paragraph para)
        | _ -> doc

/// Move cursor left
let moveCursorLeft doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.Focus.Offset = 0 -> doc
    | Some sel ->
        let newSel = {
            Anchor = { sel.Focus with Offset = sel.Focus.Offset - 1 }
            Focus = { sel.Focus with Offset = sel.Focus.Offset - 1 }
        }
        { doc with Selection = Some newSel }

/// Move cursor right
let moveCursorRight doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findCursorParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let maxOffset = TextBuffer.length para.Buffer
            if sel.Focus.Offset >= maxOffset then doc
            else
                let newSel = {
                    Anchor = { sel.Focus with Offset = sel.Focus.Offset + 1 }
                    Focus = { sel.Focus with Offset = sel.Focus.Offset + 1 }
                }
                { doc with Selection = Some newSel }

/// Extend selection left (Shift+ArrowLeft)
let extendSelectionLeft doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.Focus.Offset = 0 -> doc
    | Some sel ->
        // Keep anchor, move focus left
        let newFocus = { sel.Focus with Offset = sel.Focus.Offset - 1 }
        { doc with Selection = Some { sel with Focus = newFocus } }

/// Extend selection right (Shift+ArrowRight)
let extendSelectionRight doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        match findCursorParagraph doc sel.Focus.BlockId with
        | None -> doc
        | Some para ->
            let maxOffset = TextBuffer.length para.Buffer
            if sel.Focus.Offset >= maxOffset then doc
            else
                // Keep anchor, move focus right
                let newFocus = { sel.Focus with Offset = sel.Focus.Offset + 1 }
                { doc with Selection = Some { sel with Focus = newFocus } }

/// Get document text (for testing/debugging)
let getText doc =
    match tryGetDocumentTextCache doc.Blocks with
    | ValueSome cached ->
        cached.Text
    | ValueNone ->
        let sb =
            if documentTextCanPreSizeBuilder doc.Blocks then
                let mutable totalLength = 0L
                let mutable paragraphCount = 0L
                iterEditableParagraphs doc (fun para ->
                    totalLength <- totalLength + int64 (TextBuffer.length para.Buffer)
                    paragraphCount <- paragraphCount + 1L)
                let separatorLength = max 0L (paragraphCount - 1L)
                let capacity = min (int64 System.Int32.MaxValue) (totalLength + separatorLength) |> int
                System.Text.StringBuilder(capacity)
            else
                System.Text.StringBuilder()
        let mutable first = true
        let mutable sourceParagraphCount = 0
        iterEditableParagraphs doc (fun para ->
            if sourceParagraphCount <= documentTextCacheMaxTotalSourceParagraphs then
                sourceParagraphCount <- sourceParagraphCount + 1
            if first then
                first <- false
            else
                sb.Append '\n' |> ignore
            TextBuffer.appendToStringBuilder sb para.Buffer)
        let text = sb.ToString()
        let sourceParagraphCount =
            if text.Length > documentTextCacheMaxTotalChars
               || sourceParagraphCount >= documentTextCacheMaxTotalSourceParagraphs then
                documentTextCacheMaxTotalSourceParagraphs + 1
            else
                sourceParagraphCount
        rememberDocumentText {
            SourceBlocks = doc.Blocks
            SourceParagraphCount = sourceParagraphCount
            SourceTextCharCount =
                documentTextSourceTextCharCountFor
                    text.Length
                    sourceParagraphCount
                    doc.Blocks
            Text = text
            TextLength = text.Length
        }
        text

let private iterMatchesInBuffer matchCase (query: string) (buffer: TextBuffer.TextBuffer) (emit: FindMatch -> unit) =
    TextBuffer.iterMatchRanges matchCase query buffer (fun start finish ->
        let matchText =
            if matchCase then query
            else TextBuffer.slice start (finish - start) buffer
        emit { BlockId = ""; Start = start; End = finish; Text = matchText })

let private indexOfQueryInString matchCase (query: string) (text: string) start =
#if FABLE_COMPILER
    if start > text.Length - query.Length then
        -1
    elif matchCase then
        text.IndexOf(query, start)
    else
        let charEqualsIgnoreCase a b =
            a = b || System.Char.ToUpperInvariant a = System.Char.ToUpperInvariant b
        let mutable found = -1
        let mutable index = start
        while found < 0 && index <= text.Length - query.Length do
            let mutable offset = 0
            let mutable matches = true
            while matches && offset < query.Length do
                if not (charEqualsIgnoreCase text.[index + offset] query.[offset]) then
                    matches <- false
                offset <- offset + 1
            if matches then
                found <- index
            index <- index + 1
        found
#else
    let comparison = if matchCase then System.StringComparison.Ordinal else System.StringComparison.OrdinalIgnoreCase
    text.IndexOf(query, start, comparison)
#endif

let private iterMatchesInStringFromFirst matchCase (query: string) (text: string) firstIndex (emit: FindMatch -> unit) =
    let queryLength = query.Length
    let mutable index = firstIndex
    while index >= 0 do
        let finish = index + queryLength
        let matchText =
            if matchCase then query
            else text.Substring(index, queryLength)
        emit { BlockId = ""; Start = index; End = finish; Text = matchText }
        index <- indexOfQueryInString matchCase query text finish

let private iterMatchRangesInStringFromFirst matchCase (query: string) (text: string) firstIndex (emit: int -> int -> unit) =
    let queryLength = query.Length
    let mutable index = firstIndex
    while index >= 0 do
        let finish = index + queryLength
        emit index finish
        index <- indexOfQueryInString matchCase query text finish

let private substringEqualsOrdinal (text: string) start length (value: string) =
    if length <> value.Length then
        false
    else
        let mutable equal = true
        let mutable offset = 0
        while equal && offset < length do
            if text.[start + offset] <> value.[offset] then
                equal <- false
            offset <- offset + 1
        equal

let private allStringMatchesEqualReplacement matchCase (query: string) (replacement: string) (text: string) firstIndex =
    if query.Length <> replacement.Length then
        false
    else
        let mutable allEqual = true
        let mutable index = firstIndex
        while allEqual && index >= 0 do
            allEqual <- substringEqualsOrdinal text index query.Length replacement
            index <- indexOfQueryInString matchCase query text (index + query.Length)
        allEqual

let private bufferRangeEqualsReplacement start finish (replacement: string) (buffer: TextBuffer.TextBuffer) =
    let length = finish - start
    if length <> replacement.Length then
        false
    else
        match TextBuffer.trySliceNoCopyV start length buffer with
        | ValueSome text -> text = replacement
        | ValueNone -> TextBuffer.slice start length buffer = replacement

let private tryReplaceSingleRunTextFromFirst matchCase (query: string) (replacement: string) (text: string) firstIndex (buffer: TextBuffer.TextBuffer) =
    match buffer.Runs with
    | [ Text tr ] when tr.Start = 0 && tr.Length = text.Length ->
        let builder = System.Text.StringBuilder(text.Length)
        let appendTextSlice start finish =
            if finish > start then
#if FABLE_COMPILER
                builder.Append(text.Substring(start, finish - start)) |> ignore
#else
                builder.Append(text, start, finish - start) |> ignore
#endif
        let queryLength = query.Length
        let replacementLength = replacement.Length
        let mutable cursor = 0
        let mutable count = 0
        let mutable index = firstIndex
        while index >= 0 do
            let finish = index + queryLength
            appendTextSlice cursor index
            if replacementLength > 0 then
                builder.Append replacement |> ignore
            cursor <- finish
            count <- count + 1
            index <- indexOfQueryInString matchCase query text finish
        appendTextSlice cursor text.Length
        let replacedText = builder.ToString()
        Some (
            { Rope = RedBat.Word.Rope.ofString replacedText
              Runs = if replacedText.Length > 0 then [ Text { tr with Length = replacedText.Length } ] else [] },
            count)
    | _ ->
        None

let private findMatchesInText matchCase (query: string) (text: string) =
    let matches = ResizeArray<FindMatch>()
    let buffer = TextBuffer.ofString text
    iterMatchesInBuffer matchCase query buffer matches.Add
    resizeArrayToList matches

type private FindTextCacheEntry = {
    Matches: FindMatch array
    mutable MatchesList: FindMatch list option
    ResultTextCharCount: int
}

type private FindTextCache = {
    SourceBlocks: Block list
    SourceParagraphCount: int
    SourceTextCharCount: int
    Entries: Dictionary<struct (string * bool), FindTextCacheEntry>
    Order: Queue<struct (string * bool)>
    mutable TotalResultCount: int
    mutable TotalResultTextCharCount: int
    mutable TotalQueryCharCount: int
}

let mutable private findTextCache: FindTextCache list = []
let private findTextCacheByBlocks = Dictionary<Block list, FindTextCache>(documentInfoBlocksComparer)
let private findTextFastSlots = Array.zeroCreate<FindTextCache> 256
let private findTextCacheMaxEntries = 256
let private findTextCacheMaxSources = 256
let private findTextCacheMaxResultCount = 200000
let private findTextCacheMaxTotalResultCount = 200000
let private findTextCacheMaxTotalResultTextChars = documentSourceTextCacheMaxTotalChars
let private findTextCacheMaxTotalQueryChars = documentSourceTextCacheMaxTotalChars
let private findTextCacheMaxTotalSourceParagraphs = sourceParagraphCountCacheLimit
let private findTextCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars

let private arrayToList (items: 'T array) =
    let mutable result = []
    for i = items.Length - 1 downto 0 do
        result <- items.[i] :: result
    result

let private findTextCacheKeyCharCount (struct (query, _): struct (string * bool)) =
    query.Length

let private findTextCacheResultTextCharCount matchCase (matches: FindMatch array) =
    if matchCase then
        0
    else
        let limit = findTextCacheMaxTotalResultTextChars + 1
        let mutable total = 0
        let mutable index = 0
        while total <= findTextCacheMaxTotalResultTextChars
              && index < matches.Length do
            total <- documentAddCapped total matches.[index].Text.Length
            index <- index + 1
        min total limit

let private findTextCacheEntryList (entry: FindTextCacheEntry) =
    match entry.MatchesList with
    | Some matches -> matches
    | None ->
        let matches = arrayToList entry.Matches
        entry.MatchesList <- Some matches
        matches

let private findTextSourceParagraphCountFor resultCount resultTextCharCount keyCharCount blocks =
    if resultCount > findTextCacheMaxResultCount
       || resultTextCharCount > findTextCacheMaxTotalResultTextChars
       || keyCharCount > findTextCacheMaxTotalQueryChars then
        findTextCacheMaxTotalSourceParagraphs + 1
    else
        countSourceParagraphs blocks

let private findTextSourceTextCharCountFor resultCount resultTextCharCount keyCharCount sourceParagraphCount blocks =
    if resultCount > findTextCacheMaxResultCount
       || resultTextCharCount > findTextCacheMaxTotalResultTextChars
       || keyCharCount > findTextCacheMaxTotalQueryChars
       || sourceParagraphCount > findTextCacheMaxTotalSourceParagraphs then
        findTextCacheMaxTotalSourceTextChars + 1
    else
        documentBlocksSourceTextCharCountFor
            findTextCacheMaxTotalSourceParagraphs
            sourceParagraphCount
            blocks

let private findTextCacheMatches blocks (cached: FindTextCache) =
    obj.ReferenceEquals(cached.SourceBlocks, blocks)

let private findTextCacheRetainable (cached: FindTextCache) =
    cached.Entries.Count > 0
    && cached.Entries.Count <= findTextCacheMaxEntries
    && cached.TotalResultCount <= findTextCacheMaxTotalResultCount
    && cached.TotalResultTextCharCount <= findTextCacheMaxTotalResultTextChars
    && cached.TotalQueryCharCount <= findTextCacheMaxTotalQueryChars
    && cached.SourceParagraphCount <= findTextCacheMaxTotalSourceParagraphs
    && cached.SourceTextCharCount <= findTextCacheMaxTotalSourceTextChars

let private setFindTextCache entries =
    findTextCache <- entries
    findTextCacheByBlocks.Clear()
    clearArray findTextFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<FindTextCache>
        if not (findTextCacheByBlocks.TryGetValue(entry.SourceBlocks, &existing)) then
            findTextCacheByBlocks.[entry.SourceBlocks] <- entry
            let slot = documentInfoSlot entry.SourceBlocks findTextFastSlots.Length
            findTextFastSlots.[slot] <- entry

let private tryGetFindTextCache blocks =
    let slot = documentInfoSlot blocks findTextFastSlots.Length
    let fast = findTextFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && findTextCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<FindTextCache>
        if findTextCacheByBlocks.TryGetValue(blocks, &cached) then
            findTextFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private findTextCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<FindTextCache>
    if findTextCacheByBlocks.TryGetValue(blocks, &cached) then
        findTextCache
        |> List.filter (findTextCacheMatches blocks >> not)
    else
        findTextCache

let private removeFindTextCacheForBlocks blocks =
    let entries = findTextCacheWithoutBlocks blocks
    if not (obj.ReferenceEquals(entries, findTextCache)) then
        setFindTextCache entries

let private rememberFindTextCache cached =
    if not (findTextCacheRetainable cached) then
        removeFindTextCacheForBlocks cached.SourceBlocks
    else
        let mutable keptCount = 1
        let mutable keptResults = cached.TotalResultCount
        let mutable keptResultTextChars = cached.TotalResultTextCharCount
        let mutable keptQueryChars = cached.TotalQueryCharCount
        let mutable keptBlocks = max 1 cached.SourceParagraphCount
        let mutable keptSourceTextChars = max 1 cached.SourceTextCharCount
        let retained =
            findTextCacheWithoutBlocks cached.SourceBlocks
            |> List.choose (fun candidate ->
                let candidateResults = candidate.TotalResultCount
                let candidateResultTextChars = candidate.TotalResultTextCharCount
                let candidateQueryChars = max 1 candidate.TotalQueryCharCount
                let candidateBlocks = max 1 candidate.SourceParagraphCount
                let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                if findTextCacheRetainable candidate
                   && keptCount < findTextCacheMaxSources
                   && keptResults + candidateResults <= findTextCacheMaxTotalResultCount
                   && keptResultTextChars + candidateResultTextChars <= findTextCacheMaxTotalResultTextChars
                   && keptQueryChars + candidateQueryChars <= findTextCacheMaxTotalQueryChars
                   && keptBlocks + candidateBlocks <= findTextCacheMaxTotalSourceParagraphs
                   && keptSourceTextChars + candidateSourceTextChars <= findTextCacheMaxTotalSourceTextChars then
                    keptCount <- keptCount + 1
                    keptResults <- keptResults + candidateResults
                    keptResultTextChars <- keptResultTextChars + candidateResultTextChars
                    keptQueryChars <- keptQueryChars + candidateQueryChars
                    keptBlocks <- keptBlocks + candidateBlocks
                    keptSourceTextChars <- keptSourceTextChars + candidateSourceTextChars
                    Some candidate
                else
                    None)
        cached :: retained
        |> setFindTextCache

let private cachedFindTextMatches query matchCase doc =
    let key = struct (query, matchCase)
    match tryGetFindTextCache doc.Blocks with
    | ValueSome cached when findTextCacheRetainable cached ->
        let mutable entry = Unchecked.defaultof<FindTextCacheEntry>
        if cached.Entries.TryGetValue(key, &entry) then ValueSome entry else ValueNone
    | ValueSome cached ->
        rememberFindTextCache cached
        ValueNone
    | ValueNone -> ValueNone

let private removeFindTextCacheEntry key (cache: FindTextCache) =
    let mutable existing = Unchecked.defaultof<FindTextCacheEntry>
    if cache.Entries.TryGetValue(key, &existing) then
        cache.TotalResultCount <- max 0 (cache.TotalResultCount - existing.Matches.Length)
        cache.TotalResultTextCharCount <- max 0 (cache.TotalResultTextCharCount - existing.ResultTextCharCount)
        cache.TotalQueryCharCount <- max 0 (cache.TotalQueryCharCount - findTextCacheKeyCharCount key)
        cache.Entries.Remove key |> ignore

let private addFindTextCacheEntry key entry (cache: FindTextCache) =
    let mutable existing = Unchecked.defaultof<FindTextCacheEntry>
    let keyCharCount = findTextCacheKeyCharCount key
    if entry.Matches.Length > findTextCacheMaxResultCount
       || entry.ResultTextCharCount > findTextCacheMaxTotalResultTextChars
       || keyCharCount > findTextCacheMaxTotalQueryChars then
        removeFindTextCacheEntry key cache
    else
        let alreadyCached = cache.Entries.TryGetValue(key, &existing)
        if alreadyCached then
            cache.TotalResultCount <- max 0 (cache.TotalResultCount - existing.Matches.Length)
            cache.TotalResultTextCharCount <- max 0 (cache.TotalResultTextCharCount - existing.ResultTextCharCount)
        else
            cache.Order.Enqueue key
            cache.TotalQueryCharCount <- cache.TotalQueryCharCount + keyCharCount
        cache.Entries.[key] <- entry
        cache.TotalResultCount <- cache.TotalResultCount + entry.Matches.Length
        cache.TotalResultTextCharCount <- cache.TotalResultTextCharCount + entry.ResultTextCharCount
        while (cache.Entries.Count > findTextCacheMaxEntries
               || cache.TotalResultCount > findTextCacheMaxTotalResultCount
               || cache.TotalResultTextCharCount > findTextCacheMaxTotalResultTextChars
               || cache.TotalQueryCharCount > findTextCacheMaxTotalQueryChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> key then
                removeFindTextCacheEntry oldest cache
        if cache.Entries.Count > findTextCacheMaxEntries
           || cache.TotalResultCount > findTextCacheMaxTotalResultCount
           || cache.TotalResultTextCharCount > findTextCacheMaxTotalResultTextChars
           || cache.TotalQueryCharCount > findTextCacheMaxTotalQueryChars then
            let current = cache.Entries.[key]
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue key
            cache.Entries.[key] <- current
            cache.TotalResultCount <- current.Matches.Length
            cache.TotalResultTextCharCount <- current.ResultTextCharCount
            cache.TotalQueryCharCount <- keyCharCount

let private rememberFindTextMatchesWithSourceMetrics query matchCase doc matches matchesList sourceParagraphCountOpt sourceTextCharCountOpt =
    let key = struct (query, matchCase)
    let keyCharCount = findTextCacheKeyCharCount key
    let entry = {
        Matches = matches
        MatchesList = matchesList
        ResultTextCharCount = findTextCacheResultTextCharCount matchCase matches
    }
    match tryGetFindTextCache doc.Blocks with
    | ValueSome cached ->
        addFindTextCacheEntry key entry cached
        rememberFindTextCache cached
    | ValueNone ->
        let entries = Dictionary<struct (string * bool), FindTextCacheEntry>()
        let order = Queue<struct (string * bool)>()
        let sourceParagraphCount =
            match sourceParagraphCountOpt with
            | ValueSome count -> count
            | ValueNone ->
                findTextSourceParagraphCountFor
                    matches.Length
                    entry.ResultTextCharCount
                    keyCharCount
                    doc.Blocks
        let sourceTextCharCount =
            match sourceTextCharCountOpt with
            | ValueSome count -> count
            | ValueNone ->
                findTextSourceTextCharCountFor
                    matches.Length
                    entry.ResultTextCharCount
                    keyCharCount
                    sourceParagraphCount
                    doc.Blocks
        let cache = {
            SourceBlocks = doc.Blocks
            SourceParagraphCount = sourceParagraphCount
            SourceTextCharCount = sourceTextCharCount
            Entries = entries
            Order = order
            TotalResultCount = 0
            TotalResultTextCharCount = 0
            TotalQueryCharCount = 0
        }
        addFindTextCacheEntry key entry cache
        rememberFindTextCache cache

let private rememberFindTextMatches query matchCase doc matches matchesList =
    rememberFindTextMatchesWithSourceMetrics query matchCase doc matches matchesList ValueNone ValueNone

let private rememberEmptyFindTextMatches query matchCase doc =
    rememberFindTextMatches query matchCase doc [||] (Some [])

let private rememberEmptyFindTextMatchesWithSourceMetrics query matchCase doc sourceParagraphCount sourceTextCharCount =
    rememberFindTextMatchesWithSourceMetrics
        query
        matchCase
        doc
        [||]
        (Some [])
        (ValueSome sourceParagraphCount)
        (ValueSome sourceTextCharCount)

let private findTextMatchesArray query matchCase doc =
    if System.String.IsNullOrEmpty query then
        [||]
    else
        match cachedFindTextMatches query matchCase doc with
        | ValueSome cached ->
            cached.Matches
        | _ ->
            let matches = ResizeArray<FindMatch>()
            iterEditableParagraphs doc (fun para ->
                match TextBuffer.trySliceNoCopyV 0 (TextBuffer.length para.Buffer) para.Buffer with
                | ValueSome text when query.Length <= text.Length ->
                    let firstIndex = indexOfQueryInString matchCase query text 0
                    if firstIndex >= 0 then
                        iterMatchesInStringFromFirst matchCase query text firstIndex (fun matchInfo ->
                            matches.Add { matchInfo with BlockId = para.Id })
                | ValueSome _ -> ()
                | ValueNone ->
                    iterMatchesInBuffer matchCase query para.Buffer (fun matchInfo ->
                        matches.Add { matchInfo with BlockId = para.Id }))
            let result = matches.ToArray()
            rememberFindTextMatches query matchCase doc result None
            result

let findTextArray query matchCase doc =
    findTextMatchesArray query matchCase doc

let findText query matchCase doc =
    if System.String.IsNullOrEmpty query then
        []
    else
        match cachedFindTextMatches query matchCase doc with
        | ValueSome cached ->
            findTextCacheEntryList cached
        | _ ->
            let matches = findTextMatchesArray query matchCase doc
            match cachedFindTextMatches query matchCase doc with
            | ValueSome cached ->
                findTextCacheEntryList cached
            | _ -> arrayToList matches

let selectFindMatch (matchInfo: FindMatch) doc =
    match doc.Selection with
    | Some current
        when current.Anchor.BlockId = matchInfo.BlockId
             && current.Anchor.Offset = matchInfo.Start
             && current.Focus.BlockId = matchInfo.BlockId
             && current.Focus.Offset = matchInfo.End ->
        doc
    | _ ->
        { doc with
            Selection =
                Some {
                    Anchor = { BlockId = matchInfo.BlockId; Offset = matchInfo.Start }
                    Focus = { BlockId = matchInfo.BlockId; Offset = matchInfo.End }
                } }

let replaceAllText query replacement matchCase doc =
    if query = "" then
        (doc, 0)
    elif matchCase && query = replacement then
        (doc, 0)
    elif
        match cachedFindTextMatches query matchCase doc with
        | ValueSome cached -> cached.Matches.Length = 0
        | ValueNone -> false
    then
        (doc, 0)
    else
        let mutable foundAnyMatch = false

        let tryReplaceParagraph (para: ParagraphBlock) =
            match TextBuffer.trySliceNoCopyV 0 (TextBuffer.length para.Buffer) para.Buffer with
            | ValueSome text when query.Length <= text.Length ->
                let firstIndex = indexOfQueryInString matchCase query text 0
                if firstIndex >= 0 then
                    foundAnyMatch <- true
                    if allStringMatchesEqualReplacement matchCase query replacement text firstIndex then
                        ValueNone
                    else
                        match tryReplaceSingleRunTextFromFirst matchCase query replacement text firstIndex para.Buffer with
                        | Some (buffer, matchCount) ->
                            ValueSome ({ para with Buffer = buffer }, matchCount)
                        | None ->
                            let estimate =
                                if query.Length = 0 then 4
                                else max 4 (min 256 (text.Length / query.Length))
                            let matches = ResizeArray<struct (int * int)>(estimate)
                            iterMatchRangesInStringFromFirst matchCase query text firstIndex (fun start finish ->
                                matches.Add(struct (start, finish)))
                            ValueSome ({ para with Buffer = TextBuffer.replaceRanges matches replacement para.Buffer }, matches.Count)
                else
                    ValueNone
            | ValueSome _ -> ValueNone
            | ValueNone ->
                let mutable matches = Unchecked.defaultof<ResizeArray<struct (int * int)>>
                let mutable changesText = false
                let addMatch start finish =
                    foundAnyMatch <- true
                    if obj.ReferenceEquals(matches, null) then
                        let estimate =
                            if query.Length = 0 then 4
                            else max 4 (min 256 (TextBuffer.length para.Buffer / query.Length))
                        matches <- ResizeArray<struct (int * int)>(estimate)
                    matches.Add(struct (start, finish))
                    if not changesText && not (bufferRangeEqualsReplacement start finish replacement para.Buffer) then
                        changesText <- true
                TextBuffer.iterMatchRanges matchCase query para.Buffer addMatch
                if obj.ReferenceEquals(matches, null) || not changesText then
                    ValueNone
                else
                    ValueSome ({ para with Buffer = TextBuffer.replaceRanges matches replacement para.Buffer }, matches.Count)

        let replaceGeneric () =
            let mutable count = 0
            let mutable replacements = Unchecked.defaultof<Dictionary<BlockId, ParagraphBlock>>

            iterEditableParagraphs doc (fun para ->
                match tryReplaceParagraph para with
                | ValueSome (para, matchCount) ->
                    count <- count + matchCount
                    if obj.ReferenceEquals(replacements, null) then
                        replacements <- Dictionary<BlockId, ParagraphBlock>()
                    replacements.[para.Id] <- para
                | ValueNone -> ())

            if not foundAnyMatch then
                rememberEmptyFindTextMatches query matchCase doc

            let doc =
                if obj.ReferenceEquals(replacements, null) then doc
                else replaceParagraphs doc replacements
            (doc, count)

        let mutable count = 0
        let mutable changed = false
        let mutable nestedBlock = false
        let mutable prefixCount = 0
        let mutable blocks = Unchecked.defaultof<ResizeArray<Block>>
        let mutable pendingUnchangedBlocks = Unchecked.defaultof<ResizeArray<Block>>
        let mutable tailAfterLastChange: Block list = []
        let mutable trackEmptyFindSourceMetrics = true
        let mutable emptyFindSourceParagraphCount = 0
        let mutable emptyFindSourceTextCharCount = 0

        let trackEmptyFindParagraphSource para =
            if trackEmptyFindSourceMetrics then
                if emptyFindSourceParagraphCount <= findTextCacheMaxTotalSourceParagraphs then
                    emptyFindSourceParagraphCount <- emptyFindSourceParagraphCount + 1
                if emptyFindSourceTextCharCount <= findTextCacheMaxTotalSourceTextChars then
                    emptyFindSourceTextCharCount <-
                        documentAddCapped
                            emptyFindSourceTextCharCount
                            (documentParagraphSourceTextCharCount para)

        let stopEmptyFindSourceTracking () =
            trackEmptyFindSourceMetrics <- false

        let rememberTopLevelEmptyFind () =
            let sourceParagraphCount =
                if emptyFindSourceParagraphCount > findTextCacheMaxTotalSourceParagraphs then
                    findTextCacheMaxTotalSourceParagraphs + 1
                else
                    emptyFindSourceParagraphCount
            let sourceTextCharCount =
                if sourceParagraphCount > findTextCacheMaxTotalSourceParagraphs then
                    findTextCacheMaxTotalSourceTextChars + 1
                else
                    min emptyFindSourceTextCharCount (findTextCacheMaxTotalSourceTextChars + 1)
            rememberEmptyFindTextMatchesWithSourceMetrics
                query
                matchCase
                doc
                sourceParagraphCount
                sourceTextCharCount

        let copyPrefix () =
            blocks <- ResizeArray<Block>()
            let mutable copied = 0
            let mutable prefix = doc.Blocks
            while copied < prefixCount do
                match prefix with
                | block :: tail ->
                    blocks.Add block
                    prefix <- tail
                    copied <- copied + 1
                | [] ->
                    copied <- prefixCount

        let ensureChanged () =
            if not changed then
                changed <- true
                copyPrefix ()

        let flushPendingUnchangedBlocks () =
            if not (obj.ReferenceEquals(pendingUnchangedBlocks, null)) && pendingUnchangedBlocks.Count > 0 then
                blocks.AddRange pendingUnchangedBlocks
                pendingUnchangedBlocks.Clear()

        let addChangedBlock block tail =
            if changed then
                flushPendingUnchangedBlocks ()
            else
                ensureChanged ()
            blocks.Add block
            tailAfterLastChange <- tail

        let addUnchangedBlock block =
            if changed then
                if obj.ReferenceEquals(pendingUnchangedBlocks, null) then
                    pendingUnchangedBlocks <- ResizeArray<Block>()
                pendingUnchangedBlocks.Add block
            else
                prefixCount <- prefixCount + 1

        let mutable remainingBlocks = doc.Blocks
        while not nestedBlock && not remainingBlocks.IsEmpty do
            match remainingBlocks with
            | Paragraph para :: tail ->
                trackEmptyFindParagraphSource para
                match tryReplaceParagraph para with
                | ValueSome (para, matchCount) ->
                    stopEmptyFindSourceTracking ()
                    count <- count + matchCount
                    addChangedBlock (Paragraph para) tail
                | ValueNone ->
                    if foundAnyMatch then
                        stopEmptyFindSourceTracking ()
                    addUnchangedBlock (Paragraph para)
                remainingBlocks <- tail
            | Heading (level, para) :: tail ->
                trackEmptyFindParagraphSource para
                match tryReplaceParagraph para with
                | ValueSome (para, matchCount) ->
                    stopEmptyFindSourceTracking ()
                    count <- count + matchCount
                    addChangedBlock (Heading (level, para)) tail
                | ValueNone ->
                    if foundAnyMatch then
                        stopEmptyFindSourceTracking ()
                    addUnchangedBlock (Heading (level, para))
                remainingBlocks <- tail
            | _ ->
                nestedBlock <- true

        if nestedBlock then
            replaceGeneric ()
        elif changed then
            { doc with Blocks = prependResizeArray blocks tailAfterLastChange }, count
        else
            if not foundAnyMatch then
                rememberTopLevelEmptyFind ()
            doc, 0

/// Set selection range
let setSelection anchor focus doc =
    match doc.Selection with
    | Some current when current.Anchor = anchor && current.Focus = focus ->
        doc
    | _ ->
        { doc with Selection = Some { Anchor = anchor; Focus = focus } }

/// Select range in current block
let selectRange startOffset endOffset doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        setSelection { sel.Focus with Offset = startOffset } { sel.Focus with Offset = endOffset } doc

/// Select all text in document (all blocks)
let selectAll doc =
    match tryGetSelectAllDocument doc with
    | ValueSome cached -> cached
    | ValueNone ->
        let singleParagraphAlreadySelected (para: ParagraphBlock) (current: Selection) =
            current.Anchor.BlockId = para.Id &&
            current.Anchor.Offset = 0 &&
            current.Focus.BlockId = para.Id &&
            current.Focus.Offset = TextBuffer.length para.Buffer

        match doc.Blocks, doc.Selection with
        | [ Paragraph para ], Some current when singleParagraphAlreadySelected para current -> doc
        | [ Heading (_, para) ], Some current when singleParagraphAlreadySelected para current -> doc
        | _ ->
            let firstPara, lastPara = editableParagraphEndpoints doc
            match firstPara, lastPara with
            | ValueSome firstPara, ValueSome lastPara ->
                let lastLen = TextBuffer.length lastPara.Buffer
                match doc.Selection with
                | Some current
                    when current.Anchor.BlockId = firstPara.Id &&
                         current.Anchor.Offset = 0 &&
                         current.Focus.BlockId = lastPara.Id &&
                         current.Focus.Offset = lastLen ->
                    doc
                | _ ->
                    let result =
                        { doc with
                            Selection =
                                Some {
                                    Anchor = { BlockId = firstPara.Id; Offset = 0 }
                                    Focus = { BlockId = lastPara.Id; Offset = lastLen }
                                } }
                    rememberSelectAllDocument doc result
                    result
            | _ -> doc

let private collectSelectionBufferReplacements
    (doc: Document)
    (startPos: Position)
    (endPos: Position)
    (blockIds: BlockId list)
    (formatBuffer: int -> int -> RedBat.Word.TextBuffer.TextBuffer -> RedBat.Word.TextBuffer.TextBuffer) =
    let blockCount = blockIds.Length
    let selectedIndexes = Dictionary<BlockId, int>(blockCount)
    let mutable i = 0
    for blockId in blockIds do
        selectedIndexes.[blockId] <- i
        i <- i + 1

    let replacements = Dictionary<BlockId, ParagraphBlock>()
    let processParagraph (para: ParagraphBlock) =
        let mutable index = 0
        if selectedIndexes.TryGetValue(para.Id, &index) then
            let bufferLen = TextBuffer.length para.Buffer
            let formatStart, formatLen =
                if blockCount = 1 then
                    startPos.Offset, endPos.Offset - startPos.Offset
                elif index = 0 then
                    startPos.Offset, bufferLen - startPos.Offset
                elif index = blockCount - 1 then
                    0, endPos.Offset
                else
                    0, bufferLen
            if formatLen > 0 then
                let newBuffer = formatBuffer formatStart formatLen para.Buffer
                if not (obj.ReferenceEquals(newBuffer, para.Buffer)) then
                    replacements.[para.Id] <- { para with Buffer = newBuffer }
            true
        else
            false

    if selectedIndexes.Count = blockCount then
        let mutable seen = 0
        iterEditableParagraphsUntil doc (fun para ->
            if processParagraph para then
                seen <- seen + 1
            seen < blockCount)
    else
        iterEditableParagraphs doc (fun para -> processParagraph para |> ignore)

    replacements

let private tryApplyTopLevelSelectionBufferFormat
    (doc: Document)
    (sel: Selection)
    (formatBuffer: int -> int -> RedBat.Word.TextBuffer.TextBuffer -> RedBat.Word.TextBuffer.TextBuffer) =
    if sel.Anchor.BlockId = sel.Focus.BlockId then
        ValueNone
    else
        let blocks = ResizeArray<Block>()
        let mutable changed = false
        let mutable remaining = doc.Blocks
        let mutable collecting = false
        let mutable valid = true
        let mutable foundEnd = false
        let mutable startPos = Unchecked.defaultof<Position>
        let mutable endPos = Unchecked.defaultof<Position>
        let mutable endTarget = ""

        let addFormattedBlock
            (block: Block)
            (para: ParagraphBlock)
            (wrap: ParagraphBlock -> Block)
            isFirst
            isLast =
            let bufferLen = TextBuffer.length para.Buffer
            let formatStart, formatLen =
                if isFirst then
                    startPos.Offset, bufferLen - startPos.Offset
                elif isLast then
                    0, endPos.Offset
                else
                    0, bufferLen

            if formatLen > 0 then
                let newBuffer = formatBuffer formatStart formatLen para.Buffer
                if obj.ReferenceEquals(newBuffer, para.Buffer) then
                    blocks.Add block
                else
                    changed <- true
                    blocks.Add(wrap { para with Buffer = newBuffer })
            else
                blocks.Add block

        let startCollecting (block: Block) (para: ParagraphBlock) (wrap: ParagraphBlock -> Block) =
            let anchorMatch = para.Id = sel.Anchor.BlockId
            let focusMatch = para.Id = sel.Focus.BlockId
            if anchorMatch || focusMatch then
                collecting <- true
                if anchorMatch then
                    startPos <- sel.Anchor
                    endPos <- sel.Focus
                    endTarget <- sel.Focus.BlockId
                else
                    startPos <- sel.Focus
                    endPos <- sel.Anchor
                    endTarget <- sel.Anchor.BlockId
                addFormattedBlock block para wrap true false
            else
                blocks.Add block

        let continueCollecting (block: Block) (para: ParagraphBlock) (wrap: ParagraphBlock -> Block) =
            let isLast = para.Id = endTarget
            addFormattedBlock block para wrap false isLast
            if isLast then
                foundEnd <- true

        while valid && not foundEnd && not remaining.IsEmpty do
            match remaining with
            | block :: tail ->
                if collecting then
                    match block with
                    | Paragraph para ->
                        continueCollecting block para Paragraph
                    | Heading (level, para) ->
                        continueCollecting block para (fun para -> Heading (level, para))
                    | Table _
                    | TextBox _ ->
                        valid <- false
                else
                    match block with
                    | Paragraph para ->
                        startCollecting block para Paragraph
                    | Heading (level, para) ->
                        startCollecting block para (fun para -> Heading (level, para))
                    | Table _
                    | TextBox _ ->
                        blocks.Add block
                remaining <- tail
            | [] -> ()

        if valid && foundEnd then
            if changed then
                ValueSome { doc with Blocks = prependResizeArray blocks remaining }
            else
                ValueSome doc
        else
            ValueNone

let private applySelectionBufferFormat
    (doc: Document)
    (sel: Selection)
    (formatBuffer: int -> int -> RedBat.Word.TextBuffer.TextBuffer -> RedBat.Word.TextBuffer.TextBuffer) =
    match tryApplyTopLevelSelectionBufferFormat doc sel formatBuffer with
    | ValueSome doc ->
        doc
    | ValueNone ->
        let (startPos, endPos, blockIds) = getNormalizedRange sel doc
        let blockCount = blockIds.Length
        if blockCount = 0 then
            doc
        else
            let replacements = collectSelectionBufferReplacements doc startPos endPos blockIds formatBuffer
            replaceParagraphs doc replacements

/// Apply formatting to selection
let applyFormattingToSelection formatFn doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc  // No selection
    | Some sel when sel.Anchor.BlockId = sel.Focus.BlockId ->
        let formatStart = min sel.Anchor.Offset sel.Focus.Offset
        let formatLen = abs (sel.Focus.Offset - sel.Anchor.Offset)
        if formatLen <= 0 then
            doc
        else
            let mutable changed = false
            let updated =
                updateParagraph doc sel.Anchor.BlockId (fun para ->
                    let newBuffer = formatFn formatStart formatLen para.Buffer
                    if obj.ReferenceEquals(newBuffer, para.Buffer) then
                        para
                    else
                        changed <- true
                        { para with Buffer = newBuffer })
            if changed then updated else doc
    | Some sel ->
        applySelectionBufferFormat doc sel formatFn

/// Toggle bold on selection
let toggleBold doc =
    applyFormattingToSelection TextBuffer.toggleBold doc

/// Toggle italic on selection
let toggleItalic doc =
    applyFormattingToSelection TextBuffer.toggleItalic doc

/// Toggle underline on selection
let toggleUnderline doc =
    applyFormattingToSelection TextBuffer.toggleUnderline doc

/// Toggle all caps on selection
let toggleAllCaps doc =
    applyFormattingToSelection TextBuffer.toggleAllCaps doc

/// Toggle small caps on selection
let toggleSmallCaps doc =
    applyFormattingToSelection TextBuffer.toggleSmallCaps doc

/// Toggle strikethrough on selection
let toggleStrike doc =
    applyFormattingToSelection TextBuffer.toggleStrike doc

/// Clear direct character formatting on selection.
let clearFormatting doc =
    applyFormattingToSelection TextBuffer.clearFormatting doc

/// Set vertical alignment on selection
let setVertAlign align doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc
    | Some sel when sel.Anchor.BlockId = sel.Focus.BlockId ->
        let formatStart = min sel.Anchor.Offset sel.Focus.Offset
        let formatLen = abs (sel.Focus.Offset - sel.Anchor.Offset)
        if formatLen <= 0 then
            doc
        else
            match doc.Blocks with
            | [ Paragraph para ] when para.Id = sel.Anchor.BlockId ->
                let newBuffer = TextBuffer.setVertAlign formatStart formatLen align para.Buffer
                if obj.ReferenceEquals(newBuffer, para.Buffer) then
                    doc
                else
                    replaceParagraph doc para.Id { para with Buffer = newBuffer }
            | _ ->
                match findCursorParagraph doc sel.Anchor.BlockId with
                | Some para ->
                    let newBuffer = TextBuffer.setVertAlign formatStart formatLen align para.Buffer
                    if obj.ReferenceEquals(newBuffer, para.Buffer) then
                        doc
                    else
                        replaceParagraph doc para.Id { para with Buffer = newBuffer }
                | None -> doc
    | Some sel ->
        applySelectionBufferFormat doc sel (fun start len buffer ->
            TextBuffer.setVertAlign start len align buffer)

/// Set text color on selection
let setColor color doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc
    | Some sel when sel.Anchor.BlockId = sel.Focus.BlockId ->
        let formatStart = min sel.Anchor.Offset sel.Focus.Offset
        let formatLen = abs (sel.Focus.Offset - sel.Anchor.Offset)
        if formatLen <= 0 then
            doc
        else
            match doc.Blocks with
            | [ Paragraph para ] when para.Id = sel.Anchor.BlockId ->
                let newBuffer = TextBuffer.setColor formatStart formatLen color para.Buffer
                if obj.ReferenceEquals(newBuffer, para.Buffer) then
                    doc
                else
                    replaceParagraph doc para.Id { para with Buffer = newBuffer }
            | _ ->
                match findCursorParagraph doc sel.Anchor.BlockId with
                | Some para ->
                    let newBuffer = TextBuffer.setColor formatStart formatLen color para.Buffer
                    if obj.ReferenceEquals(newBuffer, para.Buffer) then
                        doc
                    else
                        replaceParagraph doc para.Id { para with Buffer = newBuffer }
                | None -> doc
    | Some sel ->
        applySelectionBufferFormat doc sel (fun start len buffer ->
            TextBuffer.setColor start len color buffer)

/// Set highlight color on selection
let setHighlight color doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc
    | Some sel when sel.Anchor.BlockId = sel.Focus.BlockId ->
        let formatStart = min sel.Anchor.Offset sel.Focus.Offset
        let formatLen = abs (sel.Focus.Offset - sel.Anchor.Offset)
        if formatLen <= 0 then
            doc
        else
            match doc.Blocks with
            | [ Paragraph para ] when para.Id = sel.Anchor.BlockId ->
                let newBuffer = TextBuffer.setHighlight formatStart formatLen color para.Buffer
                if obj.ReferenceEquals(newBuffer, para.Buffer) then
                    doc
                else
                    replaceParagraph doc para.Id { para with Buffer = newBuffer }
            | _ ->
                match findCursorParagraph doc sel.Anchor.BlockId with
                | Some para ->
                    let newBuffer = TextBuffer.setHighlight formatStart formatLen color para.Buffer
                    if obj.ReferenceEquals(newBuffer, para.Buffer) then
                        doc
                    else
                        replaceParagraph doc para.Id { para with Buffer = newBuffer }
                | None -> doc
    | Some sel ->
        applySelectionBufferFormat doc sel (fun start len buffer ->
            TextBuffer.setHighlight start len color buffer)

/// Set hyperlink on selection
let setHyperlink target doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc
    | Some sel when sel.Anchor.BlockId = sel.Focus.BlockId ->
        let formatStart = min sel.Anchor.Offset sel.Focus.Offset
        let formatLen = abs (sel.Focus.Offset - sel.Anchor.Offset)
        if formatLen <= 0 then
            doc
        else
            match doc.Blocks with
            | [ Paragraph para ] when para.Id = sel.Anchor.BlockId ->
                let newBuffer = TextBuffer.setHyperlink formatStart formatLen target para.Buffer
                if obj.ReferenceEquals(newBuffer, para.Buffer) then
                    doc
                else
                    replaceParagraph doc para.Id { para with Buffer = newBuffer }
            | _ ->
                match findCursorParagraph doc sel.Anchor.BlockId with
                | Some para ->
                    let newBuffer = TextBuffer.setHyperlink formatStart formatLen target para.Buffer
                    if obj.ReferenceEquals(newBuffer, para.Buffer) then
                        doc
                    else
                        replaceParagraph doc para.Id { para with Buffer = newBuffer }
                | None -> doc
    | Some sel ->
        applySelectionBufferFormat doc sel (fun start len buffer ->
            TextBuffer.setHyperlink start len target buffer)

/// Set comment metadata on selection
let setComment commentId commentText author date doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc
    | Some sel when sel.Anchor.BlockId = sel.Focus.BlockId ->
        let formatStart = min sel.Anchor.Offset sel.Focus.Offset
        let formatLen = abs (sel.Focus.Offset - sel.Anchor.Offset)
        if formatLen <= 0 then
            doc
        else
            match doc.Blocks with
            | [ Paragraph para ] when para.Id = sel.Anchor.BlockId ->
                let newBuffer = TextBuffer.setComment formatStart formatLen commentId commentText author date para.Buffer
                if obj.ReferenceEquals(newBuffer, para.Buffer) then
                    doc
                else
                    replaceParagraph doc para.Id { para with Buffer = newBuffer }
            | _ ->
                match findCursorParagraph doc sel.Anchor.BlockId with
                | Some para ->
                    let newBuffer = TextBuffer.setComment formatStart formatLen commentId commentText author date para.Buffer
                    if obj.ReferenceEquals(newBuffer, para.Buffer) then
                        doc
                    else
                        replaceParagraph doc para.Id { para with Buffer = newBuffer }
                | None -> doc
    | Some sel ->
        applySelectionBufferFormat doc sel (fun start len buffer ->
            TextBuffer.setComment start len commentId commentText author date buffer)

/// Set footnote/endnote metadata on selection
let setNote noteId noteType noteText doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc
    | Some sel when sel.Anchor.BlockId = sel.Focus.BlockId ->
        let formatStart = min sel.Anchor.Offset sel.Focus.Offset
        let formatLen = abs (sel.Focus.Offset - sel.Anchor.Offset)
        if formatLen <= 0 then
            doc
        else
            match doc.Blocks with
            | [ Paragraph para ] when para.Id = sel.Anchor.BlockId ->
                let newBuffer = TextBuffer.setNote formatStart formatLen noteId noteType noteText para.Buffer
                if obj.ReferenceEquals(newBuffer, para.Buffer) then
                    doc
                else
                    replaceParagraph doc para.Id { para with Buffer = newBuffer }
            | _ ->
                match findCursorParagraph doc sel.Anchor.BlockId with
                | Some para ->
                    let newBuffer = TextBuffer.setNote formatStart formatLen noteId noteType noteText para.Buffer
                    if obj.ReferenceEquals(newBuffer, para.Buffer) then
                        doc
                    else
                        replaceParagraph doc para.Id { para with Buffer = newBuffer }
                | None -> doc
    | Some sel ->
        applySelectionBufferFormat doc sel (fun start len buffer ->
            TextBuffer.setNote start len noteId noteType noteText buffer)

/// Set bookmark markers on selection
let setBookmark bookmarkName doc =
    let normalizedName =
        if System.String.IsNullOrWhiteSpace bookmarkName then "Bookmark"
        else bookmarkName.Trim().Replace(" ", "_")
    let bookmarkId = nextBookmarkId doc
    let updated =
        applyFormattingToSelection (fun start len buf -> TextBuffer.setBookmark start len bookmarkId normalizedName buf) doc
    if not (obj.ReferenceEquals(updated.Blocks, doc.Blocks)) then
        rememberBookmarkIdAfterChange doc updated bookmarkId normalizedName
    updated

/// Set content control metadata on selected paragraphs
let setContentControl controlId tag alias doc =
    let normalize fallback (value: string) =
        if System.String.IsNullOrWhiteSpace value then fallback else value.Trim()
    let controlId = normalize "1" controlId
    let tag = normalize "content-control" tag
    let alias = normalize tag alias
    updateSelectedParagraphMetadata true (fun metadata ->
        {
            metadata with
                ContentControlId = Some controlId
                ContentControlTag = Some tag
                ContentControlAlias = Some alias
        }) doc

/// Set named paragraph style on selected paragraphs
let setParagraphStyle styleId doc =
    let normalizedStyle =
        if System.String.IsNullOrWhiteSpace styleId then None else Some (styleId.Trim())
    updateSelectedParagraphMetadata true (fun metadata -> { metadata with Style = normalizedStyle }) doc

/// Set named character style on selection
let setCharacterStyle styleId doc =
    let normalizedStyle =
        if System.String.IsNullOrWhiteSpace styleId then None else Some (styleId.Trim())
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc
    | Some sel when sel.Anchor.BlockId = sel.Focus.BlockId ->
        let formatStart = min sel.Anchor.Offset sel.Focus.Offset
        let formatLen = abs (sel.Focus.Offset - sel.Anchor.Offset)
        if formatLen <= 0 then
            doc
        else
            match doc.Blocks with
            | [ Paragraph para ] when para.Id = sel.Anchor.BlockId ->
                let newBuffer = TextBuffer.setCharacterStyle formatStart formatLen normalizedStyle para.Buffer
                if obj.ReferenceEquals(newBuffer, para.Buffer) then
                    doc
                else
                    replaceParagraph doc para.Id { para with Buffer = newBuffer }
            | _ ->
                match findCursorParagraph doc sel.Anchor.BlockId with
                | Some para ->
                    let newBuffer = TextBuffer.setCharacterStyle formatStart formatLen normalizedStyle para.Buffer
                    if obj.ReferenceEquals(newBuffer, para.Buffer) then
                        doc
                    else
                        replaceParagraph doc para.Id { para with Buffer = newBuffer }
                | None -> doc
    | Some sel ->
        applySelectionBufferFormat doc sel (fun start len buffer ->
            TextBuffer.setCharacterStyle start len normalizedStyle buffer)

let private setRevisionMetadata revisionType revisionId author date doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc
    | Some sel when sel.Anchor.BlockId = sel.Focus.BlockId ->
        let formatStart = min sel.Anchor.Offset sel.Focus.Offset
        let formatLen = abs (sel.Focus.Offset - sel.Anchor.Offset)
        if formatLen <= 0 then
            doc
        else
            match doc.Blocks with
            | [ Paragraph para ] when para.Id = sel.Anchor.BlockId ->
                let newBuffer = TextBuffer.setRevision formatStart formatLen revisionType revisionId author date para.Buffer
                if obj.ReferenceEquals(newBuffer, para.Buffer) then
                    doc
                else
                    replaceParagraph doc para.Id { para with Buffer = newBuffer }
            | _ ->
                match findCursorParagraph doc sel.Anchor.BlockId with
                | Some para ->
                    let newBuffer = TextBuffer.setRevision formatStart formatLen revisionType revisionId author date para.Buffer
                    if obj.ReferenceEquals(newBuffer, para.Buffer) then
                        doc
                    else
                        replaceParagraph doc para.Id { para with Buffer = newBuffer }
                | None -> doc
    | Some sel ->
        applySelectionBufferFormat doc sel (fun start len buffer ->
            TextBuffer.setRevision start len revisionType revisionId author date buffer)

/// Set revision metadata on selection
let setRevision revisionType revisionId author date =
    let normalizedType = if revisionType = "delete" then "delete" else "insert"
    let revisionId = if System.String.IsNullOrWhiteSpace revisionId then "1" else revisionId
    let author = if System.String.IsNullOrWhiteSpace author then None else Some author
    let date = if System.String.IsNullOrWhiteSpace date then None else Some date
    fun doc -> setRevisionMetadata normalizedType revisionId author date doc

let private setRevisionWithOptions revisionType revisionId author date doc =
    setRevisionMetadata revisionType revisionId author date doc

let private adjacentInsertRevisionId doc (pos: Position) author date =
    findCursorParagraph doc pos.BlockId
    |> Option.bind (fun para ->
        para.Buffer.Runs
        |> List.tryPick (function
            | TextBuffer.Text tr when tr.Start + tr.Length = pos.Offset && tr.Style.RevisionType = Some "insert" && tr.Style.RevisionAuthor = author && tr.Style.RevisionDate = date -> tr.Style.RevisionId
            | _ -> None))

let private adjacentInsertRevisionIdInParagraph (para: ParagraphBlock) offset author date =
    para.Buffer.Runs
    |> List.tryPick (function
        | TextBuffer.Text tr when tr.Start + tr.Length = offset && tr.Style.RevisionType = Some "insert" && tr.Style.RevisionAuthor = author && tr.Style.RevisionDate = date -> tr.Style.RevisionId
        | _ -> None)

/// Insert text while recording the inserted run as a revision.
let insertTrackedText author date (insertedText: string) doc =
    let author = if System.String.IsNullOrWhiteSpace author then None else Some author
    let date = if System.String.IsNullOrWhiteSpace date then None else Some date
    if insertedText.Length = 0 then
        doc
    else
        match doc.Selection with
        | Some sel when sel.IsCollapsed ->
            let mutable changed = false
            let mutable newSelection: Selection option = None
            let mutable usedRevisionId: string option = None
            let updated =
                updateParagraph doc sel.Focus.BlockId (fun para ->
                    changed <- true
                    let insertionRevisionId =
                        match adjacentInsertRevisionIdInParagraph para sel.Focus.Offset author date with
                        | Some revisionId -> revisionId
                        | None -> nextRevisionId doc
                    usedRevisionId <- Some insertionRevisionId
                    let startOffset = sel.Focus.Offset
                    let buffer =
                        para.Buffer
                        |> TextBuffer.insert startOffset insertedText
                        |> TextBuffer.applyStyle startOffset insertedText.Length (fun style ->
                            { style with RevisionType = Some "insert"; RevisionId = Some insertionRevisionId; RevisionAuthor = author; RevisionDate = date })
                    newSelection <- Some {
                        Anchor = { sel.Focus with Offset = startOffset + insertedText.Length }
                        Focus = { sel.Focus with Offset = startOffset + insertedText.Length }
                    }
                    { para with Buffer = buffer })
            if changed then
                let result = { updated with Selection = newSelection }
                match usedRevisionId with
                | Some revisionId ->
                    rememberRevisionIdAfterChange doc result "insert" revisionId author date insertedText.Length
                | None -> ()
                result
            else
                doc
        | Some sel when not sel.IsCollapsed ->
            let deletionRevisionId = nextRevisionId doc
            let startPos, _, _ = getNormalizedRange sel doc
            let docForInsert =
                let marked = setRevisionWithOptions "delete" deletionRevisionId author date doc
                let result = { marked with Selection = Some { Anchor = startPos; Focus = startPos } }
                rememberRevisionIdAfterChange doc result "delete" deletionRevisionId author date 0
                result
            let insertionRevisionId =
                match docForInsert.Selection with
                | Some sel ->
                    match adjacentInsertRevisionId docForInsert sel.Focus author date with
                    | Some revisionId -> revisionId
                    | None -> nextRevisionId docForInsert
                | None -> nextRevisionId docForInsert
            let inserted = insertText insertedText docForInsert
            match inserted.Selection with
            | None -> inserted
            | Some sel ->
                let startOffset = max 0 (sel.Focus.Offset - insertedText.Length)
                let revisionSelection = {
                    Anchor = { sel.Focus with Offset = startOffset }
                    Focus = sel.Focus
                }
                let styled = setRevisionWithOptions "insert" insertionRevisionId author date { inserted with Selection = Some revisionSelection }
                rememberRevisionIdAfterChange docForInsert styled "insert" insertionRevisionId author date insertedText.Length
                { styled with Selection = Some sel }
        | Some _ -> doc
        | None -> doc

/// Mark the selected text as deleted instead of physically removing it.
let trackDeleteSelection author date doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc
    | Some sel ->
        let startPos, _, _ = getNormalizedRange sel doc
        let author = if System.String.IsNullOrWhiteSpace author then None else Some author
        let date = if System.String.IsNullOrWhiteSpace date then None else Some date
        let revisionId = nextRevisionId doc
        let marked = setRevisionWithOptions "delete" revisionId author date doc
        rememberRevisionIdAfterChange doc marked "delete" revisionId author date 0
        { marked with Selection = Some { Anchor = startPos; Focus = startPos } }

let private isDeletedRun (run: TextBuffer.TextRun) =
    run.Style.RevisionType = Some "delete"

let private nextTrackDeletableOffset para offset =
    let mutable result = None
    for run in para.Buffer.Runs do
        if result.IsNone then
            match run with
            | TextBuffer.Text tr when not (isDeletedRun tr) && tr.Start + tr.Length > offset ->
                result <- Some (max offset tr.Start)
            | _ -> ()
    result

let private previousTrackDeletableOffset para offset =
    let mutable result = None
    for run in para.Buffer.Runs do
        match run with
        | TextBuffer.Text tr when not (isDeletedRun tr) && tr.Start < offset ->
            let value = (min offset (tr.Start + tr.Length)) - 1
            if value >= 0 then
                result <- Some value
        | _ -> ()
    result

let private trackDeleteSingleWhere blockId offsetForParagraph author date doc =
    let author = if System.String.IsNullOrWhiteSpace author then None else Some author
    let date = if System.String.IsNullOrWhiteSpace date then None else Some date
    let mutable changed = false
    let mutable selection = None
    let mutable usedRevisionId: string option = None
    let updated =
        updateParagraph doc blockId (fun para ->
            match offsetForParagraph para with
            | None -> para
            | Some offset ->
                changed <- true
                let revisionId = nextRevisionId doc
                usedRevisionId <- Some revisionId
                let buffer = TextBuffer.setRevision offset 1 "delete" revisionId author date para.Buffer
                selection <- Some {
                    Anchor = { BlockId = para.Id; Offset = offset }
                    Focus = { BlockId = para.Id; Offset = offset }
                }
                { para with Buffer = buffer })
    if changed then
        let result = { updated with Selection = selection }
        match usedRevisionId with
        | Some revisionId -> rememberRevisionIdAfterChange doc result "delete" revisionId author date 0
        | None -> ()
        result
    else
        doc

let trackDeleteForward author date doc =
    match doc.Selection with
    | None -> doc
    | Some sel when not sel.IsCollapsed -> trackDeleteSelection author date doc
    | Some sel ->
        trackDeleteSingleWhere
            sel.Focus.BlockId
            (fun para -> nextTrackDeletableOffset para sel.Focus.Offset)
            author
            date
            doc

let trackDeleteBackward author date doc =
    match doc.Selection with
    | None -> doc
    | Some sel when not sel.IsCollapsed -> trackDeleteSelection author date doc
    | Some sel ->
        trackDeleteSingleWhere
            sel.Focus.BlockId
            (fun para -> previousTrackDeletableOffset para sel.Focus.Offset)
            author
            date
            doc

type private RevisionSpan = {
    Start: int
    Length: int
    ShouldDelete: bool
}

let private addRevisionSpan (spans: ResizeArray<RevisionSpan>) span =
    if spans.Count > 0 then
        let last = spans.[spans.Count - 1]
        if last.ShouldDelete = span.ShouldDelete && last.Start + last.Length = span.Start then
            spans.[spans.Count - 1] <- { last with Length = last.Length + span.Length }
        else
            spans.Add span
    else
        spans.Add span

let private selectedRevisionSpans accept para selectionRange =
    let selStart, selEnd, isCollapsed = selectionRange
    let intersects start endPos selStart selEnd isCollapsed =
        if isCollapsed then selStart >= start && selStart <= endPos
        else start < selEnd && endPos > selStart

    let spans = ResizeArray<RevisionSpan>()
    let mutable firstStart = None
    for runValue in para.Buffer.Runs do
        match runValue with
        | Text tr ->
            match tr.Style.RevisionType with
            | Some revisionType ->
                let runEnd = tr.Start + tr.Length
                if intersects tr.Start runEnd selStart selEnd isCollapsed then
                    let shouldDelete =
                        match accept, revisionType with
                        | true, "delete" -> true
                        | false, "insert" -> true
                        | _ -> false
                    addRevisionSpan spans { Start = tr.Start; Length = tr.Length; ShouldDelete = shouldDelete }
                    match firstStart with
                    | None -> firstStart <- Some tr.Start
                    | Some _ -> ()
            | None -> ()
        | Image _ -> ()
    spans, firstStart

let revisionInfos doc =
    match tryGetRevisionInfoCache doc.Blocks with
    | ValueSome cached ->
        cached.Infos
    | ValueNone ->
        let revisions = ResizeArray<RevisionInfo>()
        iterEditableParagraphs doc (fun para ->
            let mutable pendingStart = 0
            let mutable pendingEnd = 0
            let mutable pendingRevisionType: string option = None
            let mutable pendingRevisionId: string option = None
            let mutable pendingAuthor: string option = None
            let mutable pendingDate: string option = None
            let flushPending () =
                match pendingRevisionType with
                | Some revisionType ->
                    revisions.Add {
                        BlockId = para.Id
                        Start = pendingStart
                        End = pendingEnd
                        RevisionType = revisionType
                        Author = pendingAuthor
                        Date = pendingDate
                        Text = TextBuffer.slice pendingStart (pendingEnd - pendingStart) para.Buffer
                    }
                    pendingRevisionType <- None
                    pendingRevisionId <- None
                    pendingAuthor <- None
                    pendingDate <- None
                | None -> ()
            for runValue in para.Buffer.Runs do
                match runValue with
                | Text tr ->
                    match tr.Style.RevisionType with
                    | Some revisionType ->
                        let runEnd = tr.Start + tr.Length
                        let canExtend =
                            match pendingRevisionType with
                            | Some currentType ->
                                pendingEnd = tr.Start
                                && currentType = revisionType
                                && pendingRevisionId = tr.Style.RevisionId
                                && pendingAuthor = tr.Style.RevisionAuthor
                                && pendingDate = tr.Style.RevisionDate
                            | None -> false
                        if canExtend then
                            pendingEnd <- runEnd
                        else
                            flushPending ()
                            pendingStart <- tr.Start
                            pendingEnd <- runEnd
                            pendingRevisionType <- Some revisionType
                            pendingRevisionId <- tr.Style.RevisionId
                            pendingAuthor <- tr.Style.RevisionAuthor
                            pendingDate <- tr.Style.RevisionDate
                    | None ->
                        flushPending ()
                | Image _ ->
                    flushPending ()
            flushPending ())
        let infos = resizeArrayToList revisions
        let sourceParagraphCount = revisionInfoSourceParagraphCountFor infos.Length doc.Blocks
        rememberRevisionInfos {
            SourceBlocks = doc.Blocks
            SourceParagraphCount = sourceParagraphCount
            SourceTextCharCount = revisionInfoSourceTextCharCountFor infos.Length sourceParagraphCount doc.Blocks
            InfoCount = infos.Length
            Infos = infos
        }
        infos

type private RevisionNavigationSpan = {
    BlockId: BlockId
    ParagraphIndex: int
    Start: int
    End: int
}

type private RevisionNavigationCache = {
    SourceBlocks: Block list
    SourceParagraphCount: int
    SourceTextCharCount: int
    SpanCount: int
    Spans: RevisionNavigationSpan array
    ParagraphIndexes: Dictionary<BlockId, int>
    ParagraphIndexesComplete: bool
}

[<Struct>]
type private RevisionSelectionDocumentCacheKey = {
    Source: Document
    Previous: bool
}

type private RevisionSelectionDocumentCacheEntry = {
    Key: RevisionSelectionDocumentCacheKey
    Result: Document
    SourceParagraphCount: int
    SourceTextCharCount: int
}

let mutable private revisionNavigationCache: RevisionNavigationCache list = []
let private revisionNavigationCacheByBlocks = Dictionary<Block list, RevisionNavigationCache>(documentInfoBlocksComparer)
let private revisionNavigationFastSlots = Array.zeroCreate<RevisionNavigationCache> 256
let mutable private revisionSelectionDocumentCache: RevisionSelectionDocumentCacheEntry list = []
let private revisionSelectionDocumentKeyComparer =
    { new IEqualityComparer<RevisionSelectionDocumentCacheKey> with
        member _.Equals(left, right) =
            left.Previous = right.Previous
            && obj.ReferenceEquals(left.Source, right.Source)

        member _.GetHashCode(key) =
            (System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(key.Source) * 397)
            ^^^ (if key.Previous then 1 else 0) }
let private revisionSelectionDocumentCacheByKey =
    Dictionary<RevisionSelectionDocumentCacheKey, RevisionSelectionDocumentCacheEntry>(revisionSelectionDocumentKeyComparer)
let private revisionSelectionDocumentFastSlots = Array.zeroCreate<RevisionSelectionDocumentCacheEntry> 256
let private revisionNavigationCacheMaxEntries = 256
let private revisionNavigationCacheMaxTotalSpans = 200000
let private revisionNavigationCacheMaxTotalParagraphs = 100000
let private revisionNavigationCacheMaxTotalSourceTextChars = documentSourceTextCacheMaxTotalChars

let private revisionSelectionDocumentKey previous source = {
    Source = source
    Previous = previous
}

let private revisionSelectionDocumentCacheMatches previous source (cached: RevisionSelectionDocumentCacheEntry) =
    cached.Key.Previous = previous
    && obj.ReferenceEquals(cached.Key.Source, source)

let private revisionSelectionDocumentSlot previous source =
    revisionSelectionDocumentKeyComparer.GetHashCode(revisionSelectionDocumentKey previous source)
    &&& (revisionSelectionDocumentFastSlots.Length - 1)

let private setRevisionSelectionDocumentCache entries =
    revisionSelectionDocumentCache <- entries
    revisionSelectionDocumentCacheByKey.Clear()
    clearArray revisionSelectionDocumentFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<RevisionSelectionDocumentCacheEntry>
        if not (revisionSelectionDocumentCacheByKey.TryGetValue(entry.Key, &existing)) then
            revisionSelectionDocumentCacheByKey.[entry.Key] <- entry
            revisionSelectionDocumentFastSlots.[revisionSelectionDocumentSlot entry.Key.Previous entry.Key.Source] <- entry

let private tryGetRevisionSelectionDocument previous source =
    let slot = revisionSelectionDocumentSlot previous source
    let fast = revisionSelectionDocumentFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && revisionSelectionDocumentCacheMatches previous source fast then
        ValueSome fast.Result
    else
        let key = revisionSelectionDocumentKey previous source
        let mutable cached = Unchecked.defaultof<RevisionSelectionDocumentCacheEntry>
        if revisionSelectionDocumentCacheByKey.TryGetValue(key, &cached) then
            revisionSelectionDocumentFastSlots.[slot] <- cached
            ValueSome cached.Result
        else
            ValueNone

let private revisionSelectionDocumentCacheWithoutSource previous source =
    let key = revisionSelectionDocumentKey previous source
    let mutable cached = Unchecked.defaultof<RevisionSelectionDocumentCacheEntry>
    if revisionSelectionDocumentCacheByKey.TryGetValue(key, &cached)
       && revisionSelectionDocumentCacheMatches previous source cached then
        revisionSelectionDocumentCache
        |> List.filter (revisionSelectionDocumentCacheMatches previous source >> not)
    else
        revisionSelectionDocumentCache

let private revisionSelectionDocumentCacheRetainable (cached: RevisionSelectionDocumentCacheEntry) =
    cached.SourceParagraphCount <= revisionNavigationCacheMaxTotalParagraphs
    && cached.SourceTextCharCount <= revisionNavigationCacheMaxTotalSourceTextChars

let private revisionNavigationCacheMatches blocks (cached: RevisionNavigationCache) =
    obj.ReferenceEquals(cached.SourceBlocks, blocks)

let private setRevisionNavigationCache entries =
    revisionNavigationCache <- entries
    revisionNavigationCacheByBlocks.Clear()
    clearArray revisionNavigationFastSlots
    for entry in entries do
        let mutable existing = Unchecked.defaultof<RevisionNavigationCache>
        if not (revisionNavigationCacheByBlocks.TryGetValue(entry.SourceBlocks, &existing)) then
            revisionNavigationCacheByBlocks.[entry.SourceBlocks] <- entry
            let slot = documentInfoSlot entry.SourceBlocks revisionNavigationFastSlots.Length
            revisionNavigationFastSlots.[slot] <- entry

let private tryGetRevisionNavigationCache blocks =
    let slot = documentInfoSlot blocks revisionNavigationFastSlots.Length
    let fast = revisionNavigationFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && revisionNavigationCacheMatches blocks fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<RevisionNavigationCache>
        if revisionNavigationCacheByBlocks.TryGetValue(blocks, &cached) then
            revisionNavigationFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private revisionNavigationCacheWithoutBlocks blocks =
    let mutable cached = Unchecked.defaultof<RevisionNavigationCache>
    if revisionNavigationCacheByBlocks.TryGetValue(blocks, &cached) then
        revisionNavigationCache
        |> List.filter (revisionNavigationCacheMatches blocks >> not)
    else
        revisionNavigationCache

let private removeRevisionNavigationCacheForBlocks blocks =
    let entries = revisionNavigationCacheWithoutBlocks blocks
    if not (obj.ReferenceEquals(entries, revisionNavigationCache)) then
        setRevisionNavigationCache entries

let private revisionNavigationCacheRetainable (cached: RevisionNavigationCache) =
    cached.SpanCount <= revisionNavigationCacheMaxTotalSpans
    && cached.SourceParagraphCount <= revisionNavigationCacheMaxTotalParagraphs
    && cached.SourceTextCharCount <= revisionNavigationCacheMaxTotalSourceTextChars

let private revisionNavigationSourceTextCharCountFor spanCount sourceParagraphCount blocks =
    if spanCount > revisionNavigationCacheMaxTotalSpans
       || sourceParagraphCount > revisionNavigationCacheMaxTotalParagraphs then
        revisionNavigationCacheMaxTotalSourceTextChars + 1
    else
        documentBlocksSourceTextCharCountFor
            revisionNavigationCacheMaxTotalParagraphs
            sourceParagraphCount
            blocks

let private rememberRevisionNavigation cached =
    if not (revisionNavigationCacheRetainable cached) then
        removeRevisionNavigationCacheForBlocks cached.SourceBlocks
    else
        let trim entries =
            let rec loop keptCount keptSpans keptParagraphs keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateSpans = max 1 candidate.SpanCount
                    let candidateParagraphs = max 1 candidate.SourceParagraphCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    if revisionNavigationCacheRetainable candidate
                       && keptCount < revisionNavigationCacheMaxEntries
                       && keptSpans + candidateSpans <= revisionNavigationCacheMaxTotalSpans
                       && keptParagraphs + candidateParagraphs <= revisionNavigationCacheMaxTotalParagraphs
                       && keptSourceTextChars + candidateSourceTextChars <= revisionNavigationCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptSpans + candidateSpans)
                            (keptParagraphs + candidateParagraphs)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptSpans keptParagraphs keptSourceTextChars acc tail
            loop 0 0 0 0 [] entries
        (cached
         :: revisionNavigationCacheWithoutBlocks cached.SourceBlocks)
        |> trim
        |> setRevisionNavigationCache

let private revisionNavigationSpans doc =
    match tryGetRevisionNavigationCache doc.Blocks with
    | ValueSome cached ->
        cached
    | ValueNone ->
        let spans = ResizeArray<RevisionNavigationSpan>()
        let paragraphIndexes = Dictionary<BlockId, int>()
        let mutable paragraphIndex = 0
        iterEditableParagraphs doc (fun para ->
            let currentParagraphIndex = paragraphIndex
            paragraphIndex <- paragraphIndex + 1

            let mutable pendingStart = 0
            let mutable pendingEnd = 0
            let mutable pendingRevisionType: string option = None
            let mutable pendingRevisionId: string option = None
            let mutable pendingAuthor: string option = None
            let mutable pendingDate: string option = None
            let flushPending () =
                match pendingRevisionType with
                | Some _ ->
                    spans.Add {
                        BlockId = para.Id
                        ParagraphIndex = currentParagraphIndex
                        Start = pendingStart
                        End = pendingEnd
                    }
                    if not (paragraphIndexes.ContainsKey para.Id) then
                        paragraphIndexes.[para.Id] <- currentParagraphIndex
                    pendingRevisionType <- None
                    pendingRevisionId <- None
                    pendingAuthor <- None
                    pendingDate <- None
                | None -> ()
            for runValue in para.Buffer.Runs do
                match runValue with
                | Text tr ->
                    match tr.Style.RevisionType with
                    | Some revisionType ->
                        let runEnd = tr.Start + tr.Length
                        let canExtend =
                            match pendingRevisionType with
                            | Some currentType ->
                                pendingEnd = tr.Start
                                && currentType = revisionType
                                && pendingRevisionId = tr.Style.RevisionId
                                && pendingAuthor = tr.Style.RevisionAuthor
                                && pendingDate = tr.Style.RevisionDate
                            | None -> false
                        if canExtend then
                            pendingEnd <- runEnd
                        else
                            flushPending ()
                            pendingStart <- tr.Start
                            pendingEnd <- runEnd
                            pendingRevisionType <- Some revisionType
                            pendingRevisionId <- tr.Style.RevisionId
                            pendingAuthor <- tr.Style.RevisionAuthor
                            pendingDate <- tr.Style.RevisionDate
                    | None ->
                        flushPending ()
                | Image _ ->
                    flushPending ()
            flushPending ())
        let cached = {
            SourceBlocks = doc.Blocks
            SourceParagraphCount = paragraphIndex
            SourceTextCharCount = revisionNavigationSourceTextCharCountFor spans.Count paragraphIndex doc.Blocks
            SpanCount = spans.Count
            Spans = spans.ToArray()
            ParagraphIndexes = paragraphIndexes
            ParagraphIndexesComplete = paragraphIndexes.Count = paragraphIndex
        }
        rememberRevisionNavigation cached
        cached

let private rememberRevisionSelectionDocument previous source result =
    if not (obj.ReferenceEquals(source, result)) then
        let navigation = revisionNavigationSpans source
        let entry = {
            Key = revisionSelectionDocumentKey previous source
            Result = result
            SourceParagraphCount = navigation.SourceParagraphCount
            SourceTextCharCount = navigation.SourceTextCharCount
        }
        if revisionSelectionDocumentCacheRetainable entry then
            let rec trim remaining kept count totalParagraphs totalSourceTextChars =
                match remaining with
                | [] -> List.rev kept
                | candidate :: rest ->
                    let candidateParagraphs = max 1 candidate.SourceParagraphCount
                    let candidateSourceTextChars = max 1 candidate.SourceTextCharCount
                    let nextTotalParagraphs = totalParagraphs + candidateParagraphs
                    let nextTotalSourceTextChars = totalSourceTextChars + candidateSourceTextChars
                    if revisionSelectionDocumentCacheRetainable candidate
                       && count < revisionNavigationCacheMaxEntries
                       && nextTotalParagraphs <= revisionNavigationCacheMaxTotalParagraphs
                       && nextTotalSourceTextChars <= revisionNavigationCacheMaxTotalSourceTextChars then
                        trim rest (candidate :: kept) (count + 1) nextTotalParagraphs nextTotalSourceTextChars
                    else
                        trim rest kept count totalParagraphs totalSourceTextChars
            let entries = entry :: revisionSelectionDocumentCacheWithoutSource previous source
            trim entries [] 0 0 0
            |> setRevisionSelectionDocumentCache

let private applyRevisionDecision accept doc =
    let selectedBlockIds, selectionRangeForParagraph =
        match doc.Selection with
        | None ->
            [], fun _ (_: ParagraphBlock) -> None
        | Some sel when sel.IsCollapsed ->
            [sel.Focus.BlockId],
            fun blockId (_: ParagraphBlock) ->
                if sel.Focus.BlockId = blockId then Some (sel.Focus.Offset, sel.Focus.Offset, true) else None
        | Some sel ->
            let startPos, endPos, blockIds = getNormalizedRange sel doc
            let selectedBlockIdSet = HashSet<BlockId>(blockIds)
            blockIds,
            fun blockId (para: ParagraphBlock) ->
                if selectedBlockIdSet.Contains blockId then
                    let bufferLen = TextBuffer.length para.Buffer
                    let startOffset = if blockId = startPos.BlockId then startPos.Offset else 0
                    let endOffset = if blockId = endPos.BlockId then endPos.Offset else bufferLen
                    Some (startOffset, endOffset, false)
                else None

    let mutable changed = false
    let mutable firstSelection: Position option = None
    let replacements = Dictionary<BlockId, ParagraphBlock>()

    let processParagraph (para: ParagraphBlock) =
        match selectionRangeForParagraph para.Id para with
        | Some selectionRange ->
            let selStart, selEnd, isCollapsed = selectionRange
            let struct (buffer, firstStart, paragraphChanged) =
                TextBuffer.resolveRevisionsInRange accept selStart selEnd isCollapsed para.Buffer
            if paragraphChanged then
                let newPara = { para with Buffer = buffer }
                replacements.[para.Id] <- newPara
                changed <- true
                match firstSelection with
                | None ->
                    match firstStart with
                    | Some start -> firstSelection <- Some { BlockId = para.Id; Offset = start }
                    | None -> ()
                | Some _ -> ()
        | None -> ()

    match selectedBlockIds with
    | [blockId] ->
        match findCursorParagraph doc blockId with
        | Some para -> processParagraph para
        | None -> ()
    | _ ->
        iterEditableParagraphs doc processParagraph

    if changed then
        let updatedDoc = replaceParagraphs doc replacements
        let selection = firstSelection |> Option.map (fun pos -> { Anchor = pos; Focus = pos })
        { updatedDoc with Selection = selection |> Option.orElse updatedDoc.Selection }
    else doc

/// Accept selected revision runs. Insertions become normal text; deletions are removed.
let acceptSelectedRevision doc =
    applyRevisionDecision true doc

/// Reject selected revision runs. Insertions are removed; deletions become normal text.
let rejectSelectedRevision doc =
    applyRevisionDecision false doc

let private applyAllRevisionDecision accept doc =
    let replacements = Dictionary<BlockId, ParagraphBlock>()
    let mutable firstSelection: Position option = None
    iterEditableParagraphs doc (fun para ->
        let bufferLen = TextBuffer.length para.Buffer
        let struct (buffer, firstStart, paragraphChanged) =
            TextBuffer.resolveRevisionsInRange accept 0 bufferLen false para.Buffer
        if paragraphChanged then
            replacements.[para.Id] <- { para with Buffer = buffer }
            match firstSelection, firstStart with
            | None, Some start -> firstSelection <- Some { BlockId = para.Id; Offset = start }
            | _ -> ())
    if replacements.Count = 0 then
        doc
    else
        let updatedDoc = replaceParagraphs doc replacements
        let selection = firstSelection |> Option.map (fun pos -> { Anchor = pos; Focus = pos })
        { updatedDoc with Selection = selection |> Option.orElse updatedDoc.Selection }

/// Accept all revision runs in document order.
let acceptAllRevisions doc =
    applyAllRevisionDecision true doc

/// Reject all revision runs in document order.
let rejectAllRevisions doc =
    applyAllRevisionDecision false doc

let private currentRevisionPosition useSelectionStart doc =
    match doc.Selection with
    | Some sel ->
        if sel.Anchor.BlockId = sel.Focus.BlockId then
            let offset = if useSelectionStart then min sel.Anchor.Offset sel.Focus.Offset else max sel.Anchor.Offset sel.Focus.Offset
            { sel.Focus with Offset = offset }
        elif useSelectionStart then
            let startPos, _, _ = getNormalizedRange sel doc
            startPos
        else
            let _, endPos, _ = getNormalizedRange sel doc
            endPos
    | None -> { BlockId = ""; Offset = 0 }

let private currentRevisionPositionStruct useSelectionStart doc =
    match doc.Selection with
    | Some sel ->
        if sel.Anchor.BlockId = sel.Focus.BlockId then
            let offset =
                if useSelectionStart then min sel.Anchor.Offset sel.Focus.Offset
                else max sel.Anchor.Offset sel.Focus.Offset
            struct (sel.Focus.BlockId, offset)
        elif useSelectionStart then
            let startPos, _, _ = getNormalizedRange sel doc
            struct (startPos.BlockId, startPos.Offset)
        else
            let _, endPos, _ = getNormalizedRange sel doc
            struct (endPos.BlockId, endPos.Offset)
    | None ->
        struct ("", 0)

let private selectRevisionRange blockId start endPos doc =
    match doc.Selection with
    | Some current
        when current.Anchor.BlockId = blockId
             && current.Anchor.Offset = start
             && current.Focus.BlockId = blockId
             && current.Focus.Offset = endPos ->
        doc
    | _ ->
        { doc with
            Selection =
                Some {
                    Anchor = { BlockId = blockId; Offset = start }
                    Focus = { BlockId = blockId; Offset = endPos }
                } }

let private selectRevision revision doc =
    selectRevisionRange revision.BlockId revision.Start revision.End doc

let private revisionSpanAfter paragraphIndex offset (span: RevisionNavigationSpan) =
    span.ParagraphIndex > paragraphIndex
        || (span.ParagraphIndex = paragraphIndex && span.Start > offset)

let private revisionSpanBefore paragraphIndex offset (span: RevisionNavigationSpan) =
    span.ParagraphIndex < paragraphIndex
        || (span.ParagraphIndex = paragraphIndex && span.Start < offset)

let private tryEditableParagraphIndexById doc paragraphId =
    let mutable index = 0
    let mutable found = ValueNone
    iterEditableParagraphsUntil doc (fun para ->
        if para.Id = paragraphId then
            found <- ValueSome index
            false
        else
            index <- index + 1
            true)
    found

let private tryFindNextRevisionSpan doc =
    let cache = revisionNavigationSpans doc
    if cache.Spans.Length = 0 then
        ValueNone
    else
        let struct (currentBlockId, currentOffset) = currentRevisionPositionStruct false doc
        let span =
            if cache.ParagraphIndexesComplete then
                let mutable currentIndex = -1
                let hasCurrentIndex =
                    currentBlockId <> "" && cache.ParagraphIndexes.TryGetValue(currentBlockId, &currentIndex)
                if currentBlockId = "" || not hasCurrentIndex then
                    cache.Spans.[0]
                else
                    let mutable lo = 0
                    let mutable hi = cache.Spans.Length
                    while lo < hi do
                        let mid = lo + (hi - lo) / 2
                        if revisionSpanAfter currentIndex currentOffset cache.Spans.[mid] then
                            hi <- mid
                        else
                            lo <- mid + 1
                    if lo < cache.Spans.Length then cache.Spans.[lo] else cache.Spans.[0]
            else
                if currentBlockId = "" then
                    cache.Spans.[0]
                else
                    let mutable currentIndex = -1
                    let hasCurrentIndex =
                        if cache.ParagraphIndexes.TryGetValue(currentBlockId, &currentIndex) then
                            true
                        else
                            match tryEditableParagraphIndexById doc currentBlockId with
                            | ValueSome index ->
                                currentIndex <- index
                                true
                            | ValueNone ->
                                false
                    if not hasCurrentIndex then
                        cache.Spans.[0]
                    else
                        let mutable lo = 0
                        let mutable hi = cache.Spans.Length
                        while lo < hi do
                            let mid = lo + (hi - lo) / 2
                            if revisionSpanAfter currentIndex currentOffset cache.Spans.[mid] then
                                hi <- mid
                            else
                                lo <- mid + 1
                        if lo < cache.Spans.Length then cache.Spans.[lo] else cache.Spans.[0]
        ValueSome (struct (span.BlockId, span.Start, span.End))

/// Select next revision, wrapping to the first revision when needed.
let selectNextRevision doc =
    match tryGetRevisionSelectionDocument false doc with
    | ValueSome cached -> cached
    | ValueNone ->
        let result =
            match tryFindNextRevisionSpan doc with
            | ValueSome (struct (blockId, start, endPos)) -> selectRevisionRange blockId start endPos doc
            | ValueNone -> doc
        rememberRevisionSelectionDocument false doc result
        result

let private tryFindPreviousRevisionSpan doc =
    let cache = revisionNavigationSpans doc
    if cache.Spans.Length = 0 then
        ValueNone
    else
        let struct (currentBlockId, currentOffset) = currentRevisionPositionStruct true doc
        let span =
            if cache.ParagraphIndexesComplete then
                let mutable currentIndex = -1
                let hasCurrentIndex =
                    currentBlockId <> "" && cache.ParagraphIndexes.TryGetValue(currentBlockId, &currentIndex)
                if currentBlockId = "" || not hasCurrentIndex then
                    cache.Spans.[cache.Spans.Length - 1]
                else
                    let mutable lo = 0
                    let mutable hi = cache.Spans.Length
                    while lo < hi do
                        let mid = lo + (hi - lo) / 2
                        if revisionSpanBefore currentIndex currentOffset cache.Spans.[mid] then
                            lo <- mid + 1
                        else
                            hi <- mid
                    let index = lo - 1
                    if index >= 0 then cache.Spans.[index] else cache.Spans.[cache.Spans.Length - 1]
            else
                if currentBlockId = "" then
                    cache.Spans.[cache.Spans.Length - 1]
                else
                    let mutable currentIndex = -1
                    let hasCurrentIndex =
                        if cache.ParagraphIndexes.TryGetValue(currentBlockId, &currentIndex) then
                            true
                        else
                            match tryEditableParagraphIndexById doc currentBlockId with
                            | ValueSome index ->
                                currentIndex <- index
                                true
                            | ValueNone ->
                                false
                    if not hasCurrentIndex then
                        cache.Spans.[cache.Spans.Length - 1]
                    else
                        let mutable lo = 0
                        let mutable hi = cache.Spans.Length
                        while lo < hi do
                            let mid = lo + (hi - lo) / 2
                            if revisionSpanBefore currentIndex currentOffset cache.Spans.[mid] then
                                lo <- mid + 1
                            else
                                hi <- mid
                        let index = lo - 1
                        if index >= 0 then cache.Spans.[index] else cache.Spans.[cache.Spans.Length - 1]
        ValueSome (struct (span.BlockId, span.Start, span.End))

/// Select previous revision, wrapping to the last revision when needed.
let selectPreviousRevision doc =
    match tryGetRevisionSelectionDocument true doc with
    | ValueSome cached -> cached
    | ValueNone ->
        let result =
            match tryFindPreviousRevisionSpan doc with
            | ValueSome (struct (blockId, start, endPos)) -> selectRevisionRange blockId start endPos doc
            | ValueNone -> doc
        rememberRevisionSelectionDocument true doc result
        result

/// Set font size on selection
let setFontSize size doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc
    | Some sel when sel.Anchor.BlockId = sel.Focus.BlockId ->
        let formatStart = min sel.Anchor.Offset sel.Focus.Offset
        let formatLen = abs (sel.Focus.Offset - sel.Anchor.Offset)
        if formatLen <= 0 then
            doc
        else
            match doc.Blocks with
            | [ Paragraph para ] when para.Id = sel.Anchor.BlockId ->
                let newBuffer = TextBuffer.setFontSize formatStart formatLen size para.Buffer
                if obj.ReferenceEquals(newBuffer, para.Buffer) then
                    doc
                else
                    replaceParagraph doc para.Id { para with Buffer = newBuffer }
            | _ ->
                match findCursorParagraph doc sel.Anchor.BlockId with
                | Some para ->
                    let newBuffer = TextBuffer.setFontSize formatStart formatLen size para.Buffer
                    if obj.ReferenceEquals(newBuffer, para.Buffer) then
                        doc
                    else
                        replaceParagraph doc para.Id { para with Buffer = newBuffer }
                | None -> doc
    | Some sel ->
        applySelectionBufferFormat doc sel (fun start len buffer ->
            TextBuffer.setFontSize start len size buffer)

/// Set font family on selection
let setFontFamily family doc =
    match doc.Selection with
    | None -> doc
    | Some sel when sel.IsCollapsed -> doc
    | Some sel when sel.Anchor.BlockId = sel.Focus.BlockId ->
        let formatStart = min sel.Anchor.Offset sel.Focus.Offset
        let formatLen = abs (sel.Focus.Offset - sel.Anchor.Offset)
        if formatLen <= 0 then
            doc
        else
            match doc.Blocks with
            | [ Paragraph para ] when para.Id = sel.Anchor.BlockId ->
                let newBuffer = TextBuffer.setFontFamily formatStart formatLen family para.Buffer
                if obj.ReferenceEquals(newBuffer, para.Buffer) then
                    doc
                else
                    replaceParagraph doc para.Id { para with Buffer = newBuffer }
            | _ ->
                match findCursorParagraph doc sel.Anchor.BlockId with
                | Some para ->
                    let newBuffer = TextBuffer.setFontFamily formatStart formatLen family para.Buffer
                    if obj.ReferenceEquals(newBuffer, para.Buffer) then
                        doc
                    else
                        replaceParagraph doc para.Id { para with Buffer = newBuffer }
                | None -> doc
    | Some sel ->
        applySelectionBufferFormat doc sel (fun start len buffer ->
            TextBuffer.setFontFamily start len family buffer)

let private selectionTextColdCacheParagraphLimit = 4096

let private selectionTextSourceSmallEnoughForColdCache doc =
    let mutable count = 0
    let mutable smallEnough = true
    iterEditableParagraphsUntil doc (fun _ ->
        count <- count + 1
        if count > selectionTextColdCacheParagraphLimit then
            smallEnough <- false
            false
        else
            true)
    smallEnough

let private sameBlockId (left: BlockId) (right: BlockId) =
    blockIdKeyEquals left right

let private samePosition (left: Position) (right: Position) =
    left.Offset = right.Offset
    && sameBlockId left.BlockId right.BlockId

/// Get selected text as plain text
let getSelectionText doc =
    match cachedSelectionText doc with
    | ValueSome text -> text
    | _ ->
        let text =
            match doc.Selection with
            | None -> ""
            | Some sel when samePosition sel.Anchor sel.Focus -> ""  // No selection
            | Some sel when sameBlockId sel.Anchor.BlockId sel.Focus.BlockId ->
                let sliceStart = min sel.Anchor.Offset sel.Focus.Offset
                let sliceEnd = max sel.Anchor.Offset sel.Focus.Offset
                let sliceLen = sliceEnd - sliceStart
                if sliceLen <= 0 then
                    ""
                else
                    match findCursorParagraph doc sel.Anchor.BlockId with
                    | None -> ""
                    | Some para -> TextBuffer.slice sliceStart sliceLen para.Buffer
            | Some sel ->
                let selectedText = System.Text.StringBuilder()
                let mutable firstParagraph = true

                let appendSlice (para: ParagraphBlock) startOffset endOffset =
                    let bufferLen = TextBuffer.length para.Buffer
                    let sliceStart = max 0 (min bufferLen startOffset)
                    let sliceEnd = max sliceStart (min bufferLen endOffset)
                    if not firstParagraph then
                        selectedText.Append("\n") |> ignore
                    firstParagraph <- false
                    if sliceEnd > sliceStart then
                        TextBuffer.appendSliceToStringBuilder
                            selectedText
                            sliceStart
                            (sliceEnd - sliceStart)
                            para.Buffer

                let mutable collecting = false
                let mutable foundBoth = false

                iterEditableParagraphsUntil doc (fun para ->
                    let isAnchor = para.Id = sel.Anchor.BlockId
                    let isFocus = para.Id = sel.Focus.BlockId
                    if isAnchor || isFocus then
                        if not collecting then
                            collecting <- true
                            let startOffset = if isAnchor then sel.Anchor.Offset else sel.Focus.Offset
                            appendSlice para startOffset (TextBuffer.length para.Buffer)
                        else
                            let endOffset = if isAnchor then sel.Anchor.Offset else sel.Focus.Offset
                            appendSlice para 0 endOffset
                            foundBoth <- true
                    elif collecting then
                        appendSlice para 0 (TextBuffer.length para.Buffer)
                    not foundBoth)

                if foundBoth then selectedText.ToString() else ""
        if text.Length > 0
           && selectionTextSourceSmallEnoughForColdCache doc then
            rememberSelectionText doc text
        text

let private updateCurrentParagraphMaybe (update: ParagraphBlock -> ParagraphBlock option) doc =
    match doc.Selection with
    | None -> doc
    | Some sel ->
        let rec updateTopLevel remaining =
            match remaining with
            | [] -> ValueNone
            | block :: tail ->
                match block with
                | Paragraph para when para.Id = sel.Focus.BlockId ->
                    match update para with
                    | Some nextPara -> ValueSome (replaceFirstTopLevelBlock doc para.Id (Paragraph nextPara))
                    | None -> ValueSome doc
                | Heading (level, para) when para.Id = sel.Focus.BlockId ->
                    match update para with
                    | Some nextPara -> ValueSome (replaceFirstTopLevelBlock doc para.Id (Heading (level, nextPara)))
                    | None -> ValueSome doc
                | _ ->
                    updateTopLevel tail

        match updateTopLevel doc.Blocks with
        | ValueSome nextDoc -> nextDoc
        | ValueNone ->
            let mutable nestedChanged = false
            let nextDoc =
                updateParagraph doc sel.Focus.BlockId (fun para ->
                    match update para with
                    | Some nextPara when nextPara <> para ->
                        nestedChanged <- true
                        nextPara
                    | Some _ ->
                        para
                    | None -> para)
            if nestedChanged then nextDoc else doc

let private updateCurrentParagraph (update: ParagraphBlock -> ParagraphBlock) doc =
    updateCurrentParagraphMaybe (update >> Some) doc

let private updateCurrentParagraphMetadata update doc =
    updateCurrentParagraphMaybe
        (fun para ->
            let nextMetadata = update para.Metadata
            if nextMetadata = para.Metadata then
                None
            else
                Some { para with Metadata = nextMetadata })
        doc

let private updateCurrentParagraphMetadataMaybe update doc =
    updateCurrentParagraphMaybe
        (fun para ->
            update para.Metadata
            |> Option.map (fun metadata -> { para with Metadata = metadata }))
        doc

/// Toggle bullet list on current block
let toggleBulletList doc =
    updateCurrentParagraphMetadata
        (fun metadata ->
            if metadata.ListStyle = Bullet then
                { metadata with ListStyle = NoList; ListLevel = 0 }
            else
                { metadata with ListStyle = Bullet; ListLevel = max 1 metadata.ListLevel })
        doc

/// Toggle numbered list on current block
let toggleNumberedList doc =
    updateCurrentParagraphMetadata
        (fun metadata ->
            if metadata.ListStyle = Numbered then
                { metadata with ListStyle = NoList; ListLevel = 0 }
            else
                { metadata with ListStyle = Numbered; ListLevel = max 1 metadata.ListLevel })
        doc

/// Apply a numbering marker format to the current block and ensure it is a numbered list.
let setNumberFormat format doc =
    let normalizedFormat =
        match format with
        | "lowerLetter" | "upperLetter" | "lowerRoman" | "upperRoman" -> format
        | _ -> "decimal"
    updateCurrentParagraphMetadata
        (fun metadata ->
            { metadata with
                ListStyle = Numbered
                ListLevel = max 1 metadata.ListLevel
                ListNumberFormat = Some normalizedFormat
                ListLevelText = None })
        doc

/// Increase list indentation (Tab)
let increaseIndent doc =
    updateCurrentParagraphMetadataMaybe
        (fun metadata ->
            if metadata.ListStyle <> NoList && metadata.ListLevel < 9 then
                Some { metadata with ListLevel = metadata.ListLevel + 1 }
            else
                None)
        doc

/// Decrease list indentation (Shift+Tab)
let decreaseIndent doc =
    updateCurrentParagraphMetadataMaybe
        (fun metadata ->
            if metadata.ListStyle <> NoList && metadata.ListLevel > 0 then
                let newLevel = metadata.ListLevel - 1
                if newLevel = 0 then
                    Some { metadata with ListStyle = NoList; ListLevel = 0 }
                else
                    Some { metadata with ListLevel = newLevel }
            else
                None)
        doc


/// Set alignment to left
let alignLeft doc =
    updateCurrentParagraphMetadata (fun metadata -> { metadata with Alignment = Left }) doc

/// Set alignment to center
let alignCenter doc =
    updateCurrentParagraphMetadata (fun metadata -> { metadata with Alignment = Center }) doc

/// Set alignment to right
let alignRight doc =
    updateCurrentParagraphMetadata (fun metadata -> { metadata with Alignment = Right }) doc

/// Set alignment to justify
let alignJustify doc =
    updateCurrentParagraphMetadata (fun metadata -> { metadata with Alignment = Justify }) doc

let setCurrentParagraphStyle alignment indentLeft indentRight spacingBefore spacingAfter lineSpacing firstLineIndent backgroundColor doc =
    updateCurrentParagraphMetadata
        (fun metadata ->
            { metadata with
                Alignment = alignment
                IndentLeft = indentLeft
                IndentRight = indentRight
                SpacingBefore = spacingBefore
                SpacingAfter = spacingAfter
                LineSpacing = lineSpacing
                LineSpacingRule = Some "auto"
                FirstLineIndent = firstLineIndent
                BackgroundColor = backgroundColor })
        doc

let setCurrentSectionBreak sectionBreak doc =
    updateCurrentParagraphMetadata (fun metadata -> { metadata with SectionBreak = sectionBreak }) doc

let setCurrentHeaderFooterReferences headerReferences footerReferences doc =
    updateCurrentParagraphMetadata
        (fun metadata -> { metadata with HeaderReferences = headerReferences; FooterReferences = footerReferences })
        doc
