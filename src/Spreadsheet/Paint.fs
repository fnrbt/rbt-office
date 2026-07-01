module RedBat.Spreadsheet.Paint

open System.Collections.Generic
open System.Runtime.CompilerServices
open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Layout
open RedBat.Spreadsheet.CellText

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private clearArray (items: 'T array) =
    for i = 0 to items.Length - 1 do
        items.[i] <- Unchecked.defaultof<'T>

type Color = string

type TextAlign =
    | AlignLeft
    | AlignCenter
    | AlignRight

type VerticalAlign =
    | VAlignTop
    | VAlignMiddle
    | VAlignBottom

[<Struct>]
type PaintText = {
    Text: string
    /// The full cell (or header) rectangle. Per-line offset is derived from
    /// LineIndex/LineCount and font metrics in the renderer — Paint.fs does
    /// not pack baseline math into the rect.
    Rect: Rect
    Color: Color
    /// Font size in points, as written in the source XLSX. Each renderer
    /// converts to its native unit (px for raster, pt for PDF, css-pt for DOM).
    FontSize: float
    FontName: string option
    Bold: bool
    Italic: bool
    /// Render a horizontal stroke under the text (XLSX `<u/>` font flag).
    Underline: bool
    /// Render two parallel strokes for `<u val="double"/>` instead of one.
    UnderlineDouble: bool
    /// Render a horizontal stroke through the text (XLSX `<strike/>` flag).
    Strikethrough: bool
    Align: TextAlign
    /// Vertical alignment within Rect. Excel "General" = VAlignBottom.
    VerticalAlign: VerticalAlign
    /// 0-based line index within a wrapped block; 0 if not wrapped.
    LineIndex: int
    /// Total lines in the wrapped block; 1 if not wrapped.
    LineCount: int
    /// Signed pixel offset the renderer adds to the computed baseline.
    /// Used for superscript (negative — shift up) and subscript (positive
    /// — shift down). Zero for normal text.
    BaselineShift: float
    /// Cell-text rotation in degrees. 0 = horizontal (default), 1..90 =
    /// CCW (text rises left-to-right), 91..180 = CW (encoded as
    /// 90 + |angle|), 255 = stacked (each glyph on its own line). The
    /// renderer rotates the glyph stream around the cell's anchor; the
    /// rect itself is unrotated.
    Rotation: int
}

/// A single segment of a vector path. Coordinates are in the same px space
/// as everything else in PaintItem. Mirrors the renderer / canvas / PDF path
/// primitives so a chart scene can be expressed once and drawn by all three.
type VecSeg =
    | VMove of float * float
    | VLine of float * float
    | VCubic of float * float * float * float * float * float
    | VClose

type PaintItem =
    | FillRect of Rect * Color
    | StrokeRect of Rect * Color * float
    | DrawText of PaintText
    | GridLine of float * float * float * float * Color * float
    /// Arbitrary vector path with an optional fill and an optional
    /// (color, width) stroke — the primitive charts (and future drawings)
    /// emit. At least one of fill/stroke is expected.
    | VecPath of VecSeg list * fill: Color option * stroke: (Color * float) option
    /// A nested scene transformed by a 2×3 affine (a, b, c, d, e, f), i.e.
    /// x' = a·x + c·y + e, y' = b·x + d·y + f. Used to place/scale a chart
    /// scene into its anchor rect.
    | VecGroup of (float * float * float * float * float * float) * PaintItem list

[<Struct>]
type ViewportPaintItem = {
    Zone: ViewportZone
    Item: PaintItem
}

type SparseDefaultPaintParts = {
    PrefixItems: PaintItem list
    SelectionFills: PaintItem list
    GridItems: PaintItem list
    SelectionStrokes: PaintItem list
    TextItems: PaintItem list
    SuffixItems: PaintItem list
}

[<Struct>]
type private CellPaintState = {
    Cell: CellGeometry
    Format: CellFormat
    Value: CellValue
    Selected: bool
    Background: string
    Zone: ViewportZone
}

/// A hand-built vector scene used to prove the VecPath/VecGroup primitives
/// render identically through the native, canvas, and PDF paths (the
/// foundation a chart adapter would emit). Exercises a filled triangle, a
/// stroked polyline, a filled+stroked box, and a `“data bar”` strip — all
/// inside a translated group plus standalone axes.
let vectorTestScene : PaintItem list =
    [ // Plot frame (axes) as a stroked path, no group.
      VecPath ([ VMove (40.0, 20.0); VLine (40.0, 200.0); VLine (320.0, 200.0) ], None, Some ("#808080", 1.0))
      // Everything else placed by a group translated to the plot origin.
      VecGroup ((1.0, 0.0, 0.0, 1.0, 40.0, 200.0),
        [ // Filled area (semi-triangle) under a curve — fill only.
          VecPath ([ VMove (0.0, 0.0); VLine (0.0, -60.0); VCubic (40.0, -120.0, 90.0, -40.0, 130.0, -90.0); VLine (130.0, 0.0); VClose ],
                   Some "#BDD7EE", None)
          // Line on top of the area — stroke only.
          VecPath ([ VMove (0.0, -60.0); VCubic (40.0, -120.0, 90.0, -40.0, 130.0, -90.0) ], None, Some ("#1F497D", 2.0))
          // Three column bars — filled + thin stroke.
          VecPath ([ VMove (160.0, 0.0); VLine (190.0, 0.0); VLine (190.0, -70.0); VLine (160.0, -70.0); VClose ], Some "#70AD47", Some ("#000000", 0.75))
          VecPath ([ VMove (200.0, 0.0); VLine (230.0, 0.0); VLine (230.0, -110.0); VLine (200.0, -110.0); VClose ], Some "#70AD47", Some ("#000000", 0.75))
          VecPath ([ VMove (240.0, 0.0); VLine (270.0, 0.0); VLine (270.0, -45.0); VLine (240.0, -45.0); VClose ], Some "#70AD47", Some ("#000000", 0.75)) ]) ]

/// Excel's General horizontal alignment puts numbers/booleans on the right,
/// strings on the left, and errors in the center. Explicit alignment on
/// CellFormat overrides this default.
let private generalAlign (value: CellValue): TextAlign =
    match unwrapFormulaValue value with
    | Number _ -> AlignRight
    | Boolean _ -> AlignCenter
    | Error _ -> AlignCenter
    | CellValue.Empty -> AlignLeft
    | _ -> AlignLeft

let private equalsIgnoreCase (left: string) (right: string) =
    System.String.Equals(left, right, System.StringComparison.OrdinalIgnoreCase)

let private optionEqualsIgnoreCase expected value =
    match value with
    | Some actual -> equalsIgnoreCase actual expected
    | None -> false

let private cellKey row col =
    ((int64 row) <<< 32) ||| int64 col

let private isTextNumberFormat (value: string) =
    let mutable start = 0
    let mutable last = value.Length - 1
    while start <= last && System.Char.IsWhiteSpace value.[start] do
        start <- start + 1
    while last >= start && System.Char.IsWhiteSpace value.[last] do
        last <- last - 1
    start = last && start >= 0 && value.[start] = '@'

let private singleLineText (value: string) =
    let mutable firstBreak = -1
    let mutable i = 0
    while firstBreak < 0 && i < value.Length do
        if value.[i] = '\n' then firstBreak <- i
        i <- i + 1
    if firstBreak < 0 then value
    else
        let sb = System.Text.StringBuilder(value.Length - 1)
        sb.Append(value, 0, firstBreak) |> ignore
        for j = firstBreak + 1 to value.Length - 1 do
            if value.[j] <> '\n' then
                sb.Append(value.[j]) |> ignore
        sb.ToString()

let private alignFromFormat (format: CellFormat) (value: CellValue): TextAlign =
    match format.HorizontalAlign with
    | Some align when equalsIgnoreCase align "center" -> AlignCenter
    // "Center Across Selection" — the text is centred (across a run of
    // centerContinuous cells; the run rect is computed at paint time).
    | Some align when equalsIgnoreCase align "centerContinuous" -> AlignCenter
    | Some align when equalsIgnoreCase align "right" -> AlignRight
    | Some align when equalsIgnoreCase align "left" -> AlignLeft
    | _ ->
        // A `@` (Text) number format makes Excel treat the value as text, so
        // even a number is General-aligned left rather than right.
        match format.NumberFormat with
        | Some nf when isTextNumberFormat nf -> AlignLeft
        | _ -> generalAlign value

let private isSelected (selection: Selection) row col =
    match selection with
    | SingleCell (r, c) -> r = row && c = col
    | Range (sr, sc, er, ec) ->
        row >= sr && row <= er && col >= sc && col <= ec

let private hasExplicitBorder (format: CellFormat) =
    format.BorderTop.IsSome
    || format.BorderBottom.IsSome
    || format.BorderLeft.IsSome
    || format.BorderRight.IsSome
    || format.BorderDiagonal.IsSome

let private isWhiteFill color =
    equalsIgnoreCase color "#ffffff" || equalsIgnoreCase color "#fff" || equalsIgnoreCase color "white"

let private hasRenderTimeTableStyle = function
    | { Style = Some style } -> style.HeaderFill.IsSome
    | _ -> false

type private RenderTimeTableStyleCacheEntry = {
    Tables: Table list
    TableCount: int
    SourceTextCharCount: int
    HasRenderTimeStyle: bool
}

let mutable private renderTimeTableStyleCache: RenderTimeTableStyleCacheEntry list = []
let private renderTimeTableStyleCacheByTables =
    Dictionary<Table list, RenderTimeTableStyleCacheEntry>(
        { new IEqualityComparer<Table list> with
            member _.Equals(left, right) = obj.ReferenceEquals(left, right)
            member _.GetHashCode(value) = RuntimeHelpers.GetHashCode(value) })
let private renderTimeTableStyleFastSlots = Array.zeroCreate<RenderTimeTableStyleCacheEntry> 256
let private renderTimeTableStyleCacheMaxEntries = 64
let private renderTimeTableStyleCacheMaxTotalTables = 200000
let private renderTimeTableStyleCacheMaxTotalSourceTextChars = 5000000

let private renderTimeAddCapped limit value total =
    if total > limit then total
    else min (limit + 1) (total + max 0 value)

let private renderTimeTableStyleCacheEntryMatches tables (entry: RenderTimeTableStyleCacheEntry) =
    obj.ReferenceEquals(entry.Tables, tables)

let private renderTimeTableStyleCacheEntryRetainable (entry: RenderTimeTableStyleCacheEntry) =
    entry.TableCount <= renderTimeTableStyleCacheMaxTotalTables
    && entry.SourceTextCharCount <= renderTimeTableStyleCacheMaxTotalSourceTextChars

let private renderTimeTableStyleSlot tables slotCount =
    RuntimeHelpers.GetHashCode(tables) &&& (slotCount - 1)

let private setRenderTimeTableStyleCache entries =
    renderTimeTableStyleCache <- entries
    renderTimeTableStyleCacheByTables.Clear()
    clearArray renderTimeTableStyleFastSlots
    for entry in entries do
        renderTimeTableStyleCacheByTables.[entry.Tables] <- entry
        renderTimeTableStyleFastSlots.[renderTimeTableStyleSlot entry.Tables renderTimeTableStyleFastSlots.Length] <- entry

let private tryGetRenderTimeTableStyleCache tables =
    let slot = renderTimeTableStyleSlot tables renderTimeTableStyleFastSlots.Length
    let fast = renderTimeTableStyleFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && renderTimeTableStyleCacheEntryMatches tables fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<RenderTimeTableStyleCacheEntry>
        if renderTimeTableStyleCacheByTables.TryGetValue(tables, &cached) then
            renderTimeTableStyleFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private renderTimeTableStyleCacheWithoutMatching tables =
    let mutable cached = Unchecked.defaultof<RenderTimeTableStyleCacheEntry>
    if renderTimeTableStyleCacheByTables.TryGetValue(tables, &cached) then
        renderTimeTableStyleCache
        |> List.filter (renderTimeTableStyleCacheEntryMatches tables >> not)
    else
        renderTimeTableStyleCache

let private renderTimeOptionStringCharCount (value: string option) =
    match value with
    | Some text -> text.Length
    | None -> 0

let private renderTimeTableStyleStringCharCount (style: TableStyle option) =
    match style with
    | Some style ->
        style.Name.Length
        + renderTimeOptionStringCharCount style.HeaderFill
        + renderTimeOptionStringCharCount style.StripeFill
    | None -> 0

let private renderTimeTableSourceTextCharCount (tables: Table list) =
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = tables
    let limit = renderTimeTableStyleCacheMaxTotalSourceTextChars
    while total <= renderTimeTableStyleCacheMaxTotalSourceTextChars
          && seen <= renderTimeTableStyleCacheMaxTotalTables
          && not remaining.IsEmpty do
        match remaining with
        | table :: tail ->
            total <- renderTimeAddCapped limit table.Name.Length total
            if total <= limit then
                total <- renderTimeAddCapped limit (renderTimeTableStyleStringCharCount table.Style) total
            seen <- seen + 1
            remaining <- tail
        | [] -> ()
    total

let private renderTimeTableCount (tables: Table list) =
    let mutable count = 0
    let mutable remaining = tables
    while count <= renderTimeTableStyleCacheMaxTotalTables && not remaining.IsEmpty do
        count <- count + 1
        remaining <- remaining.Tail
    count

let private renderTimeTableSourceTextCharCountFor tableCount tables =
    if tableCount > renderTimeTableStyleCacheMaxTotalTables then
        renderTimeTableStyleCacheMaxTotalSourceTextChars + 1
    else
        renderTimeTableSourceTextCharCount tables

let private rememberRenderTimeTableStyle tables hasRenderTimeStyle =
    let tableCount = max 1 (renderTimeTableCount tables)
    let entry =
        { Tables = tables
          TableCount = tableCount
          SourceTextCharCount = renderTimeTableSourceTextCharCountFor tableCount tables
          HasRenderTimeStyle = hasRenderTimeStyle }
    if not (renderTimeTableStyleCacheEntryRetainable entry) then
        let entries = renderTimeTableStyleCacheWithoutMatching tables
        if not (obj.ReferenceEquals(entries, renderTimeTableStyleCache)) then
            setRenderTimeTableStyleCache entries
    else
        let trim entries =
            let rec loop keptCount keptTables keptSourceTextChars acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let candidateTables = max 1 candidate.TableCount
                    let candidateSourceTextChars = candidate.SourceTextCharCount
                    if renderTimeTableStyleCacheEntryRetainable candidate
                       && keptCount < renderTimeTableStyleCacheMaxEntries
                       && keptTables + candidateTables <= renderTimeTableStyleCacheMaxTotalTables
                       && keptSourceTextChars + candidateSourceTextChars <= renderTimeTableStyleCacheMaxTotalSourceTextChars then
                        loop
                            (keptCount + 1)
                            (keptTables + candidateTables)
                            (keptSourceTextChars + candidateSourceTextChars)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptTables keptSourceTextChars acc tail
            loop 0 0 0 [] entries
        entry :: renderTimeTableStyleCacheWithoutMatching tables
        |> trim
        |> setRenderTimeTableStyleCache
    hasRenderTimeStyle

let private hasAnyRenderTimeTableStyle (tables: Table list) =
    match tryGetRenderTimeTableStyleCache tables with
    | ValueSome cached -> cached.HasRenderTimeStyle
    | ValueNone ->
        let mutable found = false
        let mutable remaining = tables
        while not found && not remaining.IsEmpty do
            match remaining with
            | table :: tail ->
                found <- hasRenderTimeTableStyle table
                remaining <- tail
            | [] -> ()
        rememberRenderTimeTableStyle tables found

type private StyledTable = {
    StartRow0: int
    EndRow0: int
    StartCol0: int
    EndCol0: int
    HeaderRowCount: int
    Style: TableStyle
    AutoFilter: (int * int) option
}

type private TableStyleLookup = {
    Rows: Dictionary<int, ResizeArray<StyledTable>>
    IsEmpty: bool
}

let private emptyTableStyleLookup = {
    Rows = Dictionary<int, ResizeArray<StyledTable>>()
    IsEmpty = true
}

let private tryStyledTable (table: Table) =
    match table.Style with
    | Some style when style.HeaderFill.IsSome ->
        Some {
            StartRow0 = table.StartRow - 1
            EndRow0 = table.EndRow - 1
            StartCol0 = table.StartCol - 1
            EndCol0 = table.EndCol - 1
            HeaderRowCount = table.HeaderRowCount
            Style = style
            AutoFilter = table.AutoFilter
        }
    | _ -> None

let private buildTableStyleLookup (model: SpreadsheetModel) (visible: VisibleRange) =
    if not (hasAnyRenderTimeTableStyle model.Tables) then
        emptyTableStyleLookup
    else
        let mutable rows = ValueNone
        let rowLookup () =
            match rows with
            | ValueSome rowLookup -> rowLookup
            | ValueNone ->
                let rowLookup = Dictionary<int, ResizeArray<StyledTable>>()
                rows <- ValueSome rowLookup
                rowLookup
        for table in model.Tables do
            match tryStyledTable table with
            | Some styled
                when styled.EndRow0 >= visible.StartRow
                  && styled.StartRow0 < visible.EndRowExclusive
                  && styled.EndCol0 >= visible.StartCol
                  && styled.StartCol0 < visible.EndColExclusive ->
                let rowLookup = rowLookup ()
                let firstRow = max visible.StartRow styled.StartRow0
                let lastRow = min (visible.EndRowExclusive - 1) styled.EndRow0
                for row in firstRow .. lastRow do
                    match rowLookup.TryGetValue row with
                    | true, rowTables -> rowTables.Add styled
                    | false, _ ->
                        let rowTables = ResizeArray<StyledTable>()
                        rowTables.Add styled
                        rowLookup.[row] <- rowTables
            | _ -> ()
        match rows with
        | ValueSome rows -> { Rows = rows; IsEmpty = false }
        | ValueNone -> emptyTableStyleLookup

let private tableStyleFormat row col (lookup: TableStyleLookup) (format: CellFormat) =
    let mutable fmt = format
    let mutable rowTables = Unchecked.defaultof<ResizeArray<StyledTable>>
    if lookup.Rows.TryGetValue(row, &rowTables) then
        for table in rowTables do
            if col >= table.StartCol0 && col <= table.EndCol0 then
                let style = table.Style
                let tableRow0 = table.StartRow0
                let isHeader = row < tableRow0 + table.HeaderRowCount
                let dataIdx = row - tableRow0 - table.HeaderRowCount
                let isStripe = style.ShowRowStripes && dataIdx >= 0 && dataIdx % 2 = 0 && style.StripeFill.IsSome
                if isHeader then
                    let hasFilter =
                        match table.AutoFilter with
                        | Some (firstCol, lastCol) -> (col + 1) >= firstCol && (col + 1) <= lastCol
                        | None -> false
                    fmt <-
                        { fmt with
                            BgColor = fmt.BgColor |> Option.orElse style.HeaderFill
                            TextColor = fmt.TextColor |> Option.orElse (Some "#FFFFFF")
                            Bold = true
                            FilterButton = fmt.FilterButton || hasFilter }
                elif isStripe then
                    fmt <- { fmt with BgColor = fmt.BgColor |> Option.orElse style.StripeFill }
    fmt

let private iterHeaderPaintItems (model: SpreadsheetModel) (viewport: Viewport) (emit: PaintItem -> unit) =
    iterHeaders model viewport (function
        | CornerHeader rect ->
            emit (FillRect (rect, "#e8e8e8"))
            emit (StrokeRect (rect, "#a6a6a6", 1.0))
        | ColumnHeader (col, rect) ->
            emit (FillRect (rect, "#f2f2f2"))
            emit (StrokeRect (rect, "#a6a6a6", 1.0))
            emit (DrawText {
                Text = columnToLetter col
                Rect = rect
                Color = "#444444"
                FontSize = 11.0
                FontName = None
                Bold = false
                Italic = false
                Underline = false
                UnderlineDouble = false
                Strikethrough = false
                Align = AlignCenter
                VerticalAlign = VAlignMiddle
                LineIndex = 0
                LineCount = 1
                BaselineShift = 0.0
                Rotation = 0
            })
        | RowHeader (row, rect) ->
            emit (FillRect (rect, "#f2f2f2"))
            emit (StrokeRect (rect, "#a6a6a6", 1.0))
            emit (DrawText {
                Text = string (row + 1)
                Rect = rect
                Color = "#444444"
                FontSize = 11.0
                FontName = None
                Bold = false
                Italic = false
                Underline = false
                UnderlineDouble = false
                Strikethrough = false
                Align = AlignCenter
                VerticalAlign = VAlignMiddle
                LineIndex = 0
                LineCount = 1
                BaselineShift = 0.0
                Rotation = 0
            }))

let private buildHeaderPaintItems (model: SpreadsheetModel) (viewport: Viewport): PaintItem list =
    let items = ResizeArray<PaintItem>()
    iterHeaderPaintItems model viewport items.Add
    resizeArrayToList items

/// Translate an OOXML border style name into an approximate stroke width in
/// CSS pixels. Excel has many variants (hair, thin, medium, thick, double,
/// dashed/dotted in three weights, slantDashDot, etc.); we collapse them into
/// thin/medium/thick stroke widths and a default black color.
let private borderStrokeWidth (style: string) : float =
    if equalsIgnoreCase style "hair" then 0.5
    elif equalsIgnoreCase style "medium"
        || equalsIgnoreCase style "mediumdashed"
        || equalsIgnoreCase style "mediumdotted"
        || equalsIgnoreCase style "mediumdashdot"
        || equalsIgnoreCase style "mediumdashdotdot" then 1.75
    elif equalsIgnoreCase style "thick" then 3.0
    elif equalsIgnoreCase style "double" then 2.0
    elif equalsIgnoreCase style "slantdashdot" then 1.5
    else 1.0

/// Border-collapse priority for `<border style>` values. NOT the same as
/// stroke width — Excel's collapse rule follows the OOXML hierarchy where
/// `double` outranks `thick` even though `thick` paints fatter, because
/// `double` carries more visual semantics. Order: None < Hair < Dotted <
/// Dashed-family < Thin < Medium < Thick < Double. Equal priority resolves
/// in the caller's tiebreak.
let private borderPriority (style: string) : int =
    if style = "" then 0
    elif equalsIgnoreCase style "hair" then 1
    elif equalsIgnoreCase style "dotted" then 2
    elif equalsIgnoreCase style "dashed" then 3
    elif equalsIgnoreCase style "dashdot" then 4
    elif equalsIgnoreCase style "dashdotdot" then 5
    elif equalsIgnoreCase style "mediumdashed" then 6
    elif equalsIgnoreCase style "mediumdashdot" then 7
    elif equalsIgnoreCase style "mediumdashdotdot" then 8
    elif equalsIgnoreCase style "slantdashdot" then 9
    elif equalsIgnoreCase style "thin" then 10
    elif equalsIgnoreCase style "medium" then 11
    elif equalsIgnoreCase style "thick" then 12
    elif equalsIgnoreCase style "double" then 13
    else 10  // unknown → treat as thin

let private borderColor (side: BorderStyle) : string =
    side.Color |> Option.defaultValue "#000000"

type private SparseDefaultPaintGeometryKey = {
    ScrollX: float
    ScrollY: float
    Width: float
    Height: float
    HeaderHeight: float
    RowHeaderWidth: float
}

type private SparseDefaultPaintCacheEntry = {
    Cells: Map<CellAddress, CellData>
    Tables: Table list
    ConditionalFormats: Map<CellAddress, ConditionalFormat list>
    MergedCells: MergedCell list
    RichRuns: Map<CellAddress, (string * RunOverride) list>
    RowHeights: Map<int, float>
    ColumnWidths: Map<int, float>
    Filters: Map<int, ColumnFilter>
    FilterHiddenRows: int array
    Spill: Map<CellAddress, CellAddress>
    Validations: Map<CellAddress, DataValidation>
    SearchResults: CellAddress array
    CfBoldCells: Set<CellAddress>
    PageSetup: PageSetup
    RowCount: int
    ColCount: int
    Geometry: SparseDefaultPaintGeometryKey
    Measurer: TextMeasurer option
    PrefixItems: PaintItem list
    VisibleRows: (int * float * float) array
    VisibleCols: (int * float * float) array
    GridItems: PaintItem list
    TextItems: PaintItem list
    ItemCount: int
    TextCharCount: int
    SourceCellCount: int
    mutable CachedPartsSelection: Selection
    mutable CachedPartsPrefixItems: PaintItem list
    mutable CachedPartsSuffixItems: PaintItem list
    mutable CachedParts: SparseDefaultPaintParts option
}

type private SparseDefaultPaintCacheKey = {
    Cells: Map<CellAddress, CellData>
    Tables: Table list
    ConditionalFormats: Map<CellAddress, ConditionalFormat list>
    MergedCells: MergedCell list
    RichRuns: Map<CellAddress, (string * RunOverride) list>
    RowHeights: Map<int, float>
    ColumnWidths: Map<int, float>
    Filters: Map<int, ColumnFilter>
    FilterHiddenRows: int array
    Spill: Map<CellAddress, CellAddress>
    Validations: Map<CellAddress, DataValidation>
    SearchResults: CellAddress array
    CfBoldCells: Set<CellAddress>
    PageSetup: PageSetup
    RowCount: int
    ColCount: int
    Geometry: SparseDefaultPaintGeometryKey
    Measurer: TextMeasurer option
}

type private SparseDefaultPaintPartsLastEntry = {
    Cells: Map<CellAddress, CellData>
    Tables: Table list
    ConditionalFormats: Map<CellAddress, ConditionalFormat list>
    MergedCells: MergedCell list
    RichRuns: Map<CellAddress, (string * RunOverride) list>
    RowHeights: Map<int, float>
    ColumnWidths: Map<int, float>
    Filters: Map<int, ColumnFilter>
    FilterHiddenRows: int array
    Spill: Map<CellAddress, CellAddress>
    Validations: Map<CellAddress, DataValidation>
    SearchResults: CellAddress array
    CfBoldCells: Set<CellAddress>
    PageSetup: PageSetup
    RowCount: int
    ColCount: int
    Selection: Selection
    ScrollX: float
    ScrollY: float
    Width: float
    Height: float
    HeaderHeight: float
    RowHeaderWidth: float
    Measurer: TextMeasurer option
    Result: SparseDefaultPaintParts option
}

let private sparseSameMeasurer (left: TextMeasurer option) (right: TextMeasurer option) =
    match left, right with
    | None, None -> true
    | Some a, Some b -> obj.ReferenceEquals(a, b)
    | _ -> false

let private sparseMeasurerHash = function
    | Some measurer -> RuntimeHelpers.GetHashCode(measurer)
    | None -> 0

let private sparseDefaultPaintCacheKeyHash (key: SparseDefaultPaintCacheKey) =
    let mutable hashCode = RuntimeHelpers.GetHashCode(key.Cells)
    hashCode <- (hashCode * 397) ^^^ RuntimeHelpers.GetHashCode(key.Tables)
    hashCode <- (hashCode * 397) ^^^ RuntimeHelpers.GetHashCode(key.ConditionalFormats)
    hashCode <- (hashCode * 397) ^^^ RuntimeHelpers.GetHashCode(key.MergedCells)
    hashCode <- (hashCode * 397) ^^^ RuntimeHelpers.GetHashCode(key.RichRuns)
    hashCode <- (hashCode * 397) ^^^ RuntimeHelpers.GetHashCode(key.RowHeights)
    hashCode <- (hashCode * 397) ^^^ RuntimeHelpers.GetHashCode(key.ColumnWidths)
    hashCode <- (hashCode * 397) ^^^ RuntimeHelpers.GetHashCode(key.Filters)
    hashCode <- (hashCode * 397) ^^^ RuntimeHelpers.GetHashCode(key.FilterHiddenRows)
    hashCode <- (hashCode * 397) ^^^ RuntimeHelpers.GetHashCode(key.Spill)
    hashCode <- (hashCode * 397) ^^^ RuntimeHelpers.GetHashCode(key.Validations)
    hashCode <- (hashCode * 397) ^^^ RuntimeHelpers.GetHashCode(key.SearchResults)
    hashCode <- (hashCode * 397) ^^^ RuntimeHelpers.GetHashCode(key.CfBoldCells)
    hashCode <- (hashCode * 397) ^^^ RuntimeHelpers.GetHashCode(key.PageSetup)
    hashCode <- (hashCode * 397) ^^^ key.RowCount
    hashCode <- (hashCode * 397) ^^^ key.ColCount
    hashCode <- (hashCode * 397) ^^^ hash key.Geometry
    (hashCode * 397) ^^^ sparseMeasurerHash key.Measurer

let private sparseDefaultPaintCacheKeyMatches (left: SparseDefaultPaintCacheKey) (right: SparseDefaultPaintCacheKey) =
    obj.ReferenceEquals(left.Cells, right.Cells)
    && obj.ReferenceEquals(left.Tables, right.Tables)
    && obj.ReferenceEquals(left.ConditionalFormats, right.ConditionalFormats)
    && obj.ReferenceEquals(left.MergedCells, right.MergedCells)
    && obj.ReferenceEquals(left.RichRuns, right.RichRuns)
    && obj.ReferenceEquals(left.RowHeights, right.RowHeights)
    && obj.ReferenceEquals(left.ColumnWidths, right.ColumnWidths)
    && obj.ReferenceEquals(left.Filters, right.Filters)
    && obj.ReferenceEquals(left.FilterHiddenRows, right.FilterHiddenRows)
    && obj.ReferenceEquals(left.Spill, right.Spill)
    && obj.ReferenceEquals(left.Validations, right.Validations)
    && obj.ReferenceEquals(left.SearchResults, right.SearchResults)
    && obj.ReferenceEquals(left.CfBoldCells, right.CfBoldCells)
    && obj.ReferenceEquals(left.PageSetup, right.PageSetup)
    && left.RowCount = right.RowCount
    && left.ColCount = right.ColCount
    && left.Geometry = right.Geometry
    && sparseSameMeasurer left.Measurer right.Measurer

let private sparseDefaultPaintCacheKeyComparer =
    { new IEqualityComparer<SparseDefaultPaintCacheKey> with
        member _.Equals(left, right) = sparseDefaultPaintCacheKeyMatches left right
        member _.GetHashCode(value) = sparseDefaultPaintCacheKeyHash value }

let mutable private sparseDefaultPaintCache: SparseDefaultPaintCacheEntry list = []
let private sparseDefaultPaintCacheByKey =
    Dictionary<SparseDefaultPaintCacheKey, SparseDefaultPaintCacheEntry>(sparseDefaultPaintCacheKeyComparer)
let private sparseDefaultPaintFastSlots = Array.zeroCreate<SparseDefaultPaintCacheEntry> 256
let mutable private sparseDefaultPaintPartsLastSet = false
let mutable private sparseDefaultPaintPartsLast = Unchecked.defaultof<SparseDefaultPaintPartsLastEntry>
let mutable private sparseDefaultPaintPartsPreviousSet = false
let mutable private sparseDefaultPaintPartsPrevious = Unchecked.defaultof<SparseDefaultPaintPartsLastEntry>
let mutable private sparseDefaultPaintPartsOlder1Set = false
let mutable private sparseDefaultPaintPartsOlder1 = Unchecked.defaultof<SparseDefaultPaintPartsLastEntry>
let mutable private sparseDefaultPaintPartsOlder2Set = false
let mutable private sparseDefaultPaintPartsOlder2 = Unchecked.defaultof<SparseDefaultPaintPartsLastEntry>
let private sparseDefaultPaintCacheMaxEntries = 16
let private sparseDefaultPaintCacheMaxTotalItems = 200000
let private sparseDefaultPaintCacheMaxTotalTextChars = 5000000
let private sparseDefaultPaintCacheMaxTotalSourceCells = 200000

let private sparseDefaultPaintCacheEntryRetainable (entry: SparseDefaultPaintCacheEntry) =
    entry.ItemCount <= sparseDefaultPaintCacheMaxTotalItems
    && entry.TextCharCount <= sparseDefaultPaintCacheMaxTotalTextChars
    && entry.SourceCellCount <= sparseDefaultPaintCacheMaxTotalSourceCells

let private clearSparseDefaultPaintPartsLast () =
    sparseDefaultPaintPartsLastSet <- false
    sparseDefaultPaintPartsLast <- Unchecked.defaultof<SparseDefaultPaintPartsLastEntry>
    sparseDefaultPaintPartsPreviousSet <- false
    sparseDefaultPaintPartsPrevious <- Unchecked.defaultof<SparseDefaultPaintPartsLastEntry>
    sparseDefaultPaintPartsOlder1Set <- false
    sparseDefaultPaintPartsOlder1 <- Unchecked.defaultof<SparseDefaultPaintPartsLastEntry>
    sparseDefaultPaintPartsOlder2Set <- false
    sparseDefaultPaintPartsOlder2 <- Unchecked.defaultof<SparseDefaultPaintPartsLastEntry>

let private sparseDefaultPaintPartsLastMatches (model: SpreadsheetModel) (viewport: Viewport) (cached: SparseDefaultPaintPartsLastEntry) =
    obj.ReferenceEquals(cached.Cells, model.Cells)
    && obj.ReferenceEquals(cached.Tables, model.Tables)
    && obj.ReferenceEquals(cached.ConditionalFormats, model.ConditionalFormats)
    && obj.ReferenceEquals(cached.MergedCells, model.MergedCells)
    && obj.ReferenceEquals(cached.RichRuns, model.RichRuns)
    && obj.ReferenceEquals(cached.RowHeights, model.RowHeights)
    && obj.ReferenceEquals(cached.ColumnWidths, model.ColumnWidths)
    && obj.ReferenceEquals(cached.Filters, model.Filters)
    && obj.ReferenceEquals(cached.FilterHiddenRows, model.FilterHiddenRows)
    && obj.ReferenceEquals(cached.Spill, model.Spill)
    && obj.ReferenceEquals(cached.Validations, model.Validations)
    && obj.ReferenceEquals(cached.SearchResults, model.SearchResults)
    && obj.ReferenceEquals(cached.CfBoldCells, model.CfBoldCells)
    && obj.ReferenceEquals(cached.PageSetup, model.PageSetup)
    && cached.RowCount = model.RowCount
    && cached.ColCount = model.ColCount
    && cached.Selection = model.Selection
    && cached.ScrollX = viewport.ScrollX
    && cached.ScrollY = viewport.ScrollY
    && cached.Width = viewport.Width
    && cached.Height = viewport.Height
    && cached.HeaderHeight = viewport.HeaderHeight
    && cached.RowHeaderWidth = viewport.RowHeaderWidth
    && sparseSameMeasurer cached.Measurer viewport.Measurer

let private tryGetSparseDefaultPaintPartsLast model viewport =
    if sparseDefaultPaintPartsLastSet
       && sparseDefaultPaintPartsLastMatches model viewport sparseDefaultPaintPartsLast then
        ValueSome sparseDefaultPaintPartsLast.Result
    elif sparseDefaultPaintPartsPreviousSet
         && sparseDefaultPaintPartsLastMatches model viewport sparseDefaultPaintPartsPrevious then
        ValueSome sparseDefaultPaintPartsPrevious.Result
    elif sparseDefaultPaintPartsOlder1Set
         && sparseDefaultPaintPartsLastMatches model viewport sparseDefaultPaintPartsOlder1 then
        ValueSome sparseDefaultPaintPartsOlder1.Result
    elif sparseDefaultPaintPartsOlder2Set
         && sparseDefaultPaintPartsLastMatches model viewport sparseDefaultPaintPartsOlder2 then
        ValueSome sparseDefaultPaintPartsOlder2.Result
    else
        ValueNone

let private rememberSparseDefaultPaintPartsLast (model: SpreadsheetModel) (viewport: Viewport) result =
    match result with
    | Some _ ->
        if sparseDefaultPaintPartsOlder1Set then
            sparseDefaultPaintPartsOlder2Set <- true
            sparseDefaultPaintPartsOlder2 <- sparseDefaultPaintPartsOlder1
        if sparseDefaultPaintPartsPreviousSet then
            sparseDefaultPaintPartsOlder1Set <- true
            sparseDefaultPaintPartsOlder1 <- sparseDefaultPaintPartsPrevious
        if sparseDefaultPaintPartsLastSet then
            sparseDefaultPaintPartsPreviousSet <- true
            sparseDefaultPaintPartsPrevious <- sparseDefaultPaintPartsLast
        sparseDefaultPaintPartsLastSet <- true
        sparseDefaultPaintPartsLast <-
            { Cells = model.Cells
              Tables = model.Tables
              ConditionalFormats = model.ConditionalFormats
              MergedCells = model.MergedCells
              RichRuns = model.RichRuns
              RowHeights = model.RowHeights
              ColumnWidths = model.ColumnWidths
              Filters = model.Filters
              FilterHiddenRows = model.FilterHiddenRows
              Spill = model.Spill
              Validations = model.Validations
              SearchResults = model.SearchResults
              CfBoldCells = model.CfBoldCells
              PageSetup = model.PageSetup
              RowCount = model.RowCount
              ColCount = model.ColCount
              Selection = model.Selection
              ScrollX = viewport.ScrollX
              ScrollY = viewport.ScrollY
              Width = viewport.Width
              Height = viewport.Height
              HeaderHeight = viewport.HeaderHeight
              RowHeaderWidth = viewport.RowHeaderWidth
              Measurer = viewport.Measurer
              Result = result }
    | None -> ()
    result

let private sparseDefaultPaintGeometryKey (viewport: Viewport) : SparseDefaultPaintGeometryKey = {
    ScrollX = viewport.ScrollX
    ScrollY = viewport.ScrollY
    Width = viewport.Width
    Height = viewport.Height
    HeaderHeight = viewport.HeaderHeight
    RowHeaderWidth = viewport.RowHeaderWidth
}

let private sparseDefaultPaintCacheKey
    (model: SpreadsheetModel)
    (geometry: SparseDefaultPaintGeometryKey)
    measurer
    : SparseDefaultPaintCacheKey = {
    Cells = model.Cells
    Tables = model.Tables
    ConditionalFormats = model.ConditionalFormats
    MergedCells = model.MergedCells
    RichRuns = model.RichRuns
    RowHeights = model.RowHeights
    ColumnWidths = model.ColumnWidths
    Filters = model.Filters
    FilterHiddenRows = model.FilterHiddenRows
    Spill = model.Spill
    Validations = model.Validations
    SearchResults = model.SearchResults
    CfBoldCells = model.CfBoldCells
    PageSetup = model.PageSetup
    RowCount = model.RowCount
    ColCount = model.ColCount
    Geometry = geometry
    Measurer = measurer
}

let private sparseDefaultPaintCacheKeyFromEntry (entry: SparseDefaultPaintCacheEntry) : SparseDefaultPaintCacheKey = {
    Cells = entry.Cells
    Tables = entry.Tables
    ConditionalFormats = entry.ConditionalFormats
    MergedCells = entry.MergedCells
    RichRuns = entry.RichRuns
    RowHeights = entry.RowHeights
    ColumnWidths = entry.ColumnWidths
    Filters = entry.Filters
    FilterHiddenRows = entry.FilterHiddenRows
    Spill = entry.Spill
    Validations = entry.Validations
    SearchResults = entry.SearchResults
    CfBoldCells = entry.CfBoldCells
    PageSetup = entry.PageSetup
    RowCount = entry.RowCount
    ColCount = entry.ColCount
    Geometry = entry.Geometry
    Measurer = entry.Measurer
}

let private sparseDefaultPaintCacheKeyMatchesEntry (key: SparseDefaultPaintCacheKey) (cached: SparseDefaultPaintCacheEntry) =
    obj.ReferenceEquals(cached.Cells, key.Cells)
    && obj.ReferenceEquals(cached.Tables, key.Tables)
    && obj.ReferenceEquals(cached.ConditionalFormats, key.ConditionalFormats)
    && obj.ReferenceEquals(cached.MergedCells, key.MergedCells)
    && obj.ReferenceEquals(cached.RichRuns, key.RichRuns)
    && obj.ReferenceEquals(cached.RowHeights, key.RowHeights)
    && obj.ReferenceEquals(cached.ColumnWidths, key.ColumnWidths)
    && obj.ReferenceEquals(cached.Filters, key.Filters)
    && obj.ReferenceEquals(cached.FilterHiddenRows, key.FilterHiddenRows)
    && obj.ReferenceEquals(cached.Spill, key.Spill)
    && obj.ReferenceEquals(cached.Validations, key.Validations)
    && obj.ReferenceEquals(cached.SearchResults, key.SearchResults)
    && obj.ReferenceEquals(cached.CfBoldCells, key.CfBoldCells)
    && obj.ReferenceEquals(cached.PageSetup, key.PageSetup)
    && cached.RowCount = key.RowCount
    && cached.ColCount = key.ColCount
    && cached.Geometry = key.Geometry
    && sparseSameMeasurer cached.Measurer key.Measurer

let private sparseDefaultPaintCacheSlot (key: SparseDefaultPaintCacheKey) slotCount =
    sparseDefaultPaintCacheKeyHash key &&& (slotCount - 1)

let private setSparseDefaultPaintCache entries =
    sparseDefaultPaintCache <- entries
    sparseDefaultPaintCacheByKey.Clear()
    clearArray sparseDefaultPaintFastSlots
    clearSparseDefaultPaintPartsLast ()
    for entry in entries do
        let key = sparseDefaultPaintCacheKeyFromEntry entry
        sparseDefaultPaintCacheByKey.[key] <- entry
        sparseDefaultPaintFastSlots.[sparseDefaultPaintCacheSlot key sparseDefaultPaintFastSlots.Length] <- entry

let private tryGetSparseDefaultPaintCachedEntry (key: SparseDefaultPaintCacheKey) =
    let slot = sparseDefaultPaintCacheSlot key sparseDefaultPaintFastSlots.Length
    let fast = sparseDefaultPaintFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && sparseDefaultPaintCacheKeyMatchesEntry key fast then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<SparseDefaultPaintCacheEntry>
        if sparseDefaultPaintCacheByKey.TryGetValue(key, &cached) then
            sparseDefaultPaintFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private sparseDefaultPaintCacheEntryMatches (entry: SparseDefaultPaintCacheEntry) (cached: SparseDefaultPaintCacheEntry) =
    obj.ReferenceEquals(cached.Cells, entry.Cells)
    && obj.ReferenceEquals(cached.Tables, entry.Tables)
    && obj.ReferenceEquals(cached.ConditionalFormats, entry.ConditionalFormats)
    && obj.ReferenceEquals(cached.MergedCells, entry.MergedCells)
    && obj.ReferenceEquals(cached.RichRuns, entry.RichRuns)
    && obj.ReferenceEquals(cached.RowHeights, entry.RowHeights)
    && obj.ReferenceEquals(cached.ColumnWidths, entry.ColumnWidths)
    && obj.ReferenceEquals(cached.Filters, entry.Filters)
    && obj.ReferenceEquals(cached.FilterHiddenRows, entry.FilterHiddenRows)
    && obj.ReferenceEquals(cached.Spill, entry.Spill)
    && obj.ReferenceEquals(cached.Validations, entry.Validations)
    && obj.ReferenceEquals(cached.SearchResults, entry.SearchResults)
    && obj.ReferenceEquals(cached.CfBoldCells, entry.CfBoldCells)
    && obj.ReferenceEquals(cached.PageSetup, entry.PageSetup)
    && cached.RowCount = entry.RowCount
    && cached.ColCount = entry.ColCount
    && cached.Geometry = entry.Geometry
    && sparseSameMeasurer cached.Measurer entry.Measurer

let private sparseDefaultPaintCacheWithoutMatching entry =
    let key = sparseDefaultPaintCacheKeyFromEntry entry
    let mutable cached = Unchecked.defaultof<SparseDefaultPaintCacheEntry>
    if sparseDefaultPaintCacheByKey.TryGetValue(key, &cached) then
        sparseDefaultPaintCache
        |> List.filter (fun cached -> not (sparseDefaultPaintCacheEntryMatches entry cached))
    else
        sparseDefaultPaintCache

let private sparsePaintItemTextCharCount item =
    match item with
    | DrawText text -> text.Text.Length
    | _ -> 0

let private sparsePaintItemsTextCharCount (items: PaintItem list) =
    let mutable total = 0
    let mutable remaining = items
    while total <= sparseDefaultPaintCacheMaxTotalTextChars && not remaining.IsEmpty do
        total <- min (sparseDefaultPaintCacheMaxTotalTextChars + 1) (total + sparsePaintItemTextCharCount remaining.Head)
        remaining <- remaining.Tail
    total

let private sparseDefaultPrefixItems (model: SpreadsheetModel) (viewport: Viewport) =
    let prefix = ResizeArray<PaintItem>()
    prefix.Add(FillRect ({ X = 0.0; Y = 0.0; Width = viewport.Width; Height = viewport.Height }, "#ffffff"))
    iterHeaderPaintItems model viewport prefix.Add
    resizeArrayToList prefix

let private rememberSparseDefaultPaintCacheEntry entry =
    if not (sparseDefaultPaintCacheEntryRetainable entry) then
        let entries = sparseDefaultPaintCacheWithoutMatching entry
        if not (obj.ReferenceEquals(entries, sparseDefaultPaintCache)) then
            setSparseDefaultPaintCache entries
    else
        let trim entries =
            let rec loop keptCount keptItems keptTextChars keptSourceCells acc remaining =
                match remaining with
                | [] -> List.rev acc
                | candidate :: tail ->
                    let itemCount = max 1 candidate.ItemCount
                    let textChars = max 1 candidate.TextCharCount
                    let sourceCells = max 1 candidate.SourceCellCount
                    if keptCount < sparseDefaultPaintCacheMaxEntries
                       && keptItems + itemCount <= sparseDefaultPaintCacheMaxTotalItems
                       && keptTextChars + textChars <= sparseDefaultPaintCacheMaxTotalTextChars
                       && keptSourceCells + sourceCells <= sparseDefaultPaintCacheMaxTotalSourceCells then
                        loop
                            (keptCount + 1)
                            (keptItems + itemCount)
                            (keptTextChars + textChars)
                            (keptSourceCells + sourceCells)
                            (candidate :: acc)
                            tail
                    else
                        loop keptCount keptItems keptTextChars keptSourceCells acc tail
            loop 0 0 0 0 [] entries
        (entry :: sparseDefaultPaintCacheWithoutMatching entry)
        |> trim
        |> setSparseDefaultPaintCache

let private emitPaintItems (emit: ViewportZone -> PaintItem -> unit) (items: PaintItem list) =
    let mutable remaining = items
    while not remaining.IsEmpty do
        emit BodyZone remaining.Head
        remaining <- remaining.Tail

let private sparseDefaultSelectionParts
    (visibleRows: (int * float * float) array)
    (visibleCols: (int * float * float) array)
    (selection: Selection) =
    let sr, sc, er, ec =
        match selection with
        | SingleCell (row, col) -> row, col, row, col
        | Range (sr, sc, er, ec) -> sr, sc, er, ec
    let fills = ResizeArray<PaintItem>()
    let strokes = ResizeArray<PaintItem>()
    for row, y, height in visibleRows do
        if row >= sr && row <= er then
            for col, x, width in visibleCols do
                if col >= sc && col <= ec then
                    let rect = { X = x; Y = y; Width = width; Height = height }
                    fills.Add(FillRect (rect, "#E8F5E9"))
                    strokes.Add(StrokeRect (rect, "#217346", 2.0))
    resizeArrayToList fills, resizeArrayToList strokes

let private buildSparseDefaultPaintCacheEntry
    (model: SpreadsheetModel)
    (viewport: Viewport)
    (visible: VisibleRange)
    (geometry: SparseDefaultPaintGeometryKey) =
    let visibleCols = ResizeArray<int * float * float>()
    let mutable x = columnX model viewport visible.StartCol
    for col in visible.StartCol .. visible.EndColExclusive - 1 do
        let width = getColumnWidth col model
        if width > 0.0 then
            visibleCols.Add(col, x, width)
        x <- x + width

    let visibleRows = ResizeArray<int * float * float>()
    let mutable y = rowY model viewport visible.StartRow
    for row in visible.StartRow .. visible.EndRowExclusive - 1 do
        let height = getRowHeight row model
        if height > 0.0 then
            visibleRows.Add(row, y, height)
        y <- y + height

    if visibleRows.Count = 0 || visibleCols.Count = 0 then
        Some {
            Cells = model.Cells
            Tables = model.Tables
            ConditionalFormats = model.ConditionalFormats
            MergedCells = model.MergedCells
            RichRuns = model.RichRuns
            RowHeights = model.RowHeights
            ColumnWidths = model.ColumnWidths
            Filters = model.Filters
            FilterHiddenRows = model.FilterHiddenRows
            Spill = model.Spill
            Validations = model.Validations
            SearchResults = model.SearchResults
            CfBoldCells = model.CfBoldCells
            PageSetup = model.PageSetup
            RowCount = model.RowCount
            ColCount = model.ColCount
            Geometry = geometry
            Measurer = viewport.Measurer
            PrefixItems = sparseDefaultPrefixItems model viewport
            VisibleRows = visibleRows.ToArray()
            VisibleCols = visibleCols.ToArray()
            GridItems = []
            TextItems = []
            ItemCount = 0
            TextCharCount = 0
            SourceCellCount = model.Cells.Count
            CachedPartsSelection = SingleCell (-1, -1)
            CachedPartsPrefixItems = []
            CachedPartsSuffixItems = []
            CachedParts = None
        }
    else
        let textItems = ResizeArray<PaintItem>()
        let mutable compatible = true
        let mutable rowIndex = 0
        while compatible && rowIndex < visibleRows.Count do
            let row, y, height = visibleRows.[rowIndex]
            let mutable colIndex = 0
            while compatible && colIndex < visibleCols.Count do
                let col, x, width = visibleCols.[colIndex]
                let addr = { Row = row; Col = col }
                match getCell addr model with
                | Some data ->
                    if data.Format <> defaultFormat || data.Comment.IsSome then
                        compatible <- false
                    else
                        match data.Value with
                        | CellValue.Empty -> ()
                        | value ->
                            let fontPt = 11.0
                            let measure =
                                viewport.Measurer
                                |> Option.map (fun m ->
                                    fun (s: string) -> m s fontPt defaultFormat.FontName defaultFormat.Bold defaultFormat.Italic)
                            let rect = { X = x; Y = y; Width = width; Height = height }
                            let displayText = cellDisplayText measure width defaultFormat value
                            if displayText <> "" then
                                let textRect =
                                    { rect with
                                        X = rect.X + 4.0
                                        Width = max 0.0 (rect.Width - 8.0) }
                                textItems.Add(DrawText {
                                    Text = singleLineText displayText
                                    Rect = textRect
                                    Color = "#000000"
                                    FontSize = fontPt
                                    FontName = defaultFormat.FontName
                                    Bold = defaultFormat.Bold
                                    Italic = defaultFormat.Italic
                                    Underline = defaultFormat.Underline
                                    UnderlineDouble = defaultFormat.UnderlineDouble
                                    Strikethrough = defaultFormat.Strikethrough
                                    Align = alignFromFormat defaultFormat value
                                    VerticalAlign = VAlignBottom
                                    LineIndex = 0
                                    LineCount = 1
                                    BaselineShift = 0.0
                                    Rotation = defaultFormat.TextRotation
                                })
                | None -> ()
                colIndex <- colIndex + 1
            rowIndex <- rowIndex + 1

        if not compatible then
            None
        else
            let gridItems = ResizeArray<PaintItem>()
            let firstCol, minX, _ = visibleCols.[0]
            let _, lastX, lastWidth = visibleCols.[visibleCols.Count - 1]
            let firstRow, minY, _ = visibleRows.[0]
            let _, lastY, lastHeight = visibleRows.[visibleRows.Count - 1]
            let maxX = lastX + lastWidth
            let maxY = lastY + lastHeight
            let g = "#a6a6a6"

            for row, y, height in visibleRows do
                if row = firstRow then
                    gridItems.Add(GridLine (minX, y, maxX, y, g, 1.0))
                gridItems.Add(GridLine (minX, y + height, maxX, y + height, g, 1.0))

            for col, x, width in visibleCols do
                if col = firstCol then
                    gridItems.Add(GridLine (x, minY, x, maxY, g, 1.0))
                gridItems.Add(GridLine (x + width, minY, x + width, maxY, g, 1.0))

            let gridItems = resizeArrayToList gridItems
            let textItems = resizeArrayToList textItems
            let prefixItems = sparseDefaultPrefixItems model viewport
            let textCharCount = sparsePaintItemsTextCharCount prefixItems + sparsePaintItemsTextCharCount textItems
            Some {
                Cells = model.Cells
                Tables = model.Tables
                ConditionalFormats = model.ConditionalFormats
                MergedCells = model.MergedCells
                RichRuns = model.RichRuns
                RowHeights = model.RowHeights
                ColumnWidths = model.ColumnWidths
                Filters = model.Filters
                FilterHiddenRows = model.FilterHiddenRows
                Spill = model.Spill
                Validations = model.Validations
                SearchResults = model.SearchResults
                CfBoldCells = model.CfBoldCells
                PageSetup = model.PageSetup
                RowCount = model.RowCount
                ColCount = model.ColCount
                Geometry = geometry
                Measurer = viewport.Measurer
                PrefixItems = prefixItems
                VisibleRows = visibleRows.ToArray()
                VisibleCols = visibleCols.ToArray()
                GridItems = gridItems
                TextItems = textItems
                ItemCount = prefixItems.Length + gridItems.Length + textItems.Length
                TextCharCount = textCharCount
                SourceCellCount = model.Cells.Count
                CachedPartsSelection = SingleCell (-1, -1)
                CachedPartsPrefixItems = []
                CachedPartsSuffixItems = []
                CachedParts = None
            }

let private tryGetSparseDefaultPaintCacheEntry (model: SpreadsheetModel) (viewport: Viewport) =
    if Map.isEmpty model.Cells
       || not (List.isEmpty model.MergedCells)
       || model.FrozenRows <> 0
       || model.FrozenCols <> 0
       || not model.ShowGridLines
       || not (Map.isEmpty model.RichRuns) then
        None
    else
        let visible = visibleRange model viewport
        let tableStyles = buildTableStyleLookup model visible
        if not tableStyles.IsEmpty then
            None
        else
            let geometry = sparseDefaultPaintGeometryKey viewport
            let key = sparseDefaultPaintCacheKey model geometry viewport.Measurer
            let entry =
                match tryGetSparseDefaultPaintCachedEntry key with
                | ValueSome cached -> Some cached
                | ValueNone ->
                    match buildSparseDefaultPaintCacheEntry model viewport visible geometry with
                    | Some entry ->
                        rememberSparseDefaultPaintCacheEntry entry
                        Some entry
                    | None -> None
            entry

let private sparseDefaultPartsFromEntry prefixItems suffixItems (model: SpreadsheetModel) (entry: SparseDefaultPaintCacheEntry) =
    match entry.CachedParts with
    | Some parts
        when entry.CachedPartsSelection = model.Selection
             && obj.ReferenceEquals(entry.CachedPartsPrefixItems, prefixItems)
             && obj.ReferenceEquals(entry.CachedPartsSuffixItems, suffixItems) ->
        parts
    | _ ->
        let selectionFills, selectionStrokes =
            sparseDefaultSelectionParts entry.VisibleRows entry.VisibleCols model.Selection
        let parts = {
            PrefixItems = prefixItems
            SelectionFills = selectionFills
            GridItems = entry.GridItems
            SelectionStrokes = selectionStrokes
            TextItems = entry.TextItems
            SuffixItems = suffixItems
        }
        entry.CachedPartsSelection <- model.Selection
        entry.CachedPartsPrefixItems <- prefixItems
        entry.CachedPartsSuffixItems <- suffixItems
        entry.CachedParts <- Some parts
        parts

let private tryBuildSparseDefaultCellPaintParts (model: SpreadsheetModel) (viewport: Viewport) =
    tryGetSparseDefaultPaintCacheEntry model viewport
    |> Option.map (sparseDefaultPartsFromEntry [] [] model)

/// Multi-pass cell rendering. Drawing fill+stroke per cell in a single loop
/// lets the next cell's fill overwrite the previous cell's right/bottom edge
/// — invisible against white but obvious when a cell has a non-default fill
/// (Table styles, fills baked into cellXf). Splitting into three passes
/// (fills → strokes/borders → text) keeps all borders on top of all fills.
///
/// Items are tagged with the ViewportZone of the cell they came from so the
/// interactive renderer can scroll the body without sliding frozen
/// rows/cols. The static (PNG/PDF) renderer ignores the tag — at
/// ScrollX/Y=0 every zone paints at its absolute position anyway, so the
/// only visible effect of frozen panes there is the divider lines emitted
/// below.
let private iterCellPaintItems (model: SpreadsheetModel) (viewport: Viewport) (emit: ViewportZone -> PaintItem -> unit) =
    let visible = visibleRange model viewport
    let tableStyles = buildTableStyleLookup model visible

    let tryEmitEmptyDefaultCells () =
        if Map.isEmpty model.Cells
           && List.isEmpty model.MergedCells
           && tableStyles.IsEmpty
           && model.FrozenRows = 0
           && model.FrozenCols = 0
           && model.ShowGridLines then
            let selectedRange =
                match model.Selection with
                | SingleCell (row, col) -> row, col, row, col
                | Range (sr, sc, er, ec) -> sr, sc, er, ec
            let sr, sc, er, ec = selectedRange
            let fillStartRow = max visible.StartRow sr
            let fillEndRow = min (visible.EndRowExclusive - 1) er
            let fillStartCol = max visible.StartCol sc
            let fillEndCol = min (visible.EndColExclusive - 1) ec
            if fillStartRow <= fillEndRow && fillStartCol <= fillEndCol then
                for row in fillStartRow .. fillEndRow do
                    let y = rowY model viewport row
                    let height = getRowHeight row model
                    if height > 0.0 then
                        for col in fillStartCol .. fillEndCol do
                            let width = getColumnWidth col model
                            if width > 0.0 then
                                emit BodyZone (FillRect ({
                                    X = columnX model viewport col
                                    Y = y
                                    Width = width
                                    Height = height
                                }, "#E8F5E9"))

            let xBounds = ResizeArray<float>()
            let mutable minX = 0.0
            let mutable maxX = 0.0
            let mutable foundCol = false
            let mutable x = columnX model viewport visible.StartCol
            for col in visible.StartCol .. visible.EndColExclusive - 1 do
                let width = getColumnWidth col model
                if width > 0.0 then
                    if not foundCol then
                        foundCol <- true
                        minX <- x
                        xBounds.Add x
                    maxX <- x + width
                    xBounds.Add maxX
                x <- x + width

            let yBounds = ResizeArray<float>()
            let mutable minY = 0.0
            let mutable maxY = 0.0
            let mutable foundRow = false
            let mutable y = rowY model viewport visible.StartRow
            for row in visible.StartRow .. visible.EndRowExclusive - 1 do
                let height = getRowHeight row model
                if height > 0.0 then
                    if not foundRow then
                        foundRow <- true
                        minY <- y
                        yBounds.Add y
                    maxY <- y + height
                    yBounds.Add maxY
                y <- y + height

            if foundCol && foundRow then
                let g = "#a6a6a6"
                for y in yBounds do
                    emit BodyZone (GridLine (minX, y, maxX, y, g, 1.0))
                for x in xBounds do
                    emit BodyZone (GridLine (x, minY, x, maxY, g, 1.0))
            true
        else
            false

    let tryEmitSparseDefaultCells () =
        match tryBuildSparseDefaultCellPaintParts model viewport with
        | None -> false
        | Some parts ->
            emitPaintItems emit parts.SelectionFills
            emitPaintItems emit parts.GridItems
            emitPaintItems emit parts.SelectionStrokes
            emitPaintItems emit parts.TextItems
            true

    let cellState (cell: _) =
        let addr = cell.Address
        let cellData = getCell addr model
        let value = cellData |> Option.map (fun c -> c.Value) |> Option.defaultValue CellValue.Empty
        let baseFormat = cellData |> Option.map (fun c -> c.Format) |> Option.defaultValue defaultFormat
        let format =
            if tableStyles.IsEmpty then baseFormat
            else tableStyleFormat addr.Row addr.Col tableStyles baseFormat
        let selected = isSelected model.Selection addr.Row addr.Col
        let background =
            if selected then "#E8F5E9"
            else format.BgColor |> Option.defaultValue "#ffffff"
        { Cell = cell
          Format = format
          Value = value
          Selected = selected
          Background = background
          Zone = zoneForCell model addr.Row addr.Col }

    if tryEmitEmptyDefaultCells () || tryEmitSparseDefaultCells () then () else

    let estimatedVisibleCells =
        max 0 (visible.EndRowExclusive - visible.StartRow)
        * max 0 (visible.EndColExclusive - visible.StartCol)
    let states = ResizeArray<CellPaintState>(estimatedVisibleCells)
    iterVisibleCells model viewport (fun cell -> states.Add(cellState cell))
    let useSimpleBorderPass =
        if model.ShowGridLines && (model.FrozenRows > 0 || model.FrozenCols > 0) then
            false
        else
            let mutable simple = true
            let mutable index = 0
            while simple && index < states.Count do
                let state = states.[index]
                if state.Cell.Merge.IsSome || hasExplicitBorder state.Format then
                    simple <- false
                index <- index + 1
            simple

    let emitSharedGridlines () =
        if states.Count > 0 then
            let firstCell = states.[0].Cell
            let firstRow = firstCell.Address.Row
            let firstCol = firstCell.Address.Col
            let xBounds = ResizeArray<float>()
            let yBounds = ResizeArray<float>()
            let mutable minX = firstCell.Rect.X
            let mutable maxX = firstCell.Rect.X + firstCell.Rect.Width
            let mutable minY = firstCell.Rect.Y
            let mutable maxY = firstCell.Rect.Y + firstCell.Rect.Height
            for state in states do
                let cell = state.Cell
                let rect = cell.Rect
                let row = cell.Address.Row
                let col = cell.Address.Col
                minX <- min minX rect.X
                maxX <- max maxX (rect.X + rect.Width)
                minY <- min minY rect.Y
                maxY <- max maxY (rect.Y + rect.Height)
                if row = firstRow then
                    if xBounds.Count = 0 then
                        xBounds.Add rect.X
                    xBounds.Add(rect.X + rect.Width)
                if col = firstCol then
                    if yBounds.Count = 0 then
                        yBounds.Add rect.Y
                    yBounds.Add(rect.Y + rect.Height)
            let g = "#a6a6a6"
            for y in yBounds do
                emit BodyZone (GridLine (minX, y, maxX, y, g, 1.0))
            for x in xBounds do
                emit BodyZone (GridLine (x, minY, x, maxY, g, 1.0))

    // Pass 1: all backgrounds. Adjacent cells share an edge, so when a
    // cell with a non-default fill is followed by another cell, the
    // second cell's fill would overdraw the first cell's right/bottom
    // stroke if we drew them in the same per-cell loop.
    for state in states do
        if not (isWhiteFill state.Background) then
            emit state.Zone (FillRect (state.Cell.Rect, state.Background))

    // Pass 1b: conditional-formatting data bars. Drawn on top of the
    // cell background but under borders / text — a horizontal bar
    // from the cell's left edge spanning fraction × width. Excel
    // insets the bar ~2 px vertically; mirror that so the gridline
    // stays visible above and below.
    for state in states do
        let cell = state.Cell
        match state.Format.DataBar with
        | Some (frac, color) when frac > 0.0 ->
            let inset = 2.0
            let barW = (cell.Rect.Width - 1.0) * frac
            if barW > 0.5 then
                emit state.Zone (FillRect (
                    { X = cell.Rect.X
                      Y = cell.Rect.Y + inset
                      Width = barW
                      Height = max 1.0 (cell.Rect.Height - inset * 2.0) },
                    color))
        | _ -> ()

    // Pass 2: per-cell strokes + cellXf borders, drawn on top of every
    // fill. Match Excel's *interactive* view: a light gray gridline
    // around every cell — not the near-black grid that print mode emits
    // with PrintGridlines on. Our reference PNGs come from print, but the
    // renderer's target is the on-screen view.
    // Emit a styled border side as one (single solid) or two (double)
    // GridLine items. "double" in Excel is two ~0.75 px parallel rules
    // straddling the cell edge with a small gap; emit two 1 px strokes
    // offset by ±1 px from the edge so the gap is visible at 96 dpi.
    // Other styles flow through the existing single-stroke path with
    // the borderStrokeWidth weight.
    let emitBorderSide (isHorizontal: bool) (side: BorderStyle) (x1: float, y1: float) (x2: float, y2: float) =
        let color = borderColor side
        let w = borderStrokeWidth side.Style
        let style = side.Style
        // Excel's dashed / dotted / dash-dot variants render as a
        // pattern of on/off chunks. We approximate by chunking the
        // line into short solid segments separated by gaps — the
        // PaintItem.GridLine plumbing carries only width, not dash
        // info, and emitting one segment per dash keeps every
        // downstream renderer (PNG / PDF / browser) consistent.
        let dashChunks onLen offLen =
            let segs = ResizeArray<PaintItem>()
            if isHorizontal then
                let mutable x = x1
                while x < x2 do
                    let segEnd = min x2 (x + onLen)
                    segs.Add(GridLine (x, y1, segEnd, y2, color, w))
                    x <- segEnd + offLen
            else
                let mutable y = y1
                while y < y2 do
                    let segEnd = min y2 (y + onLen)
                    segs.Add(GridLine (x1, y, x2, segEnd, color, w))
                    y <- segEnd + offLen
            segs
        if equalsIgnoreCase style "double" then
            // Excel renders double borders straddling the cell edge
            // (one stroke just outside, one just inside). When the
            // adjacent cell has a wider border (e.g. `thick`, 3 px),
            // both straddling strokes get overpainted by the
            // neighbour's stroke. The caller compensates by passing
            // a `doubleInset` adjusted pair of y / x values — see
            // the per-cell emission loop for the direction-aware
            // offsets.
            let off = 1.0
            let segs = ResizeArray<PaintItem>(2)
            if isHorizontal then
                segs.Add(GridLine (x1, y1 - off, x2, y2 - off, color, 1.0))
                segs.Add(GridLine (x1, y1 + off, x2, y2 + off, color, 1.0))
            else
                segs.Add(GridLine (x1 - off, y1, x2 - off, y2, color, 1.0))
                segs.Add(GridLine (x1 + off, y1, x2 + off, y2, color, 1.0))
            segs
        elif equalsIgnoreCase style "dotted" || equalsIgnoreCase style "mediumdotted" then
            dashChunks 1.0 1.5
        elif equalsIgnoreCase style "dashed" || equalsIgnoreCase style "mediumdashed" then
            dashChunks 3.0 2.5
        elif equalsIgnoreCase style "dashdot" || equalsIgnoreCase style "mediumdashdot" then
            // 3-on, 1.5-off, 1-on, 1.5-off pattern: emit two sub-passes
            let mainDash = dashChunks 3.0 4.0  // 3 + 1 + ~ offset for dot
            if isHorizontal then
                let mutable x = x1 + 4.0
                while x < x2 do
                    let segEnd = min x2 (x + 1.0)
                    mainDash.Add(GridLine (x, y1, segEnd, y2, color, w))
                    x <- x + 8.0
            else
                let mutable y = y1 + 4.0
                while y < y2 do
                    let segEnd = min y2 (y + 1.0)
                    mainDash.Add(GridLine (x1, y, x2, segEnd, color, w))
                    y <- y + 8.0
            mainDash
        elif equalsIgnoreCase style "dashdotdot" || equalsIgnoreCase style "mediumdashdotdot" then
            dashChunks 3.0 5.0  // simplification; main dash with bigger gap
        elif equalsIgnoreCase style "slantdashdot" then
            dashChunks 3.0 3.0
        elif equalsIgnoreCase style "hair" then
            // Hair = very thin dotted. Simulate with tiny dots.
            dashChunks 0.5 1.0
        else
            let segs = ResizeArray<PaintItem>(1)
            segs.Add(GridLine (x1, y1, x2, y2, color, w))
            segs

    // Border collapse: when two adjacent cells both declare a border
    // on their shared edge, Excel draws the heavier one and skips the
    // lighter. Same-weight ties resolve in favour of the cell that
    // "owns" the edge first in the natural emission order — for
    // vertical edges that's the upper cell's bottom; for horizontal
    // edges that's the left cell's right. Without this collapse,
    // double + adjacent thin produced a doubled-line artifact and
    // dashed borders got overpainted by the neighbour's solid edge.
    // Gridlines (the always-on light cell outline) FIRST, as their own
    // sub-pass, so explicit borders always draw on top. Excel hides the
    // light gridline on any edge that carries an explicit border, so we
    // draw the grey edge only when neither this cell nor its neighbour has
    // a border there. Without this the grey gridline bleeds through a
    // double border's gap (and gets half-covered by neighbouring grid
    // strokes). Selected cells keep a full highlight box.
    if useSimpleBorderPass then
        if model.ShowGridLines then
            emitSharedGridlines()
        for state in states do
            if state.Selected then
                emit state.Zone (StrokeRect (state.Cell.Rect, "#217346", 2.0))
    else
        let borderByAddr = Dictionary<int64, CellFormat>()
        for state in states do
            borderByAddr.[cellKey state.Cell.Address.Row state.Cell.Address.Col] <- state.Format
        let neighborSide (row: int) (col: int) (side: CellFormat -> BorderStyle option) : BorderStyle option =
            match borderByAddr.TryGetValue(cellKey row col) with
            | true, format -> side format
            | false, _ -> None
        let hasMergedCells =
            let mutable found = false
            let mutable index = 0
            while not found && index < states.Count do
                if states.[index].Cell.Merge.IsSome then
                    found <- true
                index <- index + 1
            found
        let lastVisibleRow, lastVisibleCol =
            let mutable lastRow = System.Int32.MinValue
            let mutable lastCol = System.Int32.MinValue
            for state in states do
                lastRow <- max lastRow state.Cell.Address.Row
                lastCol <- max lastCol state.Cell.Address.Col
            lastRow, lastCol
        // Use OOXML border priority (not raw stroke width) for collapse:
        // `double` outranks `thick` in Excel even though it paints thinner.
        // Order: None < Hair < Dotted < Dashed-family < Thin < Medium <
        // Thick < Double.
        let weightOf (side: BorderStyle option) : int =
            match side with
            | None -> 0
            | Some s -> borderPriority s.Style

        for state in states do
            let cell = state.Cell
            let format = state.Format
            let rect = cell.Rect
            let r, c = cell.Address.Row, cell.Address.Col
            if state.Selected then
                emit state.Zone (StrokeRect (rect, "#217346", 2.0))
            // Suppress the light grey screen gridlines when the sheet has
            // `showGridLines="0"` (polished workbooks rely on explicit borders).
            elif model.ShowGridLines then
                let bordered (mine: BorderStyle option) nr nc nside =
                    mine.IsSome || (neighborSide nr nc nside).IsSome
                let g = "#a6a6a6"
                if hasMergedCells then
                    if not (bordered format.BorderTop (r - 1) c (fun f -> f.BorderBottom)) then
                        emit state.Zone (GridLine (rect.X, rect.Y, rect.X + rect.Width, rect.Y, g, 1.0))
                    if not (bordered format.BorderBottom (r + 1) c (fun f -> f.BorderTop)) then
                        emit state.Zone (GridLine (rect.X, rect.Y + rect.Height, rect.X + rect.Width, rect.Y + rect.Height, g, 1.0))
                    if not (bordered format.BorderLeft r (c - 1) (fun f -> f.BorderRight)) then
                        emit state.Zone (GridLine (rect.X, rect.Y, rect.X, rect.Y + rect.Height, g, 1.0))
                    if not (bordered format.BorderRight r (c + 1) (fun f -> f.BorderLeft)) then
                        emit state.Zone (GridLine (rect.X + rect.Width, rect.Y, rect.X + rect.Width, rect.Y + rect.Height, g, 1.0))
                else
                    if not (bordered format.BorderTop (r - 1) c (fun f -> f.BorderBottom)) then
                        emit state.Zone (GridLine (rect.X, rect.Y, rect.X + rect.Width, rect.Y, g, 1.0))
                    if r = lastVisibleRow && not (bordered format.BorderBottom (r + 1) c (fun f -> f.BorderTop)) then
                        emit state.Zone (GridLine (rect.X, rect.Y + rect.Height, rect.X + rect.Width, rect.Y + rect.Height, g, 1.0))
                    if not (bordered format.BorderLeft r (c - 1) (fun f -> f.BorderRight)) then
                        emit state.Zone (GridLine (rect.X, rect.Y, rect.X, rect.Y + rect.Height, g, 1.0))
                    if c = lastVisibleCol && not (bordered format.BorderRight r (c + 1) (fun f -> f.BorderLeft)) then
                        emit state.Zone (GridLine (rect.X + rect.Width, rect.Y, rect.X + rect.Width, rect.Y + rect.Height, g, 1.0))
        for state in states do
            let cell = state.Cell
            let format = state.Format
            let rect = cell.Rect
            let r, c = cell.Address.Row, cell.Address.Col
            // Merged cells: Excel splits a box border across the merge's
            // constituent cells (top-left carries left+top, top-right the
            // right edge, bottom-left/-right the bottom edge, etc.).
            // visibleCells yields only the merge primary — with the full
            // merge Rect — and skips the trailing cells, so we gather each
            // edge's style from the merge perimeter directly off the model
            // rather than from the primary's own format. For a non-merged
            // cell every edge list is the single cell, so this reduces to
            // the previous `format.BorderX` behavior exactly.
            let mr0, mc0, mr1, mc1 =
                match cell.Merge with
                | Some m -> m.StartRow, m.StartCol, m.EndRow, m.EndCol
                | None -> r, c, r, c
            let fmtAt row col =
                getCell { Row = row; Col = col } model
                |> Option.map (fun cd -> cd.Format)
                |> Option.defaultValue defaultFormat
            // Heaviest border style found along an edge of the merge.
            let betterSide (best: BorderStyle option) (candidate: BorderStyle option) : BorderStyle option =
                match best, candidate with
                | None, Some side -> Some side
                | Some current, Some side when borderPriority side.Style > borderPriority current.Style -> Some side
                | _ -> best
            let horizontalEdge row c0 c1 (pick: CellFormat -> BorderStyle option) =
                let mutable best: BorderStyle option = None
                for cc in c0 .. c1 do
                    best <- betterSide best (pick (fmtAt row cc))
                best
            let verticalEdge r0 r1 col (pick: CellFormat -> BorderStyle option) =
                let mutable best: BorderStyle option = None
                for rr in r0 .. r1 do
                    best <- betterSide best (pick (fmtAt rr col))
                best
            let bTop, bBottom, bLeft, bRight =
                match cell.Merge with
                | None ->
                    format.BorderTop, format.BorderBottom, format.BorderLeft, format.BorderRight
                | Some _ ->
                    horizontalEdge mr0 mc0 mc1 (fun f -> f.BorderTop),
                    horizontalEdge mr1 mc0 mc1 (fun f -> f.BorderBottom),
                    verticalEdge mr0 mr1 mc0 (fun f -> f.BorderLeft),
                    verticalEdge mr0 mr1 mc1 (fun f -> f.BorderRight)
            // Top border: skip if the cell above already drew a heavier
            // (or equal-weight) bottom edge. The above-bottom owns ties.
            // Higher priority wins (double > thick > medium > thin > dashed).
            match bTop with
            | Some side ->
                let aboveBottom = neighborSide (mr0 - 1) mc0 (fun f -> f.BorderBottom)
                let mine = weightOf (Some side)
                let neighbor = weightOf aboveBottom
                if mine > neighbor then
                    for item in emitBorderSide true side (rect.X, rect.Y) (rect.X + rect.Width, rect.Y) do
                        emit state.Zone item
            | None -> ()
            // Bottom border: draws unless the cell below has a strictly
            // heavier top border.
            match bBottom with
            | Some side ->
                let belowTop = neighborSide (mr1 + 1) mc0 (fun f -> f.BorderTop)
                let mine = weightOf (Some side)
                let neighbor = weightOf belowTop
                if mine >= neighbor then
                    let y = rect.Y + rect.Height
                    for item in emitBorderSide true side (rect.X, y) (rect.X + rect.Width, y) do
                        emit state.Zone item
            | None -> ()
            // Left border: skip if the cell to the left already drew a
            // heavier-or-equal right edge.
            match bLeft with
            | Some side ->
                let leftRight = neighborSide mr0 (mc0 - 1) (fun f -> f.BorderRight)
                let mine = weightOf (Some side)
                let neighbor = weightOf leftRight
                if mine > neighbor then
                    for item in emitBorderSide false side (rect.X, rect.Y) (rect.X, rect.Y + rect.Height) do
                        emit state.Zone item
            | None -> ()
            // Right border: draws unless the cell to the right has a
            // strictly heavier left border.
            match bRight with
            | Some side ->
                let rightLeft = neighborSide mr0 (mc1 + 1) (fun f -> f.BorderLeft)
                let mine = weightOf (Some side)
                let neighbor = weightOf rightLeft
                if mine >= neighbor then
                    let x = rect.X + rect.Width
                    for item in emitBorderSide false side (x, rect.Y) (x, rect.Y + rect.Height) do
                        emit state.Zone item
            | None -> ()
            // Diagonal strokes: down = top-left → bottom-right, up =
            // bottom-left → top-right. Cells set one or both via the
            // `diagonalUp` / `diagonalDown` border attributes. Style +
            // color come from the shared `<diagonal>` child, so both
            // strokes share the same look when both flags are set.
            match format.BorderDiagonal with
            | Some side ->
                let color = borderColor side
                let w = borderStrokeWidth side.Style
                if format.DiagonalDown then
                    emit state.Zone (GridLine (rect.X, rect.Y, rect.X + rect.Width, rect.Y + rect.Height, color, w))
                if format.DiagonalUp then
                    emit state.Zone (GridLine (rect.X, rect.Y + rect.Height, rect.X + rect.Width, rect.Y, color, w))
            | None -> ()

    // Pass 3: text on top of all strokes/borders.
    for state in states do
        let cell = state.Cell
        let format = state.Format
        let value = state.Value
        let zone = state.Zone
        let hasDisplayValue =
            match value with
            | CellValue.Empty -> false
            | _ -> true
        let fontPt =
            if hasDisplayValue then format.FontSize |> Option.map float |> Option.defaultValue 11.0
            else 11.0
        // Build the real-font measurer first so cellDisplayText can size
        // its ##### overflow run against the glyph the renderer will
        // actually draw, not the mdwPx approximation.
        let measure =
            if hasDisplayValue then
                viewport.Measurer
                |> Option.map (fun m ->
                    fun (s: string) -> m s fontPt format.FontName format.Bold format.Italic)
            else
                None
        let displayText =
            if hasDisplayValue then cellDisplayText measure cell.Rect.Width format value
            else ""
        // Cells in a hidden column (0 width) or hidden row (0 height) paint
        // nothing — Excel renders no text there, and without this guard the
        // hidden cell's text overflows into the next visible column (and a
        // 0-width cell wraps its text to one glyph per line).
        if displayText <> "" && cell.Rect.Width > 0.0 && cell.Rect.Height > 0.0 then
            // Indent past the conditional-formatting icon when present
            // so the icon and the cell text don't overlap (Excel does
            // the same — the icon occupies a ~16 px gutter).
            let iconGutter = if format.IconSet.IsSome then 16.0 else 0.0
            // Excel's `<alignment indent="N"/>` adds N units of leading
            // padding. Measured against Excel (synthetic-cell-indent 11pt
            // and the jumeirah model 12pt), one level ≈ 10 px at 96 dpi —
            // roughly font-independent and ≈ 1.4× the max-digit width, NOT
            // the 3× ("3 character widths") we used before, which indented
            // ~2× too far. Applies from the LEFT for left/distributed, from
            // the RIGHT for right align; ignored for center.
            let indentPx =
                if format.Indent <= 0 then 0.0
                else
                    let unit = mdwPx fontPt * 1.4
                    float format.Indent * unit
            let leftIndent =
                if indentPx > 0.0
                   && (optionEqualsIgnoreCase "left" format.HorizontalAlign
                       || optionEqualsIgnoreCase "distributed" format.HorizontalAlign) then indentPx
                else 0.0
            let rightIndent =
                if indentPx > 0.0 && optionEqualsIgnoreCase "right" format.HorizontalAlign then indentPx
                else 0.0
            // "Center Across Selection" (`horizontal="centerContinuous"`):
            // centre the text across the contiguous run of centerContinuous
            // cells to the right that carry no content of their own — Excel
            // does this without merging. Expand the base rect over the run.
            let baseRect =
                if optionEqualsIgnoreCase "centercontinuous" format.HorizontalAlign then
                    let row = cell.Address.Row
                    let mutable lastCol = cell.Address.Col
                    let mutable c2 = cell.Address.Col + 1
                    let mutable go = true
                    while go && c2 < model.ColCount do
                        let naddr = { Row = row; Col = c2 }
                        let nfmt = getCell naddr model |> Option.map (fun c -> c.Format) |> Option.defaultValue defaultFormat
                        let nEmpty = match getCellValue naddr model with CellValue.Empty -> true | _ -> false
                        if optionEqualsIgnoreCase "centercontinuous" nfmt.HorizontalAlign && nEmpty then lastCol <- c2; c2 <- c2 + 1
                        else go <- false
                    if lastCol > cell.Address.Col then
                        let r = cellRect model viewport row lastCol
                        { cell.Rect with Width = (r.X + r.Width) - cell.Rect.X }
                    else cell.Rect
                else cell.Rect
            let textRect =
                { baseRect with
                    X = baseRect.X + 4.0 + iconGutter + leftIndent
                    Width = max 0.0 (baseRect.Width - 8.0 - iconGutter - leftIndent - rightIndent) }
            // ShrinkToFit (single-line text only — wrap-mode escapes via
            // line breaks). When the text overflows the cell width and
            // shrinkToFit is set, scale the font size down until it fits,
            // clamped at 1 pt. We need to re-measure with the smaller
            // size so the rest of the rendering uses the actual width.
            let fontPt, measure =
                if not format.ShrinkToFit || format.WrapText then fontPt, measure
                else
                    match measure with
                    | None -> fontPt, measure
                    | Some m ->
                        let width = m displayText
                        if width <= textRect.Width || textRect.Width <= 0.0 then fontPt, measure
                        else
                            let ratio = max 0.1 (textRect.Width / width)
                            let shrunk = max 1.0 (fontPt * ratio)
                            let shrunkMeasure =
                                viewport.Measurer
                                |> Option.map (fun mm ->
                                    fun (s: string) -> mm s shrunk format.FontName format.Bold format.Italic)
                            shrunk, shrunkMeasure
            // Excel number-format section colour overrides the cell's
            // font colour: `#,##0.00;[Red]-#,##0.00` paints negatives red
            // even when the cellXf has no explicit fontColor.
            let sectionColor =
                match format.NumberFormat, value with
                | Some code, Number n -> numberFormatColorForValue code n
                | Some code, Formula (_, Number n) -> numberFormatColorForValue code n
                | _ -> None
            let color =
                sectionColor
                |> Option.orElse format.TextColor
                |> Option.defaultValue "#000000"
            let align = alignFromFormat format value
            let vAlign =
                match format.VerticalAlign with
                | Some align when equalsIgnoreCase align "top" -> VAlignTop
                | Some align when equalsIgnoreCase align "center" || equalsIgnoreCase align "middle" -> VAlignMiddle
                | _ -> VAlignBottom
            let richRuns = Map.tryFind cell.Address model.RichRuns
            match format.WrapText, richRuns, viewport.Measurer with
            | _, Some runs, Some m when not (List.isEmpty runs) ->
                let lines =
                    if format.WrapText then
                        wrapLines measure textRect.Width fontPt displayText
                    else
                        [ singleLineText displayText ]
                let lineCount = lines.Length
                // Rich-text rendering. For single-line cells we emit one
                // DrawText per run. For wrapped/multi-line cells we walk
                // a position cursor through the concatenated source text
                // and chunk runs that straddle line boundaries — each
                // line gets per-run DrawTexts at its own y, with the
                // breaking whitespace dropped by wrapLines absorbed into
                // the cursor (so a run that ends with " " before a line
                // break still aligns with the next line's start).
                let resolveRun (ovr: RunOverride) =
                    let bold = ovr.Bold |> Option.defaultValue format.Bold
                    let italic = ovr.Italic |> Option.defaultValue format.Italic
                    let underline = ovr.Underline |> Option.defaultValue format.Underline
                    let underlineDouble = ovr.UnderlineDouble |> Option.defaultValue format.UnderlineDouble
                    let strike = ovr.Strikethrough |> Option.defaultValue format.Strikethrough
                    let fName = ovr.FontName |> Option.orElse format.FontName
                    let baseSize = ovr.FontSize |> Option.defaultValue fontPt
                    let runColor = ovr.Color |> Option.defaultValue color
                    // vertAlign: shrink to ~60% of the cell font (Excel
                    // uses ~58% empirically, close enough) and shift
                    // the baseline. Negative shift = up (super); we
                    // express shift in pt at the cell's base size so
                    // the renderer's pt→px scaling stays consistent.
                    let fSize, baselineShift =
                        match ovr.VertAlign with
                        | Some VertAlignSuperscript -> baseSize * 0.6, -fontPt * 0.35
                        | Some VertAlignSubscript -> baseSize * 0.6, fontPt * 0.15
                        | _ -> baseSize, 0.0
                    bold, italic, underline, underlineDouble, strike, fName, fSize, runColor, baselineShift
                let displayConcat =
                    if lineCount <= 1 then ""
                    else
                        let displayBuilder = System.Text.StringBuilder(displayText.Length)
                        for runText, _ in runs do
                            displayBuilder.Append(runText) |> ignore
                        displayBuilder.ToString()
                let mutable runOffset = 0
                let mutable remainingRuns = runs
                let mutable currentRun =
                    match remainingRuns with
                    | head :: tail ->
                        remainingRuns <- tail
                        Some head
                    | [] -> None
                let advanceRun () =
                    match remainingRuns with
                    | head :: tail ->
                        currentRun <- Some head
                        remainingRuns <- tail
                        runOffset <- 0
                    | [] ->
                        currentRun <- None
                        runOffset <- 0
                let mutable docPos = 0
                let mutable lineIndex = 0
                for line in lines do
                    let metrics = ResizeArray<string * RunOverride * bool * bool * bool * bool * bool * string option * float * string * float * float>()
                    let mutable totalWidth = 0.0
                    let mutable lineRem = line.Length
                    while lineRem > 0 && Option.isSome currentRun do
                        match currentRun with
                        | Some (runText, runProps) ->
                            let available = runText.Length - runOffset
                            let take = min lineRem available
                            if take > 0 then
                                let textSlice =
                                    if runOffset = 0 && take = runText.Length then runText
                                    else runText.Substring(runOffset, take)
                                let bold, italic, underline, underlineDouble, strike, fName, fSize, runColor, baselineShift = resolveRun runProps
                                let width = m textSlice fSize fName bold italic
                                totalWidth <- totalWidth + width
                                metrics.Add(textSlice, runProps, bold, italic, underline, underlineDouble, strike, fName, fSize, runColor, width, baselineShift)
                            lineRem <- lineRem - take
                            runOffset <- runOffset + take
                            if runOffset >= runText.Length then
                                advanceRun ()
                        | None -> ()
                    docPos <- docPos + line.Length
                    // Skip the breaking space wrapLines drops between
                    // non-final lines. Also consume the matching space
                    // from the run queue head so positions stay aligned.
                    if lineIndex < lineCount - 1
                       && docPos < displayConcat.Length
                       && displayConcat.[docPos] = ' ' then
                        docPos <- docPos + 1
                        match currentRun with
                        | Some (rt, _) ->
                            if runOffset < rt.Length && rt.[runOffset] = ' ' then
                                runOffset <- runOffset + 1
                                if runOffset >= rt.Length then
                                    advanceRun ()
                        | None -> ()
                    // Emit the measured line segments.
                    let baseLineX =
                        match align with
                        | AlignLeft -> textRect.X
                        | AlignCenter -> textRect.X + (textRect.Width - totalWidth) / 2.0
                        | AlignRight -> textRect.X + textRect.Width - totalWidth
                    let mutable cursor = baseLineX
                    for (t, _ovr, bold, italic, underline, underlineDouble, strike, fName, fSize, runColor, width, baselineShift) in metrics do
                        emit zone (DrawText {
                            Text = t
                            Rect = { textRect with X = cursor; Width = width }
                            Color = runColor
                            FontSize = fSize
                            FontName = fName
                            Bold = bold
                            Italic = italic
                            Underline = underline
                            UnderlineDouble = underlineDouble
                            Strikethrough = strike
                            Align = AlignLeft
                            VerticalAlign = vAlign
                            LineIndex = lineIndex
                            LineCount = lineCount
                            BaselineShift = baselineShift
                            Rotation = format.TextRotation
                        })
                        cursor <- cursor + width
                    lineIndex <- lineIndex + 1
            | false, _, _ ->
                // Common plain single-line path: avoid allocating a one-item
                // line list for every visible cell.
                emit zone (DrawText {
                    Text = singleLineText displayText
                    Rect = textRect
                    Color = color
                    FontSize = fontPt
                    FontName = format.FontName
                    Bold = format.Bold
                    Italic = format.Italic
                    Underline = format.Underline
                    UnderlineDouble = format.UnderlineDouble
                    Strikethrough = format.Strikethrough
                    Align = align
                    VerticalAlign = vAlign
                    LineIndex = 0
                    LineCount = 1
                    BaselineShift = 0.0
                    Rotation = format.TextRotation
                })
            | _ ->
                // Wrapping and manual `\n` line breaks only apply when wrapText
                // is ON. With wrapText OFF, Excel renders on a single line and
                // ignores any embedded break (shown as a space), so the row is
                // not multi-line.
                let lines = wrapLines measure textRect.Width fontPt displayText
                let lineCount = lines.Length
                // Plain (or wrapped) cell text — one DrawText per line.
                let mutable lineIndex = 0
                for line in lines do
                    emit zone (DrawText {
                        Text = line
                        Rect = textRect
                        Color = color
                        FontSize = fontPt
                        FontName = format.FontName
                        Bold = format.Bold
                        Italic = format.Italic
                        Underline = format.Underline
                        UnderlineDouble = format.UnderlineDouble
                        Strikethrough = format.Strikethrough
                        Align = align
                        VerticalAlign = vAlign
                        LineIndex = lineIndex
                        LineCount = lineCount
                        BaselineShift = 0.0
                        Rotation = format.TextRotation
                    })
                    lineIndex <- lineIndex + 1

    // Pass 4: auto-filter dropdown triangles on table header cells. Excel
    // renders these as a small downward triangle on the right edge of each
    // header cell whose column falls within the table's <autoFilter>
    // range. The triangle is drawn over the header's accent fill — same
    // color as the header text (white on a Medium-style accent).
    for state in states do
        let cell = state.Cell
        let format = state.Format
        if format.FilterButton then
            let color = format.TextColor |> Option.defaultValue "#FFFFFF"
            // 7-px-wide, 4-px-tall triangle centered in a 16-px right-edge
            // button area. Drawn as five horizontal stroked lines (no
            // primitive for filled polygons; this is cheap enough).
            let buttonWidth = 16.0
            let centerX = cell.Rect.X + cell.Rect.Width - buttonWidth / 2.0
            let triHalfWidth = 3.5
            let triTopY = cell.Rect.Y + cell.Rect.Height / 2.0 - 2.0
            for i in 0 .. 3 do
                let shrink = float i * (triHalfWidth / 4.0)
                let y = triTopY + float i
                emit state.Zone (GridLine (centerX - triHalfWidth + shrink, y, centerX + triHalfWidth - shrink, y, color, 1.0))

    // Pass 5: comment-indicator red triangle in the top-right corner.
    // Excel uses a 5-px right-triangle filled with #FF0000 — we
    // approximate with horizontal strokes of decreasing width since
    // the renderer has no fill-polygon primitive. The triangle's
    // hypotenuse goes from (right-5, top) down to (right, top+5).
    for state in states do
        let cell = state.Cell
        if state.Format.HasComment then
            let color = "#C00000"  // Excel's "note marker" red
            let triSize = 5.0
            let rightX = cell.Rect.X + cell.Rect.Width
            let topY = cell.Rect.Y
            // Row k from top: from x = rightX - (triSize - k) to x = rightX.
            for k in 0 .. int triSize - 1 do
                let leftX = rightX - (triSize - float k)
                emit state.Zone (GridLine (leftX, topY + float k + 0.5, rightX, topY + float k + 0.5, color, 1.0))

    // Pass 6: data-validation list dropdown arrow. Excel only shows
    // this when the cell is selected, but for static export we
    // always paint it so the visual cue carries through to PNG /
    // PDF snapshots. Drawn as a small 7-px-wide triangle just
    // inside the cell's right edge, vertically centered, in gray.
    // FilterButton already owns the dropdown styling for table
    // header cells; skip those to avoid doubled triangles.
    for state in states do
        let cell = state.Cell
        let format = state.Format
        if format.HasListDropdown && not format.FilterButton then
            let color = "#666666"
            let buttonWidth = 16.0
            let centerX = cell.Rect.X + cell.Rect.Width - buttonWidth / 2.0
            let triHalfWidth = 3.5
            let triTopY = cell.Rect.Y + cell.Rect.Height / 2.0 - 2.0
            for i in 0 .. 3 do
                let shrink = float i * (triHalfWidth / 4.0)
                let y = triTopY + float i
                emit state.Zone (GridLine (centerX - triHalfWidth + shrink, y, centerX + triHalfWidth - shrink, y, color, 1.0))

    // Pass 7: conditional-formatting icon-set indicator. The renderer
    // has no icon artwork or circle primitive, so approximate with a
    // small filled square in a low→high color (red → amber → green
    // for a 3-set; interpolated for 4/5-sets). It conveys the same
    // low/mid/high signal the real icons do. Drawn in the 16 px
    // gutter the text was already indented past.
    for state in states do
        let cell = state.Cell
        let format = state.Format
        match format.IconSet with
        | Some (idx, count) when count > 1 ->
            // 0 = worst (red), count-1 = best (green), midpoints amber.
            let t = float idx / float (count - 1)
            let color =
                if t <= 0.5 then
                    // red → amber over the first half
                    let u = t * 2.0
                    sprintf "#FF%02X00" (int (0.0 * (1.0 - u) + 192.0 * u))
                else
                    // amber → green over the second half
                    let u = (t - 0.5) * 2.0
                    sprintf "#%02X%02X%02X"
                        (int (255.0 * (1.0 - u)))
                        (int (192.0 * (1.0 - u) + 176.0 * u))
                        (int (80.0 * u))
            let sz = 9.0
            let iconX = cell.Rect.X + 3.0
            let iconY = cell.Rect.Y + cell.Rect.Height / 2.0 - sz / 2.0
            emit state.Zone (FillRect ({ X = iconX; Y = iconY; Width = sz; Height = sz }, color))
        | _ -> ()

let private buildCellPaintItems (model: SpreadsheetModel) (viewport: Viewport): (ViewportZone * PaintItem) list =
    let items = ResizeArray<ViewportZone * PaintItem>()
    iterCellPaintItems model viewport (fun zone item -> items.Add((zone, item)))
    resizeArrayToList items

/// Freeze-divider strokes drawn on top of the cell layer. Excel renders
/// this as a single ~1 px line slightly darker than the regular gridline,
/// at the split's right (cols) / bottom (rows) edge. Matches the line
/// Excel draws in both interactive and print-snapshot views. Coordinates
/// are in the same space as cellRect (post-HeaderHeight / -RowHeaderWidth
/// offset) so the line lands flush with the bottom/right of the last
/// frozen row/column.
let private iterFreezeDividerItems (model: SpreadsheetModel) (viewport: Viewport) (emit: PaintItem -> unit) =
    // Slightly darker than the regular gridline. Excel uses a fixed
    // medium-gray (~#595959 in print mode).
    let color = "#595959"
    let w = 1.0
    if model.FrozenRows > 0 then
        let y =
            viewport.HeaderHeight
            + frozenRowsHeight model
        let totalWidth =
            viewport.RowHeaderWidth
            + columnLeft model model.ColCount
        emit (GridLine (viewport.RowHeaderWidth, y, totalWidth, y, color, w))
    if model.FrozenCols > 0 then
        let x =
            viewport.RowHeaderWidth
            + frozenColsWidth model
        let totalHeight =
            viewport.HeaderHeight
            + rowTop model model.RowCount
        emit (GridLine (x, viewport.HeaderHeight, x, totalHeight, color, w))

let private buildFreezeDividerItems (model: SpreadsheetModel) (viewport: Viewport) : PaintItem list =
    let items = ResizeArray<PaintItem>()
    iterFreezeDividerItems model viewport items.Add
    resizeArrayToList items

let private filterPaintProjectionCache = ConditionalWeakTable<SpreadsheetModel, SpreadsheetModel>()

let private estimatedPaintItemCount (model: SpreadsheetModel) (viewport: Viewport) =
    let visible = visibleRange model viewport
    let visibleRows = max 0 (visible.EndRowExclusive - visible.StartRow)
    let visibleCols = max 0 (visible.EndColExclusive - visible.StartCol)
    let visibleCells =
        if Map.isEmpty model.Cells
           && List.isEmpty model.MergedCells
           && not (hasAnyRenderTimeTableStyle model.Tables)
           && model.FrozenRows = 0
           && model.FrozenCols = 0
           && model.ShowGridLines then
            let sr, sc, er, ec =
                match model.Selection with
                | SingleCell (row, col) -> row, col, row, col
                | Range (sr, sc, er, ec) -> sr, sc, er, ec
            let selectedRows = max 0 (min (visible.EndRowExclusive - 1) er - max visible.StartRow sr + 1)
            let selectedCols = max 0 (min (visible.EndColExclusive - 1) ec - max visible.StartCol sc + 1)
            selectedRows * selectedCols
        else
            visibleRows * visibleCols
    let headers = 2 + (visibleCols * 3) + (visibleRows * 3)
    let sharedGridlines = visibleRows + visibleCols + 2
    max 16 (1 + headers + visibleCells + sharedGridlines + 2)

type private StaticPaintGeometryKey = {
    ScrollX: float
    ScrollY: float
    Width: float
    Height: float
    HeaderHeight: float
    RowHeaderWidth: float
}

type private StaticPaintItemsCacheEntry = {
    Measurer: TextMeasurer option
    Items: PaintItem list
    ItemCount: int
    ItemTextCharCount: int
}

type private StaticPaintGeometryFastEntry = {
    Key: StaticPaintGeometryKey
    Entry: StaticPaintItemsCacheEntry
}

type private StaticPaintItemsCache = {
    Model: SpreadsheetModel
    SourceItemCount: int
    SourceTextCharCount: int
    Entries: Dictionary<StaticPaintGeometryKey, StaticPaintItemsCacheEntry>
    Order: Queue<StaticPaintGeometryKey>
    EntryFastSlots: StaticPaintGeometryFastEntry array
    mutable TotalItemCount: int
    mutable TotalItemTextCharCount: int
}

let mutable private staticPaintItemsCache: StaticPaintItemsCache list = []
let private spreadsheetModelReferenceComparer =
    { new IEqualityComparer<SpreadsheetModel> with
        member _.Equals(left, right) = obj.ReferenceEquals(left, right)
        member _.GetHashCode(value) = RuntimeHelpers.GetHashCode(value) }
let private staticPaintItemsCacheByModel =
    Dictionary<SpreadsheetModel, StaticPaintItemsCache>(spreadsheetModelReferenceComparer)
let private staticPaintItemsFastSlots = Array.zeroCreate<StaticPaintItemsCache> 256
let private staticPaintItemsCacheMaxEntries = 256
let private staticPaintItemsCacheMaxModels = 256
let private staticPaintItemsCacheMaxTotalItems = 200000
let private staticPaintItemsCacheMaxTotalItemTextChars = 5000000
let private staticPaintItemsCacheMaxTotalSourceItems = 200000
let private staticPaintItemsCacheMaxTotalSourceTextChars = 5000000

let private spreadsheetModelSlot (model: SpreadsheetModel) slotCount =
    RuntimeHelpers.GetHashCode(model) &&& (slotCount - 1)

let private paintAddCapped limit value total =
    if total > limit then total
    else min (limit + 1) (total + max 0 value)

let private paintListCount limit (items: 'T list) =
    let mutable count = 0
    let mutable remaining = items
    while count <= limit && not remaining.IsEmpty do
        count <- count + 1
        remaining <- remaining.Tail
    count

let rec private paintItemRetainedItemCount limit total item =
    let total = paintAddCapped limit 1 total
    if total > limit then total
    else
        match item with
        | VecPath (segments, _, _) ->
            paintAddCapped limit (paintListCount limit segments) total
        | VecGroup (_, children) ->
            paintItemsRetainedItemCount limit total children
        | FillRect _
        | StrokeRect _
        | DrawText _
        | GridLine _ -> total

and private paintItemsRetainedItemCount limit total (items: PaintItem list) =
    let mutable total = total
    let mutable remaining = items
    while total <= limit && not remaining.IsEmpty do
        total <- paintItemRetainedItemCount limit total remaining.Head
        remaining <- remaining.Tail
    total

let private paintItemOptionStringCharCount limit total (value: string option) =
    match value with
    | Some value -> paintAddCapped limit value.Length total
    | None -> total

let rec private paintItemTextCharCount limit total item =
    if total > limit then total
    else
        match item with
        | DrawText text ->
            let total = paintAddCapped limit text.Text.Length total
            let total = paintAddCapped limit text.Color.Length total
            paintItemOptionStringCharCount limit total text.FontName
        | FillRect (_, color)
        | StrokeRect (_, color, _)
        | GridLine (_, _, _, _, color, _) ->
            paintAddCapped limit color.Length total
        | VecPath (_, fill, stroke) ->
            let total = paintItemOptionStringCharCount limit total fill
            match stroke with
            | Some (color, _) -> paintAddCapped limit color.Length total
            | None -> total
        | VecGroup (_, children) ->
            paintItemsTextCharCount limit total children

and private paintItemsTextCharCount limit total (items: PaintItem list) =
    let mutable total = total
    let mutable remaining = items
    while total <= limit && not remaining.IsEmpty do
        total <- paintItemTextCharCount limit total remaining.Head
        remaining <- remaining.Tail
    total

let private paintOptionItemCount = function
    | Some _ -> 1
    | None -> 0

let private paintFilterItemCount (filters: Map<int, ColumnFilter>) =
    if filters.Count = 0 then 0
    else
        let mutable total = 0
        use entries = (filters :> seq<_>).GetEnumerator()
        while total <= staticPaintItemsCacheMaxTotalSourceItems && entries.MoveNext() do
            let filter = entries.Current.Value
            total <- paintAddCapped staticPaintItemsCacheMaxTotalSourceItems 1 total
            total <- paintAddCapped staticPaintItemsCacheMaxTotalSourceItems filter.VisibleValues.Count total
        total

let private paintConditionalFormatItemCount (formats: Map<CellAddress, ConditionalFormat list>) =
    if formats.Count = 0 then 0
    else
        let mutable total = 0
        use entries = (formats :> seq<_>).GetEnumerator()
        while total <= staticPaintItemsCacheMaxTotalSourceItems && entries.MoveNext() do
            let items = entries.Current.Value
            total <- paintAddCapped staticPaintItemsCacheMaxTotalSourceItems 1 total
            total <- paintAddCapped staticPaintItemsCacheMaxTotalSourceItems (paintListCount staticPaintItemsCacheMaxTotalSourceItems items) total
        total

let private paintRichRunItemCount (runs: Map<CellAddress, (string * RunOverride) list>) =
    if runs.Count = 0 then 0
    else
        let mutable total = 0
        use entries = (runs :> seq<_>).GetEnumerator()
        while total <= staticPaintItemsCacheMaxTotalSourceItems && entries.MoveNext() do
            let items = entries.Current.Value
            total <- paintAddCapped staticPaintItemsCacheMaxTotalSourceItems 1 total
            total <- paintAddCapped staticPaintItemsCacheMaxTotalSourceItems (paintListCount staticPaintItemsCacheMaxTotalSourceItems items) total
        total

let private paintValidationRuleItemCount rule =
    match rule with
    | ValidationRule.ListValues values ->
        let mutable total = 0
        let mutable remaining = values
        while total <= staticPaintItemsCacheMaxTotalSourceItems && not remaining.IsEmpty do
            total <- paintAddCapped staticPaintItemsCacheMaxTotalSourceItems 1 total
            remaining <- remaining.Tail
        total
    | ValidationRule.WholeNumber _
    | ValidationRule.DecimalRange _
    | ValidationRule.TextLength _
    | ValidationRule.DateRange _ -> 0

let private paintValidationItemCount (validations: Map<CellAddress, DataValidation>) =
    if validations.Count = 0 then 0
    else
        let mutable total = 0
        use entries = (validations :> seq<_>).GetEnumerator()
        while total <= staticPaintItemsCacheMaxTotalSourceItems && entries.MoveNext() do
            let validation = entries.Current.Value
            total <- paintAddCapped staticPaintItemsCacheMaxTotalSourceItems 1 total
            total <- paintAddCapped staticPaintItemsCacheMaxTotalSourceItems (paintValidationRuleItemCount validation.Rule) total
        total

let private paintPageSetupItemCount (pageSetup: PageSetup) =
    let limit = staticPaintItemsCacheMaxTotalSourceItems
    let mutable total = paintListCount limit pageSetup.ManualRowBreaks
    if total <= limit then
        total <- paintAddCapped limit (paintListCount limit pageSetup.ManualColBreaks) total
    if total <= limit then
        total <- paintAddCapped limit (paintOptionItemCount pageSetup.PrintArea) total
    if total <= limit then
        total <- paintAddCapped limit (paintOptionItemCount pageSetup.RepeatRows) total
    if total <= limit then
        total <- paintAddCapped limit (paintOptionItemCount pageSetup.RepeatCols) total
    total

let private paintClipboardItemCount = function
    | Some clipboard -> commandRetainedClipboardItemCount clipboard
    | None -> 0

let private paintHistoryItemCount (history: History) =
    let mutable total = 0
    let mutable past = history.Past
    while total <= staticPaintItemsCacheMaxTotalSourceItems && not past.IsEmpty do
        let command = past.Head
        total <- paintAddCapped staticPaintItemsCacheMaxTotalSourceItems command.RetainedItemCount total
        past <- past.Tail
    let mutable future = history.Future
    while total <= staticPaintItemsCacheMaxTotalSourceItems && not future.IsEmpty do
        let command = future.Head
        total <- paintAddCapped staticPaintItemsCacheMaxTotalSourceItems command.RetainedItemCount total
        future <- future.Tail
    total

let private staticPaintSourceItemCount (model: SpreadsheetModel) =
    let limit = staticPaintItemsCacheMaxTotalSourceItems
    let mutable total = 0
    if total <= limit then total <- paintAddCapped limit model.Cells.Count total
    if total <= limit then total <- paintAddCapped limit (paintListCount limit model.Tables) total
    if total <= limit then total <- paintAddCapped limit (paintConditionalFormatItemCount model.ConditionalFormats) total
    if total <= limit then total <- paintAddCapped limit (paintListCount limit model.MergedCells) total
    if total <= limit then total <- paintAddCapped limit (paintRichRunItemCount model.RichRuns) total
    if total <= limit then total <- paintAddCapped limit model.RowHeights.Count total
    if total <= limit then total <- paintAddCapped limit model.ColumnWidths.Count total
    if total <= limit then total <- paintAddCapped limit (paintFilterItemCount model.Filters) total
    if total <= limit then total <- paintAddCapped limit model.FilterHiddenRows.Length total
    if total <= limit then total <- paintAddCapped limit model.Spill.Count total
    if total <= limit then total <- paintAddCapped limit (paintValidationItemCount model.Validations) total
    if total <= limit then total <- paintAddCapped limit model.SearchResults.Length total
    if total <= limit then total <- paintAddCapped limit model.CfBoldCells.Count total
    if total <= limit then total <- paintAddCapped limit (paintPageSetupItemCount model.PageSetup) total
    if total <= limit then total <- paintAddCapped limit (paintClipboardItemCount model.Clipboard) total
    if total <= limit then total <- paintAddCapped limit (paintHistoryItemCount model.History) total
    max 1 total

let private paintOptionStringCharCount (value: string option) =
    match value with
    | Some value -> value.Length
    | None -> 0

let private paintBorderStyleStringCharCount (border: BorderStyle option) =
    match border with
    | Some border -> border.Style.Length + paintOptionStringCharCount border.Color
    | None -> 0

let private paintCellValueStringCharCount (value: CellValue) =
    let mutable total = 0
    let mutable current = value
    let mutable doneCounting = false
    while not doneCounting do
        match current with
        | CellValue.String value
        | CellValue.Error value ->
            total <- total + value.Length
            doneCounting <- true
        | CellValue.Formula (formula, cachedValue) ->
            total <- total + formula.Length
            current <- cachedValue
        | CellValue.Empty
        | CellValue.Number _
        | CellValue.Boolean _ ->
            doneCounting <- true
    total

let private paintCellFormatStringCharCount (format: CellFormat) =
    paintOptionStringCharCount format.FontName
    + paintOptionStringCharCount format.TextColor
    + paintOptionStringCharCount format.BgColor
    + paintOptionStringCharCount format.HorizontalAlign
    + paintOptionStringCharCount format.VerticalAlign
    + paintOptionStringCharCount format.NumberFormat
    + paintBorderStyleStringCharCount format.BorderTop
    + paintBorderStyleStringCharCount format.BorderBottom
    + paintBorderStyleStringCharCount format.BorderLeft
    + paintBorderStyleStringCharCount format.BorderRight
    + paintBorderStyleStringCharCount format.BorderDiagonal
    + (match format.DataBar with Some (_, color) -> color.Length | None -> 0)

let private paintCellCommentStringCharCount (comment: CellComment option) =
    match comment with
    | Some comment -> comment.Text.Length + paintOptionStringCharCount comment.Author
    | None -> 0

let private paintRunOverrideStringCharCount (run: RunOverride) =
    paintOptionStringCharCount run.Color
    + paintOptionStringCharCount run.FontName

let private paintTableStyleStringCharCount (style: TableStyle option) =
    match style with
    | Some style ->
        style.Name.Length
        + paintOptionStringCharCount style.HeaderFill
        + paintOptionStringCharCount style.StripeFill
    | None -> 0

let private paintConditionalRuleStringCharCount rule =
    match rule with
    | ConditionalFormatRule.ContainsText text -> text.Length
    | ConditionalFormatRule.GreaterThan _
    | ConditionalFormatRule.LessThan _
    | ConditionalFormatRule.Between _
    | ConditionalFormatRule.Equal _ -> 0

let private paintValidationRuleStringCharCount rule =
    match rule with
    | ValidationRule.ListValues values ->
        let mutable total = 0
        let mutable seen = 0
        let mutable remaining = values
        while total <= staticPaintItemsCacheMaxTotalSourceTextChars
              && seen <= staticPaintItemsCacheMaxTotalSourceItems
              && not remaining.IsEmpty do
            total <- paintAddCapped staticPaintItemsCacheMaxTotalSourceTextChars remaining.Head.Length total
            seen <- seen + 1
            remaining <- remaining.Tail
        total
    | ValidationRule.WholeNumber _
    | ValidationRule.DecimalRange _
    | ValidationRule.TextLength _
    | ValidationRule.DateRange _ -> 0

let private paintHeaderFooterStringCharCount (headerFooter: HeaderFooter) =
    paintOptionStringCharCount headerFooter.OddHeader
    + paintOptionStringCharCount headerFooter.OddFooter
    + paintOptionStringCharCount headerFooter.EvenHeader
    + paintOptionStringCharCount headerFooter.EvenFooter
    + paintOptionStringCharCount headerFooter.FirstHeader
    + paintOptionStringCharCount headerFooter.FirstFooter

let private paintClipboardTextCharCount = function
    | Some clipboard -> commandRetainedClipboardTextCharCount clipboard
    | None -> 0

let private paintHistoryTextCharCount (history: History) =
    let mutable total = 0
    let mutable past = history.Past
    while total <= staticPaintItemsCacheMaxTotalSourceTextChars && not past.IsEmpty do
        let command = past.Head
        total <-
            paintAddCapped
                staticPaintItemsCacheMaxTotalSourceTextChars
                (command.Description.Length + command.RetainedTextCharCount)
                total
        past <- past.Tail
    let mutable future = history.Future
    while total <= staticPaintItemsCacheMaxTotalSourceTextChars && not future.IsEmpty do
        let command = future.Head
        total <-
            paintAddCapped
                staticPaintItemsCacheMaxTotalSourceTextChars
                (command.Description.Length + command.RetainedTextCharCount)
                total
        future <- future.Tail
    total

let private staticPaintSourceTextCharCount (model: SpreadsheetModel) =
    let limit = staticPaintItemsCacheMaxTotalSourceTextChars
    let cappedLimit = limit + 1
    let mutable total = paintOptionStringCharCount model.SearchQuery + model.EditBuffer.Length + model.Macros.Length
    if total <= limit then
        total <- paintAddCapped limit (paintHeaderFooterStringCharCount model.PageSetup.HeaderFooter) total
    if total <= limit then
        total <- paintAddCapped limit (paintClipboardTextCharCount model.Clipboard) total
    if total <= limit then
        total <- paintAddCapped limit (paintHistoryTextCharCount model.History) total
    if total <= limit && model.Filters.Count > 0 then
        use filters = (model.Filters :> seq<_>).GetEnumerator()
        while total <= limit && filters.MoveNext() do
            let filter = filters.Current.Value
            use values = (filter.VisibleValues :> seq<string>).GetEnumerator()
            while total <= limit && values.MoveNext() do
                total <- paintAddCapped limit values.Current.Length total
    if total <= limit && model.ConditionalFormats.Count > 0 then
        use conditionalFormats = (model.ConditionalFormats :> seq<_>).GetEnumerator()
        while total <= limit && conditionalFormats.MoveNext() do
            let mutable formats = conditionalFormats.Current.Value
            while total <= limit && not formats.IsEmpty do
                let format = formats.Head
                total <- paintAddCapped limit (paintConditionalRuleStringCharCount format.Rule) total
                total <- paintAddCapped limit (cellFormatTextCharCountFor limit format.Style) total
                formats <- formats.Tail
    if total <= limit && model.RichRuns.Count > 0 then
        use richRuns = (model.RichRuns :> seq<_>).GetEnumerator()
        while total <= limit && richRuns.MoveNext() do
            let mutable runs = richRuns.Current.Value
            while total <= limit && not runs.IsEmpty do
                let text, run = runs.Head
                total <- paintAddCapped limit (text.Length + paintRunOverrideStringCharCount run) total
                runs <- runs.Tail
    let mutable tables = model.Tables
    let mutable seenTables = 0
    while total <= limit
          && seenTables <= staticPaintItemsCacheMaxTotalSourceItems
          && not tables.IsEmpty do
        let table = tables.Head
        total <- paintAddCapped limit table.Name.Length total
        if total <= limit then
            total <- paintAddCapped limit (paintTableStyleStringCharCount table.Style) total
        seenTables <- seenTables + 1
        tables <- tables.Tail
    if total <= limit && model.Validations.Count > 0 then
        use validations = (model.Validations :> seq<_>).GetEnumerator()
        while total <= limit && validations.MoveNext() do
            let validation = validations.Current.Value
            total <- paintAddCapped limit (paintValidationRuleStringCharCount validation.Rule) total
            total <- paintAddCapped limit (paintOptionStringCharCount validation.ErrorMessage) total
            total <- paintAddCapped limit (paintOptionStringCharCount validation.InputMessage) total
    if total <= limit && model.Cells.Count > 0 then
        model.Cells
        |> Map.exists (fun _ cell ->
            total <- paintAddCapped limit (cellDataTextCharCountFor limit cell) total
            total > limit)
        |> ignore
    min total cappedLimit

let private staticPaintSourceTextCharCountFor sourceItemCount model =
    if sourceItemCount > staticPaintItemsCacheMaxTotalSourceItems then
        staticPaintItemsCacheMaxTotalSourceTextChars + 1
    else
        staticPaintSourceTextCharCount model

type private StaticPaintSourceMetrics = {
    Cells: Map<CellAddress, CellData>
    Tables: Table list
    ConditionalFormats: Map<CellAddress, ConditionalFormat list>
    MergedCells: MergedCell list
    RichRuns: Map<CellAddress, (string * RunOverride) list>
    RowHeights: Map<int, float>
    ColumnWidths: Map<int, float>
    Filters: Map<int, ColumnFilter>
    FilterHiddenRows: int array
    Spill: Map<CellAddress, CellAddress>
    Validations: Map<CellAddress, DataValidation>
    SearchResults: CellAddress array
    CfBoldCells: Set<CellAddress>
    PageSetup: PageSetup
    Clipboard: ClipboardData option
    History: History
    SearchQuery: string option
    EditBuffer: string
    Macros: string
    SourceItemCount: int
    SourceTextCharCount: int
}

let mutable private staticPaintSourceMetricsCache: StaticPaintSourceMetrics option = None

let private staticPaintSourceMetricsMatch (model: SpreadsheetModel) (cached: StaticPaintSourceMetrics) =
    obj.ReferenceEquals(cached.Cells, model.Cells)
    && obj.ReferenceEquals(cached.Tables, model.Tables)
    && obj.ReferenceEquals(cached.ConditionalFormats, model.ConditionalFormats)
    && obj.ReferenceEquals(cached.MergedCells, model.MergedCells)
    && obj.ReferenceEquals(cached.RichRuns, model.RichRuns)
    && obj.ReferenceEquals(cached.RowHeights, model.RowHeights)
    && obj.ReferenceEquals(cached.ColumnWidths, model.ColumnWidths)
    && obj.ReferenceEquals(cached.Filters, model.Filters)
    && obj.ReferenceEquals(cached.FilterHiddenRows, model.FilterHiddenRows)
    && obj.ReferenceEquals(cached.Spill, model.Spill)
    && obj.ReferenceEquals(cached.Validations, model.Validations)
    && obj.ReferenceEquals(cached.SearchResults, model.SearchResults)
    && obj.ReferenceEquals(cached.CfBoldCells, model.CfBoldCells)
    && obj.ReferenceEquals(cached.PageSetup, model.PageSetup)
    && obj.ReferenceEquals(cached.Clipboard, model.Clipboard)
    && obj.ReferenceEquals(cached.History, model.History)
    && obj.ReferenceEquals(cached.SearchQuery, model.SearchQuery)
    && obj.ReferenceEquals(cached.EditBuffer, model.EditBuffer)
    && obj.ReferenceEquals(cached.Macros, model.Macros)

let private staticPaintSourceMetrics (model: SpreadsheetModel) =
    match staticPaintSourceMetricsCache with
    | Some cached when staticPaintSourceMetricsMatch model cached ->
        cached.SourceItemCount, cached.SourceTextCharCount
    | _ ->
        let sourceItemCount = staticPaintSourceItemCount model
        let sourceTextCharCount = staticPaintSourceTextCharCountFor sourceItemCount model
        if sourceItemCount <= staticPaintItemsCacheMaxTotalSourceItems
           && sourceTextCharCount <= staticPaintItemsCacheMaxTotalSourceTextChars then
            staticPaintSourceMetricsCache <- Some {
                Cells = model.Cells
                Tables = model.Tables
                ConditionalFormats = model.ConditionalFormats
                MergedCells = model.MergedCells
                RichRuns = model.RichRuns
                RowHeights = model.RowHeights
                ColumnWidths = model.ColumnWidths
                Filters = model.Filters
                FilterHiddenRows = model.FilterHiddenRows
                Spill = model.Spill
                Validations = model.Validations
                SearchResults = model.SearchResults
                CfBoldCells = model.CfBoldCells
                PageSetup = model.PageSetup
                Clipboard = model.Clipboard
                History = model.History
                SearchQuery = model.SearchQuery
                EditBuffer = model.EditBuffer
                Macros = model.Macros
                SourceItemCount = sourceItemCount
                SourceTextCharCount = sourceTextCharCount
            }
        else
            staticPaintSourceMetricsCache <- None
        sourceItemCount, sourceTextCharCount

let private filterPaintProjectionCacheRetainable (source: SpreadsheetModel) (projected: SpreadsheetModel) =
    let sourceItemCount = staticPaintSourceItemCount source
    let projectedItemCount = staticPaintSourceItemCount projected
    sourceItemCount <= staticPaintItemsCacheMaxTotalSourceItems
    && projectedItemCount <= staticPaintItemsCacheMaxTotalSourceItems
    && staticPaintSourceTextCharCountFor sourceItemCount source <= staticPaintItemsCacheMaxTotalSourceTextChars
    && staticPaintSourceTextCharCountFor projectedItemCount projected <= staticPaintItemsCacheMaxTotalSourceTextChars

let private withFilterRowHeights (model: SpreadsheetModel) =
    if Map.isEmpty model.Filters then
        model
    else
        let mutable projected = Unchecked.defaultof<SpreadsheetModel>
        if filterPaintProjectionCache.TryGetValue(model, &projected) then
            projected
        else
            let projected = applyFilterRowHeights model
            if filterPaintProjectionCacheRetainable model projected then
                filterPaintProjectionCache.Remove model |> ignore
                filterPaintProjectionCache.Add(model, projected)
            projected

let tryBuildSparseDefaultPaintParts (model: SpreadsheetModel) (viewport: Viewport) =
    let model = model |> withFilterRowHeights
    match tryGetSparseDefaultPaintPartsLast model viewport with
    | ValueSome cached -> cached
    | ValueNone ->
        match tryGetSparseDefaultPaintCacheEntry model viewport with
        | Some entry ->
            let result = Some (sparseDefaultPartsFromEntry entry.PrefixItems [] model entry)
            if sparseDefaultPaintCacheEntryRetainable entry then
                rememberSparseDefaultPaintPartsLast model viewport result
            else
                result
        | None -> None

let private sameMeasurer (left: TextMeasurer option) (right: TextMeasurer option) =
    match left, right with
    | None, None -> true
    | Some a, Some b -> obj.ReferenceEquals(a, b)
    | _ -> false

let private staticPaintGeometryKey (viewport: Viewport) = {
    ScrollX = viewport.ScrollX
    ScrollY = viewport.ScrollY
    Width = viewport.Width
    Height = viewport.Height
    HeaderHeight = viewport.HeaderHeight
    RowHeaderWidth = viewport.RowHeaderWidth
}

let private staticPaintGeometrySlot
    (slotCount: int)
    scrollX
    scrollY
    width
    height
    headerHeight
    rowHeaderWidth =
    let mutable hashCode = scrollX.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ scrollY.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ width.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ height.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ headerHeight.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ rowHeaderWidth.GetHashCode()
    hashCode &&& (slotCount - 1)

let private staticPaintGeometryKeyMatchesViewport key (viewport: Viewport) =
    key.ScrollX = viewport.ScrollX
    && key.ScrollY = viewport.ScrollY
    && key.Width = viewport.Width
    && key.Height = viewport.Height
    && key.HeaderHeight = viewport.HeaderHeight
    && key.RowHeaderWidth = viewport.RowHeaderWidth

let private tryGetStaticPaintGeometryFastSlot (cache: StaticPaintItemsCache) (viewport: Viewport) =
    let slot =
        staticPaintGeometrySlot
            cache.EntryFastSlots.Length
            viewport.ScrollX
            viewport.ScrollY
            viewport.Width
            viewport.Height
            viewport.HeaderHeight
            viewport.RowHeaderWidth
    let fast = cache.EntryFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && staticPaintGeometryKeyMatchesViewport fast.Key viewport
       && sameMeasurer fast.Entry.Measurer viewport.Measurer then
        ValueSome fast.Entry
    else
        ValueNone

let private setStaticPaintGeometryFastSlot cache key entry =
    let slot =
        staticPaintGeometrySlot
            cache.EntryFastSlots.Length
            key.ScrollX
            key.ScrollY
            key.Width
            key.Height
            key.HeaderHeight
            key.RowHeaderWidth
    cache.EntryFastSlots.[slot] <- {
        Key = key
        Entry = entry
    }

let private removeStaticPaintItemsCacheEntry key (cache: StaticPaintItemsCache) =
    let mutable existing = Unchecked.defaultof<StaticPaintItemsCacheEntry>
    if cache.Entries.TryGetValue(key, &existing) then
        clearArray cache.EntryFastSlots
        cache.TotalItemCount <- max 0 (cache.TotalItemCount - existing.ItemCount)
        cache.TotalItemTextCharCount <- max 0 (cache.TotalItemTextCharCount - existing.ItemTextCharCount)
        cache.Entries.Remove key |> ignore

let private addStaticPaintItemsCacheEntry key entry (cache: StaticPaintItemsCache) =
    let mutable existing = Unchecked.defaultof<StaticPaintItemsCacheEntry>
    if entry.ItemCount > staticPaintItemsCacheMaxTotalItems
       || entry.ItemTextCharCount > staticPaintItemsCacheMaxTotalItemTextChars then
        removeStaticPaintItemsCacheEntry key cache
    else
        if cache.Entries.TryGetValue(key, &existing) then
            cache.TotalItemCount <- max 0 (cache.TotalItemCount - existing.ItemCount)
            cache.TotalItemTextCharCount <- max 0 (cache.TotalItemTextCharCount - existing.ItemTextCharCount)
        else
            cache.Order.Enqueue key
        cache.Entries.[key] <- entry
        cache.TotalItemCount <- cache.TotalItemCount + entry.ItemCount
        cache.TotalItemTextCharCount <- cache.TotalItemTextCharCount + entry.ItemTextCharCount
        setStaticPaintGeometryFastSlot cache key entry
        let mutable removedEntry = false
        while (cache.Entries.Count > staticPaintItemsCacheMaxEntries
               || cache.TotalItemCount > staticPaintItemsCacheMaxTotalItems
               || cache.TotalItemTextCharCount > staticPaintItemsCacheMaxTotalItemTextChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> key then
                removedEntry <- true
                removeStaticPaintItemsCacheEntry oldest cache
        if cache.Entries.Count > staticPaintItemsCacheMaxEntries
           || cache.TotalItemCount > staticPaintItemsCacheMaxTotalItems
           || cache.TotalItemTextCharCount > staticPaintItemsCacheMaxTotalItemTextChars then
            let current = cache.Entries.[key]
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue key
            cache.Entries.[key] <- current
            cache.TotalItemCount <- current.ItemCount
            cache.TotalItemTextCharCount <- current.ItemTextCharCount
            clearArray cache.EntryFastSlots
            setStaticPaintGeometryFastSlot cache key current
        elif removedEntry then
            clearArray cache.EntryFastSlots
            let mutable current = Unchecked.defaultof<StaticPaintItemsCacheEntry>
            if cache.Entries.TryGetValue(key, &current) then
                setStaticPaintGeometryFastSlot cache key current

let private staticPaintItemsCacheRetainable (cache: StaticPaintItemsCache) =
    cache.Entries.Count > 0
    && cache.TotalItemCount <= staticPaintItemsCacheMaxTotalItems
    && cache.TotalItemTextCharCount <= staticPaintItemsCacheMaxTotalItemTextChars
    && cache.SourceItemCount <= staticPaintItemsCacheMaxTotalSourceItems
    && cache.SourceTextCharCount <= staticPaintItemsCacheMaxTotalSourceTextChars

let private setStaticPaintItemsCache entries =
    staticPaintItemsCache <- entries
    staticPaintItemsCacheByModel.Clear()
    clearArray staticPaintItemsFastSlots
    for entry in entries do
        staticPaintItemsCacheByModel.[entry.Model] <- entry
        staticPaintItemsFastSlots.[spreadsheetModelSlot entry.Model staticPaintItemsFastSlots.Length] <- entry

let private tryGetStaticPaintItemsCache model =
    let slot = spreadsheetModelSlot model staticPaintItemsFastSlots.Length
    let fast = staticPaintItemsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null)) && obj.ReferenceEquals(fast.Model, model) then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<StaticPaintItemsCache>
        if staticPaintItemsCacheByModel.TryGetValue(model, &cached)
           && obj.ReferenceEquals(cached.Model, model) then
            staticPaintItemsFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private staticPaintItemsCacheWithoutModel model =
    let mutable cached = Unchecked.defaultof<StaticPaintItemsCache>
    if staticPaintItemsCacheByModel.TryGetValue(model, &cached) then
        staticPaintItemsCache
        |> List.filter (fun cached -> not (obj.ReferenceEquals(cached.Model, model)))
    else
        staticPaintItemsCache

let private removeStaticPaintItemsCacheForModel model =
    let entries = staticPaintItemsCacheWithoutModel model
    if not (obj.ReferenceEquals(entries, staticPaintItemsCache)) then
        setStaticPaintItemsCache entries

let private rememberStaticPaintItemsCache (cache: StaticPaintItemsCache) =
    if not (staticPaintItemsCacheRetainable cache) then
        removeStaticPaintItemsCacheForModel cache.Model
    else
        let mutable keptCount = 1
        let mutable keptItems = cache.TotalItemCount
        let mutable keptItemTextChars = cache.TotalItemTextCharCount
        let mutable keptSourceItems = cache.SourceItemCount
        let mutable keptSourceTextChars = cache.SourceTextCharCount
        let retained =
            staticPaintItemsCacheWithoutModel cache.Model
            |> List.choose (fun candidate ->
                let candidateItems = candidate.TotalItemCount
                let candidateItemTextChars = candidate.TotalItemTextCharCount
                let candidateSourceItems = candidate.SourceItemCount
                let candidateSourceTextChars = candidate.SourceTextCharCount
                if staticPaintItemsCacheRetainable candidate
                   && keptCount < staticPaintItemsCacheMaxModels
                   && keptItems + candidateItems <= staticPaintItemsCacheMaxTotalItems
                   && keptItemTextChars + candidateItemTextChars <= staticPaintItemsCacheMaxTotalItemTextChars
                   && keptSourceItems + candidateSourceItems <= staticPaintItemsCacheMaxTotalSourceItems
                   && keptSourceTextChars + candidateSourceTextChars <= staticPaintItemsCacheMaxTotalSourceTextChars then
                    keptCount <- keptCount + 1
                    keptItems <- keptItems + candidateItems
                    keptItemTextChars <- keptItemTextChars + candidateItemTextChars
                    keptSourceItems <- keptSourceItems + candidateSourceItems
                    keptSourceTextChars <- keptSourceTextChars + candidateSourceTextChars
                    Some candidate
                else
                    None)
        setStaticPaintItemsCache (cache :: retained)

let private iterPaintItemsPrepared (model: SpreadsheetModel) (viewport: Viewport) (emit: PaintItem -> unit) =
    emit (FillRect ({ X = 0.0; Y = 0.0; Width = viewport.Width; Height = viewport.Height }, "#ffffff"))
    iterHeaderPaintItems model viewport emit
    // Static (PNG/PDF) path: drop the zone tag; everything paints at
    // absolute cell coordinates because ScrollX/Y are zero.
    iterCellPaintItems model viewport (fun _zone item -> emit item)
    // Freeze divider lines on top of cell strokes but under any
    // selection / DrawText emitted later. The cell-item zone tag is
    // irrelevant here; we draw the divider in the model coordinate
    // system that's already baked into the cell rects above.
    iterFreezeDividerItems model viewport emit

let iterPaintItems (model: SpreadsheetModel) (viewport: Viewport) (emit: PaintItem -> unit) =
    let model = model |> withFilterRowHeights
    iterPaintItemsPrepared model viewport emit

let buildPaintItems (model: SpreadsheetModel) (viewport: Viewport): PaintItem list =
    let model = model |> withFilterRowHeights
    match tryGetStaticPaintItemsCache model with
    | ValueSome cached when staticPaintItemsCacheRetainable cached ->
        match tryGetStaticPaintGeometryFastSlot cached viewport with
        | ValueSome entry -> entry.Items
        | ValueNone ->
            let key = staticPaintGeometryKey viewport
            let mutable entry = Unchecked.defaultof<StaticPaintItemsCacheEntry>
            if cached.Entries.TryGetValue(key, &entry) && sameMeasurer entry.Measurer viewport.Measurer then
                setStaticPaintGeometryFastSlot cached key entry
                entry.Items
            else
                let values = ResizeArray<PaintItem>(estimatedPaintItemCount model viewport)
                iterPaintItemsPrepared model viewport values.Add
                let itemCount = values.Count
                let items = resizeArrayToList values
                if itemCount <= staticPaintItemsCacheMaxTotalItems then
                    addStaticPaintItemsCacheEntry key {
                        Measurer = viewport.Measurer
                        Items = items
                        ItemCount = itemCount
                        ItemTextCharCount = paintItemsTextCharCount staticPaintItemsCacheMaxTotalItemTextChars 0 items
                    } cached
                else
                    removeStaticPaintItemsCacheEntry key cached
                rememberStaticPaintItemsCache cached
                items
    | ValueSome cached ->
        rememberStaticPaintItemsCache cached
        let values = ResizeArray<PaintItem>(estimatedPaintItemCount model viewport)
        iterPaintItemsPrepared model viewport values.Add
        resizeArrayToList values
    | ValueNone ->
        let values = ResizeArray<PaintItem>(estimatedPaintItemCount model viewport)
        iterPaintItemsPrepared model viewport values.Add
        let itemCount = values.Count
        let items = resizeArrayToList values
        if itemCount <= staticPaintItemsCacheMaxTotalItems then
            let key = staticPaintGeometryKey viewport
            let entries = Dictionary<StaticPaintGeometryKey, StaticPaintItemsCacheEntry>()
            let order = Queue<StaticPaintGeometryKey>()
            let sourceItemCount, sourceTextCharCount = staticPaintSourceMetrics model
            let cache = {
                Model = model
                SourceItemCount = sourceItemCount
                SourceTextCharCount = sourceTextCharCount
                Entries = entries
                Order = order
                EntryFastSlots = Array.zeroCreate<StaticPaintGeometryFastEntry> 128
                TotalItemCount = 0
                TotalItemTextCharCount = 0
            }
            addStaticPaintItemsCacheEntry key {
                Measurer = viewport.Measurer
                Items = items
                ItemCount = itemCount
                ItemTextCharCount = paintItemsTextCharCount staticPaintItemsCacheMaxTotalItemTextChars 0 items
            } cache
            rememberStaticPaintItemsCache cache
        items

let private viewportTranslation (viewport: Viewport) zone =
    match zone with
    | BodyZone -> -viewport.ScrollX, -viewport.ScrollY
    | ColumnHeaderZone -> -viewport.ScrollX, 0.0
    | RowHeaderZone -> 0.0, -viewport.ScrollY
    | CornerZone -> 0.0, 0.0
    | FrozenRowsZone -> -viewport.ScrollX, 0.0
    | FrozenColsZone -> 0.0, -viewport.ScrollY
    | FrozenCornerZone -> 0.0, 0.0

let private translateRectBy dx dy rect =
    if dx = 0.0 && dy = 0.0 then rect
    else { rect with X = rect.X + dx; Y = rect.Y + dy }

let private translatePaintItem (viewport: Viewport) zone item =
    let dx, dy = viewportTranslation viewport zone
    if dx = 0.0 && dy = 0.0 then item else
    match item with
    | FillRect (rect, color) -> FillRect (translateRectBy dx dy rect, color)
    | StrokeRect (rect, color, width) -> StrokeRect (translateRectBy dx dy rect, color, width)
    | DrawText text -> DrawText { text with Rect = translateRectBy dx dy text.Rect }
    | GridLine (x1, y1, x2, y2, color, width) ->
        GridLine (x1 + dx, y1 + dy, x2 + dx, y2 + dy, color, width)
    | VecPath (segs, fill, stroke) ->
        // The zone offset is a pure translation; shift every path point by it.
        let mv x y = x + dx, y + dy
        let translated = ResizeArray<VecSeg>(segs.Length)
        for seg in segs do
            translated.Add(
                match seg with
                | VMove (x, y) -> let a, b = mv x y in VMove (a, b)
                | VLine (x, y) -> let a, b = mv x y in VLine (a, b)
                | VCubic (x1, y1, x2, y2, x, y) ->
                    let a1, b1 = mv x1 y1
                    let a2, b2 = mv x2 y2
                    let a, b = mv x y
                    VCubic (a1, b1, a2, b2, a, b)
                | VClose -> VClose)
        VecPath (resizeArrayToList translated, fill, stroke)
    | VecGroup ((a, b, c, d, e, f), children) ->
        // Fold the zone offset into the group's translation component;
        // children stay in group-local coordinates.
        VecGroup ((a, b, c, d, e + dx, f + dy), children)

let private iterViewportPaintItemsUncached (model: SpreadsheetModel) (viewport: Viewport) (emit: ViewportZone -> PaintItem -> unit) =
    let model = model |> withFilterRowHeights
    emit CornerZone (FillRect ({ X = 0.0; Y = 0.0; Width = viewport.Width; Height = viewport.Height }, "#ffffff"))

    iterHeaders model viewport (fun header ->
        match header with
        | CornerHeader rect ->
            emit CornerZone (FillRect (rect, "#e8e8e8"))
            emit CornerZone (StrokeRect (rect, "#a6a6a6", 1.0))
        | ColumnHeader (col, rect) ->
            let zone = ColumnHeaderZone
            emit zone (translatePaintItem viewport zone (FillRect (rect, "#f2f2f2")))
            emit zone (translatePaintItem viewport zone (StrokeRect (rect, "#a6a6a6", 1.0)))
            emit zone (translatePaintItem viewport zone (DrawText {
                Text = columnToLetter col
                Rect = rect
                Color = "#444444"
                FontSize = 11.0
                FontName = None
                Bold = false
                Italic = false
                Underline = false
                UnderlineDouble = false
                Strikethrough = false
                Align = AlignCenter
                VerticalAlign = VAlignMiddle
                LineIndex = 0
                LineCount = 1
                BaselineShift = 0.0
                Rotation = 0
            }))
        | RowHeader (row, rect) ->
            let zone = RowHeaderZone
            emit zone (translatePaintItem viewport zone (FillRect (rect, "#f2f2f2")))
            emit zone (translatePaintItem viewport zone (StrokeRect (rect, "#a6a6a6", 1.0)))
            emit zone (translatePaintItem viewport zone (DrawText {
                Text = string (row + 1)
                Rect = rect
                Color = "#444444"
                FontSize = 11.0
                FontName = None
                Bold = false
                Italic = false
                Underline = false
                UnderlineDouble = false
                Strikethrough = false
                Align = AlignCenter
                VerticalAlign = VAlignMiddle
                LineIndex = 0
                LineCount = 1
                BaselineShift = 0.0
                Rotation = 0
            })))

    iterCellPaintItems model viewport (fun zone item ->
        emit zone (translatePaintItem viewport zone item)
    )

    // Freeze dividers paint with the same zone as the body they border,
    // so they scroll with the body horizontally (row divider) or
    // vertically (col divider) and disappear under the frozen band when
    // the body scrolls past them. Excel does the same: the divider
    // moves with the scrolling region, not the frozen one.
    iterFreezeDividerItems model viewport (fun item ->
        let zone =
            match item with
            | GridLine (x1, _, x2, _, _, _) when x1 = x2 ->
                // Vertical line (cols divider) — scrolls vertically with
                // body but its X is pinned (no horizontal scroll).
                FrozenRowsZone
            | _ ->
                // Horizontal line (rows divider) — scrolls horizontally
                // with body, pinned vertically.
                FrozenColsZone
        emit zone (translatePaintItem viewport zone item)
    )

type private ViewportPaintGeometryKey = {
    ScrollX: float
    ScrollY: float
    Width: float
    Height: float
    HeaderHeight: float
    RowHeaderWidth: float
}

type private ViewportPaintItemsCacheEntry = {
    Measurer: TextMeasurer option
    Items: ViewportPaintItem list
    ItemCount: int
    ItemTextCharCount: int
}

type private ViewportPaintGeometryFastEntry = {
    Key: ViewportPaintGeometryKey
    Entry: ViewportPaintItemsCacheEntry
}

type private ViewportPaintItemsCache = {
    Model: SpreadsheetModel
    SourceItemCount: int
    SourceTextCharCount: int
    Entries: Dictionary<ViewportPaintGeometryKey, ViewportPaintItemsCacheEntry>
    Order: Queue<ViewportPaintGeometryKey>
    EntryFastSlots: ViewportPaintGeometryFastEntry array
    mutable TotalItemCount: int
    mutable TotalItemTextCharCount: int
}

let mutable private viewportPaintItemsCache: ViewportPaintItemsCache list = []
let private viewportPaintItemsCacheByModel =
    Dictionary<SpreadsheetModel, ViewportPaintItemsCache>(spreadsheetModelReferenceComparer)
let private viewportPaintItemsFastSlots = Array.zeroCreate<ViewportPaintItemsCache> 256
let private viewportPaintItemsCacheMaxEntries = 256
let private viewportPaintItemsCacheMaxModels = 256
let private viewportPaintItemsCacheMaxTotalItems = 200000
let private viewportPaintItemsCacheMaxTotalItemTextChars = 5000000
let private viewportPaintItemsCacheMaxTotalSourceItems = 200000
let private viewportPaintItemsCacheMaxTotalSourceTextChars = 5000000

let private viewportPaintGeometryKey (viewport: Viewport) = {
    ScrollX = viewport.ScrollX
    ScrollY = viewport.ScrollY
    Width = viewport.Width
    Height = viewport.Height
    HeaderHeight = viewport.HeaderHeight
    RowHeaderWidth = viewport.RowHeaderWidth
}

let private viewportPaintGeometrySlot
    (slotCount: int)
    scrollX
    scrollY
    width
    height
    headerHeight
    rowHeaderWidth =
    let mutable hashCode = scrollX.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ scrollY.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ width.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ height.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ headerHeight.GetHashCode()
    hashCode <- (hashCode * 397) ^^^ rowHeaderWidth.GetHashCode()
    hashCode &&& (slotCount - 1)

let private viewportPaintGeometryKeyMatchesViewport key (viewport: Viewport) =
    key.ScrollX = viewport.ScrollX
    && key.ScrollY = viewport.ScrollY
    && key.Width = viewport.Width
    && key.Height = viewport.Height
    && key.HeaderHeight = viewport.HeaderHeight
    && key.RowHeaderWidth = viewport.RowHeaderWidth

let private tryGetViewportPaintGeometryFastSlot (cache: ViewportPaintItemsCache) (viewport: Viewport) =
    let slot =
        viewportPaintGeometrySlot
            cache.EntryFastSlots.Length
            viewport.ScrollX
            viewport.ScrollY
            viewport.Width
            viewport.Height
            viewport.HeaderHeight
            viewport.RowHeaderWidth
    let fast = cache.EntryFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null))
       && viewportPaintGeometryKeyMatchesViewport fast.Key viewport
       && sameMeasurer fast.Entry.Measurer viewport.Measurer then
        ValueSome fast.Entry
    else
        ValueNone

let private setViewportPaintGeometryFastSlot cache key entry =
    let slot =
        viewportPaintGeometrySlot
            cache.EntryFastSlots.Length
            key.ScrollX
            key.ScrollY
            key.Width
            key.Height
            key.HeaderHeight
            key.RowHeaderWidth
    cache.EntryFastSlots.[slot] <- {
        Key = key
        Entry = entry
    }

let private viewportPaintItemsRetainedItemCount limit (items: ViewportPaintItem list) =
    let mutable total = 0
    let mutable remaining = items
    while total <= limit && not remaining.IsEmpty do
        total <- paintItemRetainedItemCount limit total remaining.Head.Item
        remaining <- remaining.Tail
    total

let private viewportPaintItemsTextCharCount limit (items: ViewportPaintItem list) =
    let mutable total = 0
    let mutable remaining = items
    while total <= limit && not remaining.IsEmpty do
        total <- paintItemTextCharCount limit total remaining.Head.Item
        remaining <- remaining.Tail
    total

let private removeViewportPaintItemsCacheEntry key (cache: ViewportPaintItemsCache) =
    let mutable existing = Unchecked.defaultof<ViewportPaintItemsCacheEntry>
    if cache.Entries.TryGetValue(key, &existing) then
        clearArray cache.EntryFastSlots
        cache.TotalItemCount <- max 0 (cache.TotalItemCount - existing.ItemCount)
        cache.TotalItemTextCharCount <- max 0 (cache.TotalItemTextCharCount - existing.ItemTextCharCount)
        cache.Entries.Remove key |> ignore

let private addViewportPaintItemsCacheEntry key entry (cache: ViewportPaintItemsCache) =
    let mutable existing = Unchecked.defaultof<ViewportPaintItemsCacheEntry>
    if entry.ItemCount > viewportPaintItemsCacheMaxTotalItems
       || entry.ItemTextCharCount > viewportPaintItemsCacheMaxTotalItemTextChars then
        removeViewportPaintItemsCacheEntry key cache
    else
        if cache.Entries.TryGetValue(key, &existing) then
            cache.TotalItemCount <- max 0 (cache.TotalItemCount - existing.ItemCount)
            cache.TotalItemTextCharCount <- max 0 (cache.TotalItemTextCharCount - existing.ItemTextCharCount)
        else
            cache.Order.Enqueue key
        cache.Entries.[key] <- entry
        cache.TotalItemCount <- cache.TotalItemCount + entry.ItemCount
        cache.TotalItemTextCharCount <- cache.TotalItemTextCharCount + entry.ItemTextCharCount
        setViewportPaintGeometryFastSlot cache key entry
        let mutable removedEntry = false
        while (cache.Entries.Count > viewportPaintItemsCacheMaxEntries
               || cache.TotalItemCount > viewportPaintItemsCacheMaxTotalItems
               || cache.TotalItemTextCharCount > viewportPaintItemsCacheMaxTotalItemTextChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if oldest <> key then
                removedEntry <- true
                removeViewportPaintItemsCacheEntry oldest cache
        if cache.Entries.Count > viewportPaintItemsCacheMaxEntries
           || cache.TotalItemCount > viewportPaintItemsCacheMaxTotalItems
           || cache.TotalItemTextCharCount > viewportPaintItemsCacheMaxTotalItemTextChars then
            let current = cache.Entries.[key]
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue key
            cache.Entries.[key] <- current
            cache.TotalItemCount <- current.ItemCount
            cache.TotalItemTextCharCount <- current.ItemTextCharCount
            clearArray cache.EntryFastSlots
            setViewportPaintGeometryFastSlot cache key current
        elif removedEntry then
            clearArray cache.EntryFastSlots
            let mutable current = Unchecked.defaultof<ViewportPaintItemsCacheEntry>
            if cache.Entries.TryGetValue(key, &current) then
                setViewportPaintGeometryFastSlot cache key current

let private viewportPaintItemsCacheRetainable (cache: ViewportPaintItemsCache) =
    cache.Entries.Count > 0
    && cache.TotalItemCount <= viewportPaintItemsCacheMaxTotalItems
    && cache.TotalItemTextCharCount <= viewportPaintItemsCacheMaxTotalItemTextChars
    && cache.SourceItemCount <= viewportPaintItemsCacheMaxTotalSourceItems
    && cache.SourceTextCharCount <= viewportPaintItemsCacheMaxTotalSourceTextChars

let private setViewportPaintItemsCache entries =
    viewportPaintItemsCache <- entries
    viewportPaintItemsCacheByModel.Clear()
    clearArray viewportPaintItemsFastSlots
    for entry in entries do
        viewportPaintItemsCacheByModel.[entry.Model] <- entry
        viewportPaintItemsFastSlots.[spreadsheetModelSlot entry.Model viewportPaintItemsFastSlots.Length] <- entry

let private tryGetViewportPaintItemsCache model =
    let slot = spreadsheetModelSlot model viewportPaintItemsFastSlots.Length
    let fast = viewportPaintItemsFastSlots.[slot]
    if not (obj.ReferenceEquals(fast, null)) && obj.ReferenceEquals(fast.Model, model) then
        ValueSome fast
    else
        let mutable cached = Unchecked.defaultof<ViewportPaintItemsCache>
        if viewportPaintItemsCacheByModel.TryGetValue(model, &cached)
           && obj.ReferenceEquals(cached.Model, model) then
            viewportPaintItemsFastSlots.[slot] <- cached
            ValueSome cached
        else
            ValueNone

let private viewportPaintItemsCacheWithoutModel model =
    let mutable cached = Unchecked.defaultof<ViewportPaintItemsCache>
    if viewportPaintItemsCacheByModel.TryGetValue(model, &cached) then
        viewportPaintItemsCache
        |> List.filter (fun cached -> not (obj.ReferenceEquals(cached.Model, model)))
    else
        viewportPaintItemsCache

let private removeViewportPaintItemsCacheForModel model =
    let entries = viewportPaintItemsCacheWithoutModel model
    if not (obj.ReferenceEquals(entries, viewportPaintItemsCache)) then
        setViewportPaintItemsCache entries

let private rememberViewportPaintItemsCache (cache: ViewportPaintItemsCache) =
    if not (viewportPaintItemsCacheRetainable cache) then
        removeViewportPaintItemsCacheForModel cache.Model
    else
        let mutable keptCount = 1
        let mutable keptItems = cache.TotalItemCount
        let mutable keptItemTextChars = cache.TotalItemTextCharCount
        let mutable keptSourceItems = cache.SourceItemCount
        let mutable keptSourceTextChars = cache.SourceTextCharCount
        let retained =
            viewportPaintItemsCacheWithoutModel cache.Model
            |> List.choose (fun candidate ->
                let candidateItems = candidate.TotalItemCount
                let candidateItemTextChars = candidate.TotalItemTextCharCount
                let candidateSourceItems = candidate.SourceItemCount
                let candidateSourceTextChars = candidate.SourceTextCharCount
                if viewportPaintItemsCacheRetainable candidate
                   && keptCount < viewportPaintItemsCacheMaxModels
                   && keptItems + candidateItems <= viewportPaintItemsCacheMaxTotalItems
                   && keptItemTextChars + candidateItemTextChars <= viewportPaintItemsCacheMaxTotalItemTextChars
                   && keptSourceItems + candidateSourceItems <= viewportPaintItemsCacheMaxTotalSourceItems
                   && keptSourceTextChars + candidateSourceTextChars <= viewportPaintItemsCacheMaxTotalSourceTextChars then
                    keptCount <- keptCount + 1
                    keptItems <- keptItems + candidateItems
                    keptItemTextChars <- keptItemTextChars + candidateItemTextChars
                    keptSourceItems <- keptSourceItems + candidateSourceItems
                    keptSourceTextChars <- keptSourceTextChars + candidateSourceTextChars
                    Some candidate
                else
                    None)
        setViewportPaintItemsCache (cache :: retained)

let buildViewportPaintItems (model: SpreadsheetModel) (viewport: Viewport): ViewportPaintItem list =
    match tryGetViewportPaintItemsCache model with
    | ValueSome cached when viewportPaintItemsCacheRetainable cached ->
        match tryGetViewportPaintGeometryFastSlot cached viewport with
        | ValueSome entry -> entry.Items
        | ValueNone ->
            let key = viewportPaintGeometryKey viewport
            let mutable entry = Unchecked.defaultof<ViewportPaintItemsCacheEntry>
            if cached.Entries.TryGetValue(key, &entry) && sameMeasurer entry.Measurer viewport.Measurer then
                setViewportPaintGeometryFastSlot cached key entry
                entry.Items
            else
                let items = ResizeArray<ViewportPaintItem>(estimatedPaintItemCount model viewport)
                iterViewportPaintItemsUncached model viewport (fun zone item -> items.Add { Zone = zone; Item = item })
                let itemCount = items.Count
                let items = resizeArrayToList items
                if itemCount <= viewportPaintItemsCacheMaxTotalItems then
                    addViewportPaintItemsCacheEntry key {
                        Measurer = viewport.Measurer
                        Items = items
                        ItemCount = itemCount
                        ItemTextCharCount = viewportPaintItemsTextCharCount viewportPaintItemsCacheMaxTotalItemTextChars items
                    } cached
                else
                    removeViewportPaintItemsCacheEntry key cached
                rememberViewportPaintItemsCache cached
                items
    | ValueSome cached ->
        rememberViewportPaintItemsCache cached
        let items = ResizeArray<ViewportPaintItem>(estimatedPaintItemCount model viewport)
        iterViewportPaintItemsUncached model viewport (fun zone item -> items.Add { Zone = zone; Item = item })
        resizeArrayToList items
    | ValueNone ->
        let items = ResizeArray<ViewportPaintItem>(estimatedPaintItemCount model viewport)
        iterViewportPaintItemsUncached model viewport (fun zone item -> items.Add { Zone = zone; Item = item })
        let itemCount = items.Count
        let items = resizeArrayToList items
        if itemCount <= viewportPaintItemsCacheMaxTotalItems then
            let key = viewportPaintGeometryKey viewport
            let entries = Dictionary<ViewportPaintGeometryKey, ViewportPaintItemsCacheEntry>()
            let order = Queue<ViewportPaintGeometryKey>()
            let sourceItemCount, sourceTextCharCount = staticPaintSourceMetrics model
            let cache = {
                Model = model
                SourceItemCount = sourceItemCount
                SourceTextCharCount = sourceTextCharCount
                Entries = entries
                Order = order
                EntryFastSlots = Array.zeroCreate<ViewportPaintGeometryFastEntry> 128
                TotalItemCount = 0
                TotalItemTextCharCount = 0
            }
            addViewportPaintItemsCacheEntry key {
                Measurer = viewport.Measurer
                Items = items
                ItemCount = itemCount
                ItemTextCharCount = viewportPaintItemsTextCharCount viewportPaintItemsCacheMaxTotalItemTextChars items
            } cache
            rememberViewportPaintItemsCache cache
        items

let iterViewportPaintItems (model: SpreadsheetModel) (viewport: Viewport) (emit: ViewportZone -> PaintItem -> unit) =
    for item in buildViewportPaintItems model viewport do
        emit item.Zone item.Item
