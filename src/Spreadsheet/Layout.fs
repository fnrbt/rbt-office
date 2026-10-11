module Rbt.Spreadsheet.Layout

open System.Collections.Generic
open Rbt.Spreadsheet.Model

[<Struct>]
type Rect = {
    X: float
    Y: float
    Width: float
    Height: float
}

/// Measure a piece of text in pixels at a given font size/face. Optional;
/// when None, the cell painter approximates width as charCount × mdwPx.
/// Set by the native runner (real Rbt.Renderer.Text.measureWidth) and by the
/// PDF emitter; left None by the Fable web canvas which doesn't have a
/// portable measurement API at this layer.
type TextMeasurer = string -> float -> string option -> bool -> bool -> float

type Viewport = {
    ScrollX: float
    ScrollY: float
    Width: float
    Height: float
    HeaderHeight: float
    RowHeaderWidth: float
    /// Optional real-font measurer used by buildCellPaintItems to decide
    /// wrap-text break points. When None, falls back to mdwPx × charCount.
    Measurer: TextMeasurer option
}

[<Struct>]
type VisibleRange = {
    StartRow: int
    EndRowExclusive: int
    StartCol: int
    EndColExclusive: int
}

[<Struct>]
type CellGeometry = {
    Address: CellAddress
    Rect: Rect
    IsMergedPrimary: bool
    Merge: MergedCell option
}

type ViewportZone =
    | BodyZone
    | ColumnHeaderZone
    | RowHeaderZone
    | CornerZone
    /// Cells in the top frozen-rows band (above the freeze split, to the
    /// right of any frozen columns). Scrolls horizontally with the body
    /// but stays pinned vertically.
    | FrozenRowsZone
    /// Cells in the left frozen-cols band. Scrolls vertically with the
    /// body but stays pinned horizontally.
    | FrozenColsZone
    /// Top-left corner: the intersection of frozen rows and frozen cols.
    /// Pinned in both axes (no scroll applied).
    | FrozenCornerZone

type HeaderGeometry =
    | ColumnHeader of col: int * rect: Rect
    | RowHeader of row: int * rect: Rect
    | CornerHeader of rect: Rect

type HitTarget =
    | Cell of CellAddress
    | HitColumnHeader of int
    | HitRowHeader of int
    | Corner
    | Empty

let defaultViewport width height = {
    ScrollX = 0.0
    ScrollY = 0.0
    Width = width
    Height = height
    HeaderHeight = 30.0
    RowHeaderWidth = 50.0
    Measurer = None
}

/// The offsets of the rows (or columns) of a sheet: `offsets.[i]` is where
/// row `i` starts - the sum of the sizes of the rows before it -, and
/// `offsets.[count]` where the sheet ends. Only the rows up to the last one
/// with a size of its own (or hidden by a filter) have an entry in memory;
/// every row past it has the default size, and its offset is worked out. A
/// sheet of 1 048 576 rows whose caret is in the last one costs what its
/// sized rows cost.
[<Sealed>]
type AxisOffsets(known: float array, count: int, size: float) =
    let last = known.Length - 1
    let lastOffset = known.[last]
    /// The offset of row `index` (0 .. count).
    member _.Item with get (index: int) : float = if index <= last then known.[index] else lastOffset + float (index - last) * size
    /// One more than the number of rows, as the array this stands for had it.
    member _.Length = count + 1
    member _.Count = count
    /// The offsets that are held in memory (the first `Known.Length - 1` rows).
    member _.Known = known
    /// The size of every row past the known ones.
    member _.DefaultSize = size
    member this.GetSlice(first: int option, last: int option) : float array =
        let first = defaultArg first 0
        let last = defaultArg last count
        Array.init (max 0 (last - first + 1)) (fun i -> this.[first + i])
    /// The first index in 0 .. count whose offset is not below `target`.
    member this.LowerBound(target: float) : int =
        let mutable lo = 0
        let mutable hi = count + 1
        while lo < hi do
            let mid = lo + ((hi - lo) / 2)
            if this.[mid] < target then lo <- mid + 1 else hi <- mid
        lo

let private buildPrefixSizeCache count sizeAt =
    let prefix = Array.zeroCreate<float> (count + 1)
    for i = 0 to count - 1 do
        prefix.[i + 1] <- prefix.[i] + max 0.0 (sizeAt i)
    prefix

/// The index whose span holds `offset` in a prefix-sum array (see
/// `SheetMetrics`), clamped to the first / last index.
let indexAtOffset (prefix: AxisOffsets) count offset =
    let target = max 0.0 offset
    let mutable lo = 0
    let mutable hi = count
    while lo < hi do
        let mid = lo + ((hi - lo) / 2)
        if prefix.[mid + 1] <= target then lo <- mid + 1 else hi <- mid
    min lo (max 0 (count - 1))

let private endIndexAtOffset (prefix: AxisOffsets) count offset =
    min count (prefix.LowerBound (max 0.0 offset))

let prefixSizeUncached sizeAt index =
    let mutable total = 0.0
    for i = 0 to index - 1 do
        total <- total + max 0.0 (sizeAt i)
    total

/// Row and column offsets of a sheet: `RowPrefix.[r]` is the top of row `r`
/// (the sum of the heights of the rows above it), `ColPrefix.[c]` the left
/// edge of column `c`; both arrays have one entry more than the sheet has
/// rows / columns (the far edge). Offsets are in sheet pixels, from the
/// top-left corner of A1.
type SheetMetrics = {
    RowPrefix: AxisOffsets
    ColPrefix: AxisOffsets
    RowCount: int
    ColCount: int
    /// What the metrics were computed from: the same maps (by identity) and
    /// counts give the same offsets.
    RowHeights: Map<int, float>
    ColumnWidths: Map<int, float>
    FilterHiddenRows: int array
    Filtered: bool
}

/// The metrics computed last, for a few sheets (a workbook is drawn sheet by
/// sheet; hit tests and paints of one sheet come in runs).
let mutable private metricsCache : SheetMetrics list = []

let private metricsMatch (model: SpreadsheetModel) (metrics: SheetMetrics) =
    metrics.RowCount = model.RowCount
    && metrics.ColCount = model.ColCount
    && obj.ReferenceEquals(metrics.RowHeights, model.RowHeights)
    && obj.ReferenceEquals(metrics.ColumnWidths, model.ColumnWidths)
    && metrics.Filtered = not (Map.isEmpty model.Filters)
    && obj.ReferenceEquals(metrics.FilterHiddenRows, model.FilterHiddenRows)

let private computeMetrics (model: SpreadsheetModel) : SheetMetrics =
    // The rows / columns that have a size of their own, and nothing for the
    // ones past the last of them: one pass over the maps, no entry per row
    // of the sheet.
    let rowCount = max 0 model.RowCount
    let colCount = max 0 model.ColCount
    let filtered = not (Map.isEmpty model.Filters)
    let offsetsOf (count: int) (defaultSize: float) (sizes: Map<int, float>) (hidden: int array) =
        let mutable known = 0
        for KeyValue(index, _) in sizes do
            if index >= 0 && index < count && index + 1 > known then known <- index + 1
        for index in hidden do
            if index >= 0 && index < count && index + 1 > known then known <- index + 1
        let own = Array.create known defaultSize
        for KeyValue(index, size) in sizes do
            if index >= 0 && index < known then own.[index] <- max 0.0 size
        for index in hidden do
            if index >= 0 && index < known then own.[index] <- 0.0
        let prefix = Array.zeroCreate<float> (known + 1)
        for i = 0 to known - 1 do
            prefix.[i + 1] <- prefix.[i] + own.[i]
        AxisOffsets(prefix, count, defaultSize)
    { RowPrefix = offsetsOf rowCount 25.0 model.RowHeights (if filtered then model.FilterHiddenRows else [||])
      ColPrefix = offsetsOf colCount 100.0 model.ColumnWidths [||]
      RowCount = rowCount
      ColCount = colCount
      RowHeights = model.RowHeights
      ColumnWidths = model.ColumnWidths
      FilterHiddenRows = model.FilterHiddenRows
      Filtered = filtered }

/// The row / column offsets of a sheet. They are computed once and reused
/// until a row height, a column width, the filter or the size of the sheet
/// changes (any of which gives the model a different map or count).
let metricsOf (model: SpreadsheetModel) : SheetMetrics =
    let cached = metricsCache
    match cached |> List.tryFind (metricsMatch model) with
    | Some metrics -> metrics
    | None ->
        let metrics = computeMetrics model
        metricsCache <- metrics :: List.truncate 7 cached
        metrics

let rowTop (model: SpreadsheetModel) (row: int): float =
    let metrics = metricsOf model
    if row <= 0 then 0.0
    elif row <= metrics.RowCount then metrics.RowPrefix.[row]
    // Past the end of the sheet: the rows there have their own or the default height.
    else metrics.RowPrefix.[metrics.RowCount] + prefixSizeUncached (fun r -> getRowHeight (metrics.RowCount + r) model) (row - metrics.RowCount)

let columnLeft (model: SpreadsheetModel) (col: int): float =
    let metrics = metricsOf model
    if col <= 0 then 0.0
    elif col <= metrics.ColCount then metrics.ColPrefix.[col]
    else metrics.ColPrefix.[metrics.ColCount] + prefixSizeUncached (fun c -> getColumnWidth (metrics.ColCount + c) model) (col - metrics.ColCount)

/// What is drawn on the edge two cells share.
type SharedEdge =
    /// Inside a merged cell: nothing.
    | EdgeInside
    /// No border: a grid line, unless the sheet shows none or a filled cell hides it.
    | EdgeGrid of shown: bool
    /// The one border of the edge (see `Model.resolveSharedBorder`).
    | EdgeBorder of BorderStyle

let private visibleBefore (prefix: AxisOffsets) (index: int) =
    let mutable at = index - 1
    while at >= 0 && prefix.[at + 1] <= prefix.[at] do
        at <- at - 1
    at

let private visibleAfter (prefix: AxisOffsets) (count: int) (index: int) =
    let mutable at = index + 1
    while at < count && prefix.[at + 1] <= prefix.[at] do
        at <- at + 1
    if at < count then at else -1

let private borderFormatAt (model: SpreadsheetModel) (row: int) (col: int) : CellFormat option =
    match Map.tryFind { Row = row; Col = col } model.Cells with
    | Some cell -> Some cell.Format
    | None -> None

let private mergeAt (model: SpreadsheetModel) (row: int) (col: int) : MergedCell option =
    if model.MergedCells.IsEmpty then None else findMergedCell { Row = row; Col = col } model

/// A cell's fill hides the grid lines around it (a merged cell's is its first cell's).
let private filledAt (model: SpreadsheetModel) (row: int) (col: int) =
    let format =
        match mergeAt model row col with
        | Some merge -> borderFormatAt model merge.StartRow merge.StartCol
        | None -> borderFormatAt model row col
    match format with
    | Some format -> format.BgColor.IsSome
    | None -> false

let private sharedEdge (model: SpreadsheetModel) (firstRow: int) (firstCol: int) (secondRow: int) (secondCol: int) (horizontal: bool) : SharedEdge =
    let hasFirst = firstRow >= 0 && firstCol >= 0
    let hasSecond = secondRow >= 0 && secondCol >= 0
    let inside =
        hasFirst && hasSecond
        && (match mergeAt model firstRow firstCol, mergeAt model secondRow secondCol with
            | Some a, Some b -> a.StartRow = b.StartRow && a.StartCol = b.StartCol
            | _ -> false)
    if inside then EdgeInside
    else
        let side (present: bool) (row: int) (col: int) (pick: CellFormat -> BorderStyle option) =
            if present then borderFormatAt model row col |> Option.bind pick else None
        let first = side hasFirst firstRow firstCol (fun format -> if horizontal then format.BorderBottom else format.BorderRight)
        let second = side hasSecond secondRow secondCol (fun format -> if horizontal then format.BorderTop else format.BorderLeft)
        match resolveSharedBorder first second with
        | Some border -> EdgeBorder border
        | None ->
            EdgeGrid (model.ShowGridLines
                      && not (hasFirst && filledAt model firstRow firstCol)
                      && not (hasSecond && filledAt model secondRow secondCol))

/// The edge above a cell: between it and the cell in the nearest row above
/// that is shown (hidden rows have no part in it).
let edgeAbove (model: SpreadsheetModel) (row: int) (col: int) : SharedEdge =
    let metrics = metricsOf model
    let above = if row <= metrics.RowCount then visibleBefore metrics.RowPrefix (min row metrics.RowCount) else row - 1
    sharedEdge model above (if above >= 0 then col else -1) row col true

/// The edge to the left of a cell.
let edgeLeft (model: SpreadsheetModel) (row: int) (col: int) : SharedEdge =
    let metrics = metricsOf model
    let left = if col <= metrics.ColCount then visibleBefore metrics.ColPrefix (min col metrics.ColCount) else col - 1
    sharedEdge model (if left >= 0 then row else -1) left row col false

/// The edge below a cell: the edge above the next row shown, or the cell's own bottom border in the last one.
let edgeBelow (model: SpreadsheetModel) (row: int) (col: int) : SharedEdge =
    let metrics = metricsOf model
    let below = if row < metrics.RowCount then visibleAfter metrics.RowPrefix metrics.RowCount row else -1
    if below >= 0 then edgeAbove model below col else sharedEdge model row col -1 -1 true

/// The edge to the right of a cell.
let edgeRight (model: SpreadsheetModel) (row: int) (col: int) : SharedEdge =
    let metrics = metricsOf model
    let right = if col < metrics.ColCount then visibleAfter metrics.ColPrefix metrics.ColCount col else -1
    if right >= 0 then edgeLeft model row right else sharedEdge model row col -1 -1 false

/// The row shown above a row, the column shown left of a column (-1: none).
let rowShownBefore (model: SpreadsheetModel) (row: int) =
    let metrics = metricsOf model
    if row <= metrics.RowCount then visibleBefore metrics.RowPrefix row else row - 1

let columnShownBefore (model: SpreadsheetModel) (col: int) =
    let metrics = metricsOf model
    if col <= metrics.ColCount then visibleBefore metrics.ColPrefix col else col - 1

/// Where a drawing object sits on the sheet, in sheet pixels (origin at the
/// top-left corner of A1, headers not included).
let drawingRect (model: SpreadsheetModel) (drawing: SheetDrawing) : Rect =
    let anchor = drawing.Anchor
    let width = anchor.Width |> Option.defaultValue 0.0
    let height = anchor.Height |> Option.defaultValue 0.0
    if anchor.Kind = "absolute" then { X = anchor.X; Y = anchor.Y; Width = width; Height = height }
    else
        let left = columnLeft model anchor.FromCol + anchor.FromOffsetX
        let top = rowTop model anchor.FromRow + anchor.FromOffsetY
        match anchor.Kind, anchor.ToCol, anchor.ToRow with
        | "twoCell", Some toCol, Some toRow ->
            let right = columnLeft model toCol + anchor.ToOffsetX
            let bottom = rowTop model toRow + anchor.ToOffsetY
            // Hidden rows / columns can collapse the span; the object's own size is the fallback.
            { X = left; Y = top
              Width = (if right - left >= 1.0 then right - left else width)
              Height = (if bottom - top >= 1.0 then bottom - top else height) }
        | _ -> { X = left; Y = top; Width = width; Height = height }

/// What a placeholder for a kept drawing object says ("Chart: Sales by region").
let drawingLabel (drawing: SheetDrawing) : string =
    let kind =
        match drawing.Content with
        | Picture _ -> "Picture"
        | KeptObject ("chart", _) -> "Chart"
        | KeptObject ("shape", _) -> "Shape"
        | KeptObject _ -> "Object"
    match drawing.Name with
    | Some name when name.Trim() <> "" -> kind + ": " + name
    | _ -> kind

let columnX (model: SpreadsheetModel) (viewport: Viewport) (col: int): float =
    viewport.RowHeaderWidth + columnLeft model col

let rowY (model: SpreadsheetModel) (viewport: Viewport) (row: int): float =
    viewport.HeaderHeight + rowTop model row

let visibleRange (model: SpreadsheetModel) (viewport: Viewport): VisibleRange =
    let metrics = metricsOf model
    let colPrefix = metrics.ColPrefix
    let rowPrefix = metrics.RowPrefix

    let startColOffset = max 0.0 (viewport.ScrollX - viewport.RowHeaderWidth)
    let endColOffset = max 0.0 (viewport.ScrollX + viewport.Width - viewport.RowHeaderWidth)
    let startRowOffset = max 0.0 (viewport.ScrollY - viewport.HeaderHeight)
    let endRowOffset = max 0.0 (viewport.ScrollY + viewport.Height - viewport.HeaderHeight)

    let startCol = if model.ColCount = 0 then 0 else indexAtOffset colPrefix model.ColCount startColOffset
    let endCol = endIndexAtOffset colPrefix model.ColCount endColOffset
    let startRow = if model.RowCount = 0 then 0 else indexAtOffset rowPrefix model.RowCount startRowOffset
    let endRow = endIndexAtOffset rowPrefix model.RowCount endRowOffset

    {
        StartRow = startRow
        EndRowExclusive = endRow
        StartCol = startCol
        EndColExclusive = endCol
    }

let cellRect (model: SpreadsheetModel) (viewport: Viewport) (row: int) (col: int): Rect =
    {
        X = columnX model viewport col
        Y = rowY model viewport row
        Width = getColumnWidth col model
        Height = getRowHeight row model
    }

let mergedRect (model: SpreadsheetModel) (viewport: Viewport) (merge: MergedCell): Rect =
    {
        X = columnX model viewport merge.StartCol
        Y = rowY model viewport merge.StartRow
        Width = columnLeft model (merge.EndCol + 1) - columnLeft model merge.StartCol
        Height = rowTop model (merge.EndRow + 1) - rowTop model merge.StartRow
    }

let visibleCells (model: SpreadsheetModel) (viewport: Viewport): CellGeometry list =
    let range = visibleRange model viewport
    let merges =
        mergedCellsIntersectingRange range.StartRow range.EndRowExclusive range.StartCol range.EndColExclusive model
    let covered = HashSet<CellAddress>()
    for merge in merges do
        for row = merge.StartRow to merge.EndRow do
            for col = merge.StartCol to merge.EndCol do
                covered.Add { Row = row; Col = col } |> ignore

    [
        for merge in merges do
            yield {
                Address = { Row = merge.StartRow; Col = merge.StartCol }
                Rect = mergedRect model viewport merge
                IsMergedPrimary = true
                Merge = Some merge
            }

        for row in range.StartRow .. range.EndRowExclusive - 1 do
            for col in range.StartCol .. range.EndColExclusive - 1 do
                let addr = { Row = row; Col = col }
                if not (covered.Contains addr) then
                    yield {
                        Address = addr
                        Rect = cellRect model viewport row col
                        IsMergedPrimary = false
                        Merge = None
                    }
    ]

let headers (model: SpreadsheetModel) (viewport: Viewport): HeaderGeometry list =
    let range = visibleRange model viewport
    [
        yield CornerHeader {
            X = 0.0
            Y = 0.0
            Width = viewport.RowHeaderWidth
            Height = viewport.HeaderHeight
        }
        for col in range.StartCol .. range.EndColExclusive - 1 do
            yield ColumnHeader (col, {
                X = columnX model viewport col
                Y = 0.0
                Width = getColumnWidth col model
                Height = viewport.HeaderHeight
            })
        for row in range.StartRow .. range.EndRowExclusive - 1 do
            yield RowHeader (row, {
                X = 0.0
                Y = rowY model viewport row
                Width = viewport.RowHeaderWidth
                Height = getRowHeight row model
            })
    ]

// ---- text overflow ------------------------------------------------------------
// Excel lets text that is wider than its cell run over the empty cells next
// to it: to the right for left-aligned text, to the left for right-aligned,
// to both sides for centred text. It stops at the first cell that holds
// something. Only text does this - a number that does not fit shows ####.

/// The sides a cell's text may overflow to: (left, right), or None when it
/// stays in its cell (not text, wrapped, shrunk to fit, rotated, or aligned
/// in a way that fills the cell).
let textOverflowSides (format: CellFormat) (value: CellValue) : (bool * bool) option =
    if format.WrapText || format.ShrinkToFit || format.TextRotation <> 0 then None
    else
        match unwrapFormulaValue value with
        | CellValue.String text when text <> "" ->
            match format.HorizontalAlign |> Option.map (fun align -> align.ToLowerInvariant()) with
            | None | Some "left" | Some "general" | Some "" -> Some (false, true)
            | Some "right" -> Some (true, false)
            | Some "center" -> Some (true, true)
            | _ -> None
        | _ -> None

/// Whether a cell stops a neighbour's text: it holds a value (a formula
/// counts, whatever it gives) or is part of a merged cell.
let private blocksTextOverflow (model: SpreadsheetModel) (row: int) (col: int) =
    let addr = { Row = row; Col = col }
    (match getCell addr model with
     | Some data -> (match data.Value with CellValue.Empty | CellValue.String "" -> false | _ -> true)
     | None -> false)
    || (not model.MergedCells.IsEmpty && (findMergedCell addr model).IsSome)

/// The columns the text of the cell (row, col) may be drawn over when it
/// needs `needLeft` / `needRight` more pixels than its cell has on that
/// side: (first column, last column). Hidden columns are passed over.
let textOverflowSpan (model: SpreadsheetModel) (row: int) (col: int) (needLeft: float) (needRight: float) : int * int =
    let mutable first = col
    let mutable last = col
    let mutable need = needRight
    let mutable next = col + 1
    while need > 0.0 && next < model.ColCount && not (blocksTextOverflow model row next) do
        last <- next
        need <- need - getColumnWidth next model
        next <- next + 1
    need <- needLeft
    next <- col - 1
    while need > 0.0 && next >= 0 && not (blocksTextOverflow model row next) do
        first <- next
        need <- need - getColumnWidth next model
        next <- next - 1
    first, last

let contains (rect: Rect) x y =
    x >= rect.X && x < rect.X + rect.Width &&
    y >= rect.Y && y < rect.Y + rect.Height

let translateRect dx dy (rect: Rect): Rect =
    { rect with X = rect.X + dx; Y = rect.Y + dy }

let toViewportRect (viewport: Viewport) zone (rect: Rect): Rect =
    match zone with
    | BodyZone -> rect |> translateRect -viewport.ScrollX -viewport.ScrollY
    | ColumnHeaderZone -> rect |> translateRect -viewport.ScrollX 0.0
    | RowHeaderZone -> rect |> translateRect 0.0 -viewport.ScrollY
    | CornerZone -> rect
    // Frozen zones: pinned axes don't translate, free axis tracks scroll.
    | FrozenRowsZone -> rect |> translateRect -viewport.ScrollX 0.0
    | FrozenColsZone -> rect |> translateRect 0.0 -viewport.ScrollY
    | FrozenCornerZone -> rect

/// Width in CSS px of the frozen-cols band (zero when FrozenCols = 0).
let frozenColsWidth (model: SpreadsheetModel) : float =
    columnLeft model model.FrozenCols

/// Height in CSS px of the frozen-rows band (zero when FrozenRows = 0).
let frozenRowsHeight (model: SpreadsheetModel) : float =
    rowTop model model.FrozenRows

/// X coordinate (in viewport space) where the freeze split sits — that is,
/// the right edge of the frozen-cols band / start of horizontally-scrolling
/// body. Returns None when there are no frozen cols.
let frozenSplitX (model: SpreadsheetModel) (viewport: Viewport) : float option =
    if model.FrozenCols <= 0 then None
    else Some (viewport.RowHeaderWidth + frozenColsWidth model)

/// Y coordinate where the freeze split sits — bottom edge of frozen-rows.
let frozenSplitY (model: SpreadsheetModel) (viewport: Viewport) : float option =
    if model.FrozenRows <= 0 then None
    else Some (viewport.HeaderHeight + frozenRowsHeight model)

/// Decide which zone a body cell belongs to based on its row/col vs the
/// freeze split. Frozen-cell zones always paint (regardless of scroll
/// position) because they sit on top of the scrolling body.
let zoneForCell (model: SpreadsheetModel) (row: int) (col: int) : ViewportZone =
    let inFrozenRows = model.FrozenRows > 0 && row < model.FrozenRows
    let inFrozenCols = model.FrozenCols > 0 && col < model.FrozenCols
    match inFrozenRows, inFrozenCols with
    | true, true -> FrozenCornerZone
    | true, false -> FrozenRowsZone
    | false, true -> FrozenColsZone
    | false, false -> BodyZone

/// Clip rect for a zone. Callers pass the model in so frozen-pane zones
/// can derive their split positions from FrozenRows/FrozenCols.
let viewportClipFor (model: SpreadsheetModel) (viewport: Viewport) zone =
    let fx = frozenColsWidth model
    let fy = frozenRowsHeight model
    let bodyX = viewport.RowHeaderWidth + fx
    let bodyY = viewport.HeaderHeight + fy
    match zone with
    | BodyZone ->
        { X = bodyX
          Y = bodyY
          Width = max 0.0 (viewport.Width - bodyX)
          Height = max 0.0 (viewport.Height - bodyY) }
    | ColumnHeaderZone ->
        { X = viewport.RowHeaderWidth
          Y = 0.0
          Width = max 0.0 (viewport.Width - viewport.RowHeaderWidth)
          Height = viewport.HeaderHeight }
    | RowHeaderZone ->
        { X = 0.0
          Y = viewport.HeaderHeight
          Width = viewport.RowHeaderWidth
          Height = max 0.0 (viewport.Height - viewport.HeaderHeight) }
    | CornerZone ->
        { X = 0.0
          Y = 0.0
          Width = viewport.RowHeaderWidth
          Height = viewport.HeaderHeight }
    | FrozenRowsZone ->
        { X = bodyX
          Y = viewport.HeaderHeight
          Width = max 0.0 (viewport.Width - bodyX)
          Height = fy }
    | FrozenColsZone ->
        { X = viewport.RowHeaderWidth
          Y = bodyY
          Width = fx
          Height = max 0.0 (viewport.Height - bodyY) }
    | FrozenCornerZone ->
        { X = viewport.RowHeaderWidth
          Y = viewport.HeaderHeight
          Width = fx
          Height = fy }

/// Backward-compat shim: callers that don't have a model still get the
/// pre-freeze clip rectangles. Internal callers should prefer
/// `viewportClipFor` so frozen panes clip correctly.
let viewportClip (viewport: Viewport) zone =
    match zone with
    | BodyZone | FrozenRowsZone | FrozenColsZone | FrozenCornerZone ->
        { X = viewport.RowHeaderWidth
          Y = viewport.HeaderHeight
          Width = max 0.0 (viewport.Width - viewport.RowHeaderWidth)
          Height = max 0.0 (viewport.Height - viewport.HeaderHeight) }
    | ColumnHeaderZone ->
        { X = viewport.RowHeaderWidth
          Y = 0.0
          Width = max 0.0 (viewport.Width - viewport.RowHeaderWidth)
          Height = viewport.HeaderHeight }
    | RowHeaderZone ->
        { X = 0.0
          Y = viewport.HeaderHeight
          Width = viewport.RowHeaderWidth
          Height = max 0.0 (viewport.Height - viewport.HeaderHeight) }
    | CornerZone ->
        { X = 0.0
          Y = 0.0
          Width = viewport.RowHeaderWidth
          Height = viewport.HeaderHeight }

let intersectRect (a: Rect) (b: Rect): Rect option =
    let x1 = max a.X b.X
    let y1 = max a.Y b.Y
    let x2 = min (a.X + a.Width) (b.X + b.Width)
    let y2 = min (a.Y + a.Height) (b.Y + b.Height)
    if x2 <= x1 || y2 <= y1 then
        None
    else
        Some { X = x1; Y = y1; Width = x2 - x1; Height = y2 - y1 }

let private columnAtOffset (model: SpreadsheetModel) offset =
    if offset < 0.0 || model.ColCount <= 0 then None
    else
        let prefix = (metricsOf model).ColPrefix
        let col = indexAtOffset prefix model.ColCount offset
        if col >= 0 && col < model.ColCount && prefix.[col] <= offset && offset < prefix.[col + 1] then Some col else None

let private rowAtOffset (model: SpreadsheetModel) offset =
    if offset < 0.0 || model.RowCount <= 0 then None
    else
        let prefix = (metricsOf model).RowPrefix
        let row = indexAtOffset prefix model.RowCount offset
        if row >= 0 && row < model.RowCount && prefix.[row] <= offset && offset < prefix.[row + 1] then Some row else None

/// The column whose span holds the sheet offset `offset` (pixels from the
/// left edge of column A), when there is one.
let columnAt (model: SpreadsheetModel) (offset: float) : int option = columnAtOffset model offset

/// The row whose span holds the sheet offset `offset` (pixels from the top of row 1).
let rowAt (model: SpreadsheetModel) (offset: float) : int option = rowAtOffset model offset

let private primaryAddressFor row col model =
    match findMergedCell { Row = row; Col = col } model with
    | Some merge -> { Row = merge.StartRow; Col = merge.StartCol }
    | None -> { Row = row; Col = col }

let hitTest (model: SpreadsheetModel) (viewport: Viewport) x y =
    if contains { X = 0.0; Y = 0.0; Width = viewport.RowHeaderWidth; Height = viewport.HeaderHeight } x y then
        Corner
    elif y < viewport.HeaderHeight then
        let frozenWidth = frozenColsWidth model
        let offset =
            if model.FrozenCols > 0 && x < viewport.RowHeaderWidth + frozenWidth then
                x - viewport.RowHeaderWidth
            else
                x + viewport.ScrollX - viewport.RowHeaderWidth
        match columnAtOffset model offset with
        | Some col -> HitColumnHeader col
        | None -> Empty
    elif x < viewport.RowHeaderWidth then
        let frozenHeight = frozenRowsHeight model
        let offset =
            if model.FrozenRows > 0 && y < viewport.HeaderHeight + frozenHeight then
                y - viewport.HeaderHeight
            else
                y + viewport.ScrollY - viewport.HeaderHeight
        match rowAtOffset model offset with
        | Some row -> HitRowHeader row
        | None -> Empty
    else
        let frozenWidth = frozenColsWidth model
        let frozenHeight = frozenRowsHeight model
        let colOffset =
            if model.FrozenCols > 0 && x < viewport.RowHeaderWidth + frozenWidth then
                x - viewport.RowHeaderWidth
            else
                x + viewport.ScrollX - viewport.RowHeaderWidth
        let rowOffset =
            if model.FrozenRows > 0 && y < viewport.HeaderHeight + frozenHeight then
                y - viewport.HeaderHeight
            else
                y + viewport.ScrollY - viewport.HeaderHeight
        match rowAtOffset model rowOffset, columnAtOffset model colOffset with
        | Some row, Some col -> Cell (primaryAddressFor row col model)
        | _ -> Empty
