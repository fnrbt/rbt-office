module RedBat.Spreadsheet.Layout

open RedBat.Spreadsheet.Model

type Rect = {
    X: float
    Y: float
    Width: float
    Height: float
}

/// Measure a piece of text in pixels at a given font size/face. Optional;
/// when None, the cell painter approximates width as charCount × mdwPx.
/// Set by the native runner (real Renderer.Text.measureWidth) and by the
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

type VisibleRange = {
    StartRow: int
    EndRowExclusive: int
    StartCol: int
    EndColExclusive: int
}

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

let columnX (model: SpreadsheetModel) (viewport: Viewport) (col: int): float =
    let mutable x = viewport.RowHeaderWidth
    for c = 0 to col - 1 do
        x <- x + getColumnWidth c model
    x

let rowY (model: SpreadsheetModel) (viewport: Viewport) (row: int): float =
    let mutable y = viewport.HeaderHeight
    for r = 0 to row - 1 do
        y <- y + getRowHeight r model
    y

let visibleRange (model: SpreadsheetModel) (viewport: Viewport): VisibleRange =
    let mutable startCol = 0
    let mutable x = viewport.RowHeaderWidth
    while startCol < model.ColCount && x < viewport.ScrollX do
        x <- x + getColumnWidth startCol model
        if x < viewport.ScrollX then
            startCol <- startCol + 1

    let mutable endCol = startCol
    let mutable xEnd = x
    while endCol < model.ColCount && xEnd < viewport.ScrollX + viewport.Width do
        xEnd <- xEnd + getColumnWidth endCol model
        endCol <- endCol + 1

    let mutable startRow = 0
    let mutable y = viewport.HeaderHeight
    while startRow < model.RowCount && y < viewport.ScrollY do
        y <- y + getRowHeight startRow model
        if y < viewport.ScrollY then
            startRow <- startRow + 1

    let mutable endRow = startRow
    let mutable yEnd = y
    while endRow < model.RowCount && yEnd < viewport.ScrollY + viewport.Height do
        yEnd <- yEnd + getRowHeight endRow model
        endRow <- endRow + 1

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
        Width = [merge.StartCol .. merge.EndCol] |> List.sumBy (fun c -> getColumnWidth c model)
        Height = [merge.StartRow .. merge.EndRow] |> List.sumBy (fun r -> getRowHeight r model)
    }

let visibleCells (model: SpreadsheetModel) (viewport: Viewport): CellGeometry list =
    let range = visibleRange model viewport
    let intersectsVisibleRange (merge: MergedCell) =
        merge.EndRow >= range.StartRow &&
        merge.StartRow < range.EndRowExclusive &&
        merge.EndCol >= range.StartCol &&
        merge.StartCol < range.EndColExclusive

    [
        for merge in model.MergedCells do
            if intersectsVisibleRange merge then
                yield {
                    Address = { Row = merge.StartRow; Col = merge.StartCol }
                    Rect = mergedRect model viewport merge
                    IsMergedPrimary = true
                    Merge = Some merge
                }

        for row in range.StartRow .. range.EndRowExclusive - 1 do
            for col in range.StartCol .. range.EndColExclusive - 1 do
                let addr = { Row = row; Col = col }
                match findMergedCell addr model with
                | Some _ -> ()
                | None ->
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
    [ 0 .. model.FrozenCols - 1 ] |> List.sumBy (fun c -> getColumnWidth c model)

/// Height in CSS px of the frozen-rows band (zero when FrozenRows = 0).
let frozenRowsHeight (model: SpreadsheetModel) : float =
    [ 0 .. model.FrozenRows - 1 ] |> List.sumBy (fun r -> getRowHeight r model)

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

let hitTest (model: SpreadsheetModel) (viewport: Viewport) x y =
    if contains { X = 0.0; Y = 0.0; Width = viewport.RowHeaderWidth; Height = viewport.HeaderHeight } x y then
        Corner
    elif y < viewport.HeaderHeight then
        headers model viewport
        |> List.tryPick (function
            | ColumnHeader (col, rect) when contains (toViewportRect viewport ColumnHeaderZone rect) x y -> Some (HitColumnHeader col)
            | _ -> None)
        |> Option.defaultValue Empty
    elif x < viewport.RowHeaderWidth then
        headers model viewport
        |> List.tryPick (function
            | RowHeader (row, rect) when contains (toViewportRect viewport RowHeaderZone rect) x y -> Some (HitRowHeader row)
            | _ -> None)
        |> Option.defaultValue Empty
    else
        // Frozen cells sit on top of body cells visually; test them first
        // and use each cell's own zone translation so clicks on a frozen
        // band don't fall through to whatever body cell happens to sit
        // under it at the current scroll position.
        visibleCells model viewport
        |> List.tryPick (fun cell ->
            let zone = zoneForCell model cell.Address.Row cell.Address.Col
            if contains (toViewportRect viewport zone cell.Rect) x y then Some (Cell cell.Address) else None)
        |> Option.defaultValue Empty
