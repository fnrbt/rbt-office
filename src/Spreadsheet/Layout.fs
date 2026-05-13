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

let viewportClip (viewport: Viewport) zone =
    match zone with
    | BodyZone ->
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
        visibleCells model viewport
        |> List.tryPick (fun cell ->
            if contains (toViewportRect viewport BodyZone cell.Rect) x y then Some (Cell cell.Address) else None)
        |> Option.defaultValue Empty
