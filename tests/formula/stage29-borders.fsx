#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model

// Cell borders: every Excel border style has a complete CSS declaration
// (width, line style, colour), and ApplyBorder's kinds ("all", "outside",
// "inside", "none", one side) change the right sides of the right cells.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-34s expected %A  got %A" label expected got)

let css style = borderCss { Style = style; Color = Some "#112233" }
eq "thin" (css "thin") "1px solid #112233"
eq "medium" (css "medium") "2px solid #112233"
eq "thick" (css "thick") "3px solid #112233"
eq "dashed" (css "dashed") "1px dashed #112233"
eq "dotted" (css "dotted") "1px dotted #112233"
eq "double" (css "double") "3px double #112233"
eq "hair" (css "hair") "1px dotted #112233"
eq "mediumDashed" (css "mediumDashed") "2px dashed #112233"
eq "dashDot" (css "dashDot") "1px dashed #112233"
eq "mediumDashDot" (css "mediumDashDot") "2px dashed #112233"
eq "dashDotDot" (css "dashDotDot") "1px dashed #112233"
eq "mediumDashDotDot" (css "mediumDashDotDot") "2px dashed #112233"
eq "slantDashDot" (css "slantDashDot") "2px dashed #112233"
eq "no colour is black" (borderCss { Style = "thin"; Color = None }) "1px solid #000000"
eq "none" (css "none") "none"
eq "every style has a line style"
    ([ "thin"; "medium"; "thick"; "dashed"; "dotted"; "double"; "hair"; "mediumDashed"; "dashDot"; "mediumDashDot"; "dashDotDot"; "mediumDashDotDot"; "slantDashDot" ]
     |> List.forall (fun style -> match borderCssParts style with Some (width, line) -> width >= 1 && line <> "" | None -> false)) true

// Sides of the cells of the block B2:D4 (rows 1-3, columns 1-3).
let block = 1, 1, 3, 3
let sides (format: CellFormat) =
    [ if format.BorderTop.IsSome then "T"
      if format.BorderBottom.IsSome then "B"
      if format.BorderLeft.IsSome then "L"
      if format.BorderRight.IsSome then "R" ] |> String.concat ""
let at side style row col = borderFormatInBlock block side style "#000000" row col defaultFormat |> sides
eq "all" (at "all" "thin" 2 2) "TBLR"
eq "outside corner" (at "outside" "thin" 1 1) "TL"
eq "outside top edge" (at "outside" "thin" 1 2) "T"
eq "outside middle" (at "outside" "thin" 2 2) ""
eq "outside far corner" (at "outside" "thin" 3 3) "BR"
eq "inside corner" (at "inside" "thin" 1 1) "BR"
eq "inside middle" (at "inside" "thin" 2 2) "TBLR"
eq "inside far corner" (at "inside" "thin" 3 3) "TL"
eq "one side" (at "left" "medium" 2 2) "L"
let boxed = borderFormatInBlock block "all" "thick" "#ff0000" 2 2 defaultFormat
eq "style and colour stored" boxed.BorderTop (Some { Style = "thick"; Color = Some "#ff0000" })
eq "none clears" (borderFormatInBlock block "none" "thin" "#000000" 2 2 boxed |> sides) ""
eq "a style of none clears that side" (borderFormatInBlock block "top" "none" "#000000" 2 2 boxed |> sides) "BLR"
eq "nothing to do keeps the format" (obj.ReferenceEquals(borderFormatInBlock block "outside" "thin" "#000000" 2 2 boxed, boxed)) true
eq "one cell: outside is its box" (borderFormatInBlock (4, 4, 4, 4) "outside" "thin" "#000000" 4 4 defaultFormat |> sides) "TBLR"
eq "one cell: inside is nothing" (borderFormatInBlock (4, 4, 4, 4) "inside" "thin" "#000000" 4 4 defaultFormat |> sides) ""

// ---- one line per edge ---------------------------------------------------------
// The edge two cells share has one border: the stronger of what the two cells
// state for it (Excel's order of styles, not the stroke width), the later
// cell's - below / to the right - when they are equally strong.
module Layout = Rbt.Spreadsheet.Layout
module Paint = Rbt.Spreadsheet.Paint
let line style color = Some { Style = style; Color = Some color }
eq "one side states a border" (resolveSharedBorder (line "thin" "#f00") None, resolveSharedBorder None (line "thin" "#00f")) (line "thin" "#f00", line "thin" "#00f")
eq "the stronger wins, whichever cell states it" (resolveSharedBorder (line "thick" "#f00") (line "thin" "#00f"), resolveSharedBorder (line "thin" "#f00") (line "medium" "#00f")) (line "thick" "#f00", line "medium" "#00f")
eq "double outranks thick, thin outranks dashed" (resolveSharedBorder (line "double" "#f00") (line "thick" "#00f"), resolveSharedBorder (line "dashed" "#f00") (line "thin" "#00f")) (line "double" "#f00", line "thin" "#00f")
eq "equally strong: the later cell's" (resolveSharedBorder (line "medium" "#f00") (line "medium" "#00f")) (line "medium" "#00f")
eq "a style of none is no border" (resolveSharedBorder (line "none" "#f00") (line "hair" "#00f"), resolveSharedBorder (line "none" "#f00") None) (line "hair" "#00f", None)

let cellWith (change: CellFormat -> CellFormat) = { Value = CellValue.Empty; Format = change defaultFormat; Comment = None }
let at2 row col : CellAddress = { Row = row; Col = col }
let sheet =
    createEmpty 20 10
    |> setCell (at2 2 2) (cellWith (fun f -> { f with BorderBottom = line "thin" "#f00"; BorderRight = line "thick" "#f00" }))
    |> setCell (at2 3 2) (cellWith (fun f -> { f with BorderTop = line "thick" "#00f" }))
    |> setCell (at2 2 3) (cellWith (fun f -> { f with BorderLeft = line "thin" "#00f" }))
    |> setCell (at2 6 6) (cellWith (fun f -> { f with BgColor = Some "#ffeeaa" }))
eq "the edge below a cell is the edge above the next" (Layout.edgeBelow sheet 2 2, Layout.edgeAbove sheet 3 2) (Layout.EdgeBorder (line "thick" "#00f").Value, Layout.EdgeBorder (line "thick" "#00f").Value)
eq "the edge to the right is the edge to the left of the next" (Layout.edgeRight sheet 2 2, Layout.edgeLeft sheet 2 3) (Layout.EdgeBorder (line "thick" "#f00").Value, Layout.EdgeBorder (line "thick" "#f00").Value)
eq "no border: a grid line" (Layout.edgeAbove sheet 10 5, Layout.edgeLeft sheet 10 5) (Layout.EdgeGrid true, Layout.EdgeGrid true)
eq "a fill hides the grid lines around its cell" (Layout.edgeAbove sheet 6 6, Layout.edgeLeft sheet 6 6, Layout.edgeBelow sheet 6 6, Layout.edgeRight sheet 6 6) (Layout.EdgeGrid false, Layout.EdgeGrid false, Layout.EdgeGrid false, Layout.EdgeGrid false)
eq "a sheet without grid lines shows none" (Layout.edgeAbove { sheet with ShowGridLines = false } 10 5) (Layout.EdgeGrid false)
// A hidden row is not part of an edge: the cells above and below it share one.
let hidden = { sheet with RowHeights = sheet.RowHeights |> Map.add 4 0.0 } |> setCell (at2 3 4) (cellWith (fun f -> { f with BorderBottom = line "medium" "#0a0" }))
eq "across a hidden row" (Layout.edgeAbove hidden 5 4) (Layout.EdgeBorder (line "medium" "#0a0").Value)
// Inside a merged cell there is no edge; its outline is that of the cells along it.
let merged =
    { sheet with MergedCells = [ { StartRow = 12; StartCol = 1; EndRow = 13; EndCol = 2 } ] }
    |> setCell (at2 12 1) (cellWith (fun f -> { f with BorderTop = line "thin" "#000"; BorderLeft = line "thin" "#000" }))
    |> setCell (at2 13 2) (cellWith (fun f -> { f with BorderBottom = line "medium" "#000"; BorderRight = line "medium" "#000" }))
eq "inside a merged cell" (Layout.edgeAbove merged 13 1, Layout.edgeLeft merged 12 2) (Layout.EdgeInside, Layout.EdgeInside)
eq "its outline" (Layout.edgeAbove merged 12 1, Layout.edgeBelow merged 13 2, Layout.edgeRight merged 13 2) (Layout.EdgeBorder (line "thin" "#000").Value, Layout.EdgeBorder (line "medium" "#000").Value, Layout.EdgeBorder (line "medium" "#000").Value)

// The portable painter: a 3 x 3 block with every border medium has one line
// on each of its 24 edges (4 lines of 3 edges each way), not one per cell side.
let blockSheet =
    [ for row in 1 .. 3 do for col in 1 .. 3 -> row, col ]
    |> List.fold (fun model (row, col) -> setCell (at2 row col) (cellWith (borderFormatInBlock (1, 1, 3, 3) "all" "medium" "#000000" row col)) model) (createEmpty 20 10)
let borderLines (model: SpreadsheetModel) =
    Paint.buildPaintItems model (Layout.defaultViewport 800.0 500.0)
    |> List.choose (function
        | Paint.GridLine (x1, y1, x2, y2, color, width) when color = "#000000" -> Some (x1, y1, x2, y2, width)
        | _ -> None)
let drawn = borderLines blockSheet
eq "the painter: one line per edge" (drawn.Length, drawn |> List.distinct |> List.length, drawn |> List.forall (fun (_, _, _, _, width) -> width = 2.0)) (24, 24, true)
// A line runs half its width past both ends: an outline has no notch at its corners.
let horizontal = drawn |> List.filter (fun (_, y1, _, y2, _) -> y1 = y2)
let cornerX = horizontal |> List.map (fun (x1, _, _, _, _) -> x1) |> List.min
let leftmostVertical = drawn |> List.filter (fun (x1, _, x2, _, _) -> x1 = x2) |> List.map (fun (x1, _, _, _, _) -> x1) |> List.min
eq "a medium line starts 1 px before the corner it leaves from" (leftmostVertical - cornerX) 1.0
// Two cells with different borders on the edge they share: the painter draws the one of resolveSharedBorder.
let coloured (model: SpreadsheetModel) color =
    Paint.buildPaintItems model (Layout.defaultViewport 800.0 500.0)
    |> List.choose (function Paint.GridLine (_, _, _, _, c, width) when c = color -> Some width | _ -> None)
let conflict first second =
    createEmpty 20 10
    |> setCell (at2 2 2) (cellWith (fun f -> { f with BorderBottom = first }))
    |> setCell (at2 3 2) (cellWith (fun f -> { f with BorderTop = second }))
eq "the painter: the stronger border of the two" (coloured (conflict (line "thin" "#ff0000") (line "thick" "#0000ff")) "#0000ff", coloured (conflict (line "thin" "#ff0000") (line "thick" "#0000ff")) "#ff0000") ([ 3.0 ], [])
eq "the painter: the upper cell's when it is stronger" (coloured (conflict (line "thick" "#ff0000") (line "thin" "#0000ff")) "#ff0000", coloured (conflict (line "thick" "#ff0000") (line "thin" "#0000ff")) "#0000ff") ([ 3.0 ], [])
eq "the painter: the later cell's on a tie" (coloured (conflict (line "medium" "#ff0000") (line "medium" "#0000ff")) "#0000ff", coloured (conflict (line "medium" "#ff0000") (line "medium" "#0000ff")) "#ff0000") ([ 2.0 ], [])
// Print uses Excel's 600-dpi solid rules, while snapshots keep screen geometry.
let widths build style =
    let model = createEmpty 20 10 |> setCell (at2 2 2) (cellWith (fun f -> { f with BorderTop = line style "#123456" }))
    build model (Layout.defaultViewport 800.0 500.0)
    |> List.choose (function Paint.GridLine (_, _, _, _, "#123456", width) -> Some width | _ -> None)
eq "print solid widths in printer dots"
    ([ "hair"; "thin"; "medium"; "thick" ] |> List.map (widths (Paint.buildPrintPaintItems (0.0, 0.0, 1.0))))
    [ [ 0.16 ]; [ 1.28 ]; [ 2.56 ]; [ 3.84 ] ]
eq "print double has two eight-dot rules" (widths (Paint.buildPrintPaintItems (0.0, 0.0, 1.0)) "double") [ 1.28; 1.28 ]
eq "hair stays one physical dot at half scale" (widths (Paint.buildPrintPaintItems (0.0, 0.0, 0.5)) "hair") [ 0.32 ]
eq "screen solid widths unchanged"
    ([ "thin"; "medium"; "thick" ] |> List.map (widths Paint.buildPaintItems))
    [ [ 1.0 ]; [ 2.0 ]; [ 3.0 ] ]
let printSegments style =
    let model = createEmpty 20 10 |> setCell (at2 2 2) (cellWith (fun f -> { f with BorderTop = line style "#123456" }))
    Paint.buildPrintPaintItems (0.0, 0.0, 1.0) model (Layout.defaultViewport 800.0 500.0)
    |> List.choose (function Paint.GridLine (x1, y1, x2, y2, "#123456", width) when y1 = y2 -> Some (x1, x2, width) | _ -> None)
let near a b = abs (a - b) < 0.00001
let hasLength length segments = segments |> List.exists (fun (a, b, _) -> near (b - a) length)
eq "dotted print uses four-dot marks" (printSegments "dotted" |> hasLength 0.64) true
eq "dashed print uses 24-dot marks" (printSegments "dashed" |> hasLength 3.84) true
eq "medium dashed print uses 72-dot marks" (printSegments "mediumDashed" |> hasLength 11.52) true
let dashDot = printSegments "dashDotDot"
eq "dash-dot-dot print has long and short marks" (hasLength 11.52 dashDot, hasLength 3.84 dashDot) (true, true)
let slantModel = createEmpty 20 10 |> setCell (at2 2 2) (cellWith (fun f -> { f with BorderTop = line "slantDashDot" "#123456" }))
let slantItems = Paint.buildPrintPaintItems (0.0, 0.0, 1.0) slantModel (Layout.defaultViewport 800.0 500.0)
eq "slant prints clipped sheared paths"
    (slantItems |> List.exists (function Paint.Clipped (_, paths) -> paths |> List.exists (function Paint.VecPath (_, Some "#123456", Some _) -> true | _ -> false) | _ -> false)) true



// The painter draws the edges the web grid draws (Layout.edgeAbove / edgeLeft): no grid line where a fill hides it, as Excel.
let gridLinesOn (model: SpreadsheetModel) row col =
    let viewport = Layout.defaultViewport 800.0 500.0
    let rect = Layout.cellRect model viewport row col
    let on (a: float) (b: float) = abs (a - b) < 0.5
    Paint.buildPaintItems model viewport
    |> List.filter (function
        | Paint.GridLine (x1, y1, x2, y2, color, _) when color = "#a6a6a6" ->
            (on y1 y2 && (on y1 rect.Y || on y1 (rect.Y + rect.Height)) && x1 >= rect.X - 0.5 && x2 <= rect.X + rect.Width + 0.5)
            || (on x1 x2 && (on x1 rect.X || on x1 (rect.X + rect.Width)) && y1 >= rect.Y - 0.5 && y2 <= rect.Y + rect.Height + 0.5)
        | _ -> false)
    |> List.length
eq "the painter: a fill hides the grid lines around its cell, a cell without one has its four" (gridLinesOn sheet 6 6, gridLinesOn sheet 10 5) (0, 4)

printfn "stage29-borders: %d passed, %d failed" pass fail
if fail > 0 then exit 1
