#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
module Layout = Rbt.Spreadsheet.Layout
module Paint = Rbt.Spreadsheet.Paint

// Excel's text overflow: text wider than its cell runs over the empty cells
// beside it (right for left-aligned text, left for right-aligned, both for
// centred) and stops at the first cell that holds something; numbers,
// wrapped text and merged cells stay in their cell. The portable painter
// gives the text the run of cells as its clip.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s expected %A  got %A" label expected got)

let text value = { Value = CellValue.String value; Format = defaultFormat; Comment = None }
let aligned align value = { text value with Format = { defaultFormat with HorizontalAlign = Some align } }
let at row col : CellAddress = { Row = row; Col = col }

// ---- which sides ---------------------------------------------------------------
eq "general text: to the right" (Layout.textOverflowSides defaultFormat (CellValue.String "abc")) (Some (false, true))
eq "left" (Layout.textOverflowSides { defaultFormat with HorizontalAlign = Some "left" } (CellValue.String "abc")) (Some (false, true))
eq "right: to the left" (Layout.textOverflowSides { defaultFormat with HorizontalAlign = Some "right" } (CellValue.String "abc")) (Some (true, false))
eq "centre: both" (Layout.textOverflowSides { defaultFormat with HorizontalAlign = Some "center" } (CellValue.String "abc")) (Some (true, true))
eq "a formula giving text" (Layout.textOverflowSides defaultFormat (CellValue.Formula ("=A1", CellValue.String "abc"))) (Some (false, true))
eq "a number never" (Layout.textOverflowSides defaultFormat (CellValue.Number 12345678.0)) None
eq "a formula giving a number never" (Layout.textOverflowSides defaultFormat (CellValue.Formula ("=A1", CellValue.Number 1.0))) None
eq "wrapped text never" (Layout.textOverflowSides { defaultFormat with WrapText = true } (CellValue.String "abc")) None
eq "justified text never" (Layout.textOverflowSides { defaultFormat with HorizontalAlign = Some "justify" } (CellValue.String "abc")) None
eq "rotated text never" (Layout.textOverflowSides { defaultFormat with TextRotation = 45 } (CellValue.String "abc")) None
eq "empty text never" (Layout.textOverflowSides defaultFormat (CellValue.String "")) None

// ---- how far ---------------------------------------------------------------------
// Columns are 100 px wide. B2 holds long text; G2 holds something.
let sheet =
    createEmpty 50 12
    |> setCell (at 1 1) (text "A long title that needs several columns")
    |> setCell (at 1 6) (text "stop")
eq "needs 250px: three more columns" (Layout.textOverflowSpan sheet 1 1 0.0 250.0) (1, 4)
eq "needs 1px: one more column" (Layout.textOverflowSpan sheet 1 1 0.0 1.0) (1, 2)
eq "needs nothing" (Layout.textOverflowSpan sheet 1 1 0.0 0.0) (1, 1)
eq "stops at the first filled cell" (Layout.textOverflowSpan sheet 1 1 0.0 5000.0) (1, 5)
eq "to the left, up to the sheet's edge" (Layout.textOverflowSpan sheet 1 1 5000.0 0.0) (0, 1)
eq "both ways" (Layout.textOverflowSpan sheet 1 3 150.0 150.0) (2, 5)
eq "a filled neighbour: stays in its cell" (Layout.textOverflowSpan (setCell (at 1 2) (text "x") sheet) 1 1 0.0 250.0) (1, 1)
eq "a neighbour with only a format is empty"
    (Layout.textOverflowSpan (setCell (at 1 2) { text "" with Value = CellValue.Empty; Format = { defaultFormat with Bold = true } } sheet) 1 1 0.0 150.0) (1, 3)
eq "a formula giving nothing still stops it"
    (Layout.textOverflowSpan (setCell (at 1 3) { text "" with Value = CellValue.Formula ("=\"\"", CellValue.String "") } sheet) 1 1 0.0 5000.0) (1, 2)
eq "a merged cell stops it" (Layout.textOverflowSpan (addMergedCell 1 4 2 5 sheet) 1 1 0.0 5000.0) (1, 3)
eq "a hidden column is passed over" (Layout.textOverflowSpan (setColumnWidth 2 0.0 sheet) 1 1 0.0 150.0) (1, 4)
eq "the last column" (Layout.textOverflowSpan sheet 3 11 0.0 500.0) (11, 11)

// ---- the portable painter --------------------------------------------------------
// A measurer that gives every character 10 px.
let viewport = { Layout.defaultViewport 1400.0 600.0 with Measurer = Some (fun text _ _ _ _ -> float text.Length * 10.0) }
let clipOf (model: SpreadsheetModel) (wanted: string) =
    Paint.buildViewportPaintItems model viewport
    |> List.tryPick (fun painted ->
        match painted.Item with
        | Paint.DrawText drawn when drawn.Text = wanted -> Some (drawn.Clip |> Option.map (fun clip -> clip.X, clip.Width))
        | _ -> None)
let cellX col = (Layout.cellRect sheet viewport 1 col).X
let long = "A long title that needs several columns"   // 39 characters: 390 px
eq "text that fits has no clip" (clipOf (setCell (at 3 1) (text "short") sheet) "short") (Some None)
eq "long text: clipped to the cells it runs over" (clipOf sheet long) (Some (Some (cellX 1, 400.0)))
eq "blocked at once: clipped to its own cell" (clipOf (setCell (at 1 2) (text "x") sheet) long) (Some (Some (cellX 1, 100.0)))
eq "blocked further on" (clipOf (setCell (at 1 3) (text "x") sheet) long) (Some (Some (cellX 1, 200.0)))
let rightAligned = createEmpty 50 12 |> setCell (at 1 5) (aligned "right" long)
eq "right-aligned: runs to the left" (clipOf rightAligned long) (Some (Some (cellX 2, 400.0)))
let centred = createEmpty 50 12 |> setCell (at 1 5) (aligned "center" long)
eq "centred: runs both ways" (clipOf centred long) (Some (Some (cellX 3, 500.0)))
let wrapped = createEmpty 50 12 |> setCell (at 1 1) { text long with Format = { defaultFormat with WrapText = true } }
// (Clipped to its own cell: the lines past its bottom are not drawn, as in Excel's PDF.)
eq "wrapped text is not let out"
    (Paint.buildViewportPaintItems wrapped viewport
     |> List.forall (fun painted -> match painted.Item with Paint.DrawText drawn when drawn.Rect.Y > 30.0 -> (match drawn.Clip with None -> true | Some clip -> abs (clip.X - cellX 1) < 1e-6 && abs (clip.Width - 100.0) < 1e-6) | _ -> true)) true
let merged = createEmpty 50 12 |> setCell (at 1 1) (text long) |> addMergedCell 1 1 1 2
eq "a merged cell's text stays in the merged cell" (clipOf merged long |> Option.bind id |> Option.map snd) (Some 200.0)

// ---- the grid lines the text crosses ------------------------------------------
// Text that runs over the cells beside its own hides the vertical grid lines
// it crosses (not the lines above and below the row, not a border, not the
// line at the cell that stops it); wrapped and rotated text crosses none.
let verticalGridLines (model: SpreadsheetModel) (row: int) =
    let top, bottom = 30.0 + Layout.rowTop model row, 30.0 + Layout.rowTop model (row + 1)
    Paint.buildPaintItems model (Layout.defaultViewport 1200.0 400.0)
    |> List.choose (function
        | Paint.GridLine (x1, y1, x2, y2, color, _) when x1 = x2 && color = "#a6a6a6" && y1 >= top - 0.5 && y2 <= bottom + 0.5 -> Some (int (x1 - 50.0) / 100)
        | _ -> None)
    |> List.distinct |> List.sort
let horizontalGridLines (model: SpreadsheetModel) (row: int) =
    let y = 30.0 + Layout.rowTop model row
    Paint.buildPaintItems model (Layout.defaultViewport 1200.0 400.0)
    |> List.filter (function Paint.GridLine (x1, y1, x2, y2, color, _) when y1 = y2 && y1 = y && color = "#a6a6a6" -> true | _ -> false)
    |> List.length
let lines = "Text that is wider than one cell, and than three of them too"
let plain = createEmpty 30 12
let across = plain |> setCell (at 2 1) (text lines)
eq "an empty row: a grid line at every column edge" (verticalGridLines plain 2 |> List.truncate 6) [ 0; 1; 2; 3; 4; 5 ]
let crossedBy (model: SpreadsheetModel) row = [ 0 .. 8 ] |> List.filter (fun edge -> not (List.contains edge (verticalGridLines model row)))
let widthNeeded = crossedBy across 2
eq "text running right: the lines it crosses are not drawn" (widthNeeded.Length >= 2 && widthNeeded.Head = 2 && widthNeeded = [ 2 .. 1 + widthNeeded.Length ]) true
eq "the lines above and below its row are whole" (horizontalGridLines across 2, horizontalGridLines across 3) (horizontalGridLines plain 2, horizontalGridLines plain 3)
eq "the rows around it keep every line" (crossedBy across 1, crossedBy across 3) ([], [])
eq "stopped by a cell: only the lines before it" (crossedBy (across |> setCell (at 2 3) (text "stop")) 2) [ 2 ]
eq "right-aligned: the lines to its left" (crossedBy (plain |> setCell (at 2 6) (aligned "right" lines)) 2 |> List.forall (fun edge -> edge <= 6), (crossedBy (plain |> setCell (at 2 6) (aligned "right" lines)) 2).Length >= 2) (true, true)
eq "centred: lines on both sides" (let crossed = crossedBy (plain |> setCell (at 2 5) (aligned "center" lines)) 2 in List.contains 5 crossed && List.contains 6 crossed) true
eq "wrapped text crosses none" (crossedBy (plain |> setCell (at 2 1) { text lines with Format = { defaultFormat with WrapText = true } }) 2) []
eq "rotated text crosses none" (crossedBy (plain |> setCell (at 2 1) { text lines with Format = { defaultFormat with TextRotation = 45 } }) 2) []
eq "a number crosses none" (crossedBy (plain |> setCell (at 2 1) { Value = CellValue.Number 123456789012345.0; Format = defaultFormat; Comment = None }) 2) []
// A border on a crossed edge is a border still.
let bordered = across |> setCell (at 2 3) { Value = CellValue.Empty; Format = { defaultFormat with BorderLeft = Some { Style = "thin"; Color = Some "#ff0000" } }; Comment = None }
eq "a border the text crosses is drawn"
    (Paint.buildPaintItems bordered (Layout.defaultViewport 1200.0 400.0) |> List.filter (function Paint.GridLine (_, _, _, _, color, _) -> color = "#ff0000" | _ -> false) |> List.length) 1

// ---- General alignment ---------------------------------------------------------
// Excel's General rule, one for every painter (the web grid's cell nodes use it too):
// numbers and dates right, text left, booleans and errors centred, explicit over it.
let alignOf align numberFormat value =
    Layout.cellTextAlign { defaultFormat with HorizontalAlign = align; NumberFormat = numberFormat } value
eq "General: a number right" (alignOf None None (CellValue.Number 12.0)) Layout.AlignRight
eq "General: a date (a number with a date format) right" (alignOf None (Some "yyyy-mm-dd") (CellValue.Number 45306.0)) Layout.AlignRight
eq "General: text left" (alignOf None None (CellValue.String "abc")) Layout.AlignLeft
eq "General: a boolean centred" (alignOf None None (CellValue.Boolean true)) Layout.AlignCenter
eq "General: an error centred" (alignOf None None (CellValue.Error "#N/A")) Layout.AlignCenter
eq "General: a formula by what it gives" (alignOf None None (CellValue.Formula ("=1+1", CellValue.Number 2.0)), alignOf None None (CellValue.Formula ("=\"a\"", CellValue.String "a"))) (Layout.AlignRight, Layout.AlignLeft)
eq "`general` written out is General" (alignOf (Some "general") None (CellValue.Number 1.0)) Layout.AlignRight
eq "a Text (@) format: a number left" (alignOf None (Some "@") (CellValue.Number 12.0)) Layout.AlignLeft
eq "explicit over General" (alignOf (Some "left") None (CellValue.Number 1.0), alignOf (Some "right") None (CellValue.String "a"), alignOf (Some "center") None (CellValue.Number 1.0), alignOf (Some "centerContinuous") None (CellValue.String "a")) (Layout.AlignLeft, Layout.AlignRight, Layout.AlignCenter, Layout.AlignCenter)
let alignPainted (model: SpreadsheetModel) (wanted: string) =
    Paint.buildViewportPaintItems model viewport
    |> List.tryPick (fun painted -> match painted.Item with Paint.DrawText drawn when drawn.Text = wanted -> Some drawn.Align | _ -> None)
let generalSheet =
    createEmpty 20 6
    |> setCell (at 1 1) { Value = CellValue.Number 1234.0; Format = defaultFormat; Comment = None }
    |> setCell (at 2 1) (text "label")
    |> setCell (at 3 1) { Value = CellValue.Boolean true; Format = defaultFormat; Comment = None }
eq "the painter places them by the same rule" (alignPainted generalSheet "1234", alignPainted generalSheet "label", alignPainted generalSheet "TRUE") (Some Layout.AlignRight, Some Layout.AlignLeft, Some Layout.AlignCenter)

// ---- the one look of a cell (Paint.cellLook): what the painters and the web grid draw ----------------
// Where the painter and the web grid disagreed, Excel's way: an indent on General text pads its left,
// a number format's [Red] section colours the text, centred across the selection spans the run of
// empty center-across cells, a line break of unwrapped text is not shown, a number shrunk to fit is
// not turned to ####.
let lookOf (model: SpreadsheetModel) row col = Paint.cellLook model row col (getCell (at row col) model) None 64.0
let lookSheet =
    createEmpty 20 10
    |> setCell (at 0 0) { text "indented" with Format = { defaultFormat with Indent = 2 } }
    |> setCell (at 1 0) { Value = CellValue.Number -5.0; Format = { defaultFormat with NumberFormat = Some "0;[Red]-0" }; Comment = None }
    |> setCell (at 2 0) (aligned "centerContinuous" "Title")
    |> setCell (at 2 1) { aligned "centerContinuous" "" with Value = CellValue.Empty }
    |> setCell (at 2 2) { aligned "centerContinuous" "" with Value = CellValue.Empty }
    |> setCell (at 3 0) (text "one\ntwo")
    |> setCell (at 4 0) { Value = CellValue.Number 123456789.0; Format = { defaultFormat with NumberFormat = Some "#,##0.00"; ShrinkToFit = true }; Comment = None }
eq "look: an indent on General text pads its left" ((lookOf lookSheet 0 0).IndentLeft > 15.0, (lookOf lookSheet 0 0).IndentRight) (true, 0.0)
eq "look: a [Red] section colours a negative number" (lookOf lookSheet 1 0).TextColor (Some "#FF0000")
eq "look: centred across the run of empty center-across cells" (lookOf lookSheet 2 0).AcrossTo 2
eq "look: unwrapped text shows no line break" (lookOf lookSheet 3 0).Text "onetwo"
eq "look: a number shrunk to fit is not ####" (lookOf lookSheet 4 0).Text "123,456,789.00"
let paintedText (model: SpreadsheetModel) row =
    let rect = Layout.cellRect model (Layout.defaultViewport 900.0 500.0) row 0
    Paint.buildPaintItems model (Layout.defaultViewport 900.0 500.0)
    |> List.tryPick (function Paint.DrawText t when abs (t.Rect.Y - rect.Y) < 1.0 && t.Rect.X >= rect.X && t.Rect.X < rect.X + 30.0 -> Some t | _ -> None)
eq "the painter: General text indented" ((paintedText lookSheet 0).Value.Rect.X - (Layout.cellRect lookSheet (Layout.defaultViewport 900.0 500.0) 0 0).X > 15.0) true
eq "the painter: centred across the run" ((paintedText lookSheet 2).Value.Rect.Width > 250.0) true

printfn "stage30-overflow: %d passed, %d failed" pass fail
if fail > 0 then exit 1
