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
eq "wrapped text is not let out"
    (Paint.buildViewportPaintItems wrapped viewport
     |> List.forall (fun painted -> match painted.Item with Paint.DrawText drawn when drawn.Rect.Y > 30.0 -> drawn.Clip.IsNone | _ -> true)) true
let merged = createEmpty 50 12 |> setCell (at 1 1) (text long) |> addMergedCell 1 1 1 2
eq "a merged cell's text stays in the merged cell" (clipOf merged long |> Option.bind id |> Option.map snd) (Some 200.0)

printfn "stage30-overflow: %d passed, %d failed" pass fail
if fail > 0 then exit 1
