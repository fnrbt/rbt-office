#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
module Ops = Rbt.Spreadsheet.SheetOps
module Formula = Rbt.Spreadsheet.Formula

// The quick sort (Sort A to Z / Z to A with one cell selected): the block it
// sorts is Excel's current region around the active cell, its header row is
// guessed as Excel guesses it, and blanks are last in both directions.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-44s expected %A  got %A" label expected got)

Formula.formulaHooksInstalled |> ignore

let at row col : CellAddress = { Row = row; Col = col }
let cell (value: CellValue) = { Value = value; Format = defaultFormat; Comment = None }
let text (value: string) = cell (CellValue.String value)
let num (n: float) = cell (CellValue.Number n)
let sheetOf cells = cells |> List.fold (fun m (addr, data) -> setCell addr data m) (createEmpty 200 30)
let range a b c d : CellRange = { From = at a b; To = at c d }
let column col (m: SpreadsheetModel) rows = [ for row in rows -> match getCell (at row col) m with Some data -> displayValue data.Value | None -> "" ]
let quick (active: CellAddress) (descending: bool) (m: SpreadsheetModel) =
    match Ops.quickSortTarget active m with
    | Some (area, options) -> Ops.sortRange "t" area [ { Column = active.Col; Descending = descending; By = Ops.SortByValue; Color = "" } ] options m
    | None -> m

// ---- the reported case: a block at R101:T105 under other data ---------------------------
// Rows 1-9 hold another list in A:D, and a column of notes runs down V: the
// rows of the sheet are filled without a gap from the top to the block.
let reported =
    sheetOf ([ at 0 0, text "Region"; at 0 1, text "Units"
               at 1 0, text "East"; at 1 1, num 30.0
               at 2 0, text "West"; at 2 1, num 10.0
               // R101:T105 (0-based rows 100..104, columns 17..19)
               at 100 17, text "Item";   at 100 18, text "Qty"; at 100 19, text "Price"
               at 101 17, text "bolt";   at 101 18, num 4.0;    at 101 19, num 1.5
               at 102 17, text "nut";    at 102 18, num 9.0
               at 103 17, text "washer"; at 103 18, num 2.0;    at 103 19, num 2.0
               at 104 17, text "pin";    at 104 18, num 7.0 ]
             @ [ for row in 3 .. 99 -> at row 21, num (float row) ])
eq "region: the block around T103" (currentRegion (at 102 19) reported) (Some (range 100 17 104 19))
eq "region: from a blank cell inside the block" (currentRegion (at 102 19) reported |> Option.map (fun r -> r.From.Row)) (Some 100)
eq "region: a text row over numbers is a header" (regionHasHeader (range 100 17 104 19) reported) true
let desc = quick (at 102 19) true reported
eq "quick sort desc: the header stays" (column 19 desc [ 100 ], column 17 desc [ 100 ]) ([ "Price" ], [ "Item" ])
eq "quick sort desc: numbers down, blanks last" (column 19 desc [ 101 .. 104 ]) [ "2"; "1.5"; ""; "" ]
eq "quick sort desc: the rows travel" (column 17 desc [ 101 .. 104 ]) [ "washer"; "bolt"; "nut"; "pin" ]
eq "quick sort desc: the other list is untouched" (column 0 desc [ 0 .. 2 ], column 21 desc [ 3; 99 ]) ([ "Region"; "East"; "West" ], [ "3"; "99" ])
let asc = quick (at 102 19) false reported
eq "quick sort asc: the header stays" (column 19 asc [ 100 ]) [ "Price" ]
eq "quick sort asc: blanks last too" (column 19 asc [ 101 .. 104 ]) [ "1.5"; "2"; ""; "" ]
eq "quick sort asc: the rows travel" (column 17 asc [ 101 .. 104 ]) [ "bolt"; "washer"; "nut"; "pin" ]
// The model-level sort the other hosts use does the same.
let viaModel = sortRegionByColumn (at 102 19) false reported
eq "sortRegionByColumn: header stays, blanks last" (column 19 viaModel [ 100 .. 104 ]) [ "Price"; "2"; "1.5"; ""; "" ]
eq "sortRegionByColumn: cells beside the block stay" (column 21 viaModel [ 100 ], column 0 viaModel [ 1 ]) ([ "" ], [ "East" ])
let byColumn = sortByColumn 19 true reported
eq "sortByColumn: the block of the column, header kept" (column 19 byColumn [ 100 .. 104 ], column 0 byColumn [ 0 .. 2 ]) ([ "Price"; "1.5"; "2"; ""; "" ], [ "Region"; "East"; "West" ])

// ---- the current region -------------------------------------------------------------------
let shapes =
    sheetOf [ at 5 5, num 1.0; at 6 6, num 2.0          // touch at a corner
              at 5 8, num 3.0                           // one empty column away: another block
              at 8 6, num 4.0                           // one empty row away
              at 20 2, text "alone" ]
eq "region: corners join" (currentRegion (at 5 5) shapes) (Some (range 5 5 6 6))
eq "region: an empty column parts blocks" (currentRegion (at 5 8) shapes) (Some (range 5 8 5 8))
eq "region: an empty row parts blocks" (currentRegion (at 8 6) shapes) (Some (range 8 6 8 6))
eq "region: an empty cell beside data joins its block" (currentRegion (at 5 6) shapes) (Some (range 5 5 6 6))
eq "region: an empty cell with nothing around" (currentRegion (at 30 10) shapes) None
eq "region: one cell" (currentRegion (at 20 2) shapes) (Some (range 20 2 20 2))
eq "region: at the sheet's corner" (currentRegion (at 0 0) (sheetOf [ at 0 0, num 1.0; at 1 0, num 2.0 ])) (Some (range 0 0 1 0))

// ---- the header guess ----------------------------------------------------------------------
let allText = sheetOf [ at 0 0, text "Fruit"; at 1 0, text "pear"; at 2 0, text "apple" ]
eq "header: text over text formatted alike is data" (regionHasHeader (range 0 0 2 0) allText) false
eq "quick sort: an all-text list sorts its first row too" (column 0 (quick (at 1 0) false allText) [ 0 .. 2 ]) [ "apple"; "Fruit"; "pear" ]
let boldHeader = allText |> setCell (at 0 0) { text "Fruit" with Format = { defaultFormat with Bold = true } }
eq "header: a first row formatted apart is a header" (regionHasHeader (range 0 0 2 0) boldHeader) true
eq "quick sort: the bold header stays" (column 0 (quick (at 1 0) false boldHeader) [ 0 .. 2 ]) [ "Fruit"; "apple"; "pear" ]
let numbers = sheetOf [ at 0 0, num 3.0; at 1 0, num 1.0; at 2 0, num 2.0 ]
eq "header: numbers have none" (regionHasHeader (range 0 0 2 0) numbers) false
let years = sheetOf [ at 0 0, text "Name"; at 0 1, num 2024.0; at 1 0, text "a"; at 1 1, text "x"; at 2 0, text "b"; at 2 1, text "y" ]
eq "header: a number over text is data" (regionHasHeader (range 0 0 2 1) years) false
let oneRow = sheetOf [ at 0 0, text "Name"; at 0 1, text "Qty" ]
eq "header: one row has none" (regionHasHeader (range 0 0 0 1) oneRow) false
let gap = sheetOf [ at 0 0, text "Name"; at 0 1, text "Qty"; at 1 0, text "a"; at 2 0, text "b"; at 2 1, num 5.0 ]
eq "header: the first filled cell below decides" (regionHasHeader (range 0 0 2 1) gap) true

// ---- tables ----------------------------------------------------------------------------------
let tableSheet =
    { sheetOf [ at 0 0, text "Name"; at 0 1, text "Kind"
                at 1 0, text "pear"; at 1 1, text "b"
                at 2 0, text "apple"; at 2 1, text "a"
                at 3 0, text "Total"; at 3 1, text "2" ] with
        Tables = [ { Name = "Fruit"; StartRow = 1; StartCol = 1; EndRow = 4; EndCol = 2; HeaderRowCount = 1; Style = None; AutoFilter = None; HasTotals = true } ] }
let tableTarget = Ops.quickSortTarget (at 1 0) tableSheet
eq "table: its header row is a header, its totals row is not sorted" (tableTarget |> Option.map (fun (area, options) -> area, options.HasHeader)) (Some (range 0 0 2 1, true))

printfn "stage46 quick sort: %d passed, %d failed" pass fail
if fail > 0 then exit 1
