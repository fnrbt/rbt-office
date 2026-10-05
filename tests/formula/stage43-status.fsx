#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
module Ops = Rbt.Spreadsheet.SheetOps

// What a status bar says of a selection (count, numbers, sum, smallest,
// largest) and the used range of a sheet.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-44s expected %A  got %A" label expected got)

let at row col : CellAddress = { Row = row; Col = col }
let cell (value: CellValue) = { Value = value; Format = defaultFormat; Comment = None }
let sheetOf cells = cells |> List.fold (fun m (addr, data) -> setCell addr data m) (createEmpty 100 10)
let range r0 c0 r1 c1 : CellRange = { From = at r0 c0; To = at r1 c1 }
let status sheet block =
    let s = Ops.selectionStatus sheet block
    s.Count, s.NumericCount, s.Sum, s.Min, s.Max

let sheet =
    sheetOf [
        at 0 0, cell (CellValue.String "Region")
        at 1 0, cell (CellValue.String "East");  at 1 1, cell (CellValue.Number 10.0)
        at 2 0, cell (CellValue.String "West");  at 2 1, cell (CellValue.Number -4.5)
        at 3 0, cell (CellValue.String "East");  at 3 1, cell (CellValue.Formula ("B2*3", CellValue.Number 30.0))
        at 4 0, cell (CellValue.String "North"); at 4 1, cell (CellValue.Boolean true)
        at 5 0, cell (CellValue.String "");      at 5 1, cell (CellValue.Error "#DIV/0!")
        at 6 1, cell (CellValue.Formula ("\"\"", CellValue.String ""))
        at 7 1, { cell CellValue.Empty with Format = { defaultFormat with Bold = true } }
        at 8 1, cell (CellValue.Formula ("\"x\"", CellValue.String "x"))
    ]

eq "numbers: count, sum, min, max" (status sheet (1, 1, 3, 1)) (3, 3, 35.5, Some -4.5, Some 30.0)
eq "text counts, is no number" (status sheet (0, 0, 3, 0)) (4, 0, 0.0, None, None)
eq "booleans and errors count, are no numbers" (status sheet (4, 1, 5, 1)) (2, 0, 0.0, None, None)
eq "empty text and formatted empty cells do not count" (status sheet (5, 0, 7, 1)) (1, 0, 0.0, None, None)
eq "a formula counts by its result" (status sheet (8, 1, 8, 1)) (1, 0, 0.0, None, None)
eq "one cell" (status sheet (1, 1, 1, 1)) (1, 1, 10.0, Some 10.0, Some 10.0)
eq "an empty block" (status sheet (20, 3, 30, 5)) (0, 0, 0.0, None, None)
eq "the block" (status sheet (0, 0, 9, 1)) (11, 3, 35.5, Some -4.5, Some 30.0)
// (A block larger than the cell-by-cell limit walks the cells that exist.)
eq "the whole sheet, by the cells that exist" (status sheet (0, 0, 1048575, 16383)) (11, 3, 35.5, Some -4.5, Some 30.0)
eq "whole columns likewise" (status sheet (0, 1, 1048575, 1)) (6, 3, 35.5, Some -4.5, Some 30.0)

// Rows the AutoFilter hides are left out.
let filtered = sheet |> setAutoFilter (Some (range 0 0 5 1)) |> setColumnFilter 0 (FilterValues (Set.ofList [ "East" ], false))
eq "the filter hides West, North and the blank" (filtered.FilterHiddenRows |> List.ofArray) [ 2; 4; 5 ]
eq "filtered rows are not counted" (status filtered (1, 0, 5, 1)) (4, 2, 40.0, Some 10.0, Some 30.0)
eq "... nor in a whole-sheet selection" (status filtered (0, 0, 1048575, 16383)) (6, 2, 40.0, Some 10.0, Some 30.0)
// Rows hidden by hand are counted.
let handHidden = Ops.hideAxis true 2 2 sheet
eq "rows hidden by hand are counted" (status handHidden (1, 1, 3, 1)) (3, 3, 35.5, Some -4.5, Some 30.0)

// The used range.
eq "an empty sheet has none" (Ops.usedExtent (createEmpty 100 10)) None
eq "from the first to the last cell that holds something" (Ops.usedExtent sheet |> Option.map Ops.rangeText) (Some "A1:B9")
let offset = sheetOf [ at 3 2, cell (CellValue.Number 1.0); at 9 4, { cell CellValue.Empty with Format = { defaultFormat with BgColor = Some "#ff0000" } } ]
eq "it need not start at A1; a formatted cell is used" (Ops.usedExtent offset |> Option.map Ops.rangeText) (Some "C4:E10")
eq "a merged cell is used" (Ops.usedExtent (addMergedCell 1 1 12 6 offset) |> Option.map Ops.rangeText) (Some "B2:G13")
eq "one cell" (Ops.usedExtent (sheetOf [ at 4 4, cell (CellValue.String "x") ]) |> Option.map Ops.rangeText) (Some "E5")

printfn "Status and used range tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
