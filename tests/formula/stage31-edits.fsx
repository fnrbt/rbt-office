#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
module Formula = Rbt.Spreadsheet.Formula
module Wb = Rbt.Spreadsheet.Workbook
module Calc = Rbt.Spreadsheet.Calc

// Edits of large blocks: clearing a block takes its cells out of the sheet
// without rebuilding the rest, parsed formulas are remembered by text (and
// the text scans that decide whether a formula may spill give what they gave),
// and a single-cell edit does not look through every cell for a cycle.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s expected %A  got %A" label expected got)

let number value = { Value = CellValue.Number value; Format = defaultFormat; Comment = None }
let formula text = { Value = CellValue.Formula (text, CellValue.Empty); Format = defaultFormat; Comment = None }
let at row col : CellAddress = { Row = row; Col = col }

// ---- clearing a block ---------------------------------------------------------
let mutable sheet = createEmpty 400 40
for row in 0 .. 299 do
    for col in 0 .. 29 do
        sheet <- setCell (at row col) (number (float (row * 100 + col))) sheet
let expectedAfter (sr, sc, er, ec) =
    sheet.Cells |> Map.filter (fun addr _ -> not (addr.Row >= sr && addr.Row <= er && addr.Col >= sc && addr.Col <= ec))
for block in [ (10, 3, 59, 12); (0, 0, 299, 29); (250, 20, 399, 39); (350, 35, 380, 39) ] do
    let sr, sc, er, ec = block
    let cleared = clearSelectionCells (Range (sr, sc, er, ec)) sheet
    eq (sprintf "block %A cleared" block) cleared.Cells (expectedAfter block)
eq "nothing to clear keeps the model" (obj.ReferenceEquals(clearSelectionCells (Range (350, 35, 380, 39)) sheet, sheet)) true
// A block too large to walk cell by cell goes through the whole map: the same result.
let wide = clearSelectionCells (Range (0, 0, 9999, 99)) sheet
eq "a very large block" (Map.count wide.Cells) 0

// ---- parsed formulas ------------------------------------------------------------
let first = Formula.parseFormula "=SUM(A1:A3)*2"
eq "parsed twice: the same tree" (Formula.parseFormula "=SUM(A1:A3)*2") first
eq "a parse error is remembered as one" (match Formula.parseFormula "=SUM(", Formula.parseFormula "=SUM(" with Result.Error a, Result.Error b -> a = b | _ -> false) true
eq "another text is another tree" (Formula.parseFormula "=SUM(A1:A3)*3" = first) false

// ---- recalculation gives what it gave -----------------------------------------------
let calcSheet =
    createEmpty 50 10
    |> setCell (at 0 0) (number 1.0) |> setCell (at 1 0) (number 2.0) |> setCell (at 2 0) (number 3.0)
    |> setCell (at 0 1) (formula "=sum(a1:a3)")
    |> setCell (at 1 1) (formula "=Index(A1:A3,2)")
    |> setCell (at 2 1) (formula "=IF(A1>0,\"offset is a word\",0)")
    |> setCell (at 3 1) (formula "=COUNTIF(A1:A3,\">1\")")
    |> setCell (at 4 1) (formula "=sequence(2)")
    |> setCell (at 7 1) (formula "=LEN(\"straße\")+SEARCH(\"E\",\"straße\")")
let recalculated = Formula.recalculateAllFormulas None None calcSheet |> Formula.spillArrays None None
let value row col = getCellValue (at row col) recalculated |> unwrapFormulaValue
eq "sum, lower case" (value 0 1) (CellValue.Number 6.0)
eq "index, mixed case" (value 1 1) (CellValue.Number 2.0)
eq "a word that looks like a function" (value 2 1) (CellValue.String "offset is a word")
eq "countif" (value 3 1) (CellValue.Number 2.0)
eq "a spilling function, lower case" (value 4 1, value 5 1) (CellValue.Number 1.0, CellValue.Number 2.0)
eq "case-insensitive search past a sharp s" (value 7 1) (CellValue.Number 12.0)

// ---- a single-cell edit in a workbook without formulas, and with a cycle ------------
let workbookOf (model: SpreadsheetModel) : Wb.Workbook =
    { Sheets = Map.ofList [ "sheet1", ({ Id = "sheet1"; Name = "Sheet1"; Model = model; Visible = true } : Wb.Sheet) ]
      SheetOrder = [ "sheet1" ]; ActiveSheet = "sheet1"; Tables = Map.empty; Names = Map.empty; NameInfo = Map.empty; Iteration = None; External = Map.empty }
let cellOf row col : Wb.SheetCell = { Sheet = "sheet1"; Cell = at row col }
let applied result = match result with Calc.IncrementalApplied _ -> "applied" | Calc.IncrementalFallback why -> "fallback: " + why
eq "plain data: an edit is applied incrementally" (applied (Calc.trySingleCellEdit (cellOf 5 5) (number 9.0) (workbookOf sheet) None)) "applied"
let withFormulas = { (recalculated |> setCell (at 4 1) (number 0.0) |> setCell (at 5 1) (number 0.0)) with Spill = Map.empty }
eq "formulas: an edit they read is applied" (applied (Calc.trySingleCellEdit (cellOf 0 0) (number 10.0) (workbookOf withFormulas) None)) "applied"
let cyclic = createEmpty 10 10 |> setCell (at 0 0) (formula "=B1+1") |> setCell (at 0 1) (formula "=A1+1") |> setCell (at 3 3) (number 1.0)
eq "a cycle: an edit that feeds it falls back"
    ((applied (Calc.trySingleCellEdit (cellOf 0 0) (formula "=B1+2") (workbookOf cyclic) None)).StartsWith "fallback") true

printfn "stage31-edits: %d passed, %d failed" pass fail
if fail > 0 then exit 1
