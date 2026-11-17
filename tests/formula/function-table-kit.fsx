#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

// The function table (Formula.fs, "The function table"): the error values Excel 365
// added, functions of the modules of src/Spreadsheet/Functions (installed at start-up
// on .NET), the generated catalog, provider functions and the cached values of their
// formulas. Expected values are those of Microsoft's function pages unless noted.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-50s expected %A  got %A" label expected got)

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let modelOf (cells: ((int * int) * CellValue) list) =
    let map = cells |> List.map (fun ((row, col), value) -> { Row = row; Col = col }, cell value) |> Map.ofList
    { createEmpty 40 12 with Cells = map }
let empty = modelOf []
let value (m: SpreadsheetModel) (formula: string) = unwrapValue (evaluateFormula formula m None)

/// The formula's value, a number rounded to `digits` decimals.
let near label (m: SpreadsheetModel) (formula: string) (digits: int) (expected: float) =
    match value m formula with
    | CellValue.Number x when abs (x - expected) <= 0.5 * (10.0 ** float -digits) -> pass <- pass + 1
    | other -> (fail <- fail + 1; printfn "FAIL %-50s %s expected %g  got %A" label formula expected other)
let is label (m: SpreadsheetModel) (formula: string) (expected: CellValue) = eq (label + " " + formula) (value m formula) expected
let text m formula (expected: string) = is "text" m formula (CellValue.String expected)
let error m formula (code: string) = is "error" m formula (CellValue.Error code)
let number m formula (expected: float) = near "number" m formula 9 expected
let approx m formula digits expected = near "approx" m formula digits expected
let arrayOf (m: SpreadsheetModel) (formula: string) =
    match parseFormula formula with
    | Ok expr -> evalArray expr (contextOf m) |> Array.map (Array.map unwrapValue)
    | Result.Error e -> failwithf "%s: %s" formula e
let N x = CellValue.Number x
let S s = CellValue.String s
let B b = CellValue.Boolean b
let E e = CellValue.Error e
