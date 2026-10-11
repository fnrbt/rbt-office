#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Input

// The 1904 date system: serial 0 is 1904-01-01. The serials of such a
// workbook are kept as they are; number formats, the date functions and
// typed dates work in the workbook's system, which every sheet of it says
// (SpreadsheetModel.Date1904). There is no system "at work": two workbooks
// in different systems are evaluated, shown, filtered and searched side by
// side, in any order.

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)
let eq label (got: 'a) (expected: 'a) =
    ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let N x = CellValue.Number x
let sheetIn (date1904: bool) = { createEmpty 10 10 with Date1904 = date1904 } |> setCell { Row = 0; Col = 0 } { Value = CellValue.Number 42000.0; Format = defaultFormat; Comment = None }

// 42000 is 2014-12-27 in the 1900 system and 2018-12-28 (a Friday) in the 1904 system.
let check (date1904: bool) (label: string) (year: int, month: int, day: int) (weekday: float) (nextMonth: float) =
    let model = sheetIn date1904
    let value formula = evaluateFormula formula model None
    let format code serial = Rbt.Spreadsheet.CellText.formatNumberWithCodeIn date1904 code serial
    eq (label + ": a serial shown as a date") (format "yyyy-mm-dd" 42000.0) (sprintf "%04d-%02d-%02d" year month day)
    eq (label + ": ... with the weekday") (format "ddd" 42000.0) (if weekday = 6.0 then "Fri" else "Sat")
    eq (label + ": DATE") (value (sprintf "=DATE(%d,%d,%d)" year month day)) (N 42000.0)
    eq (label + ": YEAR, MONTH, DAY") (value "=YEAR(A1)*10000+MONTH(A1)*100+DAY(A1)") (N (float (year * 10000 + month * 100 + day)))
    eq (label + ": WEEKDAY") (value "=WEEKDAY(A1)") (N weekday)
    eq (label + ": EDATE") (value "=EDATE(A1,1)") (N nextMonth)
    eq (label + ": DATEVALUE") (value (sprintf "=DATEVALUE(\"%04d-%02d-%02d\")" year month day)) (N 42000.0)
    eq (label + ": TEXT") (value "=TEXT(A1,\"yyyy-mm-dd\")") (CellValue.String (sprintf "%04d-%02d-%02d" year month day))
    eq (label + ": a typed date") ((parseInputFor date1904 2024 (sprintf "%04d-%02d-%02d" year month day)).Value) (N 42000.0)
    eq (label + ": dateSerialIn") (dateSerialIn date1904 year month day) 42000.0
    // Monday to Friday of that week (or the week before, in the 1900 system).
    eq (label + ": NETWORKDAYS over the five days before") (value "=NETWORKDAYS(A1-4,A1)") (N (if weekday = 6.0 then 5.0 else 4.0))
    eq (label + ": WORKDAY") (value "=WORKDAY(A1,1)") (N (if weekday = 6.0 then 42003.0 else 42002.0))
    eq (label + ": the text to edit") (editTextFor date1904 (Some { Value = N 42000.0; Format = { defaultFormat with NumberFormat = Some "yyyy-mm-dd" }; Comment = None })) (sprintf "%d/%d/%04d" month day year)

check true "1904" (2018, 12, 28) 6.0 42031.0
let format1904 code serial = Rbt.Spreadsheet.CellText.formatNumberWithCodeIn true code serial
let format code serial = Rbt.Spreadsheet.CellText.formatNumberWithCode code serial
eq "1904: serial 0 is 1904-01-01" (format1904 "yyyy-mm-dd" 0.0) "1904-01-01"
eq "1904: serial 59 and 60 are the last of February and the first of March (1904 is a leap year: 60 is Mar 1)" (format1904 "m/d" 59.0, format1904 "m/d" 60.0) ("2/29", "3/1")
eq "1904: DATE(1904,1,1)" (evaluateFormula "=DATE(1904,1,1)" (sheetIn true) None) (N 0.0)
eq "1904: a date before 1904 is not typed as a date" (match (parseInputFor true 2024 "1903-06-01").Value with CellValue.Number _ -> false | _ -> true) true
eq "1904: times are what they are" (format1904 "h:mm" 0.75) "18:00"
check false "1900" (2014, 12, 27) 7.0 42031.0
eq "1900: serial 61 is 1900-03-01" (format "yyyy-mm-dd" 61.0) "1900-03-01"
eq "1900 is the system of what says nothing" ((parseInputIn 2024 "2014-12-27").Value, dateSerial 2014 12 27, (createEmpty 1 1).Date1904) (N 42000.0, 42000.0, false)

// --- two workbooks side by side ---------------------------------------------
module Wb = Rbt.Spreadsheet.Workbook
module Ops = Rbt.Spreadsheet.SheetOps
let at row col : CellAddress = { Row = row; Col = col }
let dateCell serial = { Value = N serial; Format = { defaultFormat with NumberFormat = Some "yyyy-mm-dd" }; Comment = None }
let formulaCell text = { Value = CellValue.Formula (text, CellValue.Empty); Format = defaultFormat; Comment = None }
/// A workbook of two sheets: three days in A2:A4 of the first (under a header), formulas of them on the second.
let book (date1904: bool) : Wb.Workbook =
    let days = [ 2018, 12, 28; 2019, 1, 15; 2019, 2, 1 ]
    let first =
        days
        |> List.indexed
        |> List.fold (fun m (i, (y, mo, d)) -> setCell (at (i + 1) 0) (dateCell (dateSerialIn date1904 y mo d)) m)
                     (createEmpty 20 5 |> setCell (at 0 0) { Value = CellValue.String "Day"; Format = defaultFormat; Comment = None })
    let second =
        createEmpty 20 5
        |> setCell (at 0 0) (formulaCell "=TEXT(Sheet1!A2,\"yyyy-mm-dd\")&\"|\"&YEAR(Sheet1!A3)&\"|\"&WEEKDAY(Sheet1!A2)")
        |> setCell (at 1 0) (formulaCell "=DATE(2018,12,28)")
        |> setCell (at 2 0) (formulaCell "=EOMONTH(Sheet1!A2,0)-Sheet1!A2")
    { Wb.createWorkbook () with
        Sheets = Map.ofList [ "sheet1", ({ Id = "sheet1"; Name = "Sheet1"; Model = first; Visible = true } : Wb.Sheet)
                              "sheet2", ({ Id = "sheet2"; Name = "Sheet2"; Model = second; Visible = true } : Wb.Sheet) ]
        SheetOrder = [ "sheet1"; "sheet2" ] }
    |> Wb.withDate1904 date1904
let a, b = book true, book false
eq "two workbooks: each says its system" (Wb.date1904 a, Wb.date1904 b) (true, false)
eq "... on every sheet" ([ for KeyValue(_, s) in a.Sheets -> s.Model.Date1904 ], [ for KeyValue(_, s) in b.Sheets -> s.Model.Date1904 ]) ([ true; true ], [ false; false ])
eq "the serials differ by 1462 days" ((getCellValue (at 1 0) a.Sheets.["sheet1"].Model), (getCellValue (at 1 0) b.Sheets.["sheet1"].Model)) (N 42000.0, N 43462.0)
let valuesOf (wb: Wb.Workbook) =
    let model = wb.Sheets.["sheet2"].Model
    [ for row in 0 .. 2 ->
        match getCellValue (at row 0) model with
        | CellValue.Formula (text, _) -> evaluateFormula text model (Some { wb with ActiveSheet = "sheet2" }) |> unwrapFormulaValue
        | other -> other ]
// Evaluated in turns: a, b, a, b - neither leaves anything behind for the other.
let wantA = [ CellValue.String "2018-12-28|2019|6"; N 42000.0; N 3.0 ]
let wantB = [ CellValue.String "2018-12-28|2019|6"; N 43462.0; N 3.0 ]
eq "formulas, in turns" [ valuesOf a; valuesOf b; valuesOf a; valuesOf b ] [ wantA; wantB; wantA; wantB ]
// A full recalculation of each workbook through the engine, in turns.
let recalculated (wb: Wb.Workbook) =
    let done' = Rbt.Spreadsheet.Formula.recalculateWorkbook wb
    [ for row in 0 .. 2 -> getCellValue (at row 0) done'.Sheets.["sheet2"].Model |> unwrapFormulaValue ]
eq "recalculated, in turns" [ recalculated b; recalculated a; recalculated b; recalculated a ] [ wantB; wantA; wantB; wantA ]
// What the cells show, a filter's values and a search, in turns.
let shownIn (wb: Wb.Workbook) =
    let model = wb.Sheets.["sheet1"].Model
    [ for row in 1 .. 3 -> Rbt.Spreadsheet.CellText.cellDisplayTextIn model.Date1904 None 1.0e6 (getCell (at row 0) model).Value.Format (getCellValue (at row 0) model) ]
let days = [ "2018-12-28"; "2019-01-15"; "2019-02-01" ]
eq "shown, in turns" [ shownIn a; shownIn b; shownIn a ] [ days; days; days ]
let filterValues (wb: Wb.Workbook) =
    wb.Sheets.["sheet1"].Model |> setAutoFilter (Some { From = at 0 0; To = at 3 0 }) |> filterColumnValues 0 |> fst
eq "a filter's values, in turns" [ filterValues a; filterValues b; filterValues a ] [ days; days; days ]
let filtered (wb: Wb.Workbook) =
    let model = wb.Sheets.["sheet1"].Model
    let criteria = Ops.filterOfDescriptionIn model.Date1904 "t" { Ops.emptyFilterDescription "conditions" with Conditions = [ { Operator = "greaterOrEqual"; Value = "2019-01-15" } ] }
    (model |> setAutoFilter (Some { From = at 0 0; To = at 3 0 }) |> setColumnFilter 0 criteria).FilterHiddenRows |> List.ofArray
eq "a filter by a date, in turns" [ filtered a; filtered b; filtered a ] [ [ 1 ]; [ 1 ]; [ 1 ] ]
let filteredByValue (wb: Wb.Workbook) =
    (wb.Sheets.["sheet1"].Model |> setAutoFilter (Some { From = at 0 0; To = at 3 0 }) |> setColumnFilter 0 (FilterValues (Set.ofList [ "2019-01-15" ], false))).FilterHiddenRows |> List.ofArray
eq "a filter by a shown value, in turns" [ filteredByValue a; filteredByValue b; filteredByValue a ] [ [ 1; 3 ]; [ 1; 3 ]; [ 1; 3 ] ]
let found (wb: Wb.Workbook) =
    let model = wb.Sheets.["sheet1"].Model
    Ops.findInSheet (Ops.finderOf "Find" "2019" Ops.defaultFindOptions) Ops.defaultFindOptions None model |> List.map (fun (addr, _) -> addr.Row)
eq "a search of what is shown, in turns" [ found a; found b; found a ] [ [ 2; 3 ]; [ 2; 3 ]; [ 2; 3 ] ]
let replaced (wb: Wb.Workbook) =
    let model = wb.Sheets.["sheet1"].Model
    let options = { Ops.defaultFindOptions with InFormulas = true }
    Ops.replaceInSheet (Ops.finderOf "ReplaceAll" "12/28/2018" options) options "12/29/2018" 2024 None model
    |> List.map (fun (addr, _, after) -> addr.Row, after.Value)
eq "a replacement typed as a date, in turns" [ replaced a; replaced b; replaced a ] [ [ 1, N 42001.0 ]; [ 1, N 43463.0 ]; [ 1, N 42001.0 ] ]
eq "a sheet added to a workbook is in its system" ((Wb.addSheet "More" a).Sheets |> Map.forall (fun _ s -> s.Model.Date1904), (Wb.addSheet "More" b).Sheets |> Map.exists (fun _ s -> s.Model.Date1904)) (true, false)

printfn "\nstage50 (1904 date system): %d passed, %d failed" pass fail
if fail > 0 then exit 1
