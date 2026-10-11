#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

// The functions a model is mostly made of - IF, AND, OR, NOT, IFERROR, YEAR, MONTH, DAY,
// EOMONTH, EDATE - have a short way through the evaluator (`evalCommonFunction`), and two
// numbers under an operator one through `applyBin`. Each gives what the catalogue
// (`evalFunction`) and the general operator give: checked here call by call, over every
// kind of value an argument can have.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-50s expected %A  got %A" label expected got)

let cell v = { Value = v; Format = defaultFormat; Comment = None }
// A1 10, A2 0, A3 "hi", A4 TRUE, A5 FALSE, A6 #DIV/0!, A7 (empty), A8 "TRUE", A9 "12", A10 "", A11 #N/A
// B1 45000 (15 Mar 2023), B2 45000.75, B3 59 (28 Feb 1900), B4 60, B5 61, B6 -5, B7 2958465, B8 2958466, B9 1e12, B10 36585 (29 Feb 2000)
// C1 =A1*2 (a formula with its value), C2 =1/0
let cells =
    [ (0, 0), CellValue.Number 10.0; (1, 0), CellValue.Number 0.0; (2, 0), CellValue.String "hi"; (3, 0), CellValue.Boolean true
      (4, 0), CellValue.Boolean false; (5, 0), CellValue.Error Err.div0; (7, 0), CellValue.String "TRUE"; (8, 0), CellValue.String "12"
      (9, 0), CellValue.String ""; (10, 0), CellValue.Error Err.na
      (0, 1), CellValue.Number 45000.0; (1, 1), CellValue.Number 45000.75; (2, 1), CellValue.Number 59.0; (3, 1), CellValue.Number 60.0
      (4, 1), CellValue.Number 61.0; (5, 1), CellValue.Number -5.0; (6, 1), CellValue.Number 2958465.0; (7, 1), CellValue.Number 2958466.0
      (8, 1), CellValue.Number 1e12; (9, 1), CellValue.Number 36585.0
      (0, 2), CellValue.Formula ("=A1*2", CellValue.Number 20.0); (1, 2), CellValue.Formula ("=1/0", CellValue.Error Err.div0) ]
    |> List.map (fun ((row, col), value) -> { Row = row; Col = col }, cell value)
    |> Map.ofList
let model = { createEmpty 20 8 with Cells = cells }
let model1904 = { model with Date1904 = true }

let values = [ "A1"; "A2"; "A3"; "A4"; "A5"; "A6"; "A7"; "A8"; "A9"; "A10"; "A11"; "C1"; "C2"; "1"; "0"; "\"x\""; "TRUE"; "FALSE"; "A1>5"; "A1<5"; "A3=\"hi\""; "1/0"; "-A1"; "A1:A2"; "NA()" ]
let dates = [ "B1"; "B2"; "B3"; "B4"; "B5"; "B6"; "B7"; "B8"; "B9"; "B10"; "A3"; "A4"; "A6"; "A7"; "A9"; "45000"; "44985"; "61"; "0"; "\"3/15/2023\""; "B1+1"; "1/0" ]
let months = [ "0"; "1"; "-1"; "12"; "-13"; "1.9"; "-1.9"; "A7"; "A3"; "A4"; "A6"; "1e9"; "-1e9"; "120000"; "120001"; "\"2\"" ]

/// The call as the evaluator takes it, and as the catalogue alone answers it.
let both (m: SpreadsheetModel) (formula: string) =
    match parseFormula formula with
    | Ok (Function (name, args) as expr) ->
        let ctx = { Model = m; Workbook = None; Sheet = None; CurrentCell = Some { Row = 15; Col = 5 } }
        let outcome (run: unit -> CellValue) = try Choice1Of2 (run ()) with error -> Choice2Of2 (error.GetType().Name)
        // (On .NET the evaluator turns a date outside the calendar into #NUM!: so does this.)
        let general () = try evalFunction name args ctx with :? System.ArgumentOutOfRangeException -> CellValue.Error Err.num
        Some (outcome (fun () -> eval expr ctx), outcome general)
    | _ -> None
let mutable calls = 0
let same (m: SpreadsheetModel) (formula: string) =
    match both m formula with
    | Some (got, expected) ->
        calls <- calls + 1
        if got <> expected then (fail <- fail + 1; printfn "FAIL %-50s the catalogue gives %A, the evaluator %A" formula expected got)
    | None -> (fail <- fail + 1; printfn "FAIL %s is no call" formula)

for c in values do
    same model (sprintf "=IF(%s,1)" c)
    same model (sprintf "=IF(%s,\"yes\",A3)" c)
    same model (sprintf "=IF(A4,%s,2)" c)
    same model (sprintf "=IF(A5,2,%s)" c)
    same model (sprintf "=NOT(%s)" c)
    same model (sprintf "=IFERROR(%s,\"was an error\")" c)
    same model (sprintf "=IFERROR(A6,%s)" c)
    same model (sprintf "=AND(%s)" c)
    same model (sprintf "=OR(%s)" c)
    for d in values do
        same model (sprintf "=AND(%s,%s)" c d)
        same model (sprintf "=OR(%s,%s)" c d)
        same model (sprintf "=AND(A4,%s,%s)" c d)
        same model (sprintf "=OR(A5,%s,%s)" c d)
// (Calls with other numbers of arguments, and arguments that are ranges or arrays, go the general way: the same either way.)
for formula in [ "=IF(A4)"; "=IF(A4,1,2,3)"; "=NOT(A4,A5)"; "=IFERROR(A6)"; "=AND(A1:A5)"; "=OR(A1:A2,A7)"; "=AND(A1:A2>5)"; "=OR({0,0,1})"; "=MONTH(B1,2)"; "=EOMONTH(B1)" ] do
    same model formula
for m in [ model; model1904 ] do
    for d in dates do
        same m (sprintf "=YEAR(%s)" d)
        same m (sprintf "=MONTH(%s)" d)
        same m (sprintf "=DAY(%s)" d)
        for k in months do
            same m (sprintf "=EOMONTH(%s,%s)" d k)
            same m (sprintf "=EDATE(%s,%s)" d k)
// Every day of four years, a leap day among them, and the days around the ends of the calendar.
for m in [ model; model1904 ] do
    for serial in [ 61 .. 1600 ] @ [ 36520 .. 36900 ] @ [ 73040 .. 73120 ] @ [ 2958400 .. 2958465 ] do
        same m (sprintf "=YEAR(%d)" serial)
        same m (sprintf "=MONTH(%d)" serial)
        same m (sprintf "=DAY(%d)" serial)
eq "calls compared with the catalogue" (calls > 15000) true

// ---- a few of them by what they are -------------------------------------------------
let value (formula: string) = evaluateFormula formula model None
let n (x: float) = CellValue.Number x
eq "IF picks" (value "=IF(A1>5,\"big\",\"small\")", value "=IF(A1<5,\"big\",\"small\")") (CellValue.String "big", CellValue.String "small")
eq "IF without an else" (value "=IF(A2,1)") (CellValue.Boolean false)
eq "IF of an error" (value "=IF(A6,1,2)") (CellValue.Error Err.div0)
eq "IF of text that is no truth value" (value "=IF(A3,1,2)") (CellValue.Error Err.value)
eq "IF evaluates the branch it takes only" (value "=IF(A4,1,1/0)") (n 1.0)
eq "AND, OR" (value "=AND(A1>5,A1<20)", value "=AND(A1>5,A1<8)", value "=OR(A1<5,A1=10)", value "=OR(A2,A5)") (CellValue.Boolean true, CellValue.Boolean false, CellValue.Boolean true, CellValue.Boolean false)
eq "AND of nothing but an empty cell" (value "=AND(A7)") (CellValue.Error Err.value)
eq "AND leaves an empty cell out" (value "=AND(A7,A4)") (CellValue.Boolean true)
eq "AND gives the first error" (value "=AND(A11,A6)", value "=OR(A4,A6)") (CellValue.Error Err.na, CellValue.Error Err.div0)
eq "NOT" (value "=NOT(A2)", value "=NOT(A1)") (CellValue.Boolean true, CellValue.Boolean false)
eq "IFERROR" (value "=IFERROR(A6,7)", value "=IFERROR(A1,7)", value "=IFERROR(C2,\"e\")") (n 7.0, n 10.0, CellValue.String "e")
eq "YEAR, MONTH, DAY" (value "=YEAR(B1)", value "=MONTH(B1)", value "=DAY(B1)") (n 2023.0, n 3.0, n 15.0)
eq "... of a day with a time" (value "=YEAR(B2)", value "=MONTH(B2)", value "=DAY(B2)") (n 2023.0, n 3.0, n 15.0)
eq "... of 29 February 2000" (value "=YEAR(B10)", value "=MONTH(B10)", value "=DAY(B10)") (n 2000.0, n 2.0, n 29.0)
eq "... of text that is a number" (value "=MONTH(A9)") (n 1.0)
eq "EOMONTH, EDATE" (value "=EOMONTH(B1,0)", value "=EOMONTH(B1,-1)", value "=EDATE(B1,1)") (n 45016.0, n 44985.0, n 45031.0)
eq "EOMONTH of an error, by an error" (value "=EOMONTH(A6,1)", value "=EOMONTH(B1,A11)") (CellValue.Error Err.div0, CellValue.Error Err.na)
eq "a formula's value is read through" (value "=IF(C1=20,C1+1,0)") (n 21.0)

// ---- two numbers under an operator --------------------------------------------------
eq "numbers: + - * /" (applyBin "+" (n 2.0) (n 3.0), applyBin "-" (n 2.0) (n 3.0), applyBin "*" (n 2.0) (n 3.0), applyBin "/" (n 3.0) (n 2.0)) (n 5.0, n -1.0, n 6.0, n 1.5)
eq "numbers: a division by zero" (applyBin "/" (n 1.0) (n 0.0), applyBin "/" (n 0.0) (n 0.0)) (CellValue.Error Err.div0, CellValue.Error Err.div0)
eq "numbers: beyond the range is #NUM!" (applyBin "*" (n 1e200) (n 1e200), applyBin "+" (n 1.7e308) (n 1.7e308), applyBin "^" (n 10.0) (n 400.0), applyBin "^" (n -8.0) (n 0.5)) (CellValue.Error Err.num, CellValue.Error Err.num, CellValue.Error Err.num, CellValue.Error Err.num)
eq "numbers: a power" (applyBin "^" (n 2.0) (n 10.0)) (n 1024.0)
eq "numbers: joined as text" (applyBin "&" (n 1.0) (n 2.5)) (CellValue.String "12.5")
eq "numbers: compared"
    [ for op in [ "="; "<>"; "<"; ">"; "<="; ">=" ] -> [ for a, b in [ 1.0, 2.0; 2.0, 2.0; 3.0, 2.0; -0.0, 0.0 ] -> applyBin op (n a) (n b) = CellValue.Boolean true ] ]
    [ [ false; true; false; true ]; [ true; false; true; false ]; [ true; false; false; false ]; [ false; false; true; false ]; [ true; true; false; true ]; [ false; true; true; true ] ]
eq "an operator there is none of" (applyBin "?" (n 1.0) (n 2.0)) (CellValue.Error Err.value)
eq "a number and something else, as before" (applyBin "+" (n 1.0) (CellValue.Boolean true), applyBin "+" (n 1.0) (CellValue.String "x"), applyBin "=" (n 0.0) CellValue.Empty, applyBin "<" (n 5.0) (CellValue.String "a"), applyBin "+" (CellValue.Error Err.na) (n 1.0))
    (n 2.0, CellValue.Error Err.value, CellValue.Boolean true, CellValue.Boolean true, CellValue.Error Err.na)
eq "one operator, two, a long chain" (value "=A1-3", value "=A1-3-2", value "=A1-3-2-1-1", value "=2^3^2", value "=A1-(3-2)") (n 7.0, n 5.0, n 3.0, n 64.0, n 9.0)
eq "a chain is evaluated left to right, an error in it is the result" (value "=A1+A6+A11", value "=A11+A6", value "=A1&A3&A4") (CellValue.Error Err.div0, CellValue.Error Err.na, CellValue.String "10hiTRUE")

printfn "Stage 53 common-function tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
