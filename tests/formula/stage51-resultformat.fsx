#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
module Ops = Rbt.Spreadsheet.SheetOps

// The number format a formula's cell takes from what the formula reads when
// it is entered into a General cell (Excel's rules): references, + - * /,
// functions that pass values on, date / time / payment / rate functions.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s expected %A  got %A" label expected got)

let date, time, stamp, percent, money, fixed2 = "m/d/yyyy", "h:mm", "m/d/yyyy h:mm", "0.0%", "$#,##0.00", "0.00"
// A1 date, A2 date, A3 time, A4 time, B1 percent, B2 percent, C1 money, C2 money, D1 0.00, E1 General, F1 text, G1 date and time.
let formats =
    Map.ofList [ (0, 0), date; (1, 0), date; (2, 0), time; (3, 0), time; (0, 1), percent; (1, 1), percent; (0, 2), money; (1, 2), money
                 (0, 3), fixed2; (0, 4), "General"; (0, 5), "@"; (0, 6), stamp ]
let other = Map.ofList [ (0, 0), money ]
let formatAt (sheet: string option) (row: int) (col: int) : string option =
    match sheet with
    | None -> Map.tryFind (row, col) formats
    | Some name when name = "Other" -> Map.tryFind (row, col) other
    | Some _ -> None
let of' (formula: string) = Ops.inheritedNumberFormat formatAt formula

// ---- references -------------------------------------------------------------------
eq "a date cell" (of' "=A1") (Some date)
eq "a General cell" (of' "=E1") None
eq "a text cell says nothing" (of' "=F1") None
eq "an empty cell" (of' "=Z9") None
eq "a cell of another sheet" (of' "=Other!A1") (Some money)
eq "a constant" (of' "=1+2") None
eq "without the equals sign" (of' "A1+1") (Some date)

// ---- + and - ----------------------------------------------------------------------
eq "a date plus a number" (of' "=A1+1") (Some date)
eq "a number plus a date" (of' "=7+A1") (Some date)
eq "a date less a number" (of' "=A1-30") (Some date)
eq "a date less a date is a number of days" (of' "=A2-A1") None
eq "a date and time less a date" (of' "=G1-A1") None
eq "a time less a time is a time" (of' "=A4-A3") (Some time)
eq "a date plus a time is a date" (of' "=A1+A3") (Some date)
eq "an amount plus an amount" (of' "=C1+C2") (Some money)
eq "the first format of two" (of' "=C1+D1") (Some money)
eq "... in the order of the formula" (of' "=D1+C1") (Some fixed2)
eq "a General cell first: the next one's" (of' "=E1+C1") (Some money)
eq "a days count plus a date is a date" (of' "=D1+A1") (Some date)
eq "a negated amount" (of' "=-C1") (Some money)
eq "parentheses" (of' "=(A1+1)-2") (Some date)

// ---- * and / ----------------------------------------------------------------------
eq "a percentage doubled" (of' "=B1*2") (Some percent)
eq "twice a percentage" (of' "=2*B1") (Some percent)
eq "an amount times a rate is an amount" (of' "=C1*B1") (Some money)
eq "a rate times an amount is an amount" (of' "=B1*C1") (Some money)
eq "an amount halved" (of' "=C1/2") (Some money)
eq "an amount over an amount is a ratio" (of' "=C1/C2") None
eq "a percentage over a percentage" (of' "=B1/B2") None
eq "an amount over a rate" (of' "=C1/B1") (Some money)
eq "a date times two is no date" (of' "=A1*2") None
eq "a difference of dates times an amount" (of' "=(A2-A1)*C1") (Some money)
eq "a power, a comparison, a text" (of' "=C1^2", of' "=C1>C2", of' "=C1&\"x\"") (None, None, None)

// ---- functions that pass values on ------------------------------------------------
eq "SUM over amounts" (of' "=SUM(C1:C2)") (Some money)
eq "SUM over percentages" (of' "=SUM(B1:B2)") (Some percent)
eq "SUM: the first argument that has a format" (of' "=SUM(E1,C1:C2)") (Some money)
eq "AVERAGE, MIN, MAX, MEDIAN" (of' "=AVERAGE(C1:C2)", of' "=MIN(A1:A2)", of' "=MAX(A1:A2)", of' "=MEDIAN(B1:B2)") (Some money, Some date, Some date, Some percent)
eq "ROUND, ABS, INT" (of' "=ROUND(C1,0)", of' "=ABS(C1)", of' "=INT(D1)") (Some money, Some money, Some fixed2)
eq "IF: the value, then the other" (of' "=IF(E1>0,C1,0)", of' "=IF(E1>0,0,A1)") (Some money, Some date)
eq "IFERROR" (of' "=IFERROR(C1/E1,0)") (Some money)
eq "SUMIF: the range summed" (of' "=SUMIF(A1:A2,\">0\",C1:C2)", of' "=SUMIF(C1:C2,\">0\")") (Some money, Some money)
eq "SUMIFS, INDEX, XLOOKUP" (of' "=SUMIFS(C1:C2,A1:A2,\">0\")", of' "=INDEX(B1:B2,2)", of' "=XLOOKUP(1,E1:E2,C1:C2)") (Some money, Some percent, Some money)
eq "SUBTOTAL sums, and counts" (of' "=SUBTOTAL(9,C1:C2)", of' "=SUBTOTAL(3,C1:C2)") (Some money, None)
eq "nested" (of' "=ROUND(SUM(C1:C2)*B1,2)") (Some money)
eq "counts and parts of a date are numbers" (of' "=COUNT(C1:C2)", of' "=COUNTA(A1:A2)", of' "=YEAR(A1)", of' "=DAYS(A2,A1)", of' "=LEN(F1)") (None, None, None, None, None)

// ---- functions with a format of their own -----------------------------------------
eq "TODAY, DATE, EDATE, EOMONTH" (of' "=TODAY()", of' "=DATE(2026,1,31)", of' "=EDATE(E1,1)", of' "=EOMONTH(E1,0)") (Some "m/d/yyyy", Some "m/d/yyyy", Some "m/d/yyyy", Some "m/d/yyyy")
eq "TODAY plus a week" (of' "=TODAY()+7") (Some "m/d/yyyy")
eq "NOW" (of' "=NOW()") (Some "m/d/yyyy h:mm")
eq "TIME" (of' "=TIME(9,30,0)") (Some "h:mm AM/PM")
eq "PMT, NPV" (of' "=PMT(0.05/12,60,10000)", of' "=NPV(0.1,E1:E5)") (Some "$#,##0.00_);[Red]($#,##0.00)", Some "$#,##0.00_);[Red]($#,##0.00)")
eq "IRR, RATE" (of' "=IRR(E1:E5)", of' "=RATE(60,-100,5000)") (Some "0%", Some "0%")

// ---- what does not parse ----------------------------------------------------------
eq "a formula that does not parse" (of' "=SUM(") None

// What a number shows in a narrow cell, as Excel 2019 shows it (Calibri 11).
let shown (width: float) (code: string option) (n: float) =
    Rbt.Spreadsheet.CellText.cellDisplayText None width { defaultFormat with NumberFormat = code } (CellValue.Number n)
eq "General in a column 5.71 characters wide (39.97 px): 4E+05" (shown (5.7109375 * 7.0) None 400000.0) "4E+05"
eq "General: 1E+08, 3.142, ####" [ shown 40.0 None 123456789.0; shown 40.0 None 3.14159265358979; shown 33.0 None 400000.0 ] [ "1E+08"; "3.142"; "####" ]

printfn "Stage 51 result-format tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
