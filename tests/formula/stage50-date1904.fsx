#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Input

// The 1904 date system (Model.date1904): serial 0 is 1904-01-01. The serials
// of such a workbook are kept as they are; number formats, the date
// functions and typed dates work in the workbook's system.

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)
let eq label (got: 'a) (expected: 'a) =
    ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let model = createEmpty 10 10 |> setCell { Row = 0; Col = 0 } { Value = CellValue.Number 42000.0; Format = defaultFormat; Comment = None }
let value formula = evaluateFormula formula model None
let N x = CellValue.Number x
let format code serial = Rbt.Spreadsheet.CellText.formatNumberWithCode code serial

// 42000 is 2014-12-27 in the 1900 system and 2018-12-28 (a Friday) in the 1904 system.
let check (label: string) (year: int, month: int, day: int) (weekday: float) (nextMonth: float) =
    eq (label + ": a serial shown as a date") (format "yyyy-mm-dd" 42000.0) (sprintf "%04d-%02d-%02d" year month day)
    eq (label + ": ... with the weekday") (format "ddd" 42000.0) (if weekday = 6.0 then "Fri" else "Sat")
    eq (label + ": DATE") (value (sprintf "=DATE(%d,%d,%d)" year month day)) (N 42000.0)
    eq (label + ": YEAR, MONTH, DAY") (value "=YEAR(A1)*10000+MONTH(A1)*100+DAY(A1)") (N (float (year * 10000 + month * 100 + day)))
    eq (label + ": WEEKDAY") (value "=WEEKDAY(A1)") (N weekday)
    eq (label + ": EDATE") (value "=EDATE(A1,1)") (N nextMonth)
    eq (label + ": DATEVALUE") (value (sprintf "=DATEVALUE(\"%04d-%02d-%02d\")" year month day)) (N 42000.0)
    eq (label + ": TEXT") (value "=TEXT(A1,\"yyyy-mm-dd\")") (CellValue.String (sprintf "%04d-%02d-%02d" year month day))
    eq (label + ": a typed date") ((parseInputIn 2024 (sprintf "%04d-%02d-%02d" year month day)).Value) (N 42000.0)
    eq (label + ": dateSerial") (dateSerial year month day) 42000.0
    // Monday to Friday of that week (or the week before, in the 1900 system).
    eq (label + ": NETWORKDAYS over the five days before") (value "=NETWORKDAYS(A1-4,A1)") (N (if weekday = 6.0 then 5.0 else 4.0))
    eq (label + ": WORKDAY") (value "=WORKDAY(A1,1)") (N (if weekday = 6.0 then 42003.0 else 42002.0))

date1904 <- true
check "1904" (2018, 12, 28) 6.0 42031.0
eq "1904: serial 0 is 1904-01-01" (format "yyyy-mm-dd" 0.0) "1904-01-01"
eq "1904: serial 59 and 60 are the last of February and the first of March (1904 is a leap year: 60 is Mar 1)" (format "m/d" 59.0, format "m/d" 60.0) ("2/29", "3/1")
eq "1904: DATE(1904,1,1)" (value "=DATE(1904,1,1)") (N 0.0)
eq "1904: a date before 1904 is not typed as a date" (match (parseInputIn 2024 "1903-06-01").Value with CellValue.Number _ -> false | _ -> true) true
eq "1904: times are what they are" (format "h:mm" 0.75) "18:00"
date1904 <- false
check "1900" (2014, 12, 27) 7.0 42031.0
eq "1900: serial 61 is 1900-03-01" (format "yyyy-mm-dd" 61.0) "1900-03-01"

printfn "\nstage50 (1904 date system): %d passed, %d failed" pass fail
if fail > 0 then exit 1
