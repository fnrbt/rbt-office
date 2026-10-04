#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open System
open Rbt.Spreadsheet.Input

// Typed input → cell value + number format (Input.parseInput), the text shown
// for editing (Input.editText), and autofill (Input.fillCells).

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)
let eq label (got: 'a) (expected: 'a) =
    ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let N x = CellValue.Number x
let S x = CellValue.String x
let parse text = let p = parseInputIn 2024 text in p.Value, p.NumberFormat
let near (a: float) (b: float) = abs (a - b) < 1e-9
let numberOf text = match (parseInputIn 2024 text).Value with CellValue.Number n -> n | _ -> nan
let formatOf text = (parseInputIn 2024 text).NumberFormat
let serial y m d = dateSerial y m d

// --- numbers ---------------------------------------------------------------
eq "int"              (parse "42")            (N 42.0, None)
eq "decimal"          (parse "3.14")          (N 3.14, None)
eq "negative"         (parse "-7.5")          (N -7.5, None)
eq "plus"             (parse "+8")            (N 8.0, None)
eq "leading-dot"      (parse ".5")            (N 0.5, None)
eq "trailing-dot"     (parse "5.")            (N 5.0, None)
eq "spaces"           (parse "  12  ")        (N 12.0, None)
eq "exponent"         (parse "1.5e3")         (N 1500.0, None)
eq "exponent-neg"     (parse "2E-2")          (N 0.02, None)
eq "thousands"        (parse "1,234")         (N 1234.0, Some "#,##0")
eq "thousands-dec"    (parse "1,234,567.89")  (N 1234567.89, Some "#,##0.00")
eq "accounting-neg"   (parse "(1,234)")       (N -1234.0, Some "#,##0")
eq "paren-neg"        (parse "(12.5)")        (N -12.5, None)
eq "bad-grouping"     (parse "1,23")          (S "1,23", None)
eq "bad-grouping-2"   (parse "12,3456")       (S "12,3456", None)
eq "double-sign"      (parse "--5")           (S "--5", None)
eq "comma-list"       (parse "1,2,3")         (S "1,2,3", None)
eq "not-a-number"     (parse "12abc")         (S "12abc", None)
eq "culture-free"     (parse "1.234,56")      (S "1.234,56", None)
// --- percent ---------------------------------------------------------------
ok "percent"          (near (numberOf "12%") 0.12 && formatOf "12%" = Some "0%") (sprintf "%A" (parse "12%"))
ok "percent-decimals" (near (numberOf "12.5%") 0.125 && formatOf "12.5%" = Some "0.00%") (sprintf "%A" (parse "12.5%"))
ok "percent-space"    (near (numberOf "7 %") 0.07) (sprintf "%A" (parse "7 %"))
ok "percent-negative" (near (numberOf "-3%") -0.03) (sprintf "%A" (parse "-3%"))
eq "percent-alone"    (parse "%")             (S "%", None)
// --- currency --------------------------------------------------------------
eq "currency"         (parse "$1,234.50")     (N 1234.5, Some "$#,##0.00")
eq "currency-int"     (parse "$5")            (N 5.0, Some "$#,##0")
eq "currency-neg"     (parse "-$5.25")        (N -5.25, Some "$#,##0.00")
eq "currency-neg-2"   (parse "$-5.25")        (N -5.25, Some "$#,##0.00")
eq "currency-paren"   (parse "($1,000)")      (N -1000.0, Some "$#,##0")
eq "currency-euro"    (parse "€12.50")        (N 12.5, Some "\"€\"#,##0.00")
eq "currency-alone"   (parse "$")             (S "$", None)
eq "currency-text"    (parse "$abc")          (S "$abc", None)
eq "currency-shown"   (Rbt.Spreadsheet.CellText.formatNumberWithCode "$#,##0.00" 1234.5) "$1,234.50"
eq "percent-shown"    (Rbt.Spreadsheet.CellText.formatNumberWithCode "0%" 0.12) "12%"
// --- fractions -------------------------------------------------------------
eq "fraction"         (parse "1 1/2")         (N 1.5, Some "# ?/?")
eq "fraction-zero"    (parse "0 3/4")         (N 0.75, Some "# ?/?")
eq "fraction-div0"    (parse "1 1/0")         (S "1 1/0", None)
// --- dates -----------------------------------------------------------------
eq "date-us"          (parse "3/4/2024")      (N (serial 2024 3 4), Some "m/d/yyyy")
eq "date-us-serial"   (serial 2024 3 4)       45355.0
eq "date-us-2digit"   (parse "3/4/24")        (N (serial 2024 3 4), Some "m/d/yyyy")
eq "date-window"      (parse "1/1/30", parse "12/31/29") ((N (serial 1930 1 1), Some "m/d/yyyy"), (N (serial 2029 12 31), Some "m/d/yyyy"))
eq "date-dashes"      (parse "3-4-2024")      (N (serial 2024 3 4), Some "m/d/yyyy")
eq "date-iso"         (parse "2024-03-04")    (N (serial 2024 3 4), Some "yyyy-mm-dd")
eq "date-iso-slash"   (parse "2024/3/4")      (N (serial 2024 3 4), Some "yyyy-mm-dd")
eq "date-d-mmm"       (parse "4-Mar")         (N (serial 2024 3 4), Some "d-mmm")
eq "date-mmm-d"       (parse "Mar 4")         (N (serial 2024 3 4), Some "d-mmm")
eq "date-m/d"         (parse "3/4")           (N (serial 2024 3 4), Some "d-mmm")
eq "date-d-mmm-yy"    (parse "4-Mar-24")      (N (serial 2024 3 4), Some "d-mmm-yy")
eq "date-d-mmm-yyyy"  (parse "4 March 2024")  (N (serial 2024 3 4), Some "d-mmm-yy")
eq "date-long"        (parse "March 4, 2024") (N (serial 2024 3 4), Some "d-mmm-yy")
eq "date-month-year"  (parse "Mar 2024")      (N (serial 2024 3 1), Some "mmm-yy")
eq "date-m/yyyy"      (parse "3/2024")        (N (serial 2024 3 1), Some "mmm-yy")
eq "date-sept"        (parse "5 Sept 2023")   (N (serial 2023 9 5), Some "d-mmm-yy")
eq "date-leap"        (parse "2/29/2024")     (N (serial 2024 2 29), Some "m/d/yyyy")
eq "date-not-leap"    (parse "2/29/2023")     (S "2/29/2023", None)
eq "date-bad-month"   (parse "13/1/2024")     (S "13/1/2024", None)
eq "date-bad-day"     (parse "4/31/2024")     (S "4/31/2024", None)
eq "date-too-old"     (parse "1/1/1800")      (S "1/1/1800", None)
eq "date-mixed-sep"   (parse "3/4-2024")      (S "3/4-2024", None)
eq "date-word"        (parse "March")         (S "March", None)
eq "date-not-month"   (parse "4-Foo")         (S "4-Foo", None)
eq "date-epoch"       (serial 1900 3 1, serial 1970 1 1) (61.0, 25569.0)
eq "date-engine-agrees" (Rbt.Spreadsheet.Formula.evaluateFormula "=DATE(2024,3,4)" (createEmpty 5 5) None) (N (serial 2024 3 4))
eq "date-shown"       (Rbt.Spreadsheet.CellText.formatNumberWithCode "d-mmm" (serial 2024 3 4)) "4-Mar"
// --- times -----------------------------------------------------------------
ok "time-24h"         (near (numberOf "14:30") (14.5 / 24.0) && formatOf "14:30" = Some "h:mm") (sprintf "%A" (parse "14:30"))
ok "time-seconds"     (near (numberOf "1:02:03") ((3600.0 + 123.0) / 86400.0) && formatOf "1:02:03" = Some "h:mm:ss") (sprintf "%A" (parse "1:02:03"))
ok "time-pm"          (near (numberOf "2:30 PM") (14.5 / 24.0) && formatOf "2:30 PM" = Some "h:mm AM/PM") (sprintf "%A" (parse "2:30 PM"))
ok "time-am-12"       (near (numberOf "12:15 am") (0.25 / 24.0)) (sprintf "%A" (parse "12:15 am"))
ok "time-bare-pm"     (near (numberOf "2 PM") (14.0 / 24.0)) (sprintf "%A" (parse "2 PM"))
ok "time-over-24"     (near (numberOf "25:00") (25.0 / 24.0) && formatOf "25:00" = Some "[h]:mm:ss") (sprintf "%A" (parse "25:00"))
eq "time-bad-minute"  (parse "1:75")          (S "1:75", None)
eq "time-bad-13pm"    (parse "13:00 PM")      (S "13:00 PM", None)
eq "time-ratio-text"  (parse "a:b")           (S "a:b", None)
ok "datetime"         (near (numberOf "3/4/2024 14:30") (serial 2024 3 4 + 14.5 / 24.0) && formatOf "3/4/2024 14:30" = Some "m/d/yyyy h:mm") (sprintf "%A" (parse "3/4/2024 14:30"))
ok "datetime-iso-pm"  (near (numberOf "2024-03-04 2:30 PM") (serial 2024 3 4 + 14.5 / 24.0)) (sprintf "%A" (parse "2024-03-04 2:30 PM"))
// --- booleans, formulas, forced text, errors, blanks -----------------------
eq "bool-true"        (parse "true")          (CellValue.Boolean true, None)
eq "bool-false"       (parse "FALSE")         (CellValue.Boolean false, None)
eq "formula"          (parse "=A1+1")         (CellValue.Formula ("=A1+1", CellValue.Empty), None)
eq "formula-spaces"   (parse "  =SUM(A1:A3)") (CellValue.Formula ("=SUM(A1:A3)", CellValue.Empty), None)
eq "lone-equals"      (parse "=")             (S "=", None)
eq "apostrophe"       (parse "'123")          (S "123", None)
eq "apostrophe-formula" (parse "'=A1")        (S "=A1", None)
eq "apostrophe-date"  (parse "'3/4/2024")     (S "3/4/2024", None)
eq "apostrophe-alone" (parse "'")             (S "", None)
eq "error-literal"    (parse "#N/A")          (CellValue.Error "N/A", None)
eq "error-div0"       (parse "#div/0!")       (CellValue.Error "DIV/0!", None)
eq "empty"            (parse "")              (CellValue.Empty, None)
eq "blank"            (parse "   ")           (CellValue.Empty, None)
eq "text-kept"        (parse " hello ")       (S " hello ", None)
eq "null"             (parseInputIn 2024 null).Value CellValue.Empty

// --- the cell that results --------------------------------------------------
let general = cellFromInput 2024 None "12%"
eq "cell-format-applied" (general.Value, general.Format.NumberFormat) (N 0.12, Some "0%")
let formatted = Some { Value = N 1.0; Format = { defaultFormat with NumberFormat = Some "0.000"; Bold = true }; Comment = Some { Text = "note"; Author = None } }
let kept = cellFromInput 2024 formatted "$5"
eq "cell-format-kept"    (kept.Value, kept.Format.NumberFormat, kept.Format.Bold, kept.Comment.IsSome) (N 5.0, Some "0.000", true, true)
let generalCode = Some { Value = N 1.0; Format = { defaultFormat with NumberFormat = Some "General" }; Comment = None }
eq "cell-general-code"   (cellFromInput 2024 generalCode "3/4/2024").Format.NumberFormat (Some "m/d/yyyy")
eq "cell-plain-number"   (cellFromInput 2024 None "42").Format.NumberFormat None

// --- text for editing round-trips ------------------------------------------
let roundTrip text =
    let cell = cellFromInput 2024 None text
    let shown = editText (Some cell)
    shown, (cellFromInput 2024 (Some cell) shown).Value = cell.Value
eq "edit-date"      (roundTrip "4-Mar-2024")     ("3/4/2024", true)
eq "edit-datetime"  (roundTrip "3/4/2024 14:30") ("3/4/2024 14:30:00", true)
eq "edit-time"      (roundTrip "14:30")          ("14:30:00", true)
eq "edit-percent"   (roundTrip "7%")             ("7%", true)
eq "edit-currency"  (roundTrip "$1,234.50")      ("1234.5", true)
eq "edit-formula"   (roundTrip "=A1*2")          ("=A1*2", true)
eq "edit-text"      (roundTrip "hello")          ("hello", true)
eq "edit-forced"    (roundTrip "'123")           ("'123", true)
eq "edit-forced-date" (roundTrip "'4-Mar")       ("'4-Mar", true)
eq "edit-bool"      (roundTrip "true")           ("TRUE", true)
eq "edit-error"     (roundTrip "#N/A")           ("#N/A", true)
eq "edit-empty"     (editText None)              ""

// --- autofill ---------------------------------------------------------------
let cell v = { Value = v; Format = defaultFormat; Comment = None }
let at r c = { Row = r; Col = c }
let sheet cells = { createEmpty 100 26 with Cells = Map.ofList [ for (r, c, v) in cells -> at r c, cell v ] }
let column c rows (m: SpreadsheetModel) = [ for r in rows -> getCellValue (at r c) m ]
let rowOf r cols (m: SpreadsheetModel) = [ for c in cols -> getCellValue (at r c) m ]
let F f = CellValue.Formula (f, CellValue.Empty)
let down cells count = fillCells (0, 0, List.length cells - 1, 0) FillDown count true (sheet [ for (i, v) in List.indexed cells -> i, 0, v ])
eq "fill-123"          (down [ N 1.0; N 2.0; N 3.0 ] 3 |> column 0 [ 3 .. 5 ])         [ N 4.0; N 5.0; N 6.0 ]
eq "fill-step-5"       (down [ N 10.0; N 15.0 ] 3 |> column 0 [ 2 .. 4 ])              [ N 20.0; N 25.0; N 30.0 ]
eq "fill-descending"   (down [ N 9.0; N 7.0 ] 2 |> column 0 [ 2; 3 ])                  [ N 5.0; N 3.0 ]
eq "fill-decimal-step" (down [ N 0.1; N 0.2 ] 2 |> column 0 [ 2; 3 ])                  [ N 0.3; N 0.4 ]
eq "fill-trend"        (down [ N 1.0; N 2.0; N 4.0 ] 1 |> column 0 [ 3 ])              [ N (Math.Round(7.0 / 3.0 - 1.5 + 1.5 * 3.0, 10)) ]
eq "fill-single-copy"  (down [ N 7.0 ] 3 |> column 0 [ 1 .. 3 ])                       [ N 7.0; N 7.0; N 7.0 ]
eq "fill-weekdays"     (down [ S "Mon" ] 7 |> column 0 [ 1 .. 7 ])                     [ S "Tue"; S "Wed"; S "Thu"; S "Fri"; S "Sat"; S "Sun"; S "Mon" ]
eq "fill-weekday-step" (down [ S "Monday"; S "Wednesday" ] 3 |> column 0 [ 2 .. 4 ])   [ S "Friday"; S "Sunday"; S "Tuesday" ]
eq "fill-months"       (down [ S "November" ] 3 |> column 0 [ 1 .. 3 ])                [ S "December"; S "January"; S "February" ]
eq "fill-month-abbr"   (down [ S "JAN"; S "APR" ] 3 |> column 0 [ 2 .. 4 ])            [ S "JUL"; S "OCT"; S "JAN" ]
eq "fill-lowercase"    (down [ S "fri" ] 2 |> column 0 [ 1; 2 ])                       [ S "sat"; S "sun" ]
eq "fill-text-number"  (down [ S "Item 1" ] 3 |> column 0 [ 1 .. 3 ])                  [ S "Item 2"; S "Item 3"; S "Item 4" ]
eq "fill-quarter-step" (down [ S "Q1"; S "Q3" ] 2 |> column 0 [ 2; 3 ])                [ S "Q5"; S "Q7" ]
eq "fill-leading-zero" (down [ S "Inv-008" ] 3 |> column 0 [ 1 .. 3 ])                 [ S "Inv-009"; S "Inv-010"; S "Inv-011" ]
eq "fill-plain-text"   (down [ S "a"; S "b" ] 3 |> column 0 [ 2 .. 4 ])                [ S "a"; S "b"; S "a" ]
eq "fill-mixed-repeats" (down [ N 1.0; S "x" ] 2 |> column 0 [ 2; 3 ])                 [ N 1.0; S "x" ]
let dated = { createEmpty 50 10 with Cells = Map.ofList [ at 0 0, { cell (N (serial 2024 2 28)) with Format = { defaultFormat with NumberFormat = Some "m/d/yyyy" } } ] }
let datedFilled = fillCells (0, 0, 0, 0) FillDown 2 true dated
eq "fill-date-days"    (column 0 [ 1; 2 ] datedFilled)                                 [ N (serial 2024 2 29); N (serial 2024 3 1) ]
eq "fill-date-format"  ((getCell (at 2 0) datedFilled).Value.Format.NumberFormat)      (Some "m/d/yyyy")
let weekly = fillCells (0, 0, 1, 0) FillDown 2 true { dated with Cells = dated.Cells |> Map.add (at 1 0) (cell (N (serial 2024 3 6))) }
eq "fill-date-step"    (column 0 [ 2; 3 ] weekly)                                      [ N (serial 2024 3 13); N (serial 2024 3 20) ]
// Formulas shift; the pattern of a mixed block repeats.
let formulas = fillCells (0, 1, 0, 1) FillDown 2 true (sheet [ 0, 0, N 1.0; 0, 1, F "=A1*$B$9+A$1" ])
eq "fill-formula-down" (column 1 [ 1; 2 ] formulas)                                    [ F "=A2*$B$9+A$1"; F "=A3*$B$9+A$1" ]
let across = fillCells (0, 0, 1, 1) FillRight 3 true (sheet [ 0, 0, N 1.0; 0, 1, N 2.0; 1, 0, F "=A1+1"; 1, 1, F "=B1+1" ])
eq "fill-right-series" (rowOf 0 [ 2 .. 4 ] across)                                     [ N 3.0; N 4.0; N 5.0 ]
eq "fill-right-formula" (rowOf 1 [ 2 .. 4 ] across)                                    [ F "=C1+1"; F "=D1+1"; F "=E1+1" ]
let upward = fillCells (5, 0, 6, 0) FillUp 3 true (sheet [ 5, 0, N 10.0; 6, 0, N 12.0 ])
eq "fill-up"           (column 0 [ 2 .. 4 ] upward)                                    [ N 4.0; N 6.0; N 8.0 ]
let leftward = fillCells (0, 5, 0, 5) FillLeft 2 true (sheet [ 0, 5, S "Wed" ])
eq "fill-left"         (rowOf 0 [ 3; 4 ] leftward)                                     [ S "Mon"; S "Tue" ]
// Ctrl+D / Ctrl+R copy rather than continue.
let copied = fillCells (0, 0, 0, 1) FillDown 2 false (sheet [ 0, 0, N 1.0; 0, 1, F "=A1*2" ])
eq "fill-copy-values"  (column 0 [ 1; 2 ] copied)                                      [ N 1.0; N 1.0 ]
eq "fill-copy-formula" (column 1 [ 1; 2 ] copied)                                      [ F "=A2*2"; F "=A3*2" ]
eq "fill-blank-clears" (fillCells (0, 0, 0, 0) FillDown 1 true (sheet [ 1, 0, N 9.0 ]) |> column 0 [ 1 ]) [ CellValue.Empty ]
eq "fill-targets"      (fillTargets (0, 0, 1, 1) FillDown 1) [ at 2 0; at 2 1 ]
eq "fill-towards"      [ fillTowards (2, 2, 3, 3) 6 3; fillTowards (2, 2, 3, 3) 0 2; fillTowards (2, 2, 3, 3) 3 7; fillTowards (2, 2, 3, 3) 2 0
                         fillTowards (2, 2, 3, 3) 3 2; fillTowards (2, 2, 3, 3) 5 4 ]
                       [ Some (FillDown, 3); Some (FillUp, 2); Some (FillRight, 4); Some (FillLeft, 2); None; Some (FillDown, 2) ]

printfn ""
printfn "Stage 27 input tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
