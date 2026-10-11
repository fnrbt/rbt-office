#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

// Grid layout (0-indexed):
//   A1=10  A2=20  A3="hi"  A4=TRUE  A5=-3
//   B1=3   B2=5   B3=5     B4=8     B5=15
//   D1=1   E1="one"        (lookup table D1:E3)
//   D2=2   E2="two"
//   D3=3   E3="three"
//   F1..F5 = 5,15,25,5,35   (criteria range)
//   G1..G5 = 100,200,300,400,500  (parallel sum range)
let cell v = { Value = v; Format = defaultFormat; Comment = None }
let cells =
    [ { Row=0; Col=0 }, cell (CellValue.Number 10.0)
      { Row=1; Col=0 }, cell (CellValue.Number 20.0)
      { Row=2; Col=0 }, cell (CellValue.String "hi")
      { Row=3; Col=0 }, cell (CellValue.Boolean true)
      { Row=4; Col=0 }, cell (CellValue.Number -3.0)
      { Row=0; Col=1 }, cell (CellValue.Number 3.0)
      { Row=1; Col=1 }, cell (CellValue.Number 5.0)
      { Row=2; Col=1 }, cell (CellValue.Number 5.0)
      { Row=3; Col=1 }, cell (CellValue.Number 8.0)
      { Row=4; Col=1 }, cell (CellValue.Number 15.0)
      { Row=0; Col=3 }, cell (CellValue.Number 1.0)
      { Row=1; Col=3 }, cell (CellValue.Number 2.0)
      { Row=2; Col=3 }, cell (CellValue.Number 3.0)
      { Row=0; Col=4 }, cell (CellValue.String "one")
      { Row=1; Col=4 }, cell (CellValue.String "two")
      { Row=2; Col=4 }, cell (CellValue.String "three")
      { Row=0; Col=5 }, cell (CellValue.Number 5.0)
      { Row=1; Col=5 }, cell (CellValue.Number 15.0)
      { Row=2; Col=5 }, cell (CellValue.Number 25.0)
      { Row=3; Col=5 }, cell (CellValue.Number 5.0)
      { Row=4; Col=5 }, cell (CellValue.Number 35.0)
      { Row=0; Col=6 }, cell (CellValue.Number 100.0)
      { Row=1; Col=6 }, cell (CellValue.Number 200.0)
      { Row=2; Col=6 }, cell (CellValue.Number 300.0)
      { Row=3; Col=6 }, cell (CellValue.Number 400.0)
      { Row=4; Col=6 }, cell (CellValue.Number 500.0)
      // horizontal lookup table D7:F8
      { Row=6; Col=3 }, cell (CellValue.Number 1.0)
      { Row=6; Col=4 }, cell (CellValue.Number 2.0)
      { Row=6; Col=5 }, cell (CellValue.Number 3.0)
      { Row=7; Col=3 }, cell (CellValue.String "x")
      { Row=7; Col=4 }, cell (CellValue.String "y")
      { Row=7; Col=5 }, cell (CellValue.String "z") ]
    |> Map.ofList
let model = { createEmpty 100 26 with Cells = cells }

let mutable pass = 0
let mutable fail = 0
let check label formula (expected: CellValue) =
    let got = evaluateFormula formula model None
    if got = expected then
        pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-22s  %-34s  expected %A  got %A" label formula expected got

// Approximate compare for float results.
let checkApprox label formula (expected: float) =
    match evaluateFormula formula model None with
    | CellValue.Number got when abs (got - expected) < 1e-9 -> pass <- pass + 1
    | got ->
        fail <- fail + 1
        printfn "FAIL %-22s  %-34s  expected ~%f  got %A" label formula expected got

let n x = CellValue.Number x
let s x = CellValue.String x
let b x = CellValue.Boolean x
let e x = CellValue.Error x

// ----- aggregates -----
check "product"      "=PRODUCT(B1:B2)"        (n 15.0)
check "median-odd"   "=MEDIAN(1,2,3,4,5)"     (n 3.0)
check "median-even"  "=MEDIAN(1,2,3,4)"       (n 2.5)
check "counta"       "=COUNTA(A1:A5)"         (n 5.0)
check "countblank"   "=COUNTBLANK(A1:A10)"    (n 5.0)
check "large"        "=LARGE(F1:F5,1)"        (n 35.0)
check "large-2"      "=LARGE(F1:F5,2)"        (n 25.0)
check "small"        "=SMALL(F1:F5,1)"        (n 5.0)
checkApprox "stdevp" "=STDEVP(2,4,4,4,5,5,7,9)" 2.0
checkApprox "var"    "=VAR(2,4,4,4,5,5,7,9)"   4.571428571428571

// ----- criteria aggregates -----
check "countif-eq"   "=COUNTIF(F1:F5,5)"      (n 2.0)
check "countif-gt"   "=COUNTIF(F1:F5,\">10\")" (n 3.0)
check "sumif-gt"     "=SUMIF(F1:F5,\">10\")"  (n 75.0)
check "sumif-range"  "=SUMIF(F1:F5,5,G1:G5)"  (n 500.0)   // F1+F4 → G1+G4 = 100+400
check "averageif"    "=AVERAGEIF(F1:F5,5)"    (n 5.0)
check "sumifs"       "=SUMIFS(G1:G5,F1:F5,5)" (n 500.0)
check "countifs"     "=COUNTIFS(F1:F5,\">10\",G1:G5,\">100\")" (n 3.0)
check "sumproduct"   "=SUMPRODUCT(D1:D3,D1:D3)" (n 14.0)   // 1+4+9

// ----- math -----
check "abs"          "=ABS(A5)"               (n 3.0)
check "sqrt"         "=SQRT(16)"              (n 4.0)
check "sqrt-neg"     "=SQRT(-1)"              (e "NUM!")
check "sign"         "=SIGN(A5)"              (n -1.0)
check "int"          "=INT(3.9)"              (n 3.0)
check "int-neg"      "=INT(-3.1)"             (n -4.0)
check "trunc"        "=TRUNC(3.99)"           (n 3.0)
check "trunc-dig"    "=TRUNC(3.14159,2)"      (n 3.14)
check "round"        "=ROUND(2.5,0)"          (n 3.0)
check "round-dig"    "=ROUND(3.14159,2)"      (n 3.14)
check "roundup"      "=ROUNDUP(3.1,0)"        (n 4.0)
check "rounddown"    "=ROUNDDOWN(3.9,0)"      (n 3.0)
check "mod"          "=MOD(10,3)"             (n 1.0)
check "mod-neg"      "=MOD(-1,3)"             (n 2.0)      // sign of divisor
check "mod-div0"     "=MOD(1,0)"              (e "DIV/0!")
check "power"        "=POWER(2,8)"            (n 256.0)

// ----- logical -----
check "and-true"     "=AND(TRUE,1=1)"         (b true)
check "and-false"    "=AND(TRUE,FALSE)"       (b false)
check "or"           "=OR(FALSE,1>0)"         (b true)
check "xor"          "=XOR(TRUE,TRUE,TRUE)"   (b true)
check "not"          "=NOT(FALSE)"            (b true)
check "iferror"      "=IFERROR(1/0,\"oops\")" (s "oops")
check "iferror-ok"   "=IFERROR(5,\"oops\")"   (n 5.0)
check "ifna"         "=IFNA(#N/A,42)"         (n 42.0)
check "ifna-other"   "=IFNA(#REF!,42)"        (e "REF!")
check "ifs"          "=IFS(FALSE,1,TRUE,2)"   (n 2.0)
check "ifs-none"     "=IFS(FALSE,1,FALSE,2)"  (e "N/A")
check "switch"       "=SWITCH(2,1,\"a\",2,\"b\",\"z\")" (s "b")
check "switch-def"   "=SWITCH(9,1,\"a\",2,\"b\",\"z\")" (s "z")

// ----- info -----
check "isblank"      "=ISBLANK(Z99)"          (b true)
check "isblank-no"   "=ISBLANK(A1)"           (b false)
check "iserror"      "=ISERROR(1/0)"          (b true)
check "isna"         "=ISNA(#N/A)"            (b true)
check "isna-no"      "=ISNA(#REF!)"           (b false)
check "isnumber"     "=ISNUMBER(A1)"          (b true)
check "istext"       "=ISTEXT(A3)"            (b true)
check "islogical"    "=ISLOGICAL(A4)"         (b true)

// ----- text -----
check "concat"       "=CONCAT(\"a\",\"b\",\"c\")" (s "abc")
check "textjoin"     "=TEXTJOIN(\"-\",TRUE,\"a\",\"b\")" (s "a-b")
check "len"          "=LEN(\"hello\")"        (n 5.0)
check "trim"         "=TRIM(\"  a   b  \")"   (s "a b")
check "upper"        "=UPPER(\"aBc\")"        (s "ABC")
check "lower"        "=LOWER(\"aBc\")"        (s "abc")
check "proper"       "=PROPER(\"hELLO wORLD\")" (s "Hello World")
check "left"         "=LEFT(\"hello\",3)"     (s "hel")
check "right"        "=RIGHT(\"hello\",3)"    (s "llo")
check "mid"          "=MID(\"hello\",2,3)"    (s "ell")
check "rept"         "=REPT(\"ab\",3)"        (s "ababab")
check "find"         "=FIND(\"l\",\"hello\")" (n 3.0)
check "find-miss"    "=FIND(\"z\",\"hello\")" (e "VALUE!")
check "search-ci"    "=SEARCH(\"L\",\"hello\")" (n 3.0)
check "substitute"   "=SUBSTITUTE(\"a-b-c\",\"-\",\"+\")" (s "a+b+c")
check "substitute-i" "=SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)" (s "a-b+c")
check "replace"      "=REPLACE(\"abcdef\",2,3,\"XY\")" (s "aXYef")
check "concat-amp"   "=CONCAT(A1,A2)"         (s "1020")

// ----- lookup -----
check "choose"       "=CHOOSE(2,\"a\",\"b\",\"c\")" (s "b")
check "index-col"    "=INDEX(E1:E3,2)"        (s "two")
check "index-2d"     "=INDEX(D1:E3,3,2)"      (s "three")
check "match-exact"  "=MATCH(2,D1:D3,0)"      (n 2.0)
check "match-approx" "=MATCH(2.5,D1:D3,1)"    (n 2.0)
check "vlookup"      "=VLOOKUP(2,D1:E3,2,FALSE)" (s "two")
check "vlookup-approx" "=VLOOKUP(2.9,D1:E3,2,TRUE)" (s "two")
check "hlookup"      "=HLOOKUP(2,D7:F8,2,FALSE)" (s "y")

// ----- date -----
check "date"         "=DATE(2020,1,1)"        (n 43831.0)
check "year"         "=YEAR(43831)"           (n 2020.0)
check "month"        "=MONTH(43831)"          (n 1.0)
check "day"          "=DAY(43831)"            (n 1.0)
check "weekday"      "=WEEKDAY(43831)"        (n 4.0)        // 2020-01-01 = Wednesday
check "eomonth"      "=EOMONTH(43831,0)"      (n 43861.0)    // 2020-01-31
// EDATE / EOMONTH of whole serials are worked out without a date object: the same as with one, day for day.
do
    let epoch = System.DateTime(1899, 12, 30)
    let mutable wrong = []
    let mutable tried = 0
    for serial in [ 61 .. 53 .. 110000 ] @ [ 61; 62; 366; 425; 426; 36525; 36585; 36586; 43890; 43891; 44255; 44256; 73050; 73109; 73110; 2958465; 2958400 ] do
        for months in [ -1200; -25; -13; -12; -11; -2; -1; 0; 1; 2; 11; 12; 13; 25; 600 ] do
            let date = epoch.AddDays(float serial)
            let ok = (try Some (date.AddMonths months) with _ -> None)
            match ok with
            | Some moved when moved.Year >= 1900 && (moved - epoch).TotalDays >= 61.0 ->
                tried <- tried + 1
                let wantEdate = (moved - epoch).TotalDays
                let wantEom = (System.DateTime(moved.Year, moved.Month, System.DateTime.DaysInMonth(moved.Year, moved.Month)) - epoch).TotalDays
                let gotEdate = evaluateFormula (sprintf "=EDATE(%d,%d)" serial months) model None
                let gotEom = evaluateFormula (sprintf "=EOMONTH(%d,%d)" serial months) model None
                if gotEdate <> CellValue.Number wantEdate || gotEom <> CellValue.Number wantEom then
                    wrong <- (serial, months, gotEdate, wantEdate, gotEom, wantEom) :: wrong
            | _ -> ()
    if wrong.IsEmpty && tried > 20000 then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL edate/eomonth in whole numbers: %d of %d differ, e.g. %A" wrong.Length tried (List.truncate 3 wrong))
check "edate (a time of day stays)" "=EDATE(43831.5,1)" (n 43862.5)

printfn ""
printfn "Stage 2 formula tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
