#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

// Functions added for item 12: reference information, trigonometry, dates,
// regression/normal statistics, matrices, array shaping, database functions,
// base conversion and bit operations. Expected values are Excel's.
//
//   A1..A5 = 10,20,30,40,50      B1..B5 = 1,2,3,4,5
//   C1..C5 = 1,5,2,8,4           D1..D5 = apple,fig,kiwi,fig,pear
//   E1..E5 = 1,"2",TRUE,blank,0  G1..G3 = 2,blank,4
//   I1:K5  = Name,Qty,Price / apple,3,1.5 / fig,5,2 / kiwi,2,4 / fig,1,2.5
//   M1:N2  = Name,Qty / fig,>2
let cell v = { Value = v; Format = defaultFormat; Comment = None }
let n x = CellValue.Number x
let s x = CellValue.String x
let b x = CellValue.Boolean x
let e x = CellValue.Error x
let at r c = { Row = r; Col = c }
let column c values = values |> List.mapi (fun r v -> at r c, cell v)
let cells =
    column 0 [ n 10.0; n 20.0; n 30.0; n 40.0; n 50.0 ]
    @ column 1 [ n 1.0; n 2.0; n 3.0; n 4.0; n 5.0 ]
    @ column 2 [ n 1.0; n 5.0; n 2.0; n 8.0; n 4.0 ]
    @ column 3 [ s "apple"; s "fig"; s "kiwi"; s "fig"; s "pear" ]
    @ [ at 0 4, cell (n 1.0); at 1 4, cell (s "2"); at 2 4, cell (b true); at 4 4, cell (n 0.0) ]
    @ [ at 0 6, cell (n 2.0); at 2 6, cell (n 4.0) ]
    @ column 8 [ s "Name"; s "apple"; s "fig"; s "kiwi"; s "fig" ]
    @ column 9 [ s "Qty"; n 3.0; n 5.0; n 2.0; n 1.0 ]
    @ column 10 [ s "Price"; n 1.5; n 2.0; n 4.0; n 2.5 ]
    @ column 12 [ s "Name"; s "fig" ]
    @ column 13 [ s "Qty"; s ">2" ]
    |> Map.ofList
let model = { createEmpty 100 26 with Cells = cells }

let mutable pass = 0
let mutable fail = 0
let check label formula (expected: CellValue) =
    let got = evaluateFormula formula model None
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-22s  %-44s  expected %A  got %A" label formula expected got
let checkApprox label formula (expected: float) =
    match evaluateFormula formula model None with
    | CellValue.Number got when abs (got - expected) < 1e-9 * max 1.0 (abs expected) -> pass <- pass + 1
    | got ->
        fail <- fail + 1
        printfn "FAIL %-22s  %-44s  expected ~%.12g  got %A" label formula expected got

// ----- reference information ---------------------------------------------------
check "row"              "=ROW(B3)"                     (n 3.0)
check "column"           "=COLUMN(B3)"                  (n 2.0)
check "row-of-range"     "=ROW(C4:D9)"                  (n 4.0)
check "rows"             "=ROWS(A1:B5)"                 (n 5.0)
check "columns"          "=COLUMNS(A1:B5)"              (n 2.0)
check "rows-whole-col"   "=ROWS(A:A)"                   (n 1048576.0)
check "columns-whole-row" "=COLUMNS(1:1)"               (n 16384.0)
check "rows-array"       "=ROWS({1,2;3,4;5,6})"         (n 3.0)
check "columns-array"    "=COLUMNS({1,2;3,4;5,6})"      (n 2.0)
check "rows-indirect"    "=ROWS(INDIRECT(\"A1:A4\"))"   (n 4.0)
check "row-array-sum"    "=SUM(ROW(A1:A3))"             (n 6.0)
check "row-sumproduct"   "=SUMPRODUCT(ROW(A1:A5),B1:B5)" (n 55.0)
check "column-array"     "=SUM(COLUMN(A1:C1))"          (n 6.0)
check "address"          "=ADDRESS(2,3)"                (s "$C$2")
check "address-rel"      "=ADDRESS(2,3,4)"              (s "C2")
check "address-mixed"    "=ADDRESS(2,3,2)"              (s "C$2")
check "address-sheet"    "=ADDRESS(1,1,1,TRUE,\"My Sheet\")" (s "'My Sheet'!$A$1")
check "address-indirect" "=INDIRECT(ADDRESS(3,1))"      (n 30.0)
check "iserr"            "=ISERR(1/0)"                  (b true)
check "iserr-na"         "=ISERR(NA())"                 (b false)
check "isref-cell"       "=ISREF(A1)"                   (b true)
check "isref-value"      "=ISREF(1)"                    (b false)
check "isref-indirect"   "=ISREF(INDIRECT(\"A1\"))"     (b true)
check "error-type-na"    "=ERROR.TYPE(#N/A)"            (n 7.0)
check "error-type-div"   "=ERROR.TYPE(1/0)"             (n 2.0)
check "error-type-none"  "=ERROR.TYPE(1)"               (e "N/A")
check "type-number"      "=TYPE(1)"                     (n 1.0)
check "type-text"        "=TYPE(\"a\")"                 (n 2.0)
check "type-bool"        "=TYPE(TRUE)"                  (n 4.0)
check "type-error"       "=TYPE(#N/A)"                  (n 16.0)
check "type-range"       "=TYPE(A1:A2)"                 (n 64.0)
check "type-array"       "=TYPE({1,2})"                 (n 64.0)
check "hyperlink-name"   "=HYPERLINK(\"http://x\",\"Click\")" (s "Click")
check "hyperlink-link"   "=HYPERLINK(\"http://x\")"     (s "http://x")

// ----- trigonometry -------------------------------------------------------------
checkApprox "asin"       "=ASIN(1)"                     1.5707963267948966
check       "acos-domain" "=ACOS(2)"                    (e "NUM!")
checkApprox "acos"       "=ACOS(0.5)"                   1.0471975511965979
checkApprox "atan"       "=ATAN(1)"                     0.7853981633974483
checkApprox "atan2"      "=ATAN2(1,1)"                  0.7853981633974483
checkApprox "atan2-quad" "=ATAN2(-1,1)"                 2.356194490192345
check       "atan2-zero" "=ATAN2(0,0)"                  (e "DIV/0!")
checkApprox "sinh"       "=SINH(1)"                     1.1752011936438014
checkApprox "cosh"       "=COSH(1)"                     1.5430806348152437
checkApprox "tanh"       "=TANH(1)"                     0.7615941559557649
checkApprox "degrees"    "=DEGREES(PI())"               180.0
checkApprox "radians"    "=RADIANS(180)"                3.141592653589793

// ----- dates ----------------------------------------------------------------------
checkApprox "yearfrac-30-360" "=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1))"   0.5
checkApprox "yearfrac-actual" "=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),1)" 0.4972677595628415
checkApprox "yearfrac-act360" "=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),2)" 0.5055555555555555
checkApprox "yearfrac-act365" "=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),3)" 0.4986301369863014
checkApprox "yearfrac-multi"  "=YEARFRAC(DATE(2020,1,1),DATE(2023,1,1),1)" 3.000684462696783
check "days360"          "=DAYS360(DATE(2024,1,30),DATE(2024,3,31))"        (n 60.0)
check "days360-feb"      "=DAYS360(DATE(2024,2,29),DATE(2024,3,31))"        (n 30.0)
check "days360-eu"       "=DAYS360(DATE(2024,1,31),DATE(2024,3,31),TRUE)"   (n 60.0)
check "days360-year"     "=DAYS360(DATE(2023,1,1),DATE(2024,1,1))"          (n 360.0)
checkApprox "timevalue-pm"   "=TIMEVALUE(\"6:00 PM\")"          0.75
checkApprox "timevalue-24h"  "=TIMEVALUE(\"12:30\")"            0.5208333333333334
checkApprox "timevalue-secs" "=TIMEVALUE(\"01:02:03\")"         0.04309027777777778
checkApprox "timevalue-date" "=TIMEVALUE(\"2024-01-01 06:00\")" 0.25
check "timevalue-bad"    "=TIMEVALUE(\"x\")"                    (e "VALUE!")
// 45292 = Monday 2024-01-01
check "networkdays-intl"         "=NETWORKDAYS.INTL(45292,45298)"             (n 5.0)
check "networkdays-intl-sunday"  "=NETWORKDAYS.INTL(45292,45298,11)"          (n 6.0)
check "networkdays-intl-mask"    "=NETWORKDAYS.INTL(45292,45298,\"0000011\")" (n 5.0)
check "networkdays-intl-none"    "=NETWORKDAYS.INTL(45292,45298,\"0000000\")" (n 7.0)
check "networkdays-intl-holiday" "=NETWORKDAYS.INTL(45292,45298,1,{45293})"   (n 4.0)
check "networkdays-intl-fri-sat" "=NETWORKDAYS.INTL(45292,45298,7)"           (n 5.0)
check "workday-intl"             "=WORKDAY.INTL(45292,5)"                     (n 45299.0)
check "workday-intl-sunday"      "=WORKDAY.INTL(45296,2,11)"                  (n 45299.0)
check "workday-intl-fri-sat"     "=WORKDAY.INTL(45295,1,7)"                   (n 45298.0)
check "workday-intl-back"        "=WORKDAY.INTL(45299,-1)"                    (n 45296.0)
check "workday-intl-holiday"     "=WORKDAY.INTL(45292,1,1,{45293})"           (n 45294.0)

// ----- regression / correlation (x = B1:B5, y = C1:C5) -----------------------------
checkApprox "slope"      "=SLOPE(C1:C5,B1:B5)"          0.9
checkApprox "intercept"  "=INTERCEPT(C1:C5,B1:B5)"      1.3
checkApprox "correl"     "=CORREL(C1:C5,B1:B5)"         0.5196152422706632
checkApprox "rsq"        "=RSQ(C1:C5,B1:B5)"            0.27
checkApprox "forecast"   "=FORECAST(6,C1:C5,B1:B5)"     6.7
checkApprox "forecast-linear" "=FORECAST.LINEAR(6,C1:C5,B1:B5)" 6.7
checkApprox "slope-perfect" "=SLOPE(A1:A5,B1:B5)"       10.0
checkApprox "correl-perfect" "=CORREL(A1:A5,B1:B5)"     1.0
check "correl-size"      "=CORREL(A1:A5,B1:B4)"         (e "N/A")
checkApprox "rank-avg"   "=RANK.AVG(2,{1,2,2,3})"       2.5
checkApprox "rank-avg-asc" "=RANK.AVG(5,C1:C5,1)"       4.0
check "rank-eq"          "=RANK.EQ(5,C1:C5)"            (n 2.0)
checkApprox "percentrank" "=PERCENTRANK(C1:C5,4)"       0.5
checkApprox "percentrank-interp" "=PERCENTRANK.INC(C1:C5,3)" 0.375
check "percentrank-out"  "=PERCENTRANK(C1:C5,9)"        (e "N/A")
check "frequency-last"   "=INDEX(FREQUENCY(C1:C5,{2,5}),3)" (n 1.0)
check "frequency-first"  "=INDEX(FREQUENCY(C1:C5,{2,5}),1)" (n 2.0)
check "frequency-total"  "=SUM(FREQUENCY(C1:C5,{2,5}))" (n 5.0)
checkApprox "norm-s-dist"     "=NORM.S.DIST(1.96,TRUE)"     0.9750021048517795
checkApprox "norm-s-dist-0"   "=NORM.S.DIST(0,TRUE)"        0.5
checkApprox "norm-s-dist-pdf" "=NORM.S.DIST(0,FALSE)"       0.3989422804014327
checkApprox "norm-dist"       "=NORM.DIST(42,40,1.5,TRUE)"  0.9087887802741321
checkApprox "norm-dist-pdf"   "=NORM.DIST(42,40,1.5,FALSE)" 0.10934004978399577
checkApprox "norm-s-inv"      "=NORM.S.INV(0.975)"          1.959963984540054
checkApprox "norm-s-inv-low"  "=NORM.S.INV(0.001)"          -3.090232306167813
checkApprox "norm-inv"        "=NORM.INV(0.5,10,2)"         10.0
checkApprox "norm-inv-tail"   "=NORM.INV(0.975,10,2)"       13.919927969080108
check "norm-inv-domain"  "=NORM.S.INV(1)"               (e "NUM!")

// ----- matrices and array shaping ---------------------------------------------------
check "mmult"            "=INDEX(MMULT({1,2;3,4},{5,6;7,8}),2,1)" (n 43.0)
check "mmult-sum"        "=SUM(MMULT({1,2;3,4},{5,6;7,8}))"       (n 134.0)
check "mmult-shape"      "=MMULT({1,2},{1,2})"                    (e "VALUE!")
checkApprox "minverse"   "=INDEX(MINVERSE({4,7;2,6}),1,2)"        -0.7
checkApprox "minverse-2" "=INDEX(MINVERSE({4,7;2,6}),2,2)"        0.4
check "minverse-singular" "=MINVERSE({1,2;2,4})"                  (e "NUM!")
checkApprox "mdeterm"    "=MDETERM({1,2;3,4})"                    -2.0
checkApprox "mdeterm-3"  "=MDETERM({2,0,1;1,3,2;1,1,1})"          0.0
checkApprox "mdeterm-3b" "=MDETERM({1,2,3;0,1,4;5,6,0})"          1.0
check "transpose"        "=INDEX(TRANSPOSE({1,2;3,4}),1,2)"       (n 3.0)
check "vstack"           "=INDEX(VSTACK({1,2},{3,4}),2,1)"        (n 3.0)
check "vstack-rows"      "=ROWS(VSTACK(A1:A2,B1:B3))"             (n 5.0)
check "vstack-pad"       "=ISNA(INDEX(VSTACK({1,2},{3}),2,2))"    (b true)
check "hstack"           "=INDEX(HSTACK({1;2},{3;4}),2,2)"        (n 4.0)
check "hstack-cols"      "=COLUMNS(HSTACK(A1:A2,B1:B2))"          (n 2.0)
check "take-rows"        "=SUM(TAKE(A1:B5,2))"                    (n 33.0)
check "take-last"        "=SUM(TAKE(A1:A5,-2))"                   (n 90.0)
check "take-cols"        "=SUM(TAKE(A1:B5,2,1))"                  (n 30.0)
check "take-only-cols"   "=SUM(TAKE(A1:B5,,-1))"                  (n 15.0)
check "drop-rows"        "=SUM(DROP(A1:A5,2))"                    (n 120.0)
check "drop-last"        "=SUM(DROP(A1:A5,-1))"                   (n 100.0)
check "drop-all"         "=DROP(A1:A5,5)"                         (e "CALC!")
check "choosecols"       "=INDEX(CHOOSECOLS(A1:C5,3,1),2,1)"      (n 5.0)
check "choosecols-neg"   "=SUM(CHOOSECOLS(A1:C5,-1))"             (n 20.0)
check "choosecols-bad"   "=CHOOSECOLS(A1:C5,4)"                   (e "VALUE!")
check "chooserows"       "=SUM(CHOOSEROWS(A1:B5,-1))"             (n 55.0)
check "chooserows-two"   "=SUM(CHOOSEROWS(A1:B5,1,3))"            (n 44.0)
check "tocol"            "=INDEX(TOCOL({1,2;3,4}),3)"             (n 3.0)
check "tocol-by-col"     "=INDEX(TOCOL({1,2;3,4},,TRUE),2)"       (n 3.0)
check "tocol-skip-blank" "=ROWS(TOCOL(G1:G3,1))"                  (n 2.0)
check "torow"            "=COLUMNS(TOROW(A1:B2))"                 (n 4.0)

// ----- database functions --------------------------------------------------------------
check "dsum"             "=DSUM(I1:K5,\"Qty\",M1:M2)"             (n 6.0)
check "dsum-index"       "=DSUM(I1:K5,2,M1:M2)"                   (n 6.0)
check "dsum-number-crit" "=DSUM(I1:K5,\"Qty\",N1:N2)"             (n 8.0)
check "dsum-and"         "=DSUM(I1:K5,\"Qty\",M1:N2)"             (n 5.0)
check "dcount"           "=DCOUNT(I1:K5,\"Qty\",M1:M2)"           (n 2.0)
checkApprox "daverage"   "=DAVERAGE(I1:K5,\"Price\",M1:M2)"       2.25
check "dmax"             "=DMAX(I1:K5,\"Qty\",M1:M2)"             (n 5.0)
check "dmin"             "=DMIN(I1:K5,\"Qty\",M1:M2)"             (n 1.0)
check "dget"             "=DGET(I1:K5,\"Price\",M1:N2)"           (n 2.0)
check "dget-many"        "=DGET(I1:K5,\"Price\",M1:M2)"           (e "NUM!")
check "dsum-bad-field"   "=DSUM(I1:K5,\"Nope\",M1:M2)"            (e "VALUE!")

// ----- base conversion and bit operations ------------------------------------------------
check "bin2dec"          "=BIN2DEC(\"1010\")"           (n 10.0)
check "bin2dec-neg"      "=BIN2DEC(1111111111)"         (n -1.0)
check "bin2dec-bad"      "=BIN2DEC(\"102\")"            (e "NUM!")
check "dec2bin"          "=DEC2BIN(10)"                 (s "1010")
check "dec2bin-places"   "=DEC2BIN(10,8)"               (s "00001010")
check "dec2bin-neg"      "=DEC2BIN(-1)"                 (s "1111111111")
check "dec2bin-range"    "=DEC2BIN(512)"                (e "NUM!")
check "hex2dec"          "=HEX2DEC(\"FF\")"             (n 255.0)
check "hex2dec-lower"    "=HEX2DEC(\"1f\")"             (n 31.0)
check "hex2dec-neg"      "=HEX2DEC(\"FFFFFFFFFF\")"     (n -1.0)
check "dec2hex"          "=DEC2HEX(255)"                (s "FF")
check "dec2hex-places"   "=DEC2HEX(255,4)"              (s "00FF")
check "dec2hex-neg"      "=DEC2HEX(-1)"                 (s "FFFFFFFFFF")
check "dec2hex-zero"     "=DEC2HEX(0)"                  (s "0")
check "bitand"           "=BITAND(12,10)"               (n 8.0)
check "bitor"            "=BITOR(12,10)"                (n 14.0)
check "bitxor"           "=BITXOR(12,10)"               (n 6.0)
check "bitand-negative"  "=BITAND(-1,1)"                (e "NUM!")
check "bitand-large"     "=BITAND(4294967296,4294967297)" (n 4294967296.0)

printfn ""
printfn "Stage 23 function tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
