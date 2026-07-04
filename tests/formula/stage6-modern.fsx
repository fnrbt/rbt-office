#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// Batch 1 — modern functions: MAXIFS / MINIFS / AVERAGEIFS, XLOOKUP / XMATCH,
// LET. Expected values are Excel's documented results (cross-checked against
// real Excel via tools/excel-parity probes).

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Formula

// Grid layout (0-indexed row,col):
//   A1:A6 (col0) category  : X Y X Y X Z
//   B1:B6 (col1) value     : 10 20 30 40 50 60
//   D1:D5 (col3) keys      : 1 2 3 4 5
//   E1:E5 (col4) returns   : a b c d e
//   H1:H3 (col7) text keys : apple banana cherry
//   I1:I3 (col8) returns   : A B C
//   K1:K4 (col10) dup keys : 1 2 2 3
//   L1:L4 (col11) returns  : w x y z
let cell v = { Value = v; Format = defaultFormat; Comment = None }
let nC r c v = ({ Row = r; Col = c }, cell (CellValue.Number v))
let sC r c v = ({ Row = r; Col = c }, cell (CellValue.String v))
let cells =
    [ sC 0 0 "X"; sC 1 0 "Y"; sC 2 0 "X"; sC 3 0 "Y"; sC 4 0 "X"; sC 5 0 "Z"
      nC 0 1 10.0; nC 1 1 20.0; nC 2 1 30.0; nC 3 1 40.0; nC 4 1 50.0; nC 5 1 60.0
      nC 0 3 1.0; nC 1 3 2.0; nC 2 3 3.0; nC 3 3 4.0; nC 4 3 5.0
      sC 0 4 "a"; sC 1 4 "b"; sC 2 4 "c"; sC 3 4 "d"; sC 4 4 "e"
      sC 0 7 "apple"; sC 1 7 "banana"; sC 2 7 "cherry"
      sC 0 8 "A"; sC 1 8 "B"; sC 2 8 "C"
      nC 0 10 1.0; nC 1 10 2.0; nC 2 10 2.0; nC 3 10 3.0
      sC 0 11 "w"; sC 1 11 "x"; sC 2 11 "y"; sC 3 11 "z" ]
    |> Map.ofList
let model = { createEmpty 100 26 with Cells = cells }

let mutable pass = 0
let mutable fail = 0
let check label formula (expected: CellValue) =
    let got = evaluateFormula formula model None
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-22s  %-46s  expected %A  got %A" label formula expected got
let n x = CellValue.Number x
let s x = CellValue.String x
let e x = CellValue.Error x

// ----- MAXIFS / MINIFS / AVERAGEIFS -----
check "maxifs"        "=MAXIFS(B1:B6,A1:A6,\"X\")"               (n 50.0)
check "minifs"        "=MINIFS(B1:B6,A1:A6,\"X\")"               (n 10.0)
check "averageifs"    "=AVERAGEIFS(B1:B6,A1:A6,\"X\")"           (n 30.0)
check "maxifs-nomatch" "=MAXIFS(B1:B6,A1:A6,\"Q\")"              (n 0.0)
check "minifs-nomatch" "=MINIFS(B1:B6,A1:A6,\"Q\")"              (n 0.0)
check "avgifs-nomatch" "=AVERAGEIFS(B1:B6,A1:A6,\"Q\")"         (e "DIV/0!")
check "maxifs-2crit"  "=MAXIFS(B1:B6,A1:A6,\"X\",B1:B6,\">10\")" (n 50.0)
check "minifs-2crit"  "=MINIFS(B1:B6,A1:A6,\"X\",B1:B6,\">10\")" (n 30.0)
check "avgifs-2crit"  "=AVERAGEIFS(B1:B6,A1:A6,\"X\",B1:B6,\">10\")" (n 40.0)

// ----- XLOOKUP -----
check "xlookup-exact"     "=XLOOKUP(3,D1:D5,E1:E5)"             (s "c")
check "xlookup-notfound4" "=XLOOKUP(9,D1:D5,E1:E5,\"none\")"    (s "none")
check "xlookup-na"        "=XLOOKUP(9,D1:D5,E1:E5)"             (e "N/A")
check "xlookup-smaller"   "=XLOOKUP(3.5,D1:D5,E1:E5,\"x\",-1)"  (s "c")
check "xlookup-larger"    "=XLOOKUP(3.5,D1:D5,E1:E5,\"x\",1)"   (s "d")
check "xlookup-wildcard"  "=XLOOKUP(\"ban*\",H1:H3,I1:I3,\"x\",2)" (s "B")
check "xlookup-dup-first" "=XLOOKUP(2,K1:K4,L1:L4)"             (s "x")
check "xlookup-dup-last"  "=XLOOKUP(2,K1:K4,L1:L4,\"x\",0,-1)"  (s "y")

// ----- XMATCH -----
check "xmatch-exact"      "=XMATCH(3,D1:D5)"                    (n 3.0)
check "xmatch-smaller"    "=XMATCH(3.5,D1:D5,-1)"               (n 3.0)
check "xmatch-larger"     "=XMATCH(3.5,D1:D5,1)"                (n 4.0)
check "xmatch-wildcard"   "=XMATCH(\"ban*\",H1:H3,2)"           (n 2.0)
check "xmatch-na"         "=XMATCH(99,D1:D5)"                   (e "N/A")
check "xmatch-dup-last"   "=XMATCH(2,K1:K4,0,-1)"               (n 3.0)

// ----- LET -----
check "let-basic"     "=LET(x,5,x+1)"                            (n 6.0)
check "let-two"       "=LET(x,5,y,x*2,x+y)"                      (n 15.0)
check "let-range-sum" "=LET(r,B1:B6,SUM(r))"                     (n 210.0)
check "let-countif"   "=LET(rng,A1:A6,COUNTIF(rng,\"X\"))"       (n 3.0)
check "let-chain"     "=LET(a,2,b,a+1,c,b+1,a+b+c)"              (n 9.0)
check "let-sumifs"    "=LET(v,B1:B6,cat,A1:A6,SUMIFS(v,cat,\"X\"))" (n 90.0)
check "let-reuse"     "=LET(r,B1:B6,SUM(r)/COUNT(r))"            (n 35.0)
check "let-bad-arity" "=LET(x,5)"                                (e "VALUE!")

printfn ""
printfn "Stage 6 modern-function tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
