#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

// Build a tiny model with some seed cells: A1=10, A2=20, A3=text "hi",
// A4=TRUE, B1=3.
let cells =
    [ { Row=0; Col=0 }, { Value = CellValue.Number 10.0; Format = defaultFormat; Comment = None }
      { Row=1; Col=0 }, { Value = CellValue.Number 20.0; Format = defaultFormat; Comment = None }
      { Row=2; Col=0 }, { Value = CellValue.String "hi"; Format = defaultFormat; Comment = None }
      { Row=3; Col=0 }, { Value = CellValue.Boolean true; Format = defaultFormat; Comment = None }
      { Row=0; Col=1 }, { Value = CellValue.Number 3.0; Format = defaultFormat; Comment = None } ]
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
        printfn "FAIL %-28s  %-22s  expected %A  got %A" label formula expected got

let n x = CellValue.Number x
let s x = CellValue.String x
let b x = CellValue.Boolean x
let e x = CellValue.Error x

// arithmetic + precedence
check "add"            "=1+2"          (n 3.0)
check "mul-prec"       "=2+3*4"        (n 14.0)
check "paren"          "=(2+3)*4"      (n 20.0)
check "div"            "=10/4"         (n 2.5)
check "div0"           "=1/0"          (e "DIV/0!")
check "exp"            "=2^10"         (n 1024.0)
check "exp-left-assoc" "=2^3^2"        (n 64.0)
check "unary-neg-exp"  "=-2^2"         (n 4.0)        // Excel: (-2)^2
check "neg"            "=-5"           (n -5.0)
check "percent"        "=50%"          (n 0.5)
check "percent-expr"   "=10+5%"        (n 10.05)
check "concat"         "=\"a\"&\"b\""  (s "ab")
check "concat-num"     "=1&2"          (s "12")
check "concat-prec"    "=1+2&3"        (s "33")       // (1+2)&3

// comparisons
check "eq-true"        "=1=1"          (b true)
check "eq-false"       "=1=2"          (b false)
check "ne"             "=1<>2"         (b true)
check "lt"             "=1<2"          (b true)
check "ge"             "=2>=2"         (b true)
check "cmp-prec"       "=1+1=2"        (b true)       // (1+1)=2
check "str-cmp-ci"     "=\"ABC\"=\"abc\""  (b true)   // case-insensitive

// cell refs
check "ref"            "=A1"           (n 10.0)
check "ref-add"        "=A1+A2"        (n 30.0)
check "abs-ref"        "=$A$1"         (n 10.0)
check "mixed-ref"      "=$A1+A$2"      (n 30.0)

// coercion
check "text-num-add"   "=\"5\"+3"      (n 8.0)
check "bool-add"       "=TRUE+1"       (n 2.0)
check "text-fail"      "=\"x\"+1"      (e "VALUE!")
check "blank-add"      "=Z99+5"        (n 5.0)        // blank cell → 0

// error propagation + literals
check "err-lit"        "=#REF!"        (e "REF!")
check "err-prop"       "=#N/A+1"       (e "N/A")
check "err-prop-deep"  "=(1/0)*2"      (e "DIV/0!")

// functions
check "sum-range"      "=SUM(A1:A2)"   (n 30.0)
check "sum-mixed"      "=SUM(A1:A4)"   (n 30.0)       // 10+20; "hi" and TRUE in a range are ignored
check "average"        "=AVERAGE(A1:A2)" (n 15.0)
check "count"          "=COUNT(A1:A4)" (n 2.0)        // only 2 numbers
check "min"            "=MIN(A1:A2)"   (n 10.0)
check "max"            "=MAX(A1:B1)"   (n 10.0)
check "if-true"        "=IF(1>0,\"y\",\"n\")"  (s "y")
check "if-false"       "=IF(A1>100,1,2)"       (n 2.0)
check "if-nested"      "=IF(A1>5,IF(B1>5,1,2),3)" (n 2.0)
check "sum-of-func"    "=SUM(A1:A2)+B1"  (n 33.0)

// Unary plus is a no-op: it does not coerce its operand. A3 = "hi", A4 = TRUE.
check "plus-text-cell"     "=+A3"            (s "hi")
check "plus-text-abs"      "=+$A$3"          (s "hi")
check "plus-text-literal"  "=+\"abc\""       (s "abc")
check "plus-number"        "=+A1"            (n 10.0)
check "plus-boolean"       "=+A4"            (b true)
check "plus-plus-text"     "=++A3"           (s "hi")
check "plus-then-concat"   "=+A3&\"!\""      (s "hi!")
check "plus-error"         "=+(1/0)"         (e "DIV/0!")
check "plus-in-sum"        "=SUM(+A1,+A2)"   (n 30.0)
check "plus-blank-is-ref"  "=+Z99"           (evaluateFormula "=Z99" model None)
// Unary minus and percent coerce: numeric text converts, other text is #VALUE!.
check "minus-text-cell"    "=-A3"            (e "VALUE!")
check "minus-numeric-text" "=-\"5\""         (n -5.0)
check "minus-boolean"      "=-A4"            (n -1.0)
check "minus-blank"        "=-Z99"           (n 0.0)
check "percent-text"       "=A3%"            (e "VALUE!")
check "percent-num-text"   "=\"50\"%"        (n 0.5)
check "percent-boolean"    "=A4%"            (n 0.01)
check "minus-plus-text"    "=-+A3"           (e "VALUE!")
check "plus-minus-number"  "=+-A1"           (n -10.0)

printfn ""
printfn "Stage 1 formula tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
