#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// "Fill in gaps" — ~35 common functions across math/text/date/info.

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

let cell v = { Value = v; Format = defaultFormat; Comment = None }
// A1 = 10 (number), B1 = a formula (for ISFORMULA)
let model =
    { createEmpty 100 26 with
        Cells = Map.ofList
            [ { Row=0; Col=0 }, cell (CellValue.Number 10.0)
              { Row=0; Col=1 }, cell (CellValue.Formula("=A1+1", CellValue.Number 11.0)) ] }

let mutable pass = 0
let mutable fail = 0
let ev f = evaluateFormula f model None
let n x = CellValue.Number x
let s x = CellValue.String x
let b x = CellValue.Boolean x
let e x = CellValue.Error x
let check label f (expected: CellValue) =
    let got = ev f
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-22s %-28s expected %A got %A" label f expected got
let approx label f (expected: float) =
    match ev f with
    | CellValue.Number g when abs (g - expected) < 1e-9 -> pass <- pass + 1
    | got -> fail <- fail + 1; printfn "FAIL %-22s %-28s expected ~%f got %A" label f expected got

// ── math ──
approx "pi"        "=PI()"            System.Math.PI
approx "exp"       "=EXP(1)"          (exp 1.0)
approx "ln"        "=LN(EXP(2))"      2.0
check  "log10"     "=LOG10(1000)"     (n 3.0)
check  "log base"  "=LOG(8,2)"        (n 3.0)
approx "sin"       "=SIN(0)"          0.0
approx "cos"       "=COS(0)"          1.0
approx "radians"   "=RADIANS(180)"    System.Math.PI
approx "degrees"   "=DEGREES(PI())"   180.0
check  "quotient"  "=QUOTIENT(17,5)"  (n 3.0)
check  "ceiling"   "=CEILING(2.1,0.5)" (n 2.5)
check  "floor"     "=FLOOR(2.9,0.5)"  (n 2.5)
check  "mround"    "=MROUND(11,3)"    (n 12.0)
check  "even"      "=EVEN(3)"         (n 4.0)
check  "even-neg"  "=EVEN(-1)"        (n -2.0)
check  "odd"       "=ODD(2)"          (n 3.0)
check  "odd-1"     "=ODD(1)"          (n 1.0)
check  "sumsq"     "=SUMSQ(3,4)"      (n 25.0)
check  "fact"      "=FACT(5)"         (n 120.0)
check  "fact0"     "=FACT(0)"         (n 1.0)
check  "combin"    "=COMBIN(5,2)"     (n 10.0)
check  "gcd"       "=GCD(12,18,30)"   (n 6.0)
check  "lcm"       "=LCM(3,4,5)"      (n 60.0)

// ── text ──
check  "value"     "=VALUE(\"123.5\")"  (n 123.5)
check  "value-pct" "=VALUE(\"50%\")"    (n 0.5)
check  "exact-y"   "=EXACT(\"abc\",\"abc\")" (b true)
check  "exact-n"   "=EXACT(\"abc\",\"ABC\")" (b false)
check  "char"      "=CHAR(65)"        (s "A")
check  "code"      "=CODE(\"A\")"     (n 65.0)
check  "clean"     "=CLEAN(CHAR(9)&\"x\")" (s "x")
check  "textbefore" "=TEXTBEFORE(\"a-b-c\",\"-\")" (s "a")
check  "textbefore2" "=TEXTBEFORE(\"a-b-c\",\"-\",2)" (s "a-b")
check  "textafter" "=TEXTAFTER(\"a-b-c\",\"-\")" (s "b-c")
check  "textafter2" "=TEXTAFTER(\"a-b-c\",\"-\",2)" (s "c")
check  "fixed"     "=FIXED(1234.567,2)" (s "1,234.57")
check  "fixed-nc"  "=FIXED(1234.567,1,TRUE)" (s "1234.6")
check  "t-text"    "=T(\"hi\")"       (s "hi")
check  "t-num"     "=T(5)"            (s "")

// ── date/time ──
check  "time"      "=TIME(12,0,0)"    (n 0.5)
check  "hour"      "=HOUR(0.5)"       (n 12.0)
check  "minute"    "=MINUTE(TIME(1,30,0))" (n 30.0)
check  "second"    "=SECOND(TIME(0,0,45))" (n 45.0)
check  "edate"     "=EDATE(DATE(2020,1,31),1)" (n 43890.0) // 2020-02-29 (clamped)
check  "days"      "=DAYS(DATE(2020,1,10),DATE(2020,1,1))" (n 9.0)
check  "datevalue" "=DATEVALUE(\"2020-01-01\")" (n 43831.0)
check  "isoweek-1" "=ISOWEEKNUM(DATE(2020,1,1))" (n 1.0)
check  "isoweek-53" "=ISOWEEKNUM(DATE(2020,12,31))" (n 53.0)

// ── info ──
check  "iseven"    "=ISEVEN(4)"       (b true)
check  "isodd"     "=ISODD(3)"        (b true)
check  "isnontext" "=ISNONTEXT(5)"    (b true)
check  "isnontext-s" "=ISNONTEXT(\"x\")" (b false)
check  "n-num"     "=N(5)"            (n 5.0)
check  "n-bool"    "=N(TRUE)"         (n 1.0)
check  "n-text"    "=N(\"x\")"        (n 0.0)
check  "isformula-y" "=ISFORMULA(B1)" (b true)
check  "isformula-n" "=ISFORMULA(A1)" (b false)

printfn ""
printfn "Stage 10 extended-function tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
