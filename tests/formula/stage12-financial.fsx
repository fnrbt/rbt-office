#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// Financial (time-value-of-money) functions: PMT, FV, PV, NPER, RATE, NPV,
// IRR, SLN. Expected values match Excel.

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

let cell v = { Value = v; Format = defaultFormat; Comment = None }
// A1:A3 = -100, 50, 60  (cashflow for IRR)
let model =
    { createEmpty 100 26 with
        Cells = [ -100.; 50.; 60. ] |> List.mapi (fun i v -> { Row=i; Col=0 }, cell (CellValue.Number v)) |> Map.ofList }

let mutable pass = 0
let mutable fail = 0
let ev f = evaluateFormula f model None
let approxT label f (expected: float) (tol: float) =
    match ev f with
    | CellValue.Number g when abs (g - expected) < tol -> pass <- pass + 1
    | got -> fail <- fail + 1; printfn "FAIL %-20s %-34s expected ~%.6f got %A" label f expected got
let approx label f e = approxT label f e 1e-6

approxT "pmt"   "=PMT(0.1/12,60,10000)"   -212.4704471 1e-5
approx  "fv"    "=FV(0.05,10,-100)"        1257.789254
approx  "pv"    "=PV(0.08,20,-500)"        4909.073704
approxT "nper"  "=NPER(0.05,-100,1000)"    14.206699 1e-5
approxT "rate"  "=RATE(60,-212.470358,10000)" 0.0083333 1e-6
approx  "npv"   "=NPV(0.1,100,100,100)"    248.685199
approxT "irr"   "=IRR(A1:A3)"              0.063940   1e-5
approx  "sln"   "=SLN(10000,1000,5)"       1800.0
// round-trip: PV of the payment stream recovers the principal (positive)
approxT "pmt-roundtrip" "=PV(0.1/12,60,PMT(0.1/12,60,10000))" 10000.0 1e-4

// Last-digit agreement with Excel (values Excel cached in
// tests/real-world/milk-street-office.xlsx).
let exact label f (expected: float) =
    match ev f with
    | CellValue.Number g when g = expected -> pass <- pass + 1
    | got -> fail <- fail + 1; printfn "FAIL %-20s %-50s expected %s got %s" label f (expected.ToString "R") (match got with CellValue.Number g -> g.ToString "R" | other -> sprintf "%A" other)
// FV compounds a whole number of periods by multiplication and divides (factor - 1) by the rate first.
exact "fv-excel-digits"   "=-0.03*FV(0.025,5,-730662.5893421534)*0.6" -69130.84686587189
exact "fv-whole-periods"  "=FV(0.025,5,-730662.5893421534)" 3840602.60365955
approx "fv-fraction-periods" "=FV(0.05,2.5,-100)" ((1.05 ** 2.5 - 1.0) / 0.05 * 100.0)
approx "fv-zero-rate"     "=FV(0,10,-100,-50)" 1050.0
approx "fv-with-pv-type"  "=FV(0.06/12,10,-200,-500,1)" 2581.4033740601
// IRR converges on the rate itself, not merely on a small net present value:
// with small cash flows the old test (|NPV| < 1e-8) stopped far from the root.
let irrModel =
    { createEmpty 100 26 with
        Cells = [ -0.001; 0.0012; -100.0; 50.0; 60.0; -632835.0; 0.0; 0.0; -33539.42392; 3222649.874 ] |> List.mapi (fun i v -> { Row = i; Col = 1 }, cell (CellValue.Number v)) |> Map.ofList }
let irrOf f = match evaluateFormula f irrModel None with CellValue.Number g -> g | other -> failwithf "%s -> %A" f other
let near label (got: float) (expected: float) (tol: float) =
    if abs (got - expected) <= tol then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-20s expected %s got %s" label (expected.ToString "R") (got.ToString "R")
near "irr-small-cash-flows" (irrOf "=IRR(B1:B2)") 0.2 1e-14
near "irr-three-flows"      (irrOf "=IRR(B3:B5)") 0.06394102980498537 1e-14
let rate = irrOf "=IRR(B6:B10)"
let residual = [ -632835.0; 0.0; 0.0; -33539.42392; 3222649.874 ] |> List.mapi (fun i v -> v / ((1.0 + rate) ** float i)) |> List.sum
near "irr-converged-residual" residual 0.0 1e-8
near "irr-five-flows" rate 0.4963271306250475 1e-13

// The loan and dated cash-flow functions, by the examples of Microsoft's pages for them.
approxT "ipmt"      "=IPMT(0.1/12,1,36,8000)"                     -66.66666667 1e-6
approxT "ipmt-year" "=IPMT(0.1,3,3,8000)"                         -292.4471299 1e-6
approxT "ppmt"      "=PPMT(0.1/12,1,24,2000)"                     -75.62318601 1e-6
approxT "ppmt-fv"   "=PPMT(0.08,10,10,200000)"                    -27598.05346 1e-4
approxT "cumipmt"   "=CUMIPMT(0.09/12,30*12,125000,13,24,0)"      -11135.23213 1e-4
approxT "cumipmt-1" "=CUMIPMT(0.09/12,30*12,125000,1,1,0)"        -937.5 1e-6
approxT "cumprinc"  "=CUMPRINC(0.09/12,30*12,125000,13,24,0)"     -934.1071234 1e-5
approxT "cumprinc-1" "=CUMPRINC(0.09/12,30*12,125000,1,1,0)"      -68.27827118 1e-6
approxT "ipmt+ppmt is the payment" "=IPMT(0.05/12,7,60,25000)+PPMT(0.05/12,7,60,25000)-PMT(0.05/12,60,25000)" 0.0 1e-9
let dated =
    { model with
        Cells =
            [ for i, (v, d) in List.indexed [ -10000., 39448.; 2750., 39508.; 4250., 39751.; 3250., 39859.; 2750., 39904. ] do
                yield { Row = i; Col = 1 }, cell (CellValue.Number v)
                yield { Row = i; Col = 2 }, cell (CellValue.Number d) ]
            |> List.fold (fun cells (at, c) -> Map.add at c cells) model.Cells }
let evIn m f = evaluateFormula f m None
let near2 label m f expected tol =
    match evIn m f with
    | CellValue.Number g when abs (g - expected) < tol -> pass <- pass + 1
    | got -> fail <- fail + 1; printfn "FAIL %-20s %-34s expected ~%.9f got %A" label f expected got
near2 "xnpv" dated "=XNPV(0.09,B1:B5,C1:C5)" 2086.647602 1e-5
near2 "xirr" dated "=XIRR(B1:B5,C1:C5)" 0.373362535 1e-8
near2 "xirr-guess" dated "=XIRR(B1:B5,C1:C5,0.5)" 0.373362535 1e-8
let mirrModel =
    { model with Cells = [ -120000.; 39000.; 30000.; 21000.; 37000.; 46000. ] |> List.mapi (fun i v -> { Row = i; Col = 3 }, cell (CellValue.Number v)) |> List.fold (fun cells (at, c) -> Map.add at c cells) model.Cells }
near2 "mirr" mirrModel "=MIRR(D1:D6,0.1,0.12)" 0.126094 1e-6
near2 "mirr-three" mirrModel "=MIRR(D1:D4,0.1,0.12)" -0.048044655 1e-8
// IRR where Newton's method from 0.1 does not find the root: the rate is still found (Lyric-PF-DCF!F227).
let lyric =
    { model with Cells = [ -1.0e7; 4.0e5; 4.2e5; 4.3e5; 4.4e5; 1.04e7 ] |> List.mapi (fun i v -> { Row = i; Col = 4 }, cell (CellValue.Number v)) |> List.fold (fun cells (at, c) -> Map.add at c cells) model.Cells }
match evIn lyric "=IRR(E1:E6)", evIn lyric "=NPV(IRR(E1:E6),E2:E6)+E1" with
| CellValue.Number r, CellValue.Number residual when abs residual < 1e-4 && r > 0.0 && r < 1.0 -> pass <- pass + 1
| got -> fail <- fail + 1; printfn "FAIL irr-converges %A" got

// ----- Excel 2019's results (tools/excel-parity/eval-cells.vbs) -------------------------------
// Errors: one in an optional argument is the result; one among IRR's cash flows is #VALUE!.
let errModel =
    { createEmpty 100 26 with
        Cells = Map.ofList [ { Row = 0; Col = 0 }, cell (CellValue.Number -100.0); { Row = 1; Col = 0 }, cell (CellValue.Error "DIV/0!")
                             { Row = 2; Col = 0 }, cell (CellValue.Number 60.0); { Row = 3; Col = 0 }, cell (CellValue.Number 70.0) ] }
let evE f = evaluateFormula f errModel None
let errorIs label f (code: string) =
    match evE f with
    | CellValue.Error c when c = code -> pass <- pass + 1
    | got -> fail <- fail + 1; printfn "FAIL %-20s %-34s expected #%s got %A" label f code got
errorIs "irr-error-flow"  "=IRR(A1:A4)"            "VALUE!"
errorIs "irr-error-guess" "=IRR(A1:A4,A2)"         "DIV/0!"
errorIs "npv-error"       "=NPV(0.1,A1:A4)"        "DIV/0!"
errorIs "rate-error-fv"   "=RATE(5,,-100,A2)"      "DIV/0!"
errorIs "rate-error-pmt"  "=RATE(5,-100,A2)"       "DIV/0!"
errorIs "fv-error-pv"     "=FV(0.1,5,-100,A2)"     "DIV/0!"
errorIs "pmt-error-pv"    "=PMT(0.1,5,A2)"         "DIV/0!"
approxT "rate-empty-args" "=RATE(5,,-100,200,,)"   0.148698354992987 1e-9
// PMT to the last bit: (1 + r)^n not from 1 + r rounded. 750,000 over 300 months at 7.25% a year:
// Excel's payment, and the end of its schedule (balance - (payment - interest), 300 times) 2.5e-9.
exact "pmt-last-bit" "=PMT(0.0725/12,300,750000)" -5421.051480728742
let schedule =
    let pmt = match evaluateFormula "=-PMT(0.0725/12,300,750000)" model None with CellValue.Number p -> p | _ -> nan
    let mutable balance = 750000.0
    for _ in 1 .. 300 do balance <- balance - (pmt - balance * 0.0725 / 12.0)
    balance
if abs (schedule - 2.51657183980569e-09) < 1e-22 then pass <- pass + 1
else (fail <- fail + 1; printfn "FAIL pmt-schedule-end expected 2.51657183980569E-09 got %.15g" schedule)
// XIRR of a large outlay and little back soon: a rate just above -1, as Excel's (ROUND(.., 6)),
// and #NUM! where Excel finds none (the root 1e-13 above -1).
let xirrModel =
    let flows = [ -1000000.0; 6000.0; 6000.0; 6000.0; 6000.0; 6000.0; 6000.0; 6000.0 ]
    let dates = [ 45658.0; 45689.0; 45717.0; 45748.0; 45778.0; 45809.0; 45839.0; 45870.0 ]
    { createEmpty 100 26 with
        Cells =
            (flows |> List.mapi (fun i v -> { Row = i; Col = 0 }, cell (CellValue.Number v)))
            @ (dates |> List.mapi (fun i v -> { Row = i; Col = 1 }, cell (CellValue.Number v))) |> Map.ofList }
let xirrRounded k = evaluateFormula (sprintf "=ROUND(XIRR(A1:A%d,B1:B%d),6)" k k) xirrModel None
let xirrWanted = [ 4, -1.0; 5, -0.999999; 6, -0.999986; 7, -0.999883; 8, -0.999459 ]
for k, wanted in xirrWanted do
    match xirrRounded k with
    | CellValue.Number got when got = wanted -> pass <- pass + 1
    | got -> fail <- fail + 1; printfn "FAIL xirr-near-minus-1 (%d flows) expected %g got %A" k wanted got
match evaluateFormula "=XIRR(A1:A3,B1:B3)" xirrModel None with
| CellValue.Error "NUM!" -> pass <- pass + 1
| got -> fail <- fail + 1; printfn "FAIL xirr-too-near-minus-1 expected #NUM! got %A" got

printfn ""
printfn "Stage 12 financial tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
