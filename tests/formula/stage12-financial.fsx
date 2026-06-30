#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// Financial (time-value-of-money) functions: PMT, FV, PV, NPER, RATE, NPV,
// IRR, SLN. Expected values match Excel.

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Formula

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
let sparseNpvModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 100.0)
              { Row = 19999; Col = 0 }, cell (CellValue.Number 100.0)
              { Row = 0; Col = 1 }, cell (CellValue.Number 10.0)
              { Row = 19999; Col = 1 }, cell (CellValue.Number 20.0) ]
            |> Map.ofList }
match evaluateFormula "=NPV(0.1,A1:A20000)" sparseNpvModel None with
| CellValue.Number value when abs (value - (100.0 / 1.1 + 100.0 / 1.21)) < 1e-9 -> pass <- pass + 1
| got -> fail <- fail + 1; printfn "FAIL %-20s %-34s expected sparse blank NPV got %A" "npv-sparse" "=NPV(0.1,A1:A20000)" got
let checkSparseNpv label formula expected =
    match evaluateFormula formula sparseNpvModel None with
    | CellValue.Number value when abs (value - expected) < 1e-6 -> pass <- pass + 1
    | got -> fail <- fail + 1; printfn "FAIL %-20s %-34s expected %.10f got %A" label formula expected got
checkSparseNpv "npv-sparse-plus" "=NPV(0.1,A1:A20000+1)" 100.9090909
checkSparseNpv "npv-sparse-scale" "=NPV(0.1,A1:A20000*2)" 181.8181818
checkSparseNpv "npv-sparse-product" "=NPV(0.1,A1:A20000*B1:B20000)" 909.0909091
checkSparseNpv "npv-sparse-unary" "=NPV(0.1,-A1:A20000)" -90.90909091
let sparseIrrModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number -100.0)
              { Row = 9999; Col = 0 }, cell (CellValue.Number 50.0)
              { Row = 19999; Col = 0 }, cell (CellValue.Number 60.0)
              { Row = 0; Col = 1 }, cell (CellValue.Number 10.0)
              { Row = 9999; Col = 1 }, cell (CellValue.Number 20.0)
              { Row = 19999; Col = 1 }, cell (CellValue.Number 30.0) ]
            |> Map.ofList }
let checkSparseIrrError label formula =
    match evaluateFormula formula sparseIrrModel None with
    | CellValue.Error "NUM!" -> pass <- pass + 1
    | got -> fail <- fail + 1; printfn "FAIL %-20s %-34s expected NUM! got %A" label formula got
checkSparseIrrError "irr-sparse-plus" "=IRR(A1:A20000+1)"
checkSparseIrrError "irr-sparse-scale" "=IRR(A1:A20000*2)"
checkSparseIrrError "irr-sparse-product" "=IRR(A1:A20000*B1:B20000)"
checkSparseIrrError "irr-sparse-unary" "=IRR(-A1:A20000)"

printfn ""
printfn "Stage 12 financial tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
