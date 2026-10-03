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

printfn ""
printfn "Stage 12 financial tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
