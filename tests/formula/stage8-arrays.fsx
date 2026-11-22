#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// Batch 2, Step 1 — dynamic arrays. Array functions (SEQUENCE/RANDARRAY/UNIQUE/
// SORT/SORTBY/FILTER/TRANSPOSE) + element-wise operator broadcasting. Arrays
// flow through the engine: consumable by aggregates, collapsing to top-left in
// scalar context. (Spill placement into cells + render is Step 2.)

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let nC r c v = ({ Row=r; Col=c }, cell (CellValue.Number v))
let sC r c v = ({ Row=r; Col=c }, cell (CellValue.String v))
// A1:A8 numbers (with dups), B1:B8 = A*10, C1:C3 text (mixed case)
let model =
    { createEmpty 100 26 with
        Cells =
            ([ 3.;1.;2.;3.;1.;5.;2.;4. ] |> List.mapi (fun i v -> nC i 0 v))
            @ ([ 3.;1.;2.;3.;1.;5.;2.;4. ] |> List.mapi (fun i v -> nC i 1 (v*10.)))
            @ [ sC 0 2 "a"; sC 1 2 "A"; sC 2 2 "b" ]
            |> Map.ofList }
let ctx = contextOf model

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-26s expected %A got %A" label expected got

let arrOf f = match parseFormula f with Ok e -> evalArray e ctx | Result.Error e -> failwithf "parse: %s" e
let shape (a: CellValue[][]) = (a.Length, (if a.Length > 0 then a.[0].Length else 0))
let nums (a: CellValue[][]) = [ for row in a do for v in row -> match v with CellValue.Number n -> n | _ -> nan ]
let ev f = evaluateFormula f model None
let n x = CellValue.Number x

// ── array shapes + values ──────────────────────────────────────────────
check "sequence shape"   (shape (arrOf "=SEQUENCE(4)"))       (4,1)
check "sequence col"     (nums (arrOf "=SEQUENCE(4)"))        [1.;2.;3.;4.]
check "sequence 2x3"     (shape (arrOf "=SEQUENCE(2,3,10,5)")) (2,3)
check "sequence 2x3 vals" (nums (arrOf "=SEQUENCE(2,3,10,5)")) [10.;15.;20.;25.;30.;35.]
check "unique order"     (nums (arrOf "=UNIQUE(A1:A8)"))      [3.;1.;2.;5.;4.]
check "sort asc"         (nums (arrOf "=SORT(A1:A8)"))        [1.;1.;2.;2.;3.;3.;4.;5.]
check "sort desc"        (nums (arrOf "=SORT(A1:A8,1,-1)"))   [5.;4.;3.;3.;2.;2.;1.;1.]
check "filter >2"        (nums (arrOf "=FILTER(A1:A8,A1:A8>2)")) [3.;3.;5.;4.]
check "filter empty"     (arrOf "=FILTER(A1:A8,A1:A8>100)")   [| [| CellValue.Error "CALC!" |] |]
check "filter if-empty"  (arrOf "=FILTER(A1:A8,A1:A8>100,\"none\")") [| [| CellValue.String "none" |] |]
check "sortby B desc"    (nums (arrOf "=SORTBY(A1:A8,B1:B8,-1)")) [5.;4.;3.;3.;2.;2.;1.;1.]
check "transpose"        (shape (arrOf "=TRANSPOSE(SEQUENCE(1,4))")) (4,1)
check "unique case-insens" (arrOf "=UNIQUE(C1:C3)") [| [| CellValue.String "a" |]; [| CellValue.String "b" |] |]

// broadcasting: A1:A8>2 is a boolean array of trues for 3,3,5,4
check "broadcast bools" ((arrOf "=A1:A8>2") |> Array.sumBy (fun r -> if r.[0] = CellValue.Boolean true then 1 else 0)) 4
check "broadcast mul"   (nums (arrOf "=SEQUENCE(3)*2")) [2.;4.;6.]

// The optimized scalar/range folds and array broadcasts use the same operator
// semantics as scalar formulas, including mixed boolean and error values.
let smallestNormal = 2.0 ** -1022.0
let tinyModel =
    { createEmpty 8 2 with
        Cells = Map.ofList
            [ { Row = 0; Col = 0 }, cell (n smallestNormal)
              { Row = 1; Col = 0 }, cell (n (smallestNormal / 2.0))
              { Row = 2; Col = 0 }, cell (CellValue.Boolean true)
              { Row = 3; Col = 0 }, cell (CellValue.Error "N/A") ] }
let tinyArray formula =
    match parseFormula formula with
    | Ok expr -> evalArray expr (contextOf tinyModel)
    | Result.Error message -> failwithf "parse %s: %s" formula message
check "array subnormal, bool, error" (tinyArray "=A1:A4*0.5")
    [| [| n 0.0 |]; [| n 0.0 |]; [| n 0.5 |]; [| CellValue.Error "N/A" |] |]
check "array normal after division" (tinyArray "=A1:A2/0.5")
    [| [| n (smallestNormal * 2.0) |]; [| n smallestNormal |] |]
check "array subnormal divisor" (tinyArray "=A1:A2/A2")
    [| [| CellValue.Error "DIV/0!" |]; [| CellValue.Error "DIV/0!" |] |]
check "array power underflow" (tinyArray "=A1:A2^1")
    [| [| n smallestNormal |]; [| n 0.0 |] |]
check "optimized SUM range product" (evaluateFormula "=SUM(A1:A3*0.5)" tinyModel None) (n 0.5)
check "optimized SUM error" (evaluateFormula "=SUM(A1:A4*0.5)" tinyModel None) (CellValue.Error "N/A")

// ── consumed by aggregates (streaming foldArgs over the array) ─────────
check "sum sequence"        (ev "=SUM(SEQUENCE(10))")            (n 55.0)
check "sum sequence 3x3"    (ev "=SUM(SEQUENCE(3,3))")           (n 45.0)
check "count unique"        (ev "=COUNT(UNIQUE(A1:A8))")         (n 5.0)
check "sum filter parallel" (ev "=SUM(FILTER(B1:B8,A1:A8>2))")   (n 150.0) // B where A>2: 30+30+50+40
check "average sort"        (ev "=AVERAGE(SORT(A1:A8))")         (n 2.625) // 21/8
check "sum broadcast expr"  (ev "=SUM(SEQUENCE(5)*2)")           (n 30.0)  // 2+4+6+8+10
check "sum range plus range" (ev "=SUM(A1:A8+B1:B8)")            (n 231.0) // 21 + 210

// ── scalar context collapses to top-left ──────────────────────────────
check "scalar sequence"  (ev "=SEQUENCE(5)")     (n 1.0)
check "scalar sort"      (ev "=SORT(A1:A8)")     (n 1.0)   // smallest, top-left after sort
check "scalar unique+1"  (ev "=UNIQUE(A1:A8)+10") (n 13.0) // top-left 3, +10

printfn ""
printfn "Stage 8 dynamic-array tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
