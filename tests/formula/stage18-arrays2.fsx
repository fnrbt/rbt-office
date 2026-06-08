#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// Phase A: BYROW/BYCOL, by_col UNIQUE/SORT, multi-key SORTBY, LAMBDA IIFE.

open RedBat.Spreadsheet.Model
module F = RedBat.Spreadsheet.Formula

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let N = CellValue.Number
let S = CellValue.String

// Grid:  A1:C3 numbers          E1:E5 for unique/sort        G col keys
//  1 2 3
//  4 5 6
//  7 8 9
let cells =
    [ for r in 0..2 do for c in 0..2 -> { Row=r; Col=c }, cell (N (float (r*3 + c + 1))) ]
    @ [ { Row=0; Col=4 }, cell (S "b"); { Row=1; Col=4 }, cell (S "a"); { Row=2; Col=4 }, cell (S "b")
        { Row=3; Col=4 }, cell (S "c"); { Row=4; Col=4 }, cell (S "a") ]
    // multi-key sort source: col G = group, col H = value, col I = label
    @ [ { Row=0; Col=6 }, cell (S "x"); { Row=0; Col=7 }, cell (N 2.); { Row=0; Col=8 }, cell (S "x2")
        { Row=1; Col=6 }, cell (S "x"); { Row=1; Col=7 }, cell (N 1.); { Row=1; Col=8 }, cell (S "x1")
        { Row=2; Col=6 }, cell (S "y"); { Row=2; Col=7 }, cell (N 5.); { Row=2; Col=8 }, cell (S "y5") ]
    |> Map.ofList
let model = { createEmpty 20 12 with Cells = cells }
let ctx : F.EvalContext = { Model = model; Workbook = None; Sheet = None; CurrentCell = None }

let arr (s: string) : CellValue[][] =
    match F.parseFormula s with Ok e -> F.evalArray e ctx | Result.Error m -> [| [| CellValue.Error m |] |]
let scalar (s: string) = match F.parseFormula s with Ok e -> F.eval e ctx | Result.Error m -> CellValue.Error m
let nums (a: CellValue[][]) = a |> Array.map (Array.map (function CellValue.Number n -> n | _ -> nan))
let strs (a: CellValue[][]) = a |> Array.map (Array.map (function CellValue.String s -> s | CellValue.Number n -> string n | _ -> "?"))

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-26s expected %A got %A" label expected got

// --- BYROW / BYCOL ---
check "BYROW sum rows"  (nums (arr "BYROW(A1:C3,LAMBDA(r,SUM(r)))")) [| [|6.|]; [|15.|]; [|24.|] |]
check "BYCOL sum cols"  (nums (arr "BYCOL(A1:C3,LAMBDA(c,SUM(c)))")) [| [|12.; 15.; 18.|] |]
check "BYROW max"       (nums (arr "BYROW(A1:C3,LAMBDA(r,MAX(r)))")) [| [|3.|]; [|6.|]; [|9.|] |]

// --- LAMBDA IIFE ---
check "IIFE scalar"     (scalar "LAMBDA(x,x+1)(5)") (N 6.0)
check "IIFE two args"   (scalar "LAMBDA(a,b,a*b)(6,7)") (N 42.0)
check "IIFE nested"     (scalar "LAMBDA(x,LAMBDA(y,x+y)(10))(5)") (N 15.0)

// --- by_col UNIQUE / SORT ---
// E1:E5 = b,a,b,c,a  → unique (row-wise) = b,a,c
check "UNIQUE col"      (strs (arr "UNIQUE(E1:E5)")) [| [|"b"|]; [|"a"|]; [|"c"|] |]
// by_col on a single row: A1:C1 = 1,2,3 distinct cols stay 1,2,3
check "UNIQUE by_col"   (nums (arr "UNIQUE(A1:C1,TRUE)")) [| [|1.;2.;3.|] |]
// SORT E1:E5 ascending → a,a,b,b,c
check "SORT col"        (strs (arr "SORT(E1:E5)")) [| [|"a"|]; [|"a"|]; [|"b"|]; [|"b"|]; [|"c"|] |]
// SORT a single row by_col: 3,1,2 → 1,2,3
let row312 = { model with Cells = (model.Cells |> Map.add {Row=10;Col=0} (cell (N 3.)) |> Map.add {Row=10;Col=1} (cell (N 1.)) |> Map.add {Row=10;Col=2} (cell (N 2.))) }
let arr2 (m: SpreadsheetModel) (s: string) = match F.parseFormula s with Ok e -> F.evalArray e { ctx with Model = m } | Result.Error msg -> [| [| CellValue.Error msg |] |]
check "SORT by_col"     (nums (arr2 row312 "SORT(A11:C11,1,1,TRUE)")) [| [|1.;2.;3.|] |]

// --- multi-key SORTBY ---
// rows (G,H,I): x/2/x2, x/1/x1, y/5/y5. SORTBY(I, G asc, H asc) → x1, x2, y5
check "SORTBY multi-key" (strs (arr "SORTBY(I1:I3,G1:G3,1,H1:H3,1)")) [| [|"x1"|]; [|"x2"|]; [|"y5"|] |]
// single-key still works: SORTBY(I, H desc) → y5(5), x2(2), x1(1)
check "SORTBY single desc" (strs (arr "SORTBY(I1:I3,H1:H3,-1)")) [| [|"y5"|]; [|"x2"|]; [|"x1"|] |]

// --- A1# spill-range operator ---
// K1 = SEQUENCE(3) spills down to K1:K3 = 1,2,3; K1# references the spill.
let spillModel =
    { createEmpty 20 12 with
        Cells = Map.ofList [ { Row=0; Col=10 }, cell (CellValue.Formula ("SEQUENCE(3)", N 1.)) ] }
let spilled = F.recalcAndSpill None None spillModel
let evS (s: string) = match F.parseFormula s with Ok e -> F.eval e { ctx with Model = spilled } | Result.Error m -> CellValue.Error m
let arrS (s: string) = match F.parseFormula s with Ok e -> F.evalArray e { ctx with Model = spilled } | Result.Error m -> [| [| CellValue.Error m |] |]
check "spill materialized K3" (getCellValue { Row=2; Col=10 } spilled) (N 3.0)
check "A1# sum"    (evS "SUM(K1#)") (N 6.0)
check "A1# count"  (evS "COUNT(K1#)") (N 3.0)
check "A1# array"  (nums (arrS "K1#")) [| [|1.|]; [|2.|]; [|3.|] |]
check "A1# in expr" (evS "SUM(K1#)*2") (N 12.0)

printfn ""
printfn "Stage 18 array2 tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
