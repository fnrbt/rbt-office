#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// Batch 2 — LAMBDA + higher-order helpers (MAP/REDUCE/SCAN/MAKEARRAY).
// Lambdas are first-class via beta-reduction (expression substitution),
// consistent with LET: a LET-bound lambda call `f(5)` and a lambda passed to
// MAP/REDUCE/… are reduced and evaluated. (Direct IIFE `LAMBDA(x,..)(5)` is
// deferred — needs a parser apply-node.)

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

let cell v = { Value = v; Format = defaultFormat; Comment = None }
// A1:A4 = 1,2,3,4
let model =
    { createEmpty 100 26 with
        Cells = [ for r in 0..3 -> { Row=r; Col=0 }, cell (CellValue.Number (float (r+1))) ] |> Map.ofList }
let ctx = contextOf model

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-26s expected %A got %A" label expected got
let arrOf f = match parseFormula f with Ok e -> evalArray e ctx | Result.Error e -> failwithf "parse: %s" e
let nums (a: CellValue[][]) = [ for row in a do for v in row -> match v with CellValue.Number n -> n | _ -> nan ]
let shape (a: CellValue[][]) = (a.Length, (if a.Length > 0 then a.[0].Length else 0))
let ev f = evaluateFormula f model None
let n x = CellValue.Number x

// ── LET-bound lambdas (beta-reduced calls) ────────────────────────────
check "let-lambda call"      (ev "=LET(f,LAMBDA(x,x*2),f(5))")          (n 10.0)
check "let-lambda 2 params"  (ev "=LET(add,LAMBDA(a,b,a+b),add(3,4))")  (n 7.0)
check "let-lambda twice"     (ev "=LET(sq,LAMBDA(x,x*x),sq(3)+sq(4))")  (n 25.0)
check "let-lambda chained"   (ev "=LET(inc,LAMBDA(x,x+1),dbl,LAMBDA(x,x*2),dbl(inc(4)))") (n 10.0)
check "uncalled lambda"      (ev "=LAMBDA(x,x+1)")                      (CellValue.Error "CALC!")

// ── MAP ───────────────────────────────────────────────────────────────
check "map array"            (nums (arrOf "=MAP(A1:A4,LAMBDA(x,x*x))")) [1.;4.;9.;16.]
check "sum map"              (ev "=SUM(MAP(A1:A4,LAMBDA(x,x*x)))")      (n 30.0)
check "sum map +1"           (ev "=SUM(MAP(A1:A4,LAMBDA(x,x+1)))")      (n 14.0)
check "map via let-bound"    (ev "=LET(g,LAMBDA(x,x*3),SUM(MAP(A1:A4,g)))") (n 30.0)

// ── REDUCE ────────────────────────────────────────────────────────────
check "reduce sum"           (ev "=REDUCE(0,A1:A4,LAMBDA(a,v,a+v))")    (n 10.0)
check "reduce product"       (ev "=REDUCE(1,A1:A4,LAMBDA(a,v,a*v))")    (n 24.0)
check "reduce over array fn" (ev "=REDUCE(0,SEQUENCE(5),LAMBDA(a,v,a+v))") (n 15.0)

// ── SCAN ──────────────────────────────────────────────────────────────
check "scan running sum"     (nums (arrOf "=SCAN(0,A1:A4,LAMBDA(a,v,a+v))")) [1.;3.;6.;10.]

// ── MAKEARRAY ─────────────────────────────────────────────────────────
check "makearray shape"      (shape (arrOf "=MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))")) (2,3)
check "makearray vals"       (nums (arrOf "=MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))")) [11.;12.;13.;21.;22.;23.]
check "sum makearray"        (ev "=SUM(MAKEARRAY(2,3,LAMBDA(r,c,r*10+c)))") (n 102.0)

// ── composition: MAP of a LET-bound lambda producing into REDUCE ──────
check "reduce of map"        (ev "=REDUCE(0,MAP(A1:A4,LAMBDA(x,x*x)),LAMBDA(a,v,a+v))") (n 30.0)

printfn ""
printfn "Stage 9 lambda tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
