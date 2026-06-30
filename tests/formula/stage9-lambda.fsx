#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// Batch 2 — LAMBDA + higher-order helpers (MAP/REDUCE/SCAN/MAKEARRAY).
// Lambdas are first-class via beta-reduction (expression substitution),
// consistent with LET: a LET-bound lambda call `f(5)` and a lambda passed to
// MAP/REDUCE/… are reduced and evaluated. (Direct IIFE `LAMBDA(x,..)(5)` is
// deferred — needs a parser apply-node.)

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Formula

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

let sparseMapModel =
    { createEmpty 20010 2 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 9999; Col = 0 }, cell (CellValue.Number 3.0)
              { Row = 19999; Col = 0 }, cell (CellValue.Number 4.0) ]
            |> Map.ofList }
let evSparseMap f = evaluateFormula f sparseMapModel None
let arrSparseMap f =
    match parseFormula f with
    | Ok expr -> evalArray expr { Model = sparseMapModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None }
    | Result.Error e -> [| [| CellValue.Error e |] |]
let sparseMapPlus = arrSparseMap "=MAP(A1:A20000,LAMBDA(x,x+1))"
check "map sparse plus first" sparseMapPlus.[0].[0] (n 3.0)
check "map sparse plus blank" sparseMapPlus.[1].[0] (n 1.0)
check "map sparse plus later" sparseMapPlus.[9999].[0] (n 4.0)
check "map sparse plus tail" sparseMapPlus.[19999].[0] (n 5.0)
let sparseMapSquare = arrSparseMap "=MAP(A1:A20000,LAMBDA(x,x*x))"
check "map sparse square blank" sparseMapSquare.[1].[0] (n 0.0)
check "map sparse square later" sparseMapSquare.[9999].[0] (n 9.0)
check "scalar map sparse plus" (evSparseMap "=MAP(A1:A20000,LAMBDA(x,x+1))") (n 3.0)
check "sum map sparse square" (evSparseMap "=SUM(MAP(A1:A20000,LAMBDA(x,x*x)))") (n 29.0)
check "sum map sparse blanks" (evSparseMap "=SUM(MAP(A1:A20000,LAMBDA(x,x+1)))") (n 20009.0)
check "average map sparse blanks" (evSparseMap "=AVERAGE(MAP(A1:A20000,LAMBDA(x,x+1)))") (n 1.00045)
check "count map sparse numbers" (evSparseMap "=COUNT(MAP(A1:A20000,LAMBDA(x,x+1)))") (n 20000.0)
check "count map sparse booleans" (evSparseMap "=COUNT(MAP(A1:A20000,LAMBDA(x,x=0)))") (n 0.0)
check "min map sparse blanks" (evSparseMap "=MIN(MAP(A1:A20000,LAMBDA(x,x+1)))") (n 1.0)
check "max map sparse blanks" (evSparseMap "=MAX(MAP(A1:A20000,LAMBDA(x,x+1)))") (n 5.0)
check "sum transpose map sparse" (evSparseMap "=SUM(TRANSPOSE(MAP(A1:A20000,LAMBDA(x,x+1))))") (n 20009.0)
let longSparseMapTermCount = 60000
let longSparseMapBody = String.concat "+" [ for _ in 1..longSparseMapTermCount -> "x" ]
let longSparseMapFormula = "=SUM(MAP(A1:A20000,LAMBDA(x," + longSparseMapBody + ")))"
check "sum map sparse long body" (evSparseMap longSparseMapFormula) (n (float longSparseMapTermCount * 9.0))

let sparseMapErrorModel =
    { createEmpty 20010 2 with
        Cells = [ { Row = 99; Col = 0 }, cell (CellValue.Number 2.0) ] |> Map.ofList }
let evSparseMapError f = evaluateFormula f sparseMapErrorModel None
let arrSparseMapError f =
    match parseFormula f with
    | Ok expr -> evalArray expr { Model = sparseMapErrorModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None }
    | Result.Error e -> [| [| CellValue.Error e |] |]
let sparseMapBlankError = arrSparseMapError "=MAP(A1:A20000,LAMBDA(x,1/x))"
check "map sparse blank error" sparseMapBlankError.[0].[0] (CellValue.Error "DIV/0!")
check "map sparse value after error gap" sparseMapBlankError.[99].[0] (n 0.5)
check "sum map sparse blank error" (evSparseMapError "=SUM(MAP(A1:A20000,LAMBDA(x,1/x)))") (CellValue.Error "DIV/0!")

// ── REDUCE ────────────────────────────────────────────────────────────
check "reduce sum"           (ev "=REDUCE(0,A1:A4,LAMBDA(a,v,a+v))")    (n 10.0)
check "reduce product"       (ev "=REDUCE(1,A1:A4,LAMBDA(a,v,a*v))")    (n 24.0)
check "reduce over array fn" (ev "=REDUCE(0,SEQUENCE(5),LAMBDA(a,v,a+v))") (n 15.0)
check "reduce sparse sum"    (evSparseMap "=REDUCE(0,A1:A20000,LAMBDA(a,v,a+v))") (n 9.0)
check "reduce sparse count"  (evSparseMap "=REDUCE(0,A1:A20000,LAMBDA(a,v,a+1))") (n 20000.0)
check "reduce sparse blank error" (evSparseMapError "=REDUCE(0,A1:A20000,LAMBDA(a,v,a+1/v))") (CellValue.Error "DIV/0!")

let sparseReduceErrorModel =
    { createEmpty 20010 2 with
        Cells = [ { Row = 99; Col = 0 }, cell (CellValue.Error "N/A") ] |> Map.ofList }
let evSparseReduceError f = evaluateFormula f sparseReduceErrorModel None
check "reduce sparse cell error" (evSparseReduceError "=REDUCE(0,A1:A20000,LAMBDA(a,v,a+v))") (CellValue.Error "N/A")

// ── SCAN ──────────────────────────────────────────────────────────────
check "scan running sum"     (nums (arrOf "=SCAN(0,A1:A4,LAMBDA(a,v,a+v))")) [1.;3.;6.;10.]
let sparseScanSum = arrSparseMap "=SCAN(0,A1:A20000,LAMBDA(a,v,a+v))"
check "scan sparse sum first" sparseScanSum.[0].[0] (n 2.0)
check "scan sparse sum blank" sparseScanSum.[1].[0] (n 2.0)
check "scan sparse sum later" sparseScanSum.[9999].[0] (n 5.0)
check "scan sparse sum tail" sparseScanSum.[19999].[0] (n 9.0)
let sparseScanCount = arrSparseMap "=SCAN(0,A1:A20000,LAMBDA(a,v,a+1))"
check "scan sparse count first" sparseScanCount.[0].[0] (n 1.0)
check "scan sparse count tail" sparseScanCount.[19999].[0] (n 20000.0)
check "scalar scan sparse sum" (evSparseMap "=SCAN(0,A1:A20000,LAMBDA(a,v,a+v))") (n 2.0)
check "sum scan sparse sum" (evSparseMap "=SUM(SCAN(0,A1:A20000,LAMBDA(a,v,a+v)))") (n 70007.0)
check "average scan sparse sum" (evSparseMap "=AVERAGE(SCAN(0,A1:A20000,LAMBDA(a,v,a+v)))") (n 3.50035)
check "count scan sparse sum" (evSparseMap "=COUNT(SCAN(0,A1:A20000,LAMBDA(a,v,a+v)))") (n 20000.0)
check "min scan sparse sum" (evSparseMap "=MIN(SCAN(0,A1:A20000,LAMBDA(a,v,a+v)))") (n 2.0)
check "max scan sparse sum" (evSparseMap "=MAX(SCAN(0,A1:A20000,LAMBDA(a,v,a+v)))") (n 9.0)
check "sum scan sparse count" (evSparseMap "=SUM(SCAN(0,A1:A20000,LAMBDA(a,v,a+1)))") (n 200010000.0)
check "count scan sparse count" (evSparseMap "=COUNT(SCAN(0,A1:A20000,LAMBDA(a,v,a+1)))") (n 20000.0)
let sparseScanError = arrSparseMapError "=SCAN(0,A1:A20000,LAMBDA(a,v,a+1/v))"
check "scan sparse blank error" sparseScanError.[0].[0] (CellValue.Error "DIV/0!")
check "scan sparse error persists" sparseScanError.[99].[0] (CellValue.Error "DIV/0!")
check "sum scan sparse cell error" (evSparseReduceError "=SUM(SCAN(0,A1:A20000,LAMBDA(a,v,a+v)))") (CellValue.Error "N/A")

// ── MAKEARRAY ─────────────────────────────────────────────────────────
check "makearray shape"      (shape (arrOf "=MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))")) (2,3)
check "makearray vals"       (nums (arrOf "=MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))")) [11.;12.;13.;21.;22.;23.]
check "sum makearray"        (ev "=SUM(MAKEARRAY(2,3,LAMBDA(r,c,r*10+c)))") (n 102.0)
check "scalar makearray large" (ev "=MAKEARRAY(20000,2,LAMBDA(r,c,r+c))") (n 2.0)
check "sum makearray sparse const" (evSparseMap "=SUM(MAKEARRAY(20000,1,LAMBDA(r,c,1)))") (n 20000.0)
check "count makearray sparse const" (evSparseMap "=COUNT(MAKEARRAY(20000,1,LAMBDA(r,c,1)))") (n 20000.0)
check "average makearray sparse linear" (evSparseMap "=AVERAGE(MAKEARRAY(20000,1,LAMBDA(r,c,r+c)))") (n 10001.5)
check "max makearray sparse linear" (evSparseMap "=MAX(MAKEARRAY(20000,1,LAMBDA(r,c,r+c)))") (n 20001.0)
check "sum makearray sparse affine" (evSparseMap "=SUM(MAKEARRAY(20000,1,LAMBDA(r,c,r*10+c)))") (n 2000120000.0)
let longMakeArrayBody = String.concat "+" [ for _ in 1..longSparseMapTermCount -> "r" ]
let longMakeArrayFormula = "=SUM(MAKEARRAY(20000,1,LAMBDA(r,c," + longMakeArrayBody + ")))"
let longMakeArrayExpected = float longSparseMapTermCount * 20000.0 * 20001.0 / 2.0
check "sum makearray long body" (evSparseMap longMakeArrayFormula) (n longMakeArrayExpected)
let longMakeArrayScaledFormula = "=SUM(MAKEARRAY(20000,1,LAMBDA(r,c,(" + longMakeArrayBody + ")*2)))"
check "sum makearray scaled long body" (evSparseMap longMakeArrayScaledFormula) (n (longMakeArrayExpected * 2.0))
let longMakeArrayWrapperSuffix = String.concat "" [ for _ in 1..longSparseMapTermCount -> "+1" ]
let longMakeArrayWrapperFormula = "=SUM(MAKEARRAY(20000,1,LAMBDA(r,c,r))" + longMakeArrayWrapperSuffix + ")"
let longMakeArrayWrapperExpected = 20000.0 * 20001.0 / 2.0 + float longSparseMapTermCount * 20000.0
check "sum makearray wrapper long body" (evSparseMap longMakeArrayWrapperFormula) (n longMakeArrayWrapperExpected)
let longMakeArrayMultiplierSuffix = String.concat "" [ for _ in 1..longSparseMapTermCount -> "*1" ]
let longMakeArrayMultiplierFormula = "=SUM(MAKEARRAY(20000,1,LAMBDA(r,c,r))" + longMakeArrayMultiplierSuffix + ")"
check "sum makearray multiply wrapper long" (evSparseMap longMakeArrayMultiplierFormula) (n 200010000.0)
let longMakeArraySquareMultiplierFormula = "=SUM(MAKEARRAY(20000,1,LAMBDA(r,c,r^2))" + longMakeArrayMultiplierSuffix + ")"
let longMakeArraySquareExpected = 20000.0 * 20001.0 * 40001.0 / 6.0
check "sum makearray stream multiply long" (evSparseMap longMakeArraySquareMultiplierFormula) (n longMakeArraySquareExpected)
let longMakeArrayUnaryPrefix = String.replicate longSparseMapTermCount "+"
let longMakeArrayUnaryFormula = "=SUM(" + longMakeArrayUnaryPrefix + "MAKEARRAY(20000,1,LAMBDA(r,c,r)))"
check "sum makearray unary wrapper long" (evSparseMap longMakeArrayUnaryFormula) (n 200010000.0)
let longMakeArraySquareUnaryFormula = "=SUM(" + longMakeArrayUnaryPrefix + "MAKEARRAY(20000,1,LAMBDA(r,c,r^2)))"
check "sum makearray stream unary long" (evSparseMap longMakeArraySquareUnaryFormula) (n longMakeArraySquareExpected)
let longMakeArrayPowerBody = String.concat "+" [ for _ in 1..longSparseMapTermCount -> "r^1" ]
let longMakeArrayPowerFormula = "=SUM(MAKEARRAY(20000,1,LAMBDA(r,c," + longMakeArrayPowerBody + ")))"
check "sum makearray power long body" (evSparseMap longMakeArrayPowerFormula) (n longMakeArrayExpected)
let longMakeArrayCompareBody = String.concat "+" [ for _ in 1..longSparseMapTermCount -> "(r=r)" ]
let longMakeArrayCompareFormula = "=SUM(MAKEARRAY(20000,1,LAMBDA(r,c," + longMakeArrayCompareBody + ")))"
check "sum makearray compare long body" (evSparseMap longMakeArrayCompareFormula) (n (float longSparseMapTermCount * 20000.0))
check "count makearray direct compare" (evSparseMap "=COUNT(MAKEARRAY(20000,1,LAMBDA(r,c,r=r)))") (n 0.0)
check "sum makearray sparse product" (evSparseMap "=SUM(MAKEARRAY(20000,1,LAMBDA(r,c,r*c)))") (n 200010000.0)
check "product makearray sparse const" (evSparseMap "=PRODUCT(MAKEARRAY(20000,1,LAMBDA(r,c,1)))") (n 1.0)
check "sum makearray sparse error" (evSparseMap "=SUM(MAKEARRAY(20000,1,LAMBDA(r,c,1/(c-c))))") (CellValue.Error "DIV/0!")
check "sum makearray plus scalar large" (evSparseMap "=SUM(MAKEARRAY(20000,1,LAMBDA(r,c,1))+1)") (n 40000.0)
check "sum scalar plus makearray large" (evSparseMap "=SUM(1+MAKEARRAY(20000,1,LAMBDA(r,c,r+c)))") (n 200050000.0)
check "average makearray times scalar large" (evSparseMap "=AVERAGE(MAKEARRAY(20000,1,LAMBDA(r,c,r+c))*2)") (n 20003.0)
check "count makearray divide scalar large" (evSparseMap "=COUNT(MAKEARRAY(20000,1,LAMBDA(r,c,r+c))/2)") (n 20000.0)
check "max unary minus makearray large" (evSparseMap "=MAX(-MAKEARRAY(20000,1,LAMBDA(r,c,r+c)))") (n -2.0)
check "sum makearray compare large" (evSparseMap "=SUM(MAKEARRAY(20000,1,LAMBDA(r,c,r+c))>0)") (n 20000.0)
check "count makearray compare large" (evSparseMap "=COUNT(MAKEARRAY(20000,1,LAMBDA(r,c,r+c))>0)") (n 0.0)
check "sum sort makearray large" (evSparseMap "=SUM(SORT(MAKEARRAY(20000,1,LAMBDA(r,c,r+c))))") (n 200030000.0)

// ── composition: MAP of a LET-bound lambda producing into REDUCE ──────
check "reduce of map"        (ev "=REDUCE(0,MAP(A1:A4,LAMBDA(x,x*x)),LAMBDA(a,v,a+v))") (n 30.0)

printfn ""
printfn "Stage 9 lambda tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
