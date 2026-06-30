#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// Batch 2, Step 1 — dynamic arrays. Array functions (SEQUENCE/RANDARRAY/UNIQUE/
// SORT/SORTBY/FILTER/TRANSPOSE) + element-wise operator broadcasting. Arrays
// flow through the engine: consumable by aggregates, collapsing to top-left in
// scalar context. (Spill placement into cells + render is Step 2.)

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Formula

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
let approx label got expected =
    match got with
    | CellValue.Number g when abs (g - expected) < 1e-9 -> pass <- pass + 1
    | _ -> fail <- fail + 1; printfn "FAIL %-26s expected %A got %A" label expected got

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

let sparseUniqueModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 9999; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 9999; Col = 1 }, cell (CellValue.Number 20.0)
              { Row = 14999; Col = 0 }, cell (CellValue.Number 3.0)
              { Row = 14999; Col = 1 }, cell (CellValue.Number 30.0)
              { Row = 19999; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 19999; Col = 1 }, cell (CellValue.Number 20.0) ]
            |> Map.ofList }
let arrSparseUnique f =
    match parseFormula f with
    | Ok expr -> evalArray expr { Model = sparseUniqueModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None }
    | Result.Error e -> [| [| CellValue.Error e |] |]
check "unique sparse rows" (arrSparseUnique "=UNIQUE(A1:B20000)") [| [| CellValue.Empty; CellValue.Empty |]; [| n 2.0; n 20.0 |]; [| n 3.0; n 30.0 |] |]
check "unique sparse empty" (arrSparseUnique "=UNIQUE(D1:D20000)") [| [| CellValue.Empty |] |]
let sparseTranspose = arrSparseUnique "=TRANSPOSE(A1:B20000)"
check "transpose sparse shape" (shape sparseTranspose) (2, 20000)
check "transpose sparse blank" sparseTranspose.[0].[0] CellValue.Empty
check "transpose sparse first stored" sparseTranspose.[0].[9999] (n 2.0)
check "transpose sparse second stored" sparseTranspose.[1].[9999] (n 20.0)
check "transpose sparse tail" sparseTranspose.[1].[19999] (n 20.0)

let sparseByRowModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 1.0)
              { Row = 0; Col = 1 }, cell (CellValue.Boolean true)
              { Row = 9999; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 19999; Col = 1 }, cell (CellValue.Error "DIV/0!") ]
            |> Map.ofList }
let arrSparseByRow f =
    match parseFormula f with
    | Ok expr -> evalArray expr { Model = sparseByRowModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None }
    | Result.Error e -> [| [| CellValue.Error e |] |]
let evSparseByRow f = evaluateFormula f sparseByRowModel None
let sparseByRow = arrSparseByRow "=BYROW(A1:B20000,LAMBDA(r,SUM(r)))"
check "byrow sparse first" sparseByRow.[0] [| n 2.0 |]
check "byrow sparse blank" sparseByRow.[1] [| n 0.0 |]
check "byrow sparse later" sparseByRow.[9999] [| n 2.0 |]
check "byrow sparse error" sparseByRow.[19999] [| CellValue.Error "DIV/0!" |]
let sparseByRowCount = arrSparseByRow "=BYROW(A1:B20000,LAMBDA(r,COUNT(r)))"
check "byrow count sparse first" sparseByRowCount.[0] [| n 1.0 |]
check "byrow count sparse blank" sparseByRowCount.[1] [| n 0.0 |]
check "byrow count sparse later" sparseByRowCount.[9999] [| n 1.0 |]
check "byrow count sparse error" sparseByRowCount.[19999] [| CellValue.Error "DIV/0!" |]
let sparseByRowAverage = arrSparseByRow "=BYROW(A1:B20000,LAMBDA(r,AVERAGE(r)))"
check "byrow average sparse first" sparseByRowAverage.[0] [| n 1.0 |]
check "byrow average sparse blank" sparseByRowAverage.[1] [| CellValue.Error "DIV/0!" |]
check "byrow average sparse later" sparseByRowAverage.[9999] [| n 2.0 |]
check "byrow average sparse error" sparseByRowAverage.[19999] [| CellValue.Error "DIV/0!" |]
check "sum byrow sparse sums" (evSparseByRow "=SUM(BYROW(A1:B10000,LAMBDA(r,SUM(r))))") (n 4.0)
check "count byrow sparse sums" (evSparseByRow "=COUNT(BYROW(A1:B10000,LAMBDA(r,SUM(r))))") (n 10000.0)
check "sum byrow sparse counts" (evSparseByRow "=SUM(BYROW(A1:B10000,LAMBDA(r,COUNT(r))))") (n 2.0)
check "count byrow sparse counts" (evSparseByRow "=COUNT(BYROW(A1:B10000,LAMBDA(r,COUNT(r))))") (n 10000.0)
check "sum byrow sparse error" (evSparseByRow "=SUM(BYROW(A1:B20000,LAMBDA(r,SUM(r))))") (CellValue.Error "DIV/0!")
check "bycol sparse sums" (arrSparseByRow "=BYCOL(A1:B10000,LAMBDA(c,SUM(c)))") [| [| n 3.0; n 1.0 |] |]
check "bycol sparse error" (arrSparseByRow "=BYCOL(A1:B20000,LAMBDA(c,SUM(c)))") [| [| n 3.0; CellValue.Error "DIV/0!" |] |]
check "bycol count sparse values" (arrSparseByRow "=BYCOL(A1:B10000,LAMBDA(c,COUNT(c)))") [| [| n 2.0; n 0.0 |] |]
check "bycol count sparse error" (arrSparseByRow "=BYCOL(A1:B20000,LAMBDA(c,COUNT(c)))") [| [| n 2.0; CellValue.Error "DIV/0!" |] |]
check "bycol average sparse values" (arrSparseByRow "=BYCOL(A1:B10000,LAMBDA(c,AVERAGE(c)))") [| [| n 1.5; n 1.0 |] |]
check "bycol average sparse error" (arrSparseByRow "=BYCOL(A1:B20000,LAMBDA(c,AVERAGE(c)))") [| [| n 1.5; CellValue.Error "DIV/0!" |] |]

// broadcasting: A1:A8>2 is a boolean array of trues for 3,3,5,4
check "broadcast bools" ((arrOf "=A1:A8>2") |> Array.sumBy (fun r -> if r.[0] = CellValue.Boolean true then 1 else 0)) 4
check "broadcast mul"   (nums (arrOf "=SEQUENCE(3)*2")) [2.;4.;6.]

// ── consumed by aggregates (streaming foldArgs over the array) ─────────
check "sum sequence"        (ev "=SUM(SEQUENCE(10))")            (n 55.0)
check "sum sequence 3x3"    (ev "=SUM(SEQUENCE(3,3))")           (n 45.0)
check "sum sequence large"  (ev "=SUM(SEQUENCE(20000))")         (n 200010000.0)
check "count sequence large" (ev "=COUNT(SEQUENCE(20000))")       (n 20000.0)
check "average sequence large" (ev "=AVERAGE(SEQUENCE(20000))")   (n 10000.5)
check "max sequence descending" (ev "=MAX(SEQUENCE(20000,1,5,-1))") (n 5.0)
check "sum sequence plus scalar large" (ev "=SUM(SEQUENCE(20000)+1)") (n 200030000.0)
check "sum scalar plus sequence large" (ev "=SUM(1+SEQUENCE(20000))") (n 200030000.0)
let longSequenceWrapperTermCount = 60000
let longSequenceWrapperSuffix = String.concat "" [ for _ in 1..longSequenceWrapperTermCount -> "+1" ]
let longSequenceWrapperExpected = 20000.0 * 20001.0 / 2.0 + float longSequenceWrapperTermCount * 20000.0
check "sum sequence wrapper long" (ev ("=SUM(SEQUENCE(20000)" + longSequenceWrapperSuffix + ")")) (n longSequenceWrapperExpected)
let longSequenceMultiplierSuffix = String.concat "" [ for _ in 1..longSequenceWrapperTermCount -> "*1" ]
check "sum sequence multiply long" (ev ("=SUM(SEQUENCE(20000)" + longSequenceMultiplierSuffix + ")")) (n 200010000.0)
let longSequenceUnaryPrefix = String.replicate longSequenceWrapperTermCount "+"
check "sum sequence unary long" (ev ("=SUM(" + longSequenceUnaryPrefix + "SEQUENCE(20000))")) (n 200010000.0)
check "average sequence times scalar large" (ev "=AVERAGE(SEQUENCE(20000)*2)") (n 20001.0)
check "count sequence divide scalar large" (ev "=COUNT(SEQUENCE(20000)/2)") (n 20000.0)
check "max unary minus sequence large" (ev "=MAX(-SEQUENCE(20000))") (n -1.0)
check "sum sequence compare large" (ev "=SUM(SEQUENCE(20000)>0)") (n 20000.0)
check "count sequence compare large" (ev "=COUNT(SEQUENCE(20000)>0)") (n 0.0)
check "sum sort sequence large" (ev "=SUM(SORT(SEQUENCE(20000)))") (n 200010000.0)
check "sum randarray fixed large" (ev "=SUM(RANDARRAY(20000,1,1,1,TRUE))") (n 20000.0)
check "count randarray large" (ev "=COUNT(RANDARRAY(20000,1))") (n 20000.0)
check "average randarray fixed large" (ev "=AVERAGE(RANDARRAY(20000,1,7,7,TRUE))") (n 7.0)
check "min randarray fixed large" (ev "=MIN(RANDARRAY(20000,1,7,7,TRUE))") (n 7.0)
check "max randarray fixed large" (ev "=MAX(RANDARRAY(20000,1,7,7,TRUE))") (n 7.0)
check "sum randarray plus scalar large" (ev "=SUM(RANDARRAY(20000,1,1,1,TRUE)+1)") (n 40000.0)
check "sum scalar plus randarray large" (ev "=SUM(1+RANDARRAY(20000,1,1,1,TRUE))") (n 40000.0)
let longRandArrayWrapperSuffix = String.concat "" [ for _ in 1..longSequenceWrapperTermCount -> "+1" ]
let longRandArrayWrapperExpected = (1.0 + float longSequenceWrapperTermCount) * 20000.0
check "sum randarray wrapper long" (ev ("=SUM(RANDARRAY(20000,1,1,1,TRUE)" + longRandArrayWrapperSuffix + ")")) (n longRandArrayWrapperExpected)
let longRandArrayMultiplierSuffix = String.concat "" [ for _ in 1..longSequenceWrapperTermCount -> "*1" ]
check "sum randarray multiply long" (ev ("=SUM(RANDARRAY(20000,1,1,1,TRUE)" + longRandArrayMultiplierSuffix + ")")) (n 20000.0)
let longRandArrayUnaryPrefix = String.replicate longSequenceWrapperTermCount "+"
check "sum randarray unary long" (ev ("=SUM(" + longRandArrayUnaryPrefix + "RANDARRAY(20000,1,1,1,TRUE))")) (n 20000.0)
check "average randarray times scalar large" (ev "=AVERAGE(RANDARRAY(20000,1,2,2,TRUE)*3)") (n 6.0)
check "count randarray divide scalar large" (ev "=COUNT(RANDARRAY(20000,1,2,2,TRUE)/2)") (n 20000.0)
check "product randarray times scalar large" (ev "=PRODUCT(RANDARRAY(20000,1,1,1,TRUE)*1)") (n 1.0)
check "max unary minus randarray large" (ev "=MAX(-RANDARRAY(20000,1,2,2,TRUE))") (n -2.0)
check "count randarray invalid whole" (ev "=COUNT(RANDARRAY(20000,1,10,1,TRUE))") (CellValue.Error "VALUE!")
check "sum randarray equality large" (ev "=SUM(RANDARRAY(20000,1,1,1,TRUE)=1)") (n 20000.0)
check "sum randarray greater large" (ev "=SUM(RANDARRAY(20000,1,1,1,TRUE)>0)") (n 20000.0)
check "average randarray compare large" (ev "=AVERAGE(RANDARRAY(20000,1,1,1,TRUE)<=1)") (n 1.0)
check "min randarray compare false large" (ev "=MIN(RANDARRAY(20000,1,1,1,TRUE)<>1)") (n 0.0)
check "max scalar compare randarray large" (ev "=MAX(0<RANDARRAY(20000,1,1,1,TRUE))") (n 1.0)
check "count randarray compare large" (ev "=COUNT(RANDARRAY(20000,1,1,1,TRUE)=1)") (n 0.0)
check "sum sort randarray large" (ev "=SUM(SORT(RANDARRAY(20000,1,1,1,TRUE)))") (n 20000.0)
check "average sort randarray large" (ev "=AVERAGE(SORT(RANDARRAY(20000,1,7,7,TRUE)))") (n 7.0)
check "count sort randarray large" (ev "=COUNT(SORT(RANDARRAY(20000,1,7,7,TRUE)))") (n 20000.0)
check "min sort randarray large" (ev "=MIN(SORT(RANDARRAY(20000,1,7,7,TRUE)))") (n 7.0)
check "max sort randarray large" (ev "=MAX(SORT(RANDARRAY(20000,1,7,7,TRUE)))") (n 7.0)
check "count unique"        (ev "=COUNT(UNIQUE(A1:A8))")         (n 5.0)
check "sum filter parallel" (ev "=SUM(FILTER(B1:B8,A1:A8>2))")   (n 150.0) // B where A>2: 30+30+50+40
check "average sort"        (ev "=AVERAGE(SORT(A1:A8))")         (n 2.625) // 21/8
check "sum broadcast expr"  (ev "=SUM(SEQUENCE(5)*2)")           (n 30.0)  // 2+4+6+8+10
check "sum range plus range" (ev "=SUM(A1:A8+B1:B8)")            (n 231.0) // 21 + 210
let sparseBroadcastModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 1.0)
              { Row = 9999; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 19999; Col = 0 }, cell (CellValue.Number 3.0)
              { Row = 0; Col = 1 }, cell (CellValue.Number 10.0)
              { Row = 9999; Col = 1 }, cell (CellValue.Number 20.0)
              { Row = 19999; Col = 1 }, cell (CellValue.Number 30.0) ]
            |> Map.ofList }
let evSparseBroadcast f = evaluateFormula f sparseBroadcastModel None
let arrSparseBroadcast f =
    match parseFormula f with
    | Ok expr -> evalArray expr { Model = sparseBroadcastModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None }
    | Result.Error e -> [| [| CellValue.Error e |] |]
check "sum sparse range plus" (evSparseBroadcast "=SUM(A1:A20000+B1:B20000)") (n 66.0)
check "sum sparse range minus" (evSparseBroadcast "=SUM(B1:B20000-A1:A20000)") (n 54.0)
check "average sparse range plus" (evSparseBroadcast "=AVERAGE(A1:A20000+B1:B20000)") (n 0.0033)
check "count sparse range plus" (evSparseBroadcast "=COUNT(A1:A20000+B1:B20000)") (n 20000.0)
check "min sparse range plus" (evSparseBroadcast "=MIN(A1:A20000+B1:B20000)") (n 0.0)
check "max sparse range plus" (evSparseBroadcast "=MAX(A1:A20000+B1:B20000)") (n 33.0)
check "sum sparse range multiply" (evSparseBroadcast "=SUM(A1:A20000*B1:B20000)") (n 140.0)
check "average sparse range multiply" (evSparseBroadcast "=AVERAGE(A1:A20000*B1:B20000)") (n 0.007)
check "count sparse range multiply" (evSparseBroadcast "=COUNT(A1:A20000*B1:B20000)") (n 20000.0)
check "min sparse range multiply" (evSparseBroadcast "=MIN(A1:A20000*B1:B20000)") (n 0.0)
check "max sparse range multiply" (evSparseBroadcast "=MAX(A1:A20000*B1:B20000)") (n 90.0)
check "sum sparse range plus scalar" (evSparseBroadcast "=SUM(A1:A20000+1)") (n 20006.0)
check "average sparse range plus scalar" (evSparseBroadcast "=AVERAGE(A1:A20000+1)") (n 1.0003)
check "count sparse range plus scalar" (evSparseBroadcast "=COUNT(A1:A20000+1)") (n 20000.0)
check "min sparse range plus scalar" (evSparseBroadcast "=MIN(A1:A20000+1)") (n 1.0)
check "max sparse range plus scalar" (evSparseBroadcast "=MAX(A1:A20000+1)") (n 4.0)
check "sum scalar plus sparse range" (evSparseBroadcast "=SUM(1+A1:A20000)") (n 20006.0)
check "sum sparse range times scalar" (evSparseBroadcast "=SUM(A1:A20000*2)") (n 12.0)
check "sum sparse range divide scalar" (evSparseBroadcast "=SUM(A1:A20000/2)") (n 3.0)
check "average sparse range divide scalar" (evSparseBroadcast "=AVERAGE(A1:A20000/2)") (n 0.00015)
check "count sparse range divide scalar" (evSparseBroadcast "=COUNT(A1:A20000/2)") (n 20000.0)
check "min sparse range divide scalar" (evSparseBroadcast "=MIN(A1:A20000/2)") (n 0.0)
check "max sparse range divide scalar" (evSparseBroadcast "=MAX(A1:A20000/2)") (n 1.5)
check "sum sparse unary minus" (evSparseBroadcast "=SUM(-A1:A20000)") (n -6.0)
check "average sparse unary minus" (evSparseBroadcast "=AVERAGE(-A1:A20000)") (n -0.0003)
check "count sparse unary minus" (evSparseBroadcast "=COUNT(-A1:A20000)") (n 20000.0)
check "min sparse unary minus" (evSparseBroadcast "=MIN(-A1:A20000)") (n -3.0)
check "max sparse unary minus" (evSparseBroadcast "=MAX(-A1:A20000)") (n 0.0)
check "sum sparse unary plus" (evSparseBroadcast "=SUM(+A1:A20000)") (n 6.0)
check "average sparse unary plus" (evSparseBroadcast "=AVERAGE(+A1:A20000)") (n 0.0003)
check "sum sparse unary percent" (evSparseBroadcast "=SUM(A1:A20000%)") (n 0.06)
check "average sparse unary percent" (evSparseBroadcast "=AVERAGE(A1:A20000%)") (n 0.000003)
check "product sparse range plus scalar" (evSparseBroadcast "=PRODUCT(A1:A20000+1)") (n 24.0)
check "product scalar plus sparse range" (evSparseBroadcast "=PRODUCT(1+A1:A20000)") (n 24.0)
check "product sparse range times scalar" (evSparseBroadcast "=PRODUCT(A1:A20000*2)") (n 0.0)
check "product sparse range divide scalar" (evSparseBroadcast "=PRODUCT(A1:A20000/2)") (n 0.0)
check "product sparse range multiply" (evSparseBroadcast "=PRODUCT(A1:A20000*B1:B20000)") (n 0.0)
check "product sparse unary minus" (evSparseBroadcast "=PRODUCT(-A1:A20000)") (n 0.0)
check "product sparse unary plus" (evSparseBroadcast "=PRODUCT(+A1:A20000)") (n 0.0)
check "product sparse unary percent" (evSparseBroadcast "=PRODUCT(A1:A20000%)") (n 0.0)
check "sumsq sparse range plus scalar" (evSparseBroadcast "=SUMSQ(A1:A20000+1)") (n 20026.0)
check "sumsq scalar plus sparse range" (evSparseBroadcast "=SUMSQ(1+A1:A20000)") (n 20026.0)
check "sumsq sparse range times scalar" (evSparseBroadcast "=SUMSQ(A1:A20000*2)") (n 56.0)
check "sumsq sparse range divide scalar" (evSparseBroadcast "=SUMSQ(A1:A20000/2)") (n 3.5)
check "sumsq sparse range multiply" (evSparseBroadcast "=SUMSQ(A1:A20000*B1:B20000)") (n 9800.0)
check "sumsq sparse unary minus" (evSparseBroadcast "=SUMSQ(-A1:A20000)") (n 14.0)
check "sumsq sparse unary plus" (evSparseBroadcast "=SUMSQ(+A1:A20000)") (n 14.0)
check "sumsq sparse unary percent" (evSparseBroadcast "=SUMSQ(A1:A20000%)") (n 0.0014)
check "countblank sparse range plus scalar" (evSparseBroadcast "=COUNTBLANK(A1:A20000+1)") (n 0.0)
check "countblank sparse range times scalar" (evSparseBroadcast "=COUNTBLANK(A1:A20000*2)") (n 0.0)
check "countblank sparse range divide scalar" (evSparseBroadcast "=COUNTBLANK(A1:A20000/2)") (n 0.0)
check "countblank sparse range multiply" (evSparseBroadcast "=COUNTBLANK(A1:A20000*B1:B20000)") (n 0.0)
check "countblank sparse unary minus" (evSparseBroadcast "=COUNTBLANK(-A1:A20000)") (n 0.0)
check "countblank sparse unary percent" (evSparseBroadcast "=COUNTBLANK(A1:A20000%)") (n 0.0)
check "counta sparse range plus scalar" (evSparseBroadcast "=COUNTA(A1:A20000+1)") (n 20000.0)
check "counta sparse range times scalar" (evSparseBroadcast "=COUNTA(A1:A20000*2)") (n 20000.0)
check "counta sparse range divide scalar" (evSparseBroadcast "=COUNTA(A1:A20000/2)") (n 20000.0)
check "counta sparse range multiply" (evSparseBroadcast "=COUNTA(A1:A20000*B1:B20000)") (n 20000.0)
check "counta sparse unary minus" (evSparseBroadcast "=COUNTA(-A1:A20000)") (n 20000.0)
check "counta sparse unary percent" (evSparseBroadcast "=COUNTA(A1:A20000%)") (n 20000.0)
check "averagea sparse range plus scalar" (evSparseBroadcast "=AVERAGEA(A1:A20000+1)") (n 1.0003)
check "mina sparse range plus scalar" (evSparseBroadcast "=MINA(A1:A20000+1)") (n 1.0)
check "maxa sparse range plus scalar" (evSparseBroadcast "=MAXA(A1:A20000+1)") (n 4.0)
check "averagea sparse range multiply" (evSparseBroadcast "=AVERAGEA(A1:A20000*B1:B20000)") (n 0.007)
check "mina sparse unary minus" (evSparseBroadcast "=MINA(-A1:A20000)") (n -3.0)
check "maxa sparse unary minus" (evSparseBroadcast "=MAXA(-A1:A20000)") (n 0.0)
approx "varp sparse range plus scalar" (evSparseBroadcast "=VAR.P(A1:A20000+1)") 0.00069991000000024783
approx "vars sparse range plus scalar" (evSparseBroadcast "=VAR.S(A1:A20000+1)") 0.00069994499725011038
approx "stdevp sparse range plus scalar" (evSparseBroadcast "=STDEV.P(A1:A20000+1)") 0.026455812215848671
approx "varp sparse range multiply" (evSparseBroadcast "=VAR.P(A1:A20000*B1:B20000)") 0.48995099999999997
approx "stdevp sparse unary minus" (evSparseBroadcast "=STDEV.P(-A1:A20000)") 0.026455812215843987
approx "devsq sparse range plus scalar" (evSparseBroadcast "=DEVSQ(A1:A20000+1)") 13.9982
approx "avedev sparse range plus scalar" (evSparseBroadcast "=AVEDEV(A1:A20000+1)") 0.00059991
approx "geomean sparse range plus scalar" (evSparseBroadcast "=GEOMEAN(A1:A20000+1)") 1.000158915
approx "harmean sparse range plus scalar" (evSparseBroadcast "=HARMEAN(A1:A20000+1)") 1.000095843
approx "devsq sparse range multiply" (evSparseBroadcast "=DEVSQ(A1:A20000*B1:B20000)") 9799.02
approx "avedev sparse unary minus" (evSparseBroadcast "=AVEDEV(-A1:A20000)") 0.00059991
check "geomean sparse range invalid zero" (evSparseBroadcast "=GEOMEAN(A1:A20000*2)") (CellValue.Error "NUM!")
check "harmean sparse range invalid zero" (evSparseBroadcast "=HARMEAN(A1:A20000*2)") (CellValue.Error "NUM!")
check "median dense range plus scalar" (ev "=MEDIAN(A1:A8+1)") (n 3.5)
check "large dense range plus scalar" (ev "=LARGE(A1:A8+1,1)") (n 6.0)
check "small dense unary minus" (ev "=SMALL(-A1:A8,1)") (n -5.0)
check "percentile dense range plus scalar" (ev "=PERCENTILE(A1:A8+1,0.5)") (n 3.5)
check "quartile dense range times scalar" (ev "=QUARTILE(A1:A8*2,2)") (n 5.0)
check "mode dense range plus scalar" (ev "=MODE(A1:A8+1)") (n 4.0)
check "rank dense range plus scalar" (ev "=RANK(4,A1:A8+1)") (n 3.0)
check "sum dense range plus scalar" (ev "=SUM(A1:A8+1)") (n 29.0)
check "count dense range plus scalar" (ev "=COUNT(A1:A8+1)") (n 8.0)
check "sumsq dense range plus scalar" (ev "=SUMSQ(A1:A8+1)") (n 119.0)
approx "varp dense range plus scalar" (ev "=VAR.P(A1:A8+1)") 1.734375
check "sum dense unary minus" (ev "=SUM(-A1:A8)") (n -21.0)
approx "average dense unary percent" (ev "=AVERAGE(A1:A8%)") 0.02625
check "sum dense range plus range" (ev "=SUM(A1:A8+B1:B8)") (n 231.0)
approx "varp dense range plus range" (ev "=VAR.P(A1:A8+B1:B8)") 209.859375
check "median sparse range plus scalar" (evSparseBroadcast "=MEDIAN(A1:A20000+1)") (n 1.0)
check "large sparse range plus scalar" (evSparseBroadcast "=LARGE(A1:A20000+1,1)") (n 4.0)
check "small sparse range plus scalar" (evSparseBroadcast "=SMALL(A1:A20000+1,1)") (n 1.0)
check "percentile sparse range plus scalar" (evSparseBroadcast "=PERCENTILE(A1:A20000+1,0.5)") (n 1.0)
check "quartile sparse range plus scalar" (evSparseBroadcast "=QUARTILE(A1:A20000+1,4)") (n 4.0)
check "median sparse range product" (evSparseBroadcast "=MEDIAN(A1:A20000*B1:B20000)") (n 0.0)
check "large sparse unary minus" (evSparseBroadcast "=LARGE(-A1:A20000,1)") (n 0.0)
check "small sparse unary minus" (evSparseBroadcast "=SMALL(-A1:A20000,1)") (n -3.0)
check "mode sparse range plus scalar" (evSparseBroadcast "=MODE(A1:A20000+1)") (n 1.0)
check "mode sparse range product" (evSparseBroadcast "=MODE(A1:A20000*B1:B20000)") (n 0.0)
check "rank sparse range plus scalar" (evSparseBroadcast "=RANK(4,A1:A20000+1)") (n 1.0)
check "rank sparse range product" (evSparseBroadcast "=RANK(90,A1:A20000*B1:B20000)") (n 1.0)
let sparseBroadcastErrorModel =
    { sparseBroadcastModel with
        Cells = sparseBroadcastModel.Cells |> Map.add { Row = 100; Col = 1 } (cell (CellValue.Error "N/A")) }
check "sum sparse range binop error" (evaluateFormula "=SUM(A1:A20000+B1:B20000)" sparseBroadcastErrorModel None) (CellValue.Error "N/A")
check "average sparse range binop error" (evaluateFormula "=AVERAGE(A1:A20000+B1:B20000)" sparseBroadcastErrorModel None) (CellValue.Error "N/A")
let sparseBroadcastScalarErrorModel =
    { sparseBroadcastModel with
        Cells = sparseBroadcastModel.Cells |> Map.add { Row = 100; Col = 0 } (cell (CellValue.Error "N/A")) }
check "sum sparse range scalar error" (evaluateFormula "=SUM(A1:A20000+1)" sparseBroadcastScalarErrorModel None) (CellValue.Error "N/A")

// ── scalar context collapses to top-left ──────────────────────────────
check "scalar sequence"  (ev "=SEQUENCE(5)")     (n 1.0)
check "scalar randarray fixed" (ev "=RANDARRAY(20000,2,7,7,TRUE)") (n 7.0)
check "scalar randarray invalid bounds" (ev "=RANDARRAY(20000,1,10,1,TRUE)") (CellValue.Error "VALUE!")
check "scalar sort"      (ev "=SORT(A1:A8)")     (n 1.0)   // smallest, top-left after sort
check "scalar sortby desc" (ev "=SORTBY(A1:A8,B1:B8,-1)") (n 5.0)
check "scalar unique+1"  (ev "=UNIQUE(A1:A8)+10") (n 13.0) // top-left 3, +10
check "scalar unique rows" (ev "=UNIQUE(A1:B8)") (n 3.0)
check "scalar unique by_col" (ev "=UNIQUE(A1:B1,TRUE)") (n 3.0)
check "scalar byrow" (ev "=BYROW(A1:B8,LAMBDA(r,SUM(r)))") (n 33.0)
check "scalar bycol" (ev "=BYCOL(A1:B8,LAMBDA(c,SUM(c)))") (n 21.0)
check "scalar sparse transpose" (match parseFormula "=TRANSPOSE(A10000:B20000)" with Ok e -> eval e { Model = sparseUniqueModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None } | Result.Error e -> CellValue.Error e) (n 2.0)
check "scalar sparse sort asc" (match parseFormula "=SORT(A1:A20000)" with Ok e -> eval e { Model = sparseUniqueModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None } | Result.Error e -> CellValue.Error e) CellValue.Empty
check "scalar sparse sort desc" (match parseFormula "=SORT(A1:A20000,1,-1)" with Ok e -> eval e { Model = sparseUniqueModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None } | Result.Error e -> CellValue.Error e) (n 3.0)
check "scalar sparse unique rows" (match parseFormula "=UNIQUE(A1:B20000)" with Ok e -> eval e { Model = sparseUniqueModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None } | Result.Error e -> CellValue.Error e) CellValue.Empty
check "scalar sparse byrow" (evSparseByRow "=BYROW(A1:B20000,LAMBDA(r,SUM(r)))") (n 2.0)
check "scalar sparse bycol" (evSparseByRow "=BYCOL(A1:B10000,LAMBDA(c,SUM(c)))") (n 3.0)
check "scalar sparse sortby asc" (evSparseBroadcast "=SORTBY(A1:A20000,B1:B20000,1)") CellValue.Empty
check "scalar sparse sortby desc" (evSparseBroadcast "=SORTBY(A1:A20000,B1:B20000,-1)") (n 3.0)
check "sum sparse sort" (evSparseBroadcast "=SUM(SORT(A1:A20000))") (n 6.0)
check "count sparse sort" (evSparseBroadcast "=COUNT(SORT(A1:A20000))") (n 3.0)
check "average sparse sort" (evSparseBroadcast "=AVERAGE(SORT(A1:A20000))") (n 2.0)
check "min sparse sort" (evSparseBroadcast "=MIN(SORT(A1:A20000))") (n 1.0)
check "max sparse sort" (evSparseBroadcast "=MAX(SORT(A1:A20000))") (n 3.0)
check "product sparse sort" (evSparseBroadcast "=PRODUCT(SORT(A1:A20000))") (n 6.0)
check "sum sparse sortby" (evSparseBroadcast "=SUM(SORTBY(A1:A20000,B1:B20000,-1))") (n 6.0)
check "count sparse sortby" (evSparseBroadcast "=COUNT(SORTBY(A1:A20000,B1:B20000,-1))") (n 3.0)
check "average sparse sortby" (evSparseBroadcast "=AVERAGE(SORTBY(A1:A20000,B1:B20000,-1))") (n 2.0)
check "min sparse sortby" (evSparseBroadcast "=MIN(SORTBY(A1:A20000,B1:B20000,-1))") (n 1.0)
check "max sparse sortby" (evSparseBroadcast "=MAX(SORTBY(A1:A20000,B1:B20000,-1))") (n 3.0)
check "product sparse sortby" (evSparseBroadcast "=PRODUCT(SORTBY(A1:A20000,B1:B20000,-1))") (n 6.0)
check "sum sparse sortby source error" (evaluateFormula "=SUM(SORTBY(A1:A20000,B1:B20000,-1))" sparseBroadcastScalarErrorModel None) (CellValue.Error "N/A")
check "sum sparse filter aggregate" (evSparseBroadcast "=SUM(FILTER(A1:A20000,B1:B20000>0))") (n 6.0)
check "count sparse filter aggregate" (evSparseBroadcast "=COUNT(FILTER(A1:A20000,B1:B20000>0))") (n 3.0)
check "average sparse filter aggregate" (evSparseBroadcast "=AVERAGE(FILTER(A1:A20000,B1:B20000>0))") (n 2.0)
check "min sparse filter aggregate" (evSparseBroadcast "=MIN(FILTER(A1:A20000,B1:B20000>0))") (n 1.0)
check "max sparse filter aggregate" (evSparseBroadcast "=MAX(FILTER(A1:A20000,B1:B20000>0))") (n 3.0)
check "product sparse filter aggregate" (evSparseBroadcast "=PRODUCT(FILTER(A1:A20000,B1:B20000>0))") (n 6.0)
check "sum sparse filter empty" (evSparseBroadcast "=SUM(FILTER(A1:A20000,B1:B20000>1000))") (CellValue.Error "CALC!")
check "sum sparse filter if-empty" (evSparseBroadcast "=SUM(FILTER(A1:A20000,B1:B20000>1000,0))") (n 0.0)
let sparseFilterExpr = arrSparseBroadcast "=FILTER(A1:A20000,B1:B20000>0)"
check "filter sparse expr shape" (shape sparseFilterExpr) (3, 1)
check "filter sparse expr first" sparseFilterExpr.[0].[0] (n 1.0)
check "filter sparse expr tail" sparseFilterExpr.[2].[0] (n 3.0)
check "sum sparse filter map" (evSparseBroadcast "=SUM(MAP(FILTER(A1:A20000,B1:B20000>0),LAMBDA(x,x+1)))") (n 9.0)
check "sum sparse transpose aggregate" (evSparseBroadcast "=SUM(TRANSPOSE(A1:A20000))") (n 6.0)
check "count sparse transpose aggregate" (evSparseBroadcast "=COUNT(TRANSPOSE(A1:A20000))") (n 3.0)
check "average sparse transpose aggregate" (evSparseBroadcast "=AVERAGE(TRANSPOSE(A1:A20000))") (n 2.0)
check "min sparse transpose aggregate" (evSparseBroadcast "=MIN(TRANSPOSE(A1:A20000))") (n 1.0)
check "max sparse transpose aggregate" (evSparseBroadcast "=MAX(TRANSPOSE(A1:A20000))") (n 3.0)
check "product sparse transpose aggregate" (evSparseBroadcast "=PRODUCT(TRANSPOSE(A1:A20000))") (n 6.0)
let sparseZeroModel =
    { createEmpty 20010 2 with
        Cells = [ { Row = 9999; Col = 0 }, cell (CellValue.Number 0.0) ] |> Map.ofList }
check "scalar sparse sort zero tie" (match parseFormula "=SORT(A1:A20000)" with Ok e -> eval e { Model = sparseZeroModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None } | Result.Error e -> CellValue.Error e) CellValue.Empty
let sparseFilterScalarModel =
    { createEmpty 20010 5 with
        Cells =
            [ { Row = 9999; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 9999; Col = 1 }, cell (CellValue.Number 20.0)
              { Row = 9999; Col = 2 }, cell (CellValue.Boolean true)
              { Row = 9999; Col = 3 }, cell (CellValue.Number 1.0) ]
            |> Map.ofList }
check "scalar sparse filter hit" (match parseFormula "=FILTER(A1:B20000,C1:C20000)" with Ok e -> eval e { Model = sparseFilterScalarModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None } | Result.Error e -> CellValue.Error e) (n 2.0)
check "scalar sparse filter expr" (match parseFormula "=FILTER(A1:B20000,D1:D20000>0)" with Ok e -> eval e { Model = sparseFilterScalarModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None } | Result.Error e -> CellValue.Error e) (n 2.0)
check "scalar sparse filter empty" (match parseFormula "=FILTER(A1:B20000,E1:E20000,\"none\")" with Ok e -> eval e { Model = sparseFilterScalarModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None } | Result.Error e -> CellValue.Error e) (CellValue.String "none")
let denseFilterScalarModel =
    { createEmpty 2010 4 with
        Cells =
            [ for r in 0 .. 1999 do
                  { Row = r; Col = 0 }, cell (CellValue.Number (float r))
                  { Row = r; Col = 2 }, cell (CellValue.Boolean true)
                  { Row = r; Col = 3 }, cell (CellValue.Number 1.0) ]
            |> Map.ofList }
check "scalar dense filter hit" (match parseFormula "=FILTER(A1:B2000,C1:C2000)" with Ok e -> eval e { Model = denseFilterScalarModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None } | Result.Error e -> CellValue.Error e) (n 0.0)
check "scalar dense filter expr" (match parseFormula "=FILTER(A1:B2000,D1:D2000>0)" with Ok e -> eval e { Model = denseFilterScalarModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None } | Result.Error e -> CellValue.Error e) (n 0.0)

printfn ""
printfn "Stage 8 dynamic-array tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
