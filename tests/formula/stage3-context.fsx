#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Formula
open RedBat.Spreadsheet.Workbook

let cell v = { Value = v; Format = defaultFormat; Comment = None }

// Sheet1 (active): B1=10, B2=20, B3=30
let sheet1Model =
    { createEmpty 100 26 with
        Cells =
            [ { Row=0; Col=1 }, cell (CellValue.Number 10.0)
              { Row=1; Col=1 }, cell (CellValue.Number 20.0)
              { Row=2; Col=1 }, cell (CellValue.Number 30.0) ]
            |> Map.ofList }

// Sheet2: A1=99, A2=1, A3=2, A4=3
let sheet2Model =
    { createEmpty 100 26 with
        Cells =
            [ { Row=0; Col=0 }, cell (CellValue.Number 99.0)
              { Row=1; Col=0 }, cell (CellValue.Number 1.0)
              { Row=2; Col=0 }, cell (CellValue.Number 2.0)
              { Row=3; Col=0 }, cell (CellValue.Number 3.0) ]
            |> Map.ofList }

let workbook =
    { Sheets =
        Map.ofList
            [ "sheet1", { Id = "sheet1"; Name = "Sheet1"; Model = sheet1Model; Visible = true }
              "sheet2", { Id = "sheet2"; Name = "Sheet2"; Model = sheet2Model; Visible = true } ]
      SheetOrder = [ "sheet1"; "sheet2" ]
      ActiveSheet = "sheet1"
      Tables = Map.empty
      Names =
        namesOfTriples
            [ None,            "MyCell",    "Sheet2!$A$1"        // single cross-sheet cell
              None,            "MyRange",   "Sheet2!$A$2:$A$4"   // cross-sheet range
              None,            "LocalCell", "B2"                 // unqualified → active model
              None,            "Alias",     "MyCell"             // name → name chain
              None,            "CycA",      "CycB"               // cyclic pair
              None,            "CycB",      "CycA"
              None,            "Shared",    "B1"                 // workbook-global Shared → 10
              Some "sheet1",   "Shared",    "B3" ] }             // sheet1-local Shared → 30 (shadows)

let model = sheet1Model
let wb = Some workbook

let mutable pass = 0
let mutable fail = 0
let check label formula (expected: CellValue) =
    let got = evaluateFormula formula model wb
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-22s  %-30s  expected %A  got %A" label formula expected got

let n x = CellValue.Number x
let e x = CellValue.Error x

// ----- cross-sheet references -----
check "xsheet-cell"   "=Sheet2!A1"               (n 99.0)
check "xsheet-add"    "=Sheet2!A1+B1"            (n 109.0)        // 99 + Sheet1!B1
check "xsheet-range"  "=SUM(Sheet2!A2:A4)"       (n 6.0)
check "xsheet-range-scalar" "=Sheet2!A2:A4"       (n 1.0)
check "xsheet-avg"    "=AVERAGE(Sheet2!A2:A4)"   (n 2.0)
check "xsheet-bad"    "=NoSuchSheet!A1"          (e "REF!")
check "xsheet-ci"     "=sheet2!A1"               (n 99.0)         // sheet name match is case-insensitive
check "samesheet"     "=B2"                      (n 20.0)         // unqualified still hits active model

// ----- volatile functions -----
// RAND ∈ [0,1)
match evaluateFormula "=RAND()" model wb with
| CellValue.Number r when r >= 0.0 && r < 1.0 -> pass <- pass + 1
| got -> fail <- fail + 1; printfn "FAIL rand  expected [0,1)  got %A" got

// RANDBETWEEN(5,5) is deterministic
check "randbetween-fixed" "=RANDBETWEEN(5,5)"    (n 5.0)
match evaluateFormula "=RANDBETWEEN(1,6)" model wb with
| CellValue.Number r when r >= 1.0 && r <= 6.0 && r = floor r -> pass <- pass + 1
| got -> fail <- fail + 1; printfn "FAIL randbetween-range  got %A" got

check "offset"        "=OFFSET(B1,1,0)"          (n 20.0)         // B1 + 1 row = B2
check "offset-2"      "=OFFSET(B1,2,0)"          (n 30.0)
check "offset-neg"    "=OFFSET(B3,-2,0)"         (n 10.0)
check "offset-oob"    "=OFFSET(B1,-1,0)"         (e "REF!")
check "indirect"      "=INDIRECT(\"B2\")"        (n 20.0)
check "indirect-xs"   "=INDIRECT(\"Sheet2!A1\")" (n 99.0)
check "indirect-sum"  "=SUM(INDIRECT(\"B1\"),INDIRECT(\"B2\"))" (n 30.0)

// ----- defined names -----
check "name-cell"     "=MyCell"                  (n 99.0)
check "name-ci"       "=mycell"                  (n 99.0)         // names are case-insensitive
check "name-local"    "=LocalCell"               (n 20.0)         // unqualified ref → active model
check "name-in-expr"  "=MyCell+LocalCell"        (n 119.0)
check "name-range"    "=SUM(MyRange)"            (n 6.0)          // name standing in for a range
check "name-range-scalar" "=MyRange"             (n 1.0)
check "name-range-avg" "=AVERAGE(MyRange)"       (n 2.0)
check "name-chain"    "=Alias"                   (n 99.0)         // Alias → MyCell → Sheet2!A1
check "name-cyclic"   "=CycA"                    (e "NAME?")      // cycle guard
check "name-unknown"  "=NoSuchName"              (e "NAME?")

// sheet-scoped names: evaluateFormula resolves at the active sheet
// (sheet1), where the sheet-local "Shared" shadows the global one.
check "name-scoped-local"  "=Shared"             (n 30.0)         // sheet1-local Shared → B3
// resolving with a different (or no) sheet scope falls back to the global.
let sharedExpr = match parseFormula "=Shared" with Ok x -> x | _ -> Number 0.0
match eval sharedExpr { Model = model; Workbook = wb; Sheet = Some "sheet2"; CriteriaCache = None; CurrentCell = None } with
| CellValue.Number 10.0 -> pass <- pass + 1
| got -> fail <- fail + 1; printfn "FAIL name-scoped-global  expected 10  got %A" got

// containsVolatile detection
let volatileExpr =
    match parseFormula "=RAND()+1" with Ok x -> x | _ -> Number 0.0
let volatileArrayExpr =
    match parseFormula "=SUM(RANDARRAY(2,2))" with Ok x -> x | _ -> Number 0.0
let plainExpr =
    match parseFormula "=B1+B2" with Ok x -> x | _ -> Number 0.0
if containsVolatile volatileExpr then pass <- pass + 1
else (fail <- fail + 1; printfn "FAIL containsVolatile(RAND) should be true")
if containsVolatile volatileArrayExpr then pass <- pass + 1
else (fail <- fail + 1; printfn "FAIL containsVolatile(RANDARRAY) should be true")
if not (containsVolatile plainExpr) then pass <- pass + 1
else (fail <- fail + 1; printfn "FAIL containsVolatile(B1+B2) should be false")

printfn ""
printfn "Stage 3 formula tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
