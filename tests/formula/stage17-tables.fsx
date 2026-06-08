#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// Structured references (rich grammar) + SUBTOTAL / AGGREGATE.

open RedBat.Spreadsheet.Model
module F = RedBat.Spreadsheet.Formula
module WB = RedBat.Spreadsheet.Workbook

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let S = CellValue.String
let N = CellValue.Number

// Sheet layout (0-indexed):
//   row0 header:  Region | Amount | Sales Amount
//   row1..4 data: East/10/100, West/20/200, East/30/300, North/40/400
//   row5 totals:  "Total"| =SUBTOTAL(9,B2:B5)->100 | (blank)
let baseCells =
    [ { Row=0; Col=0 }, cell (S "Region");   { Row=0; Col=1 }, cell (S "Amount");   { Row=0; Col=2 }, cell (S "Sales Amount")
      { Row=1; Col=0 }, cell (S "East");      { Row=1; Col=1 }, cell (N 10.);        { Row=1; Col=2 }, cell (N 100.)
      { Row=2; Col=0 }, cell (S "West");      { Row=2; Col=1 }, cell (N 20.);        { Row=2; Col=2 }, cell (N 200.)
      { Row=3; Col=0 }, cell (S "East");      { Row=3; Col=1 }, cell (N 30.);        { Row=3; Col=2 }, cell (N 300.)
      { Row=4; Col=0 }, cell (S "North");     { Row=4; Col=1 }, cell (N 40.);        { Row=4; Col=2 }, cell (N 400.)
      { Row=5; Col=0 }, cell (S "Total")
      // totals cell is itself a SUBTOTAL formula (cached 100) — must be skipped
      // by an outer SUBTOTAL ranging over it.
      { Row=5; Col=1 }, cell (CellValue.Formula ("SUBTOTAL(9,B2:B5)", N 100.)) ]
    |> Map.ofList

let model = { createEmpty 20 8 with Cells = baseCells }

let table : WB.Table =
    { Name = "Sales"; SheetId = "sheet1"
      StartRow = 0; StartCol = 0; EndRow = 5; EndCol = 2
      Columns = [ "Region"; "Amount"; "Sales Amount" ]
      HasHeaders = true; HasTotals = true }

let wb : WB.Workbook =
    { Sheets = Map.ofList [ "sheet1", ({ Id="sheet1"; Name="Sheet1"; Model=model; Visible=true }: WB.Sheet) ]
      SheetOrder = [ "sheet1" ]; ActiveSheet = "sheet1"
      Tables = Map.ofList [ "Sales", table ]; Names = Map.empty }

let ctxAt (m: SpreadsheetModel) (cur: CellAddress option) : F.EvalContext =
    { Model = m; Workbook = Some wb; Sheet = Some "sheet1"; CurrentCell = cur }

let evalIn (m: SpreadsheetModel) (cur: CellAddress option) (s: string) : CellValue =
    match F.parseFormula s with
    | Ok e -> F.eval e (ctxAt m cur)
    | Result.Error msg -> CellValue.Error msg

let evalS s = evalIn model None s
let num = function CellValue.Number n -> n | v -> nan
let str = function CellValue.String s -> s | CellValue.Number n -> string n | v -> sprintf "%A" v

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-30s expected %A got %A" label expected got

// --- structured-ref grammar ---
check "Table[Col] scalar (first)" (num (evalS "Sales[Amount]")) 10.0
check "SUM Table[Col]"            (num (evalS "SUM(Sales[Amount])")) 100.0
check "multi-word column"         (num (evalS "SUM(Sales[Sales Amount])")) 1000.0
check "#Data only"                (num (evalS "SUM(Sales[[#Data],[Amount]])")) 100.0
check "#Headers scalar"           (str (evalS "Sales[[#Headers],[Amount]]")) "Amount"
check "#Totals scalar"            (num (evalS "Sales[[#Totals],[Amount]]")) 100.0
check "column span sum"           (num (evalS "SUM(Sales[[Amount]:[Sales Amount]])")) 1100.0
check "#All sum (header text ignored)" (num (evalS "SUM(Sales[[#All],[Amount]])")) 200.0  // data 100 + totals 100

// --- this-row [@Col] (qualified + unqualified), driven by CurrentCell ---
check "qualified @ row2"  (num (evalIn model (Some { Row=2; Col=0 }) "Sales[@Amount]")) 20.0
check "unqualified @ row3" (num (evalIn model (Some { Row=3; Col=0 }) "[@Amount]")) 30.0
check "@ off-table -> #VALUE!" (evalIn model (Some { Row=0; Col=0 }) "Sales[@Amount]") (CellValue.Error "VALUE!")

// --- SUBTOTAL ---
check "SUBTOTAL 9 sum"        (num (evalS "SUBTOTAL(9,B2:B5)")) 100.0
check "SUBTOTAL 1 average"    (num (evalS "SUBTOTAL(1,B2:B5)")) 25.0
check "SUBTOTAL 2 count"      (num (evalS "SUBTOTAL(2,B2:B5)")) 4.0
check "SUBTOTAL 4 max"        (num (evalS "SUBTOTAL(4,B2:B5)")) 40.0
// nested SUBTOTAL in B6 is skipped → B2:B6 still sums to 100, not 200
check "SUBTOTAL skips nested" (num (evalS "SUBTOTAL(9,B2:B6)")) 100.0
check "plain SUM counts nested" (num (evalS "SUM(B2:B6)")) 200.0

// manually-hidden row (height 0): 109 excludes it, 9 includes it
let hidden = { model with RowHeights = Map.ofList [ 3, 0.0 ] }   // row idx 3 = value 30
check "SUBTOTAL 109 skips hidden" (num (evalIn hidden None "SUBTOTAL(109,B2:B5)")) 70.0
check "SUBTOTAL 9 keeps hidden"   (num (evalIn hidden None "SUBTOTAL(9,B2:B5)")) 100.0

// filter-hidden rows excluded by both 9 and 109
let filtered = applyFilter 0 (Set.ofList [ "East" ]) model   // keeps rows 1,3 (East)
check "SUBTOTAL 9 honors filter" (num (evalIn filtered None "SUBTOTAL(9,B2:B5)")) 40.0  // 10 + 30

// --- AGGREGATE ---
let withErr = { model with Cells = Map.add { Row=3; Col=1 } (cell (CellValue.Error "DIV/0!")) model.Cells }
check "AGGREGATE 9 opt6 ignores errors" (num (evalIn withErr None "AGGREGATE(9,6,B2:B5)")) 70.0  // 10+20+40
check "AGGREGATE 9 opt4 sees error"     (evalIn withErr None "AGGREGATE(9,4,B2:B5)") (CellValue.Error "DIV/0!")
check "AGGREGATE 1 average"             (num (evalS "AGGREGATE(1,0,B2:B5)")) 25.0

// --- bridge helper: makeTableFromRegion (1-indexed XLSX region → 0-indexed) ---
// Columns omitted → read from the header row of the model.
let bridged = WB.makeTableFromRegion "sheet1" "Sales" 1 1 6 3 1 0 [] model
check "bridge name"       bridged.Name "Sales"
check "bridge 0-indexed start" (bridged.StartRow, bridged.StartCol) (0, 0)
check "bridge 0-indexed end"   (bridged.EndRow, bridged.EndCol) (5, 2)
check "bridge hasHeaders"  bridged.HasHeaders true
check "bridge cols from header" bridged.Columns [ "Region"; "Amount"; "Sales Amount" ]
// Explicit column names win over header cells.
let bridged2 = WB.makeTableFromRegion "sheet1" "T" 1 1 6 3 1 1 [ "A"; "B"; "C" ] model
check "bridge explicit cols" bridged2.Columns [ "A"; "B"; "C" ]
check "bridge hasTotals"   bridged2.HasTotals true

// --- createTable: register a real table + SUBTOTAL totals row ---
// Fresh data block: header A1:C1, data rows 2-4.
let dataModel =
    { createEmpty 20 8 with
        Cells =
            [ { Row=0; Col=0 }, cell (S "Region"); { Row=0; Col=1 }, cell (S "Q1"); { Row=0; Col=2 }, cell (S "Q2")
              { Row=1; Col=0 }, cell (S "East");   { Row=1; Col=1 }, cell (N 100.); { Row=1; Col=2 }, cell (N 200.)
              { Row=2; Col=0 }, cell (S "West");   { Row=2; Col=1 }, cell (N 150.); { Row=2; Col=2 }, cell (N 250.) ]
            |> Map.ofList }
let created = createTable "Sales" 0 0 2 2 true dataModel
let ct = List.exactlyOne created.Tables
check "createTable name"      ct.Name "Sales"
check "createTable 1-indexed"  (ct.StartRow, ct.StartCol, ct.EndCol) (1, 1, 3)
check "createTable totals grows range" ct.EndRow 4   // header(1)+data(2,3)+totals(4)
check "createTable hasTotals"  ct.HasTotals true
check "createTable autofilter" ct.AutoFilter (Some (1, 3))
check "createTable style"      (ct.Style |> Option.map (fun s -> s.Name)) (Some "TableStyleMedium2")
check "createTable totals label" (str (getCellValue { Row=3; Col=0 } created)) "Total"
// totals cells hold SUBTOTAL formulas
let totalsFormula =
    match getCellValue { Row=3; Col=1 } created with CellValue.Formula (t, _) -> t | _ -> ""
check "createTable totals formula" totalsFormula "SUBTOTAL(109,B2:B3)"

// --- table-aware filtering (filter-button scope) ---
check "tableAt header cell" (tableAt 0 0 created |> Option.isSome) true
check "tableAt data cell"   (tableAt 2 1 created |> Option.isSome) true
check "tableAt outside"     (tableAt 12 12 created |> Option.isSome) false
check "tableColUnique col0" (tableColumnUniqueValues ct 0 created) [ "East"; "West" ]
check "tableColUnique excludes totals" (tableColumnUniqueValues ct 1 created) [ "100"; "150" ]

printfn ""
printfn "Stage 17 table tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
