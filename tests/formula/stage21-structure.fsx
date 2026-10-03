#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook
open Rbt.Spreadsheet.Formula

// Structure edits: paste/fill reference shifting, the Expr → text printer,
// row/column insert + delete, sort, and workbook-wide recalculation.

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)
let eq label (got: 'a) (expected: 'a) =
    ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let num x = cell (CellValue.Number x)
let str x = cell (CellValue.String x)
let fml f = cell (CellValue.Formula (f, CellValue.Empty))
let at r c = { Row = r; Col = c }
let modelOf cells = { createEmpty 100 26 with Cells = Map.ofList cells }
let formulaAt r c (m: SpreadsheetModel) =
    match getCellValue (at r c) m with
    | CellValue.Formula (f, _) -> f
    | other -> sprintf "<%A>" other
let valueAt r c (m: SpreadsheetModel) = getCellValue (at r c) m |> unwrapFormulaValue

// --- Paste shifts references through the parser, not by scanning text -------
// This block runs before anything else touches the Formula module, so it also
// proves Model.pasteClipboard finds the rewriter on a cold start.
let pasteOf formula =
    let source = modelOf [ at 0 2, fml formula ]
    let clip = serializeSelection (SingleCell (0, 2)) source false
    pasteClipboard (at 1 3) clip source |> formulaAt 1 3

eq "paste-cold-start"     (pasteOf "=LOG10(A1)")             "=LOG10(B2)"
eq "paste-function-name"  (pasteOf "=ATAN2(A1,B1)")          "=ATAN2(B2,C2)"
eq "paste-sheet-name"     (pasteOf "=SUM(Sheet1!A1:A2)")     "=SUM(Sheet1!B2:B3)"
eq "paste-quoted-sheet"   (pasteOf "=SUM('My Sheet'!A1:A2)") "=SUM('My Sheet'!B2:B3)"
eq "paste-string-literal" (pasteOf "=IF(A1=\"B2\",1,2)")     "=IF(B2=\"B2\",1,2)"
eq "paste-col-locked"     (pasteOf "=$A1+1")                 "=$A2+1"
eq "paste-row-locked"     (pasteOf "=A$1+1")                 "=B$1+1"
eq "paste-fully-locked"   (pasteOf "=$A$1+1")                "=$A$1+1"
eq "paste-no-equals"      (pasteOf "A1*2")                   "B2*2"
eq "paste-name-kept"      (pasteOf "=Rate*A1")               "=Rate*B2"
eq "paste-unparsable"     (pasteOf "=SUM(A1")                "=SUM(A1"

// Pasting up/left past the sheet edge gives #REF!, as Excel does.
let pasteUp =
    let source = modelOf [ at 1 1, fml "=A1+B2" ]
    let clip = serializeSelection (SingleCell (1, 1)) source false
    pasteClipboard (at 0 0) clip source |> formulaAt 0 0
eq "paste-off-sheet" pasteUp "=#REF!+A1"

// Cut-then-paste moves the formula as written.
let cutPaste =
    let source = modelOf [ at 0 2, fml "=A1+$B$2" ]
    let clip = serializeSelection (SingleCell (0, 2)) source true
    pasteClipboard (at 5 5) clip source |> formulaAt 5 5
eq "cut-paste-unshifted" cutPaste "=A1+$B$2"

// A copied block keeps each formula relative to its own cell.
let blockPaste =
    let source = modelOf [ at 0 0, num 1.0; at 1 0, num 2.0; at 0 1, fml "=A1*10"; at 1 1, fml "=A2*10" ]
    let clip = serializeSelection (Range (0, 1, 1, 1)) source false
    pasteClipboard (at 0 2) clip source |> recalcAndSpill None None
eq "block-paste-formula" (formulaAt 1 2 blockPaste) "=B2*10"
eq "block-paste-value"   (valueAt 1 2 blockPaste) (CellValue.Number 200.0)

eq "fill-shift" (shiftFormulaText 2 0 "=SUM(A1:A3)/COUNT($A$1:A3)") "=SUM(A3:A5)/COUNT($A$1:A5)"

// --- Printer round trip -----------------------------------------------------
let roundTrip (formula: string) =
    match parseFormula formula with
    | Result.Error e -> ok ("print " + formula) false ("does not parse: " + e)
    | Ok expr ->
        let text = exprToFormula expr
        match parseFormula text with
        | Ok again -> ok ("print " + formula) (again = expr) (sprintf "printed %s reparsed differently" text)
        | Result.Error e -> ok ("print " + formula) false (sprintf "printed %s does not parse: %s" text e)
let prints (formula: string) (expected: string) =
    match parseFormula formula with
    | Ok expr -> eq ("text " + formula) (exprToFormula expr) expected
    | Result.Error e -> ok ("text " + formula) false e

for f in
    [ "=1+2*3"; "=(1+2)*3"; "=1-(2-3)"; "=1-2-3"; "=2^3^2"; "=2^(3^2)"; "=-2^2"; "=-(2^2)"; "=2^-3"
      "=(1+2)%"; "=-A1%"; "=(-A1)%"; "=--A1"; "=A1&B1&\"x\""; "=(A1=B1)=(C1<>D1)"; "=A1<=B1"
      "=1/(2/3)"; "=1/2/3"; "=(A1&B1)+1"; "=\"say \"\"hi\"\"\""; "=1.5E-7+2.25"; "=1E+30"
      "=$A$1+A$1+$A1+A1"; "=SUM(A1:B2,C3)"; "=Sheet2!A1+'My Sheet'!$B$2"; "='It''s'!A1:B2"
      "='2024'!A1"; "=#REF!+#N/A"; "=#DIV/0!"; "=A1#"; "=SUM(A1#)"; "=TRUE"; "=IF(FALSE,1,2)"
      "=Sales[Amount]"; "=Sales[[#Totals],[Amount]]"; "=Sales[@Amount]"; "=Sales[@[Unit Price]]"
      "=Sales[[#All],[Qty]:[Amount]]"; "=Sales[#Headers]"; "=[@Qty]*[@Price]"; "=Sales[[Qty]:[Amount]]"
      "=Sales[@]"; "=SUM(Sales[Unit Price])"
      "=LET(x,A1,y,x*2,x+y)"; "=MyName*2"; "=SUM(MyRange)"; "=IF(A1>0,\"a,b\",\"(\")"
      "=1+2+3+4+5+6+7+8+9+10"; "=A1*(B1+C1)*(D1-E1)"; "=-(-(1))"; "=+A1"; "=A1%%" ] do
    roundTrip f

prints "= sum( A1 , B1 )" "sum(A1,B1)"
prints "=(1+2)*3" "(1+2)*3"
prints "=1+(2*3)" "1+2*3"
prints "=Sheet1!A1:B2" "Sheet1!A1:B2"
prints "='My Sheet'!A1" "'My Sheet'!A1"
prints "=_xlfn.XLOOKUP(A1,B1:B9,C1:C9)" "XLOOKUP(A1,B1:B9,C1:C9)"
prints "=A1!B2+XFD1!A1+XFE1!A1" "'A1'!B2+'XFD1'!A1+XFE1!A1"

// A chain thousands of operators long must print without recursing per operator.
let longChain = "=" + String.concat "+" [ for i in 1 .. 4000 -> sprintf "A%d" i ]
ok "print-long-chain" (shiftFormulaText 1 0 longChain = "=" + String.concat "+" [ for i in 2 .. 4001 -> sprintf "A%d" i ]) "long chain shifted wrongly"

// --- Insert / delete rows and columns (single sheet) ------------------------
let sumSheet = modelOf [ at 0 0, num 1.0; at 1 0, num 2.0; at 2 0, fml "=SUM(A1:A2)" ]

let insertedMid = sumSheet |> insertRow 1 |> recalcAndSpill None None
eq "insert-row-formula-moves" (formulaAt 3 0 insertedMid) "=SUM(A1:A3)"
eq "insert-row-value"         (valueAt 3 0 insertedMid) (CellValue.Number 3.0)
eq "insert-row-count"         insertedMid.RowCount 101

let insertedTop = sumSheet |> insertRow 0 |> recalcAndSpill None None
eq "insert-above-range" (formulaAt 3 0 insertedTop) "=SUM(A2:A3)"
eq "insert-above-value" (valueAt 3 0 insertedTop) (CellValue.Number 3.0)

let insertedBelow = sumSheet |> insertRow 2 |> recalcAndSpill None None
eq "insert-after-range" (formulaAt 3 0 insertedBelow) "=SUM(A1:A2)"

let deletedTop = sumSheet |> deleteRow 0 |> recalcAndSpill None None
eq "delete-row-shrinks"  (formulaAt 1 0 deletedTop) "=SUM(A1:A1)"
eq "delete-row-value"    (valueAt 1 0 deletedTop) (CellValue.Number 2.0)

let refSheet = modelOf [ at 0 0, num 5.0; at 1 0, num 7.0; at 4 0, fml "=A1+A2"; at 5 0, fml "=SUM(A1:A2)"; at 6 0, fml "=$A$2*2" ]
let refDeleted = refSheet |> deleteRow 0 |> recalcAndSpill None None
eq "delete-ref-is-REF"     (formulaAt 3 0 refDeleted) "=#REF!+A1"
eq "delete-ref-value"      (valueAt 3 0 refDeleted) (CellValue.Error "REF!")
eq "delete-range-shrinks"  (formulaAt 4 0 refDeleted) "=SUM(A1:A1)"
eq "delete-abs-ref-moves"  (formulaAt 5 0 refDeleted) "=$A$1*2"
eq "delete-abs-ref-value"  (valueAt 5 0 refDeleted) (CellValue.Number 14.0)
let bothDeleted = refSheet |> Rbt.Spreadsheet.Model.applyStructureEdit (DeleteRows (0, 2))
eq "delete-whole-range-REF" (formulaAt 3 0 bothDeleted) "=SUM(#REF!)"

let colSheet = modelOf [ at 0 0, num 1.0; at 0 1, num 2.0; at 0 2, fml "=SUM(A1:B1)"; at 1 2, fml "=B1*10" ]
let colInserted = colSheet |> insertColumn 1 |> recalcAndSpill None None
eq "insert-col-grows"   (formulaAt 0 3 colInserted) "=SUM(A1:C1)"
eq "insert-col-ref"     (formulaAt 1 3 colInserted) "=C1*10"
eq "insert-col-value"   (valueAt 1 3 colInserted) (CellValue.Number 20.0)
eq "insert-col-count"   colInserted.ColCount 27
let colDeleted = colSheet |> deleteColumn 1 |> recalcAndSpill None None
eq "delete-col-shrinks" (formulaAt 0 1 colDeleted) "=SUM(A1:A1)"
eq "delete-col-REF"     (formulaAt 1 1 colDeleted) "=#REF!*10"
eq "delete-col-count"   colDeleted.ColCount 25

// Everything stored by position moves with the rows/columns.
let dressed =
    { modelOf [ at 4 1, num 1.0; at 6 2, str "x" ] with
        MergedCells = [ { StartRow = 3; StartCol = 0; EndRow = 5; EndCol = 1 } ]
        RowHeights = Map.ofList [ 4, 40.0 ]
        ColumnWidths = Map.ofList [ 1, 150.0 ]
        ConditionalFormats = Map.ofList [ at 4 1, [ { Rule = GreaterThan 0.0; Style = defaultFormat } ] ]
        Validations = Map.ofList [ at 6 2, { Rule = ListValues [ "x" ]; ErrorMessage = None; InputMessage = None } ]
        Spill = Map.ofList [ at 7 0, at 6 0 ]
        Tables = [ { Name = "T"; StartRow = 4; StartCol = 2; EndRow = 8; EndCol = 3; HeaderRowCount = 1; Style = None; AutoFilter = Some (2, 3); HasTotals = false } ] }
let dressedIns = insertRow 4 dressed
eq "insert-merge-grows"   dressedIns.MergedCells [ { StartRow = 3; StartCol = 0; EndRow = 6; EndCol = 1 } ]
eq "insert-row-heights"   dressedIns.RowHeights (Map.ofList [ 5, 40.0 ])
eq "insert-cond-formats"  (Map.containsKey (at 5 1) dressedIns.ConditionalFormats) true
eq "insert-validations"   (Map.containsKey (at 7 2) dressedIns.Validations) true
eq "insert-spill-map"     dressedIns.Spill (Map.ofList [ at 8 0, at 7 0 ])
eq "insert-table-grows"   (dressedIns.Tables |> List.map (fun t -> t.StartRow, t.EndRow)) [ 4, 9 ]
let dressedDel = deleteRow 4 dressed
eq "delete-merge-shrinks" dressedDel.MergedCells [ { StartRow = 3; StartCol = 0; EndRow = 4; EndCol = 1 } ]
eq "delete-row-heights"   dressedDel.RowHeights Map.empty
eq "delete-cond-formats"  dressedDel.ConditionalFormats Map.empty
eq "delete-validations"   (Map.containsKey (at 5 2) dressedDel.Validations) true
eq "delete-spill-map"     dressedDel.Spill (Map.ofList [ at 6 0, at 5 0 ])
eq "delete-table-shrinks" (dressedDel.Tables |> List.map (fun t -> t.StartRow, t.EndRow)) [ 4, 7 ]
let dressedCol = insertColumn 1 dressed
eq "insert-col-widths"    dressedCol.ColumnWidths (Map.ofList [ 2, 150.0 ])
eq "insert-col-merge"     dressedCol.MergedCells [ { StartRow = 3; StartCol = 0; EndRow = 5; EndCol = 2 } ]
eq "insert-col-table"     (dressedCol.Tables |> List.map (fun t -> t.StartCol, t.EndCol, t.AutoFilter)) [ 3, 4, Some (3, 4) ]
eq "insert-col-cf"        (Map.containsKey (at 4 2) dressedCol.ConditionalFormats) true
let dressedColDel = deleteColumn 1 dressed
eq "delete-col-widths"    dressedColDel.ColumnWidths Map.empty
eq "delete-col-merge"     dressedColDel.MergedCells [ { StartRow = 3; StartCol = 0; EndRow = 5; EndCol = 0 } ]
eq "delete-col-cells"     (Map.containsKey (at 6 1) dressedColDel.Cells && not (Map.containsKey (at 4 1) dressedColDel.Cells)) true

// --- Insert / delete across a workbook --------------------------------------
let wbOf (sheets: (string * string * SpreadsheetModel) list) : Workbook =
    { Sheets = sheets |> List.map (fun (id, name, m) -> id, { Id = id; Name = name; Model = m; Visible = true }) |> Map.ofList
      SheetOrder = sheets |> List.map (fun (id, _, _) -> id)
      ActiveSheet = (let (id, _, _) = List.head sheets in id)
      Tables = Map.empty
      Names = Map.empty }
let sheetModel id (wb: Workbook) = wb.Sheets.[id].Model

let dataSheet = modelOf [ at 0 0, num 1.0; at 1 0, num 2.0; at 2 0, fml "=SUM(Data!A1:A2)"; at 3 0, fml "=Other!A1" ]
let otherSheet = modelOf [ at 0 0, fml "=Data!A2*10"; at 1 0, fml "=SUM(Data!A1:A2)"; at 2 0, fml "=A1+1"; at 3 0, fml "=SUM(Totals)" ]
let wb0 =
    { wbOf [ "sheet1", "Data", dataSheet; "sheet2", "Other", otherSheet ] with
        Names = namesOfTriples [ None, "Totals", "Data!$A$1:$A$2"; Some "sheet1", "Local", "$A$2"; Some "sheet2", "Mine", "$A$2" ]
        Tables = Map.ofList [ "T", { Name = "T"; SheetId = "sheet1"; StartRow = 0; StartCol = 0; EndRow = 1; EndCol = 1
                                     Columns = [ "a"; "b" ]; HasHeaders = true; HasTotals = false } ] }

let wb1, undo1 = Rbt.Spreadsheet.Workbook.applyStructureEdit "sheet1" (InsertRows (1, 1)) wb0
let wb1c = recalculateWorkbook wb1
eq "wb-own-qualified"     (formulaAt 3 0 (sheetModel "sheet1" wb1)) "=SUM(Data!A1:A3)"
eq "wb-own-other-sheet"   (formulaAt 4 0 (sheetModel "sheet1" wb1)) "=Other!A1"
eq "wb-foreign-ref"       (formulaAt 0 0 (sheetModel "sheet2" wb1)) "=Data!A3*10"
eq "wb-foreign-range"     (formulaAt 1 0 (sheetModel "sheet2" wb1)) "=SUM(Data!A1:A3)"
eq "wb-foreign-local-ref" (formulaAt 2 0 (sheetModel "sheet2" wb1)) "=A1+1"
eq "wb-foreign-value"     (valueAt 0 0 (sheetModel "sheet2" wb1c)) (CellValue.Number 20.0)
eq "wb-name-value"        (valueAt 3 0 (sheetModel "sheet2" wb1c)) (CellValue.Number 3.0)
eq "wb-global-name"       (getName None "Totals" wb1) (Some "Data!$A$1:$A$3")
eq "wb-local-name"        (getName (Some "sheet1") "Local" wb1) (Some "$A$3")
eq "wb-foreign-name"      (getName (Some "sheet2") "Mine" wb1) (Some "$A$2")
eq "wb-table-grows"       (wb1.Tables.["T"].StartRow, wb1.Tables.["T"].EndRow) (0, 2)
eq "wb-foreign-recorded"  (List.length undo1.ForeignCells) 2

let wbUndone = undoStructureEdit undo1 wb1c
eq "wb-undo-own"     (formulaAt 2 0 (sheetModel "sheet1" wbUndone)) "=SUM(Data!A1:A2)"
eq "wb-undo-foreign" (formulaAt 0 0 (sheetModel "sheet2" wbUndone)) "=Data!A2*10"
eq "wb-undo-names"   wbUndone.Names wb0.Names
eq "wb-undo-tables"  wbUndone.Tables wb0.Tables
eq "wb-undo-rows"    (sheetModel "sheet1" wbUndone).RowCount 100

// A foreign formula the user re-typed after the edit is not clobbered by undo.
let wbRetyped =
    updateSheetModel "sheet2" (setCell (at 0 0) (fml "=99") (sheetModel "sheet2" wb1)) wb1
    |> undoStructureEdit undo1
eq "wb-undo-keeps-retyped" (formulaAt 0 0 (sheetModel "sheet2" wbRetyped)) "=99"

let wbDel, _ = Rbt.Spreadsheet.Workbook.applyStructureEdit "sheet1" (DeleteRows (1, 1)) wb0
eq "wb-delete-foreign-REF"   (formulaAt 0 0 (sheetModel "sheet2" wbDel)) "=#REF!*10"
eq "wb-delete-foreign-range" (formulaAt 1 0 (sheetModel "sheet2" wbDel)) "=SUM(Data!A1:A1)"
eq "wb-delete-name"          (getName None "Totals" wbDel) (Some "Data!$A$1:$A$1")

let wbCols, _ = Rbt.Spreadsheet.Workbook.applyStructureEdit "sheet1" (InsertCols (1, 1)) wb0
eq "wb-table-new-column" wbCols.Tables.["T"].Columns [ "a"; "Column1"; "b" ]
eq "wb-table-col-span"   (wbCols.Tables.["T"].StartCol, wbCols.Tables.["T"].EndCol) (0, 2)
let wbColDel, _ = Rbt.Spreadsheet.Workbook.applyStructureEdit "sheet1" (DeleteCols (1, 1)) wb0
eq "wb-table-col-removed" wbColDel.Tables.["T"].Columns [ "a" ]
let wbHeaderGone, _ = Rbt.Spreadsheet.Workbook.applyStructureEdit "sheet1" (DeleteRows (0, 1)) wb0
eq "wb-table-header-deleted" (Map.containsKey "T" wbHeaderGone.Tables) false

// Workbook-wide recalculation: a sheet earlier in tab order that reads a later one.
let chainWb =
    wbOf [ "sheet1", "First", modelOf [ at 0 0, fml "=Second!A1+1" ]
           "sheet2", "Second", modelOf [ at 0 0, fml "=Third!A1+1" ]
           "sheet3", "Third", modelOf [ at 0 0, num 10.0; at 1 0, fml "=A1" ] ]
    |> recalculateWorkbook
eq "wb-recalc-chain" (valueAt 0 0 (sheetModel "sheet1" chainWb)) (CellValue.Number 12.0)

printfn ""
printfn "Stage 21 structure tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
