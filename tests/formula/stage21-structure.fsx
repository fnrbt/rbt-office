#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open System.IO
open System.IO.Compression
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
        ConditionalFormats = [ { Id = 1; Ranges = [ { From = at 4 1; To = at 4 1 } ]; Format = { Rule = GreaterThan 0.0; Style = defaultFormat }; Priority = 1; StopIfTrue = false } ]
        Validations = [ { Id = 1; Ranges = [ { From = at 6 2; To = at 6 2 } ]; Validation = validationOf (ListValues [ "x" ]) } ]
        Spill = Map.ofList [ at 7 0, at 6 0 ]
        Tables = [ { Name = "T"; StartRow = 4; StartCol = 2; EndRow = 8; EndCol = 3; HeaderRowCount = 1; Style = None; AutoFilter = Some (2, 3); HasTotals = false } ] }
let dressedIns = insertRow 4 dressed
eq "insert-merge-grows"   dressedIns.MergedCells [ { StartRow = 3; StartCol = 0; EndRow = 6; EndCol = 1 } ]
eq "insert-row-heights"   dressedIns.RowHeights (Map.ofList [ 5, 40.0 ])
eq "insert-cond-formats"  (conditionalFormatsAt (at 5 1) dressedIns |> List.length, conditionalFormatsAt (at 4 1) dressedIns |> List.length) (1, 0)
eq "insert-validations"   ((getValidation (at 7 2) dressedIns).IsSome, (getValidation (at 6 2) dressedIns).IsSome) (true, false)
eq "insert-spill-map"     dressedIns.Spill (Map.ofList [ at 8 0, at 7 0 ])
eq "insert-table-grows"   (dressedIns.Tables |> List.map (fun t -> t.StartRow, t.EndRow)) [ 4, 9 ]
let dressedDel = deleteRow 4 dressed
eq "delete-merge-shrinks" dressedDel.MergedCells [ { StartRow = 3; StartCol = 0; EndRow = 4; EndCol = 1 } ]
eq "delete-row-heights"   dressedDel.RowHeights Map.empty
eq "delete-cond-formats"  dressedDel.ConditionalFormats []
eq "delete-validations"   ((getValidation (at 5 2) dressedDel).IsSome, (getValidation (at 6 2) dressedDel).IsSome) (true, false)
eq "delete-spill-map"     dressedDel.Spill (Map.ofList [ at 6 0, at 5 0 ])
eq "delete-table-shrinks" (dressedDel.Tables |> List.map (fun t -> t.StartRow, t.EndRow)) [ 4, 7 ]
let dressedCol = insertColumn 1 dressed
eq "insert-col-widths"    dressedCol.ColumnWidths (Map.ofList [ 2, 150.0 ])
eq "insert-col-merge"     dressedCol.MergedCells [ { StartRow = 3; StartCol = 0; EndRow = 5; EndCol = 2 } ]
eq "insert-col-table"     (dressedCol.Tables |> List.map (fun t -> t.StartCol, t.EndCol, t.AutoFilter)) [ 3, 4, Some (3, 4) ]
eq "insert-col-cf"        (conditionalFormatsAt (at 4 2) dressedCol |> List.length) 1
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
      Names = Map.empty
      NameInfo = Map.empty; Iteration = None }
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

// References on a foreign rule change, but its application range and anchor
// belong to its own sheet, not the sheet whose rows/columns were edited.
let foreignRange = [ { From = at 5 5; To = at 7 6 } ]
let formulaThreshold text = { Kind = CfvoFormula; Value = Some text }
let percentThreshold = { Kind = CfvoPercent; Value = Some "50" }
let foreignFormat id rule =
    { Id = id; Ranges = foreignRange; Format = { Rule = rule; Style = defaultFormat }; Priority = id; StopIfTrue = false }
let foreignValidation id rule =
    { Id = id; Ranges = foreignRange; Validation = { validationOf rule with Anchor = Some (at 5 5) } }
let foreignRules formula source =
    let threshold = formulaThreshold formula
    [ foreignFormat 1 (Expression formula)
      foreignFormat 2 (CellIs (OpBetween, formula, Some formula))
      foreignFormat 3 (ColorScale [ threshold, "#FF0000"; percentThreshold, "#FFFFFF" ])
      foreignFormat 4 (DataBar (threshold, percentThreshold, "#638EC6"))
      foreignFormat 5 (IconSet ("3Arrows", [ threshold; percentThreshold ], false))
      foreignFormat 6 (DataBarEx (threshold, percentThreshold, "#638EC6", defaultDataBarOptions))
      foreignFormat 7 (IconSetEx ("3Arrows", [ threshold; percentThreshold ], false, defaultIconSetOptions))
      foreignFormat 8 (Expression "=A1") ],
    [ foreignValidation 1 (Comparison (WholeKind, OpBetween, formula, Some formula))
      foreignValidation 2 (CustomFormula formula)
      foreignValidation 3 (ListSource source)
      foreignValidation 4 (CustomFormula "=A1") ]
let originalForeignFormats, originalForeignValidations = foreignRules "=Data!$B$2" "Data!$B$2"
let foreignRuleSheet =
    { modelOf [] with
        ConditionalFormats = originalForeignFormats
        Validations = originalForeignValidations }
let foreignRuleWorkbook = wbOf [ "data", "Data", modelOf [ at 1 1, num 42.0 ]; "other", "Other", foreignRuleSheet ]
let ruleEdits =
    [ "insert-row", InsertRows (1, 1), "=Data!$B$3", "Data!$B$3"
      "delete-row", DeleteRows (1, 1), "=#REF!", "#REF!"
      "insert-col", InsertCols (1, 1), "=Data!$C$2", "Data!$C$2"
      "delete-col", DeleteCols (1, 1), "=#REF!", "#REF!" ]
for (label, edit, formula, source) in ruleEdits do
    let changed, undo = Rbt.Spreadsheet.Workbook.applyStructureEdit "data" edit foreignRuleWorkbook
    let foreign = sheetModel "other" changed
    let expectedFormats, expectedValidations = foreignRules formula source
    eq (label + "-foreign-formats") foreign.ConditionalFormats expectedFormats
    eq (label + "-foreign-validations") foreign.Validations expectedValidations
    let restored = sheetModel "other" (undoStructureEdit undo changed)
    eq (label + "-undo-formats") restored.ConditionalFormats originalForeignFormats
    eq (label + "-undo-validations") restored.Validations originalForeignValidations

// An intervening rule edit is not discarded by undoing the structural edit.
// Matching only the entire list (or even the entire rule) would lose these.
let ruleChanged, ruleUndo = Rbt.Spreadsheet.Workbook.applyStructureEdit "data" (InsertRows (1, 1)) foreignRuleWorkbook
let ruleEditedForeign = sheetModel "other" ruleChanged
let editedFormats =
    ruleEditedForeign.ConditionalFormats
    |> List.map (fun entry ->
        match entry.Id with
        | 1 -> { entry with Format = { entry.Format with Rule = Expression "=99" } }
        | 2 -> { entry with Format = { entry.Format with Rule = CellIs (OpBetween, "=Data!$B$3", Some "=77") } }
        | 3 -> { entry with Priority = 20; Ranges = [ { From = at 9 9; To = at 9 9 } ]
                            Format = { entry.Format with Style = { entry.Format.Style with Bold = true } } }
        | _ -> entry)
    |> fun entries -> foreignFormat 20 (Expression "=Other!A1") :: entries
let editedValidations =
    ruleEditedForeign.Validations
    |> List.map (fun entry ->
        match entry.Id with
        | 1 ->
            { entry with Validation = { entry.Validation with
                                            Rule = Comparison (WholeKind, OpBetween, "=Data!$B$3", Some "=88")
                                            InputMessage = Some "new message" } }
        | 2 -> { entry with Validation = { entry.Validation with Rule = CustomFormula "=99" } }
        | _ -> entry)
    |> fun entries -> foreignValidation 20 (CustomFormula "=Other!A1") :: entries
let ruleRetyped =
    ruleChanged
    |> updateSheetModel "other" { ruleEditedForeign with ConditionalFormats = editedFormats; Validations = editedValidations }
    |> undoStructureEdit ruleUndo
let restoredForeign = sheetModel "other" ruleRetyped
eq "rule-undo-keeps-new-format" restoredForeign.ConditionalFormats.Head (foreignFormat 20 (Expression "=Other!A1"))
eq "rule-undo-keeps-retyped-expression" restoredForeign.ConditionalFormats.[1].Format.Rule (Expression "=99")
eq "rule-undo-restores-only-cellis-first" restoredForeign.ConditionalFormats.[2].Format.Rule
    (CellIs (OpBetween, "=Data!$B$2", Some "=77"))
eq "rule-undo-keeps-range-priority-style"
    (restoredForeign.ConditionalFormats.[3].Ranges, restoredForeign.ConditionalFormats.[3].Priority,
     restoredForeign.ConditionalFormats.[3].Format.Style.Bold, restoredForeign.ConditionalFormats.[3].Format.Rule)
    ([ { From = at 9 9; To = at 9 9 } ], 20, true,
     ColorScale [ formulaThreshold "=Data!$B$2", "#FF0000"; percentThreshold, "#FFFFFF" ])
eq "rule-undo-keeps-new-validation" restoredForeign.Validations.Head (foreignValidation 20 (CustomFormula "=Other!A1"))
eq "rule-undo-keeps-retyped-validation" restoredForeign.Validations.[2].Validation.Rule (CustomFormula "=99")
eq "rule-undo-restores-only-comparison-first"
    (restoredForeign.Validations.[1].Validation.Rule, restoredForeign.Validations.[1].Validation.InputMessage)
    (Comparison (WholeKind, OpBetween, "=Data!$B$2", Some "=88"), Some "new message")

// Workbook-wide recalculation: a sheet earlier in tab order that reads a later one.
let chainWb =
    wbOf [ "sheet1", "First", modelOf [ at 0 0, fml "=Second!A1+1" ]
           "sheet2", "Second", modelOf [ at 0 0, fml "=Third!A1+1" ]
           "sheet3", "Third", modelOf [ at 0 0, num 10.0; at 1 0, fml "=A1" ] ]
    |> recalculateWorkbook
eq "wb-recalc-chain" (valueAt 0 0 (sheetModel "sheet1" chainWb)) (CellValue.Number 12.0)

// --- Sort -------------------------------------------------------------------
// Data in rows 5-9 (0-indexed 4..8) with a header; a note sits apart in row 1.
let sortSheet =
    modelOf [ at 0 0, str "note"
              at 4 0, str "Name";  at 4 1, str "Qty";  at 4 2, str "Double"
              at 5 0, str "pear";  at 5 1, num 3.0;    at 5 2, fml "=B6*2"
              at 6 0, str "apple"; at 6 1, num 1.0;    at 6 2, fml "=B7*2"
              at 7 0, str "fig";                       at 7 2, fml "=B8*2"
              at 8 0, str "kiwi";  at 8 1, num 2.0;    at 8 2, fml "=B9*2"
              at 10 0, fml "=B6" ]
let colValues c (m: SpreadsheetModel) = [ for r in 4 .. 8 -> valueAt r c m ]
let S = CellValue.String
let N = CellValue.Number

let sortedAsc = sortSheet |> sortByColumn 1 true |> recalcAndSpill None None
eq "sort-header-stays"   (valueAt 4 0 sortedAsc, valueAt 4 1 sortedAsc) (S "Name", S "Qty")
eq "sort-in-place"       (colValues 0 sortedAsc) [ S "Name"; S "apple"; S "kiwi"; S "pear"; S "fig" ]
eq "sort-blank-last"     (valueAt 8 1 sortedAsc) CellValue.Empty
eq "sort-outside-kept"   (valueAt 0 0 sortedAsc) (S "note")
eq "sort-rows-not-moved-up" (Map.containsKey (at 1 0) sortedAsc.Cells) false
eq "sort-formula-follows" (formulaAt 5 2 sortedAsc) "=B6*2"
eq "sort-formula-values" (colValues 2 sortedAsc) [ S "Double"; N 2.0; N 4.0; N 6.0; N 0.0 ]
eq "sort-outside-formula-recalc" (valueAt 10 0 sortedAsc) (N 1.0)
eq "sort-state"          sortedAsc.SortState (Some { Column = 1; Ascending = true })

let sortedDesc = sortSheet |> sortByColumn 1 false |> recalcAndSpill None None
eq "sort-desc"            (colValues 0 sortedDesc) [ S "Name"; S "pear"; S "kiwi"; S "apple"; S "fig" ]
eq "sort-desc-blank-last" (valueAt 8 1 sortedDesc) CellValue.Empty

let sortedText = sortSheet |> sortByColumn 0 true
eq "sort-text-column" (colValues 0 sortedText) [ S "Name"; S "apple"; S "fig"; S "kiwi"; S "pear" ]

// No header when the first row holds data, not labels.
let plainSorted = modelOf [ at 0 0, num 3.0; at 1 0, num 1.0; at 2 0, num 2.0 ] |> sortByColumn 0 true
eq "sort-no-header" [ for r in 0 .. 2 -> valueAt r 0 plainSorted ] [ N 1.0; N 2.0; N 3.0 ]

// AutoFilter: the header is the first row of the filtered block, not the
// first used row of the sheet, and rows outside the block are never hidden.
let filtered = sortSheet |> applyFilter 0 (Set.ofList [ "apple" ]) |> applyFilterRowHeights
eq "filter-header-row"     (getRowHeight 4 filtered) 25.0
eq "filter-match-kept"     (getRowHeight 6 filtered) 25.0
eq "filter-others-hidden"  ([ 5; 7; 8 ] |> List.map (fun r -> getRowHeight r filtered)) [ 0.0; 0.0; 0.0 ]
eq "filter-outside-block"  (getRowHeight 0 filtered, getRowHeight 10 filtered) (25.0, 25.0)

// --- A value typed into a spill range blocks the spill ----------------------
let spillBase = modelOf [ at 0 0, fml "=SEQUENCE(3)" ] |> recalcAndSpill None None
eq "spill-initial" [ for r in 0 .. 2 -> valueAt r 0 spillBase ] [ N 1.0; N 2.0; N 3.0 ]
let spillBlocked = spillBase |> setCell (at 1 0) (num 99.0) |> recalcAndSpill None None
eq "spill-user-value-kept" (valueAt 1 0 spillBlocked) (N 99.0)
eq "spill-anchor-error"    (valueAt 0 0 spillBlocked) (CellValue.Error "SPILL!")
eq "spill-ghost-cleared"   (Map.containsKey (at 2 0) spillBlocked.Cells) false
let spillFreed = spillBlocked |> clearCell (at 1 0) |> recalcAndSpill None None
eq "spill-returns" [ for r in 0 .. 2 -> valueAt r 0 spillFreed ] [ N 1.0; N 2.0; N 3.0 ]
// Formatting a ghost cell does not turn it into an obstruction.
let spillStyled = spillBase |> setCellBgColor (at 1 0) "#FF0000" |> recalcAndSpill None None
eq "spill-styled-ghost" (valueAt 0 0 spillStyled, valueAt 1 0 spillStyled) (N 1.0, N 2.0)

// Imported pivot layouts keep line offsets relative to the moved location.
// GETPIVOTDATA depends on the labels AND the selected value, on either sheet.
let pivotLayout : PivotLayout =
    { Top = 10; Left = 2; Bottom = 14; Right = 4
      FirstHeaderRow = 1; FirstDataRow = 2; FirstDataCol = 1
      RowFields = [ "Region" ]; ColumnFields = [ "Quarter" ]; DataFields = [ "Sum of Sales" ]
      RowLines = [ { Kind = "data"; Repeat = 0; Items = [ 0 ] }
                   { Kind = "data"; Repeat = 0; Items = [ 1 ] }
                   { Kind = "grand"; Repeat = 0; Items = [] } ]
      ColumnLines = [ { Kind = "data"; Repeat = 0; Items = [ 0 ] }
                      { Kind = "grand"; Repeat = 0; Items = [] } ] }
let pivotData =
    { modelOf
        [ at 11 3, str "Q1"; at 11 4, str "Grand Total"
          at 12 2, str "East"; at 13 2, str "West"; at 14 2, str "Grand Total"
          at 12 3, num 10.0; at 12 4, num 15.0
          at 13 3, num 20.0; at 13 4, num 25.0
          at 14 3, num 30.0; at 14 4, num 40.0 ] with
        Pivots = [ pivotLayout ] }
let pivotQuery = "=GETPIVOTDATA(\"Sales\",Data!C11,\"Region\",\"East\",\"Quarter\",\"Q1\")"
let pivotTotal = "=GETPIVOTDATA(\"Sales\",Data!C11)"
let pivotOther = modelOf [ at 0 0, fml pivotQuery; at 1 0, fml pivotTotal; at 2 0, fml "=A1+1"; at 0 3, fml "=7" ]
let pivotWb = wbOf [ "pivot", "Data", pivotData; "lookup", "Other", pivotOther ] |> recalculateWorkbook
let pivotValues (book: Workbook) =
    let lookup = sheetModel "lookup" book
    [ valueAt 0 0 lookup; valueAt 1 0 lookup; valueAt 2 0 lookup ]
eq "pivot-cross-sheet-initial" (pivotValues pivotWb) [ CellValue.Number 10.0; CellValue.Number 40.0; CellValue.Number 11.0 ]
let pivotChange label edit expectedGeometry expected =
    let changed, undo = Rbt.Spreadsheet.Workbook.applyStructureEdit "pivot" edit pivotWb
    let calculated = recalculateWorkbook changed
    eq (label + "-geometry") ((sheetModel "pivot" changed).Pivots |> List.map (fun p ->
        p.Top, p.Left, p.Bottom, p.Right, p.FirstHeaderRow, p.FirstDataRow, p.FirstDataCol,
        p.RowLines |> List.map (fun line -> line.Kind), p.ColumnLines |> List.map (fun line -> line.Kind))) expectedGeometry
    eq (label + "-values") (pivotValues calculated) expected
    let restored = undoStructureEdit undo calculated |> recalculateWorkbook
    eq (label + "-undo") (pivotValues restored) (pivotValues pivotWb)
    let redone, _ = Rbt.Spreadsheet.Workbook.applyStructureEdit "pivot" edit restored
    eq (label + "-redo") (pivotValues (recalculateWorkbook redone)) expected
let kindsRow = [ "data"; "data"; "grand" ]
let kindsCol = [ "data"; "grand" ]
let pivotResult q total = [ q; total; (match q with CellValue.Number n -> CellValue.Number (n + 1.0) | _ -> q) ]
pivotChange "pivot-insert-header-row" (InsertRows (11, 1))
    [ 10, 2, 15, 4, 2, 3, 1, kindsRow, kindsCol ]
    (pivotResult (CellValue.Number 10.0) (CellValue.Number 40.0))
pivotChange "pivot-insert-data-boundary-col" (InsertCols (3, 1))
    [ 10, 2, 14, 5, 1, 2, 2, kindsRow, kindsCol ]
    (pivotResult (CellValue.Number 10.0) (CellValue.Number 40.0))
pivotChange "pivot-insert-before-row" (InsertRows (9, 2))
    [ 12, 2, 16, 4, 1, 2, 1, kindsRow, kindsCol ] (pivotResult (CellValue.Number 10.0) (CellValue.Number 40.0))
pivotChange "pivot-insert-before-col" (InsertCols (1, 2))
    [ 10, 4, 14, 6, 1, 2, 1, kindsRow, kindsCol ] (pivotResult (CellValue.Number 10.0) (CellValue.Number 40.0))
pivotChange "pivot-insert-data-row" (InsertRows (13, 1))
    [ 10, 2, 15, 4, 1, 2, 1, [ "data"; "blank"; "data"; "grand" ], kindsCol ]
    (pivotResult (CellValue.Number 10.0) (CellValue.Number 40.0))
pivotChange "pivot-insert-data-col" (InsertCols (4, 1))
    [ 10, 2, 14, 5, 1, 2, 1, kindsRow, [ "data"; "blank"; "grand" ] ]
    (pivotResult (CellValue.Number 10.0) (CellValue.Number 40.0))
pivotChange "pivot-delete-east-row" (DeleteRows (12, 1))
    [ 10, 2, 13, 4, 1, 2, 1, [ "data"; "grand" ], kindsCol ]
    (pivotResult (CellValue.Error "REF!") (CellValue.Number 40.0))
pivotChange "pivot-delete-header-row" (DeleteRows (11, 1))
    [ 10, 2, 13, 4, 1, 1, 1, kindsRow, kindsCol ]
    (pivotResult (CellValue.Error "REF!") (CellValue.Number 40.0))
pivotChange "pivot-delete-west-row" (DeleteRows (13, 1))
    [ 10, 2, 13, 4, 1, 2, 1, [ "data"; "grand" ], kindsCol ]
    (pivotResult (CellValue.Number 10.0) (CellValue.Number 40.0))
pivotChange "pivot-delete-total-col" (DeleteCols (4, 1))
    [ 10, 2, 14, 3, 1, 2, 1, kindsRow, [ "data" ] ]
    (pivotResult (CellValue.Number 10.0) (CellValue.Error "REF!"))
pivotChange "pivot-delete-q1-col" (DeleteCols (3, 1))
    [ 10, 2, 14, 3, 1, 2, 1, kindsRow, [ "grand" ] ]
    (pivotResult (CellValue.Error "REF!") (CellValue.Number 40.0))
pivotChange "pivot-delete-all-rows" (DeleteRows (10, 5)) []
    (pivotResult (CellValue.Error "REF!") (CellValue.Error "REF!"))
pivotChange "pivot-delete-all-cols" (DeleteCols (2, 3)) []
    (pivotResult (CellValue.Error "REF!") (CellValue.Error "REF!"))
// The graph must follow the value cell, not just the pivot anchor, and
// recognize a moved/deleted layout even if the formula is on another sheet.
let pivotEngine = Rbt.Spreadsheet.Recalc.build pivotWb
let pivotTyped = updateSheetModel "pivot" (setCell (at 12 3) (num 99.0) pivotData) pivotWb
let pivotTypedCalc = Rbt.Spreadsheet.Recalc.recalculate pivotEngine pivotTyped None
eq "pivot-incremental-value" (pivotTypedCalc |> Option.map pivotValues)
    (Some (pivotResult (CellValue.Number 99.0) (CellValue.Number 40.0)))
eq "pivot-incremental-selective" (pivotEngine.LastStats.EvaluatedFormulas < 4) true
let pivotRelabeled = updateSheetModel "pivot" (setCell (at 12 2) (str "North") pivotData) pivotWb
eq "pivot-incremental-label"
    (Rbt.Spreadsheet.Recalc.recalculate (Rbt.Spreadsheet.Recalc.build pivotWb) pivotRelabeled None |> Option.map pivotValues)
    (Some (pivotResult (CellValue.Error "REF!") (CellValue.Number 40.0)))
let pivotRemoved = updateSheetModel "pivot" { pivotData with Pivots = [] } pivotWb
eq "pivot-incremental-metadata"
    (Rbt.Spreadsheet.Recalc.recalculate (Rbt.Spreadsheet.Recalc.build pivotWb) pivotRemoved None |> Option.map pivotValues)
    (Some (pivotResult (CellValue.Error "REF!") (CellValue.Error "REF!")))
let pivotMoved, _ = Rbt.Spreadsheet.Workbook.applyStructureEdit "pivot" (InsertRows (13, 1)) pivotWb
let pivotMovedCalc = Rbt.Spreadsheet.Recalc.recalculate (Rbt.Spreadsheet.Recalc.build pivotWb) pivotMoved None
eq "pivot-incremental-structure" (pivotMovedCalc |> Option.map pivotValues)
    (Some (pivotResult (CellValue.Number 10.0) (CellValue.Number 40.0)))
let pivotColDeleted, _ = Rbt.Spreadsheet.Workbook.applyStructureEdit "pivot" (DeleteCols (3, 1)) pivotWb
eq "pivot-incremental-delete-col"
    (Rbt.Spreadsheet.Recalc.recalculate (Rbt.Spreadsheet.Recalc.build pivotWb) pivotColDeleted None |> Option.map pivotValues)
    (Some (pivotResult (CellValue.Error "REF!") (CellValue.Number 40.0)))

// Persist the pivot part's location and its internal line geometry, then
// reopen it through the same XLSX reader as a user file.
let pivotFile = Path.Combine(__SOURCE_DIRECTORY__, "..", "test-documents-excel", "synthetic-slicers.xlsx")
let pivotImported = Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook
                        (XlsxReader.loadBytesNative (File.ReadAllBytes pivotFile) |> Option.get)
let pivotModels = pivotImported.Sheets |> Array.map (fun sheet -> sheet.Model)
let pivotIndex = pivotModels |> Array.findIndex (fun model -> not model.Pivots.IsEmpty)
let pivotFileModel = pivotModels.[pivotIndex]
let pivotFileLayout = pivotFileModel.Pivots.Head
let pivotFileChanged =
    insertRow (pivotFileLayout.Top + pivotFileLayout.FirstDataRow + 1) pivotFileModel
    |> insertColumn (pivotFileLayout.Left + pivotFileLayout.FirstDataCol)
let savePivotFixtureFrom source pivotModel =
    let sheets = Array.copy pivotModels
    sheets.[pivotIndex] <- pivotModel
    let export : Rbt.Spreadsheet.XlsxBridge.ExportInput =
        { Sheets = sheets
          Names = pivotImported.Sheets |> Array.map (fun sheet -> sheet.Name)
          Visible = pivotImported.Sheets |> Array.map (fun sheet -> sheet.Visible)
          ActiveSheetIndex = pivotIndex
          DefinedNames = pivotImported.Context.Names; NameInfo = Map.empty
          Source = Some source; Macros = None; MacroProject = None }
    XlsxPackageWriter.writeWorkbookPackageBytes (Rbt.Spreadsheet.XlsxBridge.Export.toWorkbookOut export)
let savePivotFixture = savePivotFixtureFrom pivotImported.Source
let pivotFileSaved = savePivotFixture pivotFileChanged
let pivotReopened =
    Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook (XlsxReader.loadBytesNative pivotFileSaved |> Option.get)
    |> fun book -> book.Sheets.[pivotIndex].Model
eq "pivot-save-reopen-layout" pivotReopened.Pivots pivotFileChanged.Pivots
eq "pivot-save-reopen-cells"
    [ for r in pivotFileLayout.Top .. pivotFileLayout.Bottom + 1 do
        for c in pivotFileLayout.Left .. pivotFileLayout.Right + 1 do
            yield r, c, valueAt r c pivotReopened ]
    [ for r in pivotFileLayout.Top .. pivotFileLayout.Bottom + 1 do
        for c in pivotFileLayout.Left .. pivotFileLayout.Right + 1 do
            yield r, c, valueAt r c pivotFileChanged ]
let firstFileRow = pivotFileLayout.Top + pivotFileLayout.FirstDataRow
let firstFileLabel =
    match valueAt firstFileRow pivotFileLayout.Left pivotFileModel with
    | CellValue.String label -> label
    | value -> failwithf "pivot fixture row label: %A" value
let fileQuery =
    sprintf "=GETPIVOTDATA(\"%s\",A3,\"%s\",\"%s\")"
        pivotFileLayout.DataFields.Head pivotFileLayout.RowFields.Head firstFileLabel
let expectedFileValue = valueAt firstFileRow (pivotFileLayout.Left + pivotFileLayout.FirstDataCol) pivotFileModel
eq "pivot-save-live-lookup" (evaluateFormula fileQuery pivotFileChanged None) expectedFileValue
eq "pivot-save-reopen-lookup" (evaluateFormula fileQuery pivotReopened None) expectedFileValue
let pivotFileGone =
    Rbt.Spreadsheet.Model.applyStructureEdit
        (DeleteRows (pivotFileLayout.Top, pivotFileLayout.Bottom - pivotFileLayout.Top + 1)) pivotFileModel
let pivotGoneReopened =
    Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook
        (XlsxReader.loadBytesNative (savePivotFixture pivotFileGone) |> Option.get)
    |> fun book -> book.Sheets.[pivotIndex].Model
eq "pivot-save-reopen-whole-deletion" pivotGoneReopened.Pivots []

// An <i> can carry its data-field index and extensions that PivotLine does
// not model. The saved pivot must keep surviving items' XML byte-for-byte,
// even as structural edits add blank items or remove old ones. Qualifying
// <location> also exercises the source-span patch with namespace prefixes.
let pivotPart = "xl/pivotTables/pivotTable1.xml"
let originalPivotXml =
    pivotImported.Source.Workbook.Preserved.Parts
    |> List.find (fun (path, _) -> path = pivotPart)
    |> snd |> System.Text.Encoding.UTF8.GetString
let firstItem = "<i i=\"0\" extra=\"first\"><x v=\"1\" note=\"keep\"/><custom source=\"first\"/></i>"
let secondItem = "<i i=\"1\" extra=\"second\"><x v=\"2\" note=\"keep\"/><custom source=\"second\"/></i>"
let originalItems = "<rowItems count=\"3\"><i><x v=\"1\"/></i><i><x v=\"2\"/></i>"
eq "pivot-preserved-fixture-items" (originalPivotXml.Contains originalItems) true
let decoratedPivotXml =
    originalPivotXml.Replace(originalItems, "<rowItems count=\"3\">" + firstItem + secondItem)
    |> fun xml ->
        xml.Replace(" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"",
                    " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:p=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"")
    |> fun xml -> xml.Replace("<location ", "<p:location ")
let decoratedSource =
    let original = pivotImported.Source
    let preserved = original.Workbook.Preserved
    let parts = preserved.Parts |> List.map (fun (path, data) ->
        if path = pivotPart then path, System.Text.Encoding.UTF8.GetBytes decoratedPivotXml else path, data)
    { original with Workbook = { original.Workbook with Preserved = { preserved with Parts = parts } } }
let savedPivotXml model =
    let bytes = savePivotFixtureFrom decoratedSource model
    use archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read)
    use entry = archive.GetEntry(pivotPart).Open()
    use reader = new StreamReader(entry)
    reader.ReadToEnd()
let insertedPivotXml = savedPivotXml pivotFileChanged
eq "pivot-save-prefixed-location" (insertedPivotXml.Contains "<p:location ref=\"A3:C7\"") true
eq "pivot-save-first-i-attrs-and-child" (insertedPivotXml.Contains firstItem) true
eq "pivot-save-second-i-attrs-and-child" (insertedPivotXml.Contains secondItem) true
eq "pivot-save-inserted-blank" (insertedPivotXml.Contains "<i t=\"blank\"/>") true
let deletedPivotXml = savedPivotXml (deleteRow firstFileRow pivotFileModel)
eq "pivot-save-deleted-first-i" (deletedPivotXml.Contains firstItem) false
eq "pivot-save-surviving-second-i" (deletedPivotXml.Contains secondItem) true



printfn ""
printfn "Stage 21 structure tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
