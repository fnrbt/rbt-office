#I "../../spreadsheet-native/bin/Debug/net10.0"
#r "Rbt.Grid.dll"
#r "Rbt.Vb.dll"
#r "Rbt.OfficeUi.dll"
#r "Rbt.OfficeFormat.dll"
#r "Rbt.Renderer.dll"
#r "Rbt.Widgets.dll"
#r "Rbt.Spreadsheet.dll"
#r "Rbt.SpreadsheetWidgets.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook
open Rbt.Widgets.Composit.Ui.SpreadsheetEditor

let addr row col = { Row = row; Col = col }
let data value = { Value = value; Format = defaultFormat; Comment = None }
let value sheet row col (state: SpreadsheetEditorState) =
    state.Workbook.Value.Sheets.[sheet].Model |> getCellValue (addr row col) |> unwrapFormulaValue
let mutable checks = 0
let eq label expected actual =
    checks <- checks + 1
    if actual <> expected then failwithf "%s: expected %A, got %A" label expected actual
let key name shift ctrl state =
    handleKey state { Key = name; Shift = shift; Ctrl = ctrl; Alt = false } |> ignore
let enter text state =
    editFormula state text
    commitEdit state
let select selection state =
    applyChangeNoHistory state (fun model -> { model with Selection = selection }) |> ignore
let n value = CellValue.Number value

let first =
    createEmpty 30 12
    |> setCell (addr 0 0) (data (n 2.0))
    |> setCell (addr 0 1) (data (CellValue.Formula ("A1*5", CellValue.Empty)))
let second =
    createEmpty 30 12
    |> setCell (addr 0 0) (data (CellValue.Formula ("Data!B1+1", CellValue.Empty)))
let sheet id name model : Sheet = { Id = id; Name = name; Model = model; Visible = true }
let book =
    { Sheets = Map.ofList [ "source-17", sheet "source-17" "Data" first
                            "results-23", sheet "results-23" "Results" second ]
      SheetOrder = [ "source-17"; "results-23" ]
      ActiveSheet = "source-17"
      Tables = Map.empty
      Names = Map.empty
      NameInfo = Map.empty
      Iteration = Some { MaxIterations = 40; MaxChange = 0.0001 } }
    |> setDefinedName None "Global" "Data!$A$1" (Some "preserved")
    |> setDefinedName (Some "results-23") "Local" "Data!$A$1" None
let state = createStateFromWorkbook book
let originalNames = state.Workbook.Value.NameInfo
let iteration = state.Workbook.Value.Iteration

// Real commit and keyboard Delete take the same workbook-aware calculation path.
enter "3" state
eq "same-sheet dependent" (n 15.0) (value "source-17" 0 1 state)
eq "cross-sheet dependent" (n 16.0) (value "results-23" 0 0 state)
key "Delete" false false state
eq "deleted precedent" CellValue.Empty (value "source-17" 0 0 state)
eq "same-sheet after clear" (n 0.0) (value "source-17" 0 1 state)
eq "foreign after clear" (n 1.0) (value "results-23" 0 0 state)
key "z" false true state
eq "undo clear foreign" (n 16.0) (value "results-23" 0 0 state)
key "y" false true state
eq "redo clear foreign" (n 1.0) (value "results-23" 0 0 state)

// A selected rectangle, not just its focused corner, clears in one undo step.
select (Range (0, 0, 0, 1)) state
key "Delete" false false state
eq "range removed formula" CellValue.Empty (value "source-17" 0 1 state)
eq "range removed foreign dependent" (n 1.0) (value "results-23" 0 0 state)
key "z" false true state
eq "range undo restores formula" (n 0.0) (value "source-17" 0 1 state)

// Rebuild a source value and move it across sheets with keyboard cut/paste.
select (SingleCell (0, 0)) state
enter "7" state
key "x" false true state
eq "cut source cleared" CellValue.Empty (value "source-17" 0 0 state)
switchSheet state 1
select (SingleCell (1, 0)) state
key "v" false true state
eq "foreign paste value" (n 7.0) (value "results-23" 1 0 state)
eq "cut dependent repointed" (n 35.0) (value "source-17" 0 1 state)
eq "global name repointed" "Results!$A$2" (state.Workbook.Value.Names.[None, "GLOBAL"])
eq "local name repointed" "Results!$A$2" (state.Workbook.Value.Names.[Some "results-23", "LOCAL"])
key "z" false true state
eq "undo paste dependent" (n 0.0) (value "source-17" 0 1 state)
eq "undo paste name" "Data!$A$1" (state.Workbook.Value.Names.[None, "GLOBAL"])
key "z" false true state
eq "undo cut source" (n 7.0) (value "source-17" 0 0 state)
key "y" false true state
key "y" false true state
eq "redo move value" (n 7.0) (value "results-23" 1 0 state)
eq "redo move name" "Results!$A$2" (state.Workbook.Value.Names.[None, "GLOBAL"])

// A spilling formula reacts to typed source edits and clears orphan spill cells.
switchSheet state 1
select (SingleCell (0, 3)) state
enter "3" state
select (SingleCell (0, 4)) state
enter "=SEQUENCE(D1)" state
eq "spill initial third" (n 3.0) (value "results-23" 2 4 state)
select (SingleCell (0, 3)) state
enter "1" state
eq "spill shrunk" CellValue.Empty (value "results-23" 2 4 state)
key "z" false true state
eq "spill restored on undo" (n 3.0) (value "results-23" 2 4 state)
key "y" false true state
eq "spill shrink on redo" CellValue.Empty (value "results-23" 2 4 state)

// Copy retains the source model and shifts formulas at the target sheet.
let copyState = createStateFromWorkbook book
select (SingleCell (0, 1)) copyState
key "c" false true copyState
switchSheet copyState 1
select (SingleCell (1, 0)) copyState
enter "7" copyState
select (SingleCell (1, 1)) copyState
key "v" false true copyState
eq "foreign copy shifted formula" (n 35.0) (value "results-23" 1 1 copyState)
eq "copy keeps source" (n 10.0) (value "source-17" 0 1 copyState)
key "z" false true copyState
eq "undo foreign copy" CellValue.Empty (value "results-23" 1 1 copyState)
key "y" false true copyState
eq "redo foreign copy" (n 35.0) (value "results-23" 1 1 copyState)

// Structural commands preserve imported names and metadata; IDs never recycle.
let originalOrder = state.Workbook.Value.SheetOrder
insertSheet state "Temp" |> ignore
let insertedId = state.Workbook.Value.ActiveSheet
renameSheet state "source-17" "Inputs" |> ignore
eq "renamed UI projection" "Inputs" state.SheetNames.Value.[0]
eq "renamed dependency" (n 35.0) (value "source-17" 0 1 state)
eq "renamed name" "Results!$A$2" (state.Workbook.Value.Names.[None, "GLOBAL"])
deleteSheet state insertedId |> ignore
eq "deleted temporary sheet" false (Map.containsKey insertedId state.Workbook.Value.Sheets)
eq "delete preserves active valid identity" "results-23" state.Workbook.Value.ActiveSheet
key "z" false true state
eq "undo sheet delete identity" true (Map.containsKey insertedId state.Workbook.Value.Sheets)
key "z" false true state
eq "undo rename" "Data" state.Workbook.Value.Sheets.["source-17"].Name
eq "undo rename UI projection" "Data" state.SheetNames.Value.[0]
key "z" false true state
eq "undo insertion order" originalOrder state.Workbook.Value.SheetOrder
insertSheet state "New" |> ignore
if state.Workbook.Value.ActiveSheet = insertedId then failwith "sheet id was recycled after undo and divergent insertion"
let beforeMove = state.Workbook.Value.SheetOrder
moveSheet state "results-23" 0 |> ignore
eq "moved sheet order" "results-23" state.Workbook.Value.SheetOrder.Head
key "z" false true state
eq "undo sheet move" beforeMove state.Workbook.Value.SheetOrder
eq "imported name metadata" originalNames state.Workbook.Value.NameInfo
eq "imported iteration settings" iteration state.Workbook.Value.Iteration

// Imported structured-reference column labels survive model table edits.
let importedTable : Rbt.Spreadsheet.Model.Table =
    { Name = "Stock"; StartRow = 1; StartCol = 1; EndRow = 2; EndCol = 1
      HeaderRowCount = 1; Style = None; AutoFilter = None; HasTotals = false }
let addedTable : Rbt.Spreadsheet.Model.Table =
    { importedTable with Name = "NewStock"; StartCol = 4; EndCol = 4 }
let modelWithTable = { first with Tables = [ importedTable ] }
let tableBook =
    { book with
        Sheets = Map.add "source-17" (sheet "source-17" "Data" modelWithTable) book.Sheets
        Tables = Map.ofList [ "Stock", makeTableFromRegion "source-17" "Stock" 1 1 2 1 1 0 [ "Declared" ] modelWithTable ] }
let tableState = createStateFromWorkbook tableBook
applyChange tableState (fun m -> { m with Tables = importedTable :: [ addedTable ] }) |> ignore
eq "declared imported table column" [ "Declared" ] tableState.Workbook.Value.Tables.["Stock"].Columns
eq "added table registry" true (Map.containsKey "NewStock" tableState.Workbook.Value.Tables)
key "z" false true tableState
eq "undo added table" false (Map.containsKey "NewStock" tableState.Workbook.Value.Tables)
key "y" false true tableState
applyChange tableState (fun m -> { m with Tables = [ addedTable ] }) |> ignore
eq "removed table registry" false (Map.containsKey "Stock" tableState.Workbook.Value.Tables)

// Deleted source identities cannot be used to replay an internal clipboard move.
switchSheet tableState 0
key "c" false true tableState
deleteSheet tableState "source-17" |> ignore
if tableState.Clipboard.Value.IsSome then failwith "removed source clipboard not invalidated"
let rewritten = tableState.Workbook.Value.Sheets.["results-23"].Model |> getCellValue (addr 0 0)
match rewritten with
| CellValue.Formula (formula, _) when formula.Contains "#REF!" -> ()
| other -> failwithf "deleted sheet references not rewritten: %A" other
if pasteFromClipboard tableState then failwith "pasted from removed source identity"

// Opening a different workbook drops pending edit, drag, history and clipboard.
startEditing tableState 0 0 (Some "uncommitted")
tableState.DragAnchor.Value <- Some (2, 2)
replaceWorkbook tableState book
commitEdit tableState
eq "replacement not edited by pending input" (n 2.0) (value "source-17" 0 0 tableState)
eq "replacement drag cleared" None tableState.DragAnchor.Value
if not tableState.UndoStack.Value.IsEmpty then failwith "replacement history not cleared"

let dateBook =
    { book with Sheets = book.Sheets |> Map.map (fun _ sh -> { sh with Model = { sh.Model with Date1904 = true } }) }
let dateState = createStateFromWorkbook dateBook
insertSheet dateState "Date system" |> ignore
eq "new sheet inherits date system" true (activeModel dateState).Date1904
let freshTables = createState [| modelWithTable |] [| "Inventory" |] 0
eq "fresh model table registered" true (Map.containsKey "Stock" freshTables.Workbook.Value.Tables)
eq "invalid sheet name rejected" false (insertSheet freshTables "Bad/Name")
eq "duplicate sheet name rejected" false (insertSheet freshTables "Inventory")
eq "invalid rename rejected" false (renameSheet freshTables "sheet1" "Bad/Name")
printfn "widgets-editorhost: %d assertions passed" checks
