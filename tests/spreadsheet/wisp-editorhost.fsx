#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeUi.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Vb.dll"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../office-wisp/bin/Debug/net10.0/Rbt.OfficeWisp.dll"
// Resolve packages from the library build: dotnet fsi uses the SDK's Core,
// while the executable output carries Wisp's newer FSharp.Core runtime.
#r "../../office-wisp/bin/Debug/net10.0/Rbt.Wisp.dll"
#r "../../office-wisp/bin/Debug/net10.0/Rbt.Renderer.dll"
#r "../../office-wisp/spreadsheet-native/bin/Debug/net10.0/StbImageNext.dll"
#r "../../office-wisp/spreadsheet-native/bin/Debug/net10.0/Rbt.OfficeNative.dll"
#r "../../office-wisp/bin/Debug/net10.0/Platform.dll"

#r "../../office-wisp/spreadsheet-native/bin/Debug/net10.0/Rbt.SpreadsheetNative.dll"
#load "../../spreadsheet-wisp/EditorHost.fs"

open Rbt.Spreadsheet.Model
open Rbt.SpreadsheetWisp.EditorHost

let fail name detail =
    failwithf "%s failed: %s" name detail

let mutable checks = 0
let assertEq name expected actual =
    checks <- checks + 1
    if expected <> actual then fail name (sprintf "expected %A, got %A" expected actual)

let st = create [| createEmpty 20 10 |] [| "Sheet1" |] 0

applyCellEdit st 0 0 "2" |> ignore
applyCellEdit st 0 1 "=A1*5" |> ignore

match getCellValue { Row = 0; Col = 1 } (activeModel st) with
| Formula (_, Number 10.0) when cellEditText (activeModel st) 0 1 = "=A1*5" -> ()
| other -> fail "cell commit with formula" (sprintf "%A" other)


handleKey st { Key = "ArrowRight"; Shift = false; Ctrl = false; Alt = false } |> ignore
assertEq "arrow navigation" (0, 1) (focusCell (activeModel st))

handleKey st { Key = "Delete"; Shift = false; Ctrl = false; Alt = false } |> ignore
assertEq "delete-selection" Empty (getCellValue { Row = 0; Col = 1 } (activeModel st))

undo st |> ignore
match getCellValue { Row = 0; Col = 1 } (activeModel st) with
| Formula (_, Number 10.0) when cellEditText (activeModel st) 0 1 = "=A1*5" -> ()
| other -> fail "undo restores" (sprintf "%A" other)

redo st |> ignore
assertEq "redo delete" Empty (getCellValue { Row = 0; Col = 1 } (activeModel st))

let value st r c = getCellValue { Row = r; Col = c } (activeModel st)
let cached st r c = value st r c |> unwrapFormulaValue
let ok = function Result.Ok () -> () | Result.Error message -> failwith message
let key name = { Key = name; Shift = false; Ctrl = false; Alt = false }

let chain = create [| createEmpty 20 10; createEmpty 20 10; createEmpty 20 10 |] [| "Data"; "Middle"; "Result" |] 0
applyCellEdit chain 0 0 "4" |> ignore
switchSheet chain 1
applyCellEdit chain 0 0 "=Data!A1*2" |> ignore
switchSheet chain 2
applyCellEdit chain 0 0 "=Middle!A1+1" |> ignore
assertEq "cross-sheet chain" (Number 9.0) (cached chain 0 0)
switchSheet chain 0
handleKey chain (key "Delete") |> ignore
switchSheet chain 2
assertEq "delete precedent across sheets" (Number 1.0) (cached chain 0 0)
undo chain |> ignore
switchSheet chain 2
assertEq "undo precedent deletion" (Number 9.0) (cached chain 0 0)
redo chain |> ignore
switchSheet chain 2
assertEq "redo precedent deletion" (Number 1.0) (cached chain 0 0)

let ranged = create [| createEmpty 20 10 |] [| "Sheet1" |] 0
applyCellEdit ranged 0 0 "3" |> ignore
applyCellEdit ranged 1 0 "7" |> ignore
applyCellEdit ranged 0 1 "=SUM(A1:A2)" |> ignore
applyNoHistory ranged (fun m -> { m with Selection = Range(0,0,1,0) }) |> ignore
handleKey ranged (key "Delete") |> ignore
assertEq "range delete recalculates" (Number 0.0) (cached ranged 0 1)
undo ranged |> ignore
assertEq "range delete undo" (Number 10.0) (cached ranged 0 1)

let moved = create [| createEmpty 20 10; createEmpty 20 10 |] [| "Data"; "Other" |] 0
applyCellEdit moved 0 0 "8" |> ignore
switchSheet moved 1
applyCellEdit moved 0 1 "=Data!A1*2" |> ignore
let dataSheetId = moved.Workbook.Value.SheetOrder.[0]
let otherSheetId = moved.Workbook.Value.SheetOrder.[1]
assertEq "global move reference created" (Ok true)
    (executeName moved (Rbt.Spreadsheet.Session.SetDefinedName (None, "Target", "Data!$A$1", None)))
assertEq "local move reference created" (Ok true)
    (executeName moved (Rbt.Spreadsheet.Session.SetDefinedName (Some otherSheetId, "LocalTarget", "Data!$A$1", None)))
switchSheet moved 0
setFocus moved 0 0
cutSelection moved |> ignore
switchSheet moved 1
setFocus moved 0 0
pasteFromClipboard moved |> ignore
assertEq "cut paste dependent" (Number 16.0) (cached moved 0 1)
assertEq "cut repoints global and scoped names" ("Other!$A$1", "Other!$A$1")
    (moved.Workbook.Value.Names.[None, "TARGET"], moved.Workbook.Value.Names.[Some otherSheetId, "LOCALTARGET"])
undo moved |> ignore
assertEq "undo paste destination" Empty (value moved 0 0)
assertEq "undo move restores both names" ("Data!$A$1", "Data!$A$1")
    (moved.Workbook.Value.Names.[None, "TARGET"], moved.Workbook.Value.Names.[Some otherSheetId, "LOCALTARGET"])
undo moved |> ignore
assertEq "undo cut source" (Number 8.0) (value moved 0 0)
redo moved |> ignore
redo moved |> ignore
assertEq "redo cut paste" (Number 16.0) (cached moved 0 1)
assertEq "redo move repoints name once" "Other!$A$1" moved.Workbook.Value.Names.[None, "TARGET"]
setFocus moved 2 2
assertEq "consumed cut cannot move again" false (pasteFromClipboard moved)
assertEq "consumed cut leaves name and destination intact" ("Other!$A$1", Empty)
    (moved.Workbook.Value.Names.[None, "TARGET"], value moved 2 2)

// Copy is transient; the rich block survives sheet switches and tiles over a selection.
let copied = create [| createEmpty 8 8; createEmpty 8 8 |] [| "Data"; "Output" |] 0
applyCellEdit copied 0 0 "4" |> ignore
applyCellEdit copied 0 1 "=Data!$A$1*2" |> ignore
applyChange copied (fun m ->
    let addr = { Row = 0; Col = 1 }
    let old = getCell addr m |> Option.get
    setCell addr { old with Format = { old.Format with Bold = true } } m) |> ignore
let copyRevision = revision copied
let copyBook = copied.Workbook.Value
let savedCopy = captureSave copied
markSaved copied savedCopy
setFocus copied 0 1
assertEq "copy captures formula" true (copySelection copied)
assertEq "copy preserves revision and saved state" (copyRevision, false)
    (revision copied, isModified copied)
switchSheet copied 1
applyNoHistory copied (fun m -> { m with Selection = Range (0, 1, 1, 2) }) |> ignore
assertEq "rich copy pastes" true (pasteFromClipboard copied)
assertEq "tiled cross-sheet formula" (Number 8.0, Number 8.0) (cached copied 0 1, cached copied 1 2)
assertEq "copied cell format" true ((getCell { Row = 1; Col = 2 } (activeModel copied) |> Option.get).Format.Bold)
assertEq "copy source remains" (Number 8.0) (copied.Workbook.Value.Sheets.[copyBook.ActiveSheet].Model |> getCellValue { Row = 0; Col = 1 } |> unwrapFormulaValue)
undo copied |> ignore
assertEq "undo copy paste is clean" (Empty, false) (value copied 1 2, isModified copied)
let shiftedCopy = create [| createEmpty 8 8; createEmpty 8 8 |] [| "Data"; "Output" |] 0
applyCellEdit shiftedCopy 0 0 "2" |> ignore
applyCellEdit shiftedCopy 0 1 "=A1*5" |> ignore
setFocus shiftedCopy 0 1
copySelection shiftedCopy |> ignore
switchSheet shiftedCopy 1
applyCellEdit shiftedCopy 1 0 "7" |> ignore
setFocus shiftedCopy 1 1
assertEq "relative formula copies across sheets" true (pasteFromClipboard shiftedCopy)
assertEq "copied formula shifts its relative reference" ("=A2*5", Number 35.0)
    (cellEditText (activeModel shiftedCopy) 1 1, cached shiftedCopy 1 1)
undo shiftedCopy |> ignore
assertEq "undo shifted copy clears destination" Empty (value shiftedCopy 1 1)
redo shiftedCopy |> ignore
assertEq "redo shifted copy recalculates destination" (Number 35.0) (cached shiftedCopy 1 1)

// A cut's deletion and its subsequent cross-sheet move are independent history entries.
let renamedCut = create [| createEmpty 8 8; createEmpty 8 8 |] [| "Data"; "Output" |] 0
applyCellEdit renamedCut 0 0 "6" |> ignore
applyCellEdit renamedCut 0 1 "=Data!A1+1" |> ignore
setFocus renamedCut 0 1
cutSelection renamedCut |> ignore
renameSheet renamedCut "Inputs" |> ok
switchSheet renamedCut 1
setFocus renamedCut 2 2
assertEq "renamed clipboard formula pastes" true (pasteFromClipboard renamedCut)
assertEq "formula follows source rename" "=Inputs!A1+1" (cellEditText (activeModel renamedCut) 2 2)
assertEq "cut pastes only once without OS text" false (pasteFromClipboard renamedCut)
undo renamedCut |> ignore
assertEq "undo move restores destination" Empty (value renamedCut 2 2)
undo renamedCut |> ignore
assertEq "undo rename returns source name" "Data" renamedCut.Workbook.Value.Sheets.[renamedCut.Workbook.Value.SheetOrder.[0]].Name
undo renamedCut |> ignore
switchSheet renamedCut 0
assertEq "undo cut restores original formula" "=Data!A1+1" (cellEditText (activeModel renamedCut) 0 1)
assertEq "undo cut cancels pending move" None renamedCut.Session.Clipboard

let emptyCopy = create [| createEmpty 4 4 |] [| "Data" |] 0
let emptyBook = emptyCopy.Workbook.Value
copySelection emptyCopy |> ignore
assertEq "empty copy paste is a no-op" false (pasteFromClipboard emptyCopy)
assertEq "no-op clipboard keeps save and workbook identity" (true, 0, false)
    (obj.ReferenceEquals(emptyBook, emptyCopy.Workbook.Value), revision emptyCopy, isModified emptyCopy)

let spills = create [| createEmpty 20 10; createEmpty 20 10 |] [| "Data"; "Spills" |] 0
applyCellEdit spills 0 0 "2" |> ignore
switchSheet spills 1
applyCellEdit spills 0 0 "=SEQUENCE(Data!A1)" |> ignore
assertEq "spill initially" (Number 2.0) (cached spills 1 0)
switchSheet spills 0
applyCellEdit spills 0 0 "3" |> ignore
switchSheet spills 1
assertEq "spill refresh" (Number 3.0) (cached spills 2 0)
switchSheet spills 0
applyCellEdit spills 0 0 "1" |> ignore
switchSheet spills 1
assertEq "spill shrink clears orphan" Empty (value spills 2 0)
undo spills |> ignore
switchSheet spills 1
assertEq "undo spill shrink restores orphan" (Number 3.0) (cached spills 2 0)
redo spills |> ignore
switchSheet spills 1
assertEq "redo spill shrink clears orphan" Empty (value spills 2 0)

module W = Rbt.Spreadsheet.Workbook
module B = Rbt.Spreadsheet.XlsxBridge
let book = create [| createEmpty 20 10; createEmpty 20 10 |] [| "Data"; "Local" |] 0
applyCellEdit book 1 0 "12" |> ignore
applyChange book (createTable "Sales" 3 0 5 1 false) |> ignore
let initial = book.Workbook.Value
let freshTables = create [| activeModel book |] [| "Inventory" |] 0
assertEq "fresh workbook registers model tables" true (Map.containsKey "Sales" freshTables.Workbook.Value.Tables)
let localId = initial.SheetOrder.[1]
let info : W.DefinedNameInfo = { Name = "LocalName"; Comment = Some "note"; Hidden = true }
let enriched =
    { initial with
        Names = Map.ofList [ (None, "GLOBAL"), "Data!$A$2"; (Some localId, "LOCALNAME"), "Data!$A$2" ]
        NameInfo = Map.ofList [ (Some localId, "LOCALNAME"), info ]
        Iteration = Some { MaxIterations = 40; MaxChange = 0.01 }
        Tables = initial.Tables |> Map.map (fun _ t -> { t with Columns = [ "Declared"; "Other" ] }) }
replaceWorkbook book enriched
applyChange book (createTable "Extra" 8 0 10 1 false) |> ignore
assertEq "declared table columns retained" [ "Declared"; "Other" ] book.Workbook.Value.Tables.["Sales"].Columns
assertEq "new table registered" true (Map.containsKey "Extra" book.Workbook.Value.Tables)
applyChange book (fun m -> { m with Tables = m.Tables |> List.filter (fun t -> t.Name <> "Extra") }) |> ignore
assertEq "removed table unregistered" false (Map.containsKey "Extra" book.Workbook.Value.Tables)
switchSheet book 1
applyCellEdit book 0 0 "=Data!A2*2" |> ignore
switchSheet book 0
applyStructureEdit book (InsertRows(0, 1)) |> ignore
let localFormula () = book.Workbook.Value.Sheets.[localId].Model |> fun model -> cellEditText model 0 0
let localValue () = book.Workbook.Value.Sheets.[localId].Model |> getCellValue { Row = 0; Col = 0 } |> unwrapFormulaValue
assertEq "structural edit rewrites other-sheet formula" "=Data!A3*2" (localFormula ())
assertEq "structural edit recalculates other-sheet formula" (Number 24.0) (localValue ())
undo book |> ignore
assertEq "undo structure restores other-sheet formula" "=Data!A2*2" (localFormula ())
assertEq "undo structure restores other-sheet value" (Number 24.0) (localValue ())
redo book |> ignore
assertEq "redo structure rewrites other-sheet formula" "=Data!A3*2" (localFormula ())
assertEq "global name follows structure" "Data!$A$3" book.Workbook.Value.Names.[None, "GLOBAL"]
assertEq "local name follows structure" "Data!$A$3" book.Workbook.Value.Names.[Some localId, "LOCALNAME"]
undo book |> ignore
assertEq "undo restores names" "Data!$A$2" book.Workbook.Value.Names.[None, "GLOBAL"]
redo book |> ignore
switchSheet book 1
renameSheet book "Renamed" |> ok
addSheet book "Added" |> ok
let addedId = book.Workbook.Value.ActiveSheet
deleteSheet book |> ok
undo book |> ignore
assertEq "undo delete restores identity" addedId book.Workbook.Value.ActiveSheet
undo book |> ignore
assertEq "undo add restores active" localId book.Workbook.Value.ActiveSheet
undo book |> ignore
assertEq "undo rename" "Local" (activeSheetName book)
redo book |> ignore
redo book |> ignore
redo book |> ignore
assertEq "redo deletion order" [ "Data"; "Renamed" ] (book.Workbook.Value.SheetOrder |> List.map (fun id -> book.Workbook.Value.Sheets.[id].Name))
let reordered = create [| createEmpty 4 4; createEmpty 4 4 |] [| "First"; "Second" |] 0
let firstId, secondId = reordered.Workbook.Value.SheetOrder.[0], reordered.Workbook.Value.SheetOrder.[1]
assertEq "sheet reorder changes order" (Ok true) (executeStructural reordered (Rbt.Spreadsheet.Session.MoveSheet (secondId, 0)))
assertEq "sheet reorder retains stable identities" [ secondId; firstId ] reordered.Workbook.Value.SheetOrder
undo reordered |> ignore
assertEq "undo sheet reorder restores identities" [ firstId; secondId ] reordered.Workbook.Value.SheetOrder
let bytes = XlsxWriter.writeWorkbookPackageBytes (exportWorkbook book None)
let imported = bytes |> XlsxReader.loadBytesNative |> Option.get |> B.Import.importWorkbook
assertEq "saved iteration" book.Workbook.Value.Iteration imported.Context.Iteration
assertEq "saved global name" "Data!$A$3" imported.Context.Names.[None, "GLOBAL"]
assertEq "saved scoped name" "Data!$A$3" imported.Context.Names.[Some "sheet2", "LOCALNAME"]
assertEq "saved name metadata" info imported.Context.NameInfo.[Some "sheet2", "LOCALNAME"]
assertEq "saved table" true (Map.containsKey "Sales" imported.Context.Tables)

let reopened = createWithWorkbook imported.Context
switchSheet reopened 0
deleteSheet reopened |> ok
let removed = reopened.Workbook.Value
assertEq "surviving stable id" "sheet2" removed.ActiveSheet
let savedAgain = XlsxWriter.writeWorkbookPackageBytes (exportWorkbook reopened (Some imported.Source))
let reopenedAgain = savedAgain |> XlsxReader.loadBytesNative |> Option.get |> B.Import.importWorkbook
assertEq "local scope follows sheet deletion" true (Map.containsKey (Some "sheet1", "LOCALNAME") reopenedAgain.Context.Names)
assertEq "local metadata follows scope" info reopenedAgain.Context.NameInfo.[Some "sheet1", "LOCALNAME"]
undo reopened |> ignore
assertEq "undo restores sheet identity" [ "sheet1"; "sheet2" ] reopened.Workbook.Value.SheetOrder
redo reopened |> ignore
assertEq "redo preserves survivor identity" "sheet2" reopened.Workbook.Value.ActiveSheet

let references = create [| createEmpty 20 10; createEmpty 20 10 |] [| "Data"; "Results" |] 0
applyCellEdit references 0 0 "5" |> ignore
switchSheet references 1
applyCellEdit references 0 0 "=Data!A1*3" |> ignore
switchSheet references 0
renameSheet references "New Data" |> ok
switchSheet references 1
assertEq "rename rewrites dependent formula" "='New Data'!A1*3" (cellEditText (activeModel references) 0 0)
assertEq "rename retains dependent value" (Number 15.0) (cached references 0 0)
let savedReferences = exportWorkbook references None |> XlsxWriter.writeWorkbookPackageBytes |> XlsxReader.loadBytesNative |> Option.get
let savedResults = savedReferences |> B.Import.importWorkbook |> _.Context |> createWithWorkbook
switchSheet savedResults 1
assertEq "saved formula cache after rename" (Number 15.0) (cached savedResults 0 0)
switchSheet references 0
setFocus references 0 0
copySelection references |> ignore
deleteSheet references |> ok
assertEq "deletion invalidates dependent" (Error "REF!") (cached references 0 0)
assertEq "deleting clipboard source invalidates rich copy" None references.Session.Clipboard
assertEq "removed source cannot be pasted" false (pasteFromClipboard references)
undo references |> ignore
switchSheet references 1
assertEq "undo deletion restores dependent" (Number 15.0) (cached references 0 0)
assertEq "undo deletion does not revive stale clipboard" None references.Session.Clipboard
redo references |> ignore
assertEq "redo deletion invalidates dependent" (Error "REF!") (cached references 0 0)
addSheet references "Fresh" |> ok
assertEq "deleted identity not reused" false (references.Workbook.Value.ActiveSheet = "sheet1")
assertEq "duplicate sheet name rejected" true (renameSheet references "Results" |> Result.isError)
assertEq "invalid sheet name rejected" true (addSheet references "Bad/Name" |> Result.isError)

let events = create [| createEmpty 20 10; createEmpty 20 10 |] [| "Data"; "Results" |] 0
applyCellEdit events 0 0 "2" |> ignore
applyCellEdit events 0 1 "=A1*5" |> ignore
switchSheet events 1
applyCellEdit events 0 0 "=Data!C1*2" |> ignore
switchSheet events 0
applyChange events (fun m -> { m with Macros = "Sub Worksheet_Change(ByVal Target As Range)\nRange(\"C1\").Value = Range(\"B1\").Value\nEnd Sub" }) |> ignore
applyCellEdit events 0 0 "4" |> ignore
assertEq "event sees recalculated dependent" (Number 20.0) (cached events 0 2)
switchSheet events 1
assertEq "event writes recalculate foreign dependent" (Number 40.0) (cached events 0 0)
undo events |> ignore
assertEq "event and edit undo together" (Number 2.0) (cached events 0 0)
assertEq "event output undone" Empty (value events 0 2)
copySelection events |> ignore
startEditing events 0 0 (Some "99")
let replacement = W.createWorkbook () |> W.withDate1904 true
replaceWorkbook events replacement
assertEq "replacement cancels pending entry" false (commitEdit events)
assertEq "replacement clears clipboard" false (pasteFromClipboard events)
addSheet events "New" |> ok
assertEq "new sheet keeps workbook date system" true (activeModel events).Date1904

let lifecycle = create [| createEmpty 20 10 |] [| "Sheet1" |] 0
let initialRevision = revision lifecycle
setFocus lifecycle 2 2
assertEq "selection remains clean" false (isModified lifecycle)
assertEq "selection keeps revision" initialRevision (revision lifecycle)
applyCellEdit lifecycle 0 0 "6" |> ignore
let captured = captureSave lifecycle
applyCellEdit lifecycle 0 1 "=A1*2" |> ignore
markSaved lifecycle captured
assertEq "stale save remains dirty" true (isModified lifecycle)
addSheet lifecycle "Second" |> ok
undo lifecycle |> ignore
assertEq "undo mixed sheet operation" true (canRedo lifecycle)
undo lifecycle |> ignore
assertEq "undo returns to saved revision" false (isModified lifecycle)
redo lifecycle |> ignore
assertEq "redo is dirty" true (isModified lifecycle)
replaceWorkbook lifecycle (W.createWorkbook ())
assertEq "replacement clean" false (isModified lifecycle)
assertEq "replacement clears history" false (canUndo lifecycle)
applyCellEdit lifecycle 0 0 "9" |> ignore
markSaved lifecycle captured
assertEq "old save cannot mark replacement clean" true (isModified lifecycle)

// A document setting is an edit even before any cell exists. Undo must restore
// the date interpretation, not merely the visible cell map.
let dateHistory = create [| createEmpty 20 10 |] [| "Sheet1" |] 0
applyWorkbookChange dateHistory (W.withDate1904 true) |> ignore
assertEq "date-system-only transaction is dirty" true (isModified dateHistory)
undo dateHistory |> ignore
assertEq "undo date-system-only transaction is clean" false (isModified dateHistory)
applyCellEdit dateHistory 0 0 "1904-01-01" |> ignore
assertEq "date entry uses restored date system" (Number 1462.0) (cached dateHistory 0 0)

module S = Rbt.Spreadsheet.Session
// The shared text-paste command retains quoted TSV and selection/merge semantics;
// Wisp itself deliberately uses its internal rich clipboard instead of an OS adapter.
let textPasteModel = { createEmpty 10 8 with MergedCells = [ { StartRow = 3; StartCol = 1; EndRow = 3; EndCol = 2 } ] }
let textPaste = create [| textPasteModel |] [| "Data" |] 0
let pastedCell r c =
    textPaste.Session.Workbook.Sheets.[textPaste.Session.Workbook.ActiveSheet].Model
    |> getCellValue { Row = r; Col = c }
let selectTextPaste selection =
    let wb = textPaste.Session.Workbook
    let model = wb.Sheets.[wb.ActiveSheet].Model
    textPaste.Session.UpdateView(W.updateSheetModel wb.ActiveSheet { model with Selection = selection } wb)
selectTextPaste (SingleCell (0, 0))
assertEq "quoted multiline TSV pastes" (Ok true)
    (textPaste.Session.PasteRows(textPaste.Session.Workbook.ActiveSheet, S.clipboardRowsOfText "\"line1\nline2\"\t5\r\n\"a\"\"b\"\t6\r\n", 2026))
assertEq "quoted TSV values and numeric neighbors" (String "line1\nline2", String "a\"b", Number 5.0, Number 6.0)
    (pastedCell 0 0, pastedCell 1 0, pastedCell 0 1, pastedCell 1 1)
textPaste.Session.Undo() |> ignore
assertEq "undo text paste restores cells" (Empty, Empty) (pastedCell 0 0, pastedCell 1 1)
selectTextPaste (Range (5, 1, 6, 2))
assertEq "scalar pastes into selected range" (Ok true)
    (textPaste.Session.PasteRows(textPaste.Session.Workbook.ActiveSheet, [| [| "9" |] |], 2026))
assertEq "scalar tiles across selection" [ Number 9.0; Number 9.0; Number 9.0; Number 9.0 ]
    [ for row in 5 .. 6 do for col in 1 .. 2 -> pastedCell row col ]
selectTextPaste (SingleCell (3, 1))
assertEq "text paste into merge succeeds" (Ok true)
    (textPaste.Session.PasteRows(textPaste.Session.Workbook.ActiveSheet, [| [| "15"; "16" |] |], 2026))
assertEq "merged secondary cell stays empty" (Number 15.0, Empty) (pastedCell 3 1, pastedCell 3 2)
let guarded = create [| createEmpty 20 10 |] [| "Data" |] 0
let guardedId = guarded.Workbook.Value.ActiveSheet
applyCellEdit guarded 0 0 "1" |> ignore
undo guarded |> ignore
let guardedBook = guarded.Workbook.Value
let guardedRevision = revision guarded
assertEq "redo available before no-op" true (canRedo guarded)
renameSheet guarded "Data" |> ok
assertEq "same-name rename leaves redo intact" true (canRedo guarded)
assertEq "same-name rename leaves save clean" false (isModified guarded)
assertEq "same-name rename leaves revision intact" guardedRevision (revision guarded)
assertEq "same-name rename keeps workbook identity" true (obj.ReferenceEquals(guardedBook, guarded.Workbook.Value))
assertEq "same-index move is no-op" (Ok false) (executeStructural guarded (S.MoveSheet (guardedId, 0)))
assertEq "invalid row edit rejected" true (executeStructural guarded (S.EditStructure (guardedId, InsertRows (-1, 1))) |> Result.isError)
assertEq "invalid sheet identity rejected" true (executeStructural guarded (S.DeleteSheet "missing") |> Result.isError)
assertEq "invalid commands leave workbook identity" true (obj.ReferenceEquals(guardedBook, guarded.Workbook.Value))
assertEq "invalid commands preserve redo" true (canRedo guarded)
assertEq "invalid commands preserve save identity" false (isModified guarded)
redo guarded |> ignore
assertEq "redo after rejected structure" (Number 1.0) (cached guarded 0 0)

// The focused corner controls a mixed-range toggle; filtered rows stay untouched.
let addr r c = { Row = r; Col = c }
let cell value format = { Value = value; Format = format; Comment = None }
let formattedModel =
    createEmpty 4 4
    |> setCell (addr 0 0) (cell (String "Group") defaultFormat)
    |> setCell (addr 0 1) (cell (Number 3.0) { defaultFormat with Bold = true })
    |> setCell (addr 1 0) (cell (String "West") defaultFormat)
    |> setCell (addr 2 0) (cell (String "East") defaultFormat)
    |> setCell (addr 2 2) (cell (Formula ("RAND()", Number 0.314159)) defaultFormat)
    |> setAutoFilter (Some { From = addr 0 0; To = addr 2 2 })
    |> setColumnFilter 0 (FilterValues (Set.ofList [ "East" ], false))
let formatBook =
    let baseBook = (create [| formattedModel |] [| "Data" |] 0).Workbook.Value
    W.setDefinedName None "Tag" "$A$1" (Some "kept") baseBook
let formatting = createWithWorkbook formatBook
let formatAt r c =
    getCell (addr r c) (activeModel formatting)
    |> Option.map (fun data -> data.Format)
    |> Option.defaultValue defaultFormat
assertEq "filter hides only West" [ 1 ] (formatting |> activeModel |> fun m -> m.FilterHiddenRows |> List.ofArray)
let filteredCopy = create [| { formattedModel with RowCount = 6 } |] [| "Data" |] 0
applyNoHistory filteredCopy (fun m -> { m with Selection = Range (0, 0, 2, 0) }) |> ignore
assertEq "filtered source copied" true (copySelection filteredCopy)
setFocus filteredCopy 3 0
assertEq "filtered source pastes" true (pasteFromClipboard filteredCopy)
assertEq "filtered copy compacts visible rows" (String "Group", String "East", Empty)
    (value filteredCopy 3 0, value filteredCopy 4 0, value filteredCopy 5 0)
applyNoHistory formatting (fun m -> { m with Selection = Range (2, 2, 0, 0) }) |> ignore
let cachedBefore = getCellValue (addr 2 2) (activeModel formatting)
assertEq "reversed mixed-range bold" (Ok true) (executeFormatting formatting S.ToggleBold)
assertEq "focus toggles entire visible range on" true (formatAt 0 0).Bold
assertEq "existing bold remains on" true (formatAt 0 1).Bold
assertEq "filtered row keeps original format" false (formatAt 1 0).Bold
assertEq "formatting leaves formula cache intact" cachedBefore (getCellValue (addr 2 2) (activeModel formatting))
assertEq "formatting leaves name metadata intact" formatBook.NameInfo formatting.Workbook.Value.NameInfo
assertEq "formatting records one revision" 1 (revision formatting)
undo formatting |> ignore
assertEq "one undo restores mixed selection" (false, true) ((formatAt 0 0).Bold, (formatAt 0 1).Bold)
redo formatting |> ignore
assertEq "redo reapplies range" true (formatAt 2 2).Bold
applyNoHistory formatting (fun m -> { m with Selection = Range (0, 0, 0, 1) }) |> ignore
assertEq "merge through shared journal" (Ok true) (executeFormatting formatting S.MergeCells)
assertEq "merge retains metadata" formatBook.NameInfo formatting.Workbook.Value.NameInfo
assertEq "merge creates one range" 1 (activeModel formatting).MergedCells.Length
assertEq "unmerge uses shared journal" (Ok true) (executeFormatting formatting (S.UnmergeCells (addr 0 1)))
assertEq "unmerge removes range" 0 (activeModel formatting).MergedCells.Length
undo formatting |> ignore
assertEq "undo unmerge restores range" 1 (activeModel formatting).MergedCells.Length
undo formatting |> ignore
assertEq "undo merge alone" 0 (activeModel formatting).MergedCells.Length
assertEq "undo merge preserves formatting" true (formatAt 0 0).Bold
assertEq "top border changes selection" (Ok true) (executeFormatting formatting (S.Border ("top", "thin", "#000000")))
assertEq "top border leaves other edges alone" (Some "thin", None) ((formatAt 0 1).BorderTop |> Option.map (fun b -> b.Style), (formatAt 0 1).BorderBottom)
assertEq "format-only border keeps formula cache" cachedBefore (getCellValue (addr 2 2) (activeModel formatting))
assertEq "clear formats uses shared selection" (Ok true) (executeFormatting formatting (S.ClearCells Rbt.Spreadsheet.SheetOps.ClearFormats))
assertEq "clear formats retains value" (Number 3.0) (getCellValue (addr 0 1) (activeModel formatting))
assertEq "clear formats removes bold and border" (false, None) ((formatAt 0 1).Bold, (formatAt 0 1).BorderTop)
undo formatting |> ignore
assertEq "undo clear restores formatting atomically" (true, Some "thin") ((formatAt 0 1).Bold, (formatAt 0 1).BorderTop |> Option.map (fun b -> b.Style))


let noOpFormatting = create [| createEmpty 3 3 |] [| "Data" |] 0
applyCellEdit noOpFormatting 0 0 "3" |> ignore
undo noOpFormatting |> ignore
let cleanBook = noOpFormatting.Workbook.Value
let cleanSheets = noOpFormatting.Sheets.Value
let cleanRevision = revision noOpFormatting
assertEq "no-op format returns unchanged" (Ok false) (executeFormatting noOpFormatting (S.HorizontalAlign None))
assertEq "no-op keeps workbook projection" true (obj.ReferenceEquals(cleanBook, noOpFormatting.Workbook.Value) && obj.ReferenceEquals(cleanSheets, noOpFormatting.Sheets.Value))
assertEq "no-op keeps revision" cleanRevision (revision noOpFormatting)
assertEq "no-op keeps saved state and redo" (false, true) (isModified noOpFormatting, canRedo noOpFormatting)
assertEq "invalid formatting rejected" true (executeFormatting noOpFormatting (S.NumberFormat (Some "0[broken")) |> Result.isError)
assertEq "invalid format leaves redo and revision" (true, cleanRevision) (canRedo noOpFormatting, revision noOpFormatting)
redo noOpFormatting |> ignore
assertEq "redo survives no-op formatting" (Number 3.0) (cached noOpFormatting 0 0)

// Name transactions recalculate both owners of a shadowed name.
let named = create [| createEmpty 8 4; createEmpty 8 4 |] [| "Data"; "Report" |] 0
let dataId = named.Workbook.Value.SheetOrder.[0]
let reportId = named.Workbook.Value.SheetOrder.[1]
applyCellEdit named 0 0 "2" |> ignore
applyCellEdit named 1 0 "7" |> ignore
assertEq "name creates global entry" (Ok true) (executeName named (S.SetDefinedName (None, "Rate", "Data!$A$1", Some "global")))
applyCellEdit named 0 1 "=Rate*3" |> ignore
switchSheet named 1
applyCellEdit named 0 1 "=Rate*3" |> ignore
let globalSave = captureSave named
markSaved named globalSave
let valueIn id =
    named.Workbook.Value.Sheets.[id].Model
    |> getCellValue (addr 0 1)
    |> unwrapFormulaValue
assertEq "global name evaluates on source and foreign sheets" (Number 6.0, Number 6.0) (valueIn dataId, valueIn reportId)
assertEq "local name shadows global" (Ok true) (executeName named (S.SetDefinedName (Some reportId, "Rate", "Data!$A$2", Some "local")))
assertEq "scoped name recalculates only its owner" (Number 6.0, Number 21.0) (valueIn dataId, valueIn reportId)
assertEq "shadow metadata retained" (Some "local") named.Workbook.Value.NameInfo.[Some reportId, "RATE"].Comment
undo named |> ignore
assertEq "undo local name recalculates foreign formula" (Number 6.0) (valueIn reportId)
assertEq "undo name returns to saved revision" false (isModified named)
let beforeNameNoOp = named.Workbook.Value
let beforeNameRevision = revision named
assertEq "same global name is no-op" (Ok false) (executeName named (S.SetDefinedName (None, "Rate", "Data!$A$1", Some "global")))
assertEq "missing scope rejects name" true (executeName named (S.SetDefinedName (Some "missing", "Rate", "Data!$A$2", None)) |> Result.isError)
assertEq "no-op and reject do not publish" true (obj.ReferenceEquals(beforeNameNoOp, named.Workbook.Value))
assertEq "no-op and reject preserve revision and redo" (beforeNameRevision, true) (revision named, canRedo named)
assertEq "no-op and reject preserve saved state" false (isModified named)
redo named |> ignore
assertEq "redo local shadow restores cached result" (Number 21.0) (valueIn reportId)
assertEq "delete shadow exposes global" (Ok true) (executeName named (S.DeleteDefinedName (Some reportId, "Rate")))
assertEq "delete shadow recalculates foreign sheet" (Number 6.0) (valueIn reportId)
undo named |> ignore
assertEq "undo shadow deletion restores result" (Number 21.0) (valueIn reportId)
redo named |> ignore
assertEq "redo shadow deletion restores global result" (Number 6.0) (valueIn reportId)
replaceWorkbook named (W.createWorkbook ())
assertEq "replacement removes named workbook history" (false, false) (canUndo named, isModified named)
assertEq "replacement removes old name scopes" Map.empty named.Workbook.Value.Names

// Rule edits preserve ownership, stable IDs, partial coverage and the journal.
let ruleModel = createEmpty 8 5 |> setCell (addr 2 2) (cell (Formula ("RAND()", Number 0.314159)) defaultFormat)
let rules = create [| ruleModel; createEmpty 8 5 |] [| "Data"; "Other" |] 0
let area r1 c1 r2 c2 : CellRange = { From = addr r1 c1; To = addr r2 c2 }
let ruleId (command: Result<S.RuleCommandResult, string>) =
    match command with
    | Ok result when result.Changed && result.Id.IsSome -> result.Id.Value
    | other -> failwithf "expected rule id, got %A" other
let validationAt row col = getValidation (addr row col) (activeModel rules) |> Option.map (fun entry -> entry.Rule)
let whole = validationOf (WholeNumber (Some 1.0, Some 10.0))
let choices = validationOf (ListValues [ "Yes"; "No" ])
let initialCache = value rules 2 2
let wholeId = executeRule rules (S.SetValidation ([ area 0 0 2 0 ], whole)) |> ruleId
let choicesId = executeRule rules (S.SetValidation ([ area 1 0 1 1 ], choices)) |> ruleId
assertEq "overlapping validation replaces only covered cells"
    (Some whole.Rule, Some choices.Rule, Some whole.Rule, Some choices.Rule)
    (validationAt 0 0, validationAt 1 0, validationAt 2 0, validationAt 1 1)
assertEq "partial validation clear changes rule" (Ok true)
    (executeRule rules (S.ClearValidation [ area 1 0 1 0 ]) |> Result.map (fun result -> result.Changed))
assertEq "partial clear retains other ranges and stable IDs"
    (Some whole.Rule, None, Some whole.Rule, Some choices.Rule, true, true)
    (validationAt 0 0, validationAt 1 0, validationAt 2 0, validationAt 1 1,
     (activeModel rules).Validations |> List.exists (fun entry -> entry.Id = wholeId),
     (activeModel rules).Validations |> List.exists (fun entry -> entry.Id = choicesId))
undo rules |> ignore
assertEq "undo partial validation clear" (Some choices.Rule) (validationAt 1 0)
redo rules |> ignore
let nextValidationId = executeRule rules (S.SetValidation ([ area 3 0 3 0 ], whole)) |> ruleId
assertEq "validation IDs are not recycled after partial clear" true (nextValidationId > choicesId)
switchSheet rules 1
assertEq "validations belong only to source sheet" [] (activeModel rules).Validations
switchSheet rules 0
let red : ConditionalFormat = { Rule = GreaterThan 3.0; Style = { defaultFormat with BgColor = Some "#FF0000" } }
let blue : ConditionalFormat = { Rule = LessThan 5.0; Style = { defaultFormat with BgColor = Some "#0000FF" } }
let redId = executeRule rules (S.AddConditionalFormat ([ area 0 2 2 2 ], false, red)) |> ruleId
let blueId = executeRule rules (S.AddConditionalFormat ([ area 0 2 2 2 ], true, blue)) |> ruleId
let ruleOrder () =
    (activeModel rules).ConditionalFormats |> List.sortBy (fun entry -> entry.Priority) |> List.map (fun entry -> entry.Id)
assertEq "new rule takes priority" [ blueId; redId ] (ruleOrder ())
undo rules |> ignore
let beforeRuleNoOp = rules.Workbook.Value
let savedRuleRevision = revision rules
markSaved rules (captureSave rules)
assertEq "clear outside conditional ranges is a no-op" (Ok false)
    (executeRule rules (S.ClearConditionalFormats (Some [ area 7 4 7 4 ])) |> Result.map (fun result -> result.Changed))
assertEq "unknown conditional ID rejected" true (executeRule rules (S.RemoveConditionalFormat 9999) |> Result.isError)
assertEq "rejected rule leaves projection, revision, saved state and redo untouched"
    (true, savedRuleRevision, false, true)
    (obj.ReferenceEquals(beforeRuleNoOp, rules.Workbook.Value), revision rules, isModified rules, canRedo rules)
redo rules |> ignore
assertEq "redo restores same conditional ID and order" [ blueId; redId ] (ruleOrder ())
assertEq "reorder existing rules" (Ok true)
    (executeRule rules (S.ReorderConditionalFormats [ redId; blueId ]) |> Result.map (fun result -> result.Changed))
assertEq "reorder preserves IDs and contiguous priorities" [ (redId, 1); (blueId, 2) ]
    ((activeModel rules).ConditionalFormats |> List.sortBy (fun entry -> entry.Priority) |> List.map (fun entry -> entry.Id, entry.Priority))
assertEq "partial conditional clear affects both rules" (Ok 2)
    (executeRule rules (S.ClearConditionalFormats (Some [ area 1 2 1 2 ])) |> Result.map (fun result -> result.Affected))
assertEq "partial clear keeps outer cells on original IDs" [ redId; blueId ]
    ((activeModel rules).ConditionalFormats
     |> List.filter (fun entry ->
         entry.Ranges |> List.exists (fun range -> range.From.Row = 0 && range.To.Row = 0))
     |> List.map (fun entry -> entry.Id) |> List.sort)
undo rules |> ignore
assertEq "undo partial clear restores middle cells" true
    ((activeModel rules).ConditionalFormats |> List.forall (fun entry ->
        entry.Ranges |> List.exists (fun range -> range.From.Row <= 1 && range.To.Row >= 1)))
assertEq "conditional removal changes rule" (Ok true)
    (executeRule rules (S.RemoveConditionalFormat redId) |> Result.map (fun result -> result.Changed))
assertEq "removal leaves other identity" [ blueId ] (ruleOrder ())
undo rules |> ignore
assertEq "undo removal restores same identity" [ redId; blueId ] (ruleOrder ())
redo rules |> ignore
let replacementRuleId = executeRule rules (S.AddConditionalFormat ([ area 0 2 2 2 ], false, red)) |> ruleId
assertEq "removed rule ID is not recycled" true (replacementRuleId > blueId)
assertEq "rule changes do not recalculate volatile cells" initialCache (value rules 2 2)
replaceWorkbook rules (W.createWorkbook ())
assertEq "replacement discards rule history and metadata" (false, [], [])
    (canUndo rules, (activeModel rules).ConditionalFormats, (activeModel rules).Validations)

// Drawing commands retain opaque imported objects and share history with cells.
let pictureBytes = [| 0x47uy; 0x49uy; 0x46uy; 0x38uy; 0x39uy; 0x61uy; 2uy; 0uy; 3uy; 0uy |]
let opaque = { Id = 4; Anchor = Rbt.Spreadsheet.SheetOps.pictureAnchor (addr 0 0) 0.0 0.0 20.0 20.0
               Name = Some "Imported"; Description = None; Content = KeptObject ("shape", "<x r:id=\"rIdChart1\"/>") }
let drawingState = create [| { (createEmpty 8 4) with Drawings = [ opaque ] } |] [| "Data" |] 0
let inserted = executeDrawing drawingState (S.InsertPicture (pictureBytes, addr 1 1, 20.0, 30.0, None)) |> Result.defaultWith failwith
assertEq "drawing id follows imported ids" (Some 5) inserted.Id
let savedDrawing = captureSave drawingState
markSaved drawingState savedDrawing
executeDrawing drawingState (S.PlacePicture (5, addr 2 1, 0.0, 0.0, 40.0, 30.0)) |> Result.defaultWith failwith |> ignore
undo drawingState |> ignore
let beforeDrawingNoOp = drawingState.Workbook.Value
let drawingRevision = revision drawingState
assertEq "identical picture placement is a no-op" false ((executeDrawing drawingState (S.PlacePicture (5, addr 1 1, 0.0, 0.0, 20.0, 30.0)) |> Result.defaultWith failwith).Changed)
assertEq "opaque object cannot be deleted as picture" true (executeDrawing drawingState (S.DeletePicture 4) |> Result.isError)
assertEq "invalid geometry rejected" true (executeDrawing drawingState (S.PlacePicture (5, addr 1 1, 0.0, 0.0, -1.0, 2.0)) |> Result.isError)
assertEq "drawing no-op keeps saved state and redo" (drawingRevision, false, true, true)
    (revision drawingState, isModified drawingState, canRedo drawingState, obj.ReferenceEquals(beforeDrawingNoOp, drawingState.Workbook.Value))
redo drawingState |> ignore
assertEq "redo retains opaque object" opaque (activeModel drawingState).Drawings.Head
assertEq "picture description changes" true
    ((executeDrawing drawingState (S.DescribePicture (5, None, Some "description")) |> Result.defaultWith failwith).Changed)
assertEq "picture description is visible" (Some "description")
    ((activeModel drawingState).Drawings |> List.find (fun drawing -> drawing.Id = 5) |> fun drawing -> drawing.Description)
undo drawingState |> ignore
assertEq "undo description retains picture" (Some 5) ((activeModel drawingState).Drawings |> List.tryFind (fun drawing -> drawing.Id = 5) |> Option.map (fun drawing -> drawing.Id))
executeDrawing drawingState (S.DeletePicture 5) |> Result.defaultWith failwith |> ignore
undo drawingState |> ignore
assertEq "undo picture deletion restores identity" [ 4; 5 ] ((activeModel drawingState).Drawings |> List.map (fun d -> d.Id))
applyCellEdit drawingState 0 0 "12" |> ignore
let chartRange = { From = addr 0 0; To = addr 0 0 }
let chart = executeDrawing drawingState (S.InsertChart (chartRange, Rbt.Spreadsheet.ChartAuthoring.Column, "Chart", false, false, [ "rIdChart2" ])) |> Result.defaultWith failwith
assertEq "chart uses next drawing id" (Some 6) chart.Id
assertEq "chart avoids imported and reserved relationship ids" [ "rIdChart3" ] ((activeModel drawingState).AuthoredCharts |> Map.toList |> List.map fst)
undo drawingState |> ignore
assertEq "undo chart retains imported object and picture" [ 4; 5 ] ((activeModel drawingState).Drawings |> List.map (fun d -> d.Id))
assertEq "undo chart removes authored part" Map.empty (activeModel drawingState).AuthoredCharts
redo drawingState |> ignore
assertEq "redo restores chart and part" ([ 4; 5; 6 ], [ "rIdChart3" ])
    ((activeModel drawingState).Drawings |> List.map (fun d -> d.Id), (activeModel drawingState).AuthoredCharts |> Map.toList |> List.map fst)

// Headless and Wisp use the same raster commands for clipped objects and images.
open Rbt.Renderer
let renderScene items =
    Rbt.SpreadsheetNative.GridRender.renderSceneCommands 64.0 80.0 items
    |> Rbt.Renderer.Scene.render (Rbt.Renderer.Raster.create()) 64 80
let rgbAt (frame: Framebuffer) x y =
    let memory = frame.Px
    let pixel = memory.Span.[y * frame.W + x]
    Pixel.r pixel, Pixel.g pixel, Pixel.b pixel
let clip : Rbt.Spreadsheet.Layout.Rect = { X = 10.0; Y = 20.0; Width = 30.0; Height = 40.0 }
let fill = Rbt.Spreadsheet.Paint.FillRect (clip, "#ff0000")
let scene = Rbt.Spreadsheet.Paint.Clipped (clip, [ Rbt.Spreadsheet.Paint.VecGroup ((1.0, 0.0, 0.0, 1.0, 2.0, 3.0), [ fill ]) ])
let clipped = renderScene [ scene ]
assertEq "transformed drawing is clipped at its outer bounds"
    ((255uy, 0uy, 0uy), (255uy, 255uy, 255uy))
    (rgbAt clipped 15 25, rgbAt clipped 41 25)
let innerClip : Rbt.Spreadsheet.Layout.Rect = { X = 15.0; Y = 22.0; Width = 5.0; Height = 10.0 }
let nested = renderScene [ Rbt.Spreadsheet.Paint.Clipped (clip, [ Rbt.Spreadsheet.Paint.Clipped (innerClip, [ fill ]) ]) ]
assertEq "nested clips intersect rather than replacing the outer bounds"
    ((255uy, 0uy, 0uy), (255uy, 255uy, 255uy), (255uy, 255uy, 255uy))
    (rgbAt nested 17 25, rgbAt nested 25 25, rgbAt nested 17 35)
let emptyClip : Rbt.Spreadsheet.Layout.Rect = { X = 10.0; Y = 20.0; Width = 0.0; Height = 40.0 }
assertEq "zero-width clip paints no child pixels" (255uy, 255uy, 255uy)
    (rgbAt (renderScene [ Rbt.Spreadsheet.Paint.Clipped (emptyClip, [ fill ]) ]) 15 25)
let gif = System.Convert.FromBase64String("R0lGODlhAQABAIAAAP8AAP///yH5BAAAAAAALAAAAAABAAEAAAIBRAA7")
let picture : Rbt.Spreadsheet.Paint.PaintImage = { ContentType = "image/gif"; Data = gif; Label = "image" }
assertEq "native GIF content is painted rather than a placeholder" (255uy, 0uy, 0uy)
    (rgbAt (renderScene [ Rbt.Spreadsheet.Paint.DrawImage (clip, picture) ]) 15 25)

printfn "wisp-editorhost: %d assertions passed" checks
