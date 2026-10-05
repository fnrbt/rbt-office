#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook
open Rbt.Spreadsheet.Formula

// Structure edits, part 2: print area / print titles follow row and column
// edits, cut-and-paste re-points the formulas that referred to the moved
// cells, a stand-alone sheet rewrites references that name the sheet itself,
// and a paste tiles across a selection that is a multiple of the block.

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)
let eq label (got: 'a) (expected: 'a) =
    ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let num x = cell (CellValue.Number x)
let fml f = cell (CellValue.Formula (f, CellValue.Empty))
let at r c = { Row = r; Col = c }
let N x = CellValue.Number x
let modelOf cells = { createEmpty 100 26 with Cells = Map.ofList cells }
let formulaAt r c (m: SpreadsheetModel) =
    match getCellValue (at r c) m with
    | CellValue.Formula (f, _) -> f
    | other -> sprintf "<%A>" other
let valueAt r c (m: SpreadsheetModel) = getCellValue (at r c) m |> unwrapFormulaValue
let wbOf (sheets: (string * string * SpreadsheetModel) list) : Workbook =
    { Sheets = sheets |> List.map (fun (id, name, m) -> id, { Id = id; Name = name; Model = m; Visible = true }) |> Map.ofList
      SheetOrder = sheets |> List.map (fun (id, _, _) -> id)
      ActiveSheet = (let (id, _, _) = List.head sheets in id)
      Tables = Map.empty
      Names = Map.empty
      NameInfo = Map.empty }
let sheetModel id (wb: Workbook) = wb.Sheets.[id].Model

// --- 1. Print area, print titles and manual breaks follow the edit ----------
let printed =
    { modelOf [ at 0 0, num 1.0 ] with
        PageSetup =
            { defaultPageSetup with
                PrintArea = Some { From = at 2 1; To = at 9 4 }
                RepeatRows = Some (0, 1)
                RepeatCols = Some (0, 0)
                ManualRowBreaks = [ 5 ]
                ManualColBreaks = [ 3 ] } }
let area (m: SpreadsheetModel) = m.PageSetup.PrintArea |> Option.map (fun a -> a.From.Row, a.From.Col, a.To.Row, a.To.Col)

let rowInsAbove = insertRow 1 printed
eq "print-area-ins-above"  (area rowInsAbove) (Some (3, 1, 10, 4))
eq "print-titles-grow"     rowInsAbove.PageSetup.RepeatRows (Some (0, 2))
eq "print-cols-untouched"  rowInsAbove.PageSetup.RepeatCols (Some (0, 0))
eq "print-break-moves"     rowInsAbove.PageSetup.ManualRowBreaks [ 6 ]
let rowInsInside = insertRow 5 printed
eq "print-area-ins-inside" (area rowInsInside) (Some (2, 1, 10, 4))
eq "print-titles-kept"     rowInsInside.PageSetup.RepeatRows (Some (0, 1))
let rowInsBelow = insertRow 20 printed
eq "print-area-ins-below"  (area rowInsBelow) (Some (2, 1, 9, 4))
let rowDelInside = Rbt.Spreadsheet.Model.applyStructureEdit (DeleteRows (4, 2)) printed
eq "print-area-del-inside" (area rowDelInside) (Some (2, 1, 7, 4))
let rowDelAll = Rbt.Spreadsheet.Model.applyStructureEdit (DeleteRows (2, 8)) printed
eq "print-area-del-all"    (area rowDelAll) None
let titleDel = deleteRow 0 printed
eq "print-titles-shrink"   titleDel.PageSetup.RepeatRows (Some (0, 0))
eq "print-area-del-above"  (area titleDel) (Some (1, 1, 8, 4))
let titlesGone = Rbt.Spreadsheet.Model.applyStructureEdit (DeleteRows (0, 2)) printed
eq "print-titles-deleted"  titlesGone.PageSetup.RepeatRows None
let colIns = insertColumn 0 printed
eq "print-area-col-ins"    (area colIns) (Some (2, 2, 9, 5))
eq "print-title-cols-move" colIns.PageSetup.RepeatCols (Some (1, 1))
eq "print-col-break-moves" colIns.PageSetup.ManualColBreaks [ 4 ]
let colDel = deleteColumn 2 printed
eq "print-area-col-del"    (area colDel) (Some (2, 1, 9, 3))
eq "print-title-rows-same" colDel.PageSetup.RepeatRows (Some (0, 1))

// --- 3. A stand-alone sheet that knows its name ----------------------------
let selfNamed = modelOf [ at 0 0, num 1.0; at 1 0, num 2.0; at 2 0, fml "=SUM(Data!A1:A2)+Data!A2+A1+Other!A2" ]
eq "named-insert"   (applyStructureEditNamed "Data" (InsertRows (1, 1)) selfNamed |> formulaAt 3 0) "=SUM(Data!A1:A3)+Data!A3+A1+Other!A2"
eq "named-case"     (applyStructureEditNamed "data" (InsertRows (0, 1)) selfNamed |> formulaAt 3 0) "=SUM(Data!A2:A3)+Data!A3+A2+Other!A2"
eq "named-delete"   (applyStructureEditNamed "Data" (DeleteRows (1, 1)) selfNamed |> formulaAt 1 0) "=SUM(Data!A1:A1)+#REF!+A1+Other!A2"
// Without a name only unqualified references can be known to be the sheet's own.
eq "unnamed-insert" (insertRow 1 selfNamed |> formulaAt 3 0) "=SUM(Data!A1:A2)+Data!A2+A1+Other!A2"
// The workbook path and the named single-sheet path agree.
let viaWorkbook, _ = Rbt.Spreadsheet.Workbook.applyStructureEdit "sheet1" (InsertRows (1, 1)) (wbOf [ "sheet1", "Data", selfNamed ])
eq "named-matches-workbook"
    (applyStructureEditNamed "Data" (InsertRows (1, 1)) selfNamed |> formulaAt 3 0)
    (formulaAt 3 0 (sheetModel "sheet1" viaWorkbook))

// --- 2. Cut and paste moves the references with the cells -------------------
/// Cut `selection` on sheet1 and paste it at (row, col) of `target`.
let cutPaste (selection: Selection) (target: string) (row: int) (col: int) (wb: Workbook) =
    let source = sheetModel "sheet1" wb
    let clip = serializeSelection selection source true
    let cleared = updateSheetModel "sheet1" (clearSelectionCells selection source) wb
    let pasted = updateSheetModel target (pasteClipboard (at row col) clip (sheetModel target cleared)) cleared
    repointMovedCells "sheet1" (clip.SourceStartRow, clip.SourceStartCol, clip.SourceEndRow, clip.SourceEndCol)
                      target (row - clip.SourceStartRow, col - clip.SourceStartCol) pasted

let moveData =
    modelOf [ at 0 0, num 10.0; at 1 0, num 20.0; at 2 0, num 30.0
              at 0 1, fml "=A1*2"            // inside the moved block? no: column B
              at 4 0, fml "=A1+A2"           // reads two moved cells
              at 5 0, fml "=SUM(A1:A2)"      // range wholly moved
              at 6 0, fml "=SUM(A1:A3)"      // range only partly moved: stays
              at 7 0, fml "=$A$1+D1"         // absolute follows too; D1 is overwritten
              at 8 0, fml "=Data!A2" ]
let moveOther = modelOf [ at 0 0, fml "=Data!A1+1"; at 1 0, fml "=SUM(Data!A1:A2)"; at 2 0, fml "=A1" ]
let moveWb =
    { wbOf [ "sheet1", "Data", moveData; "sheet2", "Other", moveOther ] with
        Names = namesOfTriples [ None, "First", "Data!$A$1"; None, "Pair", "Data!$A$1:$A$2"; Some "sheet1", "Local", "$A$2" ]
        NameInfo = Map.empty }

let moved, moveUndo = cutPaste (Range (0, 0, 1, 0)) "sheet1" 0 3 moveWb
let movedCalc = recalculateWorkbook moved
eq "move-cells-arrive"   (valueAt 0 3 (sheetModel "sheet1" moved), valueAt 1 3 (sheetModel "sheet1" moved)) (N 10.0, N 20.0)
eq "move-source-empty"   (valueAt 0 0 (sheetModel "sheet1" moved)) CellValue.Empty
eq "move-ref-follows"    (formulaAt 0 1 (sheetModel "sheet1" moved)) "=D1*2"
eq "move-two-refs"       (formulaAt 4 0 (sheetModel "sheet1" moved)) "=D1+D2"
eq "move-whole-range"    (formulaAt 5 0 (sheetModel "sheet1" moved)) "=SUM(D1:D2)"
eq "move-partial-range"  (formulaAt 6 0 (sheetModel "sheet1" moved)) "=SUM(A1:A3)"
eq "move-absolute+ref"   (formulaAt 7 0 (sheetModel "sheet1" moved)) "=$D$1+#REF!"
eq "move-self-qualified" (formulaAt 8 0 (sheetModel "sheet1" moved)) "=Data!D2"
eq "move-other-sheet"    (formulaAt 0 0 (sheetModel "sheet2" moved)) "=Data!D1+1"
eq "move-other-range"    (formulaAt 1 0 (sheetModel "sheet2" moved)) "=SUM(Data!D1:D2)"
eq "move-other-own-ref"  (formulaAt 2 0 (sheetModel "sheet2" moved)) "=A1"
eq "move-name"           (getName None "First" moved) (Some "Data!$D$1")
eq "move-name-range"     (getName None "Pair" moved) (Some "Data!$D$1:$D$2")
eq "move-local-name"     (getName (Some "sheet1") "Local" moved) (Some "$D$2")
eq "move-values-kept"    (valueAt 4 0 (sheetModel "sheet1" movedCalc)) (N 30.0)
eq "move-other-value"    (valueAt 0 0 (sheetModel "sheet2" movedCalc)) (N 11.0)

let unmoved = undoRepointMovedCells moveUndo moved
eq "move-undo-ref"   (formulaAt 4 0 (sheetModel "sheet1" unmoved)) "=A1+A2"
eq "move-undo-other" (formulaAt 0 0 (sheetModel "sheet2" unmoved)) "=Data!A1+1"
eq "move-undo-names" unmoved.Names moveWb.Names

// A moved formula keeps pointing at the cells it named; references between
// moved cells travel together.
let chain = modelOf [ at 0 0, num 5.0; at 1 0, fml "=A1+C1"; at 0 2, num 7.0; at 5 5, fml "=A2" ]
let chainMoved, _ = cutPaste (Range (0, 0, 1, 0)) "sheet1" 3 1 (wbOf [ "sheet1", "Data", chain ])
eq "move-formula-internal" (formulaAt 4 1 (sheetModel "sheet1" chainMoved)) "=B4+C1"
eq "move-formula-reader"   (formulaAt 5 5 (sheetModel "sheet1" chainMoved)) "=B5"
eq "move-formula-value"    (valueAt 5 5 (sheetModel "sheet1" (recalculateWorkbook chainMoved))) (N 12.0)

// Onto another sheet: readers gain (or lose) the sheet qualifier, and the
// moved formula keeps reading its old sheet.
let crossed, _ = cutPaste (Range (0, 0, 1, 0)) "sheet2" 5 5 (wbOf [ "sheet1", "Data", chain; "sheet2", "Other", modelOf [ at 0 0, fml "=Data!A2*2" ] ])
eq "cross-reader-source"   (formulaAt 5 5 (sheetModel "sheet1" crossed)) "=Other!F7"
eq "cross-reader-target"   (formulaAt 0 0 (sheetModel "sheet2" crossed)) "=F7*2"
eq "cross-moved-formula"   (formulaAt 6 5 (sheetModel "sheet2" crossed)) "=F6+Data!C1"
eq "cross-value"           (valueAt 0 0 (sheetModel "sheet2" (recalculateWorkbook crossed))) (N 24.0)

// --- 14. Paste tiles across a selection that is a multiple of the block -----
let tileSource = modelOf [ at 0 0, num 1.0; at 0 1, fml "=A1+1" ]
let tileClip = serializeSelection (Range (0, 0, 0, 1)) tileSource false
eq "tile-anchors-rows"   (pasteAnchors (Range (2, 0, 4, 1)) tileClip) [ at 2 0; at 3 0; at 4 0 ]
eq "tile-anchors-grid"   (pasteAnchors (Range (2, 0, 3, 3)) tileClip) [ at 2 0; at 2 2; at 3 0; at 3 2 ]
eq "tile-not-multiple"   (pasteAnchors (Range (2, 0, 3, 2)) tileClip) [ at 2 0 ]
eq "tile-single-cell"    (pasteAnchors (SingleCell (2, 0)) tileClip) [ at 2 0 ]
eq "tile-same-size"      (pasteAnchors (Range (2, 0, 2, 1)) tileClip) [ at 2 0 ]
eq "tile-cut-never"      (pasteAnchors (Range (2, 0, 4, 1)) { tileClip with IsCut = true }) [ at 2 0 ]
let tiled = pasteClipboardIntoSelection (Range (2, 0, 3, 3)) tileClip tileSource |> recalcAndSpill None None
eq "tile-formulas-shift" [ formulaAt 2 1 tiled; formulaAt 2 3 tiled; formulaAt 3 1 tiled; formulaAt 3 3 tiled ] [ "=A3+1"; "=C3+1"; "=A4+1"; "=C4+1" ]
eq "tile-values"         [ valueAt 2 2 tiled; valueAt 3 3 tiled ] [ N 1.0; N 2.0 ]
eq "tile-selection"      (pasteSelection (pasteAnchors (Range (2, 0, 3, 3)) tileClip) tileClip (SingleCell (0, 0))) (Range (2, 0, 3, 3))
let single = serializeSelection (SingleCell (0, 0)) tileSource false
eq "tile-one-cell-fill"  (pasteAnchors (Range (5, 0, 6, 1)) single |> List.length) 4

printfn ""
printfn "Stage 24 move tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
