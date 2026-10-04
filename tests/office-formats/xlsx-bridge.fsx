#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// XlsxBridge: reader workbook -> editor models -> writer workbook, the one
// conversion the web component (ExportXlsx) and the native editors share.

open System
open System.IO
open Rbt.Spreadsheet.Model
module Bridge = Rbt.Spreadsheet.XlsxBridge
module Wb = Rbt.Spreadsheet.Workbook

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-34s  %s" label detail)
let eq label (got: 'a) (expected: 'a) =
    ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let load (relative: string) = XlsxReader.loadBytesNative (File.ReadAllBytes(Path.Combine(root, relative))) |> Option.get
let reread (workbook: XlsxWriter.XlsxWorkbookOut) = XlsxReader.loadBytesNative (XlsxWriter.writeWorkbookPackageBytes workbook) |> Option.get
let inputOf (imported: Bridge.ImportedWorkbook) (sheets: SpreadsheetModel array) (names: Map<Wb.SheetId option * string, string>) : Bridge.ExportInput =
    { Sheets = sheets
      Names = imported.Sheets |> Array.map (fun sheet -> sheet.Name)
      Visible = imported.Sheets |> Array.map (fun sheet -> sheet.Visible)
      ActiveSheetIndex = 0
      DefinedNames = names
      Source = Some imported.Source
      Macros = None }
let modelsOf (imported: Bridge.ImportedWorkbook) = imported.Sheets |> Array.map (fun sheet -> sheet.Model)
let cellsOf (sheet: XlsxWriter.XlsxSheetOut) = sheet.Cells |> List.map (fun cell -> (cell.Row, cell.Col), cell) |> Map.ofList
let addr row col : CellAddress = { Row = row; Col = col }

// ---- an unedited workbook goes out as the file had it ----------------------
for file in [ "real-world/milk-street-office.xlsx"; "real-world/saopaulo-condos.xlsx"; "real-world/chart-zoo.xlsx"; "test-documents-excel/phpsp-27template.xlsx" ] do
    let wb = load file
    let imported = Bridge.Import.importWorkbook wb
    let original = XlsxWriter.packageFromReaderWorkbook wb
    let out = Bridge.Export.toWorkbookOut (inputOf imported (modelsOf imported) imported.Context.Names)
    eq (file + " sheet names") (out.Sheets |> List.map (fun s -> s.Name, s.State)) (original.Sheets |> List.map (fun s -> s.Name, s.State))
    eq (file + " defined names") out.DefinedNames original.DefinedNames
    eq (file + " dxfs") out.Dxfs original.Dxfs
    ok (file + " kept parts") (obj.ReferenceEquals(out.Preserved, original.Preserved) || out.Preserved = original.Preserved) "preserved parts differ"
    for sheetOut, sheetOriginal, sheetImported in List.zip3 out.Sheets original.Sheets (List.ofArray imported.Sheets) do
        let label = sprintf "%s/%s" file sheetOut.Name
        let originals = cellsOf sheetOriginal
        // Every cell the model holds is written exactly as the file had it ...
        let changed = sheetOut.Cells |> List.filter (fun cell -> Map.tryFind (cell.Row, cell.Col) originals <> Some cell)
        eq (label + " cells as in the file") (changed |> List.truncate 3 |> List.map (fun cell -> cell.Row, cell.Col, cell.Value)) []
        // ... and nothing with content or formatting is dropped.
        let written = cellsOf sheetOut
        let dropped =
            sheetOriginal.Cells
            |> List.filter (fun cell -> not (Map.containsKey (cell.Row, cell.Col) written) && (cell.Value <> XlsxWriter.XlsxBlank || cell.Format <> XlsxWriter.defaultXlsxCellFormat))
        eq (label + " no cell dropped") (dropped |> List.truncate 3 |> List.map (fun cell -> cell.Row, cell.Col, cell.Value)) []
        eq (label + " model cell count") sheetOut.Cells.Length sheetImported.Model.Cells.Count
        eq (label + " column widths") sheetOut.ColumnWidths sheetOriginal.ColumnWidths
        eq (label + " row heights") sheetOut.RowHeights sheetOriginal.RowHeights
        eq (label + " merges") (List.sort sheetOut.MergedCells) (List.sort sheetOriginal.MergedCells)
        eq (label + " freeze / grid") (sheetOut.FrozenRows, sheetOut.FrozenCols, sheetOut.ShowGridLines) (sheetOriginal.FrozenRows, sheetOriginal.FrozenCols, sheetOriginal.ShowGridLines)
        eq (label + " tables") sheetOut.Tables sheetOriginal.Tables
        eq (label + " validations") sheetOut.DataValidations sheetOriginal.DataValidations
        eq (label + " conditional formats") sheetOut.ConditionalFormats sheetOriginal.ConditionalFormats
        eq (label + " hyperlinks") sheetOut.Hyperlinks sheetOriginal.Hyperlinks
        eq (label + " page setup") (sheetOut.PageSetup, sheetOut.PageMargins, sheetOut.PrintOptions, sheetOut.HeaderFooter, sheetOut.PrintArea, sheetOut.PrintTitleRows, sheetOut.PrintTitleCols) (sheetOriginal.PageSetup, sheetOriginal.PageMargins, sheetOriginal.PrintOptions, sheetOriginal.HeaderFooter, sheetOriginal.PrintArea, sheetOriginal.PrintTitleRows, sheetOriginal.PrintTitleCols)
        ok (label + " drawings and threads") (sheetOut.Extras = sheetOriginal.Extras) "extras differ"
    // The written package reads back with the same cells, pictures and kept objects.
    let again = reread out
    let describe (workbook: XlsxReader.Workbook) =
        workbook.Sheets
        |> List.map (fun sheet ->
            sheet.Name,
            (sheet.Rows |> List.sumBy (fun row -> row.Cells |> List.filter (fun cell -> cell.Formula.IsSome) |> List.length)),
            (sheet.Extras.Images |> List.map (fun image -> image.Anchor, image.Data.Length)),
            (sheet.Extras.Objects |> List.map (fun item -> item.Anchor, item.Kind)),
            sheet.MergedCells.Length, sheet.Tables.Length, sheet.ConditionalFormats.Length, sheet.DataValidations.Length, sheet.Hyperlinks.Length)
    eq (file + " reads back the same") (describe again) (describe wb)

// ---- edits: values, a formula, an inserted row, a merge, sizes -------------
let wb = load "real-world/milk-street-office.xlsx"
let imported = Bridge.Import.importWorkbook wb
let sheetId = Bridge.sheetIdOfIndex 0
let edited =
    let first = imported.Sheets.[0].Model
    // Find a formula cell and a number cell to work with.
    let formulaAddr, formulaText =
        first.Cells |> Map.toSeq |> Seq.pick (fun (a, cell) -> match cell.Value with CellValue.Formula (text, CellValue.Number _) when a.Row > 6 -> Some (a, text) | _ -> None)
    let withValues =
        first
        |> setCell (addr 0 30) { Value = CellValue.String "edited text"; Format = { defaultFormat with Bold = true; BgColor = Some "#FFEE00" }; Comment = None }
        |> setCell (addr 1 30) { Value = CellValue.Number 1234.5; Format = { defaultFormat with NumberFormat = Some "#,##0.00" }; Comment = Some { Text = "a note"; Author = Some "Tester" } }
        |> setCell (addr 2 30) { Value = CellValue.Formula ("=AE2*2", CellValue.Number 2469.0); Format = defaultFormat; Comment = None }
        |> setCell (addr 3 30) { Value = CellValue.Formula ("AE2+1", CellValue.Empty); Format = defaultFormat; Comment = None }
    let context = { imported.Context with Sheets = imported.Context.Sheets |> Map.add sheetId { imported.Context.Sheets.[sheetId] with Model = withValues } }
    // Insert two rows at row 5 through the workbook, so names and other sheets follow.
    let moved, _ = Wb.applyStructureEdit sheetId (InsertRows (4, 2)) context
    let model = moved.Sheets.[sheetId].Model
    let model =
        { model with MergedCells = { StartRow = 0; StartCol = 31; EndRow = 1; EndCol = 33 } :: model.MergedCells; FrozenRows = 3; FrozenCols = 1 }
        |> setColumnWidth 30 210.0
        |> setRowHeight 0 0.0
        |> setRowHeight 1 60.0
    formulaAddr, formulaText, { moved with Sheets = moved.Sheets |> Map.add sheetId { moved.Sheets.[sheetId] with Model = model } }
let formulaAddr, formulaText, editedContext = edited
let editedSheets = imported.Sheets |> Array.mapi (fun i _ -> editedContext.Sheets.[Bridge.sheetIdOfIndex i].Model)
let editedOut = Bridge.Export.toWorkbookOut (inputOf imported editedSheets editedContext.Names)
let editedBack = reread editedOut
let backSheet = editedBack.Sheets.[0]
let backCell row col =
    backSheet.Rows |> List.tryFind (fun r -> r.Index = row + 1)
    |> Option.bind (fun r -> r.Cells |> List.tryFind (fun c -> Bridge.columnIndexOfLetters c.Reference.Column = col))
let backValue row col = backCell row col |> Option.map (fun c -> XlsxReader.resolveValue editedBack c.Value)
eq "edit: text" (backValue 0 30) (Some (XlsxReader.String "edited text"))
eq "edit: text format" (backCell 0 30 |> Option.bind (fun c -> c.Style) |> Option.map (fun s -> s.Bold, s.FillColor |> Option.map (fun c -> c.ToUpperInvariant().TrimStart('#')))) (Some (true, Some "FFEE00"))
eq "edit: number" (backValue 1 30) (Some (XlsxReader.Number 1234.5))
eq "edit: number format" (backCell 1 30 |> Option.bind (fun c -> c.Style) |> Option.bind (fun s -> s.NumberFormat)) (Some "#,##0.00")
eq "edit: comment" (backSheet.Comments |> List.filter (fun c -> c.CellRef.Row = 2 && c.CellRef.Column = "AE") |> List.map (fun c -> c.Text, c.Author)) [ "a note", Some "Tester" ]
eq "edit: formula with cached value" (backCell 2 30 |> Option.map (fun c -> c.Formula, c.Value)) (Some (Some "AE2*2", XlsxReader.Number 2469.0))
eq "edit: formula without cached value" (backCell 3 30 |> Option.bind (fun c -> c.Formula)) (Some "AE2+1")
// The formula that was at formulaAddr sits two rows lower when it was below the insertion.
let movedRow = if formulaAddr.Row >= 4 then formulaAddr.Row + 2 else formulaAddr.Row
ok "edit: formulas moved with the inserted rows" ((backCell movedRow formulaAddr.Col |> Option.bind (fun c -> c.Formula)).IsSome) (sprintf "no formula at row %d" movedRow)
ok "edit: nothing left at the old place of row 5" (backSheet.Rows |> List.forall (fun r -> r.Index <> 5 && r.Index <> 6 || r.Cells |> List.forall (fun c -> c.Formula.IsNone && c.Value = XlsxReader.Empty))) "rows 5-6 still have content"
eq "edit: merge" (backSheet.MergedCells |> List.exists (fun m -> m.StartRef.Row = 1 && m.StartRef.Column = "AF" && m.EndRef.Row = 2 && m.EndRef.Column = "AH")) true
eq "edit: merges of the file moved down" (backSheet.MergedCells.Length) (wb.Sheets.[0].MergedCells.Length + 1)
eq "edit: freeze panes" (backSheet.FrozenPane |> Option.map (fun p -> p.YSplit, p.XSplit)) (Some (3, 1))
eq "edit: column width" (backSheet.Columns |> List.tryFind (fun c -> c.Index <= 31 && c.MaxIndex >= 31) |> Option.bind (fun c -> c.Width) |> Option.map (fun w -> Math.Round(w, 1))) (Some 30.0)
eq "edit: hidden row / row height" (backSheet.Rows |> List.filter (fun r -> r.Index <= 2) |> List.map (fun r -> r.Index, r.Hidden, r.Height |> Option.map (fun h -> Math.Round(h, 0)))) [ 1, true, None; 2, false, Some 45.0 ]
eq "edit: other sheets untouched" (editedOut.Sheets |> List.tail |> List.map (fun s -> s.Cells.Length)) ((XlsxWriter.packageFromReaderWorkbook wb).Sheets |> List.tail |> List.map (fun s -> s.Cells |> List.filter (fun c -> c.Value <> XlsxWriter.XlsxBlank || c.Format <> XlsxWriter.defaultXlsxCellFormat) |> List.length))
eq "edit: sheets, kept parts and named styles carried over" (editedBack.Sheets |> List.map (fun s -> s.Name), editedBack.Preserved.NamedStyles.Length, editedBack.Preserved.ThemePath) (wb.Sheets |> List.map (fun s -> s.Name), wb.Preserved.NamedStyles.Length, wb.Preserved.ThemePath)
eq "edit: defined names kept (count)" editedBack.NamedRanges.Length wb.NamedRanges.Length
// A cell that was not touched keeps what the file said, to the last digit.
let untouched =
    let originalOut = (XlsxWriter.packageFromReaderWorkbook wb).Sheets.[0] |> cellsOf
    // (Formulas above the insertion that point below it were re-pointed.)
    editedOut.Sheets.[0].Cells |> List.filter (fun cell -> cell.Row < 4 && cell.Col < 30 && (match cell.Value with XlsxWriter.XlsxFormula _ -> false | _ -> true))
    |> List.forall (fun cell -> Map.tryFind (cell.Row, cell.Col) originalOut = Some cell)
ok "edit: untouched cells are the file's" untouched "a cell above the insertion changed"

// ---- a formula whose value only became an error keeps the file's value -----
let errorModel =
    let first = imported.Sheets.[0].Model
    first |> setCell formulaAddr { first.Cells.[formulaAddr] with Value = CellValue.Formula (formulaText, CellValue.Error "NAME?") }
let errorOut = Bridge.Export.toWorkbookOut (inputOf imported (imported.Sheets |> Array.mapi (fun i sheet -> if i = 0 then errorModel else sheet.Model)) imported.Context.Names)
eq "a result that only became an error is not written" ((cellsOf errorOut.Sheets.[0]).[(formulaAddr.Row, formulaAddr.Col)]) ((cellsOf (XlsxWriter.packageFromReaderWorkbook wb).Sheets.[0]).[(formulaAddr.Row, formulaAddr.Col)])

// ---- a workbook made in the editor (no file behind it) ---------------------
let fresh =
    createEmpty 50 10
    |> setCell (addr 0 0) { Value = CellValue.String "Name"; Format = { defaultFormat with Bold = true }; Comment = None }
    |> setCell (addr 1 0) { Value = CellValue.Number 3.0; Format = defaultFormat; Comment = None }
    |> setCell (addr 2 0) { Value = CellValue.Formula ("=A2*2", CellValue.Number 6.0); Format = defaultFormat; Comment = None }
    |> setCell (addr 3 0) { Value = CellValue.Boolean true; Format = defaultFormat; Comment = None }
    |> setCell (addr 4 0) { Value = CellValue.Error "DIV/0!"; Format = defaultFormat; Comment = None }
let freshOut =
    Bridge.Export.toWorkbookOut
        { Sheets = [| fresh; createEmpty 10 5 |]; Names = [| "First"; "Hidden one" |]; Visible = [| true; false |]; ActiveSheetIndex = 0
          DefinedNames = Wb.namesOfTriples [ None, "Total", "First!$A$3"; Some (Bridge.sheetIdOfIndex 0), "Local", "First!$A$2" ]
          Source = None; Macros = None }
let freshBack = reread freshOut
eq "fresh: sheets and visibility" (freshBack.Sheets |> List.map (fun s -> s.Name, s.State)) [ "First", XlsxReader.SheetVisible; "Hidden one", XlsxReader.SheetHidden ]
eq "fresh: values" (freshBack.Sheets.[0].Rows |> List.collect (fun r -> r.Cells |> List.map (fun c -> c.Formula, XlsxReader.resolveValue freshBack c.Value)))
    [ None, XlsxReader.String "Name"; None, XlsxReader.Number 3.0; Some "A2*2", XlsxReader.Number 6.0; None, XlsxReader.Boolean true; None, XlsxReader.Error "DIV/0!" ]
eq "fresh: defined names" (freshBack.NamedRanges |> List.map (fun n -> n.Name.ToUpperInvariant(), n.RefersTo, n.Scope) |> List.sort) [ "LOCAL", "First!$A$2", Some "0"; "TOTAL", "First!$A$3", None ]
// Importing what was exported gives the models back.
let freshAgain = Bridge.Import.importWorkbook freshBack
eq "fresh: import of the export" (freshAgain.Sheets.[0].Model.Cells |> Map.map (fun _ cell -> cell.Value)) (fresh.Cells |> Map.map (fun _ cell -> match cell.Value with CellValue.Formula (text, cached) -> CellValue.Formula (text.TrimStart '=', cached) | value -> value))

// ---- import options ----------------------------------------------------------
let roomy = Bridge.Import.importWorkbookWith Bridge.editorImportOptions wb
ok "editor options leave room to grow" (roomy.Sheets.[0].Model.RowCount >= imported.Sheets.[0].Model.RowCount + 100 && roomy.Sheets.[0].Model.ColCount >= imported.Sheets.[0].Model.ColCount + 20) "model is not larger"
eq "the cells are the same either way" (roomy.Sheets.[0].Model.Cells = imported.Sheets.[0].Model.Cells) true

printfn ""
printfn "XLSX bridge tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
