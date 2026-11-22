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
let reread (workbook: XlsxWriter.XlsxWorkbookOut) = XlsxReader.loadBytesNative (XlsxPackageWriter.writeWorkbookPackageBytes workbook) |> Option.get
let inputOf (imported: Bridge.ImportedWorkbook) (sheets: SpreadsheetModel array) (names: Map<Wb.SheetId option * string, string>) : Bridge.ExportInput =
    { Sheets = sheets
      Names = imported.Sheets |> Array.map (fun sheet -> sheet.Name)
      Visible = imported.Sheets |> Array.map (fun sheet -> sheet.Visible)
      ActiveSheetIndex = 0
      NameInfo = Map.empty; DefinedNames = names
      Source = Some imported.Source
      Macros = None; MacroProject = None }
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
    // (The sheets go back as the parts they were read from: those parts, and the styles and shared strings they index, are added.)
    let withoutSheetParts (kept: XlsxReader.WorkbookPreservation) = { kept with InPlace = false; Parts = kept.Parts |> List.filter (fun (path, _) -> original.Preserved.Parts |> List.exists (fun (known, _) -> known = path)) }
    ok (file + " kept parts") (withoutSheetParts out.Preserved = original.Preserved) "preserved parts differ"
    ok (file + " sheets as they were read") (out.Preserved.InPlace && out.Sheets |> List.forall (fun sheet -> sheet.Extras.Verbatim)) "not verbatim"
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
        ok (label + " drawings and threads") ({ sheetOut.Extras with Verbatim = false } = sheetOriginal.Extras) "extras differ"
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
        |> setCell (addr 1 30) { Value = CellValue.Number 1234.5; Format = { defaultFormat with NumberFormat = Some "#,##0.00" }; Comment = Some { Text = "a note"; Author = Some "Tester"; Runs = [] } }
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
// (210 px in the digits of the workbook's Normal font: Calibri 12's 8 px.)
eq "edit: column width" (backSheet.Columns |> List.tryFind (fun c -> c.Index <= 31 && c.MaxIndex >= 31) |> Option.bind (fun c -> c.Width) |> Option.map (fun w -> Math.Round(w, 2))) (Some (Math.Round (210.0 / editedSheets.[0].MaxDigitWidth, 2)))
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

// ---- a cell whose value alone changed keeps the file's format and named style
let styledOriginal =
    (XlsxWriter.packageFromReaderWorkbook wb).Sheets
    |> List.mapi (fun index sheet -> sheet.Cells |> List.filter (fun cell -> cell.Format.NamedStyle.IsSome && (match cell.Value with XlsxWriter.XlsxNumber _ -> true | _ -> false)) |> List.map (fun cell -> index, cell))
    |> List.concat
    |> List.tryHead
match styledOriginal with
| None -> ok "a number cell based on a named style exists" false "none in milk-street-office.xlsx"
| Some (sheetIndex, styledCell) ->
    let address = addr styledCell.Row styledCell.Col
    let models = imported.Sheets |> Array.map (fun sheet -> sheet.Model)
    let retyped = models |> Array.mapi (fun i model -> if i = sheetIndex then setCell address { model.Cells.[address] with Value = CellValue.Number 98765.0 } model else model)
    let out = (Bridge.Export.toWorkbookOut (inputOf imported retyped imported.Context.Names)).Sheets.[sheetIndex] |> cellsOf
    eq "value-only edit: new value, the file's format and named style" (out.[(styledCell.Row, styledCell.Col)]) { styledCell with Value = XlsxWriter.XlsxNumber 98765.0 }
    let restyled = models |> Array.mapi (fun i model -> if i = sheetIndex then setCell address { model.Cells.[address] with Format = { model.Cells.[address].Format with Bold = not model.Cells.[address].Format.Bold } } model else model)
    let outRestyled = (Bridge.Export.toWorkbookOut (inputOf imported restyled imported.Context.Names)).Sheets.[sheetIndex] |> cellsOf
    eq "format edit: the editor's format is written" (outRestyled.[(styledCell.Row, styledCell.Col)].Format.Bold) (not styledCell.Format.Bold)

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
          NameInfo = Map.empty; DefinedNames = Wb.namesOfTriples [ None, "Total", "First!$A$3"; Some (Bridge.sheetIdOfIndex 0), "Local", "First!$A$2" ]
          Source = None; Macros = None; MacroProject = None }
let freshBack = reread freshOut
eq "fresh: sheets and visibility" (freshBack.Sheets |> List.map (fun s -> s.Name, s.State)) [ "First", XlsxReader.SheetVisible; "Hidden one", XlsxReader.SheetHidden ]
eq "fresh: values" (freshBack.Sheets.[0].Rows |> List.collect (fun r -> r.Cells |> List.map (fun c -> c.Formula, XlsxReader.resolveValue freshBack c.Value)))
    [ None, XlsxReader.String "Name"; None, XlsxReader.Number 3.0; Some "A2*2", XlsxReader.Number 6.0; None, XlsxReader.Boolean true; None, XlsxReader.Error "DIV/0!" ]
eq "fresh: defined names" (freshBack.NamedRanges |> List.map (fun n -> n.Name.ToUpperInvariant(), n.RefersTo, n.Scope) |> List.sort) [ "LOCAL", "First!$A$2", Some "0"; "TOTAL", "First!$A$3", None ]
// Importing what was exported gives the models back.
let freshAgain = Bridge.Import.importWorkbook freshBack
eq "fresh: import of the export" (freshAgain.Sheets.[0].Model.Cells |> Map.map (fun _ cell -> cell.Value)) (fresh.Cells |> Map.map (fun _ cell -> match cell.Value with CellValue.Formula (text, cached) -> CellValue.Formula (text.TrimStart '=', cached) | value -> value))

// ---- defined names: spelling, comment, scope and the hidden flag ------------
let namedBook =
    [ None, "TaxRate", "0.19", Some "VAT, standard"; Some (Bridge.sheetIdOfIndex 0), "Local_Total", "First!$A$3", None; None, "plain", "First!$A$2", None ]
    |> List.fold (fun (wb: Wb.Workbook) (scope, name, refersTo, comment) -> Wb.setDefinedName scope name refersTo comment wb)
        (Bridge.buildWorkbookContext [| fresh; createEmpty 10 5 |] [| "First"; "Hidden one" |] [| true; false |] Map.empty Map.empty 0)
let namedInput (wb: Wb.Workbook) (source: Bridge.XlsxSource option) : Bridge.ExportInput =
    { Sheets = [| fresh; createEmpty 10 5 |]; Names = [| "First"; "Hidden one" |]; Visible = [| true; false |]; ActiveSheetIndex = 0
      NameInfo = wb.NameInfo; DefinedNames = wb.Names; Source = source; Macros = None; MacroProject = None }
let namedBack = reread (Bridge.Export.toWorkbookOut (namedInput namedBook None))
eq "names: written with their spelling, comment and scope"
    (namedBack.NamedRanges |> List.map (fun n -> n.Name, n.RefersTo, n.Scope, n.Comment, n.Hidden) |> List.sort)
    [ "Local_Total", "First!$A$3", Some "0", None, false; "TaxRate", "0.19", None, Some "VAT, standard", false; "plain", "First!$A$2", None, None, false ]
let namedAgain = Bridge.Import.importWorkbook namedBack
eq "names: the import has what the export had" (Wb.definedNames namedAgain.Context) (Wb.definedNames namedBook)
// Untouched names are written as the file had them; a changed one with what it is now, the others unchanged.
let namesUntouched = reread (Bridge.Export.toWorkbookOut { namedInput namedAgain.Context (Some namedAgain.Source) with Sheets = namedAgain.Sheets |> Array.map (fun sheet -> sheet.Model) })
eq "names: an unedited workbook writes them again" (namesUntouched.NamedRanges |> List.map (fun n -> n.Name, n.Comment) |> List.sort) [ "Local_Total", None; "TaxRate", Some "VAT, standard"; "plain", None ]
let renamedBook = namedAgain.Context |> Wb.setDefinedName None "TAXRATE" "0.07" (Some "reduced") |> Wb.deleteDefinedName None "plain" |> Option.get
let namesRenamed = reread (Bridge.Export.toWorkbookOut { namedInput renamedBook (Some namedAgain.Source) with Sheets = namedAgain.Sheets |> Array.map (fun sheet -> sheet.Model) })
eq "names: an edited name is written as it is now" (namesRenamed.NamedRanges |> List.map (fun n -> n.Name, n.RefersTo, n.Scope, n.Comment) |> List.sort)
    [ "Local_Total", "First!$A$3", Some "0", None; "TAXRATE", "0.07", None, Some "reduced" ]
// A name the file hides from the name manager keeps its flag.
let filterKey = Some (Bridge.sheetIdOfIndex 0), "SECRET"
let hiddenBook =
    { namedBook with
        Names = namedBook.Names |> Map.add filterKey "First!$A$1:$A$5"
        NameInfo = namedBook.NameInfo |> Map.add filterKey { Name = "Secret"; Comment = None; Hidden = true } }
let hiddenOut = Bridge.Export.toWorkbookOut (namedInput hiddenBook None)
eq "names: a hidden name" ((reread hiddenOut).NamedRanges |> List.filter (fun n -> n.Hidden) |> List.map (fun n -> n.Name, n.Scope)) [ "Secret", Some "0" ]
eq "names: ... is hidden after the import too" ((Bridge.Import.importWorkbook (reread hiddenOut)).Context |> Wb.tryFindDefinedName (fst filterKey) "secret" |> Option.map (fun n -> n.Hidden)) (Some true)

// ---- the view of a sheet: zoom, headers, formulas, zeros, grid lines, panes ----
let viewed = { fresh with Zoom = 1.35; ShowHeaders = false; ShowFormulas = true; ShowZeros = false; ShowGridLines = false; FrozenRows = 3; FrozenCols = 2 }
let viewOut =
    Bridge.Export.toWorkbookOut
        { Sheets = [| viewed; createEmpty 10 5 |]; Names = [| "Viewed"; "Plain" |]; Visible = [| true; true |]; ActiveSheetIndex = 0
          NameInfo = Map.empty; DefinedNames = Map.empty; Source = None; Macros = None; MacroProject = None }
let viewBack = reread viewOut
eq "view: written to the sheet view" (viewBack.Sheets |> List.map (fun s -> s.View, s.ShowZeros, s.ShowGridLines, s.FrozenPane))
    [ { XlsxReader.ShowRowColHeaders = false; ShowFormulas = true; ZoomScale = 135; TabColor = None; SummaryBelow = true; SummaryRight = true; ShowOutlineSymbols = true }, false, false, Some { XlsxReader.XSplit = 2; YSplit = 3 }
      XlsxReader.defaultSheetView, true, true, None ]
let viewAgain = (Bridge.Import.importWorkbook viewBack).Sheets |> Array.map (fun sheet -> sheet.Model)
eq "view: the import has what the export had"
    (viewAgain |> Array.map (fun m -> m.Zoom, m.ShowHeaders, m.ShowFormulas, m.ShowZeros, m.ShowGridLines, m.FrozenRows, m.FrozenCols))
    [| 1.35, false, true, false, false, 3, 2; 1.0, true, false, true, true, 0, 0 |]
// (The lazy import converts a sheet later: the view comes with it.)
let viewLazy = (Bridge.Import.importWorkbookLazy (fun _ -> false) viewBack).Sheets.[0].Model |> resolveSheet
eq "view: a sheet converted on demand" (viewLazy.Zoom, viewLazy.ShowHeaders, viewLazy.ShowFormulas, viewLazy.ShowZeros) (1.35, false, true, false)

// ---- every property of a cell format set in the editor comes back ----------
let dressed =
    { defaultFormat with
        Bold = true; Italic = true; Underline = true; UnderlineDouble = true; Strikethrough = true
        FontName = Some "Georgia"; FontSize = Some 14.0; TextColor = Some "#112233"; BgColor = Some "#FFEEDD"
        HorizontalAlign = Some "right"; VerticalAlign = Some "top"; WrapText = true; Indent = 2; ShrinkToFit = false; TextRotation = 45
        NumberFormat = Some "#,##0.000"
        BorderTop = Some { Style = "thin"; Color = Some "#FF0000" }; BorderBottom = Some { Style = "double"; Color = Some "#00FF00" }
        BorderLeft = Some { Style = "dashed"; Color = Some "#0000FF" }; BorderRight = Some { Style = "thick"; Color = Some "#000000" } }
let shrunk = { defaultFormat with ShrinkToFit = true; TextRotation = 135; VerticalAlign = Some "middle"; HorizontalAlign = Some "center"; NumberFormat = Some "0.00%" }
let dressedSheet =
    createEmpty 20 6
    |> setCell (addr 0 0) { Value = CellValue.Number 1234.5; Format = dressed; Comment = None }
    |> setCell (addr 1 1) { Value = CellValue.String "shrunk"; Format = shrunk; Comment = None }
let dressedBack =
    Bridge.Export.toWorkbookOut { Sheets = [| dressedSheet |]; Names = [| "Formats" |]; Visible = [| true |]; ActiveSheetIndex = 0; NameInfo = Map.empty; DefinedNames = Map.empty; Source = None; Macros = None; MacroProject = None }
    |> reread |> Bridge.Import.importWorkbook
let dressedFormat row col = dressedBack.Sheets.[0].Model.Cells.[addr row col].Format
let sameColors (format: CellFormat) =
    let upper = Option.map (fun (color: string) -> color.ToUpperInvariant())
    let border = Option.map (fun (b: BorderStyle) -> { b with Color = upper b.Color })
    { format with TextColor = upper format.TextColor; BgColor = upper format.BgColor; BorderTop = border format.BorderTop; BorderBottom = border format.BorderBottom
                  BorderLeft = border format.BorderLeft; BorderRight = border format.BorderRight
                  VerticalAlign = (match format.VerticalAlign with Some "center" -> Some "middle" | other -> other) }
eq "formats: font, colours, alignment, wrap, indent, rotation, number format and borders round trip" (sameColors (dressedFormat 0 0)) dressed
// (A cell without a font of its own comes back with the workbook's default font spelled out.)
eq "formats: shrink to fit, a clockwise rotation, middle / centre, a percent format" { sameColors (dressedFormat 1 1) with FontName = None; FontSize = None } shrunk

// ---- validations are kept by range: a whole column is not clipped, nothing is made per cell ----
let area r0 c0 r1 c1 : CellRange = { From = addr r0 c0; To = addr r1 c1 }
let wholeCol c = area 0 c 1048575 c
let validated =
    createEmpty 60 12
    |> setCell (addr 0 0) { Value = CellValue.Number 4.0; Format = defaultFormat; Comment = None }
    |> setValidationRanges [ wholeCol 0 ] { validationOf (WholeNumber (Some 1.0, Some 10.0)) with ErrorTitle = Some "Range"; ErrorMessage = Some "1 to 10"; Style = WarningStyle }
    |> setValidationRanges [ area 1 1 20 1; area 1 3 20 3 ] { validationOf (ListSource "Lists!$A$1:$A$3") with ShowDropdown = false }
    |> setValidationRanges [ area 0 2 0 2 ] (validationOf (ListValues [ "Yes"; "No" ]))
    |> setValidationRanges [ area 2 4 9 4 ] { validationOf (CustomFormula "E3>$A$1") with InputTitle = Some "Hint"; InputMessage = Some "More than A1"; ShowInput = false }
    |> setValidationRanges [ area 0 5 0 5 ] { validationOf AnyValue with InputMessage = Some "Free text" }
    |> setValidationRanges [ area 0 6 5 6 ] (validationOf (Comparison (DateKind, OpNotBetween, "45000", Some "45100")))
let lists =
    createEmpty 10 3
    |> setCell (addr 0 0) { Value = CellValue.String "North"; Format = defaultFormat; Comment = None }
    |> setCell (addr 1 0) { Value = CellValue.String "South"; Format = defaultFormat; Comment = None }
let validatedOut =
    Bridge.Export.toWorkbookOut { Sheets = [| validated; lists |]; Names = [| "Data"; "Lists" |]; Visible = [| true; true |]; ActiveSheetIndex = 0; NameInfo = Map.empty; DefinedNames = Map.empty; Source = None; Macros = None; MacroProject = None }
let validatedFile = reread validatedOut
// (The list that reads the other sheet is in the x14 list of the sheet, which is read after the others.)
eq "validations: the sqrefs of the file (a whole column written out)" (validatedFile.Sheets.[0].DataValidations |> List.map (fun dv -> dv.Range, dv.ValidationType) |> List.sort)
   [ "A1:A1048576", "whole"; "B2:B21 D2:D21", "list"; "C1", "list"; "E3:E10", "custom"; "F1", "none"; "G1:G6", "date" ]
let validatedBack = Bridge.Import.importWorkbook validatedFile
let backModel = validatedBack.Sheets.[0].Model
eq "validations: every kind comes back for the same ranges" (backModel.Validations |> List.map (fun v -> v.Ranges, v.Validation) |> List.sort) (validated.Validations |> List.map (fun v -> v.Ranges, v.Validation) |> List.sort)
eq "validations: under ids 1.." (backModel.Validations |> List.map (fun v -> v.Id)) [ 1 .. validated.Validations.Length ]
eq "validations: a whole column makes no cell" (backModel.Cells.Count, (getValidation (addr 1048575 0) backModel).IsSome, (getValidation (addr 500000 1) backModel).IsSome) (1, true, false)
eq "validations: the lazy import gives the same" ((Bridge.Import.importWorkbookLazy (fun _ -> false) validatedFile).Sheets |> Array.map (fun sheet -> (resolveSheet sheet.Model).Validations)) [| backModel.Validations; [] |]
// Unedited, the file's validations are written as the file had them; edited, from the model.
let unedited = Bridge.Export.toWorkbookOut (inputOf validatedBack (modelsOf validatedBack) validatedBack.Context.Names)
eq "validations: unedited, as in the file" unedited.Sheets.[0].DataValidations (XlsxWriter.packageFromReaderWorkbook validatedFile).Sheets.[0].DataValidations
let editedModel = backModel |> setValidationRanges [ area 0 0 4 0 ] (validationOf (ListValues [ "a" ])) |> clearValidationRanges [ wholeCol 6 ]
let editedValidations =
    Bridge.Export.toWorkbookOut (inputOf validatedBack [| editedModel; validatedBack.Sheets.[1].Model |] validatedBack.Context.Names) |> reread
eq "validations: edited ones are written from the model" (editedValidations.Sheets.[0].DataValidations |> List.map (fun dv -> dv.Range, dv.ValidationType) |> List.sort)
   (List.sort [ "A6:A1048576", "whole"; "B2:B21 D2:D21", "list"; "C1", "list"; "E3:E10", "custom"; "F1", "none"; "A1:A5", "list" ])
eq "validations: and read back the same" ((Bridge.Import.importWorkbook editedValidations).Sheets.[0].Model.Validations |> List.map (fun v -> v.Ranges, v.Validation.Rule) |> List.sort)
   (editedModel.Validations |> List.map (fun v -> v.Ranges, v.Validation.Rule) |> List.sort)

// ---- pictures, kept objects and threaded comments are in the model ---------
// synthetic-drawing-extras.xlsx: a picture B3..D7 (two-cell, editAs oneCell), a
// chart at F (one-cell) in front of it, a threaded comment with a reply on B2.
let extrasWb = load "test-documents-excel/synthetic-drawing-extras.xlsx"
let extrasImported = Bridge.Import.importWorkbook extrasWb
// The fixture's picture is a real PNG (24 x 16), so that hosts can assert on it rendered.
let logoBytes = extrasWb.Sheets.[0].Extras.Images.[0].Data
let bigEndian (at: int) = (int logoBytes.[at] <<< 24) ||| (int logoBytes.[at + 1] <<< 16) ||| (int logoBytes.[at + 2] <<< 8) ||| int logoBytes.[at + 3]
eq "fixture picture: a PNG with a header and an end"
    (logoBytes.[0..7], System.Text.Encoding.ASCII.GetString(logoBytes, 12, 4), bigEndian 16, bigEndian 20, System.Text.Encoding.ASCII.GetString(logoBytes, logoBytes.Length - 8, 4))
    ([| 137uy; 80uy; 78uy; 71uy; 13uy; 10uy; 26uy; 10uy |], "IHDR", 24, 16, "IEND")

// ---- the PNG decoder of the native hosts ---------------------------------------
module Png = Rbt.Spreadsheet.Png
let decodedLogo = Png.tryDecode logoBytes
let pixelOf (image: Png.DecodedImage) x y = let at = (y * image.Width + x) * 4 in int image.Rgba.[at], int image.Rgba.[at + 1], int image.Rgba.[at + 2], int image.Rgba.[at + 3]
eq "png: the fixture picture decodes" (decodedLogo |> Option.map (fun image -> image.Width, image.Height, image.Rgba.Length)) (Some (24, 16, 24 * 16 * 4))
eq "png: its border, field and diagonal"
    (decodedLogo |> Option.map (fun image -> pixelOf image 0 0, pixelOf image 20 3, pixelOf image 12 8))
    (Some ((255, 255, 255, 255), (32, 96, 200, 255), (240, 140, 20, 255)))
eq "png: junk behind a PNG signature is not a picture" (Png.tryDecode [| 137uy; 80uy; 78uy; 71uy; 1uy; 2uy; 3uy; 250uy |]) None
eq "png: nothing" (Png.tryDecode [||], Png.tryDecode null) (None, None)
/// A PNG of the given colour type and bit depth from unfiltered rows (each row given without its filter byte).
let makePng (width: int) (height: int) (bitDepth: int) (colorType: int) (filter: byte) (rows: byte[] list) (extra: (string * byte[]) list) =
    let chunk (kind: string) (body: byte[]) =
        let length = [| byte (body.Length >>> 24); byte (body.Length >>> 16); byte (body.Length >>> 8); byte body.Length |]
        Array.concat [ length; System.Text.Encoding.ASCII.GetBytes kind; body; [| 0uy; 0uy; 0uy; 0uy |] ]
    let be (value: int) = [| byte (value >>> 24); byte (value >>> 16); byte (value >>> 8); byte value |]
    let raw = rows |> List.map (fun row -> Array.append [| filter |] row) |> Array.concat
    use packed = new System.IO.MemoryStream()
    (use zlib = new System.IO.Compression.ZLibStream(packed, System.IO.Compression.CompressionLevel.Optimal, true)
     zlib.Write(raw, 0, raw.Length))
    Array.concat
        [ yield [| 137uy; 80uy; 78uy; 71uy; 13uy; 10uy; 26uy; 10uy |]
          yield chunk "IHDR" (Array.concat [ be width; be height; [| byte bitDepth; byte colorType; 0uy; 0uy; 0uy |] ])
          for kind, body in extra do yield chunk kind body
          yield chunk "IDAT" (packed.ToArray())
          yield chunk "IEND" [||] ]
let pixelsOf (png: byte[]) = Png.tryDecode png |> Option.map (fun image -> [ for y in 0 .. image.Height - 1 do for x in 0 .. image.Width - 1 -> pixelOf image x y ])
eq "png: RGBA" (pixelsOf (makePng 2 1 8 6 0uy [ [| 10uy; 20uy; 30uy; 255uy; 200uy; 100uy; 50uy; 128uy |] ] [])) (Some [ (10, 20, 30, 255); (200, 100, 50, 128) ])
eq "png: greyscale with alpha" (pixelsOf (makePng 2 1 8 4 0uy [ [| 90uy; 255uy; 200uy; 0uy |] ] [])) (Some [ (90, 90, 90, 255); (200, 200, 200, 0) ])
eq "png: a palette with transparency"
    (pixelsOf (makePng 3 1 8 3 0uy [ [| 0uy; 1uy; 2uy |] ] [ "PLTE", [| 255uy; 0uy; 0uy; 0uy; 255uy; 0uy; 0uy; 0uy; 255uy |]; "tRNS", [| 255uy; 0uy |] ]))
    (Some [ (255, 0, 0, 255); (0, 255, 0, 0); (0, 0, 255, 255) ])
eq "png: a 1-bit palette" (pixelsOf (makePng 4 1 1 3 0uy [ [| 0b10100000uy |] ] [ "PLTE", [| 0uy; 0uy; 0uy; 255uy; 255uy; 255uy |] ])) (Some [ (255, 255, 255, 255); (0, 0, 0, 255); (255, 255, 255, 255); (0, 0, 0, 255) ])
eq "png: 16 bits per channel" (pixelsOf (makePng 1 1 16 2 0uy [ [| 0x12uy; 0x34uy; 0x56uy; 0x78uy; 0x9Auy; 0xBCuy |] ] [])) (Some [ (0x12, 0x56, 0x9A, 255) ])
// Row filters: the same two rows written with "up" for the second and "sub" within the rows.
eq "png: the up filter" (pixelsOf (makePng 1 2 8 0 2uy [ [| 100uy |]; [| 5uy |] ] [])) (Some [ (100, 100, 100, 255); (105, 105, 105, 255) ])
eq "png: the sub filter" (pixelsOf (makePng 3 1 8 0 1uy [ [| 10uy; 5uy; 5uy |] ] [])) (Some [ (10, 10, 10, 255); (15, 15, 15, 255); (20, 20, 20, 255) ])
eq "png: the average and Paeth filters decode to bytes"
    (pixelsOf (makePng 2 1 8 0 3uy [ [| 40uy; 10uy |] ] []), pixelsOf (makePng 2 1 8 0 4uy [ [| 40uy; 10uy |] ] []))
    (Some [ (40, 40, 40, 255); (30, 30, 30, 255) ], Some [ (40, 40, 40, 255); (50, 50, 50, 255) ])
eq "png: premultiplied BGRA for a framebuffer"
    (Png.tryDecode (makePng 2 1 8 6 0uy [ [| 10uy; 20uy; 30uy; 255uy; 200uy; 100uy; 50uy; 128uy |] ] []) |> Option.map Png.toPremultipliedBgra)
    (Some [| 0xFF0A141Eu; (128u <<< 24) ||| (100u <<< 16) ||| (50u <<< 8) ||| 25u |])
let extrasModel = extrasImported.Sheets.[0].Model
let describeDrawing (drawing: SheetDrawing) =
    (match drawing.Content with Picture (contentType, _, data) -> "picture " + contentType + " " + string data.Length | KeptObject (kind, _) -> kind),
    drawing.Name, drawing.Anchor.Kind, (drawing.Anchor.FromRow, drawing.Anchor.FromCol), (drawing.Anchor.ToRow, drawing.Anchor.ToCol), drawing.Anchor.EditAs
eq "import: drawings in z-order" (extrasModel.Drawings |> List.map describeDrawing)
    [ ("picture image/png 154", Some "Logo", "twoCell", (2, 1), (Some 6, Some 3), Some "oneCell")
      ("chart", Some "Chart 1", "oneCell", (extrasModel.Drawings.[1].Anchor.FromRow, 5), (None, None), None) ]
eq "import: anchor offsets and sizes are pixels" (extrasModel.Drawings.[0].Anchor.FromOffsetX, extrasModel.Drawings.[0].Anchor.Width, extrasModel.Drawings.[1].Anchor.Width) (1.0, Some 128.0, Some 320.0)
eq "import: threaded comment with its reply" (extrasModel.Threads |> Map.toList |> List.map (fun (a, thread) -> a.Row, a.Col, thread.Author, thread.Text, thread.Replies |> List.map (fun reply -> reply.Author, reply.Text)))
    [ 1, 1, Some "Ann Author", "Is this right?", [ Some "Bob Reply", "Yes & checked" ] ]
ok "import: the thread's cell shows a comment marker" ((extrasModel.Cells.[addr 1 1]).Format.HasComment) "no marker"
let extrasInput (model: SpreadsheetModel) = inputOf extrasImported [| model |] extrasImported.Context.Names
// Untouched: the file's drawings and threads go back exactly.
eq "export: untouched extras are the file's" ({ (Bridge.Export.toWorkbookOut (extrasInput extrasModel)).Sheets.[0].Extras with Verbatim = false }) ((XlsxWriter.packageFromReaderWorkbook extrasWb).Sheets.[0].Extras)
// Two rows inserted above everything, one column inserted before column C.
let pictureRect (model: SpreadsheetModel) = Rbt.Spreadsheet.Layout.drawingRect model model.Drawings.[0]
let movedModel = extrasModel |> applyStructureEdit (InsertRows (0, 2)) |> applyStructureEdit (InsertCols (2, 1))
eq "edit: the picture moved with its rows and kept its size (oneCell)" (describeDrawing movedModel.Drawings.[0]) ("picture image/png 154", Some "Logo", "twoCell", (4, 1), (Some 8, Some 3), Some "oneCell")
eq "edit: the chart moved down and right" ((movedModel.Drawings.[1].Anchor.FromRow, movedModel.Drawings.[1].Anchor.FromCol)) (extrasModel.Drawings.[1].Anchor.FromRow + 2, 6)
eq "edit: the thread moved with its cell" (movedModel.Threads |> Map.toList |> List.map (fun (a, _) -> a.Row, a.Col)) [ 3, 1 ]
eq "edit: the picture is drawn two rows lower" ((pictureRect movedModel).Y - (pictureRect extrasModel).Y) (2.0 * extrasModel.DefaultRowHeight)   // (two rows of the height the sheet gives a row)
let movedBack = reread (Bridge.Export.toWorkbookOut (extrasInput movedModel))
let movedExtras = movedBack.Sheets.[0].Extras
eq "export: picture re-anchored, same data" (movedExtras.Images |> List.map (fun image -> image.Anchor.FromRow, image.Anchor.FromCol, image.Anchor.ToRow, image.Anchor.ToCol, image.Anchor.FromColOffset, image.Data, image.Name))
    [ 4, 1, Some 8, Some 3, 9525.0, extrasWb.Sheets.[0].Extras.Images.[0].Data, Some "Logo" ]
eq "export: chart re-anchored, still the chart" (movedExtras.Objects |> List.map (fun item -> item.Kind, item.Name, item.Anchor.FromRow, item.Anchor.FromCol, item.Anchor.Width, item.Xml.Contains "<c:chart r:id=\"rId2\"/>"))
    [ "chart", Some "Chart 1", extrasModel.Drawings.[1].Anchor.FromRow + 2, 6, extrasWb.Sheets.[0].Extras.Objects.[0].Anchor.Width, true ]
eq "export: z-order kept" (movedExtras.Images.[0].Anchor.Order < movedExtras.Objects.[0].Anchor.Order) true
eq "export: thread at its new cell, author and reply kept" (movedExtras.ThreadedComments |> List.map (fun thread -> thread.CellRef.Column, thread.CellRef.Row, thread.Author, thread.Text, thread.Replies |> List.map (fun reply -> reply.Author, reply.Text, reply.CellRef.Row)))
    [ "B", 4, Some "Ann Author", "Is this right?", [ Some "Bob Reply", "Yes & checked", 4 ] ]
eq "export: moved source thread does not gain a legacy note" movedBack.Sheets.[0].Comments []
let insertedExtras = extrasModel |> applyStructureEdit (InsertRows (1, 1)) |> applyStructureEdit (InsertCols (1, 1))
let insertedBytes = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (extrasInput insertedExtras))
let insertedBack = XlsxReader.loadBytesNative insertedBytes |> Option.get
eq "export: inserted thread and reply keep IDs and persons"
    (insertedBack.Sheets.[0].Extras.ThreadedComments |> List.map (fun t -> t.CellRef.Column, t.CellRef.Row, t.Id, t.PersonId, t.Replies |> List.map (fun r -> r.Id, r.PersonId, r.CellRef.Column, r.CellRef.Row)), insertedBack.Preserved.Persons)
    (extrasWb.Sheets.[0].Extras.ThreadedComments |> List.map (fun t -> "C", 3, t.Id, t.PersonId, t.Replies |> List.map (fun r -> r.Id, r.PersonId, "C", 3)), extrasWb.Preserved.Persons)
ok "export: inserted thread adds no comments or VML part"
    (use stream = new MemoryStream(insertedBytes)
     use zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Read)
     zip.Entries |> Seq.forall (fun entry -> not (entry.FullName.StartsWith "xl/comments") && not (entry.FullName.EndsWith ".vml")))
    "unexpected legacy comment part"
let noteAndThread = extrasModel |> Rbt.Spreadsheet.SheetOps.setComment (addr 3 3) "real note" (Some "Note author")
let noteSource = reread (Bridge.Export.toWorkbookOut (extrasInput noteAndThread)) |> Bridge.Import.importWorkbook
let noteMoved = noteSource.Sheets.[0].Model |> applyStructureEdit (InsertRows (1, 1)) |> applyStructureEdit (InsertCols (1, 1))
let noteBack = reread (Bridge.Export.toWorkbookOut (inputOf noteSource [| noteMoved |] noteSource.Context.Names))
eq "export: real legacy note moves without creating a thread placeholder"
    (noteBack.Sheets.[0].Comments |> List.map (fun c -> c.CellRef.Column, c.CellRef.Row, c.Text, c.Author))
    [ "E", 5, "real note", Some "Note author" ]
ok "export: the chart part is still in the package" (movedBack.Preserved.Parts |> List.exists (fun (path, _) -> path.EndsWith "charts/chart1.xml")) "chart part missing"
// Deleting the rows a drawing is anchored in does not delete the drawing: it moves to the deletion.
let deletedModel = extrasModel |> applyStructureEdit (DeleteRows (1, 4))
eq "edit: deleting the anchor rows keeps the picture" (deletedModel.Drawings.Length, deletedModel.Drawings.[0].Anchor.FromRow, deletedModel.Drawings.[0].Anchor.FromOffsetY) (2, 1, 0.0)
eq "edit: deleting the thread's row removes the thread" (deletedModel.Threads.Count) 0
// A two-cell anchor that sizes with its cells grows when a row is inserted inside it.
let sizing = { extrasModel with Drawings = extrasModel.Drawings |> List.map (fun drawing -> { drawing with Anchor = { drawing.Anchor with EditAs = None } }) }
eq "edit: a move-and-size anchor grows with an inserted row" ((applyStructureEdit (InsertRows (4, 1)) sizing).Drawings.[0].Anchor |> fun anchor -> anchor.FromRow, anchor.ToRow) (2, Some 7)
// The portable painter paints the picture and a labelled frame for the chart.
let paintViewport : Rbt.Spreadsheet.Layout.Viewport = { ScrollX = 0.0; ScrollY = 0.0; Width = 1200.0; Height = 800.0; HeaderHeight = 30.0; RowHeaderWidth = 50.0; Measurer = None }
let painted = Rbt.Spreadsheet.Paint.buildPaintItems extrasModel paintViewport
eq "paint: the picture at its anchor" (painted |> List.choose (function Rbt.Spreadsheet.Paint.DrawImage (rect, image) -> Some (rect.X, rect.Y, image.ContentType, image.Data.Length, image.Label) | _ -> None))
    [ 50.0 + (pictureRect extrasModel).X, 30.0 + (pictureRect extrasModel).Y, "image/png", 154, "Picture: Logo" ]
ok "paint: a labelled frame for the chart" (painted |> List.exists (function Rbt.Spreadsheet.Paint.DrawText text -> text.Text = "Chart: Chart 1" | _ -> false)) "no chart label"
eq "paint: the viewport items scroll the drawing with the body" (Rbt.Spreadsheet.Paint.buildViewportPaintItems extrasModel { paintViewport with ScrollY = 40.0 } |> List.choose (fun item -> match item.Item with Rbt.Spreadsheet.Paint.DrawImage (rect, _) -> Some (item.Zone, rect.Y) | _ -> None))
    [ Rbt.Spreadsheet.Layout.BodyZone, 30.0 + (pictureRect extrasModel).Y - 40.0 ]
// A picture in a workbook made in the editor is written; threads too.
let pictured = { fresh with Drawings = [ extrasModel.Drawings.[0] ]; Threads = Map.ofList [ addr 0 0, { Id = "{BBBBBBBB-0000-0000-0000-000000000001}"; Author = Some "New Author"; PersonId = None; Date = None; Text = "typed here"; Replies = []; Done = false } ] }
let picturedBack = reread (Bridge.Export.toWorkbookOut { Sheets = [| pictured |]; Names = [| "First" |]; Visible = [| true |]; ActiveSheetIndex = 0; NameInfo = Map.empty; DefinedNames = Map.empty; Source = None; Macros = None; MacroProject = None })
eq "fresh: picture and thread written" (picturedBack.Sheets.[0].Extras.Images |> List.map (fun image -> image.Data.Length, image.Anchor.FromRow), picturedBack.Sheets.[0].Extras.ThreadedComments |> List.map (fun thread -> thread.Author, thread.Text))
    ([ 154, 2 ], [ Some "New Author", "typed here" ])

// ---- recalculation of an Excel-calculated workbook ---------------------------
// (milk-street-office.xlsx: 13,337 formulas, one legacy array formula.)
module Calc = Rbt.Spreadsheet.Calc
let cachedValues (workbook: Wb.Workbook) =
    [ for KeyValue(id, sheet) in workbook.Sheets do
        for KeyValue(address, data) in sheet.Model.Cells do
            match data.Value with
            | CellValue.Formula (_, cached) -> yield (id, address), cached
            | _ -> () ]
    |> Map.ofList
let cachedBefore = cachedValues imported.Context
// An edit nothing depends on is applied incrementally and changes no other value.
let freeCell : Wb.SheetCell = { Sheet = imported.Context.ActiveSheet; Cell = addr 0 60 }
match Calc.trySingleCellEdit freeCell { Value = CellValue.String "x"; Format = defaultFormat; Comment = None } imported.Context None with
| Calc.IncrementalFallback reason -> ok "an edit without dependents is incremental" false reason
| Calc.IncrementalApplied (edited, _) ->
    ok "an edit without dependents is incremental" true ""
    eq "and changes no cached value" (cachedValues edited |> Map.filter (fun key value -> Map.tryFind key cachedBefore <> Some value) |> Map.count) 0
// A full recalculation agrees with Excel's cached values to the last digits.
let recalculated = cachedValues (Rbt.Spreadsheet.Formula.recalculateWorkbook imported.Context)
let differing =
    cachedBefore
    |> Map.toList
    |> List.choose (fun (key, before) -> match Map.tryFind key recalculated with | Some after when after <> before -> Some (key, before, after) | _ -> None)
let relative (a: CellValue) (b: CellValue) = match a, b with | CellValue.Number x, CellValue.Number y when x <> 0.0 -> abs ((x - y) / x) | _ -> infinity
eq "full recalculation: nothing differs beyond the 13th digit" (differing |> List.filter (fun (_, a, b) -> not (relative a b < 1e-12)) |> List.map (fun ((sheet, cell), a, b) -> sheet, cell.Row, cell.Col, a, b) |> List.truncate 3) []
ok "full recalculation: at most 40 of 13,337 values differ in their last digits" (differing.Length <= 40) (sprintf "%d differ" differing.Length)
// The legacy array formula {=TRANSPOSE(C71:C73)} in F94:H94 is recalculated, not blocked by its own cells.
let arrayAnchor = imported.Sheets.[0].Model.Cells |> Map.tryPick (fun address data -> match data.Value with CellValue.Formula (text, _) when text.StartsWith "TRANSPOSE" -> Some address | _ -> None)
match arrayAnchor with
| None -> ok "array formula found" false "no TRANSPOSE formula on the first sheet"
| Some anchor ->
    let sheetId = Bridge.sheetIdOfIndex 0
    eq "array formula: its other cells belong to it" (imported.Sheets.[0].Model.Spill |> Map.toList |> List.filter (fun (_, owner) -> owner = anchor) |> List.map (fun (ghost, _) -> ghost.Row - anchor.Row, ghost.Col - anchor.Col)) [ 0, 1; 0, 2 ]
    eq "array formula: same values after a full recalculation" ([ 0; 1; 2 ] |> List.map (fun offset -> Map.tryFind (sheetId, addr anchor.Row (anchor.Col + offset)) recalculated |> Option.orElse (getCell (addr anchor.Row (anchor.Col + offset)) (Rbt.Spreadsheet.Formula.recalculateWorkbook imported.Context).Sheets.[sheetId].Model |> Option.map (fun data -> data.Value))) |> List.map (function Some (CellValue.Error e) -> "error " + e | Some _ -> "value" | None -> "missing")) [ "value"; "value"; "value" ]
    eq "array formula: anchor keeps Excel's value" (Map.tryFind (sheetId, anchor) recalculated) (Map.tryFind (sheetId, anchor) cachedBefore)

// ---- import options ----------------------------------------------------------
let roomy = Bridge.Import.importWorkbookWith Bridge.editorImportOptions wb
ok "editor options leave room to grow" (roomy.Sheets.[0].Model.RowCount >= imported.Sheets.[0].Model.RowCount + 100 && roomy.Sheets.[0].Model.ColCount >= imported.Sheets.[0].Model.ColCount + 20) "model is not larger"
eq "the cells are the same either way" (roomy.Sheets.[0].Model.Cells = imported.Sheets.[0].Model.Cells) true

// ---- the list of sheets: tab colours, hidden and very hidden, sheets renamed / moved / deleted / copied ----
let tabbed =
    [| { fresh with TabColor = Some "#FF8800" }; { createEmpty 10 5 with VeryHidden = true }; { createEmpty 10 5 with TabColor = Some "theme:4:-0.249977111117893" }
       { createEmpty 10 5 with TabColor = Some "indexed:10"; VeryHidden = true } |]
let tabbedInput (sheets: SpreadsheetModel array) (names: string array) (visible: bool array) (source: Bridge.XlsxSource option) : Bridge.ExportInput =
    { Sheets = sheets; Names = names; Visible = visible; ActiveSheetIndex = 0; NameInfo = Map.empty; DefinedNames = Map.empty; Source = source; Macros = None; MacroProject = None }
let tabbedBytes = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (tabbedInput tabbed [| "One"; "Two"; "Three"; "Four" |] [| true; false; true; true |] None))
let tabbedBack = XlsxReader.loadBytesNative tabbedBytes |> Option.get
eq "sheets: tab colours and states are written"
    (tabbedBack.Sheets |> List.map (fun s -> s.Name, s.State, s.View.TabColor))
    [ "One", XlsxReader.SheetVisible, Some "#FF8800"; "Two", XlsxReader.SheetVeryHidden, None; "Three", XlsxReader.SheetVisible, Some "theme:4:-0.249977111117893"
      // (A sheet with a tab is not very hidden, whatever its model says.)
      "Four", XlsxReader.SheetVisible, Some "indexed:10" ]
let tabbedImport = Bridge.Import.importWorkbook tabbedBack
eq "sheets: ... and read" (tabbedImport.Sheets |> Array.map (fun s -> s.Visible, s.Model.VeryHidden, s.Model.TabColor))
    [| true, false, Some "#FF8800"; false, true, None; true, false, Some "theme:4:-0.249977111117893"; true, false, Some "indexed:10" |]
let tabbedLazy = Bridge.Import.importWorkbookLazy (fun index -> index = 0) tabbedBack
eq "sheets: a sheet that is not converted yet knows its colour and whether it is very hidden"
    (tabbedLazy.Sheets |> Array.map (fun s -> isDeferredSheet s.Model, s.Model.VeryHidden, s.Model.TabColor))
    [| false, false, Some "#FF8800"; true, true, None; true, false, Some "theme:4:-0.249977111117893"; true, false, Some "indexed:10" |]
// Colour and visibility change on the stand-in; the sheet is written without being converted.
let recoloured =
    tabbedLazy.Sheets |> Array.mapi (fun index s ->
        match index with
        | 1 -> { s.Model with VeryHidden = false }
        | 2 -> { s.Model with TabColor = Some "#112233" }
        | 3 -> { s.Model with TabColor = None }
        | _ -> s.Model)
let recolouredBack =
    XlsxReader.loadBytesNative (XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (tabbedInput recoloured [| "One"; "Renamed"; "Three"; "Four" |] [| true; false; true; false |] (Some tabbedLazy.Source)))) |> Option.get
eq "sheets: changed without converting"
    (recolouredBack.Sheets |> List.map (fun s -> s.Name, s.State, s.View.TabColor), recoloured |> Array.map isDeferredSheet)
    ([ "One", XlsxReader.SheetVisible, Some "#FF8800"; "Renamed", XlsxReader.SheetHidden, None; "Three", XlsxReader.SheetVisible, Some "#112233"; "Four", XlsxReader.SheetHidden, None ],
     [| false; true; true; true |])
// Renamed, moved, one deleted, one copied: through the workbook operations, then the file.
let listed =
    let models = [| fresh |> setCell (addr 5 0) { Value = CellValue.Formula ("=Second!A1+1", CellValue.Number 8.0); Format = defaultFormat; Comment = None }
                    createEmpty 10 5 |> setCell (addr 0 0) { Value = CellValue.Number 7.0; Format = defaultFormat; Comment = None }
                                     |> setCell (addr 1 0) { Value = CellValue.Formula ("=First!A2*2", CellValue.Number 6.0); Format = defaultFormat; Comment = None }
                    createEmpty 10 5 |> setCell (addr 0 0) { Value = CellValue.Formula ("=Second!A1+First!A2", CellValue.Number 10.0); Format = defaultFormat; Comment = None } |]
    Bridge.buildWorkbookContext models [| "First"; "Second"; "Third" |] [| true; true; true |] Map.empty Map.empty 0
    |> Wb.setDefinedName None "Seven" "Second!$A$1" None
let relisted =
    listed
    |> Wb.renameSheetEverywhere "sheet2" "Number 2"
    |> Wb.moveSheet "sheet3" 0
    |> Wb.duplicateSheet "sheet1" "sheet4" "First (2)" 2
    |> fun wb -> fst (Wb.removeSheet "sheet1" wb).Value
    |> Wb.setSheetVisibility "sheet3" false true
let relistedOrder = List.toArray relisted.SheetOrder
let relistedBack =
    XlsxReader.loadBytesNative (XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut
        { Sheets = relistedOrder |> Array.map (fun id -> relisted.Sheets.[id].Model); Names = relistedOrder |> Array.map (fun id -> relisted.Sheets.[id].Name)
          Visible = relistedOrder |> Array.map (fun id -> relisted.Sheets.[id].Visible); ActiveSheetIndex = 1; NameInfo = relisted.NameInfo; DefinedNames = relisted.Names; Source = None; Macros = None; MacroProject = None })) |> Option.get
eq "sheets: the list after the operations" (relistedBack.Sheets |> List.map (fun s -> s.Name, s.State)) [ "Third", XlsxReader.SheetVeryHidden; "First (2)", XlsxReader.SheetVisible; "Number 2", XlsxReader.SheetVisible ]
let formulasOf (sheet: XlsxReader.Worksheet) = sheet.Rows |> List.collect (fun r -> r.Cells |> List.choose (fun c -> c.Formula))
eq "sheets: formulas name the sheets as they are now" (relistedBack.Sheets |> List.map formulasOf) [ [ "'Number 2'!A1+#REF!" ]; [ "A2*2"; "'Number 2'!A1+1" ]; [ "#REF!*2" ] ]
eq "sheets: so do defined names" (relistedBack.NamedRanges |> List.map (fun n -> n.Name.ToUpperInvariant(), n.RefersTo)) [ "SEVEN", "'Number 2'!$A$1" ]

// ---- rows and columns: hidden ones keep their size, outline levels, collapsed groups ----
module Ops = Rbt.Spreadsheet.SheetOps
let axesModel =
    fresh
    |> Ops.setAxisSize true 2 2 40.0 |> Ops.setAxisSize false 1 1 140.0
    |> Ops.hideAxis true 2 3 |> Ops.hideAxis false 1 2
    |> Ops.changeOutlineLevel true 5 8 1 |> Ops.changeOutlineLevel true 6 7 1 |> Ops.setGroupCollapsed "test" true 6 true
    |> Ops.changeOutlineLevel false 4 5 1 |> Ops.setGroupCollapsed "test" false 4 true
let axesInput (model: SpreadsheetModel) (source: Bridge.XlsxSource option) : Bridge.ExportInput =
    { Sheets = [| model |]; Names = [| "Axes" |]; Visible = [| true |]; ActiveSheetIndex = 0; NameInfo = Map.empty; DefinedNames = Map.empty; Source = source; Macros = None; MacroProject = None }
let axesFile = XlsxReader.loadBytesNative (XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (axesInput axesModel None))) |> Option.get
let rowAttributes (wb: XlsxReader.Workbook) =
    wb.Sheets.[0].Rows |> List.filter (fun r -> r.Hidden || r.OutlineLevel > 0 || r.Collapsed) |> List.map (fun r -> r.Index, r.Hidden, r.Height |> Option.map (fun h -> Math.Round (h, 2)), r.OutlineLevel, r.Collapsed)
let columnAttributes (wb: XlsxReader.Workbook) =
    wb.Sheets.[0].Columns |> List.filter (fun c -> c.Hidden || c.OutlineLevel > 0 || c.Collapsed) |> List.map (fun c -> c.Index, c.MaxIndex, c.Hidden, c.Width |> Option.map (fun w -> Math.Round (w, 2)), c.OutlineLevel, c.Collapsed)
eq "axes: hidden rows keep their height in the file; levels and collapsed marks are written" (rowAttributes axesFile)
    [ 3, true, Some 30.0, 0, false; 4, true, None, 0, false; 6, false, None, 1, false; 7, true, None, 2, false; 8, true, None, 2, false; 9, false, None, 1, true ]
eq "axes: columns likewise" (columnAttributes axesFile) [ 2, 2, true, Some 20.0, 0, false; 3, 3, true, None, 0, false; 5, 6, true, None, 1, false; 7, 7, false, None, 0, true ]
let axesImported = Bridge.Import.importWorkbook axesFile
let axesAgain = axesImported.Sheets.[0].Model
let roundedSizes (sizes: Map<int, float>) = sizes |> Map.map (fun _ size -> Math.Round size)
eq "axes: read back: what is hidden, the levels, the marks" (Ops.hiddenRuns true axesAgain, Ops.hiddenRuns false axesAgain, axesAgain.Axes.RowLevels, axesAgain.Axes.ColLevels, axesAgain.Axes.CollapsedRows, axesAgain.Axes.CollapsedCols)
    (Ops.hiddenRuns true axesModel, Ops.hiddenRuns false axesModel, axesModel.Axes.RowLevels, axesModel.Axes.ColLevels, axesModel.Axes.CollapsedRows, axesModel.Axes.CollapsedCols)
eq "axes: a hidden row / column with a size of its own comes back at it" (Ops.shownAxisSize true 2 axesAgain |> Math.Round, Ops.shownAxisSize false 1 axesAgain |> Math.Round) (40.0, 140.0)
let axesShown = axesAgain |> Ops.unhideAxis true 0 99 |> Ops.unhideAxis false 0 99
eq "axes: shown again after the import" (Math.Round axesShown.RowHeights.[2], Math.Round axesShown.ColumnWidths.[1], Ops.hiddenRuns true axesShown, Ops.hiddenRuns false axesShown) (40.0, 140.0, [], [])
// With the file as the source: untouched, then edited.
let axesSame = XlsxReader.loadBytesNative (XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (axesInput axesAgain (Some axesImported.Source)))) |> Option.get
eq "axes: an untouched sheet of a file is written as the file had it" (rowAttributes axesSame, columnAttributes axesSame) (rowAttributes axesFile, columnAttributes axesFile)
let axesEdited = axesAgain |> Ops.unhideAxis true 2 2 |> Ops.setGroupCollapsed "test" true 6 false |> Ops.changeOutlineLevel false 4 5 -1 |> Ops.hideAxis false 7 7
let axesEditedFile = XlsxReader.loadBytesNative (XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (axesInput axesEdited (Some axesImported.Source)))) |> Option.get
eq "axes: edited rows of a file's sheet" (rowAttributes axesEditedFile)
    // (A row of the file that was hidden without a height of its own still states none: the rows are the file's with what changed.)
    [ 4, true, None, 0, false; 6, false, None, 1, false; 7, false, None, 2, false; 8, false, None, 2, false; 9, false, None, 1, false ]
eq "axes: the row shown again has its height" (axesEditedFile.Sheets.[0].Rows |> List.tryFind (fun r -> r.Index = 3) |> Option.bind (fun r -> r.Height) |> Option.map (fun h -> Math.Round (h * 1.333))) (Some 40.0)
eq "axes: edited columns" (columnAttributes axesEditedFile |> List.map (fun (first, last, hidden, _, level, collapsed) -> first, last, hidden, level, collapsed)) [ 2, 2, true, 0, false; 3, 3, true, 0, false; 5, 6, true, 0, false; 8, 8, true, 0, false ]

// The outline's summary rows above / columns to the left (sheetPr outlinePr): the collapsed mark is on the row before the group.
let summaryModel =
    fresh |> Ops.setOutlineSummary false false
    |> Ops.changeOutlineLevel true 5 8 1 |> Ops.setGroupCollapsed "test" true 6 true
    |> Ops.changeOutlineLevel false 4 5 1 |> Ops.setGroupCollapsed "test" false 4 true
let summaryBytes = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (axesInput summaryModel None))
let summaryFile = XlsxReader.loadBytesNative summaryBytes |> Option.get
let sheetXmlOf (bytes: byte array) =
    use zip = new System.IO.Compression.ZipArchive(new MemoryStream(bytes))
    use reader = new StreamReader(zip.GetEntry("xl/worksheets/sheet1.xml").Open())
    reader.ReadToEnd()
ok "outline: outlinePr is written before the dimension" ((sheetXmlOf summaryBytes).Contains "<sheetPr><outlinePr summaryBelow=\"0\" summaryRight=\"0\"/></sheetPr><dimension") (sheetXmlOf summaryBytes |> fun x -> x.Substring(0, min 400 x.Length))
ok "outline: no outlinePr for Excel's default" (not ((sheetXmlOf (XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (axesInput axesModel None)))).Contains "outlinePr")) ""
// (The outline's symbols can be switched off in <outlinePr> too, not only in the sheet view.)
do
    let patched =
        use source = new System.IO.Compression.ZipArchive(new MemoryStream(summaryBytes))
        use target = new MemoryStream()
        (use out = new System.IO.Compression.ZipArchive(target, System.IO.Compression.ZipArchiveMode.Create, true)
         for entry in source.Entries do
            use reader = new StreamReader(entry.Open())
            let text = reader.ReadToEnd()
            let text = if entry.FullName = "xl/worksheets/sheet1.xml" then text.Replace("<outlinePr summaryBelow=\"0\"", "<outlinePr showOutlineSymbols=\"0\" summaryBelow=\"0\"") else text
            use writer = new StreamWriter(out.CreateEntry(entry.FullName).Open())
            writer.Write text)
        target.ToArray()
    ok "outline: the patched part says so" ((sheetXmlOf patched).Contains "showOutlineSymbols=\"0\"") ""
    eq "outline: showOutlineSymbols of outlinePr is read (XML reader)" ((XlsxReader.loadBytesNative patched |> Option.get).Sheets.[0].View.ShowOutlineSymbols) false
    eq "outline: showOutlineSymbols of outlinePr is read (text scanner)" ((XlsxReader.loadBytesNativeScanned patched |> Option.get).Sheets.[0].View.ShowOutlineSymbols) false
    eq "outline: and the sheet shows no symbols" (Bridge.Import.importWorkbook (XlsxReader.loadBytesNative patched |> Option.get)).Sheets.[0].Model.Axes.ShowSymbols false
    eq "outline: symbols are shown otherwise" (summaryFile.Sheets.[0].View.ShowOutlineSymbols) true
eq "outline: the file says where the summaries are" (summaryFile.Sheets.[0].View.SummaryBelow, summaryFile.Sheets.[0].View.SummaryRight) (false, false)
eq "outline: summary above: the mark is on the row before the group" (rowAttributes summaryFile) [ 5, false, None, 0, true; 6, true, None, 1, false; 7, true, None, 1, false; 8, true, None, 1, false; 9, true, None, 1, false ]
eq "outline: summary left: on the column before" (columnAttributes summaryFile |> List.map (fun (first, last, hidden, _, level, collapsed) -> first, last, hidden, level, collapsed)) [ 4, 4, false, 0, true; 5, 6, true, 1, false ]
let summaryImported = Bridge.Import.importWorkbook summaryFile
let summaryAgain = summaryImported.Sheets.[0].Model
eq "outline: read back" (summaryAgain.Axes.SummaryBelow, summaryAgain.Axes.SummaryRight, Ops.outlineGroups true summaryAgain, Ops.outlineGroups false summaryAgain, summaryAgain.Axes.CollapsedRows |> Set.toList)
    (false, false, [ 5, 8, 1, true ], [ 4, 5, 1, true ], [ 4 ])
let summaryExpanded = summaryAgain |> Ops.setGroupCollapsed "test" true 6 false |> Ops.setOutlineSummary true false
let summaryEditedBytes = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (axesInput summaryExpanded (Some summaryImported.Source)))
let summaryEdited = XlsxReader.loadBytesNative summaryEditedBytes |> Option.get
eq "outline: edited in a file's sheet: expanded, summary below again" (summaryEdited.Sheets.[0].View.SummaryBelow, summaryEdited.Sheets.[0].View.SummaryRight, rowAttributes summaryEdited |> List.map (fun (i, hidden, _, level, collapsed) -> i, hidden, level, collapsed))
    (true, false, [ 6, false, 1, false; 7, false, 1, false; 8, false, 1, false; 9, false, 1, false ])

// ---- the AutoFilter: its range, every kind of column filter, the rows it hides, the hidden name ----
let filterCell value = { Value = value; Format = defaultFormat; Comment = None }
let filterData =
    [ 0, "Region", "Sales"; 1, "East", "10"; 2, "West", "20"; 3, "East", "30"; 4, "North", "40"; 5, "West", "50"; 6, "South", "60" ]
    |> List.fold (fun m (row, region, sales) ->
        m
        |> setCell (addr row 1) (filterCell (CellValue.String region))
        |> setCell (addr row 2) (filterCell (match Double.TryParse sales with | true, n -> CellValue.Number n | _ -> CellValue.String sales))
        |> setCell (addr row 3) (filterCell (CellValue.String (sprintf "note %d" row)))
        |> setCell (addr row 4) { filterCell (CellValue.Number (float row)) with Format = { defaultFormat with BgColor = (if row % 2 = 0 then Some "#FFFF00" else None); TextColor = (if row = 3 then Some "#FF0000" else None) } }) (createEmpty 40 8)
    |> Ops.setAxisSize true 2 2 44.0
    |> Ops.hideAxis true 12 12
let filterRange = area 0 1 6 4
let filterKinds =
    [ "values", [ 1, FilterValues (Set.ofList [ "East"; "West" ], true) ]
      "one condition", [ 2, FilterCustom ([ OpGreaterThan, "25" ], true) ]
      "two conditions (or)", [ 2, FilterCustom ([ OpLessThan, "20"; OpGreaterOrEqual, "50" ], false) ]
      "two conditions (and)", [ 2, FilterCustom ([ OpGreaterOrEqual, "20"; OpNotEqual, "40" ], true) ]
      "text with wildcards", [ 3, FilterCustom ([ OpEqual, "*e 3*" ], true) ]
      "top", [ 2, FilterTop (2, false, false) ]
      "bottom percent", [ 2, FilterTop (40, true, true) ]
      "fill colour", [ 4, FilterColor ("#FFFF00", false) ]
      "text colour", [ 4, FilterColor ("#FF0000", true) ]
      "above average", [ 2, FilterDynamic "aboveAverage" ]
      "below average", [ 2, FilterDynamic "belowAverage" ]
      "two columns", [ 1, FilterValues (Set.ofList [ "East" ], false); 2, FilterCustom ([ OpGreaterThan, "10" ], true) ] ]
for label, columns in filterKinds do
    let model = columns |> List.fold (fun m (col, criteria) -> setColumnFilter col criteria m) (setAutoFilter (Some filterRange) filterData)
    let file = reread (Bridge.Export.toWorkbookOut (axesInput model None))
    let back = (Bridge.Import.importWorkbook file).Sheets.[0].Model
    eq ("filter " + label + ": hides rows at all") (model.FilterHiddenRows.Length > 0) true
    eq ("filter " + label + ": range and criteria come back") (back.AutoFilter, back.Filters) (model.AutoFilter, model.Filters)
    eq ("filter " + label + ": the same rows are hidden") back.FilterHiddenRows model.FilterHiddenRows
    eq ("filter " + label + ": the file hides exactly those rows (and the one the user hid)") (file.Sheets.[0].Rows |> List.filter (fun r -> r.Hidden) |> List.map (fun r -> r.Index - 1)) (List.ofArray model.FilterHiddenRows @ [ 12 ])
    eq ("filter " + label + ": the user's hidden row is still the user's") (Ops.hiddenRuns true back) [ 12, 12 ]
    eq ("filter " + label + ": applying it again changes nothing") (reapplyFilter back).FilterHiddenRows back.FilterHiddenRows
    let shown = setAutoFilter None back
    eq ("filter " + label + ": shown again, the rows have their heights") ([ for row in model.FilterHiddenRows -> row, Math.Round (getRowHeight row shown) ], getRowHeight 12 shown) ([ for row in model.FilterHiddenRows -> row, (if row = 2 then 44.0 else 25.0) ], 0.0)
let valuesModel = setAutoFilter (Some filterRange) filterData |> setColumnFilter 1 (FilterValues (Set.ofList [ "East"; "West" ], false))
let valuesOut = Bridge.Export.toWorkbookOut (axesInput valuesModel None)
let valuesBytes = XlsxPackageWriter.writeWorkbookPackageBytes valuesOut
let partText (bytes: byte[]) (path: string) =
    use archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes), System.IO.Compression.ZipArchiveMode.Read)
    use reader = new StreamReader(archive.GetEntry(path).Open())
    reader.ReadToEnd()
let valuesSheetXml = partText valuesBytes "xl/worksheets/sheet1.xml"
ok "filter: the element as Excel writes it" (valuesSheetXml.Contains "<autoFilter ref=\"B1:E7\"><filterColumn colId=\"0\"><filters><filter val=\"East\"/><filter val=\"West\"/></filters></filterColumn></autoFilter>") valuesSheetXml
ok "filter: it comes after the rows and before the merges / page setup" (valuesSheetXml.IndexOf "<autoFilter" > valuesSheetXml.IndexOf "</sheetData>" && valuesSheetXml.IndexOf "<autoFilter" < valuesSheetXml.IndexOf "<pageMargins") "order"
ok "filter: the hidden name of the sheet" ((partText valuesBytes "xl/workbook.xml").Contains "<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"0\" hidden=\"1\">'Axes'!$B$1:$E$7</definedName>") (partText valuesBytes "xl/workbook.xml")
let valuesFile = XlsxReader.loadBytesNative valuesBytes |> Option.get
// A filter with no column filtered: the range alone.
let rangeOnly = reread (Bridge.Export.toWorkbookOut (axesInput (setAutoFilter (Some filterRange) filterData) None))
eq "filter: a range without column filters" ((Bridge.Import.importWorkbook rangeOnly).Sheets.[0].Model.AutoFilter, rangeOnly.Sheets.[0].Extras.AutoFilter) (Some { Range = filterRange; Implicit = false }, Some { Ref = "B1:E7"; Columns = [] })
// With the file as the source: untouched it goes back as the file states it; edited it is written from the model.
let valuesImported = Bridge.Import.importWorkbook valuesFile
let valuesAgain = valuesImported.Sheets.[0].Model
let valuesSame = reread (Bridge.Export.toWorkbookOut (axesInput valuesAgain (Some valuesImported.Source)))
eq "filter: an untouched sheet of a file keeps its filter and hidden rows" (valuesSame.Sheets.[0].Extras.AutoFilter, valuesSame.Sheets.[0].Rows |> List.filter (fun r -> r.Hidden) |> List.map (fun r -> r.Index - 1)) (valuesFile.Sheets.[0].Extras.AutoFilter, [ 4; 6; 12 ])
let valuesLazy = (Bridge.Import.importWorkbookLazy (fun _ -> false) valuesFile).Sheets.[0].Model |> resolveSheet
eq "filter: a sheet converted on demand has it too" (valuesLazy.AutoFilter, List.ofArray valuesLazy.FilterHiddenRows) (valuesModel.AutoFilter, [ 4; 6 ])
let refiltered = valuesAgain |> setColumnFilter 2 (FilterTop (1, false, false))
let refilteredFile = reread (Bridge.Export.toWorkbookOut (axesInput refiltered (Some valuesImported.Source)))
eq "filter: another column filtered on a file's sheet" (refilteredFile.Sheets.[0].Extras.AutoFilter |> Option.map (fun f -> f.Columns |> List.map (fun c -> c.ColId, c.Kind)), refilteredFile.Sheets.[0].Rows |> List.filter (fun r -> r.Hidden) |> List.map (fun r -> r.Index - 1))
    (Some [ 0, "filters"; 1, "top10" ], [ 1; 2; 3; 4; 5; 6; 12 ])
let unfiltered = reread (Bridge.Export.toWorkbookOut (axesInput (setAutoFilter None valuesAgain) (Some valuesImported.Source)))
eq "filter: removed from a file's sheet, its rows are shown" (unfiltered.Sheets.[0].Extras.AutoFilter, unfiltered.Sheets.[0].Rows |> List.filter (fun r -> r.Hidden) |> List.map (fun r -> r.Index - 1), (partText (XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (axesInput (setAutoFilter None valuesAgain) (Some valuesImported.Source)))) "xl/workbook.xml").Contains "_FilterDatabase") (None, [ 12 ], false)
// A filter the editor does not evaluate (date groups) is kept: its XML, and the rows the file hid.
let keptColumn = { XlsxReader.emptyAutoFilterColumn with ColId = 1; Kind = "other"; Xml = "<filters><dateGroupItem year=\"2024\" month=\"3\" dateTimeGrouping=\"month\"/></filters>" }
let keptOut =
    let out = Bridge.Export.toWorkbookOut (axesInput (filterData |> Ops.setAxisSize true 3 3 0.0) None)
    { out with Sheets = out.Sheets |> List.map (fun sheet -> { sheet with Extras = { sheet.Extras with AutoFilter = Some { Ref = "B1:E7"; Columns = [ keptColumn ] } } }) }
let keptFile = reread keptOut
eq "kept filter: read as what it is" (keptFile.Sheets.[0].Extras.AutoFilter) (Some { Ref = "B1:E7"; Columns = [ keptColumn ] })
let keptImported = Bridge.Import.importWorkbook keptFile
let keptModel = keptImported.Sheets.[0].Model
eq "kept filter: in the model, with the rows the file hid" (keptModel.Filters |> Map.toList |> List.map (fun (col, f) -> col, f.Criteria), List.ofArray keptModel.FilterHiddenRows) ([ 2, FilterKept keptColumn.Xml ], [ 3 ])
eq "kept filter: applying the filters again leaves its rows hidden" (List.ofArray (reapplyFilter keptModel).FilterHiddenRows, List.ofArray (keptModel |> setColumnFilter 1 (FilterValues (Set.ofList [ "West" ], false))).FilterHiddenRows) ([ 3 ], [ 1; 3; 4; 6 ])
let keptEdited = reread (Bridge.Export.toWorkbookOut (axesInput (keptModel |> setColumnFilter 1 (FilterValues (Set.ofList [ "West" ], false))) (Some keptImported.Source)))
eq "kept filter: written back beside a new one" (keptEdited.Sheets.[0].Extras.AutoFilter |> Option.map (fun f -> f.Columns |> List.map (fun c -> c.ColId, c.Kind, c.Xml.Contains "dateGroupItem"))) (Some [ 0, "filters", false; 1, "other", true ])

// ---- comments edited through the host API (SheetOps) --------------------------
module Ops = Rbt.Spreadsheet.SheetOps
let okOf (result: Result<'a, string>) = match result with Ok value -> value | Result.Error message -> failwith message
let singleInput (model: SpreadsheetModel) (source: Bridge.XlsxSource option) : Bridge.ExportInput =
    { Sheets = [| model |]; Names = [| "First" |]; Visible = [| true |]; ActiveSheetIndex = 0; NameInfo = Map.empty; DefinedNames = Map.empty; Source = source; Macros = None; MacroProject = None }
let commentsOfModel (model: SpreadsheetModel) =
    Ops.allComments model |> List.map (fun c -> c.Cell.Row, c.Cell.Col, c.Text, c.Author, c.Thread |> Option.map (fun t -> t.Done, t.Replies |> List.map (fun r -> r.Text, r.Author, r.Date)))
// A fresh workbook: a note, and a note turned into a thread by a reply, resolved.
let commented =
    fresh
    |> Ops.setComment (addr 0 0) "a note" (Some "Ann")
    |> Ops.setComment (addr 6 2) "to discuss" (Some "Ann")
    |> fun m -> Ops.addCommentReply (addr 6 2) "agreed" (Some "Bob") (Some "2026-01-02T03:04:05.00") m |> okOf |> fst
    |> fun m -> Ops.setCommentResolved (addr 6 2) true m |> okOf
let commentedBack = reread (Bridge.Export.toWorkbookOut (singleInput commented None))
eq "comments: the thread is written with its reply, author persons and the resolved mark"
    (commentedBack.Sheets.[0].Extras.ThreadedComments |> List.map (fun t -> t.CellRef.Column, t.CellRef.Row, t.Text, t.Author, t.Done, t.Replies |> List.map (fun r -> r.Text, r.Author, r.Date)))
    [ "C", 7, "to discuss", Some "Ann", true, [ "agreed", Some "Bob", Some "2026-01-02T03:04:05.00" ] ]
eq "comments: the note, and a note beside the thread for applications without threads"
    (commentedBack.Sheets.[0].Comments |> List.map (fun c -> c.CellRef.Column, c.CellRef.Row, c.Text.StartsWith "[Threaded comment]" || c.Text = "a note", (defaultArg c.Author "").StartsWith "tc={" || c.Author = Some "Ann") |> List.sort)
    [ "A", 1, true, true; "C", 7, true, true ]
let commentedAgain = (Bridge.Import.importWorkbook commentedBack).Sheets.[0].Model
eq "comments: read back as they were" (commentsOfModel commentedAgain) (commentsOfModel commented)
// The re-imported sheet edited again: note deleted, reply removed, thread reopened.
let replyIdAgain = commentedAgain.Threads.[addr 6 2].Replies.[0].Id
let recommented =
    commentedAgain
    |> Ops.deleteComment (addr 0 0)
    |> fun m -> Ops.deleteCommentReply (addr 6 2) replyIdAgain m |> okOf
    |> fun m -> Ops.setCommentResolved (addr 6 2) false m |> okOf
    |> Ops.setComment (addr 6 2) "discussed" None
let recommentedBack = reread (Bridge.Export.toWorkbookOut (singleInput recommented (Some (Bridge.Import.importWorkbook commentedBack).Source)))
eq "comments: edits of an imported sheet are written" (commentsOfModel (Bridge.Import.importWorkbook recommentedBack).Sheets.[0].Model)
    [ 6, 2, "discussed", Some "Ann", Some (false, []) ]
eq "comments: the deleted note is not in the file" (recommentedBack.Sheets.[0].Comments |> List.map (fun c -> c.CellRef.Column, c.CellRef.Row)) [ "C", 7 ]
// Deleting the last thread drops the threaded comments of the sheet.
let uncommentedBack = reread (Bridge.Export.toWorkbookOut (singleInput (recommented |> Ops.deleteComment (addr 6 2)) (Some (Bridge.Import.importWorkbook commentedBack).Source)))
eq "comments: none left" (uncommentedBack.Sheets.[0].Comments.Length, uncommentedBack.Sheets.[0].Extras.ThreadedComments.Length) (0, 0)

// ---- hyperlinks set through the host API (SheetOps) ----------------------------
let cellRange a b c d : CellRange = { From = addr a b; To = addr c d }
let linksOfModel (model: SpreadsheetModel) = model.Hyperlinks |> List.map (fun l -> l.Range, l.Target, l.IsInternal, l.Tooltip, l.Display)
let linkedModel =
    fresh
    |> Ops.setHyperlink [ cellRange 0 0 0 0 ] "https://example.com/a?x=1&y=2#part" false None (Some "the \"site\" & more")
    |> Ops.setHyperlink [ cellRange 6 1 7 2 ] "'Hidden one'!A1" true (Some "go there") None
    |> Ops.setHyperlink [ cellRange 9 0 9 0 ] "mailto:someone@example.com" false (Some "write") (Some "mail")
let linkedInput (model: SpreadsheetModel) (source: Bridge.XlsxSource option) : Bridge.ExportInput =
    { Sheets = [| model; createEmpty 10 5 |]; Names = [| "First"; "Hidden one" |]; Visible = [| true; true |]; ActiveSheetIndex = 0; NameInfo = Map.empty; DefinedNames = Map.empty; Source = source; Macros = None; MacroProject = None }
let linkedBack = reread (Bridge.Export.toWorkbookOut (linkedInput linkedModel None))
eq "links: written with target, tooltip and display"
    (linkedBack.Sheets.[0].Hyperlinks |> List.map (fun l -> l.Range, l.Target, l.IsInternal, l.Tooltip, l.Display))
    [ "A1", "https://example.com/a?x=1&y=2#part", false, Some "the \"site\" & more", None
      "B7:C8", "'Hidden one'!A1", true, None, Some "go there"
      "A10", "mailto:someone@example.com", false, Some "mail", Some "write" ]
let linkedImported = Bridge.Import.importWorkbook linkedBack
let linkedAgain = linkedImported.Sheets.[0].Model
eq "links: read back as they were" (linksOfModel linkedAgain) (linksOfModel linkedModel)
eq "links: the display text is the cell's text, with the look of a link" (getCell (addr 6 1) linkedAgain |> Option.map (fun d -> d.Value, d.Format.TextColor, d.Format.Underline)) (Some (CellValue.String "go there", Some "#0563C1", true))
// Edited in the imported workbook: one removed, one re-pointed, one untouched.
let relinked = linkedAgain |> Ops.removeHyperlinks [ cellRange 9 0 9 0 ] |> fst |> Ops.setHyperlink [ cellRange 6 1 6 1 ] "tel:+15550100" false None None
let relinkedBack = reread (Bridge.Export.toWorkbookOut (linkedInput relinked (Some linkedImported.Source)))
eq "links: edits of an imported sheet are written"
    (relinkedBack.Sheets.[0].Hyperlinks |> List.map (fun l -> l.Range, l.Target, l.IsInternal) |> List.sort)
    ([ "A1", "https://example.com/a?x=1&y=2#part", false; "C7", "'Hidden one'!A1", true; "B8:C8", "'Hidden one'!A1", true; "B7", "tel:+15550100", false ] |> List.sort)
eq "links: none left when all are removed" ((reread (Bridge.Export.toWorkbookOut (linkedInput (Ops.removeHyperlinks [ cellRange 0 0 20 5 ] relinked |> fst) (Some linkedImported.Source)))).Sheets.[0].Hyperlinks) []

// ---- pictures placed through the host API (SheetOps) ---------------------------
let logoInfo = (Ops.imageInfo logoBytes).Value
eq "pictures: the fixture's PNG is read off its header" (logoInfo.ContentType, logoInfo.Extension, logoInfo.Width, logoInfo.Height) ("image/png", "png", 24, 16)
eq "pictures: a file's objects are numbered back to front" (extrasModel.Drawings |> List.map (fun d -> d.Id, (match d.Content with Picture _ -> "picture" | KeptObject (kind, _) -> kind))) [ 1, "picture"; 2, "chart" ]
let imagesOf (wb: XlsxReader.Workbook) = wb.Sheets.[0].Extras.Images |> List.map (fun image -> image.Anchor.Kind, image.Anchor.FromRow, image.Anchor.FromCol, image.Anchor.FromColOffset, image.Anchor.Width, image.Anchor.Height, image.Name, image.Description, image.Data.Length)
// A workbook made in the editor: two copies of a picture, one moved and described.
let twoLogos =
    fresh
    |> Ops.insertPicture 1 logoInfo logoBytes (addr 1 1) 24.0 16.0 None
    |> Ops.insertPicture 2 logoInfo (Array.copy logoBytes) (addr 8 0) 48.0 32.0 (Some "Big logo")
    |> fun m -> Ops.placePicture 1 (addr 3 2) 4.0 6.0 24.0 16.0 m |> okOf
    |> fun m -> Ops.describePicture 2 None (Some "our logo, twice the size") m |> okOf
let twoLogosBytes = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (singleInput twoLogos None))
let twoLogosBack = XlsxReader.loadBytesNative twoLogosBytes |> Option.get
eq "pictures: written at their cells with offsets, sizes, names and descriptions" (imagesOf twoLogosBack)
    [ "oneCell", 3, 2, 4.0 * 9525.0, Some (24.0 * 9525.0), Some (16.0 * 9525.0), Some "Picture 1", None, logoBytes.Length
      "oneCell", 8, 0, 0.0, Some (48.0 * 9525.0), Some (32.0 * 9525.0), Some "Big logo", Some "our logo, twice the size", logoBytes.Length ]
let entriesOf (bytes: byte[]) =
    use archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes))
    archive.Entries |> Seq.map (fun entry -> entry.FullName) |> Seq.sort |> List.ofSeq
eq "pictures: the same picture twice is one media part" (entriesOf twoLogosBytes |> List.filter (fun name -> name.StartsWith "xl/media/")) [ "xl/media/image1.png" ]
let twoLogosAgain = (Bridge.Import.importWorkbook twoLogosBack).Sheets.[0].Model
eq "pictures: read back with ids in order" (twoLogosAgain.Drawings |> List.map (fun d -> d.Id, d.Name, d.Anchor.FromRow, d.Anchor.Width)) [ 1, Some "Picture 1", 3, Some 24.0; 2, Some "Big logo", 8, Some 48.0 ]
// On a sheet of a file with a chart: a picture added, the file's one sized, then both deleted.
let withNew =
    extrasModel
    |> Ops.insertPicture (Ops.nextDrawingId extrasModel) logoInfo logoBytes (addr 12 0) 24.0 16.0 (Some "Added")
    |> fun m -> Ops.placePicture 1 (addr 2 1) 0.0 0.0 60.0 40.0 m |> okOf
let withNewBack = reread (Bridge.Export.toWorkbookOut (extrasInput withNew))
eq "pictures: beside a file's chart: the old one sized, the new one in front, the chart as it was"
    (withNewBack.Sheets.[0].Extras.Images |> List.map (fun image -> image.Name, image.Anchor.Kind, image.Anchor.FromRow, image.Anchor.Width |> Option.map (fun emu -> emu / 9525.0), image.Anchor.Order),
     withNewBack.Sheets.[0].Extras.Objects |> List.map (fun item -> item.Kind, item.Name, item.Anchor.Order, item.Xml = extrasWb.Sheets.[0].Extras.Objects.[0].Xml))
    ([ Some "Logo", "oneCell", 2, Some 60.0, 0; Some "Added", "oneCell", 12, Some 24.0, 2 ], [ "chart", Some "Chart 1", 1, true ])
let chartOnly = withNew |> Ops.deletePicture 1 |> okOf |> Ops.deletePicture 3 |> okOf
let chartOnlyBytes = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (extrasInput chartOnly))
let chartOnlyBack = XlsxReader.loadBytesNative chartOnlyBytes |> Option.get
eq "pictures: deleted ones are gone from the file, with their media; the chart stays" (chartOnlyBack.Sheets.[0].Extras.Images.Length, chartOnlyBack.Sheets.[0].Extras.Objects |> List.map (fun item -> item.Kind), entriesOf chartOnlyBytes |> List.filter (fun name -> name.StartsWith "xl/media/")) (0, [ "chart" ], [])
// The last drawing of a sheet deleted: no drawing part is left.
let noLogosBytes = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (singleInput (twoLogosAgain |> Ops.deletePicture 1 |> okOf |> Ops.deletePicture 2 |> okOf) (Some (Bridge.Import.importWorkbook twoLogosBack).Source)))
eq "pictures: none left, no drawing part and no media" (entriesOf noLogosBytes |> List.filter (fun name -> name.StartsWith "xl/media/" || name.StartsWith "xl/drawings/")) []

// ---- tables made and changed in the editor -----------------------------------------
let tableCells =
    [ addr 0 0, CellValue.String "Region"; addr 0 1, CellValue.String "Amount"; addr 0 2, CellValue.String ""
      addr 1 0, CellValue.String "East"; addr 1 1, CellValue.Number 10.0; addr 1 2, CellValue.Number 1.0
      addr 2 0, CellValue.String "West"; addr 2 1, CellValue.Number 20.0; addr 2 2, CellValue.Number 2.0
      addr 6 0, CellValue.Formula ("=SUM(Sales[Amount])", CellValue.Number 30.0) ]
    |> List.fold (fun m (at, value) -> setCell at { Value = value; Format = defaultFormat; Comment = None } m) (createEmpty 30 8)
let tableBook0 = Bridge.buildWorkbookContext [| tableCells |] [| "First" |] [| true |] Map.empty Map.empty 0
let tableId = Bridge.sheetIdOfIndex 0
let tableBook =
    (Ops.defineTable "DefineTable" tableId { From = addr 0 0; To = addr 2 2 } (Some "Sales") true (Some (Ops.tableStyleNamed "TableStyleMedium9")) tableBook0 |> fst).Book
    |> fun wb -> (Ops.setTableTotalRow "SetTableTotalRow" "Sales" true wb).Book
    |> fun wb -> (Ops.defineTable "DefineTable" tableId { From = addr 10 0; To = addr 12 1 } None false None wb |> fst).Book
let tableInput (wb: Wb.Workbook) (source: Bridge.XlsxSource option) : Bridge.ExportInput =
    { Sheets = [| wb.Sheets.[tableId].Model |]; Names = [| "First" |]; Visible = [| true |]; ActiveSheetIndex = 0; NameInfo = wb.NameInfo; DefinedNames = wb.Names; Source = source; Macros = None; MacroProject = None }
let tableBytes = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (tableInput tableBook None))
let tableBack = XlsxReader.loadBytesNative tableBytes |> Option.get
eq "tables: written with their columns, total row, filter and style"
    (tableBack.Sheets.[0].Tables |> List.map (fun t -> t.Name, (t.StartRow, t.StartCol, t.EndRow, t.EndCol), t.Columns, t.HeaderRowCount, t.TotalsRowCount, t.Totals, t.AutoFilter, t.Style |> Option.map (fun s -> s.Name, s.ShowRowStripes)))
    [ "Sales", (1, 1, 4, 3), [ "Region"; "Amount"; "Column3" ], 1, 1, [ "label", "Total"; "", ""; "function", "sum" ], Some (1, 3), Some ("TableStyleMedium9", true)
      "Table1", (11, 1, 13, 2), [ "Column1"; "Column2" ], 0, 0, [ "", ""; "", "" ], None, None ]
let tablePartText (name: string) =
    use archive = new System.IO.Compression.ZipArchive(new MemoryStream(tableBytes))
    use reader = new StreamReader(archive.GetEntry(name).Open())
    reader.ReadToEnd()
ok "tables: the filter of a table stops above its total row" ((tablePartText "xl/tables/table1.xml").Contains("<autoFilter ref=\"A1:C3\"/>")) (tablePartText "xl/tables/table1.xml")
eq "tables: one part each" (entriesOf tableBytes |> List.filter (fun name -> name.StartsWith "xl/tables/")) [ "xl/tables/table1.xml"; "xl/tables/table2.xml" ]
let tableAgain = Bridge.Import.importWorkbook tableBack
eq "tables: read back as the editor had them"
    (Ops.tablesOf tableAgain.Context |> List.map (fun t -> t.Name, Ops.rangeText t.Range, t.HasHeaders, t.HasTotals, t.Columns, t.HasAutoFilter, t.Style |> Option.map (fun s -> s.Name)))
    (Ops.tablesOf tableBook |> List.map (fun t -> t.Name, Ops.rangeText t.Range, t.HasHeaders, t.HasTotals, t.Columns, t.HasAutoFilter, t.Style |> Option.map (fun s -> s.Name)))
eq "tables: the total and a structured reference come back as formulas"
    ([ addr 3 2; addr 6 0 ] |> List.map (fun at -> match tableAgain.Sheets.[0].Model.Cells.[at].Value with CellValue.Formula (text, _) -> Rbt.Spreadsheet.Model.formulaText text | other -> sprintf "%A" other))
    [ "=SUBTOTAL(109,Sales[Column3])"; "=SUM(Sales[Amount])" ]
// Opened again and changed: renamed, resized, one deleted - the file follows.
let tableEdited =
    (Ops.renameTable "RenameTable" "Sales" "Revenue" { tableAgain.Context with Sheets = tableAgain.Context.Sheets |> Map.map (fun _ sheet -> { sheet with Model = resolveSheet sheet.Model }) }).Book
    |> fun wb -> (Ops.deleteTable "DeleteTable" "Table1" true wb).Book
    |> fun wb -> (Ops.setTableTotalRow "SetTableTotalRow" "Revenue" false wb).Book
let tableEditedBytes = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (tableInput tableEdited (Some tableAgain.Source)))
let tableEditedBack = XlsxReader.loadBytesNative tableEditedBytes |> Option.get
eq "tables: renamed, without its total row; the deleted one's part is gone"
    (tableEditedBack.Sheets.[0].Tables |> List.map (fun t -> t.Name, (t.StartRow, t.StartCol, t.EndRow, t.EndCol), t.TotalsRowCount), entriesOf tableEditedBytes |> List.filter (fun name -> name.StartsWith "xl/tables/"))
    ([ "Revenue", (1, 1, 3, 3), 0 ], [ "xl/tables/table1.xml" ])
eq "tables: formulas name the table as renamed"
    (match (Bridge.Import.importWorkbook tableEditedBack).Sheets.[0].Model.Cells.[addr 6 0].Value with CellValue.Formula (text, _) -> Rbt.Spreadsheet.Model.formulaText text | other -> sprintf "%A" other)
    "=SUM(Revenue[Amount])"
// Untouched tables of a file go out as the file had them (their total-row functions too).
let tableUntouched = reread (Bridge.Export.toWorkbookOut (tableInput tableAgain.Context (Some tableAgain.Source)))
eq "tables: untouched, they are written as the file had them"
    (tableUntouched.Sheets.[0].Tables |> List.map (fun t -> t.Name, t.Columns, t.Totals, t.AutoFilter))
    (tableBack.Sheets.[0].Tables |> List.map (fun t -> t.Name, t.Columns, t.Totals, t.AutoFilter))

// ---- page setup ---------------------------------------------------------------
// Everything a host sets through SetPageSetup / SetPageBreaks is written and read back.
let printed =
    let setup =
        Ops.applyPageSetupChanges "SetPageSetup"
            [ Ops.SetOrientation LandscapeOrientation; Ops.SetPaper 9
              Ops.SetMargin ("top", 1.0); Ops.SetMargin ("bottom", 0.9); Ops.SetMargin ("left", 0.5); Ops.SetMargin ("right", 0.6); Ops.SetMargin ("header", 0.4); Ops.SetMargin ("footer", 0.2)
              Ops.SetFitTo (Some 1, Some 0)
              Ops.SetPrintArea (Some { From = addr 0 0; To = addr 39 5 }); Ops.SetRepeatRows (Some (0, 1)); Ops.SetRepeatCols (Some (0, 0))
              Ops.SetHeaderFooterText ("header", Some "&LLeft && more&C&\"Arial,Bold\"Title&R&D"); Ops.SetHeaderFooterText ("footer", Some "&CPage &P of &N")
              Ops.SetHeaderFooterText ("firstHeader", Some "&CFirst"); Ops.SetHeaderFooterText ("evenFooter", Some "&REven &P")
              Ops.SetDifferentFirst true; Ops.SetDifferentOddEven true; Ops.SetScaleWithDocument false
              Ops.SetPrintGridLines true; Ops.SetPrintHeadings true; Ops.SetCenterHorizontally true; Ops.SetCenterVertically true
              Ops.SetPageOrder OverThenDown; Ops.SetBlackAndWhite true; Ops.SetFirstPageNumber (Some 3); Ops.SetCommentsPrint CellCommentsAtEnd ]
            defaultPageSetup
        |> Ops.setPageBreaks "SetPageBreaks" [ 20; 30 ] [ 3 ]
    { createEmpty 100 12 with PageSetup = setup } |> setCell (addr 0 0) { Value = CellValue.String "x"; Format = defaultFormat; Comment = None }
let plainSheet = createEmpty 50 8 |> setCell (addr 0 0) { Value = CellValue.Number 1.0; Format = defaultFormat; Comment = None }
let printInput (sheets: SpreadsheetModel array) (source: Bridge.XlsxSource option) : Bridge.ExportInput =
    { Sheets = sheets; Names = [| "Print me"; "Plain" |]; Visible = [| true; true |]; ActiveSheetIndex = 0; NameInfo = Map.empty; DefinedNames = Map.empty; Source = source; Macros = None; MacroProject = None }
let printedBytes = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (printInput [| printed; plainSheet |] None))
let printedFile = XlsxReader.loadBytesNative printedBytes |> Option.get
let printedBack = Bridge.Import.importWorkbook printedFile
eq "page setup: every field is read back" printedBack.Sheets.[0].Model.PageSetup printed.PageSetup
eq "page setup: a sheet without one keeps the defaults" printedBack.Sheets.[1].Model.PageSetup defaultPageSetup
eq "page setup: the lazy import gives the same" ((Bridge.Import.importWorkbookLazy (fun _ -> false) printedFile).Sheets.[0].Model |> resolveSheet).PageSetup printed.PageSetup
let textOf (bytes: byte[]) (entry: string) =
    use archive = new System.IO.Compression.ZipArchive(new MemoryStream(bytes))
    use reader = new StreamReader(archive.GetEntry(entry).Open())
    reader.ReadToEnd()
let printNamesOf (bytes: byte[]) =
    System.Text.RegularExpressions.Regex.Matches(textOf bytes "xl/workbook.xml", "<definedName name=\"(_xlnm\\.Print[^\"]*)\"([^>]*)>([^<]*)</definedName>")
    |> Seq.cast<System.Text.RegularExpressions.Match>
    |> Seq.map (fun m -> m.Groups.[1].Value, m.Groups.[3].Value.Replace("&apos;", "'"), m.Groups.[2].Value.Contains "localSheetId=\"0\"")
    |> List.ofSeq |> List.sort
eq "page setup: the print area and titles are the sheet's built-in names"
    (printNamesOf printedBytes) [ "_xlnm.Print_Area", "'Print me'!$A$1:$F$40", true; "_xlnm.Print_Titles", "'Print me'!$1:$2,'Print me'!$A:$A", true ]
let printedXml = textOf printedBytes "xl/worksheets/sheet1.xml"
ok "page setup: the sheet says paper, orientation and fit" (printedXml.Contains "paperSize=\"9\"" && printedXml.Contains "orientation=\"landscape\"" && printedXml.Contains "<pageSetUpPr fitToPage=\"1\"/>" && printedXml.Contains "fitToHeight=\"0\"") printedXml
ok "page setup: and its breaks" (printedXml.Contains "<rowBreaks count=\"2\" manualBreakCount=\"2\"><brk id=\"20\"" && printedXml.Contains "<colBreaks count=\"1\"") printedXml
// Untouched, it goes out as the file had it.
let printedUntouched = Bridge.Export.toWorkbookOut (printInput (modelsOf printedBack) (Some printedBack.Source))
let printedOriginal = XlsxWriter.packageFromReaderWorkbook printedFile
eq "page setup: untouched, it is written as the file had it"
    (printedUntouched.Sheets |> List.map (fun s -> s.PageSetup, s.PageMargins, s.PrintOptions, s.HeaderFooter, s.RowBreaks, s.ColBreaks, s.PrintArea, s.PrintTitleRows, s.PrintTitleCols))
    (printedOriginal.Sheets |> List.map (fun s -> s.PageSetup, s.PageMargins, s.PrintOptions, s.HeaderFooter, s.RowBreaks, s.ColBreaks, s.PrintArea, s.PrintTitleRows, s.PrintTitleCols))
// Changed after opening: another paper (one the editor has no name for), a scale, other breaks, no titles.
let reprinted =
    let model = printedBack.Sheets.[0].Model
    let setup =
        Ops.applyPageSetupChanges "SetPageSetup" [ Ops.SetPaper 41; Ops.SetScale 80; Ops.SetRepeatRows None; Ops.SetRepeatCols None; Ops.SetHeaderFooterText ("header", None); Ops.SetOrientation PortraitOrientation ] model.PageSetup
        |> Ops.setPageBreaks "SetPageBreaks" [ 25 ] []
    { model with PageSetup = setup }
let reprintedBytes = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (printInput [| reprinted; printedBack.Sheets.[1].Model |] (Some printedBack.Source)))
let reprintedFile = XlsxReader.loadBytesNative reprintedBytes |> Option.get
let reprintedBack = (Bridge.Import.importWorkbook reprintedFile).Sheets.[0].Model.PageSetup
eq "page setup: changed after opening, the file follows" reprintedBack reprinted.PageSetup
eq "page setup: ... a paper without a name keeps its code" (reprintedBack.PaperSize, reprintedBack.FitToPage, reprintedBack.Scale, reprintedBack.ManualRowBreaks, reprintedBack.ManualColBreaks) (Some 41, false, 80, [ 25 ], [])
eq "page setup: ... the titles' name is gone, the area's stays"
    (printNamesOf reprintedBytes |> List.map (fun (name, _, _) -> name))
    [ "_xlnm.Print_Area" ]

// What the file says about the printer stays through an edit of the setup.
let withPrinter =
    let stream = new MemoryStream()
    stream.Write(printedBytes, 0, printedBytes.Length)
    (use archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Update, true)
     let entry = archive.GetEntry "xl/worksheets/sheet1.xml"
     let text = (use reader = new StreamReader(entry.Open()) in reader.ReadToEnd())
     entry.Delete()
     let fresh = archive.CreateEntry "xl/worksheets/sheet1.xml"
     use writer = new StreamWriter(fresh.Open())
     writer.Write(text.Replace("<pageSetup ", "<pageSetup horizontalDpi=\"600\" verticalDpi=\"300\" copies=\"2\" draft=\"1\" errors=\"dash\" ")))
    stream.ToArray()
let printerBack = Bridge.Import.importWorkbook (XlsxReader.loadBytesNative withPrinter |> Option.get)
let printerEdited =
    let model = printerBack.Sheets.[0].Model
    { model with PageSetup = Ops.applyPageSetupChanges "SetPageSetup" [ Ops.SetOrientation PortraitOrientation ] model.PageSetup }
let printerXml = textOf (XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (printInput [| printerEdited; printerBack.Sheets.[1].Model |] (Some printerBack.Source)))) "xl/worksheets/sheet1.xml"
ok "page setup: the printer's attributes stay when the setup is edited"
    (printerXml.Contains "orientation=\"portrait\"" && printerXml.Contains "horizontalDpi=\"600\"" && printerXml.Contains "verticalDpi=\"300\"" && printerXml.Contains "copies=\"2\"" && printerXml.Contains "draft=\"1\"" && printerXml.Contains "errors=\"dash\"")
    printerXml

// A formula that spilled is written as an array formula over its cells: opened
// again they are its results (the spill map), and a recalculation spills over
// them instead of finding them in its way.
let spilledModel =
    let cellOf value = { Value = value; Format = defaultFormat; Comment = None }
    createEmpty 50 8
    |> setCell (addr 0 0) (cellOf (CellValue.Number 3.0))
    |> setCell (addr 0 2) (cellOf (CellValue.Formula ("=SEQUENCE(A1,2)", CellValue.Empty)))
    |> setCell (addr 6 0) (cellOf (CellValue.Formula ("=A1*2", CellValue.Empty)))
    |> Rbt.Spreadsheet.Formula.recalcAndSpill None None
let spilledBytes = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (printInput [| spilledModel; plainSheet |] None))
let spilledXml = textOf spilledBytes "xl/worksheets/sheet1.xml"
ok "spill: the formula is written as an array formula over the cells it fills, a plain formula as before"
    (spilledXml.Contains "<f t=\"array\" ref=\"C1:D3\">_xlfn.SEQUENCE(A1,2)</f>" && spilledXml.Contains "<f>A1*2</f>") spilledXml
let spilledBack = (Bridge.Import.importWorkbook (XlsxReader.loadBytesNative spilledBytes |> Option.get)).Sheets.[0].Model
eq "spill: opened again, the cells belong to the formula" (spilledBack.Spill, spilledModel.Spill.Count) (spilledModel.Spill, 5)
let respilled = spilledBack |> setCell (addr 0 0) { Value = CellValue.Number 2.0; Format = defaultFormat; Comment = None } |> Rbt.Spreadsheet.Formula.recalcAndSpill None None
eq "spill: ... and a recalculation lays the array out again" ([ for row in 0 .. 2 -> getCellValue (addr row 2) respilled |> unwrapFormulaValue ], respilled.Spill.Count) ([ CellValue.Number 1.0; CellValue.Number 3.0; CellValue.Empty ], 3)
let spilledAgain = Bridge.Export.toWorkbookOut { printInput [| spilledBack; plainSheet |] None with Source = None }
ok "spill: a sheet opened from a file writes its array formulas back" ((textOf (XlsxPackageWriter.writeWorkbookPackageBytes spilledAgain) "xl/worksheets/sheet1.xml").Contains "<f t=\"array\" ref=\"C1:D3\">") ""

// ---- iterative calculation of circular references (<calcPr iterate>) --------
do
    let wb = load "real-world/saopaulo-condos.xlsx"
    let imported = Bridge.Import.importWorkbook wb
    eq "iteration: read from the file (Excel's defaults)" imported.Context.Iteration (Some { Wb.IterativeCalculation.MaxIterations = 100; MaxChange = 0.001 })
    let out = Bridge.Export.toWorkbookOut (inputOf imported (modelsOf imported) imported.Context.Names)
    let bytes = XlsxPackageWriter.writeWorkbookPackageBytes out
    ok "iteration: an opened workbook keeps it when written" (Text.RegularExpressions.Regex.IsMatch(textOf bytes "xl/workbook.xml", "<calcPr\\b[^>]*\\biterate=\"1\"")) (textOf bytes "xl/workbook.xml")
    eq "iteration: ... and when read again" (Bridge.Import.importWorkbook (XlsxReader.loadBytesNative bytes |> Option.get)).Context.Iteration imported.Context.Iteration
    // Switched off in the editor, and set with other numbers.
    let off = XlsxPackageWriter.writeWorkbookPackageBytes (out |> Bridge.Export.withIteration None)
    ok "iteration: switched off, no calcPr says iterate" (not ((textOf off "xl/workbook.xml").Contains "iterate")) ""
    let custom = out |> Bridge.Export.withIteration (Some { MaxIterations = 50; MaxChange = 0.01 }) |> reread
    eq "iteration: other numbers written and read" custom.Preserved.CalcIteration (Some (50, 0.01))
    let plain = Bridge.Import.importWorkbook (load "real-world/chart-zoo.xlsx")
    eq "iteration: a workbook without it" plain.Context.Iteration None
    // The circular reference of the file is calculated as Excel left it: nothing moves.
    let recalculated = Rbt.Spreadsheet.Formula.recalculateWorkbook imported.Context
    let moved =
        imported.Context.SheetOrder |> List.sumBy (fun id ->
            let before, after = imported.Context.Sheets.[id].Model.Cells, recalculated.Sheets.[id].Model.Cells
            before |> Map.toSeq |> Seq.filter (fun (cellAddr, cell) ->
                match unwrapFormulaValue cell.Value, unwrapFormulaValue after.[cellAddr].Value with
                | CellValue.Number x, CellValue.Number y -> abs (x - y) > 0.01 * (1.0 + abs x)
                | CellValue.Number _, CellValue.Error "CYCLE" -> true
                | _ -> false) |> Seq.length)
    eq "iteration: the file's circular model keeps its values through a full recalculation" moved 0

// ---- notes in runs --------------------------------------------------------
// A note of a file is read with its runs (font, size, bold, colour); when the
// sheet's comments change the comments part is written from the model, and
// every note keeps its runs.
do
    let wb = load "real-world/jumeirah-hotel.xlsx"
    let imported = Bridge.Import.importWorkbook wb
    let sheetIndex = wb.Sheets |> List.findIndex (fun sheet -> sheet.Comments |> List.exists (fun c -> c.CellRef.Column = "E" && c.CellRef.Row = 49))
    let model = imported.Sheets.[sheetIndex].Model
    let runsAt (m: SpreadsheetModel) row col =
        match Ops.commentAt m (addr row col) with
        | Some info -> info.Runs |> List.map (fun (text, run) -> text.Trim().Split('\n').[0], run.Bold, run.FontName, run.FontSize)
        | None -> []
    eq "note runs: the file's note has its runs" (runsAt model 48 4 |> List.map (fun (text, bold, font, size) -> text.StartsWith "BIWS:" || text.StartsWith "4% tax", bold, font, size))
        [ true, Some true, Some "Tahoma", Some 9.0; true, None, Some "Tahoma", Some 9.0 ]
    eq "note runs: the note's text is the runs joined" (Ops.commentAt model (addr 48 4) |> Option.map (fun info -> info.Text = (info.Runs |> List.map fst |> String.concat ""))) (Some true)
    // Another note on the sheet: the comments are not the file's any more.
    let red = { emptyRunOverride with Color = Some "#FF0000"; Italic = Some true; Underline = Some true; FontSize = Some 14.0; FontName = Some "Georgia" }
    let edited = model |> Ops.setCommentRuns (addr 0 0) [ "plain ", emptyRunOverride; "red", red; "", red ] (Some "Ann") |> Ops.setComment (addr 1 0) "just text" None
    eq "note runs: runs without text are dropped" (Ops.commentAt edited (addr 0 0) |> Option.map (fun info -> info.Text, info.Runs.Length)) (Some ("plain red", 2))
    eq "note runs: a note without formatting is a plain note" ((edited |> Ops.setCommentRuns (addr 2 0) [ "a", emptyRunOverride; "b", emptyRunOverride ] None |> fun m -> Ops.commentAt m (addr 2 0)) |> Option.map (fun info -> info.Text, info.Runs)) (Some ("ab", []))
    let sheets = modelsOf imported
    sheets.[sheetIndex] <- edited
    let back = reread (Bridge.Export.toWorkbookOut (inputOf imported sheets imported.Context.Names))
    let commentOf (column: string) (row: int) = back.Sheets.[sheetIndex].Comments |> List.find (fun c -> c.CellRef.Column = column && c.CellRef.Row = row)
    eq "note runs: the file's note keeps its runs when the sheet's comments changed"
        ((commentOf "E" 49).Runs |> List.map (fun (text, props) -> text.Trim().StartsWith "BIWS:" || text.Trim().StartsWith "4% tax", props.Bold, props.FontName, props.FontSize))
        [ true, true, Some "Tahoma", Some 9.0; true, false, Some "Tahoma", Some 9.0 ]
    eq "note runs: a note made in runs is written in runs"
        ((commentOf "A" 1).Runs |> List.map (fun (text, props) -> text, props.Italic, props.Underline, props.Color |> Option.map (fun c -> c.ToUpperInvariant().EndsWith "FF0000"), props.FontName, props.FontSize))
        [ "plain ", false, false, None, None, None; "red", true, true, Some true, Some "Georgia", Some 14.0 ]
    eq "note runs: a plain note is written as one text" ((commentOf "A" 2).Text, (commentOf "A" 2).Runs) ("just text", [])
    let again = (Bridge.Import.importWorkbook back).Sheets.[sheetIndex].Model
    eq "note runs: read back as they were" (Ops.commentAt again (addr 0 0) |> Option.map (fun info -> info.Text, info.Author, info.Runs)) (Ops.commentAt edited (addr 0 0) |> Option.map (fun info -> info.Text, info.Author, info.Runs))
    eq "note runs: the file's note reads back as it was" (runsAt again 48 4) (runsAt model 48 4)

// ---- the size of a row / column that has none of its own is the sheet's (sheetFormatPr) ----
do
    let imported = Bridge.Import.importWorkbook (load "real-world/milk-street-office.xlsx")
    let annual = imported.Sheets |> Array.find (fun sheet -> sheet.Name = "Annual")
    let model = annual.Model
    let near (a: float) (b: float) = abs (a - b) < 0.01
    // (Calibri 12, the Normal font: rows of 15.75 pt, 21 px; digits of 8 px, so defaultColWidth="11" is 88 px.)
    ok "defaults: the sheet's default row height" (near model.DefaultRowHeight 21.0) (string model.DefaultRowHeight)
    ok "defaults: the sheet's default column width" (near model.DefaultColumnWidth 88.0) (string model.DefaultColumnWidth)
    // The rows and columns the grid adds past the file's (398 rows, columns to L) are of that size.
    let grown = { model with RowCount = model.RowCount + 100; ColCount = model.ColCount + 26 }
    ok "defaults: a row past the file's rows has the default height" (near (getRowHeight (model.RowCount + 20) grown) 21.0) (string (getRowHeight (model.RowCount + 20) grown))
    ok "defaults: a column past the file's columns has the default width" (near (getColumnWidth (model.ColCount + 3) grown) 88.0) (string (getColumnWidth (model.ColCount + 3) grown))
    let metrics = Rbt.Spreadsheet.Layout.metricsOf grown
    ok "defaults: ... in the layout too"
        (near (metrics.RowPrefix.[model.RowCount + 21] - metrics.RowPrefix.[model.RowCount + 20]) 21.0 && near (metrics.ColPrefix.[model.ColCount + 4] - metrics.ColPrefix.[model.ColCount + 3]) 88.0)
        (sprintf "%f %f" (metrics.RowPrefix.[model.RowCount + 21] - metrics.RowPrefix.[model.RowCount + 20]) (metrics.ColPrefix.[model.ColCount + 4] - metrics.ColPrefix.[model.ColCount + 3]))
    // ... a row hidden and shown again there, and one given the default height, are of it.
    let row = model.RowCount + 20
    let shownAgain = grown |> Ops.hideAxis true row row |> Ops.unhideAxis true row row
    ok "defaults: hidden and shown again, the row is of the default height" (near (getRowHeight row shownAgain) 21.0) (string (getRowHeight row shownAgain))
    eq "defaults: a row set to the default height has no entry of its own" (Map.containsKey row (setRowHeight row grown.DefaultRowHeight (setRowHeight row 60.0 grown)).RowHeights) false
    // A sheet that states nothing has the file format's defaults (15 pt, 64 px, as Excel reports them); a new sheet the grid's.
    let plain = (Bridge.Import.importWorkbook (load "real-world/chart-zoo.xlsx")).Sheets.[0].Model
    let stated = (load "real-world/chart-zoo.xlsx").Sheets.[0].Format
    ok "defaults: as the sheet states, or 15 pt and 64 px (Excel's default)"
        (near plain.DefaultRowHeight ((stated.DefaultRowHeight |> Option.defaultValue 15.0) * 1.333)
         && near plain.DefaultColumnWidth ((match stated.DefaultColWidth with Some width -> width * 7.0 | None -> 64.0)))
        (sprintf "%f %f" plain.DefaultRowHeight plain.DefaultColumnWidth)
    eq "defaults: a new sheet has the grid's" ((createEmpty 10 5).DefaultRowHeight, (createEmpty 10 5).DefaultColumnWidth) (25.0, 100.0)
    // Written, the sheet says what it said.
    let out = Bridge.Export.toWorkbookOut (inputOf imported (modelsOf imported |> Array.map (fun m -> if obj.ReferenceEquals(m, model) then setCell (addr (model.RowCount + 5) 1) { Value = CellValue.Number 1.0; Format = defaultFormat; Comment = None } m else m)) Map.empty)
    let back = reread out
    let backSheet = back.Sheets |> List.find (fun sheet -> sheet.Name = "Annual")
    eq "defaults: an edited sheet is written with the defaults it had" (backSheet.Format.DefaultRowHeight, backSheet.Format.DefaultColWidth) (Some 15.75, Some 11.0)
    ok "defaults: ... and the row added past its rows has no height of its own" (backSheet.Rows |> List.forall (fun r -> r.Index <> model.RowCount + 6 || r.Height.IsNone || near r.Height.Value 15.75)) "a height"


// ---- sizes as Excel 2019 shows them (tools/excel-parity/make-sizing-fixtures.py, checked in Excel) ----
do
    let sheets (file: string) = (Bridge.Import.importWorkbook (load ("test-documents-excel/" + file))).Sheets |> Array.map (fun sheet -> sheet.Model)
    let widths (model: SpreadsheetModel) n = [ for col in 0 .. n - 1 -> getColumnWidth col model ]
    let heights (model: SpreadsheetModel) n = [ for row in 0 .. n - 1 -> getRowHeight row model ]
    let arial = sheets "synthetic-font-widths.xlsx"
    eq "sizes: Arial MT 12 is digits of 9 px" arial.[0].MaxDigitWidth 9.0
    eq "sizes: ... columns of 1, 2.1, 8.9, 10, 20, 38.8, 0.5, 2 and none" (widths arial.[0] 9) [ 9.0; 19.0; 80.0; 90.0; 180.0; 349.0; 4.0; 18.0; 80.0 ]
    eq "sizes: defaultColWidth 20, baseColWidth 12" (getColumnWidth 3 arial.[1], getColumnWidth 3 arial.[2]) (180.0, 120.0)
    eq "sizes: rows of the font's 15 pt whatever defaultRowHeight says, unless it is the sheet's own" (heights arial.[0] 2 @ heights arial.[3] 2 @ heights arial.[4] 2) [ 20.0; 20.0; 20.0; 20.0; 40.0; 40.0 ]
    eq "sizes: the Normal font is the first font, not the Normal style's" ((sheets "synthetic-normal-font-first.xlsx").[0].MaxDigitWidth, (sheets "synthetic-normal-font-style.xlsx").[0].MaxDigitWidth) (9.0, 7.0)
    let fit = sheets "synthetic-row-fit.xlsx"
    eq "sizes: Calibri 11 widths, cut down" (widths fit.[13] 6) [ 64.0; 3.0; 15.0; 91.0; 7.0; 62.0 ]
    let rows = heights fit.[0] 29
    eq "sizes: rows fitted on opening (20 pt, custom, ht 30, empty 20 pt, number, ht 12, 8 pt, customHeight without ht, a row style of 20 pt, a hidden column's cell, Arial 20, Times 14)"
        [ rows.[0]; rows.[3]; rows.[4]; rows.[6]; rows.[7]; rows.[10]; rows.[11]; rows.[14]; rows.[16]; rows.[17]; rows.[20]; rows.[24]; rows.[25] ]
        [ 35.0; 20.0; 20.0; 35.0; 35.0; 35.0; 20.0; 20.0; 35.0; 35.0; 35.0; 34.0; 25.0 ]
    eq "sizes: three lines wrapped are three; a line break without wrapping one" (rows.[12], rows.[9]) (60.0, 20.0)
    eq "sizes: a pixel more for thickBot, two for thickTop and thickBot" (rows.[27], rows.[28]) (21.0, 37.0)
    eq "sizes: a sheet of its own default height is fitted only with a row of a height of its own" (heights fit.[2] 2, heights fit.[4] 1, heights fit.[3] 1) ([ 20.0; 20.0 ], [ 35.0 ], [ 20.0 ])
    eq "sizes: a column style of 20 pt makes the rows 35 px" (getRowHeight 2 fit.[11], getRowHeight 50 fit.[11]) (35.0, 35.0)
    eq "sizes: every column in a style of Arial 10: rows of 17 px, a Calibri 11 cell's 20; some columns: 20" (heights fit.[9] 3 @ [ getRowHeight 40 fit.[9]; getRowHeight 1 fit.[10] ]) [ 20.0; 17.0; 17.0; 17.0; 20.0 ]
    let styled = fit.[12]
    let fillAt row col = (Ops.shownCell styled (addr row col) |> Option.map (fun cell -> cell.Format.BgColor)) |> Option.flatten
    eq "sizes: a row's style to the last column, over a column's; not without customFormat; a cell's own" [ fillAt 2 12; fillAt 2 7; fillAt 3 7; fillAt 4 300; fillAt 6 1; fillAt 8 1 ] [ Some "#FFFF00"; Some "#FFFF00"; Some "#D9D9D9"; Some "#FFFF00"; None; None ]
    // Written again unedited, the widths are the file's; a width set is in the font's digits.
    let wb = load "test-documents-excel/synthetic-font-widths.xlsx"
    let imported = Bridge.Import.importWorkbook wb
    let unedited = reread (Bridge.Export.toWorkbookOut (inputOf imported (modelsOf imported) Map.empty))
    eq "sizes: unedited, the <col> widths are the file's" (unedited.Sheets.[0].Columns |> List.map (fun c -> c.Width)) (wb.Sheets.[0].Columns |> List.map (fun c -> c.Width))
    let models = modelsOf imported |> Array.mapi (fun i m -> if i = 0 then setColumnWidth 1 45.0 m else m)
    let edited = reread (Bridge.Export.toWorkbookOut (inputOf imported models Map.empty))
    eq "sizes: a column of 45 px is 5 digits of 9 px, the others the file's"
        (edited.Sheets.[0].Columns |> List.map (fun c -> c.Index, c.Width))
        (wb.Sheets.[0].Columns |> List.map (fun c -> c.Index, (if c.Index = 2 then Some 5.0 else c.Width)))
    eq "sizes: ... and read again, 45 px" (widths (Bridge.Import.importWorkbook edited).Sheets.[0].Model 3) [ 9.0; 45.0; 80.0 ]


// ---- fractional font sizes (tools/excel-parity/make-row-height-fixtures.py, checked in Excel) ----
do
    let wb = load "test-documents-excel/synthetic-row-heights.xlsx"
    let imported = Bridge.Import.importWorkbook wb
    let fonts = imported.Sheets |> Array.find (fun sheet -> sheet.Name = "Fonts")
    let sizeAt row = (Map.tryFind (addr row 0) fonts.Model.Cells |> Option.bind (fun cell -> cell.Format.FontSize))
    eq "fonts: sizes keep their fractions (10.5, 7.5, 13.5, 9.5, 11.5 pt)" [ for row in 0 .. 4 -> sizeAt row ] [ Some 10.5; Some 7.5; Some 13.5; Some 9.5; Some 11.5 ]
    // Excel's rows (Calibri 11 sheet): a font is measured at its pixel size - Calibri 15.5 pt is 21 px, as 16 pt.
    eq "fonts: rows of Calibri 10.5 / 7.5 / 13.5 / 9.5 / 11.5, Arial 10.5 / 7.5 / 13.5 / 11.5, Times 10.5 / 13.5, Calibri 15.5, Arial 17.5"
        [ for row in 0 .. 12 -> getRowHeight row fonts.Model ] [ 20.0; 20.0; 24.0; 20.0; 20.0; 20.0; 20.0; 23.0; 20.0; 20.0; 23.0; 28.0; 29.0 ]
    let back = reread (Bridge.Export.toWorkbookOut (inputOf imported (modelsOf imported) Map.empty))
    let backFonts = back.Sheets |> List.find (fun sheet -> sheet.Name = "Fonts")
    let backSize row = backFonts.Rows |> List.tryFind (fun r -> r.Index = row + 1) |> Option.bind (fun r -> r.Cells |> List.tryHead) |> Option.bind (fun c -> c.Style) |> Option.bind (fun style -> style.FontSize)
    eq "fonts: written back as they were" [ for row in 0 .. 4 -> backSize row ] [ Some 10.5; Some 7.5; Some 13.5; Some 9.5; Some 11.5 ]
    let sheetNamed (imported: Bridge.ImportedWorkbook) name = (imported.Sheets |> Array.find (fun sheet -> sheet.Name = name)).Model
    let heights (model: SpreadsheetModel) n = [ for row in 0 .. n - 1 -> getRowHeight row model ]
    let custom = heights (sheetNamed imported "Custom") 59
    eq "rows: heights of their own in Excel's pixels (15.65 pt 20 px, 15.7 21, 20.15 26, 20.2 27, 12.7 17, 409 545)"
        [ custom.[13]; custom.[14]; custom.[34]; custom.[35]; custom.[46]; custom.[58] ] [ 20.0; 21.0; 26.0; 27.0; 17.0; 545.0 ]
    eq "rows: a sheet's own default of 15.7, 15.95, 20, 16, 14.4 pt"
        [ for name in [ "Def157"; "Def1595"; "Def20"; "Def16"; "Def144" ] -> getRowHeight 1 (sheetNamed imported name) ] [ 21.0; 21.0; 26.0; 21.0; 19.0 ]
    eq "rows: fitted rows are their fonts' lines, not the sheet's own default (30 pt); every column Calibri 8: 15 px"
        (heights (sheetNamed imported "FitMin30") 7, heights (sheetNamed imported "FitMinCols") 7)
        ([ 53.0; 20.0; 20.0; 35.0; 20.0; 20.0; 40.0 ], [ 53.0; 15.0; 20.0; 35.0; 15.0; 15.0; 21.0 ])
    let axes = Bridge.Import.importWorkbook (load "test-documents-excel/synthetic-hidden-axes.xlsx")
    eq "rows: <col hidden> to XFD hides to XFD; defaultColWidth 0 hides the columns of no <col>"
        ([ for col in [ 2; 3; 1045; 16383 ] -> getColumnWidth col (sheetNamed axes "ColsToEnd") ], [ for col in [ 2; 3; 100 ] -> getColumnWidth col (sheetNamed axes "DefColZero") ])
        ([ 64.0; 0.0; 0.0; 0.0 ], [ 70.0; 0.0; 0.0 ])
    let zero = sheetNamed axes "ZeroHeight"
    eq "rows: zeroHeight hides the rows not listed" (heights zero 12 @ [ getRowHeight 5000 zero ]) [ 20.0; 20.0; 0.0; 20.0; 0.0; 40.0; 20.0; 0.0; 20.0; 35.0; 0.0; 0.0; 0.0 ]
    eq "rows: ... shown again at the default" (getRowHeight 2 (Ops.unhideAxis true 2 2 zero)) 20.0
    let zeroBack = reread (Bridge.Export.toWorkbookOut (inputOf axes (modelsOf axes |> Array.map (fun m -> setColumnWidth 0 50.0 m)) Map.empty))
    ok "rows: ... zeroHeight written back" ((zeroBack.Sheets |> List.find (fun sheet -> sheet.Name = "ZeroHeight")).Format.ZeroHeight) "no zeroHeight"
    let merged = Bridge.Import.importWorkbook (load "test-documents-excel/synthetic-row-merged.xlsx")
    eq "rows: merged over rows count for nothing, across columns a 20 pt line counts, wrapped its descent"
        (let m = sheetNamed merged "Merged" in [ for row in [ 2; 3; 4; 5; 6; 10 ] -> getRowHeight row m ]) [ 35.0; 20.0; 20.0; 20.0; 20.0; 23.0 ]
    eq "rows: a column of Calibri 20 under wrapped Calibri 11: its descent below the text's (3 lines, plain, 2 lines)"
        (heights (sheetNamed merged "ColStyle") 4 |> List.tail) [ 63.0; 35.0; 43.0 ]
    // Wrapped rows on opening (make-wrap-fit-fixture.py + wrap-fit-cases.tsv; Excel 2019's Rows(r).Height on opening,
    // synthetic-wrap-fit.excel.tsv): Calibri 11 / bold at 85 %, Arial 10, Calibri 14 at 85 %, Calibri 20, in columns a
    // pixel apart about where a line fits (`SheetOps.openWrapLineCount`). Not the first screen of the sheets at 85 %
    // (S2, S4): the rows Excel shows first it fits at the sheet's zoom, which is not followed.
    let wrapFit = Bridge.Import.importWorkbook (load "test-documents-excel/synthetic-wrap-fit.xlsx")
    let expected =
        File.ReadAllLines(Path.Combine(root, "test-documents-excel/synthetic-wrap-fit.excel.tsv")) |> Array.map (fun line -> line.Split '\t')
        |> Array.filter (fun parts -> not ((parts.[0] = "S2" || parts.[0] = "S4") && int parts.[1] <= 60))
    let wrong =
        [ for parts in expected do
            let got = getRowHeight (int parts.[1] - 1) (sheetNamed wrapFit parts.[0]) * 0.75
            if abs (got - float parts.[2]) > 0.01 then yield sprintf "%s!%s %g (Excel %s)" parts.[0] parts.[1] got parts.[2] ]
    eq (sprintf "rows: wrapped text fitted on opening as Excel's (%d rows)" expected.Length) (List.truncate 8 wrong) []
    // Spaces at a line's end do not count, at its start they do (Excel: "Cost Unit   " one line from a 63-px column as
    // "Cost Unit", "   Cost Unit" from 72).
    let calibri (text: string) = fun i -> Rbt.Spreadsheet.ExcelText.textWidth (fun _ -> 0.0) "Calibri" false false 15 false (string text.[i])
    eq "rows: a line's trailing spaces do not count, leading ones do"
        [ for text, col in [ "Cost Unit   ", 63.0; "Cost Unit   ", 62.0; "   Cost Unit", 72.0; "   Cost Unit", 71.0 ] -> Ops.openWrapLineCount (calibri text) (col - 7.0) text ]
        [ 1; 2; 1; 2 ]
    let baseline = Bridge.Import.importWorkbook (load "test-documents-excel/synthetic-row-baseline.xlsx")
    eq "rows: cells at the bottom meet at their baseline (Arial 12 and Arial 11 bold: 21 px), centred or at the top only their height (20)"
        (heights (sheetNamed baseline "Baseline") 11) [ 21.0; 21.0; 20.0; 20.0; 20.0; 20.0; 20.0; 63.0; 60.0; 60.0; 63.0 ]
    eq "fonts: a size set is kept to a twentieth of a point, as Excel keeps it" (List.map fontSizeAsKept [ 10.3; 8.33; 10.01; 10.5 ]) [ 10.3; 8.35; 10.0; 10.5 ]

// ---- a number's room in italic (make-number-room-fixture.py, Range.Text from Excel 2019) ----
do
    let shown (face: string) (pt: float) (italic: bool) (zoom: float) (width: float) (value: float) =
        let format = { defaultFormat with FontName = Some face; FontSize = Some pt; Italic = italic; NumberFormat = Some "0.00%" }
        let model = setCell (addr 0 0) { Value = Number value; Format = format; Comment = None } (createEmpty 2 2)
        let measure : Rbt.Spreadsheet.Layout.TextMeasurer =
            fun text fontPt family bold italic ->
                let px = Rbt.Spreadsheet.ExcelText.pixelSize fontPt zoom
                Rbt.Spreadsheet.ExcelText.textWidth (fun _ -> 0.0) (defaultArg family "Calibri") bold italic px (Rbt.Spreadsheet.ExcelText.isSmallFonts fontPt zoom) text
        let fit : Rbt.Spreadsheet.CellText.NumberFit = { Zoom = zoom; Overhang = 0.0 }
        (Rbt.Spreadsheet.Paint.cellLookFitted (Some fit) model 0 0 (Map.tryFind (addr 0 0) model.Cells) (Some measure) (Math.Floor (width * zoom + 0.5))).Text
    // oyabayo's italic Aptos Narrow 14 at 57 %: -100.00% from a 91-px column, upright from 86.
    eq "numbers: italic has round(px x 0.16) less room (Aptos Narrow 14 italic at 57 %: 2 px)"
        [ shown "Aptos Narrow" 14.0 true 0.57 90.0 -1.0; shown "Aptos Narrow" 14.0 true 0.57 91.0 -1.0; shown "Aptos Narrow" 14.0 false 0.57 85.0 -1.0; shown "Aptos Narrow" 14.0 false 0.57 86.0 -1.0 ]
        [ "########"; "-100.00%"; "########"; "-100.00%" ]
    eq "numbers: ... Calibri 11 italic at 100 % (2 px), in Small Fonts at 57 % (half its size, 4 px)"
        [ shown "Calibri" 11.0 true 1.0 61.0 -1.0; shown "Calibri" 11.0 true 1.0 62.0 -1.0; shown "Calibri" 11.0 true 0.57 72.0 -1.0; shown "Calibri" 11.0 true 0.57 73.0 -1.0 ]
        [ "#######"; "-100.00%"; "#####"; "-100.00%" ]
    eq "numbers: no # where not one fits" (shown "Calibri" 11.0 true 0.57 20.0 -1.0) ""

// ---- form controls: read for drawing, written back as the file had them -------
// (synthetic-form-controls.xlsx: built in Excel 2019 by tools/excel-parity/build-form-controls.vbs.)
let controlsWb = load "test-documents-excel/synthetic-form-controls.xlsx"
let controlsImported = Bridge.Import.importWorkbook controlsWb
let flatControls = controlsImported.Sheets.[0].Model.FormControls
eq "controls: the kinds on Flat, in the file's order"
    (flatControls |> List.map (fun c -> c.Kind))
    [ "CheckBox"; "CheckBox"; "CheckBox"; "GBox"; "Radio"; "Radio"; "Radio"; "Drop"; "List"; "Spin"; "Scroll"; "Scroll"; "Label"; "Button"; "CheckBox" ]
eq "controls: captions (the VML text box)" (flatControls |> List.map (fun c -> c.Text) |> List.filter (fun t -> t <> "") |> List.truncate 5) [ "Checked"; "Unchecked"; "Mixed"; "Group box"; "Option one" ]
eq "controls: links, ranges and saved states (ctrlProps)"
    (flatControls |> List.map (fun c -> c.LinkedCell, c.InputRange, c.Checked, c.Selected, c.Value) |> List.take 10 |> List.skip 7)
    [ Some "$J$5", Some "$L$1:$L$8", 0, 3, 0; Some "$J$6", Some "$L$1:$L$8", 0, 5, 0; Some "$J$7", None, 0, 0, 40 ]
eq "controls: flat on Flat, 3-D shading on Shaded (check boxes)"
    ([ 0; 1 ] |> List.map (fun i -> controlsImported.Sheets.[i].Model.FormControls.[0].ThreeD)) [ false; true ]
eq "controls: the spinner's and scroll bars' limits, across"
    (flatControls |> List.filter (fun c -> c.Kind = "Spin" || c.Kind = "Scroll") |> List.map (fun c -> c.Min, c.Max, c.Inc, c.Page, c.Horizontal))
    [ 0, 100, 5, 10, false; 0, 100, 1, 10, true; 0, 10, 1, 10, false ]
eq "controls: the button's caption font (Calibri 11, black)" (flatControls.[13].Font) ("Calibri", 11.0, Some "#000000", false, false)
eq "controls: ActiveX controls by their pictures (EMF)"
    (controlsImported.Sheets.[2].Model.FormControls |> List.map (fun c -> c.Kind, c.Name, c.Picture |> Option.map fst, c.LinkedCell))
    [ "ActiveX", "CheckBox1", Some "image/x-emf", Some "$J$1"; "ActiveX", "CommandButton1", Some "image/x-emf", None ]
let flatAnchor = flatControls.[0].Anchor
eq "controls: the anchor (x14 controlPr) in pixels" (flatAnchor.FromCol, flatAnchor.FromOffsetX, flatAnchor.FromRow, flatAnchor.FromOffsetY, flatAnchor.ToCol, flatAnchor.ToOffsetX) (0, 13.0, 0, 13.0, Some 2, 18.0)
// Written back - unedited, and with the linked cells changed as a click would - the controls' parts are the file's.
let controlParts (bytes: byte[]) =
    use zip = new System.IO.Compression.ZipArchive(new MemoryStream(bytes))
    zip.Entries
    |> Seq.filter (fun e -> e.FullName.StartsWith "xl/ctrlProps/" || e.FullName.StartsWith "xl/activeX/" || e.FullName.StartsWith "xl/drawings/vmlDrawing" || e.FullName.StartsWith "xl/media/")
    |> Seq.map (fun e -> use s = e.Open() in use m = new MemoryStream() in s.CopyTo m; e.FullName, m.ToArray())
    |> Seq.sortBy fst
    |> List.ofSeq
let sheetControls (bytes: byte[]) (path: string) =
    use zip = new System.IO.Compression.ZipArchive(new MemoryStream(bytes))
    use reader = new StreamReader(zip.GetEntry(path).Open())
    let xml = reader.ReadToEnd()
    let at = xml.IndexOf "<controls>"
    if at < 0 then "" else xml.Substring(at, xml.IndexOf("</controls>", at) - at)
let controlsOriginal = File.ReadAllBytes(Path.Combine(root, "test-documents-excel/synthetic-form-controls.xlsx"))
let controlsUnedited = XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (inputOf controlsImported (modelsOf controlsImported) controlsImported.Context.Names))
let controlsEdited =
    let models = modelsOf controlsImported |> Array.copy
    models.[0] <- models.[0] |> setCell (addr 0 9) { Value = CellValue.Boolean false; Format = defaultFormat; Comment = None } |> setCell (addr 6 9) { Value = CellValue.Number 45.0; Format = defaultFormat; Comment = None }
    XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (inputOf controlsImported models controlsImported.Context.Names))
for label, bytes in [ "unedited", controlsUnedited; "linked cells edited", controlsEdited ] do
    ok ("controls: ctrlProps, VML, ActiveX parts as the file's, " + label) (controlParts bytes = controlParts controlsOriginal) "parts differ"
    eq ("controls: the sheet's <controls> as the file's, " + label) (sheetControls bytes "xl/worksheets/sheet1.xml") (sheetControls controlsOriginal "xl/worksheets/sheet1.xml")
    let again = XlsxReader.loadBytesNative bytes |> Option.get
    eq ("controls: read back alike, " + label) (again.Sheets |> List.map (fun s -> s.Extras.FormControls.Length)) [ 15; 15; 2 ]

// Opening an XLSX is not the same as entering its formulas: an uninitialized
// non-iterative cycle is empty text in Excel 2019, not numeric zero. An empty
// t="str" <v> is already a cached empty string; acyclic formulas without
// caches are calculated. These values must survive a save and another import.
let cachePackage () =
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    let rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    let sheet =
        "<worksheet xmlns=\"" + ns + "\"><sheetData><row r=\"1\">"
        + "<c r=\"A1\"><f>B1+1</f></c>"
        + "<c r=\"B1\"><f>A1/2</f><v>2</v></c>"
        + "<c r=\"C1\"><f>A1+10</f><v>99</v></c>"
        + "<c r=\"D1\"><f>D1+1</f><v/></c>"
        + "<c r=\"E1\" t=\"str\"><f>E1+1</f><v/></c>"
        + "<c r=\"F1\"><f>1+2</f><v/></c>"
        + "<c r=\"G1\" t=\"str\"><f>IF(FALSE,1, \"\")</f><v></v></c>"
        + "<c r=\"H1\" t=\"str\"><f>1+2</f></c>"
        + "<c r=\"I1\"><f>A1+10</f><v/></c>"
        + "</row></sheetData></worksheet>"
    let parts =
        [ "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>"
          "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"" + rel + "/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"
          "xl/workbook.xml", "<workbook xmlns=\"" + ns + "\" xmlns:r=\"" + rel + "\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>"
          "xl/_rels/workbook.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"" + rel + "/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>"
          "xl/worksheets/sheet1.xml", sheet ]
    use buffer = new MemoryStream()
    do
        use zip = new System.IO.Compression.ZipArchive(buffer, System.IO.Compression.ZipArchiveMode.Create, true)
        for path, xml in parts do
            use writer = new StreamWriter(zip.CreateEntry(path).Open())
            writer.Write xml
    buffer.ToArray()
let cacheImported = Bridge.Import.importWorkbook (XlsxReader.loadBytesNative (cachePackage ()) |> Option.get)
let cacheValues (imported: Bridge.ImportedWorkbook) =
    [ for col in 0 .. 8 ->
        imported.Sheets.[0].Model |> getCell (addr 0 col) |> Option.map (fun c -> unwrapFormulaValue c.Value) |> Option.defaultValue CellValue.Empty ]
let expectedCacheValues =
    [ CellValue.String ""; CellValue.Number 2.0; CellValue.Number 99.0; CellValue.String ""
      CellValue.String ""; CellValue.Number 3.0; CellValue.String ""; CellValue.Number 3.0; CellValue.String "" ]
eq "opened empty versus missing formula caches" (cacheValues cacheImported) expectedCacheValues
let cacheSaved =
    XlsxPackageWriter.writeWorkbookPackageBytes
        (Bridge.Export.toWorkbookOut (inputOf cacheImported (modelsOf cacheImported) cacheImported.Context.Names))
let cacheSheetXml =
    use zip = new System.IO.Compression.ZipArchive(new MemoryStream(cacheSaved))
    use reader = new StreamReader(zip.GetEntry("xl/worksheets/sheet1.xml").Open())
    reader.ReadToEnd()
let cacheCellXml ref =
    let found = Text.RegularExpressions.Regex.Match(cacheSheetXml, "<c\\b[^>]*\\br=\"" + ref + "\"[^>]*>[\\s\\S]*?</c>")
    if found.Success then found.Value else ""
ok "saved uninitialized cycles carry empty string results"
   ([ "A1"; "D1"; "E1"; "I1" ]
    |> List.forall (fun ref ->
        let xml = cacheCellXml ref
        xml.Contains("t=\"str\"") && Text.RegularExpressions.Regex.IsMatch(xml, "<v(?:\\s[^>]*)?></v>|<v\\s*/>")))
   (sprintf "%s / %s" (cacheCellXml "A1") (cacheCellXml "I1"))
ok "saved acyclic missing caches carry calculated numbers"
   ([ "F1", "3"; "H1", "3" ]
    |> List.forall (fun (ref, value) -> (cacheCellXml ref).Contains("<v>" + value + "</v>")))
   (sprintf "%s / %s" (cacheCellXml "F1") (cacheCellXml "H1"))
let cacheReopened = Bridge.Import.importWorkbook (XlsxReader.loadBytesNative cacheSaved |> Option.get)
eq "empty and computed formula caches after save/reopen" (cacheValues cacheReopened) expectedCacheValues

printfn ""
printfn "XLSX bridge tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
