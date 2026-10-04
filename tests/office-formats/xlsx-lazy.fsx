#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// Sheets converted on demand (XlsxBridge.Import.importWorkbookLazy): a sheet
// of an opened file is a stand-in until something needs it. Converting it
// gives the model an eager import gives; a formula evaluated against the
// workbook reads a stand-in's cells (converting it); an export writes a sheet
// that was never converted as the file had it.

open System
open System.IO
open System.IO.Compression
open Rbt.Spreadsheet.Model
module Bridge = Rbt.Spreadsheet.XlsxBridge
module Wb = Rbt.Spreadsheet.Workbook
module Formula = Rbt.Spreadsheet.Formula

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-34s  %s" label detail)
let eq label (got: 'a) (expected: 'a) =
    ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let bytesOf (relative: string) = File.ReadAllBytes(Path.Combine(root, relative))
let loadBytes (bytes: byte[]) = XlsxReader.loadBytesNative bytes |> Option.get
let partsOf (bytes: byte[]) : Map<string, byte[]> =
    use archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read)
    [ for entry in archive.Entries ->
        use stream = entry.Open()
        use buffer = new MemoryStream()
        stream.CopyTo buffer
        entry.FullName, buffer.ToArray() ]
    |> Map.ofList
let inputOf (imported: Bridge.ImportedWorkbook) (sheets: SpreadsheetModel array) : Bridge.ExportInput =
    { Sheets = sheets
      Names = imported.Sheets |> Array.map (fun sheet -> sheet.Name)
      Visible = imported.Sheets |> Array.map (fun sheet -> sheet.Visible)
      ActiveSheetIndex = 0
      DefinedNames = imported.Context.Names
      Source = Some imported.Source
      Macros = None }
let export (imported: Bridge.ImportedWorkbook) (sheets: SpreadsheetModel array) =
    XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut (inputOf imported sheets))
let addr row col : CellAddress = { Row = row; Col = col }
/// What a model shows and holds, without what identifies the import.
let contentOf (model: SpreadsheetModel) =
    model.Cells, model.RowCount, model.ColCount, model.MergedCells, model.RowHeights, model.ColumnWidths, model.Tables,
    model.Validations, model.FrozenRows, model.FrozenCols, model.Hyperlinks, model.KeptRanges, model.Threads

for file in [ "real-world/milk-street-office.xlsx"; "real-world/lyric-multifamily.xlsx"; "real-world/dashboard-ecommerce-sales.xlsx"; "real-world/jumeirah-hotel.xlsx" ] do
    let label = Path.GetFileNameWithoutExtension file
    let bytes = bytesOf file
    let eager = Bridge.Import.importWorkbook (loadBytes bytes)
    let lazyBook = Bridge.Import.importWorkbookLazy (fun index -> index = 0) (loadBytes bytes)
    let count = eager.Sheets.Length
    eq (label + ": the same sheets") (lazyBook.Sheets |> Array.map (fun s -> s.Name, s.State, s.Visible)) (eager.Sheets |> Array.map (fun s -> s.Name, s.State, s.Visible))
    eq (label + ": the first sheet is converted, the others are stand-ins")
        (lazyBook.Sheets |> Array.map (fun s -> isDeferredSheet s.Model)) (Array.init count (fun index -> index > 0))
    ok (label + ": a stand-in has no cells") (lazyBook.Sheets |> Array.forall (fun s -> not (isDeferredSheet s.Model) || Map.isEmpty s.Model.Cells)) ""
    eq (label + ": tables are known before their sheet is converted") (lazyBook.Context.Tables |> Map.toList |> List.map fst) (eager.Context.Tables |> Map.toList |> List.map fst)
    eq (label + ": defined names") lazyBook.Context.Names eager.Context.Names

    // Unconverted sheets are written from the file: the package an eager import writes.
    let standIns = lazyBook.Sheets |> Array.map (fun s -> s.Model)
    let written = partsOf (export lazyBook standIns)
    let reference = partsOf (export eager (eager.Sheets |> Array.map (fun s -> s.Model)))
    eq (label + ": the same parts") (written |> Map.toList |> List.map fst) (reference |> Map.toList |> List.map fst)
    let different = written |> Map.toList |> List.filter (fun (path, data) -> Map.find path reference <> data) |> List.map fst
    eq (label + ": every part the same as with all sheets converted") different []
    ok (label + ": the export converted nothing") (standIns |> Array.mapi (fun index m -> index = 0 || isDeferredSheet m) |> Array.forall id) ""

    // Converting gives what the eager import gives.
    for index in 0 .. count - 1 do
        let converted = resolveSheet standIns.[index]
        ok (sprintf "%s: sheet %d converts to the eager model" label index) (contentOf converted = contentOf eager.Sheets.[index].Model) ""
        ok (sprintf "%s: sheet %d converts once" label index) (obj.ReferenceEquals(resolveSheet standIns.[index], converted)) ""
        ok (sprintf "%s: sheet %d is no stand-in" label index) (converted.Deferred.IsNone && not (isDeferredSheet standIns.[index])) ""
    // With everything converted the export is still the same.
    let converted = standIns |> Array.map resolveSheet
    let written = partsOf (export lazyBook converted)
    eq (label + ": converted and exported: the same parts again")
        (written |> Map.toList |> List.filter (fun (path, data) -> Map.find path reference <> data) |> List.map fst) []

// A formula evaluated against the workbook reads a stand-in's cells.
do
    let bytes = bytesOf "real-world/lyric-multifamily.xlsx"
    let eager = Bridge.Import.importWorkbook (loadBytes bytes)
    let lazyBook = Bridge.Import.importWorkbookLazy (fun index -> index = 0) (loadBytes bytes)
    let target = 2
    let name = lazyBook.Sheets.[target].Name
    let quoted = "'" + name.Replace("'", "''") + "'"
    // A cell of the third sheet that holds a number, in the eager model.
    let numberCell =
        eager.Sheets.[target].Model.Cells
        |> Map.toSeq
        |> Seq.pick (fun (cellAddr, data) -> match unwrapFormulaValue data.Value with CellValue.Number n when n <> 0.0 -> Some (cellAddr, n) | _ -> None)
    let cellAddr, expected = numberCell
    let reference = sprintf "=%s!%s%d*2" quoted (columnToLetter cellAddr.Col) (cellAddr.Row + 1)
    ok "the third sheet is a stand-in" (isDeferredSheet lazyBook.Sheets.[target].Model) ""
    let first = lazyBook.Sheets.[0].Model
    let value = Formula.evaluateFormula reference first (Some lazyBook.Context)
    eq "a formula reads the stand-in's cell" value (CellValue.Number (expected * 2.0))
    ok "which converted the sheet" (not (isDeferredSheet lazyBook.Sheets.[target].Model)) ""
    ok "and no other" (isDeferredSheet lazyBook.Sheets.[1].Model) ""
    // A recalculation of the workbook with a formula into a stand-in.
    let withFormula = setCell (addr 0 25) { Value = CellValue.Formula (reference, CellValue.Empty); Format = defaultFormat; Comment = None } first
    let recalculated = Formula.recalculateWorkbook (Wb.updateSheetModel (List.head lazyBook.Context.SheetOrder) withFormula lazyBook.Context)
    eq "a recalculation reads it too"
        (getCellValue (addr 0 25) (Map.find (List.head recalculated.SheetOrder) recalculated.Sheets).Model |> unwrapFormulaValue) (CellValue.Number (expected * 2.0))

// A workbook without cached formula values is converted whole (its values need every sheet).
do
    let sheet name cells : XlsxWriter.XlsxSheetOut =
        { (XlsxWriter.packageFromReaderWorkbook (loadBytes (bytesOf "real-world/dcf-stock-analysis.xlsx"))).Sheets.Head with
            Name = name; Cells = cells; MergedCells = []; Tables = []; ConditionalFormats = []; DataValidations = []; Hyperlinks = []
            Extras = XlsxReader.emptySheetExtras; ColumnWidths = []; RowHeights = []; PrintArea = None; PrintTitleRows = None; PrintTitleCols = None }
    let cell row col value : XlsxWriter.XlsxCellOut = { Row = row; Col = col; Value = value; Format = XlsxWriter.defaultXlsxCellFormat; Comment = None }
    let book : XlsxWriter.XlsxWorkbookOut =
        { Sheets = [ sheet "One" [ cell 0 0 (XlsxWriter.XlsxFormula ("Two!A1+1", None)) ]; sheet "Two" [ cell 0 0 (XlsxWriter.XlsxNumber 41.0) ] ]
          ActiveSheetIndex = 0; DefinedNames = []; VbaProjectSource = None; Dxfs = []; Preserved = XlsxReader.emptyWorkbookPreservation }
    let imported = Bridge.Import.importWorkbookLazy (fun index -> index = 0) (loadBytes (XlsxWriter.writeWorkbookPackageBytes book))
    ok "uncached formulas: every sheet converted" (imported.Sheets |> Array.forall (fun s -> not (isDeferredSheet s.Model))) ""
    eq "and calculated" (getCellValue (addr 0 0) imported.Sheets.[0].Model |> unwrapFormulaValue) (CellValue.Number 42.0)

// The dependency graph of a workbook whose sheets are stand-ins: the formulas
// of a sheet that is not converted are read off the file, and an edit converts
// the sheets the formulas it makes stale are on or read - by their place in
// the workbook, a hidden sheet between them included - and no other.
do
    let number value = { Value = CellValue.Number value; Format = defaultFormat; Comment = None }
    let formula text cached = { Value = CellValue.Formula (text, CellValue.Number cached); Format = defaultFormat; Comment = None }
    let sheetOf (cells: (int * int * CellData) list) =
        cells |> List.fold (fun model (row, col, data) -> setCell (addr row col) data model) (createEmpty 50 10)
    let models =
        [| sheetOf [ 0, 0, number 1.0; 0, 1, formula "Second!A1+Fourth!A1+A1" 6.0; 0, 2, number 10.0; 5, 5, number 0.0 ]   // Main
           sheetOf [ 0, 0, number 2.0 ]                                                                                 // Second
           sheetOf [ 0, 0, formula "Main!C1*2" 20.0; 1, 0, formula "A1+1" 21.0 ]                                        // Hidden (not visible)
           sheetOf [ 0, 0, number 3.0 ]                                                                                 // Fourth
           sheetOf [ 0, 0, number 4.0; 0, 1, formula "A1*5" 20.0 ] |]                                                   // Other
    let names = [| "Main"; "Second"; "Hidden"; "Fourth"; "Other" |]
    let input : Bridge.ExportInput =
        { Sheets = models; Names = names; Visible = [| true; true; false; true; true |]; ActiveSheetIndex = 0
          DefinedNames = Map.empty; Source = None; Macros = None }
    let bytes = XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input)
    let book = Bridge.Import.importWorkbookLazy (fun index -> index = 0) (loadBytes bytes)
    let standIns = book.Sheets |> Array.map (fun sheet -> sheet.Model)
    let converted () = standIns |> Array.map (fun model -> not (isDeferredSheet model))
    eq "opened: the first sheet is converted, the hidden one is not visible" (converted (), book.Sheets |> Array.map (fun sheet -> sheet.Visible)) ([| true; false; false; false; false |], [| true; true; false; true; true |])
    let ids = List.toArray book.Context.SheetOrder
    let engine = Rbt.Spreadsheet.Recalc.build book.Context
    eq "the graph holds the formulas of every sheet" (Rbt.Spreadsheet.Recalc.formulaCount engine) 4
    eq "building it converts no sheet" (converted ()) [| true; false; false; false; false |]
    let edit (cell: CellAddress) (data: CellData) (wb: Wb.Workbook) =
        let edited = Wb.updateSheetModel ids.[0] (setCell cell data (resolveSheet wb.Sheets.[ids.[0]].Model)) wb
        match Rbt.Spreadsheet.Recalc.recalculate engine edited (Some [ ids.[0], cell ]) with
        | Some recalculated -> recalculated
        | None -> failwith "declined"
    let valueOf (wb: Wb.Workbook) index row col = getCellValue (addr row col) (resolveSheet wb.Sheets.[ids.[index]].Model) |> unwrapFormulaValue
    let typed = edit (addr 5 5) { Value = CellValue.String "x"; Format = defaultFormat; Comment = None } book.Context
    eq "a text nothing reads: no sheet converted" (converted ()) [| true; false; false; false; false |]
    let main = edit (addr 0 0) (number 100.0) typed
    eq "a cell the first sheet's formula reads: the sheets that formula reads (the second and the fourth), not the hidden one between them"
        (converted ()) [| true; true; false; true; false |]
    eq "and its value" (valueOf main 0 0 1) (CellValue.Number 105.0)
    let hidden = edit (addr 0 2) (number 50.0) main
    eq "a cell the hidden sheet's formula reads: the hidden sheet" (converted ()) [| true; true; true; true; false |]
    eq "its formulas follow" (valueOf hidden 2 0 0, valueOf hidden 2 1 0) (CellValue.Number 100.0, CellValue.Number 101.0)
    eq "the last sheet was never needed" (isDeferredSheet standIns.[4]) true
    // The same edits with every sheet converted up front: the same values.
    let eager = Bridge.Import.importWorkbook (loadBytes bytes)
    let eagerEdited =
        [ addr 5 5, { Value = CellValue.String "x"; Format = defaultFormat; Comment = None }; addr 0 0, number 100.0; addr 0 2, number 50.0 ]
        |> List.fold (fun wb (cell, data) -> Wb.updateSheetModel ids.[0] (setCell cell data wb.Sheets.[ids.[0]].Model) wb) eager.Context
        |> Formula.recalculateWorkbook
    for index in 0 .. 4 do
        eq (sprintf "sheet %d: the cells of a full recalculation of the converted workbook" index) (resolveSheet hidden.Sheets.[ids.[index]].Model).Cells eagerEdited.Sheets.[ids.[index]].Model.Cells

printfn "xlsx-lazy: %d passed, %d failed" pass fail
if fail > 0 then exit 1
