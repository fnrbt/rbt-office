module Rbt.Spreadsheet.XlsxImporter

open Fable.Core
open Fable.Core.JsInterop
open Rbt.Spreadsheet.Model

/// <summary>
/// XLSX import for the browser. The conversion itself is
/// `XlsxBridge.Import` - the same typed code the native editors use; this
/// module adds the entry points JavaScript callers use (buffers, promises).
/// </summary>

/// A workbook as the Office.Formats reader returns it (`excel.load`).
type ExcelWorkbook = obj
type ExcelSheet = obj

type ImportedSheet = XlsxBridge.ImportedSheet
type ImportedWorkbook = XlsxBridge.ImportedWorkbook

[<Emit("$0 == null")>]
let private isNullOrUndefined (value: obj): bool = jsNative

/// Sheet id of the sheet at `index`. Must agree with the ids the Spreadsheet
/// component assigns (`sheetN`, 1-based).
let sheetIdOfIndex (index: int): string =
    XlsxBridge.sheetIdOfIndex index

/// Workbook context for a set of sheet models: every table on every sheet is
/// registered for structured references. `tableColumns` supplies declared
/// column names by table name; a table without them reads its header row.
let buildWorkbookContext (models: SpreadsheetModel array) (names: string array) (visible: bool array)
                         (tableColumns: Map<string, string list>)
                         (definedNames: Map<Rbt.Spreadsheet.Workbook.SheetId option * string, string>)
                         (activeIndex: int): Rbt.Spreadsheet.Workbook.Workbook =
    XlsxBridge.buildWorkbookContext models names visible tableColumns definedNames activeIndex

let private requireWorkbook (workbook: ExcelWorkbook): XlsxReader.Workbook =
    if isNullOrUndefined workbook then
        failwith "Not a readable XLSX workbook: the package has no workbook part"
    unbox<XlsxReader.Workbook> workbook

/// Convert a loaded reader workbook: every sheet, plus the workbook context
/// and what an export needs of the file (`Source`). Throws when the workbook
/// has no sheets.
let importWorkbook (workbook: ExcelWorkbook): ImportedWorkbook =
    XlsxBridge.Import.importWorkbook (requireWorkbook workbook)

    /// <summary>
    /// Import a sheet from a reader workbook into SpreadsheetModel.
    /// Formulas keep the values cached in the file. A formula the file has no
    /// value for is computed with this sheet as the only context - use
    /// `importWorkbook` when other sheets or defined names matter.
    /// Throws when the workbook is unreadable or the sheet does not exist.
    /// </summary>
    /// <param name="workbook">reader workbook object</param>
    /// <param name="sheetName">Name of sheet to import (or first sheet if not specified)</param>
    /// <returns>SpreadsheetModel populated with sheet data</returns>
let importSheet (workbook: ExcelWorkbook) (sheetName: string option): SpreadsheetModel =
    let workbook = requireWorkbook workbook
    let sheet =
        match sheetName with
        | Some name ->
            match workbook.Sheets |> List.tryFind (fun sheet -> sheet.Name = name) with
            | Some sheet -> sheet
            | None -> failwith (sprintf "XLSX workbook has no worksheet named '%s'" name)
        | None ->
            match workbook.Sheets with
            | first :: _ -> first
            | [] -> failwith "XLSX workbook has no worksheets"
    XlsxBridge.Import.importSheetWith XlsxBridge.defaultImportOptions workbook sheet

// ---------------------------------------------------------------------------
// Buffer entry points. A failure (not a zip, corrupt XML, no workbook, an
// archive over the size limits) rejects / throws: it must never come back as
// an empty sheet the user could save over the original.
//
// The `…Promise` functions are for JavaScript callers, which cannot await an
// F# Async; the Async ones remain for F# callers.
// ---------------------------------------------------------------------------

[<Emit("new Promise((resolve) => $0(() => resolve()))")>]
let private promiseOfCallback (start: (unit -> unit) -> unit): JS.Promise<unit> = jsNative

[<Emit("$0.then(() => $1)")>]
let private thenValue (promise: JS.Promise<unit>) (value: 'T): JS.Promise<'T> = jsNative

/// The formulas of a reader workbook: of its cells (found without reading
/// every row: `SheetRows.FormulaCells`) and of its defined names.
let formulasOfReaderWorkbook (workbook: XlsxReader.Workbook): string seq =
    seq {
        for sheet in workbook.Sheets do
            for cell in sheet.RowSource.FormulaCells do
                match cell.Formula, cell.Value with
                | Some formula, _ -> yield formula
                | None, XlsxReader.Formula formula -> yield formula
                | _ -> ()
        for name in workbook.NamedRanges do yield name.RefersTo }

/// The function modules a workbook's formulas need (Formula.fs, "The
/// function table"), loaded: resolves once they are installed - at once when
/// none is missing. Done before a workbook is converted, so that its first
/// calculation finds the functions.
let loadFunctionsOf (workbook: XlsxReader.Workbook): JS.Promise<unit> =
    promiseOfCallback (fun ready -> Rbt.Spreadsheet.Formula.requestFunctionsFor (formulasOfReaderWorkbook workbook) ready)

/// Read a workbook from an ArrayBuffer / Uint8Array / Blob, with the function
/// modules its formulas need loaded.
let loadReaderWorkbook (data: obj) (selected: string option option): JS.Promise<XlsxReader.Workbook> =
    let loading =
        match selected with
        | Some sheetName -> XlsxReader.loadSelected data sheetName
        | None -> XlsxReader.load data
    (loading.``then``(fun workbook ->
        match workbook with
        | Some workbook -> thenValue (loadFunctionsOf workbook) workbook
        | None -> failwith "Not a readable XLSX workbook: the package has no workbook part"))
    |> unbox<JS.Promise<XlsxReader.Workbook>>

/// Import one sheet. Only that sheet is parsed, so formulas keep their
/// cached values but nothing outside the sheet is available to recompute.
let importFromBufferPromise (buffer: JS.ArrayBuffer) (sheetName: string option): JS.Promise<SpreadsheetModel> =
    (loadReaderWorkbook buffer (Some sheetName)).``then``(fun workbook -> importSheet workbook sheetName)

/// Import the whole workbook: sheets with their visibility, and the context
/// (tables, defined names) cross-sheet formulas need.
let importWorkbookPromise (buffer: JS.ArrayBuffer): JS.Promise<ImportedWorkbook> =
    (loadReaderWorkbook buffer None).``then``(XlsxBridge.Import.importWorkbook)

/// Import every sheet in source order, paired with its name (hidden sheets
/// included - see `importWorkbookPromise` for visibility).
let importAllSheetsPromise (buffer: JS.ArrayBuffer): JS.Promise<(string * SpreadsheetModel) array> =
    (importWorkbookPromise buffer).``then``(fun imported ->
        imported.Sheets |> Array.map (fun sheet -> sheet.Name, sheet.Model))

/// The sheets the workbook lists, read from the workbook part alone (no
/// sheet is parsed, so this is quick however large the file is).
let private loadSheetInfos (buffer: JS.ArrayBuffer): JS.Promise<XlsxReader.Xlsx.SheetInfo list> =
    (XlsxReader.loadSheetInfos buffer).``then``(fun infos ->
        match infos with
        | Some infos -> infos
        | None -> failwith "Not a readable XLSX workbook: the package has no workbook part")

/// Name and visibility of every sheet: `{ name, state, visible }` where
/// state is "visible" | "hidden" | "veryHidden".
let getSheetInfosPromise (buffer: JS.ArrayBuffer): JS.Promise<obj[]> =
    (loadSheetInfos buffer).``then``(fun infos ->
        infos
        |> List.map (fun sheet ->
            let state = XlsxBridge.sheetStateText sheet.State
            createObj [ "name" ==> sheet.Name; "state" ==> state; "visible" ==> (state = "visible") ])
        |> List.toArray)

/// Names of every sheet in the workbook, hidden ones included.
let getSheetNamesPromise (buffer: JS.ArrayBuffer): JS.Promise<string[]> =
    (loadSheetInfos buffer).``then``(fun infos -> infos |> List.map (fun sheet -> sheet.Name) |> List.toArray)

/// Names of the sheets Excel shows a tab for.
let getVisibleSheetNamesPromise (buffer: JS.ArrayBuffer): JS.Promise<string[]> =
    (loadSheetInfos buffer).``then``(fun infos ->
        infos
        |> List.filter (fun sheet -> sheet.State = XlsxReader.SheetVisible)
        |> List.map (fun sheet -> sheet.Name)
        |> List.toArray)

    /// <summary>
    /// Import XLSX file from ArrayBuffer.
    /// </summary>
    /// <param name="buffer">ArrayBuffer containing XLSX file data</param>
    /// <param name="sheetName">Optional sheet name (or first sheet if not specified)</param>
    /// <returns>SpreadsheetModel populated with sheet data</returns>
let importFromBuffer (buffer: JS.ArrayBuffer) (sheetName: string option): Async<SpreadsheetModel> =
    importFromBufferPromise buffer sheetName |> Async.AwaitPromise

    /// <summary>
    /// Import every sheet in the workbook in source order, paired with
    /// its sheet name. Used by the multi-sheet load path so the chrome
    /// can render a tab strip and the user can swap between sheets
    /// without re-parsing the XLSX each time.
    /// </summary>
let importAllSheets (buffer: JS.ArrayBuffer): Async<(string * SpreadsheetModel) array> =
    importAllSheetsPromise buffer |> Async.AwaitPromise

    /// <summary>
    /// Get list of sheet names from workbook buffer.
    /// Useful for showing sheet selector before import.
    /// </summary>
    /// <param name="buffer">ArrayBuffer containing XLSX file data</param>
    /// <returns>Array of sheet names</returns>
let getSheetNames (buffer: JS.ArrayBuffer): Async<string[]> =
    getSheetNamesPromise buffer |> Async.AwaitPromise
