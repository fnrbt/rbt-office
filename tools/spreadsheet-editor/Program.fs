module SpreadsheetEditorApp

open System
open System.IO
open Renderer
open Platform
open Platform.X11Window
open Composit.Ui.Reactive
open Composit.Ui.Widget
open Composit.Ui.Layout

module XlUi = RedBat.Spreadsheet.Ui
module Model = RedBat.Spreadsheet.Model
module Editor = Composit.Ui.SpreadsheetEditor
module Chrome = Composit.Ui.SpreadsheetUiNative
module Xlsx = XlsxReader

open RedBat.Spreadsheet.Model

let private existingFile path =
    if File.Exists path then Some path else None

// ── Recent files ──────────────────────────────────────────────────

let private recentFilesPath =
    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".config/redbat-spreadsheet/recent.txt")

let private loadRecentFiles () : string list =
    if File.Exists recentFilesPath then
        try File.ReadAllLines(recentFilesPath) |> Array.toList |> List.filter (fun s -> s.Trim() <> "")
        with _ -> []
    else []

let private saveRecentFiles (list: string list) =
    try
        Directory.CreateDirectory(Path.GetDirectoryName recentFilesPath) |> ignore
        File.WriteAllLines(recentFilesPath, list |> List.toArray)
    with _ -> ()

let private recentFiles : Signal<string list> = signal (loadRecentFiles ())

let private rememberRecent (path: string) =
    let abs = Path.GetFullPath path
    let next =
        abs :: (recentFiles.Value |> List.filter (fun p -> p <> abs))
        |> List.truncate 12
    recentFiles.Value <- next
    saveRecentFiles next

let private fontDir () =
    [
        Environment.GetEnvironmentVariable("RB_WORD_FONT_DIR")
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "artifacts/fonts/windows"))
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../artifacts/fonts/windows"))
        Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "../../artifacts/fonts/windows"))
    ]
    |> List.choose (fun dir ->
        if String.IsNullOrWhiteSpace dir then None
        else Path.Combine(dir, "calibri.ttf") |> existingFile)
    |> List.tryHead

// ── Minimal native .xlsx writer ───────────────────────────────────
//
// Walks the workbook's sheets and emits the four parts a usable .xlsx
// needs: workbook.xml + per-sheet xl/worksheets/sheetN.xml + the
// content-types manifest + the relationships. Cell *values* round-trip
// (numbers, strings via shared strings, booleans, formula sources);
// styles, fonts, fills, merges, frozen panes, charts, and tables are
// dropped. Mirrors DocxWrite in tools/word-editor/Program.fs.

module XlsxWrite =
    let private escapeXml (s: string) =
        let sb = System.Text.StringBuilder()
        for c in s do
            match c with
            | '<' -> sb.Append "&lt;" |> ignore
            | '>' -> sb.Append "&gt;" |> ignore
            | '&' -> sb.Append "&amp;" |> ignore
            | '"' -> sb.Append "&quot;" |> ignore
            | _ -> sb.Append c |> ignore
        sb.ToString()

    let private cellRefOf (row: int) (col: int) =
        sprintf "%s%d" (columnToLetter col) (row + 1)

    /// Collect a deduped shared-string table from every string cell in
    /// the workbook. We index by zero-based position; the workbook XML
    /// references entries with `t="s"` + `<v>idx</v>`.
    let private buildSharedStrings (sheets: SpreadsheetModel array) : string array * Map<string, int> =
        let table = System.Collections.Generic.List<string>()
        let index = System.Collections.Generic.Dictionary<string, int>()
        for sheet in sheets do
            for KeyValue(_, data) in sheet.Cells do
                match data.Value with
                | CellValue.String s when not (index.ContainsKey s) ->
                    index.[s] <- table.Count
                    table.Add s
                | _ -> ()
        let asMap =
            let mutable m = Map.empty
            for kv in index do m <- Map.add kv.Key kv.Value m
            m
        table.ToArray(), asMap

    let private sharedStringsXml (table: string array) =
        let entries =
            table
            |> Array.map (fun s -> sprintf "<si><t xml:space=\"preserve\">%s</t></si>" (escapeXml s))
            |> String.concat ""
        sprintf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"%d\" uniqueCount=\"%d\">%s</sst>"
            table.Length table.Length entries

    let private cellXml (row: int) (col: int) (data: CellData) (sharedIndex: Map<string, int>) =
        let ref = cellRefOf row col
        match data.Value with
        | CellValue.Empty -> ""
        | CellValue.Number n -> sprintf "<c r=\"%s\"><v>%g</v></c>" ref n
        | CellValue.String s ->
            let idx = Map.find s sharedIndex
            sprintf "<c r=\"%s\" t=\"s\"><v>%d</v></c>" ref idx
        | CellValue.Boolean b ->
            sprintf "<c r=\"%s\" t=\"b\"><v>%d</v></c>" ref (if b then 1 else 0)
        | CellValue.Formula (expr, _) ->
            sprintf "<c r=\"%s\"><f>%s</f></c>" ref (escapeXml expr)
        | CellValue.Error e ->
            sprintf "<c r=\"%s\" t=\"e\"><v>%s</v></c>" ref (escapeXml e)

    let private sheetXml (model: SpreadsheetModel) (sharedIndex: Map<string, int>) =
        let cellsByRow =
            model.Cells
            |> Map.toList
            |> List.groupBy (fun (addr, _) -> addr.Row)
            |> List.sortBy fst
        let rows =
            cellsByRow
            |> List.map (fun (rowIdx, cells) ->
                let sorted = cells |> List.sortBy (fun (addr, _) -> addr.Col)
                let cellsXml =
                    sorted
                    |> List.map (fun (addr, data) -> cellXml addr.Row addr.Col data sharedIndex)
                    |> String.concat ""
                sprintf "<row r=\"%d\">%s</row>" (rowIdx + 1) cellsXml)
            |> String.concat ""
        sprintf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\
<sheetData>%s</sheetData></worksheet>" rows

    let private workbookXml (names: string array) =
        let sheets =
            names
            |> Array.mapi (fun i name ->
                sprintf "<sheet name=\"%s\" sheetId=\"%d\" r:id=\"rId%d\"/>"
                    (escapeXml name) (i + 1) (i + 1))
            |> String.concat ""
        sprintf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
<sheets>%s</sheets></workbook>" sheets

    let private workbookRelsXml (sheetCount: int) =
        let sheetRels =
            [ 1 .. sheetCount ]
            |> List.map (fun i ->
                sprintf "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet%d.xml\"/>" i i)
            |> String.concat ""
        let sharedRelId = sheetCount + 1
        let sharedRel =
            sprintf "<Relationship Id=\"rId%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>" sharedRelId
        sprintf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">%s%s</Relationships>"
            sheetRels sharedRel

    let private contentTypesXml (sheetCount: int) =
        let overrides =
            [ 1 .. sheetCount ]
            |> List.map (fun i ->
                sprintf "<Override PartName=\"/xl/worksheets/sheet%d.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" i)
            |> String.concat ""
        sprintf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
<Default Extension=\"xml\" ContentType=\"application/xml\"/>\
<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\
<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>\
%s</Types>" overrides

    let private rootRelsXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\
</Relationships>"

    /// Pack the workbook as a .xlsx and write it to `path`.
    let save (sheets: SpreadsheetModel array) (names: string array) (path: string) =
        let table, sharedIndex = buildSharedStrings sheets
        let perSheet =
            sheets
            |> Array.mapi (fun i model ->
                { ZipWriter.Path = sprintf "xl/worksheets/sheet%d.xml" (i + 1)
                  ZipWriter.Content = sheetXml model sharedIndex })
            |> Array.toList
        let files = [
            { ZipWriter.Path = "[Content_Types].xml"; ZipWriter.Content = contentTypesXml sheets.Length }
            { ZipWriter.Path = "_rels/.rels"; ZipWriter.Content = rootRelsXml }
            { ZipWriter.Path = "xl/workbook.xml"; ZipWriter.Content = workbookXml names }
            { ZipWriter.Path = "xl/_rels/workbook.xml.rels"; ZipWriter.Content = workbookRelsXml sheets.Length }
            { ZipWriter.Path = "xl/sharedStrings.xml"; ZipWriter.Content = sharedStringsXml table }
            yield! perSheet
        ]
        let bytes = ZipWriter.createPackageBytes files
        File.WriteAllBytes(path, bytes)

let private loadXlsx (path: string) : (SpreadsheetModel array * string array) option =
    try
        let bytes = File.ReadAllBytes path
        match Xlsx.loadBytesNative bytes with
        | Some workbook ->
            let sheets =
                workbook.Sheets
                |> List.map (fun sheet ->
                    // ExcelNativeRunner.toModel parks the selection off-grid
                    // at (-1,-1) so snapshot PNGs don't show the active-cell
                    // border, and it sizes RowCount/ColCount to the used
                    // range only. The interactive editor wants a real
                    // selection at A1 and enough empty rows/cols past the
                    // data that the grid fills the viewport the way Excel
                    // does — otherwise a 5x8 sheet renders as a tiny island
                    // of cells in a white void.
                    let model = ExcelNativeRunner.toModel sheet workbook
                    { model with
                        Selection = SingleCell (0, 0)
                        RowCount = max model.RowCount 200
                        ColCount = max model.ColCount 40 })
                |> List.toArray
            // Pull the workbook's VBA module source (xl/vbaProject.bin) onto each
            // sheet model so macros can run and Workbook_Open/Worksheet_Change
            // events fire. Empty when the file has no VBA project.
            let macros = VbaProject.extractFromXlsx bytes |> Option.defaultValue ""
            let sheets = if macros = "" then sheets else sheets |> Array.map (fun m -> { m with Macros = macros })
            let names =
                workbook.Sheets
                |> List.map (fun sheet -> sheet.Name)
                |> List.toArray
            if sheets.Length = 0 then None
            else Some (sheets, names)
        | None -> None
    with _ -> None

let private emptyWorkbook () =
    [| createEmpty 100 26 |], [| "Sheet1" |]

// ── Editor wiring ──────────────────────────────────────────────────

let private editsSinceSave : Signal<int> = signal 0

/// Style and metadata snapshot at the active cell — mirrors styleAtCursor
/// in tools/word-editor/Program.fs. Reads from the cell under the focus
/// of the active sheet's Selection.
let private cellFormatAt (state: Editor.SpreadsheetEditorState) =
    let model = Editor.activeModel state
    let r, c = Editor.focusCell model
    match getCell { Row = r; Col = c } model with
    | Some data -> data.Format
    | None -> defaultFormat

let private shellState (state: Editor.SpreadsheetEditorState) : XlUi.ShellState =
    let fmt () = cellFormatAt state
    let halignName () =
        (fmt ()).HorizontalAlign |> Option.defaultValue "left"
    let valignName () =
        (fmt ()).VerticalAlign |> Option.defaultValue "bottom"
    let activeName () =
        let names = state.SheetNames.Value
        let idx = state.ActiveSheet.Value
        if idx >= 0 && idx < names.Length then names.[idx] else ""
    {
        ActiveCellRef = fun () -> Editor.activeCellRef state
        FormulaText = fun () -> Editor.formulaText state
        IsBoldActive = fun () -> (fmt ()).Bold
        IsItalicActive = fun () -> (fmt ()).Italic
        IsUnderlineActive = fun () -> (fmt ()).Underline
        IsStrikeActive = fun () -> (fmt ()).Strikethrough
        IsWrapTextActive = fun () -> (fmt ()).WrapText
        IsMergedActive = fun () ->
            let model = Editor.activeModel state
            let r, c = Editor.focusCell model
            (findMergedCell { Row = r; Col = c } model).IsSome
        IsFilterActive = fun () -> (Editor.activeModel state).Filters |> Map.isEmpty |> not
        IsFreezePanesActive = fun () ->
            let model = Editor.activeModel state
            model.FrozenRows > 0 || model.FrozenCols > 0
        HorizontalAlign = halignName
        VerticalAlign = valignName
        NumberFormat = fun () -> (fmt ()).NumberFormat |> Option.defaultValue "general"
        TextColor = fun () -> (fmt ()).TextColor |> Option.defaultValue "#000000"
        BgColor = fun () -> (fmt ()).BgColor |> Option.defaultValue "#FFFFFF"
        FontFamily = fun () -> (fmt ()).FontName |> Option.defaultValue "Calibri"
        FontSize = fun () -> sprintf "%d" ((fmt ()).FontSize |> Option.defaultValue 11)
        BorderStyle = fun () -> "none"
        ZoomLevel = fun () -> "100"
        ActiveSheet = activeName
        Sheets = fun () ->
            state.SheetNames.Value
            |> Array.map (fun n -> XlUi.choice n n)
            |> Array.toList
        SelectionSummary = fun () -> Editor.selectionSummary state
        CanUndo = fun () -> not (List.isEmpty state.UndoStack.Value)
        CanRedo = fun () -> not (List.isEmpty state.RedoStack.Value)
    }

let private updateFormat (state: Editor.SpreadsheetEditorState) (update: CellFormat -> CellFormat) =
    let model = Editor.activeModel state
    let r, c = Editor.focusCell model
    let addr = { Row = r; Col = c }
    Editor.applyChange state (fun m ->
        let existing =
            match getCell addr m with
            | Some d -> d
            | None -> { Value = Empty; Format = defaultFormat; Comment = None }
        setCell addr { existing with Format = update existing.Format } m) |> ignore

let private shellActions
        (state: Editor.SpreadsheetEditorState)
        (status: Signal<string>)
        (currentPath: Signal<string option>)
        (renderState: Chrome.ShellRenderState)
        (loadFromPath: string -> unit) : XlUi.ShellActions * (unit -> unit) * (unit -> unit) =
    let say message () = status.Value <- message
    let sayArg message _ = status.Value <- message
    let parseHAlign = function
        | "center" -> Some "center"
        | "right" -> Some "right"
        | _ -> Some "left"
    let parseVAlign = function
        | "top" -> Some "top"
        | "middle" -> Some "middle"
        | _ -> Some "bottom"

    let openFindReplaceDialog () =
        XlUi.findReplaceDialog
            (fun q matchCase ->
                if q = "" then
                    status.Value <- "Enter a search term"
                    ""
                else
                    let model = Editor.activeModel state
                    let hits = searchCells q matchCase model
                    if hits.Length = 0 then
                        status.Value <- sprintf "No matches for \"%s\"" q
                        ""
                    else
                        let target = hits.[0]
                        Editor.applyChangeNoHistory state (fun m -> { m with Selection = SingleCell (target.Row, target.Col) }) |> ignore
                        status.Value <- sprintf "Match 1 of %d for \"%s\"" hits.Length q
                        "")
            (fun q r matchCase ->
                let next = replaceAll q r matchCase (Editor.activeModel state)
                Editor.applyChange state (fun _ -> next) |> ignore
                status.Value <- sprintf "Replaced occurrences of \"%s\"" q
                "")
        |> Composit.Ui.OfficeShell.openDialog renderState

    let openInsertTableDialog () =
        XlUi.insertTableDialog (fun rows cols ->
            // Register a real table at the active cell: header row ("Column N"
            // for blanks) + `rows` data rows + a SUBTOTAL totals row, styled
            // and autofiltered. Model.createTable also adds it to model.Tables
            // so it can be bridged for structured references.
            let model = Editor.activeModel state
            let r, c = Editor.focusCell model
            let name = sprintf "Table%d" (List.length model.Tables + 1)
            Editor.applyChange state (fun m ->
                createTable name r c (r + rows) (c + cols - 1) true m) |> ignore
            status.Value <- sprintf "Inserted table %s (%d×%d) with totals row" name rows cols)
        |> Composit.Ui.OfficeShell.openDialog renderState

    let openRunMacroDialog () =
        // Edit + run the workbook's VBA module. The macro reads/writes cells
        // (Vba.runOnModel, cross-platform); the edited source is kept on the
        // model so events (Worksheet_Change / Workbook_Open) can use it too.
        let model = Editor.activeModel state
        XlUi.macroDialog model.Macros (fun source ->
            let mutable msg = "Macro ran"
            Editor.applyChange state (fun m ->
                match RedBat.Spreadsheet.Vba.runOnModel { m with Macros = source } source with
                | Result.Ok m2 -> m2
                | Result.Error e -> msg <- sprintf "Macro error: %s" e; { m with Macros = source }) |> ignore
            status.Value <- msg)
        |> Composit.Ui.OfficeShell.openDialog renderState

    let openInsertPivotDialog () =
        let parseRange (s: string) =
            match s.Split(':') with
            | [| a; b |] ->
                match parseCellRef (a.Trim()), parseCellRef (b.Trim()) with
                | Some (r1, c1), Some (r2, c2) -> Some(min r1 r2, min c1 c2, max r1 r2, max c1 c2)
                | _ -> None
            | [| a |] -> parseCellRef (a.Trim()) |> Option.map (fun (r, c) -> r, c, r, c)
            | _ -> None
        let rec cellText (v: CellValue) =
            match v with
            | CellValue.Number n -> if n = floor n && abs n < 1e15 then string (int64 n) else sprintf "%g" n
            | CellValue.String s -> s
            | CellValue.Boolean b -> if b then "TRUE" else "FALSE"
            | CellValue.Error e -> "#" + e
            | CellValue.Formula (_, c) -> cellText c
            | _ -> ""
        // Live preview: resolve + compute against the CURRENT model, show the
        // formula we'll write plus the first rows of its result.
        let pivotPreview (req: XlUi.PivotUiRequest) : string =
            match parseRange req.Source with
            | None -> "Enter a source range like A1:C100."
            | Some (sr, sc, er, ec) ->
                let model = Editor.activeModel state
                let grid = [| for r in sr .. er -> [| for c in sc .. ec -> getCellValue { Row = r; Col = c } model |] |]
                let headers = if grid.Length > 0 then grid.[0] else [||]
                let rowFields = req.RowFields.Split(',') |> Array.map (fun s -> s.Trim()) |> Array.filter (fun s -> s <> "") |> Array.toList
                let colOpt = if req.ColField = "" then None else Some req.ColField
                let agg = RedBat.Spreadsheet.Pivot.parseAgg req.Agg |> Option.defaultValue RedBat.Spreadsheet.Pivot.Sum
                let formula = match RedBat.Spreadsheet.Pivot.buildFormula sr sc er ec headers rowFields colOpt req.ValueField req.Agg with Result.Ok f -> f | Result.Error _ -> ""
                match RedBat.Spreadsheet.Pivot.pivotFromNames grid rowFields colOpt req.ValueField agg with
                | Result.Error m -> m
                | Result.Ok p ->
                    let rows = p |> Array.truncate 5 |> Array.map (fun row -> row |> Array.truncate 4 |> Array.map cellText |> String.concat "    ") |> String.concat "\n"
                    (if formula <> "" then formula + "\n" else "") + rows
        XlUi.pivotTableDialog pivotPreview (fun req ->
            let model = Editor.activeModel state
            match parseRange req.Source, parseCellRef req.Dest with
            | Some (sr, sc, er, ec), Some (dr, dc) ->
                // Write a LIVE GROUPBY/PIVOTBY formula (not a static grid) — the
                // pivot is then a normal formula: recomputes on source edits,
                // composable, and the user can read/tweak it.
                let headers = [| for c in sc .. ec -> getCellValue { Row = sr; Col = c } model |]
                let rowFields = req.RowFields.Split(',') |> Array.map (fun s -> s.Trim()) |> Array.filter (fun s -> s <> "") |> Array.toList
                let colOpt = if req.ColField = "" then None else Some req.ColField
                match RedBat.Spreadsheet.Pivot.buildFormula sr sc er ec headers rowFields colOpt req.ValueField req.Agg with
                | Result.Ok formula ->
                    Editor.applyChange state (fun m ->
                        let m1 = setCell { Row = dr; Col = dc } { Value = CellValue.Formula(formula, CellValue.Empty); Format = defaultFormat; Comment = None } m
                        // Evaluate + materialize the spill so it shows immediately.
                        RedBat.Spreadsheet.Formula.recalculateAllFormulas None None m1
                        |> RedBat.Spreadsheet.Formula.spillArrays None None) |> ignore
                    status.Value <- sprintf "Inserted %s at %s" (if colOpt.IsSome then "PIVOTBY" else "GROUPBY") (formatCellRef dr dc)
                    true
                | Result.Error msg -> status.Value <- "Pivot: " + msg; false
            | _ -> status.Value <- "Pivot: invalid source range or destination cell."; false)
        |> Composit.Ui.OfficeShell.openDialog renderState

    let openFilterDialog () =
        let model = Editor.activeModel state
        let r, c = Editor.focusCell model
        // Table-aware: when the focused cell is inside a table, scope the
        // checklist to that table's data rows for the column (this is what a
        // table header filter button does); otherwise filter the whole column.
        let uniques, label =
            match tableAt r c model with
            | Some t -> tableColumnUniqueValues t c model, sprintf "%s (table %s)" (columnToLetter c) t.Name
            | None -> columnUniqueValues c model, columnToLetter c
        match uniques with
        | [] -> status.Value <- "No values to filter in this column."
        | uniques ->
            XlUi.filterDialog label uniques (fun visible ->
                Editor.applyChange state (fun m ->
                    if Set.count visible = List.length uniques then removeFilter c m else applyFilter c visible m)
                |> ignore
                status.Value <- sprintf "Filter on %s: %d of %d shown" label (Set.count visible) (List.length uniques))
            |> Composit.Ui.OfficeShell.openDialog renderState

    let openInsertSheetDialog () =
        XlUi.insertSheetDialog (fun name ->
            let sheets = state.Sheets.Value
            let names = state.SheetNames.Value
            let newSheets = Array.append sheets [| createEmpty 100 26 |]
            let newNames = Array.append names [| name |]
            state.CalcEngine.Value <- None
            state.Sheets.Value <- newSheets
            state.SheetNames.Value <- newNames
            Editor.switchSheet state (newSheets.Length - 1)
            editsSinceSave.Value <- editsSinceSave.Value + 1
            status.Value <- sprintf "Added sheet \"%s\"" name)
        |> Composit.Ui.OfficeShell.openDialog renderState

    let openRenameSheetDialog () =
        let names = state.SheetNames.Value
        let idx = state.ActiveSheet.Value
        let initial = if idx >= 0 && idx < names.Length then names.[idx] else ""
        XlUi.renameSheetDialog initial (fun name ->
            let next = Array.copy state.SheetNames.Value
            if idx >= 0 && idx < next.Length then
                state.CalcEngine.Value <- None
                next.[idx] <- name
                state.SheetNames.Value <- next
                editsSinceSave.Value <- editsSinceSave.Value + 1
                status.Value <- sprintf "Renamed sheet to \"%s\"" name)
        |> Composit.Ui.OfficeShell.openDialog renderState

    let openFileAction () =
        XlUi.openFileDialog loadFromPath
        |> Composit.Ui.OfficeShell.openDialog renderState

    let writeTo path =
        try
            XlsxWrite.save state.Sheets.Value state.SheetNames.Value path
            currentPath.Value <- Some path
            editsSinceSave.Value <- 0
            rememberRecent path
            status.Value <- sprintf "Saved %s" (Path.GetFileName path)
        with ex ->
            status.Value <- sprintf "Save failed: %s" ex.Message

    let saveFileAction () =
        let initial = currentPath.Value |> Option.defaultValue ""
        XlUi.saveFileDialog initial writeTo
        |> Composit.Ui.OfficeShell.openDialog renderState

    {
        PrintPdf = say "PDF export not wired in native shell"
        OpenFindReplace = openFindReplaceDialog
        Cut = fun () -> let _ = Editor.cutSelection state in status.Value <- "Cut"
        Copy = fun () -> let _ = Editor.copySelection state in status.Value <- "Copied"
        Paste = fun () ->
            if Editor.pasteFromClipboard state then status.Value <- "Pasted"
            else status.Value <- "Clipboard empty"
        Undo = fun () -> let _ = Editor.undo state in ()
        Redo = fun () -> let _ = Editor.redo state in ()
        ToggleBold = fun () -> updateFormat state (fun f -> { f with Bold = not f.Bold })
        ToggleItalic = fun () -> updateFormat state (fun f -> { f with Italic = not f.Italic })
        ToggleUnderline = fun () -> updateFormat state (fun f -> { f with Underline = not f.Underline })
        ToggleStrike = fun () -> updateFormat state (fun f -> { f with Strikethrough = not f.Strikethrough })
        ToggleWrapText = fun () -> updateFormat state (fun f -> { f with WrapText = not f.WrapText })
        ClearFormatting = fun () -> updateFormat state (fun _ -> defaultFormat)
        ApplyHorizontalAlign = fun v -> updateFormat state (fun f -> { f with HorizontalAlign = parseHAlign v })
        ApplyVerticalAlign = fun v -> updateFormat state (fun f -> { f with VerticalAlign = parseVAlign v })
        ApplyNumberFormat = fun v -> updateFormat state (fun f -> { f with NumberFormat = Some v })
        IncreaseDecimals = say "Increase decimals not wired"
        DecreaseDecimals = say "Decrease decimals not wired"
        ApplyFontFamily = fun v -> updateFormat state (fun f -> { f with FontName = Some v })
        ApplyFontSize = fun v ->
            match Int32.TryParse v with
            | true, n -> updateFormat state (fun f -> { f with FontSize = Some n })
            | _ -> ()
        ApplyTextColor = fun v -> updateFormat state (fun f -> { f with TextColor = Some v })
        ApplyBgColor = fun v -> updateFormat state (fun f -> { f with BgColor = Some v })
        ApplyBorderStyle = sayArg "Border style not wired"
        ApplyZoom = sayArg "Zoom not wired"
        MergeCells = fun () ->
            let model = Editor.activeModel state
            match model.Selection with
            | Range (sr, sc, er, ec) ->
                Editor.applyChange state (addMergedCell sr sc er ec) |> ignore
                status.Value <- "Merged selection"
            | _ -> status.Value <- "Select a range to merge"
        UnmergeCells = fun () ->
            let model = Editor.activeModel state
            let r, c = Editor.focusCell model
            Editor.applyChange state (removeMergedCell { Row = r; Col = c }) |> ignore
        InsertRow = fun () ->
            let r, _ = Editor.focusCell (Editor.activeModel state)
            Editor.applyChange state (insertRow r) |> ignore
        InsertColumn = fun () ->
            let _, c = Editor.focusCell (Editor.activeModel state)
            Editor.applyChange state (insertColumn c) |> ignore
        DeleteRow = fun () ->
            let r, _ = Editor.focusCell (Editor.activeModel state)
            Editor.applyChange state (deleteRow r) |> ignore
        DeleteColumn = fun () ->
            let _, c = Editor.focusCell (Editor.activeModel state)
            Editor.applyChange state (deleteColumn c) |> ignore
        SortAscending = fun () ->
            let _, c = Editor.focusCell (Editor.activeModel state)
            Editor.applyChange state (sortByColumn c true) |> ignore
        SortDescending = fun () ->
            let _, c = Editor.focusCell (Editor.activeModel state)
            Editor.applyChange state (sortByColumn c false) |> ignore
        ToggleFilter = openFilterDialog
        ClearFilter = fun () ->
            Editor.applyChange state (fun m -> { m with Filters = Map.empty; FilterHiddenRows = [||] }) |> ignore
            status.Value <- "Filters cleared"
        ToggleFreezePanes = fun () ->
            Editor.applyChange state (fun m ->
                if m.FrozenRows > 0 || m.FrozenCols > 0 then
                    { m with FrozenRows = 0; FrozenCols = 0 }
                else
                    let r, c = Editor.focusCell m
                    { m with FrozenRows = r; FrozenCols = c }) |> ignore
        SplitPanes = say "Split panes not wired"
        InsertChart = say "Insert chart not wired"
        InsertPivot = openInsertPivotDialog
        InsertTable = openInsertTableDialog
        RunMacro = openRunMacroDialog
        InsertSheet = openInsertSheetDialog
        DeleteSheet = fun () ->
            let sheets = state.Sheets.Value
            if sheets.Length <= 1 then status.Value <- "Can't delete the last sheet"
            else
                let idx = state.ActiveSheet.Value
                let nextSheets = Array.append sheets.[..idx-1] sheets.[idx+1..]
                let nextNames = Array.append state.SheetNames.Value.[..idx-1] state.SheetNames.Value.[idx+1..]
                state.CalcEngine.Value <- None
                state.Sheets.Value <- nextSheets
                state.SheetNames.Value <- nextNames
                state.ActiveSheet.Value <- max 0 (min (nextSheets.Length - 1) idx)
                editsSinceSave.Value <- editsSinceSave.Value + 1
                status.Value <- "Deleted sheet"
        RenameSheet = openRenameSheetDialog
        SwitchSheet = fun name ->
            let idx =
                state.SheetNames.Value
                |> Array.tryFindIndex (fun n -> n = name)
            match idx with
            | Some i -> Editor.switchSheet state i
            | None -> ()
        EditFormula = fun text -> Editor.editFormula state text
        SubmitFormula = fun () -> Editor.submitFormula state
    }, openFileAction, saveFileAction

// ── App state and event loop ───────────────────────────────────────

type AppState = {
    mutable ShiftDown: bool
    mutable CtrlDown: bool
    mutable AltDown: bool
    mutable LayoutRoot: LayoutNode option
    mutable Width: int
    mutable Height: int
    mutable LastHover: (bool -> unit) option
    mutable LastMouseX: float
    mutable LastMouseY: float
    mutable LastCmds: Cmd list
    Window: IWindow
    EditorState: Editor.SpreadsheetEditorState
    RenderState: Chrome.ShellRenderState
    Actions: XlUi.ShellActions
    /// File ops sit on AppState rather than ShellActions because the
    /// spreadsheet's ShellActions record (shared with the Fable shell)
    /// doesn't yet include open/save — Ctrl+O / Ctrl+S route through
    /// these closures.
    OpenFile: unit -> unit
    SaveFile: unit -> unit
    ControlServer: ControlServer.ControlServer
}

let private handleInput (state: AppState) (ev: InputEvent) =
    match ev with
    | KeyDown(keycode, _) ->
        match keycode with
        | 50 | 62 -> state.ShiftDown <- true
        | 37 | 105 -> state.CtrlDown <- true
        | 64 | 108 -> state.AltDown <- true
        | _ -> ()

        if not (KeyMap.isModifier keycode) then
            match KeyMap.translate keycode state.ShiftDown with
            | Some key ->
                if key = "Escape" && state.RenderState.OpenMenu.Value.IsSome then
                    state.RenderState.OpenMenu.Value <- None
                    state.Window.RequestRedraw()
                elif state.CtrlDown && (key = "s" || key = "S") then
                    state.SaveFile ()
                    state.Window.RequestRedraw()
                elif state.CtrlDown && (key = "o" || key = "O") then
                    state.OpenFile ()
                    state.Window.RequestRedraw()
                elif state.CtrlDown && (key = "f" || key = "F") then
                    state.Actions.OpenFindReplace ()
                    state.Window.RequestRedraw()
                elif state.CtrlDown && (key = "t" || key = "T") then
                    state.Actions.InsertTable ()
                    state.Window.RequestRedraw()
                else
                    let ke : KeyEvent = { Key = key; Shift = state.ShiftDown; Ctrl = state.CtrlDown; Alt = state.AltDown }
                    match state.LayoutRoot with
                    | Some root ->
                        let _ = dispatchKey root ke
                        state.Window.RequestRedraw()
                    | None -> ()
            | None -> ()

    | KeyUp(keycode, _) ->
        match keycode with
        | 50 | 62 -> state.ShiftDown <- false
        | 37 | 105 -> state.CtrlDown <- false
        | 64 | 108 -> state.AltDown <- false
        | _ -> ()

    | MouseButton(button, pressed, x, y) ->
        state.LastMouseX <- x
        state.LastMouseY <- y
        match state.LayoutRoot with
        | Some root when button = 0 ->
            if pressed then
                let (onClick, _) = hitTestAll root x y
                onClick |> Option.iter (fun f -> f())
                match pointerHitTestAll root x y with
                | Some(h, lx, ly) -> h (PointerDown(lx, ly))
                | None -> ()
            else
                match pointerHitTestAll root x y with
                | Some(h, lx, ly) -> h (PointerUp(lx, ly))
                | None -> ()
            state.Window.RequestRedraw()
        | Some root when button = 2 && pressed ->
            match contextMenuHitTestAll root x y with
            | Some handler -> handler (x, y)
            | None -> ()
            state.Window.RequestRedraw()
        | _ -> ()

    | MouseMove(x, y) ->
        state.LastMouseX <- x
        state.LastMouseY <- y
        match state.LayoutRoot with
        | Some root ->
            let (_, onHover) = hitTestAll root x y
            state.LastHover |> Option.iter (fun f -> f false)
            onHover |> Option.iter (fun f -> f true)
            state.LastHover <- onHover
            match pointerHitTestAll root x y with
            | Some(h, lx, ly) -> h (PointerMove(lx, ly))
            | None -> ()
            state.Window.RequestRedraw()
        | None -> ()

    | MouseScroll(_, dy) ->
        match state.LayoutRoot with
        | Some root ->
            match scrollHitTestAll root state.LastMouseX state.LastMouseY with
            | Some(_, sy) ->
                sy.Value <- sy.Peek() + dy * 60.0
                state.Window.RequestRedraw()
            | None -> ()
        | None -> ()

    | Resize(w, h) ->
        state.Width <- w
        state.Height <- h

    | Close ->
        state.Window.Close()

    | MouseLeave ->
        state.LastHover |> Option.iter (fun f -> f false)
        state.LastHover <- None

let private renderFrame (state: AppState) (buildWidget: unit -> Widget) (w: int) (h: int) : Cmd list =
    state.ControlServer.ProcessCommands()
    try
        state.Width <- w
        state.Height <- h
        Animation.tick()
        TextEdit.tickCaret 0.016

        let widgetTree = buildWidget()
        let bounds = { X = 0.0; Y = 0.0; Width = float w; Height = float h }
        clearOverlays()
        initOverlayBounds bounds
        let root = layout widgetTree bounds
        state.LayoutRoot <- Some root

        let mainCmds = paint root
        let overlayCmds = getOverlays() |> List.collect paint
        let bg = Composit.Ui.OfficeShell.palette XlUi.chrome
        let cmds = [ Clear bg.PaperDesk; yield! mainCmds; yield! overlayCmds ]
        state.LastCmds <- cmds
        cmds
    with ex ->
        eprintfn "[render] %O" ex
        [ Clear (Color.rgb 255uy 0uy 0uy) ]

// ── Main ───────────────────────────────────────────────────────────

[<EntryPoint>]
let main argv =
    let sw = Diagnostics.Stopwatch.StartNew()
    Platform.now <- fun () -> sw.Elapsed.TotalSeconds

    fontDir () |> Option.iter Composit.Ui.Layout.initFont

    let sheets, names =
        match argv |> Array.tryHead with
        | Some path ->
            match loadXlsx path with
            | Some (s, n) ->
                rememberRecent path
                s, n
            | None ->
                eprintfn "Failed to parse %s; opening blank workbook" path
                emptyWorkbook ()
        | None -> emptyWorkbook ()

    let editor = Editor.createState sheets names 0
    let status = signal "Ready"
    let currentPath = signal (argv |> Array.tryHead)
    let renderState = Chrome.closed

    let loadFromPath path =
        match loadXlsx path with
        | Some (s, n) ->
            // Fire the workbook's Workbook_Open handler on the active sheet
            // (a no-op unless its macro module defines the sub).
            let s =
                if s.Length > 0 then
                    let opened = RedBat.Spreadsheet.Vba.fireEvent "Workbook_Open" "" s.[0]
                    s |> Array.mapi (fun i m -> if i = 0 then opened else m)
                else s
            editor.Sheets.Value <- s
            editor.SheetNames.Value <- n
            editor.ActiveSheet.Value <- 0
            editor.CalcEngine.Value <- None
            editor.UndoStack.Value <- []
            editor.RedoStack.Value <- []
            currentPath.Value <- Some path
            editsSinceSave.Value <- 0
            rememberRecent path
            status.Value <- sprintf "Opened %s" (Path.GetFileName path)
        | None -> status.Value <- sprintf "Failed to open %s" path

    let state = shellState editor
    let actions, openFileAction, saveFileAction = shellActions editor status currentPath renderState loadFromPath

    let statusText () =
        let counts = Editor.selectionSummary editor
        match status.Value with
        | message when message <> "Ready" && message <> "" -> message + "   |   " + counts
        | _ -> counts

    let titleText () =
        let prefix = if editsSinceSave.Value > 0 then "• " else ""
        let file =
            match currentPath.Value with
            | Some path -> Path.GetFileName path
            | None -> "Untitled"
        sprintf "%sRedBat Calc - %s" prefix file

    let chrome = XlUi.chrome
    // Chrome bands: menu 32 + toolbar (32 * 3 rows) + formula bar 26 + sheet
    // tabs 30 + status 22. The grid itself is `expanded` so this is only the
    // hint used to size scroll content.
    let bodyOverhead = 32.0 + 96.0 + 26.0 + 30.0 + 22.0

    let initialWidth = int chrome.WindowWidth
    let initialHeight = 800
    let x11win = X11Window(titleText (), initialWidth, initialHeight)
    let window = x11win :> IWindow

    let controlServer = new ControlServer.ControlServer()

    let appState = {
        ShiftDown = false; CtrlDown = false; AltDown = false
        LayoutRoot = None
        Width = initialWidth; Height = initialHeight
        LastHover = None
        LastMouseX = 0.0; LastMouseY = 0.0
        LastCmds = []
        Window = window
        EditorState = editor
        RenderState = renderState
        Actions = actions
        OpenFile = openFileAction
        SaveFile = saveFileAction
        ControlServer = controlServer
    }

    let openGridContextMenu (x, y) =
        let item label action =
            RedBat.Office.Ui.menuItem (sprintf "ctx-%s" label) label None action
        let items = [
            item "Cut" actions.Cut
            item "Copy" actions.Copy
            item "Paste" actions.Paste
            item "Clear" (fun () ->
                let r, c = Editor.focusCell (Editor.activeModel editor)
                Editor.applyChange editor (clearCell { Row = r; Col = c }) |> ignore)
            item "Insert row" actions.InsertRow
            item "Insert column" actions.InsertColumn
            item "Delete row" actions.DeleteRow
            item "Delete column" actions.DeleteColumn
        ]
        Composit.Ui.OfficeShell.openContextMenu renderState x y items

    // Tools menu unique to the native shell — Preferences, the advanced
    // Format Cells dialog, and the About box. Bypasses
    // SpreadsheetUi.ShellActions so the Fable shell's ShellActions
    // record stays unchanged.
    let openPreferences () =
        XlUi.preferencesDialog (fun values ->
            status.Value <- sprintf "Saved %d preferences" (Map.count values))
        |> Composit.Ui.OfficeShell.openDialog renderState
    let openFormatCells () =
        XlUi.formatCellsAdvancedDialog (fun values ->
            let cat = Map.tryFind "number-category" values |> Option.defaultValue "general"
            status.Value <- sprintf "Format Cells applied (number=%s, %d fields)" cat (Map.count values))
        |> Composit.Ui.OfficeShell.openDialog renderState
    let openAbout () =
        XlUi.aboutDialog "0.1.0" "dev"
        |> Composit.Ui.OfficeShell.openDialog renderState
    let toolsMenu : RedBat.Office.Ui.MenuSpec = {
        TestId = "xl-tools-menu"
        PanelTestId = "xl-tools-menu-panel"
        Label = "Tools"
        Items = [
            RedBat.Office.Ui.menuItem "xl-menu-format-cells-advanced" "Format cells..." None openFormatCells
            RedBat.Office.Ui.menuItem "xl-menu-prefs"  "Preferences..."        None openPreferences
            RedBat.Office.Ui.menuItem "xl-menu-about" "About RedBat Calc..." None openAbout
        ]
    }

    let buildWidget () =
        let bodyHeight = max 200.0 (float appState.Height - bodyOverhead)
        let workspace =
            Chrome.workspace editor bodyHeight actions.InsertSheet
            |> onContextMenu openGridContextMenu
        Chrome.viewWith renderState (titleText ()) statusText state actions [ toolsMenu ] workspace

    Platform.requestRedraw <- fun () -> window.RequestRedraw()
    Platform.setTimeout <- fun ms cb -> window.SetTimeout(ms, cb)
    Platform.clearTimeout <- fun id -> window.ClearTimeout(id)
    Platform.onTextInputFocus <- fun _ -> ()
    Platform.onTextInputBlur <- fun () -> ()
    Platform.onCursorRectChange <- fun _ -> ()

    let rasterBackend = Raster.create()
    controlServer.ScreenshotProvider <- Some (fun () ->
        let cmds = appState.LastCmds
        if cmds.IsEmpty then Array.empty
        else
            rasterBackend.Init appState.Width appState.Height
            rasterBackend.Submit cmds
            let fb = rasterBackend.Finish()
            PngEncode.encode fb)
    controlServer.ClickAtHandler <- Some (fun (x, y) ->
        handleInput appState (MouseButton(0, true, x, y))
        handleInput appState (MouseButton(0, false, x, y)))
    controlServer.RightClickAtHandler <- Some (fun (x, y) ->
        handleInput appState (MouseButton(2, true, x, y))
        handleInput appState (MouseButton(2, false, x, y)))
    controlServer.MoveAtHandler <- Some (fun (x, y) ->
        handleInput appState (MouseMove(x, y)))
    controlServer.SendKeyHandler <- Some (fun keycode pressed ->
        if pressed then handleInput appState (KeyDown(keycode, 0u))
        else handleInput appState (KeyUp(keycode, 0u)))
    controlServer.ExitHandler <- Some (fun () -> window.Close())
    controlServer.OnCommandReceived <- Some (fun () -> window.RequestRedraw())
    controlServer.Start()

    Focus.request editor.FocusId

    window.Run(
        (fun ev -> handleInput appState ev),
        (fun w h -> renderFrame appState buildWidget w h))

    (controlServer :> System.IDisposable).Dispose()
    0
