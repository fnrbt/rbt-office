module SpreadsheetVbaLayout

open System
open System.IO
open System.Globalization
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Layout
open Rbt.Spreadsheet.Paint
open Rbt.Spreadsheet.CellText
module Xlsx = XlsxReader

/// Tool that loads an XLSX and dumps a Spreadsheet-side per-cell layout
/// TSV in the same schema as `tools/excel-parity/extract-excel-layout.vbs`.
/// Lets `compare-vba-layout-results.mjs` diff our output against Excel's
/// own COM-reported `cell.Text` / `cell.Width` / etc. without bouncing
/// through Fable.

let private usage () =
    eprintfn "Usage: dotnet run --project tools/spreadsheet-vba-layout -- <input.xlsx> <output.tsv> [sheetIndex]"
    eprintfn "Defaults: sheetIndex = 0"

let private fullPath (p: string) =
    if Path.IsPathRooted p then p
    else Path.GetFullPath(p, Directory.GetCurrentDirectory())

let private columnLetters (col0: int) =
    // 0-indexed → A1 letters.
    let mutable n = col0 + 1
    let sb = System.Text.StringBuilder()
    while n > 0 do
        let r = (n - 1) % 26
        sb.Insert(0, char (int 'A' + r)) |> ignore
        n <- (n - 1) / 26
    sb.ToString()

let private valueRepr (v: CellValue) =
    match v with
    | CellValue.Empty -> ""
    | CellValue.Number n -> n.ToString("R", CultureInfo.InvariantCulture)
    | CellValue.String s -> s
    | CellValue.Boolean b -> if b then "TRUE" else "FALSE"
    | CellValue.Error e -> "#" + e
    | CellValue.Formula (expr, _) -> "=" + expr

let private formulaRepr (v: CellValue) =
    match v with
    | CellValue.Formula (expr, _) -> "=" + expr
    | _ -> ""

let private cachedValue (v: CellValue) =
    match v with
    | CellValue.Formula (_, cached) -> cached
    | other -> other

let private isMergedTrailing (model: SpreadsheetModel) (r: int) (c: int) =
    // True when the cell is *inside* a merged range but not the top-left.
    model.MergedCells
    |> List.exists (fun m ->
        r >= m.StartRow && r <= m.EndRow &&
        c >= m.StartCol && c <= m.EndCol &&
        not (r = m.StartRow && c = m.StartCol))

let private mergeFor (model: SpreadsheetModel) (r: int) (c: int) =
    model.MergedCells
    |> List.tryFind (fun m ->
        r >= m.StartRow && r <= m.EndRow &&
        c >= m.StartCol && c <= m.EndCol)

let private pxToPt (v: float) = v * 72.0 / 96.0

let private safe (s: string) =
    (s |> Option.ofObj |> Option.defaultValue "").Replace("\t", " ").Replace("\n", " ").Replace("\r", " ")

let private emitSheet (out: StreamWriter) (docName: string) (workbook: Xlsx.Workbook) (sheetIndex: int) =
    let sheet = workbook.Sheets.[sheetIndex]
    let model = ExcelNativeRunner.toModel sheet workbook
    let sheetName = sheet.Name

    // Cell-text measurer: feeds `Paint.cellDisplayText` real glyph widths
    // so date / number overflow decisions match what the PDF renderer
    // produces (and what Excel reports via COM cell.Text). Without a
    // measure function, cellDisplayText falls back to char-count heuristics
    // and hash-overflows dates that visually fit (e.g. pandas-testskiprows
    // C4 "2015-01-01 00:00:00" hashes by char-count but renders by measure).
    let fontCache = System.Collections.Generic.Dictionary<string option * bool * bool, _>()
    let getFont (familyName: string option) (bold: bool) (italic: bool) =
        let key = familyName, bold, italic
        match fontCache.TryGetValue key with
        | true, f -> f
        | _ ->
            let f = Rbt.Widgets.Composit.Ui.SpreadsheetGrid.rendererFontFor familyName bold italic
            fontCache.[key] <- f
            f
    let measureFor (format: CellFormat) =
        try
            let fontPt = format.FontSize |> Option.map float |> Option.defaultValue 11.0
            let fontPx = fontPt * 96.0 / 72.0
            let font = getFont format.FontName format.Bold format.Italic
            Some (fun (s: string) -> Rbt.Renderer.Text.measureWidth font fontPx s)
        with ex ->
            eprintfn "measureFor failed: %s" ex.Message
            None

    // Detect cells whose Bold flag was added by the table-style header
    // overlay rather than coming from the cell's actual font. Excel COM
    // reports `cell.Font.Bold = False` for these — the table-style bold is
    // a render-time overlay, not a baked font attribute. We read straight
    // off `model.Tables` (no workbook handle needed) since the runner now
    // translates Xlsx.TableDefinition onto the model. The bold overlay
    // fires for any TableStyleMedium-named table's header row.
    //
    // 1-indexed coordinates to match Model.Table.
    let isTableStyleHeaderBoldOverlay (r0: int) (c0: int) =
        let row1 = r0 + 1
        let col1 = c0 + 1
        model.Tables
        |> List.exists (fun t ->
            row1 >= t.StartRow && row1 < t.StartRow + t.HeaderRowCount
            && col1 >= t.StartCol && col1 <= t.EndCol
            && (match t.Style with
                | Some s -> s.Name.StartsWith "TableStyleMedium"
                | None -> false))

    // Determine the row/col range to dump. Excel's COM .UsedRange walks
    // every cell from A1 to the last non-empty row/col. We mirror that by
    // iterating 0..maxRow / 0..maxCol from the model's populated cells.
    let maxRow, maxCol =
        if Map.isEmpty model.Cells then 0, 0
        else
            let mutable mr = 0
            let mutable mc = 0
            for KeyValue(addr, _) in model.Cells do
                if addr.Row > mr then mr <- addr.Row
                if addr.Col > mc then mc <- addr.Col
            mr, mc

    // Walk used range. For each (r, c) compute the position cumulatively.
    let rowHeight r =
        Map.tryFind r model.RowHeights |> Option.defaultValue 20.0
    let colWidth c =
        Map.tryFind c model.ColumnWidths |> Option.defaultValue 64.0

    // Cumulative top y at each row index.
    let yAt r =
        let mutable y = 0.0
        for r0 in 0 .. r - 1 do y <- y + rowHeight r0
        y
    let xAt c =
        let mutable x = 0.0
        for c0 in 0 .. c - 1 do x <- x + colWidth c0
        x

    for r in 0 .. maxRow do
        for c in 0 .. maxCol do
            // Skip merged trailing cells — Excel reports geometry on the
            // top-left of the merge only.
            if not (isMergedTrailing model r c) then
                let addr = { Row = r; Col = c }
                let cellData = Map.tryFind addr model.Cells
                let value = cellData |> Option.map (fun d -> d.Value) |> Option.defaultValue CellValue.Empty
                let format = cellData |> Option.map (fun d -> d.Format) |> Option.defaultValue defaultFormat
                let mergeInfo = mergeFor model r c
                let mergedRows, mergedCols =
                    match mergeInfo with
                    | Some m -> m.EndRow - m.StartRow + 1, m.EndCol - m.StartCol + 1
                    | None -> 1, 1
                // Excel's `cell.Width` / `cell.Height` for a merged cell
                // returns the primary cell's OWN column-width / row-height,
                // not the full merge span. Match that — the comparison
                // harness can multiply via mergedRows/mergedCols if needed.
                let rectWidth = colWidth c
                let rectHeight = rowHeight r
                let fontPt = format.FontSize |> Option.map float |> Option.defaultValue 11.0
                let displayText =
                    let cached = cachedValue value
                    cellDisplayText (measureFor format) rectWidth format cached
                // Match `extract-excel-layout.vbs`'s IsEmptyCell filter:
                // emit only cells with a non-empty value OR a formula. Cells
                // that exist only for their col/cell formatting (e.g.
                // `<c r="A1" s="1"/>` with no <v>) don't appear in Excel's
                // UsedRange enumeration there, so they shouldn't appear in
                // ours either. Empty merge primaries also skip — VBS sees
                // them as empty and skips them too (pandas-df_header_oob
                // has B1:C1 merged with no value, and Excel doesn't emit
                // anything for that merge).
                let hasFormula =
                    match value with
                    | CellValue.Formula _ -> true
                    | _ -> false
                let isInteresting =
                    (value <> CellValue.Empty)
                    || hasFormula
                if isInteresting then
                    let row =
                        [ docName
                          sheetName
                          string (sheetIndex + 1)
                          string (r + 1)
                          string (c + 1)
                          columnLetters c + string (r + 1)
                          (pxToPt (xAt c)).ToString("F2", CultureInfo.InvariantCulture)
                          (pxToPt (yAt r)).ToString("F2", CultureInfo.InvariantCulture)
                          (pxToPt rectWidth).ToString("F2", CultureInfo.InvariantCulture)
                          (pxToPt rectHeight).ToString("F2", CultureInfo.InvariantCulture)
                          string mergedRows
                          string mergedCols
                          (pxToPt (rowHeight r)).ToString("F2", CultureInfo.InvariantCulture)
                          (pxToPt (colWidth c)).ToString("F2", CultureInfo.InvariantCulture)
                          (if format.WrapText then "true" else "false")
                          (format.HorizontalAlign |> Option.defaultValue "general")
                          (format.VerticalAlign |> Option.defaultValue "bottom")
                          (format.FontName |> Option.defaultValue "Calibri") |> safe
                          (fontPt.ToString("F2", CultureInfo.InvariantCulture))
                          (if format.Bold
                              && not (isTableStyleHeaderBoldOverlay r c)
                              && not (model.CfBoldCells.Contains addr) then "true" else "false")
                          (if format.Italic then "true" else "false")
                          (format.NumberFormat |> Option.defaultValue "General") |> safe
                          ""  // style — not surfaced
                          formulaRepr value |> safe
                          valueRepr value |> safe
                          displayText |> safe ]
                    out.WriteLine(String.concat "\t" row)
    eprintfn "spreadsheet vba layout: sheet=%s rows=%d cols=%d"
        sheetName (maxRow + 1) (maxCol + 1)

let private dump (inputPath: string) (outputPath: string) (sheetIndex: int option) =
    let bytes = File.ReadAllBytes inputPath
    let workbook =
        match Xlsx.loadBytesNative bytes with
        | Some wb -> wb
        | None -> failwithf "Could not parse XLSX: %s" inputPath
    let docName = Path.GetFileName inputPath
    use out = new StreamWriter(outputPath, false, System.Text.Encoding.UTF8)
    let header =
        [ "doc"; "sheet"; "sheetIndex"; "row"; "column"; "address"
          "x"; "y"; "width"; "height"
          "mergedRows"; "mergedCols"
          "rowHeight"; "columnWidth"
          "wrapText"; "hAlign"; "vAlign"
          "fontName"; "fontSize"; "bold"; "italic"
          "numberFormat"; "style"; "formula"; "value"; "text" ]
    out.WriteLine(String.concat "\t" header)
    match sheetIndex with
    | Some idx ->
        if idx < 0 || idx >= workbook.Sheets.Length then
            failwithf "Sheet index %d out of range (workbook has %d sheets)" idx workbook.Sheets.Length
        emitSheet out docName workbook idx
    | None ->
        // Match extract-excel-layout.vbs's "For Each sheet In Worksheets"
        // loop — emit every visible sheet. Default for the 2-arg invocation.
        for idx in 0 .. workbook.Sheets.Length - 1 do
            emitSheet out docName workbook idx
    printfn "spreadsheet vba layout: %s -> %s" inputPath outputPath

[<EntryPoint>]
let main argv =
    try
        match argv with
        | [| input; output |] ->
            dump (fullPath input) (fullPath output) None
            0
        | [| input; output; sheetIdx |] ->
            dump (fullPath input) (fullPath output) (Some (int sheetIdx))
            0
        | _ ->
            usage ()
            2
    with ex ->
        eprintfn "%O" ex
        1
