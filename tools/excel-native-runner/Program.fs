module ExcelNativeRunner

open System
open System.IO
open Renderer

module Xlsx = XlsxReader

open RedBat.Spreadsheet.Model

let private usage () =
    eprintfn "Usage: dotnet run --project tools/excel-native-runner -- <input.xlsx> <output.png> [width] [height] [sheetIndex]"
    eprintfn "Defaults: width=1100 height=720 sheetIndex=0"

let private fullPath (path: string) =
    if Path.IsPathRooted path then path
    else Path.GetFullPath(path, Directory.GetCurrentDirectory())

let private existingFile path =
    if File.Exists path then Some path else None

let private findFont () =
    let cwd = Directory.GetCurrentDirectory()
    [
        Environment.GetEnvironmentVariable("RB_WORD_FONT_DIR")
        Path.GetFullPath(Path.Combine(cwd, "artifacts/fonts/windows"))
        Path.GetFullPath(Path.Combine(cwd, "../../artifacts/fonts/windows"))
        // Sibling office repo's font cache, when running from office-spreadsheet/.
        // SpreadsheetGrid.fontDirs already includes this — keep the runner's
        // own font lookup in sync so the default UI font (row/column headers,
        // chrome) loads from the same Calibri Windows TTF Excel rendered with.
        Path.GetFullPath(Path.Combine(cwd, "../office/artifacts/fonts/windows"))
    ]
    |> List.choose (fun dir ->
        if String.IsNullOrWhiteSpace dir then None
        else Path.Combine(dir, "calibri.ttf") |> existingFile)
    |> List.tryHead

let private initFont () =
    match findFont () with
    | Some path -> Composit.Ui.Layout.initFont path
    | None ->
        eprintfn "warning: calibri.ttf not found in RB_WORD_FONT_DIR or known font dirs; row/column headers will fall back to NotoSans"

let private columnIndex (letters: string) : int =
    let mutable acc = 0
    for ch in letters do
        if Char.IsLetter ch then
            acc <- acc * 26 + (int (Char.ToUpper ch) - int 'A' + 1)
    acc - 1

let private resolveValue (workbook: Xlsx.Workbook) (value: Xlsx.CellValue) : CellValue =
    match value with
    | Xlsx.Number n -> Number n
    | Xlsx.String s -> String s
    | Xlsx.SharedString idx when idx >= 0 && idx < workbook.SharedStrings.Length ->
        String workbook.SharedStrings.[idx]
    | Xlsx.SharedString _ -> Empty
    | Xlsx.Boolean b -> Boolean b
    | Xlsx.Date s -> String s
    | Xlsx.Formula expr -> Formula (expr, Empty)
    | Xlsx.Empty -> Empty

/// Translate an Xlsx.RunProps into a RunOverride. Drops empty fields so the
/// renderer can fall back to the cell's base CellFormat.
let private runOverrideOf (props: Xlsx.RunProps) : RunOverride =
    let boolOpt v = if v then Some true else None
    let hexColor (rgb: string) =
        let trimmed = rgb.TrimStart('#')
        if trimmed.Length = 6 then Some ("#" + trimmed) else None
    {
        Bold = boolOpt props.Bold
        Italic = boolOpt props.Italic
        Underline = boolOpt props.Underline
        Strikethrough = boolOpt props.Strikethrough
        Color = props.Color |> Option.bind hexColor
        FontName = props.FontName
        FontSize = props.FontSize
    }

/// Look up rich-text runs for a cell whose value is a shared-string ref.
/// Returns Some runs only when the SI has more than one run (single-run
/// shared strings are fully described by the cell's base format).
let private cellRichRuns (workbook: Xlsx.Workbook) (value: Xlsx.CellValue) : (string * RunOverride) list option =
    match value with
    | Xlsx.SharedString idx when
        idx >= 0 && idx < workbook.SharedStringRuns.Length
        && not (List.isEmpty workbook.SharedStringRuns.[idx]) ->
        let runs = workbook.SharedStringRuns.[idx]
        if List.length runs <= 1 then None
        else Some (runs |> List.map (fun (text, props) -> text, runOverrideOf props))
    | _ -> None

let private cellFormatFromStyle (style: Xlsx.CellStyle) : CellFormat =
    let hexColor (rgb: string) =
        let trimmed = rgb.TrimStart('#')
        // Excel often prefixes ARGB; drop alpha when 8 chars
        let cleaned = if trimmed.Length = 8 then trimmed.Substring(2) else trimmed
        if cleaned.Length = 6 then Some ("#" + cleaned) else None
    let borderSide (side: Xlsx.CellBorderSide option) : BorderStyle option =
        side
        |> Option.map (fun s ->
            {
                Style = s.Style
                Color = s.Color |> Option.bind hexColor
            })
    {
        Bold = style.Bold
        Italic = style.Italic
        Underline = style.Underline
        Strikethrough = style.Strikethrough
        FontName = style.FontName
        TextColor = style.FontColor |> Option.bind hexColor
        BgColor = style.FillColor |> Option.bind hexColor
        FontSize = style.FontSize |> Option.map int
        HorizontalAlign = style.HorizontalAlign
        VerticalAlign = style.VerticalAlign
        WrapText = style.WrapText
        NumberFormat = style.NumberFormat
        BorderTop = borderSide style.BorderTop
        BorderBottom = borderSide style.BorderBottom
        BorderLeft = borderSide style.BorderLeft
        BorderRight = borderSide style.BorderRight
        FilterButton = false
    }

    // Layer the workbook's table style (e.g. `TableStyleMedium7`) on top of
    // the cellXf-derived CellFormat. Excel renders the table header row with
    // a solid accent fill + white bold text, and (when ShowRowStripes is set)
    // tints alternating data rows with a very light blend of the same accent.
    // We only override fields the cellXf didn't already set, so explicit
    // colors in the file still win.
let private tintHex (hex: string) (t: float) =
    let parseByte (s: string) =
        try System.Convert.ToByte(s, 16) with _ -> 0uy
    let lerp (c: byte) =
        let f = float c
        byte (int (System.Math.Round(f + t * (255.0 - f))))
    if hex.Length < 6 then hex
    else
        sprintf "#%02X%02X%02X"
            (lerp (parseByte (hex.Substring(0,2))))
            (lerp (parseByte (hex.Substring(2,2))))
            (lerp (parseByte (hex.Substring(4,2))))

let private styleNumber (name: string) =
    let digits = name |> Seq.skipWhile (System.Char.IsLetter) |> Seq.toArray |> System.String
    match System.Int32.TryParse digits with
    | true, n -> Some n
    | _ -> None

/// Map a `TableStyleMediumN` name to an accent color from the workbook theme.
/// Medium2-7 → accent1-6; higher Medium variants cycle through the same six
/// accents (the rendering differences between Medium7/Medium14/Medium21 are
/// header banding details we don't reproduce yet).
let private tableAccentColor (themeAccents: string array) (styleName: string) : string option =
    let n = styleNumber styleName
    match n with
    | Some n when n >= 2 ->
        let idx = (n - 2) % 6
        if idx < themeAccents.Length then Some ("#" + themeAccents.[idx]) else None
    | _ -> None

type private TableOverlay = {
    BgColor: string option
    TextColor: string option
    Bold: bool
    /// Set on header cells when the table declares an <autoFilter> covering
    /// this column. Renderer paints a dropdown triangle on the right side.
    FilterButton: bool
}

/// For each cell coordinate inside a table range, return the table-derived
/// overlay (background + bold + font color). Excel's TableStyleMedium
/// variants don't draw heavy inner borders — the accent fill + banded
/// stripes are the only visual separation, with the workbook's default
/// gridlines showing through where they survive the fill. Returns None
/// for cells outside any table.
let private tableOverrideFor (tables: Xlsx.TableDefinition list) (themeAccents: string array) (row: int) (col: int) : TableOverlay option =
    // row/col are 1-indexed to match Xlsx.TableDefinition.
    tables
    |> List.tryPick (fun t ->
        if row < t.StartRow || row > t.EndRow || col < t.StartCol || col > t.EndCol then None
        else
            match t.Style with
            | None -> None
            | Some style ->
                // Light/Dark variants don't carry the same accent banding;
                // restrict the overlay to the Medium family for now.
                if not (style.Name.StartsWith "TableStyleMedium") then None
                else
                    match tableAccentColor themeAccents style.Name with
                    | None -> None
                    | Some accentHex ->
                        let isHeader = row < t.StartRow + t.HeaderRowCount
                        if isHeader then
                            let hasFilter =
                                match t.AutoFilter with
                                | Some (fc1, fc2) -> col >= fc1 && col <= fc2
                                | None -> false
                            Some {
                                BgColor = Some accentHex
                                TextColor = Some "#FFFFFF"
                                Bold = true
                                FilterButton = hasFilter }
                        elif style.ShowRowStripes then
                            // Excel bands every other data row. Pixel-sampled
                            // against artifacts/excel/pandas-test_decimal-
                            // headings.png: the first data row (immediately
                            // below the header) renders at #E2EFDA (light
                            // accent tint); the second is white; alternates.
                            let dataIdx = row - t.StartRow - t.HeaderRowCount
                            if dataIdx % 2 = 0 then
                                Some {
                                    BgColor = Some (tintHex (accentHex.TrimStart('#')) 0.80)
                                    TextColor = None
                                    Bold = false
                                    FilterButton = false }
                            else None
                        else None)

let toModel (sheet: Xlsx.Worksheet) (workbook: Xlsx.Workbook) : SpreadsheetModel =
    let mutable maxRow = 0
    let mutable maxCol = 0
    let mutable richRunsMap : Map<CellAddress, (string * RunOverride) list> = Map.empty
    let cells =
        sheet.Rows
        |> List.collect (fun row ->
            row.Cells
            |> List.choose (fun cell ->
                let r = cell.Reference.Row - 1
                let c = columnIndex cell.Reference.Column
                if r < 0 || c < 0 then None
                else
                    let value = resolveValue workbook cell.Value
                    let baseFormat =
                        cell.Style
                        |> Option.map cellFormatFromStyle
                        |> Option.defaultValue defaultFormat
                    let format =
                        match tableOverrideFor sheet.Tables workbook.ThemeAccents cell.Reference.Row (c + 1) with
                        | None -> baseFormat
                        | Some overlay ->
                            { baseFormat with
                                BgColor = baseFormat.BgColor |> Option.orElse overlay.BgColor
                                TextColor = baseFormat.TextColor |> Option.orElse overlay.TextColor
                                Bold = baseFormat.Bold || overlay.Bold
                                FilterButton = baseFormat.FilterButton || overlay.FilterButton }
                    let cellData = { Value = value; Format = format; Comment = None }
                    match cellRichRuns workbook cell.Value with
                    | Some runs -> richRunsMap <- Map.add { Row = r; Col = c } runs richRunsMap
                    | None -> ()
                    if r > maxRow then maxRow <- r
                    if c > maxCol then maxCol <- c
                    Some ({ Row = r; Col = c }, cellData)))
        |> Map.ofList

    let mergedCells =
        sheet.MergedCells
        |> List.map (fun mc ->
            {
                StartRow = mc.StartRef.Row - 1
                StartCol = columnIndex mc.StartRef.Column
                EndRow = mc.EndRef.Row - 1
                EndCol = columnIndex mc.EndRef.Column
            })

    // Excel col.Width is in character units. The chars-to-pixels ratio is
    // the "Maximum Digit Width" (MDW) of the workbook's default font —
    // technically the rendered pixel width of the widest digit 0-9 in that
    // font at the workbook's default size, as Excel computes it. Use the
    // real loaded font instead of a Calibri-anchored linear fit, since
    // fonts vary enough at the same point size to skew column widths
    // (Verdana 10 needs 8 px/char, not the 7 a Calibri-only estimate gives).
    let defaultFontSize =
        if workbook.Styles.Length > 0 then
            workbook.Styles.[0].FontSize |> Option.defaultValue 11.0
        else 11.0
    let defaultFontName =
        if workbook.Styles.Length > 0 then
            workbook.Styles.[0].FontName
        else None
    let mdw =
        try
            let font = Composit.Ui.SpreadsheetGrid.rendererFontFor defaultFontName false false
            let pxSize = defaultFontSize * 96.0 / 72.0
            // Excel's MDW is max-of-digits rounded up to the next integer.
            // We approximate by measuring the literal '0' (typically the
            // widest digit in monospaced-figure fonts like Calibri/Verdana)
            // and ceiling. Floor at 7 px so the Calibri-anchored fallback
            // chain doesn't drift sub-pixel.
            let zero = Renderer.Text.measureWidth font pxSize "0"
            max 7.0 (System.Math.Ceiling zero)
        with _ ->
            // Renderer unavailable (font cache miss, etc.) — fall back to
            // the original Calibri-anchored linear estimate.
            max 7.0 (System.Math.Round(defaultFontSize * 7.0 / 11.0))
    let charsToPixels (chars: float) = chars * mdw

    // sheetFormatPr defaults. Excel uses these for every column without an
    // explicit <col width="..."> entry. Without seeding the defaults across
    // used columns, our renderer falls back to 100 px per column and drifts
    // every column past A. Excel's implicit default when sheetFormatPr is
    // absent is 8.43 character units (the same value Excel writes when you
    // explicitly request "AutoFit Column Width" on an unsized column).
    let defaultColWidthPx =
        match sheet.Format.DefaultColWidth with
        | Some w -> charsToPixels w
        | None ->
            match sheet.Format.BaseColWidth with
            // For baseColWidth Excel auto-adds about one character of margin
            // padding to the effective rendered width. With baseColWidth=10
            // pandas-test1 default columns render at 88 px ≈ 11 chars * 8.
            | Some baseW -> charsToPixels (baseW + 1.0)
            | None -> charsToPixels 8.43

    let defaultRowHeightPx =
        sheet.Format.DefaultRowHeight |> Option.map (fun h -> h * (96.0 / 72.0))

    // <col min=M max=N width=W hidden=1/> applies W to every column in M..N
    // (inclusive). Hidden columns collapse to width=0 regardless of the
    // declared width — Excel hides them by zero-width, not by skipping the
    // entry, so a hidden col with width=11 still renders as 0 px. poi-51585
    // has 748 hidden cols in its template; without this they'd push the
    // real data way off-screen.
    let explicitColumns =
        sheet.Columns
        |> List.collect (fun col ->
            let widthPx =
                if col.Hidden then Some 0.0
                else col.Width |> Option.map charsToPixels
            match widthPx with
            | None -> []
            | Some px -> [ for idx in col.Index .. col.MaxIndex -> idx - 1, px ])
        |> Map.ofList

    let explicitRows =
        sheet.Rows
        |> List.choose (fun row ->
            row.Height |> Option.map (fun h -> row.Index - 1, h * (96.0 / 72.0)))
        |> Map.ofList

    // Seed every used row/column with the sheet default so unspecified
    // entries no longer fall back to the generic 100 px / 25 px model
    // defaults.
    let columnWidths =
        [ 0 .. maxCol ]
        |> List.fold (fun acc c ->
            if Map.containsKey c explicitColumns then acc
            else Map.add c defaultColWidthPx acc) explicitColumns

    let rowHeights =
        match defaultRowHeightPx with
        | None -> explicitRows
        | Some defHeight ->
            [ 0 .. maxRow ]
            |> List.fold (fun acc r ->
                if Map.containsKey r explicitRows then acc
                else Map.add r defHeight acc) explicitRows

    let baseModel = createEmpty (max 1 (maxRow + 1)) (max 1 (maxCol + 1))
    // Excel's printed/screenshot output never shows the active-cell border, so
    // park the selection off-grid for snapshot renders.
    { baseModel with
        Cells = cells
        Selection = SingleCell(-1, -1)
        RowHeights = rowHeights
        ColumnWidths = columnWidths
        MergedCells = mergedCells
        RichRuns = richRunsMap
        FrozenRows = sheet.FrozenPane |> Option.map (fun p -> p.YSplit) |> Option.defaultValue 0
        FrozenCols = sheet.FrozenPane |> Option.map (fun p -> p.XSplit) |> Option.defaultValue 0 }

let private escapeJson (value: string) =
    let sb = System.Text.StringBuilder(value.Length + 2)
    sb.Append('"') |> ignore
    for ch in value do
        match ch with
        | '"' -> sb.Append("\\\"") |> ignore
        | '\\' -> sb.Append("\\\\") |> ignore
        | '\n' -> sb.Append("\\n") |> ignore
        | '\r' -> sb.Append("\\r") |> ignore
        | '\t' -> sb.Append("\\t") |> ignore
        | c when int c < 0x20 -> sb.AppendFormat("\\u{0:x4}", int c) |> ignore
        | c -> sb.Append(c) |> ignore
    sb.Append('"') |> ignore
    sb.ToString()

let private cellInfoJson (sheetName: string) (model: SpreadsheetModel) =
    let columnLetter col = RedBat.Spreadsheet.Model.columnToLetter col
    let viewport = RedBat.Spreadsheet.Layout.defaultViewport 0.0 0.0
    let records =
        model.Cells
        |> Map.toList
        |> List.map (fun (addr, data) ->
            let rect = RedBat.Spreadsheet.Layout.cellRect model viewport addr.Row addr.Col
            let display =
                match data.Value with
                | Number n -> string n
                | String s -> s
                | Boolean b -> if b then "TRUE" else "FALSE"
                | Formula (_, _) -> ""
                | Empty -> ""
                | Error e -> e
            let fields = [
                "\"Sheet\":" + escapeJson sheetName
                "\"Address\":" + escapeJson (sprintf "%s%d" (columnLetter addr.Col) (addr.Row + 1))
                "\"Row\":" + string (addr.Row + 1)
                "\"Column\":" + string (addr.Col + 1)
                "\"X\":" + sprintf "%.2f" (rect.X - viewport.RowHeaderWidth)
                "\"Y\":" + sprintf "%.2f" (rect.Y - viewport.HeaderHeight)
                "\"Width\":" + sprintf "%.2f" rect.Width
                "\"Height\":" + sprintf "%.2f" rect.Height
                "\"Value\":" + escapeJson display
                "\"Text\":" + escapeJson display
                "\"NumberFormat\":" + escapeJson (data.Format.NumberFormat |> Option.defaultValue "")
            ]
            "{ " + String.concat ", " fields + " }")
    "[\n  " + String.concat ",\n  " records + "\n]\n"

#if EXCEL_NATIVE_RUNNER_MAIN
let private renderXlsx (inputPath: string) (outputPath: string) (width: int) (height: int) (sheetIndex: int) (jsonPath: string option) =
    initFont ()
    let bytes = File.ReadAllBytes inputPath
    let workbook =
        match Xlsx.loadBytesNative bytes with
        | Some parsed -> parsed
        | None -> failwithf "Could not parse XLSX: %s" inputPath

    if workbook.Sheets.IsEmpty then
        failwithf "No sheets in workbook: %s" inputPath

    let chosenIndex = min (max 0 sheetIndex) (workbook.Sheets.Length - 1)
    let sheet = workbook.Sheets.[chosenIndex]
    let model = toModel sheet workbook

    let widget = Composit.Ui.SpreadsheetGrid.widget (float width) (float height) model
    let backend = Raster.create()
    let framebuffer = Composit.Ui.Render.renderWithBg Color.white backend width height widget
    let png = PngEncode.encode framebuffer

    Directory.CreateDirectory(Path.GetDirectoryName outputPath) |> ignore
    File.WriteAllBytes(outputPath, png)

    match jsonPath with
    | Some path ->
        Directory.CreateDirectory(Path.GetDirectoryName path) |> ignore
        File.WriteAllText(path, cellInfoJson sheet.Name model)
    | None -> ()

    let cellCount = Map.count model.Cells
    printfn "excel native runner ok: %s [sheet=%s cells=%d] -> %s (%dx%d)" inputPath sheet.Name cellCount outputPath width height

[<EntryPoint>]
let main argv =
    try
        let parseInt fallback (value: string) =
            match Int32.TryParse value with
            | true, parsed -> parsed
            | false, _ -> fallback
        let jsonEnv = Environment.GetEnvironmentVariable("EXCEL_NATIVE_JSON")
        let jsonOpt = if String.IsNullOrWhiteSpace jsonEnv then None else Some (fullPath jsonEnv)
        match argv with
        | [| input; output |] ->
            renderXlsx (fullPath input) (fullPath output) 1100 720 0 jsonOpt
            0
        | [| input; output; width |] ->
            renderXlsx (fullPath input) (fullPath output) (parseInt 1100 width) 720 0 jsonOpt
            0
        | [| input; output; width; height |] ->
            renderXlsx (fullPath input) (fullPath output) (parseInt 1100 width) (parseInt 720 height) 0 jsonOpt
            0
        | [| input; output; width; height; sheet |] ->
            renderXlsx (fullPath input) (fullPath output) (parseInt 1100 width) (parseInt 720 height) (parseInt 0 sheet) jsonOpt
            0
        | [| input; output; width; height; sheet; jsonOut |] ->
            renderXlsx (fullPath input) (fullPath output) (parseInt 1100 width) (parseInt 720 height) (parseInt 0 sheet) (Some (fullPath jsonOut))
            0
        | _ ->
            usage()
            2
    with ex ->
        eprintfn "%O" ex
        1
#endif
