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
    [
        Environment.GetEnvironmentVariable("RB_WORD_FONT_DIR")
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "artifacts/fonts/windows"))
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../artifacts/fonts/windows"))
    ]
    |> List.choose (fun dir ->
        if String.IsNullOrWhiteSpace dir then None
        else Path.Combine(dir, "calibri.ttf") |> existingFile)
    |> List.tryHead

let private initFont () =
    findFont ()
    |> Option.iter Composit.Ui.Layout.initFont

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

let private cellFormatFromStyle (style: Xlsx.CellStyle) : CellFormat =
    let hexColor (rgb: string) =
        let trimmed = rgb.TrimStart('#')
        // Excel often prefixes ARGB; drop alpha when 8 chars
        let cleaned = if trimmed.Length = 8 then trimmed.Substring(2) else trimmed
        if cleaned.Length = 6 then Some ("#" + cleaned) else None
    {
        Bold = style.Bold
        Italic = style.Italic
        Underline = false
        TextColor = style.FontColor |> Option.bind hexColor
        BgColor = style.FillColor |> Option.bind hexColor
        FontSize = style.FontSize |> Option.map int
        HorizontalAlign = None
        VerticalAlign = None
        WrapText = false
        NumberFormat = style.NumberFormat
        BorderTop = None
        BorderBottom = None
        BorderLeft = None
        BorderRight = None
    }

let private toModel (sheet: Xlsx.Worksheet) (workbook: Xlsx.Workbook) : SpreadsheetModel =
    let mutable maxRow = 0
    let mutable maxCol = 0
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
                    let format =
                        cell.Style
                        |> Option.map cellFormatFromStyle
                        |> Option.defaultValue defaultFormat
                    let cellData = { Value = value; Format = format; Comment = None }
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

    // Excel col.Width is in character units. The empirical ratio between
    // character units and CSS pixels at Calibri 11 / 96 DPI is ~8.0 px per
    // character for this corpus (column A in pandas-test1.xlsx: 11.83 chars
    // → Excel reports 71.25 pt = 95 px, so 95/11.83 ≈ 8.03).
    let charsToPixels (chars: float) = chars * 8.0

    // sheetFormatPr defaults. Excel uses these for every column without an
    // explicit <col width="..."> entry. Without seeding the defaults across
    // used columns, our renderer falls back to 100 px per column and drifts
    // ~12 px past every column beyond A.
    let defaultColWidthPx =
        match sheet.Format.DefaultColWidth with
        | Some w -> Some (charsToPixels w)
        | None ->
            // For baseColWidth Excel auto-adds about one character of margin
            // padding to the effective rendered width. With baseColWidth=10
            // pandas-test1 default columns render at 88 px ≈ 11 chars * 8.
            sheet.Format.BaseColWidth
            |> Option.map (fun baseW -> charsToPixels (baseW + 1.0))

    let defaultRowHeightPx =
        sheet.Format.DefaultRowHeight |> Option.map (fun h -> h * (96.0 / 72.0))

    let explicitColumns =
        sheet.Columns
        |> List.choose (fun col ->
            col.Width |> Option.map (fun w -> col.Index - 1, charsToPixels w))
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
        match defaultColWidthPx with
        | None -> explicitColumns
        | Some defWidth ->
            [ 0 .. maxCol ]
            |> List.fold (fun acc c ->
                if Map.containsKey c explicitColumns then acc
                else Map.add c defWidth acc) explicitColumns

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
        MergedCells = mergedCells }

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
