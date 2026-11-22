module SpreadsheetPdfRunner

open System
open System.IO
module Xlsx = XlsxReader
module Charts = Rbt.Spreadsheet.Charts

let private usage () =
    eprintfn "Usage: dotnet run --project tools/spreadsheet-pdf-runner -- <input.xlsx> <output.pdf> [sheetIndex]"
    eprintfn "Defaults: sheetIndex=0"

let private fullPath (path: string) =
    if Path.IsPathRooted path then path
    else Path.GetFullPath(path, Directory.GetCurrentDirectory())

/// The faces of the PDF: those of the families the sheet's text is set in (`Pdf.requiredFaces`, by the names the
/// workbook gives them) that the font directories have (RB_WORD_FONT_DIR: Aptos Narrow, Tahoma, Segoe UI Black
/// in their own faces; a family the machine has none of is set in Calibri), and Calibri, Arial, Cambria, Times
/// New Roman and Verdana.
let private pdfOptions (sheetName: string) (fileName: string) (required: (string * bool * bool) array) =
    let families =
        [ "Calibri"; "Arial"; "Cambria"; "Times New Roman"; "Verdana" ]
        @ (required |> Array.map (fun (family, _, _) -> family) |> Array.filter Rbt.OfficeNative.Fonts.hasFamily |> List.ofArray)
        |> List.distinct
    let faces : Rbt.Spreadsheet.Pdf.Render.FontFace list =
        [ for family in families do
            for bold, italic in [ false, false; true, false; false, true; true, true ] do
                let font = Rbt.OfficeNative.Fonts.resolve (Some family) bold italic
                let suffix =
                    if bold && italic then "-BoldItalic"
                    elif bold then "-Bold"
                    elif italic then "-Italic"
                    else ""
                yield {
                    FaceName = family.Replace(" ", "") + suffix
                    Family = family
                    Bold = bold
                    Italic = italic
                    Data = font.OtFont.Data
                } ]
    let cjkFallback : Rbt.Spreadsheet.Pdf.Render.FontFace list =
        [ for bold in [ false; true ] do
            for name, font in Rbt.OfficeNative.Fonts.fallbackFaces bold do
                yield {
                    FaceName = Path.GetFileNameWithoutExtension(name) + (if bold then "-Bold" else "")
                    Family = "_cjk"
                    Bold = bold
                    Italic = false
                    Data = font.OtFont.Data
                } ]
    ({
        DefaultFontFamily = "Calibri"
        Fonts = faces
        CjkFallback = cjkFallback
        SheetName = sheetName
        FileName = fileName
    } : Rbt.Spreadsheet.Pdf.Render.RenderOptions)

/// The charts and shapes of a sheet of a workbook as overlays of its PDF, by the layout of the grid.
let private drawingOverlays (workbook: Xlsx.Workbook) (context: Rbt.Spreadsheet.Workbook.Workbook) (sheetIndex: int) =
    let text (data: byte[]) = System.Text.Encoding.UTF8.GetString data
    let parts = workbook.Preserved.Parts |> Map.ofList
    let theme =
        workbook.Preserved.ThemePath
        |> Option.bind (fun path -> Map.tryFind path parts)
        |> Option.map (text >> Charts.parseTheme)
        |> Option.defaultValue Charts.defaultTheme
    let models =
        workbook.Sheets
        |> List.mapi (fun index sheet ->
            sheet.Name, lazy ((Map.find (sprintf "sheet%d" (index + 1)) context.Sheets).Model))
    let data : Charts.ChartData =
        { Cells = (fun name r0 c0 r1 c1 ->
            models
            |> List.tryFind (fun (known, _) -> String.Equals (known, name, StringComparison.OrdinalIgnoreCase))
            |> Option.map (fun (_, sheet) ->
                let cells = sheet.Value
                Array.init (r1 - r0 + 1) (fun r -> Array.init (c1 - c0 + 1) (fun c -> Rbt.Spreadsheet.Model.getCell { Row = r0 + r; Col = c0 + c } cells))))
          Display = (fun cell -> Rbt.Spreadsheet.CellText.cellDisplayText None 1000.0 cell.Format cell.Value)
          Format = (fun code value -> Rbt.Spreadsheet.CellText.formatNumberWithCodeIn false code value) }
    let target (id: string) =
        workbook.Sheets.[sheetIndex].Extras.DrawingRelationships
        |> List.tryPick (fun (known, _, target, isExternal) -> if known = id && not isExternal then Some (ZipInterop.normalizePath target) else None)
    let partOf (id: string) = target id |> Option.bind (fun path -> Map.tryFind path parts |> Option.map (fun bytes -> path, bytes))
    let drawing : Charts.DrawingParts =
        { Theme = theme
          Chart = (fun id ->
            partOf id |> Option.bind (fun (path, bytes) ->
                Charts.parseChart theme (text bytes)
                |> Option.map (Charts.withUserShapes (fun name -> Map.tryFind name parts |> Option.map text) path)))
          Picture = (fun id -> partOf id |> Option.map (fun (name, bytes) -> "data:image/" + Path.GetExtension(name).TrimStart('.') + ";base64," + Convert.ToBase64String bytes))
          Slicer = Rbt.Spreadsheet.Slicers.drawer workbook }
    let drawn = (snd models.[sheetIndex]).Value
    // (Anchored to cells: as wide as those cells print.)
    drawn.Drawings, Rbt.Spreadsheet.Paint.drawingOverlays (Charts.objectScene drawing data) (Rbt.Spreadsheet.Pdf.Pagination.withPrintSizes drawn)

let private renderPdf (inputPath: string) (outputPath: string) (sheetIndex: int) =
    let bytes = File.ReadAllBytes inputPath
    let workbook =
        match Xlsx.loadBytesNative bytes with
        | Some wb -> wb
        | None -> failwithf "Could not parse XLSX: %s" inputPath
    if sheetIndex < 0 || sheetIndex >= workbook.Sheets.Length then
        failwithf "Sheet index %d out of range (workbook has %d sheets)" sheetIndex workbook.Sheets.Length
    let sheet = workbook.Sheets.[sheetIndex]
    // (The bridge's model as the browser has it - rows fitted on opening as Excel 2019 fits them, columns at its
    // widths: ExcelNativeRunner's native geometry fits wrapped rows with the renderer's widths, a line more than
    // Excel's on acre's MF sheets.)
    let context = (Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook workbook).Context
    let model = (Map.find (sprintf "sheet%d" (sheetIndex + 1)) context.Sheets).Model
    // Page size + margins now live on model.PageSetup (defaulted to US
    // Letter portrait when the source file didn't specify a pageSetup).
    // Excel's `&F` substitutes the workbook's display name (no extension)
    // on Windows; match that so our footer reads "synthetic-print" rather
    // than "synthetic-print.xlsx".
    // Charts and shapes as the grid draws them, anchored to their cells.
    let _, overlays = drawingOverlays workbook context sheetIndex
    let required = Rbt.Spreadsheet.Pdf.Render.requiredFaces "Calibri" [| Rbt.Spreadsheet.Pdf.Render.pdfSheet model sheet.Name overlays |]
    let opts = pdfOptions sheet.Name (Path.GetFileNameWithoutExtension inputPath) required
    let pdfBytes = Rbt.Spreadsheet.Pdf.Render.toBytesWith opts model overlays
    Directory.CreateDirectory(Path.GetDirectoryName outputPath) |> ignore
    File.WriteAllBytes(outputPath, pdfBytes)
    printfn "spreadsheet pdf runner ok: %s [sheet=%s cells=%d] -> %s (%d bytes)"
        inputPath sheet.Name (model.Cells.Count) outputPath pdfBytes.Length

[<EntryPoint>]
let main argv =
    try
        match argv with
        | [| "--vectortest"; output |] ->
            let bytes = Rbt.Spreadsheet.Pdf.Render.sceneToBytes 400.0 280.0 Rbt.Spreadsheet.Paint.vectorTestScene
            System.IO.File.WriteAllBytes(fullPath output, bytes)
            0
        | [| input; output |] ->
            renderPdf (fullPath input) (fullPath output) 0
            0
        | [| input; output; sheetIdx |] ->
            renderPdf (fullPath input) (fullPath output) (int sheetIdx)
            0
        | _ ->
            usage ()
            2
    with ex ->
        eprintfn "%O" ex
        1
