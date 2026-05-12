module SpreadsheetPdfRunner

open System
open System.IO
module Xlsx = XlsxReader

let private usage () =
    eprintfn "Usage: dotnet run --project tools/spreadsheet-pdf-runner -- <input.xlsx> <output.pdf> [sheetIndex]"
    eprintfn "Defaults: sheetIndex=0"

let private fullPath (path: string) =
    if Path.IsPathRooted path then path
    else Path.GetFullPath(path, Directory.GetCurrentDirectory())

let private existingDir path =
    if Directory.Exists path then Some path else None

let private fontDirs () =
    [
        Environment.GetEnvironmentVariable("RB_WORD_FONT_DIR")
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "artifacts/fonts/windows"))
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../artifacts/fonts/windows"))
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../office/artifacts/fonts/windows"))
    ]
    |> List.choose (fun p -> if String.IsNullOrWhiteSpace p then None else existingDir p)

let private tryFontFile fileName =
    fontDirs ()
    |> List.tryPick (fun dir ->
        let path = Path.Combine(dir, fileName)
        if File.Exists path then Some path else None)

let private tryFace faceName family bold italic fileName : RedBat.Spreadsheet.Pdf.FontFace option =
    tryFontFile fileName
    |> Option.map (fun path ->
        {
            FaceName = faceName
            Family = family
            Bold = bold
            Italic = italic
            Data = File.ReadAllBytes path
        })

let private pdfOptions (pageWidth: float) (pageHeight: float) : RedBat.Spreadsheet.Pdf.RenderOptions =
    let faces =
        [
            tryFace "Calibri" "Calibri" false false "calibri.ttf"
            tryFace "Calibri-Bold" "Calibri" true false "calibrib.ttf"
            tryFace "Calibri-Italic" "Calibri" false true "calibrii.ttf"
            tryFace "Calibri-BoldItalic" "Calibri" true true "calibriz.ttf"
            tryFace "Arial" "Arial" false false "arial.ttf"
            tryFace "Arial-Bold" "Arial" true false "arialbd.ttf"
            tryFace "Arial-Italic" "Arial" false true "ariali.ttf"
            tryFace "Arial-BoldItalic" "Arial" true true "arialbi.ttf"
            // Cambria ships as a TrueType collection (.ttc); FsPdf only
            // parses single-face TTFs today. Skip the regular weight; load
            // the standalone bold/italic faces.
            tryFace "Cambria-Bold" "Cambria" true false "cambriab.ttf"
            tryFace "Cambria-Italic" "Cambria" false true "cambriai.ttf"
            // PDF /Name tokens can't contain literal spaces — use a flat
            // face name internally and keep the spaceful Family for matching
            // the cell's `<name val="Times New Roman"/>`.
            tryFace "TimesNewRoman" "Times New Roman" false false "times.ttf"
            tryFace "TimesNewRoman-Bold" "Times New Roman" true false "timesbd.ttf"
            tryFace "TimesNewRoman-Italic" "Times New Roman" false true "timesi.ttf"
            tryFace "TimesNewRoman-BoldItalic" "Times New Roman" true true "timesbi.ttf"
        ]
        |> List.choose id

    {
        DefaultFontFamily = "Calibri"
        Fonts = faces
        PageWidth = pageWidth
        PageHeight = pageHeight
    }

let private renderPdf (inputPath: string) (outputPath: string) (sheetIndex: int) =
    let bytes = File.ReadAllBytes inputPath
    let workbook =
        match Xlsx.loadBytesNative bytes with
        | Some wb -> wb
        | None -> failwithf "Could not parse XLSX: %s" inputPath
    if sheetIndex < 0 || sheetIndex >= workbook.Sheets.Length then
        failwithf "Sheet index %d out of range (workbook has %d sheets)" sheetIndex workbook.Sheets.Length
    let sheet = workbook.Sheets.[sheetIndex]
    let model = ExcelNativeRunner.toModel sheet workbook

    // Letter portrait at 96 DPI in CSS pixels (8.5 × 11 inch → 816 × 1056).
    // Letter is the closest paper to what Excel print defaults look like.
    let pageWidth = 816.0
    let pageHeight = 1056.0

    let pdfBytes = RedBat.Spreadsheet.Pdf.toBytes (pdfOptions pageWidth pageHeight) model
    Directory.CreateDirectory(Path.GetDirectoryName outputPath) |> ignore
    File.WriteAllBytes(outputPath, pdfBytes)
    printfn "spreadsheet pdf runner ok: %s [sheet=%s cells=%d] -> %s (%d bytes)"
        inputPath sheet.Name (model.Cells.Count) outputPath pdfBytes.Length

[<EntryPoint>]
let main argv =
    try
        match argv with
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
