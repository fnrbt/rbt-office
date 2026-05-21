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

let private tryFace faceName family bold italic fileName : RedBat.Spreadsheet.Pdf.Render.FontFace option =
    tryFontFile fileName
    |> Option.map (fun path ->
        {
            FaceName = faceName
            Family = family
            Bold = bold
            Italic = italic
            Data = File.ReadAllBytes path
        })

let private pdfOptions (sheetName: string) (fileName: string) : RedBat.Spreadsheet.Pdf.Render.RenderOptions =
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
            // Cambria ships as a TrueType collection (.ttc); FsPdf now
            // pulls font 0 out of any .ttc so we can load the regular too.
            tryFace "Cambria" "Cambria" false false "cambria.ttc"
            tryFace "Cambria-Bold" "Cambria" true false "cambriab.ttf"
            tryFace "Cambria-Italic" "Cambria" false true "cambriai.ttf"
            // PDF /Name tokens can't contain literal spaces — use a flat
            // face name internally and keep the spaceful Family for matching
            // the cell's `<name val="Times New Roman"/>`.
            tryFace "TimesNewRoman" "Times New Roman" false false "times.ttf"
            tryFace "TimesNewRoman-Bold" "Times New Roman" true false "timesbd.ttf"
            tryFace "TimesNewRoman-Italic" "Times New Roman" false true "timesi.ttf"
            tryFace "TimesNewRoman-BoldItalic" "Times New Roman" true true "timesbi.ttf"
            tryFace "Verdana" "Verdana" false false "verdana.ttf"
            tryFace "Verdana-Bold" "Verdana" true false "verdanab.ttf"
            tryFace "Verdana-Italic" "Verdana" false true "verdanai.ttf"
            tryFace "Verdana-BoldItalic" "Verdana" true true "verdanaz.ttf"
        ]
        |> List.choose id

    // CJK fallback chain — matches the GDI font-link order Windows uses when
    // Calibri/Arial don't cover a codepoint. PDF /Name tokens can't contain
    // spaces, so use single-word internal names; Family doesn't matter here
    // since these faces are picked by codepoint coverage, not family lookup.
    let cjkFallback =
        [
            tryFace "YuGothic-Regular" "_cjk" false false "YuGothR.ttc"
            tryFace "YuGothic-Bold" "_cjk" true false "YuGothB.ttc"
            tryFace "MSGothic" "_cjk" false false "msgothic.ttc"
            tryFace "MSYaHei" "_cjk" false false "msyh.ttc"
            tryFace "MSYaHei-Bold" "_cjk" true false "msyhbd.ttc"
            tryFace "MSJhengHei" "_cjk" false false "msjh.ttc"
            tryFace "MSJhengHei-Bold" "_cjk" true false "msjhbd.ttc"
            tryFace "MalgunGothic" "_cjk" false false "malgun.ttf"
            tryFace "MalgunGothic-Bold" "_cjk" true false "malgunbd.ttf"
        ]
        |> List.choose id

    {
        DefaultFontFamily = "Calibri"
        Fonts = faces
        CjkFallback = cjkFallback
        SheetName = sheetName
        FileName = fileName
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
    // Page size + margins now live on model.PageSetup (defaulted to US
    // Letter portrait when the source file didn't specify a pageSetup).
    // Excel's `&F` substitutes the workbook's display name (no extension)
    // on Windows; match that so our footer reads "synthetic-print" rather
    // than "synthetic-print.xlsx".
    let opts = pdfOptions sheet.Name (Path.GetFileNameWithoutExtension inputPath)
    let pdfBytes = RedBat.Spreadsheet.Pdf.Render.toBytes opts model
    Directory.CreateDirectory(Path.GetDirectoryName outputPath) |> ignore
    File.WriteAllBytes(outputPath, pdfBytes)
    printfn "spreadsheet pdf runner ok: %s [sheet=%s cells=%d] -> %s (%d bytes)"
        inputPath sheet.Name (model.Cells.Count) outputPath pdfBytes.Length

[<EntryPoint>]
let main argv =
    try
        match argv with
        | [| "--vectortest"; output |] ->
            let bytes = RedBat.Spreadsheet.Pdf.Render.sceneToBytes 400.0 280.0 RedBat.Spreadsheet.Paint.vectorTestScene
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
