module PdfRunner

open System
open System.IO
open RedBat.Word

let private usage () =
    eprintfn "Usage: dotnet run --project tools/pdf-runner -- <input.docx> <output.pdf>"

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
    ]
    |> List.choose (fun path -> if String.IsNullOrWhiteSpace path then None else existingDir path)

let private findFallbackFont () =
    [
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../renderer/NotoSans-Regular.ttf"))
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../renderer/NotoSans-Regular.ttf"))
        "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf"
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    ]
    |> List.tryFind File.Exists

let private tryFontFile fileName =
    fontDirs ()
    |> List.tryPick (fun dir ->
        let path = Path.Combine(dir, fileName)
        if File.Exists path then Some path else None)

let private tryFace faceName family bold italic fileName =
    tryFontFile fileName
    |> Option.map (fun path ->
        {
            FaceName = faceName
            Family = family
            Bold = bold
            Italic = italic
            Data = File.ReadAllBytes path
        }: RedBat.Word.Pdf.FontFace)

let private pdfOptions () : RedBat.Word.Pdf.RenderOptions =
    let windowsFaces =
        [
            tryFace "Calibri" "Calibri" false false "calibri.ttf"
            tryFace "Calibri-Bold" "Calibri" true false "calibrib.ttf"
            tryFace "Calibri-Italic" "Calibri" false true "calibrii.ttf"
            tryFace "Calibri-BoldItalic" "Calibri" true true "calibriz.ttf"
            tryFace "Calibri Light" "Calibri Light" false false "calibril.ttf"
            tryFace "Calibri Light-Italic" "Calibri Light" false true "calibrili.ttf"
            tryFace "Cambria-Bold" "Cambria" true false "cambriab.ttf"
            tryFace "Cambria-Italic" "Cambria" false true "cambriai.ttf"
            tryFace "Cambria-BoldItalic" "Cambria" true true "cambriaz.ttf"
            tryFace "Arial" "Arial" false false "arial.ttf"
            tryFace "Arial-Bold" "Arial" true false "arialbd.ttf"
            tryFace "Arial-Italic" "Arial" false true "ariali.ttf"
            tryFace "Arial-BoldItalic" "Arial" true true "arialbi.ttf"
            tryFace "Times New Roman" "Times New Roman" false false "times.ttf"
            tryFace "Times New Roman-Bold" "Times New Roman" true false "timesbd.ttf"
            tryFace "Times New Roman-Italic" "Times New Roman" false true "timesi.ttf"
            tryFace "Times New Roman-BoldItalic" "Times New Roman" true true "timesbi.ttf"
        ]
        |> List.choose id

    let fallbackPath =
        findFallbackFont ()
        |> Option.defaultWith (fun () -> failwith "Could not find a TTF font for PDF rendering")
    let fallbackBytes = File.ReadAllBytes fallbackPath
    let face faceName family =
        {
            FaceName = faceName
            Family = family
            Bold = false
            Italic = false
            Data = fallbackBytes
        }: RedBat.Word.Pdf.FontFace

    let defaultFamily =
        if windowsFaces |> List.exists (fun face -> face.Family = "Calibri" && not face.Bold && not face.Italic) then "Calibri"
        else "Noto Sans"

    {
        DefaultFontFamily = defaultFamily
        Fonts = windowsFaces @ [
            face "NotoSans" "Noto Sans"
        ]
        MeasureText = None
        ShapeText = None
        Images = []
    }

let private decodeImageResources (wordDoc: DocxReader.WordDocument) : RedBat.Word.Pdf.ImageResource list =
    wordDoc.Images
    |> Map.toList
    |> List.choose (fun (_, dataUrl) ->
        match FsPdf.Image.fromDataUrl dataUrl with
        | Ok info ->
            Some ({
                Source = dataUrl
                Width = info.Width
                Height = info.Height
                ColorSpace = info.ColorSpace
                BitsPerComponent = info.BitsPerComponent
                Data = info.Data
                Filter = info.Filter
            }: RedBat.Word.Pdf.ImageResource)
        | Error _ -> None)

let private writeProjectPdf (inputPath: string) (outputPath: string) =
    let raw = File.ReadAllBytes inputPath
    let wordDoc =
        match DocxReader.loadBytesNative raw with
        | Some parsed -> parsed
        | None -> failwithf "Could not parse DOCX: %s" inputPath
    let doc = NativeRunner.importDoc wordDoc
    let options = { pdfOptions() with Images = decodeImageResources wordDoc }
    let bytes = RedBat.Word.Pdf.toBytes options doc
    Directory.CreateDirectory(Path.GetDirectoryName outputPath) |> ignore
    File.WriteAllBytes(outputPath, bytes)
    printfn "pdf runner ok: %s -> %s (%d bytes)" inputPath outputPath bytes.Length

[<EntryPoint>]
let main argv =
    try
        match argv with
        | [| input; output |] ->
            writeProjectPdf (fullPath input) (fullPath output)
            0
        | _ ->
            usage()
            2
    with ex ->
        eprintfn "%O" ex
        1
