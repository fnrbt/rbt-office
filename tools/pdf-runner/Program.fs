module PdfRunner

open System
open System.IO
open RedBat.Word

let private usage () =
    eprintfn "Usage: dotnet run --project tools/pdf-runner -- <input.docx> <output.pdf>"

let private fullPath (path: string) =
    if Path.IsPathRooted path then path
    else Path.GetFullPath(path, Directory.GetCurrentDirectory())

let private findFont () =
    let candidates = [
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../renderer/NotoSans-Regular.ttf"))
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../renderer/NotoSans-Regular.ttf"))
        "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf"
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    ]
    candidates |> List.tryFind File.Exists

let private pdfOptions () : RedBat.Word.Pdf.RenderOptions =
    let fontPath =
        findFont ()
        |> Option.defaultWith (fun () -> failwith "Could not find a TTF font for PDF rendering")
    let bytes = File.ReadAllBytes fontPath
    let face faceName family =
        {
            FaceName = faceName
            Family = family
            Bold = false
            Italic = false
            Data = bytes
        }: RedBat.Word.Pdf.FontFace
    {
        DefaultFontFamily = "Noto Sans"
        Fonts = [
            face "NotoSans" "Noto Sans"
            face "Calibri" "Calibri"
            face "Arial" "Arial"
            face "TimesNewRoman" "Times New Roman"
        ]
        MeasureText = None
        ShapeText = None
        Images = []
    }

let private writeProjectPdf (inputPath: string) (outputPath: string) =
    let doc =
        match DocxReader.loadBytesNative (File.ReadAllBytes inputPath) with
        | Some parsed -> NativeRunner.importDoc parsed
        | None -> failwithf "Could not parse DOCX: %s" inputPath

    let bytes = RedBat.Word.Pdf.toBytes (pdfOptions()) doc
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
