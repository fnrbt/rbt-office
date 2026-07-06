module NativeRunner

open System
open System.IO
open Renderer

let importDoc = RedBat.Word.Import.importDoc

let private usage () =
    eprintfn "Usage: dotnet run --project tools/native-runner -- <input.docx> <output.png> [width] [height]"
    eprintfn "Defaults: width=900 height=1600"

let private fullPath (path: string) =
    if Path.IsPathRooted path then path
    else Path.GetFullPath(path, Directory.GetCurrentDirectory())

let private existingFile path =
    if File.Exists path then Some path else None

let private findWordFont () =
    [
        Environment.GetEnvironmentVariable("RB_WORD_FONT_DIR")
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "artifacts/fonts/windows"))
        Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "../../artifacts/fonts/windows"))
    ]
    |> List.choose (fun dir ->
        if String.IsNullOrWhiteSpace dir then None
        else Path.Combine(dir, "calibri.ttf") |> existingFile)
    |> List.tryHead

let private initWordFont () =
    findWordFont ()
    |> Option.iter Composit.Ui.Layout.initFont

let private renderDocx (inputPath: string) (outputPath: string) (width: int) (height: int) =
    initWordFont ()
    let bytes = File.ReadAllBytes inputPath
    let doc =
        match DocxReader.loadBytesNative bytes with
        | Some parsed -> parsed
        | None -> failwithf "Could not parse DOCX: %s" inputPath

    let wordDoc = { RedBat.Word.Import.importDoc doc with Selection = None }
    let widget = Composit.Ui.WordEditor.scrollEditor (float width) (float height) wordDoc
    let backend = Raster.create()
    let framebuffer = Composit.Ui.Render.renderWithBg Color.white backend width height widget
    let png = PngEncode.encode framebuffer

    Directory.CreateDirectory(Path.GetDirectoryName outputPath) |> ignore
    File.WriteAllBytes(outputPath, png)
    printfn "native runner ok: %s -> %s (%dx%d)" inputPath outputPath width height

[<EntryPoint>]
let main argv =
    try
        match argv with
        | [| input; output |] ->
            renderDocx (fullPath input) (fullPath output) 900 1600
            0
        | [| input; output; width |] ->
            renderDocx (fullPath input) (fullPath output) (Int32.Parse width) 1600
            0
        | [| input; output; width; height |] ->
            renderDocx (fullPath input) (fullPath output) (Int32.Parse width) (Int32.Parse height)
            0
        | _ ->
            usage()
            2
    with ex ->
        eprintfn "%O" ex
        1
