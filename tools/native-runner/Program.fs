module NativeRunner

open System
open System.IO
open Rbt.Renderer
module PngEncode = Rbt.OfficeNative.Png

let private usage () =
    eprintfn "Usage: dotnet run --project tools/native-runner -- <input.docx> <output.png> [width] [height]"
    eprintfn "Defaults: width=900 height=1600"

let private fullPath (path: string) =
    if Path.IsPathRooted path then path
    else Path.GetFullPath(path, Directory.GetCurrentDirectory())

let private renderDocx (inputPath: string) (outputPath: string) (width: int) (height: int) =
    let bytes = File.ReadAllBytes inputPath
    let doc =
        match DocxReader.loadBytesNative bytes with
        | Some parsed -> parsed
        | None -> failwithf "Could not parse DOCX: %s" inputPath

    let wordDoc = { Rbt.Word.Import.importDoc doc with Selection = None }
    let settings = Rbt.WordNative.Render.layoutSettings wordDoc
    let measure = Rbt.WordNative.Render.measureTextFor wordDoc
    let _, paint, _ =
        Rbt.Word.Paint.collectDocumentCommands settings "Calibri" measure (Rbt.Word.Layout.rangeMeasurerFromTextMeasurer measure) wordDoc
    let resolve = Rbt.WordNative.Render.createImageResolver ()
    let commands =
        paint
        |> List.collect (Rbt.WordNative.Render.toRendererCommandFor wordDoc resolve)
        |> Rbt.WordNative.Render.viewTransform (float width) wordDoc
    let framebuffer = Scene.render (Raster.create()) width height (Clear Color.white :: commands)
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
