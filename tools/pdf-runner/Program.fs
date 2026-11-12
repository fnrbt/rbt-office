module PdfRunner

open System
open System.IO
open Rbt.Word

module Doc = Rbt.Word.Document
module WordImport = Rbt.Word.Import
module WordLayout = Rbt.Word.Layout
module WordPdf = Rbt.Word.Pdf

let private usage () =
    eprintfn "Usage: dotnet run --project tools/pdf-runner -- <input.docx> <output.pdf>"

let private fullPath (path: string) =
    if Path.IsPathRooted path then path
    else Path.GetFullPath(path, Directory.GetCurrentDirectory())

let private pdfOptions () : Rbt.Word.Pdf.RenderOptions =
    let faces : Rbt.Word.Pdf.FontFace list =
        [ for family in [ "Calibri"; "Calibri Light"; "Cambria"; "Arial"; "Times New Roman" ] do
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
    {
        DefaultFontFamily = "Calibri"
        Fonts = faces
        MeasureText = None
        ShapeText = None
        Images = []
    }

let private decodeImageResources (wordDoc: DocxReader.WordDocument) : Rbt.Word.Pdf.ImageResource list =
    wordDoc.Images
    |> Map.toList
    |> List.choose (fun (_, dataUrl) ->
        match Rbt.Pdf.Image.fromDataUrl dataUrl with
        | Ok info ->
            Some ({
                Source = dataUrl
                Width = info.Width
                Height = info.Height
                ColorSpace = info.ColorSpace
                BitsPerComponent = info.BitsPerComponent
                Data = info.Data
                Filter = info.Filter
            }: Rbt.Word.Pdf.ImageResource)
        | Error _ -> None)

// 1-based index of the page whose band contains y, using the actual per-page
// geometry (so page numbers are correct when sections have different page sizes).
let private pageNumberForY (pages: WordLayout.PageInfo list) y =
    let mutable idx = 0
    pages |> List.iteri (fun i info -> if y >= info.Top - 0.01 then idx <- i)
    idx + 1

let private displayPageNumberForPhysicalPage (restartPoints: (int * int) list) physicalPage =
    restartPoints
    |> List.filter (fun (restartPhysicalPage, _) -> restartPhysicalPage <= physicalPage)
    |> List.tryLast
    |> Option.map (fun (restartPhysicalPage, restartDisplayPage) -> restartDisplayPage + physicalPage - restartPhysicalPage)
    |> Option.defaultValue physicalPage


let private pageMapFromLayouts (doc: Doc.Document) (pages: WordLayout.PageInfo list) (layouts: WordLayout.BlockLayout list) : Doc.PageMap =
    let restartPoints =
        layouts
        |> List.choose (fun layout ->
            Doc.findParagraph doc layout.BlockId
            |> Option.bind (fun para ->
                para.Metadata.SectionPageNumberStart
                |> Option.map (fun start -> pageNumberForY pages layout.Y, max 1 start)))
        |> List.sortBy fst

    let displayPageNumberForY y =
        y
        |> pageNumberForY pages
        |> displayPageNumberForPhysicalPage restartPoints

    let fieldAnchors =
        layouts
        |> List.collect (fun layout ->
            layout.Lines
            |> List.collect (fun line ->
                line.Segments
                |> List.choose (function
                    | WordLayout.TextSegment (_, style, _, _, startOffset) ->
                        style.FieldCode
                        |> Option.map (fun code ->
                            ({
                                PageFieldBlockId = layout.BlockId
                                PageFieldStartOffset = startOffset
                                PageFieldCode = code
                                PageNumber = displayPageNumberForY line.Y
                            }: Doc.PageFieldAnchor))
                    | WordLayout.ImageSegment _ -> None
                    | WordLayout.MathSegment _ -> None)))

    let bookmarkAnchors =
        layouts
        |> List.collect (fun layout ->
            layout.Lines
            |> List.collect (fun line ->
                line.Segments
                |> List.choose (function
                    | WordLayout.TextSegment (_, style, _, _, startOffset) ->
                        Rbt.Word.TextBuffer.bookmarkName style
                        |> Option.map (fun name ->
                            ({
                                PageBookmarkName = name
                                PageBookmarkBlockId = layout.BlockId
                                PageBookmarkStartOffset = startOffset
                                PageNumber = displayPageNumberForY line.Y
                            }: Doc.PageBookmarkAnchor))
                    | WordLayout.ImageSegment _ -> None
                    | WordLayout.MathSegment _ -> None)))
        |> List.distinctBy (fun anchor -> anchor.PageBookmarkName)

    let blockAnchors =
        layouts
        |> List.map (fun layout ->
            ({
                PageBlockId = layout.BlockId
                PageNumber = displayPageNumberForY layout.Y
            }: Doc.PageBlockAnchor))

    let pageCount = max 1 pages.Length

    {
        PageCount = pageCount
        BlockAnchors = blockAnchors
        Fields = fieldAnchors
        Bookmarks = bookmarkAnchors
    }

let private refreshFieldsForPdf (options: WordPdf.RenderOptions) (doc: Doc.Document) =
    let measureText = WordPdf.createMeasurer options
    let measureRange = WordLayout.rangeMeasurerFromTextMeasurer measureText
    let layouts, pages = WordLayout.layoutDocumentPages (WordLayout.documentSettings 14.0 21.0 0.25 doc) measureText measureRange doc
    Doc.refreshFieldsWithPageMap (pageMapFromLayouts doc pages layouts) doc

let private writeProjectPdf (inputPath: string) (outputPath: string) =
    let raw = File.ReadAllBytes inputPath
    let wordDoc =
        match DocxReader.loadBytesNative raw with
        | Some parsed -> parsed
        | None -> failwithf "Could not parse DOCX: %s" inputPath
    let options = { pdfOptions() with Images = decodeImageResources wordDoc }
    let doc = Rbt.Word.Import.importDoc wordDoc |> refreshFieldsForPdf options
    let bytes = Rbt.Word.Pdf.toBytes options doc
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
