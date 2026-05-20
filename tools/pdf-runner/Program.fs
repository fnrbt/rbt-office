module PdfRunner

open System
open System.IO
open RedBat.Word

module Doc = RedBat.Word.Document
module WordLayout = RedBat.Word.Layout
module WordModelLayout = RedBat.Word.WordLayout
module WordPdf = RedBat.Word.Pdf

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
            // Regular Cambria ships only inside cambria.ttc (font index 0);
            // parseTrueType in both the shaper and PdfFont auto-detects
            // ttcf and extracts font 0, so we can register the collection
            // directly under the bare "Cambria" face name.
            tryFace "Cambria" "Cambria" false false "cambria.ttc"
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

let private pageNumberForY (doc: Doc.Document) y =
    let stride = doc.PageSetup.PageHeight + WordModelLayout.Page.gap
    int (Math.Floor(max 0.0 y / stride)) + 1

let private displayPageNumberForPhysicalPage (restartPoints: (int * int) list) physicalPage =
    restartPoints
    |> List.filter (fun (restartPhysicalPage, _) -> restartPhysicalPage <= physicalPage)
    |> List.tryLast
    |> Option.map (fun (restartPhysicalPage, restartDisplayPage) -> restartDisplayPage + physicalPage - restartPhysicalPage)
    |> Option.defaultValue physicalPage

let private pdfLayoutSettings (doc: Doc.Document) : WordLayout.LayoutSettings =
    let lineHeight = 21.0
    let footerReserve =
        if not doc.Footers.IsEmpty || not doc.FooterParts.IsEmpty then doc.PageSetup.FooterDistance + lineHeight * 1.5
        else doc.PageSetup.MarginBottom
    {
        Width = doc.PageSetup.PageWidth
        PageHeight = doc.PageSetup.PageHeight
        FontSize = 14.0
        LineHeight = lineHeight
        PaddingLeft = doc.PageSetup.MarginLeft
        PaddingTop = WordModelLayout.Page.bodyTopWithRunningHeader doc
        MarginRight = doc.PageSetup.MarginRight
        MarginBottom = max doc.PageSetup.MarginBottom footerReserve
        InTableCell = false
        TableContinuationHeaderHeight = 0.0
        LongTrailingSpaceFitSafety = 0.25
    }

let private pageMapFromLayouts (doc: Doc.Document) (layouts: WordLayout.BlockLayout list) : Doc.PageMap =
    let restartPoints =
        layouts
        |> List.choose (fun layout ->
            Doc.findParagraph doc layout.BlockId
            |> Option.bind (fun para ->
                para.Metadata.SectionPageNumberStart
                |> Option.map (fun start -> pageNumberForY doc layout.Y, max 1 start)))
        |> List.sortBy fst

    let displayPageNumberForY y =
        y
        |> pageNumberForY doc
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
                    | WordLayout.ImageSegment _ -> None)))

    let bookmarkAnchors =
        layouts
        |> List.collect (fun layout ->
            layout.Lines
            |> List.collect (fun line ->
                line.Segments
                |> List.choose (function
                    | WordLayout.TextSegment (_, style, _, _, startOffset) ->
                        style.BookmarkName
                        |> Option.map (fun name ->
                            ({
                                PageBookmarkName = name
                                PageBookmarkBlockId = layout.BlockId
                                PageBookmarkStartOffset = startOffset
                                PageNumber = displayPageNumberForY line.Y
                            }: Doc.PageBookmarkAnchor))
                    | WordLayout.ImageSegment _ -> None)))
        |> List.distinctBy (fun anchor -> anchor.PageBookmarkName)

    let blockAnchors =
        layouts
        |> List.map (fun layout ->
            ({
                PageBlockId = layout.BlockId
                PageNumber = displayPageNumberForY layout.Y
            }: Doc.PageBlockAnchor))

    let pageCount =
        layouts
        |> List.map (fun layout -> pageNumberForY doc (layout.Y + layout.Height))
        |> List.fold max 1

    {
        PageCount = pageCount
        BlockAnchors = blockAnchors
        Fields = fieldAnchors
        Bookmarks = bookmarkAnchors
    }

let private refreshFieldsForPdf (options: WordPdf.RenderOptions) (doc: Doc.Document) =
    let measureText = WordPdf.createMeasurer options
    let measureRange = WordLayout.rangeMeasurerFromTextMeasurer measureText
    let layouts = WordLayout.layoutDocumentWithRange (pdfLayoutSettings doc) measureText measureRange doc
    Doc.refreshFieldsWithPageMap (pageMapFromLayouts doc layouts) doc

let private writeProjectPdf (inputPath: string) (outputPath: string) =
    let raw = File.ReadAllBytes inputPath
    let wordDoc =
        match DocxReader.loadBytesNative raw with
        | Some parsed -> parsed
        | None -> failwithf "Could not parse DOCX: %s" inputPath
    let options = { pdfOptions() with Images = decodeImageResources wordDoc }
    let doc = NativeRunner.importDoc wordDoc |> refreshFieldsForPdf options
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
