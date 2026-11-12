// The shared native Word painter draws a document where the browser lays it out.
// The fixture passes through import, layout, paint translation and the headless
// software rasterizer without a UI toolkit.
//
// The page, line and shape expectations are shared with the browser fixture;
// pixels also check the text, shape and table borders.
//
// tests/word/native-painter/expected.json: the page, where each line of text
// is, the rectangle of the shape. tools/word-browser's native-fixture check
// asserts the same numbers of the browser's layout of the same document, so
// the two painters are compared through them:
//
//   * the draw calls: one run of glyphs per line of text, starting at the
//     line's x with its baseline inside the line's box; the shape's path with
//     the shape's rectangle as its bounds, filled red;
//   * the pixels: ink in every line's box and none in the space between the
//     paragraphs, the large line's ink as high as its size makes it, the red
//     pixels of the shape within two pixels of its rectangle, the table's
//     borders around its cells' text.
//
// (Line boxes do not depend on the widths of the glyphs - the paragraphs are
// one line each - so the native font, Noto Sans where the document says
// Calibri, gives the positions the browser has.)
//
//   dotnet build office-wisp/word-render/Rbt.WordNative.fsproj
//   dotnet fsi tests/word/native-painter.fsx
#r "nuget: Fable.Core, 4.3.0"
#r "../../office-wisp/word-render/bin/Debug/net10.0/Rbt.Renderer.dll"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../office-wisp/word-render/bin/Debug/net10.0/Rbt.Word.dll"
#r "../../office-wisp/word-render/bin/Debug/net10.0/Rbt.OfficeNative.dll"
#r "../../office-wisp/word-render/bin/Debug/net10.0/Rbt.WordNative.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open System.IO.Compression
open System.Text
open System.Text.Json
open Rbt.Renderer

module Painter = Rbt.WordNative.Render

let mutable pass = 0
let mutable fail = 0
let check (label: string) (ok: bool) (detail: string) =
    if ok then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %s: %s" label detail

let fixtureDir = Path.Combine(__SOURCE_DIRECTORY__, "native-painter")

// ---- the fixture as a package ---------------------------------------------
let docxBytes =
    use stream = new MemoryStream()
    (use zip = new ZipArchive(stream, ZipArchiveMode.Create, true)
     let add (name: string) (text: string) =
         let entry = zip.CreateEntry name
         use writer = new StreamWriter(entry.Open(), UTF8Encoding(false))
         writer.Write text
     add "[Content_Types].xml" "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>"
     add "_rels/.rels" "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>"
     add "word/document.xml" (File.ReadAllText(Path.Combine(fixtureDir, "document.xml"))))
    stream.ToArray()

// ---- what the browser's layout of it is ------------------------------------
type Line = { Text: string; X: float; Y: float; Height: float }
let expected = JsonDocument.Parse(File.ReadAllText(Path.Combine(fixtureDir, "expected.json"))).RootElement
let number (element: JsonElement) (name: string) = element.GetProperty(name).GetDouble()
let pageWidth = number (expected.GetProperty "page") "width"
let pageHeight = number (expected.GetProperty "page") "height"
let lines =
    [ for line in (expected.GetProperty "lines").EnumerateArray() ->
        { Text = line.GetProperty("text").GetString(); X = number line "x"; Y = number line "y"; Height = number line "height" } ]
let shape = expected.GetProperty "shape"
let shapeX, shapeY, shapeW, shapeH = number shape "x", number shape "y", number shape "width", number shape "height"

// ---- through the native path -----------------------------------------------
let fontPath = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "../../office-wisp/word-render/bin/Debug/net10.0/NotoSans-Regular.ttf"))
let font = Rbt.Renderer.Text.loadFontFile fontPath

let doc =
    match DocxReader.loadBytesNative docxBytes with
    | Some parsed -> { Rbt.Word.Import.importDoc parsed with Selection = None }
    | None -> failwith "the fixture could not be read"

check "the page is the fixture's" (doc.PageSetup.PageWidth = pageWidth && doc.PageSetup.PageHeight = pageHeight) (sprintf "%g x %g, the browser has %g x %g" doc.PageSetup.PageWidth doc.PageSetup.PageHeight pageWidth pageHeight)

let width = int pageWidth
let height = int pageHeight

// The draw calls.
let settings = Painter.layoutSettings doc
let measure = Painter.measureTextWithFont font
let _, paints, _ = Rbt.Word.Paint.collectDocumentCommands settings "Calibri" measure (Rbt.Word.Layout.rangeMeasurerFromTextMeasurer measure) doc
let commands = paints |> List.collect (Painter.toRendererCommandWith (fun _ -> font) (fun _ -> None))
// Embedded media and its source crop must reach the renderer rather than a
// placeholder; the same paint path is used by both hosts and the runner.
let imageSource = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAABCAYAAAD5PA/NAAAAEklEQVR4nGP4z8DwH4SBqAGEATzXBvv75BwiAAAAAElFTkSuQmCC"
let imagePaint = Rbt.Word.Paint.DrawImage(imageSource, 10.0, 20.0, 30.0, 40.0, Some "50000", None, None, None)
let imageCommands = Painter.toRendererCommand (Painter.createImageResolver ()) imagePaint
let imageFrame = Scene.render (Raster.create ()) 60 80 (Clear Color.white :: imageCommands)
let imagePixel x y =
    let memory = imageFrame.Px
    memory.Span.[y * imageFrame.W + x]
let insideImage = imagePixel 18 40
let outsideImage = imagePixel 5 40
check "cropping removes the red half and composites translucent blue over white"
    (Pixel.r insideImage >= 126uy && Pixel.r insideImage <= 128uy
     && Pixel.g insideImage = Pixel.r insideImage && Pixel.b insideImage = 255uy
     && Pixel.r outsideImage = 255uy && Pixel.g outsideImage = 255uy && Pixel.b outsideImage = 255uy)
    (sprintf "inside %08x, outside %08x" insideImage outsideImage)

// The pixels.
let framebuffer = Scene.render (Raster.create ()) width height (Clear Color.white :: commands)
let pixels : uint32[] = framebuffer.Px.ToArray()
let pixelAt (x: int) (y: int) = pixels.[y * framebuffer.W + x]
let isRed (p: uint32) = Pixel.r p > 200uy && Pixel.g p < 80uy && Pixel.b p < 80uy
let isInk (p: uint32) = int (Pixel.r p) + int (Pixel.g p) + int (Pixel.b p) < 384 && not (isRed p)
/// The rows of [top, bottom) with ink in the columns [left, right).
let inkRows (left: int) (right: int) (top: int) (bottom: int) =
    [ for y in max 0 top .. min (height - 1) (bottom - 1) do
        if seq { max 0 left .. min (width - 1) (right - 1) } |> Seq.exists (fun x -> isInk (pixelAt x y)) then yield y ]

// The red of the shape.
let redBounds =
    let mutable minX, minY, maxX, maxY = Int32.MaxValue, Int32.MaxValue, -1, -1
    for y in 0 .. height - 1 do
        for x in 0 .. width - 1 do
            if isRed (pixelAt x y) then
                minX <- min minX x; minY <- min minY y; maxX <- max maxX x; maxY <- max maxY y
    if maxX < 0 then None else Some (float minX, float minY, float (maxX - minX + 1), float (maxY - minY + 1))
check "the red pixels are the shape's rectangle, within two pixels"
    (match redBounds with Some (x, y, w, h) -> abs (x - shapeX) <= 2.0 && abs (y - shapeY) <= 2.0 && abs (w - shapeW) <= 2.0 && abs (h - shapeH) <= 2.0 | None -> false)
    (sprintf "the shape is at %g, %g, %g x %g; the red pixels: %A" shapeX shapeY shapeW shapeH redBounds)

// The ink of each line, in a strip to the right of where the line starts (inside its cell for a line of the table).
let strip (line: Line) = int line.X + 1, int line.X + 60
let inkOf (line: Line) = let left, right = strip line in inkRows left right (int line.Y - 6) (int (line.Y + line.Height) + 6)
for line in lines do
    let rows = inkOf line
    check (sprintf "\"%s\" is painted in its line's box" line.Text)
        (not rows.IsEmpty && float (List.min rows) >= line.Y - 2.0 && float (List.max rows) <= line.Y + line.Height + 2.0)
        (sprintf "the line's box is y %g to %g; ink in rows %s" line.Y (line.Y + line.Height) (if rows.IsEmpty then "none" else sprintf "%d to %d" (List.min rows) (List.max rows)))
// Nothing is painted in the space between two paragraphs: of the three lines above the shape's line, whose
// paragraphs are a third of an inch apart. (A descender may reach a few pixels below its line's box - the
// native font is not the document's - so the space is taken from eight pixels below the box.)
let spaced = lines |> List.takeWhile (fun line -> line.Text <> "Line of the shape")
check "the fixture has three lines above the shape's line" (spaced.Length = 3) (sprintf "%A" (spaced |> List.map (fun line -> line.Text)))
for (above, below) in List.pairwise spaced do
    let top, bottom = int (ceil (above.Y + above.Height)) + 8, int below.Y - 2
    let rows = inkRows 96 400 top bottom
    check (sprintf "no ink between \"%s\" and \"%s\"" above.Text below.Text) (bottom - top >= 10 && rows.IsEmpty) (sprintf "rows %d to %d should be blank; ink in %A" top bottom rows)
// The large line is painted large: its ink is higher than the first line's by about the ratio of the sizes (28 pt and 11 pt).
match lines |> List.tryFind (fun line -> line.Text = "Large line"), lines |> List.tryFind (fun line -> line.Text = "First line") with
| Some large, Some first ->
    let inkHeight (line: Line) = match inkOf line with [] -> 0 | rows -> List.max rows - List.min rows + 1
    check "the large line's ink is as high as its size makes it" (float (inkHeight large) > 2.0 * float (inkHeight first) && float (inkHeight large) < 3.2 * float (inkHeight first))
        (sprintf "ink of the large line %d rows, of the first %d" (inkHeight large) (inkHeight first))
| _ -> check "the fixture has its large and its first line" false "missing in expected.json"
// The table: a border above and below its cells' text, and between the two cells.
match lines |> List.tryFind (fun line -> line.Text = "Cell one"), lines |> List.tryFind (fun line -> line.Text = "Cell two") with
| Some one, Some two ->
    let column = int one.X + 100
    let borderRows = [ for y in int one.Y - 14 .. int (one.Y + one.Height) + 14 do if isInk (pixelAt column y) then yield y ]
    check "the table's borders are above and below the text of its cells"
        (borderRows |> List.exists (fun y -> float y < one.Y + 2.0) && borderRows |> List.exists (fun y -> float y > one.Y + one.Height - 2.0))
        (sprintf "the cell's line is y %g to %g; ink in column %d at rows %A" one.Y (one.Y + one.Height) column borderRows)
    let row = int (one.Y + one.Height / 2.0)
    let between = [ for x in int one.X + 90 .. int two.X - 1 do if isInk (pixelAt x row) then yield x ]
    check "and between its two cells" (not between.IsEmpty && float (List.max between) < two.X) (sprintf "ink in row %d between x %g and %g: %A" row (one.X + 90.0) two.X between)
| _ -> check "the fixture has its table" false "missing in expected.json"

printfn "Word native painter tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
