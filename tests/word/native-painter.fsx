// The native painter (word-native: the Word editor as a widget of
// rbt-widgets, painted by rbt-renderer) draws a document where the layout
// puts it, and where the browser's painter does.
//
// The painter was only ever built. Here a fixture page
// (tests/word/native-painter/document.xml) goes through the whole native
// path - the .NET reader, the import, the layout with the native text
// measurer, Paint's commands, word-native's renderer commands, the software
// rasteriser - into a bitmap, headless, and is held against the numbers of
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
// Needs word-native built (it is not one of the projects run-all builds):
//   dotnet build word-native/Rbt.WordWidgets.fsproj
//   dotnet fsi tests/word/native-painter.fsx
// run-all runs it when word-native's build output is there.
#r "nuget: Fable.Core, 4.3.0"
#r "../../word-native/bin/Debug/net10.0/Rbt.Renderer.dll"
#r "../../word-native/bin/Debug/net10.0/Rbt.Widgets.dll"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../word-native/bin/Debug/net10.0/Rbt.Word.dll"
#r "../../word-native/bin/Debug/net10.0/Rbt.WordWidgets.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open System.IO.Compression
open System.Text
open System.Text.Json
open Rbt.Renderer

module Editor = Rbt.Widgets.Composit.Ui.WordEditor

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
let fontPath = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "../../word-native/bin/Debug/net10.0/NotoSans-Regular.ttf"))
if File.Exists fontPath then Rbt.Widgets.Composit.Ui.Layout.initFont fontPath

let doc =
    match DocxReader.loadBytesNative docxBytes with
    | Some parsed -> { Rbt.Word.Import.importDoc parsed with Selection = None }
    | None -> failwith "the fixture could not be read"

check "the page is the fixture's" (doc.PageSetup.PageWidth = pageWidth && doc.PageSetup.PageHeight = pageHeight) (sprintf "%g x %g, the browser has %g x %g" doc.PageSetup.PageWidth doc.PageSetup.PageHeight pageWidth pageHeight)

let width = int pageWidth
let height = int pageHeight

// The draw calls.
let commands = Editor.commands pageWidth pageHeight (fun _ -> None) None doc
let rec flatten (cmds: Cmd list) : Cmd list =
    cmds
    |> List.collect (fun cmd ->
        match cmd with
        | Cmd.Group (_, _, inner) | Cmd.Clip (_, _, inner) | Cmd.AntiAlias (_, inner) | Cmd.SaveRestore inner | Cmd.Filtered (_, inner) | Cmd.Blend (_, inner) -> flatten inner
        | other -> [ other ])
let flat = flatten commands
let glyphRuns = flat |> List.choose (function Cmd.DrawGlyphs (origin, _, _) -> Some (float origin.X, float origin.Y) | _ -> None)
check "one run of glyphs per line of text" (glyphRuns.Length = lines.Length) (sprintf "%d runs of glyphs, %d lines: %A" glyphRuns.Length lines.Length glyphRuns)
for line in lines do
    let found = glyphRuns |> List.filter (fun (x, baseline) -> abs (x - line.X) <= 1.0 && baseline > line.Y && baseline < line.Y + line.Height + 1.0)
    check (sprintf "the glyphs of \"%s\" start at the line's x, their baseline in the line's box" line.Text) (found.Length = 1)
        (sprintf "the line is at x %g, y %g to %g; glyphs are drawn at %A" line.X line.Y (line.Y + line.Height) glyphRuns)
let redPaths =
    flat
    |> List.choose (function
        | Cmd.FillPath (path, Solid color, _) when color.R > 200uy && color.G < 60uy && color.B < 60uy ->
            let points = path |> List.collect (function MoveTo p | LineTo p -> [ float p.X, float p.Y ] | QuadTo (_, p) -> [ float p.X, float p.Y ] | CubicTo (_, _, p) -> [ float p.X, float p.Y ] | _ -> [])
            if points.IsEmpty then None
            else
                let xs, ys = points |> List.map fst, points |> List.map snd
                Some (List.min xs, List.min ys, List.max xs - List.min xs, List.max ys - List.min ys)
        | _ -> None)
check "the shape is one path filled red, its bounds the shape's rectangle"
    (match redPaths with [ (x, y, w, h) ] -> abs (x - shapeX) <= 1.0 && abs (y - shapeY) <= 1.0 && abs (w - shapeW) <= 1.0 && abs (h - shapeH) <= 1.0 | _ -> false)
    (sprintf "the shape is at %g, %g, %g x %g; red paths: %A" shapeX shapeY shapeW shapeH redPaths)

// The pixels.
let framebuffer = Rbt.Widgets.Composit.Ui.Render.renderWithBg Color.white (Raster.create ()) width height (Editor.editor pageWidth pageHeight doc)
check "the bitmap is the page" (framebuffer.W = width && framebuffer.H = height) (sprintf "%d x %d" framebuffer.W framebuffer.H)
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
