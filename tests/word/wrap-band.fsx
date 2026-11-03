// Text beside a floating object: not in a band narrower than an inch, on the
// side(s) the object says (wrapText), clear of it by its distances from text.
// float-image-wrap-square.docx: the picture is anchored at 1463040 EMU (1.6 in
// = 153.6 px) from the page's left edge and 1371600 EMU (1.5 in = 144 px)
// from its top, 1463040 x 1097280 EMU (153.6 x 115.2 px), wrapSquare
// wrapText="bothSides", distL = distR = 114300 EMU (12 px), distT = distB = 0.
// The margins are 1 in (96 px), the text 624 px wide: left of the picture there
// are 153.6 - 96 - 12 = 45.6 px - less than an inch, so no text goes there -
// and right of it 624 - (57.6 + 153.6 + 12) = 400.8 px, where the lines that
// pass the picture (144 .. 259.2 px down the page) stand, from 223.2 px on.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open Rbt.Word.TextBuffer
module Doc = Rbt.Word.Document
module Layout = Rbt.Word.Layout

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-50s expected %A got %A" label expected got

let root = IO.Path.GetFullPath(IO.Path.Combine(__SOURCE_DIRECTORY__, "../.."))
let importFile (path: string) =
    match DocxReader.loadBytesNative (IO.File.ReadAllBytes(IO.Path.Combine(root, path))) with
    | Some parsed -> Rbt.Word.Import.importDoc parsed
    | None -> failwithf "Could not parse %s" path
let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let settingsOf (doc: Doc.Document) : Layout.LayoutSettings =
    let s = doc.PageSetup
    { Width = s.PageWidth; PageHeight = s.PageHeight; FontSize = 16.0; LineHeight = 24.0; PaddingLeft = s.MarginLeft; PaddingTop = s.MarginTop
      MarginRight = s.MarginRight; MarginBottom = s.MarginBottom; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5; DefaultTabStop = s.DefaultTabStop }

/// The text of each line with where it starts and ends across the text area and where it is down the page.
let textLines (doc: Doc.Document) =
    let layouts, _ = Layout.layoutDocumentPages (settingsOf doc) measureText measureRange doc
    [ for layout in layouts do
        for line in layout.Lines do
            // (Without the space a line ends with: that may stand beyond the line's end.)
            let spans = line.Segments |> List.choose (function Layout.TextSegment (text, style, x, _, _) when text.Trim() <> "" -> Some (x, x + measureText (text.TrimEnd()) style) | _ -> None)
            if not spans.IsEmpty then
                yield (spans |> List.map fst |> List.min) + Layout.alignmentOffset layout line, (spans |> List.map snd |> List.max) + Layout.alignmentOffset layout line, line.Y, line.Height ]
let beside top bottom lines = lines |> List.filter (fun (_, _, y, height) -> y + height > top + 0.5 && y < bottom - 0.5)

let doc = importFile "tests/repros/word/float-image-wrap-square.docx"
let lines = textLines doc
let passing = lines |> beside 144.0 259.2
let clear = lines |> List.filter (fun (_, _, y, _) -> y >= 259.2)
check "lines pass the picture" (passing.Length >= 3) true
check "none of them has text left of the picture (the band there is 45.6 px)" (passing |> List.filter (fun (left, _, _, _) -> left < 223.2 - 0.5) |> List.map (fun (left, _, y, _) -> Math.Round left, Math.Round y)) []
check "they start right of the picture, 12 px clear of it" (passing |> List.map (fun (left, _, _, _) -> Math.Round(left, 1)) |> List.distinct) [ 223.2 ]
check "and end within the text" (passing |> List.forall (fun (_, right, _, _) -> right <= 624.5)) true
check "below the picture the lines start at the margin again" (clear.Length >= 2, clear |> List.map (fun (left, _, _, _) -> left) |> List.distinct) (true, [ 0.0 ])

// The side the text may stand on, and a band wide enough on both sides.
let withPicture (change: ImageRun -> ImageRun) =
    { doc with
        Blocks =
            doc.Blocks
            |> List.map (function
                | Doc.Paragraph para -> Doc.Paragraph { para with Buffer = { para.Buffer with Runs = para.Buffer.Runs |> List.map (function Image image -> Image (change image) | run -> run) } }
                | block -> block) }
let image = Doc.paragraphsOfBlocks doc.Blocks |> List.pick (fun para -> para.Buffer.Runs |> List.tryPick (function Image image -> Some image | _ -> None))
check "the picture as read" (image.WrapText, image.Wrap) (Some "bothSides", Some "square")
// (Moved right by 2 in: 237.6 px of text left of it, 624 - 249.6 - 153.6 - 12 = 208.8 px right of it.)
let moved (side: string) = withPicture (fun image -> { image with PositionX = image.PositionX |> Option.map (fun x -> x + 192.0); WrapText = Some side })
let sides side =
    let passing = textLines (moved side) |> beside 144.0 259.2
    passing |> List.exists (fun (left, _, _, _) -> left < 237.6), passing |> List.exists (fun (_, right, _, _) -> right > 415.2 + 20.0), passing |> List.forall (fun (left, right, _, _) -> right <= 237.6 + 0.5 || left >= 415.2 - 0.5)
check "both sides: text left and right of the picture, none over it" (sides "bothSides") (true, true, true)
check "left: text left of it only" (sides "left") (true, false, true)
check "right: text right of it only" (sides "right") (false, true, true)
check "largest: the wider side (the left one)" (sides "largest") (true, false, true)

printfn "word wrap-band: %d passed, %d failed" pass fail
if fail > 0 then exit 1
