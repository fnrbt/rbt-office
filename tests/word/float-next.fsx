// The paragraphs after a square-wrapped picture's own, as Word 2019 (and 2010) lays them out, against
// Word's own results for the fixtures of tools/word-parity/fixtures/float-next.py
// (tests/word/fixtures/float-next/*.docx; Word's PDFs and word/geometry.tsv - pdf-geometry.py):
// - the picture keeps its box, distB under it, clear of the text of the paragraphs after its own too, in a
//   section with columns as on a page (col-*: ours kept it from its own paragraph only - the showcase's
//   magazine-feature.docx, "Moving a turbine fifty metres is" beside its turbines picture);
// - a later paragraph's first line goes beside it when its top less its paragraph's space before is
//   above the picture's bottom plus distB: the space before counts (b6-d9-m55 beside, -m65 not), the
//   paragraph above's space after does not (a6-d9-m05 not), nor its line spacing's extra (l15-*); distB
//   0, 9, 18 alike (d*-m05 not beside, d*-p05 beside); Word 2010's layout alike (-c14).
// Text is measured by its length here, as in float-wrap.fsx: which lines go beside the picture is
// compared (each of Word's lines of B, C and D down to below the picture: page, top, left), not their words.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-70s expected %A got %A" label expected got

let dir = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/float-next")
let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let pt (px: float) = px * 0.75
// (pdftotext's box of a Calibri word starts 0.199 of its size below its line's top: see float-wrap.fsx.)
let glyphTop (top: float) (bottom: float) = 0.1991 * (bottom - top) / 1.00364

let geometry = File.ReadAllLines(Path.Combine(dir, "word/geometry.tsv")) |> Array.map (fun l -> l.Split '\t')
/// Word's lines of B, C and D (page, top, left, first word) down to 20pt below the picture's bottom.
let wordLines name =
    let image = geometry |> Array.find (fun r -> r.[0] = name && r.[1] = "image")
    let bottom = float image.[4] + float image.[6]
    geometry
    |> Array.filter (fun r -> r.[0] = name && r.[1] = "line" && not (r.[7].StartsWith "a"))
    |> Array.map (fun r -> int r.[2], float r.[3] - glyphTop (float r.[3]) (float r.[4]), float r.[5], r.[7])
    |> Array.filter (fun (_, top, _, _) -> top < bottom + 20.0)
    |> Array.toList

/// Our lines: page, top and left of their text in points from the page's top-left.
let ours name =
    let doc = File.ReadAllBytes(Path.Combine(dir, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = { Layout.documentSettings 16.0 24.0 3.5 doc with SpaceShrink = 0.0; FirstPageExtraRoom = 0.0 }
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    let pageTop (page: int) = (List.item (page - 1) pages).Top
    [ for layout in layouts do
        for line in layout.Lines do
            let starts =
                line.Segments |> List.choose (function
                    | Layout.TextSegment (text, _, x, _, _) when text.Trim() <> "" -> Some x
                    | _ -> None)
            if not starts.IsEmpty then
                let page = pageOf line.Y
                yield page, pt (line.Y - pageTop page), pt (settings.PaddingLeft + List.min starts + Layout.alignmentOffset layout line) ]

let fixtures = Directory.GetFiles(dir, "*.docx") |> Array.map Path.GetFileNameWithoutExtension |> Array.sort |> Array.toList
check "the fixtures are all there" fixtures.Length 32
check "Word's results for each" (fixtures |> List.filter (fun name -> not (geometry |> Array.exists (fun r -> r.[0] = name)))) []
for name in fixtures do
    let ourLines = ours name
    let missing =
        wordLines name
        |> List.filter (fun (page, top, left, _) ->
            not (ourLines |> List.exists (fun (p, t, l) -> p = page && abs (t - top) <= 0.75 && abs (l - left) <= 0.6)))
        |> List.map (fun (page, top, left, first) -> page, Math.Round(top, 1), Math.Round(left, 1), first)
    check (sprintf "%s: lines where Word's are (page, top, left, first word)" name) missing []

printfn "float-next: %d passed, %d failed" pass fail
if fail > 0 then exit 1
