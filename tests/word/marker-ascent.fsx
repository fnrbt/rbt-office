// How far a list marker in a taller font raises its line (Layout's ascentOfText, the first line's growth in
// layoutParagraph), against Word 2019's PDFs of the fixtures tools/word-parity/fixtures/marker-ascent.py
// generates (tests/word/fixtures/marker-ascent; Word's lines in word/geometry.tsv - pdf-geometry.py). Each
// document is 30 one-line list paragraphs (or plain ones); compared: their pitch on page 1, Word's from its
// text's tops. What Word showed (docs/layout-parity-flow.md, ledger):
// - the first line is its text's line at the paragraph's spacing plus the marker's ascent above the text's
//   tallest, each a font's single line less its usWinDescent (its line gap above the ascender), the marker at
//   its own size: Calibri 11pt at 1.15 under Word's Symbol bullet 16.03pt, a 20pt "1." 24.01;
// - the multiple does not scale what the marker adds (single 22.00, double 35.43 under the 20pt "1.");
// - what of the marker is below the text's descent adds nothing (Calibri 20pt under a 22pt Courier New "o":
//   28.08, as plain);
// - an at-least line is the more of its height and the single line so raised (15pt: 22.00; 30pt: 30.00);
//   an exact line is not grown.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/marker-ascent")

// Word's: the tops of each document's "Line n" texts on page 1.
let wordTops =
    File.ReadAllLines(Path.Combine(folder, "word/geometry.tsv"))
    |> Array.map (fun l -> l.Split('\t'))
    // (A row's last word: a big marker's row can take the text in, "1. Line 1 of the list"; in `mix-*` the
    // 16pt "Big" is a row of its own.)
    |> Array.filter (fun c -> c.Length > 8 && c.[1] = "line" && c.[2] = "1" && (if c.[0].StartsWith "mix" then c.[8] = "Big" else c.[8] = "list"))
    |> Array.groupBy (fun c -> c.[0])
    |> Array.map (fun (doc, rows) -> doc, rows |> Array.map (fun c -> float c.[3]))
let pitch (tops: float[]) = (tops.[tops.Length - 1] - tops.[0]) / float (tops.Length - 1)

// (Line heights only: the advances do not matter.)
let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let toPt (px: float) = px * 72.0 / 96.0

let oursPitch (name: string) =
    let doc = File.ReadAllBytes(Path.Combine(folder, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageTwo = match pages with _ :: second :: _ -> second.Top | _ -> infinity
    let tops = layouts |> List.collect (fun l -> l.Lines) |> List.map (fun l -> l.Y) |> List.filter (fun y -> y < pageTwo) |> Array.ofList
    toPt (pitch tops), tops.Length

let mutable pass = 0
let mutable fail = 0
for name, tops in wordTops |> Array.sortBy fst do
    let word = pitch tops
    let ours, count = oursPitch name
    if abs (ours - word) < 0.05 && count = tops.Length then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-24s Word %.3fpt a line (%d on page 1), ours %.3f (%d)" name word tops.Length ours count

printfn "List markers' lines against Word: %d documents as Word, %d not" pass fail
if fail > 0 then exit 1
