// A paragraph that starts with a w:lastRenderedPageBreak (where Word last broke the page), against Word 2019's
// PDFs of the fixtures tools/word-parity/fixtures/heading-mark.py generates (tests/word/fixtures/heading-mark;
// Word's lines in word/geometry.tsv - pdf-geometry.py). Compared: the probe's page and its top on it (Word's glyph
// top less 2.81pt for the 14pt bold heading, 2.16 for the 11pt paragraph: 72 at the margin). What Word showed
// (docs/layout-parity-flow.md, ledger): the mark starts no page - a heading with it stays in the middle of the
// page, and one the flow takes to a page's top is there without its space before, as one without it; the
// layout takes no mark as a page's start (stale-*).
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/heading-mark")
let rows = File.ReadAllLines(Path.Combine(folder, "word/geometry.tsv")) |> Array.map (fun l -> l.Split('\t'))

let wordOf (name: string) =
    let probe = rows |> Array.find (fun c -> c.[0] = name && c.[1] = "line" && c.[7] = "H-probe")
    int probe.[2], float probe.[3] - (if name.Contains "plain" then 2.16 else 2.81)

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let toPt (px: float) = px * 72.0 / 96.0

let oursOf (name: string) =
    let doc = File.ReadAllBytes(Path.Combine(folder, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    let line =
        layouts |> List.collect (fun l -> l.Lines)
        |> List.find (fun line ->
            let text = line.Segments |> List.map (function Layout.TextSegment (t, _, _, _, _) -> t | _ -> "") |> String.concat ""
            text.StartsWith "H-probe")
    let page = pageOf line.Y
    page, toPt (line.Y - pages.[page - 1].Top)

let mutable pass = 0
let mutable fail = 0
for name in rows |> Array.map (fun c -> c.[0]) |> Array.distinct |> Array.sort do
    let wordPage, wordTop = wordOf name
    let ourPage, ourTop = oursOf name
    if wordPage = ourPage && abs (wordTop - ourTop) < 0.3 then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-22s Word: page %d at %.2fpt; ours page %d at %.2fpt" name wordPage wordTop ourPage ourTop

printfn "Rendered page break marks against Word: %d documents as Word, %d not" pass fail
if fail > 0 then exit 1
