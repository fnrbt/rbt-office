// How high an empty paragraph is (DocxReader.paragraphMarkFontSize), against Word 2019's PDFs of the fixtures
// tools/word-parity/fixtures/empty-mark.py generates (tests/word/fixtures/empty-mark; Word's lines in
// word/geometry.tsv - pdf-geometry.py). Compared: how far below "A 1" the probe after five empty paragraphs
// stands (one font above and below: the glyphs' offset from their lines' tops cancels). What Word showed
// (docs/layout-parity-flow.md, ledger): an empty paragraph's mark has its style's size - a bare <w:p/> as much
// as one with properties, after a paragraph or a table, Word 2010's layout and 2013's alike.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/empty-mark")
let rows = File.ReadAllLines(Path.Combine(folder, "word/geometry.tsv")) |> Array.map (fun l -> l.Split('\t'))

let wordOf (name: string) =
    let top word = rows |> Array.find (fun c -> c.[0] = name && c.[1] = "line" && c.[7] = word) |> fun c -> float c.[3]
    top "P-probe" - top "A"

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let toPt (px: float) = px * 72.0 / 96.0

let oursOf (name: string) =
    let doc = File.ReadAllBytes(Path.Combine(folder, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, _ = Layout.layoutDocumentPages settings measureText measureRange doc
    let top (start: string) =
        layouts |> List.collect (fun l -> l.Lines)
        |> List.find (fun line -> line.Segments |> List.exists (function Layout.TextSegment (t, _, _, _, _) -> t.StartsWith start | _ -> false))
        |> fun line -> line.Y
    toPt (top "P-probe" - top "A 1")

let mutable pass = 0
let mutable fail = 0
for name in rows |> Array.map (fun c -> c.[0]) |> Array.distinct |> Array.sort do
    let word = wordOf name
    let ours = oursOf name
    if abs (word - ours) < 0.3 then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-20s Word: the probe %.2fpt below A 1; ours %.2fpt" name word ours

printfn "Empty paragraphs against Word: %d documents as Word, %d not" pass fail
if fail > 0 then exit 1
