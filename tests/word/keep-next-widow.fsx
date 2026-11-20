// A keepNext paragraph (or two) at a page's foot before a paragraph that cannot leave one line behind, against Word
// 2019's PDFs of the fixtures tools/word-parity/fixtures/keep-next-widow.py generates (tests/word/fixtures/keep-next-widow;
// Word's lines in word/geometry.tsv - pdf-geometry.py). Compared: the page of H-probe, K-probe and each line P1 ... Pn
// of the paragraph after them. What Word showed (docs/layout-parity-flow.md, ledger): the keepNext paragraphs stay
// only where the next paragraph can start on the page - its first lines that its widow control and keepLines keep
// together.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/keep-next-widow")
let rows = File.ReadAllLines(Path.Combine(folder, "word/geometry.tsv")) |> Array.map (fun l -> l.Split('\t'))
let probe (word: string) = word = "H-probe" || word = "K-probe" || (word.StartsWith "P" && word.Length <= 3)

let wordOf (name: string) =
    rows |> Array.filter (fun c -> c.[0] = name && c.[1] = "line" && probe c.[7]) |> Array.map (fun c -> c.[7], int c.[2]) |> Array.sort

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText

let oursOf (name: string) =
    let doc = File.ReadAllBytes(Path.Combine(folder, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    layouts |> List.collect (fun l -> l.Lines)
    |> List.choose (fun line ->
        let text = line.Segments |> List.map (function Layout.TextSegment (t, _, _, _, _) -> t | _ -> "") |> String.concat ""
        let first = text.Trim().Split(' ').[0]
        if probe first then Some (first, pageOf line.Y) else None)
    |> List.toArray |> Array.sort

let show (lines: (string * int)[]) = lines |> Array.map (fun (w, p) -> sprintf "%s:%d" w p) |> String.concat " "
let mutable pass = 0
let mutable fail = 0
for name in rows |> Array.map (fun c -> c.[0]) |> Array.distinct |> Array.sort do
    let word, ours = wordOf name, oursOf name
    if word = ours then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-16s Word: %s\n     %-16s ours: %s" name (show word) "" (show ours)

printfn "keepNext before a paragraph kept together against Word: %d documents as Word, %d not" pass fail
if fail > 0 then exit 1
