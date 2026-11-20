// Which section's bottom margin a page takes when a continuous section break starts a section part-way down
// it (Layout's layoutSections, FirstPageExtraRoom), against Word 2019's pages of the fixtures
// tools/word-parity/fixtures/section-margins.py generates (tests/word/fixtures/section-margins; Word's pages,
// page-lines.vbs, in word/*.tsv). Compared: the first line of every page. What Word showed
// (docs/layout-parity-flow.md, ledger): a page's bottom margin is that of the section the page starts in -
// one, two and three sections on a page, one column and two, Word 2010's layout and 2013's alike; a section
// that starts part-way down a page takes its own from the next page on (and its top margin there).
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/section-margins")

let wordOf (name: string) =
    File.ReadAllLines(Path.Combine(folder, "word", name + ".tsv"))
    |> Array.skip 1 |> Array.map (fun line -> line.Split('\t').[2]) |> Array.toList

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText

let lineText (line: Layout.MeasuredLine) =
    line.Segments |> List.map (function Layout.TextSegment (text, _, _, _, _) -> text | _ -> "") |> String.concat "" |> fun s -> s.Trim()

let oursOf (name: string) =
    let doc = File.ReadAllBytes(Path.Combine(folder, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    let lines = layouts |> List.collect (fun l -> l.Lines) |> List.filter (fun l -> lineText l <> "")
    // (Columns: the first line of a page is its topmost leftmost - the lines come column by column.)
    [ for page in 1 .. pages.Length do
        match lines |> List.filter (fun l -> pageOf l.Y = page) with
        | [] -> ()
        | first :: _ -> yield lineText first ]

let mutable pass = 0
let mutable fail = 0
for name in Directory.GetFiles(folder, "*.docx") |> Array.map Path.GetFileNameWithoutExtension |> Array.sort do
    let word = wordOf name
    let ours = oursOf name
    if word = ours then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-22s Word's pages start %A; ours %A" name word ours

printfn "Section margins against Word: %d documents as Word, %d not" pass fail
if fail > 0 then exit 1
