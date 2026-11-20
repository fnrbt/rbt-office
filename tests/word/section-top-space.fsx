// The space before of the first paragraph of a section that starts a new page (Layout's pageTopSpaceBeforeOverlap,
// the section before's last paragraph above it), against Word 2019's PDFs of the fixtures tools/word-parity/fixtures/section-top-space.py
// generates (tests/word/fixtures/section-top-space; Word's lines in word/geometry.tsv - pdf-geometry.py).
// Compared: the probe's top on its page (Word's glyph top less 2.16pt: Calibri 11pt's, 74.16 at the 72pt
// margin). What Word showed (docs/layout-parity-flow.md, ledger): the probe keeps of its space before what is
// more than the space after of the paragraph the section before ends with (12pt before after 6pt after: 6pt;
// after 12, 18 or 14 - auto -: none; 24pt before: 18, 12, 6, 10) - whatever that section is (continuous, in
// columns), the page size, keepNext; Word 2010's layout and 2013's alike.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/section-top-space")
let rows = File.ReadAllLines(Path.Combine(folder, "word/geometry.tsv")) |> Array.map (fun l -> l.Split('\t'))

let wordOf (name: string) =
    let probe = rows |> Array.find (fun c -> c.[0] = name && c.[1] = "line" && c.[7] = "P-probe")
    int probe.[2], float probe.[3] - 2.16

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
        |> List.find (fun line -> line.Segments |> List.exists (function Layout.TextSegment (t, _, _, _, _) -> t.StartsWith "P-probe" | _ -> false))
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
        printfn "FAIL %-20s Word: page %d at %.2fpt; ours page %d at %.2fpt" name wordPage wordTop ourPage ourTop

printfn "A section's first space before against Word: %d documents as Word, %d not" pass fail
if fail > 0 then exit 1
