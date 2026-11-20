// Where a table cell's lines break at a page's foot, against Word 2019's PDFs of two fixture families (Word's
// lines in each folder's word/geometry.tsv - pdf-geometry.py). Compared: the page each of the cells' lines and the
// paragraph after the table is on. What Word showed (docs/layout-parity-flow.md, ledger):
// - cell-mark (tools/word-parity/fixtures/cell-mark.py; Courier New 10pt): a w:lastRenderedPageBreak in a cell
//   starts no page - at a paragraph's start or at a line's, in one cell or both, fitting or not - and takes
//   nothing of widow control or keepLines: each cell breaks where it would without it; keepLines in a cell, as
//   widow control, only in Word 2013's layout (kl-base-f55: broken in mode 14, the row moved in 15);
// - cell-foot (tools/word-parity/fixtures/cell-foot.py; Calibri 11pt): a cell's line stays when all of it - the
//   extra of its spacing multiple too, not only its text as in the body -, the cell's bottom margin and the
//   bottom border fit above the bottom margin (*-s115-*: a 1.15 line 0.25pt past it goes on).
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let labelRx = Text.RegularExpressions.Regex(@"^([PQ](L\d|w\d+)|R\d|C)$")

let expectedOf (folder: string) =
    File.ReadAllLines(Path.Combine(folder, "word/geometry.tsv"))
    |> Array.map (fun l -> l.Split('\t'))
    |> Array.filter (fun c -> c.Length > 7 && c.[1] = "line" && labelRx.IsMatch c.[7])
    |> Array.groupBy (fun c -> c.[0])
    |> Array.map (fun (name, rows) -> name, rows |> Array.map (fun c -> c.[7], int c.[2]) |> Map.ofArray)

// (Courier New: every character six tenths of the size across; the Calibri fixtures' lines are short.)
let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.6
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText

let oursOf (folder: string) (name: string) =
    let doc = File.ReadAllBytes(Path.Combine(folder, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    [ for layout in layouts do
        for line in layout.Lines do
            let text = line.Segments |> List.choose (function Layout.TextSegment (t, _, _, _, _) -> Some t | _ -> None) |> String.concat ""
            let first = text.Trim().Split(' ').[0]
            if labelRx.IsMatch first then first, pageOf line.Y ]
    |> Map.ofList

let mutable pass = 0
let mutable fail = 0
for family in [ "cell-mark"; "cell-foot" ] do
    let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures", family)
    for name, word in expectedOf folder do
        let ours = oursOf folder name
        let wrong = word |> Map.toList |> List.filter (fun (label, page) -> Map.tryFind label ours <> Some page)
        if wrong.IsEmpty then pass <- pass + 1
        else
            fail <- fail + 1
            let show = wrong |> List.truncate 6 |> List.map (fun (label, page) -> sprintf "%s Word p%d ours %s" label page (match Map.tryFind label ours with Some p -> sprintf "p%d" p | None -> "-"))
            printfn "FAIL %s/%-24s %s" family name (String.Join(", ", show))

printfn "Cell breaks against Word: %d fixtures as Word, %d not" pass fail
if fail > 0 then exit 1
