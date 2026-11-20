// Where a table row broken across pages goes on, against Word 2019's PDFs of the fixtures
// tools/word-parity/fixtures/row-continue.py generates (tests/word/fixtures/row-continue; Word's side,
// the top of every label, in word/tops.tsv). Arial 10pt at exactly 20pt a line; a label's top in
// Word's PDF is 7.09pt below its line's (f01 at 79.09 for a line at the 72pt margin).
// What Word showed (docs/layout-parity-flow.md, ledger), in compatibility mode 14 and 15 alike:
// the part of a row that goes on on the next page starts as a row of its own does there - below the
// repeated header rows (or the page's top), the row's top border and its cells' top margin above its
// text again (h0-m100-b24: r04 8pt below the margin; h1-m100-b24: the header 3 + 5 + 20 + 5, then 3 + 5).
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/row-continue")
let labelGap = 7.09

// Word's: fixture -> label -> (page, the line's top on the page, pt).
let expected =
    File.ReadAllLines(Path.Combine(folder, "word/tops.tsv"))
    |> Array.filter (fun l -> l.Length > 0)
    |> Array.map (fun l -> l.Split('\t'))
    |> Array.filter (fun c -> Text.RegularExpressions.Regex.IsMatch(c.[4], @"^[a-z]\d\d$"))
    |> Array.groupBy (fun c -> c.[0])
    |> Array.map (fun (name, rows) ->
        name, rows |> Array.map (fun c -> (c.[4], int c.[1]), float c.[3] - labelGap) |> Map.ofArray)
    |> Map.ofArray

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let labelRx = Text.RegularExpressions.Regex(@"\b([a-z]\d\d)\b")

let oursOf (file: string) =
    let doc = File.ReadAllBytes file |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    [ for layout in layouts do
        for line in layout.Lines do
            let text = line.Segments |> List.choose (function Layout.TextSegment (t, _, _, _, _) -> Some t | _ -> None) |> String.concat " "
            for m in labelRx.Matches text ->
                let page = pageOf line.Y
                (m.Value, page), (line.Y - (pages |> List.item (page - 1)).Top) * 0.75 ]
    |> Map.ofList

let mutable pass = 0
let mutable fail = 0
for KeyValue (name, word) in expected do
    let ours = oursOf (Path.Combine(folder, name + ".docx"))
    let wrong =
        word |> Map.toList |> List.filter (fun (key, top) ->
            match Map.tryFind key ours with
            | Some t -> abs (t - top) > 0.3
            | None -> true)
    if wrong.IsEmpty then pass <- pass + 1
    else
        fail <- fail + 1
        let show =
            wrong |> List.truncate 4 |> List.map (fun ((label, page), top) ->
                sprintf "%s p%d Word %.2f ours %s" label page top (match Map.tryFind (label, page) ours with Some t -> sprintf "%.2f" t | None -> "-"))
        printfn "FAIL %-18s %s" name (String.Join(", ", show))

printfn "Row continuations against Word: %d fixtures as Word, %d not" pass fail
if fail > 0 then exit 1
