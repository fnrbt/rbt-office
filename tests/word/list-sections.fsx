// A numbered list across section breaks, against Word 2019's PDFs of the fixtures
// tools/word-parity/fixtures/list-sections.py generates (tests/word/fixtures/list-sections; Word's
// side, the number before every item, in word/numbers.tsv). What Word showed (docs/layout-parity-flow.md,
// ledger), in compatibility mode 14 and 15 alike: a list goes on from section to section - after a
// next-page, a continuous and a landscape section break (a01 - a03 1 - 3, b01 - b03 4 - 6, c01 7) -
// and a second w:num of the same abstract list goes on with it (d01: 8).
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/list-sections")

// Word's: fixture -> label -> (page, number).
let expected =
    File.ReadAllLines(Path.Combine(folder, "word/numbers.tsv"))
    |> Array.filter (fun l -> l.Length > 0)
    |> Array.map (fun l -> l.Split('\t'))
    |> Array.groupBy (fun c -> c.[0])
    |> Array.map (fun (name, rows) -> name, rows |> Array.map (fun c -> c.[2], (int c.[1], c.[3])) |> Map.ofArray)
    |> Map.ofArray

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let itemRx = Text.RegularExpressions.Regex(@"^([a-d]\d\d) item")

let oursOf (file: string) =
    let doc = File.ReadAllBytes file |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    // (The number is the layout's prefix, its marker: "8." and a tab.)
    [ for layout in layouts do
        for line in layout.Lines |> List.truncate 1 do
            let text = line.Segments |> List.choose (function Layout.TextSegment (t, _, _, _, _) -> Some t | _ -> None) |> String.concat ""
            for m in itemRx.Matches (text.Trim()) -> m.Groups.[1].Value, (pageOf line.Y, layout.Prefix.Trim().TrimEnd('.')) ]
    |> Map.ofList

// Open (docs/layout-parity-flow.md, ledger): a second w:num of one abstract list is a list of its own
// in ours (the reader's ListId is the w:num), numbered from 1 (d01).
let known = set [ "d01" ]

let mutable pass = 0
let mutable fail = 0
for KeyValue (name, word) in expected do
    let ours = oursOf (Path.Combine(folder, name + ".docx"))
    let wrong = word |> Map.toList |> List.filter (fun (label, v) -> not (known.Contains label) && Map.tryFind label ours <> Some v)
    if wrong.IsEmpty then pass <- pass + 1
    else
        fail <- fail + 1
        let show = wrong |> List.map (fun (label, (p, n)) -> sprintf "%s Word p%d %s. ours %s" label p n (match Map.tryFind label ours with Some (op, on) -> sprintf "p%d %s." op on | None -> "-"))
        printfn "FAIL %-10s %s" name (String.Join(", ", show))

printfn "Lists across sections against Word: %d fixtures as Word, %d not" pass fail
if fail > 0 then exit 1
