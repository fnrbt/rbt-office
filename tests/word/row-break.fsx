// Where a table row breaks across pages, against Word 2019's own pages of the
// fixtures tools/word-parity/fixtures/row-break.py generates (tests/word/fixtures/row-break;
// Word's side, the page every label is on in Word's PDF, in word/labels.tsv).
// Every line of a fixture is exactly 20pt and one label, so the pages do not depend
// on how text is measured. What Word showed (docs/layout-parity-flow.md, ledger):
// - each cell of a row breaks by itself, at its own last break that fits (cell-two:
//   a cell of 40pt lines beside one of 20pt; nest2-text: a nested row splits between
//   its lines while the text beside it breaks a line further down) - between the lines
//   of a paragraph, between the rows of a nested table or inside one of them by this
//   same rule (depth2-merge), a nested cantSplit row going on whole (nest-cantsplit);
// - a vertical merge sets no bounds: the rows of a merge group break apart, at the
//   top level and nested alike (top-merge*, nest-merge*, small-*, tall-nestmerge*);
// - when a cell would keep nothing on the page - its first line does not fit
//   (first-tall, first-space), its nested table's first row is cantSplit (nest-first-
//   cantsplit), orphan control or keepLines takes its paragraph (widow-two_f31,
//   nest-first-orphan_f31, keeplines-two) - the whole row goes to the next page, the
//   other cells' lines with it; a nested row so goes on within its cell (nest-orphan);
//   behind a repeated header row it goes under the header again (header-first-tall);
// - a cantSplit row that does not fit goes to the next page whole (top-cantsplit); one
//   taller than a page breaks there all the same, at the page's bottom (tall-cantsplit);
// - widow/orphan control: a first line alone stays not, a last line alone takes the one
//   before it along (widow-*: 10 lines with room for 9 keep 8), in the body and in cells
//   alike; with no w:widowControl anywhere it is on (widow-unset*);
// - a line of an exact height fits when all of it does, not only its text (exact-body).
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/row-break")
let expected =
    File.ReadAllLines(Path.Combine(folder, "word/labels.tsv"))
    |> Array.filter (fun l -> l.Length > 0)
    |> Array.map (fun l -> let c = l.Split('\t') in c.[0], (c.[1], int c.[2]))
    |> Array.groupBy fst
    |> Array.map (fun (name, rows) -> name, rows |> Array.map snd |> Map.ofArray)
    |> Map.ofArray

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let labelRx = Text.RegularExpressions.Regex(@"\b([a-z]\d\d[a-z]?|end)\b")

let oursOf (file: string) =
    let doc = File.ReadAllBytes file |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    [ for layout in layouts do
        for line in layout.Lines do
            let text = line.Segments |> List.choose (function Layout.TextSegment (t, _, _, _, _) -> Some t | _ -> None) |> String.concat " "
            for m in labelRx.Matches text -> m.Value, pageOf line.Y ]
    |> Map.ofList

let mutable pass = 0
let mutable fail = 0
let failedFamilies = Collections.Generic.Dictionary<string, int>()
for KeyValue (name, word) in expected do
    let ours = oursOf (Path.Combine(folder, name + ".docx"))
    let wrong = word |> Map.toList |> List.filter (fun (label, page) -> Map.tryFind label ours <> Some page)
    if wrong.IsEmpty then pass <- pass + 1
    else
        fail <- fail + 1
        let family = name.Split('_').[0]
        failedFamilies.[family] <- (match failedFamilies.TryGetValue family with | true, n -> n + 1 | _ -> 1)
        let show = wrong |> List.truncate 6 |> List.map (fun (label, page) -> sprintf "%s Word p%d ours %s" label page (match Map.tryFind label ours with Some p -> sprintf "p%d" p | None -> "-"))
        printfn "FAIL %-28s %s" name (String.Join(", ", show))

printfn ""
for KeyValue (family, n) in failedFamilies do printfn "  %s: %d" family n
printfn "Row breaks against Word: %d fixtures as Word, %d not" pass fail
if fail > 0 then exit 1
