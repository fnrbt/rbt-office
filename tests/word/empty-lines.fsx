// Empty and blank lines, against Word 2019's PDFs of the fixtures tools/word-parity/fixtures generates
// (docs/layout-parity-flow.md, ledger):
//   mark-font         an empty paragraph's line is its mark's font and size (w:pPr/w:rPr, its w:rStyle, the
//                     paragraph style's - Normal's for a bare <w:p/> - not the document defaults')
//   blank-run-height  a run of only spaces or tabs takes no part in its line's height
//                     (word/heights.tsv: each probe paragraph's line, from Topq / Botq - fixtures' --table)
//   keep-next-fit     a keep-with-next paragraph stays when the next one's first line (two lines under
//                     widow control) fits where it stands, below the space between them
//   empty-foot        an empty paragraph past the page's foot starts the next page, after a section of
//                     another page size too
//                     (word/*.tsv, page-lines.vbs: the last page's first line with text)
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let fixtures = Path.Combine(__SOURCE_DIRECTORY__, "fixtures")
let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText

let layoutOf (file: string) =
    let doc = File.ReadAllBytes file |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    Layout.layoutDocumentPages settings measureText measureRange doc

let lineText (line: Layout.MeasuredLine) =
    line.Segments |> List.choose (function Layout.TextSegment (t, _, _, _, _) -> Some t | _ -> None) |> String.concat "" |> fun t -> t.Trim()

let mutable pass = 0
let mutable fail = 0
let check name ok detail =
    if ok then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-28s %s" name detail

// The probe's line (pt): (Botq's top - Topq's top - 13.43) / 10, as the generators' --table.
for family in [ "mark-font"; "blank-run-height" ] do
    let folder = Path.Combine(fixtures, family)
    for row in File.ReadAllLines(Path.Combine(folder, "word/heights.tsv")) |> Array.filter (fun l -> l.Length > 0) do
        let cells = row.Split('\t')
        let name, word = cells.[0], float cells.[1]
        let layouts, _ = layoutOf (Path.Combine(folder, name + ".docx"))
        let top text = layouts |> List.collect (fun l -> l.Lines) |> List.find (fun line -> lineText line = text) |> fun line -> line.Y
        let ours = ((top "Botq" - top "Topq") * 0.75 - 13.43) / 10.0
        check (family + "/" + name) (abs (ours - word) < 0.02) (sprintf "Word %.3f ours %.3f" word ours)

// The last page's first line with a letter or number.
for family in [ "keep-next-fit"; "empty-foot" ] do
    let folder = Path.Combine(fixtures, family)
    for file in Directory.GetFiles(Path.Combine(folder, "word"), "*.tsv") |> Array.sort do
        let name = Path.GetFileNameWithoutExtension file
        let rows = File.ReadAllLines file |> Array.filter (fun l -> l.Length > 0 && not (l.StartsWith "pages"))
        let word = (rows |> Array.last).Split('\t').[4]
        let layouts, pages = layoutOf (Path.Combine(folder, name + ".docx"))
        let lastTop = (pages |> List.last).Top
        let ours =
            layouts |> List.collect (fun l -> l.Lines) |> List.filter (fun line -> line.Y >= lastTop - 0.01)
            |> List.map lineText |> List.tryFind (fun t -> t |> Seq.exists Char.IsLetterOrDigit) |> Option.defaultValue ""
        check (family + "/" + name) (ours.StartsWith (word.Substring(0, min 12 word.Length))) (sprintf "Word %s / ours %s" word ours)

printfn "Empty and blank lines against Word: %d fixtures as Word, %d not" pass fail
if fail > 0 then exit 1
