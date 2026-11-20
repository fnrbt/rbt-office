// Where lines break round a word longer than its line, in fields' results and at the
// foot of a page, against Word 2019's own lines of the fixtures
// tools/word-parity/fixtures/long-word.py generates (tests/word/fixtures/long-word;
// Word's side, every line of every fixture from Word's PDF, in word/lines.tsv).
// Text is measured with the faces of fonts/ that have the Office fonts' widths
// (Liberation Mono for Courier New, Carlito for Calibri), as ShaperText measures it.
// What Word showed (docs/layout-parity-flow.md, ledger):
// - a word longer than the line goes to the next line, and is broken there where the
//   line is full, its next lines likewise (body-*, cell-*: Courier New, 30 characters
//   a line; with one character's room left, none, in a cell, in compatibility mode 14,
//   over two runs, after a glued "§", before a hyphen, two in a row);
// - a field's result and a link's text break as any text (fld-*, fldq-*, hf-quote*);
// - docx4j-headers-footers' narrow cell: "§Error! " / "Reference" / "source" / "not" /
//   "found." - "Reference" (46.24pt of Calibri Bold, not kerned) fits its 46.25pt line;
// - a row broken at the page's foot keeps the lines of its cells that fit (foot_f*,
//   with widow control; hf-foot_f*).
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Unicode.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Bidi.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.OpenType.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Shaper.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Text.Engine.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/ShaperText/bin/Debug/net10.0/Rbt.Word.ShaperText.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
module Layout = Rbt.Word.Layout
module Shaper = Rbt.Word.ShaperText

let root = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "../.."))
let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/long-word")
let face family bold file : Shaper.FontFace =
    { Family = family; Bold = bold; Italic = false; Data = File.ReadAllBytes(Path.Combine(root, "fonts", file)) }
let fontSet =
    Shaper.createFontSet "Calibri" [
        face "Calibri" false "Carlito-Regular.ttf"; face "Calibri" true "Carlito-Bold.ttf"
        face "Courier New" false "LiberationMono-Regular.ttf"; face "Courier New" true "LiberationMono-Bold.ttf" ]
let measureText = Shaper.createMeasurer fontSet
let measureRange = Shaper.createRangeMeasurer fontSet

let squeeze (text: string) = String(text.ToCharArray() |> Array.filter (fun c -> not (Char.IsWhiteSpace c)))

// Word's lines: fixture -> [page, text] in order (the cells of a row side by side make one line).
let expected =
    File.ReadAllLines(Path.Combine(folder, "word/lines.tsv"))
    |> Array.filter (fun l -> l.Length > 0)
    |> Array.map (fun l -> let c = l.Split('\t') in c.[0], (int c.[1], squeeze c.[3]))
    |> Array.groupBy fst
    |> Array.map (fun (name, rows) -> name, rows |> Array.map snd |> Array.toList)
    |> Map.ofArray

let oursOf (file: string) =
    let doc = File.ReadAllBytes file |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    [ for layout in layouts do
        for line in layout.Lines do
            let segments = line.Segments |> List.choose (function Layout.TextSegment (t, _, x, _, _) -> Some (x, t) | _ -> None)
            match segments with
            | [] -> ()
            | (x, _) :: _ -> yield pageOf line.Y, line.Y, x, segments |> List.map snd |> String.concat "" ]
    // (Lines within 2px of one another are one line of Word's PDF, left to right.)
    |> List.sortBy (fun (page, y, x, _) -> page, y, x)
    |> List.fold (fun acc (page, y, x, text) ->
        match acc with
        | (p, y0, t) :: rest when p = page && abs (y - y0) < 2.0 -> (p, y0, t + text) :: rest
        | _ -> (page, y, text) :: acc) []
    |> List.rev
    |> List.map (fun (page, _, text) -> page, squeeze text)
    |> List.filter (fun (_, text) -> text <> "")

let mutable pass = 0
let mutable fail = 0
for KeyValue (name, word) in expected do
    let ours = oursOf (Path.Combine(folder, name + ".docx"))
    if ours = word then pass <- pass + 1
    else
        fail <- fail + 1
        let show lines = lines |> List.map (fun (p, t) -> sprintf "%d:%s" p t) |> String.concat " | "
        printfn "FAIL %s\n  Word %s\n  ours %s" name (show word) (show ours)

printfn "Long words against Word: %d fixtures as Word, %d not" pass fail
if fail > 0 then exit 1
