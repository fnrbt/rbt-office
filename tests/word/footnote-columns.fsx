// The footnotes of a section with columns as Word 2019 lays them out (Word 2013's layout, compatibility
// mode 15), against Word's own PDFs of the fixtures of tools/word-parity/fixtures/footnote-columns.py
// (tests/word/fixtures/footnote-columns/*.docx; where Word put every line, from pdf-geometry.py, in
// word/geometry.tsv):
// - a page's notes are set in the columns of the section of its first note's reference, each at its
//   column's width, one column after the other and balanced as the columns at a section's end are -
//   a note's paragraph broken between columns with widow control (a 3-line note is not), each column's
//   notes from the top of the tallest; one separator above them, at the first column;
// - the text of every column of the page runs down to the separator, not only the reference's column;
// - a line whose note's room comes over the text of a column before it: the page is set again with
//   that room under every column - the lines it takes going on into the next column with the line -,
//   or, when the line is then no longer on the page, the page ends before the line.
// Checked: the first and last line of each column of each page, and how many note lines each column's
// foot holds (for the notes of one long sentence, whose lines are the measurer's here, only which columns
// hold them; `page-parity.mjs --dist dist --dir tests/word/fixtures/footnote-columns` checks those pages
// with Word's fonts).
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-60s\n     expected %A\n     got      %A" label expected got

let dir = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/footnote-columns")
let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText

/// Word's lines: (page, left from the left margin in points, first word, last word), from its PDF.
let wordRows =
    File.ReadAllLines(Path.Combine(dir, "word", "geometry.tsv"))
    |> Array.map (fun l -> l.Split '\t')
    |> Array.filter (fun r -> r.[1] = "line")
    |> Array.groupBy (fun r -> r.[0])
    |> Map.ofArray

let isText (word: string) = word = "Body" || word = "Top" || word = "After"

/// The columns' left edges (from the left margin, points): where Word's text lines start.
let columnLefts name =
    wordRows.[name]
    |> Array.filter (fun r -> isText r.[7])
    |> Array.map (fun r -> Math.Round(float r.[5] - 72.0))
    |> Array.distinct
    |> Array.sort

let columnOf (lefts: float[]) (x: float) = (lefts |> Array.filter (fun left -> left <= x + 2.0)).Length - 1

/// A page's columns: per page, per column, the first and last text line ("Body 12"), and the number of note lines.
type Shape = { Text: Map<int * int, string * string>; Notes: Map<int * int, int> }

let shapeOf (lines: (int * float * string * bool) list) (lefts: float[]) =
    let text =
        lines
        |> List.filter (fun (_, _, _, note) -> not note)
        |> List.groupBy (fun (page, x, _, _) -> page, columnOf lefts x)
        |> List.map (fun (key, ls) -> key, ((let (_, _, t, _) = List.head ls in t), (let (_, _, t, _) = List.last ls in t)))
        |> Map.ofList
    let notes =
        lines
        |> List.filter (fun (_, _, _, note) -> note)
        |> List.countBy (fun (page, x, _, _) -> page, columnOf lefts x)
        |> Map.ofList
    { Text = text; Notes = notes }

let word name =
    let rows = wordRows.[name] |> Array.sortBy (fun r -> int r.[2], float r.[5], float r.[3])
    [ for r in rows do
        let first, last = r.[7], r.[8]
        let x = float r.[5] - 72.0
        if isText first then yield int r.[2], x, first + " " + last, false
        // (A note's line; not its mark alone, nor the separator's word.)
        elif not (first = last && first |> Seq.forall Char.IsDigit) && first <> "Sepword" then yield int r.[2], x, "", true ]

let ours name =
    let doc = File.ReadAllBytes(Path.Combine(dir, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    [ for l in layouts do
        let note = l.BlockId.StartsWith Layout.footnoteBlockIdPrefix
        for line in l.Lines do
            match line.Segments |> List.tryPick (function Layout.TextSegment (t, _, x, _, _) -> Some (x, t) | _ -> None) with
            | Some (x, t) -> yield pageOf line.Y, x * 0.75, (if note then "" else t.Trim()), note
            | None -> () ]
    |> List.sortBy (fun (page, x, _, _) -> page, Math.Round x)

let fixtures =
    Directory.GetFiles(dir, "*.docx") |> Array.map Path.GetFileNameWithoutExtension |> Array.sort |> Array.toList
check "the fixtures are all there" fixtures.Length 47

// Word 2010's layout (m14-*: each note under its reference's column) is not ours; huge-c1's note
// (120 lines) runs on over the page, its page's text not as Word's.
let pending = set [ "m14-one-c2"; "m14-three-c1"; "m14-two-c1c2-tall2"; "m14-wrap-c1"; "m14-bal-c2"; "huge-c1" ]
// The notes of one long sentence: their lines are the measurer's (only the columns they use are checked).
let measured (name: string) = name.Contains "wrap" || name.StartsWith "uneq" || name.StartsWith "mixed" || name.StartsWith "four"

for name in fixtures do
    if not (pending.Contains name) then
        let lefts = columnLefts name
        let w = shapeOf (word name) lefts
        let o = shapeOf (ours name) lefts
        // (Where the text of a page with a long sentence's note ends depends on the note's lines: those
        // are checked in the browser with Word's fonts - page-parity.mjs --dir.)
        if not (measured name) then check (sprintf "%s: each column's first and last line" name) o.Text w.Text
        if measured name then check (sprintf "%s: the columns that hold notes" name) (o.Notes |> Map.toList |> List.map fst) (w.Notes |> Map.toList |> List.map fst)
        else check (sprintf "%s: note lines by column" name) o.Notes w.Notes

printfn "footnote-columns: %d passed, %d failed" pass fail
if fail > 0 then exit 1
