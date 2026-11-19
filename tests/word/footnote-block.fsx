// The footnote block as Word 2019 lays it out, against Word's own results for the fixtures of
// tools/word-parity/fixtures/footnote-block.py (tests/word/fixtures/footnote-block/*.docx; Word's pages
// in word/*.tsv - page-parity.mjs --dir --refresh-word -, its PDFs' baselines and rules in
// word/pdf-lines.tsv - pdf-baselines.py - and its paragraph and footnote line positions in
// word-positions/*.tsv - para-positions.vbs; Information(6) is in whole pixels, 0.75pt):
// - the notes are stacked up from the bottom margin: the last note's last line, and its space after,
//   end at it (Word's PDFs: the last baseline 2.69pt above 720pt for Calibri 10pt);
// - each note is its paragraphs with their own spacing (FootnoteText's, or what it inherits: 1.15 and
//   8pt after under such a w:pPrDefault), the first one's space before and the last one's space after
//   included, nothing between the notes;
// - a line of a note is as tall as its text: the raised reference mark (or a superscript letter on
//   every line, or a custom mark) does not make a 10pt Calibri line taller (12.21pt);
// - the separator is a paragraph of its own above the notes (footnotes.xml's w:type="separator"):
//   double-spaced, or in 20pt, it takes that much more room;
// - the body runs down to the separator's top: a line whose text ends below it goes on.
// Not yet ours (the cont-* fixtures, recorded in docs/layout-parity-flow.md): a note too long for the
// room under its reference is broken there, as a paragraph with widow control - the reference's line
// stays when two of its lines fit (one without widow control), the rest goes on under the next
// page's continuation separator.
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
        printfn "FAIL %-60s expected %A got %A" label expected got
let near (tolerance: float) label (got: float) (expected: float) =
    if abs (got - expected) <= tolerance then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-60s expected %.3f got %.3f" label expected got

let dir = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/footnote-block")
let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let toPt (px: float) = px * 72.0 / 96.0

/// Our layout of a fixture: the page (1-based) of each body paragraph in order, and the footnote
/// lines as (page, top in points from the page's top).
let ours name =
    let doc = File.ReadAllBytes(Path.Combine(dir, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    let pageTop (page: int) = (List.item (page - 1) pages).Top
    let seen = Collections.Generic.HashSet<string>()
    let body =
        [ for l in layouts do
            if not (l.BlockId.StartsWith "__") && seen.Add l.BlockId then
                yield (match l.Lines with line :: _ -> pageOf line.Y | [] -> pageOf l.Y) ]
    let notes =
        [ for l in layouts do
            if l.BlockId.StartsWith Layout.footnoteBlockIdPrefix then
                for line in l.Lines do
                    let page = pageOf line.Y
                    yield page, toPt (line.Y - pageTop page), toPt line.Height ]
        |> List.sortBy (fun (page, top, _) -> page, top)
    body, notes

/// Word's: the page of each body paragraph, and the tops of the footnotes' lines.
let word name =
    let rows = File.ReadAllLines(Path.Combine(dir, "word-positions", name + ".tsv")) |> Array.map (fun l -> l.Split '\t')
    let body = [ for r in rows do if r.[0] |> Seq.forall Char.IsDigit then yield int r.[1] ]
    let notes = [ for r in rows do if r.[0].StartsWith "fn" then yield float r.[2] ]
    body, notes

let fixtures =
    Directory.GetFiles(dir, "*.docx") |> Array.map Path.GetFileNameWithoutExtension |> Array.sort |> Array.toList
// The note broken at its reference (not ours yet), and two where Word's single line of Calibri 11pt
// (13.4307pt in its PDFs, ours 13.4277) puts the 45th line's text 0.03 and 0.3pt below the limit.
let pending = set [ "cont-40"; "cont-k44"; "cont-k45"; "cont-k45-nowidow"; "cont-k46-nowidow"; "limit-p00"; "dflt115-inherit-two" ]
check "the fixtures are all there" (fixtures.Length) 35

for name in fixtures do
    let ourBody, ourNotes = ours name
    let wordBody, wordNotes = word name
    if not (pending.Contains name) then
        check (sprintf "%s: the pages of the body's paragraphs" name) ourBody wordBody
        if not (name.StartsWith "cont") then
            check (sprintf "%s: footnote lines" name) ourNotes.Length wordNotes.Length
            // (Word's tops are whole pixels, rounded either way: 707.79pt is 943 or 944px.)
            List.iteri2 (fun i (_, top, _) wordTop -> near 0.75 (sprintf "%s: footnote line %d's top" name (i + 1)) top wordTop) ourNotes wordNotes

// The block ends at the bottom margin, the last line's bottom at 720pt.
let lastBottom name = let _, notes = ours name in notes |> List.last |> fun (_, top, height) -> top + height
for name in [ "one-1"; "one-3"; "two-1"; "custom-1"; "nosuper-1"; "super-10" ] do
    near 0.01 (sprintf "%s: the last footnote line ends at the bottom margin" name) (lastBottom name) 720.0
// ... and the last note's space after (12pt; 8pt inherited) with it.
near 0.01 "fnspaced-2: the last line and its 12pt after end at 720pt" (lastBottom "fnspaced-2" + 12.0) 720.0
near 0.01 "dflt115-inherit-2: the last line and its inherited 8pt after end at 720pt" (lastBottom "dflt115-inherit-2" + 8.0) 720.0
// A 10pt Calibri note line is 12.21pt, its reference mark raised or not.
for name in [ "one-1"; "nosuper-1"; "custom-1" ] do
    let _, notes = ours name
    near 0.01 (sprintf "%s: the note's line is 12.21pt" name) (notes.Head |> fun (_, _, h) -> h) 12.207
// The separator is its paragraph: Body 46 (its text 689.8pt down) fits above a single-spaced
// Calibri 11pt separator (13.43pt) and a one-line note, not above a double-spaced or a 20pt one.
check "one-1: Body 46 on page 1" (fst (ours "one-1") |> List.item 45) 1
check "sep-double-1: Body 46 on page 2" (fst (ours "sep-double-1") |> List.item 45) 2
check "sep-big-1: Body 46 on page 2" (fst (ours "sep-big-1") |> List.item 45) 2

printfn "footnote-block: %d passed, %d failed" pass fail
if fail > 0 then exit 1
