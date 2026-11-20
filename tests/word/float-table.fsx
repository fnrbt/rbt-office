// Text round a floating table (w:tblpPr) as Word 2019 lays it out, against Word's own results for
// the fixtures of tools/word-parity/fixtures/float-table.py (tests/word/fixtures/float-table/*.docx;
// Word's pages and PDFs - page-parity.mjs --dir --refresh-word -, and in word/geometry.tsv where its
// PDFs put each line's text, across and down, the table's cells' lines among them -
// tools/word-parity/pdf-geometry.py):
// - the table is placed from the paragraph after it (vertAnchor text: that paragraph's top, its space
//   before above it, plus tblpY), from the page's margin or the page (tblpY, or tblpYSpec top /
//   center / bottom), and across from the text's (the margin's) or the page's left edge (tblpX), or
//   where tblpXSpec puts its box (left, center, right; inside / outside on an odd page) - in Word
//   2010's layout hanging out by its cells' margin on the side it is aligned to, as a table in the
//   flow does on the left (base, y*, ym*, vm-*, vp-*, ys-*, h*-x*, ind-*, xs-*, *-c14);
// - it takes no room in the flow: the text goes round it on both sides as round a square-wrapped
//   picture - every line whose box reaches into the table's box widened by its distances from text
//   (left / right / top / bottomFromText; across never nearer than half a point to its outer border),
//   in a band a quarter inch wide or more that a word fits in; the paragraphs after its anchor too
//   (base, cal, d-*, gap-*, gapr-*, gapil-*, wide, short, empty, sb12, sz12);
// - it is placed once: lines above its paragraph that it reaches go round it, the paragraph moves down
//   (ym20, ys-top);
// - broken at a page's foot as a table in the flow is, its paragraph goes on beside the part of it on
//   the page it ends on, from the top of that page (pb, pb-big);
// - the first-line indent of a line beside it starts from the table (fl-b), and a table is not moved
//   in by Normal's first-line indent (fl-normal, fl-normal-inline);
// - a table in the flow that would reach into its box (its distances from text across included) starts
//   below it: below its bottomFromText in Word 2019's layout (flow-base, -db18, -ind, -left, -near,
//   -two, -fixed*, -center), below its border in Word 2010's for a fixed-width table (flow-fixed-c14,
//   -fixed-db18-c14, -fixed-center-c14); one clear of it stays beside it (flow-narrow*). Word 2010's
//   layout narrows an autofit one to the room beside it instead (flow-base-c14, -db18-c14, -center-c14:
//   not here - flow-center-c14 is open).
// The cells' paragraphs of a table with a table style (ts-*): Normal's first-line indent, and its
// font size, over the table style's (the rule: tests/word/table-style-props.fsx).
// Text is measured here by its length, so only where lines are and where their pieces start (and end,
// when justified) is compared, and where the table's first cell's text is. Not ours yet, recorded in
// docs/layout-parity-flow.md: ov-never (tblOverlap never: Word moves the second table to the right of
// the first), cell, nested (a floating table in a table's cell), vm-y0 (vertAnchor margin, tblpY 0:
// Word lays it out in the flow), ys-*-text (tblpYSpec with vertAnchor text: Word puts the table a
// fraction of a point above its paragraph, and the line above goes beside it), sz12 (where the
// table's first cell's text is: half a 1.5pt border).
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
        printfn "FAIL %-70s expected %A got %A" label expected got

let dir = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/float-table")
let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let pt (px: float) = px * 0.75

type Row = { Page: int; Top: float; Left: float; Right: float; First: string; Last: string; Single: bool; Final: bool }

// (pdftotext's box of a Calibri word starts 0.199 of its size below its line's top.)
let glyphTop (top: float) (bottom: float) = 0.1991 * (bottom - top) / 1.00364
// The last words of the fixtures' paragraphs: a justified paragraph's last line is not justified.
let finals = set [ "a40"; "b020"; "b120"; "b200"; "c60"; "i40"; "p10" ]
let isCell (word: string) = word.Length = 3 && (word.[0] = 't' || word.[0] = 'u' || word.[0] = 'n' || word.[0] = 'f') && Char.IsDigit word.[1]

/// Word's lines of a fixture, in the PDF's order.
let wordRows =
    File.ReadAllLines(Path.Combine(dir, "word/geometry.tsv")) |> Array.map (fun l -> l.Split '\t')
    |> Array.filter (fun r -> r.[1] = "line")
    |> Array.groupBy (fun r -> r.[0])
    |> Array.map (fun (name, rs) ->
        name, rs |> Array.map (fun r -> { Page = int r.[2]; Top = float r.[3] - glyphTop (float r.[3]) (float r.[4]); Left = float r.[5]; Right = float r.[6]
                                          First = r.[7]; Last = r.[8]; Single = r.[7] = r.[8]; Final = finals.Contains r.[8] }) |> Array.toList)
    |> Map.ofArray

/// Our lines: each piece of one (page, top, left and right of its text in points from the page's
/// top-left, whether it is its paragraph's last), the table's cells' lines among them.
let ours name =
    let doc = File.ReadAllBytes(Path.Combine(dir, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    let pageTop (page: int) = (List.item (page - 1) pages).Top
    [ for layout in layouts do
        let last = List.tryLast layout.Lines
        for line in layout.Lines do
            let spans =
                line.Segments |> List.choose (function
                    | Layout.TextSegment (text, style, x, _, _) when text.Trim() <> "" -> Some (x, x + measureText (text.TrimEnd()) style, text.Trim())
                    | _ -> None)
            if not spans.IsEmpty then
                let page = pageOf line.Y
                let offset = Layout.alignmentOffset layout line
                // (A line beside a table is pieces of text with the table between them: 144pt or more.)
                let pieces =
                    spans |> List.fold (fun acc (left, right, text) ->
                        match acc with
                        | (l, r, first) :: rest when left - r < 100.0 -> (l, right, first) :: rest
                        | _ -> (left, right, text) :: acc) [] |> List.rev
                let isLast = match last with Some l -> obj.ReferenceEquals(l, line) | None -> false
                for left, right, first in pieces do
                    yield page, pt (line.Y - pageTop page), pt (settings.PaddingLeft + left + offset), pt (settings.PaddingLeft + right + offset), first, isLast ]

let fixtures = Directory.GetFiles(dir, "*.docx") |> Array.map Path.GetFileNameWithoutExtension |> Array.sort |> Array.toList
check "the fixtures are all there" fixtures.Length 96
check "Word's results for each" (fixtures |> List.filter (fun name -> not (wordRows.ContainsKey name))) []

let open_ = set [ "ov-never"; "cell"; "nested"; "vm-y0"; "ys-top-text"; "ys-center-text"; "ys-bottom-text"
                  // (Word 2010's layout narrows an autofit table in the flow to the room beside a floating
                  // one, and puts it at the margin: not here.)
                  "flow-center-c14" ]
// (Word's table stands half its 1.5pt left border further in than a table in the flow of ours: in
// Word 2019's layout a table's outer border, not its grid, is at the margin - open; the text round
// it is compared.)
let cellOpen = set [ "sz12" ]
// The paragraphs after the anchor's are compared where they go beside the table too.
let withC = set [ "short"; "empty"; "pb-big" ]

for name in fixtures |> List.filter (fun name -> not (open_.Contains name)) do
    let ourRows = ours name
    let rows = wordRows.[name]
    let ourAt (row: Row) =
        ourRows |> List.tryFind (fun (page, top, left, _, _, _) -> page = row.Page && abs (top - row.Top) <= 0.75 && abs (left - row.Left) <= 0.6)
    // Where the table is: its first cell's text, each table's.
    let firstCells = rows |> List.filter (fun row -> isCell row.First && row.First.EndsWith "11")
    let missingCells = firstCells |> List.filter (fun row -> (ourAt row).IsNone) |> List.map (fun row -> row.First, row.Page, Math.Round(row.Top, 1), Math.Round(row.Left, 1))
    if not (cellOpen.Contains name) then
        check (sprintf "%s: the table's first cell where Word's is (page, top, left)" name) missingCells []
    // Word's lines of the text up to C's first (C may start further down here: B's words are measured
    // otherwise), each with a line of ours at its place.
    let text = rows |> List.filter (fun row -> not (isCell row.First))
    let checkedRows =
        if withC.Contains name then text
        else
            let firstC = text |> List.tryFindIndex (fun row -> row.First = "c01") |> Option.defaultValue text.Length
            text |> List.truncate firstC
    // (Where a line's first piece ends at the table's first cell's text, pdftotext ran them together.)
    let placed = checkedRows |> List.filter (fun row -> not row.Single && not row.Final && not (isCell row.Last))
    let missing = placed |> List.filter (fun row -> (ourAt row).IsNone) |> List.map (fun row -> row.Page, Math.Round(row.Top, 1), Math.Round(row.Left, 1))
    check (sprintf "%s: lines where Word's are (page, top, left)" name) missing []
    // A justified piece of a line (not a paragraph's last, here or in Word, nor one word) ends where
    // Word's does: at the table less its distance from text, or at the margin.
    let wrongEnds =
        placed |> List.choose (fun row ->
            match ourAt row with
            | Some (_, _, _, right, _, false) when abs (right - row.Right) > 1.0 -> Some (row.Page, Math.Round(row.Top, 1), Math.Round(row.Right, 1), Math.Round(right, 1))
            | _ -> None)
    check (sprintf "%s: justified lines end where Word's do" name) wrongEnds []

printfn "float-table: %d passed, %d failed" pass fail
if fail > 0 then exit 1
