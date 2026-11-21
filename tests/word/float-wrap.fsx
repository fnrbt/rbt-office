// Text round a floating picture as Word 2019 lays it out, against Word's own results for the
// fixtures of tools/word-parity/fixtures/float-wrap.py (tests/word/fixtures/float-wrap/*.docx; Word's
// pages and PDFs - page-parity.mjs --dir --refresh-word -, and in word/geometry.tsv where its PDFs put
// each line's text, across and down, and each picture - tools/word-parity/pdf-geometry.py):
// - a picture placed from its paragraph stands at the paragraph's top (plus its offset) wherever in
//   the paragraph its anchor run is, and every line of the paragraph whose box reaches into the
//   picture's box (widened by distT / distB) goes beside it - square - or below it, distB under it -
//   top and bottom (sq-start / sq-mid / sq-end, sq-h4*, sq-top-*, sq-dist*, tb-*);
// - it is placed once: a line above its paragraph that it reaches goes round it (or below it) too,
//   and the paragraph moves down for that, the picture staying where it was (sq-start: A's last line
//   beside it; kl-tb: a line above it moves below it);
// - the text goes on both sides, on the left or right only, or on the wider side (side-*), into a
//   band a quarter inch wide or more when a word fits in it (gap-*, gapil-*: none in 16pt, words in
//   18pt, in compatibilityMode 15, 14, 12 and none);
// - beside it a justified line is justified to the end of its piece of the line (all of them);
// - in a section with columns `column` is the anchor's column and `margin` the page's margins
//   (col-colright, col-coloff, col-marginright, col-pageoff: where the picture is across);
// - a shape that text goes round, in a section with columns, is placed from its paragraph's top in the
//   column the paragraph starts in, and the column's lines go round it; the paragraph's lines that fit
//   below it stay (widow control as on a page), the rest go on in the next column; when none fit the
//   paragraph goes on in the next column whole, the shape with it in Word 2019's layout, left where it
//   was in Word 2010's (colq-*, col-tb-low-start; where the painter draws the shape, from its text
//   "Pull quote");
// - a paragraph that ends in a column break leaves its mark at the next column's top, a line as high as
//   the mark's font, its space after below it (cbr-*, and B in the col-* fixtures).
// Every line is Calibri 11pt; text is measured here by its length, so only where lines are and where
// they start (and end, when justified) is compared - not which words they hold. Not ours yet, recorded
// in docs/layout-parity-flow.md: col-marginoff (the first column's text going round a picture anchored
// in the second), col-tb-low, col-tb-low-c14, colq-k43-end-* (a shape whose anchor is in its
// paragraph's part in the next column: the anchor's place in a paragraph is not read), split2-* (a
// picture's paragraph across a page break), kl-one-c14 / kl-two-c14 / rep* (Word 2010's layout moving a
// keepLines paragraph down beside the picture).
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

let dir = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/float-wrap")
let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let pt (px: float) = px * 0.75

type Row = { Page: int; Top: float; Left: float; Right: float; Single: bool; Final: bool; Quote: bool }

// (pdftotext's box of a Calibri word starts 0.199 of its size below its line's top: 2.19pt at 11pt.
// Its height is the size times 1.00364.)
let glyphTop (top: float) (bottom: float) = 0.1991 * (bottom - top) / 1.00364
// The last words of the fixtures' paragraphs: a justified paragraph's last line is not justified.
let finals = set [ "a40"; "b040"; "b048"; "b120"; "b200"; "b300"; "c60"; "c800"; "k30"; "t05"; "Kept" ]

/// Word's lines and pictures of a fixture.
let wordRows, wordImages =
    let rows = File.ReadAllLines(Path.Combine(dir, "word/geometry.tsv")) |> Array.map (fun l -> l.Split '\t')
    let lines =
        rows |> Array.filter (fun r -> r.[1] = "line")
        |> Array.groupBy (fun r -> r.[0])
        |> Array.map (fun (name, rs) ->
            name, rs |> Array.map (fun r -> { Page = int r.[2]; Top = float r.[3] - glyphTop (float r.[3]) (float r.[4]); Left = float r.[5]; Right = float r.[6]; Single = r.[7] = r.[8]; Final = finals.Contains r.[8]; Quote = r.[7] = "Pull" }) |> Array.toList)
        |> Map.ofArray
    let images =
        rows |> Array.filter (fun r -> r.[1] = "image")
        |> Array.groupBy (fun r -> r.[0])
        |> Array.map (fun (name, rs) -> name, rs |> Array.map (fun r -> int r.[2], float r.[3], float r.[4]) |> Array.toList)
        |> Map.ofArray
    lines, images

/// Where Word put each fixture's shape (page, left, top), from its text "Pull quote": the shape's insets
/// (0.1in across, 0.05in down) and the word's box in its line before it.
let wordShapes =
    File.ReadAllLines(Path.Combine(dir, "word/geometry.tsv")) |> Array.map (fun l -> l.Split '\t')
    |> Array.filter (fun r -> r.[1] = "line" && r.[7] = "Pull")
    |> Array.map (fun r -> r.[0], (int r.[2], float r.[5] - 0.02 - 7.2, float r.[3] - glyphTop (float r.[3]) (float r.[4]) - 3.6))
    |> Map.ofArray

/// Our lines (each piece of one: page, top, left and right of its text in points from the page's
/// top-left) and pictures (page, left, top).
let ours name =
    let doc = File.ReadAllBytes(Path.Combine(dir, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    // (Not the narrowing of a justified line's spaces - Layout's SpaceShrink, tests/word/justify-squeeze.fsx:
    // this measurer's spaces are half an em, Calibri's 0.226 of one, and a quarter of them takes some words
    // more onto a line than Word's does.)
    let settings = { Layout.documentSettings 16.0 24.0 3.5 doc with SpaceShrink = 0.0; FirstPageExtraRoom = 0.0; NoteColumns = [] }
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    let pageTop (page: int) = (List.item (page - 1) pages).Top
    let rows =
        [ for layout in layouts do
            let last = List.tryLast layout.Lines
            for line in layout.Lines do
                let spans =
                    line.Segments |> List.choose (function
                        | Layout.TextSegment (text, style, x, _, _) when text.Trim() <> "" -> Some (x, x + measureText (text.TrimEnd()) style)
                        | _ -> None)
                if not spans.IsEmpty then
                    let page = pageOf line.Y
                    let left = (spans |> List.map fst |> List.min) + Layout.alignmentOffset layout line
                    let right = (spans |> List.map snd |> List.max) + Layout.alignmentOffset layout line
                    let isLast = match last with Some l -> obj.ReferenceEquals(l, line) | None -> false
                    yield page, pt (line.Y - pageTop page), pt (settings.PaddingLeft + left), pt (settings.PaddingLeft + right), isLast ]
    let images =
        [ for layout in layouts do
            for line in layout.Lines do
                for segment in line.Segments do
                    match segment with
                    | Layout.ImageSegment (_, x, yOffset, _, _, _, true, _, _, _, _, _, _, _) ->
                        let y = line.Y + yOffset
                        let page = pageOf y
                        yield page, pt (settings.PaddingLeft + x), pt (y - pageTop page)
                    | _ -> () ]
    // (The shapes where the painter draws them: page, left, top.)
    let shapes =
        Rbt.Word.Paint.bodyShapeRects settings measureText measureRange doc layouts pages
        |> List.map (fun (_, x, y, _, _) -> let page = pageOf y in page, pt x, pt (y - pageTop page))
    rows, images, shapes

let fixtures = Directory.GetFiles(dir, "*.docx") |> Array.map Path.GetFileNameWithoutExtension |> Array.sort |> Array.toList
check "the fixtures are all there" fixtures.Length 165
check "Word's results for each" (fixtures |> List.filter (fun name -> not (wordRows.ContainsKey name))) []

// Not ours yet (see the top), and where Word's picture is at the very bottom of the line above it
// (gap-36-c14: the line goes beside it in Word, not here - a hundredth of a point). The mag-* documents
// (the showcase's magazine, its pull quote changed) are Word's record only: their faces are not measured here.
let open_ = set [ "col-marginoff"; "col-tb-low"; "col-tb-low-c14"; "colq-k43-end-c14"; "colq-k43-end-c15"; "split2-sq-start"; "split2-sq-end"; "split2-tb-end"
                  "kl-one-c14"; "kl-two-c14"; "rep"; "rep-noempty"; "rep-single"; "gap-36-c14" ]
// In the showcase's faces (Segoe UI, Georgia) the words' boxes stand otherwise in their lines: the picture only.
// Where A's last line goes on on both sides of the picture its words decide whether A takes a third
// line, and B moves down with it: where the pieces of B's lines start across is compared.
let acrossLines = set [ "side-both-c"; "side-both-c-c14"; "side-both-r" ]
let pictureOnly name = (name: string).StartsWith "rep" || acrossLines.Contains name

for name in fixtures |> List.filter (fun name -> not (open_.Contains name) && not (name.StartsWith "mag-")) do
    let ourRows, ourImages, ourShapes = ours name
    match Map.tryFind name wordShapes with
    | Some (wordPage, wordLeft, wordTop) ->
        match ourShapes |> List.tryFind (fun (page, left, top) -> page = wordPage && abs (left - wordLeft) <= 0.6 && abs (top - wordTop) <= 0.6) with
        | Some _ -> pass <- pass + 1
        | None ->
            fail <- fail + 1
            printfn "FAIL %s: the shape at (%d, %.1f, %.1f), ours %A" name wordPage wordLeft wordTop (ourShapes |> List.map (fun (p, l, t) -> p, Math.Round(l, 1), Math.Round(t, 1)))
    | None -> ()
    let wordImageList = wordImages |> Map.tryFind name |> Option.defaultValue []
    check (sprintf "%s: pictures" name) ourImages.Length wordImageList.Length
    for (wordPage, wordLeft, wordTop), (page, left, top) in List.zip wordImageList (List.truncate wordImageList.Length ourImages) do
        let same = page = wordPage && abs (left - wordLeft) <= 0.6 && abs (top - wordTop) <= 0.6
        if same then pass <- pass + 1
        else
            fail <- fail + 1
            printfn "FAIL %s: the picture at (%d, %.1f, %.1f), ours (%d, %.1f, %.1f)" name wordPage wordLeft wordTop page left top
    if acrossLines.Contains name then
        let starts (rows: float list) = rows |> List.map (fun x -> Math.Round(x, 0)) |> List.distinct |> List.sort
        let wordStarts = wordRows.[name] |> List.filter (fun row -> not row.Single && not row.Final) |> List.map (fun row -> row.Left) |> starts
        check (sprintf "%s: where the pieces of lines start across" name) (ourRows |> List.map (fun (_, _, left, _, _) -> left) |> starts) wordStarts
    if not (pictureOnly name) then
        // Word's lines of A, B and the kept paragraph (C may start further down here: B's words are
        // measured otherwise, so it may have more lines), each with a line of ours at its place.
        let rows = wordRows.[name]
        let ourAt (row: Row) =
            ourRows |> List.tryFind (fun (page, top, left, _, _) -> page = row.Page && abs (top - row.Top) <= 0.75 && abs (left - row.Left) <= 0.6)
        // (Word's rows up to C's first, "c01"; in two columns, where the rows are in order down the
        // page and not of the text, all but C's.)
        let checkedRows =
            let names = File.ReadAllLines(Path.Combine(dir, "word/geometry.tsv")) |> Array.filter (fun l -> l.StartsWith(name + "\tline\t")) |> Array.map (fun l -> (l.Split '\t').[7])
            if name.StartsWith "colq-" || name.StartsWith "cbr-" then
                List.zip rows (List.ofArray names) |> List.filter (fun (_, first) -> not (first.StartsWith "c")) |> List.map fst
            else
            let firstC = names |> Array.tryFindIndex (fun first -> first = "c01") |> Option.defaultValue names.Length
            rows |> List.truncate firstC
        // (A word alone in a piece of a line may stand anywhere in it, justified; and a paragraph's
        // last line holds what is left of its words, which are measured otherwise here. The shape's
        // own text, "Pull quote", is not in the body's lines.)
        let placed = checkedRows |> List.filter (fun row -> not row.Single && not row.Final && not row.Quote)
        let missing = placed |> List.filter (fun row -> (ourAt row).IsNone) |> List.map (fun row -> row.Page, Math.Round(row.Top, 1), Math.Round(row.Left, 1))
        check (sprintf "%s: lines where Word's are (page, top, left)" name) missing []
        // A justified piece of a line (not a paragraph's last, here or in Word, nor one word) ends where
        // Word's does: at the picture less its distance from text, or at the margin.
        let wrongEnds =
            checkedRows
            |> List.filter (fun row -> not row.Single && not row.Final)
            |> List.choose (fun row ->
                match ourAt row with
                | Some (_, _, _, right, false) when abs (right - row.Right) > 1.0 -> Some (row.Page, Math.Round(row.Top, 1), Math.Round(row.Right, 1), Math.Round(right, 1))
                | _ -> None)
        check (sprintf "%s: justified lines end where Word's do" name) wrongEnds []

printfn "float-wrap: %d passed, %d failed" pass fail
if fail > 0 then exit 1
