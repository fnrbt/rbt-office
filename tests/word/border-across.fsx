// Where a paragraph's borders run, across and down, against Word 2019's own PDFs of the fixtures
// `tools/word-parity/fixtures/para-border.py --across` generates (tests/word/fixtures/para-border-across;
// Word's borders in word/geometry.tsv - `pdf-geometry.py --paths`, the filled rectangles Word draws a
// border's lines with). Every fixture: a01, the bordered paragraph b01 (or an empty one), c01, Arial
// 10pt at exactly 20pt a line, Letter, 1in margins (the text from 72pt to 540pt).
// What Word showed (docs/layout-parity-flow.md, ledger):
// - a top or bottom border runs from the left edge of the text (the left indent, or the first line's
//   where further left) less 1.5pt to the right indent plus 1.5pt, whatever its width, space or line
//   style (bottom, bottom-sz*, -sp10, -double, top), with text or without (bottom-empty), whatever
//   the alignment (bottom-center), the indents (bottom-ind, -neg, -hang, -first), in a table's cell
//   (cell-*: the cell's text area) and in a column (cols-*), Word 2010's layout alike (*-c14);
// - a side border's line is outside that, its space further out (box-sp0, -sp4, -sp10, box-ind), its
//   width outwards from there (box-sz48, box-left48); the top and bottom borders run out to the side
//   borders' outer edges (bottom-left, bottom-right, box-*);
// - down: the top border's line is above the text, its space between (top); the bottom one's below
//   the text after its space (bottom-sp10); a double line is two lines of its width, a width apart
//   (bottom-double); the side borders run from the top border's outer edge to the bottom one's, or
//   from the text's top / to its bottom where there is none (bottom-left, bottom-right).
// Our side: Paint's commands (StrokeLine, centred on its line) for the layout of each fixture; across
// in points from the page's left edge, down from the top of a01's line (Word's at the margin, 72pt).
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout
module Paint = Rbt.Word.Paint

// (And `--corners`, tests/word/fixtures/para-border-corners: double lines at a box's corners - Word draws
// two frames, the outer lines meeting at the outer corner, the inner at the inner one: dbl-box, -c14,
// -sz12, -mixed, -group; where a double border meets a single one, the double one's lines run out to
// the single one's outer edge, the single one's to the double one's (dbl-topbottom, dbl-sides).)
let families = [ "fixtures/para-border-across", 32; "fixtures/para-border-corners", 8 ]
let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let pt (px: float) = px * 0.75

/// A border's painted area: horizontal or vertical, left, top, right, bottom (pt).
type Band = { Horizontal: bool; Left: float; Top: float; Right: float; Bottom: float }

/// Rectangles merged into the bands they make: those of one row (same top and bottom) that touch,
/// across; of one column, down. A square (a corner) joins both.
let bands (rects: (float * float * float * float) list) =
    let merge horizontal (rs: (float * float * float * float) list) =
        let key (l, t, r, b) = if horizontal then (t, b) else (l, r)
        let span (l, t, r, b) = if horizontal then (l, r) else (t, b)
        rs
        |> List.groupBy (fun r -> let a, b = key r in Math.Round(a * 4.0), Math.Round(b * 4.0))
        |> List.collect (fun (_, group) ->
            let a, b = key group.Head
            group |> List.map span |> List.sortBy fst
            |> List.fold (fun acc (s, e) ->
                match acc with
                | (s0, e0) :: rest when s <= e0 + 0.13 -> (s0, max e0 e) :: rest
                | _ -> (s, e) :: acc) []
            |> List.map (fun (s, e) ->
                if horizontal then { Horizontal = true; Left = s; Top = a; Right = e; Bottom = b }
                else { Horizontal = false; Left = a; Top = s; Right = b; Bottom = e }))
    let w (l, _, r, _) = r - l
    let h (_, t, _, b) = b - t
    let horizontal = rects |> List.filter (fun r -> w r >= h r - 0.13)
    let vertical = rects |> List.filter (fun r -> h r >= w r - 0.13)
    // (A side's band is only a side's where it is longer than wide: corners alone are not one.)
    let across = merge true horizontal |> List.filter (fun b -> b.Right - b.Left > b.Bottom - b.Top + 0.2)
    // (A side's line drawn up to the top border's inner edge or over it is the same on the page: it is
    // taken up to the top border's outer edge, and down to the bottom one's.)
    let down =
        merge false vertical |> List.filter (fun b -> b.Bottom - b.Top > b.Right - b.Left + 0.2)
        |> List.map (fun v ->
            let meets (h: Band) = h.Left <= v.Left + 0.13 && h.Right >= v.Right - 0.13
            let top = across |> List.filter (fun h -> meets h && h.Top <= v.Top + 0.13 && h.Bottom >= v.Top - 0.13) |> List.fold (fun t h -> min t h.Top) v.Top
            let bottom = across |> List.filter (fun h -> meets h && h.Top <= v.Bottom + 0.13 && h.Bottom >= v.Bottom - 0.13) |> List.fold (fun b h -> max b h.Bottom) v.Bottom
            { v with Top = top; Bottom = bottom })
    across @ down |> List.sortBy (fun b -> b.Horizontal, b.Top, b.Left)

let wordBandsOf dir =
    File.ReadAllLines(Path.Combine(dir, "word/geometry.tsv")) |> Array.map (fun l -> l.Split '\t')
    |> Array.filter (fun r -> r.[1] = "path" && r.[2] = "1")
    |> Array.groupBy (fun r -> r.[0])
    |> Array.map (fun (name, rs) ->
        // (Down from a01's line's top: Word's at the margin.)
        name, rs |> Array.map (fun r -> float r.[3], float r.[4] - 72.0, float r.[5], float r.[6] - 72.0) |> Array.distinct |> Array.toList |> bands)
    |> Map.ofArray

let oursOf dir name =
    let doc = File.ReadAllBytes(Path.Combine(dir, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let a01 =
        layouts |> List.pick (fun layout ->
            layout.Lines |> List.tryPick (fun line ->
                let text = line.Segments |> List.choose (function Layout.TextSegment (t, _, _, _, _) -> Some t | _ -> None) |> String.concat ""
                if text = "a01" then Some line.Y else None))
    let _, commands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Arial" measureText measureRange doc layouts pages
    Paint.flattenedCommands commands
    |> List.choose (function
        | Paint.StrokeLine (_, width, _, _, x1, y1, x2, y2) when abs (y1 - y2) < 0.01 ->
            Some (pt (min x1 x2), pt (y1 - width / 2.0 - a01), pt (max x1 x2), pt (y1 + width / 2.0 - a01))
        | Paint.StrokeLine (_, width, _, _, x1, y1, x2, y2) when abs (x1 - x2) < 0.01 ->
            Some (pt (x1 - width / 2.0), pt (min y1 y2 - a01), pt (x1 + width / 2.0), pt (max y1 y2 - a01))
        | _ -> None)
    |> bands

let mutable pass = 0
let mutable fail = 0
for family, count in families do
  let dir = Path.Combine(__SOURCE_DIRECTORY__, family)
  let wordBands = wordBandsOf dir
  let fixtures = Directory.GetFiles(dir, "*.docx") |> Array.map Path.GetFileNameWithoutExtension |> Array.sort
  if fixtures.Length <> count || fixtures |> Array.exists (fun name -> not (wordBands.ContainsKey name)) then
    printfn "FAIL %s: the fixtures and Word's borders of each: %d fixtures, %d with Word's" family fixtures.Length wordBands.Count
    fail <- fail + 1
  for name in fixtures do
    let word = wordBands.[name]
    let ours = oursOf dir name
    // (Word draws on a 300dpi grid: its edges are within 0.25pt of where they are. In a table's cell
    // Word's text area is 0.5pt narrower on each side than ours - a table without cell margins, open
    // in the ledger -: the border's ends there are compared to 0.7pt.)
    let acrossBy = if name.StartsWith "cell-" then 0.7 else 0.3
    let same (a: Band) (b: Band) =
        a.Horizontal = b.Horizontal && abs (a.Left - b.Left) <= acrossBy && abs (a.Right - b.Right) <= acrossBy
        && abs (a.Top - b.Top) <= 0.3 && abs (a.Bottom - b.Bottom) <= 0.3
    let missing = word |> List.filter (fun w -> not (ours |> List.exists (same w)))
    let extra = ours |> List.filter (fun o -> not (word |> List.exists (same o)))
    if missing.IsEmpty && extra.IsEmpty then pass <- pass + 1
    else
        fail <- fail + 1
        let show (b: Band) = sprintf "%s %.2f,%.2f-%.2f,%.2f" (if b.Horizontal then "H" else "V") b.Left b.Top b.Right b.Bottom
        printfn "FAIL %-20s Word's not ours: %s | ours not Word's: %s" name (missing |> List.map show |> String.concat "; ") (extra |> List.map show |> String.concat "; ")

printfn "border-across: %d fixtures as Word, %d not" pass fail
if fail > 0 then exit 1
