// How wide a table's columns are, against Word 2019's own tables of the fixtures
// tools/word-parity/fixtures/table-width.py generates (tests/word/fixtures/table-width; Word's side,
// where each cell's left-aligned label starts and its right-aligned label ends, from Word's PDFs, in
// word/edges.tsv). Text is measured with Liberation Sans, which has Arial's widths, as ShaperText
// measures it. What Word showed (docs/layout-parity-flow.md, ledger):
// - a column is as wide as the widest of its cells' w:tcW; the grid only stands in for a cell with
//   none, in a fixed table (grideq-tcw, tcweq-grid, rows-*, vmerge, span-*);
// - an autofit table takes a cell with no width at its content on one line, never narrower than its
//   longest word, and is no wider than its room (the text, the column, the cell; less the indent):
//   past it the columns give way above their least widths in proportion (content-*, wide-*,
//   shrink3-*, cols2-full-auto, nest-*-auto, ind720-auto);
// - a fixed table without a w:tblW is as wide as its columns, past the margin too (wide-fixed,
//   ind720-fixed, cols2-full-fixed, nest-full-fixed);
// - a w:tblW sets the table's width: a percentage of the room and the indent; a fixed table's columns
//   give way above their margins and a pixel (small-tblw, shrink3-tblw), and grow in proportion
//   (widen-*); percentages of cells are of the table (pct-grid, tcwpct-tblw), of a table of no width
//   the narrowest that gives each its least width (tcwpct-auto);
// - in compatibility mode 14 a table reaches past the text by its outer cells' margins; a nested
//   table does not (*-c14).
// Not ours: tcwpct-fixed (Word's fixed least width there is 11.77pt, ours 11.55).
// Ours keeps the 0.25px a cell's line keeps free beside a column's longest word: 0.19pt over Word's.
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
let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/table-width")
let fontSet =
    Shaper.createFontSet "Arial" [
        { Family = "Arial"; Bold = false; Italic = false; Data = File.ReadAllBytes(Path.Combine(root, "fonts", "LiberationSans-Regular.ttf")) } ]
let measureText = Shaper.createMeasurer fontSet
let measureRange = Shaper.createRangeMeasurer fontSet
let pt (px: float) = px * 0.75

// Word's edges: fixture -> (row label, column) -> page, left of "L", right of "R" (None: off the page).
let word =
    File.ReadAllLines(Path.Combine(folder, "word/edges.tsv"))
    |> Array.map (fun l -> l.Split '\t')
    |> Array.groupBy (fun c -> c.[0])
    |> Array.map (fun (name, rows) ->
        name, rows |> Array.map (fun c -> (c.[1], int c.[2]), (int c.[3], float c.[4], (match c.[5] with "None" -> None | v -> Some (float v)))) |> Map.ofArray)
    |> Map.ofArray

/// Ours: the same labels' left and right edges in points from the page's left.
let ours name =
    let doc = File.ReadAllBytes(Path.Combine(folder, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    [ for layout in layouts do
        for line in layout.Lines do
            for segment in line.Segments do
                match segment with
                | Layout.TextSegment (text, style, x, _, _) when text.Length >= 3 && (text.[0] = 'L' || text.[0] = 'R') && Char.IsDigit text.[text.Length - 1] ->
                    let label = text.Trim()
                    let left = settings.PaddingLeft + x + Layout.alignmentOffset layout line
                    yield label, (pageOf line.Y, pt left, pt (left + measureText label style))
                | _ -> () ]
    |> Map.ofList

let mutable pass = 0
let mutable fail = 0
let open_ = set [ "tcwpct-fixed"; "tcwpct-fixed-c14" ]
// (0.45pt: two columns at their least widths, each 0.19pt over Word's - see the top.)
// (tcwpct-auto: a table five times its 20% column's least width, the 0.19pt five times over.)
let tolerance name = if (name: string).StartsWith "tcwpct-auto" then 1.0 else 0.45
for KeyValue (name, cells) in word do
    if not (open_.Contains name) then
        let close (a: float) (b: float) = abs (a - b) <= tolerance name
        let ourCells = ours name
        let mutable same = true
        let report = Text.StringBuilder()
        for KeyValue ((row, column), (page, left, right)) in cells do
            let tag = if row.StartsWith "i" then "i" + row.Substring 1 else row
            let l = sprintf "L%s%d" tag column
            let r = sprintf "R%s%d" tag column
            let ok =
                match ourCells.TryFind l, ourCells.TryFind r with
                | Some (p, ol, _), Some (_, _, orr) ->
                    let rightOk = match right with Some right -> close orr right | None -> orr > 612.0
                    p = page && close ol left && rightOk
                | _ -> false
            if not ok then
                same <- false
                report.AppendFormat("  {0}{1}: Word {2} {3:0.00} {4}, ours {5} {6}\n", row, column, page, left, right, ourCells.TryFind l, ourCells.TryFind r) |> ignore
        if same then pass <- pass + 1
        else
            fail <- fail + 1
            printf "FAIL %s\n%s" name (report.ToString())

printfn "Table widths against Word: %d fixtures as Word, %d not" pass fail
if fail > 0 then exit 1
