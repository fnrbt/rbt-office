// How paragraph and table borders take vertical space, against Word 2019's own pages of
// the fixtures tools/word-parity/fixtures/para-border.py and table-border.py generate
// (tests/word/fixtures/para-border, table-border; Word's side, where every label's glyphs
// start on its page, in word/positions.tsv). Every line is exactly 20pt and one label:
// a01, the paragraphs or table rows under test, c01. The gaps between consecutive labels
// less 20pt are what borders, cell margins and spacing took; ours must be Word's within
// 0.3pt (Word's PDF positions are good to about 0.12pt).
// What Word showed (docs/layout-parity-flow.md, ledger) - paragraphs (para-border):
// - a top or bottom border takes its space and its width, nothing more: sz 4/12/24/48
//   (0.5-6pt) with space 0/1/4/10 (top-*, bottom-*); a double line three widths;
// - space before and after go outside the border (box1-before, box1-after);
// - adjoining paragraphs with the same borders are one box: the top border above the
//   first, the bottom below the last, nothing between (box3, topbottom3), whatever
//   their styles (grp-style-*), shading (grp-shade) or spacing (grp-after, grp-before:
//   the spacing stays); a different side, width, space, colour or line style ends it
//   (grp-sides, grp-width, grp-space, grp-color, grp-val, grp-left-only), and so does a
//   different right indent or left edge of the text - the left indent, or the first
//   line's where that is further left (grp-indent-*, grp-edge-*: hanging indents with
//   the first line at the same place group);
// - a between border, not part of what groups paragraphs, takes its space and width
//   below the paragraph that has it and its space above the one that has it (between3:
//   2 x space + width between; between-first-*, between-second-*, grp-between-off);
// - a line style's line: double three widths, triple five, thick-thin with a small gap the
//   width and 1.5pt, a medium gap two widths, a large gap the width and 2.25pt, thin-thick-
//   thin small the width and 3pt (line-*; dotted, dashed and outset one width). The wave
//   and 3D lines follow no such rule in three sizes: not checked (open in the ledger);
// - space after and the next paragraph's space before: the larger, not both (space-sum).
// Tables (table-border; Word 2010 compatibility mode, -c14, alike):
// - every horizontal border takes its width: the table's top above the first row, the
//   inside one between rows, the bottom below the last (all-sz*, insideH-*, topbottom-*);
//   the vertical ones nothing (sides-*); a double line three widths (double-sz6);
// - between two rows the widest of the borders there - the table's, a cell's own top
//   below or bottom above, in any column (cell-*, cell2col-*);
// - cell margins add to that (mar*, style-grid43*: a table style's alike);
// - an exact row height holds its border (exact-sz24: 30pt a row); an at-least height
//   is the content's, the border comes on top (atleast-sz24: 33pt a row).
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let fixtures =
    [ for family in [ "para-border"; "table-border" ] do
        let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures", family)
        let rows =
            File.ReadAllLines(Path.Combine(folder, "word/positions.tsv"))
            |> Array.filter (fun l -> l.Length > 0)
            |> Array.map (fun l -> let c = l.Split('\t') in c.[0], (c.[1], float c.[3]))
            |> Array.groupBy fst
        for name, labels in rows -> family + "/" + name, Path.Combine(folder, name + ".docx"), labels |> Array.map snd |> Array.sortBy (fun (label, y) -> y, label) ]

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText

let oursOf (file: string) =
    let doc = File.ReadAllBytes file |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    let layouts, _ = Layout.layoutDocumentPages settings measureText measureRange doc
    [ for layout in layouts do
        for line in layout.Lines do
            let text = line.Segments |> List.choose (function Layout.TextSegment (t, _, _, _, _) -> Some t | _ -> None) |> String.concat ""
            if text.Length = 3 then text, line.Y * 0.75 ]
    |> Map.ofList

let gaps (positions: (string * float) array) =
    positions |> Array.pairwise |> Array.map (fun ((a, ya), (b, yb)) -> sprintf "%s-%s" a b, yb - ya - 20.0)

// (Word's numbers for these are kept; the rule for them is not known yet.)
let notYet = [ "line-wave-"; "line-doubleWave-"; "line-threeDEmboss-" ]
let mutable pass = 0
let mutable fail = 0
for name, file, word in fixtures |> List.filter (fun (name, _, _) -> not (notYet |> List.exists (fun prefix -> name.Contains("/" + prefix)))) do
    let ours = oursOf file
    let oursPositions = word |> Array.map (fun (label, _) -> label, Map.tryFind label ours |> Option.defaultValue nan)
    let wrong =
        Array.zip (gaps word) (gaps oursPositions)
        |> Array.filter (fun ((_, w), (_, o)) -> not (abs (w - o) <= 0.3))
    if wrong.Length = 0 then pass <- pass + 1
    else
        fail <- fail + 1
        let show = wrong |> Array.map (fun ((gap, w), (_, o)) -> sprintf "%s Word %.2f ours %.2f" gap w o)
        printfn "FAIL %-36s %s" name (String.Join(", ", show))

printfn "Paragraph and table borders against Word: %d fixtures as Word, %d not" pass fail
if fail > 0 then exit 1
