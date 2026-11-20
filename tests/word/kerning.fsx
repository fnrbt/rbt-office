// Whether text is kerned (and its ligatures formed), against the widths Word 2019 gives
// words in its PDFs of the fixtures tools/word-parity/fixtures/kerning.py generates
// (tests/word/fixtures/kerning/word/widths.tsv). What Word showed (docs/layout-parity-flow.md):
// - without w:kern, or with w:kern above the text's size (kern-28 at 11pt), a word is the sum
//   of its glyphs' advances - no kerning ("AVAVAVAVAV" 63.04pt of Calibri 11pt), no
//   ligatures ("office"), in compatibility mode 14 and 15;
// - w:kern at or below the size kerns (kern-2, kern-28-16: 58.73pt), and w14:ligatures
//   fuses ("office" 0.25pt narrower) - which ours does not do yet (w:kern and
//   w14:ligatures are not read), so those fixtures are only listed.
// Ours: ShaperText's measuring with the faces of fonts/ that have the Office fonts'
// advances (Carlito, Liberation Sans, Liberation Serif, Caladea), kerning pairs and all.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Unicode.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Bidi.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.OpenType.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Shaper.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Text.Engine.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/ShaperText/bin/Debug/net10.0/Rbt.Word.ShaperText.dll"

open System
open System.IO
module Shaper = Rbt.Word.ShaperText
module TB = Rbt.Word.TextBuffer

let root = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "../.."))
let face family file : Shaper.FontFace =
    { Family = family; Bold = false; Italic = false; Data = File.ReadAllBytes(Path.Combine(root, "fonts", file)) }
let fontSet =
    Shaper.createFontSet "Calibri" [
        face "Calibri" "Carlito-Regular.ttf"; face "Arial" "LiberationSans-Regular.ttf"
        face "Times New Roman" "LiberationSerif-Regular.ttf"; face "Cambria" "Caladea-Regular.ttf" ]
let measure = Shaper.createMeasurer fontSet
let measureRange = Shaper.createRangeMeasurer fontSet

let rows =
    File.ReadAllLines(Path.Combine(__SOURCE_DIRECTORY__, "fixtures/kerning/word/widths.tsv"))
    |> Array.filter (fun l -> l.Length > 0)
    |> Array.map (fun l -> let c = l.Split('\t') in c.[0], c.[1], float c.[2], c.[3], float c.[4])

// (Word's widths are its PDF's: glyph positions in 1/600in, and Caladea's advances are not all Cambria's.)
let tolerance = 0.25
let unkerned = set [ "kern-none"; "kern-none-16"; "kern-28"; "kern-none-c14" ]
let mutable pass = 0
let mutable fail = 0
let mutable listed = 0
for fixture, font, size, word, width in rows do
    if font <> "Cambria" then
        let style = { TB.defaultStyle with FontFamily = Some font; FontSize = size }
        let ours = measure word style
        let ranged = measureRange word style 0 word.Length
        if unkerned.Contains fixture then
            if abs (ours - width) <= tolerance && abs (ranged - ours) < 1e-9 then pass <- pass + 1
            else
                fail <- fail + 1
                printfn "FAIL %-14s %-16s %4.0f %-12s Word %.3f ours %.3f (range %.3f)" fixture font size word width ours ranged
        else listed <- listed + 1

printfn "Kerning against Word: %d widths as Word, %d not (%d kerned or fused by Word, not checked)" pass fail listed
if fail > 0 then exit 1
