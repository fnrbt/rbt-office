// A justified line's spaces narrowed for more text (Layout's LayoutSettings.SpaceShrink, lineLimit;
// justifyLineToEdge), against Word 2019's lines of the fixtures tools/word-parity/fixtures/justify-squeeze.py
// generates (tests/word/fixtures/justify-squeeze; Word's side, per probe whether its first line held all of
// its tokens, in word/squeeze.tsv). Every paragraph is a probe: n spaces between tokens of digits, the line
// narrowed by a right indent to the text's width less n x o; it fits when Word lets each space give up o.
// What Word showed (docs/layout-parity-flow.md, ledger):
// - in Word 2013's layout (compatibility mode 15) a justified line (both, distribute, thaiDistribute) fits
//   when each of its spaces gives up no more than a quarter of its width - whatever the font, its size, how
//   many spaces (1 to 9); its last line too; in a table's cell and in a column alike;
// - a word is not hyphenated when it fits so, and the part of a hyphenated word fits so too;
// - not in Word 2010's layout and before (14, 12, none), not left, centred or right-aligned.
// The fonts: Carlito, Liberation Mono / Serif / Sans for Calibri, Courier New, Times New Roman, Arial (their
// advances are those fonts').
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
let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/justify-squeeze")
let face family file : Shaper.FontFace =
    { Family = family; Bold = false; Italic = false; Data = File.ReadAllBytes(Path.Combine(root, "fonts", file)) }
let fontSet =
    Shaper.createFontSet "Calibri"
        [ face "Calibri" "Carlito-Regular.ttf"; face "Courier New" "LiberationMono-Regular.ttf"
          face "Times New Roman" "LiberationSerif-Regular.ttf"; face "Arial" "LiberationSans-Regular.ttf" ]
let measureText = Shaper.createMeasurer fontSet
let measureRange = Shaper.createRangeMeasurer fontSet
Rbt.Word.Hyphenation.register Rbt.Word.HyphenPatternsEnUs.patterns (Rbt.Word.HyphenPatternsEnUs.exceptions + " " + Rbt.Word.HyphenWordsEnUs.words)

// Word's: doc -> probe -> fits; and which documents' probes are their paragraphs' last lines.
let rows =
    File.ReadAllLines(Path.Combine(folder, "word/squeeze.tsv")) |> Array.skip 1 |> Array.filter (fun l -> l.Length > 0)
    |> Array.map (fun l -> let c = l.Split('\t') in c.[0], int c.[1], c.[3] = "1")
let lastDocs =
    File.ReadAllLines(Path.Combine(folder, "probes.tsv")) |> Array.skip 1 |> Array.filter (fun l -> l.Length > 0)
    |> Array.map (fun l -> l.Split('\t')) |> Array.filter (fun c -> c.[9] = "1") |> Array.map (fun c -> c.[0]) |> Set.ofArray

// Ours: per probe (in order) whether its first line held all its tokens - the line after it " end" alone,
// or for a last line no second line.
let oursOf (name: string) =
    let doc = File.ReadAllBytes(Path.Combine(folder, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, _ = Layout.layoutDocumentPages settings measureText measureRange doc
    // (A table's cells are among the layouts, a block each paragraph.)
    layouts
    |> List.map (fun layout ->
        [ for line in layout.Lines ->
            line.Segments |> List.choose (function Layout.TextSegment (t, _, _, _, _) -> Some t | _ -> None) |> String.concat "" ])
    |> List.filter (fun ls -> ls |> List.exists (fun t -> t.Trim() <> ""))
    |> List.map (fun ls ->
        if lastDocs.Contains name then ls.Length = 1
        // (hyph-part*: the first line kept "accessibil-" - its hyphen is drawn, not in the text.)
        elif name.StartsWith "hyph-part" then ls.Length > 1 && ls.[1].StartsWith "ity"
        else ls.Length > 1 && ls.[1].Trim() = "end")
    |> Array.ofList

let mutable pass = 0
let mutable fail = 0
for name, probes in rows |> Array.groupBy (fun (n, _, _) -> n) do
    let ours = oursOf name
    let wrong = [ for _, i, fits in probes do if i >= ours.Length || ours.[i] <> fits then yield i, fits ]
    pass <- pass + probes.Length - wrong.Length
    fail <- fail + wrong.Length
    if not wrong.IsEmpty then
        printfn "FAIL %s: %s" name (wrong |> List.map (fun (i, f) -> sprintf "%d (Word %s)" i (if f then "fits" else "goes on")) |> String.concat ", ")

printfn "Justified lines' spaces narrowed against Word: %d probes as Word, %d not" pass fail
if fail > 0 then exit 1
