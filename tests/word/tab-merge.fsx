// Which tab stops a paragraph has when its style has some and it states its own, against Word 2019's PDFs of
// the fixtures of tools/word-parity/fixtures/tab-merge.py (tests/word/fixtures/tab-merge/*.docx; every word's
// box in word/words.tsv): its style's stops (down the basedOn chain) with its own added, its own kind at a
// position both have, a w:val="clear" one taking the style's away. Checked: the line each word is on and
// where it starts (Carlito for Calibri: their advances are the same).
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
module Shaper = Rbt.Word.ShaperText
module Layout = Rbt.Word.Layout

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-50s expected %A got %A" label expected got
let near (tolerance: float) label (got: float) (expected: float) =
    if abs (got - expected) <= tolerance then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-50s expected %.2f got %.2f" label expected got

let root = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "../.."))
let dir = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/tab-merge")
let fontSet =
    Shaper.createFontSet "Calibri" [ { Family = "Calibri"; Bold = false; Italic = false; Data = File.ReadAllBytes(Path.Combine(root, "fonts", "Carlito-Regular.ttf")) } ]
let measure = Shaper.createMeasurer fontSet
let measureRange = Shaper.createRangeMeasurer fontSet

/// Word's words: (word, line index, left in points from the page's edge).
let word =
    File.ReadAllLines(Path.Combine(dir, "word", "words.tsv"))
    |> Array.map (fun l -> l.Split '\t')
    |> Array.groupBy (fun r -> r.[0])
    |> Array.map (fun (name, rows) ->
        let tops = rows |> Array.map (fun r -> Math.Round(float r.[4], 1)) |> Array.distinct |> Array.sort
        name, [ for r in rows -> r.[1], Array.IndexOf(tops, Math.Round(float r.[4], 1)), float r.[2] ] |> List.sortBy (fun (w, _, _) -> w))
    |> Map.ofArray

let ours name =
    let doc = File.ReadAllBytes(Path.Combine(dir, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    let layouts, _ = Layout.layoutDocumentPages settings measure measureRange doc
    let lines = [ for l in layouts do yield! l.Lines ]
    [ for index, line in List.indexed lines do
        for segment in line.Segments do
            match segment with
            | Layout.TextSegment (text, _, x, _, _) when text.Trim() <> "" -> yield text.Trim(), index, 72.0 + x * 0.75
            | _ -> () ]
    |> List.sortBy (fun (w, _, _) -> w)

// (para-only: a right stop past the right margin - Word breaks "Bbb" there; not ours.)
let fixtures = Directory.GetFiles(dir, "*.docx") |> Array.map Path.GetFileNameWithoutExtension |> Array.sort |> Array.filter (fun name -> name <> "para-only")
check "the fixtures are all there" fixtures.Length 8
for name in fixtures do
    let w = word.[name]
    let o = ours name
    check (sprintf "%s: the words and their lines" name) (o |> List.map (fun (t, line, _) -> t, line)) (w |> List.map (fun (t, line, _) -> t, line))
    if o.Length = w.Length then
        List.iter2 (fun (t, _, x) (_, _, wx) -> near 0.75 (sprintf "%s: %s starts" name t) x wx) o w

printfn "tab-merge: %d passed, %d failed" pass fail
if fail > 0 then exit 1
