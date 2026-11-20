// Where Word breaks a line inside a word - after a hyphen, a dash, a soft hyphen - against Word
// 2019's lines of the fixtures tools/word-parity/fixtures/hyphen-break.py generates
// (tests/word/fixtures/hyphen-break; Word's side, every line of every fixture from Word's PDF, in
// word/lines.tsv). Every paragraph is a probe: a label and 'a's (40 characters), a space, a token
// that does not fit after them, " end" (Courier New 10pt, 60 characters a line). What is compared:
// how many of the token's characters Word and ours put on the first line.
// What Word showed (docs/layout-parity-flow.md, ledger; Layout.breaksAfterInWord):
// - a break after "-", U+2010, U+2012, U+2013, U+2014 and w:softHyphen, wherever in the word: between
//   letters, digits, in web addresses ("http://www.ecma-" / "international.org/news/"), paths, at the
//   word's start; never before them, and the character must fit;
// - none after U+2011, w:noBreakHyphen, U+2015, U+2212, / \ . , : ; _ ? ! & = + % ( ) ~ | * # @ ·;
// - alike in compatibility modes 14 and 15, under East Asian layout, in a w:hyperlink and a
//   HYPERLINK field's result;
// - a soft hyphen not broken at is not seen and takes no width; broken at, it is a hyphen.
// Left out: a06 (a literal U+00AD in w:t: Word draws a hyphen, 6pt, and does not break after it;
// ours reads it as w:softHyphen).
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
let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/hyphen-break")
let face family bold file : Shaper.FontFace =
    { Family = family; Bold = bold; Italic = false; Data = File.ReadAllBytes(Path.Combine(root, "fonts", file)) }
let fontSet =
    Shaper.createFontSet "Courier New" [
        face "Courier New" false "LiberationMono-Regular.ttf"; face "Courier New" true "LiberationMono-Bold.ttf" ]
let measureText = Shaper.createMeasurer fontSet
let measureRange = Shaper.createRangeMeasurer fontSet

let leftOut = set [ "a06" ]

/// A probe's label and how many of its token's characters its first line holds.
let probeOf (text: string) =
    if text.Length >= 4 && Char.IsLetter text.[0] && Char.IsDigit text.[1] && Char.IsDigit text.[2] && text.[3] = 'a' then
        let i = text.IndexOf ' '
        Some (text.Substring(0, 3), if i < 0 then 0 else text.Length - i - 1)
    else None

let expected =
    File.ReadAllLines(Path.Combine(folder, "word/lines.tsv"))
    |> Array.filter (fun l -> l.Length > 0)
    |> Array.map (fun l -> let c = l.Split('\t') in c.[0], c.[3].Replace("&quot;", "\"").Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">").TrimEnd())
    |> Array.groupBy fst
    |> Array.map (fun (name, rows) -> name, rows |> Array.choose (snd >> probeOf) |> Map.ofArray)
    |> Map.ofArray

let oursOf (file: string) =
    let doc = File.ReadAllBytes file |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, _ = Layout.layoutDocumentPages settings measureText measureRange doc
    [ for layout in layouts do
        match layout.Lines with
        | first :: _ ->
            let text = first.Segments |> List.choose (function Layout.TextSegment (t, _, _, _, _) -> Some t | _ -> None) |> String.concat ""
            let text = text.TrimEnd()
            // (A soft hyphen is seen only where the line breaks at it.)
            let text = if text.EndsWith("\u00AD", StringComparison.Ordinal) then text.Replace("\u00AD", "") + "-" else text.Replace("\u00AD", "")
            match probeOf text with
            | Some probe -> yield probe
            | None -> ()
        | [] -> () ] |> Map.ofList

let mutable pass = 0
let mutable fail = 0
for KeyValue (name, word) in expected do
    let ours = oursOf (Path.Combine(folder, name + ".docx"))
    let wrong = [ for KeyValue (label, n) in word do if not (leftOut.Contains label) && Map.tryFind label ours <> Some n then yield label, n ]
    pass <- pass + (word |> Map.filter (fun l _ -> not (leftOut.Contains l))).Count - wrong.Length
    fail <- fail + wrong.Length
    if not wrong.IsEmpty then
        printfn "FAIL %s: %s" name (wrong |> List.map (fun (l, n) -> sprintf "%s (Word %d, ours %A)" l n (Map.tryFind l ours)) |> String.concat ", ")

printfn "Breaks inside words against Word: %d probes as Word, %d not" pass fail
if fail > 0 then exit 1
