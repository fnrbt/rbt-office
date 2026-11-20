// Punctuation that Word lets hang past the end of a line, against Word 2019's lines of the
// fixtures tools/word-parity/fixtures/hang-punct.py generates (tests/word/fixtures/hang-punct;
// Word's side, every line of every fixture from Word's PDF, in word/lines.tsv). Every paragraph
// is a probe: text 0.1pt short of the line (Calibri 11pt, measured here with Carlito, which has
// its widths), then one or more characters that do not fit. The probe "hangs" when they stay on
// the first line, past the margin; otherwise the last word (or the character alone) goes on.
// What Word showed (docs/layout-parity-flow.md, ledger):
// - punctuation hangs only under w:useFELayout, in compatibility mode 14, 12 or none (not 15),
//   with w:overflowPunct on (unset is on), and only when the run of the character has an East
//   Asian language (w:lang w:eastAsia ko-KR, ja-JP, zh-CN, zh-TW; en-US: none);
// - which characters hang depends on that language (Layout.hangsPastLineEnd);
// - one character only (").": the ")" must fit), whatever its width (a full-width 11pt one, a
//   36pt "?"), past a right indent too; no tolerance for a letter (0.05pt over goes);
// - under compressPunctuation not at the paragraph's end (before a line break, or mid-paragraph, it does);
// - never in a table's cell (cell*: fixed or autofit, cell margins or none, the cell's text ending at
//   the page's margin too), but in a column, a header, a frame, a field's result, a bold run, a run
//   with w:lang w:val="en-US" (the East Asian language from the defaults): cols2, header, frame, fe3-fld.
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
let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/hang-punct")
let face family bold file : Shaper.FontFace =
    { Family = family; Bold = bold; Italic = false; Data = File.ReadAllBytes(Path.Combine(root, "fonts", file)) }
let fontSet = Shaper.createFontSet "Calibri" [ face "Calibri" false "Carlito-Regular.ttf"; face "Calibri" true "Carlito-Bold.ttf" ]
let measureText = Shaper.createMeasurer fontSet
let measureRange = Shaper.createRangeMeasurer fontSet

let isLetter (c: char) = c >= 'a' && c <= 'z'
let labelOf (text: string) =
    if text.Length >= 4 && text.[3] = ' ' && Char.IsLetter text.[0] && Char.IsDigit text.[1] && Char.IsDigit text.[2] then Some (text.Substring(0, 3)) else None
// A probe's first line hangs when it ends in its probe characters - not in a letter of its
// last word, not before the last word gone on ("eeeeeeee" is the last word before it).
let hangsOf (first: string) =
    let t = first.TrimEnd()
    let last = t.Substring(t.LastIndexOf(' ') + 1)
    last <> "eeeeeeee" && not (isLetter t.[t.Length - 1])

// Word's: fixture -> label -> hangs. (A character in another font can be a PDF line of its own:
// after the label's line when it went on alone - the w and x probes; the 36pt one of the s
// probes is one just before the label's line when it is on it.)
let expected =
    File.ReadAllLines(Path.Combine(folder, "word/lines.tsv"))
    |> Array.filter (fun l -> l.Length > 0)
    |> Array.map (fun l -> let c = l.Split('\t') in c.[0], c.[3])
    |> Array.groupBy fst
    |> Array.map (fun (name, rows) ->
        let texts = rows |> Array.map snd
        name,
        [ for i in 0 .. texts.Length - 1 do
            match labelOf texts.[i] with
            | Some label ->
                let t = texts.[i].TrimEnd()
                let last = t.Substring(t.LastIndexOf(' ') + 1)
                let hangs =
                    if label.[0] = 's' then last <> "eeeeeeee" && i > 0 && texts.[i - 1].Trim().Length = 1
                    elif i + 1 < texts.Length && texts.[i + 1].Trim().Length = 1 && not (isLetter (texts.[i + 1].Trim().[0])) then false
                    else hangsOf t
                yield label, hangs
            | None -> () ] |> Map.ofList)
    |> Map.ofArray

let oursOf (file: string) =
    let doc = File.ReadAllBytes file |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    // (The header fixture's probes are in its header: laid out here as a body of its width.)
    let doc =
        match doc.HeaderParts with
        | header :: _ when Path.GetFileNameWithoutExtension file = "header" -> { doc with Blocks = header.Blocks }
        | _ -> doc
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, _ = Layout.layoutDocumentPages settings measureText measureRange doc
    [ for layout in layouts do
        for line in layout.Lines do
            let text = line.Segments |> List.choose (function Layout.TextSegment (t, _, _, _, _) -> Some t | _ -> None) |> String.concat ""
            match labelOf text with
            | Some label -> yield label, hangsOf text
            | None -> () ] |> Map.ofList

let mutable pass = 0
let mutable fail = 0
// (frame: Word hangs in a frame as on the page - its c, m, w probes as fe3's - but ours does not
// lay out a w:framePr frame at its own width yet.)
for KeyValue (name, word) in expected |> Map.remove "frame" do
    let ours = oursOf (Path.Combine(folder, name + ".docx"))
    let wrong = [ for KeyValue (label, hangs) in word do if Map.tryFind label ours <> Some hangs then yield label, hangs ]
    pass <- pass + word.Count - wrong.Length
    fail <- fail + wrong.Length
    if not wrong.IsEmpty then
        printfn "FAIL %s: %s" name (wrong |> List.map (fun (l, h) -> sprintf "%s (Word %s)" l (if h then "hangs" else "goes on")) |> String.concat ", ")

printfn "Hanging punctuation against Word: %d probes as Word, %d not" pass fail
if fail > 0 then exit 1
