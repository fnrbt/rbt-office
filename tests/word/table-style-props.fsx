// The paragraphs in the cells of a table with a table style, as Word 2019 sets them, against Word's
// own results for the fixtures of tools/word-parity/fixtures/table-style-props.py
// (tests/word/fixtures/table-style-props/*.docx; Word's pages and PDFs - page-parity.mjs --dir
// --refresh-word -, and in word/geometry.tsv where its PDFs put each line's text, across and down -
// tools/word-parity/pdf-geometry.py). Each cell paragraph ("a01 ..." / "b01 ..." in row 1, "d01 ..." /
// "e01 ..." in row 2) has no properties of its own; Word's font size of it (its glyphs' height), its
// first-line indent (its first line's start against its second's), its justification (where its lines
// start and end), its line pitch and the room after it are compared with ours:
// - the style hierarchy: document defaults, the table style (its conditional formats over the whole
//   table's), the paragraph's style (Normal) - Normal's properties beat the table style's (first-line
//   indent, space after, line spacing, justification, size, face; bold is a toggle: Normal's b 0 does
//   not cancel the table style's b);
// - except in Word 2010's layout and before (compatibilityMode 14 or less, or none) unless the
//   document's overrideTableStyleFontSizeAndJustification is on, where Normal's 12pt and Normal's left
//   justification (Normal's own, inherited by the paragraph's style or not) give way to the table
//   style's size and justification for the cell (a conditional format's, the whole table's, or else
//   the document defaults'; a size of 10pt is no size) - in a table whose style has paragraph or run
//   properties somewhere, or in a cell a conditional format of it applies to.
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

let dir = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/table-style-props")
let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let pt (px: float) = px * 0.75

type Line = { Top: float; Bottom: float; Left: float; Right: float; First: string }

/// Word's lines of each cell paragraph of a fixture ('a', 'b', 'd', 'e').
let wordParagraphs =
    File.ReadAllLines(Path.Combine(dir, "word/geometry.tsv")) |> Array.map (fun l -> l.Split '\t')
    |> Array.filter (fun r -> r.[1] = "line" && "abde".Contains r.[7].[0])
    |> Array.groupBy (fun r -> r.[0])
    |> Array.map (fun (name, rs) ->
        name,
        rs |> Array.map (fun r -> { Top = float r.[3]; Bottom = float r.[4]; Left = float r.[5]; Right = float r.[6]; First = r.[7] })
        |> Array.groupBy (fun line -> line.First.[0]) |> Map.ofArray)
    |> Map.ofArray

/// What Word gave a paragraph: its size (pdftotext's box of a word is 1.0036 of a Calibri size, 1.111
/// of a Times New Roman one), first-line indent, justification and line pitch, in points.
type Props = { Size: float; FirstLine: float; Align: string; Pitch: float }

let wordProps name (lines: Line[]) =
    let times = (name: string).Contains "times"
    let size = Math.Round((lines.[0].Bottom - lines.[0].Top) / (if times then 1.111 else 1.0036) * 2.0) / 2.0
    let cellLeft = lines |> Array.map (fun l -> l.Left) |> Array.min
    let last = lines.[lines.Length - 1]
    // The cell's text runs 77.66 - 354.94 (72.02 - 349.30 in Word 2010's layout: hanging out by its margin).
    let textLeft, textRight = if name.EndsWith "-c14" || name.EndsWith "-c12" || name.EndsWith "-cnone" then 72.02, 349.30 else 77.66, 354.94
    let align =
        if abs (last.Right - textRight) < 0.5 && last.Left > textLeft + 2.0 then "right"
        elif abs ((last.Left + last.Right) / 2.0 - (textLeft + textRight) / 2.0) < 1.0 && last.Left > textLeft + 2.0 then "center"
        // (Justified: every line but the last reaches the cell's right edge - Word's end up to 0.35pt short.)
        elif lines.Length > 1 && lines |> Array.take (lines.Length - 1) |> Array.forall (fun l -> l.Right > textRight - 0.4) then "both"
        else "left"
    { Size = size; FirstLine = (if lines.Length > 1 then Math.Round(lines.[0].Left - lines.[1].Left, 1) else 0.0)
      Align = align; Pitch = (if lines.Length > 1 then lines.[1].Bottom - lines.[0].Bottom else 0.0) }

/// Ours: each cell paragraph's lines - its first segment's style, where each line starts and its top.
let ours name =
    let doc = File.ReadAllBytes(Path.Combine(dir, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    let layouts, _ = Layout.layoutDocumentPages settings measureText measureRange doc
    [ for layout in layouts do
        let lines =
            layout.Lines |> List.choose (fun line ->
                line.Segments |> List.tryPick (function
                    | Layout.TextSegment (text, style, x, _, _) when text.Trim() <> "" -> Some (text.Trim(), style, x + Layout.alignmentOffset layout line, line.Y)
                    | _ -> None))
        match lines with
        | (text, _, _, _) :: _ when "abde".Contains text.[0] && text.Length >= 3 && Char.IsDigit text.[1] -> yield text.[0], (layout, lines)
        | _ -> () ] |> Map.ofList

let fixtures = Directory.GetFiles(dir, "*.docx") |> Array.map Path.GetFileNameWithoutExtension |> Array.sort |> Array.toList
check "the fixtures are all there" fixtures.Length 103
check "Word's results for each" (fixtures |> List.filter (fun name -> not (wordParagraphs.ContainsKey name))) []

for name in fixtures do
    let paras = ours name
    for KeyValue (key, lines) in wordParagraphs.[name] do
        let w = wordProps name lines
        match Map.tryFind key paras with
        | None -> check (sprintf "%s %c: a paragraph of ours" name key) false true
        | Some (layout, ourLines) ->
            let _, style, x0, y0 = ourLines.Head
            check (sprintf "%s %c: size" name key) (pt style.FontSize) w.Size
            let align =
                match layout.Metadata.Alignment with
                | Rbt.Word.Document.Left -> "left" | Rbt.Word.Document.Center -> "center" | Rbt.Word.Document.Right -> "right" | Rbt.Word.Document.Justify -> "both"
            check (sprintf "%s %c: justification" name key) align w.Align
            match ourLines with
            | _ :: (_, _, x1, y1) :: _ ->
                check (sprintf "%s %c: first-line indent" name key) (Math.Round(pt (x0 - x1), 1)) w.FirstLine
                let pitch = pt (y1 - y0)
                if abs (pitch - w.Pitch) > 0.3 then check (sprintf "%s %c: line pitch" name key) (Math.Round(pitch, 2)) w.Pitch
            | _ -> ()
            // From this paragraph's last line to the next one's first in the cell (a to b, d to e).
            let next = char (int key + 1)
            match Map.tryFind next wordParagraphs.[name], Map.tryFind next paras with
            | Some nextLines, Some (_, (_, _, _, yNext) :: _) when key = 'a' || key = 'd' ->
                let wordGap = nextLines.[0].Bottom - lines.[lines.Length - 1].Bottom
                let _, _, _, yLast = List.last ourLines
                let gap = pt (yNext - yLast)
                if abs (gap - wordGap) > 0.4 then check (sprintf "%s %c: to the next paragraph" name key) (Math.Round(gap, 2)) wordGap
            | _ -> ()
            if Environment.GetEnvironmentVariable "VERBOSE" = "1" then
                printfn "%s %c: ours size %.1f bold %b family %A align %A | word size %.1f fl %.1f %s pitch %.2f"
                    name key (pt style.FontSize) style.Bold style.FontFamily align w.Size w.FirstLine w.Align w.Pitch

printfn "table-style-props: %d passed, %d failed" pass fail
if fail > 0 then exit 1
