// A vertically merged cell's continuation cells (w:vMerge without restart) and the rows' heights, against
// Word 2019's own PDFs of the fixtures tools/word-parity/fixtures/vmerge-content.py generates
// (tests/word/fixtures/vmerge-content; Word's lines in word/geometry.tsv - pdf-geometry.py).
// What Word showed (docs/layout-parity-flow.md, ledger): a continuation cell takes no room - its empty
// paragraph, with Normal's 10pt after and 1.15 lines, makes its row no taller (empty, empty-center,
// empty-c14, two-merged; none-spacing alike), and Word neither shows nor makes room for text there
// (text, tall: "x21" ... not on the page, the rows one line high - open here: ours lays it out); what the
// merged cell's content needs beyond the rows it spans goes to the last (restart-tall, -empty).
// Every line of Word's (a01, the cells' "t12" ..., "m11" ..., c01) has one of ours at its place, and ours
// has no other.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let dir = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/vmerge-content")
let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText

// (pdftotext's box of a Calibri word starts 0.199 of its size below its line's top.)
let glyphTop (top: float) (bottom: float) = 0.1991 * (bottom - top) / 1.00364

/// Word's lines: (word, top of its line, left) by fixture.
let wordLines =
    File.ReadAllLines(Path.Combine(dir, "word/geometry.tsv")) |> Array.map (fun l -> l.Split '\t')
    |> Array.filter (fun r -> r.[1] = "line")
    |> Array.groupBy (fun r -> r.[0])
    |> Array.map (fun (name, rs) -> name, rs |> Array.map (fun r -> r.[7], float r.[3] - glyphTop (float r.[3]) (float r.[4]), float r.[5]) |> Array.toList)
    |> Map.ofArray

let ours name =
    let doc = File.ReadAllBytes(Path.Combine(dir, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    let layouts, _ = Layout.layoutDocumentPages settings measureText measureRange doc
    [ for layout in layouts do
        for line in layout.Lines do
            for segment in line.Segments do
                match segment with
                | Layout.TextSegment (text, _, x, _, _) when text.Trim() <> "" -> yield text.Trim(), line.Y * 0.75, (settings.PaddingLeft + x + Layout.alignmentOffset layout line) * 0.75
                | _ -> () ]

// (Ours lays out the text of a continuation cell - Word hides it; and the merged cell's extra height
// in its last row leaves out that row's bottom border, 0.5pt: restart-tall*.)
let open_ = set [ "text"; "tall"; "restart-tall"; "restart-tall-empty" ]
let mutable pass = 0
let mutable fail = 0
let fixtures = Directory.GetFiles(dir, "*.docx") |> Array.map Path.GetFileNameWithoutExtension |> Array.sort
if fixtures.Length <> 9 || fixtures |> Array.exists (fun name -> not (wordLines.ContainsKey name)) then
    printfn "FAIL the fixtures and Word's lines of each"
    fail <- fail + 1
for name in fixtures |> Array.filter (fun name -> not (open_.Contains name)) do
    let word = wordLines.[name]
    let ourLines = ours name
    let missing =
        word |> List.filter (fun (text, top, left) ->
            not (ourLines |> List.exists (fun (t, y, x) -> t = text && abs (y - top) <= 0.3 && abs (x - left) <= 0.6)))
    let extra = ourLines |> List.filter (fun (t, _, _) -> not (word |> List.exists (fun (text, _, _) -> text = t)))
    if missing.IsEmpty && extra.IsEmpty then pass <- pass + 1
    else
        fail <- fail + 1
        let show (t, y, x) = sprintf "%s %.2f %.2f" t y x
        let oursOf (t, _, _) = ourLines |> List.tryFind (fun (o, _, _) -> o = t) |> Option.map show |> Option.defaultValue "none"
        printfn "FAIL %-20s Word's not ours: %s | ours: %s | ours not Word's: %s" name
            (missing |> List.map show |> String.concat ", ") (missing |> List.map oursOf |> String.concat ", ") (extra |> List.map show |> String.concat ", ")

printfn "vmerge-content: %d fixtures as Word, %d not" pass fail
if fail > 0 then exit 1
