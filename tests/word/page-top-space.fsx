// A paragraph's space before at the top of a page or a column, against Word 2019's PDFs of the
// fixtures tools/word-parity/fixtures/page-top-space.py generates (tests/word/fixtures/page-top-space;
// Word's side, the top of every probe word "P-...", in word/tops.tsv). Every probe is a Calibri 11pt
// paragraph with 12pt before; it is at the top margin (its space gone) or 12pt below it.
// What Word showed (docs/layout-parity-flow.md, ledger), in compatibility mode 14 and 15 alike unless said:
// - gone when the paragraph comes to the page by the text's flow: pushed down by the lines above it
//   (flow, a column's too), after a page break in a paragraph before it (one of its own, as
//   docx4j-headers-footers' "pgMar (Page Margins)", or at the end of one's text: brkpara, brkend);
// - kept on the document's first paragraph, after a next-page section break, and after a page break
//   at the start of the paragraph's own text (brkin) - unless w:suppressSpBfAfterPgBrk, in mode 14;
// - w:pageBreakBefore: kept in mode 14, gone in 15;
// - a column break in a paragraph of its own: its paragraph mark goes on to the next column as a line.
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
let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/page-top-space")
let face family bold file : Shaper.FontFace =
    { Family = family; Bold = bold; Italic = false; Data = File.ReadAllBytes(Path.Combine(root, "fonts", file)) }
let fontSet = Shaper.createFontSet "Calibri" [ face "Calibri" false "Carlito-Regular.ttf"; face "Calibri" true "Carlito-Bold.ttf" ]
let measureText = Shaper.createMeasurer fontSet
let measureRange = Shaper.createRangeMeasurer fontSet

// (A Calibri 11pt word's top in Word's PDF is 2.16pt below its line's top: 74.16 for a line at the 72pt margin.)
let wordAscentGap = 2.16

// Word's: fixture -> probe -> (page, the line's top below the margin, pt).
let expected =
    File.ReadAllLines(Path.Combine(folder, "word/tops.tsv"))
    |> Array.filter (fun l -> l.Length > 0)
    |> Array.map (fun l -> l.Split('\t'))
    |> Array.filter (fun c -> c.[4].StartsWith "P-")
    |> Array.groupBy (fun c -> c.[0])
    |> Array.map (fun (name, rows) -> name, rows |> Array.map (fun c -> c.[4], (int c.[1], float c.[3] - wordAscentGap - 72.0)) |> Map.ofArray)
    |> Map.ofArray

let oursOf (file: string) =
    let doc = File.ReadAllBytes file |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    [ for layout in layouts do
        for line in layout.Lines do
            match line.Segments |> List.tryPick (function Layout.TextSegment (t, _, _, _, _) when t.StartsWith "P-" -> Some t | _ -> None) with
            | Some text ->
                let page = pageOf line.Y
                let top = (line.Y - (pages |> List.item (page - 1)).Top) * 0.75 - 72.0
                yield text.Split(' ').[0], (page, top)
            | None -> () ] |> Map.ofList

// What ours does not do yet (docs/layout-parity-flow.md, ledger: open):
// - a page break inside a paragraph's text: the reader makes it a break after the paragraph (brkmid);
// - w:suppressSpBfAfterPgBrk is not read (pages-c14-sup's brkin);
// - a column break in a paragraph of its own: Word puts its paragraph mark at the next column's top.
let known =
    set [ for d in [ "pages-c14"; "pages-c14-sup"; "pages-c15"; "pages-c15-sup" ] -> d, "P-brkmid"
          for d in [ "sect-c14"; "sect-c15" ] -> d, "P-s-brkmid"
          yield "pages-c14-sup", "P-brkin"
          for d in [ "cols-brkpara-c14"; "cols-brkpara-c15" ] -> d, "P-col-brkpara" ]

let mutable pass = 0
let mutable fail = 0
for KeyValue (name, word) in expected do
    let ours = oursOf (Path.Combine(folder, name + ".docx"))
    for KeyValue (probe, (page, top)) in word do
        if known.Contains (name, probe) then () else
        match Map.tryFind probe ours with
        | Some (p, t) when p = page && abs (t - top) < 0.5 -> pass <- pass + 1
        | o ->
            fail <- fail + 1
            printfn "FAIL %s %s: Word page %d, %.2fpt below the margin; ours %s" name probe page top
                (match o with Some (p, t) -> sprintf "page %d, %.2fpt" p t | None -> "none")

printfn "Space before at a page's top against Word: %d probes as Word, %d not (%d known, not done)" pass fail known.Count
if fail > 0 then exit 1
