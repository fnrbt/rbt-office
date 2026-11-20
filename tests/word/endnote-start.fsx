// Where the endnotes start at the end of the document (Layout's endnoteLayoutsAfterBody, noteSeparatorOf),
// against Word 2019's PDFs of the fixtures tools/word-parity/fixtures/endnote-start.py generates
// (tests/word/fixtures/endnote-start; Word's lines in word/geometry.tsv - pdf-geometry.py --paths). Compared:
// the page the note is on, and its first line's top below the last body line's top (on the page it starts
// below the text) or below the page's top. What Word showed (docs/layout-parity-flow.md, ledger):
// - the notes start below the last paragraph - its space after too, and an empty paragraph after it - by
//   endnotes.xml's separator, a paragraph of its own as high as its line (13.43pt in Calibri 11pt, 26.86
//   double-spaced, 24.41 at 20pt), not its space before;
// - when the separator and the note's lines do not all fit under the text (a 3-line note with room for 2 of
//   them too: widow control), they start the next page, the separator atop it.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/endnote-start")
let rows = File.ReadAllLines(Path.Combine(folder, "word/geometry.tsv")) |> Array.map (fun l -> l.Split('\t'))

// Word's glyph tops are the line's top plus 0.198 of the size in Calibri (the body 11pt, the note 10pt).
let wordOf (name: string) =
    let lines = rows |> Array.filter (fun c -> c.[0] = name && c.[1] = "line")
    let note = lines |> Array.find (fun c -> c.[7] = "Note")
    let notePage = int note.[2]
    let noteTop = float note.[3] - 1.98
    let body = lines |> Array.filter (fun c -> c.[7] = "Body" && int c.[2] = notePage)
    let below = if body.Length = 0 then noteTop - 72.0 else noteTop - (float body.[body.Length - 1].[3] - 2.18)
    notePage, below

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let toPt (px: float) = px * 72.0 / 96.0

let oursOf (name: string) =
    let doc = File.ReadAllBytes(Path.Combine(folder, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    let note = layouts |> List.find (fun l -> l.BlockId.StartsWith Layout.endnoteBlockIdPrefix)
    let noteTop = note.Lines.Head.Y
    let notePage = pageOf noteTop
    let bodyLines =
        layouts |> List.filter (fun l -> not (l.BlockId.StartsWith Layout.endnoteBlockIdPrefix))
        |> List.collect (fun l -> l.Lines) |> List.filter (fun l -> pageOf l.Y = notePage && not l.Segments.IsEmpty)
    let below =
        match List.tryLast bodyLines with
        | Some last -> toPt (noteTop - last.Y)
        | None -> toPt (noteTop - (pages.[notePage - 1].Top + settings.PaddingTop))
    notePage, below

let mutable pass = 0
let mutable fail = 0
for name in rows |> Array.map (fun c -> c.[0]) |> Array.distinct |> Array.sort do
    let wordPage, wordBelow = wordOf name
    let ourPage, ourBelow = oursOf name
    if wordPage = ourPage && abs (wordBelow - ourBelow) < 0.3 then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-16s Word: page %d, %.2fpt below; ours page %d, %.2fpt" name wordPage wordBelow ourPage ourBelow

printfn "Endnotes' start against Word: %d documents as Word, %d not" pass fail
if fail > 0 then exit 1
