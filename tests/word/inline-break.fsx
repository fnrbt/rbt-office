// A page or column break inside a paragraph's text (DocxReader's placeInlineBreaks, Layout's lineEndsInBreak),
// against Word 2019's PDFs of the fixtures tools/word-parity/fixtures/inline-break.py generates
// (tests/word/fixtures/inline-break; Word's lines in word/geometry.tsv - pdf-geometry.py). Compared, for the
// short probes (one line a part): the page each part of the probe paragraph and the first tail paragraph
// are on, in which column, and whether at the top of it. What Word showed (docs/layout-parity-flow.md,
// ledger): the text after the break starts the next page - after a column break in a section with columns
// the next column, in one column the next page - at its top (no space before, no first-line indent), Word
// 2010's layout and 2013's alike. And the break is written back where it was.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System.IO
open System.IO.Compression
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout

let folder = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/inline-break")
let rows = File.ReadAllLines(Path.Combine(folder, "word/geometry.tsv")) |> Array.map (fun l -> l.Split('\t'))
let probes = [ "Probe start"; "Probe rest"; "Probe third"; "Tail 1" ]

// (page, column - 0 or 1 -, at the top of the page) of each probe Word has.
let wordOf (name: string) =
    let lines = rows |> Array.filter (fun c -> c.[0] = name && c.[1] = "line")
    [ for probe in probes do
        let words = probe.Split(' ')
        match lines |> Array.tryFind (fun c -> c.[7] = words.[0] && c.[8] = words.[1]) with
        | Some c -> yield probe, (int c.[2], (if float c.[5] > 300.0 then 1 else 0), float c.[3] < 76.0)
        | None -> () ]

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let toPt (px: float) = px * 72.0 / 96.0

let oursOfDoc doc =
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let pageOf (y: float) = pages |> List.filter (fun p -> p.Top <= y + 0.01) |> List.length
    [ for layout in layouts do
        for line in layout.Lines do
            let text = line.Segments |> List.map (function Layout.TextSegment (t, _, _, _, _) -> t | _ -> "") |> String.concat "" |> fun s -> s.Trim()
            match probes |> List.tryFind (fun p -> text = p) with
            | Some probe ->
                let page = pageOf line.Y
                // (A segment's x is from the text's left: the second column 252pt in.)
                let x = (line.Segments |> List.tryPick (function Layout.TextSegment (_, _, x, _, _) -> Some x | _ -> None) |> Option.defaultValue 0.0)
                yield probe, (page, (if toPt x > 200.0 then 1 else 0), toPt (line.Y - pages.[page - 1].Top) < 72.5)
            | None -> () ]

let load name = File.ReadAllBytes(Path.Combine(folder, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc

let mutable pass = 0
let mutable fail = 0
let names = Directory.GetFiles(folder, "*.docx") |> Array.map Path.GetFileNameWithoutExtension |> Array.filter (fun n -> not (n.Contains "-long")) |> Array.sort
for name in names do
    let word = wordOf name
    let ours = oursOfDoc (load name)
    if word = ours then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-18s Word %A; ours %A" name word ours

// Written back where it was: the probe paragraph keeps its break between its parts (an edit in it, so the
// paragraph is written from the model), and the saved document lays out as before.
let roundTrip name =
    let doc = File.ReadAllBytes(Path.Combine(folder, name + ".docx")) |> Rbt.Word.Docx.loadBytesNative |> Option.get
    let probe = Rbt.Word.Document.paragraphsOfBlocks doc.Blocks |> List.find (fun p -> (toString p.Buffer).StartsWith "Probe start")
    let edited = doc |> Rbt.Word.Document.setCursor probe.Id 0 |> Rbt.Word.Document.insertText "X"
    let package = Rbt.Word.Docx.exportTypedDocumentPackage edited
    let bytes = DocumentWriter.writeTypedPackageBytes { package with Unchanged = false }
    use stream = new MemoryStream(bytes)
    use zip = new ZipArchive(stream, ZipArchiveMode.Read)
    use reader = new StreamReader(zip.GetEntry("word/document.xml").Open())
    let xml = reader.ReadToEnd()
    let start = xml.IndexOf "XProbe start"
    let rest = xml.IndexOf "Probe rest"
    let between = if start >= 0 && rest > start then xml.Substring(start, rest - start) else ""
    let after = if rest >= 0 then xml.Substring(rest, xml.IndexOf("</w:p>", rest) - rest) else ""
    let reloaded = DocxReader.loadBytesNative bytes |> Option.get |> Rbt.Word.Import.importDoc
    let placed = oursOfDoc reloaded |> List.map (fun (p, at) -> (if p = "Probe start" then "XProbe start" else p), at)
    between.Contains "<w:br w:type=\"column\"/>" || between.Contains "<w:br w:type=\"page\"/>"
    , not (after.Contains "<w:br")
    , placed |> List.filter (fun (p, _) -> p <> "XProbe start") = (oursOfDoc (load name) |> List.filter (fun (p, _) -> p <> "Probe start"))
for name in [ "column-c2"; "page-c1" ] do
    match roundTrip name with
    | true, true, true -> pass <- pass + 1
    | inPlace, notAtEnd, same ->
        fail <- fail + 1
        printfn "FAIL %-18s saved: the break between the parts %b, none after them %b, laid out as before %b" name inPlace notAtEnd same

printfn "Breaks inside a paragraph against Word: %d checks as Word, %d not" pass fail
if fail > 0 then exit 1
