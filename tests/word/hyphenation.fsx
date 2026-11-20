// Automatic hyphenation (w:autoHyphenation; Hyphenation.fs, Layout's hyphenationBreak) against Word
// 2019's lines of the fixtures tools/word-parity/fixtures/hyphen-rules.py and hyphen-words.py make
// (tests/word/fixtures/hyphen-rules: Word's lines in word/lines.tsv; hyphen-words, hyphen-words-c14:
// Word's break of every probe in word/points.tsv). Courier New 10pt / 6pt, 60 characters a line.
// What Word showed (docs/layout-parity-flow.md, ledger):
// - a word is broken at the last of its points where it and the hyphen fit (all 37167 probes);
// - Word 2013's layout (mode 15): no hyphenation zone; at least 2 letters before a break, 3 after;
//   Word 2010's (mode 14): the zone (360 twips when none is stated: a word whose start is less than it
//   from the line's end is not broken), 2 letters after a break, and other points;
// - w:consecutiveHyphenLimit n: at most n lines in a row end in a hyphen;
// - left and justified text alike, trailing punctuation and a leading "(" aside, after a w:br too.
// Ours uses Liang's patterns (hyph-en-us), which are not Word's: the probes are checked for the
// agreement measured when they were chosen (mode 15: 90.6% of the probes; mode 14: 88.6%); the rule
// fixtures line for line with Word's own points in place of the patterns (its points of the word lists,
// mode 14's for the mode-14 documents), so that what they check is the rules.
// Left out, as ours does otherwise on purpose: words with digits (dig*), all capitals (caps*: Word breaks
// them unless w:doNotHyphenateCaps); not yet read: w:noProof, w:lang (nop, de, fr), w:suppressAutoHyphens
// (sup, supstyle); a justified line's spaces narrowed in mode 15 (flow-m15-both-*), mode 14's points
// (flow-m14-*).
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
let fixtures = Path.Combine(__SOURCE_DIRECTORY__, "fixtures")
let face family bold file : Shaper.FontFace =
    { Family = family; Bold = bold; Italic = false; Data = File.ReadAllBytes(Path.Combine(root, "fonts", file)) }
let fontSet =
    Shaper.createFontSet "Courier New" [
        face "Courier New" false "LiberationMono-Regular.ttf"; face "Courier New" true "LiberationMono-Bold.ttf" ]
let measureText = Shaper.createMeasurer fontSet
let measureRange = Shaper.createRangeMeasurer fontSet

let docOf (file: string) = File.ReadAllBytes file |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc

/// Our lines of a document's body: their text, a hyphen after a hyphenated one.
let linesOf (doc: Rbt.Word.Document.Document) =
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, _ = Layout.layoutDocumentPages settings measureText measureRange doc
    [ for layout in layouts do
        for line in layout.Lines do
            let text = line.Segments |> List.choose (function Layout.TextSegment (t, _, _, _, _) -> Some t | _ -> None) |> String.concat ""
            let text = String.Join(" ", text.Split([| ' ' |], StringSplitOptions.RemoveEmptyEntries))
            if text <> "" then yield text + (if line.Hyphenated then "-" else "") ]

let unescape (t: string) = t.Replace("&apos;", "'").Replace("&quot;", "\"").Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">").TrimEnd()

let mutable failures = 0
let fail (message: string) =
    failures <- failures + 1
    printfn "FAIL %s" message

// The rule fixtures: Word's lines and ours, document by document.
let rulesFolder = Path.Combine(fixtures, "hyphen-rules")
let wordLines =
    File.ReadAllLines(Path.Combine(rulesFolder, "word/lines.tsv"))
    |> Array.filter (fun l -> l.Length > 0)
    |> Array.map (fun l -> let c = l.Split('\t') in c.[0], unescape c.[3])
    |> Array.groupBy fst
    |> Array.map (fun (name, rows) -> name, rows |> Array.map snd |> Array.toList)
    |> Map.ofArray
// (A probe's lines: its label's, and the one after.)
let leftOutProbes = set [ "dig1"; "dig2"; "dig3"; "caps12"; "caps06"; "nop"; "de"; "fr"; "sup"; "supstyle" ]
let labelOf (line: string) =
    let token = line.Split(' ').[0]
    token.TrimEnd('a')
let checkedDocs =
    wordLines |> Map.toList |> List.map fst
    |> List.filter (fun n -> not (n.StartsWith "flow-m14") && not (n.StartsWith "flow-m15-both") && not (n.StartsWith "squeeze") && not (n.StartsWith "bottom") && not (n.StartsWith "colbottom") && n <> "caps-dnh" && n <> "none" && not (n.StartsWith "base-m14"))
/// Word's points of a word list, as exceptions the hyphenator takes (no patterns).
let wordPointsAsExceptions (folderName: string) =
    File.ReadAllLines(Path.Combine(fixtures, folderName, "word/points.tsv"))
    |> Array.filter (fun l -> l.StartsWith "word\t")
    |> Array.map (fun l ->
        let c = l.Split('\t')
        let word = c.[1]
        let points = if c.Length > 2 && c.[2] <> "" then c.[2].Split(',') |> Array.map int else [||]
        let mutable last = 0
        let parts = ResizeArray<string>()
        for p in points do
            parts.Add(word.Substring(last, p - last))
            last <- p
        parts.Add(word.Substring last)
        String.Join("-", parts))
    |> String.concat " "
let wordPoints15 = wordPointsAsExceptions "hyphen-words"
let wordPoints14 = wordPointsAsExceptions "hyphen-words-c14"
let mutable linesSame = 0
for name in checkedDocs do
    Rbt.Word.Hyphenation.register "" (if name.Contains "-m14" then wordPoints14 else wordPoints15)
    let word = wordLines.[name]
    let ours = linesOf (docOf (Path.Combine(rulesFolder, name + ".docx")))
    // Drop the left-out probes' two lines from both.
    let drop (lines: string list) =
        let arr = List.toArray lines
        [ let mutable skip = 0
          for line in arr do
            if skip > 0 then skip <- skip - 1
            elif leftOutProbes.Contains (labelOf line) then skip <- 1
            else yield line ]
    let word, ours = drop word, drop ours
    if word = ours then linesSame <- linesSame + word.Length
    else
        let i = List.zip (List.truncate (min word.Length ours.Length) word) (List.truncate (min word.Length ours.Length) ours) |> List.tryFindIndex (fun (a, b) -> a <> b)
        match i with
        | Some i -> fail (sprintf "%s line %d: Word %A, ours %A" name (i + 1) word.[i] ours.[i])
        | None -> fail (sprintf "%s: Word %d lines, ours %d" name word.Length ours.Length)
printfn "rule fixtures: %d documents, %d lines as Word's" checkedDocs.Length linesSame

// A hyphenated line at a page's or column's foot (Word's own points again): Word 2013's layout moves it
// on (bottom-m15-last: the probe's first line atop page 2; colbottom-m15: atop the second column), Word
// 2010's keeps it (bottom-m14-last: the page's last line).
let probeLine (name: string) (label: string) =
    let doc = docOf (Path.Combine(rulesFolder, name + ".docx"))
    let settings = Layout.documentSettings 16.0 24.0 0.25 doc
    let layouts, _ = Layout.layoutDocumentPages settings measureText measureRange doc
    [ for layout in layouts do
        for line in layout.Lines do
            match line.Segments with
            | Layout.TextSegment (t, _, x, _, _) :: _ when t.StartsWith label ->
                yield int (floor (line.Y / (settings.PageHeight + Layout.pageGap))), layout.Indent + x, line.Hyphenated
            | _ -> () ] |> List.head
for name, label, page, column in [ "bottom-m15-last", "p32", 1, 0; "bottom-m15-allow-last", "p32", 1, 0; "bottom-m14-last", "p32", 0, 0
                                   "bottom-m14-allow-last", "p32", 0, 0; "bottom-m15", "p31", 0, 0; "colbottom-m15", "p23", 0, 1 ] do
    let ourPage, x, hyphenated = probeLine name label
    let ourColumn = if x > 400.0 then 1 else 0
    if (ourPage, ourColumn, hyphenated) <> (page, column, true) then
        fail (sprintf "%s: %s on page %d column %d (hyphenated %b), Word page %d column %d" name label (ourPage + 1) (ourColumn + 1) hyphenated (page + 1) (column + 1))
Rbt.Word.Hyphenation.register Rbt.Word.HyphenPatternsEnUs.patterns Rbt.Word.HyphenPatternsEnUs.exceptions

// The word lists: how many probes break where Word's do.
let probeAgreement (folderName: string) =
    let folder = Path.Combine(fixtures, folderName)
    let wordBreaks =
        File.ReadAllLines(Path.Combine(folder, "word/points.tsv"))
        |> Array.filter (fun l -> l.StartsWith "probe\t")
        |> Array.map (fun l -> let c = l.Split('\t') in int c.[1], int c.[4])
        |> Map.ofArray
    let mutable same = 0
    let mutable total = 0
    let watch = Diagnostics.Stopwatch.StartNew()
    for file in Directory.GetFiles(folder, "w*.docx") |> Array.sort do
        let lines = linesOf (docOf file) |> List.toArray
        // (A probe's first line starts with its zero-padded number.)
        for i in 0 .. lines.Length - 1 do
            let parts = lines.[i].Split(' ')
            if parts.[0].Length >= 12 && parts.[0] |> Seq.forall Char.IsDigit then
                let number = int parts.[0]
                let ours = if parts.Length > 1 && parts.[1].EndsWith "-" then parts.[1].Length - 1 else 0
                match Map.tryFind number wordBreaks with
                | Some word ->
                    total <- total + 1
                    if word = ours then same <- same + 1
                | None -> ()
    printfn "%s: %d of %d probes break as Word's (%.1f%%), laid out in %d ms" folderName same total (100.0 * float same / float (max 1 total)) watch.ElapsedMilliseconds
    float same / float (max 1 total), total, wordBreaks.Count

let m15, n15, w15 = probeAgreement "hyphen-words"
if n15 <> w15 then fail (sprintf "hyphen-words: %d probes read of %d" n15 w15)
if m15 < 0.905 then fail (sprintf "hyphen-words: agreement %.3f below 0.905" m15)
let m14, n14, w14 = probeAgreement "hyphen-words-c14"
if n14 <> w14 then fail (sprintf "hyphen-words-c14: %d probes read of %d" n14 w14)
if m14 < 0.885 then fail (sprintf "hyphen-words-c14: agreement %.3f below 0.885" m14)

// What hyphenation costs: the flow fixture laid out with and without it.
let flow = docOf (Path.Combine(rulesFolder, "flow-m15-left-znone.docx"))
let time (doc: Rbt.Word.Document.Document) =
    linesOf doc |> ignore
    let watch = Diagnostics.Stopwatch.StartNew()
    for _ in 1 .. 20 do linesOf doc |> ignore
    float watch.ElapsedMilliseconds / 20.0
let withHyphens = time flow
let without = time { flow with PageSetup = { flow.PageSetup with Hyphenation = None } }
printfn "flow-m15-left (660 words, 98 lines): %.2f ms a layout with hyphenation, %.2f ms without" withHyphens without

if failures > 0 then
    printfn "%d failures" failures
    exit 1
printfn "hyphenation: all passed"
