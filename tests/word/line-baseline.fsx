// Where a line's baseline is inside the line, against Word 2019's own PDFs of the fixtures
// `tools/word-parity/fixtures/line-baseline.py` generates (tests/word/fixtures/line-baseline; each probe
// word's baseline from the text matrix of Word's PDF in word/baselines.tsv). Every fixture: Letter, 1in
// margins, header and footer 0.5in from the edge, one font, size and line spacing on every paragraph of
// the header ("Hdrq"), the body ("Bodya", "Bodyb"), a one-cell table ("Cella", "Cellb"), the body again
// ("Refp", with the footnote's reference), the footnote ("Notea", "Noteb") and the footer ("Ftrq").
// What Word showed (docs/layout-parity-flow.md, ledger):
// - at single spacing and at every multiple the baseline is the text's ascent (its single line less its
//   descent) below the line's top - what the multiple adds is all below the text: Calibri 11pt 10.46pt
//   at single, 1.15, 1.5 and double alike; in the header, a cell and a footnote as in the body;
// - at an at-least height above the text's single line, the text's descent above the line's bottom;
// - at an exact height four fifths of the line below its top, whatever the font and its size;
// - a footer ends its footer distance above the page's bottom, however tall it is - its last paragraph's space
//   after included (-fa12: 12pt higher; -fb12 and -f2: the last line where it was);
// - the paragraph mark's size takes no part (-mk*); Word 2010's layout alike (-c14).
// Our side: Paint's DrawText commands for the layout of each fixture, down from their page's top.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open Rbt.Word.TextBuffer
module Layout = Rbt.Word.Layout
module Paint = Rbt.Word.Paint

let dir = Path.Combine(__SOURCE_DIRECTORY__, "fixtures/line-baseline")
let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText
let pt (px: float) = px * 0.75
let probes = set [ "Hdrq"; "Bodya"; "Bodyb"; "Cella"; "Cellb"; "Refp"; "Notea"; "Noteb"; "Ftrp"; "Ftrq" ]

let word =
    File.ReadAllLines(Path.Combine(dir, "word/baselines.tsv")) |> Array.map (fun l -> l.Split '\t')
    |> Array.filter (fun r -> r.Length >= 5 && r.[2] = "1")
    |> Array.groupBy (fun r -> r.[0])
    |> Array.map (fun (name, rs) -> name, rs |> Array.map (fun r -> r.[1], float r.[3]) |> Map.ofArray)
    |> Map.ofArray

let oursOf name =
    let doc = File.ReadAllBytes(Path.Combine(dir, name + ".docx")) |> DocxReader.loadBytesNative |> Option.get |> Rbt.Word.Import.importDoc
    let settings = Layout.documentSettings 16.0 24.0 3.5 doc
    let layouts, pages = Layout.layoutDocumentPages settings measureText measureRange doc
    let firstTop = (List.head pages).Top
    let _, commands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange doc layouts pages
    Paint.flattenedCommands commands
    |> List.choose (function
        | Paint.DrawText (text, _, _, baseline, _) when probes.Contains (text.Trim()) -> Some (text.Trim(), pt (baseline - firstTop))
        | _ -> None)
    |> List.distinctBy fst
    |> Map.ofList

let mutable pass = 0
let mutable fail = 0
let fixtures = Directory.GetFiles(dir, "*.docx") |> Array.map Path.GetFileNameWithoutExtension |> Array.sort
if fixtures.Length <> 86 || fixtures |> Array.exists (fun name -> not (word.ContainsKey name)) then
    printfn "FAIL the fixtures and Word's baselines of each: %d fixtures, %d with Word's" fixtures.Length word.Count
    fail <- fail + 1
for name in fixtures do
    match Map.tryFind name word with
    | None -> ()
    | Some words ->
        let ours = oursOf name
        // (Word's PDFs place text in 0.12pt steps, and ours sums each line's height: 0.25pt.)
        let off =
            words |> Map.toList |> List.choose (fun (probe, at) ->
                match Map.tryFind probe ours with
                | Some mine when abs (mine - at) <= 0.25 -> None
                | Some mine -> Some (sprintf "%s word %.2f ours %.2f" probe at mine)
                | None -> Some (sprintf "%s word %.2f ours none" probe at))
        if off.IsEmpty then pass <- pass + 1
        else
            fail <- fail + 1
            printfn "FAIL %s: %s" name (String.Join("; ", off))
printfn "line-baseline: %d passed, %d failed" pass fail
if fail > 0 then exit 1
