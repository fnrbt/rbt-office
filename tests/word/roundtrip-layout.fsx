// A document saved by the editor is laid out as the document it was opened
// from: the same pages, every block at the same place with the same lines and
// the same resolved formatting.
//
// The writer (DocumentWriter) runs in the browser only, so the save itself is
// checked there, for the synthetic set and the whole corpus, unedited and
// edited: tools/docx-roundtrip/roundtrip.mjs. This suite is the .NET half:
//
//   * LayoutFingerprint (what both halves compare layouts by) tells two
//     layouts apart by a moved block, a changed line count, changed paragraph
//     or run formatting, and lets a sub-pixel difference pass;
//   * every document of the corpus read twice gives the same fingerprint (the
//     comparison has nothing in it that changes from one load to the next);
//   * documents saved by the browser writer and kept under
//     tests/word/roundtrip-saved read back, with the .NET reader and layout,
//     as their originals: <name>.model.docx was written from the model alone
//     (what a changed block is written as), <name>.pass1.docx is the editor's
//     save of the unedited document (its body is the text that was read).
//     The original is <name>.original.docx next to them, or the corpus file
//     of that name.
//
// With a directory argument the last part runs over that directory instead:
//   node tools/docx-roundtrip/roundtrip.mjs --dist $DIST --out DIR
//   dotnet fsi tests/word/roundtrip-layout.fsx DIR
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
open Rbt.Word.TextBuffer
module Doc = Rbt.Word.Document
module Layout = Rbt.Word.Layout
module Fingerprint = Rbt.Word.LayoutFingerprint

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-46s %s" label detail

let measureText (text: string) (style: TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = Layout.rangeMeasurerFromTextMeasurer measureText

let settingsOf (doc: Doc.Document) : Layout.LayoutSettings =
    let lineHeight = 24.0
    let hasFooter = not doc.Footers.IsEmpty || not doc.FooterParts.IsEmpty
    let footerReserve = if hasFooter then doc.PageSetup.FooterDistance + lineHeight * 1.5 else doc.PageSetup.MarginBottom
    { Width = doc.PageSetup.PageWidth
      PageHeight = doc.PageSetup.PageHeight
      FontSize = 16.0
      LineHeight = lineHeight
      PaddingLeft = doc.PageSetup.MarginLeft
      PaddingTop = Rbt.Word.WordLayout.Page.bodyTopWithRunningHeader doc
      MarginRight = doc.PageSetup.MarginRight
      MarginBottom = max doc.PageSetup.MarginBottom footerReserve
      InTableCell = false
      TableContinuationHeaderHeight = 0.0
      LongTrailingSpaceFitSafety = 3.5; DefaultTabStop = 48.0; CompatibilityMode = 12 }

/// Page count and fingerprint of a document.
let layoutOf (doc: Doc.Document) =
    let layouts, pages = Layout.layoutDocumentPages (settingsOf doc) measureText measureRange doc
    let tops = pages |> List.map (fun page -> page.Top) |> List.toArray
    let pageOf (y: float) =
        let mutable index = 0
        for i in 0 .. tops.Length - 1 do
            if y >= tops.[i] - 0.01 then index <- i
        index + 1
    pages.Length, Fingerprint.ofLayouts pageOf layouts

let load (path: string) =
    File.ReadAllBytes path |> DocxReader.loadBytesNative |> Option.map Rbt.Word.Import.importDoc

// ---- the fingerprint tells layouts apart ----------------------------------------

let para id metadata text : Doc.ParagraphBlock = { Id = id; Buffer = ofString text; Metadata = metadata }
let words count = String.concat " " [ for i in 1 .. count -> sprintf "word%d" i ]
let sample (second: Doc.ParagraphBlock) =
    { Doc.empty with
        Blocks = [ Doc.Paragraph (para "a" Doc.defaultMetadata (words 60)); Doc.Paragraph second; Doc.Paragraph (para "c" Doc.defaultMetadata (words 30)) ]
        Styles = []
        Selection = None }
let plain = sample (para "b" Doc.defaultMetadata (words 40))
let _, plainPrint = layoutOf plain
let difference (doc: Doc.Document) = Fingerprint.firstDifference 0.5 plainPrint (snd (layoutOf doc))
let mentions (what: string) (found: string option) = found |> Option.exists (fun text -> text.Contains what)

ok "the same document: no difference" (difference plain = None) (sprintf "%A" (difference plain))
ok "spacing after: the next block moves" (difference (sample (para "b" { Doc.defaultMetadata with SpacingAfter = Some 30.0 } (words 40))) |> Option.isSome) "no difference found"
ok "an indent: paragraph format" (difference (sample (para "b" { Doc.defaultMetadata with IndentLeft = Some 48.0 } (words 40))) |> Option.isSome) "no difference found"
ok "more text: other text" (mentions "text" (difference (sample (para "b" Doc.defaultMetadata (words 41))))) "no text difference"
let boldBuffer =
    let text = words 40
    { Rope = Rbt.Word.Rope.ofString text; Runs = [ Text { Start = 0; Length = text.Length; Style = { defaultStyle with Bold = true; Explicit = ExplicitFormat.Bold } } ] }
ok "bold text: run format" (mentions "run format" (difference (sample { para "b" Doc.defaultMetadata (words 40) with Buffer = boldBuffer }))) "no run format difference"
ok "a block fewer" (mentions "blocks became" (difference { plain with Blocks = List.take 2 plain.Blocks })) "block count not reported"
let nudged = plainPrint |> List.map (fun block -> { block with Y = block.Y + 0.3 })
ok "a third of a pixel is within the tolerance" (Fingerprint.firstDifference 0.5 plainPrint nudged = None) "difference reported"
let moved = plainPrint |> List.mapi (fun index block -> if index = 2 then { block with Y = block.Y + 2.0 } else block)
ok "two pixels are not" (mentions "y" (Fingerprint.firstDifference 0.5 plainPrint moved)) "no difference reported"

// ---- the corpus, read twice --------------------------------------------------------

let repo = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "../.."))
let corpus =
    Directory.GetDirectories(Path.Combine(repo, "tests"), "test-documents*")
    |> Array.collect (fun dir -> Directory.GetFiles(dir, "*.docx", SearchOption.AllDirectories))
    |> Array.sort
ok "the corpus is there" (corpus.Length >= 60) (sprintf "%d documents" corpus.Length)
// (The documents that are megabytes of pictures or thousands of blocks are left to the browser run.)
let small = corpus |> Array.filter (fun path -> FileInfo(path).Length < 300_000L)
let mutable readTwice = 0
for path in small do
    match load path, load path with
    | Some first, Some second ->
        let pagesA, printA = layoutOf first
        let pagesB, printB = layoutOf second
        readTwice <- readTwice + 1
        ok ("read twice: " + Path.GetFileName path) (pagesA = pagesB && Fingerprint.firstDifference 0.0 printA printB = None)
            (sprintf "%d pages / %d pages, %A" pagesA pagesB (Fingerprint.firstDifference 0.0 printA printB))
    | _ -> ()
ok "documents read twice" (readTwice >= 40) (sprintf "%d" readTwice)

// ---- saved documents against their originals -----------------------------------------

let savedDir =
    match fsi.CommandLineArgs |> Array.tryItem 1 with
    | Some dir -> Path.GetFullPath dir
    | None -> Path.Combine(__SOURCE_DIRECTORY__, "roundtrip-saved")
let originalOf (name: string) =
    let beside = Path.Combine(savedDir, name + ".original.docx")
    if File.Exists beside then Some beside
    else corpus |> Array.tryFind (fun path -> Path.GetFileNameWithoutExtension path = name)
let savedFiles =
    [ for kind in [ "model"; "pass1" ] do
        for path in Directory.GetFiles(savedDir, "*." + kind + ".docx") |> Array.sort do
            yield kind, Path.GetFileName(path).Replace("." + kind + ".docx", ""), path ]
ok "saved documents are there" (savedFiles.Length >= 3) (sprintf "%d under %s" savedFiles.Length savedDir)
let mutable compared = 0
let originals = Collections.Generic.Dictionary<string, (int * Fingerprint.BlockFingerprint list) option>()
for kind, name, path in savedFiles do
    let label = sprintf "%s.%s" name kind
    let original =
        match originals.TryGetValue name with
        | true, known -> known
        | _ ->
            let found = originalOf name |> Option.bind load |> Option.map layoutOf
            originals.[name] <- found
            found
    match original, load path with
    | Some (pages, print), Some saved ->
        let savedPages, savedPrint = layoutOf saved
        compared <- compared + 1
        ok (label + ": pages") (savedPages = pages) (sprintf "%d pages became %d" pages savedPages)
        let found = Fingerprint.firstDifference 0.5 print savedPrint
        ok (label + ": layout") found.IsNone (found |> Option.defaultValue "")
    // (A fixture the harness changes before it saves, and the document made in the editor, have none.)
    | None, _ -> ()
    | _, None -> ok (label + ": opens") false "the .NET reader could not open the saved document"

ok "saved documents compared with their originals" (compared >= 3) (sprintf "%d" compared)

printfn "word-roundtrip-layout: %d passed, %d failed" pass fail
if fail > 0 then exit 1
