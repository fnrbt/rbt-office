// A font set that gains faces (the host fetched a family it had set in a
// fallback): the measurers measure text in those families otherwise, all other
// text as before. Layout.invalidateFacesInIncrementalLayout forgets what was
// laid out in them; the next incremental pass lays those paragraphs out again
// and keeps the lines of the others. The result must be exactly the layout a
// pass from nothing makes with the new measurers - for every test document
// and each of the families it names.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

open System
open System.IO
module Doc = Rbt.Word.Document
module Layout = Rbt.Word.Layout
module TB = Rbt.Word.TextBuffer

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-50s %s" label detail)

let settingsOf (doc: Doc.Document) : Layout.LayoutSettings =
    { Width = doc.PageSetup.PageWidth
      PageHeight = doc.PageSetup.PageHeight
      FontSize = 16.0
      LineHeight = 24.0
      PaddingLeft = doc.PageSetup.MarginLeft
      PaddingTop = doc.PageSetup.MarginTop
      MarginRight = doc.PageSetup.MarginRight
      MarginBottom = doc.PageSetup.MarginBottom
      InTableCell = false
      TableContinuationHeaderHeight = 0.0
      LongTrailingSpaceFitSafety = 3.5 }

let defaultFamily = "Calibri"
let familyOf (style: TB.TextStyle) =
    (style.FontFamily |> Option.filter (fun name -> name.Trim() <> "") |> Option.defaultValue defaultFamily).Trim().ToLowerInvariant()

/// Measurers of a text stack that has faces of `families`: text in one of them is wider.
let measurers (families: Set<string>) : Layout.TextMeasurer * Layout.TextRangeMeasurer =
    let factor (style: TB.TextStyle) = if families.Contains (familyOf style) then 0.61 else 0.5
    let measureText (text: string) (style: TB.TextStyle) = float text.Length * style.FontSize * factor style
    let measureRange (text: string) (style: TB.TextStyle) (start: int) (length: int) =
        float (max 0 (min length (text.Length - start))) * style.FontSize * factor style
    measureText, measureRange

let familiesOfLayouts (layouts: Layout.BlockLayout list) =
    let found = Collections.Generic.HashSet<string>()
    for layout in layouts do
        match layout.PrefixStyle with
        | Some style -> found.Add (familyOf style) |> ignore
        | None -> ()
        for line in layout.Lines do
            for segment in line.Segments do
                match segment with
                | Layout.TextSegment (_, style, _, _, _) -> found.Add (familyOf style) |> ignore
                | _ -> ()
    found |> Seq.sort |> Seq.toList

let root = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "../.."))
let documents =
    Directory.GetDirectories(Path.Combine(root, "tests"), "test-documents*")
    |> Array.append [| Path.Combine(root, "tests/repros/word") |]
    |> Array.collect (fun dir -> Directory.GetFiles(dir, "*.docx", SearchOption.AllDirectories))
    |> Array.sort

let mutable cases = 0
let mutable brokenSelective = 0
let mutable brokenFull = 0
let mutable untouched = 0
for path in documents do
    let name = Path.GetFileName path
    match (try DocxReader.loadBytesNative (File.ReadAllBytes path) with _ -> None) with
    | None -> ()
    | Some parsed ->
        let doc = Rbt.Word.Import.importDoc parsed
        let settings = settingsOf doc
        let before = measurers Set.empty
        let layoutWith (state: Layout.IncrementalLayoutState) (scope: obj) (measureText, measureRange) =
            Layout.withTablePlanScope scope (fun () -> Layout.layoutDocumentPagesIncremental state settings measureText measureRange doc)
        let first, _ = layoutWith (Layout.createIncrementalLayoutState ()) (obj ()) before
        let families = familiesOfLayouts first
        // Each family by itself, all of them at once, and a family the document does not name.
        let arrivals =
            (families |> List.truncate 5 |> List.map (fun family -> [ family ]))
            @ (if families.Length > 1 then [ families ] else [])
            @ [ [ "a family the document does not name" ] ]
        for arriving in arrivals do
            cases <- cases + 1
            let label = sprintf "%s + %s" name (String.concat ", " arriving)
            let names = Set.ofList arriving
            let after = measurers names
            // The layout as it stood, then the faces arrive.
            let state = Layout.createIncrementalLayoutState ()
            layoutWith state (obj ()) before |> ignore
            Layout.resetIncrementalLayoutMeasurements state
            let count = Layout.invalidateFacesInIncrementalLayout state (fun style -> names.Contains (familyOf style))
            let selective = layoutWith state (obj ()) after
            let selectiveBroken = Layout.incrementalLayoutMeasuredBlockCount state
            // The layout from nothing with the new measurers.
            let fresh = Layout.createIncrementalLayoutState ()
            let cold = layoutWith fresh (obj ()) after
            let coldBroken = Layout.incrementalLayoutMeasuredBlockCount fresh
            brokenSelective <- brokenSelective + selectiveBroken
            brokenFull <- brokenFull + coldBroken
            ok (label + ": the layout") (selective = cold)
                (let a, b = fst selective, fst cold
                 if a.Length <> b.Length then sprintf "%d layouts against %d" a.Length b.Length
                 else
                    match List.zip a b |> List.tryFindIndex (fun (x, y) -> x <> y) with
                    | Some index -> sprintf "layout %d differs (block %s); %d paragraphs named the faces" index a.[index].BlockId count
                    | None -> "the pages differ")
            if arriving = [ "a family the document does not name" ] then
                if count = 0 && selectiveBroken = 0 then untouched <- untouched + 1
                ok (label + ": nothing is laid out again") (count = 0 && selectiveBroken = 0) (sprintf "%d paragraphs said to name it, %d broken again" count selectiveBroken)
            // And again from there: the state is one any pass can go on from.
            let again = layoutWith state (obj ()) after
            ok (label + ": the state stands") (again = cold) ""

printfn "faces relayout: %d cases over %d documents; lines broken again for %d paragraphs where a layout from nothing breaks %d; %d passed, %d failed" cases documents.Length brokenSelective brokenFull pass fail
ok "paragraphs in other faces keep their lines" (brokenSelective * 10 < brokenFull * 7) (sprintf "%d of %d" brokenSelective brokenFull)
if fail > 0 then exit 1
