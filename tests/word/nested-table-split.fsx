#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Pdf.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Unicode.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Bidi.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.OpenType.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Shaper.Core.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Text.Engine.dll"
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"
#r "../../src/ShaperText/bin/Debug/net10.0/Rbt.Word.ShaperText.dll"
#r "../../src/Word.Pdf/bin/Debug/net10.0/Rbt.Word.Pdf.dll"
#r "../../src/Word.Import/bin/Debug/net10.0/Rbt.WordImport.dll"

// A table row that is split across a page break, holding a nested table
// (tests/test-documents/sample-formatted.docx). The borders of the outer
// cells used to be painted with the row's whole height - which runs from the
// row's top on one page to its bottom on the next and so includes the page
// break - in each fragment: they ran through the bottom margin, the gap
// between the pages and on below the table. The nested table's rows were
// drawn where a table laid out on its own would have them, not where the
// split had put them. The paint commands are checked directly: the editor
// (whole canvas and viewport) and the PDF writer both paint from them.

open System
open System.IO
open Rbt.Word.Document
open Rbt.Word.Layout
open Rbt.Word.Paint
module TB = Rbt.Word.TextBuffer
module Pdf = Rbt.Word.Pdf

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s %s" label detail)

let root = Path.GetFullPath(Path.Combine(__SOURCE_DIRECTORY__, "../.."))
let source =
    match DocxReader.loadBytesNative (File.ReadAllBytes(Path.Combine(root, "tests/test-documents/sample-formatted.docx"))) with
    | Some parsed -> Rbt.Word.Import.importDoc parsed
    | None -> failwith "sample-formatted.docx could not be read"
let outer =
    source.Blocks
    |> List.pick (function
        | Table table when table.Rows |> List.exists (fun row -> row.Cells |> List.exists (fun cell -> not cell.NestedTables.IsEmpty)) -> Some table
        | _ -> None)
let nested = outer.Rows |> List.collect (fun row -> row.Cells |> List.collect (fun cell -> cell.NestedTables)) |> List.head
ok "the document has a table with a nested table of two rows" (nested.Rows.Length = 2) (string nested.Rows.Length)

// Narrow text, so that the table is as it is in the editor with a shaped text stack.
let k = 0.45
let measureText (text: string) (style: TB.TextStyle) = float text.Length * style.FontSize * k
let measureRange (_: string) (style: TB.TextStyle) (_: int) (length: int) = float length * style.FontSize * k
let paragraph id (text: string) : ParagraphBlock = { Id = id; Buffer = TB.empty |> TB.insert 0 text; Metadata = defaultMetadata }
// The table after `fillers` paragraphs and one more with `gap` pixels above it: where on the page it starts.
let build (fillers: int) (gap: float) =
    let spacer = { paragraph "spacer" "spacer" with Metadata = { defaultMetadata with SpacingBefore = Some gap; Explicit = ParagraphExplicit.SpacingBefore } }
    { source with
        Blocks = [ for i in 1 .. fillers -> Paragraph (paragraph (sprintf "filler-%d" i) "filler") ] @ [ Paragraph spacer; Table outer; Paragraph (paragraph "after" "after the table") ]
        Selection = None }
let verticals (commands: PaintCommand list) =
    commands |> List.choose (function StrokeLine (_, _, _, _, x1, y1, x2, y2) when abs (x1 - x2) < 0.1 -> Some (x1, min y1 y2, max y1 y2) | _ -> None)
let horizontals (commands: PaintCommand list) =
    commands |> List.choose (function StrokeLine (_, _, _, _, x1, y1, x2, y2) when abs (y1 - y2) < 0.1 -> Some (y1, min x1 x2, max x1 x2) | _ -> None)
let linesOf (layouts: BlockLayout list) (para: ParagraphBlock) = layouts |> List.filter (fun layout -> layout.BlockId = para.Id) |> List.collect (fun layout -> layout.Lines)
let cellParagraphs (table: TableBlock) = table.Rows |> List.collect (fun row -> row.Cells |> List.collect (fun cell -> cell.Content))

/// Check the borders of one painted document. Returns whether the outer row was split.
let checkBorders (label: string) (settings: LayoutSettings) (layouts: BlockLayout list) (commands: PaintCommand list) : bool =
    let pitch = settings.PageHeight + pageGap
    let pageOf (y: float) = int (floor (y / pitch))
    let contentTop page = float page * pitch + settings.PaddingTop
    let contentBottom page = float page * pitch + settings.PageHeight - settings.MarginBottom
    let outerLines = cellParagraphs outer |> List.collect (linesOf layouts)
    let nestedLines = cellParagraphs nested |> List.collect (linesOf layouts)
    let allLines = outerLines @ nestedLines
    let tableBottom = allLines |> List.map (fun line -> line.Y + line.Height) |> List.max
    let pagesUsed = allLines |> List.map (fun line -> pageOf line.Y) |> List.distinct |> List.sort
    let lines = verticals commands
    // Every side of a cell stays on one page, inside its text area (a row that just fits may overhang by a few pixels) ...
    let outside =
        lines |> List.filter (fun (_, top, bottom) ->
            let page = pageOf top
            top < contentTop page - 1.0 || bottom > contentBottom page + 6.0)
    ok (label + ": no border leaves its page's text area") outside.IsEmpty (sprintf "%A" (List.truncate 3 outside))
    // ... and none goes on below the table.
    let below = lines |> List.filter (fun (_, _, bottom) -> bottom > tableBottom + 24.0)
    ok (label + ": no border runs on below the table") below.IsEmpty (sprintf "table ends at %.0f: %A" tableBottom (List.truncate 3 below))
    // The text of every cell is between sides of cells: the borders are where the content is.
    let uncovered =
        allLines |> List.filter (fun line ->
            let middle = line.Y + line.Height / 2.0
            not (lines |> List.exists (fun (_, top, bottom) -> top <= middle && middle <= bottom)))
    ok (label + ": every line of the table has cell sides beside it") uncovered.IsEmpty (sprintf "%d lines, first at %.0f" uncovered.Length (if uncovered.IsEmpty then 0.0 else uncovered.Head.Y))
    if pagesUsed.Length > 1 then
        // Split: the table is closed at its bottom on the second page, and has sides on both pages.
        let second = pagesUsed.[1]
        let sidesOnBoth = pagesUsed |> List.forall (fun page -> lines |> List.exists (fun (_, top, _) -> pageOf top = page))
        ok (label + ": sides on both pages") sidesOnBoth ""
        let bottoms = horizontals commands |> List.filter (fun (y, _, _) -> pageOf y = second && abs (y - tableBottom) < 6.0)
        ok (label + ": the table's bottom border is under its last line") (not bottoms.IsEmpty) (sprintf "bottom %.0f" tableBottom)
        // The nested table's rows are drawn where their text is.
        let nestedOnSecond = nestedLines |> List.filter (fun line -> pageOf line.Y = second)
        if not nestedOnSecond.IsEmpty then
            let nestedBottom = nestedOnSecond |> List.map (fun line -> line.Y + line.Height) |> List.max
            let closing = horizontals commands |> List.filter (fun (y, _, _) -> pageOf y = second && y >= nestedBottom - 1.0 && y <= nestedBottom + 24.0)
            ok (label + ": the nested table's moved row has its bottom border on the second page") (closing.Length >= 1) (sprintf "nested text ends at %.0f" nestedBottom)
    pagesUsed.Length > 1

// ---- the editor's paint commands -------------------------------------------------
let hasFooter = not source.Footers.IsEmpty || not source.FooterParts.IsEmpty
let settings : LayoutSettings =
    { Width = source.PageSetup.PageWidth
      PageHeight = source.PageSetup.PageHeight
      FontSize = 16.0
      LineHeight = 24.0
      PaddingLeft = source.PageSetup.MarginLeft
      PaddingTop = Rbt.Word.WordLayout.Page.bodyTopWithRunningHeader source
      MarginRight = source.PageSetup.MarginRight
      MarginBottom = max source.PageSetup.MarginBottom (if hasFooter then source.PageSetup.FooterDistance + 36.0 else source.PageSetup.MarginBottom)
      InTableCell = false
      TableContinuationHeaderHeight = 0.0
      LongTrailingSpaceFitSafety = 3.5 }
let mutable editorSplits = 0
let mutable editorNestedSplits = 0
for fillers in 20 .. 27 do
    for step in 0 .. 16 do
        let gap = float step * 2.0
        let doc = build fillers gap
        let layouts, pages = layoutDocumentPages settings measureText measureRange doc
        let _, commands, _ = collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange doc layouts pages
        if checkBorders (sprintf "editor %d+%.0f" fillers gap) settings layouts commands then
            editorSplits <- editorSplits + 1
            let pitch = settings.PageHeight + pageGap
            let nestedPages = cellParagraphs nested |> List.collect (linesOf layouts) |> List.map (fun line -> int (floor (line.Y / pitch))) |> List.distinct
            if nestedPages.Length > 1 then editorNestedSplits <- editorNestedSplits + 1
ok "editor: the scan has positions where the row is split across the page break" (editorSplits >= 3) (string editorSplits)
ok "editor: and where the nested table's rows are on different pages" (editorNestedSplits >= 1) (string editorNestedSplits)

// ---- the PDF writer's paint commands (its own settings) ------------------------------
let mutable pdfSplits = 0
for fillers in 20 .. 32 do
    for step in 0 .. 8 do
        let gap = float step * 4.0
        let doc = build fillers gap
        let pdfSettings, commands, _ = Pdf.collectPaintCommands "Calibri" measureText measureRange doc
        let layouts, _ = layoutDocumentPages pdfSettings measureText measureRange doc
        if checkBorders (sprintf "pdf %d+%.0f" fillers gap) pdfSettings layouts commands then pdfSplits <- pdfSplits + 1
ok "pdf: the scan has positions where the row is split across the page break" (pdfSplits >= 3) (string pdfSplits)

printfn ""
printfn "Nested table across a page break: %d passed, %d failed (%d editor and %d PDF split positions)" pass fail editorSplits pdfSplits
if fail > 0 then exit 1
