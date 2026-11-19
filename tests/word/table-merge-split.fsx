// A vertically merged cell whose rows cross a page break. Its content is in
// the cell that starts the merge and was laid out in one piece from that
// row's top: it ran on through the bottom margin, the gap between the pages
// and under the header rows repeated on the next page, while the rows beside
// it went to the next page. It is split with the rows it spans now: lines
// that do not fit above the last row of the page continue at the top of the
// first row on the next one (below the repeated header rows); a row that may
// not break ("allow row to break across pages" off) goes to the next page
// whole and the merged content with it; what the rows do not hold makes the
// last row taller; and a row that is kept with the next (the first paragraph
// of its first cell) is not left behind at the foot of a page. The table's
// plan (what the borders are painted from) has the same rows.
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

open Rbt.Word.Document
open Rbt.Word.Layout
module TB = Rbt.Word.TextBuffer
module Paint = Rbt.Word.Paint

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-50s expected %A got %A" label expected got

let settings : LayoutSettings =
    { Width = 816.0; PageHeight = 1056.0; FontSize = 16.0; LineHeight = 24.0; PaddingLeft = 96.0; PaddingTop = 96.0
      MarginRight = 96.0; MarginBottom = 96.0; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5; DefaultTabStop = 48.0; CompatibilityMode = 12 }
let measureText (text: string) (style: TB.TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = rangeMeasurerFromTextMeasurer measureText
let para id (text: string) : ParagraphBlock = { Id = id; Buffer = TB.ofString text; Metadata = defaultMetadata }
let words n = String.concat " " [ for i in 1 .. n -> sprintf "w%d" i ]
let pitch = settings.PageHeight + pageGap
let pageOf (y: float) = int (floor (y / pitch))
let contentTop page = float page * pitch + settings.PaddingTop
let contentBottom page = float page * pitch + settings.PageHeight - settings.MarginBottom
let emptyTable rows =
    let position : Position = { BlockId = "p0"; Offset = 0 }
    let made = insertTable rows 2 { empty with Blocks = [ Paragraph (para "p0" "before") ]; Selection = Some { Anchor = position; Focus = position } }
    made, (made.Blocks |> List.pick (function Table table -> Some table | _ -> None))

/// A table of `rows` rows after `fillers` paragraphs: a header row, then the
/// first column merged over all the other rows with `mergedParagraphs`
/// paragraphs of three lines; `shape` changes a row (by index).
let build (fillers: int) (rows: int) (mergedParagraphs: int) (shape: int -> TableRow -> TableRow) =
    let made, table = emptyTable rows
    let merged =
        { table with
            Rows =
                table.Rows |> List.mapi (fun r row ->
                    { row with
                        IsHeader = (r = 0)
                        Cells =
                            row.Cells |> List.mapi (fun c cell ->
                                if r = 0 then { cell with Content = [ para (sprintf "h%d" c) "HEAD" ] }
                                elif c = 0 then
                                    { cell with
                                        VerticalMerge = Some (if r = 1 then RestartMerge else ContinueMerge)
                                        Content = (if r = 1 then [ for i in 1 .. mergedParagraphs -> para (sprintf "m%d" i) (words 40) ] else [ para (sprintf "c%d" r) "" ]) }
                                else { cell with Content = [ para (sprintf "r%d" r) (sprintf "row %d" r) ] }) }
                    |> shape r) }
    { made with Blocks = [ for i in 1 .. fillers -> Paragraph (para (sprintf "f%d" i) "filler") ] @ [ Table merged; Paragraph (para "after" "after") ]; Selection = None }, merged

let layoutsOf (layouts: BlockLayout list) (id: string) = layouts |> List.filter (fun layout -> layout.BlockId = id)
// (The second cell of a row holds "r<row>", or - the tall row of the third case - "t1" to "t8".)
let rowTop (layouts: BlockLayout list) r =
    match layoutsOf layouts (sprintf "r%d" r) with
    | first :: _ -> first.Y
    | [] -> (layoutsOf layouts "t1" |> List.head).Y
let rowBottom (layouts: BlockLayout list) r =
    match layoutsOf layouts (sprintf "r%d" r) with
    | [] -> let l = layoutsOf layouts "t8" |> List.last in l.Y + l.Height
    | found -> let l = List.last found in l.Y + l.Height
let mergedLines (layouts: BlockLayout list) count = [ for i in 1 .. count do for layout in layoutsOf layouts (sprintf "m%d" i) do yield! layout.Lines ]

/// What holds for every split: the merged lines are in order, each inside
/// the text area of its page and beside rows of the table - not above the
/// first of the rows it spans on that page, not below the last.
let checkSplit (label: string) (doc: Document) (rows: int) (mergedParagraphs: int) =
    let layouts, pages = layoutDocumentPages settings measureText measureRange doc
    let lines = mergedLines layouts mergedParagraphs
    let rowsOn page = [ for r in 1 .. rows - 1 do if pageOf (rowTop layouts r) = page then yield r ]
    let splitRowTop page =
        [ for i in 1 .. 8 do for layout in layoutsOf layouts (sprintf "t%d" i) do for line in layout.Lines do if pageOf line.Y = page then yield line.Y ]
        |> function [] -> System.Double.PositiveInfinity | ys -> List.min ys
    let inside =
        lines |> List.forall (fun line ->
            let page = pageOf line.Y
            match rowsOn page with
            | [] -> false
            | onPage ->
                // (A row that is split has its second part above the first row that starts on the page.)
                let firstTop = min (rowTop layouts (List.head onPage)) (splitRowTop page)
                line.Y >= contentTop page - 0.5 && line.Y + line.Height <= contentBottom page + 8.5
                && line.Y >= firstTop - 0.5)
    check (label + ": the merged lines are in the text area of a page, at or below the top of its first spanned row") inside true
    check (label + ": in order, none on top of another") (lines |> List.pairwise |> List.forall (fun (a, b) -> b.Y >= a.Y + a.Height - 0.5)) true
    // No merged line under the header rows repeated on a later page.
    let headers = layoutsOf layouts "h1"
    check (label + ": none under a repeated header row") (lines |> List.exists (fun line -> headers |> List.exists (fun header -> line.Y < header.Y + header.Height - 0.5 && line.Y + line.Height > header.Y + 0.5))) false
    layouts, pages, lines

// ---- the rows cross a page break ----
do
    let doc, _ = build 20 30 12 (fun _ row -> row)
    let layouts, pages, lines = checkSplit "a page break between the rows" doc 30 12
    let firstPageRows = [ for r in 1 .. 29 do if pageOf (rowTop layouts r) = 0 then yield r ]
    let lastOnFirst, firstOnSecond = List.last firstPageRows, List.last firstPageRows + 1
    check "the rows are on two pages, the header row repeated" (pages.Length, firstPageRows.Length > 3 && firstPageRows.Length < 29, (layoutsOf layouts "h1").Length) (2, true, 2)
    let onFirst, onSecond = lines |> List.partition (fun line -> pageOf line.Y = 0)
    check "the merged content is on both pages" (onFirst.Length > 3, onSecond.Length > 3, onFirst.Length + onSecond.Length) (true, true, 36)
    check "on the first page it ends with the last row there (it ran through the bottom margin)" (onFirst |> List.forall (fun line -> line.Y + line.Height <= rowBottom layouts lastOnFirst + 0.5)) true
    check "on the second page it starts at the top of the first row there" (abs ((List.head onSecond).Y - rowTop layouts firstOnSecond) < 2.0) true
    check "as much of it on the first page as fits beside the rows" (rowBottom layouts lastOnFirst - (let last = List.last onFirst in last.Y + last.Height) < 2.0 * (List.head onFirst).Height + 12.0) true
    check "the paragraph after the table is below its last row" ((layoutsOf layouts "after" |> List.head).Y >= rowBottom layouts 29 - 0.5) true
    // The same from an incremental layout, and after an edit of the merged text.
    let state = createIncrementalLayoutState ()
    check "the same incrementally" (layoutDocumentPagesIncremental state settings measureText measureRange doc = (layouts, pages)) true
    let edited = { doc with Blocks = doc.Blocks |> mapParagraphsOfBlocks (fun p -> if p.Id = "m2" then para "m2" (words 90) else p) }
    check "an edit of the merged text: as a layout from nothing" (layoutDocumentPagesIncremental state settings measureText measureRange edited = layoutDocumentPages settings measureText measureRange edited) true
    checkSplit "after the edit" edited 30 12 |> ignore
    // Painted: every text of the merged cell has its baseline in the text area of a page.
    let _, commands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange doc layouts pages
    let stray =
        commands |> List.choose (function
            | Paint.DrawText (text, _, _, baseline, _) when text.StartsWith "w" && (baseline < contentTop (pageOf baseline) || baseline > contentBottom (pageOf baseline) + 8.0) -> Some baseline
            | _ -> None)
    check "painted: no text of the merged cell in a margin or between the pages" stray []

// ---- more content than the rows hold: the last row is taller ----
do
    let doc, _ = build 20 30 20 (fun _ row -> row)
    let layouts, pages, lines = checkSplit "more content than the rows hold" doc 30 20
    let last = List.last lines
    let after = (layoutsOf layouts "after" |> List.head).Y
    check "the content ends above the paragraph after the table" (last.Y + last.Height <= after + 0.5, lines.Length) (true, 60)
    let tall = layoutDocumentPages settings measureText measureRange (fst (build 0 30 20 (fun _ row -> row)))
    check "(on one page the last row is as tall as the content needs, too)" ((layoutsOf (fst tall) "after" |> List.head).Y >= (let l = mergedLines (fst tall) 20 |> List.last in l.Y + l.Height) - 0.5) true

// ---- a row that may not break goes to the next page whole, a row that may is split ----
do
    let tallRow canSplit r (row: TableRow) =
        if r <> 6 then row
        else { row with CanSplit = Some canSplit; Cells = row.Cells |> List.mapi (fun c cell -> if c = 1 then { cell with Content = [ for i in 1 .. 8 -> para (sprintf "t%d" i) (sprintf "tall %d" i) ] } else cell) }
    // 27 fillers: row 6 starts near the foot of the first page.
    let whole, _ = build 27 12 9 (tallRow false)
    let layouts, _, lines = checkSplit "a row that may not break" whole 12 9
    let tallLines = [ for i in 1 .. 8 do for layout in layoutsOf layouts (sprintf "t%d" i) do yield! layout.Lines ]
    check "the row that may not break is on one page" (tallLines |> List.map (fun line -> pageOf line.Y) |> List.distinct |> List.length, pageOf (List.head tallLines).Y) (1, 1)
    check "the merged content goes on beside it, not above it on that page" (lines |> List.filter (fun line -> pageOf line.Y = 1) |> List.forall (fun line -> line.Y >= (List.head tallLines).Y - 0.5)) true
    // (25 paragraphs before it: the row starts on the first page and goes on to the second.)
    let split, _ = build 25 12 9 (tallRow true)
    let splitLayouts, _, splitLines = checkSplit "a row that is split" split 12 9
    let splitTall = [ for i in 1 .. 8 do for layout in layoutsOf splitLayouts (sprintf "t%d" i) do yield! layout.Lines ]
    check "the row that may break is on both pages" (splitTall |> List.map (fun line -> pageOf line.Y) |> List.distinct) [ 0; 1 ]
    check "and the merged content beside both parts of it" (splitLines |> List.map (fun line -> pageOf line.Y) |> List.distinct) [ 0; 1 ]

// ---- keep with next: a row is not left at the foot of a page without the row it is kept with ----
do
    let made, table = emptyTable 12
    let keep = { defaultMetadata with KeepNext = Some true }
    let rowsOf (kept: int list) =
        { table with
            Rows =
                table.Rows |> List.mapi (fun r row ->
                    { row with
                        CanSplit = Some false
                        Cells = row.Cells |> List.mapi (fun c cell ->
                            let text = if r = 8 && c = 1 then [ for i in 1 .. 6 -> para (sprintf "k%d-%d-%d" r c i) "tall" ] else [ para (sprintf "k%d-%d" r c) (sprintf "row %d" r) ]
                            { cell with Content = text |> List.map (fun p -> if c = 0 && List.contains r kept then { p with Metadata = keep } else p) }) }) }
    let docWith kept = { made with Blocks = [ for i in 1 .. 24 -> Paragraph (para (sprintf "f%d" i) "filler") ] @ [ Table (rowsOf kept); Paragraph (para "after" "after") ]; Selection = None }
    let pageOfRow (layouts: BlockLayout list) r = pageOf (layoutsOf layouts (sprintf "k%d-0" r) |> List.head).Y
    let plain, _ = layoutDocumentPages settings measureText measureRange (docWith [])
    check "without it: the tall row 8 starts the second page, row 7 ends the first" (pageOfRow plain 7, pageOfRow plain 8) (0, 1)
    let kept, _ = layoutDocumentPages settings measureText measureRange (docWith [ 7 ])
    check "row 7 kept with the next: it goes to the second page with row 8; row 6 stays" (pageOfRow kept 6, pageOfRow kept 7, pageOfRow kept 8) (0, 1, 1)
    let chain, _ = layoutDocumentPages settings measureText measureRange (docWith [ 5; 6; 7 ])
    check "rows 5 to 7 kept with the next: all of them go" (pageOfRow chain 4, pageOfRow chain 5, pageOfRow chain 6, pageOfRow chain 7, pageOfRow chain 8) (0, 1, 1, 1, 1)
    check "the rows are in order, one below the other" ([ 5 .. 8 ] |> List.map (fun r -> (layoutsOf chain (sprintf "k%d-0" r) |> List.head).Y) |> List.pairwise |> List.forall (fun (a, b) -> b > a)) true
    let state = createIncrementalLayoutState ()
    check "the same incrementally" (layoutDocumentPagesIncremental state settings measureText measureRange (docWith [ 5; 6; 7 ]) |> fst = chain) true
    // The plan the borders are painted from has the rows on the same pages.
    let _, commands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange (docWith [ 5; 6; 7 ]) chain (snd (layoutDocumentPages settings measureText measureRange (docWith [ 5; 6; 7 ])))
    let horizontals = commands |> List.choose (function Paint.StrokeLine (_, _, _, _, x1, y1, x2, y2) when abs (y1 - y2) < 0.1 && abs (x2 - x1) > 50.0 -> Some y1 | _ -> None)
    let row5 = (layoutsOf chain "k5-0" |> List.head).Y
    let verticals = commands |> List.choose (function Paint.StrokeLine (_, _, _, _, x1, y1, x2, y2) when abs (x1 - x2) < 0.1 -> Some (min y1 y2, max y1 y2) | _ -> None)
    let row4Bottom = (layoutsOf chain "k4-0" |> List.head).Y + 40.0
    check "the borders are where the rows are: the cells' sides start at the top of row 5 on the second page, nothing below row 4 on the first"
        (verticals |> List.exists (fun (top, _) -> abs (top - row5) < 6.0), verticals |> List.exists (fun (top, bottom) -> pageOf top = 0 && bottom > row4Bottom), horizontals |> List.exists (fun y -> pageOf y = 0 && y > row4Bottom))
        (true, false, false)

printfn "word table-merge-split: %d passed, %d failed" pass fail
if fail > 0 then exit 1
