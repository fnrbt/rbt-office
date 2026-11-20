// A table in a section of two or more columns goes from column to column and
// page to page row by row. It was placed whole: a table longer than what was
// left of a column left that empty and started the next one, and a table
// longer than a column ran on below the page's text area. The rows that fit
// stand under the rows above them, the next ones start the next column below
// the table's header rows, which are repeated; the borders are painted
// where the rows are.
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
      MarginRight = 96.0; MarginBottom = 96.0; InTableCell = false; TableContinuationHeaderHeight = 0.0; LongTrailingSpaceFitSafety = 3.5; DefaultTabStop = 48.0; CompatibilityMode = 12; HangingPunctuation = Rbt.Word.Layout.NoHanging; SpaceShrink = 0.0 }
let measureText (text: string) (style: TB.TextStyle) = float text.Length * style.FontSize * 0.5
let measureRange = rangeMeasurerFromTextMeasurer measureText
let para id (text: string) : ParagraphBlock = { Id = id; Buffer = TB.ofString text; Metadata = defaultMetadata }
let pitch = settings.PageHeight + pageGap
let pageOf (y: float) = int (floor (y / pitch))
let contentTop page = float page * pitch + settings.PaddingTop
let contentBottom page = float page * pitch + settings.PageHeight - settings.MarginBottom
let gap = 24.0
let columnWidth = (settings.Width - settings.PaddingLeft - settings.MarginRight - gap) / 2.0
let columns : SectionColumns = { Count = 2; Gap = gap; Separator = false; Widths = [] }

let tableOf (rows: int) (headerRows: int) =
    let position : Position = { BlockId = "p0"; Offset = 0 }
    let made = insertTable rows 2 { empty with Blocks = [ Paragraph (para "p0" "before") ]; Selection = Some { Anchor = position; Focus = position } }
    let table = made.Blocks |> List.pick (function Table table -> Some table | _ -> None)
    made,
    { table with
        Rows =
            table.Rows |> List.mapi (fun r row ->
                { row with
                    IsHeader = r < headerRows
                    Cells = row.Cells |> List.mapi (fun c cell -> { cell with Content = [ para (sprintf "r%d-%d" r c) (if r < headerRows then "HEAD" else sprintf "row %d" r) ] }) }) }
let build (fillers: int) (rows: int) (headerRows: int) =
    let made, table = tableOf rows headerRows
    { made with
        Blocks = [ for i in 1 .. fillers -> Paragraph (para (sprintf "f%d" i) "filler") ] @ [ Table table; Paragraph (para "after" "after the table") ]
        PageSetup = { made.PageSetup with Columns = Some columns }
        Selection = None }
let layoutsOf (layouts: BlockLayout list) (id: string) = layouts |> List.filter (fun layout -> layout.BlockId = id)
/// (page, column) of a layout.
let placeOf (layout: BlockLayout) = pageOf layout.Y, (if layout.Indent > columnWidth then 1 else 0)

// ---- a table longer than a page of two columns ----
do
    let rows = 90
    let doc = build 6 rows 1
    let layouts, pages = layoutDocumentPages settings measureText measureRange doc
    let rowLayout r = layoutsOf layouts (sprintf "r%d-0" r) |> List.head
    let bodyRows = [ 1 .. rows - 1 ] |> List.map rowLayout
    let places = bodyRows |> List.map placeOf |> List.distinct
    check "the rows are in the columns of more than one page, in order" (places.Length >= 3, places = List.sort places, pages.Length >= 2) (true, true, true)
    check "it starts under the paragraphs before it, in the first column" (placeOf (rowLayout 1), (rowLayout 0).Y > (layoutsOf layouts "f6" |> List.head).Y) ((0, 0), true)
    check "every row is in the text area of its page (the table ran on below it)" (bodyRows |> List.filter (fun layout -> layout.Y < contentTop (pageOf layout.Y) - 0.5 || layout.Y + layout.Height > contentBottom (pageOf layout.Y) + 0.5) |> List.length) 0
    let byPlace = bodyRows |> List.groupBy placeOf
    check "in a column the rows stand under one another" (byPlace |> List.forall (fun (_, column) -> column |> List.pairwise |> List.forall (fun (a, b) -> b.Y >= a.Y + a.Height - 0.5 && b.Y < a.Y + a.Height + 12.0))) true
    check "a column is filled before the next one is started" (byPlace |> List.take (byPlace.Length - 1) |> List.forall (fun ((page, _), column) -> let last = List.last column in contentBottom page - (last.Y + last.Height) < 2.5 * last.Height + 12.0)) true
    // The header row is repeated at the top of every column the table goes on in.
    let headers = layoutsOf layouts "r0-0"
    check "the header row is at the top of every column of the table" (headers.Length, headers |> List.map placeOf) (places.Length, places)
    check "in the later columns at the top of the text area, the first row below it" (headers |> List.tail |> List.forall (fun header -> abs (header.Y - contentTop (pageOf header.Y)) < 8.0), byPlace |> List.forall (fun (place, column) -> let header = headers |> List.find (fun h -> placeOf h = place) in (List.head column).Y >= header.Y + header.Height - 0.5)) (true, true)
    let after = layoutsOf layouts "after" |> List.head
    let last = rowLayout (rows - 1)
    check "the paragraph after the table is below its last row, in that column" (placeOf after, after.Y >= last.Y + last.Height - 0.5) (placeOf last, true)
    // Incremental = cold, also after an edit before and inside the table.
    let state = createIncrementalLayoutState ()
    check "the same incrementally" (layoutDocumentPagesIncremental state settings measureText measureRange doc = (layouts, pages)) true
    let edited = { doc with Blocks = doc.Blocks |> mapParagraphsOfBlocks (fun p -> if p.Id = "f2" then para "f2" (String.replicate 30 "longer ") elif p.Id = "r40-1" then para "r40-1" (String.replicate 12 "wrap ") else p) }
    check "after edits: as a layout from nothing" (layoutDocumentPagesIncremental state settings measureText measureRange edited = layoutDocumentPages settings measureText measureRange edited) true
    // Painted: the borders are where the rows are.
    let _, commands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange doc layouts pages
    let verticals = commands |> List.choose (function Paint.StrokeLine (_, _, _, _, x1, y1, x2, y2) when abs (x1 - x2) < 0.1 -> Some (x1, min y1 y2, max y1 y2) | _ -> None)
    let horizontals = commands |> List.choose (function Paint.StrokeLine (_, _, _, _, x1, y1, x2, y2) when abs (y1 - y2) < 0.1 -> Some (y1, min x1 x2, max x1 x2) | _ -> None)
    let leftOf (layout: BlockLayout) = settings.PaddingLeft + (if snd (placeOf layout) = 1 then columnWidth + gap else 0.0)
    let bordered (layout: BlockLayout) =
        let middle = layout.Y + layout.Height / 2.0
        verticals |> List.exists (fun (x, top, bottom) -> abs (x - leftOf layout) < 3.0 && top <= middle && bottom >= middle)
        && verticals |> List.exists (fun (x, top, bottom) -> abs (x - (leftOf layout + columnWidth)) < 3.0 && top <= middle && bottom >= middle)
    check "every row has its left and right border beside it, in its column" (bodyRows @ headers |> List.filter (fun layout -> not (bordered layout)) |> List.map (fun layout -> layout.BlockId)) []
    check "no border outside the text area of a page" (verticals |> List.filter (fun (_, top, bottom) -> top < contentTop (pageOf top) - 2.0 || bottom > contentBottom (pageOf top) + 2.0) |> List.length, horizontals |> List.filter (fun (y, _, _) -> y < contentTop (pageOf y) - 2.0 || y > contentBottom (pageOf y) + 2.0) |> List.length) (0, 0)
    check "no border across the gap between the columns" (horizontals |> List.filter (fun (_, left, right) -> left < settings.PaddingLeft + columnWidth + 2.0 && right > settings.PaddingLeft + columnWidth + gap - 2.0) |> List.length) 0
    // Every column of the table is closed: a line under its last row and over its first.
    let closed =
        byPlace |> List.forall (fun (place, column) ->
            let lastRow = List.last column
            let header = headers |> List.find (fun h -> placeOf h = place)
            horizontals |> List.exists (fun (y, left, _) -> abs (left - leftOf lastRow) < 3.0 && y >= lastRow.Y + lastRow.Height - 6.0 && y <= lastRow.Y + lastRow.Height + 8.0)
            && horizontals |> List.exists (fun (y, left, _) -> abs (left - leftOf header) < 3.0 && y <= header.Y + 2.0 && y >= header.Y - 10.0))
    check "every column of the table has a line over its first row and under its last" closed true

// ---- a point is in the layout it is in sideways: the cell of a row, the column of a page ----
do
    let doc = build 6 90 1
    let layouts, _ = layoutDocumentPages settings measureText measureRange doc
    let hit (x: float) (y: float) = hitTestLayoutsWithRange measureRange settings.PaddingLeft layouts x y |> Option.map fst
    let middleOf (id: string) = let layout = layoutsOf layouts id |> List.head in settings.PaddingLeft + layout.Indent + 6.0, layout.Y + layout.Height / 2.0
    let hitAt id = let x, y = middleOf id in hit x y
    check "a point in the second cell of a row is in that cell (it was in the first cell of the row)" (hitAt "r3-0", hitAt "r3-1") (Some "r3-0", Some "r3-1")
    let inSecondColumn = [ 1 .. 89 ] |> List.find (fun r -> snd (placeOf (layoutsOf layouts (sprintf "r%d-0" r) |> List.head)) = 1)
    check "a point in a row of the second column is in that row (it was in what stands at that height in the first column)" (hitAt (sprintf "r%d-0" inSecondColumn), hitAt (sprintf "r%d-1" inSecondColumn)) (Some (sprintf "r%d-0" inSecondColumn), Some (sprintf "r%d-1" inSecondColumn))
    let fillerX, fillerY = middleOf "f2"
    check "a point in a paragraph, and right of its text in the gap between the columns" (hit fillerX fillerY, hit (settings.PaddingLeft + columnWidth + 4.0) fillerY) (Some "f2", Some "f2")
    check "a point left of the text area is in the first cell" (hit 10.0 (snd (middleOf "r3-0"))) (Some "r3-0")

// ---- a table that fits what is left of the column stays where it is ----
do
    let doc = build 6 8 0
    let layouts, _ = layoutDocumentPages settings measureText measureRange doc
    let rowsAt = [ 0 .. 7 ] |> List.map (fun r -> layoutsOf layouts (sprintf "r%d-0" r)) 
    check "a short table: one layout a row, all in the first column" (rowsAt |> List.forall (fun found -> found.Length = 1), rowsAt |> List.map (fun found -> placeOf found.Head) |> List.distinct) (true, [ (0, 0) ])

// ---- a table that starts near the foot of a column: its first rows stay, the rest goes on ----
do
    let doc = build 30 12 0
    let layouts, _ = layoutDocumentPages settings measureText measureRange doc
    let places = [ 0 .. 11 ] |> List.map (fun r -> placeOf (layoutsOf layouts (sprintf "r%d-0" r) |> List.head))
    check "the first rows stay in the column the table starts in (it was moved to the next column whole)" (List.head places, places |> List.distinct |> List.length) ((0, 0), 2)

// ---- three header rows; a section of two columns before a section of one ----
do
    let made, table = tableOf 70 2
    let section : SectionPageSetup =
        { PageWidth = 816.0; PageHeight = 1056.0; MarginLeft = 96.0; MarginRight = 96.0; MarginTop = 96.0; MarginBottom = 96.0
          HeaderDistance = 48.0; FooterDistance = 48.0; DifferentFirstPage = false; EvenAndOddHeaders = false }
    let doc =
        { made with
            Blocks =
                [ Paragraph (para "intro" "two columns")
                  Table table
                  Paragraph { para "end" "end of the section" with Metadata = { defaultMetadata with SectionBreak = Some "nextPage"; SectionColumns = Some columns; SectionPageSetup = Some section } }
                  Paragraph (para "single" "one column") ]
            Selection = None }
    let layouts, pages = layoutDocumentPages settings measureText measureRange doc
    let firstHeader, secondHeader = layoutsOf layouts "r0-0", layoutsOf layouts "r1-0"
    let places = [ 2 .. 69 ] |> List.map (fun r -> placeOf (layoutsOf layouts (sprintf "r%d-0" r) |> List.head)) |> List.distinct
    check "a table of a two-column section of a document with sections: its rows go from column to column" (places.Length >= 2, firstHeader.Length = places.Length, secondHeader.Length = places.Length) (true, true, true)
    check "both header rows are repeated, one under the other" (List.zip firstHeader secondHeader |> List.forall (fun (a, b) -> placeOf a = placeOf b && b.Y >= a.Y + a.Height - 0.5)) true
    let _, commands, _ = Paint.collectDocumentCommandsWithLayoutsAndPages settings "Calibri" measureText measureRange doc layouts pages
    let verticals = commands |> List.choose (function Paint.StrokeLine (_, _, _, _, x1, y1, x2, y2) when abs (x1 - x2) < 0.1 -> Some (x1, min y1 y2, max y1 y2) | _ -> None)
    let second = layoutsOf layouts "r40-0" |> List.head
    let leftOfSecond = settings.PaddingLeft + (if snd (placeOf second) = 1 then columnWidth + gap else 0.0)
    check "painted with its borders in that column" (verticals |> List.exists (fun (x, top, bottom) -> abs (x - leftOfSecond) < 3.0 && top <= second.Y + 4.0 && bottom >= second.Y + 4.0)) true
    let state = createIncrementalLayoutState ()
    check "the same incrementally" (layoutDocumentPagesIncremental state settings measureText measureRange doc = (layouts, pages)) true

printfn "word table-columns: %d passed, %d failed" pass fail
if fail > 0 then exit 1
