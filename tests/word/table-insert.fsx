// A row or column inserted into a table takes the properties of the cell at
// the insertion position (after a cell: the following cell, or the last when
// appending). A new row copies its neighbouring row's properties, and a new
// cell gets one empty paragraph with the donor's properties. The grid and
// an absolute-width fixed table gain or lose the inserted/deleted width.
#r "../../src/Word/bin/Debug/netstandard2.0/Rbt.Word.dll"

open Rbt.Word.Document
module TB = Rbt.Word.TextBuffer

let mutable pass = 0
let mutable fail = 0
let check label got expected =
    if got = expected then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %-50s expected %A got %A" label expected got

let para id (text: string) : ParagraphBlock = { Id = id; Buffer = TB.ofString text; Metadata = defaultMetadata }
let position id : Position = { BlockId = id; Offset = 0 }
let made =
    insertTable 2 2 { empty with Blocks = [ Paragraph (para "p0" "before") ]; Selection = Some { Anchor = position "p0"; Focus = position "p0" } }
let tableOf (doc: Document) = doc.Blocks |> List.pick (function Table table -> Some table | _ -> None)
let border color : TableBorder = { Style = Some "single"; Color = Some color; Width = Some 1.0 }
let width value : TableWidth = { Kind = DxaWidth; Value = Some value }
// The table dressed: every cell with borders, shading, margins, a vertical alignment and a width of its own; its
// first paragraph centred with a style and spacing; the rows 40 and 60 high; the grid 1500 and 3000 wide.
let dressed =
    let table = tableOf made
    let rows =
        table.Rows
        |> List.mapi (fun r row ->
            { row with
                Height = Some (40.0 + 20.0 * float r); HeightRule = Some "atLeast"; CanSplit = Some false; IsHeader = (r = 0)
                Cells =
                    row.Cells
                    |> List.mapi (fun c cell ->
                        let name = sprintf "r%dc%d" r c
                        { cell with
                            Content = [ { para name name with Metadata = { defaultMetadata with Style = Some "TableText"; Alignment = Center; ExplicitAlignment = true; SpacingBefore = Some 4.0; SpacingAfter = Some 2.0; ParagraphMarkFontSize = Some 18.0; PageBreakBefore = (r = 1) } } ]
                            BackgroundColor = Some (sprintf "#0%d0%d00" r c)
                            Borders = Some { Top = Some (border "#FF0000"); Right = None; Bottom = Some (border (sprintf "#00%d%d00" r c)); Left = None; InsideH = None; InsideV = None }
                            VerticalAlignment = Some "center"
                            PaddingTop = Some 3.0; PaddingLeft = Some (5.0 + float c)
                            Width = Some (100.0 + 100.0 * float c); PreferredWidth = Some (width (1500.0 + 1500.0 * float c))
                            VerticalMerge = (if c = 1 then Some (if r = 0 then RestartMerge else ContinueMerge) else None) }) })
    { made with Blocks = made.Blocks |> List.map (function Table t -> Table { t with LayoutMode = Some FixedTableLayout; PreferredWidth = Some (width 4500.0); Rows = rows; GridColumns = [ width 1500.0; width 3000.0 ] } | block -> block) }
let at id (doc: Document) = { doc with Selection = Some { Anchor = position id; Focus = position id } }
/// What of a cell a new cell next to it is to have, and what of its paragraph.
let look (cell: TableCell) =
    cell.BackgroundColor, cell.Borders, cell.VerticalAlignment, (cell.PaddingTop, cell.PaddingLeft), (cell.Width, cell.PreferredWidth)
let paragraphLook (cell: TableCell) =
    cell.Content |> List.map (fun p -> TB.toString p.Buffer, p.Metadata.Style, p.Metadata.Alignment, p.Metadata.ExplicitAlignment, p.Metadata.SpacingBefore, p.Metadata.SpacingAfter, p.Metadata.ParagraphMarkFontSize, p.Metadata.PageBreakBefore)
let emptyLook = [ "", Some "TableText", Center, true, Some 4.0, Some 2.0, Some 18.0, false ]
let ids (table: TableBlock) = table.Rows |> List.collect (fun row -> row.Cells |> List.collect (fun cell -> cell.Content |> List.map (fun p -> p.Id)))
let rowLook (row: TableRow) = row.Height, row.HeightRule, row.CanSplit, row.IsHeader

// ---- a row below ----
let below = tableOf (insertTableRowAfterSelection (dressed |> at "r1c0"))
check "a row below: three rows, the new one after the second" (below.Rows.Length, below.Rows.[1].Cells.[0].Content.Head.Id) (3, "r1c0")
check "its cells are like those of the row above it" (below.Rows.[2].Cells |> List.map look) (below.Rows.[1].Cells |> List.map look)
check "each with one empty paragraph like the neighbour's (not its page break)" (below.Rows.[2].Cells |> List.map paragraphLook) [ emptyLook; emptyLook ]
check "the row is like its neighbour" (rowLook below.Rows.[2]) (Some 60.0, Some "atLeast", Some false, false)
check "the new cells are not part of a vertical merge" (below.Rows.[2].Cells |> List.map (fun cell -> cell.VerticalMerge)) [ None; None ]
check "the paragraphs have ids of their own" ((ids below |> List.distinct).Length) 6
check "the grid is as it was" (below.GridColumns |> List.map (fun w -> w.Value)) [ Some 1500.0; Some 3000.0 ]

// ---- a row above ----
let above = tableOf (insertTableRowBeforeSelection (dressed |> at "r0c1"))
check "a row above the first: like the first, and not a header row" (above.Rows.Length, above.Rows.[0].Cells |> List.map look, rowLook above.Rows.[0], above.Rows.[0].Cells |> List.map paragraphLook) (3, above.Rows.[1].Cells |> List.map look, (Some 40.0, Some "atLeast", Some false, false), [ emptyLook; emptyLook ])
check "the first row of before is still the header row" (rowLook above.Rows.[1]) (Some 40.0, Some "atLeast", Some false, true)

// ---- a column to the right ----
let right = tableOf (insertTableColumnAfterSelection (dressed |> at "r0c0"))
check "a column to the right of the first: three cells in every row" (right.Rows |> List.map (fun row -> row.Cells.Length)) [ 3; 3 ]
check "the new cells take the properties of the cell at the insertion position" (right.Rows |> List.map (fun row -> look row.Cells.[1])) (right.Rows |> List.map (fun row -> look row.Cells.[2]))
check "with an empty paragraph like that cell's" (right.Rows |> List.map (fun row -> paragraphLook row.Cells.[1])) [ emptyLook; emptyLook ]
check "the cells that were there are where they were" (right.Rows |> List.map (fun row -> row.Cells.[0].Content.Head.Id, row.Cells.[2].Content.Head.Id)) [ "r0c0", "r0c1"; "r1c0", "r1c1" ]
check "the grid copies the following column and the fixed table grows by its width" (right.GridColumns |> List.map (fun w -> w.Value), right.PreferredWidth.Value.Value) ([ Some 1500.0; Some 3000.0; Some 3000.0 ], Some 7500.0)

// ---- a column to the left ----
let left = tableOf (insertTableColumnBeforeSelection (dressed |> at "r1c1"))
check "a column to the left of the second: like the second" (left.Rows |> List.map (fun row -> row.Cells.Length, look row.Cells.[1] = look row.Cells.[2], row.Cells.[1].VerticalMerge, row.Cells.[2].Content.Head.Id)) [ (3, true, None, "r0c1"); (3, true, None, "r1c1") ]
check "the grid copies the second column and the fixed table grows by its width" (left.GridColumns |> List.map (fun w -> w.Value), left.PreferredWidth.Value.Value) ([ Some 1500.0; Some 3000.0; Some 3000.0 ], Some 7500.0)
let appended = tableOf (insertTableColumnAfterSelection (dressed |> at "r0c1"))
check "appending copies the last cell's formatting and width" (appended.Rows |> List.map (fun row -> look row.Cells.[2] = look row.Cells.[1]), appended.GridColumns |> List.map (fun w -> w.Value), appended.PreferredWidth.Value.Value) ([ true; true ], [ Some 1500.0; Some 3000.0; Some 3000.0 ], Some 7500.0)
let prepended = tableOf (insertTableColumnBeforeSelection (dressed |> at "r0c0"))
check "prepending copies the first cell's formatting and width" (prepended.Rows |> List.map (fun row -> look row.Cells.[0] = look row.Cells.[1]), prepended.GridColumns |> List.map (fun w -> w.Value), prepended.PreferredWidth.Value.Value) ([ true; true ], [ Some 1500.0; Some 1500.0; Some 3000.0 ], Some 6000.0)
let percent =
    let withPercent block =
        match block with
        | Table table -> Table { table with PreferredWidth = Some { Kind = PercentWidth; Value = Some 50.0 } }
        | other -> other
    { dressed with Blocks = dressed.Blocks |> List.map withPercent }
let percentAfter = tableOf (insertTableColumnAfterSelection (percent |> at "r0c0"))
check "a percentage table keeps its preferred-width unit and value" percentAfter.PreferredWidth (Some { Kind = PercentWidth; Value = Some 50.0 })

// ---- a cell that spans: the new cell of a column is one column wide, that of a row spans as its neighbour ----
let spanning =
    { dressed with Blocks = dressed.Blocks |> List.map (function Table t -> Table { t with Rows = t.Rows |> List.mapi (fun r row -> if r = 0 then { row with Cells = [ { row.Cells.[0] with ColSpan = 2 } ] } else row) } | block -> block) }
let spanRow = tableOf (insertTableRowAfterSelection (spanning |> at "r0c0"))
check "a row below a row with a cell over two columns: one cell over two columns" (spanRow.Rows.[1].Cells |> List.map (fun cell -> cell.ColSpan)) [ 2 ]
let spanColumn = tableOf (insertTableColumnAfterSelection (spanning |> at "r1c0"))
check "a column: its cells are one column wide" (spanColumn.Rows |> List.map (fun row -> row.Cells |> List.map (fun cell -> cell.ColSpan))) [ [ 2; 1 ]; [ 1; 1; 1 ] ]

// ---- a column deleted ----
let deleted = tableOf (deleteSelectedTableColumn (dressed |> at "r0c0"))
check "a column deleted: the grid and fixed width lose its column" (deleted.Rows |> List.map (fun row -> row.Cells |> List.map (fun cell -> cell.Content.Head.Id)), deleted.GridColumns |> List.map (fun w -> w.Value), deleted.PreferredWidth.Value.Value) ([ [ "r0c1" ]; [ "r1c1" ] ], [ Some 3000.0 ], Some 3000.0)
let roundTrip = tableOf (deleteSelectedTableColumn (insertTableColumnAfterSelection (dressed |> at "r0c0") |> fun doc -> at (tableOf doc).Rows.[0].Cells.[1].Content.Head.Id doc))
check "a column put in and taken out again: the grid and fixed width that were" (roundTrip.GridColumns |> List.map (fun w -> w.Value), roundTrip.PreferredWidth.Value.Value, roundTrip.Rows |> List.map (fun row -> row.Cells.Length)) ([ Some 1500.0; Some 3000.0 ], Some 4500.0, [ 2; 2 ])

printfn "word table-insert: %d passed, %d failed" pass fail
if fail > 0 then exit 1
