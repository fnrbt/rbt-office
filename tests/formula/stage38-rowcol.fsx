#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
module Ops = Rbt.Spreadsheet.SheetOps
module Layout = Rbt.Spreadsheet.Layout

// Rows and columns as the host API changes them: sizes, hidden rows and
// columns that come back at the size they had, auto-fit against a text
// measure, the outline (groups, collapse / expand), and how all of it
// follows inserted and deleted rows and columns.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s expected %A  got %A" label expected got)
let throws label (run: unit -> unit) (part: string) =
    try
        run ()
        fail <- fail + 1
        printfn "FAIL %-40s did not throw" label
    with error ->
        if error.Message.Contains part then pass <- pass + 1
        else (fail <- fail + 1; printfn "FAIL %-40s threw %s" label error.Message)

let at row col : CellAddress = { Row = row; Col = col }
let text (value: string) = { Value = CellValue.String value; Format = defaultFormat; Comment = None }
let num (n: float) = { Value = CellValue.Number n; Format = defaultFormat; Comment = None }
let withFormat (change: CellFormat -> CellFormat) (cell: CellData) = { cell with Format = change cell.Format }
let sheetOf cells = cells |> List.fold (fun m (addr, cell) -> setCell addr cell m) (createEmpty 50 10)
let empty = createEmpty 50 10

// ---- sizes ---------------------------------------------------------------------------------
let sized = empty |> Ops.setAxisSize false 1 3 80.0 |> Ops.setAxisSize true 4 4 40.0
eq "sizes: columns B..D and row 5" (sized.ColumnWidths, sized.RowHeights) (Map.ofList [ 1, 80.0; 2, 80.0; 3, 80.0 ], Map.ofList [ 4, 40.0 ])
eq "sizes: read back" ([ 0; 1; 3; 4 ] |> List.map (fun c -> Ops.axisSize false c sized), Ops.axisSize true 4 sized, Ops.axisSize true 5 sized) ([ 100.0; 80.0; 80.0; 100.0 ], 40.0, 25.0)
eq "sizes: the default size is no entry" ((sized |> Ops.setAxisSize false 1 2 100.0).ColumnWidths) (Map.ofList [ 3, 80.0 ])
eq "sizes: the same size again is the same model" (obj.ReferenceEquals(Ops.setAxisSize false 1 3 80.0 sized, sized)) true
eq "sizes: a span is clipped to the sheet" ((empty |> Ops.setAxisSize false 8 99 60.0).ColumnWidths) (Map.ofList [ 8, 60.0; 9, 60.0 ])
eq "sizes: a span off the sheet changes nothing" (obj.ReferenceEquals(Ops.setAxisSize true 60 70 30.0 empty, empty)) true

// ---- hidden rows and columns ---------------------------------------------------------------
let hidden = sized |> Ops.hideAxis false 2 4 |> Ops.hideAxis true 3 5
eq "hide: size 0 in the size maps (what every layout reads)" (hidden.ColumnWidths, hidden.RowHeights)
    (Map.ofList [ 1, 80.0; 2, 0.0; 3, 0.0; 4, 0.0 ], Map.ofList [ 3, 0.0; 4, 0.0; 5, 0.0 ])
eq "hide: the sizes they had are kept (default ones need none)" (hidden.Axes.HiddenColSizes, hidden.Axes.HiddenRowSizes) (Map.ofList [ 2, 80.0; 3, 80.0 ], Map.ofList [ 4, 40.0 ])
eq "hide: asked" ([ 1; 2; 4; 5 ] |> List.map (fun c -> Ops.isAxisHidden false c hidden), Ops.shownAxisSize false 2 hidden, Ops.shownAxisSize false 4 hidden) ([ false; true; true; false ], 80.0, 100.0)
eq "hide: runs" (Ops.hiddenRuns false hidden, Ops.hiddenRuns true hidden) ([ 2, 4 ], [ 3, 5 ])
eq "hide: runs as text" (Ops.hiddenRuns false hidden |> List.map (Ops.axisRunText false), Ops.hiddenRuns true hidden |> List.map (Ops.axisRunText true)) ([ "C:E" ], [ "4:6" ])
eq "hide: hiding again keeps the kept size" ((hidden |> Ops.hideAxis false 2 2).Axes.HiddenColSizes) hidden.Axes.HiddenColSizes
eq "hide: the layout gives them no room" ((Layout.metricsOf hidden).ColPrefix.[2..5] |> Array.toList, (Layout.metricsOf hidden).RowPrefix.[3..6] |> Array.toList) ([ 180.0; 180.0; 180.0; 180.0 ], [ 75.0; 75.0; 75.0; 75.0 ])
let shown = hidden |> Ops.unhideAxis false 0 9 |> Ops.unhideAxis true 0 49
eq "unhide: everything is as before" (shown.ColumnWidths, shown.RowHeights, shown.Axes) (sized.ColumnWidths, sized.RowHeights, sized.Axes)
let partly = hidden |> Ops.unhideAxis false 3 3
eq "unhide: one of them" (partly.ColumnWidths, partly.Axes.HiddenColSizes, Ops.hiddenRuns false partly) (Map.ofList [ 1, 80.0; 2, 0.0; 3, 80.0; 4, 0.0 ], Map.ofList [ 2, 80.0 ], [ 2, 2; 4, 4 ])
eq "unhide: a whole-sheet span costs what is hidden" ((hidden |> Ops.unhideAxis true 0 1048575).RowHeights) sized.RowHeights
eq "size 0 hides" ((sized |> Ops.setAxisSize false 1 1 0.0).Axes.HiddenColSizes, Ops.hiddenRuns false (sized |> Ops.setAxisSize false 1 1 0.0)) (Map.ofList [ 1, 80.0 ], [ 1, 1 ])
let resized = hidden |> Ops.setAxisSize false 2 2 55.0
eq "a size shows a hidden column and forgets the kept one" (resized.ColumnWidths.[2], resized.Axes.HiddenColSizes) (55.0, Map.ofList [ 3, 80.0 ])
// A filter's hidden rows are its own.
let filtered = sheetOf [ at 0 0, text "h"; at 1 0, text "a"; at 2 0, text "b"; at 3 0, text "a" ] |> applyFilter 0 (Set.ofList [ "a" ]) |> applyFilterRowHeights
eq "filter: its rows are not hidden rows" (isFilterHiddenRow 2 filtered, Ops.hiddenRuns true filtered, (Ops.unhideAxis true 0 49 filtered).RowHeights) (true, [], filtered.RowHeights)

// ---- inserted and deleted rows / columns ---------------------------------------------------
let moved = hidden |> insertRow 4 |> insertColumn 0
eq "structure: hidden rows and their sizes move down" (Ops.hiddenRuns true moved, moved.Axes.HiddenRowSizes) ([ 3, 3; 5, 6 ], Map.ofList [ 5, 40.0 ])
eq "structure: hidden columns move right" (Ops.hiddenRuns false moved, moved.Axes.HiddenColSizes) ([ 3, 5 ], Map.ofList [ 3, 80.0; 4, 80.0 ])
let cut = hidden |> deleteRow 4 |> deleteColumn 2
eq "structure: a deleted hidden row / column is gone" (Ops.hiddenRuns true cut, cut.Axes.HiddenRowSizes, Ops.hiddenRuns false cut, cut.Axes.HiddenColSizes) ([ 3, 4 ], Map.empty, [ 2, 3 ], Map.ofList [ 2, 80.0 ])

// ---- auto-fit ---------------------------------------------------------------------------------
// Ten pixels a character at 11 px, in proportion for other sizes; bold is a tenth wider.
let measure : Ops.TextMeasure = fun text size _ bold _ -> float text.Length * 10.0 * size / 11.0 * (if bold then 1.1 else 1.0)
let content =
    sheetOf [ at 0 0, text "abcd"; at 1 0, text "abcdefgh"; at 2 0, num 1234.5 |> withFormat (fun f -> { f with NumberFormat = Some "#,##0.00" })
              at 0 1, text "abcd" |> withFormat (fun f -> { f with Bold = true; FontSize = Some 22 })
              at 0 2, text "abcdefghij" |> withFormat (fun f -> { f with Indent = 2 })
              at 0 3, text "one\ntwo three" |> withFormat (fun f -> { f with WrapText = true })
              at 0 4, text "merged text that is long"; at 3 5, text "turned" |> withFormat (fun f -> { f with TextRotation = 90 })
              at 5 6, text "" ]
    |> addMergedCell 0 4 0 5
let widths = Ops.autoFitColumnWidths measure content (0, 9) |> Map.ofList
eq "fit: the widest cell of a column, what it displays, plus the padding" widths.[0] 90.0          // "1,234.50" and "abcdefgh": 8 characters
eq "fit: in the cell's font (22 px bold)" widths.[1] (ceil (4.0 * 10.0 * 2.0 * 1.1) + 10.0)
eq "fit: with its indent" widths.[2] (100.0 + 10.0 + 2.0 * Rbt.Spreadsheet.CellText.mdwPx 11.0 * 1.4 |> ceil)
eq "fit: wrapped text counts with its longest line" widths.[3] 100.0
eq "fit: a cell merged across columns does not count; turned text is a line high" (Map.containsKey 4 widths, widths.[5]) (false, 30.0)
eq "fit: a column without text is left alone" (Map.containsKey 6 widths, Map.containsKey 7 widths) (false, false)
eq "fit: only the columns asked for" (Ops.autoFitColumnWidths measure content (1, 2) |> List.map fst) [ 1; 2 ]
eq "fit: hidden rows are left out" ((Ops.autoFitColumnWidths measure (content |> Ops.hideAxis true 1 2) (0, 0)) |> List.map snd) [ 50.0 ]
eq "fit: a hidden column stays hidden" (Ops.autoFitColumnWidths measure (content |> Ops.hideAxis false 0 0) (0, 1) |> List.map fst) [ 1 ]
let fitted = Ops.applyAxisSizes false (Ops.autoFitColumnWidths measure content (0, 3)) content
eq "fit: applied" ([ 0; 1; 2; 3 ] |> List.map (fun c -> Ops.axisSize false c fitted)) [ 90.0; widths.[1]; widths.[2]; 100.0 ]
eq "fit: a very long text is capped" (Ops.autoFitColumnWidths measure (sheetOf [ at 0 0, text (String.replicate 400 "x") ]) (0, 0)) [ 0, Ops.maxColumnWidthPx ]

eq "wrap: fits on a line" (Ops.wrappedLineCount (fun t -> float t.Length * 10.0) 100.0 "abc def") 1
eq "wrap: broken at spaces" (Ops.wrappedLineCount (fun t -> float t.Length * 10.0) 100.0 "abcd efgh ijkl mnop") 2
eq "wrap: a line break always breaks" (Ops.wrappedLineCount (fun t -> float t.Length * 10.0) 100.0 "a\nb\n\nc") 4
eq "wrap: a long word is broken where it must be" (Ops.wrappedLineCount (fun t -> float t.Length * 10.0) 100.0 "abcdefghijklmnopqrstuvwxy") 3
eq "wrap: a long word after a short one" (Ops.wrappedLineCount (fun t -> float t.Length * 10.0) 100.0 "ab abcdefghijklmnopqrstu z") 4

let tall =
    sheetOf [ at 0 0, text "abcd"; at 1 0, text "abcd efgh ijkl mnop qrst" |> withFormat (fun f -> { f with WrapText = true })
              at 2 0, text "abcd" |> withFormat (fun f -> { f with FontSize = Some 22 }); at 3 0, text "a\nb\nc"
              at 4 0, text "abcd efgh ijkl mnop qrst" |> withFormat (fun f -> { f with WrapText = true })
              at 6 0, text "tall" ; at 7 0, text "turned" |> withFormat (fun f -> { f with TextRotation = 90 }) ]
    |> addMergedCell 4 0 4 1 |> addMergedCell 6 0 7 0
    |> Ops.setAxisSize true 5 5 60.0
let heights = Ops.autoFitRowHeights measure tall (0, 8) |> Map.ofList
// (A line is as tall as Excel auto-fits a row to it: 20 px for 11 pt, 39 for 22; n lines n of them.)
eq "fit rows: one line of the default font" heights.[0] 20.0
eq "fit rows: wrapped text takes the lines it needs at the column's width" heights.[1] (3.0 * 20.0)      // 90 px for the text: "abcd efgh", "ijkl mnop", "qrst"
eq "fit rows: a larger font" heights.[2] (ceil (22.0 * 96.0 / 72.0 * 1.3))
eq "fit rows: line breaks count without wrapping" heights.[3] (3.0 * 20.0)
eq "fit rows: text wrapped in merged columns has their width" heights.[4] (2.0 * 20.0)                    // 190 px
eq "fit rows: a row without text gets the default; a cell merged across rows does not count" (heights.[5], heights.[6], heights.[8]) (20.0, 20.0, 20.0)
eq "fit rows: a narrower column, more lines" ((Ops.autoFitRowHeights measure (tall |> Ops.setAxisSize false 0 0 60.0) (1, 1))) [ 1, 5.0 * 20.0 ]
eq "fit rows: hidden rows stay hidden" (Ops.autoFitRowHeights measure (tall |> Ops.hideAxis true 1 2) (0, 3) |> List.map fst) [ 0; 3 ]
eq "fit rows: a whole-sheet span looks at the rows that have something" (Ops.autoFitRowHeights measure { tall with RowCount = 1000000 } (0, 999999) |> List.map fst) [ 0; 1; 2; 3; 4; 5 ]

// ---- outline --------------------------------------------------------------------------------
let grouped = empty |> Ops.changeOutlineLevel true 2 8 1 |> Ops.changeOutlineLevel true 4 5 1
eq "group: levels" (grouped.Axes.RowLevels |> Map.toList) [ 2, 1; 3, 1; 4, 2; 5, 2; 6, 1; 7, 1; 8, 1 ]
eq "group: the groups, outer first" (Ops.outlineGroups true grouped) [ 2, 8, 1, false; 4, 5, 2, false ]
eq "group: the group of a row at its own level" (Ops.outlineGroupAt true 3 grouped, Ops.outlineGroupAt true 4 grouped, Ops.outlineGroupAt true 9 grouped) (Some (2, 8, 1), Some (4, 5, 2), None)
eq "group: at most seven levels" ((List.fold (fun m _ -> Ops.changeOutlineLevel true 0 0 1 m) empty [ 1 .. 9 ]).Axes.RowLevels) (Map.ofList [ 0, 7 ])
eq "ungroup: one level down, level 0 is no entry" ((grouped |> Ops.changeOutlineLevel true 3 4 -1).Axes.RowLevels |> Map.toList) [ 2, 1; 4, 1; 5, 2; 6, 1; 7, 1; 8, 1 ]
eq "ungroup: rows that are in no group" (obj.ReferenceEquals(Ops.changeOutlineLevel true 20 30 -1 grouped, grouped)) true
let inner = grouped |> Ops.setGroupCollapsed "T" true 4 true
eq "collapse: the inner group is hidden, the row after it marked" (Ops.hiddenRuns true inner, inner.Axes.CollapsedRows |> Set.toList, Ops.outlineGroups true inner) ([ 4, 5 ], [ 6 ], [ 2, 8, 1, false; 4, 5, 2, true ])
let outer = inner |> Ops.setGroupCollapsed "T" true 2 true
eq "collapse: the outer group" (Ops.hiddenRuns true outer, outer.Axes.CollapsedRows |> Set.toList) ([ 2, 8 ], [ 6; 9 ])
let reopened = outer |> Ops.setGroupCollapsed "T" true 7 false
eq "expand: the outer group opens, the collapsed group inside stays shut" (Ops.hiddenRuns true reopened, reopened.Axes.CollapsedRows |> Set.toList) ([ 4, 5 ], [ 6 ])
eq "expand: then the inner one" ((reopened |> Ops.setGroupCollapsed "T" true 5 false).RowHeights, (reopened |> Ops.setGroupCollapsed "T" true 5 false).Axes.CollapsedRows) (Map.empty, Set.empty)
throws "collapse: a row in no group" (fun () -> Ops.setGroupCollapsed "SetGroupCollapsed" true 20 true grouped |> ignore) "SetGroupCollapsed: row 20 is not in a group"
let sizedGroup = empty |> Ops.setAxisSize true 3 3 44.0 |> Ops.changeOutlineLevel true 2 4 1 |> Ops.setGroupCollapsed "T" true 3 true |> Ops.setGroupCollapsed "T" true 3 false
eq "collapse and expand: a row keeps its height" sizedGroup.RowHeights (Map.ofList [ 3, 44.0 ])
eq "ungroup: the mark of a group that is gone goes too" ((inner |> Ops.changeOutlineLevel true 2 8 -1 |> Ops.changeOutlineLevel true 2 8 -1).Axes.CollapsedRows) Set.empty
let columns = empty |> Ops.changeOutlineLevel false 1 3 1 |> Ops.setGroupCollapsed "T" false 2 true
eq "columns: grouped and collapsed" (Ops.outlineGroups false columns, Ops.hiddenRuns false columns, columns.Axes.CollapsedCols |> Set.toList) ([ 1, 3, 1, true ], [ 1, 3 ], [ 4 ])
let shifted = inner |> insertRow 0 |> deleteRow 4
eq "structure: levels and marks follow the rows" (shifted.Axes.RowLevels |> Map.toList, shifted.Axes.CollapsedRows |> Set.toList, Ops.hiddenRuns true shifted) ([ 3, 1; 4, 2; 5, 2; 6, 1; 7, 1; 8, 1 ], [ 6 ], [ 4, 5 ])

printfn "\nRow and column tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
