#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
module Ops = Rbt.Spreadsheet.SheetOps
module Layout = Rbt.Spreadsheet.Layout
module Input = Rbt.Spreadsheet.Input
module Formula = Rbt.Spreadsheet.Formula

// Sorting a range (order of values, ties, several keys, colours, a header
// row, formulas that move) and the sheet's AutoFilter: every kind of column
// filter, the rows it hides and when they are worked out, and what an edit
// of a filtered selection (copy, clear, fill) leaves alone.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-44s expected %A  got %A" label expected got)
let throws label (run: unit -> unit) (part: string) =
    try
        run ()
        fail <- fail + 1
        printfn "FAIL %-44s did not throw" label
    with error ->
        if error.Message.Contains part then pass <- pass + 1
        else (fail <- fail + 1; printfn "FAIL %-44s threw %s" label error.Message)

// (The formula module installs what Model needs to re-point and display.)
Formula.formulaHooksInstalled |> ignore

let at row col : CellAddress = { Row = row; Col = col }
let cell (value: CellValue) = { Value = value; Format = defaultFormat; Comment = None }
let text (value: string) = cell (CellValue.String value)
let num (n: float) = cell (CellValue.Number n)
let formula (f: string) = cell (CellValue.Formula (f, CellValue.Empty))
let withFormat (change: CellFormat -> CellFormat) (data: CellData) = { data with Format = change data.Format }
let sheetOf cells = cells |> List.fold (fun m (addr, data) -> setCell addr data m) (createEmpty 60 10)
let range a b c d : CellRange = { From = at a b; To = at c d }
let key col : Ops.SortKey = { Column = col; Descending = false; By = Ops.SortByValue; Color = "" }
let plain : Ops.SortOptions = { HasHeader = false; CaseSensitive = false }
let header = { plain with HasHeader = true }
let column col (m: SpreadsheetModel) rows = [ for row in rows -> match getCell (at row col) m with Some data -> displayValue data.Value | None -> "" ]
let formulaAt row col (m: SpreadsheetModel) = match getCell (at row col) m with Some { Value = CellValue.Formula (f, _) } -> f | _ -> ""

// ---- sort: the order of values -----------------------------------------------------------
let mixed =
    sheetOf [ at 0 0, text "pear"; at 1 0, num 3.0; at 2 0, cell (CellValue.Boolean true); at 3 0, cell (CellValue.Error "N/A")
              at 5 0, text "Apple"; at 6 0, num -1.0; at 7 0, cell (CellValue.Boolean false); at 8 0, text "apple"
              // a second column that travels with the rows
              at 0 1, num 0.0; at 1 1, num 1.0; at 2 1, num 2.0; at 3 1, num 3.0; at 4 1, num 4.0; at 5 1, num 5.0; at 6 1, num 6.0; at 7 1, num 7.0; at 8 1, num 8.0 ]
let ascending = Ops.sortRange "t" (range 0 0 8 1) [ key 0 ] plain mixed
eq "sort: numbers, text, FALSE, TRUE, errors, blank" (column 0 ascending [ 0 .. 8 ]) [ "-1"; "3"; "Apple"; "apple"; "pear"; "FALSE"; "TRUE"; "#N/A"; "" ]
eq "sort: the other cells of the rows come along" (column 1 ascending [ 0 .. 8 ]) [ "6"; "1"; "5"; "8"; "0"; "7"; "2"; "3"; "4" ]
let descending = Ops.sortRange "t" (range 0 0 8 1) [ { key 0 with Descending = true } ] plain mixed
eq "sort: descending reverses, blanks stay last" (column 0 descending [ 0 .. 8 ]) [ "#N/A"; "TRUE"; "FALSE"; "pear"; "Apple"; "apple"; "3"; "-1"; "" ]
let cased = Ops.sortRange "t" (range 0 0 8 1) [ key 0 ] { plain with CaseSensitive = true } mixed
eq "sort: case sensitive puts lower case first" (column 0 cased [ 2; 3 ]) [ "apple"; "Apple" ]
eq "sort: sorted already is the same cells" (obj.ReferenceEquals((Ops.sortRange "t" (range 0 0 8 1) [ key 0 ] plain ascending).Cells, ascending.Cells)) true
eq "sort: sets the sort state" (ascending.SortState, descending.SortState) (Some { Column = 0; Ascending = true }, Some { Column = 0; Ascending = false })

// ---- sort: stable, several keys, a header row, cells outside the range ---------------------
let people =
    sheetOf [ at 0 0, text "Team"; at 0 1, text "Name"; at 0 2, text "Score"; at 0 4, text "beside"
              at 1 0, text "red"; at 1 1, text "Zoe"; at 1 2, num 7.0; at 1 4, text "r1"
              at 2 0, text "blue"; at 2 1, text "Yan"; at 2 2, num 9.0; at 2 4, text "r2"
              at 3 0, text "red"; at 3 1, text "Abe"; at 3 2, num 7.0; at 3 4, text "r3"
              at 4 0, text "blue"; at 4 1, text "Bo"; at 4 2, num 3.0; at 4 4, text "r4"
              at 5 0, text "below"; at 5 1, text "x" ]
let byTeam = Ops.sortRange "t" (range 0 0 4 2) [ key 0 ] header people
eq "sort: the header row stays" (column 0 byTeam [ 0 .. 4 ]) [ "Team"; "blue"; "blue"; "red"; "red" ]
eq "sort: ties keep their order" (column 1 byTeam [ 1 .. 4 ]) [ "Yan"; "Bo"; "Zoe"; "Abe" ]
eq "sort: cells beside and below the range stay" (column 4 byTeam [ 1 .. 4 ], column 0 byTeam [ 5 ]) ([ "r1"; "r2"; "r3"; "r4" ], [ "below" ])
let twoKeys = Ops.sortRange "t" (range 0 0 4 2) [ key 0; { key 2 with Descending = true }; key 1 ] header people
eq "sort: three keys" (column 1 twoKeys [ 1 .. 4 ]) [ "Yan"; "Bo"; "Abe"; "Zoe" ]
let withoutHeader = Ops.sortRange "t" (range 0 0 4 2) [ key 0 ] plain people
eq "sort: without a header the first row sorts too" (column 0 withoutHeader [ 0 .. 4 ]) [ "blue"; "blue"; "red"; "red"; "Team" ]
let wholeColumns = Ops.sortRange "t" (range 0 0 Ops.maxSheetRow 2) [ key 2 ] header (people |> clearCell (at 5 0) |> clearCell (at 5 1))
eq "sort: whole columns are clipped to the used rows" (column 1 wholeColumns [ 1 .. 4 ]) [ "Bo"; "Zoe"; "Abe"; "Yan" ]

// ---- sort: by colour -------------------------------------------------------------------------
let coloured =
    sheetOf [ at 0 0, text "a"; at 1 0, text "b" |> withFormat (fun f -> { f with BgColor = Some "#ff0000" }); at 2 0, text "c" |> withFormat (fun f -> { f with TextColor = Some "#0000FF" })
              at 3 0, text "d" |> withFormat (fun f -> { f with BgColor = Some "#FF0000" }); at 4 0, text "e" ]
let fillFirst = Ops.sortRange "t" (range 0 0 4 0) [ { key 0 with By = Ops.SortByCellColor; Color = "#FF0000" } ] plain coloured
eq "sort: a fill colour first, the rest in order" (column 0 fillFirst [ 0 .. 4 ]) [ "b"; "d"; "a"; "c"; "e" ]
let fontLast = Ops.sortRange "t" (range 0 0 4 0) [ { key 0 with By = Ops.SortByFontColor; Color = "#0000ff"; Descending = true } ] plain coloured
eq "sort: a text colour last" (column 0 fontLast [ 0 .. 4 ]) [ "a"; "b"; "d"; "e"; "c" ]
eq "sort: formats move with the cells" ((getCell (at 0 0) fillFirst).Value.Format.BgColor) (Some "#ff0000")
let noFill = Ops.sortRange "t" (range 0 0 4 0) [ { key 0 with By = Ops.SortByCellColor; Color = "" } ] plain coloured
eq "sort: no fill first" (column 0 noFill [ 0 .. 4 ]) [ "a"; "c"; "e"; "b"; "d" ]

// ---- sort: formulas, comments, rich text, hyperlinks, what stays -----------------------------
let calc =
    { sheetOf [ at 0 0, num 30.0; at 0 1, formula "=A1*2"; at 0 2, formula "=A1+$A$3+A$1+Other!A1"
                at 1 0, num 10.0; at 1 1, formula "=A2*2"; at 1 2, formula "=SUM(A1:A2)"
                at 2 0, num 20.0; at 2 1, formula "=A3*2"; at 2 2, text "plain"
                at 0 5, formula "=A1"; at 4 0, formula "=SUM(A1:A3)" ] with
        RichRuns = Map.ofList [ at 2 2, [ "plain", emptyRunOverride ] ]
        Threads = Map.ofList [ at 0 0, { Id = "t1"; Author = None; PersonId = None; Date = None; Text = "thirty"; Replies = []; Done = false } ]
        Hyperlinks = [ { Range = range 1 0 1 0; Target = "https://example.com"; IsInternal = false; Tooltip = None; Display = None }; { Range = range 0 0 2 0; Target = "Sheet2!A1"; IsInternal = true; Tooltip = None; Display = None } ]
        RowHeights = Map.ofList [ 0, 40.0 ] }
    |> setCell (at 1 0) { num 10.0 with Comment = Some { Text = "ten"; Author = None; Runs = [] } }
let sortedCalc = Ops.sortRange "t" (range 0 0 2 2) [ key 0 ] plain calc
eq "sort: values in order" (column 0 sortedCalc [ 0 .. 2 ]) [ "10"; "20"; "30" ]
eq "sort: a formula keeps pointing at its own row" ([ 0 .. 2 ] |> List.map (fun row -> formulaAt row 1 sortedCalc)) [ "=A1*2"; "=A2*2"; "=A3*2" ]
eq "sort: relative rows shift, locked rows and other sheets as a copy does" (formulaAt 2 2 sortedCalc) "=A3+$A$3+A$1+Other!A3"
eq "sort: a range reference moves like a copy (off the sheet is #REF!)" (formulaAt 0 2 sortedCalc) "=SUM(#REF!)"
eq "sort: formulas outside the range are not re-pointed" (formulaAt 0 5 sortedCalc, formulaAt 4 0 sortedCalc) ("=A1", "=SUM(A1:A3)")
eq "sort: the comment moves with its cell" ((getCell (at 0 0) sortedCalc).Value.Comment |> Option.map (fun c -> c.Text)) (Some "ten")
eq "sort: rich text and threads move" (sortedCalc.RichRuns |> Map.toList |> List.map fst, sortedCalc.Threads |> Map.toList |> List.map fst) ([ at 1 2 ], [ at 2 0 ])
eq "sort: a one-cell hyperlink moves, a range one stays" (sortedCalc.Hyperlinks |> List.map (fun link -> link.Range)) [ range 0 0 0 0; range 0 0 2 0 ]
eq "sort: row heights stay" sortedCalc.RowHeights (Map.ofList [ 0, 40.0 ])

// ---- sort: what is refused -------------------------------------------------------------------
throws "sort: no key" (fun () -> Ops.sortRange "SortRange" (range 0 0 4 2) [] plain people |> ignore) "SortRange: no sort key"
throws "sort: a key outside the range" (fun () -> Ops.sortRange "SortRange" (range 0 0 4 2) [ key 4 ] plain people |> ignore) "key column E is outside the range A1:C5"
throws "sort: merged cells" (fun () -> Ops.sortRange "SortRange" (range 0 0 4 2) [ key 0 ] header (addMergedCell 2 0 3 0 people) |> ignore) "merged cells"
eq "sort: a merge in the header row is no obstacle" (Ops.sortProblem (range 0 0 4 2) [ key 0 ] header (addMergedCell 0 0 0 1 people)) None
throws "sort: a dynamic array" (fun () -> Ops.sortRange "SortRange" (range 0 0 4 2) [ key 0 ] header { people with Spill = Map.ofList [ at 3 1, at 2 1 ] } |> ignore) "dynamic array"

// ---- filter: wildcards ----------------------------------------------------------------------
eq "wildcards" ([ wildcardMatch "a*" "Apple"; wildcardMatch "*pl?" "apple"; wildcardMatch "a?c" "abcd"; wildcardMatch "*" ""; wildcardMatch "a~*b" "a*b"; wildcardMatch "a~*b" "axb"; wildcardMatch "*~?" "why?"; wildcardMatch "" "x" ])
                [ true; true; false; true; true; false; true; false ]
eq "wildcards: escape" (escapeWildcards "50% *off* ~?") "50% ~*off~* ~~~?"

// ---- filter: the AutoFilter and each kind of column filter -----------------------------------
// Row 0 is the header; rows 1..8 the data; row 9 is empty; row 10 lies below the range.
let data =
    sheetOf [ at 0 0, text "Region"; at 0 1, text "Sales"; at 0 2, text "Note"
              at 1 0, text "East"; at 1 1, num 10.0; at 1 2, text "first order"
              at 2 0, text "West"; at 2 1, num 20.0; at 2 2, text "Second"
              at 3 0, text "east"; at 3 1, num 30.0
              at 4 0, text "North"; at 4 1, num 40.0; at 4 2, text "50% off"
              at 5 0, text "West"; at 5 1, num 50.0; at 5 2, text "last order"
              at 6 1, num 60.0; at 6 2, text "no region"
              at 7 0, text "South"; at 7 1, text "n/a"; at 7 2, text "what?"
              at 8 0, text "East"; at 8 1, num 20.0 |> withFormat (fun f -> { f with BgColor = Some "#FFFF00"; NumberFormat = Some "0.00" })
              at 10 0, text "West"; at 10 1, num 1.0 ]
let filtered = data |> setAutoFilter (Some (range 0 0 8 2))
let hiddenBy col criteria = (filtered |> setColumnFilter col criteria).FilterHiddenRows |> List.ofArray
eq "filter: an AutoFilter alone hides nothing" (filtered.AutoFilter, filtered.FilterHiddenRows) (Some { Range = range 0 0 8 2; Implicit = false }, [||])
eq "filter: the rows it covers" (filterDataRows filtered) (Some (1, 8))
eq "filter: values (case-insensitive)" (hiddenBy 0 (FilterValues (Set.ofList [ "East" ], false))) [ 2; 4; 5; 6; 7 ]
eq "filter: values and blanks" (hiddenBy 0 (FilterValues (Set.ofList [ "west" ], true))) [ 1; 3; 4; 7; 8 ]
eq "filter: only blanks" (hiddenBy 0 (FilterValues (Set.empty, true))) [ 1; 2; 3; 4; 5; 7; 8 ]
eq "filter: values match the text shown" (hiddenBy 1 (FilterValues (Set.ofList [ "20.00" ], false))) [ 1; 2; 3; 4; 5; 6; 7 ]
eq "filter: values match the plain number too" (hiddenBy 1 (FilterValues (Set.ofList [ "20" ], false))) [ 1; 3; 4; 5; 6; 7 ]
eq "filter: greater than" (hiddenBy 1 (FilterCustom ([ OpGreaterThan, "30" ], true))) [ 1; 2; 3; 7; 8 ]
eq "filter: between, as two conditions" (hiddenBy 1 (FilterCustom ([ OpGreaterOrEqual, "20"; OpLessOrEqual, "40" ], true))) [ 1; 5; 6; 7 ]
eq "filter: either of two" (hiddenBy 1 (FilterCustom ([ OpLessThan, "20"; OpGreaterThan, "50" ], false))) [ 2; 3; 4; 5; 7; 8 ]
eq "filter: equals a number" (hiddenBy 1 (FilterCustom ([ OpEqual, "20" ], true))) [ 1; 3; 4; 5; 6; 7 ]
eq "filter: not equal" (hiddenBy 0 (FilterCustom ([ OpNotEqual, "west" ], true))) [ 2; 5 ]
eq "filter: contains" (hiddenBy 2 (FilterCustom ([ OpEqual, "*order*" ], true))) [ 2; 3; 4; 6; 7; 8 ]
eq "filter: does not contain" (hiddenBy 2 (FilterCustom ([ OpNotEqual, "*order*" ], true))) [ 1; 5 ]
eq "filter: begins with" (hiddenBy 2 (FilterCustom ([ OpEqual, "s*" ], true))) [ 1; 3; 4; 5; 6; 7; 8 ]
eq "filter: ends with an escaped ?" (hiddenBy 2 (FilterCustom ([ OpEqual, "*~?" ], true))) [ 1; 2; 3; 4; 5; 6; 8 ]
eq "filter: text greater than" (hiddenBy 0 (FilterCustom ([ OpGreaterOrEqual, "south" ], true))) [ 1; 3; 4; 6; 8 ]
eq "filter: equals nothing is the blanks" (hiddenBy 0 (FilterCustom ([ OpEqual, "" ], true))) [ 1; 2; 3; 4; 5; 7; 8 ]
eq "filter: top 2" (hiddenBy 1 (FilterTop (2, false, false))) [ 1; 2; 3; 4; 7; 8 ]
eq "filter: bottom 2 takes the ties" (hiddenBy 1 (FilterTop (2, false, true))) [ 3; 4; 5; 6; 7 ]
eq "filter: top 30 percent of seven numbers" (hiddenBy 1 (FilterTop (30, true, false))) [ 1; 2; 3; 4; 7; 8 ]
eq "filter: above average" (hiddenBy 1 (FilterDynamic "aboveAverage")) [ 1; 2; 3; 7; 8 ]
eq "filter: below average" (hiddenBy 1 (FilterDynamic "belowAverage")) [ 4; 5; 6; 7 ]
eq "filter: fill colour" (hiddenBy 1 (FilterColor ("#ffff00", false))) [ 1; 2; 3; 4; 5; 6; 7 ]
eq "filter: no fill" (hiddenBy 1 (FilterColor ("", false))) [ 8 ]
eq "filter: a kept filter hides nothing by itself" (hiddenBy 1 (FilterKept "<filters/>")) []
let both = filtered |> setColumnFilter 0 (FilterValues (Set.ofList [ "East"; "West" ], false)) |> setColumnFilter 1 (FilterCustom ([ OpGreaterOrEqual, "20" ], true))
eq "filter: two columns, a row has to pass both" (List.ofArray both.FilterHiddenRows) [ 1; 4; 6; 7 ]
eq "filter: rows outside the range are never hidden" (isFilterHiddenRow 10 both, isFilterHiddenRow 0 both, isFilterHiddenRow 1 both) (false, false, true)
eq "filter: one column's filter off" (List.ofArray (clearColumnFilter (Some 0) both).FilterHiddenRows, (clearColumnFilter (Some 0) both).AutoFilter.IsSome) ([ 1; 7 ], true)
eq "filter: every column's filter off keeps the range" ((clearColumnFilter None both).FilterHiddenRows, (clearColumnFilter None both).Filters.IsEmpty, (clearColumnFilter None both).AutoFilter) ([||], true, filtered.AutoFilter)
eq "filter: removed" ((setAutoFilter None both).AutoFilter, (setAutoFilter None both).Filters.IsEmpty, (setAutoFilter None both).FilterHiddenRows) (None, true, [||])
eq "filter: a narrower range drops the filters outside it" ((setAutoFilter (Some (range 0 1 8 2)) both).Filters |> Map.toList |> List.map fst) [ 1 ]
eq "filter: the checklist of a column" (filterColumnValues 0 filtered, filterColumnValues 1 filtered) ((([ "East"; "North"; "South"; "West" ]), true), ([ "10"; "20"; "20.00"; "30"; "40"; "50"; "60"; "n/a" ], false))
eq "filter: hidden rows have no height" (getRowHeight 1 both, getRowHeight 2 both, (Layout.metricsOf both).RowPrefix.[3]) (0.0, 25.0, 50.0)

// ---- filter: when the rows are worked out ----------------------------------------------------
let east = filtered |> setColumnFilter 0 (FilterValues (Set.ofList [ "East" ], false))
let edited = east |> setCell (at 1 0) (text "West") |> setCell (at 2 0) (text "East")
eq "filter: an edit neither hides nor shows a row" (edited.FilterHiddenRows) east.FilterHiddenRows
eq "filter: until it is applied again" (List.ofArray (reapplyFilter edited).FilterHiddenRows) [ 1; 4; 5; 6; 7 ]
eq "filter: applying again with nothing changed is the same model" (obj.ReferenceEquals(reapplyFilter east, east)) true
let inserted = east |> insertRow 3
eq "filter: hidden rows and the range follow inserted rows" (List.ofArray inserted.FilterHiddenRows, inserted.AutoFilter |> Option.map (fun f -> f.Range)) ([ 2; 5; 6; 7; 8 ], Some (range 0 0 9 2))
eq "filter: an inserted row is not hidden" (isFilterHiddenRow 3 inserted) false
let removed = east |> deleteRow 2
eq "filter: ... and deleted rows" (List.ofArray removed.FilterHiddenRows, removed.AutoFilter |> Option.map (fun f -> f.Range)) ([ 3; 4; 5; 6 ], Some (range 0 0 7 2))
let movedColumns = both |> insertColumn 0
eq "filter: column filters and the range follow inserted columns" (movedColumns.Filters |> Map.toList |> List.map fst, movedColumns.AutoFilter |> Option.map (fun f -> f.Range), List.ofArray movedColumns.FilterHiddenRows) ([ 1; 2 ], Some (range 0 1 8 3), [ 1; 4; 6; 7 ])
let lostColumn = both |> deleteColumn 0
eq "filter: a deleted column takes its filter along" (lostColumn.Filters |> Map.toList |> List.map fst, List.ofArray lostColumn.FilterHiddenRows) ([ 0 ], [ 1; 7 ])
eq "filter: deleting the whole range removes the filter" ((both |> applyStructureEdit (DeleteRows (0, 9))).AutoFilter, (both |> applyStructureEdit (DeleteRows (0, 9))).Filters.IsEmpty) (None, true)

// ---- filter: without a range (the grid's own filter buttons) ---------------------------------
let quick = data |> applyFilter 0 (Set.ofList [ "West" ])
eq "implicit: the block of data around the column" (quick.AutoFilter) (Some { Range = range 0 0 8 2; Implicit = true })
eq "implicit: rows hidden" (List.ofArray quick.FilterHiddenRows) [ 1; 3; 4; 6; 7; 8 ]
eq "implicit: gone with its last column filter" ((removeFilter 0 quick).AutoFilter, (removeFilter 0 quick).FilterHiddenRows) (None, [||])
eq "implicit: an explicit range stays" ((filtered |> applyFilter 0 (Set.ofList [ "West" ]) |> removeFilter 0).AutoFilter) filtered.AutoFilter

// ---- filtered selections: copy, clear, fill, sort ---------------------------------------------
let copied = serializeSelection (Range (1, 0, 8, 1)) east false
eq "copy: only the rows shown, one after the other" (copied.Cells |> List.map (fun (r, c, d) -> r, c, displayValue d.Value), copied.SourceEndRow) ([ 0, 0, "East"; 0, 1, "10"; 1, 0, "east"; 1, 1, "30"; 2, 0, "East"; 2, 1, "20" ], 3)
let cut = serializeSelection (Range (1, 0, 8, 0)) east true
eq "cut: the rows shown, at their places" (cut.Cells |> List.map (fun (r, _, _) -> r), cut.SourceEndRow) ([ 0; 2; 7 ], 8)
let cleared = clearSelectionCells (Range (1, 0, 8, 2)) east
eq "clear: hidden rows keep their cells" (column 0 cleared [ 1 .. 8 ]) [ ""; "West"; ""; "North"; "West"; ""; "South"; "" ]
let clearedFormats = Ops.clearCells Ops.ClearContents (1, 1, 8, 1) east
eq "clear contents: likewise" (column 1 clearedFormats [ 1 .. 8 ]) [ ""; "20"; ""; "40"; "50"; "60"; "n/a"; "" ]
let filled = Input.fillCells (1, 0, 1, 0) Input.FillDown 7 false east
eq "fill down: only the rows shown" (column 0 filled [ 1 .. 8 ]) [ "East"; "West"; "East"; "North"; "West"; ""; "South"; "East" ]
let sortedShown = Ops.sortRange "t" (range 0 0 8 2) [ { key 1 with Descending = true } ] header east
eq "sort: hidden rows stay where they are" (column 1 sortedShown [ 1 .. 8 ]) [ "30"; "20"; "20"; "40"; "50"; "60"; "n/a"; "10" ]
eq "sort: the filter's rows are unchanged" sortedShown.FilterHiddenRows east.FilterHiddenRows
eq "rows shown" (filterShownRows 0 9 east) [ 0; 1; 3; 8; 9 ]

// ---- criteria as plain data ----------------------------------------------------------------------
let roundTrip criteria = Ops.filterOfDescription "t" (Ops.describeFilter criteria)
for criteria in [ FilterValues (Set.ofList [ "a"; "b" ], true); FilterCustom ([ OpGreaterThan, "5"; OpLessOrEqual, "9" ], true); FilterCustom ([ OpEqual, "x*y" ], true)
                  FilterCustom ([ OpEqual, "*50% ~*off~**" ], true); FilterCustom ([ OpNotEqual, "*x*"; OpEqual, "*end" ], false); FilterCustom ([ OpEqual, "start*" ], true)
                  FilterTop (5, true, true); FilterColor ("#FF0000", true); FilterColor ("", false); FilterDynamic "belowAverage" ] do
    eq (sprintf "description round trip %A" criteria) (roundTrip criteria) criteria
let described = Ops.describeFilter (FilterCustom ([ OpEqual, "*50% ~*off~**"; OpNotEqual, "*x*" ], false))
eq "description: text operators" (described.Conditions |> List.map (fun c -> c.Operator, c.Value), described.And) ([ "contains", "50% *off*"; "notContains", "x" ], false)
eq "description: begins / ends / equals" ([ "ab*"; "*ab"; "a*b"; "ab" ] |> List.map (fun pattern -> (Ops.describeFilter (FilterCustom ([ OpEqual, pattern ], true))).Conditions.Head.Operator)) [ "beginsWith"; "endsWith"; "equals"; "equals" ]
let condition op value = { Ops.emptyFilterDescription "conditions" with Conditions = [ { Operator = op; Value = value } ] }
eq "description: a date becomes its serial" (Ops.filterOfDescription "t" (condition "greaterOrEqual" "2024-03-31")) (FilterCustom ([ OpGreaterOrEqual, "45382" ], true))
eq "description: values with an empty text are the blanks" (Ops.filterOfDescription "t" { Ops.emptyFilterDescription "values" with Values = [ "a"; "" ] }) (FilterValues (Set.ofList [ "a" ], true))
throws "description: unknown operator" (fun () -> Ops.filterOfDescription "SetFilter" (condition "like" "x") |> ignore) "SetFilter: unknown filter operator \"like\""
throws "description: no condition" (fun () -> Ops.filterOfDescription "SetFilter" (Ops.emptyFilterDescription "conditions") |> ignore) "at least one condition"
throws "description: three conditions" (fun () -> Ops.filterOfDescription "SetFilter" { condition "equals" "a" with Conditions = List.replicate 3 { Operator = "equals"; Value = "a" } } |> ignore) "at most two"
throws "description: contains nothing" (fun () -> Ops.filterOfDescription "SetFilter" (condition "contains" "") |> ignore) "needs a text"
throws "description: top 0" (fun () -> Ops.filterOfDescription "SetFilter" { Ops.emptyFilterDescription "top" with Count = 0 } |> ignore) "at least 1"
throws "description: 150 percent" (fun () -> Ops.filterOfDescription "SetFilter" { Ops.emptyFilterDescription "top" with Count = 150; Percent = true } |> ignore) "above 100"
throws "description: a bad colour" (fun () -> Ops.filterOfDescription "SetFilter" { Ops.emptyFilterDescription "color" with Color = "red" } |> ignore) "not a colour"
throws "description: an unknown dynamic filter" (fun () -> Ops.filterOfDescription "SetFilter" { Ops.emptyFilterDescription "dynamic" with Dynamic = "today" } |> ignore) "unknown dynamic filter"
throws "description: an unknown kind" (fun () -> Ops.filterOfDescription "SetFilter" (Ops.emptyFilterDescription "icons") |> ignore) "unknown filter kind"
throws "filter range outside the sheet" (fun () -> Ops.autoFilterRange "SetAutoFilter" (range 100 0 120 2) data |> ignore) "outside the sheet"
eq "filter range clipped to the sheet" (Ops.autoFilterRange "SetAutoFilter" (range 0 0 Ops.maxSheetRow 2) data) (range 0 0 59 2)

printfn "\nSort and filter tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
