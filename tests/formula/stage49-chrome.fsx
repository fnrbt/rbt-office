#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
module Ops = Rbt.Spreadsheet.SheetOps

// What the grid's chrome is drawn from: the outline with its summary rows
// below or above (columns right or left) and its level buttons.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s expected %A  got %A" label expected got)

let empty = createEmpty 50 10

// ---- the outline: where the summary is ----------------------------------------------------------
// Rows 3..9 at level 1, rows 5..6 at level 2 inside them.
let nested = empty |> Ops.changeOutlineLevel true 2 8 1 |> Ops.changeOutlineLevel true 4 5 1
eq "summary: below by default" (Ops.summaryAfter true nested, Ops.summaryAfter false nested, Ops.summaryIndex true 2 8 nested, Ops.summaryIndex false 1 3 nested) (true, true, 9, 4)
let above = nested |> Ops.setOutlineSummary false true
eq "summary: above" (Ops.summaryAfter true above, Ops.summaryIndex true 2 8 above, Ops.summaryIndex true 4 5 above, Ops.summaryIndex false 1 3 above) (false, 1, 3, 4)
let collapsedAbove = above |> Ops.setGroupCollapsed "T" true 4 true
eq "summary above: the mark is on the row before the group" (collapsedAbove.Axes.CollapsedRows |> Set.toList, Ops.hiddenRuns true collapsedAbove, Ops.outlineGroups true collapsedAbove) ([ 3 ], [ 4, 5 ], [ 2, 8, 1, false; 4, 5, 2, true ])
let outerAbove = collapsedAbove |> Ops.setGroupCollapsed "T" true 2 true
eq "summary above: the outer group" (outerAbove.Axes.CollapsedRows |> Set.toList, Ops.hiddenRuns true outerAbove) ([ 1; 3 ], [ 2, 8 ])
let reopened = outerAbove |> Ops.setGroupCollapsed "T" true 2 false
eq "summary above: expanding the outer group leaves the inner one collapsed" (reopened.Axes.CollapsedRows |> Set.toList, Ops.hiddenRuns true reopened) ([ 3 ], [ 4, 5 ])
eq "summary above: ungrouping drops the mark" ((collapsedAbove |> Ops.changeOutlineLevel true 2 8 -1 |> Ops.changeOutlineLevel true 2 8 -1).Axes.CollapsedRows) Set.empty
let atTop = empty |> Ops.setOutlineSummary false false |> Ops.changeOutlineLevel true 0 2 1 |> Ops.setGroupCollapsed "T" true 1 true
eq "summary above: a group at the first row has its mark off the sheet" (atTop.Axes.CollapsedRows |> Set.toList, Ops.outlineGroups true atTop) ([ -1 ], [ 0, 2, 1, true ])
// Moving the summaries keeps what is collapsed.
let moved = collapsedAbove |> Ops.setOutlineSummary true true
eq "summary: moved below: the marks move, the groups stay collapsed" (moved.Axes.CollapsedRows |> Set.toList, Ops.outlineGroups true moved, Ops.hiddenRuns true moved) ([ 6 ], [ 2, 8, 1, false; 4, 5, 2, true ], [ 4, 5 ])
eq "summary: the same setting again is the same model" (obj.ReferenceEquals(Ops.setOutlineSummary true true moved, moved)) true
let leftCols = empty |> Ops.setOutlineSummary true false |> Ops.changeOutlineLevel false 2 4 1 |> Ops.setGroupCollapsed "T" false 3 true
eq "summary left: columns" (leftCols.Axes.CollapsedCols |> Set.toList, Ops.hiddenRuns false leftCols, Ops.outlineGroups false leftCols) ([ 1 ], [ 2, 4 ], [ 2, 4, 1, true ])

// ---- the level buttons ---------------------------------------------------------------------------
eq "depth" (Ops.outlineDepth true nested, Ops.outlineDepth false nested, Ops.outlineDepth true empty) (2, 0, 0)
let level1 = nested |> Ops.showOutlineLevel true 1
eq "level 1: every group collapsed" (Ops.hiddenRuns true level1, level1.Axes.CollapsedRows |> Set.toList, Ops.outlineGroups true level1 |> List.map (fun (_, _, _, collapsed) -> collapsed)) ([ 2, 8 ], [ 6; 9 ], [ true; true ])
let level2 = level1 |> Ops.showOutlineLevel true 2
eq "level 2: the outer group shown, the inner one collapsed" (Ops.hiddenRuns true level2, level2.Axes.CollapsedRows |> Set.toList) ([ 4, 5 ], [ 6 ])
let level3 = level2 |> Ops.showOutlineLevel true 3
eq "level 3: everything shown" (Ops.hiddenRuns true level3, level3.Axes.CollapsedRows |> Set.toList) ([], [])
eq "level: rows keep their heights" ((empty |> Ops.setAxisSize true 4 4 44.0 |> Ops.changeOutlineLevel true 2 8 1 |> Ops.showOutlineLevel true 1 |> Ops.showOutlineLevel true 2).RowHeights) (Map.ofList [ 4, 44.0 ])
eq "level: no groups, nothing happens" (obj.ReferenceEquals(Ops.showOutlineLevel true 1 empty, empty)) true
eq "level: with the summary above" ((above |> Ops.showOutlineLevel true 2).Axes.CollapsedRows |> Set.toList, Ops.hiddenRuns true (above |> Ops.showOutlineLevel true 2)) ([ 3 ], [ 4, 5 ])
let colLevels = empty |> Ops.changeOutlineLevel false 1 3 1 |> Ops.showOutlineLevel false 1
eq "level: columns" (Ops.hiddenRuns false colLevels, colLevels.Axes.CollapsedCols |> Set.toList) ([ 1, 3 ], [ 4 ])

// ---- spilled arrays: the range a dynamic array fills ----------------------------------------------
let at row col : CellAddress = { Row = row; Col = col }
let spilling = { empty with Spill = Map.ofList [ at 1 2, at 1 1; at 2 1, at 1 1; at 2 2, at 1 1; at 3 1, at 1 1; at 3 2, at 1 1; at 6 5, at 5 5 ] }
let b2c4 : CellRange = { From = at 1 1; To = at 3 2 }
eq "spill: the anchor" (Ops.spillRangeAt (at 1 1) spilling) (Some (at 1 1, b2c4))
eq "spill: a cell it filled" (Ops.spillRangeAt (at 3 2) spilling) (Some (at 1 1, b2c4))
eq "spill: another array of the sheet" (Ops.spillRangeAt (at 6 5) spilling |> Option.map (fun (anchor, range) -> anchor, Ops.rangeText range)) (Some (at 5 5, "F6:F7"))
eq "spill: a cell beside it" (Ops.spillRangeAt (at 4 1) spilling, Ops.spillRangeAt (at 0 0) spilling) (None, None)
eq "spill: a sheet without arrays" (Ops.spillRangeAt (at 1 1) empty) None
let respilled = { spilling with Spill = Map.ofList [ at 1 2, at 1 1 ] }
eq "spill: worked out again when the arrays change" (Ops.spillRangeAt (at 1 1) respilled |> Option.map (fun (_, range) -> Ops.rangeText range), Ops.spillRangeAt (at 3 2) respilled) (Some "B2:C2", None)
// Addresses order by row, then column, and are equal by value (the comparison is written out).
eq "addresses: ordered by row, then column" ([ at 2 1; at 1 5; at 1 2; at 0 9 ] |> List.sort) [ at 0 9; at 1 2; at 1 5; at 2 1 ]
eq "addresses: equal by value, as keys too" (at 3 4 = at 3 4, at 3 4 = at 4 3, (Map.ofList [ at 3 4, "x" ]).[at 3 4], (dict [ at 3 4, "y" ]).[at 3 4], compare (at 1 1) (at 1 1)) (true, false, "x", "y", 0)

printfn "\nGrid chrome tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
