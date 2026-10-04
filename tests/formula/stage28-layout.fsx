#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
module Layout = Rbt.Spreadsheet.Layout

// Layout.rowTop / columnLeft / visibleRange / hitTest read cached row and
// column offsets (Layout.metricsOf), and Model.findMergedCell an index of
// the merges by row. Both must give what the plain computation gives, and
// must notice every change of what they were computed from.

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)
let eq label (got: 'a) (expected: 'a) =
    ok label (got = expected) (sprintf "expected %A  got %A" expected got)

// The sum of the sizes before an index, computed the slow way.
let plainRowTop (model: SpreadsheetModel) row = Layout.prefixSizeUncached (fun r -> getRowHeight r model) row
let plainColumnLeft (model: SpreadsheetModel) col = Layout.prefixSizeUncached (fun c -> getColumnWidth c model) col
let sameOffsets label (model: SpreadsheetModel) =
    let rows = [ 0; 1; 2; 7; 8; 9; 49; 50; model.RowCount - 1; model.RowCount; model.RowCount + 3 ]
    let cols = [ 0; 1; 2; 3; 4; model.ColCount - 1; model.ColCount; model.ColCount + 2 ]
    eq (label + ": rowTop") (rows |> List.map (Layout.rowTop model)) (rows |> List.map (plainRowTop model))
    eq (label + ": columnLeft") (cols |> List.map (Layout.columnLeft model)) (cols |> List.map (plainColumnLeft model))

let sheet = createEmpty 200 12
sameOffsets "default sizes" sheet
eq "default: row 8 starts at 200" (Layout.rowTop sheet 8) 200.0
eq "default: column 3 starts at 300" (Layout.columnLeft sheet 3) 300.0

let sized = sheet |> setRowHeight 2 40.0 |> setRowHeight 8 0.0 |> setColumnWidth 1 30.0 |> setColumnWidth 3 0.0
sameOffsets "own sizes" sized
eq "a taller row pushes the rows below" (Layout.rowTop sized 3) 90.0
eq "a hidden row takes no space" (Layout.rowTop sized 9 - Layout.rowTop sized 8) 0.0
// The first sheet is still answered for itself (the cache is per set of sizes).
eq "the other sheet's offsets are its own" (Layout.rowTop sheet 3, Layout.columnLeft sheet 2) (75.0, 200.0)

// Every way sizes change is seen.
let resized = setRowHeight 2 60.0 sized
eq "a changed row height" (Layout.rowTop resized 3) 110.0
eq "and the model it came from is unchanged" (Layout.rowTop sized 3) 90.0
let widened = setColumnWidth 1 130.0 sized
eq "a changed column width" (Layout.columnLeft widened 2) 230.0
let inserted = insertRow 1 sized
sameOffsets "after a row insert" inserted
eq "an inserted row moves the tall row down" (Layout.rowTop inserted 4 - Layout.rowTop inserted 3) 40.0
let deleted = deleteColumn 1 sized
sameOffsets "after a column delete" deleted
eq "a deleted column takes its width with it" (Layout.columnLeft deleted 1) 100.0
let grown = { sized with RowCount = 400 }
sameOffsets "a larger sheet" grown
eq "its end" (Layout.rowTop grown 400) (plainRowTop grown 400)

// A filter hides rows: their height is 0 while it is active.
let withData =
    [ 0, "Kind"; 1, "a"; 2, "b"; 3, "a"; 4, "b"; 5, "a" ]
    |> List.fold (fun m (row, text) -> setCell { Row = row; Col = 0 } { Value = CellValue.String text; Format = defaultFormat; Comment = None } m) (createEmpty 50 4)
sameOffsets "before the filter" withData
let filtered = applyFilter 0 (Set.ofList [ "a" ]) withData |> applyFilterRowHeights
sameOffsets "filtered" filtered
ok "filtered rows take no space" (Layout.rowTop filtered 6 < Layout.rowTop withData 6) (sprintf "%f" (Layout.rowTop filtered 6))
let unfiltered = removeFilter 0 filtered
sameOffsets "filter removed" unfiltered

// Visible range and hit test agree with the offsets.
let viewport = { Layout.defaultViewport 800.0 400.0 with ScrollX = 250.0; ScrollY = 1000.0 }
let range = Layout.visibleRange sized viewport
ok "visible range: first row holds the scroll offset" (plainRowTop sized range.StartRow <= 1000.0 - 30.0 && plainRowTop sized (range.StartRow + 1) > 1000.0 - 30.0) (sprintf "%A" range)
ok "visible range: last row is past the bottom edge" (plainRowTop sized range.EndRowExclusive >= 1000.0 + 400.0 - 30.0 || range.EndRowExclusive = sized.RowCount) (sprintf "%A" range)
match Layout.hitTest sized viewport 300.0 200.0 with
| Layout.Cell addr ->
    let x = 300.0 + 250.0 - 50.0
    let y = 200.0 + 1000.0 - 30.0
    ok "hit test finds the cell under a point" (plainColumnLeft sized addr.Col <= x && x < plainColumnLeft sized (addr.Col + 1) && plainRowTop sized addr.Row <= y && y < plainRowTop sized (addr.Row + 1)) (sprintf "%A" addr)
| other -> ok "hit test finds a cell" false (sprintf "%A" other)
eq "rowAt / columnAt" (Layout.rowAt sized 95.0, Layout.columnAt sized 125.0, Layout.rowAt sized -1.0) (Some 3, Some 1, None)
eq "a point in a hidden column belongs to the next" (Layout.columnAt sized 230.0) (Some 4)

// A large sheet: offsets are found without walking the rows each time.
let large = { createEmpty 200000 30 with RowHeights = Map.ofList [ for r in 0 .. 1999 -> r * 100, 31.0 ] }
let watch = System.Diagnostics.Stopwatch.StartNew()
let mutable sum = 0.0
for i in 0 .. 49999 do
    sum <- sum + Layout.rowTop large (i * 4)
watch.Stop()
ok "50,000 row offsets of a 200,000 row sheet in well under a second" (watch.ElapsedMilliseconds < 1000L) (sprintf "%d ms" watch.ElapsedMilliseconds)
eq "and they are right" (Layout.rowTop large 199999) (plainRowTop large 199999)

// ---- merges ------------------------------------------------------------------
let plainFind (addr: CellAddress) (model: SpreadsheetModel) = model.MergedCells |> List.tryFind (fun merge -> isInMergedRange addr merge)
let merged =
    createEmpty 100 20
    |> addMergedCell 1 1 2 3
    |> addMergedCell 5 0 5 4
    |> addMergedCell 10 2 30 2
    |> addMergedCell 0 6 99 6          // a merged column
    |> addMergedCell 40 0 20000 0      // taller than the index lists row by row
let cells = [ for row in 0 .. 99 do for col in 0 .. 9 -> { Row = row; Col = col } ] @ [ { Row = 15000; Col = 0 }; { Row = 20001; Col = 0 } ]
eq "findMergedCell: same as a scan of the list" (cells |> List.map (fun addr -> findMergedCell addr merged)) (cells |> List.map (fun addr -> plainFind addr merged))
eq "a cell of a merge" (findMergedCell { Row = 2; Col = 3 } merged |> Option.map (fun m -> m.StartRow, m.StartCol)) (Some (1, 1))
eq "a cell next to it" (findMergedCell { Row = 2; Col = 4 } merged) None
eq "no merges" (findMergedCell { Row = 2; Col = 3 } (createEmpty 10 10)) None
let unmerged = removeMergedCell { Row = 1; Col = 1 } merged
eq "after an unmerge" (findMergedCell { Row = 2; Col = 3 } unmerged, (findMergedCell { Row = 5; Col = 2 } unmerged).IsSome) (None, true)
eq "the model it came from still has it" (findMergedCell { Row = 2; Col = 3 } merged).IsSome true
let shifted = insertRow 0 merged
eq "after a row insert" (cells |> List.map (fun addr -> findMergedCell addr shifted)) (cells |> List.map (fun addr -> plainFind addr shifted))
// Overlapping merges (a broken file): the first of the list wins, as with a scan.
let overlapping = createEmpty 50 10 |> addMergedCell 0 0 3 3 |> addMergedCell 2 2 5 5 |> addMergedCell 0 2 9000 2
let overlapCells = [ for row in 0 .. 8 do for col in 0 .. 6 -> { Row = row; Col = col } ]
eq "overlapping merges" (overlapCells |> List.map (fun addr -> findMergedCell addr overlapping)) (overlapCells |> List.map (fun addr -> plainFind addr overlapping))
// Many merges: a lookup does not scan them all.
let many = { createEmpty 20000 10 with MergedCells = [ for row in 0 .. 9999 -> { StartRow = row * 2; StartCol = 0; EndRow = row * 2; EndCol = 1 } ] }
let mergeWatch = System.Diagnostics.Stopwatch.StartNew()
let mutable found = 0
for row in 0 .. 19999 do
    if (findMergedCell { Row = row; Col = 1 } many).IsSome then found <- found + 1
mergeWatch.Stop()
eq "10,000 merges: every other row" found 10000
ok "20,000 lookups among 10,000 merges in well under a second" (mergeWatch.ElapsedMilliseconds < 1000L) (sprintf "%d ms" mergeWatch.ElapsedMilliseconds)

printfn ""
printfn "Stage 28 (layout metrics, merge index): %d passed, %d failed" pass fail
if fail > 0 then exit 1
