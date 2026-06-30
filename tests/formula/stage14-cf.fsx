#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// Conditional-format visuals computed over a range: data bars / color scale /
// icon set, baked into each cell's CellFormat (which the renderers paint).

open RedBat.Spreadsheet.Model

let cell v = { Value = v; Format = defaultFormat; Comment = None }
// A1:A5 = 0,25,50,75,100
let model =
    { createEmpty 20 8 with
        Cells = [ 0.;25.;50.;75.;100. ] |> List.mapi (fun i v -> { Row=i; Col=0 }, cell (CellValue.Number v)) |> Map.ofList }

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-22s expected %A got %A" label expected got
let allocatedBytesFor (f: unit -> unit) =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetAllocatedBytesForCurrentThread()
    f()
    System.GC.GetAllocatedBytesForCurrentThread() - before
let fmtAt r (m: SpreadsheetModel) = match getCell { Row=r; Col=0 } m with Some cd -> cd.Format | None -> defaultFormat

// ── data bars (0-baseline, all ≥ 0) ──
let db = applyDataBars 0 0 4 0 "#638EC6" model
check "databar min"  ((fmtAt 0 db).DataBar) (Some (0.0, "#638EC6"))
check "databar max"  ((fmtAt 4 db).DataBar) (Some (1.0, "#638EC6"))
check "databar mid"  ((fmtAt 2 db).DataBar) (Some (0.5, "#638EC6"))
let dbReapplied = applyDataBars 0 0 4 0 "#638EC6" db
check "databar reapply cells" (obj.ReferenceEquals(dbReapplied.Cells, db.Cells)) true
let dbSplit =
    model
    |> applyDataBars 0 0 1 0 "#638EC6"
    |> applyDataBars 2 0 4 0 "#638EC6"
applyDataBars 0 0 1 0 "#638EC6" dbSplit |> ignore
applyDataBars 2 0 4 0 "#638EC6" dbSplit |> ignore
check "databar alternating range no-op" (obj.ReferenceEquals(applyDataBars 0 0 1 0 "#638EC6" dbSplit, dbSplit)) true
let dbPartialStale =
    match getCell { Row=2; Col=0 } db with
    | Some cd -> { db with Cells = Map.add { Row=2; Col=0 } { cd with Format = { cd.Format with DataBar = None } } db.Cells }
    | None -> failwith "missing databar cell"
let dbPartialReapplied = applyDataBars 0 0 4 0 "#638EC6" dbPartialStale
check "databar partial repaired" ((fmtAt 2 dbPartialReapplied).DataBar) (Some (0.5, "#638EC6"))
let dbDifferentColor = applyDataBars 0 0 4 0 "#FF0000" db
check "databar cache color invalidates" ((fmtAt 4 dbDifferentColor).DataBar) (Some (1.0, "#FF0000"))
check "databar cache keeps first source" (obj.ReferenceEquals((applyDataBars 0 0 4 0 "#638EC6" db).Cells, db.Cells)) true
let wideCfLastRow = 511
let wideCfCells =
    [ for row in 0 .. wideCfLastRow do
        for col in 0 .. 15 ->
            { Row = row; Col = col }, cell (CellValue.Number (float row)) ]
    |> Map.ofList
let wideCfModel = { createEmpty (wideCfLastRow + 2) 20 with Cells = wideCfCells }
let mutable wideDataBarModel = wideCfModel
for col in 0 .. 15 do
    wideDataBarModel <- applyDataBars 0 col wideCfLastRow col "#638EC6" wideDataBarModel
applyDataBars 0 0 wideCfLastRow 0 "#638EC6" wideDataBarModel |> ignore
for col in 1 .. 15 do
    applyDataBars 0 col wideCfLastRow col "#638EC6" wideDataBarModel |> ignore
let wideDataBarNoOpBytes =
    allocatedBytesFor (fun () ->
        for _ in 1 .. 3 do
            for col in 0 .. 15 do
                applyDataBars 0 col wideCfLastRow col "#638EC6" wideDataBarModel |> ignore)
check "databar cache keeps many ranges" (wideDataBarNoOpBytes < 250000L) true
let cfManyModelCount = 24
let cfManyRowCount = 2000
let cfManyBaseModels =
    [| for modelIndex in 0 .. cfManyModelCount - 1 ->
        { createEmpty (cfManyRowCount + 1) 4 with
            Cells =
                [ for row in 0 .. cfManyRowCount - 1 ->
                    { Row = row; Col = 0 }, cell (CellValue.Number (float ((row + modelIndex) % 100))) ]
                |> Map.ofList } |]
let cfManyModels =
    [| for source in cfManyBaseModels -> applyDataBars 0 0 (cfManyRowCount - 1) 0 "#638EC6" source |]
let cfManyColorScaleModels =
    [| for source in cfManyBaseModels -> applyColorScale 0 0 (cfManyRowCount - 1) 0 "#FF0000" None "#00FF00" source |]
let cfManyIconSetModels =
    [| for source in cfManyBaseModels -> applyIconSet 0 0 (cfManyRowCount - 1) 0 3 source |]
let cfManyRun passes =
    let mutable hits = 0
    for _pass in 1 .. passes do
        for source in cfManyModels do
            let reapplied = applyDataBars 0 0 (cfManyRowCount - 1) 0 "#638EC6" source
            if obj.ReferenceEquals(reapplied.Cells, source.Cells) then
                hits <- hits + 1
        for source in cfManyColorScaleModels do
            let reapplied = applyColorScale 0 0 (cfManyRowCount - 1) 0 "#FF0000" None "#00FF00" source
            if obj.ReferenceEquals(reapplied.Cells, source.Cells) then
                hits <- hits + 1
        for source in cfManyIconSetModels do
            let reapplied = applyIconSet 0 0 (cfManyRowCount - 1) 0 3 source
            if obj.ReferenceEquals(reapplied.Cells, source.Cells) then
                hits <- hits + 1
    hits
cfManyRun 3 |> ignore
let cfManySw = System.Diagnostics.Stopwatch.StartNew()
let cfManyHits = cfManyRun 120
cfManySw.Stop()
check "cf visual cache many models"
    (cfManySw.Elapsed.TotalMilliseconds < 20.0 && cfManyHits = cfManyModelCount * 120 * 3)
    true
let cfNoNumericRows = 20000
let cfNoNumericCols = 4
let cfNoNumericModel =
    { createEmpty cfNoNumericRows cfNoNumericCols with
        Cells =
            [ for row in 0 .. cfNoNumericRows - 1 do
                for col in 0 .. cfNoNumericCols - 1 ->
                    { Row = row; Col = col }, cell (CellValue.String (sprintf "text-%d-%d" row col)) ]
            |> Map.ofList }
check "cf no-number databar identity"
    (obj.ReferenceEquals(applyDataBars 0 0 (cfNoNumericRows - 1) (cfNoNumericCols - 1) "#638EC6" cfNoNumericModel, cfNoNumericModel))
    true
check "cf no-number colorscale identity"
    (obj.ReferenceEquals(applyColorScale 0 0 (cfNoNumericRows - 1) (cfNoNumericCols - 1) "#FF0000" None "#00FF00" cfNoNumericModel, cfNoNumericModel))
    true
check "cf no-number iconset identity"
    (obj.ReferenceEquals(applyIconSet 0 0 (cfNoNumericRows - 1) (cfNoNumericCols - 1) 3 cfNoNumericModel, cfNoNumericModel))
    true
let cfNoNumericRun passes =
    let mutable hits = 0
    for _pass in 1 .. passes do
        if obj.ReferenceEquals(applyDataBars 0 0 (cfNoNumericRows - 1) (cfNoNumericCols - 1) "#638EC6" cfNoNumericModel, cfNoNumericModel) then
            hits <- hits + 1
        if obj.ReferenceEquals(applyColorScale 0 0 (cfNoNumericRows - 1) (cfNoNumericCols - 1) "#FF0000" None "#00FF00" cfNoNumericModel, cfNoNumericModel) then
            hits <- hits + 1
        if obj.ReferenceEquals(applyIconSet 0 0 (cfNoNumericRows - 1) (cfNoNumericCols - 1) 3 cfNoNumericModel, cfNoNumericModel) then
            hits <- hits + 1
    hits
cfNoNumericRun 3 |> ignore
let cfNoNumericSw = System.Diagnostics.Stopwatch.StartNew()
let cfNoNumericHits = cfNoNumericRun 120
cfNoNumericSw.Stop()
check "cf no-number cache perf"
    (cfNoNumericHits = 120 * 3 && cfNoNumericSw.Elapsed.TotalMilliseconds < 20.0)
    true
let cfApplyCacheRetainedRowCount = 10000
let cfApplyCacheRetainedModelCount = 80
let cfApplyCacheRetainedModel modelIndex =
    { createEmpty cfApplyCacheRetainedRowCount 4 with
        Cells =
            [ for row in 0 .. cfApplyCacheRetainedRowCount - 1 ->
                { Row = row; Col = modelIndex % 4 }, cell (CellValue.Number (float ((row + modelIndex) % 100))) ]
            |> Map.ofList }
let cfApplyCacheRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. cfApplyCacheRetainedModelCount - 1 do
        let col = modelIndex % 4
        applyDataBars 0 col (cfApplyCacheRetainedRowCount - 1) col "#638EC6" (cfApplyCacheRetainedModel modelIndex) |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "databar apply cache bounded source memory" (cfApplyCacheRetainedBytes < 120000000L) true
let cfApplyCacheSourceTextModel modelIndex =
    { createEmpty 4 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number (float modelIndex))
              { Row = 1; Col = 1 }, cell (CellValue.String (System.String(char (65 + (modelIndex % 26)), 2000000))) ]
            |> Map.ofList }
let cfApplyCacheSourceTextBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    for modelIndex in 0 .. 39 do
        applyDataBars 0 0 0 0 "#638EC6" (cfApplyCacheSourceTextModel modelIndex) |> ignore
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
check "databar apply cache source text memory" (cfApplyCacheSourceTextBytes < 30000000L) true
let cfApplyCacheOversizedSourceModel marker =
    let payload = System.String(marker, 6 * 1024 * 1024)
    { createEmpty 4 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 10.0)
              { Row = 1; Col = 1 }, cell (CellValue.String payload) ]
            |> Map.ofList }
let databarApplyOversizedSourceCellsWeakReference () =
    let applied = applyDataBars 0 0 0 0 "#638EC6" (cfApplyCacheOversizedSourceModel 'D')
    applyDataBars 0 0 0 0 "#638EC6" applied |> ignore
    System.WeakReference(applied.Cells)
let databarApplyOversizedSourceCellsWeak = databarApplyOversizedSourceCellsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "databar cache skips oversized source"
    databarApplyOversizedSourceCellsWeak.IsAlive
    false
let colorScaleApplyOversizedSourceCellsWeakReference () =
    let applied = applyColorScale 0 0 0 0 "#FF0000" None "#00FF00" (cfApplyCacheOversizedSourceModel 'C')
    applyColorScale 0 0 0 0 "#FF0000" None "#00FF00" applied |> ignore
    System.WeakReference(applied.Cells)
let colorScaleApplyOversizedSourceCellsWeak = colorScaleApplyOversizedSourceCellsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "colorscale cache skips oversized source"
    colorScaleApplyOversizedSourceCellsWeak.IsAlive
    false
let iconSetApplyOversizedSourceCellsWeakReference () =
    let applied = applyIconSet 0 0 0 0 3 (cfApplyCacheOversizedSourceModel 'I')
    applyIconSet 0 0 0 0 3 applied |> ignore
    System.WeakReference(applied.Cells)
let iconSetApplyOversizedSourceCellsWeak = iconSetApplyOversizedSourceCellsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "iconset cache skips oversized source"
    iconSetApplyOversizedSourceCellsWeak.IsAlive
    false

// ── 2-color scale (red→green) ──
let cs = applyColorScale 0 0 4 0 "#FF0000" None "#00FF00" model
check "colorscale min" ((fmtAt 0 cs).BgColor) (Some "#FF0000")
check "colorscale max" ((fmtAt 4 cs).BgColor) (Some "#00FF00")
check "colorscale mid" ((fmtAt 2 cs).BgColor) (Some "#7F7F00")   // halfway lerp
let csReapplied = applyColorScale 0 0 4 0 "#FF0000" None "#00FF00" cs
check "colorscale reapply cells" (obj.ReferenceEquals(csReapplied.Cells, cs.Cells)) true
let csSplit =
    model
    |> applyColorScale 0 0 1 0 "#FF0000" None "#00FF00"
    |> applyColorScale 2 0 4 0 "#FF0000" None "#00FF00"
applyColorScale 0 0 1 0 "#FF0000" None "#00FF00" csSplit |> ignore
applyColorScale 2 0 4 0 "#FF0000" None "#00FF00" csSplit |> ignore
check "colorscale alternating no-op" (obj.ReferenceEquals(applyColorScale 0 0 1 0 "#FF0000" None "#00FF00" csSplit, csSplit)) true
let csDifferentHi = applyColorScale 0 0 4 0 "#FF0000" None "#0000FF" cs
check "colorscale cache color invalidates" ((fmtAt 4 csDifferentHi).BgColor) (Some "#0000FF")
check "colorscale cache keeps first source" (obj.ReferenceEquals((applyColorScale 0 0 4 0 "#FF0000" None "#00FF00" cs).Cells, cs.Cells)) true
let colorScaleApplyOversizedKeyCellsWeakReference () =
    let midColor = System.String('M', 6 * 1024 * 1024)
    let cells =
        [ { Row = 0; Col = 0 },
          { Value = CellValue.Number 0.0
            Format = { defaultFormat with BgColor = Some "#FF0000" }
            Comment = None }
          { Row = 1; Col = 0 },
          { Value = CellValue.Number 100.0
            Format = { defaultFormat with BgColor = Some "#00FF00" }
            Comment = None } ]
        |> Map.ofList
    let source = { createEmpty 2 1 with Cells = cells }
    applyColorScale 0 0 1 0 "#FF0000" (Some midColor) "#00FF00" source |> ignore
    System.WeakReference(cells)
let colorScaleApplyOversizedKeyCellsWeak = colorScaleApplyOversizedKeyCellsWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
check "colorscale cache skips oversized key"
    colorScaleApplyOversizedKeyCellsWeak.IsAlive
    false
let cfFirstApplyRows = 25000
let cfFirstApplyCols = 4
let cfFirstApplyModel =
    { createEmpty (cfFirstApplyRows + 1) (cfFirstApplyCols + 1) with
        Cells =
            [ for row in 0 .. cfFirstApplyRows - 1 do
                for col in 0 .. cfFirstApplyCols - 1 ->
                    { Row = row; Col = col }, cell (CellValue.Number (float (row * cfFirstApplyCols + col))) ]
            |> Map.ofList }
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let cfFirstApplyBefore = System.GC.GetAllocatedBytesForCurrentThread()
let cfFirstApplySw = System.Diagnostics.Stopwatch.StartNew()
let cfFirstApplied =
    applyColorScale 0 0 (cfFirstApplyRows - 1) (cfFirstApplyCols - 1) "#FF0000" None "#00FF00" cfFirstApplyModel
cfFirstApplySw.Stop()
let cfFirstApplyBytes = System.GC.GetAllocatedBytesForCurrentThread() - cfFirstApplyBefore
check "colorscale first apply perf"
    (cfFirstApplySw.Elapsed.TotalMilliseconds < 120.0 && cfFirstApplied.Cells.Count = cfFirstApplyModel.Cells.Count)
    true
check "colorscale first apply allocation"
    (cfFirstApplyBytes < 55000000L)
    true

// ── icon set (3 buckets) ──
let ic = applyIconSet 0 0 4 0 3 model
check "iconset low"  ((fmtAt 0 ic).IconSet) (Some (0, 3))   // 0
check "iconset mid"  ((fmtAt 2 ic).IconSet) (Some (1, 3))   // 50 → bucket 1
check "iconset high" ((fmtAt 4 ic).IconSet) (Some (2, 3))   // 100 → clamped to 2
let icReapplied = applyIconSet 0 0 4 0 3 ic
check "iconset reapply cells" (obj.ReferenceEquals(icReapplied.Cells, ic.Cells)) true
let icSplit =
    model
    |> applyIconSet 0 0 1 0 3
    |> applyIconSet 2 0 4 0 3
applyIconSet 0 0 1 0 3 icSplit |> ignore
applyIconSet 2 0 4 0 3 icSplit |> ignore
check "iconset alternating no-op" (obj.ReferenceEquals(applyIconSet 0 0 1 0 3 icSplit, icSplit)) true
let icPartialStale =
    match getCell { Row=2; Col=0 } ic with
    | Some cd -> { ic with Cells = Map.add { Row=2; Col=0 } { cd with Format = { cd.Format with IconSet = None } } ic.Cells }
    | None -> failwith "missing iconset cell"
let icPartialReapplied = applyIconSet 0 0 4 0 3 icPartialStale
check "iconset partial repaired" ((fmtAt 2 icPartialReapplied).IconSet) (Some (1, 3))
let icDifferentCount = applyIconSet 0 0 4 0 4 ic
check "iconset cache count invalidates" ((fmtAt 4 icDifferentCount).IconSet) (Some (3, 4))
check "iconset cache keeps first source" (obj.ReferenceEquals((applyIconSet 0 0 4 0 3 ic).Cells, ic.Cells)) true
let mutable wideIconSetModel = wideCfModel
for col in 0 .. 15 do
    wideIconSetModel <- applyIconSet 0 col wideCfLastRow col 3 wideIconSetModel
applyIconSet 0 0 wideCfLastRow 0 3 wideIconSetModel |> ignore
for col in 1 .. 15 do
    applyIconSet 0 col wideCfLastRow col 3 wideIconSetModel |> ignore
let wideIconSetNoOpBytes =
    allocatedBytesFor (fun () ->
        for _ in 1 .. 3 do
            for col in 0 .. 15 do
                applyIconSet 0 col wideCfLastRow col 3 wideIconSetModel |> ignore)
check "iconset cache keeps many ranges" (wideIconSetNoOpBytes < 250000L) true

let cfAddr = { Row = 2; Col = 2 }
let cfRule = { Rule = GreaterThan 10.0; Style = { defaultFormat with BgColor = Some "#FFFF00" } }
let cfBase = createEmpty 20 8
check "cf remove missing identity" (obj.ReferenceEquals(removeConditionalFormats cfAddr cfBase, cfBase)) true
let cfAdded = addConditionalFormat cfAddr cfRule cfBase
check "cf add same identity" (obj.ReferenceEquals(addConditionalFormat cfAddr cfRule cfAdded, cfAdded)) true
check "cf add same count" (Map.find cfAddr cfAdded.ConditionalFormats).Length 1
let cfRemoved = removeConditionalFormats cfAddr cfAdded
check "cf remove existing clears" (Map.containsKey cfAddr cfRemoved.ConditionalFormats) false

// Sparse large ranges should only format populated numeric cells.
let sparse =
    { createEmpty 10000 1000 with
        Cells =
            [ { Row=0; Col=0 }, cell (CellValue.Number 0.0)
              { Row=5000; Col=500 }, cell (CellValue.Number 50.0)
              { Row=9999; Col=999 }, cell (CellValue.Number 100.0)
              { Row=100; Col=100 }, cell (CellValue.String "skip") ]
            |> Map.ofList }
let sparseCs = applyColorScale 0 0 9999 999 "#FF0000" None "#00FF00" sparse
check "sparse colorscale count" sparseCs.Cells.Count sparse.Cells.Count
check "sparse colorscale min" ((getCell { Row=0; Col=0 } sparseCs).Value.Format.BgColor) (Some "#FF0000")
check "sparse colorscale mid" ((getCell { Row=5000; Col=500 } sparseCs).Value.Format.BgColor) (Some "#7F7F00")
check "sparse colorscale text" ((getCell { Row=100; Col=100 } sparseCs).Value.Format.BgColor) None
let sparseDb = applyDataBars 0 0 9999 999 "#638EC6" sparse
check "sparse databar max" ((getCell { Row=9999; Col=999 } sparseDb).Value.Format.DataBar) (Some (1.0, "#638EC6"))
let sparseIc = applyIconSet 0 0 9999 999 3 sparse
check "sparse iconset mid" ((getCell { Row=5000; Col=500 } sparseIc).Value.Format.IconSet) (Some (1, 3))

printfn ""
printfn "Stage 14 CF-visual tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
