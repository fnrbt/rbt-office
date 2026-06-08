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
let fmtAt r (m: SpreadsheetModel) = match getCell { Row=r; Col=0 } m with Some cd -> cd.Format | None -> defaultFormat

// ── data bars (0-baseline, all ≥ 0) ──
let db = applyDataBars 0 0 4 0 "#638EC6" model
check "databar min"  ((fmtAt 0 db).DataBar) (Some (0.0, "#638EC6"))
check "databar max"  ((fmtAt 4 db).DataBar) (Some (1.0, "#638EC6"))
check "databar mid"  ((fmtAt 2 db).DataBar) (Some (0.5, "#638EC6"))

// ── 2-color scale (red→green) ──
let cs = applyColorScale 0 0 4 0 "#FF0000" None "#00FF00" model
check "colorscale min" ((fmtAt 0 cs).BgColor) (Some "#FF0000")
check "colorscale max" ((fmtAt 4 cs).BgColor) (Some "#00FF00")
check "colorscale mid" ((fmtAt 2 cs).BgColor) (Some "#7F7F00")   // halfway lerp

// ── icon set (3 buckets) ──
let ic = applyIconSet 0 0 4 0 3 model
check "iconset low"  ((fmtAt 0 ic).IconSet) (Some (0, 3))   // 0
check "iconset mid"  ((fmtAt 2 ic).IconSet) (Some (1, 3))   // 50 → bucket 1
check "iconset high" ((fmtAt 4 ic).IconSet) (Some (2, 3))   // 100 → clamped to 2

printfn ""
printfn "Stage 14 CF-visual tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
