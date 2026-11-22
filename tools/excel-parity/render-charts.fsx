// dotnet fsi tools/excel-parity/render-charts.fsx WORKBOOK OUTDIR [W H]: every chart of every sheet of a workbook laid
// out as ours (Charts.objectScene) at W x H px (480 x 288: 360 x 216 pt) and written as OUTDIR/<sheet>-<chart>.svg, for
// comparing with Excel's own pictures of them (Chart.Export: chart-probe.vbs, build-chartex-zoo.vbs). A chart sheet's
// chart is written as OUTDIR/<sheet>-chart.svg at the size its page gives it. shoot-svgs.mjs turns them into PNGs.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open System
open System.IO
open Rbt.Spreadsheet.Model
module Charts = Rbt.Spreadsheet.Charts
module Bridge = Rbt.Spreadsheet.XlsxBridge

let args = fsi.CommandLineArgs |> Array.skip 1
let file, outDir = args.[0], args.[1]
let w, h = (if args.Length > 3 then float args.[2], float args.[3] else 480.0, 288.0)
Directory.CreateDirectory outDir |> ignore
let text (data: byte[]) = Text.Encoding.UTF8.GetString data
let wb = XlsxReader.loadBytesNative (File.ReadAllBytes file) |> Option.get
let book = Bridge.Import.importWorkbook wb
let parts = wb.Preserved.Parts |> Map.ofList
let theme = wb.Preserved.ThemePath |> Option.bind (fun path -> Map.tryFind path parts) |> Option.map (text >> Charts.parseTheme) |> Option.defaultValue Charts.defaultTheme
let names (name: string) = wb.NamedRanges |> List.tryPick (fun known -> if known.Name = name then Some known.RefersTo else None)
let data : Charts.ChartData =
    { Cells = (fun name r0 c0 r1 c1 ->
        book.Sheets |> Array.tryFind (fun sheet -> String.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
        |> Option.map (fun sheet ->
            let model = resolveSheet sheet.Model
            Array.init (r1 - r0 + 1) (fun r -> Array.init (c1 - c0 + 1) (fun c -> getCell { Row = r0 + r; Col = c0 + c } model))))
      Display = (fun cell -> Rbt.Spreadsheet.CellText.cellDisplayText None 1000.0 cell.Format cell.Value)
      Format = (fun code value -> Rbt.Spreadsheet.CellText.formatNumberWithCodeIn false code value) }
for sheet in wb.Sheets do
    let target (id: string) = sheet.Extras.DrawingRelationships |> List.tryPick (fun (known, _, target, ext) -> if known = id && not ext then Some (ZipInterop.normalizePath target) else None)
    let drawing : Charts.DrawingParts =
        { Theme = theme
          Chart = (fun id -> target id |> Option.bind (fun path -> Map.tryFind path parts |> Option.bind (fun bytes -> Charts.parseChartWith names theme (text bytes) |> Option.map (Charts.withUserShapes (fun name -> Map.tryFind name parts |> Option.map text) path))))
          Picture = (fun _ -> None); Slicer = (fun _ _ _ _ _ -> None) }
    for item in sheet.Extras.Objects |> List.filter (fun item -> item.Kind = "chart") do
        match Charts.objectScene drawing data item.Xml w h with
        | Some scene ->
            let name = (defaultArg item.Name "chart").Replace(" ", "_")
            File.WriteAllText(Path.Combine(outDir, sprintf "%s-%s.svg" (sheet.Name.Replace(" ", "_")) name), Charts.sceneSvg scene)
        | None -> ()
