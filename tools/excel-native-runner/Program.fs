module ExcelNativeRunnerCli

open System
open System.IO
open Rbt.Renderer
open Rbt.Spreadsheet.Model
open ExcelNativeRunner

module Xlsx = XlsxReader
module PngEncode = Rbt.OfficeNative.Png

module GridRender = Rbt.SpreadsheetNative.GridRender

let private renderGridPng width height model extra =
    let viewport = GridRender.viewportOf model (float width) (float height)
    let commands =
        Clear Color.white
        :: (GridRender.commands model viewport @ GridRender.paintItemsToCommands extra)
    Scene.render (Raster.create()) width height commands |> PngEncode.encode

let private renderScenePng width height items =
    GridRender.renderSceneCommands (float width) (float height) items
    |> Scene.render (Raster.create()) width height
    |> PngEncode.encode

let private usage () =
    eprintfn "Usage: dotnet run --project tools/excel-native-runner -- <input.xlsx> <output.png> [width] [height] [sheetIndex|all] [jsonOut]"
    eprintfn "       Pass 'all' as the sheet selector to render every sheet to <output>-sheet<N>.png."
    eprintfn "Defaults: width=1100 height=720 sheetIndex=0"

let private fullPath (path: string) =
    if Path.IsPathRooted path then path
    else Path.GetFullPath(path, Directory.GetCurrentDirectory())
let private escapeJson (value: string) =
    let sb = System.Text.StringBuilder(value.Length + 2)
    sb.Append('"') |> ignore
    for ch in value do
        match ch with
        | '"' -> sb.Append("\\\"") |> ignore
        | '\\' -> sb.Append("\\\\") |> ignore
        | '\n' -> sb.Append("\\n") |> ignore
        | '\r' -> sb.Append("\\r") |> ignore
        | '\t' -> sb.Append("\\t") |> ignore
        | c when int c < 0x20 -> sb.AppendFormat("\\u{0:x4}", int c) |> ignore
        | c -> sb.Append(c) |> ignore
    sb.Append('"') |> ignore
    sb.ToString()

let private cellInfoJson (sheetName: string) (model: SpreadsheetModel) =
    let columnLetter col = Rbt.Spreadsheet.Model.columnToLetter col
    let viewport = Rbt.Spreadsheet.Layout.defaultViewport 0.0 0.0
    let records =
        model.Cells
        |> Map.toList
        |> List.map (fun (addr, data) ->
            let rect = Rbt.Spreadsheet.Layout.cellRect model viewport addr.Row addr.Col
            let display =
                match data.Value with
                | Number n -> string n
                | String s -> s
                | Boolean b -> if b then "TRUE" else "FALSE"
                | Formula (_, _) -> ""
                | Empty -> ""
                | Error e -> e
            let fields = [
                "\"Sheet\":" + escapeJson sheetName
                "\"Address\":" + escapeJson (sprintf "%s%d" (columnLetter addr.Col) (addr.Row + 1))
                "\"Row\":" + string (addr.Row + 1)
                "\"Column\":" + string (addr.Col + 1)
                "\"X\":" + sprintf "%.2f" (rect.X - viewport.RowHeaderWidth)
                "\"Y\":" + sprintf "%.2f" (rect.Y - viewport.HeaderHeight)
                "\"Width\":" + sprintf "%.2f" rect.Width
                "\"Height\":" + sprintf "%.2f" rect.Height
                "\"Value\":" + escapeJson display
                "\"Text\":" + escapeJson display
                "\"NumberFormat\":" + escapeJson (data.Format.NumberFormat |> Option.defaultValue "")
            ]
            "{ " + String.concat ", " fields + " }")
    "[\n  " + String.concat ",\n  " records + "\n]\n"

let private renderSheetToPath
    (inputPath: string) (outputPath: string) (jsonPath: string option)
    (width: int) (height: int)
    (sheetIndex: int) (sheet: Xlsx.Worksheet) (sheetNames: string array) (wbContext: Rbt.Spreadsheet.Workbook.Workbook) =
    // The shared workbook import computes only uncached formulas and keeps
    // cached results intact, including references to other sheets and names.
    let sheetId = sprintf "sheet%d" (sheetIndex + 1)
    let baseModel = (Map.find sheetId wbContext.Sheets).Model
    let model =
        let m = baseModel
        // Opt-in: recompute pivot tables from their cache and place the live
        // result (cached pivot cells already render, so this is off by default
        // to leave the parity corpus untouched).
        if Environment.GetEnvironmentVariable("EXCEL_RECOMPUTE_PIVOTS") = "1" && sheetIndex = 0 then
            ChartRender.parsePivots inputPath
            |> List.fold (fun acc (p: ChartRender.PivotPlacement) ->
                let mutable a = acc
                for r in 0 .. p.Grid.Length - 1 do
                    for c in 0 .. p.Grid.[r].Length - 1 do
                        a <- setCell { Row = p.Row + r; Col = p.Col + c } { Value = p.Grid.[r].[c]; Format = defaultFormat; Comment = None } a
                a) m
        else m

    // Overlay charts and sparklines at their anchor / cell rects.
    let viewport = GridRender.viewportOf model (float width) (float height)
    let chartItems =
        ChartRender.parseAllCharts inputPath
        |> List.collect (fun pl ->
            let f = Rbt.Spreadsheet.Layout.cellRect model viewport pl.FromRow pl.FromCol
            let t = Rbt.Spreadsheet.Layout.cellRect model viewport pl.ToRow pl.ToCol
            let left = f.X + pl.FromColOffPx
            let top = f.Y + pl.FromRowOffPx
            let w = (t.X + pl.ToColOffPx) - left
            let h = (t.Y + pl.ToRowOffPx) - top
            [ Rbt.Spreadsheet.Paint.VecGroup((1.0, 0.0, 0.0, 1.0, left, top), ChartRender.buildChart pl.Spec w h) ])
    let source : Rbt.Spreadsheet.Sparklines.Source =
        fun name ->
            let target = name |> Option.defaultValue sheet.Name
            sheetNames
            |> Array.tryFindIndex (fun known -> String.Equals(known,target,StringComparison.OrdinalIgnoreCase))
            |> Option.map (fun index ->
                let other = (Map.find (sprintf "sheet%d" (index+1)) wbContext.Sheets).Model
                fun row col -> getCell { Row = row; Col = col } other)
    let sparkItems =
        Rbt.Spreadsheet.Paint.sparklineCellItems source model viewport
        |> List.map snd
    let slicerItems =
        ChartRender.parseSlicers inputPath
        |> List.collect (fun sl ->
            let f = Rbt.Spreadsheet.Layout.cellRect model viewport sl.FromRow sl.FromCol
            let t = Rbt.Spreadsheet.Layout.cellRect model viewport sl.ToRow sl.ToCol
            let rect : Rbt.Spreadsheet.Layout.Rect =
                { X = f.X + sl.FromColOffPx; Y = f.Y + sl.FromRowOffPx
                  Width = (t.X + sl.ToColOffPx) - (f.X + sl.FromColOffPx)
                  Height = (t.Y + sl.ToRowOffPx) - (f.Y + sl.FromRowOffPx) }
            ChartRender.buildSlicer sl rect)
    let png = renderGridPng width height { model with Sparklines = [] } (chartItems @ sparkItems @ slicerItems)

    Directory.CreateDirectory(Path.GetDirectoryName outputPath) |> ignore
    File.WriteAllBytes(outputPath, png)

    match jsonPath with
    | Some path ->
        Directory.CreateDirectory(Path.GetDirectoryName path) |> ignore
        File.WriteAllText(path, cellInfoJson sheet.Name model)
    | None -> ()

    let cellCount = Map.count model.Cells
    printfn "excel native runner ok: %s [sheet=%s cells=%d] -> %s (%dx%d)" inputPath sheet.Name cellCount outputPath width height

/// Insert `suffix` before the file extension so `out/foo.png` → `out/foo-sheet2.png`.
let private withSheetSuffix (path: string) (suffix: string) =
    let dir = Path.GetDirectoryName path
    let stem = Path.GetFileNameWithoutExtension path
    let ext = Path.GetExtension path
    let leaf = stem + "-" + suffix + ext
    if String.IsNullOrEmpty dir then leaf else Path.Combine(dir, leaf)

let private renderXlsx (inputPath: string) (outputPath: string) (width: int) (height: int) (sheetSel: Choice<int, unit>) (jsonPath: string option) =
    let bytes = File.ReadAllBytes inputPath
    let workbook =
        match Xlsx.loadBytesNative bytes with
        | Some parsed -> parsed
        | None -> failwithf "Could not parse XLSX: %s" inputPath

    if workbook.Sheets.IsEmpty then
        failwithf "No sheets in workbook: %s" inputPath

    // Build the formula-engine context once: every sheet's model + the
    // workbook's defined names. Shared across all sheets so cross-sheet
    // references and named ranges resolve during recalculation.
    let wbContext = buildWorkbookContext workbook
    let sheetNames = workbook.Sheets |> List.map (fun s -> s.Name) |> List.toArray

    match sheetSel with
    | Choice1Of2 sheetIndex ->
        let chosenIndex = min (max 0 sheetIndex) (workbook.Sheets.Length - 1)
        let sheet = workbook.Sheets.[chosenIndex]
        renderSheetToPath inputPath outputPath jsonPath width height chosenIndex sheet sheetNames wbContext
    | Choice2Of2 () ->
        // --all: render every sheet to `<base>-sheet<N>.png` (1-indexed,
        // matching how Excel labels them in the tab strip and how pandas
        // writes sheet names like "Sheet1"). When there's only one sheet,
        // skip the suffix so the output path stays predictable.
        let sheets = workbook.Sheets
        sheets
        |> List.iteri (fun i sheet ->
            let path =
                if sheets.Length = 1 then outputPath
                else withSheetSuffix outputPath (sprintf "sheet%d" (i + 1))
            let json =
                jsonPath |> Option.map (fun p ->
                    if sheets.Length = 1 then p
                    else withSheetSuffix p (sprintf "sheet%d" (i + 1)))
            renderSheetToPath inputPath path json width height i sheet sheetNames wbContext)

/// Parse the sheet argument: `all` / `-a` selects every sheet; an integer
/// selects that 0-based sheet index. Anything else falls back to sheet 0.
let private parseSheetArg (value: string) : Choice<int, unit> =
    if value = "all" || value = "-a" || value = "--all" then Choice2Of2 ()
    else
        match Int32.TryParse value with
        | true, parsed -> Choice1Of2 parsed
        | false, _ -> Choice1Of2 0

[<EntryPoint>]
let main argv =
    try
        let parseInt fallback (value: string) =
            match Int32.TryParse value with
            | true, parsed -> parsed
            | false, _ -> fallback
        let jsonEnv = Environment.GetEnvironmentVariable("EXCEL_NATIVE_JSON")
        let jsonOpt = if String.IsNullOrWhiteSpace jsonEnv then None else Some (fullPath jsonEnv)
        let renderDialogPng (path: string) (spec: Rbt.Office.Ui.DialogSpec) (w: int) (h: int) : int =
            Wisp.Layout.Scheduler.requestFrame <- ignore
            let mutable root = Unchecked.defaultof<Wisp.Layout.Element>
            let dispose = Wisp.Reactive.createRoot (fun dispose ->
                root <- Wisp.OfficeShell.dialogPreview Rbt.Spreadsheet.Ui.chrome spec
                dispose)
            root.Layout(Wisp.Layout.BoxConstraints.tightWH (float w) (float h))
            Wisp.Layout.flushLayout ()
            let cmds = Clear Color.white :: root.Paint(0.0, 0.0)
            let backend = Raster.create()
            let framebuffer = Rbt.Renderer.Scene.render backend w h cmds
            dispose ()
            File.WriteAllBytes(path, PngEncode.encode framebuffer)
            0
        match argv with
        | [| "--charttest"; input; idx; output |] ->
            // Parse the workbook's Nth chart and render it through the vector
            // PaintItem path (proves the chart bridge end-to-end).
            let chartIdx = parseInt 1 idx
            match ChartRender.parseChart (fullPath input) chartIdx with
            | None -> eprintfn "no chart found"; 2
            | Some spec ->
                let w, h = ChartRender.parseAnchorSize (fullPath input) chartIdx
                let items = ChartRender.buildChart spec w h
                File.WriteAllBytes(fullPath output, renderScenePng (int (ceil w)) (int (ceil h)) items)
                0
        | [| "--pivottest"; output |] ->
            // Self-contained pivot demo: sample sales data (left) → computed
            // Region × Product cross-tab with Grand Totals (right), rendered.
            let S = CellValue.String
            let N = CellValue.Number
            let source =
                [| [| S "Region"; S "Product"; S "Sales" |]
                   [| S "East"; S "Apples"; N 100.0 |]
                   [| S "East"; S "Bananas"; N 200.0 |]
                   [| S "West"; S "Apples"; N 150.0 |]
                   [| S "West"; S "Bananas"; N 250.0 |]
                   [| S "East"; S "Apples"; N 50.0 |]
                   [| S "West"; S "Cherries"; N 300.0 |]
                   [| S "East"; S "Cherries"; N 120.0 |] |]
            // Nested rows (Region → Product) with subtotals + two measures.
            let pivot = Rbt.Spreadsheet.Pivot.computePivotN source [ 0; 1 ] [ (2, Rbt.Spreadsheet.Pivot.Sum); (2, Rbt.Spreadsheet.Pivot.Count) ]
            let bold = { defaultFormat with Bold = true }
            let endsTotal (v: CellValue) = match v with CellValue.String s -> s.EndsWith "Total" | _ -> false
            let place (grid: CellValue[][]) r0 c0 boldTotals (model: SpreadsheetModel) =
                let mutable m = model
                for r in 0 .. grid.Length - 1 do
                    let rowBold = r = 0 || (boldTotals && grid.[r].Length > 0 && endsTotal grid.[r].[0])
                    for c in 0 .. grid.[r].Length - 1 do
                        m <- setCell { Row = r0 + r; Col = c0 + c } { Value = grid.[r].[c]; Format = (if rowBold then bold else defaultFormat); Comment = None } m
                m
            let model = createEmpty 40 16 |> place source 0 0 false |> place pivot 0 5 true
            let w, h = 1000, 360
            File.WriteAllBytes(fullPath output, renderGridPng w h model [])
            0
        | [| "--dialogtest"; output |] ->
            // Render the Insert-Pivot dialog standalone (headless) — proves the
            // pivot UI's DialogSpec paints through the native shell.
            let samplePreview (_: Rbt.Spreadsheet.Ui.PivotUiRequest) =
                "=GROUPBY(A2:A8,C2:C8,\"sum\")\nEast    470\nWest    700\nGrand Total    1170"
            let spec = Rbt.Spreadsheet.Ui.pivotTableDialog samplePreview (fun _ -> true)
            renderDialogPng (fullPath output) spec 560 780
        | [| "--filterdialogtest"; output |] ->
            let spec = Rbt.Spreadsheet.Ui.filterDialog "Region" [ "East"; "North"; "South"; "West" ] (fun _ -> ())
            renderDialogPng (fullPath output) spec 480 460
        | [| "--macrodialogtest"; output |] ->
            let spec =
                Rbt.Spreadsheet.Ui.macroDialog
                    "Sub Main()\n  Range(\"A1\").Value = 42\nEnd Sub\nMain()"
                    (fun _ -> ())
            renderDialogPng (fullPath output) spec 560 520
        | [| "--formatdialogtest"; output |] ->
            // Format-cells dialog with a live "Sample" preview that re-renders the
            // chosen number format through the real cell renderer.
            let spec = Rbt.Spreadsheet.Ui.formatCellsDialog "currency" "#000000" "#FFFFFF" (fun _ -> ())
            renderDialogPng (fullPath output) spec 560 520
        | [| "--autofiltertest"; outputBefore; outputAfter |] ->
            // AutoFilter end-to-end: render the same sheet unfiltered, then with
            // Region filtered to {East, West} — proves filtered-out rows collapse
            // (height 0) while survivors keep their original row numbers.
            let S = CellValue.String
            let N = CellValue.Number
            let bold = { defaultFormat with Bold = true }
            let rows =
                [ [ S "Region"; S "Product"; S "Sales" ]
                  [ S "East"; S "Apples"; N 100.0 ]
                  [ S "North"; S "Bananas"; N 200.0 ]
                  [ S "West"; S "Apples"; N 150.0 ]
                  [ S "South"; S "Cherries"; N 250.0 ]
                  [ S "East"; S "Bananas"; N 50.0 ]
                  [ S "North"; S "Cherries"; N 300.0 ]
                  [ S "West"; S "Apples"; N 120.0 ]
                  [ S "South"; S "Bananas"; N 90.0 ] ]
            let baseModel =
                rows
                |> List.mapi (fun r cols -> r, cols)
                |> List.fold (fun m (r, cols) ->
                    cols |> List.mapi (fun c v -> c, v)
                    |> List.fold (fun acc (c, v) ->
                        setCell { Row = r; Col = c } { Value = v; Format = (if r = 0 then bold else defaultFormat); Comment = None } acc) m)
                    (createEmpty 20 8)
            let filtered = applyFilter 0 (Set.ofList [ "East"; "West" ]) baseModel
            let w, h = 360, 280
            let renderTo path model =
                File.WriteAllBytes(fullPath path, renderGridPng w h model [])
            renderTo outputBefore baseModel
            renderTo outputAfter filtered
            0
        | [| "--tabletest"; input |] ->
            // Import a table-bearing xlsx, bridge its tables into the workbook
            // context, recompute sheet1, then prove a structured reference
            // resolves to the same cells as the equivalent A1 range.
            let bytes = File.ReadAllBytes (fullPath input)
            match Xlsx.loadBytesNative bytes with
            | None -> eprintfn "parse failed"; 2
            | Some wb ->
                let ctx = buildWorkbookContext wb
                printfn "Bridged tables: %d" (Map.count ctx.Tables)
                ctx.Tables
                |> Map.iter (fun name (t: Rbt.Spreadsheet.Workbook.Table) ->
                    printfn "  %s  rows %d..%d cols %d..%d  headers=%b totals=%b  cols=[%s]"
                        name t.StartRow t.EndRow t.StartCol t.EndCol t.HasHeaders t.HasTotals (String.concat "," t.Columns))
                let sheetId = "sheet1"
                let model0 = (Map.find sheetId ctx.Sheets).Model
                let model = Rbt.Spreadsheet.Formula.recalculateAllFormulas (Some ctx) (Some sheetId) model0
                let ev (s: string) = Rbt.Spreadsheet.Formula.evaluateFormula s model (Some ctx)
                match ctx.Tables |> Map.toList with
                | (name, t) :: _ when not (List.isEmpty t.Columns) ->
                    let col0 = List.head t.Columns
                    let colLetter = Rbt.Spreadsheet.Model.columnToLetter t.StartCol
                    let r0 = t.StartRow + (if t.HasHeaders then 1 else 0) + 1
                    let r1 = t.EndRow + 1
                    // COUNTA over the data column (non-zero, regardless of text
                    // vs number) and the header cell value, each via the
                    // structured ref and the equivalent A1 range.
                    let structured = ev (sprintf "COUNTA(%s[%s])" name col0)
                    let plain = ev (sprintf "COUNTA(%s%d:%s%d)" colLetter r0 colLetter r1)
                    let hdrStructured = ev (sprintf "%s[[#Headers],[%s]]" name col0)
                    let hdrPlain = ev (sprintf "%s%d" colLetter (t.StartRow + 1))
                    printfn "COUNTA(%s[%s])  = %A   COUNTA range = %A" name col0 structured plain
                    printfn "%s[[#Headers],[%s]] = %A   header cell = %A" name col0 hdrStructured hdrPlain
                    if structured = plain && structured <> CellValue.Number 0.0 && hdrStructured = hdrPlain
                    then (printfn "MATCH (structured refs resolve to the same cells)"; 0)
                    else (printfn "MISMATCH"; 1)
                | _ -> printfn "no bridged table with columns"; 1
        | [| "--tablecreatetest"; output |] ->
            // Build plain data, register a table with a SUBTOTAL totals row via
            // Model.createTable, bridge it, recompute, and render — proving the
            // totals row computes and a structured ref over the new table works.
            let S = CellValue.String
            let N = CellValue.Number
            let bold = { defaultFormat with Bold = true }
            let rows =
                [ [ S "Region"; S "Q1"; S "Q2" ]
                  [ S "East"; N 100.; N 200. ]
                  [ S "West"; N 150.; N 250. ]
                  [ S "North"; N 120.; N 180. ] ]
            let m0 =
                rows
                |> List.mapi (fun r cols -> r, cols)
                |> List.fold (fun m (r, cols) ->
                    cols |> List.mapi (fun c v -> c, v)
                    |> List.fold (fun acc (c, v) ->
                        setCell { Row = r; Col = c } { Value = v; Format = (if r = 0 then bold else defaultFormat); Comment = None } acc) m)
                    (createEmpty 20 8)
            // Register the table (header row 0, data rows 1..3) + totals row.
            let withTable = Rbt.Spreadsheet.Model.createTable "Sales" 0 0 3 2 true m0
            // A structured-ref demo cell below the table.
            let withDemo =
                setCell { Row = 6; Col = 0 } { Value = S "SUM(Sales[Q1]) ="; Format = defaultFormat; Comment = None } withTable
                |> setCell { Row = 6; Col = 1 } { Value = CellValue.Formula ("SUM(Sales[Q1])", N 0.); Format = bold; Comment = None }
            // Bridge the model table into a workbook so structured refs resolve.
            let mt = List.last withDemo.Tables
            let wtable =
                Rbt.Spreadsheet.Workbook.makeTableFromRegion
                    "sheet1" mt.Name mt.StartRow mt.StartCol mt.EndRow mt.EndCol
                    mt.HeaderRowCount (if mt.HasTotals then 1 else 0) [] withDemo
            let wb : Rbt.Spreadsheet.Workbook.Workbook =
                { Sheets = Map.ofList [ "sheet1", ({ Id="sheet1"; Name="Sheet1"; Model=withDemo; Visible=true }: Rbt.Spreadsheet.Workbook.Sheet) ]
                  SheetOrder = [ "sheet1" ]; ActiveSheet = "sheet1"
                  Tables = Map.ofList [ mt.Name, wtable ]; Names = Map.empty; NameInfo = Map.empty; Iteration = None; External = Map.empty }
            let recomputed = Rbt.Spreadsheet.Formula.recalculateAllFormulas (Some wb) (Some "sheet1") withDemo
            let w, h = 360, 260
            File.WriteAllBytes(fullPath output, renderGridPng w h recomputed [])
            // Also echo the computed values for a textual check.
            let v (r:int) (c:int) = Rbt.Spreadsheet.Model.displayValue (getCellValue { Row=r; Col=c } recomputed)
            printfn "totals row: %s | %s | %s" (v 4 0) (v 4 1) (v 4 2)
            printfn "SUM(Sales[Q1]) = %s" (v 6 1)
            0
        | [| "--tablefiltertest"; input; outBefore; outAfter |] ->
            // Imported table + an applied column filter → filtered rows collapse
            // while the styled header and its filter buttons stay put (Excel's
            // table-filter behaviour). Proves the table filter button's model
            // path (applyFilter → applyFilterRowHeights) end to end.
            let bytes = File.ReadAllBytes (fullPath input)
            match Xlsx.loadBytesNative bytes with
            | None -> eprintfn "parse failed"; 2
            | Some wb ->
                let ctx = buildWorkbookContext wb
                let model = (Map.find "sheet1" ctx.Sheets).Model
                // Keep only column-0 values "1" and "3" (collapses the others).
                let filtered = applyFilter 0 (Set.ofList [ "1"; "3" ]) model
                let w, h = 700, 280
                let renderTo path m =
                    File.WriteAllBytes(fullPath path, renderGridPng w h m [])
                renderTo outBefore model
                renderTo outAfter filtered
                printfn "visible rows after filter: %A" (getVisibleRows filtered |> Set.toList |> List.sort)
                0
        | [| "--cftest"; output |] ->
            // Conditional-format visuals: data bars / color scale / icon set,
            // computed over real numbers and rendered.
            let N = CellValue.Number
            let bold = { defaultFormat with Bold = true }
            let vals = [ 10.; 45.; 30.; 80.; 55.; 20.; 95.; 65.; 5.; 70. ]
            let putCol col (vs: float list) m =
                vs |> List.mapi (fun i v -> i, v)
                |> List.fold (fun acc (i, v) -> setCell { Row = i + 1; Col = col } { Value = N v; Format = defaultFormat; Comment = None } acc) m
            let header col t m = setCell { Row = 0; Col = col } { Value = CellValue.String t; Format = bold; Comment = None } m
            let lastRow = List.length vals
            let model =
                createEmpty 20 8
                |> header 0 "Data bars" |> putCol 0 vals |> applyDataBars 1 0 lastRow 0 "#638EC6"
                |> header 2 "Color scale" |> putCol 2 vals |> applyColorScale 1 2 lastRow 2 "#F8696B" (Some "#FFEB84") "#63BE7B"
                |> header 4 "Icon set" |> putCol 4 vals |> applyIconSet 1 4 lastRow 4 3
            let w, h = 560, 320
            File.WriteAllBytes(fullPath output, renderGridPng w h model [])
            0
        | [| "--vectortest"; output |] ->
            // Render the hand-built vector scene (VecPath/VecGroup) through the
            // native rasterizer — proves the new paint primitives end-to-end.
            let w, h = 400, 280
            File.WriteAllBytes(fullPath output, renderScenePng w h Rbt.Spreadsheet.Paint.vectorTestScene)
            0
        | [| input; output |] ->
            renderXlsx (fullPath input) (fullPath output) 1100 720 (Choice1Of2 0) jsonOpt
            0
        | [| input; output; width |] ->
            renderXlsx (fullPath input) (fullPath output) (parseInt 1100 width) 720 (Choice1Of2 0) jsonOpt
            0
        | [| input; output; width; height |] ->
            renderXlsx (fullPath input) (fullPath output) (parseInt 1100 width) (parseInt 720 height) (Choice1Of2 0) jsonOpt
            0
        | [| input; output; width; height; sheet |] ->
            renderXlsx (fullPath input) (fullPath output) (parseInt 1100 width) (parseInt 720 height) (parseSheetArg sheet) jsonOpt
            0
        | [| input; output; width; height; sheet; jsonOut |] ->
            renderXlsx (fullPath input) (fullPath output) (parseInt 1100 width) (parseInt 720 height) (parseSheetArg sheet) (Some (fullPath jsonOut))
            0
        | _ ->
            usage()
            2
    with ex ->
        eprintfn "%O" ex
        1
