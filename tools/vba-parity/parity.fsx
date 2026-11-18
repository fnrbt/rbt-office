#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Vb.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// VBA parity with Microsoft Excel: the same macros run in Excel (tools/vba-parity/run-on-vm.sh) and here,
// and what they leave in the cells is compared.
//
//   dotnet fsi tools/vba-parity/parity.fsx build CASES DOCS        each CASES/*.bas as DOCS/<name>.xlsm
//   tools/vba-parity/run-on-vm.sh DOCS excel.tsv                   Excel's results
//   dotnet fsi tools/vba-parity/parity.fsx ours DOCS ours.tsv      ours, in the same form
//   dotnet fsi tools/vba-parity/parity.fsx compare excel.tsv ours.tsv
//
// A case is a standard module with `Sub Main()` that writes what it finds into cells. It runs inside
// RbtParityMain, which traps an error and writes "#ERR <number> <description>" to Worksheets(1)!Z1 (an
// unhandled error would stop Excel with a dialog).

open System
open System.IO
open Rbt.Spreadsheet.Model
module Bridge = Rbt.Spreadsheet.XlsxBridge

let invariant = Globalization.CultureInfo.InvariantCulture

/// `P label, value` writes, on the next row of the first sheet, the label (A), the value as a cell takes it (B),
/// TypeName (C) and CStr as text (D) - one probe per line of a case.
let probe = "\n\nSub P(label, v)\n    rbtOk = True\n    rbtRow = rbtRow + 1\n    Worksheets(1).Cells(rbtRow, 1).Value = label\n    Worksheets(1).Cells(rbtRow, 3).Value = TypeName(v)\n    If IsObject(v) Or IsArray(v) Then Exit Sub\n    On Error Resume Next\n    Worksheets(1).Cells(rbtRow, 2).Value = v\n    If Err.Number <> 0 Then Worksheets(1).Cells(rbtRow, 2).Value = \"#cell \" & Err.Number: Err.Clear\n    Worksheets(1).Cells(rbtRow, 4).Value = \"'\" & CStr(v)\n    If Err.Number <> 0 Then Worksheets(1).Cells(rbtRow, 4).Value = \"#CStr \" & Err.Number\nEnd Sub\nSub PE(label, number)\n    rbtRow = rbtRow + 1\n    Worksheets(1).Cells(rbtRow, 1).Value = label\n    Worksheets(1).Cells(rbtRow, 2).Value = \"#ERR \" & number\nEnd Sub\n"

/// A probe line `P "label", expression` on its own: an error in the expression is that probe's row
/// ("#ERR n" in B) and the case goes on, in both Excel and ours, so later probes stay on their rows.
let guardedProbes (source: string) =
    let pattern = Text.RegularExpressions.Regex("^(\\s*)P (\"[^\"]*\"), (.+)$")
    source.Replace("\r\n", "\n").Split('\n')
    |> Array.map (fun line ->
        let m = pattern.Match line
        if not m.Success then line
        // (P sets rbtOk first thing: not set, the expression failed. Err is left as the case has it.)
        else sprintf "%sOn Error Resume Next: rbtOk = False: P %s, %s: If Not rbtOk Then PE %s, Err.Number: Err.Clear" m.Groups.[1].Value m.Groups.[2].Value m.Groups.[3].Value m.Groups.[2].Value)
    |> String.concat "\n"

// (Declarations come before the first procedure: VBA allows nothing else there.)
let wrapped (source: string) =
    "Private rbtRow As Long\nPrivate rbtOk As Boolean\n" + guardedProbes (source.TrimEnd()) + probe + "\n\nSub RbtParityMain()\n    On Error GoTo failed\n    Main\n    Exit Sub\nfailed:\n    Worksheets(1).Range(\"Z1\").Value = \"#ERR \" & Err.Number & \" \" & Err.Description\nEnd Sub\n"

let build (cases: string) (docs: string) =
    Directory.CreateDirectory docs |> ignore
    for file in Directory.GetFiles(cases, "*.bas") |> Array.sort do
        let wb = Rbt.Spreadsheet.Workbook.createWorkbook ()
        let sheet = wb.Sheets.[wb.ActiveSheet]
        let project = VbaProject.writeProjectModules [ "Module1", VbaProject.StandardModule, wrapped (File.ReadAllText file) ]
        let input : Bridge.ExportInput =
            { Sheets = [| sheet.Model |]; Names = [| sheet.Name |]; Visible = [| true |]; ActiveSheetIndex = 0
              NameInfo = Map.empty; DefinedNames = Map.empty; Source = None; Macros = None; MacroProject = Some project }
        File.WriteAllBytes(Path.Combine(docs, Path.GetFileNameWithoutExtension file + ".xlsm"), XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input))
        // (Excel gets the module itself: run-cases.vbs puts it into a blank workbook's project.)
        File.WriteAllText(Path.Combine(docs, Path.GetFileNameWithoutExtension file + ".bas"), (wrapped (File.ReadAllText file)).Replace("\r\n", "\n").Replace("\n", "\r\n"))
    printfn "built %d cases into %s" (Directory.GetFiles(docs, "*.xlsm").Length) docs

let colorLong (color: string option) =
    match color with
    | Some text when text.StartsWith "#" && text.Length >= 7 ->
        Some (Convert.ToInt32(text.Substring(1, 2), 16) + Convert.ToInt32(text.Substring(3, 2), 16) * 256 + Convert.ToInt32(text.Substring(5, 2), 16) * 65536)
    | _ -> None

let ours (docs: string) (output: string) =
    let lines = ResizeArray<string>()
    for file in Directory.GetFiles(docs, "*.xlsm") |> Array.sort do
        let name = Path.GetFileNameWithoutExtension file
        let reader = XlsxReader.loadBytesNative (File.ReadAllBytes file) |> Option.get
        let imported = Bridge.Import.importWorkbookWith Bridge.editorImportOptions reader
        let project = Rbt.Spreadsheet.MacroProject.openProject (Some reader.Preserved)
        let started = DateTime.UtcNow
        let outcome = project.Run(imported.Context, Rbt.Spreadsheet.MacroApi.defaultHooks, "RbtParityMain")
        let seconds = (DateTime.UtcNow - started).TotalSeconds
        match outcome with
        | Rbt.Spreadsheet.MacroApi.MacroFailed failure ->
            lines.Add(sprintf "%s\t#RUN\trun-error %d %s" name failure.Number failure.Description)
        | Rbt.Spreadsheet.MacroApi.MacroDone (wb, _) ->
            lines.Add(sprintf "%s\t#RUN\tok" name)
            if seconds > 2.0 then eprintfn "%s took %.1f s" name seconds
            let date1904 = Rbt.Spreadsheet.Workbook.date1904 wb
            for id in wb.SheetOrder do
                let sheet = wb.Sheets.[id]
                let model = sheet.Model
                for KeyValue(address, cell) in model.Cells |> Seq.sortBy (fun kv -> kv.Key.Row, kv.Key.Col) do
                    let format = cell.Format
                    let value = unwrapFormulaValue cell.Value
                    let fill = colorLong format.BgColor |> Option.map string |> Option.defaultValue "none"
                    let fontColor = colorLong format.TextColor |> Option.defaultValue 0
                    if not (value = CellValue.Empty && fill = "none" && not format.Bold && fontColor = 0) then
                        let width = model.ColumnWidths |> Map.tryFind address.Col |> Option.defaultValue model.DefaultColumnWidth
                        let text = Rbt.Spreadsheet.CellText.cellDisplayTextIn date1904 None width format cell.Value
                        let t, v =
                            match value with
                            | Number n -> "n", n.ToString("0.00000000000000E+00", invariant)
                            | String s -> "s", s.Replace("\t", "\\t").Replace("\r", "\\r").Replace("\n", "\\n")
                            | Boolean b -> "b", (if b then "True" else "False")
                            | Error _ -> "e", text
                            | _ -> "x", ""
                        let numberFormat = format.NumberFormat |> Option.defaultValue "General"
                        lines.Add(String.Join("\t", [| name; sheet.Name; columnToLetter address.Col + string (address.Row + 1); t; v; text.Replace("\t", " "); numberFormat
                                                       (if format.Bold then "True" else "False"); fill; string fontColor |]))
    lines.Add "#DONE"
    File.WriteAllLines(output, lines)
    printfn "ours: %s" output

type Row = { Kind: string; Value: string; Text: string; Format: string; Bold: string; Fill: string; Font: string }

let read (path: string) =
    let runs = Collections.Generic.Dictionary<string, string>()
    let cells = Collections.Generic.Dictionary<string, Map<string * string, Row>>()
    for line in File.ReadAllLines path do
        let f = line.Split '\t'
        if f.Length >= 3 && f.[1] = "#RUN" then runs.[f.[0]] <- f.[2]
        elif f.Length >= 10 then
            let row = { Kind = f.[3]; Value = f.[4]; Text = f.[5]; Format = f.[6]; Bold = f.[7]; Fill = f.[8]; Font = f.[9] }
            let known = match cells.TryGetValue f.[0] with | true, m -> m | _ -> Map.empty
            cells.[f.[0]] <- Map.add (f.[1], f.[2]) row known
    runs, cells

let sameValue (a: Row) (b: Row) =
    if a.Kind <> b.Kind then false
    elif a.Kind = "n" then
        match Double.TryParse(a.Value, Globalization.NumberStyles.Float, invariant), Double.TryParse(b.Value, Globalization.NumberStyles.Float, invariant) with
        | (true, x), (true, y) -> x = y || abs (x - y) <= 1e-13 * max (abs x) (abs y)
        | _ -> a.Value = b.Value
    else a.Value = b.Value

let compare (excelPath: string) (oursPath: string) =
    let excelRuns, excelCells = read excelPath
    let ourRuns, ourCells = read oursPath
    let mutable passed = 0
    let failed = ResizeArray<string>()
    for case in excelRuns.Keys |> Seq.sort do
        let problems = ResizeArray<string>()
        let ourRun = match ourRuns.TryGetValue case with | true, r -> r | _ -> "(not run)"
        if excelRuns.[case] <> ourRun then problems.Add(sprintf "run: Excel %s / ours %s" excelRuns.[case] ourRun)
        let e = match excelCells.TryGetValue case with | true, m -> m | _ -> Map.empty
        let o = match ourCells.TryGetValue case with | true, m -> m | _ -> Map.empty
        for key in Set.union (e |> Map.keys |> Set.ofSeq) (o |> Map.keys |> Set.ofSeq) do
            let sheet, address = key
            match Map.tryFind key e, Map.tryFind key o with
            | Some a, None -> problems.Add(sprintf "%s!%s: Excel %s %A, ours nothing" sheet address a.Kind a.Value)
            | None, Some b -> problems.Add(sprintf "%s!%s: Excel nothing, ours %s %A" sheet address b.Kind b.Value)
            | Some a, Some b ->
                if not (sameValue a b) then problems.Add(sprintf "%s!%s value: Excel %s %A / ours %s %A" sheet address a.Kind a.Value b.Kind b.Value)
                // (Excel's #### is a value too wide for its column, which this dump does not measure.)
                elif a.Text <> b.Text && not (a.Text.Length > 0 && a.Text |> Seq.forall ((=) '#')) then problems.Add(sprintf "%s!%s text: Excel %A / ours %A" sheet address a.Text b.Text)
                if a.Format <> b.Format then problems.Add(sprintf "%s!%s format: Excel %A / ours %A" sheet address a.Format b.Format)
                if a.Bold <> b.Bold || a.Fill <> b.Fill || a.Font <> b.Font then
                    problems.Add(sprintf "%s!%s look: Excel bold %s fill %s font %s / ours bold %s fill %s font %s" sheet address a.Bold a.Fill a.Font b.Bold b.Fill b.Font)
            | None, None -> ()
        if problems.Count = 0 then passed <- passed + 1
        else
            failed.Add case
            printfn "FAIL %s" case
            for p in problems |> Seq.truncate (match Environment.GetEnvironmentVariable "VBA_PARITY_ALL" with null -> 12 | _ -> 10000) do printfn "     %s" p
            if problems.Count > 12 then printfn "     ... %d more" (problems.Count - 12)
    printfn "\nVBA parity: %d of %d cases as Excel%s" passed excelRuns.Count (if failed.Count = 0 then "" else "; differ: " + String.Join(", ", failed))

match fsi.CommandLineArgs |> Array.toList |> List.tail with
| [ "build"; cases; docs ] -> build cases docs
| [ "ours"; docs; output ] -> ours docs output
| [ "compare"; excel; output ] -> compare excel output
| _ -> eprintfn "usage: parity.fsx build CASES DOCS | ours DOCS OUT.tsv | compare EXCEL.tsv OURS.tsv"; exit 2
