#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Vb.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// Showcase macros (tests/macros/showcase/*.bas, *.cls): visual VBA found on the web, each run on a fresh
// workbook through Excel's object model (MacroHost) with a limit on its instructions. Reports whether it
// ran, the error where it stopped (number, module, line), how long it took and what it changed.
//
//   dotnet build src/Spreadsheet && dotnet fsi tests/macros/showcase.fsx [--dump DIR] [--xlsm DIR] [name ...]
//
// --dump DIR writes DIR/<name>.tsv: the active sheet's cells with a value or a fill (address, Value2,
// fill as a BGR long or -), as tools/excel-parity/run-showcase-macro.vbs writes Excel's; --xlsm DIR writes
// DIR/<name>__<Entry>[__<Before>].xlsm, the macro in a blank macro-enabled workbook, for that script.
//
// A macro's header comments say how to run it:
//   ' Entry: Main             the Sub to run (default Main)
//   ' Input: 7                what InputBox answers (one line per call, the last one repeats)
//   ' Limit: 900000000        its instruction limit (default 600 million)
//   ' Before: Seed            a Sub run first (what a user would set up by hand, in a <name>.Prep.bas)
//   ' Ticks: 40               then up to 40 procedures it asked for with Application.OnTime, in order
//   ' Press: 5={DOWN}; 9={LEFT}   keys pressed before those ticks: their Application.OnKey procedures run
// A macro "is" the files <name>.bas / <name>.<module>.bas / <name>.<class>.cls; a file whose name starts
// with "_" is left out.

open System
open System.IO
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook
module M = Rbt.Spreadsheet.MacroHost

let folder = Path.Combine(__SOURCE_DIRECTORY__, "showcase")
let arguments = fsi.CommandLineArgs |> Array.skip 1 |> Array.toList
let rec options (args: string list) =
    match args with
    | "--dump" :: dir :: rest -> let d, x, names = options rest in Some dir, x, names
    | "--xlsm" :: dir :: rest -> let d, x, names = options rest in d, Some dir, names
    | name :: rest -> let d, x, names = options rest in d, x, name :: names
    | [] -> None, None, []
let dumpDir, xlsmDir, names = options arguments
let only = Set.ofList names

let header (key: string) (source: string) =
    source.Replace("\r\n", "\n").Split('\n')
    |> Array.choose (fun line ->
        let t = line.Trim()
        let prefix = "' " + key + ":"
        if t.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) then Some (t.Substring(prefix.Length).Trim()) else None)
    |> Array.toList

let groups =
    Directory.GetFiles folder
    |> Array.filter (fun f -> (f.EndsWith ".bas" || f.EndsWith ".cls") && not ((Path.GetFileName f).StartsWith "_"))
    |> Array.groupBy (fun f -> (Path.GetFileName f).Split('.').[0])
    |> Array.sortBy fst
    |> Array.filter (fun (name, _) -> only.IsEmpty || only.Contains name)

let moduleOf (file: string) =
    let source = File.ReadAllText file
    let parts = (Path.GetFileNameWithoutExtension file).Split('.')
    let stream = if parts.Length > 1 then parts.[parts.Length - 1] else "Module1"
    let m = M.macroModule stream source
    if file.EndsWith ".cls" then { m with Kind = Rbt.Vb.Api.Class } else m

type Outcome = { Name: string; Status: string; Detail: string; Seconds: float; Changes: string }

let describe (before: Workbook) (after: Workbook) =
    let sheets = after.SheetOrder |> List.map (fun id -> after.Sheets.[id])
    let cells = sheets |> List.sumBy (fun s -> s.Model.Cells.Count)
    let values = sheets |> List.sumBy (fun s -> s.Model.Cells |> Map.filter (fun _ c -> c.Value <> CellValue.Empty) |> Map.count)
    let formulas = sheets |> List.sumBy (fun s -> s.Model.Cells |> Map.filter (fun _ c -> match c.Value with Formula _ -> true | _ -> false) |> Map.count)
    let fills = sheets |> List.collect (fun s -> s.Model.Cells |> Map.toList |> List.choose (fun (_, c) -> c.Format.BgColor))
    let widths = sheets |> List.sumBy (fun s -> s.Model.ColumnWidths.Count)
    let heights = sheets |> List.sumBy (fun s -> s.Model.RowHeights.Count)
    let merged = sheets |> List.sumBy (fun s -> s.Model.MergedCells.Length)
    let drawings = sheets |> List.sumBy (fun s -> s.Model.Drawings.Length)
    sprintf "%d sheet(s), %d cells (%d values, %d formulas), %d filled (%d colours), %d col widths, %d row heights, %d merges, %d drawings"
        sheets.Length cells values formulas fills.Length (fills |> List.distinct |> List.length) widths heights merged drawings

let bgr (color: string) =
    let r = Convert.ToInt32(color.Substring(1, 2), 16)
    let g = Convert.ToInt32(color.Substring(3, 2), 16)
    let b = Convert.ToInt32(color.Substring(5, 2), 16)
    r + g * 256 + b * 65536

let columnName (col: int) =
    let mutable n = col + 1
    let mutable name = ""
    while n > 0 do
        name <- string (char (int 'A' + (n - 1) % 26)) + name
        n <- (n - 1) / 26
    name

/// The active sheet's cells with a value or a fill: address, Value2 as text, fill (BGR) or "-".
let dump (wb: Workbook) =
    let model = wb.Sheets.[wb.ActiveSheet].Model
    model.Cells |> Map.toList
    |> List.sortBy (fun (a, _) -> a.Row, a.Col)
    |> List.choose (fun (a, cell) ->
        let text =
            match unwrapFormulaValue cell.Value with
            | CellValue.Empty -> ""
            | Number n -> n.ToString("R", Globalization.CultureInfo.InvariantCulture)
            | String s -> s.Replace("\t", " ").Replace("\n", " ")
            | Boolean b -> if b then "True" else "False"
            | Error e -> "#" + e
            | other -> string other
        let fill = cell.Format.BgColor |> Option.map (bgr >> string) |> Option.defaultValue "-"
        if text = "" && fill = "-" then None else Some (sprintf "%s%d\t%s\t%s" (columnName a.Col) (a.Row + 1) text fill))

/// The modules in a blank macro-enabled workbook (Office compiles the source-only project when it opens it).
let writeXlsm (path: string) (modules: Rbt.Vb.Api.ModuleSource list) =
    let blank = File.ReadAllBytes(Path.Combine(__SOURCE_DIRECTORY__, "../test-documents-excel/pandas-blank.xlsx"))
    let injected = VbaProject.injectIntoXlsx blank "Module1" ""
    let project =
        modules |> List.map (fun m ->
            let kind = match m.Kind with Rbt.Vb.Api.Class -> VbaProject.ClassModule | Rbt.Vb.Api.Document -> VbaProject.DocumentModule | _ -> VbaProject.StandardModule
            m.Name, kind, m.Source)
        |> VbaProject.writeProjectModules
    use stream = new MemoryStream()
    stream.Write(injected, 0, injected.Length)
    (
        use archive = new IO.Compression.ZipArchive(stream, IO.Compression.ZipArchiveMode.Update, true)
        archive.GetEntry("xl/vbaProject.bin").Delete()
        use entry = archive.CreateEntry("xl/vbaProject.bin").Open()
        entry.Write(project, 0, project.Length)
    )
    File.WriteAllBytes(path, stream.ToArray())

let runOne (name: string, files: string[]) =
    let sources = files |> Array.map File.ReadAllText |> String.concat "\n"
    let entry = header "Entry" sources |> List.tryHead |> Option.defaultValue "Main"
    let inputs = header "Input" sources
    let limit = header "Limit" sources |> List.tryHead |> Option.map int |> Option.defaultValue 600_000_000
    let mutable inputIndex = 0
    let output = ResizeArray<string>()
    let ticks = header "Ticks" sources |> List.tryHead |> Option.map int |> Option.defaultValue 0
    let presses =
        header "Press" sources |> List.collect (fun line -> line.Split(';') |> Array.toList)
        |> List.choose (fun p -> match p.Trim().Split('=') with [| t; k |] -> Some (int t, k.Trim().ToUpperInvariant()) | _ -> None)
    let scheduled = Collections.Generic.Queue<string>()
    let keys = Collections.Generic.Dictionary<string, string>()
    let hooks =
        { M.defaultHooks with
            Output = output.Add
            FileName = name + ".xlsm"
            MessageBox = (fun prompt _ _ -> output.Add("MsgBox: " + prompt); 1)
            InputBox = (fun _ _ _ ->
                match inputs with
                | [] -> None
                | _ ->
                    let answer = inputs.[min inputIndex (inputs.Length - 1)]
                    inputIndex <- inputIndex + 1
                    Some answer)
            OnTime = (fun _ procedure schedule -> if schedule then scheduled.Enqueue procedure)
            OnKey = (fun key procedure -> match procedure with Some p -> keys.[key.ToUpperInvariant()] <- p | None -> keys.Remove(key.ToUpperInvariant()) |> ignore) }
    let modules = files |> Array.sort |> Array.map moduleOf |> Array.toList
    let beforeName = match header "Before" sources with b :: _ -> "__" + b | [] -> ""
    xlsmDir |> Option.iter (fun dir -> Directory.CreateDirectory dir |> ignore; writeXlsm (Path.Combine(dir, sprintf "%s__%s%s.xlsm" name entry beforeName)) modules)
    match (try M.MacroProgram.Compile modules with e -> Result.Error [ { Detail = ""; Kind = Rbt.Vb.Api.Compile; Number = 0; Description = "engine exception: " + e.Message; Source = ""; Line = 0; Column = 0; Procedure = ""; Module = "" } ]) with
    | Result.Error failures ->
        { Name = name; Status = "COMPILE"; Detail = failures |> List.truncate 3 |> List.map Rbt.Vb.Api.VbFailure.format |> String.concat "; "; Seconds = 0.0; Changes = "" }
    | Result.Ok program ->
        // (A macro that wants the user to have done something first: a Sub run before it, on the fresh workbook.)
        let start =
            match header "Before" sources with
            | before :: _ ->
                match program.Run(createWorkbook (), hooks, before) with
                | M.MacroDone (wb, _) -> wb
                | M.MacroFailed f -> failwithf "%s: %s" before (Rbt.Vb.Api.VbFailure.format f)
            | [] -> createWorkbook ()
        let watch = Diagnostics.Stopwatch.StartNew()
        let result = try Choice1Of2 (program.Run(start, hooks, entry, maxInstructions = limit)) with e -> Choice2Of2 e
        // (The procedures it asked for later, as the grid runs them when idle; keys pressed between them.)
        let mutable result = result
        let mutable tick = 0
        while tick < ticks && scheduled.Count > 0 && (match result with Choice1Of2 (M.MacroDone _) -> true | _ -> false) do
            tick <- tick + 1
            let runOn (procedure: string) =
                match result with
                | Choice1Of2 (M.MacroDone (wb, _)) -> result <- (try Choice1Of2 (program.Run(wb, hooks, procedure, maxInstructions = limit)) with e -> Choice2Of2 e)
                | _ -> ()
            for (at, key) in presses do
                if at = tick then match keys.TryGetValue key with | true, p when p <> "" -> runOn p | _ -> ()
            runOn (scheduled.Dequeue())
        if ticks > 0 then output.Add(sprintf "%d ticks" tick)
        let seconds = watch.Elapsed.TotalSeconds
        let printed = if output.Count = 0 then "" else " | output: " + (output |> Seq.truncate 3 |> String.concat " / ")
        match result with
        | Choice1Of2 (M.MacroDone (wb, _)) ->
            dumpDir |> Option.iter (fun dir ->
                Directory.CreateDirectory dir |> ignore
                File.WriteAllLines(Path.Combine(dir, name + ".tsv"), sprintf "status\tok\t%.2f" seconds :: dump wb))
            { Name = name; Status = "ok"; Detail = printed.TrimStart(' ', '|'); Seconds = seconds; Changes = describe start wb }
        | Choice1Of2 (M.MacroFailed failure) ->
            { Name = name; Status = "FAIL"; Detail = sprintf "error %d %s%s (%s.%s line %d)%s" failure.Number failure.Description (if failure.Detail = "" then "" else " [" + failure.Detail + "]") failure.Module failure.Procedure failure.Line printed; Seconds = seconds; Changes = "" }
        | Choice2Of2 e -> { Name = name; Status = "CRASH"; Detail = e.GetType().Name + ": " + e.Message.Split('\n').[0]; Seconds = seconds; Changes = "" }

let outcomes = groups |> Array.map (fun g -> let o = runOne g in printfn "%-5s %-22s %6.2fs  %s  %s" (if o.Status = "ok" then "ok" else o.Status) o.Name o.Seconds o.Detail o.Changes; o)
let passed = outcomes |> Array.filter (fun o -> o.Status = "ok") |> Array.length
printfn "\n%d of %d showcase macros ran" passed outcomes.Length
