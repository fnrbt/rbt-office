// dotnet fsi tools/excel-parity/function-parity-engine.fsx EXCEL.tsv [CELLS.txt] [--misses N] [--only NAME]
//
//   e.g. dotnet fsi tools/excel-parity/function-parity-engine.fsx tools/excel-parity/bonds-excel.tsv
//
// Compares the engine with Excel on a probe of worksheet functions: EXCEL.tsv what eval-cells-on-vm.sh
// returned for the probe's cells (bonds-excel.tsv, complex-excel.tsv, engineering-excel.tsv here), the
// cells (sheet, address, formula) CELLS.txt or, when not given, what make-<family>-probe.py next to
// EXCEL.tsv writes (the probes are deterministic). Each formula is evaluated by
// Rbt.Spreadsheet.Formula.evaluateFormula on an empty sheet and compared with Excel's value: the kind
// (number, text, error, logical), text and errors exactly, numbers by their exact binary value, their
// 15 significant digits and their relative error. Prints a table by function (cells, exact, 15 digits,
// worst relative error, kind or text mismatches) and the first N mismatches (default 40). Exits 1
// when a cell's kind or text differs. Build src/Spreadsheet first.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
open System
open System.Globalization
open System.IO
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

let args = fsi.CommandLineArgs |> Array.skip 1
let option name = args |> Array.tryFindIndex ((=) name) |> Option.map (fun i -> args.[i + 1])
let positional =
    args |> Array.indexed
    |> Array.filter (fun (i, a) -> not (a.StartsWith "--") && not (i > 0 && args.[i - 1].StartsWith "--"))
    |> Array.map snd
let excelPath = positional.[0]
let missLimit = option "--misses" |> Option.map int |> Option.defaultValue 40
let only = option "--only"

/// The probe's cells: the file given, else the generator's output.
let cellLines : string[] =
    if positional.Length > 1 then File.ReadAllLines positional.[1]
    else
        let family = Path.GetFileName(excelPath).Replace("-excel.tsv", "")
        let generator = Path.Combine(Path.GetDirectoryName(Path.GetFullPath excelPath), "make-" + family + "-probe.py")
        let psi = Diagnostics.ProcessStartInfo("python3", "\"" + generator + "\"")
        psi.RedirectStandardOutput <- true
        psi.UseShellExecute <- false
        psi.StandardOutputEncoding <- Text.Encoding.UTF8
        psi.Environment.["PYTHONIOENCODING"] <- "utf-8"
        use proc = Diagnostics.Process.Start psi
        let text = proc.StandardOutput.ReadToEnd()
        proc.WaitForExit()
        text.Split('\n') |> Array.map (fun l -> l.TrimEnd('\r'))

Rbt.Spreadsheet.Functions.All.installed |> ignore
let model = createEmpty 100 26

let formulas =
    cellLines
    |> Array.choose (fun line ->
        let parts = line.Split([| '\t' |], 3)
        if parts.Length = 3 then Some ((parts.[0], parts.[1]), parts.[2]) else None)
    |> Map.ofArray

/// v = (hi*2^26 + lo) * 2^e
let exactOf (text: string) =
    let neg = text.StartsWith "-"
    let p = text.TrimStart('-').Split(':')
    let m = float (Int64.Parse p.[0]) * 67108864.0 + float (Int64.Parse p.[1])
    let v = m * Math.Pow(2.0, float (Int32.Parse p.[2]))
    if neg then -v else v

let digits15 (x: float) = x.ToString("E14", CultureInfo.InvariantCulture)
let fnName (formula: string) =
    let i = formula.IndexOf '('
    if i > 1 then formula.Substring(1, i - 1) else formula

type Row = { mutable Cells: int; mutable Exact: int; mutable Digits: int; mutable Worst: float; mutable Kind: int }
let table = Collections.Generic.SortedDictionary<string, Row>()
let misses = ResizeArray<string>()
let mutable kindMisses = 0

for line in File.ReadAllLines excelPath do
    let p = line.Split('\t')
    if p.Length >= 5 && p.[0] <> "CIRCULAR" then
        let formula = formulas.[(p.[0], p.[1])]
        let name = fnName formula
        if only.IsNone || only = Some name then
            let row =
                match table.TryGetValue name with
                | true, r -> r
                | _ -> let r = { Cells = 0; Exact = 0; Digits = 0; Worst = 0.0; Kind = 0 } in table.[name] <- r; r
            row.Cells <- row.Cells + 1
            let got = unwrapValue (evaluateFormula formula model None)
            let describe (v: CellValue) =
                match v with
                | CellValue.Number n -> n.ToString("R", CultureInfo.InvariantCulture)
                | CellValue.String s -> "\"" + s + "\""
                | CellValue.Error e -> "#" + e
                | CellValue.Boolean b -> (if b then "TRUE" else "FALSE")
                | CellValue.Empty -> "(empty)"
                | other -> sprintf "%A" other
            let kindMiss () =
                row.Kind <- row.Kind + 1
                kindMisses <- kindMisses + 1
                misses.Add(sprintf "KIND  %-60s excel %s %s  engine %s" formula p.[2] p.[4] (describe got))
            match p.[2], got with
            | "n", CellValue.Number x ->
                let e = exactOf p.[5]
                if x = e then
                    row.Exact <- row.Exact + 1
                    row.Digits <- row.Digits + 1
                else
                    if digits15 x = digits15 e then row.Digits <- row.Digits + 1
                    let rel = if e = 0.0 then abs x else abs (x - e) / abs e
                    row.Worst <- max row.Worst rel
                    if digits15 x <> digits15 e then
                        misses.Add(sprintf "DIFF  %-60s excel %s  engine %s  rel %.2e" formula (e.ToString("R", CultureInfo.InvariantCulture)) (describe got) rel)
            | "s", CellValue.String s when s = p.[3] -> row.Exact <- row.Exact + 1; row.Digits <- row.Digits + 1
            | "e", CellValue.Error e when "#" + e = p.[4] -> row.Exact <- row.Exact + 1; row.Digits <- row.Digits + 1
            | "b", CellValue.Boolean b when (if b then "True" else "False") = p.[3] -> row.Exact <- row.Exact + 1; row.Digits <- row.Digits + 1
            | "x", CellValue.Empty -> row.Exact <- row.Exact + 1; row.Digits <- row.Digits + 1
            | _ -> kindMiss ()

printfn "%-14s %6s %6s %6s %10s %6s" "function" "cells" "exact" "15dig" "worst rel" "kind"
for KeyValue(name, r) in table do
    printfn "%-14s %6d %6d %6d %10.2e %6d" name r.Cells r.Exact r.Digits r.Worst r.Kind
let total f = table.Values |> Seq.sumBy f
printfn "%-14s %6d %6d %6d %10s %6d" "TOTAL" (total (fun r -> r.Cells)) (total (fun r -> r.Exact)) (total (fun r -> r.Digits)) "" (total (fun r -> r.Kind))
// Kind mismatches first, then the digit misses.
let ordered = Seq.append (misses |> Seq.filter (fun m -> m.StartsWith "KIND")) (misses |> Seq.filter (fun m -> m.StartsWith "DIFF"))
for m in ordered |> Seq.truncate missLimit do printfn "%s" m
if kindMisses > 0 then exit 1
