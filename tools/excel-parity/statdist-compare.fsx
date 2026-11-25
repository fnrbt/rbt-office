// dotnet fsi tools/excel-parity/statdist-compare.fsx CELLS-DIR [RESULTS.tsv] [--mismatches OUT.tsv] [--reference REF.tsv]
//
// Evaluates the probe cells made by make-statdist-probe.py (CELLS-DIR/statdist-cells-*.tsv)
// with the engine and compares them with what Excel computed (RESULTS.tsv, by default
// tools/excel-parity/statdist-excel-results.tsv: the eval-cells-on-vm.sh outputs of the
// batches, concatenated). Prints, per function, the cells, how many agree in Excel's 15
// significant digits, how many are the same double, the worst relative difference of the
// numbers and the cells whose kind differs (a number against an error, or two errors).
// With the high-precision values of statdist-reference.py (--reference, by default
// tools/excel-parity/statdist-reference.tsv when it exists) it also counts, of the cells
// it covers, how many of Excel's and of the engine's values are the true value rounded
// to 15 digits.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
open System
open System.Globalization
open System.IO
open System.Text.RegularExpressions
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

let installed = Rbt.Spreadsheet.Functions.All.installed
let args = fsi.CommandLineArgs |> Array.skip 1
let cellsDir = args.[0]
let option name =
    match args |> Array.tryFindIndex ((=) name) with
    | Some i -> Some args.[i + 1]
    | None -> None
let optionValues = [ option "--mismatches"; option "--reference" ] |> List.choose id
let positional = args |> Array.filter (fun a -> not (a.StartsWith "--") && not (List.contains a optionValues))
let resultsPath =
    if positional.Length > 1 then positional.[1]
    else Path.Combine(__SOURCE_DIRECTORY__, "statdist-excel-results.tsv")
let mismatchPath = option "--mismatches"
let referencePath =
    match option "--reference" with
    | Some p -> Some p
    | None ->
        let p = Path.Combine(__SOURCE_DIRECTORY__, "statdist-reference.tsv")
        if File.Exists p then Some p else None

/// The true value rounded to 15 significant digits, from its 21-digit text.
let referenceDigits (text: string) =
    let i = text.IndexOf 'e'
    let negative = text.StartsWith "-"
    let digits = text.Substring(0, i).TrimStart('-').Replace(".", "")
    let exponent = int (text.Substring(i + 1))
    let mutable kept = Numerics.BigInteger.Parse(digits.Substring(0, 15))
    if int digits.[15] - int '0' >= 5 then kept <- kept + Numerics.BigInteger.One
    let mutable mantissa = kept.ToString()
    let mutable e = exponent
    if mantissa.Length > 15 then
        mantissa <- mantissa.Substring(0, 15)
        e <- e + 1
    if mantissa.Trim('0') = "" then "0", 0 else (if negative then "-" else "") + mantissa, e
let reference =
    match referencePath with
    | Some p ->
        File.ReadAllLines p
        |> Array.map (fun l -> let q = l.Split('\t') in (q.[0], q.[1]), referenceDigits q.[3])
        |> dict
    | None -> dict []

/// Excel's results: (sheet, address) -> (type, 15-digit text, exact value).
let excel =
    File.ReadAllLines resultsPath
    |> Array.choose (fun line ->
        let p = line.Split('\t')
        if p.Length >= 4 && p.[0] <> "CIRCULAR" then
            let exact =
                if p.[2] = "n" && p.Length >= 6 then
                    let negative = p.[5].StartsWith "-"
                    let parts = p.[5].TrimStart('-').Split(':')
                    let m = float (int64 parts.[0] * 67108864L + int64 parts.[1])
                    let v = m * Math.Pow(2.0, float (int parts.[2]))
                    Some (if negative then -v else v)
                else None
            Some ((p.[0], p.[1]), (p.[2], p.[3], exact))
        else None)
    |> dict

let parseCell (text: string) =
    match Double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture) with
    | true, n -> CellValue.Number n
    | _ ->
        match text.ToUpperInvariant() with
        | "TRUE" -> CellValue.Boolean true
        | "FALSE" -> CellValue.Boolean false
        | _ -> CellValue.String text

let addressOf (a: string) =
    let m = Regex.Match(a, "^([A-Z]+)([0-9]+)$")
    let col = m.Groups.[1].Value |> Seq.fold (fun acc ch -> acc * 26 + int ch - int 'A' + 1) 0
    { Row = int m.Groups.[2].Value - 1; Col = col - 1 }

/// The 15 significant digits Excel shows: mantissa digits and exponent.
let digits15 (x: float) =
    if x = 0.0 then "0", 0
    else
        let s = x.ToString("E14", CultureInfo.InvariantCulture)
        let i = s.IndexOf 'E'
        s.Substring(0, i).Replace(".", ""), int (s.Substring(i + 1))
let excelDigits (text: string) =
    let i = text.IndexOf 'E'
    let mantissa = text.Substring(0, i).Replace(".", "")
    let exponent = int (text.Substring(i + 1))
    if mantissa.TrimStart('-').Trim('0') = "" then "0", 0 else mantissa, exponent

let functionOf (formula: string) =
    let m = Regex.Match(formula, "^=(?:INDEX\(|ROWS\(|COLUMNS\()?([A-Z][A-Z0-9.]*)\(")
    if m.Success then m.Groups.[1].Value else "?"

type Tally = { mutable Cells: int; mutable Digits15: int; mutable Same: int; mutable Numbers: int; mutable Worst: float; mutable WorstAt: string; mutable KindMismatch: int
               mutable Covered: int; mutable ExcelTrue: int; mutable OursTrue: int }
let tallies = Collections.Generic.Dictionary<string, Tally>()
let mismatches = ResizeArray<string>()

for file in Directory.GetFiles(cellsDir, "statdist-cells-*.tsv") |> Array.sort do
    let lines = File.ReadAllLines file |> Array.map (fun l -> l.Split('\t'))
    let data = lines |> Array.filter (fun p -> not (p.[2].StartsWith "="))
    let cells =
        data |> Array.map (fun p -> addressOf p.[1], { Value = parseCell p.[2]; Format = defaultFormat; Comment = None }) |> Map.ofArray
    let model = { createEmpty 1000 26 with Cells = cells }
    for p in lines |> Array.filter (fun p -> p.[2].StartsWith "=") do
        let formula = p.[2]
        let name = functionOf formula
        let t =
            match tallies.TryGetValue name with
            | true, t -> t
            | _ ->
                let t = { Cells = 0; Digits15 = 0; Same = 0; Numbers = 0; Worst = 0.0; WorstAt = ""; KindMismatch = 0; Covered = 0; ExcelTrue = 0; OursTrue = 0 }
                tallies.[name] <- t
                t
        t.Cells <- t.Cells + 1
        let ours = try unwrapValue (evaluateFormula formula model None) with ex -> CellValue.String ("exception " + ex.Message)
        let kind, shown, exact = excel.[(p.[0], p.[1])]
        let report (what: string) = mismatches.Add(sprintf "%s\t%s\t%s\t%s" name formula shown what)
        match kind, ours with
        | "n", CellValue.Number x ->
            t.Numbers <- t.Numbers + 1
            let e = exact.Value
            if x = e then t.Same <- t.Same + 1
            let truth =
                match reference.TryGetValue((p.[0], p.[1])) with
                | true, d -> Some d
                | _ -> None
            match truth with
            | Some d ->
                t.Covered <- t.Covered + 1
                if excelDigits shown = d then t.ExcelTrue <- t.ExcelTrue + 1
                if digits15 x = d then t.OursTrue <- t.OursTrue + 1
            | None -> ()
            let truthNote =
                match truth with
                | Some d when excelDigits shown = d && digits15 x <> d -> "\tExcel true"
                | Some d when digits15 x = d && excelDigits shown <> d -> "\tours true"
                | Some d when digits15 x <> d -> sprintf "\tneither (true %s e%d)" (fst d) (snd d)
                | _ -> ""
            if digits15 x = excelDigits shown then t.Digits15 <- t.Digits15 + 1
            else
                let rel = if e = 0.0 then abs x else abs (x - e) / abs e
                if rel > t.Worst then
                    t.Worst <- rel
                    t.WorstAt <- formula
                report (x.ToString("R", CultureInfo.InvariantCulture) + sprintf "\trel %.2e" rel + truthNote)
        | "e", CellValue.Error code when "#" + code = shown || code = shown -> t.Digits15 <- t.Digits15 + 1; t.Same <- t.Same + 1
        | _, other ->
            t.KindMismatch <- t.KindMismatch + 1
            report (sprintf "%A" other)

printfn "%-18s %6s %8s %6s %10s %5s %8s %7s %7s  %s" "function" "cells" "15digit" "same" "worst-rel" "kind" "covered" "excel=T" "ours=T" "worst at"
let mutable total = 0
let mutable agree = 0
for KeyValue(name, t) in tallies |> Seq.sortBy (fun kv -> kv.Key) do
    total <- total + t.Cells
    agree <- agree + t.Digits15
    printfn "%-18s %6d %8d %6d %10s %5d %8d %7d %7d  %s" name t.Cells t.Digits15 t.Same (if t.Worst = 0.0 then "-" else sprintf "%.1e" t.Worst) t.KindMismatch t.Covered t.ExcelTrue t.OursTrue t.WorstAt
printfn "total %d cells, %d agree to 15 digits (or the same error)" total agree
match mismatchPath with
| Some path -> File.WriteAllLines(path, mismatches)
| None -> ()
