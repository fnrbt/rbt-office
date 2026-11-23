// dotnet fsi tools/excel-parity/xirr-cells-engine.fsx /tmp/xirr-threshold-cells.tsv
// Evaluate the formulas of any cells TSV (sheet, address, formula or label; as
// made by make-xirr-domain-probe.py or make-xirr-threshold-probe.py) with the
// engine, one result per formula line: sheet, address, value.
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
open System
open System.Globalization
open System.IO
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
let model = createEmpty 100 26
for line in File.ReadAllLines(fsi.CommandLineArgs.[1]) do
    let parts = line.Split('\t')
    if parts.Length >= 3 && parts.[2].StartsWith("=") then
        let value =
            match evaluateFormula parts.[2] model None with
            | CellValue.Number n -> n.ToString("R", CultureInfo.InvariantCulture)
            | CellValue.Error e -> "#" + e
            | other -> sprintf "%A" other
        printfn "%s\t%s\t%s" parts.[0] parts.[1] value
