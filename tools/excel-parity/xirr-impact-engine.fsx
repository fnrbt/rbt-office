// dotnet fsi tools/excel-parity/xirr-impact-engine.fsx /tmp/xirr-impact.xlsx
// Recalculate the generated workbook without cached formula values.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
open System
open System.IO
open Rbt.Spreadsheet.Model
let path = fsi.CommandLineArgs.[1]
let imported = Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook (XlsxReader.loadBytesNative (File.ReadAllBytes path) |> Option.get)
let cleared =
    { imported.Context with
        Sheets =
            imported.Context.Sheets
            |> Map.map (fun _ sheet ->
                let model = resolveSheet sheet.Model
                { sheet with Model = { model with Cells = model.Cells |> Map.map (fun _ cell -> match cell.Value with Formula (text, _) -> { cell with Value = Formula (text, Empty) } | _ -> cell) } }) }
let result = Rbt.Spreadsheet.Formula.recalculateWorkbook cleared
let sheet = result.Sheets |> Map.toSeq |> Seq.map snd |> Seq.find (fun s -> s.Name = "Impact")
let show value =
    match value with
    | Number n -> n.ToString("R", Globalization.CultureInfo.InvariantCulture)
    | Boolean b -> if b then "TRUE" else "FALSE"
    | String s -> s
    | Error e -> if e.StartsWith "#" then e else "#" + e
    | x -> sprintf "%A" x
for row in [1..11] @ [14..25] do
    let values = [0; 6; 8; 9; 10; 11] |> List.map (fun col -> getCellValue { Row = row; Col = col } sheet.Model |> unwrapFormulaValue |> show)
    printfn "%s" (String.concat "\t" values)
let explicit = getCellValue { Row = 25; Col = 5 } sheet.Model |> unwrapFormulaValue |> show
printfn "month-12-explicit-guess\t%s" explicit
