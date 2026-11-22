#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open System
open System.IO
open Rbt.Spreadsheet.Model
module Bridge = Rbt.Spreadsheet.XlsxBridge
module Spark = Rbt.Spreadsheet.Sparklines

// Sparklines as Excel 2019 draws them. tests/test-documents-excel/synthetic-sparkline-looks.xlsx was made in Excel
// (tools/excel-parity/build-sparklines.vbs): one sparkline a row in A1:A19, cells of 200 x 40 px, its data in C:N.
// The numbers below are Excel's picture of it (shoot-sheets.ps1 at 100 %): pixels from the cell's top left grid line.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-56s expected %A  got %A" label expected got)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let model =
    let wb = XlsxReader.loadBytesNative (File.ReadAllBytes(Path.Combine(root, "test-documents-excel/synthetic-sparkline-looks.xlsx"))) |> Option.get
    let book = Bridge.Import.importWorkbook wb
    resolveSheet book.Sheets.[0].Model

eq "19 groups, one sparkline each, read from the sheet's extension list" (model.Sparklines |> List.sumBy (fun g -> g.Items.Length)) 19
let groupAt (row: int) = model.Sparklines |> List.find (fun g -> g.Items |> List.exists (fun (at, _) -> at = { Row = row - 1; Col = 0 }))
let valuesAt (row: int) =
    let _, formula = (groupAt row).Items.Head
    let range = formula.Substring(formula.IndexOf '!' + 1).Replace("$", "")
    let parts = range.Split ':'
    let col (text: string) = (text |> Seq.takeWhile Char.IsLetter |> Seq.fold (fun a c -> a * 26 + int c - int 'A' + 1) 0) - 1
    [| for c in col parts.[0] .. col parts.[1] ->
        match getCell { Row = row - 1; Col = c } model |> Option.map (fun cell -> unwrapFormulaValue cell.Value) with
        | Some (CellValue.Number v) -> Some v
        | _ -> None |]
let marksAt (row: int) =
    let values = valuesAt row
    Spark.marks (groupAt row) values (Spark.range values |> Option.get) 200.0 40.0
/// A column as Excel's pixels: first and last column, first and last row.
let bars (row: int) =
    marksAt row |> List.choose (function
        | Spark.SparkBar (x, y, w, h, _) when h > 1.0 || w < 100.0 -> Some (int x, int (x + w) - 1, int y, int (y + h) - 1)
        | _ -> None)

eq "group kinds and options as written" ((groupAt 12).Kind, (groupAt 17).Kind, (groupAt 3).Markers, (groupAt 8).EmptyAs, (groupAt 9).LineWeight)
    ("column", "stacked", true, "span", 2.25)
eq "ColRamp 1..5: the lowest a 1-px line, each from the bottom of the plot (row 36), tops 36 29 21 13 5"
    (bars 11) [ (9, 35, 36, 36); (48, 74, 29, 36); (87, 113, 21, 36); (126, 152, 13, 36); (165, 191, 5, 36) ]
eq "ColNeg: up from 0 (row 25), down from row 26"
    (bars 12 |> List.map (fun (_, _, y0, y1) -> y0, y1)) [ (12, 25); (26, 32); (9, 25); (26, 29); (5, 25); (19, 25); (26, 36) ]
eq "ColMany 1..12: tops as Excel's (7 at row 19.09 drawn from 19)"
    (bars 14 |> List.map (fun (_, _, y0, _) -> y0)) [ 36; 34; 31; 28; 25; 22; 19; 17; 14; 11; 8; 5 ]
eq "WinLoss: wins rows 5..20, losses 21..36"
    (bars 17 |> List.map (fun (_, _, y0, y1) -> y0, y1)) [ (5, 20); (21, 36); (5, 20); (5, 20); (21, 36); (5, 20); (21, 36) ]
eq "ColHigh: the first, highest and last columns red (low = first)"
    (marksAt 16 |> List.map (function Spark.SparkBar (_, _, _, _, c) -> c | _ -> "") ) [ "#D00000"; "#376092"; "#376092"; "#376092"; "#376092"; "#D00000"; "#D00000" ]
let path (row: int) = marksAt row |> List.pick (function Spark.SparkPath (points, _, w) -> Some (points, w) | _ -> None)
eq "LineZigzag: 7 points from x 16.4 to 183.6, y 5.5 (highest) .. 36.5 (lowest)"
    (path 2 |> fst |> List.map (fun (x, y) -> Math.Round(x, 1), Math.Round(y, 1)) |> fun ps -> List.head ps, List.last ps, ps |> List.map snd |> List.min, ps |> List.map snd |> List.max)
    ((16.4, 36.5), (183.6, 28.8), 5.5, 36.5)
eq "LineGap: two pieces, the lone 9 not drawn" (marksAt 6 |> List.filter (function Spark.SparkPath _ -> true | _ -> false) |> List.length) 2
eq "LineSpan: one piece over the empty cells" (marksAt 8 |> List.filter (function Spark.SparkPath _ -> true | _ -> false) |> List.length) 1
eq "LineWeight 2.25 pt: 2 px" (path 9 |> snd) 2.0
eq "LineFlat: all at the top" (path 10 |> fst |> List.map snd |> List.distinct) [ 5.5 ]
eq "LineAxis: the axis at 0 (row 25), x 3 .. 198"
    (marksAt 5 |> List.pick (function Spark.SparkBar (x, y, w, _, c) -> Some (x, y, x + w - 1.0, c) | _ -> None)) (3.0, 25.0, 198.0, "#000000")

printfn "stage65-sparklines: %d passed, %d failed" pass fail
if fail > 0 then exit 1
