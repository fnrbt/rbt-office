#load "function-table-kit.fsx"
open System
open System.IO
open Rbt.Spreadsheet.Model
open ``Function-table-kit``

// CEILING, FLOOR, MROUND and ECMA.CEILING against Excel 2019, bit for bit. The cells are
// tools/excel-parity/rounding-multiples-cells.tsv and Excel's values rounding-multiples-excel.tsv
// (eval-cells-on-vm.sh): sheet S a grid of numbers x significances the rules were found on, S2 900
// random cells they were checked on afterwards (all exact without refitting), S3 ECMA.CEILING next to
// CEILING. Excel writes a number as `[-]hi:lo:exp` = (hi x 2^26 + lo) x 2^exp.
let m = empty
let dir = Path.Combine(__SOURCE_DIRECTORY__, "..", "..", "tools", "excel-parity")
let lines (name: string) =
    File.ReadAllLines(Path.Combine(dir, name)) |> Array.filter (fun l -> l.Length > 0) |> Array.map (fun l -> l.Split '\t')
let excel = lines "rounding-multiples-excel.tsv" |> Array.map (fun f -> (f.[0], f.[1]), f) |> dict

let decode (text: string) =
    let negative = text.StartsWith "-"
    let parts = text.TrimStart('-').Split ':'
    let v = (float (int64 parts.[0]) * 67108864.0 + float (int64 parts.[1])) * Math.Pow(2.0, float (int parts.[2]))
    if negative then -v else v

for cell in lines "rounding-multiples-cells.tsv" do
    let x = excel.[(cell.[0], cell.[1])]
    let got = value m cell.[2]
    let ok =
        match x.[2], got with
        | "n", CellValue.Number v -> BitConverter.DoubleToInt64Bits (v + 0.0) = BitConverter.DoubleToInt64Bits (decode x.[5] + 0.0)
        | "e", CellValue.Error code -> "#" + code = x.[4]
        | "s", CellValue.String s -> s = x.[4]
        | "b", CellValue.Boolean b -> (if b then "TRUE" else "FALSE") = x.[4].ToUpperInvariant()
        | _ -> false
    if ok then pass <- pass + 1
    else
        fail <- fail + 1
        printfn "FAIL %s!%s %s: Excel %s, engine %A" cell.[0] cell.[1] cell.[2] x.[4] got

printfn "stage71 rounding to multiples: %d passed, %d failed" pass fail
if fail > 0 then exit 1
