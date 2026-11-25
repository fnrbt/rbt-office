#load "function-table-kit.fsx"
open System
open System.IO
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open ``Function-table-kit``

// The functions no other stage named (found by matching every Excel function name against the tests,
// 2026-11-25), against Excel 2019 bit for bit: ACOSH ASINH ATANH COT TAN, COVAR COVARIANCE.P
// COVARIANCE.S PEARSON STEYX STDEV VARP MODE.SNGL, PERCENTILE.INC QUARTILE.INC, DEC2OCT OCT2DEC
// (sheet S, typed on an empty sheet) and DCOUNTA DPRODUCT (sheet D: a database typed into F1:I7 and
// criteria into K1:Q2, which Excel needs as ranges); then SIN COS TAN COT SEC CSC (T: the cells the
// x87-pi reduction of ExcelTrig was found on, T2: 540 random ones it was checked on after), every base
// conversion (B) and MODE / QUARTILE's argument rules (M). Cells tools/excel-parity/gap-functions-cells.tsv,
// Excel's values gap-functions-excel.tsv (eval-cells-on-vm.sh). Numbers bit for bit, except where
// Excel's sums are not reproduced to the last bit (STDEV, VARP, PEARSON, STEYX: the order it adds in)
// and SEC / CSC (4 of the 540 random cells an ulp off): there 15 significant digits. Then RANDARRAY
// (Microsoft 365 only: its documented rules) and the provider functions, without a provider and with one.
let dir = Path.Combine(__SOURCE_DIRECTORY__, "..", "..", "tools", "excel-parity")
let lines (name: string) =
    File.ReadAllLines(Path.Combine(dir, name)) |> Array.filter (fun l -> l.Length > 0) |> Array.map (fun l -> l.Split '\t')
let cells = lines "gap-functions-cells.tsv"
let excel = lines "gap-functions-excel.tsv" |> Array.map (fun f -> (f.[0], f.[1]), f) |> dict

let decode (text: string) =
    let negative = text.StartsWith "-"
    let parts = text.TrimStart('-').Split ':'
    let v = (float (int64 parts.[0]) * 67108864.0 + float (int64 parts.[1])) * Math.Pow(2.0, float (int parts.[2]))
    if negative then -v else v

let addressOf (a: string) =
    let col = int a.[0] - int 'A'
    (int (a.Substring 1) - 1, col)
let constant (text: string) =
    match Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
    | true, x -> CellValue.Number x
    | _ -> CellValue.String text
let databaseSheet =
    modelOf [ for c in cells do if c.[0] = "D" && not (c.[2].StartsWith "=") then addressOf c.[1], constant c.[2] ]

for c in cells do
    if c.[2].StartsWith "=" then
        let x = excel.[(c.[0], c.[1])]
        let m = if c.[0] = "D" then databaseSheet else empty
        let got = value m (c.[2].Replace("D!", ""))
        let toDigits = [ "=STDEV("; "=VARP("; "=PEARSON("; "=STEYX("; "=SEC("; "=CSC(" ] |> List.exists c.[2].StartsWith
        let ok =
            match x.[2], got with
            | "n", CellValue.Number v when toDigits -> Rbt.Spreadsheet.NumberText.sameToSignificantDigits 15 v (decode x.[5])
            | "n", CellValue.Number v -> BitConverter.DoubleToInt64Bits (v + 0.0) = BitConverter.DoubleToInt64Bits (decode x.[5] + 0.0)
            | "e", CellValue.Error code -> "#" + code = x.[4]
            | "s", CellValue.String s -> s = x.[4]
            | "b", CellValue.Boolean b -> (if b then "TRUE" else "FALSE") = x.[4].ToUpperInvariant()
            | _ -> false
        if ok then pass <- pass + 1
        else
            fail <- fail + 1
            printfn "FAIL %s!%s %s: Excel %s, engine %A" c.[0] c.[1] c.[2] x.[4] got

// ----- RANDARRAY (Microsoft 365; its documented rules: rows and columns default to 1, min 0, max 1,
// whole_number FALSE; max < min is #VALUE!; numbers in [min, max), whole numbers in [min, max])
let n m f = match value m f with CellValue.Number x -> x | v -> (printfn "FAIL %s %A" f v; nan)
eq "RANDARRAY() is one number" (arrayOf empty "=RANDARRAY()" |> Array.map Array.length) [| 1 |]
eq "RANDARRAY(3,4) is 3 x 4" (arrayOf empty "=RANDARRAY(3,4)" |> fun a -> a.Length, a.[0].Length) (3, 4)
eq "RANDARRAY(2) is 2 x 1" (arrayOf empty "=RANDARRAY(2)" |> fun a -> a.Length, a.[0].Length) (2, 1)
let between lo hi whole (a: CellValue[][]) =
    a |> Array.forall (Array.forall (function
        | CellValue.Number x -> x >= lo && x <= hi && (not whole || x = Math.Floor x)
        | _ -> false))
eq "RANDARRAY in [0,1)" (arrayOf empty "=RANDARRAY(50,4)" |> between 0.0 1.0 false) true
eq "RANDARRAY(…,5,7)" (arrayOf empty "=RANDARRAY(50,4,5,7)" |> between 5.0 7.0 false) true
eq "RANDARRAY whole numbers" (arrayOf empty "=RANDARRAY(50,4,1,6,TRUE)" |> between 1.0 6.0 true) true
eq "RANDARRAY(…,5,5,TRUE) is all 5" (n empty "=SUM(RANDARRAY(10,10,5,5,TRUE))") 500.0
error empty "=RANDARRAY(2,2,5,1)" Err.value
error empty "=RANDARRAY(-1)" Err.value
error empty "=RANDARRAY(2,2,1,1,\"x\")" Err.value

// ----- the provider functions: without a provider each gives its error; with one, its answer
for (name, noProvider, _) in providerFunctionSpecs do
    let formula = sprintf "=%s(\"a\",\"b\")" name
    unregisterFunctionProvider name
    error empty formula noProvider
    eq ("register " + name) (registerFunctionProvider name { Invoke = fun call answer -> answer [| [| S ("answer of " + call.Name) |] |] }) true
    text empty formula ("answer of " + name)
    unregisterFunctionProvider name
    error empty formula noProvider

printfn "stage72 functions no other stage named: %d passed, %d failed" pass fail
if fail > 0 then exit 1
