// Compile errors, ours against Excel's: each case of a folder (tests/vba-parity/compile-cases: a standard
// module <case>.bas, class modules <case>.<Class>.cls) compiled as Excel's object model compiles a project
// (MacroHost: Excel's type library, early binding), its first compile error or "ok", next to what Excel
// gave (EXCEL.tsv from compile-cases.ps1, run with tools/vba-parity/compile-on-vm.sh).
//
//   dotnet build src/Spreadsheet && dotnet fsi tools/vba-parity/compile-parity.fsx CASES EXCEL.tsv
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Vb.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open System
open System.IO
module M = Rbt.Spreadsheet.MacroHost

let folder, excelPath = fsi.CommandLineArgs.[1], fsi.CommandLineArgs.[2]
let normal (text: string) = Text.RegularExpressions.Regex.Replace(text.Trim(), @"\s+", " ")
let excel =
    File.ReadAllLines excelPath
    |> Array.choose (fun line -> match line.Split('\t') with [| case; result |] -> Some (case, normal result) | _ -> None)
    |> Map.ofArray
let mutable same = 0
let cases = Directory.GetFiles(folder, "*.bas") |> Array.sort
for path in cases do
    let case = Path.GetFileNameWithoutExtension path
    let modules =
        Rbt.Vb.Api.ModuleSource.standard "Module1" (File.ReadAllText path)
        :: [ for cls in Directory.GetFiles(folder, case + ".*.cls") ->
                let name = Path.GetFileNameWithoutExtension(cls).Substring(case.Length + 1)
                Rbt.Vb.Api.ModuleSource.classModule name (File.ReadAllText cls) ]
    let ours =
        match M.MacroProgram.Compile modules with
        | Result.Ok _ -> "ok"
        | Result.Error (f :: _) -> normal ("Compile error: " + f.Description)
        | Result.Error [] -> "ok"
    let theirs = excel |> Map.tryFind case |> Option.defaultValue "(no result)"
    if ours = theirs then same <- same + 1
    else printfn "DIFF %-30s Excel: %s / ours: %s" case theirs ours
printfn "Compile parity: %d of %d cases as Excel" same cases.Length
