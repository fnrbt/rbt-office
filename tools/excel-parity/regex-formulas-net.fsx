// dotnet fsi tools/excel-parity/regex-formulas-net.fsx OUT.jsonl
// The .NET engine's values (as arrays: a formula that spills, all of it) of the formulas of
// regex-formulas.jsonl (make-regex-formulas.py), and the engine's answers on the cases of
// regex-perl.jsonl, written where regex-js-check.mjs compares the browser build with them.
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
open System.IO
open System.Text.Json
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
module Pcre = Rbt.Spreadsheet.Functions.Pcre

let here = __SOURCE_DIRECTORY__
let out = fsi.CommandLineArgs.[1]
let model = createEmpty 40 12
let valueText (v: CellValue) =
    match unwrapValue v with
    | CellValue.Number n -> "n:" + n.ToString("R")
    | CellValue.String s -> "s:" + s
    | CellValue.Boolean b -> "b:" + (if b then "1" else "0")
    | CellValue.Error e -> "e:" + e
    | _ -> "x"
let lines = ResizeArray<string>()
for line in File.ReadAllLines(Path.Combine(here, "regex-formulas.jsonl")) do
    let formula = JsonSerializer.Deserialize<string> line
    let values =
        match parseFormula formula with
        | Ok expr -> evalArray expr (contextOf model) |> Array.map (Array.map valueText)
        | Result.Error e -> [| [| "parse:" + e |] |]
    lines.Add(JsonSerializer.Serialize(dict [ "formula", box formula; "values", box values ]))
for line in File.ReadAllLines(Path.Combine(here, "regex-perl.jsonl")) do
    let root = (JsonDocument.Parse line).RootElement
    let pattern = root.GetProperty("pattern").GetString()
    let ci = root.GetProperty("i").GetBoolean()
    let cps = Pcre.codePoints (root.GetProperty("subject").GetString())
    let answer =
        try
            let re = Pcre.compileRegex pattern ci
            let first = match Pcre.search re cps 0 with Some c -> box (Array.ofSeq c) | None -> null
            box (dict [ "first", first; "all", box (Pcre.allMatches re cps |> List.map (fun c -> [| c.[0]; c.[1] |]) |> Array.ofList) ])
        with
        | Pcre.PatternError _ -> box "error"
        | Pcre.MatchLimitExceeded -> box "limit"
    lines.Add(JsonSerializer.Serialize(dict [ "pattern", box pattern; "i", box ci; "subject", box (root.GetProperty("subject").GetString()); "answer", answer ]))
File.WriteAllLines(out, lines)
printfn "%d lines" lines.Count
