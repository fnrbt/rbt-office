// What our engine computes for every formula cell of the corpus, against what Excel computes
// (tools/excel-parity/dump-values.vbs: values.tsv). Each workbook is imported, every formula's
// cached result is thrown away, and the workbook is recalculated from nothing.
//
//   dotnet fsi tools/excel-parity/compare-values.fsx VALUES.tsv DOCS_DIR [--list N] [--cached | --cached-new]
//
// --cached: every workbook is compared against the values its file keeps (what the Excel that saved it
// computed: Excel 365 for most), not against VALUES.tsv - every formula cell and every cell a dynamic
// array spilled into. --cached-new: so are the workbooks Excel 2019 cannot calculate - those whose
// formulas call a function it does not have (`newFunctions`: dynamic arrays, `@`, LET ...) or that keep
// dynamic arrays (xl/metadata.xml); the others against VALUES.tsv.
//
// Cells are left out where Excel cannot say: #NAME? (a function this Excel does not have: the VM's
// is 2019's engine) and the volatile functions (RAND, NOW, TODAY ...). Numbers agree to 1e-9 relative;
// those within 1e-6 are counted apart as near (XIRR: Excel stops within its stated 0.000001 percent,
// we give the root).
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open System
open System.IO
open Rbt.Spreadsheet.Model
module Bridge = Rbt.Spreadsheet.XlsxBridge
module Wb = Rbt.Spreadsheet.Workbook

let args = fsi.CommandLineArgs |> Array.skip 1
let valuesPath, docsDir = args.[0], args.[1]
let listCount = match Array.tryFindIndex ((=) "--list") args with Some i -> int args.[i + 1] | None -> 5
let cachedAll = Array.contains "--cached" args
let cachedNew = Array.contains "--cached-new" args

type Expected = { Sheet: string; Address: string; Kind: string; Text: string }
let expected =
    File.ReadAllLines(valuesPath, Text.Encoding.Unicode)
    |> Array.choose (fun line ->
        match line.Split '\t' with
        | [| book; sheet; address; kind; text |] when kind <> "ERROR" -> Some (book, { Sheet = sheet; Address = address; Kind = kind; Text = text })
        | _ -> None)
    |> Array.groupBy fst
    |> Array.map (fun (book, rows) -> book, rows |> Array.map snd)

/// Functions Excel 2019 does not have (its #NAME?, or its implicit intersection where 365 spills).
let newFunctions =
    Text.RegularExpressions.Regex(
        "\\b(SINGLE|ANCHORARRAY|SEQUENCE|FILTER|SORT|SORTBY|UNIQUE|XLOOKUP|XMATCH|LET|LAMBDA|RANDARRAY|MAP|REDUCE|SCAN|BYROW|BYCOL|MAKEARRAY|ISOMITTED|TEXTBEFORE|TEXTAFTER|TEXTSPLIT|VSTACK|HSTACK|TOROW|TOCOL|WRAPROWS|WRAPCOLS|TAKE|DROP|CHOOSEROWS|CHOOSECOLS|EXPAND|GROUPBY|PIVOTBY|PERCENTOF|ARRAYTOTEXT|VALUETOTEXT|TRIMRANGE|REGEXTEST|REGEXEXTRACT|REGEXREPLACE|IMAGE)\\s*\\(",
        Text.RegularExpressions.RegexOptions.IgnoreCase)
let keepsDynamicArrays (path: string) =
    use zip = IO.Compression.ZipFile.OpenRead path
    zip.Entries |> Seq.exists (fun entry -> entry.FullName = "xl/metadata.xml")

let volatile = Text.RegularExpressions.Regex("\\b(RAND|RANDBETWEEN|RANDARRAY|NOW|TODAY|INFO|CELL|OFFSET|INDIRECT)\\s*\\(", Text.RegularExpressions.RegexOptions.IgnoreCase)

let addressOf (address: string) =
    let m = Text.RegularExpressions.Regex.Match(address, "^([A-Z]+)([0-9]+)$")
    let col = m.Groups.[1].Value |> Seq.fold (fun acc c -> acc * 26 + int c - 64) 0
    { Row = int m.Groups.[2].Value - 1; Col = col - 1 }

let agree (want: Expected) (got: CellValue) =
    match want.Kind, got with
    | "n", Number n ->
        let w = Double.Parse(want.Text, Globalization.CultureInfo.InvariantCulture)
        abs (n - w) <= 1e-9 * max 1.0 (abs w)
    | "n", Boolean b -> false
    | "s", String s -> s.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ') = want.Text
    | "s", Empty -> want.Text = ""
    | "b", Boolean b -> (if b then "True" else "False") = want.Text || (if b then "TRUE" else "FALSE") = want.Text
    | "e", Error code -> (if code.StartsWith "#" then code else "#" + code) = want.Text
    | "x", Empty -> true
    | "n", Empty -> Double.Parse(want.Text, Globalization.CultureInfo.InvariantCulture) = 0.0
    | _ -> false

let near (want: Expected) (got: CellValue) =
    match want.Kind, got with
    | "n", Number n ->
        let w = Double.Parse(want.Text, Globalization.CultureInfo.InvariantCulture)
        abs (n - w) <= 1e-6 * max 1e-3 (abs w)
    | _ -> false

let show (value: CellValue) =
    match value with
    | Number n -> sprintf "n %.15g" n
    | String s -> sprintf "s %s" s
    | Boolean b -> sprintf "b %b" b
    | Error e -> sprintf "e %s" e
    | Empty -> "x"
    | Formula _ -> "formula"

let mutable totalCompared = 0
let mutable totalAgreed = 0
let mutable totalSkipped = 0
let mutable totalNear = 0
let lines = ResizeArray<string>()
let colName c = let rec go c = if c < 26 then string (char (65 + c)) else go (c / 26 - 1) + string (char (65 + c % 26)) in go c
let kindOf (value: CellValue) =
    match value with
    | Number n -> "n", n.ToString("0.00000000000000E+00", Globalization.CultureInfo.InvariantCulture)
    | String s -> "s", s.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ')
    | Boolean b -> "b", (if b then "True" else "False")
    | Error e -> "e", (if e.StartsWith "#" then e else "#" + e)
    | _ -> "x", ""
for book, cells in expected |> Array.sortBy fst do
    let path = Path.Combine(docsDir, book)
    if File.Exists path then
        try
            let imported = Bridge.Import.importWorkbook (XlsxReader.loadBytesNative (File.ReadAllBytes path) |> Option.get)
            let context = imported.Context
            let models = context.Sheets |> Map.toList |> List.map (fun (_, sheet) -> sheet.Name, resolveSheet sheet.Model)
            let usesNew =
                models |> List.exists (fun (_, model) -> model.Cells |> Map.exists (fun _ cell -> match cell.Value with Formula (text, _) -> newFunctions.IsMatch text | _ -> false))
            let fromFile = cachedAll || (cachedNew && (usesNew || keepsDynamicArrays path))
            // (Against the file: every formula cell and every spilled cell, with the value the file keeps.)
            let cells =
                if not fromFile then cells
                else
                    [| for name, model in models do
                        for KeyValue (at, cell) in model.Cells do
                            let kept =
                                match cell.Value with
                                | Formula (_, cached) -> Some (unwrapFormulaValue cached)
                                | value when Map.containsKey at model.Spill -> Some value
                                | _ -> None
                            match kept with
                            | Some value when value <> Empty ->
                                let kind, text = kindOf value
                                yield { Sheet = name; Address = sprintf "%s%d" (colName at.Col) (at.Row + 1); Kind = kind; Text = text }
                            | _ -> () |]
            // (Every result thrown away: the engine computes each from nothing.)
            let cleared =
                { context with
                    Sheets =
                        context.Sheets
                        |> Map.map (fun _ sheet ->
                            let model = resolveSheet sheet.Model
                            { sheet with Model = { model with Cells = model.Cells |> Map.map (fun _ cell -> match cell.Value with Formula (text, _) -> { cell with Value = Formula (text, Empty) } | _ -> cell) } }) }
            let recalculated = Rbt.Spreadsheet.Formula.recalculateWorkbook cleared
            let byName = recalculated.Sheets |> Map.toList |> List.map (fun (_, sheet) -> sheet.Name, sheet.Model) |> Map.ofList
            let mutable compared = 0
            let mutable agreed = 0
            let mutable skipped = 0
            let mutable nearly = 0
            let differences = ResizeArray<string>()
            for want in cells do
                match Map.tryFind want.Sheet byName with
                | Some model ->
                    let at = addressOf want.Address
                    let cell = getCell at model
                    let formula = match cell |> Option.map (fun c -> c.Value) with Some (Formula (text, _)) -> text | _ -> ""
                    if (want.Text = "#NAME?" && not fromFile) || volatile.IsMatch formula then skipped <- skipped + 1
                    else
                        compared <- compared + 1
                        let got = getCellValue at model |> unwrapFormulaValue
                        if agree want got then agreed <- agreed + 1
                        elif near want got then nearly <- nearly + 1
                        elif differences.Count < listCount then differences.Add (sprintf "    %s!%s =%s  Excel %s %s, ours %s" want.Sheet want.Address formula want.Kind want.Text (show got))
                | None -> skipped <- skipped + 1
            totalCompared <- totalCompared + compared
            totalAgreed <- totalAgreed + agreed
            totalSkipped <- totalSkipped + skipped
            totalNear <- totalNear + nearly
            lines.Add (sprintf "%s %s: %d of %d agree%s (%d left out)%s" (if agreed + nearly = compared then "ok  " else "DIFF") book agreed compared (if nearly > 0 then sprintf ", %d near" nearly else "") skipped (if fromFile then " [against the file's values]" else ""))
            lines.AddRange differences
        with error -> lines.Add (sprintf "FAIL %s: %s" book (error.Message.Split('\n').[0]))
for line in lines do printfn "%s" line
printfn "\nFORMULA VALUES AGAINST EXCEL: %d of %d cells agree (%.2f%%), %d near (1e-6), %d left out, %d workbooks"
    totalAgreed totalCompared (100.0 * float totalAgreed / float (max 1 totalCompared)) totalNear totalSkipped expected.Length
