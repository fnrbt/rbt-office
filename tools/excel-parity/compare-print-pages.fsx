// dotnet fsi tools/excel-parity/compare-print-pages.fsx BOOK.xlsx EXCEL_DIR
// Our pages of every visible sheet of BOOK.xlsx (Pagination.paginatePrinted: what our PDF prints) against Excel 2019's
// (excel-print-pages.sh: EXCEL_DIR/<book>.tsv, the page breaks of Page Break Preview, and EXCEL_DIR/<book>-<n>.pdf,
// counted with pdfinfo). A line per sheet: pages, the rows and columns pages start at, and the scale.
// Build first: dotnet build tools/spreadsheet-pdf-runner/SpreadsheetPdfRunner.fsproj
#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
#r "../../src/Spreadsheet.Pdf/bin/Debug/net10.0/Rbt.Spreadsheet.Pdf.dll"

open System
open System.IO
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Pdf

let book = fsi.CommandLineArgs.[1]
let excelDir = fsi.CommandLineArgs.[2]
let stem = Path.GetFileNameWithoutExtension book
let tsv = Path.Combine(excelDir, stem + ".tsv")
let facts =
    if File.Exists tsv then
        File.ReadAllLines tsv |> Array.choose (fun line ->
            match line.Split '\t' with
            | [| sheet; n; key; value |] -> Some (sheet, int n, key, value)
            | _ -> None)
    else [||]
let pdfPages (path: string) =
    if not (File.Exists path) then -1 else
    let p = Diagnostics.Process.Start(Diagnostics.ProcessStartInfo("pdfinfo", "\"" + path + "\"", RedirectStandardOutput = true))
    let text = p.StandardOutput.ReadToEnd()
    p.WaitForExit()
    text.Split('\n') |> Array.tryPick (fun l -> if l.StartsWith "Pages:" then Some (int (l.Substring(6).Trim())) else None) |> Option.defaultValue -1

let colName (c: int) =
    let mutable c = c
    let mutable s = ""
    while c > 0 do
        let r = (c - 1) % 26
        s <- string (char (65 + r)) + s
        c <- (c - 1) / 26
    s

let imported = Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook (XlsxReader.loadBytesNative (File.ReadAllBytes book) |> Option.get)
let mutable sheets, same = 0, 0
imported.Sheets |> Array.iteri (fun index sheet ->
    if sheet.Visible then
        let n = index + 1
        let mine = facts |> Array.filter (fun (_, k, _, _) -> k = n)
        let get key = mine |> Array.tryPick (fun (_, _, k, v) -> if k = key then Some v else None)
        let breaks key = mine |> Array.choose (fun (_, _, k, v) -> if k = key then Some (int (v.Split(' ').[0])) else None) |> Array.distinct |> Array.sort |> List.ofArray
        let ourP = Pagination.paginatePrinted sheet.Model
        let ourRows = ourP.Pages |> List.map (fun p -> p.RowStart + 1) |> List.distinct |> List.sort |> (fun xs -> if xs.IsEmpty then xs else xs.Tail)
        let ourCols = ourP.Pages |> List.map (fun p -> p.ColStart + 1) |> List.distinct |> List.sort |> (fun xs -> if xs.IsEmpty then xs else xs.Tail)
        let area = ourP.PrintBounds
        let excelPages = match get "pdf" with Some name when name.EndsWith ".pdf" -> pdfPages (Path.Combine(excelDir, name)) | other -> -1
        let xr, xc = breaks "hbreak", breaks "vbreak"
        // (Excel's breaks of Page Break Preview are the sheet's; ours are the print area's.)
        let ok = excelPages = ourP.Pages.Length && xr = ourRows && xc = ourCols
        sheets <- sheets + 1
        if ok then same <- same + 1
        let show xs = xs |> List.map string |> String.concat ","
        printfn "%s\t%s\tpages %d/%d\trows %s | %s\tcols %s | %s\tscale %.3f zoom %s fit %s x %s\tarea %s%d:%s%d (excel %s used %s)"
            (if ok then "same" else "DIFF") sheet.Name ourP.Pages.Length excelPages
            (show ourRows) (show xr) (show ourCols) (show xc)
            ourP.Scale (defaultArg (get "setup.zoom") "-") (defaultArg (get "setup.fitwide") "-") (defaultArg (get "setup.fittall") "-")
            (colName (area.From.Col + 1)) (area.From.Row + 1) (colName (area.To.Col + 1)) (area.To.Row + 1)
            (defaultArg (get "setup.printarea") "-") (defaultArg (get "used") "-"))
printfn "%s: %d of %d sheets paged as Excel" stem same sheets
