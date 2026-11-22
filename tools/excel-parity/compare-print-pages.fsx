// dotnet fsi tools/excel-parity/compare-print-pages.fsx BOOK.xlsx EXCEL_DIR
// Compare direct PDF page counts with paginatePrinted and, separately, Page Break Preview breaks with
// paginatePreview. New oracle PDFs are direct exports; their preview breaks do not describe those PDFs.
// Older oracle PDFs were exported after preview and are compared only against preview pagination.
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
let mutable sheets, previewSame, directSheets, directSame = 0, 0, 0, 0
imported.Sheets |> Array.iteri (fun index sheet ->
    if sheet.Visible then
        let n = index + 1
        let mine = facts |> Array.filter (fun (_, k, _, _) -> k = n)
        let get key = mine |> Array.tryPick (fun (_, _, k, v) -> if k = key then Some v else None)
        let breaks key = mine |> Array.choose (fun (_, _, k, v) -> if k = key then Some (int (v.Split(' ').[0])) else None) |> Array.distinct |> Array.sort |> List.ofArray
        let direct = Pagination.paginatePrinted sheet.Model
        let preview = Pagination.paginatePreview sheet.Model
        let rows = preview.Pages |> List.map (fun p -> p.RowStart + 1) |> List.distinct |> List.sort |> (fun xs -> if xs.IsEmpty then xs else xs.Tail)
        let cols = preview.Pages |> List.map (fun p -> p.ColStart + 1) |> List.distinct |> List.sort |> (fun xs -> if xs.IsEmpty then xs else xs.Tail)
        let area = preview.PrintBounds
        let excelPages = match get "pdf" with Some name when name.EndsWith ".pdf" -> pdfPages (Path.Combine(excelDir, name)) | _ -> -1
        let fresh = (get "breaks.provenance").IsSome
        let rowKey, colKey = if fresh then "preview.hbreak", "preview.vbreak" else "hbreak", "vbreak"
        let xr, xc = breaks rowKey |> List.filter (fun r -> r <= area.To.Row + 1), breaks colKey |> List.filter (fun c -> c <= area.To.Col + 1)
        let previewOk = xr = rows && xc = cols && (fresh || excelPages < 0 || excelPages = preview.Pages.Length)
        sheets <- sheets + 1
        if previewOk then previewSame <- previewSame + 1
        if fresh && excelPages >= 0 then
            directSheets <- directSheets + 1
            if excelPages = direct.Pages.Length then directSame <- directSame + 1
        let show xs = xs |> List.map string |> String.concat ","
        let directPages = if fresh && excelPages >= 0 then sprintf "%d/%d" direct.Pages.Length excelPages else "unavailable"
        printfn "%s\t%s\tdirect pages %s\tpreview rows %s | %s\tpreview cols %s | %s\tpreview scale %.3f zoom %s fit %s x %s\tarea %s%d:%s%d (excel %s used %s)"
            (if previewOk then "preview-same" else "preview-DIFF") sheet.Name directPages
            (show rows) (show xr) (show cols) (show xc)
            preview.Scale (defaultArg (get "setup.zoom") "-") (defaultArg (get "setup.fitwide") "-") (defaultArg (get "setup.fittall") "-")
            (colName (area.From.Col + 1)) (area.From.Row + 1) (colName (area.To.Col + 1)) (area.To.Row + 1)
            (defaultArg (get "setup.printarea") "-") (defaultArg (get "used") "-"))
printfn "%s: preview breaks %d/%d; direct PDF page counts %d/%d" stem previewSame sheets directSame directSheets
