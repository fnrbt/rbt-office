#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

// The browser build reads the rows of a sheet from the text of its part
// (Xlsx.scanSheetRows) instead of building a DOM for them. It must give
// exactly the workbook the XML parsers give: every workbook of the corpus is
// loaded both ways and compared.

open System
open System.IO

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-34s  %s" label detail)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let files =
    [ "real-world"; "test-documents-excel" ]
    |> List.collect (fun dir -> Directory.GetFiles(Path.Combine(root, dir)) |> Array.toList)
    |> List.filter (fun file -> file.EndsWith ".xlsx" || file.EndsWith ".xlsm")
    |> List.sort

// NaN <> NaN: floats are compared through their text.
let describeCell (cell: XlsxReader.Cell) =
    sprintf "%s%d %A f=%A a=%A s=%A runs=%A" cell.Reference.Column cell.Reference.Row cell.Value cell.Formula cell.ArrayRef cell.StyleIndex cell.InlineRuns
let describeRow (row: XlsxReader.Row) =
    sprintf "%d %A %A %A %b" row.Index row.Height row.RawHeight row.CustomHeight row.Hidden

let mutable compared = 0
let mutable cells = 0
let mutable formulaCells = 0
let mutable deferredSheets = 0
let mutable quickSheets = 0
for file in files do
    let name = Path.GetFileName file
    let bytes = File.ReadAllBytes file
    let load (reader: byte[] -> XlsxReader.Workbook option) = try Choice1Of2 (reader bytes) with ex -> Choice2Of2 (ex.GetType().Name)
    match load XlsxReader.loadBytesNative, load XlsxReader.loadBytesNativeScanned with
    | Choice1Of2 (Some expected), Choice1Of2 (Some scanned) ->
        compared <- compared + 1
        // (A second reading, whose rows nothing has asked for.)
        let fresh = XlsxReader.loadBytesNativeScanned bytes
        ok (name + " sheets") (expected.Sheets.Length = scanned.Sheets.Length) "sheet count differs"
        for a, b in List.zip expected.Sheets scanned.Sheets do
            let label = sprintf "%s/%s" name a.Name
            ok (label + " row count") (a.Rows.Length = b.Rows.Length) (sprintf "%d vs %d" a.Rows.Length b.Rows.Length)
            if a.Rows.Length = b.Rows.Length then
                let mutable firstDifference = None
                for rowA, rowB in List.zip a.Rows b.Rows do
                    if firstDifference.IsNone then
                        if describeRow rowA <> describeRow rowB then firstDifference <- Some (describeRow rowA + "  vs  " + describeRow rowB)
                        elif rowA.Cells.Length <> rowB.Cells.Length then firstDifference <- Some (sprintf "row %d: %d cells vs %d" rowA.Index rowA.Cells.Length rowB.Cells.Length)
                        else
                            for cellA, cellB in List.zip rowA.Cells rowB.Cells do
                                cells <- cells + 1
                                if firstDifference.IsNone && (not (obj.ReferenceEquals(cellA.Style, cellB.Style)) && cellA.Style <> cellB.Style || describeCell cellA <> describeCell cellB) then
                                    firstDifference <- Some (describeCell cellA + "  vs  " + describeCell cellB)
                ok (label + " cells") firstDifference.IsNone (defaultArg firstDifference "")
            // The cells with a formula, asked for before the rows are read, are the rows' formula cells.
            let isFormula (cell: XlsxReader.Cell) = cell.Formula.IsSome || (match cell.Value with XlsxReader.Formula _ -> true | _ -> false)
            let fromRows = a.Rows |> List.collect (fun row -> row.Cells |> List.filter isFormula) |> List.map describeCell
            match fresh |> Option.bind (fun (w: XlsxReader.Workbook) -> w.Sheets |> List.tryFind (fun s -> s.Name = a.Name)) with
            | Some unread ->
                let wasLoaded = unread.RowSource.IsLoaded
                // What the rows say of themselves, likewise without their cells.
                let headers = unread.RowSource.Headers
                // (Rows that say nothing of themselves may be left out.)
                let telling (rows: XlsxReader.Row list) =
                    rows |> List.filter (fun row -> row.RawHeight.IsSome || row.Hidden || row.OutlineLevel > 0 || row.Collapsed)
                    |> List.map (fun row -> describeRow row, row.OutlineLevel, row.Collapsed)
                ok (label + " row headers") (telling headers = telling a.Rows && (wasLoaded || not unread.RowSource.IsLoaded)) "differ"
                ok (label + " row headers are rows") (headers |> List.forall (fun row -> a.Rows |> List.exists (fun other -> other.Index = row.Index)) || a.Rows.Length > 3000) "a row that is not there"
                let quick = unread.FormulaCells |> List.map describeCell
                formulaCells <- formulaCells + quick.Length
                if not wasLoaded then
                    deferredSheets <- deferredSheets + 1
                    if not unread.RowSource.IsLoaded then quickSheets <- quickSheets + 1
                ok (label + " formula cells") (quick = fromRows) (sprintf "%d vs %d; first difference %A" quick.Length fromRows.Length (List.zip (List.truncate (min quick.Length fromRows.Length) quick) (List.truncate (min quick.Length fromRows.Length) fromRows) |> List.tryFind (fun (x, y) -> x <> y)))
            | None -> ok (label + " read again") false "missing"
            // Everything that is not a row comes from the same parser, on the part without its rows.
            let rest (s: XlsxReader.Worksheet) =
                sprintf "%A" (s.MergedCells, s.Columns, s.Format, s.FrozenPane, s.ShowZeros, s.ShowGridLines, s.Hyperlinks, s.PageSetup, s.PageMargins, s.PrintOptions, s.HeaderFooter, s.RowBreaks, s.ColBreaks, s.Extras.DataTables, s.Extras.KeptExtensions)
            ok (label + " rest of the sheet") (rest a = rest b) "differs"
            // The fast native reader skips rules; compare them when it read them.
            if not a.ConditionalFormats.IsEmpty || not a.DataValidations.IsEmpty then
                ok (label + " rules") (sprintf "%A" (a.ConditionalFormats, a.DataValidations) = sprintf "%A" (b.ConditionalFormats, b.DataValidations)) "differ"
    | Choice1Of2 None, Choice1Of2 None -> ()
    | Choice2Of2 a, Choice2Of2 b -> ok (name + " both fail alike") (a = b) (a + " vs " + b)
    | a, b -> ok (name + " same outcome") false (sprintf "%A vs %A" (match a with Choice1Of2 w -> string w.IsSome | Choice2Of2 e -> e) (match b with Choice1Of2 w -> string w.IsSome | Choice2Of2 e -> e))

ok "the corpus was read" (compared > 100 && cells > 100000) (sprintf "%d workbooks, %d cells" compared cells)

// The rows of a sheet stay unread until they are asked for, and its formula cells are found without them.
ok "rows are left for later" (deferredSheets > 120 && quickSheets = deferredSheets && formulaCells > 10000)
    (sprintf "%d sheets deferred, %d gave their formula cells without their rows, %d formula cells" deferredSheets quickSheets formulaCells)

printfn ""
printfn "XLSX row scanner: %d workbooks, %d cells compared; %d passed, %d failed" compared cells pass fail
if fail > 0 then exit 1
