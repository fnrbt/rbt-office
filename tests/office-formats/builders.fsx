#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)

let cell value = WorkbookBuilder.createCell value None None None None None

let sheet = WorkbookBuilder.SheetBuilder("Rows")
sheet.AddRow([ cell "r1" ]) |> ignore
sheet.AddRow([ cell "r2" ]) |> ignore
sheet.AddMergedRange(0, 0, 0, 1) |> ignore
sheet.AddMergedRange(2, 0, 2, 1) |> ignore
let builtSheet = sheet.Build()
ok "sheet-row-order"
   (builtSheet.Rows |> List.map (fun row -> row.Cells.Head.Value) = [ "r1"; "r2" ])
   (sprintf "%A" builtSheet.Rows)
ok "sheet-merge-order"
   (builtSheet.MergedCells |> List.map (fun range -> range.StartRow, range.EndRow) = [ (0, 0); (2, 2) ])
   (sprintf "%A" builtSheet.MergedCells)

printfn ""
printfn "OfficeFormat builder tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
