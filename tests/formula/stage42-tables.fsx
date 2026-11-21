#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
module Wb = Rbt.Spreadsheet.Workbook
module Ops = Rbt.Spreadsheet.SheetOps

// Tables as the host API changes them: making one of a range (names of the
// columns, both records of the table), resizing, renaming (structured
// references follow), the total row, converting to a range (structured
// references become plain ones) and deleting with the data.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-44s expected %A  got %A" label expected got)
let throws label (run: unit -> unit) (part: string) =
    try
        run ()
        fail <- fail + 1
        printfn "FAIL %-44s did not throw" label
    with error ->
        if error.Message.Contains part then pass <- pass + 1
        else (fail <- fail + 1; printfn "FAIL %-44s threw %s" label error.Message)

let at row col : CellAddress = { Row = row; Col = col }
let num (n: float) = { Value = CellValue.Number n; Format = defaultFormat; Comment = None }
let txt (s: string) = { Value = CellValue.String s; Format = defaultFormat; Comment = None }
let fml (text: string) = { Value = CellValue.Formula (text, CellValue.Empty); Format = defaultFormat; Comment = None }
let sheetOf cells = cells |> List.fold (fun m (addr, cell) -> setCell addr cell m) (createEmpty 40 12)
let range r0 c0 r1 c1 : CellRange = { From = at r0 c0; To = at r1 c1 }

// Data: Region | Amount | (blank header) | Amount, three rows; formulas beside it and on another sheet.
let data =
    sheetOf [ at 0 0, txt "Region"; at 0 1, txt "Amount"; at 0 3, txt "Amount"
              at 1 0, txt "East"; at 1 1, num 10.0; at 1 2, num 1.0; at 1 3, num 100.0
              at 2 0, txt "West"; at 2 1, num 20.0; at 2 2, num 2.0; at 2 3, num 200.0
              at 3 0, txt "East"; at 3 1, num 30.0; at 3 2, num 3.0; at 3 3, num 300.0
              at 1 6, fml "=SUM(Sales[Amount])"; at 2 6, fml "=Sales[@Amount]*2"; at 3 6, fml "=COUNTA(Sales[#Headers])"
              at 4 6, fml "=SUM(Sales[[Amount]:[Amount2]])"; at 5 6, fml "=ROWS(Sales[#All])" ]
let other = sheetOf [ at 0 0, fml "=SUM(Sales[Amount2])"; at 1 0, fml "=Sales[[#Totals],[Amount2]]"; at 2 0, fml "=SUM(Big)" ]
let sheetEntry id name model : Wb.SheetId * Wb.Sheet = id, { Id = id; Name = name; Model = model; Visible = true }
let wbBare : Wb.Workbook =
    { Sheets = Map.ofList [ sheetEntry "sheet1" "Data" data; sheetEntry "sheet2" "Other" other ]
      SheetOrder = [ "sheet1"; "sheet2" ]; ActiveSheet = "sheet1"; Tables = Map.empty; Names = Map.empty; NameInfo = Map.empty; Iteration = None; External = Map.empty }
let wb0 = wbBare |> Wb.setDefinedName None "Big" "Sales[Amount2]" None
let formulaAt sheet row col (wb: Wb.Workbook) =
    match getCellValue (at row col) wb.Sheets.[sheet].Model with
    | CellValue.Formula (text, _) -> formulaText text
    | value -> sprintf "%A" value
let valueAt sheet row col (wb: Wb.Workbook) = getCellValue (at row col) wb.Sheets.[sheet].Model |> unwrapFormulaValue
let calc = recalculateWorkbook
let infoOf name (wb: Wb.Workbook) = Ops.tablesOf wb |> List.find (fun table -> table.Name = name)
let style = Some (Ops.tableStyleNamed Ops.defaultTableStyleName)

// ---- names ---------------------------------------------------------------------------
eq "column names from header cells" (Ops.uniqueColumnNames [ "Region"; "Amount"; ""; "Amount"; "amount"; " " ]) [ "Region"; "Amount"; "Column3"; "Amount2"; "amount3"; "Column6" ]
eq "the first free TableN" (Ops.freeTableName wb0) "Table1"
eq "a built-in style's fills" (Ops.tableStyleFills "TableStyleMedium2", Ops.tableStyleFills "TableStyleLight1" |> fst) (("#4472C4", "#D9E1F2"), "#000000")

// ---- define ----------------------------------------------------------------------------
let defined, name = Ops.defineTable "DefineTable" "sheet1" (range 0 0 3 3) (Some "Sales") true style wb0
let wb1 = calc defined.Book
eq "the name is returned" name "Sales"
eq "the table as hosts see it" (infoOf "Sales" wb1 |> fun t -> Ops.rangeText t.Range, t.HasHeaders, t.HasTotals, t.Columns, t.HasAutoFilter, t.Style |> Option.map (fun s -> s.Name))
    ("A1:D4", true, false, [ "Region"; "Amount"; "Column3"; "Amount2" ], true, Some "TableStyleMedium2")
eq "blank and repeated header cells are written" (valueAt "sheet1" 0 2 wb1, valueAt "sheet1" 0 3 wb1, defined.Cells) (CellValue.String "Column3", CellValue.String "Amount2", [ "sheet1", at 0 2; "sheet1", at 0 3 ])
eq "the sheet's record (1-based)" (wb1.Sheets.["sheet1"].Model.Tables |> List.map (fun t -> t.Name, t.StartRow, t.StartCol, t.EndRow, t.EndCol, t.HeaderRowCount, t.AutoFilter)) [ "Sales", 1, 1, 4, 4, 1, Some (1, 4) ]
eq "structured references resolve" (valueAt "sheet1" 1 6 wb1, valueAt "sheet1" 2 6 wb1, valueAt "sheet1" 3 6 wb1, valueAt "sheet1" 4 6 wb1, valueAt "sheet1" 5 6 wb1, valueAt "sheet2" 0 0 wb1, valueAt "sheet2" 2 0 wb1)
    (CellValue.Number 60.0, CellValue.Number 40.0, CellValue.Number 4.0, CellValue.Number 666.0, CellValue.Number 4.0, CellValue.Number 600.0, CellValue.Number 600.0)
let unnamed, defaultName = Ops.defineTable "DefineTable" "sheet2" (range 5 0 7 1) None false None wb1
eq "a default name, no headers" (defaultName, infoOf "Table1" unnamed.Book |> fun t -> t.HasHeaders, t.Columns, t.HasAutoFilter, t.Style, unnamed.Cells) ("Table1", (false, [ "Column1"; "Column2" ], false, None, []))
throws "a name that is taken" (fun () -> Ops.defineTable "DefineTable" "sheet2" (range 5 0 7 1) (Some "sales") true style wb1 |> ignore) "already a table named"
throws "a name that is a defined name" (fun () -> Ops.defineTable "DefineTable" "sheet2" (range 5 0 7 1) (Some "Big") true style wb1 |> ignore) "is a defined name"
throws "a name that is a reference" (fun () -> Ops.defineTable "DefineTable" "sheet2" (range 5 0 7 1) (Some "A1") true style wb1 |> ignore) "reads as a cell reference"
throws "an overlapping range" (fun () -> Ops.defineTable "DefineTable" "sheet1" (range 3 3 6 5) None true style wb1 |> ignore) "overlaps the table \"Sales\""
throws "outside the sheet" (fun () -> Ops.defineTable "DefineTable" "sheet2" (range 0 0 50 1) None true style wb1 |> ignore) "outside the sheet"
throws "a header row alone" (fun () -> Ops.defineTable "DefineTable" "sheet2" (range 9 0 9 3) None true style wb1 |> ignore) "at least one more row"
throws "merged cells" (fun () -> Ops.defineTable "DefineTable" "sheet2" (range 9 0 12 3) None true style { wb1 with Sheets = Map.add "sheet2" { wb1.Sheets.["sheet2"] with Model = addMergedCell 10 0 10 1 other } wb1.Sheets } |> ignore) "merged cells"

// ---- total row --------------------------------------------------------------------------
let totalled = Ops.setTableTotalRow "SetTableTotalRow" "Sales" true wb1
let wb2 = calc totalled.Book
eq "the total row is the row below" (infoOf "Sales" wb2 |> fun t -> Ops.rangeText t.Range, t.HasTotals) ("A1:D5", true)
eq "a label and the sum of the last column" (valueAt "sheet1" 4 0 wb2, formulaAt "sheet1" 4 3 wb2, valueAt "sheet1" 4 3 wb2) (CellValue.String "Total", "=SUBTOTAL(109,Sales[Amount2])", CellValue.Number 600.0)
eq "the cells it wrote" totalled.Cells [ "sheet1", at 4 0; "sheet1", at 4 3 ]
eq "#Totals and #All see it; the data does not" (valueAt "sheet2" 1 0 wb2, valueAt "sheet1" 5 6 wb2, valueAt "sheet1" 1 6 wb2) (CellValue.Number 600.0, CellValue.Number 5.0, CellValue.Number 60.0)
eq "the sheet's record follows" (wb2.Sheets.["sheet1"].Model.Tables |> List.map (fun t -> t.EndRow, t.HasTotals)) [ 5, true ]
eq "shown twice is no change" (obj.ReferenceEquals((Ops.setTableTotalRow "SetTableTotalRow" "Sales" true wb2).Book, wb2)) true
let untotalled = Ops.setTableTotalRow "SetTableTotalRow" "Sales" false wb2
eq "removed: the row is emptied and the table shorter" (infoOf "Sales" untotalled.Book |> fun t -> Ops.rangeText t.Range, t.HasTotals, valueAt "sheet1" 4 0 untotalled.Book, valueAt "sheet1" 4 3 untotalled.Book) ("A1:D4", false, CellValue.Empty, CellValue.Empty)
throws "the row below is in use" (fun () -> Ops.setTableTotalRow "SetTableTotalRow" "Sales" true { wb1 with Sheets = Map.add "sheet1" { wb1.Sheets.["sheet1"] with Model = setCell (at 4 1) (num 1.0) wb1.Sheets.["sheet1"].Model } wb1.Sheets } |> ignore) "is not empty"
let textTable, _ = Ops.defineTable "DefineTable" "sheet2" (range 5 0 7 1) (Some "Words") true style { wb1 with Sheets = Map.add "sheet2" { wb1.Sheets.["sheet2"] with Model = other |> setCell (at 5 0) (txt "k") |> setCell (at 5 1) (txt "v") |> setCell (at 6 1) (txt "a") |> setCell (at 7 1) (txt "b") } wb1.Sheets }
let countTotals = calc (Ops.setTableTotalRow "SetTableTotalRow" "Words" true textTable.Book).Book
eq "a column of text is counted" (formulaAt "sheet2" 8 1 countTotals, valueAt "sheet2" 8 1 countTotals) ("=SUBTOTAL(103,Words[v])", CellValue.Number 2.0)
throws "an unknown table" (fun () -> Ops.setTableTotalRow "SetTableTotalRow" "Nope" true wb1 |> ignore) "there is no table \"Nope\""

// ---- resize -------------------------------------------------------------------------------
let grownSheet = wb1.Sheets.["sheet1"].Model |> setCell (at 4 0) (txt "North") |> setCell (at 4 1) (num 40.0) |> setCell (at 4 3) (num 400.0) |> setCell (at 0 4) (txt "Region")
let wbGrow = { wb1 with Sheets = Map.add "sheet1" { wb1.Sheets.["sheet1"] with Model = grownSheet } wb1.Sheets }
let resized = Ops.resizeTable "ResizeTable" "Sales" (range 0 0 4 4) wbGrow
let wb3 = calc resized.Book
eq "resized: more rows and a column" (infoOf "Sales" wb3 |> fun t -> Ops.rangeText t.Range, t.Columns) ("A1:E5", [ "Region"; "Amount"; "Column3"; "Amount2"; "Region2" ])
eq "the new header cell gets its unique name" (valueAt "sheet1" 0 4 wb3, resized.Cells) (CellValue.String "Region2", [ "sheet1", at 0 4 ])
eq "references read the new rows" (valueAt "sheet1" 1 6 wb3, valueAt "sheet2" 0 0 wb3, valueAt "sheet1" 5 6 wb3) (CellValue.Number 100.0, CellValue.Number 1000.0, CellValue.Number 5.0)
eq "the sheet's record and its filter buttons follow" (wb3.Sheets.["sheet1"].Model.Tables |> List.map (fun t -> t.EndRow, t.EndCol, t.AutoFilter)) [ 5, 5, Some (1, 5) ]
let shrunk = calc (Ops.resizeTable "ResizeTable" "Sales" (range 0 1 2 3) wb1).Book
eq "shrunk from the left and the bottom" (infoOf "Sales" shrunk |> fun t -> Ops.rangeText t.Range, t.Columns, valueAt "sheet1" 1 6 shrunk) ("B1:D3", [ "Amount"; "Column3"; "Amount2" ], CellValue.Number 30.0)
let resizedTotals = calc (Ops.resizeTable "ResizeTable" "Sales" (range 0 0 5 3) { wb2 with Sheets = Map.add "sheet1" { wb2.Sheets.["sheet1"] with Model = wb2.Sheets.["sheet1"].Model |> setCell (at 4 3) (num 400.0) |> setCell (at 4 0) (txt "North") } wb2.Sheets }).Book
eq "with a total row: it moves to the new last row" (infoOf "Sales" resizedTotals |> fun t -> Ops.rangeText t.Range, t.HasTotals, valueAt "sheet1" 5 0 resizedTotals, valueAt "sheet1" 5 3 resizedTotals, valueAt "sheet1" 4 3 resizedTotals)
    ("A1:D6", true, CellValue.String "Total", CellValue.Number 600.0, CellValue.Empty)
throws "the first row has to stay" (fun () -> Ops.resizeTable "ResizeTable" "Sales" (range 1 0 4 3) wb1 |> ignore) "has to keep its first row"
throws "it has to overlap" (fun () -> Ops.resizeTable "ResizeTable" "Sales" (range 0 5 3 7) wb1 |> ignore) "has to overlap"
throws "at least one row of data" (fun () -> Ops.resizeTable "ResizeTable" "Sales" (range 0 0 0 3) wb1 |> ignore) "at least one row of data"
eq "the same range is no change" (obj.ReferenceEquals((Ops.resizeTable "ResizeTable" "Sales" (range 0 0 3 3) wb1).Book, wb1)) true

// ---- rename -------------------------------------------------------------------------------
let renamed = Ops.renameTable "RenameTable" "Sales" "Revenue" wb2
let wb4 = calc renamed.Book
eq "renamed in both records" (Ops.tablesOf wb4 |> List.map (fun t -> t.Name), wb4.Sheets.["sheet1"].Model.Tables |> List.map (fun t -> t.Name), Map.containsKey "Sales" wb4.Tables) ([ "Revenue" ], [ "Revenue" ], false)
eq "formulas follow" (formulaAt "sheet1" 1 6 wb4, formulaAt "sheet1" 2 6 wb4, formulaAt "sheet2" 1 0 wb4, formulaAt "sheet1" 4 3 wb4) ("=SUM(Revenue[Amount])", "=Revenue[@Amount]*2", "=Revenue[[#Totals],[Amount2]]", "=SUBTOTAL(109,Revenue[Amount2])")
eq "defined names follow" (Wb.getName None "Big" wb4) (Some "Revenue[Amount2]")
eq "and still resolve" (valueAt "sheet1" 1 6 wb4, valueAt "sheet2" 2 0 wb4, valueAt "sheet2" 1 0 wb4) (CellValue.Number 60.0, CellValue.Number 600.0, CellValue.Number 600.0)
eq "the cells that changed" (renamed.Cells |> List.length) 8
eq "a change of case is a rename" ((Ops.renameTable "RenameTable" "Sales" "SALES" wb2).Book |> Ops.tablesOf |> List.map (fun t -> t.Name)) [ "SALES" ]
throws "a name that is taken" (fun () -> Ops.renameTable "RenameTable" "Sales" "Table1" unnamed.Book |> ignore) "already a table named"
throws "a name that is not one" (fun () -> Ops.renameTable "RenameTable" "Sales" "my table" wb2 |> ignore) "may hold letters"

// ---- convert to a range, delete -----------------------------------------------------------------
let converted = Ops.deleteTable "DeleteTable" "Sales" true wb2
let wb5 = calc converted.Book
eq "converted: the table is gone from both records" (Ops.tablesOf wb5, wb5.Sheets.["sheet1"].Model.Tables) ([], [])
eq "structured references become plain ones" (formulaAt "sheet1" 1 6 wb5, formulaAt "sheet1" 2 6 wb5, formulaAt "sheet1" 3 6 wb5, formulaAt "sheet1" 4 6 wb5, formulaAt "sheet1" 5 6 wb5, formulaAt "sheet1" 4 3 wb5)
    ("=SUM($B$2:$B$4)", "=$B$3*2", "=COUNTA($A$1:$D$1)", "=SUM($B$2:$D$4)", "=ROWS($A$1:$D$5)", "=SUBTOTAL(109,$D$2:$D$4)")
eq "on another sheet they name the sheet" (formulaAt "sheet2" 0 0 wb5, formulaAt "sheet2" 1 0 wb5, Wb.getName None "Big" wb5) ("=SUM(Data!$D$2:$D$4)", "=Data!$D$5", Some "Data!$D$2:$D$4")
eq "the values stand" (valueAt "sheet1" 1 6 wb5, valueAt "sheet1" 2 6 wb5, valueAt "sheet2" 0 0 wb5, valueAt "sheet2" 2 0 wb5, valueAt "sheet1" 0 0 wb5, valueAt "sheet1" 4 0 wb5)
    (CellValue.Number 60.0, CellValue.Number 40.0, CellValue.Number 600.0, CellValue.Number 600.0, CellValue.String "Region", CellValue.String "Total")
let deleted = Ops.deleteTable "DeleteTable" "Sales" false wb2
let wb6 = calc deleted.Book
eq "deleted with its data: the cells are empty" ([ for row in 0 .. 4 do for col in 0 .. 3 -> valueAt "sheet1" row col wb6 ] |> List.forall (fun value -> value = CellValue.Empty), Ops.tablesOf wb6) (true, [])
eq "references to it are errors" (valueAt "sheet1" 1 6 wb6, formulaAt "sheet1" 1 6 wb6) (CellValue.Error "REF!", "=SUM(Sales[Amount])")
throws "an unknown table" (fun () -> Ops.deleteTable "DeleteTable" "Nope" true wb2 |> ignore) "there is no table \"Nope\""

// ---- style ----------------------------------------------------------------------------------
let restyled = (Ops.setTableStyle "SetTableStyle" "Sales" (Some { Ops.tableStyleNamed "TableStyleLight10" with ShowRowStripes = false }) wb1).Book
eq "another style" (infoOf "Sales" restyled |> fun t -> t.Style |> Option.map (fun s -> s.Name, s.ShowRowStripes, s.Elements |> List.tryFind (fun e -> e.Kind = "headerRow") |> Option.bind (fun e -> e.Format.Fill))) (Some ("TableStyleLight10", false, Some "#ED7D31"))
eq "no style" ((Ops.setTableStyle "SetTableStyle" "Sales" None wb1).Book |> infoOf "Sales" |> fun t -> t.Style) None
eq "the same style is no change" (obj.ReferenceEquals((Ops.setTableStyle "SetTableStyle" "Sales" style wb1).Book, wb1)) true

// ---- rows and columns that move --------------------------------------------------------------
let inserted, _ = Wb.applyStructureEdit "sheet1" (InsertRows (2, 1)) wb2
let wb7 = calc inserted
eq "a row inserted inside: the table grows" (infoOf "Sales" wb7 |> fun t -> Ops.rangeText t.Range, t.HasTotals, valueAt "sheet1" 1 6 wb7) ("A1:D6", true, CellValue.Number 60.0)
let column, _ = Wb.applyStructureEdit "sheet1" (InsertCols (1, 1)) wb2
let wb8 = calc column
eq "a column inserted inside: it is a column of the table" (infoOf "Sales" wb8 |> fun t -> Ops.rangeText t.Range, t.Columns.Length, valueAt "sheet1" 1 7 wb8) ("A1:E5", 5, CellValue.Number 60.0)

printfn "Table tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
