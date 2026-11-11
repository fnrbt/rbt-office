#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
module Wb = Rbt.Spreadsheet.Workbook
module Ops = Rbt.Spreadsheet.SheetOps
module Paint = Rbt.Spreadsheet.Paint
module Layout = Rbt.Spreadsheet.Layout

// Tables as typing changes them (SheetOps.tableEntry): a header cell names its
// column (structured references follow, blank and repeated names as Excel
// makes them), a table grows when something is typed directly below or right
// of it, a formula typed in a column fills the column; and what a built-in
// table style gives each cell (Model.tableCellLook), in the painter too.

let mutable pass = 0
let mutable fail = 0
let eq label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-44s expected %A  got %A" label expected got)

let at row col : CellAddress = { Row = row; Col = col }
let num (n: float) = { Value = CellValue.Number n; Format = defaultFormat; Comment = None }
let txt (s: string) = { Value = CellValue.String s; Format = defaultFormat; Comment = None }
let fml (text: string) = { Value = CellValue.Formula (text, CellValue.Empty); Format = defaultFormat; Comment = None }
let sheetOf cells = cells |> List.fold (fun m (addr, cell) -> setCell addr cell m) (createEmpty 40 12)
let range r0 c0 r1 c1 : CellRange = { From = at r0 c0; To = at r1 c1 }

// Sales: Region | Amount | Qty | Total (= Amount * Qty, a calculated column), three rows.
let data =
    sheetOf [ at 0 0, txt "Region"; at 0 1, txt "Amount"; at 0 2, txt "Qty"; at 0 3, txt "Total"
              at 1 0, txt "East"; at 1 1, num 10.0; at 1 2, num 1.0; at 1 3, fml "=[@Amount]*[@Qty]"
              at 2 0, txt "West"; at 2 1, num 20.0; at 2 2, num 2.0; at 2 3, fml "=[@Amount]*[@Qty]"
              at 3 0, txt "East"; at 3 1, num 30.0; at 3 2, num 3.0; at 3 3, fml "=[@Amount]*[@Qty]"
              at 1 6, fml "=SUM(Sales[Amount])"; at 2 6, fml "=Sales[@Amount]*2"; at 3 6, fml "=SUM(Sales[[Amount]:[Qty]])"
              at 4 6, fml "=SUM(Other[Amount])+LEN(\"Amount\")" ]
let other = sheetOf [ at 0 0, fml "=SUM(Sales[Qty])"; at 1 0, fml "=SUM(Sales[Amount])"; at 2 0, fml "=SUM(BigOnes)"
                      at 4 0, txt "Amount"; at 5 0, num 7.0 ]
let sheetEntry id name model : Wb.SheetId * Wb.Sheet = id, { Id = id; Name = name; Model = model; Visible = true }
let bare : Wb.Workbook =
    { Sheets = Map.ofList [ sheetEntry "sheet1" "Data" data; sheetEntry "sheet2" "Other" other ]
      SheetOrder = [ "sheet1"; "sheet2" ]; ActiveSheet = "sheet1"; Tables = Map.empty; Names = Map.empty; NameInfo = Map.empty; Iteration = None }
let style = Some (Ops.tableStyleNamed Ops.defaultTableStyleName)
let wb0 =
    let named = bare |> Wb.setDefinedName None "BigOnes" "Sales[Amount]" None
    let sales, _ = Ops.defineTable "T" "sheet1" (range 0 0 3 3) (Some "Sales") true style named
    let both, _ = Ops.defineTable "T" "sheet2" (range 4 0 5 0) (Some "Other") true style sales.Book
    recalculateWorkbook both.Book
let formulaAt sheet row col (wb: Wb.Workbook) =
    match getCellValue (at row col) wb.Sheets.[sheet].Model with
    | CellValue.Formula (text, _) -> formulaText text
    | value -> sprintf "%A" value
let valueAt sheet row col (wb: Wb.Workbook) = getCellValue (at row col) wb.Sheets.[sheet].Model |> unwrapFormulaValue
let infoOf name (wb: Wb.Workbook) = Ops.tablesOf wb |> List.find (fun table -> table.Name = name)
let entered row col (cell: CellData option) (wb: Wb.Workbook) = Ops.tableEntry "sheet1" (at row col) cell wb
let isNone row col cell wb = (entered row col cell wb).IsNone
let after row col cell wb = (entered row col cell wb).Value.Book |> recalculateWorkbook

// ---- a header cell names its column ------------------------------------------------------
let renamed = after 0 1 (Some (txt "Price")) wb0
eq "header: the column has the name" (infoOf "Sales" renamed).Columns [ "Region"; "Price"; "Qty"; "Total" ]
eq "header: references by the table's name follow" (formulaAt "sheet1" 1 6 renamed, formulaAt "sheet1" 2 6 renamed, formulaAt "sheet1" 3 6 renamed, formulaAt "sheet2" 1 0 renamed)
    ("=SUM(Sales[Price])", "=Sales[@Price]*2", "=SUM(Sales[[Price]:[Qty]])", "=SUM(Sales[Price])")
eq "header: references without it, in the table's own cells" (formulaAt "sheet1" 1 3 renamed, formulaAt "sheet1" 3 3 renamed) ("=[@Price]*[@Qty]", "=[@Price]*[@Qty]")
eq "header: another table's column of that name, and text, are left" (formulaAt "sheet1" 4 6 renamed) "=SUM(Other[Amount])+LEN(\"Amount\")"
eq "header: a defined name follows" (renamed.Names |> Map.toList |> List.map snd) [ "Sales[Price]" ]
eq "header: the values are what they were" ([ valueAt "sheet1" 1 6 renamed; valueAt "sheet1" 2 6 renamed; valueAt "sheet1" 3 3 renamed; valueAt "sheet2" 2 0 renamed ]) [ CellValue.Number 60.0; CellValue.Number 40.0; CellValue.Number 90.0; CellValue.Number 60.0 ]
eq "header: the sheet's own record of the table is as it was" (renamed.Sheets.["sheet1"].Model.Tables) (wb0.Sheets.["sheet1"].Model.Tables)
let repeated = after 0 2 (Some (txt "region")) wb0
eq "header: a name another column has gets a number" ((infoOf "Sales" repeated).Columns, valueAt "sheet1" 0 2 repeated, formulaAt "sheet2" 0 0 repeated, formulaAt "sheet1" 1 3 repeated)
    ([ "Region"; "Amount"; "region2"; "Total" ], CellValue.String "region2", "=SUM(Sales[region2])", "=[@Amount]*[@region2]")
let blanked = after 0 2 None wb0
eq "header: cleared, it is ColumnN again" ((infoOf "Sales" blanked).Columns, valueAt "sheet1" 0 2 blanked, formulaAt "sheet2" 0 0 blanked) ([ "Region"; "Amount"; "Column3"; "Total" ], CellValue.String "Column3", "=SUM(Sales[Column3])")
let numbered = after 0 2 (Some (num 2024.0)) wb0
eq "header: a number is its text" ((infoOf "Sales" numbered).Columns.[2], valueAt "sheet1" 0 2 numbered) ("2024", CellValue.String "2024")
let sameAgain = entered 0 1 (Some (txt "Amount")) wb0
eq "header: the same name again changes no formula" (sameAgain.Value.Cells, (infoOf "Sales" sameAgain.Value.Book).Columns) ([ "sheet1", at 0 1 ], [ "Region"; "Amount"; "Qty"; "Total" ])
let cased = after 0 1 (Some (txt "AMOUNT")) wb0
eq "header: another spelling of its own name" ((infoOf "Sales" cased).Columns.[1], formulaAt "sheet1" 1 6 cased) ("AMOUNT", "=SUM(Sales[AMOUNT])")

// ---- typing below the table -----------------------------------------------------------------
let grownChange = (entered 4 0 (Some (txt "North")) wb0).Value
let grown = recalculateWorkbook grownChange.Book
eq "below: the table has the row" (Ops.rangeText (infoOf "Sales" grown).Range, grown.Sheets.["sheet1"].Model.Tables |> List.map (fun t -> t.EndRow)) ("A1:D5", [ 5 ])
eq "below: the calculated column is in the new row" (formulaAt "sheet1" 4 3 grown, valueAt "sheet1" 4 3 grown) ("=[@Amount]*[@Qty]", CellValue.Number 0.0)
eq "below: the cells of the change" (grownChange.Cells |> List.sort) [ "sheet1", at 4 0; "sheet1", at 4 3 ]
eq "below: formulas over the column read the row" (valueAt "sheet1" 1 6 (after 4 1 (Some (num 40.0)) wb0)) (CellValue.Number 100.0)
let blocked = { wb0 with Sheets = wb0.Sheets |> Map.add "sheet1" { wb0.Sheets.["sheet1"] with Model = setCell (at 4 2) (txt "note") wb0.Sheets.["sheet1"].Model } }
eq "below: not over a cell that holds something" (isNone 4 0 (Some (txt "North")) blocked) true
eq "below: nothing typed, nothing grows" (isNone 4 0 None wb0, isNone 4 0 (Some (txt "")) wb0) (true, true)
eq "below: two rows down is not the table's" (isNone 5 0 (Some (txt "North")) wb0) true
eq "below: beside the row below is not the table's" (isNone 4 4 (Some (txt "x")) wb0, isNone 4 5 (Some (txt "x")) wb0) (true, true)
let totalled = (Ops.setTableTotalRow "T" "Sales" true wb0).Book
eq "below: a table with a total row does not grow" (isNone 5 0 (Some (txt "North")) totalled) true
let merged = { wb0 with Sheets = wb0.Sheets |> Map.add "sheet1" { wb0.Sheets.["sheet1"] with Model = addMergedCell 4 1 5 2 wb0.Sheets.["sheet1"].Model } }
eq "below: not into a merged cell" (isNone 4 0 (Some (txt "North")) merged) true

// ---- typing right of the table -----------------------------------------------------------------
let wider = after 2 4 (Some (num 5.0)) wb0
eq "right: the table has the column, named in the header row" (Ops.rangeText (infoOf "Sales" wider).Range, (infoOf "Sales" wider).Columns, valueAt "sheet1" 0 4 wider) ("A1:E4", [ "Region"; "Amount"; "Qty"; "Total"; "Column5" ], CellValue.String "Column5")
let named = after 0 4 (Some (txt "Tax")) wb0
eq "right: typed in the header row, the text is the name" ((infoOf "Sales" named).Columns, Ops.rangeText (infoOf "Sales" named).Range) ([ "Region"; "Amount"; "Qty"; "Total"; "Tax" ], "A1:E4")
let taxChange = (entered 1 4 (Some (fml "=[@Amount]*0.1")) wb0).Value
let taxed = recalculateWorkbook taxChange.Book
eq "right: a formula there is a calculated column at once" ([ for row in 1 .. 3 -> formulaAt "sheet1" row 4 taxed, valueAt "sheet1" row 4 taxed ]) [ "=[@Amount]*0.1", CellValue.Number 1.0; "=[@Amount]*0.1", CellValue.Number 2.0; "=[@Amount]*0.1", CellValue.Number 3.0 ]
eq "right: the cells of the change" (taxChange.Cells |> List.sort) [ for row in 0 .. 3 -> "sheet1", at row 4 ]
let blockedRight = { wb0 with Sheets = wb0.Sheets |> Map.add "sheet1" { wb0.Sheets.["sheet1"] with Model = setCell (at 3 4) (txt "note") wb0.Sheets.["sheet1"].Model } }
eq "right: not over a cell that holds something" (isNone 1 4 (Some (num 1.0)) blockedRight) true

// ---- a formula in a column ---------------------------------------------------------------------
let emptied = { wb0 with Sheets = wb0.Sheets |> Map.add "sheet1" { wb0.Sheets.["sheet1"] with Model = [ 1; 2; 3 ] |> List.fold (fun m row -> clearCell (at row 3) m) wb0.Sheets.["sheet1"].Model } }
let filled = after 2 3 (Some (fml "=B3*C3")) emptied
eq "calculated column: an empty column takes the formula, row by row" ([ for row in 1 .. 3 -> formulaAt "sheet1" row 3 filled ]) [ "=B2*C2"; "=B3*C3"; "=B4*C4" ]
let replaced = after 1 3 (Some (fml "=[@Amount]+[@Qty]")) wb0
eq "calculated column: a new formula replaces the column's" ([ for row in 1 .. 3 -> formulaAt "sheet1" row 3 replaced ]) [ "=[@Amount]+[@Qty]"; "=[@Amount]+[@Qty]"; "=[@Amount]+[@Qty]" ]
let single = after 1 1 (Some (fml "=5+5")) wb0
eq "calculated column: a column of values keeps them" ([ for row in 1 .. 3 -> formulaAt "sheet1" row 1 single ]) [ "=5+5"; "Number 20.0"; "Number 30.0" ]
eq "a value typed in the table is an entry of the table (nothing else changes)" ((entered 2 1 (Some (num 99.0)) wb0).Value.Cells, (after 2 1 (Some (num 99.0)) wb0).Tables = wb0.Tables) ([ "sheet1", at 2 1 ], true)
eq "a cell away from every table is no table entry" (isNone 10 10 (Some (num 1.0)) wb0, Ops.entryConcernsTable (at 10 10) wb0.Sheets.["sheet1"].Model, Ops.entryConcernsTable (at 4 4) wb0.Sheets.["sheet1"].Model) (true, false, true)

// ---- what a style gives a cell --------------------------------------------------------------------
let tableWith (styleName: string) (change: TableStyle -> TableStyle) (totals: bool) : Table =
    { Name = "T"; StartRow = 1; StartCol = 1; EndRow = (if totals then 6 else 5); EndCol = 4; HeaderRowCount = 1
      Style = Some (change (Ops.tableStyleNamed styleName)); AutoFilter = None; HasTotals = totals }
let look styleName change totals row col = (tableCellLook (tableWith styleName change totals) row col).Value
let fillOf (l: TableCellLook) = l.Fill
let edge (line: Rbt.Spreadsheet.TableStyles.TableLine option) = line |> Option.map (fun line -> line.Color, line.Style)
eq "Medium2: header, bands, plain rows" ([ 0; 1; 2; 3 ] |> List.map (fun row -> let l = look "TableStyleMedium2" id false row 1 in l.Fill, l.Bold))
    [ Some "#4472C4", true; Some "#D9E1F2", false; None, false; Some "#D9E1F2", false ]
eq "Medium2: the header's text is white" (look "TableStyleMedium2" id false 0 1).Text (Some "#FFFFFF")
eq "Medium2: the total row is bold under a double line" (let l = look "TableStyleMedium2" id true 5 1 in l.Fill, l.Bold, edge l.Top) (None, true, Some ("#4472C4", "double"))
eq "outside the table, and a table without a style" (tableCellLook (tableWith "TableStyleMedium2" id false) 5 1, tableCellLook { tableWith "TableStyleMedium2" id false with Style = None } 1 1) (None, None)
eq "no banded rows" (look "TableStyleMedium2" (fun s -> { s with ShowRowStripes = false }) false 1 1 |> fillOf) None
eq "banded columns" ([ 0; 1; 2 ] |> List.map (fun col -> look "TableStyleMedium2" (fun s -> { s with ShowRowStripes = false; ShowColumnStripes = true }) false 2 col |> fillOf)) [ Some "#D9E1F2"; None; Some "#D9E1F2" ]
eq "first and last column" ([ 0; 1; 3 ] |> List.map (fun col -> (look "TableStyleMedium2" (fun s -> { s with ShowFirstColumn = true; ShowLastColumn = true }) false 2 col).Bold)) [ true; false; true ]
eq "Light1: no header fill, lines above and below it, grey bands" (let h = look "TableStyleLight1" id false 0 0 in h.Fill, h.Bold, edge h.Top, edge h.Bottom, (look "TableStyleLight1" id false 1 0).Fill, edge (look "TableStyleLight1" id false 4 0).Bottom)
    (None, true, Some ("#000000", "thin"), Some ("#000000", "thin"), Some "#D9D9D9", Some ("#000000", "thin"))
eq "Light9: a filled header, a line above each band" (let h = look "TableStyleLight9" id false 0 0 in h.Fill, h.Text, (look "TableStyleLight9" id false 1 0).Fill, edge (look "TableStyleLight9" id false 2 1).Top) (Some "#4472C4", Some "#FFFFFF", None, Some ("#4472C4", "thin"))
eq "Light16: a grid in the colour, a medium line under the header" (edge (look "TableStyleLight16" id false 2 1).Left, edge (look "TableStyleLight16" id false 2 1).Top, edge (look "TableStyleLight16" id false 0 1).Bottom) (Some ("#4472C4", "thin"), Some ("#4472C4", "thin"), Some ("#4472C4", "medium"))
eq "Medium9: every cell filled, the total row and the first column in the colour" ((look "TableStyleMedium9" id true 1 1).Fill, (look "TableStyleMedium9" id true 2 1).Fill, (look "TableStyleMedium9" id true 5 1).Fill, (look "TableStyleMedium9" (fun s -> { s with ShowFirstColumn = true }) true 2 0).Fill)
    (Some "#B4C6E7", Some "#D9E1F2", Some "#4472C4", Some "#4472C4")
eq "Medium16: grey bands" ((look "TableStyleMedium16" id false 1 1).Fill, (look "TableStyleMedium16" id false 2 1).Fill) (Some "#D9D9D9", None)
eq "Dark2: dark fills and white text" (let d = look "TableStyleDark2" id false 1 1 in (look "TableStyleDark2" id false 0 1).Fill, d.Fill, d.Text, (look "TableStyleDark2" id false 2 1).Fill) (Some "#000000", Some "#305496", Some "#FFFFFF", Some "#4472C4")
eq "the style's colour is the theme's accent" ((look "TableStyleMedium3" id false 0 0).Fill) (Some "#ED7D31")
eq "a cell's own colours win" (let f = formatWithTableLook (look "TableStyleMedium2" id false 0 0) { defaultFormat with BgColor = Some "#FFFF00" } in f.BgColor, f.TextColor, f.Bold) (Some "#FFFF00", Some "#FFFFFF", true)
eq "... but black text is the default font's: the style's colour replaces it" ((formatWithTableLook (look "TableStyleMedium2" id false 0 0) { defaultFormat with TextColor = Some "#000000" }).TextColor, (formatWithTableLook (look "TableStyleMedium2" id false 0 0) { defaultFormat with TextColor = Some "#FF0000" }).TextColor) (Some "#FFFFFF", Some "#FF0000")

// ---- the painter ------------------------------------------------------------------------------------
let paintedModel (styleName: string) =
    let m = wb0.Sheets.["sheet1"].Model
    { m with Tables = m.Tables |> List.map (fun t -> { t with Style = Some (Ops.tableStyleNamed styleName) }) }
let itemsOf styleName = Paint.buildPaintItems (paintedModel styleName) (Layout.defaultViewport 800.0 500.0)
let fillCount color items = items |> List.sumBy (function Paint.FillRect (_, c) when c = color -> 1 | _ -> 0)
let boldTexts items = items |> List.choose (function Paint.DrawText t when t.Bold -> Some t.Text | _ -> None) |> List.sort
eq "painter, Medium2: header fills, band fills, bold header" (fillCount "#4472C4" (itemsOf "TableStyleMedium2"), fillCount "#D9E1F2" (itemsOf "TableStyleMedium2"), boldTexts (itemsOf "TableStyleMedium2")) (4, 8, [ "Amount"; "Qty"; "Region"; "Total" ])
eq "painter, Light1: no header fill, grey bands" (fillCount "#000000" (itemsOf "TableStyleLight1"), fillCount "#D9D9D9" (itemsOf "TableStyleLight1")) (0, 8)
eq "painter, Dark2: every cell of the table filled" (fillCount "#000000" (itemsOf "TableStyleDark2") + fillCount "#305496" (itemsOf "TableStyleDark2") + fillCount "#4472C4" (itemsOf "TableStyleDark2")) 16

printfn ""
printfn "Stage 48 table entry tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
