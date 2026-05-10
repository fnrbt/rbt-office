#!/usr/bin/env dotnet fsi

// Test script for workbook with multiple sheets and tables

#load "Spreadsheet/Model.fs"
#load "Spreadsheet/Formula.fs"
#load "Spreadsheet/Calc.fs"
#load "Spreadsheet/Workbook.fs"

open System
open RedBat.Grid.Spreadsheet.Model
open RedBat.Grid.Spreadsheet.Workbook

printfn "=== Workbook with Multiple Sheets and Tables ==="
printfn ""

// Create workbook
let mutable workbook = createWorkbook()

printfn "1. Create Workbook"
let stats = getWorkbookStats workbook
printfn "   Sheets: %d" stats.SheetCount
printfn "   Active: %s" workbook.ActiveSheet
printfn ""

// Add more sheets
printfn "2. Adding Sheets"
workbook <- addSheet "Budget" workbook
workbook <- addSheet "Expenses" workbook
workbook <- addSheet "Summary" workbook

let sheetNames = getSheetNames workbook
printfn "   Sheets:"
for (id, name) in sheetNames do
    let active = if id = workbook.ActiveSheet then " (active)" else ""
    printfn "     - %s [%s]%s" name id active
printfn ""

// Rename a sheet
printfn "3. Rename Sheet"
match workbook.SheetOrder.[1] with
| sheetId ->
    workbook <- renameSheet sheetId "Monthly Budget" workbook
    printfn "   Renamed sheet to 'Monthly Budget'"
printfn ""

// Switch active sheet
printfn "4. Switch Active Sheet"
match workbook.SheetOrder.[2] with
| sheetId ->
    workbook <- setActiveSheet sheetId workbook
    printfn "   Active sheet: %s" (Map.find workbook.ActiveSheet workbook.Sheets).Name
printfn ""

// Define tables
printfn "5. Define Tables"

let salesTable = {
    Name = "Sales"
    SheetId = workbook.SheetOrder.[0]
    StartRow = 0
    StartCol = 0
    EndRow = 10
    EndCol = 3
    Columns = ["Date"; "Product"; "Amount"; "Region"]
    HasHeaders = true
}
workbook <- defineTable salesTable workbook

let expensesTable = {
    Name = "Expenses"
    SheetId = workbook.SheetOrder.[2]
    StartRow = 0
    StartCol = 0
    EndRow = 20
    EndCol = 2
    Columns = ["Category"; "Amount"; "Date"]
    HasHeaders = true
}
workbook <- defineTable expensesTable workbook

let tableNames = getTableNames workbook
printfn "   Tables defined:"
for name in tableNames do
    match getTable name workbook with
    | Some table ->
        let sheetName = (Map.find table.SheetId workbook.Sheets).Name
        let range = sprintf "%s:%s"
                        (formatCellRef table.StartRow table.StartCol)
                        (formatCellRef table.EndRow table.EndCol)
        printfn "     - %s on sheet '%s' (%s)" name sheetName range
        printfn "       Columns: %s" (String.concat ", " table.Columns)
    | None -> ()
printfn ""

// Cross-sheet references
printfn "6. Cross-Sheet References"
printfn "   Examples:"
printfn "   - =Sheet1!A1 (reference to Sheet1 cell A1)"
printfn "   - ='Monthly Budget'!B5 (reference with sheet name containing spaces)"
printfn "   - =SUM(Sheet1!A1:A10)"
printfn ""

let testRefs = [
    "Sheet1!A1"
    "'Monthly Budget'!B5"
    "Expenses!C10"
    "A1"
]

for ref in testRefs do
    let (sheetName, cellRef) = parseSheetReference ref
    match sheetName with
    | Some name ->
        printfn "   %s -> Sheet: %s, Cell: %s" ref name cellRef
        match resolveSheetName name workbook with
        | Some id -> printfn "     (Sheet ID: %s)" id
        | None -> printfn "     (Sheet not found)"
    | None ->
        printfn "   %s -> Current sheet, Cell: %s" ref cellRef
printfn ""

// Table references
printfn "7. Table References (Structured)"
printfn "   Examples:"
printfn "   - =Sales[Amount] (reference to Amount column in Sales table)"
printfn "   - =SUM(Sales[Amount])"
printfn "   - =AVERAGE(Expenses[Amount])"
printfn ""

let testTableRefs = [
    "Sales[Amount]"
    "Sales[Product]"
    "Expenses[Category]"
    "Invalid[Column]"
]

for ref in testTableRefs do
    match parseTableReference ref with
    | Some (tableName, columnName) ->
        printfn "   %s -> Table: %s, Column: %s" ref tableName columnName
        match getTableColumn tableName columnName workbook with
        | Some (sheetId, startAddr, endAddr) ->
            let sheetName = (Map.find sheetId workbook.Sheets).Name
            printfn "     Range: %s!%s:%s"
                sheetName
                (formatCellRef startAddr.Row startAddr.Col)
                (formatCellRef endAddr.Row endAddr.Col)
        | None ->
            printfn "     (Table or column not found)"
    | None ->
        printfn "   %s -> Not a table reference" ref
printfn ""

// Delete sheet
printfn "8. Delete Sheet"
match workbook.SheetOrder |> List.tryItem 3 with
| Some sheetId ->
    let sheetName = (Map.find sheetId workbook.Sheets).Name
    match deleteSheet sheetId workbook with
    | Some wb ->
        workbook <- wb
        printfn "   Deleted sheet '%s'" sheetName
        printfn "   Remaining sheets: %d" (getWorkbookStats workbook).SheetCount
    | None ->
        printfn "   Cannot delete last sheet"
| None -> ()
printfn ""

// Reorder sheets
printfn "9. Reorder Sheets"
printfn "   Original order:"
for (id, name) in getSheetNames workbook do
    printfn "     %d. %s" ((List.findIndex ((=) id) workbook.SheetOrder) + 1) name

let newOrder = List.rev workbook.SheetOrder
workbook <- reorderSheets newOrder workbook

printfn "   New order (reversed):"
for (id, name) in getSheetNames workbook do
    printfn "     %d. %s" ((List.findIndex ((=) id) workbook.SheetOrder) + 1) name
printfn ""

// Final statistics
printfn "10. Final Workbook Statistics"
let finalStats = getWorkbookStats workbook
printfn "   Sheets: %d" finalStats.SheetCount
printfn "   Tables: %d" finalStats.TableCount
printfn "   Total cells: %d" finalStats.TotalCells
printfn "   Total formulas: %d" finalStats.TotalFormulas
printfn ""

printfn "=== Test Complete ==="
printfn ""
printfn "Key Features Demonstrated:"
printfn "1. Multiple sheets in a workbook"
printfn "2. Sheet operations (add, delete, rename, reorder)"
printfn "3. Active sheet tracking"
printfn "4. Structured tables with named columns"
printfn "5. Cross-sheet references (Sheet1!A1)"
printfn "6. Table references (Sales[Amount])"
printfn "7. Sheet name resolution"
printfn "8. Table column range lookup"
