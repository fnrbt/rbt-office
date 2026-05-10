# Multiple Sheets and Structured Tables

This document explains how the spreadsheet supports multiple sheets (tabs) and structured tables (named ranges with column references), similar to Excel and Google Sheets.

## Multiple Sheets (Tabs)

### Workbook Structure

A workbook contains multiple sheets:

```fsharp
type Workbook = {
    Sheets: Map<SheetId, Sheet>
    SheetOrder: SheetId list  // Tab order
    ActiveSheet: SheetId
    Tables: Map<string, Table>
}
```

Each sheet has:
- Unique ID (string)
- Name (displayed in tab)
- Model (SpreadsheetModel - cells, selection, etc.)
- CalcEngine (dependency graph, calc chain)
- Visibility flag

### Sheet Operations

**Create Workbook:**
```fsharp
let workbook = createWorkbook()  // Creates with Sheet1
```

**Add Sheet:**
```fsharp
let workbook' = addSheet "Budget" workbook
```

**Delete Sheet:**
```fsharp
match deleteSheet "sheet2" workbook with
| Some wb -> wb  // Success
| None -> workbook  // Can't delete last sheet
```

**Rename Sheet:**
```fsharp
let workbook' = renameSheet "sheet1" "Revenue" workbook
```

**Switch Active Sheet:**
```fsharp
let workbook' = setActiveSheet "sheet2" workbook
```

**Reorder Sheets:**
```fsharp
let newOrder = ["sheet3"; "sheet1"; "sheet2"]
let workbook' = reorderSheets newOrder workbook
```

### Cross-Sheet References

Reference cells in other sheets using `SheetName!CellRef` syntax:

```
=Sheet1!A1           // Reference A1 in Sheet1
='Budget Sheet'!B5   // Spaces require quotes
=SUM(Sheet2!A1:A10) // Range in another sheet
```

**Parsing:**
```fsharp
parseSheetReference "Sheet1!A1"
// Returns: (Some "Sheet1", "A1")

parseSheetReference "'My Budget'!B5"
// Returns: (Some "My Budget", "B5")
```

**Resolution:**
```fsharp
resolveSheetName "Budget" workbook
// Returns: Some "sheet2" (the sheet ID)
```

### Tabs UI Component

Visual tab navigation created with `Tabs.fs`:

```javascript
const tabHandlers = {
    OnTabClick: (sheetId) => { /* Switch sheets */ },
    OnTabAdd: () => { /* Create new sheet */ },
    OnTabRename: (sheetId, newName) => { /* Rename */ },
    OnTabDelete: (sheetId) => { /* Delete */ },
    OnTabReorder: (newOrder) => { /* Reorder */ }
};

const tabs = createTabs(workbook, tabHandlers);
tabs.Mount(container);

// Update tabs when workbook changes
tabs.Update(updatedWorkbook);
```

**Features:**
- Click tab to switch sheets
- Double-click to rename
- Click × to delete (protects last sheet)
- Click + to add new sheet
- Visual indication of active sheet

## Structured Tables

### Table Definition

Tables define named ranges with column headers:

```fsharp
type Table = {
    Name: string           // Table name (e.g., "Sales")
    SheetId: SheetId       // Which sheet it's on
    StartRow: int          // Top-left row (0-indexed)
    StartCol: int          // Top-left column
    EndRow: int            // Bottom-right row
    EndCol: int            // Bottom-right column
    Columns: string list   // Column names
    HasHeaders: bool       // First row is headers?
}
```

**Example:**
```fsharp
let salesTable = {
    Name = "Sales"
    SheetId = "sheet1"
    StartRow = 0
    StartCol = 0
    EndRow = 10
    EndCol = 3
    Columns = ["Date"; "Product"; "Amount"; "Region"]
    HasHeaders = true
}

let workbook' = defineTable salesTable workbook
```

### Table References (Structured)

Reference table columns using `TableName[ColumnName]` syntax:

```
=Sales[Amount]              // Column reference
=SUM(Sales[Amount])         // Aggregate function
=AVERAGE(Sales[Amount])     // Average of column
=Sales[@Amount]             // Current row (future)
```

**Parsing:**
```fsharp
parseTableReference "Sales[Amount]"
// Returns: Some ("Sales", "Amount")
```

**Column Range Lookup:**
```fsharp
getTableColumn "Sales" "Amount" workbook
// Returns: Some (sheetId, startAddr, endAddr)
// Example: ("sheet1", {Row=1; Col=2}, {Row=10; Col=2})
```

### Benefits of Tables

1. **Readable Formulas:**
   - `=Sales[Amount]` instead of `=C2:C11`
   - Self-documenting code

2. **Auto-Expand:**
   - Add rows to table, formulas automatically include them
   - No need to update range references

3. **Column Renaming:**
   - Rename "Amount" to "Revenue"
   - All formulas update automatically

4. **Scope Control:**
   - Table references are workbook-scoped
   - No ambiguity across sheets

### Table Operations

**Define Table:**
```fsharp
let workbook' = defineTable table workbook
```

**Delete Table:**
```fsharp
let workbook' = deleteTable "Sales" workbook
```

**Get Table:**
```fsharp
match getTable "Sales" workbook with
| Some table -> // Use table
| None -> // Not found
```

**List Tables:**
```fsharp
let names = getTableNames workbook
// Returns: ["Sales"; "Budget"; "Expenses"]
```

## Examples

### Multi-Sheet Budget

```fsharp
// Create workbook
let mutable wb = createWorkbook()

// Add sheets
wb <- addSheet "Income" wb
wb <- addSheet "Expenses" wb
wb <- addSheet "Summary" wb

// Define tables
let incomeTable = {
    Name = "Income"
    SheetId = "sheet2"  // Income sheet
    StartRow = 0
    StartCol = 0
    EndRow = 10
    EndCol = 2
    Columns = ["Source"; "Amount"; "Date"]
    HasHeaders = true
}
wb <- defineTable incomeTable wb

// Summary sheet uses cross-sheet references
// Cell A1: =SUM(Income!B2:B11)
// Or with table: =SUM(Income[Amount])
```

### Sales Dashboard

```javascript
// JavaScript/TypeScript usage

// Create workbook with sheets
let workbook = createWorkbook();
workbook = addSheet('Sales', workbook);
workbook = addSheet('Products', workbook);
workbook = addSheet('Dashboard', workbook);

// Define Sales table
const salesTable = {
    Name: 'Sales',
    SheetId: 'sheet2',
    StartRow: 0,
    StartCol: 0,
    EndRow: 100,
    EndCol: 4,
    Columns: ['Date', 'Product', 'Quantity', 'Price', 'Total'],
    HasHeaders: true
};
workbook = defineTable(salesTable, workbook);

// Dashboard formulas can reference:
// =SUM(Sales[Total])
// =AVERAGE(Sales[Price])
// =COUNTIF(Sales[Product], "Widget")
```

## Implementation Details

### Sheet Isolation

Each sheet has:
- Independent cell storage (SpreadsheetModel)
- Separate calculation engine (CalcEngine)
- Own dependency graph

Benefits:
- Changes in Sheet1 don't affect Sheet2's calc chain
- Parallel calculation possible (future)
- Sheet deletion doesn't break other sheets

### Cross-Sheet Dependencies

To support cross-sheet formulas:
1. Parse formula AST to detect sheet references
2. Build cross-sheet dependency edges
3. Topological sort across all sheets
4. Recalculate in correct order

**Future Work:** Currently each sheet calculates independently. Full cross-sheet support requires:
- Global dependency graph
- Multi-sheet calculation chain
- Reference invalidation across sheets

### Table Performance

Tables don't affect calculation performance:
- Table references resolve to cell ranges at formula time
- Formulas still operate on CellAddress ranges
- No extra indirection during calculation

### Memory Overhead

Per-table:
- 32 bytes (table metadata)
- 8 bytes per column name

Per-workbook with 10 sheets, 20 tables:
- ~1-2 KB overhead
- Negligible compared to cell data

## Testing

### Console Test

`test-workbook.fsx` demonstrates:
- Creating multiple sheets
- Sheet operations (add, rename, delete, reorder)
- Defining tables
- Cross-sheet reference parsing
- Table reference parsing and resolution

Run:
```bash
dotnet fsi test-workbook.fsx
```

### HTML Demo

`test-sheets-tabs.html` provides interactive demo:
- Visual tabs for sheet navigation
- Click tabs to switch sheets
- Double-click to rename
- Click + to add, × to delete
- Pre-populated sample data
- Toolbar integration
- Real-time stats

Open: http://localhost:8888/test-sheets-tabs.html

## Comparison with Excel/Sheets

### Excel

**Sheets:**
- ✓ Multiple sheets: Same
- ✓ Tab navigation: Same
- ✓ Cross-sheet formulas: Same syntax

**Tables:**
- ✓ Named tables: Same
- ✓ Structured references: Same syntax
- ✓ Auto-expand: Same concept
- ✗ Table styles: Not implemented
- ✗ Slicers/filters: Not implemented

### Google Sheets

**Sheets:**
- ✓ Multiple sheets: Same
- ✓ Tab operations: Same
- ✓ Cross-sheet formulas: Same

**Tables:**
- ✗ Named ranges: Different (Sheets uses named ranges, not tables)
- ✗ INDIRECT() required: Our tables are first-class

## Future Enhancements

### Advanced Table Features

1. **Calculated Columns:**
   - Define formula for entire column
   - Auto-fill new rows

2. **Table Totals Row:**
   - Automatic SUM/AVERAGE at bottom
   - Configurable aggregation

3. **Table Filtering:**
   - Filter rows by criteria
   - Hide/show rows dynamically

4. **Table Sorting:**
   - Sort by column
   - Multi-column sort

### Sheet Features

1. **Sheet Protection:**
   - Lock sheets from editing
   - Password protection

2. **Hidden Sheets:**
   - Hide sheets from tab bar
   - Still referenceable

3. **Sheet Groups:**
   - Group related sheets
   - Collapsible groups

4. **Sheet Templates:**
   - Pre-defined layouts
   - Quick sheet creation

### Cross-Sheet Improvements

1. **3D References:**
   - `=SUM(Sheet1:Sheet3!A1)`
   - Aggregate across multiple sheets

2. **Dynamic Ranges:**
   - `=INDIRECT("Sheet"&A1&"!B2")`
   - Computed sheet references

3. **Sheet-Level Names:**
   - Name scoped to sheet
   - Avoid global namespace pollution

## References

- Excel Tables: https://support.microsoft.com/en-us/office/overview-of-excel-tables
- Google Sheets: https://support.google.com/docs/answer/9055595
- LibreOffice Calc: Similar sheet/table model
