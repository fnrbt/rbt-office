# Excel Features Implementation Status

This document tracks the implementation status of Excel-like features in RedBat.Spreadsheet.

## ✅ Phase 1: Foundation Features (COMPLETE)

### 1. Range References in Formulas ✅
**Status:** Fully implemented

- **Parser:** `Formula.fs` lines 132-149 - Parses `A1:B10` syntax with colon token
- **Evaluator:** `Formula.fs` lines 229-239 - `expandRangeRef` expands ranges to cell lists
- **Functions:** SUM, AVERAGE work with ranges: `=SUM(A1:A10)`
- **Dependency Tracking:** `Calc.fs` lines 55-62 - Extracts all cells in range for dependency graph

**Examples:**
```
=SUM(A1:A10)       → Sum cells A1 through A10
=AVERAGE(B2:B20)   → Average of range
=COUNT(C1:C100)    → Count non-empty cells
```

### 2. Copy/Paste/Cut ✅
**Status:** Fully implemented with undo support

- **Model:** `Model.fs` lines 80-438
  - `ClipboardData` type (lines 80-86)
  - `serializeSelection` (lines 382-410) - Captures selected cells
  - `pasteClipboard` (lines 412-438) - Pastes with formula adjustment
  - `adjustFormulaReferences` (lines 323-380) - Adjusts relative refs when pasting
- **UI:** `Spreadsheet.fs` lines 232-329
  - **Ctrl+C** (lines 232-236): Copy to clipboard
  - **Ctrl+X** (lines 238-254): Cut and clear source
  - **Ctrl+V** (lines 256-329): Paste with formula adjustment and undo command

**Formula Adjustment:**
- Copying `=A1+10` from E1 to E2 → becomes `=A2+10`
- Relative references adjust by delta (row/col offset)
- Undo/redo supported for all paste operations

### 3. More Functions ✅
**Status:** All implemented in `Formula.fs`

| Function | Lines | Description | Example |
|----------|-------|-------------|---------|
| SUM | 304-322 | Sum of numbers | `=SUM(A1:A10)` |
| AVERAGE | 324-345 | Average of numbers | `=AVERAGE(B1:B20)` |
| COUNT | 347-368 | Count non-empty cells | `=COUNT(C1:C100)` |
| MIN | 392-413 | Minimum value | `=MIN(A1:A10)` |
| MAX | 415-436 | Maximum value | `=MAX(A1:A10)` |
| IF | 370-390 | Conditional logic | `=IF(A1>100, "High", "Low")` |

**IF Function Details:**
- Booleans: TRUE/FALSE work as expected
- Numbers: Non-zero = TRUE, zero = FALSE (Excel behavior)
- Strings: Non-empty = TRUE, empty = FALSE
- Errors propagate through

### 4. Undo/Redo ✅
**Status:** Fully implemented with command pattern

- **Model:** `Model.fs` lines 174-489
  - `Command` type (lines 175-179) - Undo/Redo functions + description
  - `History` type (lines 182-186) - Past/Future stacks with size limit (50 default)
  - `pushCommand` (lines 440-455) - Add command, clear future
  - `undo` (lines 457-472) - Execute undo, move to future
  - `redo` (lines 474-489) - Execute redo, move to past
- **UI:** `Spreadsheet.fs` lines 331-347
  - **Ctrl+Z** (lines 331-338): Undo last command
  - **Ctrl+Y** (lines 340-347): Redo last undone command

**Commands with Undo Support:**
- Cell edits (SetCell)
- Paste operations
- Formatting changes
- Sorting
- Insert/delete rows/columns

---

## ✅ Phase 2: Data Operations (MOSTLY COMPLETE)

### 5. Column Sorting ✅
**Status:** Fully implemented with UI

- **Model:** `Model.fs` lines 519-574
  - `sortByColumn` - Sort by column with ascending/descending
  - `SortState` type - Track current sort column and direction
  - `compareCellValues` - Compare numbers < strings < booleans < empty
- **UI:** `Spreadsheet.fs` lines 406-465, 1478-1524
  - **Click column header** - Toggle sort direction (lines 443-465)
  - **Sort indicators** - Shows ▲ (ascending) or ▼ (descending) in header (lines 433-439)
  - **API methods:** `SortCurrentColumnAsc()`, `SortCurrentColumnDesc()`
- **Undo:** Full undo/redo support for sorting operations

**Sorting Behavior:**
- First click: Ascending
- Second click: Descending
- Third click: Remove sort (back to original order)
- Maintains row integrity (all cells in row move together)

### 6. Column Filtering ⚠️
**Status:** Model implemented, NO UI

- **Model:** `Model.fs` lines 94-98, 586-622
  - `ColumnFilter` type - Stores visible values per column
  - `applyFilter` / `removeFilter` - Add/remove filters
  - `getVisibleRows` - Returns set of visible row indices (AND logic for multiple filters)
  - `isRowVisible` used in rendering (`Spreadsheet.fs` line 112)
- **Missing:**
  - ❌ Filter icon/button in column headers
  - ❌ Dropdown UI showing unique values with checkboxes
  - ❌ API methods to show/hide filter UI

**TODO:** Add filter dropdown UI to column headers

### 7. Number Formatting ✅
**Status:** Fully implemented

- **Model:** `Model.fs` lines 274-299
  - `formatNumber` - Applies format codes to numbers
  - `NumberFormat` field in `CellFormat` (line 37)
- **Formats Supported:**
  - `"currency"` → `$1,234.56`
  - `"percentage"` → `45.67%`
  - `"date"` → `01/15/2024` (Excel serial date)
  - `"decimal2"` → `123.46` (2 decimals)
  - `"number"` → `1,234.56` (comma separators)

**TODO:** Add toolbar buttons for quick formatting

### 8. Insert/Delete Rows & Columns ✅
**Status:** Fully implemented

- **Model:** `Model.fs` lines 624-756
  - `insertRow` (lines 624-654) - Shift rows down, update selection
  - `deleteRow` (lines 656-689) - Remove row, shift up
  - `insertColumn` (lines 691-721) - Shift columns right
  - `deleteColumn` (lines 723-756) - Remove column, shift left
- **API:** Public methods available on Spreadsheet class
- **Formula Updates:** References automatically adjusted when rows/columns shift

**Missing:** Right-click context menu UI

---

## ⚠️ Phase 3: Polish Features (PARTIAL)

### 9. Row/Column Resizing ✅
**Status:** Fully implemented

- **Model:** `Model.fs` lines 152-156, 758-774
  - `RowHeights` / `ColumnWidths` maps - Custom sizes per row/column
  - Default: 25px rows, 100px columns
- **UI:** `Spreadsheet.fs` lines 466-493 (columns), 527-554 (rows)
  - Drag column divider in header
  - Drag row divider in header
  - Visual feedback during resize
  - Persistent after resize

### 10. Find/Replace ⚠️
**Status:** Model implemented, NO Ctrl+F UI

- **Model:** `Model.fs` lines 158-165, 777-867
  - `SearchQuery` / `SearchResults` / `SearchIndex` - Search state
  - `searchCells` (lines 777-792) - Find all matches
  - `startSearch` (lines 795-801) - Initialize search
  - `nextSearchResult` (lines 803-815) - Cycle through results
  - `replaceAtSelection` (lines 817-844) - Replace current
  - `replaceAll` (lines 846-867) - Batch replace
- **API:** `Spreadsheet.fs` lines 1277-1305
  - `StartSearch(query, caseSensitive)` - Begin search
  - `FindNext()` - Navigate to next match
  - `ReplaceAtSelection(find, replace, caseSensitive)` - Replace current
  - `ReplaceAll(find, replace, caseSensitive)` - Replace all
- **Missing:**
  - ❌ Ctrl+F dialog/modal
  - ❌ Find/Replace input fields
  - ❌ Match case checkbox
  - ❌ "Find Next" / "Replace" / "Replace All" buttons

**TODO:** Add modal dialog UI for find/replace

### 11. Conditional Formatting ⚠️
**Status:** Model implemented, NO UI

- **Model:** `Model.fs` lines 100-112, 167-168, 869-897
  - `ConditionalFormatRule` type - GreaterThan, LessThan, Between, Equal, ContainsText
  - `ConditionalFormat` type - Rule + Style to apply
  - `evaluateRule` (lines 870-878) - Check if value matches rule
  - `getConditionalFormat` (lines 880-887) - Get format for cell
  - `addConditionalFormat` (lines 889-893) - Add rule to cell
  - `removeConditionalFormats` (lines 895-897) - Clear rules
- **Missing:**
  - ❌ UI for defining rules
  - ❌ Visual rule builder
  - ❌ Rule management dialog

**TODO:** Add conditional format rule builder UI

### 12. CSV Import/Export ✅
**Status:** Fully implemented

- **Model:** `Model.fs` lines 899-975
  - `parseCSV` (lines 900-909) - Parse CSV text to rows/columns
  - `importCSV` (lines 911-942) - Populate cells from CSV
  - `exportCSV` (lines 944-975) - Serialize to CSV with escaping
- **Features:**
  - Auto-detect numbers vs strings
  - Formula import (cells starting with `=`)
  - Quote escaping for commas
  - Boolean parsing

**TODO:** Add file upload/download UI buttons

---

## 🎨 Additional Implemented Features

### Merged Cells ✅
**Status:** Fully implemented

- **Model:** `Model.fs` lines 114-121, 171-172, 977-1013
  - `MergedCell` type - Define rectangular merge range
  - `findMergedCell` - Find merge containing address
  - `addMergedCell` / `removeMergedCell` - Manage merges
- **Rendering:** `Spreadsheet.fs` - Spans cell rendering
- **Selection:** Clicking merged cell selects entire range

### Cell Borders ✅
**Status:** Fully implemented

- **Model:** `Model.fs` lines 20-24, 38-41
  - `BorderStyle` type - Style (thin/medium/thick/dashed/dotted/double) + color
  - Four borders per cell: Top, Bottom, Left, Right
- **Rendering:** `Spreadsheet.fs` - CSS borders applied
- **Import:** `XlsxImporter.fs` imports border styles from Excel files

### Cell Comments/Notes ✅
**Status:** Fully implemented

- **Model:** `Model.fs` lines 62-66, 72
  - `CellComment` type - Text + optional Author
- **Rendering:** `Spreadsheet.fs` - Red triangle indicator + hover tooltip
- **API:** `SetComment(row, col, text, author)` and `RemoveComment(row, col)`
- **Import:** `XlsxImporter.fs` imports comments from Excel files

### Cell Alignment ✅
**Status:** Fully implemented

- **Model:** `Model.fs` lines 34-35
  - `HorizontalAlign`: left, center, right, justify
  - `VerticalAlign`: top, middle, bottom
- **Rendering:** Applied via CSS
- **Import:** `XlsxImporter.fs` imports alignment from Excel

### Text Wrapping ✅
**Status:** Fully implemented

- **Model:** `Model.fs` line 36 - `WrapText: bool`
- **Rendering:** CSS `white-space: normal` vs `nowrap`
- **Import:** Preserved from Excel files

### Table References ✅
**Status:** Fully implemented for tables defined in workbook

- **Syntax:** `TableName[ColumnName]`
- **Example:** `=SUM(Sales[Amount])`
- **Parser:** `Formula.fs` lines 120-129
- **Evaluator:** `Formula.fs` lines 242-257
- **Workbook:** Requires `Workbook` context with table definitions

---

## 📊 Summary Statistics

| Category | Implemented | Partial | Missing |
|----------|-------------|---------|---------|
| **Phase 1: Foundation** | 4/4 | 0 | 0 |
| **Phase 2: Data Ops** | 3/4 | 1 | 0 |
| **Phase 3: Polish** | 2/4 | 2 | 0 |
| **Additional Features** | 6/6 | 0 | 0 |
| **TOTAL** | **15/18** | **3** | **0** |

**Completion:** 83% core features, 100% model logic

---

## 🚧 TODO: Missing UI Components

1. **Column Filtering** - Add dropdown to headers with value checkboxes
2. **Find/Replace Dialog** - Add Ctrl+F modal with search inputs
3. **Conditional Formatting Builder** - Add UI for defining rules
4. **Number Format Toolbar** - Quick buttons for currency, percentage, etc.
5. **CSV Upload/Download Buttons** - File I/O UI
6. **Right-Click Context Menus** - Insert/delete rows/columns

All model logic exists for these features - only UI layer needed.

---

## 🧪 Testing

- **Manual Test:** `test-phase1-features.html` - Tests range refs, functions, copy/paste, undo/redo
- **Playwright:** `tests/test-phase1-features.ts` - Automated testing (WIP)
- **Excel Import:** `tests/test-xlsx-import.ts` - XLSX file import testing
- **Excel Features:** `tests/test-excel-features.ts` - Borders, merged cells, comments

---

## 📝 Architecture Notes

### Calculation Engine
- **Dependency Graph:** `Calc.fs` tracks which cells depend on which
- **Topological Sort:** Ensures formulas calculate in correct order
- **Smart Recalc:** Only recalculates dirty cells (cells that changed or depend on changed cells)
- **Circular Detection:** Detects and handles circular references

### Event Flow
1. User edits cell → `SetCell` API call
2. Update model → `setCell` in Model.fs
3. Extract formula dependencies → `extractReferences` in Calc.fs
4. Mark dependents dirty → `markDirty` in Calc.fs
5. Rebuild calc chain → `rebuildCalcChain`
6. Recalculate dirty cells → `smartRecalculate`
7. Render updated cells → `render()`

### Undo System
- Command pattern with Undo/Redo functions
- Captures old state before mutation
- Stack limited to 50 commands (configurable)
- Clearing future stack on new command (standard behavior)

---

## 🎯 Next Steps

**High Priority (UI Only):**
1. Add filter dropdown to column headers
2. Add Ctrl+F find/replace modal
3. Add conditional format rule builder

**Medium Priority:**
1. Number format toolbar buttons
2. CSV upload/download buttons
3. Right-click context menus

**Low Priority:**
1. Data bars for conditional formatting
2. Color scales for ranges
3. Icon sets (arrows, flags, etc.)
4. Named ranges
5. Freeze panes
6. Charts/graphs
