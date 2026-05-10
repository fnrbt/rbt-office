# XLSX Import Integration - office-format

## Overview

RedBat.Grid now supports importing Excel files (.xlsx) using the office-format library. This document describes the integration architecture and usage.

## Architecture

### New Module: `XlsxImporter.fs`

Created `components/RedBat.Grid/Spreadsheet/XlsxImporter.fs` as an adapter between the office-format library and RedBat.Grid's SpreadsheetModel.

**Key Functions:**
- `importFromBuffer`: Import XLSX from ArrayBuffer → SpreadsheetModel
- `getSheetNames`: Extract sheet names from workbook
- `importSheet`: Core conversion logic (internal)

**Type Conversions:**

```fsharp
office-format Cell Type → RedBat.Grid CellValue
==========================================
Number                 → CellValue.Number
String                 → CellValue.String
Boolean                → CellValue.Boolean
Date (serial)          → CellValue.Number
Formula (text)         → CellValue.Formula (re-evaluated by grid)
Empty                  → CellValue.Empty
```

```fsharp
office-format Style    → RedBat.Grid CellFormat
===========================================
bold                   → Bold: bool
italic                 → Italic: bool
fontColor              → TextColor: string option
fillColor              → BgColor: string option
fontSize               → FontSize: int option
numberFormat           → NumberFormat: string option
```

**Layout Conversions:**
- Column widths: Excel units → pixels (×7)
- Row heights: Points → pixels (×1.333 at 96 DPI)

### Public API

Added to `Spreadsheet.fs`:

```javascript
// Import XLSX from ArrayBuffer
await spreadsheet.ImportXlsxFromBuffer(buffer, sheetName)

// Get sheet names before importing
const sheets = await spreadsheet.GetXlsxSheetNames(buffer)
```

**Usage Pattern:**

```javascript
// From File input
const file = event.target.files[0];
const buffer = await file.arrayBuffer();

// List sheets (if multiple)
const sheets = await spreadsheet.GetXlsxSheetNames(buffer);

// Import (null = first sheet)
await spreadsheet.ImportXlsxFromBuffer(buffer, sheetName || null);
```

## Features Imported

### Cell Data
- ✅ Numbers, strings, booleans, dates
- ✅ Formula text (re-evaluated by RedBat.Grid)
- ✅ Empty cells (sparse storage)

### Cell Formatting
- ✅ Bold, Italic
- ✅ Font color, background color
- ✅ Font size
- ✅ Number formats

### Layout
- ✅ Column widths (custom per-column)
- ✅ Row heights (custom per-row)

### Multi-Sheet Support
- ✅ Sheet selector UI
- ✅ Import specific sheet by name
- ✅ Default to first sheet

## Not Implemented (office-format limitations)

The office-format library doesn't support these yet:
- Underline formatting
- Merged cells
- Charts, images
- Pivot tables
- Data validation rules
- Conditional formatting (in file format)
- Cell comments

## Undo Support

XLSX import is undo-able (Ctrl+Z):
- Restores previous cell data
- Restores previous column widths / row heights
- Uses command pattern with undo/redo functions

## Test Page

Created `test-xlsx-import.html` demonstrating:
- File upload UI
- Sheet selection (for multi-sheet workbooks)
- Import button with status display
- Integration with existing spreadsheet features

## Integration with office-format

The adapter expects the office-format library to be available as:

```javascript
import('~/lib/office-format')
```

### office-format API Used

```javascript
const api = (await import('~/lib/office-format')).default();

// Load workbook from ArrayBuffer
const workbook = await api.excel.load(buffer);

// Get sheet names
const sheets = api.excel.sheetNames(workbook);

// Get specific sheet
const sheet = api.excel.sheet(sheetName, workbook);

// Get cell values as 2D array
const data = api.excel.toArray(workbook, sheet);

// Get cell with style
const cell = api.excel.cellWithStyleAt('A1', sheet);
// Returns: { value, formula, bold, italic, fontColor, fillColor, fontSize, numberFormat }

// Get layout info
const width = api.excel.columnWidth(col, sheet);
const height = api.excel.rowHeight(row, sheet);
```

## Preparation Steps

To use XLSX import, ensure:

1. **office-format Library Compiled:**
   ```bash
   cd ~/Dev/office-format
   pnpm run build
   ```

2. **Library Available in Frontend:**
   ```bash
   # Copy or link office-format output to src/lib/office-format/
   cd hosts/mega/frontend
   mkdir -p src/lib/office-format
   cp -r ~/Dev/office-format/dist/* src/lib/office-format/
   ```

3. **Import Path Configured:**
   - XlsxImporter.fs uses `[<ImportAll("~/lib/office-format")>]`
   - Vite should resolve `~/lib/` to `src/lib/`

## Error Handling

If import fails:
- Console logs error message
- Returns empty SpreadsheetModel (doesn't crash UI)
- User can retry with different file

## Future Enhancements

When office-format adds support:
- Underline formatting
- Merged cells (span multiple cells)
- Cell comments / notes
- Conditional formatting import
- Charts (render as images or skip)

## File Structure

```
components/RedBat.Grid/
├── Spreadsheet/
│   ├── XlsxImporter.fs         # NEW: XLSX import adapter
│   ├── Spreadsheet.fs           # Updated: Added Import* methods
│   ├── Model.fs                 # Existing: CellValue, CellFormat types
│   └── ...
├── test-xlsx-import.html        # NEW: Test page
└── XLSX-IMPORT-INTEGRATION.md   # This file
```

## Testing

1. Create a sample .xlsx file with:
   - Numbers, formulas, text
   - Bold/italic formatting
   - Background colors
   - Different column widths

2. Open `test-xlsx-import.html` in browser

3. Click "Choose File" and select the .xlsx

4. If multiple sheets, select one from dropdown

5. Click "Import XLSX"

6. Verify:
   - Data appears correctly
   - Formulas are re-evaluated
   - Formatting is preserved
   - Column widths match source

7. Test undo (Ctrl+Z) to revert import

## Notes for Future Developers

- **Formula Evaluation:** office-format extracts formula TEXT only. RedBat.Grid re-parses and evaluates formulas using its own engine (Formula.fs).

- **Sparse Storage:** Only non-empty cells are stored in the model. Empty cells are not materialized.

- **Default Sizes:** Minimum grid is 100 rows × 26 columns (A-Z), even if imported sheet is smaller.

- **File Input:** The Spreadsheet component expects ArrayBuffer, not File objects. JavaScript side must call `file.arrayBuffer()` first.

- **Sheet Names:** Passing `null` or omitting sheetName imports the first sheet.

## Build Status

✅ XlsxImporter.fs compiles successfully
✅ Spreadsheet.fs integration compiles successfully
⚠️ Pre-existing errors in Hooks.fs/Grid.fs (unrelated namespace conflicts)

All new XLSX import code is ready to use once office-format library is available at runtime.
