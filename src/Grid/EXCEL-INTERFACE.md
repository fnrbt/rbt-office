# Excel-Like Interface Components

> **Heads up:** The per-category toolbar classes documented below (`FormatToolbar`, `AlignmentToolbar`, `NumberFormatToolbar`, `InsertDeleteToolbar`, `DataToolbar`, `ClipboardToolbar`, `FindReplaceToolbar`) have been consolidated into a single `SpreadsheetShell` that renders a menubar + one focused toolbar row + an inline find/replace bar — matching the Word web shell's chrome. See `TOOLBAR-USAGE.md` for the current API. The component descriptions below remain as a reference for which actions live where in the new shell.

## Overview

RedBat.Grid now provides a comprehensive set of Excel-like interface components that host applications can compose as needed. This document describes all available toolbars and UI components.

## Architecture

All interface components are **host-assembled** - the spreadsheet component provides the building blocks, but the host application decides where to place them and how to wire them up.

```
┌─────────────────────────────────────────────┐
│  Host Application (decides layout)         │
│  ┌──────────────────────────────────────┐  │
│  │ Title Bar                            │  │
│  ├──────────────────────────────────────┤  │
│  │ Toolbar Row 1 (Format + Alignment)   │  │
│  │ Toolbar Row 2 (Data + Insert/Delete) │  │
│  ├──────────────────────────────────────┤  │
│  │ Formula Bar                          │  │
│  ├──────────────────────────────────────┤  │
│  │ Spreadsheet Grid                     │  │
│  ├──────────────────────────────────────┤  │
│  │ Status Bar                           │  │
│  └──────────────────────────────────────┘  │
└─────────────────────────────────────────────┘
```

## Available Components

### 1. Formula Bar

Shows the active cell reference and allows editing cell formulas/values.

**Create:**
```javascript
import { createFormulaBar } from './Toolbar.js';

const formulaBar = createFormulaBar(
    // onCellRefChange - called when user types in cell ref input
    (ref) => console.log('Cell ref:', ref),

    // onFormulaChange - called as user types in formula input
    (formula) => console.log('Formula:', formula),

    // onFormulaSubmit - called when user presses Enter
    () => {
        const formula = document.querySelector('.formula-input').value;
        spreadsheet.SetCurrentCellFromFormula(formula);
    }
);

formulaBar.Mount(container);
```

**Methods:**
- `SetCellRef(ref: string)` - Update cell reference display (e.g., "A1")
- `SetFormula(formula: string)` - Update formula input field

**Typical Usage:**
```javascript
// Update formula bar when selection changes
function updateFormulaBar() {
    const cellRef = spreadsheet.GetSelectionRef();      // "A1" or "A1:B10"
    const formula = spreadsheet.GetCurrentCellFormula(); // "=SUM(A1:A10)" or "123"
    formulaBar.SetCellRef(cellRef);
    formulaBar.SetFormula(formula);
}

// Poll for changes (in real app, spreadsheet would emit events)
setInterval(updateFormulaBar, 100);
```

### 2. Clipboard Toolbar

Cut, copy, paste, undo, redo operations.

**Create:**
```javascript
import { createClipboardToolbar } from './Toolbar.js';

const clipboardToolbar = createClipboardToolbar((action) => {
    switch (action) {
        case 'cut': spreadsheet.Cut(); break;
        case 'copy': spreadsheet.Copy(); break;
        case 'paste': spreadsheet.Paste(); break;
        case 'undo': spreadsheet.Undo(); break;
        case 'redo': spreadsheet.Redo(); break;
    }
});

clipboardToolbar.Mount(container);
```

**Buttons:**
- ✂ Cut (also Ctrl+X)
- 📋 Copy (also Ctrl+C)
- 📄 Paste (also Ctrl+V)
- ↶ Undo (also Ctrl+Z)
- ↷ Redo (also Ctrl+Y)

### 3. Format Toolbar

Text formatting: bold, italic, underline, font size, colors.

**Create:**
```javascript
import { createFormatToolbar } from './Toolbar.js';

const formatToolbar = createFormatToolbar((action) => {
    if (action === 'textColor') {
        spreadsheet.ApplyTextColor('#ff0000');
    } else if (action === 'bgColor') {
        spreadsheet.ApplyBgColor('#ffff00');
    } else {
        spreadsheet.ApplyFormat(action);  // bold, italic, underline, etc.
    }
});

formatToolbar.Mount(container);
```

**Buttons:**
- **B** Bold
- *I* Italic
- <u>U</u> Underline
- A+ Font increase
- A- Font decrease
- 🎨 Text color
- ⬛ Background color

### 4. Alignment Toolbar

Text alignment within cells.

**Create:**
```javascript
import { createAlignmentToolbar } from './Toolbar.js';

const alignmentToolbar = createAlignmentToolbar((action) => {
    spreadsheet.ApplyFormat(action);
});

alignmentToolbar.Mount(container);
```

**Buttons:**
- ⬅ Align left
- ↔ Align center
- ➡ Align right

### 5. Number Format Toolbar

Number formatting: currency, percentage, decimals.

**Create:**
```javascript
import { createNumberFormatToolbar } from './Toolbar.js';

const numberToolbar = createNumberFormatToolbar((action) => {
    switch (action) {
        case 'formatNumber': spreadsheet.ApplyNumberFormat('number'); break;
        case 'formatCurrency': spreadsheet.ApplyNumberFormat('currency'); break;
        case 'formatPercent': spreadsheet.ApplyNumberFormat('percentage'); break;
        // decimalIncrease, decimalDecrease not yet implemented
    }
});

numberToolbar.Mount(container);
```

**Buttons:**
- 1,234.56 Number format (2 decimals, thousands separator)
- $ Currency format
- % Percentage format
- .0+ Increase decimal places
- .0- Decrease decimal places

### 6. Insert/Delete Toolbar

Insert and delete rows and columns.

**Create:**
```javascript
import { createInsertDeleteToolbar } from './Toolbar.js';

const insertDeleteToolbar = createInsertDeleteToolbar((action) => {
    switch (action) {
        case 'insertRow':
            const rowIdx = parseInt(prompt('Insert row at index:') || '0');
            spreadsheet.InsertRow(rowIdx);
            break;
        case 'deleteRow':
            const delRowIdx = parseInt(prompt('Delete row at index:') || '0');
            spreadsheet.DeleteRow(delRowIdx);
            break;
        case 'insertColumn':
            const colIdx = parseInt(prompt('Insert column at index:') || '0');
            spreadsheet.InsertColumn(colIdx);
            break;
        case 'deleteColumn':
            const delColIdx = parseInt(prompt('Delete column at index:') || '0');
            spreadsheet.DeleteColumn(delColIdx);
            break;
    }
});

insertDeleteToolbar.Mount(container);
```

**Buttons:**
- ⬇+ Insert row
- ⬇− Delete row
- ➡+ Insert column
- ➡− Delete column

**Note:** Currently uses prompts for row/column index. In a real app, you'd determine the index from the current selection automatically.

### 7. Data Toolbar

Sorting and filtering operations.

**Create:**
```javascript
import { createDataToolbar } from './Toolbar.js';

const dataToolbar = createDataToolbar((action) => {
    switch (action) {
        case 'sortAsc': spreadsheet.SortCurrentColumnAsc(); break;
        case 'sortDesc': spreadsheet.SortCurrentColumnDesc(); break;
        case 'filter': /* not yet implemented */ break;
        case 'clearFilter': /* not yet implemented */ break;
    }
});

dataToolbar.Mount(container);
```

**Buttons:**
- ↓A Sort ascending (A→Z)
- ↓Z Sort descending (Z→A)
- ⚲ Filter (infrastructure in place, UI not implemented)
- ✕ Clear filters

### 8. Find/Replace Toolbar

Search and replace text in cells.

**Create:**
```javascript
import { createFindReplaceToolbar } from './Toolbar.js';

const findReplaceToolbar = createFindReplaceToolbar(
    // onFind
    (query) => {
        spreadsheet.StartSearch(query, false);
        const count = spreadsheet.GetSearchResultsCount();
        console.log(`Found ${count} result(s)`);
        if (count > 0) spreadsheet.FindNext();
    },

    // onReplace
    (findText, replaceText) => {
        spreadsheet.ReplaceAll(findText, replaceText, false);
        console.log(`Replaced all "${findText}" with "${replaceText}"`);
    }
);

findReplaceToolbar.Mount(container);
```

**UI Elements:**
- Find input field
- "Find" button
- Replace input field
- "Replace" button

**Typical Usage:**
```javascript
// Show/hide with Ctrl+F
document.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.key === 'f') {
        e.preventDefault();
        findReplaceContainer.classList.add('visible');
        document.querySelector('.find-input').focus();
    }
});
```

## Spreadsheet Methods

New methods added to support toolbars:

### Formula Bar Integration

```javascript
// Get current selection as cell reference string
const ref = spreadsheet.GetSelectionRef();  // "A1" or "A1:B10"

// Get formula/value of active cell
const formula = spreadsheet.GetCurrentCellFormula();  // "=SUM(A1:A10)" or "123"

// Set cell value from formula bar input
spreadsheet.SetCurrentCellFromFormula(input);  // Handles formulas and values
```

### Clipboard Operations

```javascript
spreadsheet.Cut();   // Cut selection to clipboard
spreadsheet.Copy();  // Copy selection to clipboard
spreadsheet.Paste(); // Paste from clipboard
```

### Undo/Redo

```javascript
spreadsheet.Undo();  // Undo last operation
spreadsheet.Redo();  // Redo last undone operation
```

### Sorting

```javascript
spreadsheet.SortCurrentColumnAsc();   // Sort active column ascending
spreadsheet.SortCurrentColumnDesc();  // Sort active column descending
```

## Layout Patterns

### Excel-Style Ribbon

```html
<div class="ribbon">
    <div class="ribbon-row">
        <div class="toolbar-section" id="clipboard"></div>
        <div class="toolbar-section" id="format"></div>
        <div class="toolbar-section" id="alignment"></div>
    </div>
    <div class="ribbon-row">
        <div class="toolbar-section" id="insert-delete"></div>
        <div class="toolbar-section" id="data"></div>
    </div>
</div>
```

### Vertical Toolbar Stack

```html
<div class="toolbar-stack">
    <div id="clipboard-toolbar"></div>
    <div id="format-toolbar"></div>
    <div id="data-toolbar"></div>
</div>
```

### Minimal Interface

```html
<!-- Just formula bar + spreadsheet -->
<div id="formula-bar"></div>
<div id="spreadsheet"></div>
```

## Status Bar Pattern

```javascript
// Update status bar
function updateStatusBar() {
    const cellCount = Object.keys(spreadsheet.GetModel().Cells).length;
    const selection = spreadsheet.GetSelectionRef();

    document.getElementById('cell-count').textContent = `${cellCount} cells`;
    document.getElementById('selection-info').textContent = `Selection: ${selection}`;
}
```

## Keyboard Shortcuts

The spreadsheet already handles:
- **Ctrl+C** - Copy
- **Ctrl+X** - Cut
- **Ctrl+V** - Paste
- **Ctrl+Z** - Undo
- **Ctrl+Y** - Redo
- **Enter** - Confirm cell edit
- **F2** - Edit cell
- **Delete** - Clear cell
- **Arrow keys** - Navigate

Additional shortcuts can be added by the host:
```javascript
document.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.key === 'b') {
        e.preventDefault();
        spreadsheet.ApplyFormat('bold');
    }
    if (e.ctrlKey && e.key === 'f') {
        e.preventDefault();
        showFindReplace();
    }
});
```

## Example: Full Interface

See `test-full-interface.html` for a complete example showing:
- Title bar
- Multi-row ribbon with all toolbars
- Formula bar with auto-update
- Spreadsheet grid
- Status bar with selection info
- Collapsible find/replace (Ctrl+F)

## Styling

All toolbar components use inline styles for portability. Host applications can override with CSS:

```css
/* Override toolbar button styles */
.toolbar-button {
    padding: 8px 16px;
    border-radius: 4px;
    background: #f0f0f0;
}

.toolbar-button:hover {
    background: #e0e0e0;
}

/* Custom formula bar styling */
.formula-bar {
    background: #ffffff;
    border-bottom: 2px solid #e0e0e0;
    padding: 8px;
}

.formula-input {
    font-family: 'Monaco', 'Consolas', monospace;
    font-size: 12px;
}
```

## Future Enhancements

Planned additions:
- Context menus (right-click on cells/headers)
- Cell comments/notes toolbar
- Conditional formatting UI
- Chart toolbar
- Font family selector
- More number format options (date, time, scientific)
- Toolbar customization API
- Keyboard shortcut configuration

## Notes for Host Applications

1. **Selection Change Events:** Currently, the host must poll for selection changes. Future versions will emit events when selection changes.

2. **Toolbar State:** Toolbars don't automatically update state (e.g., bold button doesn't show pressed state when active cell is bold). Host applications can implement this by querying cell format and updating button styles.

3. **Modal Dialogs:** Some operations (like insert/delete row) currently use `prompt()`. Host applications should replace these with proper modal dialogs.

4. **Customization:** All toolbars are standard JavaScript objects. Host applications can inspect button elements and modify them as needed.

5. **Localization:** All button labels are currently English. To localize, either:
   - Create wrapper functions that pass translated labels
   - Modify button text after mounting: `button.textContent = translations[buttonId]`

## Build Status

✅ All toolbar components compile successfully
✅ All spreadsheet integration methods compile successfully
✅ Test page demonstrates full interface

The only remaining compile errors are pre-existing namespace conflicts in Hooks.fs/Grid.fs (unrelated to this work).
