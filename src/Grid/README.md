# RedBat.Grid - Spreadsheet Components

Independent, composable spreadsheet components for building Excel-like applications.

## Key Principle

**No UI assumptions are baked in.** Components are separate layers that apps compose freely.

```
Data Models ──→ Renderers ──→ UI Components
(no UI)         (no workbook)  (no grid)
```

## Components

### Data Layer (No UI)

**Model.fs** - Cell data structures
```fsharp
SpreadsheetModel, CellValue, CellAddress
```

**Calc.fs** - Calculation engine
```fsharp
Dependency graph, topological sort, smart recalc
```

**Workbook.fs** - Multi-sheet structure
```fsharp
Workbook, Sheet, Table (structured references)
```

### Rendering Layer (No Workbook)

**Spreadsheet.fs** - Grid renderer
```javascript
const grid = createSpreadsheet(container);
grid.Mount();
```
Renders **one** SpreadsheetModel. Doesn't know about workbooks or tabs.

### UI Layer (Optional)

**Tabs.fs** - Sheet navigation
```javascript
const tabs = createTabs(workbook, handlers);
tabs.Mount(container);
```
Doesn't know about the grid. Just renders tabs.

**Toolbar.fs** - Formatting buttons
```javascript
const toolbar = createFormatToolbar(handleAction);
toolbar.Mount(container);
```
Doesn't know about the grid. Just renders buttons.

## Usage Patterns

### Single Sheet (No Tabs)

```javascript
import createSpreadsheet from './Spreadsheet.js';

const grid = createSpreadsheet(container);
grid.Mount();
grid.SetCell(0, 0, '=A1+1');
```

### Multi-Sheet (Excel Clone)

```javascript
import createSpreadsheet from './Spreadsheet.js';
import { createWorkbook, addSheet } from './Workbook.js';
import { createTabs } from './Tabs.js';

// Data
let workbook = createWorkbook();
workbook = addSheet('Budget', workbook);

// Renderer
const grid = createSpreadsheet(gridContainer);
grid.Mount();

// Navigation (separate)
const tabs = createTabs(workbook, {
    OnTabClick: (sheetId) => {
        // App switches grid to show different sheet
        grid.SetModel(workbook.Sheets.get(sheetId).Model);
    }
});
tabs.Mount(tabsContainer);
```

### Custom Navigation

```javascript
// Use workbook model, but custom UI (no Tabs component)
let workbook = createWorkbook();
const grid = createSpreadsheet(container);

// Sidebar, dropdown, or any custom navigation
document.querySelectorAll('.sheet-link').forEach(link => {
    link.addEventListener('click', () => {
        grid.SetModel(workbook.Sheets.get(link.dataset.sheetId).Model);
    });
});
```

## Features

### Calculation Engine
- Dependency graph tracking
- Topological sort (dependencies-first order)
- Dirty flagging (only recalc affected cells)
- Smart recalculation (100x+ speedup)
- Circular reference detection
- Formula caching (parse once, eval many)

See: `CALC-ENGINE.md`

### Multiple Sheets
- Workbook with multiple sheets
- Sheet operations (add, delete, rename, reorder)
- Active sheet tracking
- Cross-sheet references (`=Sheet2!A1`)

See: `SHEETS-AND-TABLES.md`

### Structured Tables
- Named ranges with column headers
- Table references (`=Sales[Amount]`)
- Auto-expand (future)
- Column-based formulas

See: `SHEETS-AND-TABLES.md`

### Formatting
- Text formatting (bold, italic, underline)
- Font size, colors (text, background)
- Text alignment
- Number formats (future)

### Formulas
- Arithmetic operators (+, -, *, /)
- Cell references (A1, B2)
- Functions (SUM, AVERAGE)
- Cross-sheet refs (`=Sheet2!A1`)
- Table refs (`=Sales[Amount]`)

## Architecture

```
┌─────────────────────────────────────────┐
│           App Layer                      │
│  (Composes components freely)            │
└─────────────────────────────────────────┘
         │           │           │
         ▼           ▼           ▼
  ┌──────────┐ ┌──────────┐ ┌──────────┐
  │Workbook  │ │Spreadsheet│ │  Tabs    │
  │(data)    │ │(renderer) │ │  (UI)    │
  └──────────┘ └──────────┘ └──────────┘
```

See: `ARCHITECTURE.md`

## Examples

### Test Pages

**Composition Patterns:** `test-composition-patterns.html`
- Grid only (no tabs)
- Grid + tabs (separate)
- Grid + sidebar (custom nav)
- Grid + dropdown (mobile)

**Calculation Engine:** `test-calc-visualization.html`
- Dependency graph visualization
- Dirty cell highlighting
- Performance metrics

**Formatting:** `test-spreadsheet.html`
- Excel-style grid
- Toolbar integration
- Formula examples

**Multi-Sheet:** `test-sheets-tabs.html`
- Multiple sheets with tabs
- Cross-sheet formulas
- Structured tables

### Test Scripts

**Calculation:** `test-calc-engine.fsx`
```bash
dotnet fsi test-calc-engine.fsx
```

**Workbook:** `test-workbook.fsx`
```bash
dotnet fsi test-workbook.fsx
```

## Building

Compile F# to JavaScript with Fable:

```bash
cd hosts/mega/frontend
pnpm fable ../../../components/RedBat.Grid --outDir src/lib/grid
```

Output: `src/lib/grid/Spreadsheet/*.js`

## Documentation

- `ARCHITECTURE.md` - Component separation, composition patterns
- `CALC-ENGINE.md` - Calculation optimizations, performance
- `SHEETS-AND-TABLES.md` - Multi-sheet, structured tables
- `TOOLBAR-USAGE.md` - Formatting toolbar API

## Design Philosophy

### Separation of Concerns

**Data, Rendering, and UI are separate layers.** Apps wire them together.

### No Baked-In Assumptions

The grid doesn't assume tabs exist. Tabs don't assume a grid exists. Apps compose freely.

### Composable Components

Each component works independently. Mix and match as needed.

### Testable

Pure data models test without UI. Renderers test without data. UI tests with mocks.

## License

Part of RedBatNet project.
