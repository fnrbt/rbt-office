# Architecture: Independent Components

The spreadsheet system is built from independent, composable components. **No UI assumptions are baked in** - apps wire components together however they want.

## Three Independent Layers

### 1. Data Models (No UI)

**Spreadsheet/Model.fs**
- Pure data: `SpreadsheetModel`, `CellValue`, `CellAddress`
- No rendering, no events, no UI
- Just cells, formulas, formatting

**Spreadsheet/Calc.fs**
- Pure computation: dependency graph, topological sort
- No UI, no rendering
- Just calculation engine

**Spreadsheet/Workbook.fs**
- Pure data: multiple sheets, tables
- No UI, no rendering
- Just workbook structure and operations

### 2. Rendering Components (Single Sheet)

**Spreadsheet/Spreadsheet.fs**
- Renders **one** SpreadsheetModel
- Doesn't know about workbooks
- Doesn't know about tabs
- Just a grid renderer

**Key Point:** The spreadsheet component takes a container and renders cells. Period.

```fsharp
type Spreadsheet(container: Element) =
    let mutable model = createEmpty 1000 26  // Just one sheet
```

### 3. Optional UI Components

**Spreadsheet/Tabs.fs**
- Renders sheet tabs
- Doesn't know about the grid
- Doesn't render cells
- Just navigation UI

**Spreadsheet/Toolbar.fs**
- Renders formatting buttons
- Doesn't know about the grid
- Just toolbar UI

**Key Point:** These are **optional**. Apps can use them, replace them, or omit them.

## Clean Separation

```
┌─────────────────────────────────────────────────────┐
│                   App Layer                          │
│  (Composes components however it wants)              │
└─────────────────────────────────────────────────────┘
           │              │              │
           ▼              ▼              ▼
    ┌──────────┐   ┌──────────┐   ┌──────────┐
    │Workbook  │   │Spreadsheet│   │  Tabs    │
    │(data)    │   │(renderer) │   │  (UI)    │
    └──────────┘   └──────────┘   └──────────┘
         │              │              │
         │              │              │
    No UI deps    No workbook     No grid deps
                  No tab deps
```

## Composition Patterns

### Pattern 1: Grid Only

Single sheet, no tabs, no workbook.

```javascript
const grid = createSpreadsheet(container);
grid.Mount();
```

**Use cases:**
- Embedded calculator
- Simple data entry form
- Single-sheet editor

### Pattern 2: Grid + Tabs (Excel Clone)

Multiple sheets with tab navigation.

```javascript
// Create workbook (data)
let workbook = createWorkbook();
workbook = addSheet('Budget', workbook);

// Create grid (renderer)
const grid = createSpreadsheet(container);
grid.Mount();

// Create tabs (navigation UI)
const tabs = createTabs(workbook, {
    OnTabClick: (sheetId) => {
        // Switch grid to show different sheet
        grid.SetModel(workbook.Sheets.get(sheetId).Model);
    }
});
tabs.Mount(tabContainer);
```

**Use cases:**
- Excel clone
- Multi-sheet workbook
- Data analysis tool

### Pattern 3: Grid + Custom Navigation

Use Workbook model but custom UI.

```javascript
let workbook = createWorkbook();
workbook = addSheet('Budget', workbook);

const grid = createSpreadsheet(container);
grid.Mount();

// Custom sidebar navigation (no Tabs component)
document.querySelectorAll('.sheet-item').forEach(item => {
    item.addEventListener('click', () => {
        const sheetId = item.dataset.sheetId;
        grid.SetModel(workbook.Sheets.get(sheetId).Model);
    });
});
```

**Use cases:**
- Custom dashboard
- Sidebar navigation
- Mobile apps (dropdown navigation)

### Pattern 4: No Workbook

Single sheet, use grid directly.

```javascript
const grid = createSpreadsheet(container);
grid.Mount();

// Just one sheet, no workbook data structure
// App manages state however it wants
```

**Use cases:**
- Simple spreadsheet
- Embedded in larger app
- Custom state management

## Anti-Pattern: Don't Bake UI Into Data

❌ **Wrong:** Workbook renders itself
```fsharp
// BAD: Workbook shouldn't render
type Workbook = {
    Sheets: Map<SheetId, Sheet>
    RenderTabs: unit -> Element  // ❌ NO!
}
```

✓ **Right:** Workbook is pure data
```fsharp
// GOOD: Workbook is just data
type Workbook = {
    Sheets: Map<SheetId, Sheet>
    SheetOrder: SheetId list
}
```

Apps decide how to render:
```javascript
// App chooses: tabs? sidebar? dropdown?
const tabs = createTabs(workbook, handlers);
// OR
const sidebar = createSidebar(workbook, handlers);
// OR
const dropdown = createDropdown(workbook, handlers);
```

## Component APIs

### Spreadsheet (Grid Renderer)

**Input:** Container element
**State:** SpreadsheetModel (one sheet)
**Output:** Rendered grid

```javascript
const grid = createSpreadsheet(container);
grid.Mount();
grid.SetCell(0, 0, '=A1+1');
grid.ApplyFormat('bold');
```

**No workbook knowledge!** Just renders one model.

### Tabs (Navigation UI)

**Input:** Workbook + handlers
**State:** None (stateless, controlled by app)
**Output:** Tab bar

```javascript
const tabs = createTabs(workbook, {
    OnTabClick: (id) => { /* app handles */ },
    OnTabAdd: () => { /* app handles */ },
    OnTabRename: (id, name) => { /* app handles */ },
    OnTabDelete: (id) => { /* app handles */ }
});
tabs.Mount(container);
tabs.Update(newWorkbook);  // When workbook changes
```

**No grid knowledge!** Just renders tabs, fires events.

### Workbook (Data Model)

**Input:** Operations
**State:** Sheets, tables, active sheet
**Output:** Updated workbook

```javascript
let workbook = createWorkbook();
workbook = addSheet('Budget', workbook);
workbook = setActiveSheet('sheet2', workbook);
workbook = defineTable(salesTable, workbook);
```

**No UI!** Just data transformations.

## Why This Matters

### Flexibility

Apps can:
- Use grid without tabs
- Use tabs with custom grid
- Use workbook model without tabs UI
- Build completely custom navigation

### Testability

Each component tests independently:
- Workbook: Test data operations (no UI needed)
- Grid: Test rendering (no workbook needed)
- Tabs: Test UI (mock workbook)

### Reusability

Components work in different contexts:
- Spreadsheet component: Excel clone, embedded calculator, data entry
- Tabs component: File tabs, page tabs, category tabs
- Workbook model: Desktop app, web app, mobile app

### Maintainability

Changes isolated:
- Change tab styling? Only Tabs.fs affected
- Change grid rendering? Only Spreadsheet.fs affected
- Change workbook structure? Only Workbook.fs affected

## Implementation Notes

### Current State

✓ Spreadsheet.fs: No workbook dependencies
✓ Workbook.fs: No UI dependencies
✓ Tabs.fs: Separate component
✓ Toolbar.fs: Separate component
✓ Calc.fs: Pure computation

**All components are properly separated!**

### App Layer Responsibilities

Apps must:
1. Create workbook (if multi-sheet)
2. Create grid(s) - one per sheet or reuse one
3. Create navigation UI (tabs, sidebar, dropdown, etc.)
4. Wire events: tab click → switch grid model
5. Manage state: active sheet, grid visibility

### Example: Complete Excel Clone

```javascript
// 1. Data layer
let workbook = createWorkbook();
workbook = addSheet('Budget', workbook);
workbook = addSheet('Sales', workbook);

// 2. Rendering layer - one grid, swap models
const grid = createSpreadsheet(gridContainer);
grid.Mount();

// 3. Navigation layer - tabs
const tabs = createTabs(workbook, {
    OnTabClick: (sheetId) => {
        const sheet = workbook.Sheets.get(sheetId);
        grid.SetModel(sheet.Model);  // Swap displayed model
        workbook = setActiveSheet(sheetId, workbook);
        tabs.Update(workbook);
    },
    OnTabAdd: () => {
        workbook = addSheet(`Sheet${workbook.Sheets.size + 1}`, workbook);
        tabs.Update(workbook);
    }
});
tabs.Mount(tabsContainer);

// 4. Optional toolbar
const toolbar = createFormatToolbar((action) => {
    grid.ApplyFormat(action);
});
toolbar.Mount(toolbarContainer);
```

## Testing

See `test-composition-patterns.html` for live examples:
- Pattern 1: Grid only (no tabs)
- Pattern 2: Grid + Tabs (Excel style)
- Pattern 3: Grid + Sidebar (custom nav)
- Pattern 4: Grid + Dropdown (mobile style)

Open: http://localhost:8888/test-composition-patterns.html

## Summary

**Three independent layers:**
1. Data models (Workbook, Model, Calc) - No UI
2. Renderers (Spreadsheet) - No workbook assumptions
3. UI components (Tabs, Toolbar) - No grid assumptions

**Apps compose them freely.**

**No baked-in UI assumptions.**

This is the correct architecture for a reusable spreadsheet library.
