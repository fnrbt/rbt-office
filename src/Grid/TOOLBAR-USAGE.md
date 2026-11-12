# Spreadsheet Shell Usage

The spreadsheet shell combines a menubar, selection-aware toolbar, Format Cells and data dialogs, and inline find/replace. Mount it alongside the grid; formatting, history, search, sorting, filtering, validation, XLSX files and printing run directly against that grid.

## Overview

`createSpreadsheetShell(spreadsheet, onAction)` returns an object with `Mount(container)`, `GetElement()`, `ShowFindReplace()`, `Dispose()` and `Commands` (the shared command registry). Format → Format cells or Ctrl/Cmd+1 opens a tabbed dialog with the current selection's formats; mixed fields remain unchanged until edited. The toolbar reads font, size, alignment and toggle states from the selection after edits, undo, sheet changes and import.

The find/replace bar is available from Edit → Find and replace, `ShowFindReplace()` or Ctrl/Cmd+F **while this shell or its grid owns the keyboard**. Results search the active sheet's current values; Find next selects the next match, Replace all changes actual cells as one undo step. Pending entries rejected by validation remain editable instead of being discarded.

## Callbacks

```ts
createSpreadsheetShell(
  spreadsheet: Spreadsheet,
  onAction: (action: string) => void
): { Mount(container): void; GetElement(): Element; ShowFindReplace(): void;
     Dispose(): void; Commands: CommandRegistry }
```

`onAction` remains for only `cut`, `copy`, `paste`, `insertRow`, `deleteRow`, `insertColumn` and `deleteColumn`. Hosts implement those actions as needed. It is **not called** for formatting, history, search, files, PDF, sorting, filtering or validation. `Commands.Run(id, argument?)` accepts parameterized font family, size and colors; `State(id)` provides enabled/checked/mixed values. Call the disposer returned by `Commands.OnChange(listener)` when the host stops observing it.

File → Open confirms unsaved replacement; Save validates and commits an active entry before downloading the opened filename (initially `Workbook1.xlsx`); Print PDF calls the grid directly. Ctrl/Cmd+O/S are scoped to the shell and grid. Dispose removes listeners and the dialog renderer.

Data → Custom sort, Filter and Data validation open selection-aware forms. They validate before applying, retain errors for correction and reject stale workbook/sheet state. Cancel leaves the workbook unchanged. Quick sort and Clear filters also run directly. Commands are `data.sort`, `data.sortAsc`, `data.sortDesc`, `data.filter`, `data.clearFilter` and `data.validation`; details and supported criteria are in [the host API](../../docs/spreadsheet-host-api.md#browser-shell-command-routing-and-selection-dialogs).

## Wiring example
For unbundled browser examples, load `node_modules/jszip/dist/jszip.min.js` and map the bare `jszip` import to a module exporting `window.JSZip` before loading the compiled spreadsheet modules (as in `test-full-interface.html`).

```js
import createSpreadsheet from '../../dist/spreadsheet/Spreadsheet.js';
import { createSpreadsheetShell, createFormulaBar } from '../../dist/spreadsheet/Toolbar.js';

const spreadsheet = createSpreadsheet(document.getElementById('spreadsheet-root'));
// The PDF exporter does not embed text without host-provided TrueType faces.
// Serve these from your app's own assets (including bold/italic faces if needed).
const fontResponse = await fetch('/fonts/LiberationSans-Regular.ttf');
if (!fontResponse.ok) throw new Error('Print font is unavailable');
spreadsheet.SetPdfFonts([{
  family: 'Calibri', bold: false, italic: false,
  data: new Uint8Array(await fontResponse.arrayBuffer())
}]);
spreadsheet.Mount();

const dispatchAction = (action) => {
  switch (action) {
    case 'cut': spreadsheet.Cut(); break;
    case 'copy': spreadsheet.Copy(); break;
    case 'paste': spreadsheet.Paste(); break;
    // Row and column structure actions remain host-owned.
  }
};

const shell = createSpreadsheetShell(spreadsheet, dispatchAction);
shell.Mount(document.getElementById('shell-container'));
// Dispose before removing the shell; no global Ctrl+F handler is needed.
```

## Formula bar

The formula bar is still a separate concern, mounted next to the shell:

```js
const formulaBar = createFormulaBar(onCellRefChange, onFormulaChange, onFormulaSubmit);
formulaBar.Mount(document.getElementById('formula-bar-container'));
formulaBar.SetCellRef('A1');
formulaBar.SetFormula('=A1+B1');
```

## Styling

Buttons carry a `data-sheet-icon="<id>"` attribute so downstream CSS can upgrade the ASCII labels (`B`, `+R`, `A-Z`, `Cut`, etc.) to icon-font or SVG renderings without touching the F#. Menu triggers carry `data-sheet-menu="<id>"`; menu items carry `data-sheet-menu-item="<id>"`. The filename is exposed as `.sheet-file-name` with `data-sheet-filename` and `data-sheet-dirty` attributes.

The shell uses a flat, white-first palette with hairline borders, matching the Word web shell.
