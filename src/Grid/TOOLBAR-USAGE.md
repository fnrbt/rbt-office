# Spreadsheet Shell Usage

The spreadsheet ships with a unified shell component that combines a menubar, a single focused toolbar row, and an inline find/replace bar. Apps mount it once and route every command through one action dispatcher.

## Overview

`createSpreadsheetShell(onAction, onFind, onReplace)` returns an object with a `Mount(container)` method. The shell renders:

- **Menubar**: File, Edit, Insert, Format, Data
- **Toolbar row** (compact, ASCII glyphs): B, I, U | A+, A- | A (text color), Fill | L, C, R (alignment) | $, %, .0+, .0- (number formats) | Cut, Copy, Paste | Undo, Redo
- **Find/replace bar**: hidden by default; revealed via Edit > Find and replace, the `ShowFindReplace()` method, or a Ctrl+F handler your app installs

Niche actions (insert/delete row/col, sort, filter, find/replace) live in the menus only — the toolbar stays focused on the common ones.

## Callbacks

```ts
createSpreadsheetShell(
  onAction: (action: string) => void,
  onFind: (query: string) => void,
  onReplace: (pair: [string, string]) => void
): { Mount(container): void; GetElement(): Element; ShowFindReplace(): void }
```

`onAction` receives one of:
- Format: `bold`, `italic`, `underline`, `fontIncrease`, `fontDecrease`, `textColor`, `bgColor`
- Alignment: `alignLeft`, `alignCenter`, `alignRight`
- Number: `formatNumber`, `formatCurrency`, `formatPercent`, `decimalIncrease`, `decimalDecrease`
- Clipboard / history: `cut`, `copy`, `paste`, `undo`, `redo`
- Structure: `insertRow`, `deleteRow`, `insertColumn`, `deleteColumn`
- Data: `sortAsc`, `sortDesc`, `filter`, `clearFilter`
- File: `open`, `save`, `printPdf`

## Wiring example

```js
import createSpreadsheet from './Spreadsheet.js';
import { createSpreadsheetShell, createFormulaBar } from './Toolbar.js';

const spreadsheet = createSpreadsheet(document.getElementById('spreadsheet-root'));
spreadsheet.Mount();

const dispatchAction = (action) => {
  switch (action) {
    case 'textColor':      spreadsheet.ApplyTextColor('#ff0000'); break;
    case 'bgColor':        spreadsheet.ApplyBgColor('#ffff00'); break;
    case 'formatNumber':   spreadsheet.ApplyNumberFormat('number'); break;
    case 'formatCurrency': spreadsheet.ApplyNumberFormat('currency'); break;
    case 'formatPercent':  spreadsheet.ApplyNumberFormat('percentage'); break;
    case 'cut':            spreadsheet.Cut(); break;
    case 'copy':           spreadsheet.Copy(); break;
    case 'paste':          spreadsheet.Paste(); break;
    case 'undo':           spreadsheet.Undo(); break;
    case 'redo':           spreadsheet.Redo(); break;
    case 'sortAsc':        spreadsheet.SortCurrentColumnAsc(); break;
    case 'sortDesc':       spreadsheet.SortCurrentColumnDesc(); break;
    case 'insertRow':      spreadsheet.InsertRow(promptIndex()); break;
    // ...
    default:               spreadsheet.ApplyFormat(action);
  }
};

const shell = createSpreadsheetShell(
  dispatchAction,
  (query)         => spreadsheet.StartSearch(query, false),
  ([find, repl])  => spreadsheet.ReplaceAll(find, repl, false)
);
shell.Mount(document.getElementById('shell-container'));

// Optional: bind Ctrl+F to the find bar.
document.addEventListener('keydown', (e) => {
  if (e.ctrlKey && e.key === 'f') { e.preventDefault(); shell.ShowFindReplace(); }
});
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

Buttons carry a `data-sheet-icon="<id>"` attribute so downstream CSS can upgrade the ASCII labels (`B`, `+R`, `A-Z`, `Cut`, etc.) to icon-font or SVG renderings without touching the F#. Menu triggers carry `data-sheet-menu="<id>"`; menu items carry `data-sheet-menu-item="<id>"`.

The shell uses a flat, white-first palette with hairline borders, matching the Word web shell.
