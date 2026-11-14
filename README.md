# rbt-office

Document and spreadsheet editing for the browser: open, edit and save `.docx` and `.xlsx`
files, with layout that follows Office's. Written in F# and compiled to JavaScript with Fable;
nothing is sent to a server.

**[Try the demo](https://fnrbt.github.io/rbt-office-demo/)**: open a sample or your own file
([demo source](https://github.com/fnrbt/rbt-office-demo)).

## Use

The package `@rbt/office` ships the compiled modules under `dist/`. A page can mount the
built-in shells (menus, toolbar, open, save, print to PDF):

```js
import "@rbt/office/office-formats/src/Main.js";
import * as Host from "@rbt/office/word/WordHost.js";
import * as Shell from "@rbt/office/word/WordShell.js";
import * as Dom from "@rbt/office/word/FableUI/Dom.js";
import createSpreadsheet from "@rbt/office/spreadsheet/Spreadsheet.js";
import { createSpreadsheetShell } from "@rbt/office/spreadsheet/Toolbar.js";

// Documents
const host = Host.create("", "word-shell");
Dom.mountApp("#word", Shell.view(host));
Host.openFile(host, file);                // a File holding a .docx

// Spreadsheets
const grid = createSpreadsheet(document.querySelector("#grid"));
grid.Mount();
createSpreadsheetShell(grid, () => {}).Mount(document.querySelector("#toolbar"));
await grid.ImportXlsx(bytes);             // a Uint8Array holding an .xlsx
```

Or drive the engines directly: see [`docs/word-host-api.md`](docs/word-host-api.md) and
[`docs/spreadsheet-host-api.md`](docs/spreadsheet-host-api.md).

## Build

`pnpm install`, then `pnpm run build` (needs the .NET SDK; `dotnet tool restore` first).
Tests: `bash tests/run-all.sh --core`.

## Licence

MIT; see [LICENSE](LICENSE).
