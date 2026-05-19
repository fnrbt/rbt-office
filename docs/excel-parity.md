# Excel Rendering Parity Work

Current goal: iteratively improve RedBat Spreadsheet render/layout parity against Microsoft Excel for real XLSX files and focused generated repros. Mirrors `docs/word-parity.md`.

Active tracker: `docs/excel-parity-tracker.md`. Visual fix history:
`docs/excel-parity-fixes.html`.

## Confirmed Repo Map

- Browser/Fable Spreadsheet path: `src/Spreadsheet/`
  - Layout core: `src/Spreadsheet/Layout.fs`
  - Paint commands: `src/Spreadsheet/Paint.fs`
  - Calculation/formula engine: `src/Spreadsheet/Formula.fs`, `src/Spreadsheet/Calc.fs`
  - Workbook/sheet model: `src/Spreadsheet/Workbook.fs`
  - Shared UI specs: `src/Spreadsheet/SpreadsheetUi.fs` (consumes `RedBat.Office.Ui` types)
  - Browser shell entry points: `src/Spreadsheet/Spreadsheet.fs`, `src/Spreadsheet/Toolbar.fs`, `src/Spreadsheet/Tabs.fs`
- Native path: `spreadsheet-native/`
  - Native grid widget: `spreadsheet-native/SpreadsheetGrid.fs`
  - Native chrome wiring: `spreadsheet-native/SpreadsheetUiNative.fs`
  - Uses the shared chrome renderer in `office-native/OfficeShell.fs` plus `../widgets`.
- XLSX read path: `src/Spreadsheet/XlsxImporter.fs`
- Shared UI primitives: `src/OfficeUi/Specs.fs` (consumed by both Word and Spreadsheet).
- Shared native chrome: `office-native/OfficeShell.fs` (rendered by `Word`, `Spreadsheet`, and future Office-style apps via the `OfficeShell.view`/`ShellConfig` API).

## Fixtures

Excel fixtures live under (to be populated):

- `tests/repros/excel/` — focused, deterministically generated XLSX repros.
- `tests/test-documents-excel/` — captured real-world workbooks.

The Word VM (`wordvm`) used by the parity tooling also ships Excel; the same VBS-on-task-scheduler infrastructure works for Excel automation.

## Comparison Loop

1. Pick a real XLSX fixture or generate a focused repro with the XLSX writer.
2. Query Microsoft Excel for cell/row/column/range geometry through VBA.
3. Query our browser path via the in-progress `getLayoutCells`/`getRangeInformation` browser exports.
4. Query native path via `spreadsheet-native`/`../widgets`.
5. Compare cell rectangles, header positions, merged-range geometry, and rendered cell text first; then inspect screenshots.
6. Fix the smallest layout/import/export discrepancy that explains the measured delta.
7. Keep the repro fixture and comparison output.

For each identified bug, prefer the stricter loop used by the Word parity work:

1. Add or generate a focused XLSX repro under `tests/repros/excel/`.
2. Capture the old wrong rendering before editing code. Use a temporary git worktree at the pre-fix commit and write PNG/JSON outputs under `artifacts/baseline/`; these artifacts are for records and are not committed.
3. Fix the importer/layout/rendering bug.
4. Re-run Excel/browser/native as appropriate.
5. Commit the repro and the code fix together, or commit repro/tooling separately when retrofitting an already-fixed bug.

## Current Machine State

- `ssh wordvm hostname` works and returns `DESKTOP-OP2QS0N`. The same VM hosts Microsoft Excel.
- Excel automation lives in checked-in `.vbs` files invoked with `cscript`; avoid one-off SSH/PowerShell command lines.
- The Linux wrappers default to `WORD_VM_RUN_MODE=task`, copying the checked-in VBS to the VM and running it through an interactive-user scheduled task. Excel needs the same path because the SSH-launched `cscript` cannot instantiate `Excel.Application` directly.
- Windows/Excel fonts come from `tools/word-parity/fetch-windows-fonts.sh`. The same fonts cover Excel cell rendering.

## Added Tooling (skeleton — wire up as needed)

- `tools/excel-parity/extract-excel-layout.vbs` — extracts Microsoft Excel cell geometry/format/value to TSV, including row/column index, merged ranges, and the active sheet name.
- `docs/excel-vba-layout-api.md` — tracks the Microsoft Excel VBA layout-query API surface and which parts our browser API mirrors.
- `tools/excel-parity/print-excel-pdf.vbs` — opens an XLSX in Excel and prints it to PDF.
- `tools/excel-parity/run-excel-vbs-on-vm.sh` — copies a VBS script and XLSX to `wordvm`, runs `cscript`, and copies the output back. Mirrors `run-word-vbs-on-vm.sh`.
- `tools/excel-parity/run-excel-vbs-task.ps1` — runs the copied VBS through Windows Task Scheduler with an interactive-user token. Mirrors `run-word-vbs-task.ps1`.
- `tools/excel-parity/excel-layout-tsv.sh` — wraps the layout extractor.
- `tools/excel-parity/excel-print-pdf.sh` — wraps the PDF generator.
- `tools/excel-parity/compare-vba-layout-results.mjs` — runs Excel TSV extraction plus browser layout-call result extraction for matched cell ranges; writes per-document layout diffs and a ranked summary.
- `tools/excel-parity/capture-browser-xlsx.mjs` — builds/serves the Fable
  Spreadsheet bundle and captures an imported XLSX through the browser
  portable-canvas path.
- `tools/excel-parity/validate-fixes-gallery.mjs` — checks
  `docs/excel-parity-fixes.html` for missing image references and broken
  row cells.

Run on the Windows VM from an interactive session:

```cmd
cscript //nologo tools\excel-parity\extract-excel-layout.vbs C:\path\input.xlsx C:\path\excel-layout.tsv
```

Or from Linux, using the VM wrapper:

```sh
tools/excel-parity/excel-layout-tsv.sh tests/test-documents-excel/sample.xlsx artifacts/excel/sample.tsv
tools/excel-parity/excel-print-pdf.sh tests/test-documents-excel/sample.xlsx artifacts/excel/sample.pdf
```

The TSV columns are:

```text
doc	sheet	row	column	x	y	width	height	mergedRows	mergedCols	format	value
```

Excel positions are reported in points relative to the worksheet origin. The comparator converts them to CSS pixels with `96 / 72` before comparing against the browser cell rectangles.

## Next Setup Tasks

- Land an `XlsxWriter` (parallel to `DocumentWriter.fs`) so the create-focused-repros script can emit deterministic XLSX fixtures.
- Add browser-side `getLayoutCells`, `getRangeInformation`, and `computeRangeStatistics` exports in `RedBat.Spreadsheet.Spreadsheet`/`RedBat.Spreadsheet.API`.
- Add a native geometry export alongside the native grid widget if we need numeric native cell positions, not just raster snapshots.
- Resume parity fixes on cell rendering, header geometry, and merged ranges using real-world fixtures and the focused XLSX writer.
