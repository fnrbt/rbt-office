# Excel Advanced Features Tracker

Last updated: 2026-06-07 UTC.

This tracker is about **functionality** — formulas, calc, tables, pivots,
sparklines, charts, VBA, validation — i.e. *what the spreadsheet can do*. It
is the functional counterpart to `docs/excel-parity-tracker.md`, which tracks
**render fidelity** (does our pixels match Excel's). A feature can be "present"
here while still owing a parity pass there, and vice-versa.

## How features get verified (e2e, X11, look at screenshots)

The native interactive editor is the system under test:

- Binary: `tools/spreadsheet-editor/bin/Debug/net10.0/SpreadsheetEditor`
  (build: `dotnet build tools/spreadsheet-editor/SpreadsheetEditor.fsproj`).
  It opens a real X11 window (`Platform.X11Window`) running the full chrome
  (toolbar, formula bar, grid) over `Rbt.Spreadsheet`.
- It exposes a **control server** (`widgets/x11-viewer/ControlServer.fs`) over
  named-pipe FIFOs `/tmp/x11viewer-ctl-{pid}` + `-resp`. Frames are
  big-endian `>II` (cmdId, payloadLen). Supported commands:
  `Screenshot=1`, `Exit=2`, `GetText=10`, `ClickAt=11`, `RightClickAt=12`,
  `SendKey=14`, `MoveAt=15`, `SetScroll=6`, … (see the `CommandId` enum).
- `Screenshot` returns a **framebuffer PNG** rasterized from the actual paint
  commands (`Raster` + `PngEncode`), so screenshots are pixel-exact, not a
  window-manager grab. `tools/word-editor/snapshot.sh` and
  `widgets/x11-viewer/screenshot.sh` are the reference drivers.

**The e2e loop for a feature:** launch the editor → drive it via the FIFO
(`ClickAt` a cell, `SendKey` a formula, Enter) → `Screenshot` → open the PNG
with the Read tool and inspect the rendered result against the expected Excel
behavior. Real Excel ground truth, when needed, comes from the VBS/COM helpers
in `tools/excel-parity/` (run on the Office VM). A `tools/spreadsheet-editor/
drive.mjs|drive.py` thin client is planned (Batch 1, step 0) to script
cell→keystroke→screenshot sequences repeatably.

Status legend: ✅ done & screenshot-verified · 🟡 partial / model-only / one
render path · ⛔ missing.

---

## 1. Functions (formula engine — `src/Spreadsheet/Formula.fs`)

The evaluator dispatches ~80 functions (`Formula.fs` ~line 964,
`match name.ToUpperInvariant()`).

### Present ✅
Math/stats: SUM, PRODUCT, AVERAGE, MIN, MAX, MEDIAN, COUNT, COUNTA,
COUNTBLANK, STDEV, STDEVP, VAR, LARGE, SMALL, ABS, SQRT, SIGN, INT, TRUNC,
ROUND, ROUNDUP, ROUNDDOWN, MOD, POWER, RAND, RANDBETWEEN.
Conditional aggregates: COUNTIF, SUMIF, AVERAGEIF, SUMIFS, COUNTIFS, SUMPRODUCT.
Logical: IF, AND, OR, XOR, NOT, IFERROR, IFNA, **IFS**, **SWITCH**, TRUE, FALSE,
NA. IS-family: ISBLANK, ISERROR, ISNA, ISNUMBER, ISTEXT, ISLOGICAL.
Text: **CONCAT**, **CONCATENATE**, **TEXTJOIN**, LEN, TRIM, UPPER, LOWER,
PROPER, LEFT, RIGHT, MID, REPT, FIND, SEARCH, SUBSTITUTE, REPLACE, TEXT.
Lookup/ref: CHOOSE, INDEX, MATCH, VLOOKUP, HLOOKUP, OFFSET, INDIRECT.
Date: TODAY, NOW, DATE, YEAR, MONTH, DAY, WEEKDAY, EOMONTH.

### Batch 1 — landed ✅ (unit-verified, `tests/formula/stage6-modern.fsx`)
- Conditional aggregates: **MAXIFS**, **MINIFS**, **AVERAGEIFS** (reuse the
  SUMIFS criteria machinery; AVERAGEIFS → `#DIV/0!` empty, MAX/MINIFS → 0 empty).
- Lookup: **XLOOKUP**, **XMATCH** — match modes 0/-1/1/2 (exact / next-smaller
  / next-larger / wildcard), search 1/-1, `if_not_found`.
- Scoping: **LET** (named sub-results; inlines value *expressions* so range
  names stay ranges and dependency extraction still sees real refs).
- Still owed for Batch 1: real-Excel COM cross-check + a rendered screenshot
  pass (headless native-runner PNG; the interactive X11 editor won't launch in
  the current sandbox).

### Batch 2 functional — landed ✅ (`tests/formula/stage9-lambda.fsx`)
- **LAMBDA** + **MAP, REDUCE, SCAN, MAKEARRAY** via beta-reduction (expression
  substitution, consistent with LET): a LET-bound lambda call `f(5)` and a
  lambda passed to MAP/REDUCE/… are reduced and evaluated — no `CellValue`
  change. Render-verified (`synthetic-lambda.xlsx`: MAP/SCAN/MAKEARRAY spill,
  REDUCE + LET-bound lambda scalars).
- Deferred: direct IIFE `LAMBDA(x,…)(5)` (needs a parser apply node) and
  BYROW/BYCOL (array-valued lambda args).
- Dynamic arrays (need the spill engine, see §2): **FILTER**, **SORT**,
  **SORTBY**, **UNIQUE**, **SEQUENCE**, **RANDARRAY**, **TOROW/TOCOL**,
  **TEXTSPLIT**, **VSTACK/HSTACK**.
### Extended batch — landed ✅ (`tests/formula/stage10-extended.fsx`, 56 assertions)
Math: PI, EXP, LN, LOG, LOG10, SIN, COS, TAN, RADIANS, DEGREES, QUOTIENT,
CEILING, FLOOR, MROUND, EVEN, ODD, SUMSQ, FACT, COMBIN, GCD, LCM.
Text: VALUE, EXACT, CHAR, CODE, UNICHAR, UNICODE, CLEAN, TEXTBEFORE, TEXTAFTER,
FIXED, T. Date/time: TIME, HOUR, MINUTE, SECOND, EDATE, DAYS, DATEVALUE,
ISOWEEKNUM. Info: ISEVEN, ISODD, ISNONTEXT, N, ISFORMULA.

### Stats/date/lookup batch — landed ✅ (`tests/formula/stage11-stats.fsx`, 31)
Stats: AVERAGEA, MAXA, MINA, MODE/MODE.SNGL, RANK/RANK.EQ, PERCENTILE(.INC),
QUARTILE(.INC), DEVSQ, AVEDEV, GEOMEAN, HARMEAN, VAR.S/VAR.P/VARP, STDEV.S/.P.
Date: WEEKNUM, DATEDIF, NETWORKDAYS, WORKDAY. Lookup: LOOKUP (vector). Text:
NUMBERVALUE, **TEXTSPLIT** (array). ~150 functions total now.

### Financial batch — landed ✅ (`tests/formula/stage12-financial.fsx`, 9)
Time-value-of-money: PMT, FV, PV, NPER, RATE (Newton-iterated), NPV, IRR
(Newton-iterated), SLN. Excel sign convention; verified incl. a PMT↔PV
round-trip. ~160 functions total now.

- Still nice-to-have: IPMT/PPMT, XNPV/XIRR, depreciation (SYD/DB/DDB),
  YEARFRAC, DAYS360, FORMULATEXT, AGGREGATE, SUBTOTAL.

## 2. Dynamic arrays / spill 🟡 (engine + render-side spill landed)
- **Step 1 — array engine ✅** (`Formula.fs`, `tests/formula/stage8-arrays.fsx`,
  25 assertions): `ArrVal` (2-D value) + `evalArray` + element-wise operator
  **broadcasting** (`A1:A8>5`, `SEQUENCE(5)*2`). Functions: **SEQUENCE,
  RANDARRAY, UNIQUE, SORT, SORTBY, FILTER, TRANSPOSE**. Arrays flow through the
  engine — consumed by the streaming aggregates (`SUM(FILTER(…))`,
  `COUNT(UNIQUE(…))`) and criteria/lookup family — and collapse to their
  top-left element in scalar context. `#CALC!` on an empty FILTER/UNIQUE.
- **Step 2 — render-side spill ✅** (`Formula.spillArrays`, wired in the native
  runner): a formula using a spilling function writes its array into the anchor
  + neighbouring cells (row-major); an occupied target → `#SPILL!`. Verified by
  headless render of `tests/test-documents-excel/synthetic-dynamic-arrays.xlsx`
  (SORT/UNIQUE/FILTER/SEQUENCE, incl. a 2-D `SEQUENCE(3,2)` spill).
- **Step 3 — LAMBDA + higher-order ✅** (see §1): LAMBDA / MAP / REDUCE / SCAN
  / MAKEARRAY via beta-reduction; render-verified.
- **Still owed:** interactive spill (clear-on-recalc + spill-cell dependencies
  in `Calc.fs`), `#SPILL!` in the canvas/PDF paths, `by_col` UNIQUE/SORT,
  multi-key SORTBY, BYROW/BYCOL, direct LAMBDA IIFE. The streaming foundation
  underneath is in `docs/calc-streaming-fusion-design.md`.

## 3. Excel Tables / ListObjects ✅ (core complete)
- Model `Table` (+ `Name`, `HasTotals`) + `TableStyle` carry region, header
  count, style, autofilter range. Imported from `xl/tables/tableN.xml`
  (`XlsxReader` now also reads `displayName`, `<tableColumn>` names, and
  `totalsRowCount`).
- **Structured references — full grammar** (`Formula.fs`): item specifiers
  `[#All]`/`[#Data]`/`[#Headers]`/`[#Totals]`, this-row `[@Col]` (qualified +
  unqualified, via `EvalContext.CurrentCell`), column spans `Table[[A]:[B]]`,
  and multi-word column names. Lexer consumes a balanced `[...]` into one
  `TTableSpec`; `parseTableSpec` → `TableSel` AST; `expandTableSel`/
  `tableSelArea` resolve it.
- **Imported tables feed structured refs** (`Workbook.makeTableFromRegion` +
  the native runner's `buildWorkbookContext`): `Table1[Amount]` formulas
  resolve on recompute of real workbooks (was `#REF!`). Proof: runner
  `--tabletest`.
- **`SUBTOTAL` / `AGGREGATE`** (fn 1–12): filter-hidden + manually-hidden row
  exclusion (101–111 / opts 1/3/5/7), nested-subtotal skip, error skip
  (AGGREGATE opts 2/3/6/7).
- **Insert Table / Ctrl+T** (`Model.createTable`): registers a real styled +
  autofiltered table with an optional `SUBTOTAL(109,…)` **totals row**; editor
  binds Ctrl+T. Proof: runner `--tablecreatetest` (totals row + structured ref
  both compute).
- **Table filtering**: filter buttons render (native PNG); applying a filter
  collapses the filtered rows via `applyFilterRowHeights` while the styled
  header + buttons stay (proof: runner `--tablefiltertest`). Editor's filter
  dialog is table-aware (`tableAt` / `tableColumnUniqueValues` scope the
  checklist to the table's data rows).
- **Table styling rendered from `model.Tables`** (`Model.applyTableStyles`, run
  in `Paint.fs` `buildPaintItems`/`buildViewportPaintItems`): app-created tables
  render styled (header accent fill + white bold + filter buttons + banded rows)
  in the interactive grid + PDF + runner. Gated on `TableStyle.HeaderFill` being
  set, so it's a strict no-op for imported tables (whose runner bakes
  theme-accent styling into cells) — corpus renders byte-for-byte identical
  (md5-verified). Proof: runner `--tablecreatetest`.
- Tests: `stage17-tables` (50 assertions).
- Deferred ⛔: TableStyleLight/Dark overlays; table auto-extend on adjacent data
  entry; structured-ref autocomplete; AGGREGATE 13–19 (MODE/LARGE/SMALL/
  PERCENTILE/QUARTILE).

## 4. Pivot tables 🟡 (engine landed)
- **Engine ✅** — `src/Spreadsheet/Pivot.fs`:
  - `computePivot` — single row field, optional **column field** (cross-tab),
    one value field, Sum/Count/Average/Max/Min, Grand-Total row + column.
  - `computePivotN` — **nested row fields** (tabular, one column per level) with
    a **subtotal row per non-leaf group**, **multiple value fields** (one measure
    column each), and a Grand Total row — the "grouped list with subtotals" pivot.
  - Pure + unit-tested (`tests/formula/stage13-pivot.fsx`, 25 assertions:
    group-by, cross-tab, count/avg/max, nested subtotals, multi-measure, grand
    totals).
- **Render ✅** — native runner `--pivottest` renders a nested Region→Product
  pivot with Sum + Count measures, bold subtotals (East 470/4, West 700/3) and
  Grand Total (1170/7), verified visually.
- **Better-than-Excel direction ✅ — pivot as a LIVE spilling formula.**
  `GROUPBY(row_range, value_range, [agg])` and `PIVOTBY(row, col, value, [agg])`
  are dynamic-array functions (`Formula.fs` → `Pivot.groupBy`/`pivotBy`) that
  **spill**, so a pivot is a normal formula: recomputes on source edits (no
  Refresh), composable (`SUM(GROUPBY(...))`, FILTER/SORT it), and transparent/
  diffable. Render-verified (`synthetic-groupby.xlsx`: GROUPBY → East 470/West
  700/Grand 1170; PIVOTBY → full cross-tab).
- **UI ✅ — formula-writing builder with live preview.** The Insert-Pivot dialog
  (`SpreadsheetUi.pivotTableDialog`, an `Office.Ui.DialogSpec`) **writes a live
  `=GROUPBY`/`=PIVOTBY` formula** (`Pivot.buildFormula`) instead of static
  cells, and shows a **live preview** (the new `DialogComputedField` —
  reactive, re-runs against the model as fields change) of the formula + its
  first result rows. Wired in `spreadsheet-editor`. Resolution + formula
  builder unit-tested (`stage13-pivot.fsx`, 49 assertions incl. the full
  dialog→formula→spill pipeline). Screenshot-verified via
  `OfficeShell.dialogPreview` + runner `--dialogtest` (full dialog incl. live
  preview pane).
- **Nested column fields ✅** — `PIVOTBY` / `Pivot.pivotBy` now accept a
  multi-column column range → a multi-level column header (verified:
  `stage13-pivot.fsx` "ncol *").
- **Interactive spill auto-refresh ✅** — `Model.Spill` + `Formula.recalcAndSpill`;
  `SpreadsheetEditor.writeCell` recalcs + re-spills on every edit, so a
  GROUPBY/PIVOTBY (or any spilling formula) refreshes live in the editor and
  clears stale cells on shrink (verified: `stage13-pivot.fsx` "ispill *").
- **OOXML import-recompute ✅ (common case)** — `Pivot.fromCache` rebuilds a
  pivot from parsed cache rows + field layout (nested rows via `computePivotN`,
  cross-tab via `pivotBy`). The runner's `ChartRender.parsePivots` reads
  `pivotCacheDefinition`/`pivotCacheRecords` (shared-item resolution) +
  `pivotTableDefinition` (row/col/data field indices, subtotal, location) and
  places the recomputed grid (opt-in `EXCEL_RECOMPUTE_PIVOTS=1`, off by default
  so the cached-render corpus is untouched). Verified:
  `synthetic-pivot-import.xlsx` → Region/Sum of Sales recomputed from 5 cache
  records (East 350 / West 400 / Grand 750). Owed: multiple data fields, exact
  match of Excel's item ordering/subtotals on large multi-cache workbooks,
  following rels (currently convention-named cache1).
- Still owed: report filters / page fields, pivot styles, a drag-drop
  field-list pane.

## 5. Sparklines 🟡 (native runner)
In-cell mini-charts, parsed from the x14 `sparklineGroups` extLst and drawn as
vector paint at the cell rect (`tools/excel-native-runner/ChartRender.fs`:
`parseSparklines` + `buildSparkline`, overlaid like charts). All three types
render: **line ✅** (pre-existing), **column ✅** and **win/loss ✅** (added —
group `type` parsed, negatives drawn in `colorNegative`). Render-verified:
`tests/test-documents-excel/synthetic-sparklines.xlsx` (line / column±neg /
win-loss). Gaps: high/low/first/last **markers**, axis options, and wiring into
the canvas/PDF paths (currently native-runner only, like charts).

## 6. Charts 🟡 (native + PDF)
- Vector paint foundation `VecPath`/`VecGroup` (`Paint.fs` ~74) renders in all
  three paths (native/canvas/PDF), proven on a hand-built scene.
- `tools/excel-native-runner/ChartRender.fs` parses chart XML and renders
  column/bar/line/pie/area/scatter at parity (native); wired into the native
  import pipeline as anchored overlays (`renderXlsx` + `parseAllCharts`).
- **PDF ✅** — `Render.toBytesWith` takes anchored overlays (cell anchor + px
  offset + local paint items) and places each on the page whose vertical slice
  spans its anchor row (fallback page 0), positioned with the print model's own
  widths so charts paginate with the grid. The PDF runner reuses ChartRender
  (`parseAllCharts`/`buildChart`) to build them. Verified: chart-column and
  chart-zoo render through `pdftoppm` (column/line+markers/area).
- Missing ⛔: **canvas/web** chart wiring — blocked on a Fable-safe chart-XML
  parser (ChartRender uses `System.IO.Compression`/`Xml.Linq`, not Fable);
  interactive chart insertion; legends/multi-series/secondary-axis polish.

## 7. Data validation 🟡
Model carries a list-dropdown flag; `Paint.fs` (~782) draws the dropdown
arrow. Missing ⛔: enforcement on edit, the dropdown value picker, and the
other validation types (whole/decimal/date/textLength/custom), input/error
messages.

## 8. Conditional formatting 🟡
Model + `cellIs`/`containsText`/comparison rules; `resolveOperand` handles cell
references; imported and rendered. Missing ⛔: data bars, color scales, icon
sets, top/bottom & duplicate rules, and any rule-builder UI.

## 9. AutoFilter / filtering ✅
Model (`ColumnFilter`, `getVisibleRows`), checklist dropdown UI
(`Ui.filterDialog` — a checkbox per value + live "N of M values" preview),
and **render-side row collapsing** (`Model.applyFilterRowHeights` sets
filtered-out data rows to height 0 in `buildPaintItems`/`buildViewportPaintItems`,
so all three render paths collapse hidden rows while survivors keep their
original row numbers — strict no-op when no filter is active). Header row of
the data block is never hidden. Sorting is fully wired. Tests: `stage16-filter`
(13 assertions). Render proof: runner `--autofiltertest` (before/after).
Deferred: blue hidden-row boundary glyph; per-value numeric/date filters.

## 10. Freeze panes 🟡
Status unconfirmed in the interactive grid (import may carry pane state). Needs
a verification pass; flagged for a later batch.

## 11. Named ranges / defined names 🟡
Some plumbing (3 files); INDIRECT resolves. Workbook-scope defined names as
first-class formula operands not confirmed. Later batch.

## 12. VBA / macros 🟡 (more complete than the rest)
- Real engine: `FsVb` (compiled to `dist/spreadsheet/fsvb/FsVb.js`), driven by
  `src/Spreadsheet/SpreadsheetVba.fs`, both VBA and VBScript dialects
  (`run`/`runVbs`, plus snapshot/model variants).
- Host object model: Range (+ `.Value`/`.Formula`/`.Address`/resize), Cells,
  Worksheet, Workbook, Application, CreateObject.
- **Cross-platform write-capable runner ✅** — `Vba.runOnModel : model → source
  → Result<model,_>` runs a macro with read+write host objects over a
  `SpreadsheetModel` (no DOM), folding writes back + recomputing. Runs natively
  *and* in the web build (FsVb is plain F#). Tested via `stage19-vba` (12
  assertions, native `dotnet fsi`).
- **Run-from-UI ✅** — native editor's Insert→"Run macro..." opens a textarea
  dialog (`Ui.macroDialog`), runs `runOnModel`, keeps the source on
  `model.Macros`. Render proof: runner `--macrodialogtest`.
- **Events ✅** — `Vba.fireEvent`/`hasSub`; `Worksheet_Change` fires on each
  cell edit (`SpreadsheetEditor.writeCell`, passing the cell as Target),
  `Workbook_Open` on load (`loadFromPath`). No-op without the sub; writes apply
  directly so they can't recurse.
- `SpreadsheetModel.Macros` stores the workbook's module source.
- **`xl/vbaProject.bin` read + write ✅** (`Office.Formats/VbaProject.fs`): an
  [MS-CFB] compound-file reader/writer + [MS-OVBA] codec. `readMacroSource` /
  `extractFromXlsx` recover module source from real files (anchored on the
  `Attribute VB_Name` prefix, skipping any p-code); `writeProject` /
  `injectIntoXlsx` emit a **source-only** vbaProject.bin (Office recompiles from
  source). The editor loads macros from an opened `.xlsm` onto `model.Macros`.
  Tested by self round-trip + full `.xlsm` inject→extract→run (`stage20-vbabin`,
  9 assertions) **and verified against a real Excel-authored `.xlsm`**: the
  reader recovers the exact module source from genuine Office output (p-code +
  real OVBA compression) and it runs on our engine. Harness: `tools/vba/`
  (`gen-macro-xlsm.vbs` + `validate-real-xlsm.sh`); needs an Active VM session
  (`vm-rdp-keepalive.sh ensure`) and `AccessVBOM=1` in the Trust Center.
- **Compile cache + bytecode persistence ✅** — `Vba.compileCached` (reuse the
  BytecodeProgram), `compileToBytes`/`runBytesOnModel` (our serialized p-code),
  builtin-name lowercasing hoisted to compile time.
- Missing ⛔: UserForms, broader host API (Charts/ListObjects/PivotTables/
  Names/Selection), and emitting a `.xlsm` from the editor's Save (the
  `injectIntoXlsx` capability exists; the Save button isn't wired to it yet).

## 13. Slicers ⛔ / Images & shapes 🟡
Slicers: none. Images: anchored picture rendering is partial (flagged in the
dashboards note of the parity tracker). Shapes: preset geometry exists on the
Word side; spreadsheet drawing-ML shapes not wired.

---

## Batch roadmap

Ordered for value × independence × testability. Each batch ends with an
x11/screenshot e2e pass and tracker check-offs.

### Batch 1 — Modern scalar functions (NOW)
`MAXIFS, MINIFS, AVERAGEIFS, XLOOKUP, XMATCH, LET`. All evaluate to scalars,
all fit the current engine (one render path: calc), all directly requested.
Step 0 stands up the repeatable x11 driver. Proves the screenshot test loop on
low-risk additions before touching the engine. (Detail below.)

### Batch 2 — Dynamic-array spill engine + array functions (NEXT)
Array-valued `CellValue` + spill placement + `#SPILL!` + Calc/render wiring,
then `SEQUENCE, UNIQUE, SORT, SORTBY, FILTER, RANDARRAY`, and finally
`LAMBDA` (+ MAP/REDUCE/BYROW). The foundational engine change that unlocks the
rest of the modern-function suite. (Detail below.)

### Later batches (unordered backlog)
- Interactive Excel Tables (Ctrl+T, totals row, filter dropdown across paths).
- Charts into canvas/PDF + interactive insertion.
- Sparklines (model + import + in-cell vector render).
- VBA: run-from-UI + events + vbaProject import.
- Pivot tables (cache + field model + recompute).
- Data-validation enforcement + picker; conditional-formatting data
  bars/scales/icons; AutoFilter dropdown UI; freeze panes; named ranges;
  slicers; images/shapes.

---

## Batch 1 detail — Modern scalar functions

Target functions and semantics:
- **MAXIFS / MINIFS / AVERAGEIFS** — reuse the SUMIFS/COUNTIFS criteria
  evaluation (`Formula.fs` ~1047). AVERAGEIFS → `#DIV/0!` on empty match;
  MAXIFS/MINIFS → 0 on empty match (Excel behavior).
- **XLOOKUP(lookup, lookup_array, return_array, [if_not_found], [match_mode],
  [search_mode])** — match modes 0 exact (default), -1 exact-or-next-smaller,
  1 exact-or-next-larger, 2 wildcard; search 1 first / -1 last. Returns the
  aligned `return_array` element; `if_not_found` else `#N/A`.
- **XMATCH(lookup, lookup_array, [match_mode], [search_mode])** — same modes,
  returns 1-based position.
- **LET(name1, value1, …, calculation)** — parser binds name/value pairs into
  a scoped environment, later names may reference earlier ones; evaluate the
  final calculation in that env. No spill required.

Tests (x11 + screenshot, plus `tests/formula` unit fixtures):
- A fixture sheet exercising each function against known data; type each
  formula into the editor, Enter, Screenshot, and read the cell result.
- Cross-check a handful against Excel via a `tools/excel-parity` VBS probe.

Done when: all six evaluate correctly, screenshot-verified in the x11 editor,
unit fixtures pass, and these rows are checked off here.

## Batch 2 detail — Dynamic arrays + spill

1. `CellValue.Array of CellValue[][]` (or a dedicated spill result) + an
   evaluator that can return a rectangle.
2. Spill placement: a formula in the anchor cell owns the rectangle; spilled
   cells are read-only ghosts; collision → `#SPILL!`. Calc.fs dependency graph
   must treat the whole rectangle as outputs of one formula.
3. Render: native/canvas/PDF show spilled values (and the spill-range border
   when the anchor is selected).
4. Functions: SEQUENCE, UNIQUE, SORT, SORTBY, FILTER, RANDARRAY.
5. LAMBDA + first-class function values, then MAP/REDUCE/SCAN/BYROW/BYCOL.

Tests: `=SEQUENCE(5)` spills a column; `=SORT(UNIQUE(A1:A20))`; `=FILTER(...)`
with no match → `#CALC!`; a `#SPILL!` collision case — each screenshot-verified.
