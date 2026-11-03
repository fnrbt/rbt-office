# Excel Parity Tracker

Last updated: 2026-05-20 UTC.

This tracker is for deciding what to fix next and recording which parity
surfaces have actually been checked. It complements
`docs/excel-parity-fixes.html`, which is the visual history of fixed issues.

## Current Checks

- [x] Gallery image references: `docs/excel-parity-fixes.html` has 167 image
  references and 0 missing targets.
- [x] Fixture inventory: 85 `.xlsx` files currently exist under
  `tests/test-documents-excel/`.
- [x] Print helper tests: `dotnet fsi tests/print/helpers.fsx` passes
  16/16.
- [x] PDF runner build:
  `dotnet build tools/spreadsheet-pdf-runner/SpreadsheetPdfRunner.fsproj --no-restore`
  succeeds.
- [x] Whitespace/syntax diff check: `git diff --check` passes.
- [x] VBA-layout corpus gate:
  `SKIP_SMOKE=1 tools/excel-parity/check-parity.sh` passes with cached
  Excel COM TSVs (`78/78` layout fixtures clean, 10 print-PDF references
  within thresholds, print helpers passed).
- [x] Full PDF smoke render: 85/85 XLSX fixtures render through
  `tools/spreadsheet-pdf-runner` with outputs written under
  `tmp/parity-smoke`.
- [x] Native screenshot smoke diffs:
  `pandas-dimension_large` and `pandas-test1` are pixel-identical to the
  saved native baselines (`mae=0.0000`, `bbox=None`).
- [x] Core print PDF-to-PDF comparison pass:
  `tools/excel-parity/compare-print-pdfs.py --reuse-excel --out artifacts/pdf-compare-excel`
  over `synthetic-header-footer-runs.xlsx`,
  `synthetic-header-footer-runs-footer.xlsx`, `synthetic-print.xlsx`, and
  `synthetic-print-chrome.xlsx`.
- [x] Gallery row validator:
  `tools/excel-parity/validate-fixes-gallery.mjs` passes with 55 rows,
  167 images, and no missing render cells.
- [x] Focused browser/Fable XLSX import screenshot pass:
  `tools/excel-parity/capture-browser-xlsx.mjs` captures
  `synthetic-browser-styles.xlsx` with 26 imported cells.
- [x] Focused browser/Fable rich-text screenshot pass:
  `tools/excel-parity/capture-browser-xlsx.mjs` captures
  `synthetic-inline-rich.xlsx` with 3 cells and 2 rich-run entries.
- [x] Small browser/Fable imported-XLSX corpus pass:
  `synthetic-browser-styles`, `pandas-test_converters`, `poi-47813`,
  `synthetic-inline-rich`, and `synthetic-fmt5` capture successfully through
  `tools/excel-parity/capture-browser-xlsx.mjs`.
- [ ] Full native corpus screenshot pass against saved baselines.
- [ ] Full PDF-to-PDF comparison pass for all print fixtures.
- [x] Browser/Fable screenshot pass for imported XLSX fixtures: 85/85
  workbook captures pass through `tools/excel-parity/capture-browser-xlsx.mjs`.


## Real-world: dcf-stock-analysis.xlsx + conditional-formatting-tutorial.xlsx

- dcf-stock-analysis (Main, Financial Model): at parity. Stock Snapshot / Key
  Events / Management / Profile boxes (light-blue headers + borders) match;
  ~3px global origin offset only.
- conditional-formatting-tutorial (Demographics, Sales): at parity. Despite
  the name it has NO conditional formatting — two dense data tables. Columns
  align to 0–2px; elevated MAE (~22) is dense small-text ClearType-vs-grayscale
  AA (text in every cell × 11 cols). Only nit: a 1-char text-clip difference
  on a couple of overflow cells (column-width rounding).
- Status: both at render parity, no fix needed.

## Real-world goal: milk-street-office.xlsx

- Both sheets (Annual, Monthly) at AA-level parity via the global fixes — no
  model-specific bug. Aligned MAE ~8.5 (Annual) / ~7 (Monthly). Verified
  region-by-region: acquisition overview, assumptions, renovation/exit/
  financing (Holdback, LIBOR, Mezzanine), and the Monthly rent-roll + tenant
  sections. Values, yellow inputs, dates, currency, and grouped-row outline
  all match. The only embedded drawing is a legacy VML drawing (not a chart).
- Status: at render parity.

## Dashboards (assessment only — large multi-feature targets)

dashboard-sales-analysis.xlsx (8 sheets; 5 charts / 16 pivots / 2 images) and
dashboard-ecommerce-sales.xlsx (19 charts / 30 pivots / 11 images).

Rendered the sales-analysis Dashboard sheet: the new chart pipeline already
draws several charts in-place (bar "Sales By Category/Manager", line "Sales By
Chain"). NOT yet rendered: slicers (filter-button controls), in-cell
sparklines, images/logos; pivot tables show as cached cell values (OK).
Reaching parity needs: slicer rendering, sparklines, image embedding, and
chart-position/layout polish. Not started — flagged as the big next goal.

## Chart zoo (goal: render chart-zoo.xlsx at parity, one type at a time)

chart-zoo.xlsx (build-chart-zoo.vbs): one of each common type. All six render
at parity via ChartRender.fs (--charttest <xlsx> <n> <out>):
- Column ✓, Bar ✓ (bars within 1–2px), Line ✓, Pie ✓ (palette slices,
  clockwise from 12 o'clock), Area ✓ (edge-to-edge fill, MAE ~4),
  Scatter ✓ (numeric axes + full gridline grid, MAE ~4).
- Each logged as a gallery row.
- Polish done: pie radius/centre calibrated (MAE ~4); palette read from
  theme1.xml accents; charts WIRED INTO THE NORMAL PIPELINE — renderXlsx
  overlays parseAllCharts results at anchor cell rects (Layout.cellRect) via
  SpreadsheetGrid.widgetWith, so the whole file renders with charts in-place
  (full-sheet MAE ~8.5 vs Excel).
- Remaining: anchor->chart mapping assumes document order; chart sizing uses
  default 64/20 col/row px; only sheet1 drawing; canvas/PDF paths render the
  vector primitives but charts not yet wired into those runners.

## Charts (new)

- Vector paint foundation (commit 80d5379): PaintItem += VecPath/VecGroup,
  rendered by native/canvas/PDF; proven identical on a hand-built scene.
- First chart instance AT PARITY: clustered column chart
  (synthetic-chart-column.xlsx, build-chart-column.vbs). ChartRender.fs parses
  chart1.xml (barDir/categories/values/title/gapWidth/gridlines) + the drawing
  anchor and lays out an Excel-faithful column chart (nice-number Y-axis 0–35,
  gridlines, 34px bars from gapWidth=219) -> PaintItems. Plot-region MAE ~6 vs
  Excel; bar widths/positions/heights/colors/title all match. Rendered via the
  native runner `--charttest` mode.
- TODO: theme colors (currently hard-coded accent1), legends, multi-series,
  line/pie/area types, wiring charts into the normal render pipeline (read
  charts in XlsxReader -> model -> Paint), and the chart-heavy dashboards.

## Real-world goal: saopaulo-condos.xlsx

- Both sheets (Assumptions, Monthly-CF) at AA-level parity via the global fixes
  (gridlines off, indent, centerContinuous, number formats, MS Calibri). Aligned
  MAE ~10 (sheet 1) / ~14 (sheet 2, denser). No model-specific layout diff.
- The sheet-2 conditional format (E11 `cellIs greaterThan $E$10`) does NOT fire
  (E11=3.92 < E10=6), so no visible CF highlight in Excel or ours — but it
  surfaced a latent bug: our `cellIs` parsed the threshold as a numeric constant
  only, so a cell reference (`$E$10`) never evaluated. Fixed (`resolveOperand`
  resolves single-cell refs) + `synthetic-cf-cellref.xlsx` verifies a triggering
  case (12/20 > C1=10 → pink, matching Excel).
- Status: at render parity.

## Real-world goal: jumeirah-hotel.xlsx (full parity)

- One real, model-specific bug found + fixed: a hard line break `\n` in cell
  G68 ("Uses of \nFunds:") with wrapText OFF was treated as a line break,
  doubling row 68 (31.5pt vs Excel 15.75pt) and drifting the section below.
  Fixed (commit 52e80cb): wrapping/`\n` only apply when wrapText is on; the
  break is stripped on a single line otherwise.
- Everything else already matched via the lyric-derived global fixes
  (showGridLines off, centerContinuous, borders/fills/number-formats, font
  bytes = MS Calibri). Verified region-by-region against Excel screenshots
  (top, sensitivity tables, financing, sources/uses, market data, assumptions
  detail, property pro-forma, far-right FY21/FY22) — content matches.
- Row heights: an Excel COM dump (probe-rowheights.vbs) shows every row is
  15.75pt except the title (row 2 = 18.75pt); ours computes identical heights
  for every row. No cumulative drift — an earlier band-spacing "163 vs 168"
  reading was ±3px detection noise; deeper-region MAE is denser number rows
  (more AA pixels), not misalignment.
- Residual: ClearType-vs-grayscale text antialiasing only.
- Status: at render parity (commit 52e80cb).

## Real-world goal: lyric-multifamily.xlsx (full parity)

- Sheet 1 (Lyric-PF-DCF) visible region: ~94% match (MAE ~6%), dominated by
  ClearType-vs-grayscale text AA; navy bands, yellow inputs, gray sub-headers,
  number formats, borders all match.
- [x] Comp-Sales: hidden columns (D/E `hidden="1"`) no longer paint their text
  (was overflowing onto visible columns) and no longer inflate row auto-height
  (the hidden "733 Summit Avenue" was wrapping to 1 glyph/line). `Paint.fs`
  skips 0-width/height cells; native-runner auto-grow skips 0-width columns.
  Verified vs Excel Comp-Sales screenshot.
- [x] Sheet 1 Operating Assumptions FY headers (`"FY"yy`): date-format quoted
  literals were lowercased ("fy18"); `excelToNetFormat` now preserves literal
  case ("FY18"). Verified vs Excel (scrolled to row 50).
- [x] Apt-Comps, Replacement-Cost, and sheet-1 lower/far-right regions (cash
  flows, wrapped "Stabilized Year:" header) compared and match Excel at the
  AA-tolerance level — no further structural diffs found.
- Residual: sub-pixel ClearType-vs-grayscale text antialiasing (Excel
  screenshots use subpixel AA; we use grayscale) — within the goal's "a couple
  pixels off here and there" tolerance, not a structural/content bug.
- Status: renders at Excel parity within tolerance (commits 3d38033, d88d0a2).

## P0: Print/PDF Fidelity

### Header/Footer Runs

Status: partially fixed.

Current evidence:

- `HeaderFooter.resolve` now emits styled runs inside each left/center/right
  section.
- `&B` / `&I` toggles, `&"Font,Style"` face/style changes, `&NN` size
  changes, and `&Kxxxxxx` color changes are preserved through PDF render.
- `synthetic-header-footer-runs.xlsx` exercises seven style runs in one
  center header; the generated PDF crop shows bold, red, italic, 16 pt, and
  regular 10 pt segments in one section.
- Excel/COM capture through `wordvm` is back online for this fixture;
  `tools/excel-parity/excel-print-pdf.sh` generated the real Excel reference
  PDF/crop.
- Header baseline positioning now accounts for the largest run size in the
  header band. The mixed-run fixture's header dark-pixel top now matches
  Excel at y=32 in the 96 dpi crop.
- Footer baseline positioning now also accounts for large runs. The
  large-footer fixture's footer dark-pixel bounds now match Excel at
  y=129..150 in the 96 dpi bottom crop.

Checks to add:

- [x] Synthetic workbook with mixed H/F runs in one section:
  `tests/test-documents-excel/synthetic-header-footer-runs.xlsx`.
- [x] Synthetic workbook with H/F color (`&K`) and font size changes.
- [x] PDF crop generated for the mixed-run fixture and added as a
  row in `docs/excel-parity-fixes.html`.
- [x] Replace the semantic reference crop with a real Excel PDF crop from
  `wordvm`.
- [x] PDF crop comparison for header vertical placement.
- [x] Synthetic fixture with large footer runs:
  `tests/test-documents-excel/synthetic-header-footer-runs-footer.xlsx`.

Likely files:

- `src/Spreadsheet.Pdf/HeaderFooter.fs`
- `src/Spreadsheet.Pdf/Pdf.fs`
- `tests/print/helpers.fsx`
- `tools/excel-parity/compare-print-pdfs.py`

### Print Geometry And Chrome

Status: partially checked.

Current evidence:

- Print headings and print gridlines are implemented.
- The latest chrome fixture renders all four PDF pages with expected header
  variants and top-left heading corner.
- PDF cell text now treats `PaintText.FontSize` as points for drawing and
  uses the 96/72 conversion only for metrics/line-height. This improved
  `synthetic-print.xlsx` page-1 top-crop MAE from `7.0589` to `6.3969`.
- PDF print-heading text now uses 10 pt sizing, centers row numbers, and gets
  a heading-specific baseline shift. `synthetic-print-chrome.xlsx` page-1
  top-crop MAE improved from `12.4314` to `11.6033`.
- `&D` / `&T` header-footer fields now use Excel-like default Windows output
  (`M/d/yyyy`, `h:mm AM/PM`) instead of ISO strings.
- RESOLVED (2026-05-20): the body-row pitch now matches Excel. PDF text
  extraction (`pdftotext -bbox`) shows Excel and ours both emit a `14.52 pt`
  body-row pitch on pages 1 and 4; every data-row text baseline aligns within
  `0.094 pt` and the row-boundary gridlines within `1 px` (both ~`19.4 px`/row
  at 96 dpi). Page-4 body MAE is now `8.51` (was `9.4584`) and the
  column/row-heading chrome band is pixel-identical (`MAE 0.0`). The earlier
  `15 pt` discrepancy was closed by the Phase 8 print-row-pitch work; the
  rejected global-scale retune is no longer relevant. (The original note,
  measuring a `15 pt` emit, is kept here for history.)
- Remaining chrome diff is dominated by cross-renderer text rasterization
  weight: in the contrast diff every glyph reads as a ghosted/doubled stroke
  (Excel's PDF text engine vs ours), which is not a layout bug and is the
  reason dense data-row crops carry the highest residual MAE.
- FIXED (2026-05-20): the cumulative column-width drift (our rendered columns
  ~2 pt narrower, every cell right of col A creeping left). Root cause was
  `Pdf.fs withPrintColumnWidths` scaling only explicit `<col>` widths by the
  print MDW (15/14); the workbook-default width was left at the floored screen
  MDW. Now all positive widths scale. Verified on `synthetic-diagonal-borders`
  (col-D label 10.2 pt → 0.06 pt) with 11/11 print fixtures still within MAE.

Checks to add:

- [x] Quantified PDF-to-PDF crop summary for `synthetic-print.xlsx`.
- [x] Quantified PDF-to-PDF crop summary for
  `synthetic-print-chrome.xlsx`.
- [x] Gallery row for PDF print text sizing, date fields, and print-heading
  row-number alignment.
- [ ] Turn the current crop summary into enforced thresholds.
- [ ] Border/gridline weight fixture with explicit thin/medium/thick/double
  cell borders in print.
- [x] Center-on-page support in the PDF renderer
  (`PageSetup.Options.HorizontalCentered` / `VerticalCentered`). Each page
  shifts by half the unfilled printable area when the option is set; gated
  on the flag so existing `horizontalCentered="0"` fixtures are unchanged
  (verified: re-running `compare-print-pdfs.py` on the four core fixtures
  leaves every crop's MAE within sub-pixel noise of the prior run).
- [x] Center-on-page fixture: `synthetic-print-centered.xlsx`
  (= `synthetic-print.xlsx` with `printOptions horizontalCentered="1"
  verticalCentered="1"`). Excel reference PDF captured via
  `tools/excel-parity/excel-print-pdf.sh` on `wordvm`. Page 1 body MAE vs
  Excel = `5.81` (residual is content-extent measurement, not centering).
  Gallery row added at the top of `docs/excel-parity-fixes.html`.

Likely files:

- `src/Spreadsheet.Pdf/Pdf.fs`
- `src/Spreadsheet.Pdf/Pagination.fs`
- `tools/spreadsheet-pdf-runner/Program.fs`

## P0: Browser XLSX Import

Status: partially fixed.

Current evidence:

- `synthetic-browser-styles.xlsx` is an Excel-authored fixture with visible
  styles on empty cells (`B1:B7`), larger font, double underline, and diagonal
  borders.
- Before the fix, browser/Fable import stayed on the initial blank sheet for
  this class of workbook (`cells=0` in the capture harness).
- The Office.Format browser API now returns JS-array sheet names and merged
  ranges, flattened cell style objects, comments, effective row heights, and
  effective column widths.
- The browser importer now preserves style-only cells and imports fill,
  font, border, diagonal, double-underline, strikethrough, comment, and
  list-dropdown fields.
- The browser importer now imports inline/shared rich-text runs into
  `SpreadsheetModel.RichRuns`; the browser portable canvas path installs a
  text measurer and paints run-level color, size, bold, italic, underline,
  and strikethrough.
- The Fable output no longer loses the multi-sheet workbook loader to a
  same-name `SetWorkbook` overload; the table-context setter is now
  `SetWorkbookContext`.
- `tools/excel-parity/capture-browser-xlsx.mjs` renders the focused fixture
  through the browser portable canvas path and reports `cells=26`.
- `tools/excel-parity/capture-browser-xlsx.mjs` renders
  `synthetic-inline-rich.xlsx` through the same path and reports
  `cells=3 richRuns=2`.

Checks to add:

- [x] Excel-authored focused fixture:
  `tests/test-documents-excel/synthetic-browser-styles.xlsx`.
- [x] Browser screenshot capture for the focused fixture.
- [x] Gallery row with real Excel screenshot, browser after, and browser
  before images.
- [x] Browser rich-text import/render capture:
  `tests/test-documents-excel/synthetic-inline-rich.xlsx`.
- [ ] Full browser/Fable screenshot pass over imported XLSX fixtures.
- [ ] Browser import coverage for comments, list validations, merged cells,
  and conditional formatting in one fixture.

Likely files:

- `src/Office.Formats/src/Main.fs`
- `src/Spreadsheet/XlsxImporter.fs`
- `src/Spreadsheet/Spreadsheet.fs`
- `src/Spreadsheet/Formula.fs`
- `tools/excel-parity/capture-browser-xlsx.mjs`

## P0: Harness And Gallery Hygiene

Status: partially checked.

Current evidence:

- The gallery HTML is not large: about 191 KB, 37 body rows, 113 images.
- Viewport and row-level Playwright screenshots work.
- Full-page Playwright `page.screenshot({ fullPage: true })` crashes headless
  Chromium in this environment on the rendered gallery surface; this is a
  capture-path problem, not evidence that the HTML is too large.
- `tools/word-parity/capture-fixes-gallery.mjs` now relaunches before viewport
  fallback, instead of reusing a page after a renderer crash.

Checks to add:

- [ ] Dedicated Excel gallery capture script with row-level output.
- [ ] Script that validates every gallery row has Excel / After / Before
  images when the row claims them.
- [ ] Script that writes a small manifest for generated before/after assets.
- [ ] Rule for not adding a row until both before and after images exist.

Likely files:

- `docs/excel-parity-fixes.html`
- `tools/word-parity/capture-fixes-gallery.mjs`
- `tools/excel-parity/validate-fixes-gallery.mjs`

## P1: Tables

Status: open.

Known gaps:

- Every built-in style (Light1-21, Medium1-28, Dark1-11) is drawn from Excel's own
  definition (`src/Spreadsheet/TableStyles.fs`, from ECMA-376 `presetTableStyles.xml`);
  custom styles from `<tableStyles>` of styles.xml. Elements are laid over each other in
  Excel's order, each element's lines on the edges of the block it covers.
- Not compared with Excel yet (the VM was unreachable): `synthetic-table-styles.xlsx`
  is the fixture to compare.
- Tints are 1 off in a channel against Excel (floating-point HLS, `Xlsx.applyColorTint`).
- Filter button rendering is native PNG only; PDF and Fable/web paths ignore it.

Checks to add:

- [x] Light/Dark table style synthetic fixture (`tests/test-documents-excel/synthetic-table-styles.xlsx`,
  `tools/excel-parity/make-table-styles-fixture.py`; `tests/formula/stage54-tablestyles.fsx`).
- [x] Table with explicit cell styles overriding table style (stage48: a cell's own fill wins,
  black text is the default font's).
- [ ] Excel's PDF of `synthetic-table-styles.xlsx` against ours.
- [ ] Filter button crop in native PNG, PDF, and Fable/web.

Likely files:

- `tools/excel-native-runner/Program.fs`
- `src/Spreadsheet/Paint.fs`
- `src/Spreadsheet.Pdf/Pdf.fs`

## P1: Rich Text And Wrapping

Status: open.

Known gaps:

- Single-line rich shared strings render with per-run formatting.
- Browser/Fable now imports and renders single-line inline/shared rich runs in
  the portable canvas path.
- Multi-line wrap with per-run formatting still falls back to plain rendering.
- CJK fallback is implemented, but wrapped rich text needs per-run line
  chunking.

Checks to add:

- [x] Browser/Fable capture for single-line inline rich strings:
  `synthetic-inline-rich.xlsx`.
- [ ] Rich text with wrap across multiple lines.
- [ ] Rich text with CJK fallback inside wrapped lines.
- [ ] Native PNG and PDF parity crops.

Likely files:

- `src/Spreadsheet/Paint.fs`
- `spreadsheet-native/SpreadsheetGrid.fs`
- `src/Spreadsheet.Pdf/Pdf.fs`

## P1: Merged Cells And Borders

Status: open.

Risk:

- Merged geometry, border ownership, and vertical alignment tend to expose
  differences between Excel and our paint stream.
- Need stress cases for merged empty cells, diagonal borders, and thick/double
  borders across merge boundaries.

Checks to add:

- [x] Merged header cell with centered text and explicit borders:
  `synthetic-merged-text.xlsx`. VERIFIED PASSING (no fix needed) — horizontal
  centering across the full merge and top/center/bottom vertical alignment over
  the merge height already match Excel, since the merge primary carries the full
  merge rect. Kept as a regression fixture.
- [x] Empty merged cells with borders on all sides: `synthetic-diagonal-borders.xlsx`
  `E6:F7`. Fixed merge-perimeter border sourcing in `Paint.fs buildCellPaintItems`
  (the box now traces the full merge instead of only the primary cell's edges).
- [x] Diagonal borders inside merged and unmerged cells: `synthetic-diagonal-borders.xlsx`
  diagonal down / up / both render correctly (native PNG + PDF).
- [x] Border stroke crispness + gridlines-under-borders. Three coupled fixes,
  verified vs Excel screenshot on `synthetic-diagonal-borders.xlsx`:
  1. **Pixel-snap** (`GridLine` rendering, native + canvas): round width; centre
     odd widths on n+0.5, even on n. A 3 px border centred on an integer
     straddled the grid → 2 black rows + two 50%-grey edges ("half the line is
     lighter"); now 3 solid-black rows. Butt caps (Square AA-blobbed corners).
  2. **Gridline z-order + suppression** (shared `Paint.fs buildCellPaintItems`):
     the always-on grey `#a6a6a6` cell outline is now per-edge `GridLine`s drawn
     before borders, and skipped on any edge where this cell or its neighbour
     has an explicit border. Fixes (a) neighbour gridline overdrawing a border,
     and (b) grey bleeding through a `double` border's 1 px gap (gap pixel
     166-grey → 245-white). `double` now reads as two crisp lines with a white
     gap.
  3. **Print filter** (`Pdf.fs isInteractiveCellStroke`): now also drops
     `#a6a6a6` `GridLine`s (not just `StrokeRect`) so the screen grid still
     doesn't print. 11/11 print-PDF within MAE. Fable build clean.
- [x] Native PNG and PDF parity crops for the border fixture.

Likely files:

- `src/Spreadsheet/Layout.fs`
- `src/Spreadsheet/Paint.fs`
- `src/Spreadsheet.Pdf/Pdf.fs`

## P1: Charts, Images, And Shapes

Status: open.

Risk:

- Real workbooks often include charts, logos, and simple shapes.
- The parser has chart metadata, but visual rendering is not first-class.

Checks to add:

- [ ] Workbook with one embedded image anchored to cells.
- [ ] Workbook with a simple bar/line chart.
- [ ] Placeholder geometry parity before full chart rendering.

Likely files:

- `src/Office.Formats/src/XlsxReader.fs`
- `tools/excel-native-runner/Program.fs`
- `src/Spreadsheet/Paint.fs`

## P2: Number Formats

Status: open.

Risk:

- General/date/common custom formats are much better than before, but Excel's
  format language is deep.
- Remaining gaps likely include conditional sections, color sections, elapsed
  time, locale/accounting variants, and text sections.

Checks to add:

- [x] Hidden/empty sections + `@` text format on numbers:
  `synthetic-numfmt-hide.xlsx`. Empty section renders "" (`0.00;;` hides
  zeros/negatives, `0;;` hides zeros); `@` on a number shows its General text
  and `alignFromFormat` left-aligns it (Excel treats it as text). Engineering
  notation (`##0.0E+0`) still TODO (no native .NET support).
- [x] Trailing-comma scaling: `#,##0,` (thousands), `#,##0.0,`, `0.0,,`
  (millions), `#,##0.0,,"M"`. `synthetic-numfmt-scale.xlsx`. `CellText.fs`
  `trailingScale` strips end-of-section commas and divides by 1000ⁿ (.NET only
  scales a comma left of the decimal). 24/24 vs an Excel probe incl. rounding.
- [x] Custom format with positive/negative/zero/text + CONDITIONAL sections:
  `synthetic-numfmt-sections.xlsx` (built by `build-numfmt-sections-fixture.vbs`),
  ground truth from `probe-numfmt-sections.vbs` (72 pairs). `CellText.fs`
  `sectionForValue` now evaluates `[op N]` conditions (first match wins, else the
  last unconditioned default); a negative on a default section gets an auto "−";
  the date/time sniffer strips colour/condition/quoted content so `[Red]`(→d) /
  `"High"`(→h) no longer mis-route to the date path. 72/72 text vs Excel.
- [x] Color section fixture: covered by the same `synthetic-numfmt-sections.xlsx`
  (`[Red]`/`[Blue]` by sign and by condition). Colour tag is now matched anywhere
  among a section's leading tags (`[<0][Red]`, `[Red][>100]`). Verified vs the
  Excel screenshot (red negatives, blue conditionals, including value 0 → blue).
- [x] Elapsed time fixture (`[h]:mm:ss`, `[m]`, `[s]`): `synthetic-elapsed-time.xlsx`
  (built by `tools/excel-parity/build-elapsed-fixture.vbs`). All 120
  (serial × format) pairs match Excel's `cell.Text` per
  `tools/excel-parity/probe-elapsed.vbs`; implemented as `formatElapsedTime`
  in `src/Spreadsheet/Paint.fs`. Whole-second rounding for plain `h:mm:ss`
  fixed alongside.
- [ ] Accounting and locale code fixture.

Likely files:

- `src/Spreadsheet/Paint.fs`
- `src/Office.Formats/src/XlsxReader.fs`

## P2: Formula And Recalc

Status: open.

Risk:

- Cached values hide many formula gaps on import.
- Editing and uncalculated workbooks need broader formula coverage.

Checks to add:

- [ ] Structured references.
- [ ] Array formulas / dynamic arrays.
- [ ] External references behavior.
- [ ] Volatile functions.
- [ ] Cross-sheet dirty propagation stress fixture.

Likely files:

- `src/Spreadsheet/Formula.fs`
- `src/Spreadsheet/Calc.fs`
- `tests/formula/`

## P2: Browser/Fable Parity

Status: partially fixed.

Known gaps:

- Some older fixes are native PNG only; the browser import harness now covers
  the full 85-workbook XLSX fixture corpus as a smoke pass.
- Fable/web currently does not load per-family font files the same way native
  rendering does.
- Layout-cell export for browser parity is planned but not wired.

Checks to add:

- [x] Browser/Fable capture for an Excel-authored style-only-cell fixture:
  `synthetic-browser-styles.xlsx`.
- [x] Browser/Fable capture for inline rich text:
  `synthetic-inline-rich.xlsx`.
- [x] Browser/Fable capture for imported `pandas-test_converters.xlsx`.
- [x] Browser/Fable capture for a non-Calibri fixture (`poi-47813.xlsx`,
  Verdana 10).
- [ ] Layout-cell JSON parity for representative fixtures.

Likely files:

- `src/Spreadsheet/Spreadsheet.fs`
- `src/Spreadsheet/Paint.fs`
- Browser runner tooling.

## Render fixes landed (2026-05-20)

- Elapsed-time number formats (`[h]`/`[m]`/`[s]`): `formatElapsedTime` in
  `CellText.fs`; `synthetic-elapsed-time.xlsx`; 120/120 vs Excel.
- Number-format engine extracted from `Paint.fs` into `CellText.fs`
  (behaviour-preserving; paint now only paints resolved elements).
- Merged-cell box borders trace the full merge perimeter
  (`Paint.fs buildCellPaintItems`, merge-aware edge sourcing);
  `synthetic-diagonal-borders.xlsx`.
- Print column width: `Pdf.fs withPrintColumnWidths` now scales ALL columns by
  15/14 (the non-floored print MDW), not just explicit `<col>` widths. Default
  columns were left at 64 px while Excel prints ~68.6 px, drifting every cell
  right of column A leftward (column-D label was 10.2 pt off on
  `synthetic-diagonal-borders.xlsx`; now 0.06 pt). 11/11 print fixtures within
  MAE — no regression. This is the column drift previously flagged as the open
  residual in "Print Geometry And Chrome".
- Cell-text rotation in the native/interactive grid (`SpreadsheetGrid.fs`).
  Direction + vertical match Excel (CCW rises, CW descends; `rot 90` at y 88–122
  vs Excel 88–121; row height honours `ht=80 customHeight`).
- Rotated-text HORIZONTAL anchor — RULE DERIVED + IMPLEMENTED (native). Built
  `tools/excel-parity/build-rotation-sweep.vbs`: 13 angles (orientation −90..90
  step 15) × 3 text lengths → `synthetic-rotation-sweep.xlsx`, screenshotted and
  measured. Pattern (offset of text centre from the 89 px column centre, "Hi"):
  CCW angled 15..75 ≈ −30 (left); CCW vertical 90 ≈ +32 (right); CW angled
  105..165 ≈ +30 (right); CW vertical 180 ≈ −34 (left); flat 0 left (General).
  → **Rule: right-align the rotated bounding box when 90 ≤ textRotation < 180,
  else left-align; always bottom-align.** Implemented in `SpreadsheetGrid.fs`
  (rotate about the run anchor, then translate so the rotated bbox is L/R + bottom
  aligned). Verified on `synthetic-text-rotation.xlsx`: rot45 −24/Excel −10 (L),
  rot90 +34/+37 (R), rot−45 +22/+10 (R), rot−90 −35/−39 (L) — all four sides now
  match (magnitudes within ~14 px; small diff = bbox metrics vs Excel).
- Rotation placement now applied in ALL THREE render paths:
  - native (`SpreadsheetGrid.fs`) — verified vs Excel screenshot.
  - PDF (`Pdf.fs`) — anchor computed in CSS px then converted; verified by
    pdf↔pdf vs Excel's print PDF (`synthetic-text-rotation.xlsx` top-crop
    MAE 4.85, the four rotated runs overlay Excel's).
  - web `<canvas>` (`Spreadsheet.fs`) — previously did NOT rotate at all; now
    rotates via a new `canvasTextRotated` (save/translate/rotate/fillText) with
    the same L/R-bbox + bottom rule. Fable build clean. (Browser-pixel
    verification still TODO, but it shares the derived rule with native/PDF.)

## Next Suggested Work

1. DONE (2026-05-20): the `synthetic-print-chrome.xlsx` row-pitch diff is
   resolved — body pitch now matches Excel at `14.52 pt` and page-4 body MAE
   dropped to `8.51`. See "Print Geometry And Chrome" above. The column-width
   drift flagged here is now also FIXED (`withPrintColumnWidths` scales all
   columns, 2026-05-20); the remaining print residual is cross-renderer text
   rasterization weight, which is not a layout bug.
2. Expand the browser/Fable capture pass from `synthetic-browser-styles.xlsx`
   to a small imported-XLSX corpus now that the loader path is working.
3. Turn `tools/excel-parity/compare-print-pdfs.py` output into enforced
   thresholds for the core print fixtures.
4. Add enforced thresholds for the large header/footer run fixtures.
5. Separate table-style overlay from baked cell format
   (`tableHeaderOverlay` currently bakes Bold=true into the cell's format;
   Excel COM keeps it as a render-time overlay). Fixing it would let the
   VBA-layout harness compare bold/italic/fill cleanly on table headers
   and stripes.

## VBA-Layout Harness Status (2026-05-19)

`compare-vba-layout.py` joins Excel COM TSVs (extracted on the wordvm via
`extract-excel-layout.vbs`) with our `spreadsheet-vba-layout-tsv.sh`
output by `(sheet, address)` and diffs every per-cell field. Current
state across the eight reference fixtures:

| fixture | excel rows | our rows | remaining diffs |
| --- | --- | --- | --- |
| `pandas-blank` | 0 | 0 | clean |
| `pandas-test1` | 96 | 96 | clean |
| `pandas-test_decimal` | 30 | 30 | 6 bold (table-style overlay) |
| `pandas-test_none_type` | 8 | 8 | clean |
| `pandas-dimension_large` | 13 | 13 | clean |
| `pandas-multiindex_no_index_names` | 17 | 17 | clean |
| `phpsp-26template` | 16 | 16 | clean |
| `synthetic-print` | 186 | 186 | clean |

Fixes that landed via the harness this round:
- OOXML §18.3.1.13 column-width formula: stored `<col width>` and
  `defaultColWidth` are encoded with the 5-px padding baked in
  (`pixels = stored * MDW`); only the implicit chars-based default keeps
  `chars * MDW + 5`. baseColWidth → `chars * MDW + 8`.
- `applyXxx` flags default to "apply" when absent (Excel's behavior) —
  fixes `pandas-test_none_type`'s `@` text-format cells inheriting `General`.
- Font-derived row auto-height when `customHeight="false"`: Calibri
  lookup table (11pt→15, 12pt→15.75, 18pt→23.25, etc.) drives per-row
  height. Fixes `pandas-dimension_large` cumulative y-drift entirely.
- General-format trim formula: `(cellWidth - 5) / MDW` (no extra MDW
  buffer). Fixes the 1-digit-short General representations on
  `pandas-test1` (e.g. `0.98026851`).
- Hash-overflow budget: `(cellWidth - 5) / MDW` for the `#` count and a
  char-count-vs-measure guard on the fit check so date cells render as
  `#####` whenever Excel does (regardless of HarfBuzz's narrower glyph
  measurement).
- `spreadsheet-vba-layout` now walks every sheet and skips
  formatted-but-empty cells to match `extract-excel-layout.vbs`'s
  `IsEmptyCell` filter, eliminating spurious `ours-only` rows.
