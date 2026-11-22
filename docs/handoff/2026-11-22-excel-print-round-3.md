# Handoff: print against Excel 2019's PDF, round 3 (2026-11-22)

Branch `worktree-agent-a4770b7e47e13fe02` (fast-forwarded to master 6364de84 first). Stopped on request; no
product code changed yet. Only the fixture generator and the print probe were changed (tools only).

## Items

1. AutoFit columns (`bestFit="1"`) printing wider - **in progress** (mechanism found, width model not yet).
2. Painter: icon sets as Excel's 16-px icons, sparklines in PNG / PDF - **not started** (Excel's PDFs taken, see below).
3. Hairline / all border styles in print - **not started** (fixture made, Excel's PDF taken).
4. Segoe UI text 1 px low, a workbook's right-aligned numbers 4.6 pt right, a blank last page Excel prints - **not started**.
5. Data-validation arrow painted on every list cell in PNG / PDF - **not started** (Excel's PDF taken; the code to
   drop is Paint.fs "Pass 6", nothing tests it).

## 1. AutoFit columns: what is confirmed

Confirmed on Excel 2019 (the probe now writes each column's `ColumnWidth` / `Width` before and after the export):

- Laying out pages (Page Break Preview, the export) **changes the sheet's own column widths**: a `bestFit="1"`
  column with a number in it is re-fitted, and `ColumnWidth` reports the new width afterwards (an earlier probe
  that only opened the file saw the stored widths: 0.58 for a width-1 column; after Page Break Preview, 1.57).
  The print then prints that screen width by the usual d / m rule. So the rule is "re-fit on pagination",
  measured in screen px, not a print-only width.
- Text-only and empty `bestFit` columns are left alone; text cells do not count in a column that has numbers
  (TwoCells: a long text under 12345 leaves the column at 12345's fit).
- The new width is the widest number of the column (TwoCells: 12345 and 1234567 -> 1234567's fit), at least the old
  width plus 1 px (WideCols: "1" in columns stored 8 / 10 / 12 / 15 -> 57 / 71 / 85 / 106 px, ours 56 / 70 / 84 /
  105; "$1,234,567 " -> 88, 88, 88, 106). (bestfit-figures' "Wide" earlier showed no +1 px: not explained.)
- A General number is fitted by the count of its integer figures only (0.5, 0.125, 1.25 fit as "1"; decimals can
  be dropped): round(k a) + 11 px, a = 8.33 for 10-11 pt, 10 for 12 pt, 6.67 for 9 pt, 11.67 for 16 pt
  (multiples of 5/3 px), whatever the face (Calibri 10 = Arial 10 = Calibri 11 = Aptos Narrow 11 in px).
- A formatted number is fitted by its shown text measured in a font that is **not** the cell's face: Calibri 11
  gives "1" 16 px but "5" / "7" / "9" 19, "11111" 42 vs "55555" 53, "M" wider than "W", bold "$1,234,567 " 109 vs
  regular 88. GDI's own Calibri / Arial (regular and bold, every ppem 5-31, scaled linearly) do not fit
  (`fit.py`: ~2 px off on average). Calibri 20 / Arial 8 behave as the real face (digits tabular).

Hypotheses not confirmed (next steps):
- A fixed fallback face picked by point size (10 and 11 pt alike, a bitmap-like set of sizes) measures the text;
  candidates to test: Small Fonts (GdiAdvances.smallFonts has its advances by em; its "1" is narrower), MS Sans
  Serif. Fit `fit.py`'s per-character estimates (in Calibri-11 px: 1 6.4, 0 8.0, other digits 8.5, i 3.6,
  . 3.6, , 3.3, % 11.5, space 3.0, M 14.8, W 13.4, - 4.8; constant ~10.5) against those faces' advances.
- Which fonts / sizes take the fallback and which the real face (real: Arial 8, Arial 14, Calibri 20, Segoe UI
  Black 11; fallback: Calibri 9-16, Arial 10 / 12, Times 12, Verdana 10, Tahoma 10, Aptos Narrow 11).
- Where to implement once found: at pagination (Pagination.paginatePrinted / the PDF path) and arguably in the grid
  as Excel shows the new width after printing; for parity of print only, the print path.

Data (scratchpad, not checked in):
`/tmp/claude-1000/-home-dave-Dev-rbt-office/c821990b-f07a-46bb-a142-dfe16abbb344/scratchpad/bf3/xl/`
(synthetic-print-pages-bestfit-chars{,-calibri12,-calibri20,-arial10,-arial8}: Excel's PDFs and tsv with col /
colafter), `.../bf3/gdi-out.tsv` (GDI extents, `gdiext.ps1`), `.../bf3/fit.py`; earlier rounds' PDFs in
`.../pq/bf`, `.../pq/bf2` (`pq/bfw.sh STEM DIR` gives a page's column widths in dots).

## Excel PDFs taken for items 2, 3, 5 (not yet compared)

`/tmp/claude-1000/-home-dave-Dev-rbt-office/c821990b-f07a-46bb-a142-dfe16abbb344/scratchpad/bf4/xl/`:
synthetic-icon-sets-1.pdf, synthetic-sparkline-looks-1.pdf, synthetic-print-pages-validation-1.pdf;
`.../bf3/xl/synthetic-print-pages-borders-1.pdf`. Plan noted for item 2: draw icons from the same shapes as
`ConditionalFormatting.iconSvg` as PaintItems (VecPath) at 2 px / 16 px as the grid; sparklines as PDF overlays
(`Paint.drawingOverlays`-like, from a values resolver - the grid's `sparklineIn` / `sparklineValues` in
Spreadsheet.fs would move to Sparklines.fs to be shared by the grid, ExportPdf and the PDF runner).

## Tools and fixtures added (uncommitted fixtures: generated, not checked in)

- `tools/excel-parity/probe-print-pages.vbs`: writes `col` / `colafter` (index, ColumnWidth, Width) for the used
  columns and one more, before and after each sheet's export. Consumers ignore unknown keys.
- `tools/excel-parity/make-print-pages-fixture.py`: sets `bestfit-chars` (+ `-calibri20`, `-arial10`, `-arial8`,
  `-calibri12`: literal characters after a number, percentages, digits, decimals, other cell fonts, wide columns, two
  numbers in a column), `borders` (every border style as a box and as a bottom line, no gridlines), `validation`
  (a list validation's cells). No *.excel.tsv checked in for them (helpers.fsx "print pages:" reads every
  synthetic-print-pages-*.excel.tsv and would fail on bestfit-chars until the rule is implemented).

## Gates

None run (no product code changed). Nothing known broken on the branch.

## VM

My probes quit their own Excel. Seven instances of mine left by a broken fixture (Open failed) were stopped by PID.
At hand-off the VM has EXCEL.EXE 10016, 6680, 9804, 8364 (started 00:04-00:05 VM time, another agent's) and 6052
(01:28 VM time, not started by any probe of mine as far as I can tell) - left running. Note: a probe whose
`Workbooks.Open` fails leaves its Excel running (probe-print-pages.vbs has no error path there).
