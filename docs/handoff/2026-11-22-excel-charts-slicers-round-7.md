# Handoff: Excel charts and slicers, round 7 (2026-11-22)

Branch `worktree-agent-a8e5c2927c8806b7e`. Goal: fix the Excel 2019 discrepancies that earlier rounds left, using the
parity method (find the rule, test it on generated fixtures in Excel 2019 on the VM, implement it, add a test and a tracker
row). The tracker section for this round is "Slicer text, cameras, Pareto, sunburst labels, filled maps against Excel 2019,
2026-11-22" in docs/excel-parity-tracker.md.

## Items

1. **Slicer caption and button cut ("Centr...")**: done.
   - A slicer's text is in the theme's body font, measured by GDI widths at 15 px.
   - A button's text stays whole if it is no wider than the button less 18 px. Otherwise Excel keeps as many characters as
     it can (spaces included) and adds "...", provided the result is narrower than that room. If even "..." does not fit,
     it shows as many stops as do fit (two, one or none).
   - The caption follows the same rule, with a room of the slicer width less 72 px.
   - Buttons start 9 px in at the left and end 9.67 px in at the right. Columns are 3 px apart, and each button sits on
     whole pixels, rounded half down.
   - Evidence: `build-slicer-cut.vbs` made 400 slicers, 1 px apart in width, in Calibri and in Tw Cen MT. That gives
     1600 texts, and each one matches Excel's ink. The e-commerce dashboard's "Centr..." also matches.
   - A table slicer on a different sheet from its table now shows its buttons in the browser.
2. **Camera text narrower than Excel's**: done.
   - Cause: Lato is an Office cloud font. On the VM it is under `%LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts`. It was
     missing from our fonts, so the browser drew a fallback.
   - Fix part 1: Lato was copied into `artifacts/fonts/windows-extra` (local and git-ignored, like Aptos before it).
   - Fix part 2: Lato was added to the GDI advance tables and to FontMetricsData. The measurements came from
     `probe-font-faces.vbs`, which also covers Tw Cen MT; that font had been measured as Arial before.
   - A camera's cell text is now cut at the next non-empty cell, as the grid does.
   - `[Green]` in a number format is #00FF00.
   - Evidence: `build-missing-font.vbs` pictures and the dashboard's KPI YOY boxes.
3. **Pareto 6 px wider**: done.
   - A chartEx chart with a second value axis keeps 12 pt at its sides, not 6.
   - Value labels stand 6 pt from the plot, not 7.5. This also moved every histogram and box chart onto Excel's left edge.
   - Evidence: `build-pareto-layout.vbs` gives Excel plots starting at 36, 42, 49 … 67 px and ending at 428. Ours are
     within 1 px.
4. **Sunburst label breaks**: done.
   - Labels break to the ring width, measured by design advances.
   - A label is left out when its lines plus 4.5 pt are longer than the part's arc at its middle radius.
   - Evidence: `build-sunburst-labels.vbs` plus Sunburst3 in synthetic-chartex-looks.
5. **Map charts (regionMap)**: done, with a first approximation of the layout.
   - Maps are now drawn from the part's `cx:geoCache`. That cache is base64 of a raw deflate stream, decoded by the new
     `Inflate.fs`. Region outlines use Bing's point compression.
   - The United States uses an Albers projection with Alaska and Hawaii as insets. Other countries use Mercator.
   - Colours: if the part has `cx:valueColors`, its min/mid/max scale is used. Otherwise Excel's two colours are used:
     80 % lighter to 75 % of the colour style's first colour.
   - Regions without data are #E0E0E0. Outlines take the series' line colour.
   - The legend and the provider credits are drawn.
   - Colours match Excel's pictures exactly on both dashboards. The layout constants come from a single picture.
   - The 8 MB geography cache is cut out unread at parse time and only decoded when the map is drawn.
   - Not checked: Excel drawing the map offline. The VM was online and shared, so I did not cut it off.

## Not confirmed / next

- Narrow slicers: header buttons (Excel drops Multi-Select below about 47 px).
- A timeline's caption cut.
- A camera's text in a Lato row sits 4 px lower than Excel's.
- Pareto charts with labels of 4 or more figures are 1 px off.
- Sunburst labels:
  - In the 376 x 228 pt chart, Excel leaves out labels that it breaks in the 360 x 216 one. Not understood.
  - Excel draws a label black on light parts; ours stays white.
- Maps:
  - The lower 48 states are about 4 % narrower than Excel's; Australia is wider than Excel's.
  - Legend value text size and bar width match one picture only.
  - A three-colour legend shows only the largest and middle values, as seen in one picture.
  - No synthetic map fixture: Excel needs Bing to make one. The tests use the two dashboard workbooks that are already in
    tests/real-world.

## Added

- **Probes and tools** in tools/excel-parity:
  - `build-slicer-cut.vbs`
  - `build-missing-font.vbs`
  - `build-pareto-layout.vbs`
  - `build-sunburst-labels.vbs`
  - `font-metrics/probe-font-faces.vbs`
  - `make-gdi-advances.py` now adds Lato and takes `RB_FONTS_DIR`.
  - `make-font-metrics.py` now adds Lato and Tw Cen MT descents.
- **Fixtures** in tests/test-documents-excel (all made by Excel through the probes):
  - synthetic-slicer-cut.xlsx and synthetic-slicer-cut-twcen.xlsx
  - synthetic-missing-font.xlsx and synthetic-missing-font-twcen.xlsx
  - synthetic-pareto-layout.xlsx
  - synthetic-sunburst-labels.xlsx
- **Tests**:
  - stage56-slicers: cut texts and button geometry.
  - stage52-charts: Pareto and histogram spans, sunburst labels.
  - stage55-chartex: maps, point decoding, inflate.
  - New xlsx-browser checks: `camera-excel-parity` and `map-charts-excel-parity`.
  - Extended xlsx-browser check: `slicers-in-grid`.
  - Updated xlsx-browser checks: `lazy-sheets` (the map reads 'Map Pivot') and `sheet-pdf-default-fonts` (Carlito-Bold
    is fetched for headers).
- **Code**: Slicers.fs, Charts.fs (`RegionMapPlot`, `layoutRegionMapChartEx`, `decodeMapRing`, Pareto and sunburst
  changes), Inflate.fs, CellText.fs, FontMetrics.fs and its data, GdiAdvances.fs, Spreadsheet.fs (camera clipping, table
  slicer redraw).

## Gates

- `pnpm run build:fable`: passes.
- xlsx-browser, full run: two failures. Both were fixed afterwards and the two checks rerun in isolation; both pass. The
  full run was not repeated.
  - `lazy-sheets`: the map now reads its sheet.
  - `sheet-pdf-default-fonts`: the expectation changed now that captions use the theme font.
- `tests/run-all.sh --core`: was still running when I stopped. Every stage finished so far had 0 failures (through
  xlsx-parts); the last stage, xlsx-patched-corpus, had not finished.

## Known breakage

None known.

## VM

My probes all quit their own Excel instances. `shoot-sheets.ps1` and my VBS scripts call Quit; none timed out. At stop
there were 5 EXCEL.EXE processes in the console session. They are not mine as far as I can tell, so I left them running.
