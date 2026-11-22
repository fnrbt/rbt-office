# Excel Parity Tracker

Last updated: 2026-11-22 UTC.

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


## Charts and grid chrome against Excel 2019, 2026-11-21

The 15 charts of the real-estate models (A.CRE's unit-split pies, the house-flipping doughnuts, the investment-property
calculator's pie, oyabayo's five column charts and two pies, vertex42's mortgage balance over a date axis and its
interest-rate scatter) against Excel 2019 on the VM (`shoot-sheets.ps1 -KeepZoom`, `chart-probe.vbs`: each chart's
axes' scales, plot area, labels, legend from COM, and `Chart.Export`), ours from `shoot-browser-sheets.mjs` (which
now really scrolls to the cell asked for). Each rule then confirmed on charts made in Excel through COM (sweeps of
axis lengths and fonts) and on synthetic-chart-parity.xlsx (`make-chart-fixtures.py`, 12 charts, checked with
`chart-probe.vbs` and pictures). Checks: tests/formula/stage52-charts.fsx "Excel 2019:", xlsx-browser
`charts-excel-parity`.

| Rule (Excel 2019) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| A value axis: zero an end unless the values' spread is at most a sixth of the largest (then the low end is the lowest less half the spread); 5 % of the spread beyond the highest (below the lowest under zero); the step the smallest 1, 2, 5 x 10^n giving no more intervals (ends whole steps) than N; zeros only 0 .. 1; no values as values of 0 .. 1 | 125 COM charts of 25 ranges x 5 heights (50 .. 60 is 44 .. 62 by 2, 100 .. 103 98.5 .. 103.5, -5 .. 20 -10 .. 25), fixture ColScale / ColClose / ColZero / Skip | 50 .. 60 as 50 .. 61; 0 .. 1.1 for no values; the step from the range alone | `Charts.valueScale` |
| N up the side = floor(plot area's height / (1.2207 x the labels' size)) - 2, at most 10; along the bottom floor(width / (3.8 x size)) | 12,000 COM charts of fonts 8 - 24 pt, every 2 pt of height, then every 0.25 pt about the steps: exact for every font; 22,000 bar charts for widths | always up to 10 intervals | `Charts.axisIntervals` (ColShort 0 .. 10 by 5, BarWide 0 .. 300,000 by 100,000) |
| The plot area's box 11 pt in from the chart's edges; a title or legend 6 pt in, 11 pt from the plot; legend entries a line apart and 5.85 pt, the box 3 pt round them, a side legend centred on the plot's box; a legend box of the file's own size shares its height out evenly; `c:legendEntry` text looks | layout probe of 4 kinds x 4 sizes x 8 title / legend variants (COM rectangles; the chart-area origin from a picture), PieLegend entries 50.55 pt apart in 151.65 | 6 - 8 px margins, entries bunched | classic layout |
| A date axis (`c:dateAx` over dates): a band per month (`baseTimeUnit`) from the first date to the last, a label every `majorUnit` `majorTimeUnit` from the first, in its format, turned by `bodyPr rot` | COM: 45292 .. 56219 for vertex42's 360 months; DateLine | the dates as 1,561 text categories: the line in a quarter of the plot | `dated` remap in the classic layout |
| A scatter's x values that are text, or a range of two columns and rows, are 1, 2, 3 ...; the x axis to the last point drawn, in the x values' format | ScatterText 0 .. 150 by 50; vertex42 0 .. 400 | x of 6,000,000 "%" | `Charts.xValuesOf` |
| A pie's label without a separator is a line a part, broken at spaces to a fifth of the chart's width; best fit inside the slice when it holds it, else just outside; moved by its manual layout from the series' place; the pie 1.5625 x the labels' size smaller each side for labels that may stand outside; a point's own `c:dLbl` | COM label rectangles of PieLabels / PieOut / A.CRE's two pies ("Two\nBR/One\nBA\n14%", 4 lines) | "Two BR/Two BA; 14%" in one line, cut at the chart's edge | pie labels |
| Chart text in its typeface (`a:latin`), Segoe UI Black as Segoe UI bold where the weight is missing; measured with GDI's widths | pictures | all Calibri | `PFace`, `Charts.faceWidth` (`ExcelText`) |
| A line of w EMU is round(w / 9525) px, at least 1 (0.25 - 1 pt 1 px, 1.5 pt 2, 2.25 pt 3, 3 pt 4, 4.5 pt 6); a series stating no line: 1.5 pt under chart styles 1 - 8, 3.5 pt under 25 - 32, else 2.5 pt | synthetic-chart-lines "Widths" (`Chart.Export` pixels), "Plain2" / "Plain34" / "PlainNone" / "ScatterPlain", `chart-style-probe.vbs` (48 styles through COM) | 0.75 px floor, unrounded; 2.25 px for a series stating none | `Charts.linePixels` |
| Legend keys 26 px long wherever a line is in the legend (a column's key then a box that long, a swatch high), as wide as their series' line, the marker in the middle; else a square swatch | "Widths" / "Keys9" / "Keys14" (26 px at 9 and 14 pt; COM's LegendKey.Width says 12 pt and is not what is drawn) | 2.25 px lines of 1.44 em, no marker | `KeyLook` |
| A slice's outline is its point's own look (`c:dPt` `c:spPr`): its `a:ln`, none where it has none; else the series'; stating none, the style's (a pixel of white for 9 - 16, of a shade for 33 - 40, else none) | "PiePoints" (1 pt red, none, a look without a line: none - not the series' 3 pt black -, 4.5 pt white), "PiePlain", style probe | the series' line, else a white pixel | `PointLines` |
| Category labels along the bottom: every s-th, s the least for which the labels stand in s bands - in a line when the widest is no wider than s bands, else broken at spaces when the widest word is, else (the angle left to Excel: no `rot`, or -60000000) turned -45 when s bands are at least sqrt 2 lines (1.2207 em), -90 when a line; `c:tickLblSkip` sets s | `chart-label-probe.vbs`: 13,400 column charts (labels 1, 2 ..., "Cat 1", "Category 1", "Q1 2000", mixed; 6 - 100 categories; 8 - 11 pt; 90 - 760 pt): spacing 97 % (2,752 of 2,754 at 9 pt in 2 pt steps), the angle from the axis' height; "LabelsWrap" / "Labels45" / "Labels90" / "LabelsSkip" | every third of oyabayo's 60 (Excel every second); rot -60000000 drawn turned -1000 deg. | `categoryPlan` |
| Clustered bars: a bar is a band / (series - (series - 1) x overlap + gap) - the gap a share of one bar | dashboard-sales-analysis "Sales By Manager": 7.5 px of a 29.6 px band, 2 series, gap 182 % | band / (1 + gap) / series | bars |
| The primary axes are the first value axis written that a plot uses (not the first plot's) | dashboard-ecommerce-sales' area of profit before its columns of sales: $ 400K at the left | the profit's | `primaryValueAxis` |
| A pivot chart (`c:pivotSource`) whose range holds an outer field's rows (years among months, subtotals) is drawn from the file's points; its series named as cached ("Total", not the cell's "Sum of Sales"); multi-level categories (`c:multiLvlStrCache`) are the file's levels: the outer a row each below the points' labels (a turned column beside bars), centred on its group, lines between groups | "Sales By Chain" / "By Manager", "Top 5", phpsp-32's years and quarters | 23 rows of years and months, values shifted; 24 quarters | `cellsOf`, `Outer` |
| Chart text in the theme's fonts: `+mn-lt` / `+mj-lt`, and text naming no face in the minor font | sales-analysis in Garamond, e-commerce in Tw Cen MT | Calibri | `parseTheme`, `faceOf` |
| A display-units label alone stands at the far end of its axis (the top of one up the side, the right of one along the bottom) | "Thousands" of the sales-analysis charts | the axis' middle | axis titles |
| In a legend box of the file's size a too-long entry is broken at its spaces (its key by the first line) | e-commerce "Most Sold Items": "GBC Premium / Transparent / Covers" | cut with an ellipsis | `drawEntryIn` |
| A pie with labels that may stand outside (end, best fit - a pie's default), a doughnut with any labels, is an eighth of the chart's shorter side smaller, whatever the labels' size | `chart-pie-probe.vbs`: 153 pt for 178 in 300 x 200, 196.75 for 228 in 400 x 250 and 250 x 300, labels 8, 10, 14 pt; e-commerce "Pie Chart" 165 pt | 1.5625 x the labels' size each side, doughnuts never | pie |
| A marker of `c:size` n is n pt across | chart-zoo scatter: 7 px for 5 | 1.1 n px | markers |

Grid chrome (synthetic-error-checks.xlsx, `make-error-check-fixtures.py`, `error-probe.vbs`; xlsx-browser
`grid-chrome-excel-parity`):

| Rule (Excel 2019) | Evidence | Now |
| --- | --- | --- |
| Frozen panes: a screen pixel of #9B9B9B along each frozen edge, across the window and the headers, on the grid line after the last frozen row / column | choudhary's and the house-flipping model's pictures | `.spreadsheet-freeze-line` in the frozen corner's pane |
| A green triangle (#008000, 6 px legs, inside the grid lines) at the top left of a cell a default check finds: a number stored as text ("1,000", " 12", "12%", "$5", "(7)", "0012" - not a date), a text date of a two-digit year, a formula's error but #N/A, a formula unlike the alike formulas above and below it (else either side of it; none reading another sheet), a total in line after its range that leaves out the number just before it, a formula in an unlocked cell (not on a protected sheet); not where `<ignoredErrors>` says so | the fixture's 17 marks, none on its "Ignored" sheet; Excel's `Range.Errors` over rows 1-80, columns A-AD of the 65 sheets of the 12 models: 5,700 marks, all ours, none more | `ErrorChecks.markedCells` |

Left: best fit placement of pie labels is an approximation (inside at 0.6 of the radius when the slice holds the
label); -45 against -90 near sqrt 2 lines is wrong for 9 % of the probe's turned axes (Excel turns -45 at less room
when it leaves labels out: 14.4 pt at every second, 12.5 at every third); category label spacing at bands under a
pixel is a label or two short of Excel's; the error triangle is 6 px whatever the zoom; a table's inconsistent
calculated column, data-validation and misleading number format checks are not run; the axis rules were measured on
Calibri labels (other faces use Calibri's line height: the sales-analysis Garamond axis 0 .. 3000 by 1000, Excel's
by 500). From the corpus: chart user shapes (`c:userShapes`, the e-commerce doughnut's "Most Sold Items") are not
drawn; rich titles of several runs are one run (phpsp-32); a legend box's border; a doughnut beside a legend of manual
layout is smaller than Excel's; Tw Cen MT (an Office font) is not among the fonts, so the e-commerce charts measure
and draw in a fallback; Excel's own default marker sequence and grey marker outlines; area and line charts on ticks
leave no room for the last label's half (chart-zoo "Area" plot 7 px wider than Excel's); chartEx and chart sheets
were not compared (chart-probe reads ChartObjects only).

## Charts: faces, shapes, looks against Excel 2019, 2026-11-22

What the two rounds above left. Each rule found on Excel's charts of generated workbooks on the VM (`probe-on-vm.sh`
runs a probe in Excel's desktop session: `chart-axis-face-probe.vbs`, `chart-probe.vbs`, `chart-marker-probe.vbs`,
`chart-marker-style-probe.vbs`, `chart-label-probe.vbs`), the new fixture synthetic-chart-looks.xlsx
(`make-chart-fixtures.py`, 18 charts) and the dashboards; checks: tests/formula/stage52-charts.fsx "Excel 2019:",
xlsx-browser `charts-excel-parity`.

| Rule (Excel 2019) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| A value axis up the side: N = floor(plot area's height / a line of its labels' face) - 2, the line GDI's: tmHeight + tmExternalLeading (OS/2 win ascent + descent, or the hhea line when longer): Calibri 1.2207 em, Garamond 1.125, Arial / Times 1.1499, Segoe UI 1.3301, Georgia 1.1362, Tw Cen MT 1.0889 | 15,700 column charts of 10 faces x 9, 10, 12 pt x heights 90 - 420 pt x tops 60 - 100: 15,697 as Excel (Calibri's line for all: 96 %) | Calibri's line for every face: sales-analysis' Garamond axis 0 .. 3000 by 1000 | `lineIn` (`GdiAdvances.lineExtents`, `make-gdi-advances.py`); 0 .. 3000 by 500 |
| Along the bottom: one interval per 7.5 widths of the face's widest digit (Calibri 0.5068 em: the old 3.8 x size; Garamond 0.4688, Verdana 0.6357) | 6,651 bar charts of the same faces: 6,649 | 3.8 x size whatever the face (Verdana 88 %) | `axisIntervals`, `ExcelText.digitEm` |
| Chart text is spaced by its face's line (titles, legend entries, labels) | as above | Calibri's | `lineFor` |
| A chart's shapes (`c:userShapes`: the drawing part of its `chartUserShapes` relationship) are drawn over it: `cdr:relSizeAnchor` from a corner to a corner in fractions of the chart, `cdr:absSizeAnchor` from a corner by its EMU size; a linked text box shows its cell | picture of "UserShapes" (two text boxes, one linked; a box of its own size, filled and outlined); the sales-analysis and e-commerce doughnuts' labels | not drawn | `withUserShapes`, `elementItems` |
| A title or an axis title of runs of their own looks: each run in its size, colour, weight, slant, the line as high as its largest | picture of "RichTitle" ("Sales" 18 pt red bold, "by" 12 pt italic, "Region" 18 pt blue; a turned axis title of two runs); phpsp-32 | one run in the first run's look | `TitleRuns`, `drawRuns` |
| A legend's own fill and outline (`c:legend` `c:spPr`) round its entries; the plot area's outline (`c:plotArea` `c:spPr` `a:ln`) | "LegendRight" (legend 312.8, 96.2, 37.2 x 50.6 pt), "LegendBottom" | neither drawn | `Legend.Fill` / `Border`, `PlotBorder` |
| Rounded corners (`c:roundedCorners` on, or not written) 10 px round, the chart's fill too | pictures of "RoundedBare" (grey frame, no element) and "LegendRight" | 8 px, the fill square | frame |
| Markers a series does not name (a line chart with `c:marker val=1`, a scatter of markers; `c:symbol auto`): diamond, square, triangle, x, star, circle, plus, dash in turn; a size the series does not state the style's: 6 pt for styles 1 - 8, 10 pt for 25 - 32, else 8 pt; filled and outlined in the series' colour (25 - 32 with a shadow, 41 - 48 lighter: not drawn) | "Markers" / "Markers34" / "Markers42" / "MarkersXY" pictures; `MarkerSize` under every style 1 - 48 through COM | circles of 5 pt | `markerOf` |
| A line or scatter series that does not write `c:smooth` is smoothed | the same charts: curves, `Series.Smooth` True | straight | `seriesOf`; charts the editor makes write `smooth val="0"`, as Excel does |
| The seventh to twelfth automatic colours are the accents lighter (HSL lightness x 0.6 + 0.4: Excel's #8FA2D4 / #F1A78A, ours #8FAADC / #F4B183) | "Markers" series 7, 8 | darker | `autoColor` |
| On the ticks (`crossBetween midCat`, no bars), the first and the last category labels are centred on the plot's ends with half of each inside the plot area's box; their spacing is the points' (n - 1 gaps) | "AreaTicks" / "LineTicks": inside 31.22 .. 326.46 pt (box 7 .. 345; halves of "January 2024", "June 2024"), labels in a line; chart-zoo "Area" 259.72 pt wide (ours 259.9, was 265) | the plot to the box's ends, labels cut and broken | `tickHalves`, `planFor` |
| Turned category labels have the plot's width shared among the labels shown - every s-th of n, ceil(n / s) labels -: -90 from a line each, -45 from sqrt 2 lines, decided before the -45 overhang narrows the plot | `chart-label-probe.vbs` (kind 3, 24 and 60 categories, 9 pt, 90 - 760 pt by 2): 665 of 672 as Excel (every 12th of 60 in 60.56 pt, not every 11th: 6 labels of 10.09 pt) | every s-th by s bands: 653 | `categoryPlan` |
| A shape's text run that names no typeface is in the theme's minor font (lines broken by its widths) | Excel's picture of the e-commerce doughnut: "Most Sold Items" in Tw Cen MT, 2 lines | Calibri, 3 lines | `shapeItems` (ours still 3 lines: "Most Sold" 61 px in a box of 60 - our group's scale) |
| Tw Cen MT (an Office font: the e-commerce dashboard's theme) is measured with GDI's hinted advances and drawn in its face | picture of "FaceTwCenMT" | a fallback | `GdiAdvances` (`fetch-windows-fonts.sh` fetches TCM*.TTF into artifacts/fonts/windows-extra) |

Left: the remaining 7 turned-label cases are our plot 0.5 pt wider than Excel's (the side margins); the -45 / -90 and
spacing rules for flat and broken labels still use s bands (not measured at sub-pixel bands); markers' shadows (styles
25 - 32) and the lighter markers of 41 - 48 are not drawn; series 13+ colours not measured; underline, strike and
baseline offsets of title runs not drawn; a picture in a chart's shapes (relative to the shapes' own part) is not drawn;
legend entries sit 1.5 pt further in than Excel's in a boxed legend; chartEx (waterfall, funnel, treemap, histogram) and
chart sheets were not compared again (the probes read ChartObjects; the VM's Excel is stopped by other runs that kill
it by name, so long probes are run in short pieces).

## chartEx, chart sheets, colours and legends against Excel 2019, 2026-11-22

The newer chart kinds (chartEx: histogram, Pareto, box and whisker, treemap, sunburst; waterfall and funnel again),
chart sheets, and what the round above left. Fixtures made by Excel itself through COM on the VM
(`build-chartex-zoo.vbs` -> synthetic-chartex-zoo.xlsx, 25 charts; `make-chartex-looks.py` -> synthetic-chartex-looks.xlsx,
the same with what COM cannot set written into the parts; `chartsheet-probe.ps1` -> synthetic-chartsheets.xlsx;
`make-chart-fixtures.py` -> synthetic-chart-colors.xlsx), each checked on Excel's pictures (Chart.Export, the window
through PrintWindow) and COM (`chartex-treemap-probe.vbs`: the points' rectangles of 160 treemaps; `legacy-chart-probe.vbs`).
Ours drawn with `render-charts.fsx` + `shoot-svgs.mjs`. Checks: stage52 "Excel 2019: chartEx", "chart sheet",
"automatic colours", "shape's text"; stage55; xlsx-browser `charts-excel-parity`.

| Rule (Excel 2019) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| A histogram's bins: from the least value (or the underflow), Scott's width - 3.5 x the sample standard deviation / n^(1/3) - to two significant digits, as many as reach the greatest; a set width or count; bins closed on the right `[a, b]`, `(a, b]` (`intervalClosed="l"`: `[a, b)` .. `[a, b]`), under- and overflow bins `≤ 100` / `> 800`, the last regular bin ending at the overflow; text categories a bar each (their values' sum) | 8 data sets of 5 - 200 values (100 values: 12.654 -> 13, `[11.87, 24.87]` .. `(76.87, 89.87]`; 183.1 -> 180; 0.304 -> 0.3; 5 values one bin `[2, 97]`), "HistCount", "HistWidth", "Hist2" | placeholder | `histogramBins` |
| A Pareto chart: the bins from the most to the least, their running share a 2.25 pt line of the second accent on a 0 - 100 % axis at the right | "Pareto", "ParetoBins" (`(44.2, 61.2]` first) | placeholder | `classicOfChartEx` |
| Box and whisker: quartiles QUARTILE.EXC's (`exclusive`) or .INC's; whiskers to the farthest values within 1.5 x the box; the mean's x, the points inside and outside (what a part that says nothing shows); the box and its lines the series' colour 80 % (#365B9D); a category's boxes side by side in the band less its gap (gapWidth 1: half a band) | "BoxCat", "BoxTwo" / "BoxInclusive" (1 .. 9, 30: 2.75 .. 8.25 / 3.25 .. 7.75, COM's boxes 0.818 the height) | placeholder | `boxFiguresOf`, `BoxesPlot` |
| chartEx layout: the plot's box 6 pt in, 8 pt below the title (4.5 pt from the top), value labels 7.5 pt from their axis, category labels 6.5 pt below; text #595959, the title 14 pt not bold, labels 9 pt; axis lines and gridlines #D9D9D9, the frame #D9D9D9; a title that says nothing "Chart Title"; bars of a histogram touching, a pixel of white between | pictures of the zoo (Hist5: plot 28 .. 471 px, top gridline 42, labels 269 px) | classic margins (11 pt) | `ChartEx` layout constants |
| A treemap: a strip at a time along the shorter side (a column where the rectangle is at least as wide as high), as many in it as bring the strip's smallest rectangle's width / height nearer 1 - not squarify's longer / shorter; each node's children in its rectangle, largest first; colours by the outer level's first row (the colours part: accents, then 60 % darker - the seventh #264478 -, ...); 1.5 pt white round each | `chartex-treemap-probe.vbs`: 160 of 160 flat treemaps (40 sets x 4 sizes; squarify's rule 128); the zoo's three treemaps | placeholder | `squarify`, `chartExColor` |
| A treemap's labels: a tile's at its bottom left, the outer level's name at its top left over its children (`overlapping`) or in a grey banner #A6A6A6 a line high above them (`banner`); "A, 50" for a name and a value | "Treemap", "Treemap3", "TreemapFlat" of synthetic-chartex-looks | - | `layoutHierarchyChartEx` |
| A sunburst: a hole and rings each the radius / (levels + 1), from twelve o'clock clockwise, largest first, children within their parent; labels along their part's middle radius, turned over on the left half, broken to the ring's width | "Sunburst" / "Sunburst3" / "SunburstFlat" (121 px: holes of 30, 40, 61) | placeholder | `layoutHierarchyChartEx` |
| A waterfall's legend: squares 5.25 pt, text 5.25 pt after, 6.5 pt between; the plot 80 .. 257 px of 288 | "WaterNeg" (gridlines 80, 116, 151, 186, 222, 257) | legend 3 px high, plot 77 .. 260 | `layoutPlottedChartEx` |
| A chart sheet: no cells; a grey window #F6F6F6; the page less its margins (Letter landscape, a chart sheet's default; 0.7 / 0.75 in) fitted to the window at a whole per cent when `zoomToFit` (Excel's default) - the page's fit, not the chart's -, else `zoomScale`; the page in the middle of the window, the chart in the middle of the page | `chartsheet-probe.ps1`: Letter landscape 117 %, portrait 86 %, A4 121 %, margins of 2 in 149 % in 1898 x 787 px; pandas-chartsheet 117 %; the chart at 415 px, 5 px down | the chart over an empty grid at 100 % | `chartSheetView`, `chartSheetPlacement`, `.spreadsheet-chartsheet` |
| A chart of Excel's old API (`Charts.Add`, `ChartObjects.Add`: a part naming a chart style, no look of its own): title 18 pt bold, text 10 pt, gridlines, axes and frame #898989 | `legacy-chart-probe.vbs` (COM: 18, 10, #898989), the chart sheets' pictures | 10 pt title, black lines | parse defaults (`plain`) |
| More than six series that state no colour: each round of six a shade or tint in linear light of the accents, round p of n series (p - q / 2) x 1.4 / (q + 2) from them, q = n / 6 whole | synthetic-chart-colors (7, 12, 13, 18, 19, 24, 25, 30, 36 series through COM: 7 -> 0.767 shades then tints #3B64AD / #8FA2D4; 12 -> 0.65 shades then the accents), each within a unit of 255 | the first six the accents, then lighter | `autoColorOf` |
| Text in a shape and a legend's entries laid out by the face's design advances (no hinting) | e-commerce "Most Sold" 59.7 px (GDI 61) on a line in Excel's box of 60 px; "LegendBottom" 172 .. 309 px (ours 174 .. 305, was 166 .. 313) | GDI's hinted widths | `ExcelText.idealWidth` |
| A legend's entries 5.85 pt apart in a row, square keys 0.55 of the text's size (4.94 pt at 9 pt), a side legend's entries 1.5 pt further in than its padding | COM's LegendKey / LegendEntry of "LegendRight" / "LegendBottom", pictures (keys 7 px in from the box) | 8 pt apart, 0.6, 5 px in | legend |

Left: a Pareto chart's plot is about 6 px wider than Excel's (Excel leaves more room at both sides when there is a second
axis); a box chart of one category draws its gridlines to the chart's right edge in Excel (ours stop at the plot); the box
chart's points are 5 px circles (not measured); sunburst labels that do not fit their part are left out (Excel draws them
broken or clipped: "Osaka" fits in Excel, ours "Osak / a"); a treemap's banner text 2 px low; map charts stay
placeholders; the chartEx layout constants were measured at 360 x 216 pt only. A chart sheet's chart is the file's size
(the drawing's `absoluteAnchor`; Excel recomputes it from the printer: 682.05 x 495.5 pt for 682.4 x 493.9); Excel's
selection handles are not drawn; a chart sheet prints as before (Pagination is another branch's). Flat and broken category
labels at bands under a pixel, marker shadows (styles 25 - 32) and the lighter markers of 41 - 48 were not taken up.

## Cell text at a zoom, 2026-11-21

The browser grid's cell text against Excel 2019 at 50-200 % (`ExcelText`, `GdiAdvances`, `CellText.NumberFit`).
Found on the real-estate models (Excel's zoomed text 3-13 % wider than ours: other wrapping, overflow and `####`);
each rule below confirmed on generated cells in Excel on the VM (`shoot-sheets.ps1 -KeepZoom` of a column of
texts between black marker cells, 10 fonts x 13 zooms and 18 more cases; ink extents measured; `Range.Text` of
numbers in 7 column widths x 12 zooms). Fixture `synthetic-text-zoom.xlsx`
(`tools/excel-parity/make-text-zoom-fixture.py`, checked in Excel), xlsx-browser `grid-text-excel-parity`.

| Rule (Excel 2019, GDI) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| Text is drawn at a whole pixel size: round(points x 96/72 x zoom), a half up | 80 % and 85 % of Calibri 11 alike (12 px), 67 % and 70 % alike; Arial 10 at 100 % 13 px (twenty zeros 140 px); Arial 13.5 at 75 % (13.5 px) as 14 px | the browser's size (14.67 px at 100 %, any zoom continuous) | `ExcelText.pixelSize`; the cell's CSS size is that over the zoom |
| From 11 px, every character as wide as GDI hints it: a whole pixel, the design width rounded or a pixel either side | the widths of 10 and 20 repeats of `0` and `n` in 10 fonts x 13 zooms give each advance: FreeType's v35 interpreter has all 465 TrueType cases (Calibri `0` 7 px at 15 px, `n` 7 px at 12 px) | the browser's advances (Calibri 11 at 80 %: a line 9 % narrower) | `GdiAdvances` (generated, `make-gdi-advances.py`: design advances and hinted differences at 11-36 px of 56 faces of the Office fonts); other fonts: the browser's advance a character, rounded. The browser draws its own glyphs, spaced (`letter-spacing`) to Excel's width over the line; no kerning |
| Under 8 points on screen (points x zoom, to a twentieth: see the re-survey below), whatever the font, Windows' raster Small Fonts: its face as high as the pixel size or the next lower (10 px: the 120-dpi 6-pt face; 10.5 px: the 11-px face); bold a pixel wider a character | 185 cases of Calibri, Arial, Times, Verdana at 50-90 % and Arial 6-7.5 at 100 % agree within 1 px (a bitmap font in the shots: every font as wide); Calibri 8 at 100 % and Arial 10 at 80 % (8 pt) are TrueType | Calibri 11 at 67 %: a line 12 % narrower | Small Fonts' advances (from `smalle.fon` / `smallf.fon`) in `GdiAdvances`; drawn in a wide sans serif made for small sizes as Small Fonts is (Verdana, else DejaVu Sans), spaced to that width (a bitmap font is not drawn) |
| A line of text is set in 2 px from the cell's left side and 3 px from its right (the grid line at its right is the cell's), x the zoom | left-, right- and centre-aligned lines of 18 cases at 50-200 % | 5 px each side | style sheet `padding: 3px 3px 3px 2px`; an indent adds to that side; a text overflows when wider than the cell less 5 px |
| A number at a zoom is fitted to the column on screen (its 100 % pixels x the zoom, rounded): the room is the cell less 5 px x the zoom (superseded below: less a margin of the font's digit); a formatted number fits when its text is no wider; a General number takes the longest form whose characters, each a digit wide (a minus sign as itself), fit; else as many `#` as fit in the cell less 5 px (x the zoom from 100 %), `#` measured | `Range.Text` at the sheet's zoom: 816 of 840 cells agree (the rest at 50-70 %, in Small Fonts) | 644 (fitted at 100 %'s digit width whatever the zoom) | `CellText.NumberFit`, `cellDisplayTextFitted`, `Paint.cellLookFitted` (the browser grid only: other callers as before) |

Left: General numbers at 50-70 % (Small Fonts) are fitted a little differently by Excel (24 of 840 cells); in Small
Fonts the text is drawn 1 px nearer the cell's left side than Excel's bitmap glyphs; the hinted widths come from
FreeType's GDI-like interpreter, not GDI (all 465 measured cases agree, larger sizes not measured past 32 px);
the grid's cell edges at a zoom are rounded where they fall (a column of 75 px at 90 % is 67 or 68 px on screen,
Excel's always 68); autofit still measures with the browser's widths at 100 %.

## Real-estate re-survey: text, wrapping, fonts, conditional formats, 2026-11-22

Every sheet of the 12 real-estate models (59 sheets) shot in Excel 2019 (`shoot-sheets.ps1 -KeepZoom`) and in the
browser grid (`shoot-browser-sheets.mjs`), stacked and pixel-diffed. Ranked by what still differed: (1)
acre-property-acquisition's six sheets, all in Segoe UI Black: other `####` than Excel's; (2) oyabayo's sheet of Aptos
Narrow 14 at 57 % drawn in Small Fonts where Excel draws TrueType, so other widths, `####` and wraps; (3) wrapped
headers broken at another word (acre-apartment's "Releasing Cost / per Unit" on 7 sheets); (4) numbers' room for
wide fonts and at zooms other than 100 %; (5) conditional formats' "cell value" rules comparing within 1e-9. Each rule
below confirmed on generated workbooks in Excel on the VM, then implemented.

| Rule (Excel 2019) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| A wrapped text breaks where the next word would make the line wider than the cell on screen less 5 px, the line measured in Excel's widths at the zoom (Small Fonts below 8 points) | synthetic-wrap-budget: one text in columns a pixel apart (Calibri 11 and bold, Arial 10 and 11 at 100 %, Calibri at 85 % and 70 %): bold Calibri 11 "Releasing Cost" (90 px) one line from a 95-px column, not 94; the same rule at every zoom | the browser wrapped the text with the letter spacing of the whole text: other breaks | `excelWrapLines` in the grid; each line a block of its own spaced to Excel's width (not justified / distributed, rich or turned text); xlsx-browser `wrap-excel-parity` |
| Small Fonts under 8 points on screen, the points x zoom taken to the nearest twentieth of a point | synthetic-small-fonts: Calibri and Arial 7-16 pt at 23 zooms (569 cells; a raster font draws both faces alike): 14 pt at 57 % and 10.5 at 76 % (7.98) TrueType; 15 at 53 % (7.95), 11 at 72 % and 12 at 66 % (7.92), 10.5 at 75 % Small Fonts. A cloud font (Aptos Narrow) the same | under 7.98 exactly: oyabayo's whole sheet in Small Fonts | `ExcelText.isSmallFonts`; every case's face as Excel's, widths within 3 px of its ink |
| Segoe UI Black is drawn in GDI's hinted advances, like the other Office fonts | synthetic-font-faces "Black": $1,592,811 in columns of 40-85 px as Excel shows it at 9, 10, 11 pt; twenty zeros 140 / 160 / 180 px | not in the tables: the browser's face, whatever it had | `GdiAdvances` (seguibl / seguibli); xlsx-browser `font-faces-excel-parity` |
| A number's room in its cell is the cell on screen less 2 x ceil(d / 4) + 1 px, d the widest digit of its font as drawn (the padding a column width has for that digit); bold in Small Fonts a pixel less. `#`s fill the same room. The `_x` of a format is a space wide whatever x | synthetic-font-faces "Room*": 10 fonts (Segoe UI Black 9-14, Arial Black, Segoe UI and bold, Calibri and bold) at 50-200 %, 6210 cells' `Range.Text` all as Excel's (`synthetic-font-faces.excel.json`): Segoe UI Black 11 (d 9) 7 px, Calibri 11 5 px at 75-100 %, 7 at 125 %, 9 at 200 %; `"$"#,##0_)` and `_W` of $21,000,000 fit alike | 5 px x the zoom whatever the font | `CellText.fittedRoom`, `NumberFit.Overhang` |
| A General number none of whose forms fits as digits shows its shortest form where its own characters fit | 1234567 in Calibri 11 at 75 %: `1E+06` (28 px) in a 33-px cell | `#####` | `generalNumberFitted` |
| A conditional format's "cell value" rule compares numbers as `=` does: equal to 15 significant digits, else their order, no tolerance near 0 | synthetic-cf-compare: 20 rules, `DisplayFormat.Interior.Color` read in Excel: 0.1+0.2 = 0.3 and between 0.3 and 0.5; 1E-10 > 0 and <> 0; -1E-10 < 0; 1+22 ulp = 1, 1+23 ulp not; an operand 100.00000000000001 is 100 | within 1e-9: 1E-10 = 0, -1E-10 not < 0 | `ConditionalFormatting.numberHolds` (`Formula.compareNumbers`); stage26 `cf-compare-*` |

Found, not matched: rows Excel fits on opening wrap about 9 % earlier than it draws them (synthetic-wrap-budget
rows 6+: "Cost per Unit" bold, 81 px, one line from an 88-px column, "The quick brown" Calibri from 111, Arial 10 from
103), and depending on the zoom (85 %: 90) - the measure is not GDI's at the drawn size (vertex42's Help rows a line
higher in Excel); text paddings for wide fonts (the room is wider, the text's left edge measured the same). (The
open-time wrap is matched now: "Wrapped rows, print text, outline strips, numbers in italic" below.)

## Wrapped rows, print text, outline strips, numbers in italic against Excel 2019, 2026-11-22

The rest of the re-survey: rows of wrapped text Excel fits on opening, the PDF's cell text (the painter still wrapped at
the cell less 22 px with 4 px paddings), the outline strips beside / above the headers, oyabayo's `-100.00%` that Excel
shows as `########`. Each rule found on generated workbooks in Excel on the VM, then implemented. Fixtures and probes:
`make-wrap-fit-fixture.py` (+ `wrap-fit-cases.tsv`: synthetic-wrap-fit.xlsx, Excel's heights in
synthetic-wrap-fit.excel.tsv), `probe-wrap-fit.vbs` (heights on opening, after AutoFit at the sheet's zoom and at 100 %),
`probe-open-heights.vbs`; `make-print-text-fixture.py` (`print-excel-pdf.vbs`, the PDF's words and grid lines);
`make-outline-fixture.py` (synthetic-outline.xlsx, `shoot-sheets.ps1`); `make-number-room-fixture.py`,
`probe-text-shown.vbs` (`Range.Text` with each sheet shown at its zoom). Checks: `tests/office-formats/xlsx-bridge.fsx`
("rows: wrapped text fitted on opening", "rows: a line's trailing spaces", "numbers: italic"), `tests/print/helpers.fsx`
("print:"), xlsx-browser `chrome-outline`.

| Rule (Excel 2019) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| Fitting a row on opening, a wrapped line fits when g + max(0, floor((g - 57) / 28)) <= the column less 2 ceil(d / 4) + 3 px: g its width at 100 % in GDI's advances (whatever the sheet's zoom), d the font's widest digit at 100 % (7 px for Calibri 11 / Arial 10, 9 for Calibri 14 / Arial 16 / Cambria 12 / Segoe UI Black 11, 11 for Calibri 20 and 24). 80 px of Calibri 11 fit a column of 87 px, 88 px one of 96, 240 one of 253 | 3 sweeps, 27 fonts and sizes (Calibri 8-24 and bold, Arial 8-16, Times, Verdana, Segoe UI, Tahoma, Cambria, Segoe UI Black), 2-word and 1-word texts (runs of `n`, `i`, `0`) in columns a pixel apart, 75,000 rows: every threshold as the rule (the 28-px step: the same for every font) | the column less 6 px, the line in the browser's widths | `SheetOps.openWrapLineCount`, `fitRowsOnOpen` (native: `ExcelText` too) |
| ... a line's spaces at its end do not count, at its start they do; a text of runs is measured run by run in its runs' fonts | "Cost Unit   " one line from 63 px as "Cost Unit", "   Cost Unit" from 72; choudhary's `_(@_)` headers one line; vertex42's Help (bold runs) | the whole shown text, the cell's font | `openWrapLineCount` (per-character advances) |
| The rows Excel shows first - those that fit its window at their unfitted heights, from the top - of a sheet at a zoom are fitted at that zoom, the active sheet's all; the rest at 100 % | the same case in every row of sheets at 50, 85, 150 %: rows 1-78 / 1-46 / 1-26 differ (780 px of 20-px rows at the zoom), the active sheet's 300 | - | not followed (the zoomed measure not found: Calibri as at 100 %, Arial 10 a pixel off either way; the window's size is Excel's own) |
| Text in print (PDF) is set at 600 dpi: the font a whole number of dots high (11 pt is 92 dots, 11.04 pt in the PDF; Arial 10 9.96), each character a whole number of dots wide; a line 16/600 in (2.56 px) from either side of its cell | make-print-text-fixture.py: 4 fonts x 27 columns, left and right aligned digits 1.93 pt from the grid lines (+-0.07), the PDF's `Tf` | 4 px either side, the font's advances unrounded | `Paint.printPadX`; `Pdf.measurePdf` |
| ... a wrapped line breaks as on screen (`CellText.excelWrapLines`) with the printed column less 28/600 in (4.48 px) for room; lines past the cell's bottom are not printed | the paragraphs of the same 108 columns: 101 as Excel's PDF (before: 8) | the cell less 22 px; lines drawn over the rows below | `Paint.printWrapMargin`, the text clipped to its cell |
| The outline strips at 100 %: a column of 19 px a level and one more, 5 px to spare (62 px beside the row numbers for two levels of groups, 81 for three; 43 above the letters for one); level and +/- buttons 19 px (a +/- at the top of its summary row / column); bars 2 px wide, 6 px into their column, a 6-px tick at the end away from the summary; a 2-px dot in the middle of every grouped row / column, in the column of the level below its own, unless a button is there | pictures of synthetic-outline.xlsx ("Rows", "Cols", "Both", "Above") | 15-px columns, 13-px buttons, 1-px brackets, no dots | `renderOutline` in the grid; xlsx-browser `chrome-outline` |
| A number in italic has round(px x 0.16) px less room (2 px at 11-15 px, 3 at 16-21, 4 at 24, 5 at 29); italic in Small Fonts (made up by GDI) half its size, cut down (3 at 7 px, 4 at 8, 5 at 10), bold there a pixel more. The `#`s fill the same room; none where not one fits (the cell shows nothing) | make-number-room-fixture.py: Aptos Narrow 14, Calibri, Arial upright, italic, bold, bold italic at 57 and 100 %: 9,684 of 9,696 cells as Excel (before 8,407); italic Calibri, Arial, Times at 8-22 pt and in Small Fonts: 10,066 of 10,100 (Calibri 16 italic, 21 px, has 4) | italic as upright; at least one `#` | `ExcelText.overhang`, `CellText.fittedHashes` |

On the 12 real-estate models (every sheet, rows 1-150, Excel's heights on opening): 4,032 of 4,038 rows as Excel's
(before: 4,027 with the native measure); oyabayo's `-100.00%` (italic Aptos Narrow 14 at 57 %, a 90-px column) is
`########` as in Excel. The reader kept a no-break space at the end of a run: .NET's `Trim` took it (vertex42's "and
payment schedule" was "andpayment").

Left: the first screen of rows of a sheet at a zoom (above); vertex42's Help rows 1-2 px lower than Excel's (the lines
of bold runs, unchanged); in print, Excel prints from column A where the PDF starts at the first column with
something (synthetic print fixture: our page starts at B), and a 300-pt row prints a little lower than ours (both
matched now: "Print pages against Excel 2019's PDF"); a General
number in Small Fonts at 57 % (4 cells), bold italic Aptos Narrow at 57 % (a pixel more room than the rule); text
boxes' wrapping not measured.

## Print pages against Excel 2019's PDF, 2026-11-22

Every visible sheet of the 12 real-estate models, three real-world workbooks and the print fixtures of
tests/test-documents-excel, exported by Excel 2019 sheet by sheet (`Worksheet.ExportAsFixedFormat`), against our
pages (`Pagination.paginatePrinted`, what our PDF prints): the pages, where they break (Page Break Preview's
`HPageBreaks` / `VPageBreaks`), the print area, the scale. Then the words of the pages (`compare-pdf-words.py`).
Tools: `probe-print-pages.vbs` / `excel-print-pages.sh` (each sheet's PDF, page setup and breaks), `compare-print-pages.fsx`,
`compare-pdf-words.py`, `probe-col-widths.vbs`; fixtures `make-print-pages-fixture.py` (sets range, widths, heights,
fit, breaks, titles, rowunit-*, bestfit), checked in as synthetic-print-pages-*.xlsx with Excel's facts in
*.excel.tsv; checks: tests/print/helpers.fsx ("print pages:", "print sizes:").

Sheets paged as Excel (pages, row and column breaks): 181 of 186 - 59 real-estate sheets, 8 real-world, 13 print
fixtures, 106 synthetic (before: 113). Words of the print fixtures' pages: within 0.02 pt across (were 0.50 left).

Ranked differences found (sheets of the corpus they changed): what prints without a print area (17), how wide a
column prints (9), how high a row prints (11), the fit-to-pages scale and the page's room (6), repeat rows inside the
print area and print areas written as formulas (2, one of them no page at all), sheets with nothing to print (2),
the grid's place on the page (every page, 0.48 pt), header / footer sections (every page, 0.96 pt), columns
Excel's AutoFit left (2, not matched).

| Rule (Excel 2019) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| Without a print area a sheet prints from A1, whatever its first row and column with something | "range" StartC5 (C5:F12 prints A1:F12; StartC5Wide breaks at row 48, column J as from A1); acre's MF Rents (breaks J, Q, ... not K, R) | from the first row / column with a cell | `Pagination.usedRange` |
| ... to the last cell with a value (a formula giving "" too), a fill or a border; a font or number format alone does not print; a merged cell with a value or fill to its end; a drawing; text running over empty cells to its right to where it ends; a row's own fill (`<row s customFormat>`) down to that row and across to the last column with a format of its own, a column's own fill not at all | "range": FillPast / BorderPast to F12, BoldPast / FormatPast A1:C5, FormulaBlank, MergePast / MergeFill to H14, Overflow to F7, OverflowFar 3 pages, RowFill A1:C12, ColFill A1:C5; bramosman's Amortization formatted to row 383 prints to row 194; acre-property's Summary (text to X); choudhary's Rental Assumptions (to U) | every cell in the model | `printsCell`, `overflowEnd` |
| A sheet with nothing to print has no pages: Excel's export fails ("We didn't find anything to print"), its workbook's PDF leaves it out | "range" Empty; acre's "Calculation Modules-->" | one blank page | `paginate` (no pages), `Render.toBytesOfSheets` |
| A column prints at d / m dots of 600 dpi a pixel of its width on the screen, rounded to whole dots: m the Normal font's maximum digit width on the screen, d its widest digit in print (the font at its 600-dpi pixel size, points x 600 / 72 rounded: 47 for Calibri 11, 46 Arial 10, 51 Calibri 12, 56 Arial 12, 47 Aptos Narrow 11) | "widths": 36 columns of 0.1 to 25 characters in 5 Normal fonts, 180 columns within a dot of Excel's grid lines | 7.5 / m px a pixel for every font (Calibri 12 7 % narrow, Arial 12 15 %) | `Pagination.printColumnWidthPx`, `ExcelText.printDigitWidth`, `SpreadsheetModel.PrintDigitWidth` |
| A row prints at floor(h x R / r + 0.14) dots, h its height in pixels unrounded (points x 4/3), r the Normal font's row on the screen, R the font's row in print (121 dots Calibri 11, 103 Arial 10, 130 Calibri 12, 127 Aptos Narrow 11, 108 Tahoma 10: 32 fonts and sizes measured) | "heights": rows of 7.5 to 409.5 pt in 9 Normal fonts, 151 of 153 within a dot; "rowunit-*" (24 fonts); bramosman (Calibri 12: 15.24 pt a row, we 15.48), oyabayo / acre-property (Aptos Narrow: 15.24, we 14.52), vertex42 mortgage (Tahoma) | 19.36 px for a 20-px row, 20.64 for 21, else the screen's | `Pagination.printRowHeightPx`, `ExcelText.printRowScale` |
| ... h from the row's points (`ht`, or the sheet's own default height) while they are the pixels the row shows: a row of 20 pt shows at 26 px and prints as 26.67 px would | "heights" (10 and 20 pt rows); houseflipping's HOME (rows of 15.7 pt: Excel's break at row 38, ours 39 from the 21 px) | the pixels | `SpreadsheetModel.RowPoints` / `DefaultRowPoints`, `Pagination.exactRowPx` |
| At a scale every row and column is its printed size times the scale in whole dots, and a page is filled to a pixel short of its printable width / height | "fit": Scale90 ten 387-dot columns a page (eleven would be 4,257 of 4,260 dots), Scale55 84 rows of 67 dots | sizes unrounded, the whole page | `Pagination.paginate` (`rangesAt`) |
| Fit to pages is a whole percent: the share of a page (a pixel short) the area needs, cut down, and lower while the area takes more pages than it may | "fit": Fit10 98 % (99.07 cut to 99 fills the width), Fit11 89, Fit13 76, Fit17 58, Fit23 42, FitTall 72 on 2 pages, Fit2x2 67 (each the PDF's column pitch) | the exact share (99.4 %, 3 pages for FitTall) | `paginate` |
| Repeat rows / columns are printed above / beside a page that begins past them; a page that begins at or before them has them where they are | "titles": TitlesInside (row 20 of A1:F137: page 1 rows 1-47, then 47 with the title), TitlesTop, TitlesAbove, TitlesCols, TitlesColsInside; vertex42 mortgage (row 49 of A1:L410) | the rows before the titles were not printed at all (vertex42's first 48 rows) | `PageSlice.TitleRows` / `TitleCols`, `Render` |
| A print area written as a formula is the area it comes to | "titles" TitlesInside (`OFFSET(...,100+H1,6)`: A1:F137); vertex42 balloon and mortgage (`OFFSET(...,ROW($A$49)+1+$K$25,...)`: A1:L410) | not read (the mortgage sheet printed no page) | `Formula.evaluateFormulaArea`, the importer; the reader and writer keep the formula |
| The grid starts 4 dots (0.48 pt) right of the left margin and below the top one at any scale; centred, 4 dots short of the middle | the widths fixture's first lines at 18.48 / 54.48 pt; Fit10 at 98 % 50.88 for 50.4; synthetic-print-centered (96.98 pt past the margin for 97.44) | at the margins | `gridInsetPt` (text kept where it was) |
| A left header / footer section starts 8 dots (0.96 pt) right of the left margin, a right one ends 8 dots left of the right margin - not a section with underlined, struck, raised or lowered runs | synthetic-print, -centered, -scale, -first-page-number, header-footer-runs (51.36 pt for 50.4; 740.6 for 741.6); hf-style-codes at the margins | at the margins | `appendSectionOps` |
| A header / footer's `&Ktt+nnn` (`-nnn`) is a theme colour (index tt, tint nnn %), six characters of code | vertex42 mortgage's footer `&K01+034https://...` prints the address alone | "+034" printed, the colour code read as hex | `PrintHeaderFooter` (the colour kept: the theme is not known there) |

Found, not matched: a column Excel's AutoFit left (`bestFit="1"`) prints wider than its width - 7 dots more, or as
wide as a number in it needs ("bestfit": a 5-character column of 123456789 at 577 dots, not 235; acre's Annual Cash
Flow and choudhary's Rental break columns earlier than ours); where a cell's text stands in it in print depends on the
Normal font (a line 16 dots from the cell's left in Calibri 11 / Arial 10, 17 in Calibri 12, 18 in Arial 12, 21 in
Cambria 14; a bottom-aligned line of Calibri 11 3.58 pt above the cell's bottom, ours 2.96; a top-aligned one 11.04 pt
below its top, ours 10.47); the row of a Normal font not among the 32 measured is taken at 6.2 dots a pixel (they
run 5.55 to 7.1); a header's vertical place (0.2-0.3 pt); houseflipping's "2 REPAIR ESTIMATOR" (fit 1 x 6 of a
print area of 1,092 rows with hidden ones: 2 pages in Excel, 6 ours); dcf-stock-analysis's Financial Model and
acre-property's Summary (a page more than Excel). Words of the real-estate pages (354 pages, `compare-pdf-words.py`): 47 within 1 pt of
Excel's; the rest mostly from elsewhere - rows fitted on opening a line higher than Excel's (acre's MF sheets: a
two-line header of 28.8 pt fitted to three lines, every row below 14.7 pt lower), the PDF runner's faces (Calibri,
Arial, Cambria, Times New Roman, Verdana: Aptos Narrow, Tahoma, Arial MT are set in Calibri - oyabayo and
acre-property 4 pt across), the columns AutoFit left (above). (Most of this is matched now: "Print text and pages,
round 2" below.)

## Print text and pages against Excel 2019's PDF, round 2, 2026-11-22

What "Print pages" left: the words of the real-estate pages, AutoFit columns in print, where text stands in its cell,
the three sheets still paged otherwise. Each rule found on generated workbooks exported by Excel 2019 on the VM
(`excel-print-pages.sh`), then implemented. Fixtures: `make-print-place-fixture.py` (Place: 7 cell fonts in rows of
15 - 45 pt at the bottom, middle and top, right aligned; Pads: `_x` widths; Blank*: pages with nothing on them),
checked in as synthetic-print-pages-place / -blank; `make-print-pages-fixture.py` "filter" (synthetic-print-pages-filter),
"bestfit-sweep", "bestfit-figures-<face><size>" (15 Normal fonts, not checked in: the rule is not found). Checks:
tests/print/helpers.fsx ("print pages:" reads every synthetic-print-pages-*.excel.tsv; "print: the text's inset by the
Normal font", "a format's widths of characters", "text at the bottom of a row on one baseline", "... the top and in
the middle"). Measured with `compare-print-pages.fsx` and `compare-pdf-words.py` (a page is within 1 pt when its
paired words' median move is, across and down).

Real-estate PDFs, word by word: 306 of 370 pages within 1 pt of Excel's (before: 49 of 354 by the same count - the
previous round's 47 - with Excel's PDFs of that round; more pages compare now that they pair). Sheets paged as Excel: 183 of 186 (before 181): Annual Cash Flow, Rental (AutoFit columns) and oyabayo (below) differ.

| Rule (Excel 2019) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| The PDF runner prints the workbook as the browser has it: rows fitted on opening by Excel's rule | acre's MF sheets (Excel's row 5 30 pt; our runner's native geometry fitted it to 45 with the renderer's widths, every row below 14.7 pt lower) | `ExcelNativeRunner.buildWorkbookContext` | `XlsxBridge.Import.importWorkbook`'s model (tools/spreadsheet-pdf-runner) |
| Text is set and measured in its own face where the machine has it | Excel's PDFs embed Segoe UI Black, Aptos, Tahoma ...; acre-property's words 4 pt across in Calibri | Calibri, Arial, Cambria, Times, Verdana only | `Fonts`: a family found by the name tables (Windows family, OS/2 style) of the files in the font directories, before the known sets' prefixes; RB_WORD_FONT_DIR a PATH-separated list (windows:windows-extra); the runner embeds the faces `Pdf.requiredFaces` names |
| A row an AutoFilter hides prints nothing, as a hidden row | houseflipping's "2 REPAIR ESTIMATOR" (1,051 rows hidden in the filter's range, columns of no criteria): fit 1 x 6 at 55 % on 2 pages; "filter": FilterFit / FilterBreaks 1 page | the rows printed (6 pages at 24 %) | `Pagination.rowHeightPx` |
| Pages at the end with nothing on them are left out (no cell that prints, no text running in, no merged cell, no drawing); a blank page before one with something prints; &N counts the pages printed; a column's or row's own fill or font, gridlines, a drawing in another column of pages do not count | Blank (pages 2-3 of 4 blank, printed), BlankLast 3 of 4, BlankLastTwo 4 of 6, BlankLastGrid, -ColFill, -RowFill, -Shape 3, -ColFont 4 of 6; acre-property Summary 7, dcf Financial Model 9 | every page | `Pagination.paginate` |
| A cell's text stands ceil(d / 4) + 4 dots in from its sides, d the Normal font's widest digit in print, whatever the cell's font | the rowunit / heights fixtures' first text, 33 Normal fonts: Calibri 9 14 dots, Calibri 11 16, Calibri 14 19, Cambria 14 21, Verdana 10 18; the print-text fixture's Arial and Times cells under Calibri 11 16 | 16 dots | `Paint.printPadXOf`; all 33 to the dot |
| A format's `_x` leaves the width of x in the cell's font | Pads: `0_)` 28 dots of Calibri 11, `0_W` 82, `0_.` 23, a literal space 21; acre's MF sheets' `"$"#,##0_)` | a space | `CellText.widthPads`, Paint (within 2 dots) |
| Text at the bottom of its cell: every such cell of a row on one baseline, (the row's largest descent of that text, `FontMetrics.lineOf`, + 1) screen px above the bottom, scaled as the row (`PrintRowScale`) | Place: a row with Calibri 20 sets Calibri 11, Arial, Segoe UI, Times, Trebuchet 51.5 dots up under Calibri 11 (8.4 - 8.5 px at each Normal font's row scale: 54.5 under Segoe UI 10, 50.5 Times 12); rowunit: 22 of 24 fonts within a pixel | each cell by its own font's Windows descent, 21 - 28 dots low in such rows | `PaintText.Baseline`; within 2 dots |
| At the top and in the middle by the font's own metrics in print whatever the Normal font (at the top 4.5 dots lower), from the grid's lines | Place: Calibri 11 90.5 dots below the top, Arial 10 79.5, Segoe UI 10 91.5 under Calibri 11 and Segoe UI 10 alike | raised 4 dots with the rest of the text (`TextRaisePt`) | top within 4.5 dots, middle 1 (Segoe UI 5) |

Found, not matched:
- AutoFit columns (`bestFit="1"`): one with text alone prints as its width; one with a number prints wider - a pixel
  more on some sheets (bestfit-sweep's One, Fixed2, Dates: 85 px for 84) but not others (bestfit-figures' Wide, the
  same columns and values), or as wide as its numbers "need" in print: a General number of k figures needs
  round(k a) + 11 px, a = 8.33 px for Calibri 10 / 11, Arial 10, Tahoma 10, Verdana 10, Aptos Narrow 11, 10 for
  Calibri 12 / Arial 12 / Times 12, 6.67 Calibri 9, 11.67 Calibri 16, else the screen digit (Arial 8, Arial 14,
  Calibri 20, Segoe UI Black 11) - not the digit's width at the screen or in print, nor Excel's own AutoFit on the
  screen (General 7 px a figure there); formatted numbers need less a character than figures, and 50.0% more than
  12.5%. acre's Annual Cash Flow ("$"#,##0_) values of 7 figures, bold totals) prints its columns at 106 px for 70,
  more than any of these. Data: bestfit-sweep / bestfit-figures PDFs (generator sets checked in); left for another
  round. acre's Annual Cash Flow and choudhary's Rental page otherwise for it.
- oyabayo's last page (rows 74 - 116 x column CD, nothing on it) is in Excel's PDF though blank pages at the end are
  left out everywhere else (no fill, drawing or text running there found); ours leaves it out: 38 pages, Excel 39.
- Text placement: Segoe UI and Segoe UI Black a pixel lower than Excel at the bottom (their descent in the GDI table
  a pixel short of what Excel places them by): acre-property's words 0.7 pt low; at the top Trebuchet 1 px.
- oyabayo's right-aligned numbers in some columns 4.6 pt right of Excel's (not found).
- The PDF's text of a glyph two characters share is the other one: Calibri's hyphen is U+2010 in our PDFs' text
  ("31‐Jan‐16"), from Rbt.Pdf's ToUnicode (its extra glyph mappings do not override the cmap's); words differ.
- Fonts the machine has no file of (Aptos, Book Antiqua, Century Gothic, Myriad Pro) are set in Calibri.

## Rich grid objects: pivot styles, sparklines, icon sets, data bars against Excel 2019, 2026-11-22

The objects "fancy" templates are made of, against Excel 2019 on the VM: the dashboards' pivot tables and sparklines
(dashboard-sales-analysis, -ecommerce-sales), synthetic-condformat, synthetic-slicers, synthetic-sparklines, and
fixtures made in Excel by VBS (`probe-on-vm.sh`; pictures with `shoot-sheets.ps1`, ours with
`shoot-browser-sheets.mjs`): `build-pivot-styles.vbs` (synthetic-pivot-styles.xlsx, 12 pivot tables in 12 styles,
with Excel's `DisplayFormat` of every cell around each in synthetic-pivot-styles.excel.tsv),
`probe-pivot-style-order.vbs` (which element of a pivot style covers which cell, which wins where two meet),
`build-sparklines.vbs` (synthetic-sparkline-looks.xlsx, 19 cases in cells of 200 x 40 px), `build-icon-sets.vbs`
(synthetic-icon-sets.xlsx: the 20 icon sets, reversed, icons only, centred, a tall row, 4 kinds of data bar).
Checks: tests/formula stage64-pivot-styles, stage65-sparklines, stage66-icon-sets.

Ranked by how far off we were: (1) pivot tables drew no style at all - no fills, bold, lines, buttons - on every
pivot of the dashboards (16 and 30), and marked their text row labels with error triangles; (2) sparklines were not
drawn (the sales dashboard's 24); (3) icon sets were characters of a 10-px font (a traffic light a 4-px dot);
(4) form controls (check boxes, option buttons, combo / list boxes, spinners, scroll bars, group boxes, buttons:
legacy VML and ctrlProps) are not drawn at all - none in the corpus or the real-estate models but
houseflipping's 5 zero-height filter drops (not shown by Excel either); (5) data bars 1-2 px off, their gradient
and right-to-left lengths; (6) slicers: Excel's header icons (multi-select, clear filter), item order, no-data look,
the timeline's caption ("Mar - Dec 2012", ours "Mar 2012 - Dec 2012") and scroll bar; (7) table styles: already at
parity (0 of 272 cells, 2026-11-10).

| Rule (Excel 2019) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| A pivot table is drawn in its style (`<pivotTableStyleInfo>`): the built-in PivotStyleLight1-28 / Medium1-28 / Dark1-28 are presetTableStyles.xml's, whose pivot styles number their dxfs from 0 (table styles from 1) | `DisplayFormat` of the 1,779 cells of synthetic-pivot-styles.xlsx (Light16's header fill #D9E1F2, bold, a bottom line #8EA9DB) | plain cells | `TableStyles.pivotPresets` (`make-table-style-presets.py --pivot`), `Model.pivotCellEntry`, `Layout.styleEntryAt` |
| The parts: the whole table and its page fields' cells; the header rows (to the first data row) when column headers are on; the row-label columns from the table's top when row headers are on; the first header cell (the header rows but the last, the label columns); banded rows (every line, the grand total too) and columns (every data column, the grand total too); a row subheading for each line of an outer item (compact / outline: the whole line; tabular: the item's label column down its lines); column subheadings in the column-item rows; subtotal rows; subtotal columns from the first column-item row; the grand total row; the grand total column (when the last column is on); page field labels / values | the 12 pivot tables: compact, outline, tabular, page fields, subtotals at the top / bottom, 2 and 3 row levels, 2 column levels, stripes, headers off | - | `TableStyles.pivotBlocks` |
| Fills and fonts are laid in this order (a later one wins): whole table, column stripes, page fields, row stripes, first column, header row, first header cell, subtotal columns, column subheadings, blank rows, subtotal rows, row subheadings, grand total column, grand total row | `probe-pivot-style-order.vbs`: every two of the 25 elements that fill one cell, alone, in 3 pivot tables: no contradiction | - | `pivotOrder` |
| Lines do not follow one order: an edge between rows shows the top line of the block below it, unless the block above is the header row or its bottom line is heavier; either beats a line between a block's rows | the same probe's edges (180 of 417 edge-or-inside pairs go to the inside line - by side, not by element); Deep (a stripe's top over the subheading above), Tabular (the header's bottom over the subheading below), Dark12 (a medium subheading bottom over a thin stripe top) | - | `pivotLooks`: 1,779 of 1,779 cells as Excel |
| A pivot table's buttons: a drop-down (17 px, 1 px inside the cell's bottom right, grey edge #A6ACB3, white to #EFF2F7, a 7-px arrow) on each page field's value, "Row Labels" / "Column Labels" in compact form, each field's caption otherwise (not the Values field); a +/- (9 px, cut corners, edges #AAB5C9 / #8A9EBF, #EBEDED to #D8D8D8, a 5-px #3D4E7C minus) 5 px right of the grid line, centred down, before the label of an item with items under it; the label 14 px right (in compact form every item label) | pictures of the 12 pivot tables (Compact: the box at x 94-102 of a column from 89, "2023" from 106) | none | `Model.pivotButtons`, the grid's `drawCell` (the web grid only; Excel prints none) |
| No error-check marks inside a pivot table | the sales dashboard's "2016" / "2017" (text) unmarked | green triangles | the grid's marks |
| A sparkline's plot is the cell from 3 px right of its left grid line to 2 px left of its right one, rows 5 .. h - 4; n points at the middles of n slots; a column its slot less 15 % a side; the scale from the lowest value to the highest (or the group's, or custom); a column from 0 (or the scale's end nearer 0), its top the pixel row its value falls in, an eighth down (ColMany 7 of 1..12 at 19.09: row 19); all-positive data: the lowest a 1-px line; win / loss: a win the upper half (rows 5-20), a loss the lower; empty cells break the line (gap), are 0 or are spanned; the line its weight rounded to pixels (0.75 pt 1 px, 2.25 pt 2 px); markers 4 px; first / last / high / low / negative points in their colours | Excel's picture of synthetic-sparkline-looks.xlsx: every column edge of 4 column cases within a pixel, line ends within a pixel; the sales dashboard's 24 | not drawn | `Sparklines.marks` / `svg`, `XlsxBridge.sparklineGroupsOf`, data from any sheet (a lazily read sheet is read for it) |
| An icon is a 16 x 16 picture 2 px right of the left grid line, 2 px above the bottom one (top / centre: as the value), Excel's flat colours: red #D65532, amber #EAC282, green #68A490, grey #808080, dark #505050, blue #4A7DB1, light #B3B3B3, each with a darker edge | Excel's picture of synthetic-icon-sets.xlsx (every icon's box 2-17 x 3-18 of a 20-px row; 23-38 of a 40-px one) | a 10-px character, other colours | `ConditionalFormatting.iconSvg` (the web grid; the painter still draws characters) |
| A data bar runs from 2 px right of the left grid line to 2 px short of the right one, 2 px in from top and bottom | the 4 bar rows of synthetic-icon-sets.xlsx (a full bar of a 64-px column: pixels 2 .. 61) | 1 .. 63 | the grid's bar room |

Found, not matched: form controls are not drawn (nothing read of `ctrlProps` / VML `x:ClientData` but notes);
ActiveX placeholders not drawn; a data-validation list's arrow (Excel: only beside the selected cell; ours: the
painter draws it on every such cell in PNG / PDF); icon pictures are drawn to Excel's size and colours, not
pixel for pixel (arrows' and flags' shapes), the canvas / PDF painter still draws characters, a left-aligned text
is not moved past its icon; data bars: Excel 2010+'s gradient fades to a light tint of the colour (ours to white),
the right-to-left bars of synthetic-icon-sets.xlsx are 2-6 px longer than Excel's, the bar of the lowest value
under "lowest value" is empty in Excel (ours: Excel 2007's minimum length); colour scales not re-measured;
sparklines: a date axis, hidden cells, columns of 7 in 200 px off by a pixel at 2 edges (Excel's sub-pixel
rounding), the line's anti-aliasing (ours bleeds a pixel above a peak), sparklines are not moved by inserted rows
(drawn where the file put them); the +/- of a collapsed item is drawn from the lines (an item without lines
under it shows +), pivot buttons do not respond to clicks, compact-form indents of 2+ levels are the cell's
indent (Excel's ~12 px a level, ours ~10); pivot tables without `pivotTableStyleInfo` / with custom styles
are read but not compared with Excel; slicer details (6 above). (Form controls, ActiveX pictures, slicers, data
bars and compact indents: since done, "Form controls, slicers, data bars, indents" below.)

## Form controls, slicers, data bars, indents against Excel 2019, 2026-11-22

Items (4)-(6) of the ranking above and the compact-indent finding, each from fixtures made in Excel 2019 on the VM
(`probe-on-vm.sh`) and pictures of them (`shoot-sheets.ps1` at 100 %) against ours (`shoot-browser-sheets.mjs`),
compared pixel by pixel: `build-form-controls.vbs` (synthetic-form-controls.xlsx: every form control flat and with
3-D shading, linked cells, two ActiveX controls; and copies whose linked cells hold other values), `build-slicer-looks.vbs`
(synthetic-slicer-looks.xlsx: Region slicers sorted up / down, no data in place, hidden, a filtering Product slicer,
an unfiltered two-column one), `make-timeline-variants.py` (synthetic-slicers.xlsx's timeline at 12 selections and
levels; `SlicerCaches.Add2(..., xlTimeline)` is refused from a script), `build-data-bars.vbs` (synthetic-data-bars.xlsx:
six colours, right to left, negative values with an automatic / middle / no axis, a border), `build-pivot-indent.vbs`
(synthetic-pivot-indent.xlsx: compact pivots of CompactRowIndent 1-3, in Calibri 11 and 16 cells) and a probe of
indent levels in five cell fonts under three Normal fonts. Checks: xlsx-browser `controls-excel-parity`,
`databars-excel-parity`, `indent-excel-parity`; tests/formula stage56-slicers (orders, header buttons, 12 timeline
wordings); tests/office-formats xlsx-bridge ("controls:": read, and written back byte for byte, edited or not).

| Rule (Excel 2019) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| A form control is painted over its anchor grown by a pixel at the left and top (a combo box's, group box's, button's edges on the pixel left of / above the anchor and on its last pixel right / below) | Flat / Shaded: every control's edges | not drawn | `FormControls.svg` over (w + 1) x (h + 1) |
| Check box / option button: a 13-px box or circle 4 px into the anchor, centred down; flat black, 3-D Windows' sunken edges (#A0A0A0 / #646464 above-left, #F0F0F0 inside right-below); a 7-px check mark or dot 3 px in; mixed: a chequer (flat) or a 7x7 dotted pattern (3-D); the caption (the VML text box's font, Segoe UI 8 here) 20 px into the anchor | pixel maps of both sheets | - | masks of Excel's pixels |
| Group box: a frame on the rectangle (flat black; 3-D #A0A0A0 with white inside), open for the caption from 8 px to its GDI width + 19, caption baseline 7 px down | "Group box" | - | `svg` "GBox" |
| Combo box: frame #828790 (flat) or sunken (3-D), a raised button 23 px wide (`dx` 22 + 1) with an 11 x 6 arrow, the item in Tahoma 8; list box: black frame (flat), items 13 px apart from 3 px down, the selected #0078D7 / white, a 21-px scroll bar (thumb visible / items of the track) when they do not fit, items cut at the frame | Drop / List | - | `svg` "Drop" / "List" |
| Spinner: two classic raised buttons (#F0F0F0 / white, #A0A0A0 / black) with black arrows; scroll bar: Windows 10's flat look with or without shading - white line along its top (left), track #F0F0F0, #606060 chevrons in square buttons of its thickness, thumb #CDCDCD thickness - 13 px long at value / (max - min) of the track | Spin / Scroll at 70 % and 7/10 | - | `svg` "Spin" / "Scroll" |
| Button: #F0F0F0 with #A0A0A0 / #646464 at the right and bottom, its caption (Calibri 11 here) centred | Button 1 | - | `svg` "Button" (its macro runs on a click) |
| What a control shows is its linked cell's, not the state saved in `ctrlProps` (`checked`, `sel`, `val`): TRUE or a number not 0 checks a check box, FALSE / 0 clears it, #N/A makes it mixed; an option group, combo or list box selects the item the number's whole part gives (0 or past the buttons: none; past a combo's items: its last); a scroll bar's value is the number kept to min .. max; text or an empty cell: the saved state | 4 copies of the Flat sheet with other cell values (1, -2, 0, #N/A, "x", empty, 0 / 7 / 12, 2.7, -5, 150) | - | `FormControls.checkedOf` / `selectedOf` / `valueOf`; option groups = the smallest group box a button lies in (else the sheet), in file order |
| A click: a check box toggles TRUE / FALSE, an option button writes its place, a list item its number, a combo box drops its list (DropLines items) and writes the item picked, a spinner's / scroll bar's arrows step by `inc`, the track by `page`, kept to min .. max - as an entry of the cell (one undo step; a linked cell on another sheet too) | (interaction: no Excel picture) | - | `pressFormControl`; without a linked cell the state is the grid's own |
| An ActiveX control is shown by its picture (`controlPr r:id`, an EMF of Excel's drawing: fills by brush, text, a DIB) | the CheckBox1 / CommandButton1 pictures | not drawn | `Emf.svg` (EMF records to SVG: mapping, brushes, pens, rectangles, ellipses, polygons, lines, BITBLT / STRETCHDIBITS as BMP, EXTTEXTOUTW) |
| A slicer's items: those with data first, each part sorted up or down (`sortOrder`); `showItemsWithNoData`: all in their sorted places; `x15:slicerCacheHideItemsWithNoData` (in the cache's own extLst) hides them | RegionUp East North South / Central West; RegionDown South North East / West Central; RegionInPlace; RegionHide | the no-data items of a hiding cache shown (the extension looked for in `<tabular>`) | reader fixed |
| A slicer's header buttons: Multi-Select (three ticks before lines) 60 px and Clear Filter (a funnel) 28 px left of its right edge, 6 / 7 px down; Clear Filter has a red #D65532 cross while some item is not selected, all light grey (#B0ADB0) while none is filtered | ProductS, RegionUp (red), RegionDown, ProductAll (grey) | none | `Slicers.headerButtons` |
| A slicer is framed a pixel wider and higher than its anchor; its buttons and header line 9 px in from the frame; buttons of whole pixels (241300 EMU = 25 px, 28 apart) with corners rounded by 2 px; the caption stops before the header's buttons | frames at 12 .. 204 of an anchor 13 .. 204, button edges 21 .. 195 | 10 px in, 25.33 px, square | `objectRect`, `margin`, rounded `PRect` |
| A timeline's selection label says what both ends share once: "Mar - Dec 2012", "Nov 2012 - Feb 2013", "Jun 2012"; "Q2 - Q3 2012", "Q4 2012 - Q1 2013", "Q2 2012"; "Mar 5 - 20 2012" (no comma), "Mar 28 - Apr 3, 2012", "Dec 30, 2012 - Jan 2, 2013", "Mar 5, 2012"; "2012 - 2013", "2012"; a day level's month labels in capitals ("MAR 2012") | the 12 variants' labels | "Mar 2012 - Dec 2012" | `Slicers.selectionText` |
| A gradient data bar fades from its colour to 85 % of the way to white, channel by channel (not to white), each pixel the colour at its left edge - Excel 2007-form bars (no x14 part) too | 6 colours: #638EC6 to #E7EEF6, #FF0000 to #FFD9D9, #000080 to #D9D9EC (Excel's last pixel 60/61 of the way) | to white | `ConditionalFormatting.dataBarFadeColor`, gradient stops 0.5 px in |
| A data bar's room is 61 px of a 64-px column (from 2 px right of the left grid line to 2 short of the right one); a bar is round(fraction x room) whole pixels from where it grows (right to left: from the right end: 25 % 15 px, 50 % 31, 75 % 46); with an axis (a dashed line on the pixel floor(place x room)) a negative bar ends at the pixel before it, a positive one starts at the pixel after it | every bar's first / last pixel in synthetic-data-bars.xlsx and synthetic-icon-sets.xlsx | 60 px of percentages; RTL bars 1-2 px longer | `dataBarRoom`, the grid's bar nodes |
| A level of indent is three spaces of the workbook's Normal font as GDI draws them, whatever the cell's font: 9 px (Calibri 11), 12 (Arial 12), 21 (Verdana 14) | indent 0-4 in Calibri 11 / 16, Arial 10, Times 12, Verdana 10 cells under three Normal fonts: 9 / 12 / 21 px a level throughout | 1.4 digit widths of the cell's font (~10 px) | `SpreadsheetModel.IndentUnit` (`XlsxBridge.indentUnitOf`) |
| A compact pivot table moves a level of row items in by its CompactRowIndent of indent units, at least 13 px: 13 / 18 / 27 px for 1 / 2 / 3 in a Calibri 11 workbook, the same in Calibri 16 cells; the file says `indent="0"` for 1, nothing for 2, `indent="2"` for 3 (the cells' own `indent` is CompactRowIndent x level); the +/- 5 px into the cell | +/- boxes at 4, 17 / 22 / 31 px; leaves at 42 / 52 | the cells' indent (~10 px a level), the +/- 4 px in | the grid's pivot labels, `PivotStyleInfo.CompactIndent` |

Left (most since done: "Text insets, baselines, fills, moving controls, timelines, ActiveX pictures" below): a timeline's Clear Filter button, its scroll bar's arrows and its text sizes (Excel's selection label and period
labels are a point or two smaller); combo and list box text is the browser's Tahoma, not GDI's ClearType; form
controls are not moved by inserted rows or columns (drawn at the file's anchors) and not printed (the PNG / PDF
painter does not draw them); a control's caption is one line per VML line (no wrapping); the 3-D radio and the
check marks are Excel's pixels, the arrows' sizes at other control sizes are a rule of thumb; an ActiveX picture is
the file's (Excel draws the live control: a check box linked to an empty cell grey); EMF clipping, ROPs and paths
are not drawn; a combo box's list drops in the page's own look; a text cell aligned left (`horizontal="left"`)
starts 2 px further right in Excel than ours (synthetic-cell-indent.xlsx: 8 vs 6 px at indent 0 - not looked into).

## Text insets, baselines, fills, moving controls, timelines, ActiveX pictures against Excel 2019, 2026-11-22

The "Left" items of the form-controls round and the text inset found on its fixture, each from generated workbooks in
Excel 2019 on the VM and pictures of them (`shoot-sheets.ps1`, at the sheet's zoom) against ours
(`shoot-browser-sheets.mjs`, which now hides the fill handle too). Fixtures and probes: `make-text-inset-fixture.py`
(7 Normal fonts x their alignments - general, left, right, centre, indents, fill, distributed, justify, numbers,
wrapped - and 30 cell fonts / sizes at 75-200 %, each case between black marker cells; `measure-text-inset.mjs`
lays our ink over Excel's: synthetic-text-inset.xlsx, synthetic-text-inset-align.xlsx), `probe-control-moves.vbs`
(15 form controls at each Placement through 6 row / column edits: synthetic-control-moves.xlsx +
synthetic-control-moves.excel.tsv), `probe-shape-places.vbs` (where Excel puts the shapes of our saved file),
`make-timeline-variants.py` (13 timelines), `build-activex.vbs` (14 ActiveX controls, saved while shown:
synthetic-activex.xlsx). Checks: xlsx-browser `text-inset-excel-parity`, `grid-text-excel-parity`,
`grid-excel-parity`, `controls-move-excel-parity`, `slicers-in-grid`, `activex-emf-excel-parity`; tests/formula
stage56-slicers.

| Rule (Excel 2019) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| A line of text stands ceil(d / 4) px after the grid line at the cell's left and ceil(d / 4) + 1 before the one at its right (that grid line is the cell's), d the widest digit of the cell's own font as drawn at the zoom - not the Normal font's: Calibri 11 2 / 3 px, Calibri 14-18 3 / 4, Calibri 20 / Arial 20 4 / 5, Calibri 36 6 / 7, Calibri 11 at 150 % 3 / 4, at 200 % 4 / 5. The same for general, left, right, indented (the indent added), wrapped and distributed text and for numbers (a number's room between them: `fittedRoom`) | make-text-inset-fixture.py: 446 of 447 cases within a pixel of Excel, 366 on it (the rest the grid's rounding of cell edges at zooms); then every 100 % sweep (Calibri, bold Calibri, Arial, Times) on Excel's pixels | 2 / 3 px x the zoom whatever the font, measured from the node's box - which starts on the grid line at its left -, so a pixel left of Excel everywhere (synthetic-text-zoom: "T" 2 px after the grid line, Excel's 3) | `CellText.textInset`; paddings ceil(d / 4) + 1 / ceil(d / 4) of the node (style sheet 3 / 2 px); overflow, shrink and wrap room the cell less 2 ceil(d / 4) + 1 |
| Fill repeats the text as many whole times as fit in the cell less its left inset and a pixel; distributed spreads its one line from inset to inset, its indent on both sides | Calibri 11 "H" 18 times in 173 px (9 px each), Calibri 20 "H" 20 times in 346; "Hx" 10 / 11 / 10 at 100 / 75 / 90 %; "H H H" across 166 of 172 px, indented 1: 9 px in from either side | drawn once; distributed as left | `displayText` of the grid, `text-align-last: justify` |
| The bottom-aligned lines of a row stand on one baseline, the largest GDI descent of them above the row's grid line (Calibri 11 4 px; in a row with Arial 20 6 px; Calibri 36 13); a wrapped text's last line too | the bottoms of the ink of every case of the fixture: at 100 % every sweep on Excel's pixel | the browser's line box a pixel below the row's top: 2 px high in a 20-px row of Calibri 11, 4 in a row of Calibri 36, wrapped 3 | the grid's `gridRowDescents`, `fontBox` (where the browser lays the font's line out) |
| A fill covers the cell from grid line to grid line: the grid lines at its right and below it are in its fill (a black cell's block 1 px wider and higher than its node) | the markers of the inset fixture: Excel's black blocks 33 x 22 px, ours 32 x 21 | the grid line left in the neighbour's colour | an outset shadow where nothing is drawn there, else the neighbour's first column / row of pixels in the fill (`contain: size layout style` for such cells) |
| Form controls move and size with their cells by their Placement: "move and size" (controlPr `moveWithCells` + `sizeWithCells`; VML neither `x:MoveWithCells` nor `x:SizeWithCells` - buttons, spinners, labels) both corners with their cells; "move" (`moveWithCells` alone; VML `x:SizeWithCells` - check boxes, option buttons, combo / list / group boxes, scroll bars; a check box refuses "move and size") as large as they were; "don't move or size" (neither; VML both) where they were. A deleted near corner's cell puts it at the deletion | probe-control-moves.vbs: every control after each of 6 edits (insert rows 1:2, column A, row 6 40 pt, column D 25, delete rows 7:8, column C) where Excel has it (1 px); Excel opening our saved file puts every control where it put them itself | drawn at the file's anchors | `Model` (structure edits: `moveAnchor` of drawings and controls), `keepAnchoredSizes` (sizes, hiding, autofit); pictures and shapes anchored "oneCell" kept their size the same way |
| ... and the file written has the controls' anchors where they are now: `<controls>` (`<from>` / `<to>`) and the VML's `x:Anchor` | xlsx-browser controls-move-excel-parity: written, read again from `<controls>` and from the VML alone | the file's | `SheetExtras.ControlAnchors`, `XlsxPackageWriter.withControlAnchors` / `withVmlControlAnchors` |
| A timeline: no fill of its own in a custom style without one; its header's Clear Filter button 25 px from the right (red cross while a period is picked); the selection label, level and labels a pixel lower than ours were; the level's arrow a light grey (#B2B2B2) triangle; periods by their days - a month 1.6 px a day (48 / 49 / 50 px), a quarter 0.3771 px a day, a year 54, a day 30 - from 7 px in, a pixel between blocks 17 px high, only the cache's bounds' periods, cut 10 px from the right; a selection out of sight of the scroll position centred (Mar 5 - 20 at the day level from Mar 4); the scroll bar - no track, white arrow buttons framed #C4C4C4 7 px from the left and 9 from the right, a white thumb as long as the shown part of the bounds where it is scrolled to, none when all the bounds show, the selected days light green on the track | synthetic-slicers.xlsx and the 13 timelines of make-timeline-variants.py: every block on Excel's pixels but one 1-px end; the thumb 255 of 529 px | white, no Clear Filter button, a grey track and thumb, months 48.5 px, a black "▾", blocks 16 px; days from the scroll position | `Slicers.timelineItems`, `clearFilterButton` |
| Combo / list boxes and captions: each line as wide as GDI draws it at its pixel size (Tahoma 8 "North" 27 px); a list box's items 5 px in | synthetic-form-controls.xlsx: the items' ink starts on Excel's pixel | the browser's advances; 4 px | `FormControls` text `textLength` |
| An ActiveX control's picture (EMF): pattern brushes (CREATEMONOBRUSH: 0s in the text colour, 1s in the background colour - a pressed toggle button, scroll bar tracks), a mask then a picture painted through it (SRCAND, SRCPAINT: option buttons), clipping (INTERSECTCLIPRECT, EXTSELECTCLIPRGN, SELECTCLIPPATH; saved with the DC), opaque / clipped text (ETO_OPAQUE, ETO_CLIPPED) at each character's advance, hung from its top by the font's Windows ascent (Calibri 15 px: 14), paths (BEGINPATH .. FILLPATH / STROKEPATH / STROKEANDFILLPATH), POLYLINETO, POLYBEZIER(TO), POLYPOLYLINE, MODIFYWORLDTRANSFORM | synthetic-activex.xlsx: every control as Excel shows it but a list box's selection (Excel draws the live control, without its highlight; its picture has it) | none of these (grey toggles, black option rings, text 3 px high); a BITBLT's source read a field early | `Emf.svg` |

Left: text at zooms other than 100 % within the grid's own rounding of cell edges (1 px either way); Segoe UI Black's
bottom line a pixel lower than Excel's; a distributed line stands 2 px higher in Excel than the row's other lines; a
row's descent counts every bottom-aligned cell of it, not only those drawn; a fill's corner pixel at the right of its
top grid line; combo / list box text in the browser's grayscale, not ClearType; the real-estate models' first screens
were shot again (53 sheets; on the sheets the header-aligned diff lines up, acre-apartment's and acre-basic's, 0.2-1.1
points fewer differing pixels; others not aligned by it), the dashboards not re-surveyed; the old note that left-aligned text starts 2 px further right in Excel was the 1-px box offset
above (synthetic-cell-indent.xlsx needs no rule of its own).

## Formula oracle on real-world models, 2026-11-21

Eleven real-world financial models (local only, private-use licences: not in the repo) against Excel
2019's full recalculation (`dump-values.vbs`, `compare-values.fsx`), and - for the two that need Excel
365 (`@`, SEQUENCE, XMATCH, dynamic arrays) - against the values the file keeps (`--cached-new`).
Before: 86,045 of 97,011 cells agreed. After: every cell agrees (99,977 of 100,235 at 1e-9, the 258
others within 1e-6: XIRR within Excel's own accuracy, and a model with iterative calculation, which
stops within its 0.001). The corpus oracle is unchanged: 50,480 of 50,496 (the same 16 XIRR).
Each rule below was confirmed on generated cells in Excel 2019 (`eval-cells-on-vm.sh`):

- A formula whose value is an empty cell (`=B9`, `=IF(c,"",B9)`) is 0: `=A1=""` reads FALSE,
  `CONCATENATE(blank," ",x)` gives "0" for it. So is an empty cell of a dynamic array. No negative zero.
- A circular reference only through the branch of an IF / IFS / IFERROR / IFNA / CHOOSE the values do
  not take is none: `=IF(B1="",5,A2)` in A1 with `=A1` in A2 gives 5 and 5 (`Formula.liveReferences`,
  in both passes). With the branch taken it is one (below).
- Cancellation to 0 at the end of a formula's addition: a result under 8 units in the last place of
  the larger operand (7 are 0, 8 are not, for any mantissa), was 2^-48 of it. SUM's last addition
  the same, inside any expression (`=1*SUM(..)` too; `=SUM(a,-b,0)` not). Not SUMPRODUCT, not an
  operation inside a function. Parentheses around the whole formula are an operation of their own
  (2026-11-21, 80 shapes with A1 1.5 and B1 3 units above it): `=(A1-B1)`, `=((A1-B1))`, `=( A1-B1 )`,
  `=(SUM(A1)-B1)` are not 0; `=(A1)-(B1)`, `=(A1*1)-B1`, `=(0.1+0.2)-0.3`, `{=A1-B1}` are, and so is
  `=(SUM(A1,-B1))` (SUM's own). `Formula.wrappedForms`, marked by the parser.
- SUMIF / SUMIFS: an error among the cells added is the result.
- Another workbook's cells (`[1]Sheet!A1`, `'[1]My sheet'!A1`) are what xl/externalLinks keeps of
  them (`Workbook.External`); a sheet with `refreshError` is #REF!. `'Sheet'!#REF!` is #REF!.
- IRR with an error among its cash flows is #VALUE! (NPV, XIRR, MIRR: the error). An error in an
  optional argument of PMT, FV, PV, NPER, RATE, IPMT, PPMT, IRR, XIRR is the result (was the default).
- XIRR finds a rate just above -1 (a large outlay, little back soon) down to ~1e-9 above it, as Excel.
- PMT: (1 + r)^n as exp(n ln(1 + r)) with ln(1 + r) from r - the power of 1 + r rounded was off by
  up to n/2 units in the last place, and an amortization schedule ended on another 1e-7 residual.
  Excel's PMT is within a unit of the exact value, not always the same unit.
- SEQUENCE / RANDARRAY with an argument left empty (`SEQUENCE(20,,0)`) take its default.


### Circular references, 2026-11-21

Generated workbooks with values in the file, opened in Excel 2019 (`eval-file-on-vm.sh`, modes none /
calc / full), and cells typed one by one (`eval-cells-on-vm.sh`, `#ITERATE`, `#NOCALC`):

- Without iterative calculation the formulas of a circular reference **and every formula that reads
  one** (directly or not) are not calculated: they keep the value they have - the file's (`=A1+C1` /
  `=B1*0.5` at 100 / 50 stay 100 / 50; `=D1*2` reading `=B1+1` keeps 4; `=ISNUMBER(B1)` keeps 0), 0 for
  one just typed (`=A1+5` beside a cycle). Not even when another input of the reader changes (`=A1+C1`
  stays 0 when C1 is typed again; `=C1*2` beside it follows). Was: the cycle error, and what read it had it.
  `Recalc` (`keptOf`, `Node.Held`: also for an engine that only looked for the cycles; a sort that moves
  such a formula goes to the full pass) and the per-sheet pass (`Formula.keptValue`).
- With iterative calculation (`calcPr iterate`): sweeps over the formulas of the cycles **and what reads
  them**, in reading order (sheet, row, column - `=C2*0.5` / `=100+0.1*C1` at 20 / 50 give 25 / 102.5;
  `=E2*3` above `=E2/2+1` reads it a sweep behind), each from the values the others have then, starting
  from the file's. They stop when nothing calculated in the sweep changed by the maximum change - what
  reads the cycle counts (`=A6/2+1` from 0 takes 11 sweeps at 0.001 alone, 13 with `=A6*3` beside it) - or
  after the number of iterations (1: one sweep); the last sweep is kept (was dropped). Opening a workbook
  calculates it; `CalculateFull` then sweeps again (one more, when settled); and every calculation -
  any edit - sweeps once more (`=A6/2+1` typed is 2 - 2^-10, after five more entries 2 - 2^-15). All 6
  iteration settings x 3 cycles agree exactly with Excel (one run against the open, two against open +
  CalculateFull). A workbook calculated from nothing (no values yet) is swept in the order of what the
  formulas read instead (from nothing, reading order lets an #N/A go round a cycle for good: saopaulo-condos).
  The real-estate model with iterative calculation: 4196 of 4225 cells exact (was 3972), the 29 others near:
  ours starts from nothing, Excel from its own last values.
- Not reproduced: the formula that closes a cycle, typed in the editor, has the value it was given as it
  was entered (the editor calculates a typed formula once before the recalculation); Excel's is 0.
- Not reproduced: a formula with no value in the file (`<f>` without `<v>`) is "" in Excel (so `100+0.1*B1`
  on a cycle is #VALUE!); ours starts it from 0.

### Comparisons at 15 significant digits, 2026-11-21

Cells typed into Excel 2019 (`eval-cells-on-vm.sh`, values exact) and a generated workbook opened in it
(`eval-file-on-vm.sh`). Regression tests: `tests/formula/stage22-semantics.fsx` (`cmp-*`, `literal-*`,
the functions over `cmpModel`).

| Rule (Excel 2019) | Evidence | Where |
| --- | --- | --- |
| `=` `<>` `<` `>` `<=` `>=` compare two numbers rounded to 15 significant digits: equal there is equal (`<` FALSE, `<=` TRUE), otherwise the binary order | `=0.1+0.2=0.3` TRUE, `>` FALSE, `<=` TRUE; 1+22 ulp = 1, 1+23 ulp <> 1 and = 1+45 ulp; 1+22 ulp <> 1+23 ulp (one ulp apart, either side of 1.000000000000005) | `Formula.compareNumbers`: the general operator, `applyBin`, FILTER's conditions (three), the array streams of SEQUENCE / MAKEARRAY / SUMPRODUCT (`--(a=b)`), the constant-term and min / max shortcuts (a comparison of terms a constant apart is no longer taken for the constants' order) |
| The rounding is of the exact binary value, a half away from zero (not to even) | 1E15+5 = 1E15+10, <> 1E15; 1E15+25 = 1E15+30, <> 1E15+20; 1E14+2.5 = 1E14+3; -(1E15+5) = -(1E15+10); 1E15+4 <> 1E15+6; 10-2x2^-49 = 10, 10-3x2^-49 <> 10 | `NumberText.sameToSignificantDigits` (exact digits, rounded half up), only for numbers within 2e-14 of each other |
| Relative only: no absolute tolerance near 0 | `=1E-16=0` FALSE, `1E-16>0` TRUE; `=1*(0.1+0.2-0.3)=0` FALSE (5.55e-17); 1E-300 x (1+22 ulp) = 1E-300, x (1+23 ulp) not; `EXP(-740)=0` FALSE, `=EXP(-740)=EXP(-741)` FALSE (subnormal) | `compareNumbers` |
| The criteria of COUNTIF, COUNTIFS, SUMIF, SUMIFS, AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS, DSUM / DCOUNT compare the same way, a number criterion or one in text (`"=0.3"`, `"0.3"`, `"<=0.3"`, `">"&A1`) | over 0.1, 0.1+0.2, 0.5: COUNTIF 0.3 / "=0.3" / "0.3" 1, "<=0.3" 2, ">0.3" 1, "<>0.3" 2, "<0.3" 1, ">=0.3" 2; SUMIF 2, SUMIFS ">0.3" 3, MAXIFS 2, MINIFS ">0.3" 3, DSUM 2; COUNTIF(1E15..1E15+30, 1E15+10) 2 | `compileCriteria` |
| SWITCH compares like `=` | `=SWITCH(0.1+0.2,0.3,1,0)` 1 | `SWITCH` (`compareOperands`) |
| MATCH (0, 1, -1), VLOOKUP / HLOOKUP (exact and approximate), LOOKUP compare the binary values | `MATCH(0.3,{0.1;0.1+0.2;0.5},0)` #N/A, `,1)` 1; `MATCH(0.1+0.2,{0.5;0.3;0.1},-1)` 1; VLOOKUP FALSE #N/A, TRUE 1; HLOOKUP the same; LOOKUP 1 | unchanged (`compareValuesNoError`) |
| So do RANK (#N/A for a number not in the list), MODE, FREQUENCY's bins, DELTA, GESTEP; EXACT compares the texts (TRUE) | `RANK(0.3,{0.1;0.1+0.2;0.5})` #N/A (was 3: ours took a number within 1e-12 as in the list), `MODE(0.1+0.2,0.3)` #N/A, `INDEX(FREQUENCY(..,0.3),1)` 1, `DELTA(0.1+0.2,0.3)` 0, `GESTEP(0.3,0.1+0.2)` 0, `EXACT` TRUE | RANK / RANK.EQ: equal binary values (`rankFold`, `rankSelected`, `rankCompressed`) |
| A number written in a formula, typed into a cell, or text read as a number (VALUE, NUMBERVALUE, `--"..."`, `"..."*1`, a criteria's number) keeps its first 15 significant digits, the rest cut, not rounded; so does a formula in a file (`<f>`), not a value (`<v>0.30000000000000004</v>` is kept) | `=1.0000000000000099` 1, `=123456789012345.6` 123456789012345, `=0.1234567890123456789` 0.123456789012345, `=1000000000000005` 1E15, `=9.9999999999999999` 9.99999999999999, typed the same; `VALUE("1.0000000000000099")` 1; `COUNTIF({1+45 ulp;1},"<1.0000000000000099")` 0; `<f>(0.30000000000000004-0.3)*1</f>` 0, `<v>` kept: `(A1-A2)*1` 5.55e-17 | `NumberText.cutToSignificantDigits`: the formula tokenizer (array constants too), `Input` (cell entry), `coerceNumber` and the other text-to-number reads of `Formula`, VALUE, NUMBERVALUE, criteria |

Not probed (Excel 2019 has none of them): XLOOKUP, XMATCH - taken to compare as MATCH does (binary);
SORT / SORTBY / UNIQUE - left binary. FILTER's condition is an operator and rounds.

Found, not matched: Excel keeps no subnormal number from `*`, `/` and `^` (`=2^-1074`, `=1E-307/1E10`,
`=1E-200*1E-120`, `=(2^-1022/4)*4` are 0) nor from a file (`<v>4.9E-324</v>` = 0 TRUE); `+`, `-`, SUM, EXP
and negation do keep one (`=2^-1022*1.5-2^-1022` is 2^-1023, `EXP(-740)` 4.2e-322). Ours keeps them all.

### IRR / RATE / XIRR (found, not matched)

Excel's IRR is within ~1e-13 of the root, not on it, and depends on the guess (`-1000,300,400,500`: guess
0.1 -3.5e-17, 0.05 -3.6e-13, 0.5 +2.5e-14), also for two cash flows (results 1.2 +- an ulp minus 1). No
Newton / secant / finite-difference variant in r, 1/(1+r) or 1+r (analytic or numeric derivative, 3 x 4
x 7 step sizes) reproduces the digits on 120 + 120 probes; the errors are close to Newton in 1/(1+r)
(about 1.15 times) without being it. XIRR is coarser: within ~1e-8 relative, the result the guess plus
steps of a power of two (guess 0.1 and 0.2 give the same digits, `-1000, 1100` a year later is
0.0999999940395 = 3355443 x 2^-25 with guess 0.1, 26843545 x 2^-28 with 0.5 or 1), a negative guess with a
positive root #NUM!; a stepping search (step-halving, bracket and bisect, 6 starts x 10 tolerances) does
not reproduce it either. Both stay within 1e-6 of ours (the 16 corpus XIRR cells and 5 of the real-estate
models are "near").

## Grid sizes and theme fills against Excel 2019, 2026-11-21

Found on the real-estate models (houseflippingspreadsheet's Arial MT 12 columns 29 % too narrow and black
buttons, choudhary's row styles stopping at the row's last cell, rows of wrapped text not fitted); each
rule then confirmed on generated workbooks opened in Excel 2019 on the VM - Excel's own numbers read with
COM (`Columns(i).Width`, `Rows(i).Height`, `StandardWidth`, `StandardHeight`) and its pictures taken with
`shoot-sheets.ps1`. Fixtures: `tools/excel-parity/make-sizing-fixtures.py` (synthetic-font-widths,
synthetic-normal-font-first / -style, synthetic-row-fit, synthetic-theme-shapes); checks: xlsx-browser
`sizing-excel-parity`, `tests/office-formats/xlsx-bridge.fsx` ("sizes:").

| Rule (Excel 2019) | Evidence | Where |
| --- | --- | --- |
| A column's width unit is the maximum digit width (mdw) of the workbook's first `<font>` - not the Normal style's xf font. Pixels = trunc(((256 w + trunc(128 / mdw)) / 256) x mdw): Arial MT 12 (drawn as Arial) mdw 9, widths 1, 2.098, 8.898, 10, 20, 38.797, 0.5, 2 are 9, 19, 80, 90, 180, 349, 4, 18 px; Calibri 11 mdw 7, 9.140625, 0.5, 2.098, 13, 1, 8.898 are 64, 3, 15, 91, 7, 62 px (truncated, not rounded) | COM widths of synthetic-font-widths / -row-fit "Widths"; fonts[0] Arial 12 + Normal xf Calibri 11 is mdw 9, the other way round 7 | `FontMetrics.columnWidthPx`, `XlsxBridge` import |
| mdw and a line's height per face and size: a sweep of 16 faces x 16 sizes and 14 more x 9 sizes as the Normal font (ColumnWidth 10 and 20: mdw = difference / 10; padding = 2 ceil(mdw / 4) + 1, 5 px at 7-8, 7 at 9-12); a face no font has ("NoSuchFont Xyz"), Helvetica and Arial MT measure as Arial | `StandardHeight` and widths of each, all consistent with the padding and default-width formulas | `FontMetrics` table (a weight of a family - "Segoe UI Black" - as its family: acre-property-acquisition's rows of it are Segoe UI's 22 px in Excel) |
| The default column (no `defaultColWidth`) is `baseColWidth` (8) x mdw + padding, up to a multiple of 8 px: 64 at mdw 7, 80 at 9, baseColWidth 12 at 9 is 120; `defaultColWidth` is a width like any other (20 at 9: 180 px) | synthetic-font-widths "Widths", "BaseCol", "DefaultCol" | `FontMetrics.baseColumnWidthPx` |
| The default row height is a line of the Normal font (Arial 12: 15 pt) whatever `defaultRowHeight` says (15.7, 30), unless `sheetFormatPr customHeight="1"` (then that, in whole pixels cut down: 20 pt is 26 px, 16 is 21); the empty cells of a row are in their column's style: the tallest font of the columns' styles counts, and the Normal font's while a column has no style (a `<col style>` of Calibri 20 on one column: every row 26.25 pt; Arial 10 on all 16384: rows of 12.75 pt, as dcf-stock-analysis's) | "RowDefault", "RowCustom", "S7", trigger "T12", "ColStyle", "ColStyleAll" (stdH 12.75, a Calibri 11 cell's row 15), "ColStyleSome" (15) | `XlsxBridge` import |
| Rows are fitted to their cells when Excel opens a file: every row without `customHeight`, whatever its `ht` (ht 30 of one short line is 15 pt, ht 12 of Calibri 20 is 26.25; synthetic-grid-parity's row of ht 15 in Arial 10 is 12.75 - the earlier rule "ht without customHeight is the height" was wrong). A cell is a line of its font (table), a wrapped text n lines (n x the line: 7 lines of Calibri 11 are 105 pt), a line break without wrapping one line, an empty cell with a font of its own counts, so does a row's style font and a cell in a hidden column; nothing makes a row lower than the default (8 pt text: 15) | synthetic-row-fit "Rows": 26.25, 105, 105, 15 (custom), 15, 150, 26.25 (empty 20 pt cell), 26.25, 15, 15, 26.25, 15, 45 ... ; same with `fileVersion` / `x14ac:dyDescent` added | `SheetOps.fitRowsOnOpen` |
| A row that says `thickBot="1"` or `thickTop="1"` is fitted a pixel higher for each (the attribute counts, not the cells' borders: medium, thick, double bottom borders without it change nothing) - choudhary's and bramosman's rows of `thickBot` are 21 / 22 px | probe of 12 border cases; "Rows" 28, 29: 15.75, 27.75 pt | `XlsxReader.Row.ThickEdges`, `XlsxBridge` |
| ... except on a sheet whose default height is its own (`sheetFormatPr customHeight="1"`) none of whose rows has a height of its own (customHeight, an `ht` other than the default, hidden) | S2, S3 (ht = default) not fitted; S4 (a custom row), S6 / T9 (ht 30 / 16), T10, T11 (hidden row) fitted | `XlsxBridge` import (`fitsRows`) |
| A row's style (`<row s customFormat="1">`) is every cell of the row the row does not list, to XFD, over a column's style; without `customFormat` the `s` is not drawn; a cell's own style is its own | picture of synthetic-row-fit "RowStyle" | `Model.RowStyles`, `SheetOps.shownCell` / `effectiveFormat` |
| A shape's style reference is the theme's formatting scheme: `fillRef idx` 0 none, 1-3 the scheme's fills, 1001+ its background fills, `phClr` the reference's colour; `lnRef idx` likewise (width and colour); DrawingML `tint` / `shade` mix with white / black in linear light (black tint 50 % is ~BCBCBC, Excel's pixel C0C0C0 near that end of the 2007 theme's gradient), `satMod` scales the saturation | picture of synthetic-theme-shapes: grey gradient E8E8E8 -> C0C0C0 for idx 2 in dk1 (was solid black), D5E0C4 for 1002 in accent3 (ours D5E0C4), 376092 -> B9CDE5 for accent1 lumMod 75 % / lumMod 40 % lumOff 60 % | `Charts.styleFill`, `styleLine`, `colorOfNodeWith` |

Fitting is right where the file's own `ht` is not: files Excel saved at another DPI keep heights of it (acre-apartment's `ht="15.6"` in a sheet of `defaultRowHeight="14.4"`), Excel at 96 DPI fits them again (21 px). On the 12 real-estate models (rows 1-40, columns 1-20 of 65 sheets, Excel's COM numbers): columns 1291 of 1300 as Excel (before: 1107; the rest are columns a sheet hides past its last), rows 2547 of 2600 (before: 1451; the rest: rows a mixed-font row 1 px off, wrapped help texts, rows a sheet hides past its last).

Print: a column prints at 7.5 px a digit of the sheet's maximum digit width whatever the font (`Pagination.printColumnWidthPx`: 15/14 of the screen at 7 px, 15/16 at 8; milk-street-office's Monthly still prints A1:R58 on its first page, as Excel's PDF).

Export: a column's width is written back as pixels / mdw (Excel reads that as the same pixels), a row's
height as pixels x 0.75; widths and heights the user left alone are written as the file had them (xlsx-browser
`sizing-excel-parity`: unedited `<col>`s identical, a column set to 45 px is `width="5"` at mdw 9 and reads
back 45 px).

Left: a row with `x14ac:dyDescent` on the sheet and `defaultRowHeight="30"` keeps rows of one Calibri 11 line at 30 pt in Excel (18 and 12.75 do not: unexplained, not followed); text turned 90 degrees fits a few pixels lower than Excel (79 px for "Turned text"); a sheet added in the editor to a workbook whose Normal font is not 7 px wide writes its widths in 7 px digits. (Mixed fonts, merged rows, the larger column style under wrapped text, columns hidden to XFD: see "Rows, fonts and hidden axes" below.)

## Rows, fonts and hidden axes against Excel 2019, 2026-11-21

Found on the real-estate models (houseflippingspreadsheet's HOME rows of `ht="15.7" customHeight` 1 px low, rows of a
bold or a second font 1 px low in vertex42's sheets, choudhary's columns hidden to XFD shown from ANF, the
investment-property calculator's `zeroHeight` / `defaultColWidth="0"` sheets showing what Excel hides, 10.5 pt drawn
as 10); each rule then confirmed on generated workbooks opened in Excel 2019 on the VM (`Rows(i).Height`,
`Columns(i).Hidden`, `AutoFit` of 17,000 cells of 30 faces x 30 sizes x 4 styles in sheets of four Normal fonts,
GDI's `TEXTMETRIC` of the same fonts). Fixtures: `tools/excel-parity/make-row-height-fixtures.py`
(synthetic-row-heights, synthetic-hidden-axes, synthetic-row-merged); data and probes:
`tools/excel-parity/font-metrics/` (`make-font-metrics.py` writes `FontMetricsData.fs`); checks: xlsx-browser
`rows-excel-parity`, `tests/office-formats/xlsx-bridge.fsx` ("fonts:", "rows:").

| Rule (Excel 2019) | Evidence | Where |
| --- | --- | --- |
| A font size is kept to a twentieth of a point (Font.Size 10.3 stays 10.3, 8.33 is 8.35, 10.01 is 10); a file's `sz` as it is | COM: Font.Size of 13 values | `CellFormat.FontSize` a float, `Model.fontSizeAsKept` (SetFontSize, ApplyFormatPatch, VBA Font.Size, the desktop toolbar); written back as read |
| A font is measured at its pixel size, round(points x 4/3): 8 and 8.5 pt (11 px) have the same maximum digit width and row, 15.5 pt is 16 pt's 28 px | 144 of 144 fractional sizes of 27 faces as the Normal font agree with the integer size of their pixel size (6 Segoe UI Black widths are not Segoe UI's: now a face of its own) | `FontMetrics.pixelSize`, the table by pixel size (`FontMetricsData.regular`, 30 faces) |
| A row's or default's height in points is shown in whole pixels of 15 twentieths of a point, cut down - a height a twentieth short of a pixel being that pixel: 15.65 pt is 20 px, 15.7 is 21; 20.15 is 26, 20.2 27; 20 is 26, 12.7 17, 409 545 | 59 custom heights, 5 sheets of a default height of their own | `FontMetrics.pointsToRowPx` |
| Fitting a row on opening, its fonts meet at their baseline: the row is the most any cell has above the baseline plus the most any has below (a font's own line R, its descent D = GDI's `tmDescent`; under 8 pt Small Fonts'); a wrapped text of n lines is n R - D above. Arial 12 (line 20, descent 3) in a Calibri 11 sheet (20, 4): 17 + 4 = 21 px; a column of Calibri 20 under 7 wrapped lines of Calibri 11: 143 px (7 x 20 + 7 - 4) | 10,482 of 10,742 AutoFit rows of 30 faces x 30 sizes x 4 styles in Calibri 11 / Arial 10 / Tahoma 10 sheets (the tallest line alone: 9,279); ColStyle 107.25 / 47.25 / 32.25 pt | `FontMetrics.lineOf`, `SheetOps.fitRowsOnOpen` |
| Only the cells aligned at the bottom (the default) meet at the baseline; a cell centred or at the top is as high as its own text | synthetic-row-baseline: an empty Arial 12 cell and Arial 11 bold text in an Arial 10 sheet 21 px at the bottom, 20 centred, at the top or mixed (vertex42's rental cash flow); 3 wrapped lines under an empty Calibri 20 cell 63 / 60 / 60 | `fitRowsOnOpen` |
| A bold or italic font's line is its own (Arial 12 bold 21 px, regular 20; Times New Roman 8 bold 14, regular 15) | the same AutoFit rows in an Arial 6 sheet (the font's own line) | `FontMetricsData.styled` |
| A fitted row is as low as its fonts' lines - the empty cells' (columns' styles, the Normal font while a column has none) - not the sheet's own default height | FitMin30 (default 30 pt): a row of 8 pt text or none is 15 pt; FitMinCols (every column Calibri 8): 11.25 pt | `XlsxBridge` (the base fonts), `fitRowsOnOpen` |
| A cell merged across rows counts for nothing; merged across columns, a wrapped text only with its descent (Calibri 20 wrapped in a Calibri 11 row: 23 px), an unwrapped one a line | synthetic-row-merged "Merged" | `fitRowsOnOpen` |
| `<col hidden>` to XFD hides every column to XFD; `defaultColWidth="0"` hides the columns no `<col>` sizes | ColsToEnd (Columns(16384).Hidden), DefColZero | `XlsxBridge` import (hidden columns to 16384) |
| `sheetFormatPr zeroHeight="1"`: every row the sheet does not list is hidden, between its rows and past them (StandardHeight 0); a listed row with nothing is at the default, a listed row of no `customHeight` fitted | ZeroHeight: rows 3, 8, 11.. hidden, `<row r="7"/>` 15 pt, `ht="20"` 15 | `XlsxReader.SheetFormat.ZeroHeight`, `XlsxBridge` (DefaultRowHeight 0, `SheetAxes.DefaultRowShownHeight` for Unhide), written back |

The tracker's earlier "a row whose wrapped cell is merged across columns is 26.25 pt in Excel" was a misreading of
synthetic-row-fit's rows (row 16 is 15 pt in Excel, as ours; 26.25 is row 14's). Excel does not break a wrapped line
at a no-break space (Nbsp: 4 lines, with spaces 5).

On the 12 real-estate models (rows 1-40, columns 1-20 of 65 sheets, Excel's COM numbers): rows 2595 of 2600 as Excel (before: 2547), columns 1300 of 1300 (before: 1291).

Left: vertex42's Help sheets: rows of wrapped Arial 11 text in an Arial 10 sheet 1-2 px higher in Excel (59 for 3 lines, ours 57; 58 / 39 in the mortgage calculator's) and two of them a line more in Excel (59 / 135, ours 38 / 114: the wrap measure); 260 of the 10,742 AutoFit rows (mostly Roboto, Gill Sans MT, Palatino, Trebuchet MS, Century Gothic and Book Antiqua, whose line has more above or below than GDI's descent says); a face GDI lacks (Roboto, Myriad Pro) has Arial's descent, Aptos its file's; unlisted rows below the default of a sheet with both styled columns and the Normal font keep the tallest line (not the baseline sum: not measured); `x14ac:dyDescent` with a default of 30 pt (rows of 30 pt in Excel) not followed.

## Excel 2019 checks of 2026-11-10 (0.1.29, branch `sheet/excel-checks`)

How: the VM (`winvm`) had no desktop session, so Excel ran straight from SSH (`cscript` over `ssh winvm`;
it needs `C:\Windows\System32\config\systemprofile\Desktop` and the SysWOW64 one, created on the VM for
this). Excel's pictures of a range: `tools/excel-parity/export-ranges.vbs` (CopyPicture as on screen,
pasted into an empty chart and exported as PNG at 100%; or a PDF of the range); ours:
`tools/excel-parity/shoot-browser-sheets.mjs` (the web component, the sheet scrolled to the same cell).
Colours were compared cell by cell, positions by the rows and columns of ink in both pictures.

- **Open-all** (`run.mjs --keep`, `open-all.vbs`): 111 workbooks, 333 saves (as opened, a cell typed, a row
  and a column inserted). Excel refuses none of a workbook whose original it opens. It refused
  `synthetic-slicers.xlsx` itself (and so its saves): a timeline style's formats pointed into `<x14:dxfs>`;
  Excel keeps them in `<x15:dxfs>` (fixed in the fixture and the reader; the new fixture opens; its saves
  and those of the new `synthetic-chartex.xlsx` were not opened yet: the VM went off the network). Differences in used cells / characters are the check's own edits (18 characters typed, a table's
  new column named `Column1`), as before.
- **Table styles** (`synthetic-table-styles.xlsx`, 11 tables): 80 of 272 cells had another fill than
  Excel's, every one a tint 1-2 off in a channel (`#DAE3F3` for Excel's `#D9E1F2` ...). Tints are now
  Excel's integer HLS (`Xlsx.applyColorTint`, 250 of 250 against Excel, `tests/formula/excel-2019-tints.tsv`):
  0 of 272 differ. Line colours cannot be compared exactly in Excel's picture (its lines are blended); their
  places are the same. pandas-test_decimal (Medium7) and the dashboards' tables: same.
- **Slicers**: margins measured on Excel's picture: buttons and header line 10 px in from the frame (were 6),
  the line 27 px down (25), a button's text at its left 8 px in (was centred); now within 1 px. Left: the
  header's two buttons (multi-select, clear filter) are not drawn; the caption and button text are wider in
  Chromium than Excel's Calibri ("Region" bold 48 px vs 41) - the browser's font, not the layout; the
  dashboards' Century Gothic slicers ("Centr..." in Excel) fit their text in ours.
- **Timeline**: was a strip with labels under it, 40 px a month; Excel puts the year and each month's label
  above the block, at its left, 48.5 px a month, blocks 96 px down; a custom style draws only its own
  elements. Now within 1 px vertically (caption, selection, year, months, strip 96-111 px both). Left: the
  scroll bar is a plain track (Excel's has arrows and is 4 px lower); the clear-filter button is not drawn;
  ours draws the last, partly shown month to the frame's edge (Excel stops 15 px before it).
- **Waterfall** (`synthetic-chartex.xlsx`, Excel's own; the e-commerce "Waterfall Chart"): connector lines
  were `#A6A6A6`, Excel's are light grey as the gridlines (`#D9D9D9`, 0.75 pt); legend 8 px and plot top 24 px
  too high (Excel: legend 50 px, axis top 78-80 px down at 480 x 288 px; ours now 49.6 and 77.6). Labels
  outside the bars: same place. Left: Excel's plot bottom is ~5 px higher; its data labels are a little
  smaller.
- **Funnel** (`synthetic-chartex.xlsx`): bars and labels in place (bar left 57 px Excel, 60 ours). Left: our
  bars are 46.5 px of a 49 px pitch (gap ~2.5 px), Excel's 44.5 of 48.75 (gap ~4.25 px) with the same
  `gapWidth` 0.06; Excel draws a light vertical axis line at the bars' left, ours none.
- **The e-commerce dashboard's "Category wise Profit" waterfall** was an empty frame until its data sheet was
  visited (a chartEx part keeps no copy of its points; the sheet was not converted): now drawn on first
  display.
- Found, not fixed (the grid's cell painting is another branch's): in the web component a number in a cell
  of General alignment is drawn at the cell's left (`text-align: start`); Excel puts it at the right
  (every number of the table-styles and chartEx sheets).
- **Formula oracle** (`dump-values.vbs`, `compare-values.fsx`) after the sort/recalc changes of 0.1.29:
  50480 of 50496 cells agree, the 16 others XIRR within Excel's stated accuracy, 77 left out, 21 workbooks
  (one more than at the last run: synthetic-table-styles' 21 total-row formulas). No new disagreement.

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
Chain"). Slicers are drawn since the merge after 0.1.28 (caption, buttons in their
columns, selected / no-data looks from the slicer style; see the roadmap's
backlog for what is approximated). Since 2026-11-22 the in-cell sparklines and the pivot tables' styles and
buttons are drawn ("Rich grid objects" above); NOT yet rendered: images/logos.
Reaching parity needs: sparklines, image embedding, and
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
- Compared with Excel 2019's picture of `synthetic-table-styles.xlsx` (2026-11-10, see above): every
  cell's fill is Excel's since tints are worked out in Excel's integer HLS (`Xlsx.applyColorTint`).
- Filter button rendering is native PNG only; PDF and Fable/web paths ignore it.

Checks to add:

- [x] Light/Dark table style synthetic fixture (`tests/test-documents-excel/synthetic-table-styles.xlsx`,
  `tools/excel-parity/make-table-styles-fixture.py`; `tests/formula/stage54-tablestyles.fsx`).
- [x] Table with explicit cell styles overriding table style (stage48: a cell's own fill wins,
  black text is the default font's).
- [x] Excel's picture of `synthetic-table-styles.xlsx` against ours (fills: 0 of 272 cells differ).
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
