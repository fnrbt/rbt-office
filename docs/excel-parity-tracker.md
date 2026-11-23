# Excel Parity Tracker

Last updated: 2026-11-23 UTC.

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

**Oracle provenance:** The measurements below were collected with Page Break Preview before PDF export. Preview can
refit `bestFit` columns and its breaks can differ from a fresh direct export. `probe-print-pages.vbs` now exports
the PDFs before activation or inspection, then reopens the workbook for separately marked preview breaks and widths.
The best-fit rules below describe preview-induced layout, not necessarily direct-export behavior.

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
previous round's 47 - with Excel's PDFs of that round; more pages compare now that they pair). Sheets paged as Excel: 183 of 186 at that round (before 181): Annual Cash Flow and Rental (AutoFit columns) differ;
oyabayo's last page no longer differs (Print round 4, below).

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
| A fresh direct PDF export uses the stored column widths, including for `bestFit="1"`; a number that does not fit may print as hashes. Page Break Preview is a separate, potentially mutating operation: it can refit numeric best-fit columns, changing their widths and the subsequent PDF's page breaks. Browser-visible preview pagination is not a direct-export oracle. | `synthetic-print-pages-bestfit-refit`: direct NarrowTrigger prints three wide values on page 1 and the narrow value on page 2; preview refits 63→63.75pt and 5.25→14.25pt and prints two values per page. Direct and preview word positions are in `bestfit-refit-direct-reference.tsv`. | flag not imported; widths stored unchanged | `XlsxReader.ColumnInfo.BestFit`, `XlsxBridge`; direct `Pagination.paginatePrinted`, preview `Pagination.paginatePreview` |

The public `bestfit` fixture's `BestFit` sheet also distinguishes overflow: B2 (`bestFit`) and D2 (ordinary) both
contain 123456789 in 5-character columns. Fresh direct export prints hashes in both; after preview B2 expands from
26.25 to 64.5 pt and prints the number, while D2 stays 26.25 pt and prints hashes.

Verified not reproduced: oyabayo's last page (rows 74-116 x column CD) - both Excel's direct PDF and ours have 39
pages with the same starts, and the page is blank in both; see "Print round 4", below.

Found, not matched:
- Text placement: Segoe UI and Segoe UI Black a pixel lower than Excel at the bottom (their descent in the GDI table
  a pixel short of what Excel places them by): acre-property's words 0.7 pt low; at the top Trebuchet 1 px.
- oyabayo's right-aligned *italic* numbers stand 2.8 - 3.2 pt right of Excel's (Aptos Narrow 14/11; the round-3 figure
  "4.6 pt" is Excel's own word-box inset from the cell border, not our delta - see "Print round 5", below). Upright
  numbers and currency pair within 0.03 - 0.36 pt. The offset is a placement rule, not glyph ink: a font-metric
  formula (italic angle x ascent) fits the two faces it was derived on but is refuted out of sample (Arial, Times,
  Cambria, hashes, trailing `_x`); no product code changed.
- The PDF's text of a glyph two characters share is the other one: Calibri's hyphen is U+2010 in our PDFs' text
  ("31‐Jan‐16"), from Rbt.Pdf's ToUnicode (its extra glyph mappings do not override the cmap's); words differ.
  Fixed 2026-11-23: the writers now emit per-source-sequence CIDs, so U+002D and U+2010 keep their own
  ToUnicode entries while `CIDToGIDMap` draws the one glyph. The first attempt still failed because the
  office repo pinned `Rbt.OpenType.Core` 0.1.3 while `Rbt.Pdf` 0.1.25 is built against 0.1.9, so the
  subset call threw and every font took the old glyph-ID fallback; `RbtOpenTypePackageVersion` now pins
  0.1.9. Subset failure also retries the full font with the same CID mapping, so only fonts with more
  than 65,535 distinct `(glyph, source)` pairs lose colliding-character identity.
- Fonts the machine has no file of (Aptos, Book Antiqua, Century Gothic, Myriad Pro) are set in Calibri.

### Direct-PDF numeric overflow, 2026-11-22

Fresh Excel 2019 direct exports, not Page Break Preview, of `synthetic-print-overflow.xlsx`
(`make-print-overflow-fixture.py`, oracle `print-overflow-direct.excel.tsv`): a
Calibri 11 `$1,234,567` prints as eight hashes at column width 9.50 and as the
value at 9.75; a holdout `0.00%` value `1234567.89%` hashes through 11.50 and
prints at 11.75. Arial 10 has the same transitions. General `1234567` prints
at width 8.00. Empty versus occupied neighbours make no difference. At 9.50
the Calibri cell is 53.16 pt wide and the printed currency glyphs take 50.53
pt; their 2.63 pt remainder is smaller than the existing five-pixel cell
inset (3.75 pt), while the 9.75 cell's 54.84 pt leaves 4.31 pt and fits.
That same printer-space comparison, with the font measured at 600 dpi,
explains the holdout; using screen GDI widths or a generous PDF glyph-only
check does not.

Formats matter independently of the number: a trailing `_)` delays the full
value through width 10.00 (10.25 fits) but does **not** reduce the number of
hashes. A leading `_(` in an accounting format occupies one hash position
before overflow (one fewer hash at a given width), even without `*` fill.
At 75% print scale both the column and printer advances change: some
thresholds and hash counts differ from 100%, so the PDF fit measures its font
at the page's scale before choosing text. The browser grid's screen fitting
rule is unchanged.

The remaining printer-rounding uncertainty in the public occupied-row controls:
at 100% width 10.50 the accounting format prints eight hashes here versus
Excel's nine, and the percentage holdout prints nine versus Excel's ten.
At 75% the trailing-`_)` format at width 10.125 shows the value here where
Excel prints nine hashes (both show the value at 10.25). The other 117 of
120 occupied controls match, with no value erroneously hashed. Do not change
the browser's screen rule or add a workbook-specific width allowance to
conceal these narrow print-rounding differences.


On the private apartment acquisition workbook's Annual Cash Flow direct PDF,
the 27 million-dollar values in G6:O6, G8:O8 and G12:O12 print as nine
hashes (Year 3 through Year 11); the former native export printed all 27
values. The cells are numeric, Calibri 11 bold, number format
`"$"#,##0_);\("$"#,##0\)` and width 10. They are not text-overflow
candidates; their printed 50-point glyphs plus the format's trailing space
do not fit the column with its cell inset. This workbook remains local-only.

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
| A linked picture (camera) inside a group is a part of the group: the group's shapes drawn, the picture of its range in its own rectangle | dashboard-ecommerce-sales' Dashboard (Excel 2019's picture): the four KPI cards with their small YOY boxes | the whole group drawn as one picture of the range (the cards covered by a 900-px "▲ 20.6") | `cameraRange` (not for a group), `DrawingParts.Camera` |
| An ActiveX control's picture (EMF): pattern brushes (CREATEMONOBRUSH: 0s in the text colour, 1s in the background colour - a pressed toggle button, scroll bar tracks), a mask then a picture painted through it (SRCAND, SRCPAINT: option buttons), clipping (INTERSECTCLIPRECT, EXTSELECTCLIPRGN, SELECTCLIPPATH; saved with the DC), opaque / clipped text (ETO_OPAQUE, ETO_CLIPPED) at each character's advance, hung from its top by the font's Windows ascent (Calibri 15 px: 14), paths (BEGINPATH .. FILLPATH / STROKEPATH / STROKEANDFILLPATH), POLYLINETO, POLYBEZIER(TO), POLYPOLYLINE, MODIFYWORLDTRANSFORM | synthetic-activex.xlsx: every control as Excel shows it but a list box's selection (Excel draws the live control, without its highlight; its picture has it) | none of these (grey toggles, black option rings, text 3 px high); a BITBLT's source read a field early | `Emf.svg` |

Left: text at zooms other than 100 % within the grid's own rounding of cell edges (1 px either way); Segoe UI Black's
bottom line a pixel lower than Excel's; a distributed line stands 2 px higher in Excel than the row's other lines; a
row's descent counts every bottom-aligned cell of it, not only those drawn; a fill's corner pixel at the right of its
top grid line; combo / list box text in the browser's grayscale, not ClearType; the real-estate models' first screens
were shot again (53 sheets; on the sheets the header-aligned diff lines up, acre-apartment's and acre-basic's, 0.2-1.1
points fewer differing pixels; others not aligned by it), of the dashboards only dashboard-ecommerce-sales' Dashboard looked at again (its KPI values $22,97,200.86 where Excel
shows $2,297,200.86 - format [$$-1009]#,##0.00 -, the camera's text a darker green than Excel's, the cards' corners
square where Excel's are rounded: not looked into); the old note that left-aligned text starts 2 px further right in Excel was the 1-px box offset
above (synthetic-cell-indent.xlsx needs no rule of its own).

## Dashboards: linked text, format colours, shape outlines, shadows, lines, hidden rows against Excel 2019, 2026-11-22

The three items left on dashboard-ecommerce-sales' Dashboard by the round above, then a re-survey of the workbooks with
shapes: the house-flipping model's five shown sheets (shapes with shadows on all of them) and dashboard-sales-analysis' Dashboard,
each shot in Excel 2019 on the VM (`shoot-sheets.ps1 -KeepZoom`; the house-flipping sheets need the `Goto` in a `try`)
against ours (`shoot-browser-sheets.mjs`). Each rule from a generated workbook checked in Excel: `make-shape-looks-fixture.py`
(synthetic-shape-looks.xlsx: "Links" - cells in locale-tagged formats and shapes linked to them with stale saved text;
"Colors" - [Black] .. [Cyan] and [Color1] .. [Color56]; "Rounds" - rounded rectangles of 8 sizes / adjustments,
lines of 0.75-3 pt, a group with a shadow and a shadowed shape, connectors with arrowheads of every kind and size and
no height, ellipses, a theme effect and its override, a shape of nothing; "Hidden" - shapes over hidden rows),
pictures of it, and a probe of every shape's text, Adjustments, Line.Weight, Top / Height and every cell's
`DisplayFormat.Font.Color`; a probe of 43 formats x 4 values with `Range.Text` and `TEXT()`. Checks: xlsx-browser
`shape-looks-excel-parity`; tests/formula stage27 (locale tags, colours), stage52 (shapes).

| Rule (Excel 2019) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| A locale tag (`[$$-1009]`, `[$-4009]` en-IN, `[$-439]` hi-IN, 1009, 409, 809, 410, 407, 40C, 411, 804, 449, 44D, 44E, 445, 420, 463, 461, 45B, 468, `-1004009`, `-404009`) never groups by lakhs: `#,##0.00` is thousands whatever the tag (`#,##,##0` is read as `###,##0`); `[$-2000000]` gives Arabic-Indic digits, `[$-F400]` the system time format (neither matched) | 2297200.86 in each: $2,297,200.86 / 2,297,200.86, by `Range.Text` and `TEXT()` alike | (already so) | the dashboard's "$22,97,200.86" was the text saved with the shape (the workbook last saved on an en-IN machine), not a format rule |
| A shape linked to a cell (`textlink`, a `TxLink` field) shows the cell's text as it is now, the number in its format whatever the column's width (no `####`), from any sheet | "Links": eight shapes whose saved text is "$22,97,200.86" / "stale": $2,297,200.86, -$98,765.40, ..., 1,234,567.89 (a 3.7-wide column), $286,397.02 (another sheet) | the saved text when the cell's sheet was not converted yet (the dashboard's KPI sheet) | a one-cell read of a sheet not converted converts it (`chartBlock`) |
| A line feed in a shape's text (a linked cell's CHAR(10), a field's saved text) breaks the line as `a:br` | dashboard-sales-analysis: the pie's labels on three lines | run together, wrapped by width | `shapeItems` |
| A number format's colour names are the palette's first eight - [Green] #00FF00, not #008000 - and `[ColorN]` the palette's Nth (N 1-56) | "Colors": `DisplayFormat.Font.Color` of all 64; the dashboard's "▲20.62%" (`[Green]\▲0.00%;[Red]\▼0.00%;`) bright green | [Green] #008000; `[ColorN]` not read (and left in the text) | `CellText.colorTagValueAt` |
| A shape's outline is centred on its edge, half of it outside the shape, `round(w)` whole pixels (at least 1); an odd one covers the pixel of the edge (1 px: the pixels of the left and right edges; 2: the pixel before too; 3: one each side; 4: two before, one after); rounded corners `adj` of the shorter side (default 1/6) | "Rounds": 0.75 / 1 / 1.25 / 1.5 / 2 / 3 pt lines 1 / 1 / 2 / 2 / 3 / 4 px, every edge pixel as Excel's; 8 rounded rectangles' corners pixel for pixel | clipped to the anchor (half the outline gone - 1 px lines grey, the cards' 2-px lines 1 px, their corners looking square), fractional | `sceneSvgSpilling` (shapes, connectors and groups draw past their rectangle; charts and pictures not), outline half a pixel over for odd widths; frozen panes: a margin round a part where the drawing ends in it |
| Outer shadows (`a:outerShdw`): a shape's own effect list (an empty one: none), else its style's `effectRef` (the theme's effect style of that index; Office's default theme's third: 1.5 pt down, 4.5-pt blur, 63 %), a group's for its shapes; offset `dist` at `dir`, a Gaussian of deviation `blurRad` / 2, the colour at its alpha | the shadowed card: right of its 2-px line #AAAAAA #B5B5B5 #C6C6C6 #D7D7D7 #E4E4E4 #F1F1F1 #F8F8F8 #FCFCFC #FEFEFE, ours #ABABAB #B8B8B8 #C7C7C7 #D6D6D6 #E5E5E5 #F0F0F0 #F8F8F8 #FCFCFC #FEFEFE; the theme effect all round, its override none; the dashboard's KPI cards | none | `PShadow` (`feDropShadow`), `outerShadowOf`, the default scheme's effect styles |
| Arrowheads (`headEnd` / `tailEnd`): `w` and `len` sm / med / lg 2 / 3 / 5 times the line's width, the line taken as at least 2 pt; triangle, stealth (notched), arrow (open), its tip at the line's end; oval and diamond centred on it | "Rounds": a 1-pt line's med / lg triangle 8 x 13 px, lg / lg 13 x 13, sm / sm 5 x 5; a 2.25-pt line's med / med 9 x 9 | none | `arrowhead` |
| A line of no height (or no width) is drawn; a flow chart connector is an ellipse | "Rounds"; house-flipping's progress line and arrow on three sheets, its step circles | not drawn (an SVG of no height draws nothing; drawings under 1 px skipped); a rectangle | lines kept, their SVG a pixel at least |
| A shape of no fill, no line (none given and no style) and no text draws nothing | "Rounds" "Empty"; house-flipping's clickable rectangles over cells on four sheets | a placeholder frame with its name ("S..." over each check box, two frames over a table) | an empty scene |
| An anchor's offset in a hidden row or column is none; a span the file gives that hidden rows take away is none - the object is not shown | "Hidden": Shape.Top / Height 45 / 0, 18.75 / 26.25, 33.75 / 33.75 pt; house-flipping's picture in hidden rows not shown | offsets inside hidden rows counted (a 12-px bar), the object's own size when nothing was left | `Layout.drawingRect` |

Left: the camera's text on the dashboard's cards is narrower than Excel's (the range drawn into the picture's
rectangle); a slicer button's caption cut ("Centr...") where Excel cuts it; the map chart; text shadows (`a:rPr`
effects), `rotWithShape`, a group's shadow drawn for the group as a whole (alike where its shapes do not overlap),
glows / soft edges / reflections; arrowheads' line not shortened under an open arrow; a pane part's margin can show
a shadow past a frozen line; a camera's own outline; whether an offset larger than its (shown) cell is cut to it (not
looked into: kept as the file has it). Converting the sheet a one-cell reference reads also converts the
sheets a chart's one-cell series names read (sooner than before; the chart's values from its cache until then); an
export keeps the calculation chain for such a sheet still standing in for the file's.

## Slicer text, cameras, Pareto, sunburst labels, filled maps against Excel 2019, 2026-11-22

What the earlier rounds left on the dashboards and the chartEx zoo, each from generated workbooks made by Excel 2019 on
the VM (`probe-on-vm.sh`) and Excel's pictures (`shoot-sheets.ps1`, Chart.Export) against ours (`shoot-browser-sheets.mjs`,
`render-charts.fsx` + `shoot-svgs.mjs`). Fixtures and probes: `build-slicer-cut.vbs` (synthetic-slicer-cut.xlsx and
-twcen.xlsx: one table cache shown by 400 slicers 1 px apart, Calibri and a Tw Cen MT theme), `build-missing-font.vbs`
(synthetic-missing-font.xlsx and -twcen.xlsx: cells in Lato, a made-up face and installed ones; linked pictures of them),
`font-metrics/probe-font-faces.vbs` (Lato and Tw Cen MT as Normal fonts at 6 - 48 pt, their cells' rows),
`build-pareto-layout.vbs` (synthetic-pareto-layout.xlsx: Pareto charts of labels 1 - 7 figures wide),
`build-sunburst-labels.vbs` (synthetic-sunburst-labels.xlsx: leaves 2.75 px apart in width, leaves of 1 - 12 beside a large
one). Checks: tests/formula stage56-slicers, stage52-charts, stage55-chartex; xlsx-browser `slicers-in-grid`,
`camera-excel-parity`, `map-charts-excel-parity`.

| Rule (Excel 2019) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| A slicer's text is in the theme's body font (Tw Cen MT in the e-commerce dashboard), at 15 px by GDI's widths | the Tw Cen book: every button and caption cut by Tw Cen MT's widths | Calibri | `Slicers.drawer` (`PFace`), `slicerFitted` |
| A button's text is whole where it is no wider than the button less 18 px; else its most characters (spaces kept: "Ab ...") and three full stops "..." narrower than that; where even "..." is not, as many stops as are ("..", ".", none). The caption likewise in the slicer less 72 px | 400 slicers 1 px apart: each of the 1600 texts (3 buttons and the caption, 2 fonts) one ink width a text of ours and the other way round; the dashboard's "Centr..." in a 61-px button ("Central" 44 px of Tw Cen MT, room 43) | cut at the button less 12 px by Calibri's widths with "…" | `Slicers.slicerFitted` |
| A slicer's buttons and header line from 9 px inside the frame at the left to 9.67 at the right (8 or 9 whole pixels by where the anchor falls), columns 3 px apart, each button on whole pixels rounded half down (108.25-px buttons 108, 108, 109, 108) | the cut fixture (button w - 18 px), synthetic-slicer-looks' "Two columns" (123 + 3 + 123), the dashboard's 61 / 108.25-px buttons 3 apart | 9 / 9 px, 4 apart, fractional | `rightMargin`, `columnGap`, `snap` |
| A table's slicer on another sheet than its table shows its column's values (that sheet read first) | the cut fixture's B1 .. C2 sheets of the Data sheet's table | no buttons in the browser | `objectSvg` draws again once the sheet is converted |
| Lato (an Office cloud font: %LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts) is drawn by Excel in Lato: as wide as its own advances, its rows as Lato's (11 pt: 24 px); a face no one has is Arial | synthetic-missing-font's Fonts sheet; the e-commerce KPI YOY cells (Lato bold) and their cameras' YOY boxes | the browser's fallback (narrower, taller), rows of Arial | Lato in artifacts/fonts/windows-extra (copied from the VM), `GdiAdvances` (Lato, regular and bold), `FontMetricsData` (Lato and Tw Cen MT: rows and widest figure as a Normal font, bold / italic rows; descents from winDescent); "Tw Cen MT" no longer measured as "Tw Cen" (Arial) |
| A linked picture's cell text is cut where the grid cuts it: at the next cell that is not empty, on the side it is aligned to; numbers and wrapped text keep to their cell | synthetic-missing-font's Camera sheet: "Hamburgevo|" at C2 | running over its neighbour | `cameraSvg` clip paths |
| [Green] in a number format is #00FF00 | the dashboard's KPI YOY cells and YOY boxes (#00FF00 ink) | #008000 | `CellText` colour tags |
| A chartEx chart with a second value axis (Pareto) keeps 12 pt at either side, not 6; a value axis' labels (GDI widths) stand 6 pt from the plot, not 7.5 | Pareto plots from 36, 42, 49, 55, 61, 67 to 428 px (labels from 17 px, the right ones to 462); every histogram and box chart of the zoo from Excel's left pixel | 6 px wider; histograms 2 px right | `layoutClassicChart` `sideMargin`, `valueGap` |
| A sunburst label is broken to its ring by the design advances (no extra room) and left out where its lines and 4.5 pt are longer than its part's arc at the middle radius | "Osaka" 29.6 px whole in a ring of 30.25, "0000iii" 32.5 broken; "Osaka / J" in 37.4 px of arc kept, "Osaka / I" in 33.6 left out; Sunburst3's "Ameri / ca" (42.3) kept, Italy (14.8) and "Madri / d" (31.2) left out | GDI widths and 2 px more ("Osak / a"); left out only where the lines alone did not fit | `layoutHierarchyChartEx` `brokenTo` |
| A filled map (regionMap) is drawn from its part's geography cache (`cx:geoCache`: base64 of raw deflate; regions' outlines in Bing's point compression): the country's divisions, those of the data in the colour of their value (a scale of the part's own - min, mid at the values' middle, max -, else 80 % lighter to 75 % of the colour style's first colour, in proportion in sRGB), the others #E0E0E0, outlines in the series' line colour (else white); the United States as an Albers map with Alaska (0.35) and Hawaii in insets; the legend (the series' name, the scale's bar, its ends' or middle value) at the right; the provider's credits under the map | e-commerce "Map Chart": California #1482AC, New York #51A5C5, Texas #8CC7DD, the least #D2EFFA, Alaska #E0E0E0; "Sales by State": Texas #679491 (three colours, "$ 229.30K" the middle); sales-analysis' Australia: every state's grey (#000000 .. #C6C6C6) | a placeholder "Map chart (not drawn)" | `RegionMapPlot`, `Inflate.raw`, `decodeMapRing`, `layoutRegionMapChartEx`, `withUserShapes` (colour style) |

Left: the slicer's header buttons of narrow slicers (Excel leaves the Multi-Select out under ~47 px) and a timeline's
caption cut (not looked into); a button's whole pixels with an anchor at a fraction of a pixel (ours from the snapped frame:
a pixel either way); a camera's text in a Lato row stands 4 px lower than Excel's (the row's baseline in a picture), the
dashboard's YOY boxes 1 px narrower; Pareto plots of 4 figures and more and histograms of one start a pixel off Excel's;
sunbursts of the 376 x 228-pt chart leave out labels Excel's 360 x 216 one breaks (not understood); a sunburst label on a
light part is black in Excel (Sunburst's "Bread" / "Bakery" on yellow), white in ours. Maps: Excel 2019 on the VM was
online (Bing answered), so that it draws these without the network is the cache's design, not checked by cutting the VM
off (shared with another agent); the lower 48 states are about 4 % narrower in ours than Excel's projection, Australia
wider (Mercator; Excel's projection of other countries not found); the legend's value labels are 7.5 pt (the Map Chart
sheet) where the dashboard's are larger, its bar 6 px wide where the dashboard's is 8; the three-colour legend shows the
largest and the middle value only (as the one Excel picture does); a world map or one of several countries is not tried;
the map tests use the dashboards (no synthetic map: Excel makes one only with Bing).

## Formula oracle on real-world models, 2026-11-21

Eleven real-world financial models (local only, private-use licences: not in the repo) against Excel
2019's full recalculation (`dump-values.vbs`, `compare-values.fsx`), and - for the two that need Excel
365 (`@`, SEQUENCE, XMATCH, dynamic arrays) - against the values the file keeps (`--cached-new`).
The original run agreed on 86,045 of 97,011 compared cells. Replaying the current engine against the
same references gives 100,201 of 100,235 within `1e-9 * max(1, abs(reference))`. The 34 outside
that band are five waterfall results (Excel 2019 recalculation) and 29 iterative leasing-model
results (saved values). All 34 pass the comparator's looser band; 586 cells were skipped. Neither
band tests displayed rounding or investment thresholds. The separate public corpus had 16 XIRR
differences out of 50,496 compared cells.

**Impact, not just tolerance.** In the waterfall, two unrounded XIRRs differ by
3.55e-9 and 5.71e-9 in rate units. On the same dated flows, the discounted balance at the
Excel rate misses zero by about 19 and 28 cents; at the engine root the residual on its inputs
is below a cent. Using the other engine's cash flows at the Excel rate changes that balance by
less than a millionth of a dollar. This isolates the solver discrepancy from input drift; it
does not identify Excel's iteration algorithm. Both reported percentages display identically,
and all 58 checked rolling six-decimal returns agree. No payout or hurdle was found to depend
on the differing unrounded results. Separately, all 4,209 numeric formula results in the
iterative model round to the same hundredth as its **saved** values (not a fresh Excel run).

The [synthetic XIRR workbook generator](../tools/excel-parity/make-xirr-impact-fixture.py)
and [Excel/engine observations](../tools/excel-parity/xirr-impact-results.tsv) show where this
distinction matters: at a deliberately chosen rounding boundary, Excel gives `0.100000` and the
engine `0.100001` after `ROUND(rate,6)`. Both display `10.00%`, but a comparison can flip.
For a positive one-year return with a negative starting guess, Excel returned `#NUM!` where the
engine found a root. Neither case was the traced workbook's decision. Retain the solver; a
constant bias, rounding all XIRRs, or a blanket guess-sign restriction is not supported. The
2026-11-23 threshold probe below separates this stopping difference from root selection.

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
- When a new formula closes a cycle and iteration is off, Excel starts that formula at zero—even
  when it replaces a number or a calculated formula. Previously calculated dependents keep their
  values; breaking the cycle recalculates them. Typing `=B1+5`, then `=A1+2`, then `=A1+B1`
  into A1:C1 gives 5, 0, 0. Cell entry now leaves the new formula's cache empty for recalculation
  rather than storing a result computed against the pre-edit sheet. Validation still sees the
  prospective value. Browser regression: `api-circular-entry` (incremental and full calculation).
- On opening an XLSX, a formula on a non-iterative cycle (or reading it) with no cached
  value is empty text in Excel 2019, whether `<v>` is absent, `<v/>` is empty, or `t="str"` has
  an empty `<v>`; numeric caches remain numeric. Previously ours started these formulas
  from 0. The reader now distinguishes an empty string cache from an absent result, and
  the import calculation keeps only uncached formulas held by a cycle as empty text;
  formulas newly typed into a cycle still start at 0. Acyclic uncached formulas still
  recalculate on import (Excel opens `=1+2` as 3). An unchanged export keeps
  the original package byte for byte; reopening it recalculates the same
  uncached formulas. After an edit, saving writes the calculated caches,
  including empty-text cycle results. Native import/roundtrip and browser
  `xlsx-open-empty-formula-caches` cover both export paths and breaking a cycle.

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

Matched (Excel 2019): multiplication, division and exponentiation discard a **subnormal result** to zero,
not a subnormal operand. Let `A1=2^-1022` (minimum normal), `A2=A1*1.5-A1=2^-1023`, `A3=-A2`:
`A1*0.5`, `A1/2`, `A2*1`, `A2^1`, `2^-1074` are zero, also with negative signs; `A1*1`,
`A2*2`, `A2/0.5`, `A3*2`, `A3/0.5` are signed minimum normals. `A2^0.5` and
`A2^-0.5` amplify to normals; `A2^-1`, `(-A2)^-2` and even `A1^-1.001` are #DIV/0!
because the corresponding positive-power denominator is subnormal, though
`A2^-1`'s mathematical reciprocal fits a double; `A2^-0.99` is finite.
`A1/A2`, `0/A2`, `A1/A3` also give #DIV/0! (subnormal divisor).
`(-A2)^0.5` and `(-A2)^-0.5` give #NUM!. `POWER` mirrors `^`; postfix `%`
discards a subnormal result. Addition, subtraction, SUM, EXP, negation,
comparison and IF keep subnormals (`SUM(A2,A2)=A1`, `A2<>0`, `IF(A2,1,2)=1`,
`EXP(-740)>0`). The same operator rule holds in array arithmetic
(`SUM(A1:A3*0.5)=0.5` for `A3=TRUE`). Excel's PRODUCT drops each subnormal
intermediate (`PRODUCT(A1,0.5,4)=0`, reversed order `PRODUCT(A1,4,0.5)=2*A1`);
SUMPRODUCT drops each subnormal elementwise product before summing: for
`C1:C3=A1` and `D1:D3=0.5`, `SUMPRODUCT(C1:C3,D1:D3)=0`.
Plain numeric and cached-formula XLSX `<v>` values below minimum normal (including
`2.225073858507201E-308` and negatives) open as zero; exact `±2.2250738585072014E-308`
and larger normals survive; `<v>0.30000000000000004</v>` still keeps its binary value.
Implemented in `Formula.applyBin` / `applyBinGeneral`, `POWER`, `PRODUCT`,
`SUMPRODUCT` and their numeric array/stream folds, plus `XlsxReader.Xlsx.cellValueFromParts`.
Symbolic MAKEARRAY simplifications with tiny coefficients decline to its ordered
numeric stream; no global numeric/text/style parse flush. Entered formula text,
VALUE and typed cell input follow separate rules; do not apply the stored-`<v>` rule.
Separate Excel 2019 COM-entry observation (not imposed on the UI): assigning
`1E-310` through `.Formula` stores text, assigning `=1E-310` is rejected,
and `VALUE("1E-310")` is #VALUE!. These are input/locale/parse behaviors,
not the stored-XLSX-`<v>` normalization.

### IRR / RATE / XIRR (found, not matched)

Excel's IRR is within ~1e-13 of the root, not on it, and depends on the guess (`-1000,300,400,500`: guess
0.1 -3.5e-17, 0.05 -3.6e-13, 0.5 +2.5e-14), also for two cash flows (results 1.2 +- an ulp minus 1). No
Newton / secant / finite-difference variant in r, 1/(1+r) or 1+r (analytic or numeric derivative, 3 x 4
x 7 step sizes) reproduces the digits on 120 + 120 probes; the errors are close to Newton in 1/(1+r)
(about 1.15 times) without being it. XIRR is coarser: within ~1e-8 relative, the result the guess plus
steps of a power of two (guess 0.1 and 0.2 give the same digits, `-1000, 1100` a year later is
0.0999999940395 = 3355443 x 2^-25 with guess 0.1, 26843545 x 2^-28 with 0.5 or 1).
A stepping search (step-halving, bracket and bisect, 6 starts x 10 tolerances) did not reproduce it.

The [guess-domain probe](../tools/excel-parity/make-xirr-domain-probe.py) compares 25 cash-flow
series at 29 guesses in Excel 2019 ([results](../tools/excel-parity/xirr-domain-results.tsv)).
All 50 guesses of −1 and −2 returned `#NUM!`, including cases with a valid root; `IFERROR`
took its fallback. Reject guesses at or below −1 for XIRR, without changing IRR or XNPV.
Within the valid domain the selection rule remains open. A tiny positive root can return a
number with a negative guess. Of two annual series with two sign changes, one with roots
10% and 20% returned `#NUM!` even at a 15% guess; one with roots of opposite signs returned
the positive root at the default guess. Neither rejecting all multiple-root series nor choosing
a root by the sign of the guess matches Excel.

### XIRR six-decimal boundaries: stopping, not selection, 2026-11-23

The public [impact probe](../tools/excel-parity/xirr-impact-results.tsv) has one boundary pair
(`-1000, 1100.0005` one year apart, exact root 0.1000005) whose `ROUND(...,6)` differs: Excel returns
0.100000494718552 (`0.100000`), the engine the root (`0.100001`); both guesses reproduce it. Its other
18 numeric rows agree after rounding; at the boundary the rounded hurdle test also flips
(`reject` vs `accept`). The [guess-domain probe](../tools/excel-parity/xirr-domain-results.tsv)
has **no** same-root six-decimal flip in its 725 direct XIRR outcomes: of its 441 outcomes where both sides
return a number, 16 differ by more than 2e-8 and also by `ROUND(...,6)`, and all 16 are root selection (the two
annual series with roots 0 and 0.5, and -0.366025 and 1.366025); 207 more cells are Excel `#NUM!`
where the engine returns a valid root, mostly for guesses on the wrong side of zero for the
one-root series. Its rounded disagreements are therefore selection, not a threshold effect.

The [threshold probe generator](../tools/excel-parity/make-xirr-threshold-probe.py) and
[Excel/engine observations](../tools/excel-parity/xirr-threshold-results.tsv) test this directly
(the engine side by [`xirr-cells-engine.fsx`](../tools/excel-parity/xirr-cells-engine.fsx)).
Its 18 public synthetic cases were solved in decimal to exact roots and their Excel/engine
`ROUND(...,6)` results predicted in `CASES` **before** their measurement (`pred_*_ok` columns in the
results). Eleven single-root one-year rows move the root through 0.1000004, 0.10000049, 0.1000005,
0.100000503, 0.10000051 and 0.1000006 (plus the negative mirror, a 1e6-scale copy and one repeat at
guess 0.5); the rest are multi-root selection cases and two held-out shapes (an irregular three-flow
and a 13-flow monthly waterfall, each with its last flow solved to a root just above the boundary).

| case | exact root | Excel 2019 | engine | `ROUND(...,6)` | cause |
| --- | --- | --- | --- | --- | --- |
| T3 at boundary | 0.1000005 | 0.100000494719 | 0.10000050000000013 | 0.1 vs 0.100001 | stopping (same root) |
| T10 at boundary, 1e6 scale | 0.1000005 | 0.100000494719 | 0.1000005 | 0.1 vs 0.100001 | stopping (same root) |
| T4 +3e-9 | 0.100000503 | 0.100000506639 | 0.10000050300000007 | 0.100001 both | agree (predicted down) |
| T4 +3e-9, guess 0.5 | 0.100000503 | 0.100000504404 | 0.10000050299999991 | 0.100001 both | agree (predicted down) |
| T7 negative boundary | -0.1000005 | -0.100000499561 | -0.10000049999999996 | -0.1 both | agree (predicted -0.100001) |
| M3 irregular +3e-9 | 0.100000503 | 0.100000506639 | 0.10000050299999964 | 0.100001 both | agree (predicted down) |
| M4 monthly +3e-9 | 0.100000503 | 0.100000506639 | 0.10000050299999966 | 0.100001 both | agree |
| M1 two positive roots | 0.100000503, 0.3 | `#NUM!` | 0.10000050299999938 | `#NUM!` vs 0.100001 | Excel rejects a valid root |
| M2 opposite-signed, guess 0 | -0.100000503, 0.3 | 0.299999995117 | -0.10000050300000012 | 0.3 vs -0.100001 | root selection |

Only the two cases whose exact root is the decimal half-way value flip, and the flip follows
Excel's stopping point rather than a bias: near this boundary Excel's numeric answers take a few
values (0.100000494719, 0.100000504404, 0.100000506639, 0.100000602007) with errors of -5.3e-9,
+1.4e-9, +3.6e-9 and +2.0e-9 against their roots. The same input at guess 0.5 stops at
0.100000504404 where the default guess stops at 0.100000506639, and T4/T5 share 0.100000506639
although their roots differ by 7e-9, so the error is neither a constant nor a function of the root;
the pre-registered negative-bias prediction failed on T4, T4 at guess 0.5, T7 and M3 (and the engine
prediction failed on T7 and on M2's default guess, where Newton finds the positive root like Excel).
Reproducing the flip would need Excel's exact iterate and stopping rule,
which the 120 + 120 IRR probes and the XIRR stepping search above did not reproduce; a constant offset
or a snap to the observed grid would falsify T4/T7 and replace a root good to about 1e-15 (the tests
allow 1e-12) with one off by ~5e-9. **Keep the solver.** The [`Formula.rootOf`](../src/Spreadsheet/Formula.fs) candidate ordering
is also unchanged: M1/M2 show Excel can reject a valid root or pick the other one, and its choices are not
consistent enough for a general selection rule (two positive roots always `#NUM!`; opposite-signed roots
positive at guess 0 but `#NUM!` for negative guesses). The new `stage12-financial` cases
(`xirr-boundary-root`, `xirr-boundary-round6`, `xirr-neg-boundary`, `xirr-multiple-positive-root`,
`xirr-opposite-sign-guess0`) pin the accurate root and keep valid roots alive: 58 pass, 0 fail on
2026-11-23. No solver source changed.

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

## Print round 3: validation and borders

- Excel 2019's `synthetic-print-pages-validation-1.pdf` has no list-validation
  dropdown arrows. Static PNG/PDF cell painting no longer draws them; the
  selected-cell interactive control and table filter indicators are separate.
- In `synthetic-print-pages-borders-1.pdf`, hair, thin, medium and thick
  solid rules measure 1, 8, 16 and 24 dots at 600 dpi (0.12, 0.96, 1.92
  and 2.88 pt). PDF painting uses these widths; screen/PNG geometry is unchanged.
  Dotted, dashed and dash-dot patterns use measured 600-dpi mark/gap lengths,
  with page-origin phase and continuity across cell edges. Hair remains one
  physical dot under scaling. Double rules are each 8 dots with an 8-dot gap.
  Slant rules use clipped sheared marks; at 100% dotted gaps paint white,
  matching Excel's opaque tile. Double-border joins and exact printer-dot
  positioning remain open.
- Excel 2019 COM reports built-in formats 37–40 with `_ )` padding (without
  the space: `_)`) in the positive section. Import and save now preserve that
  padding; the visually similar unpadded codes remain custom formats.

## Print icon sets and sparklines against Excel 2019, round 3 (2026-11-22)

Evidence: synthetic `bf4` icon-set and sparkline-look Excel PDFs (not checked in). Icons are embedded at 24×24 pixels/144 ppi, i.e. 16×16 screen px; their positions agree with the existing browser grid's 2-px left inset and vertical alignment. Each sparkline is embedded as a 447×80-pixel image at about 200 ppi across a printed cell that was 200×40 px on screen: plot in screen-cell coordinates, then stretch to its printed row/column size. The original grid geometry leaves 3/2 px horizontally and 5/4 px vertically for the plot.

Shared geometry: `ConditionalFormatting.iconArtwork` supplies SVG and vector-native/PDF paths; `Sparklines.inCell` supplies range parsing, source values and axis bounds for the grid and print. The PDF and native PNG runners resolve explicit, quoted, cross-sheet and nonactive-sheet references against the workbook; printed overlays are drawn once, including cells without stored values. Native PNG no longer uses the legacy XML sparkline reader (which crashes on `$`-qualified cell references).

Remaining visual limits: Excel rasterizes both into PDF images, whereas ours emits vector paths. Rounded corners on some icon-set boxes/traffic-light housings and stroke caps/anti-aliasing of sparkline lines may differ at pixel level. Standalone model-only static painting cannot resolve an explicitly named *different* sheet without a workbook source callback; workbook PDF/PNG runners supply that callback.

## Print round 3: inherited columns and bottom text

- Excel 2019's `blockhold.xlsx`/`blockhold.tsv`: an inherited column fill or
  border can retain a trailing page when a visible participating row on that
  page shares an absolute 16-row block with a used-cell extent reaching the
  styled column. Source cells may be hidden or formula-empty; an empty merge
  counts. A font-only column, absent/hidden target row, or a row with only a
  nonprinting row font style does not. The existing implicit used-range and
  trailing-page trim remain otherwise unchanged.
- Printer bottom text baseline uses the larger 96-dpi GDI descent/internal
  leading of the bottom-aligned row's fonts, with a Normal-font floor. The
  600-dpi Normal metric sets the dot offset (not the screen row-height scale).
  Checked normal sizes: Calibri 8/11/20, Arial 8/10/12, Segoe UI 9/10,
  Times New Roman 12, Cambria 14, Tahoma 10, Aptos Narrow 10/11/12,
  Aptos/Display 11 and Lato 11. Six tested body-font families use measured
  GDI levels; other normals retain previous placement. Screen baselines,
  row heights, line pitches, and top/middle print alignment are unchanged.
- Corpus check: oyabayo now prints 39 pages, matching Excel; the affected
  right-aligned number pages moved from roughly 4.5 pt right to 0.21–0.27 pt
  right (median). The Segoe-heavy Summary sample now aligns vertically within
  0.02 pt on pages 1 and 3; Financials still has pages about 0.73 pt low.

## Print round 4: oyabayo's trailing blank page, 2026-11-23

The round-2 item "oyabayo's last page ... ours leaves it out: 38 pages, Excel 39" does not reproduce with the current
native exporter. A fresh direct export (each sheet's `Worksheet.ExportAsFixedFormat` before activating a sheet or
entering Page Break Preview, `excel-print-pages.sh`) has 39 pages; `tools/spreadsheet-pdf-runner` has 39. Excel's
`preview.hbreak` starts 35 and 74 and its `preview.vbreak` starts 6, 12, 18, 25, 32, 39, 46, 53, 60, 67, 73, 82;
our 39 page slices start at rows 1, 35, 74 and columns 1, 6, 12, 18, 25, 32, 39, 46, 53, 60, 67, 73, 82. Excel's
`used` is `$A$1:$CD$77`; our model's print bounds are A1:CD116 (the drawing anchor extends the row extent), at scale 1,
and the last slice is rows 74-116 x column CD. Rasterized at 40 dpi with `pdftoppm -gray`, page 39 has no non-white
dot in either PDF (all of pages 36-39 are blank in both), so this is a page-count match, not a page filled with an
overlay.

Eligibility trace: CD64:CD68 are cells with a fill, so they are in `printsCell` and take the implicit used range to
column CD; columns CA:XFD carry the same printable fill as a `<col ... style>` (through column 16,384). Page 39 is
inside that last column page, after the page with CD64:CD68. `Pagination.paginate`'s trailing trim keeps it because
`holdsInherited` finds the used extent reaching column CD in absolute 16-row block rows 65-80, which also contains
visible used rows 74-77 at the top of page 39. This is the existing "Print round 3" rule; no pagination or PDF change
was needed.

Public controls (checked in, fresh Excel 2019 direct counts in the `.excel.tsv` siblings; both are generated by
`make-print-pages-fixture.py`):
- `synthetic-print-pages-blockhold.xlsx` (`blockhold`): BlockHit 4 pages, BlockMiss 3, NoStyleHit 3, FillOnly 3;
  ours is 4/3/3/3 with the same starts (manual row break after 41, automatic column break at K). BlockHit prints a
  blank page 4; BlockMiss (target in the next 16-row block) trims it; NoStyleHit (no column fill) trims it; FillOnly
  (column fill with no used target row) trims it.
- `synthetic-print-pages-blockhold-holdout.xlsx` (`blockhold-holdout`), generated after the rule to test it, not used
  to derive it: different row break (after 47) and manual column page at I. Hit 4 pages, Miss 3; ours 4/3 with starts
  row 48 and column I.

Fresh `direct.pages` were also appended to `synthetic-print-pages-blank.excel.tsv`: all 14 sheets' direct counts
(Place 2, Blank 4, BlankGrid 4, BlankFill 4, BlankBorder 4, BlankLast 3, BlankLastTwo 4, BlankLastGrid 3, Pads 2,
BlankOverflow 3, BlankLastColFill 3, BlankLastColFont 4, BlankLastRowFill 3, BlankLastShape 3) equal
`Pagination.paginatePrinted`.

Verification: `dotnet fsi tests/print/helpers.fsx` -> 227 passed, 0 failed; oyabayo's direct PDF is 39 pages and the
runner's PDF is 39.

## Print round 5: oyabayo's right-aligned italic numbers, 2026-11-23

The round-3 item "a workbook's right-aligned numbers 4.6 pt right" reproduces only for **italic** right-aligned text,
and not at 4.6 pt as a delta. Fresh Excel 2019 direct export (`excel-print-pages.sh`, before any sheet activation or
Page Break Preview; `/tmp/oyabayo-fresh/excel/oyabayo-1.pdf`, 39 pages, sha256 `025a65bd...`; the workbook's sha256
`ea15d9ae...` equals `artifacts/real-estate/oyabayo-pari-passu-pref-return-waterfall-v1.1.xlsx`) against native
`a0477030` with `RB_WORD_FONT_DIR` = `artifacts/fonts/windows:artifacts/fonts/windows-extra`
(`/tmp/oyabayo-fresh/ours-aptos.pdf`):

- Cell Y25/Y28 (Aptos Narrow 14 italic, `0.00%`, right): Excel's word edge xMax 118.2996 pt, ours 121.3701 pt, so
  ours is **+3.07 pt** right. A neighbouring upright 14 pt cell Y21 (`#,##0_);[Red](#,##0)`, `(90,000)`) has Excel
  121.08768, ours 121.3724 (+0.28): the columns, borders and page origin agree.
- All 122 paired negative percent words: delta median **+3.17 pt**, range +0.90 to +3.20; the 26 positive percent
  words have median +0.07 and maximum +3.18 (the shifted ones are its italic subgroup). Page 13's upright numeric
  median is +0.21.
- Removing the Aptos face (the older `ours.pdf` with only Calibri embedded) moves the *left* edge by ~4.5 pt (the
  glyphs are narrower) but leaves the right edge at 121.373: this is not font substitution. The old 4.6 pt figure is
  Excel's own *word-box* inset from the cell's right border (~5.1 pt: pad 1.92 + italic placement offset 2.8),
  measured from a naive column edge (123.39 pt) rather than from Excel's word edge; the actual delta is ~3.0 pt.
- Remaining +0.2 - +0.3 pt on upright table numbers is page/column geometry: our column is 72.4800 pt, Excel's
  cumulative edges average 72.4927 pt and are quantized per boundary; the shared grid inset is 0.48 pt. It is inside
  the existing print tolerance and is not the reported offset.

Synthetic controls (`tools/excel-parity/make-print-right-edge-fixture.py`,
`tests/test-documents-excel/synthetic-print-right-edge.xlsx`; one variable per sheet, fresh direct Excel export and
native; oracle table `tools/excel-parity/print-right-edge-reference.tsv`). The measured quantity is **Poppler's word
box** (`pdftotext -bbox` xMax) to the printed `#FFF2CC` fill's right edge - a word-box gap, not raster ink and not a
glyph outline. Excel vs ours: Aptos Narrow 14 `0.00%` upright 2.41/1.98, italic 5.22/1.98; Aptos Narrow 14 `0.00`
italic 5.12/1.96; Calibri 11 `0.00%` upright 2.11/1.97, italic 4.14/1.97; a 30-character column 5.19/1.99
(width-independent); accounting `_)` formats add no reserve on top of the pad (Excel 6.68 vs ours 6.28). The inferred
italic placement offset (Excel italic gap - Excel upright gap, same face/size) is 0.200 x point size for Aptos
Narrow at 9/11/14/20 pt (1.808/2.198/2.810/3.992) and 0.186 x point size for Calibri at 8/11/14/20 pt
(1.440/2.040/2.640/3.720); ours applies 0.000 at every size. Held out: Aptos 14 italic hashes reserve 1.440, not
2.794 (same-string upright twin gives gap 3.450 vs italic 4.890; widths identical, dxMin == dxMax == 1.440);
Calibri 12 italic `#,##0.0` width 24 reserve 2.280 against 2.221 predicted (+0.059, pass). Center-aligned italic
shifts about half (-0.995 at Aptos 14); left-aligned is unaffected (+0.321).

Falsified: the offset is **not** the last glyph's outline overhang. From the installed matching faces (real Calibri
`artifacts/fonts/windows/calibrii.ttf`, Aptos Narrow Italic `artifacts/fonts/windows-extra/aptosnarrow-28903934312.ttf`),
italic minus upright `(outline xMax - horizontal advance)/em`: Aptos `%` +0.0039, `#` -0.0547, `0` -0.0229, `9`
-0.0273; Calibri `%` +0.0181, `#` +0.0430, `0` +0.0322, `9` +0.0352. At print sizes those are +/-0.1 - 0.8 pt and of
the wrong sign for `%`: Aptos 14 predicts +0.05 pt for `%` and -0.77 for `#` against observed offsets 2.81 and 1.44.
The word-box xMax is advance-based besides, so outline overhang is not what the gap measures.

Refuted out of sample: the candidate `fontPt * tan(|italicAngle|) * ascentEm` (typo ascent when OS/2
USE_TYPO_METRICS, else usWinAscent) fits the faces it was fitted on, but a pre-registered validation does not
(`print-right-edge-predictions.tsv`, written before the export; results `print-right-edge-validation.tsv`;
acceptance |observed - predicted| <= 0.10 pt). Fails: Arial 14 -0.174, Arial 20 -0.248, Arial 14 digit -0.174, Arial
14 at 75 % -0.220, Times New Roman 14 -1.256, Cambria 14 -0.438, Aptos 14 hashes -1.354, Aptos 14 with a trailing
`0_)` -2.779. Passes: Aptos Narrow 14 at 75 % +0.003, Calibri 12 `#,##0.0` +0.059, Arial 8 -0.099 (at the edge); the
rows the candidate was
fitted on (Aptos 9/11/14/20 and Calibri 8/11/14/20 at 100 %) also pass. Observed coefficients at 100 % are Arial
0.180, Times 0.171, Cambria 0.178, Aptos `%` 0.201 but `#` 0.103, Calibri 12 0.190. The `#`/`%` split inside one face,
and identical italic/upright word widths with dxMin == dxMax for those strings, rule out both a single per-face
constant and a last-glyph advance artifact. No model covers the corpus, so no product code is shipped and screen
placement is unchanged. What survives is the phenomenon and scope (print, right-aligned, italic; width-independent;
about half for centered italic; none for left-aligned; none on trailing uninked `_x` pads), the Aptos/Calibri fitted
rows, and the falsification of the outline-overhang account.

Verification: `pdfinfo /tmp/oyabayo-fresh/excel/oyabayo-1.pdf` (39 pages);
`RB_WORD_FONT_DIR=$PWD/artifacts/fonts/windows:$PWD/artifacts/fonts/windows-extra dotnet run --project
tools/spreadsheet-pdf-runner -- artifacts/real-estate/oyabayo-pari-passu-pref-return-waterfall-v1.1.xlsx /tmp/ours.pdf 0`;
`python3 tools/excel-parity/compare-pdf-words.py /tmp/oyabayo-fresh/excel/oyabayo-1.pdf /tmp/ours.pdf` (page 13
median +0.21, largest cell text +3.20 on the italic `-97.54%`); targeted right edges with
`pdftotext -bbox -f 13 -l 13 <pdf> - | grep -E '98\.41%|103,000|\(90,000\)'`.

## Print italic reserve: GDI printer measurement and a shipped right-alignment rule, 2026-11-23

Round 5 left the printed italic offset with no model. This round measured the printer device directly, found that no
GDI text-extent metric produces it, and found an empirical layout rule that does: Excel 2019's printed right alignment
reserves a fifth of GDI's ascent for an italic run, cut down to whole printer dots.

**Device and DC.** Excel 2019's default printer on the VM is `Microsoft Print to PDF` (driver `Microsoft Print To PDF`,
port `PORTPROMPT:`; `Win32_Printer` Horizontal/VerticalResolution 600). The direct PDFs set font sizes in whole 600-dpi
dots (14 pt is `/F1 14.04 Tf`, 20 pt 20.04, 10.5 pt at 75 % 10.56), so the layout is 600 dpi. `CreateDC("WINSPOOL",
"Microsoft Print to PDF")` in the desktop session (1600 % CPU session? no: `run-vbs-in-session.ps1`) realizes the real
faces (`GetTextFaceW`: Arial, Calibri, Cambria, Times New Roman, Georgia, Tahoma, Aptos Narrow loaded privately from
`artifacts/fonts/windows-extra`). `GetDeviceCaps`: see `devcap` rows of the probe output (TECHNOLOGY 2, LOGPIXELSX/Y
600; PDF media units are points, the layout is dots). `GetTextMetricsW` for every real italic and every synthesized
italic (Arial Black, Lucida Console, MS Sans Serif, Small Fonts) has **tmOverhang = 0**; `GetTextExtentPoint32W`
equals the sum of `GetCharABCWidthsW` A+B+C to the dot for every string measured; `GetTextExtentExPointW` agrees;
`GetCharABCWidthsFloatW` agrees; `GetGlyphOutlineW` GGO_METRICS is not available on this printer DC. So the reserve is
**not** tmOverhang, not an ABC side bearing, not a hinted whole-run width, and not the last glyph's outline overhang
(the latter is +/-0.1 - 0.8 pt and the wrong sign for `%`, as round 5 found).
Tools: `tools/excel-parity/gdi-print-italic-metrics.ps1` (printer DC, faces, strings, synthesized italics, Aptos
private), run through `run-vbs-in-session.ps1`. The format controls are
`tools/excel-parity/make-print-italic-format-controls.py` and `print-italic-format-controls-predictions.tsv`
(excel export in `/tmp/italic-fmt`, not retained; the predictions and measured conclusions are).

**Rule.** For a right-aligned printed italic run, Excel's layout width is the font's advance width plus
`floor(tmAscent / 5)` printer dots, where `tmAscent` is GDI's ascent of the run's face at the printed pixel height
`round(points x scale x 600 / 72)`. In points the reserve is `floor(tmAscent / 5) x 72/600` (3.6 printer dots/pt).
It is a layout offset, not a glyph metric: it is width-independent, applies to literal and formatted displayed
strings alike, and is invisible at left alignment. The PDF renderer computes it from the face it actually draws
(`winMetrics primaryTtf` and `.UnitsPerEm` in `Spreadsheet.Pdf`), including a fallback face, as
`round(usWinAscent x heightDots / upm)`, then floors by five; this is a **candid approximation of GDI's tmAscent**:
GDI's integer ascent differs by one dot for a few face/size pairs, giving at most 0.12 pt of reserve error (Arial 14
at 75 %: GDI 79, formula 80, so 1.92 against Excel's 1.80). Cambria is the other known residual (formula 2.64 against
2.49-2.52). No per-face offsets were added, and no generated per-face metric table is shipped.

Applied in `Paint.buildCellPaintItems` as the semantic flag `PaintText.PrintItalicEligible` (print only, italic,
right-aligned, unrotated cell text); `Spreadsheet.Pdf` measures and subtracts the reserve from the aligned x for
`AlignRight`. Suppression is the **active section's unquoted `*` fill** (Excel's accounting/fill number layout draws
several runs), not the `_)` pad: the format controls (`print-italic-format-controls-predictions.tsv`) measured a
numeric `0.00_)` at the full 2.782 pt, `0.00"_)"` full, `0.00;[Red]-0.00_)` full even on the negative (pad-only)
section, while `* 0.00`, `_(* 0.00_)` and standard accounting are 0.008-0.022 pt; literal cells under the same codes
all take 2.768-2.782 pt. `CellText.numberFormatUsesFillForValue` selects the section by Excel's sign/condition rules
and ignores quoted/escaped `*` and inactive sections. Also suppressed for `###` overflow placeholders whose raw
formatted text differs from the displayed hashes; a numeric cell with a custom `"#######"` format is not overflow and
**does** get the reserve. Centred italic is still uncorrected (about half the reserve, -1.0 pt at Aptos Narrow 14);
left alignment needs none.

**Blind holdouts.** The floor variant was preregistered in
`print-italic-holdout-predictions.tsv` before `synthetic-print-italic-holdout.xlsx` was exported
(sha256 `3b0d9f46...`). Fresh direct Excel export via `excel-print-pages.sh`, then Poppler word-box gaps (italic sheet
less upright sheet) in `print-italic-holdout-results.tsv`. Verdana 13 at 85 % (2.160 predicted/observed), Segoe UI
10.5 at 90 % (2.040/2.047), formatted hashes 2.160/2.160 and literal accounting 2.760/2.781 passed; Georgia 15 at
100 % observed 2.7525 against the floor prediction 2.64, a **fail (+0.11)**. The rounded-ascent variant predicts 2.76
and was then tested blind on a second workbook, `synthetic-print-italic-holdout2.xlsx` (sha256 `f9268670...`),
predictions `print-italic-holdout2-predictions.tsv` written before export: Tahoma 16 at 80 % 2.520/2.520, Garamond 11
at 115 % 2.160/2.1465, Verdana 9 at 110 % 1.920/1.920, Calibri 10.5 at 95 % 1.800/1.800 - all pass
(`print-italic-holdout2-results.tsv`). The first workbook's other cases also pass under the rounded variant. This is
honest post-hoc revision: the floor model was falsified by a blind case, the rounded model is blind-validated on four
fresh face/size/scale combinations.

**Format path.** `0.00_)` full, `0.00"_)"` full, `0.00;[Red]-0.00_)` full for both signs, `* 0.00` and
`_(* 0.00_)` and standard accounting ~0 for numeric; every literal twin full. So the discriminator is an active
unquoted `*` fill token, not a pad or an accounting label.

**End-to-end improvement (ours, PDF runner).** Same first workbook, ours rendered with
`RB_WORD_FONT_DIR=artifacts/fonts/windows:artifacts/fonts/windows-extra` and the new rule; reserve deltas
ours - Excel: Verdana 13/85 % +0.000, Georgia 15 +0.008, Segoe UI 10.5/90 % -0.007, formatted hashes +0.000, numeric
accounting -0.016, literal accounting -0.016, overflow hashes -1.440 (then unchanged: Excel's overflow placeholder
has its own layout; its placement was modelled separately on 2026-11-23, below). Second workbook: Tahoma +0.000, Garamond +0.012, Verdana +0.000, Calibri
+0.000. Before the change all ordinary/hash/literal reserves were 0, so the improvement is the full reserve
(1.44-2.76 pt). On oyabayo page 13 the italic `-98.41%` word xMax is now 118.6101 against Excel's 118.2996
(+0.3105, was +3.0705); the upright `(90,000)` remains +0.285 (page/column geometry), so the italic-specific error
is gone and the residual matches the sheet's upright text.

Verification commands:
`python3 tools/excel-parity/make-print-italic-holdout.py tests/test-documents-excel/synthetic-print-italic-holdout.xlsx`;
`tools/excel-parity/excel-print-pages.sh tests/test-documents-excel/synthetic-print-italic-holdout.xlsx /tmp/italic-holdout`;
`python3 tools/excel-parity/measure-print-italic-holdout.py /tmp/italic-holdout /tmp/italic-holdout` (with
`ours-<n>.pdf` from the runner); `dotnet fsi tests/print/helpers.fsx` (240 passed, 0 failed);
`dotnet build src/Spreadsheet/Rbt.Spreadsheet.fsproj --no-restore` and
`dotnet build src/Spreadsheet.Pdf/Rbt.Spreadsheet.Pdf.fsproj --no-restore`.

Left: the GDI ascent formula is rounded, not a GDI-table lookup (Arial 14 at 75 % and Cambria residuals above);
centred italic text outside a true overflow placeholder still takes no half reserve (the placeholder's centred
placement and half reserve are shipped, below); linked-cell/camera pictures are built from the screen
paint path and do not get the reserve; rich-text runs use the cell format's italic flag for the whole block.

## Printed hash geometry: true overflow is centred, 2026-11-23

Fresh direct Excel 2019 exports of the public `synthetic-print-hash-geometry.xlsx`
(`make-print-hash-geometry.py`; 15 sheets changing one of alignment, face/size, print scale, or
overflow/literal; every test column carries its own solid fill whose printed rectangle is the physical
cell edge). Measured through `HashGeometryOracle`'s tools (`predict-print-hash-geometry.py`,
`measure-print-hash-geometry.py`, `analyze-print-hash-geometry.py`): Poppler's word box
(`pdftotext -bbox`) against both fill edges, an advance box rather than ink. Predictions were locked
before the export: `print-hash-geometry-predictions-round1.tsv` (slope 1 right / 0.5 centred / 0 left
by the cell's alignment) was refuted; `-round2.tsv` locked the alternative before its confirming
export; measurements in `print-hash-geometry-reference.tsv`.

| Rule (Excel 2019) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| A true numeric overflow placeholder (`###`, where the displayed hashes differ from the formatted text) is **centred in the cell's content box whatever the cell's horizontal alignment**: d(right edge from cell left)/d(physical width) = 0.50. | R/L/C Aptos Narrow 14 overflow rows identical (max |ΔxMin| = |ΔxMax| = 0.000 over 21 widths); 53 fixed-count groups over right/left/centre, upright/italic, Aptos 14, Calibri 11, Arial 12, 100 % and 75 %: slope mean +0.501 (0.454..0.518). At the 7->8 boundary the jump is dxMin = -a/2 + dW/2, dxMax = +a/2 + dW/2. | followed the cell's alignment (right +2.80 pt, left -2.45 pt from Excel; centre already +0.43) | `PaintText.PrintOverflowHashes`; the PDF centres that run |
| The same placeholder in italic takes **half the printed italic reserve** (Aptos 14 -1.394 pt, Aptos 14 at 75 % -1.025, Calibri 11 -1.018); run width unchanged. | the oracle's italic vs upright sheets. | no shift | PDF subtracts reserve/2 for `PrintOverflowHashes && Italic` |
| Literal hash text and a custom number format whose quoted text is hashes are ordinary text/numbers: they keep the cell's alignment (right slope +1.0004, centre +0.4997, left 0.0000) and the ordinary italic reserve. | literal sheets of the oracle; a scratch export with custom `"#####"` in all three alignments gives custom geometry identical to the literal cells on every sheet. | already so | unchanged; the flag is set only when the raw formatted text differs from the displayed hash run |
| The hash **count** is the existing fit rule, unchanged: the largest run of the face's `#` advance that fits the printed cell's content width. | 213/245 overflow/short rows exact, 32 off by one; literal counts 44/44. | existing rule | unchanged here (placement only); the count rule is superseded in the next section |

Quantified on the controlled sweep (97 fixed-count overflow pairs, same counts): absolute centre error
vs Excel mean 2.229 pt / max 5.646 before and 0.397 / 0.455 after; by alignment right +2.795 -> -0.389,
left -2.445 -> -0.406, centre +0.434 -> -0.406. The residual -0.4 pt is a word-box convention, not
placement: Excel's centred ordinary and literal text also reads +0.41 from the fill centre through
Poppler while ours reads 0.00. Literal rows are identical before and after (88/88); custom-format and
browser/screen alignment paths are untouched, and no count or per-face/size constant changed.

Limits: the metric is a word box; the half reserve is applied to true overflow only, not to other
centred italic text; the pre-existing count rule keeps its off-by-one cases; the 1.44 pt overflow
italic residual quoted in the italic-reserve section is now explained by the half reserve, not a
separate layout.

Verification:
`python3 tools/excel-parity/make-print-hash-geometry.py tests/test-documents-excel/synthetic-print-hash-geometry.xlsx tools/excel-parity/print-hash-geometry.meta.tsv`;
`tools/excel-parity/excel-print-pages.sh tests/test-documents-excel/synthetic-print-hash-geometry.xlsx /tmp/hashgeo/excel`;
`python3 tools/excel-parity/measure-print-hash-geometry.py tools/excel-parity/print-hash-geometry.meta.tsv tools/excel-parity/print-hash-geometry-predictions-round2.tsv /tmp/hashgeo/excel /tmp/hashgeo/results.tsv`;
`python3 tools/excel-parity/analyze-print-hash-geometry.py /tmp/hashgeo/results.tsv /tmp/hashgeo/summary.tsv`;
`RB_WORD_FONT_DIR=artifacts/fonts/windows:artifacts/fonts/windows-extra tools/spreadsheet-pdf-runner/bin/Debug/net10.0/SpreadsheetPdfRunner tests/test-documents-excel/synthetic-print-hash-geometry.xlsx /tmp/hashgeo/ours-<n>.pdf <n-1>`;
`dotnet fsi tests/print/helpers.fsx` (250 passed, 0 failed); rebuild `src/Spreadsheet`, `src/Spreadsheet.Pdf` and the runner from source.
Browser/screen path: `bash scripts/build-dist.sh dist`, then load `dist/spreadsheet` through `tools/browser-runner`
(`npm --prefix tools/browser-runner run dev`, import the fixture, screenshot/inspect the cells); the new flag is
print-only and no DOM/grid code consumes it, so alignment there is unchanged. Independent Chromium 147 gate
(HashGeometryOracle, frozen dist + runner, Fable `Spreadsheet.ExportPdf` against native): 11 PDFs, equal page
counts, browser vs native max 0.0002 pt; same-count overflow |centre residual| max 0.456 / mean 0.396 pt,
literal/custom right-gap max 0.469 and xMaxLeft max 0.503 pt; gate passed.

## Printed hash count: printer-unit recount, 2026-11-23

The count left open by the centred-placement round (`213/245` exact, 32 off by one) is now modelled
and shipped. Trace: `Paint.cellLookFitted` still decides *whether* a number overflows; only a true
placeholder (`PrintOverflowHashes`, raw formatted text differs from the all-`#` display) reaches the
new recount. The PDF renderer (`Pdf.appendDrawTextOps`) replaces that run's length from the face it
actually draws; placement, alignment, the overflow decision, literal/custom hash text and the
browser/screen path are unchanged.

| Rule (Excel 2019, printer units of 600 dpi) | Evidence | Ours before | Now |
| --- | --- | --- | --- |
| A true overflow run has `n = floor((W - c) / a) - slots`; `a = round(hashEm x dots) x 72/600`, `dots = round(points x scale x 600/72)`; `W` is the full cell content width (not the text rect's padded width); `c_upright = digitEm x points x scale / 2 + 0.875 pt`; italic adds the **full** right-align reserve (`FontMetrics.printItalicShiftPt`), not the screen overhang; `slots` counts the active section's leading `_x` pads. It is not a constant inset: Calibri 11 is 3.66 pt, Arial 12 4.21, Aptos 14 4.42. | Direct-export geometry sweep: 243/245 exact (the two misses, Calibri 10.5 and Arial 10.25, sit 0.02-0.06 pt below a transition because the PDF fill edge is 0.10-0.15 pt wider than the layout width); literal 44/44 and the row-2 controls unchanged. CountTransitions' dense 1/16-char training: 781/786 with measured W (overflow 695/700, full 86/86), 769/786 with the pre-registered W model (overflow 688/700, full 81/86); every miss is one hash at a transition or the hash/full edge and there is no pitch drift. The current rule replicated on that sample is 621/786 (539/700 overflow). | c9fe8034 runtime measured here: 213/245, 32 off by one (the oracle's Python replica of `fittedRoom`/`fittedHashes` reads 194/245; the runtime's `fittedHashes` reserve and exact measure put it at 213). | `PaintText.PrintHashSlots` / `PrintHashCellWidth`; PDF recount above |
| The formatted number still **fits** when its printed text width is at most `W - c` (the same `c`, italic reserve included). | Training full transitions: every fitted `c_text` interval contains the count `c`. Blind holdout: 127 full-number rows predicted exactly, including Georgia italic at 16.6875/16.75 chars (Excel prints 8 hashes although the text is 0.02-0.05 pt inside `W-c`, because the measured PDF W is wider than the layout W); the first full is 16.8125. | existing fit rule | unchanged here (overflow decision kept) |

**Blind holdout.** `make-print-count-fixture.py`'s second workbook (Verdana 13 upright 100 % `0.00%`;
Georgia 15 italic 90 % `#,##0.00`; widths 8.00-15.00 and 16.00-21.00 by 1/16 char) was not exported
until `print-count-holdout-predictions-countmodel.tsv` was locked: **388/388 count predictions exact**
(overflow 261, full 127), with no count contradiction. Rendering the workbook through the shipped
native path agrees with all 383 predictions on cells where the existing overflow decision matches
Excel; the 5 residual rows are decision mismatches outside this slice - Verdana 17.375/17.4375/17.5
(the old fit prints 8 hashes where Excel prints the full number) and Georgia italic 16.6875/16.75
(the old fit prints the full number where Excel prints 8 hashes). The count rule is not applied to
those because the decision is intentionally untouched.

Limits: the 600-dpi pitch is the rounded design advance; GDI's own A+B+C equals it in the measured
core faces but a marginal face could differ by a dot; `W` is the layout cell width, so Excel's PDF
fill edge being 0.10-0.15 pt wider leaves the two coarse-fixture boundary cases and the train
one-less transition misses. A narrow-cell control (Aptos 14, widths 0.5/1/2/3/4 chars, fresh direct
Excel PDF) prints nothing at 0.5 and 1, one hash at 2 and 3, two at 4; ours prints nothing at 0.5/1/2
and one/two at 3/4, the width-2 case being the same one-dot layout-width boundary, not a new hide
rule; the recount clamps a flagged placeholder to at least one hash, so it can never blank a cell
Paint already printed as hashes. The `0.875 pt` part of `c` is empirical: half the digit advance is the font's, the rest was
fitted to the printer intercept intervals and has no separate metric derivation. The overflow
decision's own full/hash borders (above) remain old; only true numeric placeholders are recounted
(literal and custom `"#####"` formats keep their strings).

Verification:
`dotnet build src/Spreadsheet/Rbt.Spreadsheet.fsproj --no-restore`;
`dotnet build src/Spreadsheet.Pdf/Rbt.Spreadsheet.Pdf.fsproj --no-restore`;
`dotnet build tools/spreadsheet-pdf-runner/SpreadsheetPdfRunner.fsproj --no-restore`;
`RB_WORD_FONT_DIR=$PWD/artifacts/fonts/windows:$PWD/artifacts/fonts/windows-extra tools/spreadsheet-pdf-runner/bin/Debug/net10.0/SpreadsheetPdfRunner
tests/test-documents-excel/synthetic-print-hash-geometry.xlsx /tmp/countmodel-ours/ours-<n>.pdf <n-1>`;
`python3 tools/excel-parity/measure-print-hash-geometry.py /tmp/hashfinal/synthetic-print-hash-geometry.meta.tsv
tools/excel-parity/print-hash-geometry-predictions-round2.tsv /tmp/countmodel-meas /tmp/countmodel-meas/results-ours.tsv`
(243/245 overflow row-1 counts, 0 non-overflow/literal count changes);
`RB_WORD_FONT_DIR=... SpreadsheetPdfRunner tests/test-documents-excel/synthetic-print-count-holdout.xlsx /tmp/countmodel-holdout/holdout-<n>.pdf <n-1>`;
`python3 tools/excel-parity/measure-print-count-transitions.py tools/excel-parity/print-count-holdout.meta.tsv /tmp/countmodel-holdout /tmp/countmodel-holdout/results-ours.tsv`
(383/383 decision-agreeing counts equal the locked predictions);
`dotnet fsi tests/print/helpers.fsx` (254 passed, 0 failed; includes the consumer-level regression render of
Aptos 14 at 8.25 characters, which asserts printed `######`, plus ordinary/-98.41, full/42 and literal/`#####`
controls through `Render.toBytes` + `pdftotext`).
