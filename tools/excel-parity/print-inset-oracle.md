# Printed cell-text insets, hinted widths and the fill count

Question (from `print-fill-layout-gate.md`, "Where the 0.34-0.74 pt no-fill box gap comes from"): what
places a printed cell's text, and does that explain the trailing-fill count without a constant per face?

## Files

| role | path |
| --- | --- |
| generator | `tools/excel-parity/make-print-inset-fixture.py` (`--set train/holdout`, `--normal-face/--normal-size`) |
| workbooks | `tests/test-documents-excel/synthetic-print-inset-{train,holdout,normal}.xlsx` |
| manifests | `tools/excel-parity/print-inset-{train,holdout,normal}-cells.tsv` |
| hinted widths | `tools/excel-parity/print-inset-gdi.fsx` -> `print-inset-gdi.tsv` (the engine's GDI advance tables) |
| locked predictions | `tools/excel-parity/print-inset-{train,holdout,normal}-predictions.tsv` (commit `79c26de0`) |
| measurement | `tools/excel-parity/measure-print-inset.py` |
| lock / check | `tools/excel-parity/lock-print-inset.py` |

Plain numbers under `0.00` (1234.5, -1234.5, 0) and the text "1234.50", no fill, no border, columns of
14 and 18 characters with a solid fill each, one sheet per (print scale, alignment). Full grid:
Aptos Narrow, Arial, Calibri, Georgia, Tahoma, Verdana x 9/10/11/12/14 pt x 60/75/100 %, Normal font
Calibri 11. Training = the first four faces x 9/11/14 x 60/100 (576 cells); holdout = the rest
(1,584 cells: two unseen faces, two unseen sizes and an unseen scale); `normal` = the training grid
with a Georgia 12 Normal font (576 cells).

## Cell geometry in Excel's PDF

Each solid cell fill is drawn one 600-dpi dot wider than its column: the rectangle's right edge is the
next column's left edge + 1 dot. Interior boundaries clustered from all edges therefore give the true
width; the last fill column of a page reads 1 dot wide (`measure-print-border-fill.py` and the earlier
`border-*`/`trail-*` tables carry that dot on the final column). True width = rectangle width - 1 dot.

## Rule (all in 600-dpi device dots)

- `ppem = round(points x scale x 600 / 72)`, the pixel size the cell font prints at.
- `P = ceil(d / 4) + 4`, d = the **cell** font's widest digit (hinted GDI advance at ppem). The Normal
  font plays no part (the earlier `printPadXOf` rule came from fixtures whose cells used the Normal font).
- `W` = the text's hinted GDI width at ppem (`ExcelText.printWidthDots`), not the face's design width.
- left aligned:   `x0 = cell_left + P`
- right aligned:  `x0 = cell_right - P - W`
- centred:        `x0 = cell_left + floor((cw - W) / 2) + 4`

Numbers and text behave identically (the `text` rows equal the `pos` rows on every cell).

## Result

Predictions locked on the training export before the holdout and Normal-font workbooks were exported:

| set | cells | left residual | right residual (vs fill edge) | centre residual (vs fill width) | failures |
| --- | --- | --- | --- | --- | --- |
| train | 576 | 0.000 | +0.917..+1.200 | +2.900..+3.600 | 0 |
| holdout | 1,584 | 0.000 | +0.917..+1.200 | +2.900..+3.600 | 0 |
| normal (Georgia 12) | 576 | 0.000 | +0.917..+1.200 | +2.900..+3.667 | 0 |

The locked residuals are against the fill rectangle (so right reads P + 1 and centre reads 3 dots,
because the rectangle is 1 dot wide); against the true width both are the rule above. The sub-dot
spread is where the cell edge falls between dots.

## Fill count

With the same quantities, the trailing `*x` repeat count is

`count = max(0, floor((cw - 2P + 2 - w) / a))`

with cw the true column width, w the hinted width of the number run and a the hinted advance of the
fill character. It is exact on every Excel fill cell that has been measured: border sweep 3,240/3,240
(Aptos Narrow 14/100 %, Arial 11/60 %, Calibri 11/100 %, four border weights), trailing training
420/420 and trailing holdout 480/480 (true widths from the fill rectangles). The law was found on the
border sweep and checked on the two trailing sets without refitting. The `+2` is one integer, the same
on every face and scale. The earlier face-dependent intervals came from (a) the Normal-font pad,
(b) design rather than hinted widths and (c) the extra dot on the last column.

## Engine

- `ExcelText.printPx`, `printWidthDots`, `printCellPadDots`.
- Paint (print only): each cell's text rect is inset by its own font's `P`; a font the tables do not know
  keeps `printPadXOf`. The fill layout counts with hinted widths over `cw - 2P + 2` dots.
- PDF: `PaintText.PrintGdiAlign` (cell text and fill runs only, not chart/shape/header/rich text) aligns
  right/centre by the hinted width, centred 4 dots right of the middle, and sets the run as one TJ whose
  adjustments place each glyph at its hinted advance, as Excel's PDF does. A true `###` overflow
  placeholder keeps its earlier centring.

Ours vs Excel after the change (native, same workbooks, 2,736 cells): text origin left 0.000, right
-0.25..+0.03, centre -0.25..+0.53 dots (before: -7..+9). Run advance within 1 dot on 552/576 training
and 1,536/1,584 holdout cells (before 354 and 732); the rest are Tahoma 9 and 14 at 60 % and Verdana 11
at 100 %, where Excel's own glyph steps differ from the hinted table by 1-4 dots although its alignment
follows the table.

## Trailing-fill paired gate (candidate vs bf132a48)

`run-print-numfill-trail-gate.sh` with a dist and Release native runner from this tree, baseline
`/var/tmp/fill-engine/dist` + a bf132a48 native runner:

| set | absolute 0.35 | relative checks | count errors vs Excel | count better | browser == native |
| --- | --- | --- | --- | --- | --- |
| train (Aptos Narrow 14/100) | PASS (was FAIL) | 12/12 PASS | 0 (was 60) | 342 | 0 diffs |
| holdout (Arial 11/60) | PASS (was FAIL) | 12/12 PASS | 0 (was 54) | 423 | 0 diffs |

The 15 one-glyph count regressions of the earlier candidate (9 train, 6 holdout) are gone. Worst
newly worse cell: x1 +0.070 pt (29 train cells, all within 0.35). The control check now requires the
no-fill controls to keep text/hash/count and to move no further from Excel (an error growing by more
than 0.12 pt, or newly over 0.35 pt, fails). Every control moved toward Excel: x0 better on 189/189
train and 225/225 holdout; train max |x0/x1 error| 0.648 -> 0.202 pt (cells over 0.35 pt: 252 -> 0),
holdout 0.573 -> 0.057 pt (261 -> 0). PAIRED_RELEASE_GATE PASS. The numeric-fill gate passes on the same
dist (browser == native on 1,035 / 630 / 630 rows), and so does the text-pad gate (0 failures; max
error vs Excel 0.229 / 0.128 / 0.060 / 0.210 pt train / Georgia / Tahoma / escapes, was 0.645 /
0.598 / 0.591 / 0.614). `tests/print/helpers.fsx` passes 336/336.

## Commands

```
python3 tools/excel-parity/make-print-inset-fixture.py --set train \
  --out tests/test-documents-excel/synthetic-print-inset-train.xlsx --cells tools/excel-parity/print-inset-train-cells.tsv
tools/excel-parity/excel-print-pages.sh tests/test-documents-excel/synthetic-print-inset-train.xlsx /var/tmp/inset-oracle/train-excel
python3 tools/excel-parity/measure-print-inset.py tools/excel-parity/print-inset-train-cells.tsv \
  /var/tmp/inset-oracle/train-excel /var/tmp/inset-oracle/train-excel/synthetic-print-inset-train.tsv /var/tmp/inset-oracle/train-excel.tsv
python3 tools/excel-parity/lock-print-inset.py check tools/excel-parity/print-inset-train-predictions.tsv /var/tmp/inset-oracle/train-excel.tsv
(cd tools/excel-parity && dotnet fsi print-inset-gdi.fsx > print-inset-gdi.tsv)
```
