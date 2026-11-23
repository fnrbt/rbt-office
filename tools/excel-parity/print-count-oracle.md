# Printed hash-count oracle: transitions, intercepts, full-number fade-in

Owner: CountTransitions (Excel oracle). Runtime model: CountModel. Source of truth: fresh Excel 2019
direct `ExportAsFixedFormat` PDFs, before any activation or Page Break Preview. Physical cell edges come
from per-column solid fill rectangles; the printed run is Poppler's `pdftotext -bbox` word box.

## Inputs and tools

| Path | What |
| --- | --- |
| `tools/excel-parity/make-print-count-fixture.py` | deterministic generator (`train` / `holdout` split) |
| `tests/test-documents-excel/synthetic-print-count-train.xlsx` | 16 sheets, 786 test columns, 1/16-char width grids across several adjacent count transitions |
| `tests/test-documents-excel/synthetic-print-count-holdout.xlsx` | 2 unseen-face/size/scale+format sheets, 388 test columns; `0.00%` and `#,##0.00` |
| `tools/excel-parity/print-count-train.meta.tsv`, `print-count-holdout.meta.tsv` | per-column face/size/italic/scale/format/value/width |
| `tools/excel-parity/measure-print-count-transitions.py` | PDF measurement (reuses the hash-geometry fill/word readers) |
| `tools/excel-parity/print-count-train-results.tsv`, `print-count-holdout-results.tsv` | one row per printed word; physical `cellWidth`, `runWidth`, `nHashes` |
| `tools/excel-parity/analyze-print-count-transitions.py` | per-case pitch/intercept/drift/fade-in analysis |
| `tools/excel-parity/print-count-train-analysis.tsv`, `print-count-holdout-analysis.tsv` | analyser output |
| `tools/excel-parity/predict-print-count.py` | candidate model + pre-registered W(w) |
| `tools/excel-parity/print-count-holdout-predictions-countmodel.tsv` | CountModel's predictions, locked before export |
| `tools/excel-parity/print-count-holdout-predictions-counttransitions.tsv` | this oracle's independent lock (Georgia file wrong, see limits) |
| `tools/excel-parity/print-count-model-comparison.tsv`, `print-count-model-errors.tsv` | denominators and every counterexample |

Regenerate and measure:

```bash
python3 tools/excel-parity/make-print-count-fixture.py \
  tests/test-documents-excel/synthetic-print-count-train.xlsx \
  tools/excel-parity/print-count-train.meta.tsv train
tools/excel-parity/excel-print-pages.sh \
  tests/test-documents-excel/synthetic-print-count-train.xlsx /tmp/count-train
python3 tools/excel-parity/measure-print-count-transitions.py \
  tools/excel-parity/print-count-train.meta.tsv /tmp/count-train \
  tools/excel-parity/print-count-train-results.tsv
python3 tools/excel-parity/analyze-print-count-transitions.py \
  tools/excel-parity/print-count-train.meta.tsv \
  tools/excel-parity/print-count-train-results.tsv \
  tools/excel-parity/print-count-train-analysis.tsv
```

## Measured rule (Excel 2019 print)

For a true numeric overflow run of `n` hashes:

```
dots        = round(size_pt * print_scale * 600 / 72)          # 600 dpi font height
pitch a     = round(hash_advance_em * dots) * 72/600           # GDI's A+B+C advance
offset c    = digit_advance_em * size_pt * print_scale / 2 + 0.875
              + (italic ? printItalicShiftPt(face, size_pt, scale) : 0)
              + slots * a          # slots = leading `_x` tokens of the active section
n           = max(0, floor((W_pt - c) / a) - slots)
full value  = same c; print the number when its per-character rounded-dot width <= W_pt - c
```

`W_pt` in this oracle is the measured fill-rectangle width. A pre-registered width model is
`W = floor(screenPx(w) * 47/7 + 0.5) * 72/600 * scale`; on the dense training columns it is
0.10-0.15 pt under the PDF fill width, which moves a few exact-boundary columns.

Evidence:

- **Pitch.** `runWidth / n` equals the quantized design `#` advance to within 0.006 pt in every
  training case (Aptos Narrow 14 6.9647 vs 6.9600; Calibri 11 5.5177 vs 5.5200; Arial 12 6.7149
  vs 6.7200; Aptos Narrow 14 at 75 % 5.2770 vs 5.2800). The PDF word box is slightly narrower than
  the GDI A+B+C pitch, so the count rule must use the design/GDI advance, not `runWidth/n`.
- **Constant offset, no drift.** For every case the intersection of the adjacent `n -> n+1` intercept
  intervals is non-empty when `a` is the quantized pitch: Aptos 14 (4.344, 4.581] pt, Calibri 11
  (3.621, 3.722], Arial 12 (4.102, 4.342], Aptos 14 at 75 % (3.480, 3.720], and the italic twins
  (6.983, 7.320], (5.666, 5.906], (5.639, 5.785]. A pitch error would make these intervals walk;
  they do not. Least-squares boundary slopes scatter by +-0.13 pt around `a`, inside the half-grid
  quantization of the 1/16-char sweep; transition residuals are at most +-0.37 pt.
- **Fit of `c`.** The formula above lands inside every count interval. Largest deviation from an
  interval mid-point is 0.20 pt on the single-transition full-only Aptos 14 at 75 % window; among
  multi-transition count cases it is 0.14 pt (Aptos 14 italic at 75 %, formula low); every other case
  is within 0.10 pt.
- **Italic** adds the full printed right-alignment reserve, measured as `c_italic - c_upright`:
  +2.69 pt Aptos 14 (formula 2.76), +2.12 Calibri 11 (2.04), +2.11 Aptos 14 at 75 % (2.04).
- **Accounting pads.** A leading `_(` costs exactly one pitch: Aptos 14 accounting interval
  (11.304, 11.541] vs 4.423 + 6.96 = 11.383; Calibri 11 (9.121, 9.261] vs 3.663 + 5.52 = 9.183.
  The trailing `_)` changes the full-value fit only.
- **Full-number fade-in.** Dedicated dense windows bracket hash -> number for Aptos 14 upright/
  italic, Calibri 11 upright/italic, Arial 12, accounting, and Aptos 14 at 75 %. The measured
  `c_text = W - printed_run` intervals contain the same `c` as the count rule in all cases once the
  `_x` pad advances are subtracted from the run (e.g. Aptos 14 upright (3.922, 4.785]); no separate
  full-number constant is needed. The one
  borderline is Georgia 15 italic at 90 %: columns 126/127 at W 84.961/84.984 print 8 hashes while
  `text = 77.52 <= W - 7.418` by 0.02-0.05 pt; first full is column 128 at W 85.680.
- **Non-overflow control** (`-98.41`, format `0.00`, row 2) is full text in all 786 training columns.

## Excel vs current vs candidate

`print-count-model-comparison.tsv` (accuracy = exact columns / all row-1 columns):

| model | train dense (1/16-char) | coarse hash-geometry (1/4-char) | holdout (unseen) |
| --- | --- | --- | --- |
| current rule replicated from `CellText.fs` | **621/786 (79.0 %)**, overflow 539/700, full 82/86 | **194/245 (79.2 %)**, overflow 184/235, full 10/10 | Verdana subset 186/194 (95.9 %, overflow 130/135, full 56/59) |
| candidate, measured W | **781/786 (99.36 %)**, overflow 695/700, full 86/86 | **244/245 (99.59 %)** | 386/388 post-hoc*, 194/194 Verdana |
| candidate, pre-registered W(w) | **769/786 (97.84 %)** | — | **388/388 CountModel locked** |
| fixed 4 pt inset | 625/786 (79.5 %) | 203/245 (82.9 %) | — |

`current` is a faithful Python replication of `fittedRoom`/`fittedHashes`
(`src/Spreadsheet/CellText.fs`) on the measured physical width, not a frozen c9fe8034 binary: the
shared tree was being rebuilt by CountModel during the run, so an `ours` PDF measurement was
discarded. On the dense boundary-concentrated sample the current rule misses 165/786; on the coarse sweep it
misses 51/245 (41 overflow rows). CountModel's shipped runtime reports 243/245 on the same coarse
sweep; the one-row difference is the runtime's layout W versus the measured PDF fill width, not a
rule change. The older `213/245` in the placement section is not directly comparable: it predates
this production-path replication and used a different predictor.

Candidate counterexamples with measured W (5/786 train, 1/245 coarse): all are Aptos Narrow 14 at
75 % and all are one hash off at a transition or the hash->full edge, consistent with the formula's
`c` being ~0.1 pt (one 600-dpi dot) low at that scale:
`T_aptos14_it_0_75` columns 42/43 (width 11.25/11.3125: Excel 7 hashes, candidate 8);
`T_aptos14_up_0_75_full` columns 16-18 (17.375-17.5: Excel 13 hashes, candidate full);
coarse `R_aptos14_italic_75` column 15 (11.25: Excel 7, candidate 8). No other training case errs.

## Holdout protocol and result

CountModel wrote `print-count-holdout-predictions-countmodel.tsv` and this oracle wrote
`print-count-holdout-predictions-counttransitions.tsv` before the holdout workbook was exported.
The holdout is Verdana 13 upright at 100 % with `0.00%` (value 12345.6789) and Georgia 15 italic at
90 % with `#,##0.00` (value 1234567), widths 8.00-15.00 and 16.00-21.00 in 1/16 steps.

- **CountModel locked: 388/388 exact** (261/261 overflow, 127/127 full).
- **CountTransitions locked: 384/388.** All 4 misses are Georgia italic columns 63/64 and 91/92 and
  are an input mistake, not a rule difference: the local `windows/GARAIT.TTF` is Garamond Italic, so
  the lock used `a = 9.00` instead of GDI's 8.76 and the wrong `c`. With the real Georgia Italic
  metrics and measured W the same rule gives 386/388; the two remaining rows are the full-number
  fade-in at columns 126/127 above (the pre-registered W model happens to land on the right side).

Verdana holdout: 388/388 for both locks. The pre-registered W(w) model was 0.13 pt mean / 0.21 pt max
from the measured cell width on the holdout.

## Limits

- Word-box measurement, not raster ink: `runWidth/n` is 0.002-0.006 pt below the quantized pitch.
- The 1/16-char sweep locates a transition to ~0.18 pt; intercept intervals are 0.1-0.6 pt wide.
  Differences below ~0.1 pt between two `c` formulas are not resolved, except where a boundary column
  separates them.
- Aptos Narrow 14 at 75 % sits at the low edge of its intercept interval; the base formula is one dot
  low there. A scale-correction fitted to five boundary columns was deliberately not added.
- Georgia Italic and Verdana metrics are not in `artifacts/fonts/windows` (the GARA files are
  Garamond); the holdout used the PDF-embedded face names and countmodel's GDI pitches. Do not copy
  the wrong Georgia metrics into product code.
