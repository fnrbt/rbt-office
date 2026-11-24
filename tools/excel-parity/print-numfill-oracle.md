# Numeric active-`*` fill layout: oracle, fixtures and blind holdout

Oracle for the print behaviour of a numeric format's active `*x` fill token: the segment before the
fill, the repeated fill character, and the segment after it, against matched no-fill formats. The
measurement is from the PDF content stream (`Tm`/`TJ` positions and font advances), not Poppler word
boxes, so whitespace fills do not bias the amount or currency-sign position.

## Owned files

| role | path |
| --- | --- |
| fixture generator | `tools/excel-parity/make-print-numfill-fixture.py` |
| training workbook | `tests/test-documents-excel/synthetic-print-numfill.xlsx` |
| holdout Georgia (12 pt, 75 %) | `tests/test-documents-excel/synthetic-print-numfill-holdout-georgia.xlsx` |
| holdout Tahoma (10 pt, 100 %) | `tests/test-documents-excel/synthetic-print-numfill-holdout-tahoma.xlsx` |
| manifests | `tools/excel-parity/print-numfill-cells.tsv`, `print-numfill-holdout-{georgia,tahoma}-cells.tsv` |
| locked predictions | `tools/excel-parity/print-numfill-holdout-predictions.tsv` (sha256 `96db0191...`) |
| operator measurement | `tools/excel-parity/measure-print-numfill.py` |
| prediction checks | `tools/excel-parity/analyze-print-numfill.py` |
| freeze gate | `tools/excel-parity/run-print-numfill-gate.sh` |

Short formats: `#,##0.00`, `0.00`, `* #,##0.00`, `*.#,##0.00`, `0.00* `, `0.00*.`,
`"*"#,##0.00`, `\*#,##0.00`, `0.00;[Red]* -0.00;0.00`, `0.00*.;0.00`. Wide formats:
`$* #,##0.00`, `$*.#,##0.00`, standard accounting, accounting with `*.`, `$0.00;[Red]$* -0.00;0.00`.
Each sheet has left/centre/right alignment rows, positive/negative/zero values and eight 0.25-char
width steps across repeat transitions. Every test cell has a distinct solid fill and a thin border;
the Georgia/Tahoma direct exports omit the thin border path, so the tool falls back to the printed fill
rectangle edges for the column geometry.

## Locked design and predictions

Predictions were written before either holdout was exported. Holdout Excel PDFs are in
`/var/tmp/numfill-oracle/holdout-{georgia,tahoma}`. The checks are relational, not absolute page
coordinates: fill-before amount right gap equals the right-aligned no-fill twin; fill-after amount
left gap equals the left content inset (Excel's sub-fill jitter on centre/right recorded as a range);
repeat count is monotone with width, step at most 2; count 0 falls back to the no-fill alignment;
quoted/escaped `*` and inactive sections stay literal; currency `$` stays at the left inset; accounting
number right gap is the no-fill gap plus the trailing `_)` pad.

## Training (Aptos Narrow 14, 100 %, direct Excel 2019)

Excel truth checks used by the locked predictions all pass on the measured training cells (0
`excel_failures` in `train-results.tsv`). The measured model:

- fill-before `* #,##0.00`, `*.#,##0.00`: number right gap equals the right-aligned no-fill twin;
  repeat count is non-decreasing and drops to 0 rather than hashing when the fill no longer fits.
- fill-after `0.00* `, `0.00*.`: number keeps the cell's own alignment and the fill runs rightward;
  the fill is dropped at the tight boundary.
- `$* #,##0.00`: `$` at the left inset, number at the right inset.
- quoted `"*"`, escaped `\*`, and the inactive positive/zero sections of `0.00;[Red]* -0.00;0.00`
  are literal/no-expansion and match the no-fill controls.
- The standard accounting `$` is 0.6-1.4 pt right on centre/right sheets; this sub-fill residual is
  left open and recorded, not tuned.

## Blind holdout

Georgia 12 pt 75 % and Tahoma 10 pt 100 %, exported after the prediction lock. Excel truth: every
checked relation passes except `ANCHOR_L_S2` on Tahoma (3 cells outside the preregistered
`[-0.3, 2.5] pt` left-inset range), recorded as a falsified bound. The number/fill counts, currency
left inset, accounting right gap, no-overflow, text and literal/inactive relations pass on both.

Our engine snapshot at the time of this writing (working-tree DLL built before FillEngine's latest
`Paint.fs` edit) still fails the screen-hash boundary and some suffix/inactive cases; those are
quantified in `/var/tmp/numfill-oracle/holdout-*-results.tsv` and are the evidence FillEngine is
using for the remaining prefix-empty/suffix/screen-hash controls. Do not read those ours-side counts
as the frozen engine result.

## Commands

```
python3 tools/excel-parity/make-print-numfill-fixture.py --set train \
  --out tests/test-documents-excel/synthetic-print-numfill.xlsx \
  --cells tools/excel-parity/print-numfill-cells.tsv
tools/excel-parity/excel-print-pages.sh tests/test-documents-excel/synthetic-print-numfill.xlsx /var/tmp/numfill-oracle/train
python3 tools/excel-parity/measure-print-numfill.py tools/excel-parity/print-numfill-cells.tsv \
  /var/tmp/numfill-oracle/train /var/tmp/numfill-oracle/train/synthetic-print-numfill.tsv \
  OURS_DIR /var/tmp/numfill-oracle/train-comparison.tsv
python3 tools/excel-parity/analyze-print-numfill.py \
  tools/excel-parity/print-numfill-holdout-predictions.tsv COMPARISON.tsv WORKBOOK
```

## Final freeze (f7104bb6 / a27e8eca / a6ab371e / dist b38ad676)

FillEngine's strict gate and my independent run on this freeze both pass: train 1035, Georgia 630,
Tahoma 630 ours rows; browser vs native max 0.00000 pt and 0 text/count/coord differences on all
three. The gate requires
nonempty equal ours-only key sets relative to the cell manifests, so it fails on a silently missing
measurement row.

Corrected paired comparison against baseline `6fe0e6e2`, same workbooks, browser PDFs only
(`print-numfill-baseline-vs-candidate.tsv`):

| workbook | expected ours | common | hash base→cand | improved/regressed | printed lost/gained | gap mean/max base | gap mean/max cand |
| --- | --- | --- | --- | --- | --- | --- | --- |
| train | 1035 | 1035 | 898→918 | 20/0 | 0/34 | 13.148/87.288 | 1.986/32.599 |
| Georgia | 630 | 630 | 570→576 | 6/0 | 0/16 | 9.857/66.326 | 1.029/21.981 |
| Tahoma | 630 | 630 | 588→588 | 0/0 | 0/10 | 12.577/85.528 | 1.962/34.178 |

No measurement rows missing, no printed numeric rows lost, no newly regressed hash flags. The earlier
sign regressions are gone. The remaining candidate max errors are suffix/inactive-fill text shapes
that stay on the prior path; they are not caused by this freeze.

Final Chromium proof command:
`run-print-numfill-gate.sh DIST EXCEL_ROOT OUT`. It exports all six sheets per workbook through
`grid.ExportPdf({what:'sheet'})`, measures the browser PDFs, compares only ours rows with native, and
exits non-zero on any missing oracle key or text/count/coordinate difference.
