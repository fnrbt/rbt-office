# Trial verification: scaled-dot printed column pitch

Independent acceptance check of the trial now committed as `2e691804`
("Snap printed column widths to printer dots at the page scale") in
`src/Spreadsheet/Pagination.fs` + `src/Spreadsheet.Pdf/Pdf.fs` plus its unit
test.  No product source was changed by this check and nothing was committed
by it.

## The trial

`Pagination.printColumnWidthPxAtScale` snaps a column's 600-dpi dot width *after*
the page scale is applied; `Pdf.withPrintColumnWidthsAtScale` re-expresses the
PDF geometry widths at that scale.  Pagination is still computed before that
step, from the same scaled `dotted` budget, so page breaks should be unchanged.
At 100 % the change is an identity.

## Source freeze and builds

Hashes before and after both dist builds (unchanged):

```
Pdf.fs         4e5bb851251b1fe845fd3f463d4d81d32a1a55639a8184d13234e7b3f8dc94e0
Pagination.fs  e922f8aebe95ead2421a23b4c12f19f2e78338a5deabb6fd60b5ad4a113699ea
CellText.fs    ee6d6ea8609bf338cee3f406b86510665572fdc9c970e542313e80399dcfb430
Paint.fs       e9bec9d3fb9901240ca9495077c33c3d0b684bf8228e9cd25f5c32376527548b
```

Baseline is the same tree at `af0270fb` (`Pdf.fs` 8f04d59a, `Pagination.fs`
dd6d4d86); CellText/Paint are identical.  The trial build's source hashes
equal `HEAD` (`2e691804`) before and after the build, so the dist below is
from the committed content.

| role | path |
| --- | --- |
| trial Fable dist | `/var/tmp/border-oracle/dist-trial` |
| baseline Fable dist | `/var/tmp/border-oracle/dist-base` |
| trial native runner | `tools/spreadsheet-pdf-runner/bin/Debug/net10.0/SpreadsheetPdfRunner` |
| baseline native runner | `/var/tmp/geometry-model/baseline/tools/spreadsheet-pdf-runner/bin/Debug/net10.0/SpreadsheetPdfRunner` |
| browser harness | `tools/xlsx-browser/browser-count-proof.mjs` (system Chrome; the headless shell hit `ERR_INSUFFICIENT_RESOURCES` under host load) |
| browser fonts | `/var/tmp/fitgate/fonts` |

The trial bundle contains the final `PrintFitFullText` glyph prepass:
`dist-trial/spreadsheet-pdf/Pdf.js` lines 2034-2038 call `recordTextGids` for
`t.Text` and then for `t.PrintFitFullText` when it differs.  It also contains
`withPrintColumnWidthsAtScale` (Pdf.js:1155, called at 2079) and
`printColumnWidthPxAtScale` (Pagination.js).

## Native acceptance

| corpus | baseline | trial |
| --- | --- | --- |
| column-edges 90 % `E90_edge` | **45/81 wrong** (all false hash) | **0/81** |
| column-edges 90 % `H90_shift` | **45/81 wrong** | **0/81** |
| column-edges 85 % `H85_edge` | 0/81 | 0/81 |
| Georgia italic 90 % count holdout (388 cols) | 388/388 decisions, 261/261 counts | **same** |
| fullhash blind holdout (Cambria 80 %, Trebuchet 85 %) | 4/96 decisions wrong | **0/96**, counts 33/33 both |
| fullhash anchor (344 cols) | 15 wrong | **10 wrong**; exactly 5 flips, all Georgia 15 upright 90 % 16.2188-16.3438 hash→full; no other flip |
| count train (786 cols) | 6 decision / 12 count mismatches | **identical** |
| hash-geometry (289 cols) | 0 decision, 277/279 counts | **identical**, 0 row-1 flips |
| page-break corpus direct pages (`breaks`) | 10, 8, 12, 8, 10 | **identical** |
| fit-to-page corpus direct pages (`fit`) | 1,1,1,1,1,2,4,2,2,6,8 | **identical** |
| `dotnet fsi tests/print/helpers.fsx` | - | **283 passed, 0 failed** |

Column-edge border widths and decisions vs the fresh direct Excel 2019 PDFs:

| sheet | baseline border error mean/max | trial border error mean/max | baseline wrong | trial wrong |
| --- | --- | --- | --- | --- |
| `E90_edge` | 0.057 / 0.090 pt | **0.017 / 0.030** | 45 | **0** |
| `H85_edge` | 0.014 / 0.034 | **0.016 / 0.030** | 0 | **0** |
| `H90_shift` | 0.057 / 0.090 | **0.017 / 0.030** | 45 | **0** |

## Chromium acceptance (real `grid.ExportPdf` from each dist)

| corpus | baseline | trial |
| --- | --- | --- |
| column-edges `E90_edge` / `H90_shift` | 45 / 45 wrong; border error 0.057 pt | **0 / 0**; border error **0.017 pt** |
| column-edges `H85_edge` | 0 | 0 |
| Georgia italic 90 % count holdout | 388/388 decisions, 261/261 counts | **same** |
| fullhash blind holdout | 4/96 wrong | **0/96**; counts 33/33 both |
| fullhash anchor | 15 wrong; counts 154/154 | **10 wrong**; **same 5 flips, all gains**; counts 154/154 |
| count train | 6 decision / 12 count mismatches | **identical** |
| hash-geometry | 0 decision, 277/279 counts, 5 missing literal row-2 controls | **identical**, 0 row-1 flips |
| page-break corpus page counts | 10, 8, 12, 8, 10 (matches Excel `pages`) | **identical** |

No regression was found in any corpus: every row-1 change is a gain, and every
pre-existing mismatch is unchanged.  The remaining 10 anchor misses are the
Verdana 13 at 75 % advance residual (70.098 pt printed vs the 69.000 design
model), explicitly out of this change's scope.

## Verdict

Retain the trial.  It fixes all five Georgia 15 upright 90 % columns in the
native and Chromium paths, improves the blind Cambria/Trebuchet holdout from
4 wrong to 0, leaves the Georgia italic cutoff corpus exact, leaves all page
counts and pre-existing residuals unchanged, and its compiled bundle carries
the glyph prepass.  The only remaining anchor errors are the pre-existing
Verdana 75 % advance residual, not caused by this change.

## Reproduction

```
# native
RB_WORD_FONT_DIR=$PWD/artifacts/fonts/windows:$PWD/artifacts/fonts/windows-extra \
  tools/spreadsheet-pdf-runner/bin/Debug/net10.0/SpreadsheetPdfRunner BOOK OUT.pdf SHEET
# browser (system Chrome; harness copy at /var/tmp/border-oracle/debug-proof.mjs)
node tools/xlsx-browser/browser-count-proof.mjs --dist /var/tmp/border-oracle/dist-trial \
  --fixture BOOK --out DIR --fonts /var/tmp/fitgate/fonts
# measurements
python3 tools/excel-parity/measure-print-column-edges.py ...
python3 tools/excel-parity/measure-print-count-transitions.py ...
python3 /var/tmp/fitgate/gate.py measure|compare|diff ...
```
