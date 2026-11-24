# Printed text under the fourth number-format section: final Chromium gate

Frozen sources (no commit; verified before and after the one dist build):

```
CellText.fs                  7bf6858ca9c364ca999276497b6a9122f473e6d0c32b73918f368f1d1495d2fc
Paint.fs                     d1414458421b046e86514499d0f3174d06a0327678de595252d0e6a7c879351c
XlsxReader.fs                ce94d44a145ff41cd71e86393c6d8736d1493cee66c719bbe9bb06b65c51cee2
tests/print/helpers.fsx      9bc25657fdd408a87235312c7bf6041f7c9f55ad44fa61803eaf196cb8ec0b03
tests/office-formats/xlsx-reader.fsx  10cd3c4a2a6bc1ffb9897ffe9aec1892a6ef2d705dc9ab1b18160fc5d3ec9644
docs/excel-parity-tracker.md 52560827a6dbeb52b0263c52fc74da00facf80dc58efeeeaf0520e418859a145
```

| role | path |
| --- | --- |
| frozen dist (one Fable build) | `/var/tmp/textpad-dist` |
| native runner (same tree) | `tools/spreadsheet-pdf-runner/bin/Debug/net10.0/SpreadsheetPdfRunner` |
| Chromium harness | `tools/xlsx-browser/browser-textpad-proof.mjs` |
| measurement | `tools/excel-parity/measure-print-textpad.py` |
| gate summary | `tools/excel-parity/print-textpad-chromium-gate.tsv` |
| gate runner | `tools/excel-parity/run-print-textpad-gate.sh` |
| fixtures | `tests/test-documents-excel/synthetic-print-textpad*.xlsx` |
| Excel truth / PDFs | `/var/tmp/textpad-oracle/{train-excel,holdout-*-excel,escapes-excel}` |
| browser/native PDFs + full tables | `/var/tmp/textpad-gate/out` |

Every export is a real Chromium PDF from the Fable `Spreadsheet` component
(`ExportPdf({what:'sheet'})`), measured against the fresh direct Excel 2019 PDFs with Poppler's word box
against the printed solid-fill cell edges. The browser PDFs embed Aptos Narrow, Georgia and Tahoma; no
page errors.

```
node tools/xlsx-browser/browser-textpad-proof.mjs --dist /var/tmp/textpad-dist \
  --fixture tests/test-documents-excel/synthetic-print-textpad.xlsx \
  --out DIR --fonts artifacts/fonts/windows-extra --label synthetic-print-textpad
PLAYWRIGHT=tools/browser-runner/node_modules/playwright/index.mjs ...
tools/excel-parity/run-print-textpad-gate.sh /var/tmp/textpad-dist /var/tmp/textpad-oracle /var/tmp/textpad-gate/out
```

## Result

| set | text rows | max |delta| | mean | text failures | numeric non-F09 max | F09 numeric max | browser - native |
| --- | --- | --- | --- | --- | --- | --- | --- |
| training Aptos Narrow 14, 100 % | 54 | 0.645 pt | 0.523 | 0 | 0.603 | 12.627 | 0.00000 pt |
| holdout Georgia 12, 75 % | 36 | 0.598 | 0.453 | 0 | 0.573 | 10.756 | 0.00000 |
| holdout Tahoma 10, 100 % | 36 | 0.591 | 0.232 | 0 | 0.711 | 10.013 | 0.00000 |
| numeric escapes Aptos Narrow 14 | 33 | 0.614 | 0.507 | 0 | 0.681 | - | 0.00000 |

Browser and native word boxes agree to 0.00000 pt on every text and numeric cell and every displayed
string, so `print-textpad-reference.tsv` and `print-textpad-holdout-reference.tsv` (native, Excel vs
before/after) are also the row-level Chromium result. There is no text regression and no counterexample
failure: quoted/escaped literals print, empty fourth sections hide text, two-/three-section `@` codes stay
as stored, and centre/right/left pad geometry matches.

The three numeric misses per text set are all standard accounting (F09) fill-layout cells: Excel's `$` and
number are placed by its `*` fill expansion, ours differently. They are identical in the pre-change PDFs and
in Chromium vs native, are outside the fourth-section text change, and were excluded by the preregistered
numeric prediction. Every other numeric control is within the ordinary 0.6-0.7 pt page/font delta.
