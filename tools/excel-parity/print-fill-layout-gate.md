# Printed active `*` fill layout: formatter → painter → PDF, and the F09 numeric gate

Owner: FillEngine. Baseline commit `6fe0e6e2` plus the recorded text-section F09 outcome in
`print-textpad-browser-gate.md`. One general mechanism, no accounting-specific path. Predictions were
locked before the textpad holdout was read (`print-fill-layout-predictions.tsv`); the general rule was then
preregistered in `print-fill-layout-general-predictions.tsv` for the numfill holdout (FillOracle), whose
export had not happened at that point. The textpad holdout F09 rows had already been observed by then and
are reported here as results, not predictions.

## Where the number is made

1. **Formatter** — `CellText.formatExcelNumber` selects the value's section, and
   `stripExcelWidthPlaceholders` turns `_(`/`_)` into one space each, `*x` into one fill character, and `?`
   into `#`; .NET then renders e.g. `" $ 1,234.50 "`. `numberFormatUsesFillForValue` records that the
   active section has an unquoted `*`.
2. **Painter** — `Paint.buildCellPaintItems` (print) asks `cellLookFitted` for the display text with the
   print measurer. The print fit cutoff (`printFitRunFor`) excludes active fills, so Paint's screen-shaped
   fit/hash decision stands. `CellText.printFillLayoutRuns(code, n, formatted, measure, available)` splits
   the active section at its **one active `*x` token**, renders the prefix section with the value, takes the
   suffix as the text after that fill char, and sets
   `count = max 0 floor((available − width(prefix) − width(suffix)) / width(fill))`
   with `available` the printed content width after both `_x` pad advances. Paint emits **two runs**:
   prefix + repeated fill char anchored **left**, suffix anchored **right**; the sub-fill leftover stays
   between the fill run and the suffix (invisible for space fills). Count 0 drops the fill char.
   Fit is then decided on the fixed run (`prefix + suffix`, fill chars excluded) against the printer's
   window — `fixed + adv("0")/2 + 0.875 pt <= cell width` at the page scale, the same intercept as the
   non-fill "Print fit cutoff" — not the inset content width. When it fits, the number prints even where
   Paint's screen-shaped rule had produced hashes; when it does not, the display stays Paint's hashes and
   the PDF's counted hash run (genuine numeric overflow). The earlier F09 gate's hash counts are unchanged
   (the narrow accounting test still prints its counted hashes).
3. **PDF** — `Pdf.appendDrawTextOps` draws each run from the embedded face at `FontSize × scale`; the left
   run at `rectX`, the right run at `rectRight − width`. Fill layouts carry no print-fit run and no italic
   reserve (both checks exclude them), so the hash recount and the italic reserve are untouched. Poppler
   extracts `"$"` and `"1,234.50"` separately; the measurement compares their boxes with the cell fill edges.

Screen is untouched (one collapsed run, the cell's own alignment). Kept on the prior path: no active fill,
two or more active fills, no numeric placeholder, quoted/escaped/inactive stars, rich runs,
wrap/rotation/shrink-to-fit/across-selection, and a fixed run that does not fit (hashes, as before).

## Frozen sources

```
CellText.fs, Paint.fs, tests/print/helpers.fsx            (sha256 in the final message)
print-fill-layout-predictions.tsv         5e8b75bea7ee0a8893b849715067fe4f5a95921abd6943add5d85bbefb4006fe
print-fill-layout-general-predictions.tsv
print-fill-layout-results.tsv
print-fill-layout-gate.md
```

## Result (native PDFs; Chromium identical)

Native runner built from this tree; dist `/var/tmp/fill-engine/dist` from one Fable build of the same tree.
Excel truth: `/var/tmp/textpad-oracle/{train-excel,holdout-*-excel,escapes-excel}`.

| set | face/scale | F09 numeric group Δ before | after | F09 text rows | other rows changed | non-F09 numeric max Δ | text max Δ |
| --- | --- | --- | --- | --- | --- | --- | --- |
| training | Aptos Narrow 14, 100 % | 11.702 / 6.959 / 12.627 pt | **0.452 / 1.172 / 1.052** | 0 | 0 | 0.603 | 0.645 |
| holdout Georgia | Georgia 12, 75 % | 10.552 / 6.088 / 10.756 | **0.461 / 0.941 / 0.461** | 0 | 0 | 0.573 | 0.598 |
| holdout Tahoma | Tahoma 10, 100 % | 8.334 / 5.461 / 10.013 | **0.170 / 1.150 / 1.390** | 0 | 0 | 0.711 | 0.591 |
| numeric escapes | Aptos Narrow 14, 100 % | – (no F09) | – | – | 0 | 0.681 | 0.614 |

Provisional numfill training signal (FillOracle's public `synthetic-print-numfill.xlsx`, its Excel PDFs and
`measure-print-numfill.py`; the fixture/tools were being edited concurrently, so this is a direction check,
not a frozen gate): 813 matching cells. Hash agreement 180 match / 105 mismatch (141 mismatches before the
fit override, 108 before the intercept); amount gap within 1 pt 595 cells (was 565); fill-count equal 479
(was 455). Residuals are single-step fit-boundary cells in both directions, with our fixed run 0.15 pt
above FillOracle's measured advance for "1,234.50"; that calibration belongs to the numfill face.

Per-cell symbol/amount boxes: `print-fill-layout-results.tsv`. The amount is right-anchored in every
alignment in both Excel and ours (ours uniform, spread 0.000); the `$` prefix is left-anchored in ours. Excel
itself moves `$` 0.45-1.39 pt right on centre/right sheets, inside one fill advance; that residual is
open and is what the numfill width sweep (FillOracle) should settle.

Chromium `ExportPdf({what:'sheet'})` from the frozen dist equals native on every measured word box:
training 222 rows, Georgia 150, Tahoma 150, escapes 132 — max word-box delta **0.00000 pt**, 0 text
differences; no page errors.

## Evidence for the general shape (FillOracle training, adopted)

FillOracle's fresh Excel training (public scratch workbook, Aptos Narrow 14, 100 %, 0.25-char width steps)
showed: the before-fill segment is anchored left and the after-fill segment right irrespective of cell
alignment; `* #,##0.00` right-anchors its number; `0.00*-` keeps the number left and runs the `-`s rightward;
the repeat count is the largest whole count that fits and can be 0; quoted/escaped/inactive `*` matches the
no-fill twin. Tests here cover each of those shapes at the helper and Paint levels, plus count 0 vs a real
all-hash overflow.

## Commands

```
dotnet build src/Spreadsheet/Rbt.Spreadsheet.fsproj
dotnet build tools/spreadsheet-pdf-runner/SpreadsheetPdfRunner.fsproj
RB_WORD_FONT_DIR=artifacts/fonts/windows:artifacts/fonts/windows-extra \
  tools/spreadsheet-pdf-runner/bin/Debug/net10.0/SpreadsheetPdfRunner tests/test-documents-excel/synthetic-print-textpad.xlsx OUT/ours-N.pdf N-1
python3 tools/excel-parity/measure-print-textpad.py tools/excel-parity/print-textpad-cells.tsv \
  /var/tmp/textpad-oracle/train-excel /var/tmp/textpad-oracle/train-excel/synthetic-print-textpad.tsv OUT OUT.tsv
PLAYWRIGHT=tools/browser-runner/node_modules/playwright/index.mjs \
  node tools/xlsx-browser/browser-textpad-proof.mjs --dist /var/tmp/fill-engine/dist \
  --fixture tests/test-documents-excel/synthetic-print-textpad.xlsx --out OUT --fonts artifacts/fonts/windows-extra --label synthetic-print-textpad
bash scripts/build-dist.sh /var/tmp/fill-engine/dist
dotnet fsi tests/print/helpers.fsx      # 322 passed, 0 failed
```
