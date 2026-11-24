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
   The split checks that the fill character really sits where the rendered prefix ends. A negative value
   under a single-section code is formatted with a leading `-` (`* #,##0.00` -> `- 1,234.50`); that sign is
   carried as part of the left-anchored prefix, so the fixed width includes the sign and the number is never
   printed without it. A sign where the fill character belongs makes the section fall back to the prior path.
   A fill token with nothing after it (`0.00* `, `0.00*.`, `0.00*-`) has no right-hand suffix: Paint emits
   **one run** — number + repeated fill char — placed with the cell's own alignment, not left-anchored.
   Excel pads the number with the fill characters using the same left-anchored count and then centres or
   right-aligns the whole padded string, so count 0 is byte-equal to the no-fill twin and the number moves
   left by half/whole fill advances on centre/right. A trailing `.` fill is consumed by .NET (`0.00*.`
   renders `1,234.50`): when the token is last and the display is exactly the prefix rendering, the fill
   run starts at the formatter's number end; no dot is inserted into the number string. The fit window
   still uses the number-only fixed run (fill characters excluded).
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

Numfill baseline→candidate on the frozen native runner (FillOracle's corpus and measurement; identical
workbooks, baseline `6fe0e6e2`):
- train 1035 cells: Excel hashes 183/183 reproduced (baseline 183/183); Excel numbers hashed by the engine
  fall 137 -> 117 (20 fixed, 0 newly hashed); 153/153 negative two-number cells keep their sign; amount gap
  mean 13.148 -> 1.550 pt, max 87.288 -> 32.599; 0 cells newly over 1 pt, 88 newly within 1 pt.
- Georgia holdout 630: 45/45 Excel hashes reproduced, 0 sign mismatches; gap mean 9.857 -> 0.575, max
  66.326 -> 6.161 (FillOracle's browser/native comparison; strict Chromium check below).
- Tahoma holdout 630: no Excel hash cells, 0 sign mismatches; gap mean 12.577 -> 1.166, max 85.528 -> 19.287.
FillOracle's independent strict gate on the same freeze also passes (ours-only nonempty key sets checked
against the cell manifests, so a missing row fails): train 1035 / Georgia 630 / Tahoma 630 ours rows,
browser-vs-native 0.00000 pt, 0 diffs. Its corrected paired table
(`print-numfill-baseline-vs-candidate.tsv`) records 898 -> 918 hash rows train (20 improved / 0 regressed),
570 -> 576 Georgia (6/0), 588 -> 588 Tahoma (0/0), with zero missing rows, zero lost printed rows and zero
new hash regressions. The fit window uses FillOracle's calibrated intercept. The suffix-empty layouts that
previously stayed on the prior path are now the single aligned run described above (see the trailing-fill
result below); inactive-section text shapes are unchanged and remain on the prior path.

### Trailing fills (`0.00* `, `0.00*.`, `0.00*-`): suffix-empty run

Training Excel truth (Aptos Narrow 14, 100 %, direct Excel 2019; S2 `0.00* ` and D2 `0.00*.` against the
matched N1 `0.00`, 144 cells over three alignments, three values and eight widths): 0 hash mismatches and
0 amount-text mismatches. Emitted text-order fill counts are exact on 117/144 cells and one high on 15
(never low, never more than one) — the engine's content-width available is a fraction of a point wider at
exactly the width steps where one more fill still fits, so counts move one half-step early. Allowing the
count difference, the suffix-specific amount position matches Excel within 0.13 pt (median 0.011, 42
count-equal cells compared after removing the engine's general N1-control offset). The trailing rule adds
no hash change: the narrow S2 cells that Paint had hashed now print their number at count 0 (Excel prints
the number too), and no training cell newly hashes.

Fresh held-out export (Arial 11 pt at 60 %, TrailingTruth's locked predictions): the runtime and Chromium
PDFs were produced from this tree before the Excel PDFs existed. On those PDFs, checked against the
engine's own N1 controls (the side-relative relations in the lock), all 729 trailing cells pass hash/text/
sign, inside, residual and count-law (|N_pred−N| ≤ 1; 0 violations), with a maximum centre/right shift
residual of 0.075 pt against the 0.35 pt bound. Chromium equals native on every run of that workbook
(largest origin difference 0.0001 pt, one run). The Excel-side check still requires the locked fresh
export, which had not been taken when this tree was frozen.

Limitations: (1) the measured fill advance can put the engine one fill glyph ahead of Excel at a boundary
width (count +1 above; not a sign/hash/fit error). (2) A trailing dot is not part of the .NET display
string, so the rule reads the fill run from the formatter's number end rather than from a guessed dot; a
non-space/non-dot trailing token is covered only by the fresh dash fixture. (3) Screen, quoted/escaped and
inactive-section paths are untouched: the new branch runs only for printing with one active fill and an
empty suffix section.

Per-cell symbol/amount boxes: `print-fill-layout-results.tsv`. The amount is right-anchored in every
alignment in both Excel and ours (ours uniform, spread 0.000); the `$` prefix is left-anchored in ours. Excel
itself moves `$` 0.45-1.39 pt right on centre/right sheets, inside one fill advance; that residual is
open and is what the numfill width sweep (FillOracle) should settle.

Chromium `ExportPdf({what:'sheet'})` from the frozen dist equals native on every measured word box:
training 222 rows, Georgia 150, Tahoma 150, escapes 132 — max word-box delta **0.00000 pt**, 0 text
differences; no page errors. The numfill gate (`run-print-numfill-gate.sh`) on the same dist passes
strictly: train 1035, Georgia 630, Tahoma 630 ours rows, 0 text/count/coord differences, same 0.00000 pt.

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
dotnet fsi tests/print/helpers.fsx      # 324 passed, 0 failed
```
