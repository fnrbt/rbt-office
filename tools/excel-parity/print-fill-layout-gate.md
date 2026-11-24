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
PDFs were produced from this tree before the Excel PDFs existed, then measured against the locked export.
All 12 locked checks pass on Excel truth and on ours (ours_failures 0 on both the training and holdout
workbooks): hash/text/sign, inside, residual, advance, monotonicity, count law (|N_pred−N| ≤ 1) and the
side-relative left/centre/right relations. Chromium equals native on every measurement row of both
workbooks (identical long-table sha256; largest run-origin difference 0.0001 pt). This is a paired
improvement, not absolute Excel parity: on the holdout, candidate-vs-Excel is hash 18 → 0, count 423 → 54,
amount_x0 413 → 233, amount_x1 595 → 366 versus the pre-rule dist, and text 0 throughout. Paired with the
same baseline over the lock's classes: the no-fill/default-alignment N1 cells are byte-identical (243/243
per set), genuine overflow hash runs are unchanged (0 text differences on 267 holdout / 363 training
overflow cells), and D2 dot-as-decimal keeps the decimal point and matches Excel's amount text on all
174 holdout / 150 training non-hash cells. Where candidate and baseline counts both equal Excel's, no
cell's x0/x1 error worsens by more than 0.10 pt (0 of 531 holdout and 0 of 459 training cells); the
count distance to Excel improves on 417 vs 6 (holdout) and 333 vs 9 (training) cells. The fallback cannot
drop an arbitrary trailing glyph: it only
runs when the token is the section's last element (`suffixSection = ""`) and the whole display equals the
formatter's prefix rendering; a non-empty suffix (quoted literal included), a surviving token char, two
active fills, or a full display wider than the cell all keep the prior path, and the sign is taken from the
prefix so it is never duplicated.

Limitations: (1) the engine's content-width available is a fraction of a point wider than Excel's at a
boundary width, so a fill count can be one high: paired strict error counts 5 (holdout) and 12 (training)
cells where the count moves 0→1 against Excel, and the holdout count is exact on 675/729 trailing cells
with the rest ±1 (never more). No sign, text, hash or fit change comes with it. (2) A trailing `.` is not
part of the .NET display string, so the rule reads the fill run from the formatter's number end rather than
from a guessed dot; a non-space/non-dot trailing token is covered by the fresh dash fixture, not by a
separate fallback. (3) Absolutely positioned x0/x1 versus Excel is not gated here: the no-fill N1 controls
alone already differ by >0.35 pt on roughly two thirds of the cells at these printer scales (pre-existing
face/scale placement), which is why the lock uses side-relative relations. No per-face or per-scale
calibration was added. (4) Screen, quoted/escaped and inactive-section paths are untouched: the branch runs
only for printing with one active fill and an empty suffix section. A whole-corpus diff against the pre-rule
dist confirms it: on FillOracle's 1035-cell training workbook, only `S2_space_after`, `D2_dot_after` and
`I2_fill_pos_only` (the suffix-empty classes) change; leading-fill `S1`/`D1`, currency and accounting
`W_C1`/`W_C2`/`W_A1`/`W_A2`/`W_I3`, quoted `Q1`, escaped `E1`, inactive `I1` and both no-fill controls are
identical in text, hash, count and every measured coordinate.

### Repeat-width / over-fill investigation (falsified for a no-calibration runtime rule)

The paired gate's worst new error is one extra fill on a right-aligned cell: 4.120 pt on training
`T3_right` row 11 width 10.5, F8 zero (Excel 1 dash, ours 2; one dash advance is 4.296) and 1.510 pt on
the holdout D2 negative. Filed against the pre-rule dist the same cells were within 0.25-0.49 pt, so this
is a real over-fill, not a placement regression elsewhere.

Tested as one principled rule on both faces using the no-fill controls:
- Advance is not the cause (measured vs 600-dpi model: space 2.626/1.802 vs 2.640/1.800, dot 3.594-3.650
  vs 3.600, dash 4.296/2.158 vs 4.320/2.160).
- `N = floor((W − lin − rin − number)/adv)` with `lin`/`rin` from the same-width N1 left/right controls is
  never high, and low on 45/420 (training) and 66/480 (holdout): the fill region is wider than the no-fill
  text box. Quantizing the slack and advance to 600-dpi dots (floor/round/ceil) is identical or worse on
  the Excel side; on engine-side quantities floor helps training (372 vs 366 exact) and hurts the holdout
  (384 vs 423). A symmetric `W − 2·lin` box is worse still. A constant repeat pad exists only per face
  (training interval [4.136, 4.236], holdout empty by 0.013), and choosing it is calibration.
- Cross-fixture check on that sub-finding: the training and holdout exports draw DIFFERENT border ink —
  trail-train has a 0.96 pt filled strip plus a 0.14 pt stroked centreline, trail-holdout has 0.60 pt
  filled strips and no stroke — yet both have the same exact correction interval **[0.24, 0.35]**
  (adding a constant c in that interval to `W − lin − rin − number` makes the count law exact on all
  420 training and 480 holdout fill cells). Width-proportional models are therefore falsified across the
  pair: S/2 would be 0.48 vs 0.30, S/4 0.24 vs 0.15, T/2 0.07 vs 0. A constant device quantum of about
  2.5 printer dots survives; whether it is a border-presence term or a text/measurement term is what the
  `none` sheet decides. (Earlier wording that called this half a 0.6 pt strip was wrong.)
- It is not implementable without calibration today: the engine's own no-fill insets already make its
  text box 0.342-0.420 pt wider than Excel's on training and 0.449-0.738 pt wider on the holdout, so the
  constant term cannot be applied to engine inputs either; the best single engine-side correction
  (0.36 pt) still leaves 6/420 and 11/480 mismatches.

So the runtime rule is falsified as stated; no runtime change, no new calibration, absolute gate intact.
The border sweep has now resolved the border question (TrailingTruth's direct sweep,
`print-border-fill-oracle.md`: 1080 fill cases on 1728 measured cells per face, 12 sheets x 12 rows x
12 widths — the 100 % sheets' final 11.25 column paginates horizontally and is carried by identical
row y positions — none/thin/medium/thick, Arial 11/60, Aptos Narrow 14/100, Calibri 11/100):

- Border weight does nothing: count, amount x0/x1, composite end and residual are identical under all
  four weights (0 differing cells across all 1728 measured cells per face; transition content width
  shift 0.000 pt; residual minimum style-invariant: Arial 1.176, Aptos 2.019, Calibri 1.654 pt).
- Border ink does vary: Arial strips 0.60/1.20/1.68 pt (no stroke); Aptos and Calibri 0.96+0.14
  stroke / 1.92 / 2.88 pt — so the absence of any count/geometry change is not an absence of border ink.
- The correction interval is identical across weights for each face, and the no-border sheet carries the
  same ~0.24-0.33 correction. Border-presence and width-proportional HALF_STRIP are falsified; NONE (0)
  fails on every face.
- With **Excel's measured number width**, one global constant `c ∈ [0.25, 0.32]` gives 0 count failures
  on all 1080 fill cases (`c = 0.24` sits at the rounding edge and misses 20: Aptos 4, Arial 12,
  Calibri 4). With the **600-dpi predicted number width** the intervals are face-dependent and do not
  intersect (Aptos [0.159, 0.279], Arial [0.310, 0.417], Calibri [0.313, 0.403]; best single constant
  8/1080 at 0.32-0.36, 40/1080 at 0.24-0.26), so the remaining blocker is the predicted text
  width/advance model, not the fill boundary.
- That is exactly the advance stage of the decomposition below (ours +0.03 pt Aptos, +0.14 pt Arial
  against Excel's measured width). A constant is not implementable as an engine rule until the
  text-width/placement model is fixed; no runtime change, no fitted constant.

Frozen report/evidence: `print-border-fill-oracle.md` (`180faee5...`), `measure-print-border-fill.py`
(`121fb9d4...`), `border-report.txt` (`0980d7e7...`), `border-cases.tsv` (`5fc5c58c...`),
`border-summary.tsv` (`23b403d2...`). Re-locked predictions for the actual design were
`border-sweep-predictions-v2.tsv` (`9a98145f...`); v1 (`78c78b56...`) assumed only thin/medium/thick
and is superseded.

Resolved by `print-inset-oracle.md`. The pad is the cell font's `ceil(d/4)+4` dots, not the Normal font's; widths are hinted GDI advances; Excel's fill rectangles are one dot wider than the column. With those, `count = floor((cw - 2P + 2 - w)/a)` is exact on all 4,140 fill cells. The engine now uses it, and the paired gate has 0 count errors.

#### Where the 0.34-0.74 pt no-fill box gap comes from

Cross-sheet decomposition of `slack = W − lin − rin − number` on all fill cells with an amount (left inset
from the left-aligned N1 control, right inset from the right-aligned one):

| face | border ΔW | advance −Δw | left inset −Δlin | right inset −Δrin | gap before s/2 | gap after s/2 |
| --- | --- | --- | --- | --- | --- | --- |
| Aptos Narrow 14/100 | −0.007 | +0.028 | +0.780 | −0.192 | +0.610 | +0.370 (0.342-0.390) |
| Arial 11/60 | −0.024 | +0.135 | +0.290 | +0.363 | +0.764 | +0.524 (0.420-0.600) |

The cell border span contributes nothing net. The number glyph advance is small (ours 0.03-0.14 pt
narrower). The placement insets dominate. The engine's inset is one symmetric pad fixed by the workbook
Normal font (Calibri-11 47-dot digit) and the print scale: 1.920 pt at 100 %, 1.440 at 75 %, 1.152 at
60 %, the same for every cell font. Excel's insets follow the cell font/size and are asymmetric: at
100 % Aptos Narrow 14 has left 2.700 / right 1.78 and Tahoma 10 has 2.340 / 1.47; Georgia 12 at 75 % has
1.920; Arial 11 at 60 % has 1.44-1.47 / 1.49-1.63. The exact cell-font -> inset law (including the
0.87-0.92 pt left-right asymmetry on Aptos/Tahoma) is underdetermined; the next observable is a direct
export grid of plain `0.00`/`1234.50`/`-1234.50`, left and right aligned at fixed column width, over cell
fonts {Aptos Narrow, Arial, Tahoma, Georgia, Calibri} x sizes {9,10,11,12,14} x scales {60,75,100} with
the Normal font held fixed, measuring border span, number advance, left inset and right inset in printer
dots. Acceptance: one zero-parameter rule predicts all four within one 600-dpi dot. Evidence:
`/var/tmp/fill-engine/trail-proof/no-fill-box-decomposition.txt` (`d03d8d80...`).

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
