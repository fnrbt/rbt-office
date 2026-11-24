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
| trailing fixture generator sets | `--set trail-train`, `--set trail-holdout` in the same generator |
| trailing training workbook (Aptos Narrow 14, 100 %) | `tests/test-documents-excel/synthetic-print-numfill-trail.xlsx` |
| trailing holdout (Arial 11, 60 %; fresh font and scale) | `tests/test-documents-excel/synthetic-print-numfill-trail-holdout.xlsx` |
| trailing manifests | `tools/excel-parity/print-numfill-trail{,-holdout}-cells.tsv` |
| trailing locked predictions | `tools/excel-parity/print-numfill-trail-holdout-predictions.tsv` (sha256 `fdeeaf00...`) |
| trailing model / checks | `tools/excel-parity/analyze-print-numfill-trail.py` |
| trailing freeze gate | `tools/excel-parity/run-print-numfill-trail-gate.sh` |

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

This freeze is the frozen previous release: origin/master `bf132a48`, dist manifest `b38ad676`,
preserved at `/var/tmp/fill-engine/dist`. That dist is the baseline for every later round. The table
below was produced against an older snapshot (`6fe0e6e2`) and is kept as history; do not use
`6fe0e6e2` as the baseline again.

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

## Trailing suffix-empty fills (`0.00* `, `0.00*.`, `0.00*-;0.00`)

A second, denser oracle isolates the fill token when it ends the active section. Four formats
`0.00` (control), `0.00* `, `0.00*.`, `0.00*-;0.00`; values +1234.5/-1234.5/0; left/centre/right;
widths 5.5..12.0 in 0.25-char steps split into four band sheets per alignment (12 sheets per
workbook), so every count transition from 0 upward for zero and from the hash boundary for positive
and negative is bracketed. The F8 negative section is `0.00`; a multi-section format prints the
negative section on the absolute value, so F8 negative is `1234.50` with no minus and no fill (its
fit/hash control is the positive control, not the negative one).

Training is Aptos Narrow 14 at 100 % (972 measured cells). Holdout is Arial 11 at 60 %: a fresh face
and a fresh print scale. Its predictions were locked before export
(`print-numfill-trail-holdout-predictions.tsv`, sha256 `fdeeaf000c51c7693044f6482a25e6a072329acc4cefb4a71a008b8e90e3770b`).

Measured model (Excel truth, both faces):

- `N = max(0, floor((cell_width - pad - w(number)) / w(fill)))` with one `pad` per face/scale:
  Aptos 14/100 % pad in `[4.22, 4.32]` pt (mid 4.27), Arial 11/60 % pad exactly 2.64 pt; the law is
  exact for all 140/160 left-aligned fill cells. `w` is the 600-dpi printer advance
  (`round(glyphEm*round(size*scale*600/72))*72/600`). `pad` is smaller than the sum of the text
  insets (train 4.27 vs 4.479, holdout 2.64 vs ~2.96); it is not derivable from them.
- `number + N*fill` is one run placed by the cell's own alignment: left keeps the control's amount
  x0/x1 (max delta 0.000 pt); centre shifts x0 by `N*w/2` (residual -0.025..+0.098 pt); right by
  `N*w` (-0.050..+0.212). Count 0 is the control cell exactly (<= 0.011 pt), with no fill glyph.
- The fill run is maximal but leaves 0.8-3.0 pt before the border. Count is monotone in width: step
  <= 1 per 0.25 char for space/dash, <= 2 for dot.
- Trailing fill never changes overflow: hash equals the matched control at every width/alignment.
- Holdout advance check: space/dot 1.802-1.835 pt (model 1.800), dash 2.158-2.198 (model 2.160).

All 12 locked checks pass on the fresh holdout Excel truth with 0 failures (edges, centre/right
shift, text, count-0 identity, inside-cell, residual, advance, hash equivalence, monotonicity, sign,
and the content-width law with same-PDF insets to |N_pred-N| <= 1).

Measurement correction (tooling): `measure-print-numfill.py` now counts fill glyphs in emitted text
order. Excel's first trailing fill glyph starts about 0.07 pt before the number's `/Widths` advance
end (ink/advance boxes overlap), so an x-threshold count silently dropped it: 8 emitted spaces were
counted as 7, 6 dots as 5. The corrected count adds one wherever a fill glyph exists, on Excel and
ours alike. New trailing columns `composite_x1` and `fill_adv` are appended; old comparison TSVs are
not rewritten.

Baseline for the trailing round is the frozen previous release (`bf132a48`, dist manifest `b38ad676`,
kept at `/var/tmp/fill-engine/dist`). On the new workbooks it is browser==native exactly, but against
Excel it still caps `0.00* ` at one space and emits no dots for `0.00*.`:
holdout ours-vs-Excel hash 18, count 423, amount_x0 413, amount_x1 595; count mismatches by
format/alignment S2 50, D2 54, F8 37 (each of left/centre/right); train hash 51, count 327,
amount_x0 481, amount_x1 481. Candidate numbers from the trailing gate replace these when the
candidate freeze lands.

Gate: `run-print-numfill-trail-gate.sh CANDIDATE_DIST EXCEL_ROOT OUT [BASELINE_DIST]
[CANDIDATE_NATIVE_DIR] [BASELINE_NATIVE_DIR]`. It keeps two verdicts separate:

- `ABSOLUTE_0.35`: candidate ours vs Excel absolute amount x0/x1 within 0.35 pt on every cell. This
  stays a reported FAIL when the pre-existing face/scale glyph-model offset puts no-fill controls
  outside 0.35 pt; it is not the promotion gate, and the failure list is written to
  `<stem>-absolute-failures.tsv`.
- `PAIRED_RELEASE`: the scoped change gate. Required: complete ours-only keys against the manifest
  for candidate and baseline, browser==native (text/hash/count/coordinates), N1 and F8-negative
  no-fill controls unchanged from baseline, no sign/text/hash regression (baseline correct ->
  candidate wrong), all preregistered relative checks passing on candidate browser and native, and
  every newly worsened count/geometry cell written with its magnitude to
  `<stem>-paired-regressions.tsv`. Material geometry regressions (a cell newly outside 0.35 pt) fail
  the paired gate; +/-1 count regressions are reported for explanation rather than hidden in an
  aggregate.

Expected baseline here is the frozen previous release `/var/tmp/fill-engine/dist` (`bf132a48`, dist
manifest `b38ad676`); the candidate owns a distinct dist and native runner. Do not rebuild the older
`6fe0e6e2` snapshot. Commands:

```
python3 tools/excel-parity/make-print-numfill-fixture.py --set trail-train \
  --out tests/test-documents-excel/synthetic-print-numfill-trail.xlsx \
  --cells tools/excel-parity/print-numfill-trail-cells.tsv
tools/excel-parity/excel-print-pages.sh tests/test-documents-excel/synthetic-print-numfill-trail.xlsx \
  /var/tmp/numfill-oracle/trail-train
python3 tools/excel-parity/measure-print-numfill.py tools/excel-parity/print-numfill-trail-cells.tsv \
  /var/tmp/numfill-oracle/trail-train /var/tmp/numfill-oracle/trail-train/synthetic-print-numfill-trail.tsv \
  OURS_DIR /var/tmp/numfill-oracle/trail-train-comparison.tsv
python3 tools/excel-parity/analyze-print-numfill-trail.py train \
  /var/tmp/numfill-oracle/trail-train-comparison.tsv /var/tmp/numfill-oracle/trail \
  --face 'Aptos Narrow' --size 14 --scale 100
python3 tools/excel-parity/analyze-print-numfill-trail.py check \
  tools/excel-parity/print-numfill-trail-holdout-predictions.tsv \
  /var/tmp/numfill-oracle/trail-holdout-comparison.tsv synthetic-print-numfill-trail-holdout \
  --face Arial --size 11 --scale 60
```

## Trailing candidate paired result (candidate dist-suffix-empty; baseline b38ad676)

Candidate: `/var/tmp/fill-engine/dist-suffix-empty` (CellText.js `985a3020…`, Paint.js `9be221cc…`)
and native runner `/var/tmp/fill-engine/native-suffix/Release/net10.0`. Baseline: the frozen previous
release `/var/tmp/fill-engine/dist` with the repo Debug native runner. Both releases print browser ==
native exactly (0 text/count/hash/coordinate differences on both workbooks), and every ours key set
equals the manifest. All 12 preregistered relative checks pass on both, browser and native.

The strict absolute verdict stays FAIL, as it must for this pre-existing face/scale glyph-model
offset; the failed cell list is in `<stem>-absolute-failures.tsv`:

| set | absolute 0.35 (control x0/x1, trailing x0/x1, count) | count fixed / regressed | count-error direction | geometry newly outside 0.35 | worst new error |
| --- | --- | --- | --- | --- | --- |
| train | 100 / 100 / 354 / 347 / 60 | 333 / 9 | 60 × +1 | 3 (all right) | x0 4.100 pt (baseline 0.186) |
| holdout | 92 / 116 / 233 / 366 / 54 | 417 / 6 | 24 × +1, 30 × −1 | 6 (2 count-driven, 4 geometry-only) | x1 1.510 pt (baseline 0.323) |

Paired release verdict: controls (N1 and F8-negative no-fill) are byte-identical candidate vs
baseline; sign/text/hash regressions are 0; no number is lost; `ours-vs-Excel` hash is 0 on both sets
(baseline: train 51, holdout 18). Counts improve 333 and 417 cells. The regressions are one extra
trailing glyph at the 0→1 boundary, always where Excel emits 0 and the candidate emits 1 (train:
D2 zero 5.75, D2 neg 10.25, F8 pos 10.5; holdout: D2 neg 9.0, F8 pos 8.5). Because centre/right place
`number + N*fill` as one aligned block, that one wrong glyph shifts the number by half/one fill
advance, taking those cells outside 0.35 pt: train right D2 zero 5.75 x0 0.249→3.394, right D2 neg
10.25 x0 0.206→3.438, right F8 pos 10.5 x0 0.186→4.100; holdout centre D2 neg 9.0 x0 0.332→1.248,
right D2 neg 9.0 x1 0.323→1.510. Four further holdout cells (D2 pos 8.5/8.75 and D2 neg 9.25/9.5,
centre x1) cross 0.35 by 0.22–0.26 pt with unchanged count/text/hash. The training face is uniformly
over by one (60/60), the fresh face has 24 over and 30 under: the engine's reserved content width is a
boundary-level approximation, and per-face/scale tuning is out of bounds, so these stay documented
residuals rather than a calibration. The paired gate therefore exits FAIL on the newly-outside-0.35
condition (3 train cells, 6 holdout cells); everything else in the paired gate passes. Promotion
needs an explicit decision on those boundary cells, not a relaxed 0.35 absolute rule.

Per-cell detail: `<stem>-paired-regressions.tsv` (baseline-vs-candidate count and x0/x1 error deltas),
`<stem>-absolute-failures.tsv`, `<stem>-candidate-{browser,native}-check.tsv`. Run:

```
tools/excel-parity/run-print-numfill-trail-gate.sh /var/tmp/fill-engine/dist-suffix-empty \
  /var/tmp/numfill-oracle OUT /var/tmp/fill-engine/dist \
  /var/tmp/fill-engine/native-suffix/Release/net10.0 \
  tools/spreadsheet-pdf-runner/bin/Debug/net10.0
```
