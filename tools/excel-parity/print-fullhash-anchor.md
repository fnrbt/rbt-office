# Printed full→hash cutoff: independent Excel anchors and blind holdout

Owner: FitCutoffOracle (independent Excel measurement). Runtime candidate: FitCutoffEngine
(`tools/excel-parity/print-fullhash-engine-predictions.tsv` for the holdout lock). Source of truth: fresh
Excel 2019 direct `ExportAsFixedFormat` PDFs, exported before any preview or activation. Physical cell
edges come from per-column solid fill rectangles; the printed run is Poppler's `pdftotext -bbox` word box.
Print-only: nothing here changes hash sizing, placement, counts, or the screen path.

## Question

The shipped printed fit rule (cell less a font inset) and the printer-unit count model disagree with Excel
on 5 known full/hash rows: Verdana 13 upright `0.00%` at widths 17.375/17.4375/17.5 (old fit prints 8
hashes, Excel the number) and Georgia 15 italic 90 % `#,##0.00` at 16.6875/16.75 (old fit prints the
number, Excel 8 hashes). This experiment measures the two independent quantities the decision needs:

- **printed word advance** of the full formatted text in a wide cell (no clipping) - the advance Excel
  actually lays out, not our font-metric estimate;
- **physical cell bounds** at the tight full→hash cutoff - the `W` at which Excel flips, which bounds the
  room constant `c` (`fits iff advance <= W - c`) from both sides.

Matched strings of the same length but different glyph composition (1s vs 8s), literal-string twins,
italic twins, an accounting fill/pad format, print scales 75/80/85/90/100 %, and a blind holdout on a
face/format/scale absent from the engine fixtures are included.

## Tools and data

| Path | What |
| --- | --- |
| `tools/excel-parity/make-print-fullhash-anchor.py` | deterministic generator (`anchor` / `holdout`) |
| `tests/test-documents-excel/synthetic-print-fullhash-anchor.xlsx` | 7 sheets, 344 test columns (7 wide + dense 1/32-char sweeps), row 2 `-98.41` control |
| `tests/test-documents-excel/synthetic-print-fullhash-holdout.xlsx` | 2 sheets, 96 test columns; blind until the lock |
| `tools/excel-parity/print-fullhash-anchor.meta.tsv`, `print-fullhash-holdout.meta.tsv` | per-cell face/size/italic/scale/format/value/role/text (`row` 1 test, 2 control) |
| `tools/excel-parity/measure-print-fullhash-anchor.py` | PDF measurement: fill → cell edges, word box → advanced run, `nHashes`, per-word span (last word for accounting numbers) |
| `tools/excel-parity/predict-print-fullhash-anchor.py` | printer-unit candidate, `--measure` evaluates at the measured `W` |
| `tools/excel-parity/analyze-print-fullhash-anchor.py` | per-column table + per-sheet summary (implied `c`, transitions, false hash/full, newly wrong) |
| `tools/excel-parity/print-fullhash-anchor-results.tsv`, `print-fullhash-holdout-results.tsv` | Excel measurements (one row per printed line) |
| `tools/excel-parity/print-fullhash-anchor-analysis.tsv`(`.summary`), `print-fullhash-holdout-analysis.tsv`(`.summary`) | analysis output |
| `tools/excel-parity/print-fullhash-holdout-predictions-oracle.tsv` | this oracle's lock, written before the holdout export (byte-identical to the draft kept at 01:00, export 01:06) |
| `tools/excel-parity/print-fullhash-engine-predictions.tsv` | FitCutoffEngine's holdout lock (its own candidate, W+0.15 variant, baseline) |

Candidate model used for the lock (same family as `predict-print-count.py`, real Georgia Italic):

```
dots  = floor(size_pt*scale*600/72 + 0.5)
a     = floor(hashEm*dots + 0.5) * 72/600
c     = digitEm*size_pt*scale/2 + 0.875
        + italic ? floor(floor(usWinAscent*dots/upm + 0.5)/5) * 72/600 : 0
W(w)  = floor(screenPx(w)*47/7 + 0.5) * 72/600 * scale,  screenPx(w) = floor(((256*w + 18)/256)*7)
fits iff sum(floor(glyphEm*dots + 0.5) * 72/600) <= W - c        # per-character rounded advances
else n = max(1, floor((W-c)/a) - leading `_x` slots)
```

Faces are the real installed files: `artifacts/fonts/windows/verdana{,i}.ttf`,
`artifacts/fonts/windows-extra/georgia{i}.ttf`, face 0 of `artifacts/fonts/windows/cambria.ttc`,
`artifacts/fonts/windows-extra/trebucit.ttf`. The PDF-embedded names confirm them for every sheet
(`Verdana(-Italic)`, `Georgia(-Italic)`, `Cambria`, `TrebuchetMS-Italic`); the local `windows/GARA*.TTF`
is Garamond and is not used.

## Blind protocol

1. Anchor workbook generated and exported first; there was no earlier Excel export of either workbook.
2. Holdout workbook and meta generated, published to FitCutoffEngine before export; the engine locked
   `print-fullhash-engine-predictions.tsv` from the meta alone and replied "locked, export".
3. This oracle's holdout predictions (`print-fullhash-holdout-predictions-oracle.tsv`, W from the
   pre-registered formula only) were checked to be identical to the pre-export draft.
4. Only then was the holdout exported and measured. The holdout dense windows were re-centred once
   before the lock because the first choice did not straddle a transition; nothing was exported at that
   point.

## Measured result

### Printed advance in a wide cell

Wide columns (22 characters) print the full formatted string; `runWidth` is the Poppler word-box span
(advance box, side bearings included, blank `_x` pads excluded). For accounting, the number word is
measured separately from the leading `$`, which Excel parks at the physical left edge.

| Sheet (face/size/italic/scale/format) | first wide string | model pt | printed pt | bias |
| --- | --- | --- | --- | --- |
| A_V13up_pct (Verdana 13 up 100 % `0.00%`) | `1234567.89%` | 93.120 | 93.142 | +0.022 |
| A_V13it_pct (Verdana 13 it 100 % `0.00%`) | `1234567.89%` | 93.120 | 93.142 | +0.022 |
| A_V13up_pct75 (Verdana 13 up 75 % `0.00%`) | `1234567.89%` | 69.000 | **70.098** | **+1.098** |
| A_V13up_acct (Verdana 13 up 100 % accounting) | `1,234,567` | 67.320 | 67.262 | -0.058 |
| A_G15up_grp (Georgia 15 up 90 % `#,##0.00`) | `1,234,567.00` | 77.640 | 77.658 | +0.018 |
| A_G15it_grp (Georgia 15 it 90 % `#,##0.00`) | `1,234,567.00` | 77.520 | 77.536 | +0.016 |
| A_G15it_pct (Georgia 15 it 100 % `0.00%`) | `1234567.89%` | 89.400 | 89.340 | -0.060 |
| H_Cambria16up_grp80 (Cambria 16 up 80 % `#,##0.00`) | `9,876,543.21` | 71.640 | 71.671 | +0.031 |
| H_Trebuchet12it_pct85 (Trebuchet MS 12 it 85 % `0.00%`) | `123456.78%` | 53.040 | 52.932 | -0.108 |

The model advance is right within 0.11 pt at 80-100 % scale. **At 75 % print scale it is 1.10 pt too
small**: the printed advance implies a layout height of 82 dots where `floor(13*0.75*600/72+0.5)` gives
81 (13*0.75*600/72 = 81.25). The hash pitch at the same sheet is unambiguously the 81-dot advance
(`runWidth/8 = 7.9239` vs `a(81) = 7.9200`, `a(82) = 8.0400`), so the number layout and the `###` run
differ in this one quantisation - an advance-model effect, not a room-constant effect.

Matched strings: Verdana digits are tabular - `1111111.00%` and `8888888.00%` print identical 93.142 pt,
and `111111.00%`/`888888.00%` 84.797; Georgia at 90 % italic is proportional - same-length
`1,111,111.11` prints 63.018 pt and `8,888,888.88` 83.309 (model 63.72/83.16, residual <= 0.14). A model
that treats Georgia's digits as uniform (or uses the Garamond metrics) cannot fit the holdout/anchors.

Right gap (`cellRight - xMax` in the wide cell): Verdana 100 % upright 2.74-2.83, italic 5.26-5.35 -> one
reserve 2.52 pt; Georgia 90 % upright 2.72-2.92, italic 5.12-5.32 -> 2.40 pt; both equal
`printItalicShiftPt` for those faces, independently of the word-box width metric. Accounting 100 %
upright 8.69-8.76 (leading `$` + trailing `_)` pad).

### Full→hash cutoffs and the implied room constant

`c` below is `W_transition - printed_model_text`; the interval is set by the last hash column (strict
lower bound) and the first full column (closed upper bound) at the measured fill width. The candidate's
`c` is the formula above.

| Sheet | last hash w (W pt) | first full w (W pt) | implied c | formula c | candidate | baseline |
| --- | --- | --- | --- | --- | --- | --- |
| A_V13up_pct | 17.3438 (97.582) | 17.375 (98.402) | (4.462, 5.282] | 5.007 | 33/33 | 28/33 (5 false hash) |
| A_V13it_pct | 17.7812 (100.098) | 17.8125 (100.797) | (6.978, 7.677] | 7.527 | 41/41 | 37/41 (4 false hash) |
| A_V13up_pct75 | 17.5 (73.801) | 17.5312 (74.543) | (4.801, 5.543] | 3.974 | **31/41 (10 false full)** | 36/41 (5 false full) |
| A_V13up_acct | - (all 49 full) | - | - | 4.419 | 49/49 | 49/49 |
| A_G15up_grp | 16.1875 (82.102) | 16.2188 (82.797) | (4.462, 5.157] | 5.018 | **36/41 (5 false hash)** | 41/41 |
| A_G15it_grp | 16.7812 (84.957) | 16.8125 (85.703) | (7.437, 8.183] | 7.418 | 41/41 | 36/41 (5 false full) |
| A_G15it_pct | 17.3438 (97.586) | 17.375 (98.426) | (8.186, 9.026] | 8.238 | 41/41 | 36/41 (5 false hash) |
| H_Cambria16up_grp80 | 16.7812 (75.598) | 16.8125 (76.223) | (3.958, 4.583] | 4.419 | **37/41 (4 false hash)** | 37/41 (4 false hash) |
| H_Trebuchet12it_pct85 | 12.1875 (58.320) | 12.2188 (58.922) | (5.280, 5.882] | 5.470 | 41/41 | 32/41 (9 false full) |

Seven of the eight transitions contain the formula `c` once the documented fill-vs-layout width offset
(~0.1 pt) is allowed for. The exception is Verdana 13 at 75 %: the transition is consistent with the same
`c` only when the printed advance 70.098 is used instead of the 81-dot model width 69.0 -
`73.801 - 70.098 < 3.974 <= 74.543 - 70.098`. So that residual lives in the advance quantisation at
75 %, not in the fit rule or the room constant.

The five known rows are inside these transitions: Verdana 17.375/17.4375/17.5 now fall on the full side
(W 98.40-98.43, last hash 97.58); Georgia italic 16.6875/16.75 fall on the hash side (last hash 84.96,
first full 85.70). Both are reproduced by the candidate at the pre-registered `W`.

### Blind holdout

- **Oracle lock: 78/82 exact**, every miss a Cambria false hash (4 rows, widths 16.8125-16.9062) where
  the pre-registered `W` is 0.03 pt low (`W+0.15` makes it 82/82, and at the *measured* fill width it is
  82/82). Trebuchet: 41/41, including the discriminating slice 12.0938-12.1875 where Excel prints 9
  hashes.
- **Engine lock (`print-fullhash-engine-predictions.tsv`): 78/82**, same four Cambria rows; its
  `candidate_full_W+0.15` column scores 82/82.
- **Baseline (frozen runner PDF, same measurement): 69/82.** It prints the full number in the whole
  Trebuchet 12.0938-12.1875 slice (9 false full) and hashes the four Cambria rows (4 false hash).

### Totals (369 dense columns: 287 anchor + 82 holdout)

| decision rule | exact | false hash | false full | newly wrong vs baseline |
| --- | --- | --- | --- | --- |
| baseline (current print fit, measured on its own PDF) | 332/369 (90.0 %) | 18 | 19 | - |
| candidate at pre-registered W | 350/369 (94.9 %) | 9 | 10 | 10 |
| candidate at measured fill W | 354/369 (95.9 %) | 0 | 15 | 5 |

Newly wrong at the pre-registered W: the 5 Georgia 90 % upright false hashes (W formula 0.04-0.15 pt low;
0 at measured W) and the 5 extra Verdana 75 % false fulls. At that W the Verdana 75 % block (10 columns) is
the only failure not explained by the width formula: the Georgia upright and Cambria misses vanish at the
measured fill width, and the Georgia italic 90 % transition sits 0.019 pt above the formula `c` at the
fill width, inside the documented 0.1 pt fill-vs-layout allowance. At the measured W the 5 Georgia italic
90 % false fulls coincide with the baseline's own false fulls, so only the Verdana 75 % block is newly
wrong there.

Counts: 187/187 printed hash counts are exact wherever the decision agrees with Excel; the only count
mismatches (10) are the same Verdana 75 % columns that the candidate wrongly prints full. No independent
count error was observed (consistent with the count model's earlier 388/388 blind holdout).

## Limits

- Word-box metric, not ink: the box is an advance box with side bearings; residual 0.02-0.14 pt for
  tabular faces, larger for Georgia's proportional digits. It is used only to compare relative advances
  and to bound transitions, never as the model's own width.
- The 1/32-char grid locates a transition to ~0.4-0.7 pt at non-100 % scale because Excel quantises the
  physical column width; at 100 % steps are ~0.12-0.3 pt. `c` intervals are correspondingly wide.
- Fill width is 0.10-0.15 pt wider than the layout width (reproduced here: model `W` is 0.09-0.12 pt below
  the fill edge); a rule that must choose at a boundary needs the product's layout `W`, and the Cambria
  rows show a W-precision miss is a decision miss.
- The 82-dot number-layout height at 75 % is inferred from the printed advance and the transition; the
  mechanism (which Excel code path rounds up) is not established, and only one face/size is at a .25-dot
  height in this sweep.
- Accounting: the `*` fill spreads `$` to the left edge, so the whole printed span is not the number's
  advance; the fit is still bounded by the number word + pads. The 100 % window (17.5-19.0 chars) shows
  no transition, so its cutoff is not localised here.
- No Chromium/browser gate: this is a print-only measurement; screen overflow semantics are untouched.

## Reproduce

```bash
python3 tools/excel-parity/make-print-fullhash-anchor.py \
  tests/test-documents-excel/synthetic-print-fullhash-anchor.xlsx \
  tools/excel-parity/print-fullhash-anchor.meta.tsv anchor
python3 tools/excel-parity/make-print-fullhash-anchor.py \
  tests/test-documents-excel/synthetic-print-fullhash-holdout.xlsx \
  tools/excel-parity/print-fullhash-holdout.meta.tsv holdout
tools/excel-parity/excel-print-pages.sh tests/test-documents-excel/synthetic-print-fullhash-anchor.xlsx /var/tmp/fh/anchor
python3 tools/excel-parity/measure-print-fullhash-anchor.py \
  tools/excel-parity/print-fullhash-anchor.meta.tsv /var/tmp/fh/anchor /var/tmp/fh/anchor.tsv
# lock the holdout from its meta alone, then export it
python3 tools/excel-parity/predict-print-fullhash-anchor.py \
  tools/excel-parity/print-fullhash-holdout.meta.tsv /var/tmp/fh/holdout-pred.tsv
tools/excel-parity/excel-print-pages.sh tests/test-documents-excel/synthetic-print-fullhash-holdout.xlsx /var/tmp/fh/holdout
python3 tools/excel-parity/measure-print-fullhash-anchor.py \
  tools/excel-parity/print-fullhash-holdout.meta.tsv /var/tmp/fh/holdout /var/tmp/fh/holdout.tsv
python3 tools/excel-parity/analyze-print-fullhash-anchor.py \
  tools/excel-parity/print-fullhash-anchor.meta.tsv /var/tmp/fh/anchor.tsv \
  /var/tmp/fh/anchor-pred.tsv /var/tmp/fh/anchor-analysis.tsv
```

Baseline decisions in the totals come from our own PDF for the same workbook, rendered with a frozen
runner copy and measured with the same tool (`--map` supplies sheet → PDF); the checked-in analysis
files carry `ours_decision`/`ours_n` per column.
