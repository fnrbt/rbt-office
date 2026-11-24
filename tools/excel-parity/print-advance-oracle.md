# Printed advance oracle: Verdana 13 at 75 % and 100 % (2026-11-24)

No product sources were edited.  Generator, measurement, lock and comparison:
`make-print-advance-oracle.py`, `measure-print-advance-oracle.py`,
`analyze-print-advance-oracle.py`, `lock-print-advance-oracle.py`,
`compare-print-advance-holdout.py`.  Fixtures:
`tests/test-documents-excel/synthetic-print-advance-{train,holdout}.xlsx`.
Data: `print-advance-oracle-*.{meta,results,analysis,lock,comparison}.tsv`.

Every sheet was exported by `excel-print-pages.sh` (fresh direct Excel 2019 PDF,
before any preview or sheet activation).  Each cell has its own solid fill, so
the physical edges come from the PDF fill rectangles.

## 1. The 70.098 pt is a real pen advance, not a word-box artifact

For all **668/668** row-1 and control cells (`advance agreement` in
`print-advance-oracle-train.analysis.tsv`):

```
max |PDF positioning advance - pdftotext bbox width| = 0.000000 pt
```

The measurement reconstructs the pen advance from `/Widths`, `Tc` and the `TJ`
adjustments (tracking the PDF `q`/`Q` text-state stack) and it equals Poppler's
word box exactly.  For `1234567.89%` at Verdana 13 / 75 % the operators are
`9.72 Tf`, `0.0581 Tc`, `[(1234567.)12(89%)] TJ`:

| quantity | pt |
| --- | --- |
| pen advance excluding the trailing `Tc` (= bbox) | **70.09844** |
| pen advance including the trailing `Tc` | 70.15654 |
| previous engine model (integer dots at 81) | 69.00000 |
| exact-height candidate (integer dots at 81.25) | 70.20000 |

The gap is therefore in the glyph advances, not in the extraction convention.

## 2. Slope, not intercept, at 75 %

Regression of the measured advance of `1...1.00%` on the count of repeated `1`s
(`print-advance-oracle-train.analysis.tsv`):

| sheet | measured slope / intercept | engine slope / intercept | exact-height slope / intercept |
| --- | --- | --- | --- |
| Verdana 13 @ 75 % | **6.24471 / 26.36333** | 6.12000 / 26.16000 | 6.24000 / 26.52000 |
| Verdana 13 @ 100 % | **8.28314 / 35.14244** | 8.28000 / 35.28000 | 8.28000 / 35.16000 |

The same split by the PDF operator strategy (the exporter switches from per-glyph
`TJ` corrections to a uniform `Tc` as the run grows):

| sheet | `Tc` | implied per-digit advance |
| --- | --- | --- |
| Verdana 75 % | 0 | 6.23052 |
| Verdana 75 % | 0.0581 | 6.24002 |
| Verdana 100 % | 0 | 8.26686 |
| Verdana 100 % | 0.0374 | 8.27996 |

At 75 % the engine's 6.12 pt/digit is **0.110-0.120 pt low per digit**; 9 digits
account for 0.99-1.08 pt of the 1.098 pt gap.  At 100 % the engine's 8.28 already
matches.  The exact-height candidate predicts the slope but runs **0.10-0.18 pt
wide per string** (a small residual intercept); it is not an exact model of GDI
hinting.

## 3. Tight cells independently verify the cutoffs

Each sweep is 33 columns at 1/32 character; `W` is the PDF fill width, the
advance is the independently measured wide-cell run, so the implied `c` does not
use the predictor.  Every sweep brackets its transition
(`c_bracket(measuredAdv)` contains the model `c`):

| sheet | text | last hash W | first full W | implied c | model c | engine wrong | exact wrong |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Verdana 75 % | `1234567.89%` | 73.8008 | 74.5430 | (3.7024, 4.4446] | 3.9742 | 10 | 0 |
| Verdana 75 % | `1111111111.00%` | 92.5430 | 93.2422 | (3.7245, 4.4237] | 3.9742 | 9 | 0 |
| Verdana 100 % | `1234567.89%` | 97.5859 | 98.4023 | (4.4442, 5.2606] | 5.0073 | 1 | 1 |
| Verdana 100 % | `1111111111.00%` | 122.6602 | 123.3789 | (4.6786, 5.3973] | 5.0073 | 0 | 0 |
| Georgia 90 % | `1,234,567.00` | 82.0781 | 82.8242 | (4.4200, 5.1661] | 5.0179 | 4 | 0 |
| Georgia 90 % | `1,111,111.11` | 68.2773 | 69.0234 | (4.6945, 5.4406] | 5.0179 | 0 | 5 |

Totals over the 198 tight cells: **engine 24 wrong, exact-height 6 wrong**.  The
exact-height candidate removes all 19 Verdana-75 misses and does not change
Verdana 100 %.  It regresses five Georgia 90 % `1`-heavy columns: at H = 112.5
dots the eight `1` glyphs round 48.34 -> 48 in the formula while Excel's hinting
uses 49; the rounded-height engine happens to match there.  A global switch to
exact-height is therefore not safe for unseen faces without hinted advances.

## 4. Georgia: the column width is a separate residual

Measured physical `W` minus the engine's `W(width_chars)` model:

| sheet | kind | n | mean | min | max |
| --- | --- | --- | --- | --- | --- |
| Georgia 90 % | wide | 24 | **+0.0882** | +0.0712 | +0.0947 |
| Verdana 75 % | wide | 56 | +0.1339 | +0.1186 | +0.1420 |
| Verdana 100 % | wide | 56 | +0.1405 | +0.1386 | +0.1425 |

The model `W` is low for every face by ~0.09-0.14 pt (the Georgia 90 % `16.2188-
16.3438` columns are full in Excel while the engine predicts hash).  That is a
width-model residual, independent of the Verdana advance slope; computing the same
cutoff at the measured `W` makes the engine correct on both Georgia strings.

## 5. Holdout

Reserved before export: unseen scale **Verdana 13 upright 140 %** and unseen face
**Tahoma 15 upright 90 %**, both `0.00%`.  The lock
(`print-advance-oracle-holdout-lock.tsv`) records the workbook and metadata
SHA-256 and the shipping engine's and exact-height candidate's decisions for all
244 row-1 cells; no holdout PDF existed when it was written.  Engine lock: 150
full / 94 hash.  Exact-height lock: 180 full / 64 hash.

The engine agent's lock (`print-advance-oracle-holdout-engine-lock.tsv`, mtime
2026-11-24 05:28:04Z, sha256 `fab70837...483a3e`) verified the same workbook and
metadata hashes and was present before the export; the PDFs were produced at
05:28.  Its candidate (FreeType-v35 hinted advances at ppem 152 and 113) is
numerically identical to its baseline on every cell, so this holdout tests the
integer-height model against the exact-height model, **not** hinting.

Fresh direct Excel 2019 result (488 cells; PDF advance vs bbox still 0.000000 pt;
engine lock's advances equal this experiment's rounded-height columns exactly):

| sheet | text | cells | Excel full | engine full | exact-height full | engine wrong | exact wrong |
| --- | --- | --- | --- | --- | --- | --- | --- |
| H_V13up_pct140 | `1234567.89%` | 33 | 12 | 12 | 17 | **0** | 5 (false full 74-78) |
| H_V13up_pct140 | `1111111111.00%` | 33 | 7 | 7 | 17 | **0** | 10 (false full 107-116) |
| H_Tahoma15up_pct90 | `1234567.89%` | 33 | 12 | 12 | 17 | **0** | 5 (false full 74-78) |
| H_Tahoma15up_pct90 | `1111111111.00%` | 33 | 7 | 7 | 17 | **0** | 10 (false full 107-116) |

Per-digit advance on the holdout (measured, engine, exact-height):

| sheet | measured slope | engine | exact-height |
| --- | --- | --- | --- |
| Verdana 13 @ 140 % | 11.64667 | 11.64000 | 11.52000 |
| Tahoma 15 @ 90 % | 7.44583 | 7.44000 | 7.32000 |

Wide-cell engine residuals are -0.045 .. +0.165 pt; the exact-height candidate is
0.35-1.75 pt narrow and grows with length.  The reason is the same boundary
effect in the opposite direction: at 140 % the exact height is 151.667 dots and
`floor(0.6357 x 151.667 + 0.5) = 96`, while Excel/engine use 97 at the rounded
ppem; at Tahoma 90 % the exact height 112.5 gives 61 for `1` while Excel uses 62.

**Totals over 330 tight cells: engine 24 wrong, exact-height 36 wrong.**  The
engine's 24 are the Verdana-75 slice (19) plus five Georgia/100 % width-model
misses; exact-height's 36 are its six train misses (five Georgia false-full, one
Verdana-100 false-hash) plus 30 holdout false-fulls.  The exact-height arithmetic candidate is therefore not a
general replacement: it fits the 75 % slice by crossing a glyph-rounding boundary
that happens to coincide there.  The honest next discriminator is the engine's
hinted-advance path at ppem = 81 on the 75 % training columns; this holdout's
ppem 152/113 have no hinting delta and cannot decide it.


## 6. Candidate vs baseline scoreboard

`engine` = afe6665d shipping rounded-height model; `exact` = per-glyph integer
rounding at the unrounded height (`floor(em * size*scale*600/72 + 0.5)`).

| set | tight cells | engine wrong | exact-height wrong | exact-height miss class |
| --- | --- | --- | --- | --- |
| train (Verdana 75/100, Georgia 90) | 198 | 24 | 6 | Georgia 90 false-full 5, Verdana 100 false-hash 1 |
| blind holdout (Verdana 140, Tahoma 90) | 132 | **0** | 30 | false-full 30 |
| **all** | **330** | **24** | **36** | |

Engine misses: Verdana 75 false-full on the 19 columns the reported residual
predicts (`1234567.89%` cols 65-74, `1111111111.00%` cols 98-106), Georgia 90
false-hash 4 (`1,234,567.00` cols 43-46, the W model 0.07-0.09 pt low), and
Verdana 100 false-hash 1 (col 74, model W quantisation).  Exact-height trades
those for 30 holdout false-fulls and does not change the V100/Georgia width
classes.

Recommendation: keep the integer-height baseline.  The 75 % Verdana slice needs
hinted per-glyph advances at ppem = 81 (the engine's `cand_*` path), not the
exact-height arithmetic; the latter fits 75 % only because its digit boundary
happens to agree there and is demonstrably wrong at ppem 152 and 113.  This
experiment supplies the locked blind evidence for that decision but does not
edit production.
