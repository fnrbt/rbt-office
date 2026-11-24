# Border weight and the trailing-fill boundary: discriminating direct-Excel sweep

Question: is the trailing `*x` fill repeat boundary the border's inner edge (width-proportional), a
fixed device quantum, or the no-fill text content width? Three faces and four border weights on
identical transition-adjacent rows.

## Files

| role | path |
| --- | --- |
| generator | `tools/excel-parity/make-print-border-fill-fixture.py` |
| workbooks | `tests/test-documents-excel/synthetic-print-border-fill-{arial,aptos,calibri}.xlsx` |
| manifests | `tools/excel-parity/print-border-fill-{arial,aptos,calibri}-cells.tsv` (same content) |
| locked predictions | `tools/excel-parity/print-border-fill-predictions.tsv` sha256 `ffb1bd16…` |
| measurement | `tools/excel-parity/measure-print-border-fill.py` |
| analysis | `tools/excel-parity/analyze-print-border-fill.py` |

12 sheets per workbook `B{none,thin,medium,thick}_{left,center,right}`, identical column widths in
every sheet, widths 8.5..11.25 in 0.25-char steps, formats N1 `0.00`, S2 `0.00* `, D2 `0.00*.`,
F8 `0.00*-;0.00`, values +1234.5/-1234.5/0. Faces: Arial 11/60 %, Aptos Narrow 14/100 %, Calibri
11/100 %. Cell geometry is taken from the tall solid cell-fill rectangles (drawn from the cell edge
in every style), so W/L/R are identical across weights; borders are measured separately as ink.

Printed border ink from the PDF operators:

| face / scale | none | thin (strip / stroke) | medium | thick |
| --- | --- | --- | --- | --- |
| Arial 11 / 60 % | — | 0.60 / — | 1.20 | 1.68 |
| Aptos Narrow 14 / 100 % | — | 0.96 / 0.14 | 1.92 | 2.88 |
| Calibri 11 / 100 % | — | 0.96 / 0.14 | 1.92 | 2.88 |

## Result 1: border weight does nothing

For every cell the count, amount x0/x1, composite end and residual are identical under no, thin,
medium and thick borders: 0/432 differing cells on Arial and 0/396 on Aptos and Calibri. The count
transition content width moves 0.000 pt between weights (all transitions, all formats/values). The
residual minimum is style-invariant per face (Arial 1.176, Aptos 2.019, Calibri 1.654 pt).

## Result 2: the correction is a constant, not border geometry

Model `count = floor((W - L - R + c - w)/a)`, with L/R the same-sheet N1 controls, W the cell width,
and w/a 600-dpi font advances. The feasible c interval is *identical* for none/thin/medium/thick on
every face:

| face | basis for w | c interval |
| --- | --- | --- |
| Arial 11 / 60 % | predicted advance | [0.310, 0.417] |
| Arial 11 / 60 % | Excel amount width x1-x0 | [0.244, 0.360] |
| Aptos 14 / 100 % | predicted advance | [0.159, 0.279] |
| Aptos 14 / 100 % | Excel amount width | [0.240, 0.360] |
| Calibri 11 / 100 % | predicted advance | [0.313, 0.403] |
| Calibri 11 / 100 % | Excel amount width | [0.240, 0.330] |

With Excel's measured number width, one global constant `c in [0.25, 0.32]` gives **0 count failures
on all 1080 fill cells** (3 faces x 4 weights x 90 fill cells per face); `c = 0.24` is the sharp lower
edge and misses 20 cells only by boundary rounding. With the predicted number width no single constant
works (best 8/1080 at 0.32-0.36, 40/1080 at 0.24-0.26): the face-dependent predicted text width, not
the fill boundary, is the blocker.

Key sets are complete: 1728 measured cells per face (12 sheets x 12 rows x 12 widths). At 100 % the
last width column paginates horizontally; the measurement carries page-0 row labels to later pages by
their identical row y positions, so no cell is silently dropped.

Completeness is mandatory, not advisory: after extraction `measure-print-border-fill.py` compares the
written `(sheet, row, width)` counters with the manifest and exits 1 on any missing, extra or duplicate
key, printing the counts and up to ten examples. `verify-print-border-fill-keys.py` reuses the same
gate on a stored measured TSV. Exercised on 2026-11-24:

```
full extraction (each face) -> keys expected 1728 measured 1728 missing 0 extra 0 duplicate 0 PASS, rc=0
throwaway, 11.25 column rows dropped (paginated page) -> expected 1728 / measured 1584 / missing 144, FAIL rc=1
throwaway manifest expecting width 99 -> extraction FAIL missing 144, rc=1
throwaway duplicated row -> duplicate 1, FAIL rc=1
```

Verification command:

```
python3 tools/excel-parity/verify-print-border-fill-keys.py \
  tools/excel-parity/print-border-fill-<face>-cells.tsv /var/tmp/numfill-oracle/border-<face>-measured.tsv
```

## Verdicts

- **Half border width: falsified.** It predicts transition shifts of 0.30/0.60/0.84 pt (Arial) and
  0.48/0.96/1.44 pt (Aptos/Calibri); observed 0.000 pt, and the residual minimum does not move.
- **Border presence: falsified.** The `none` sheet has the same correction interval as the bordered
  sheets (~0.24-0.33 with measured width). The term is not produced by drawn border ink.
- **Fixed device quantum: supported only against Excel's own text extent.** One constant in
  [0.25, 0.32] reproduces every count; 2 device dots (0.24) sits at the boundary. Applied to the
  600-dpi predicted text width it fails on Arial and Calibri, so it cannot be used as an engine rule
  without fixing the text-width/inset model first.
- **No correction (0): falsified** on every face.

Evidence: `/var/tmp/numfill-oracle/border-{arial,aptos,calibri}-measured.tsv`,
`/var/tmp/numfill-oracle/border-cases.tsv` (per-cell observed/predicted/signed error for HALF, FIXED,
NONE on both width bases), `border-summary.tsv`, `border-report.txt`.

## Commands

```
python3 tools/excel-parity/make-print-border-fill-fixture.py --face Arial --size 11 --scale 60 \
  --out tests/test-documents-excel/synthetic-print-border-fill-arial.xlsx \
  --cells tools/excel-parity/print-border-fill-arial-cells.tsv
tools/excel-parity/excel-print-pages.sh tests/test-documents-excel/synthetic-print-border-fill-arial.xlsx \
  /var/tmp/numfill-oracle/border-arial
python3 tools/excel-parity/measure-print-border-fill.py tools/excel-parity/print-border-fill-arial-cells.tsv \
  /var/tmp/numfill-oracle/border-arial /var/tmp/numfill-oracle/border-arial/synthetic-print-border-fill-arial.tsv \
  /var/tmp/numfill-oracle/border-arial-measured.tsv
python3 tools/excel-parity/analyze-print-border-fill.py /var/tmp/numfill-oracle/border \
  /var/tmp/numfill-oracle/border-{arial,aptos,calibri}-measured.tsv
```

Evidence-only: this fixture does not touch or gate the runtime; the absolute 0.35 gate remains the
strict verdict.
