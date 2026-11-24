# Print fit cutoff: final Chromium gate

No product changes, no commit. Frozen sources (hashes identical before and after the single build):

```
CellText.fs  ee6d6ea8609bf338cee3f406b86510665572fdc9c970e542313e80399dcfb430
Paint.fs     e9bec9d3fb9901240ca9495077c33c3d0b684bf8228e9cd25f5c32376527548b
Pdf.fs       8f04d59a1597440a3d3b06bf3b5766a14e119ae25b936138c2fff99778880ab3
```

| role | path |
| --- | --- |
| final dist (built once, 01:26:09) | `/var/tmp/fitgate/final-dist` |
| frozen baseline dist (e3d55f93) | `/home/dave/count-proof/final3-dist` |
| final native runner (same freeze) | `/var/tmp/fitgate/final-runner` |
| pre-change native runner (e3d55f93) | `/var/tmp/fh-oracle/runner` |
| browser gate tool | `/var/tmp/fitgate/gate.py`; harness `tools/xlsx-browser/browser-count-proof.mjs` (font list extended with Cambria, Trebuchet MS, Aptos Narrow; 27 faces resolved, no page errors) |
| fonts | `/var/tmp/fitgate/fonts` (includes Cambria face 0 extracted from `cambria.ttc`) |

Every export is a real Chromium PDF from the Fable `Spreadsheet` component (`grid.ExportPdf({what:'sheet'})`),
measured with the same fill-rectangle/word-box oracle as the Excel truth. No truth row was dropped:
row2 literal controls with no second word and accounting `$`-span text differences are reported below.

## Result per set (row 1 only; Excel direct PDF is truth)

| set | row1 cols | final exact | final wrong | baseline (e3d55f93) wrong | gains | regressions | final counts |
| --- | --- | --- | --- | --- | --- | --- | --- |
| new anchor | 344 | 329 (95.6 %) | 15 (5 false hash, 10 false full) | 24 | 19 | **10** | 154/154 |
| blind holdout | 96 | 92 (95.8 %) | 4 (4 false hash) | 13 | 9 | 0 | 33/33 |
| previous 388 holdout | 388 | **388 (100 %)** | 0 | 5 | 5 | 0 | **261/261** |
| hash-geometry (old overflow/literal) | 289 | 289 | 0 | 0 | 0 | 0 | 277/279 |
| count-train accounting | 167 | 165 | 2 (pre-existing) | 2 | 0 | 0 | 134/137 |

Counts are exact/checked among row-1 columns whose decision matches Excel; the 2 hash-geometry and 3
accounting count misses are pre-existing (identical in the baseline), not caused by this change.

Final Chromium row-1 decisions and counts are identical to the final native runner on the new anchor and
blind holdout (0 differences in 440 row-1 cells).

## Regressions (complete, 10 columns)

All are in the new anchor workbook; no other set regressed.

1. `A_V13up_pct75` columns 31-35, widths **17.2188-17.3438** (also 36-40 at 17.375-17.5 are wrong in both
   builds): Excel prints 8 hashes, baseline printed 8, final prints `1234567.89%`. This is the known
   75 %-scale residual: the printed advance there is 70.098 pt while the 81-dot model gives 69.000
   (hash pitch is still 81-dot, so only a hinted-advance model would fix it). 10 false fulls total,
   29.1 % of the 344 anchor columns.
2. `A_G15up_grp` columns 24-28, widths **16.2188-16.3438**: Excel and baseline print `1,234,567.00`,
   final prints 8 hashes. The runtime layout `W` is 0.038-0.15 pt under Excel's fill edge; evaluating
   the same rule at the measured fill width gives 0 wrong on this sheet.

## Gains (33 columns)

- anchor 19: Verdana 13 upright 100 % at 17.375-17.5 (the 3 known rows + 2), Verdana 13 italic 100 %
  at 17.8125-17.9062, Georgia 15 italic 90 % `#,##0.00` at 16.6562-16.7812 (the 2 known rows + 3),
  Georgia 15 italic 100 % at 17.375-17.5.
- blind holdout 9: Trebuchet MS 12 italic 85 % at 11.9375-12.1875 (baseline printed the number,
  Excel/final print 9 hashes; includes the pre-registered 12.0938-12.1875 slice).
- previous 388 holdout 5: Verdana 13 at 17.375/17.4375/17.5 (now full) and Georgia 15 italic 90 % at
  16.6875/16.75 (now 8 hashes). 388/388 decisions and 261/261 counts vs Excel.

## Reported non-decisions (no silent exclusion)

- hash-geometry: 5 missing rows are **row 2** literal controls whose second word does not exist
  (`missing-word`), in both builds; 2 count misses both builds: `R_calibri11` col 12 w 10.5
  (Excel 10, browser 9) and `R_arial12` col 3 w 10.25 (Excel 8, browser 7) - the known fill-vs-layout
  edge cases.
- accounting: final and baseline are byte-for-byte equal in decisions/counts. Pre-existing mismatches:
  `T_calibri11_up_acct` cols 47-48 w 10.8125/10.875 (Excel full, both print 9 hashes) and cols 40-42
  w 10.375-10.5 (Excel 9 hashes, both print 8). Active `*`-fill accounting is deliberately outside the
  new cutoff, so this is the prior path preserved.

Evidence: `tools/excel-parity/print-fit-browser-gate.tsv` lists every non-agreeing row with
set/sheet/column/width/Excel/final/baseline and class (`gain`, `regression`, `both-wrong`,
`count-mismatch`).
