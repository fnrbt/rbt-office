# Browser Fable proof: printed numeric overflow counts, blind workbook

No product changes. The proof mounts the actual Fable `Spreadsheet` component from two dists, imports
the blind holdout workbook, calls `grid.SetPdfFonts(...)` + `grid.ExportPdf({what:'sheet', sheetIndex})`,
and measures the resulting PDFs with the same fill-rectangle/word-box oracle as the Excel side.

## Freeze and dists

CountModel confirmed the final frozen sources (sha256, rechecked around the build):

```
src/Spreadsheet.Pdf/Pdf.fs   8738b08b295253cee6965b2e897d3f84d9b018564f6e4457fe2ddfe8f3b1242d
src/Spreadsheet/Paint.fs     2145bf894d756067c8670d000497dd26bf308fe8278225ed09883550537825ea
src/Spreadsheet/CellText.fs  0309d36fe01a7a66fc8a44e19735f171d3275f5640dafb9ec1ccb97e148b466c
```

| dist | provenance |
| --- | --- |
| baseline | copy of the published `dist/` dated 2026-11-23 19:57 = release `c9fe8034`; not rebuilt after |
| final | `bash scripts/build-dist.sh /home/dave/count-proof/final3-dist` from the frozen sources at 22:35; `Pdf.fs` hash identical before and after the build |

The first final build raced a late (`LineCount = 1` / `max 1`) hot-path edit; it was discarded. The
reported `final3` build is the only one used, and `Pdf.fs` was 8738b08 before and after it.

## Blind workbook and browser proof

`tests/test-documents-excel/synthetic-print-count-holdout.xlsx` (Verdana 13 upright 100 % `0.00%`;
Georgia 15 italic 90 % `#,##0.00`; widths 8.00-15.00 and 16.00-21.00 in 1/16 steps; predictions locked
earlier in `print-count-holdout-predictions-countmodel.tsv`).

```bash
node tools/xlsx-browser/browser-count-proof.mjs \
  --dist /home/dave/count-proof/final3-dist \
  --fixture tests/test-documents-excel/synthetic-print-count-holdout.xlsx \
  --out /home/dave/count-proof/final3-pdf --fonts /home/dave/count-proof/fonts --label final3
python3 tools/excel-parity/measure-print-count-transitions.py \
  tools/excel-parity/print-count-holdout.meta.tsv /home/dave/count-proof/final3-pdf \
  tools/excel-parity/print-count-browser-final-results.tsv
```

The same commands with `--dist /home/dave/count-proof/baseline-dist` produce
`print-count-browser-baseline-results.tsv`. The font directory has the 16 PDF faces used by the
workbook, including a real Georgia Italic (`georgiai.ttf`); the browser resolved all 16 and reported
no page errors. Both exports made two one-page PDFs per workbook.

## Results (388 row-1 columns; Excel direct PDF is truth)

| build | exact | overflow | full-number |
| --- | --- | --- | --- |
| **final Fable (8738b08)** | **383/388 (98.7 %)** | **259/261** | **124/127** |
| frozen c9fe8034 Fable | 357/388 (92.0 %) | 233/261 | 124/127 |

- 26 columns change count between baseline and final; **all 26 now equal Excel**. No count that was
  right in the baseline became wrong.
- Final browser counts equal the locked native predictions on all 383 cells where the overflow
  decision agrees with Excel.
- Remaining 5 rows are the **overflow decision**, which this change deliberately does not touch, and
  are identical in final and baseline:
  - `H_verdana13_up_pct` columns 137-139 (widths 17.375/17.4375/17.5): Excel prints the full
    `1234567.89%`, both browsers print 8 hashes.
  - `H_georgia15_it_grp` columns 126/127 (widths 16.6875/16.75): Excel prints 8 hashes, both browsers
    print `1,234,567.00`.

## Guards

- **No newly wrong fully displayed numbers.** Excel-full cells 127; final wrong 3; baseline wrong 3;
  newly wrong 0; newly fixed 0. The 3 are the pre-existing Verdana decision rows above.
- **No decision flip.** 0 columns move between hash and full between baseline and final.
- **Literal hashes unchanged.** Four literal-hash sheets of `synthetic-print-hash-geometry.xlsx`
  exported through both browsers: 78/78 rows identical text, and 78/78 identical `xMin`/`xMax`
  (`print-count-browser-literal-{final,baseline}-results.tsv`). That partial set needs the explicit
  opt-in; the holdout measurement stays strict:
  `python3 tools/excel-parity/measure-print-count-transitions.py tools/excel-parity/print-hash-geometry.meta.tsv /home/dave/count-proof/final3-lit /home/dave/count-proof/final3-lit-results.tsv --allow-missing-sheets`
- **Centering unchanged.** For the 362 unchanged-count columns the word box is byte-identical between
  baseline and final (max xMin/xMax delta 0.000 pt); for the 26 recounted columns the run centre moves
  at most 0.003 pt. Hash-run centre residual vs the cell fill centre is identical in both builds
  (mean -0.576 pt, max 1.205 pt: a word-box convention, not placement).

## Paths

| Path | What |
| --- | --- |
| `tools/xlsx-browser/browser-count-proof.mjs` | browser proof harness (serves dist/fixture/fonts, writes PDFs + sheet map) |
| `tools/excel-parity/print-count-browser-final-results.tsv` | final browser measurements |
| `tools/excel-parity/print-count-browser-baseline-results.tsv` | c9fe8034 browser measurements |
| `tools/excel-parity/print-count-browser-literal-{final,baseline}-results.tsv` | literal-hash guard rows |
| `tools/excel-parity/print-count-browser-comparison.tsv` | summary table |
| `tools/excel-parity/print-count-oracle.md` | Excel-side oracle and native model evidence |

Raw PDFs and the two dists are in `/home/dave/count-proof/` (final3-dist, baseline-dist, final3-pdf,
baseline-pdf, final3-lit, baseline-lit); the shared checkout's `dist/` was not modified.
