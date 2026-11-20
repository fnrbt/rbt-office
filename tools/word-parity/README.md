# Word parity tools

The scripts that drive Microsoft Word on the Windows VM and compare its layout
with ours. `docs/word-parity.md` describes the VM, the run modes and the older
geometry tools.

## Page parity (`page-parity.mjs`)

Our page count and the first and last line of every page, against Word's, for
every `tests/test-documents*/*.docx`:

```
pnpm run build:fable
RB_WORD_FONT_DIR=$PWD/artifacts/fonts/windows \
PLAYWRIGHT=tools/browser-runner/node_modules/playwright/index.mjs \
  node tools/word-parity/page-parity.mjs --dist dist [--detail] [--json out.json] [name ...]
```

- Word's side is committed: `tests/word/word-pages/<dir>__<name>.tsv`, one per
  document - Word's page count (`ComputeStatistics(wdStatisticPages)` after
  `Repaginate`) and for each page the character position it starts at, its
  first and last line, and its first and last line with a letter or number.
  `page-lines.vbs` writes them; `--refresh-word` stages the corpus on the VM
  (`WORD_VM`, default `winvm`; `WORD_VM_DIR`, default
  `C:/Users/olivia/RbtExcelParity`; direct `cscript` over ssh), runs it, and
  brings back the TSVs and Word's PDFs (`artifacts/word-pages/`). Rerun it when
  a document is added or changed.
- Our side is the browser layout of `dist` in Chromium, with Word's fonts
  registered as web fonts under the family their name tables give: the faces
  of `RB_WORD_FONT_DIR` and of `RB_WORD_EXTRA_FONT_DIR` (default `windows-extra`
  beside it: Consolas, Courier New, Tahoma, Georgia ...). Both come from the VM
  with `fetch-windows-fonts.sh`. Text is measured as a host that installs
  ShaperText's measurers measures it (`setTextMeasurers`: the fonts' own
  advances, from the same files); `--measurer canvas` leaves the canvas's
  `measureText`, which in headless Chromium gives whole pixels. The fonts a
  document embeds are used as the product uses them: the editor makes them
  canvas fonts itself (`loadEmbeddedFonts`, waited for), and the ShaperText font
  set gets them for the families it has no face of (`getEmbeddedFonts`,
  `ShaperText.addEmbeddedFaces`). The lines are `Word.getLayoutLines`' body lines.
- A table's header row repeated at the top of a later page is not where that
  page starts in the text (Word's side is the character position a page starts
  at): its lines are left out of the comparison.
- A page matches when its first text line is Word's (case, spaces and
  punctuation ignored, at most 40 characters compared). The table gives Word's
  pages, ours, how many of Word's pages start where ours do, and the first page
  that does not with both lines; `--dir <folder>` compares the documents of a
  folder instead (Word's side in `<folder>/word`, made with `--refresh-word`) -
  the small documents a rule is checked on; `--detail` lists every page whose first or last
  line differs; `--lines <name>` prints our lines of one document (page, y,
  height, block, text) for finding where a break moved.

Run it before and after any layout change: a change must not make a document
that matched stop matching.

## Fixtures (`fixtures/`)

Generators of small documents that each vary one thing, for checking a layout
rule in Word (`docs/layout-parity-flow.md`). `fixtures/row-break.py` writes
`tests/word/fixtures/row-break/` (where a table row breaks across pages); its
docstring has the regenerate / ask Word / extract-labels steps, and
`tests/word/row-break.fsx` checks our layout against Word's pages of them.
`fixtures/para-border.py` and `fixtures/table-border.py` write
`tests/word/fixtures/para-border/` and `table-border/` (what paragraph and table
borders take vertically); `para-border.py --positions <dir>` reads where Word put
every line from its PDFs into `<dir>/word/positions.tsv`, and
`tests/word/border-space.fsx` checks our gaps between lines against it.
`para-border.py --across` writes `tests/word/fixtures/para-border-across/` (where the
borders run across and where their lines are); `pdf-geometry.py --paths` reads the paths Word's
PDFs stroke and fill (its borders are filled rectangles) into its `word/geometry.tsv`, and
`tests/word/border-across.fsx` checks Paint's border commands against it. `fixtures/vmerge-content.py`
(`tests/word/fixtures/vmerge-content/`: what a vertical merge's continuation cells hold and the rows'
heights) is read with `pdf-geometry.py` for `tests/word/vmerge-content.fsx`.

`fixtures/float-wrap.py` writes `tests/word/fixtures/float-wrap/` (text round
floating pictures and shapes); where Word's lines and pictures are in its PDFs
comes from `pdf-geometry.py` (each line's pieces across and down, each picture's
box) into `word/geometry.tsv`, which `tests/word/float-wrap.fsx` checks against. `fixtures/float-table.py` writes
`tests/word/fixtures/float-table/` (where a floating table, `w:tblpPr`, is placed and the text round
it), read the same way into its `word/geometry.tsv` for `tests/word/float-table.fsx`.

`fixtures/long-word.py` (where lines break round a word longer than the line, in fields' results,
at a page's foot in a cell; `--lines <dir>` reads every line of Word's PDFs into `<dir>/word/lines.tsv`,
`tests/word/long-word.fsx`), `fixtures/kerning.py` (the width Word gives a word, kerned or not;
`--widths <dir>`, `tests/word/kerning.fsx`) and `fixtures/line-fit.py` (how many characters Word
puts on a line; read with `long-word.py --lines`) and `fixtures/hang-punct.py` (punctuation that
hangs past a line's end under East Asian layout, by setting and language; read with `long-word.py --lines`,
`tests/word/hang-punct.fsx`) and `fixtures/hyphen-break.py` (where a line breaks inside a word: after
which hyphens, dashes and punctuation; read with `long-word.py --lines`, `hyphen-break.py --table`,
`tests/word/hyphen-break.fsx`). A page that starts inside a hyperlink (a field) has its start and
first line in `page-lines.vbs` at the field's start - not where Word's page starts. A REF field is updated by Word's PDF export: its
line in the PDF is not Word's layout of it - use a QUOTE field in fixtures.

`fixtures/hyphen-words.py` (where Word hyphenates: thousands of probes, each a word Word must break or
move; `--points <dir>` reads Word's break of each from its PDFs, four columns a page, into `word/points.tsv`;
`--c14` the same in compatibility mode 14) and `fixtures/hyphen-rules.py` (when it does: the zone, the limit,
capitals, digits, languages, page and column feet; read with `long-word.py --lines`, `hyphen-rules.py --table`)
are checked by `tests/word/hyphenation.fsx`.

`fixtures/justify-squeeze.py` (how far Word narrows a justified line's spaces to fit more text: fonts, sizes,
space counts, compatibility modes, alignments, last lines, hyphenation, cells, columns; `--table <dir>` reads
Word's PDFs - or `--pdfs <dir>` another set - into `word/squeeze.tsv`, a row per probe) is checked by
`tests/word/justify-squeeze.fsx`.

`fixtures/table-width.py` writes `tests/word/fixtures/table-width/` (how wide Word makes a table's
columns: grid, cell and table widths, fixed and autofit, content, indents, columns, nesting, spans,
compatibility mode 14); `table-width.py --edges <pdfs>` reads where each cell's left-aligned label
starts and its right-aligned one ends from Word's PDFs into `word/edges.tsv`, which
`tests/word/table-width.fsx` checks ours against (the same `--edges` reads tools/pdf-runner's PDFs).
`fixtures/nested-width.py` writes `tests/word/fixtures/nested-width/` with the same labels and page
(how far a column whose cell holds a nested table gives way; how a nested table is laid out in a
fixed or an autofit table; aligned tables' room in compatibility mode 14), read with
`table-width.py --edges` and checked by the same `table-width.fsx`.
