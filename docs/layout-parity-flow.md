# Layout parity: the flow

How the Word layout is brought to what Microsoft Word does (and, the same way, the spreadsheet to Excel).
One loop, run again and again; never a fix for one document.

1. **Find discrepancies.** `node tools/word-parity/page-parity.mjs --dist dist --detail` compares every
   corpus document's pages with Word 2019's (cached in `tests/word/word-pages/`; `--refresh-word` asks the
   Windows VM again). Pick the document whose pages diverge first or most, and find the first line where our
   layout and Word's part (`--lines <doc>` dumps ours; Word's PDF is in `artifacts/word-pages/`).
2. **Reason out the rule.** What is there - the construct, the spacing, the break - and what is the simplest
   general rule Word could be following that gives its result? Prefer a rule that also explains what Word
   does elsewhere over one that only explains this document.
3. **State the hypothesis** in the ledger (below) as a rule, with the predictions that would refute it.
4. **Build fixtures.** A generator (under `tools/word-parity/fixtures/`, committed, with how to regenerate)
   makes a family of small documents that vary one thing at a time across the boundaries the rule predicts -
   sizes, positions, nesting, content, flags, page size. The corpus is not evidence of a rule; fixtures are.
5. **Compare against Word.** `page-parity.mjs --dir <fixtures>` runs them in Word and in ours. Refuted: say
   what Word did, revise the hypothesis, back to 4. Confirmed by every fixture: go on.
6. **Implement the rule**, removing whatever special cases it makes unnecessary. The fixtures and Word's
   results stay in the repo as regression data and tests run against them.
7. **Check the corpus.** Rerun all of it: a document that gets worse needs a Word-backed explanation or a
   new hypothesis (back to 2). Record the outcome in the ledger.

## Ledger

One line per hypothesis, refuted ones too. Status: testing, confirmed (implemented), refuted, open.

| Rule | Status | Evidence |
|---|---|---|
| A line is its font's single-line height times the spacing multiple, in body text, tables, headings and TOC entries alike | confirmed | Word: Calibri 11pt 13.43pt single, 15.44pt at 1.15; Times New Roman 12pt 13.80pt (bf94fc0e) |
| No `w:pPrDefault`: Word's own defaults (8pt after, 1.158 x single); an empty one: single spacing | confirmed | Calibri 11pt 15.54pt, 22pt 31.10pt, Arial 14.70pt, Times New Roman 14.64pt (bf94fc0e) |
| A paragraph mark's run properties (`w:pPr/w:rPr`) format the mark only, not the paragraph's runs | confirmed | a run under a Calibri 11pt mark is set in its style's Times New Roman 12pt (f9ee9884) |
| A table style inherits paragraph and run properties from the style it is based on | confirmed | ElementTable on Table Grid: single-spaced cells, 13.44pt lines (f9ee9884) |
| A line stays on the page when its text at single spacing fits; the extra of its spacing multiple and the paragraph's space after need not | confirmed | 1.15-1.62x lines kept up to 5pt past the margin; a single-spaced line 1.6pt over moved |
| The body starts below the header and ends above the footer as they are laid out | confirmed | docx4j-sample-docx body at 62.9pt in Word (bdcf1198) |
| A word in several runs breaks only where it would in one run (a space, after a hyphen) | confirmed | Word's own wrapping of per-letter colour runs (5eebf850) |
| A row breaks only where every cell can break at the same position | refuted | fixtures `row-break` (221, `tools/word-parity/fixtures/row-break.py`): `cell-two` (40pt lines beside 20pt), `nest2-text`, `widow-two_f23` - each cell breaks at its own last line that fits |
| A vertical-merge group has no break point inside it | refuted | `top-merge*`, `top-merge-tall*`, `nest-merge{1-2,3-6,7-9,1-10}`, `small-nest-merge3-6`, `tall-nestmerge*`, `depth2-merge`: merge groups break apart, top level and nested; the calibre variants with and without the merge move alike |
| Each cell breaks on its own at its last break that fits; a nested table between rows or inside a row by the same rule | confirmed | `cell-text`, `cell-two`, `nest-plain`, `nest2-text`, `depth2-merge`, `small-nest-plain` |
| A row moves whole only when some cell would keep nothing on the page (a nested row likewise in its cell); one at a page's top (or under repeated headers) stays and breaks | confirmed | `first-tall`, `first-space`, `nest-first-cantsplit`, `nest-first-orphan_f31`, `widow-two_f31`, `keeplines-two_f28-30`, `nest-orphan_f30`, `header-first-tall` |
| A nested cantSplit row that does not fit moves to the next page within the split outer row | confirmed | `nest-cantsplit_f27-30` |
| A cantSplit row moves whole; one taller than a page starts the next page and breaks at its bottom | confirmed | `top-cantsplit`, `tall-cantsplit_f10` |
| Widow control: an orphan goes on with the rest; a widow pulls one line along (n-1 fit: n-2 kept, or all go if that leaves one); body and cells alike | confirmed | `widow-body`, `widow-on`, `widow-two`, `widow-off`, `nest-first-orphan_f29` |
| No `w:widowControl` in paragraph, styles or defaults: on (against the spec's default) | confirmed | `widow-unset*` |
| A line of exact height fits only whole, not by its text | confirmed | `exact-body` (40pt lines: 100pt keeps 2, 60pt 1, 20pt 0), `cell-two`, `first-tall` |
| The calibre family's moved row is the embedded Ubuntu font, not a row rule | open | Word moves it in all three variants (original, no merge, no "Three"); with `--embedded-fonts` our page 4 is Word's |
| Word drops the first paragraph's automatic space before | open | seen in Word, not yet probed |
| A list marker's font does not make its line taller (only the text does) | refuted | fixtures `list-marker-height`: Word's pitch is 15.44pt plain, with a Calibri bullet and numbered, but 16.03pt with Word's Symbol bullet and 21.05pt with a 16pt Symbol bullet - the marker counts. Ours: 16.07pt, already right |
| The footnote block ends at the bottom margin: its last line's bottom is the page height less the bottom margin, and the body may run down to it less the separator | testing | docx4j-getting-started (45 -> 46 pages): every paragraph of page 4 within 0.5pt of Word's, but our footnotes start 4.16-4.37pt above Word's on pages 4, 10 and 12 and end ~4.5pt above 720pt - so Word fits one more body line (`para-positions.vbs`, the footnotes' `Information(6)`) |
| A line with a footnote reference (superscript) is taller than the footnote text's single line | open | Word's two one-line footnotes on page 12 of docx4j-getting-started are 12.75pt apart; ours 12.21pt |
| Word fits ~1pt less per 468pt line than we do (test-long-document): kerning only with `w:kern`? | open | |
| A paragraph border takes its space and twice its width above (below) the text | refuted | fixtures `para-border` (`tools/word-parity/fixtures/para-border.py`, `tests/word/border-space.fsx`): `top-*`, `bottom-*` - sz 4/12/24/48 with space 0/1/4/10 take space + width (sz 24 space 4: 7pt, not 10) |
| A border takes its space and its line: a single, dotted, dashed line its width, double three widths, triple five, thick-thin small gap width + 1.5pt, medium gap two widths, large gap width + 2.25pt, thin-thick-thin small width + 3pt | confirmed | `top-*`, `bottom-*`, `line-*` (sz 6/12/24) |
| Wave and 3D border lines: what they take | open | `line-wave-*` 2.92/3.76/2.92pt, `line-doubleWave-*` 5.2pt at every size, `line-threeDEmboss-*` 2.2/2.92/5.92pt - no rule fits three sizes; kept as Word's numbers, not checked |
| Adjoining paragraphs with the same borders form one box only when they have the same style | refuted | `grp-style-two`, `grp-style-based`, `grp-style-direct`: one box across styles (and direct vs style borders) - docx4j-headers-footers' XML schema boxes (Schema Fragment, then Schema Fragment Last) |
| Adjoining paragraphs are one bordered group when their four borders (style, width, space, colour) are the same and so are their right indent and the left edge of their text (the left indent, or the first line's where further left); styles, shading, spacing and the between border do not matter. The top border goes above the first only, the bottom below the last only; space before / after stay inside the group | confirmed | `box3`, `topbottom3`, `grp-*` (`grp-sides`, `-width`, `-space`, `-color`, `-val`, `-left-only`, `-indent-left`, `-indent-right`, `-edge-hang-off`, `-edge-first-off` split; `-indent-hang`, `-indent-first`, `-edge-hang`, `-edge-both`, `-shade`, `-after`, `-before`, `-ctx` group) |
| Inside a group the between border takes 2 x space + width | refuted in part | `between3` (2.5pt at sz 4 space 1), `between3-sz24-sp4` (11pt) fit it, but `between-first-*` / `between-second-*`: the paragraph with the between border takes space + width below it, the one after takes its own between's space above it - confirmed in that form |
| Space after and the next paragraph's space before: the larger of the two | confirmed | `space-sum` (10 then 5: 10pt; 0 then 15: 15pt) - already ours |
| Table rows take only their cell margins; borders are drawn over them | refuted | fixtures `table-border` (`tools/word-parity/fixtures/table-border.py`): `all-sz*`, `insideH-sz24`, `topbottom-sz24` - every horizontal border takes its width; ours took none (or the widest of a row's borders, only where a cell's paragraphs inherited their space after) |
| A row's content is below the widest horizontal border on its top edge (the table's top or inside border, a cell's own top, the bottom of a cell above, any column); the last row's bottom border is below it; vertical borders take nothing; cell margins add; an exact row height holds the border above, an at-least height comes before both borders; compatibility mode 14 alike | confirmed | `all-sz4..48`, `insideH-*`, `topbottom-*`, `sides-*`, `double-sz6`, `cell-*`, `cell2col-*`, `mar*`, `style-grid43*`, `exact-sz24`, `atleast-sz24`, `atleast-low-sz24`, all also `-c14`; docx4j-headers-footers' element tables now at Word's row pitch (71.94pt vs Word's 71.88, a row at 730.78pt vs 730.85) |
| A row broken at the page's foot keeps as many of its cell's lines as fit | open | docx4j-headers-footers p14: the last row's narrow cell ("§" + a REF field's "Error! Reference source not found."): Word keeps "§Error! R", "eference" on the page (it breaks the word "Reference" inside, too long for the 46pt cell), ours keeps "§" alone and goes on with "Error!" - two lines from row 730.8pt fit above the 769.9pt margin |

## Queue

The corpus's discrepancies as of 2026-11-19 (master df3230ad, canvas measurer: 54 of 67 documents at Word's
page count, 51 with every page starting at Word's line), worked a round at a time, worst first:

| Documents | Word / ours | First divergence | Round |
|---|---|---|---|
| calibre family (calibre-demo x2, calibre-demo-docx, complex-tables, wikipedia-sample, microsoft-word-sample, sample-formatted) | 8 / 9 | p4: a row with a nested, vertically merged table | row rule done (canvas: all 8 pages start right, one extra at the end); next: embedded fonts (Ubuntu) |
| docx4j-getting-started | 45 / 46 | p4: the footnote block | footnote block (running) |
| docx4j-headers-footers | 31 / 34 (was 36), 17 starts (15 before this round, the repeated header row left out) | p15: a row broken at the page's foot (see the ledger) | borders round done: grouped paragraph borders, table border widths; next: the broken row's cell |
| test-long-document | 8 / 8, 7 starts (was 7 / 3) | widow control, not kerning | mostly done by the row round |
| docx4j-legal-contributor | 3 / 4 canvas, 3 / 3 ShaperText | p3 | |
| file-sample | 2 / 3 | | |
| docx-5mb | 4 / 4, all starts | - | done by the row round |
| docx4j-conventions | 9 / 9, 5 starts | p6 | |
| docx4j-repeating-section | 8 / 8, 7 starts | p8 | |
| docx4j-binding-simple | 1 / 1, 0 starts | p1's text: content-control data binding, not layout | |
