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
| A line stays on the page when its text at single spacing fits; the extra of its spacing multiple and the paragraph's space after need not | confirmed | 1.15-1.62x lines kept up to 5pt past the margin; a single-spaced line 1.6pt over moved | (but see "however little" below)
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
| `w:rFonts`: each level of the chain replaces the inherited face per slot - an explicit face lower down beats a theme face above; on one element the theme face beats the named one | confirmed | fixtures `rfonts-precedence` (10, `tests/word/fixtures/rfonts-precedence`): Word's face for each; ours matched but for the one-element case, now fixed |
| A list marker's font does not make its line taller (only the text does) | refuted | fixtures `list-marker-height`: Word's pitch is 15.44pt plain, with a Calibri bullet and numbered, but 16.03pt with Word's Symbol bullet and 21.05pt with a 16pt Symbol bullet - the marker counts. Ours: 16.07pt for the 11pt bullet (but see the two rows below) |
| The footnote block ends at the bottom margin: the notes are stacked up from it (the last line's bottom, and its space after, at the page height less the bottom margin); the body may run down to the block's top less the separator | confirmed | fixtures `footnote-block` (35, Word's pages, PDF baselines and `para-positions.vbs` positions in `tests/word/fixtures/footnote-block/`, tests `tests/word/footnote-block.fsx`): the last note line's baseline 717.45pt (Calibri 10pt: its bottom at 720.1, PDF positions are 0.12pt steps) with 1, 2, 3, 10 lines, one or two notes; one-line note: Body 46 (text to 689.8pt) stays, a 2-line note or two notes: it goes; `limit-*`: the body's limit above a one-line note is in [693.9, 694.4), 720 - 12.21 - 13.43 = 694.36. Ours ended the block 4.5pt (6px pad) above 720 and took a 12pt separator. docx4j-getting-started: the notes of pages 4, 10, 12 now at 683.38, 695.59, 707.79pt - Word's 683.25, 695.25, 707.25-708 to its whole pixels |
| A line with a footnote reference (superscript) is taller than the footnote text's single line | refuted | `one-1`, `nosuper-1`, `custom-1` (mark not raised, `*`): the same 717.47pt baseline; `super-10` (a superscript on every line) = `one-10`; note lines 12.24/12.12pt apart in the PDFs. The 12.75pt was `Information(6)`'s whole pixels (17px) |
| The footnote separator is a paragraph of its own (footnotes.xml's `w:type="separator"`, `continuationSeparator` above a note going on from the page before), as high as its line; its rule 2in long, its top 0.703 x the font size below the line's top, 0.066 x the size thick | confirmed | `sep-double-1` (line 480) and `sep-big-1` (20pt mark): Body 46 goes, as it would with a 26.9 / 24.4pt separator; the rule at 702.10pt (11pt, 0.72 thick) over a one-line note, 688.66 double, 697.42 at 20pt (1.32 thick) |
| A note is its paragraphs with their own spacing, the first one's space before and the last one's after included, nothing between notes | confirmed | `fnspaced-2` (6 before, 12 after, 1.5): lines 18.24pt apart, the last baseline 699.35 (12pt after under it to 720), the rule 6pt higher; `dflt115-inherit-*` (FootnoteText inheriting 1.15 / 8pt after): notes 22.04pt apart, Body 44 0.4pt over the limit goes; `fnbefore-*` (12pt before): the rule 12pt higher, Body 46 goes; `two-1`, `two-para`: 12.24pt between notes |
| A note too long for the room under its reference is broken there like a paragraph with widow control: the reference's line stays when 2 of the note's lines fit (1 without widow control), the rest goes on at the next page's foot under the continuation separator, the body ending above it | open (Word confirmed, not ours yet) | `cont-k44..47` (10-line note, room for 3, 2, 1, 0 lines): 3+7, 2+8, the line moves, the line moves; `cont-k46-nowidow`: 1+9; `cont-40`: 13 lines on page 1, 27 on page 2 (Body 36-57 above them). Ours moves the reference's line whenever the whole note does not fit |
| A list marker taller than the text makes the paragraph's first line taller and moves its other lines down | confirmed | `list-marker-height` `bullet-symbol-3lines`: paragraphs 46.88pt apart = 16.03 + 2 x 15.44 (16pt bullet: 51.94 = 21.05 + 30.88). Ours grew the first line but not the paragraph (46.33); docx4j-getting-started page 6: the 0.75pt that kept a line Word moves |
| A list marker in a font size of its own sets the line with that size | open | `bullet-symbol16`: Word 21.05pt, ours 16.07 (the marker's line factor at the text's size) |
| A body line whose text ends below the bottom of the page's text goes on, however little | confirmed | `limit-nofn-p00`: Body 47's text 0.2pt below 720pt goes (ours kept up to 0.375pt over: 0.5px); `limit-p00` (0.03 over the footnote limit) goes too, but ours puts it 0.14 higher (next row) |
| Word's single line of Calibri 11pt is 13.4307pt, not 1.2207 x 11 = 13.4277 | open | baselines of 43 lines in the `footnote-block` PDFs: 577.52pt; 0.14pt over 45 lines (`limit-p00`, `dflt115-inherit-two`) |
| Word fits ~1pt less per 468pt line than we do (test-long-document): kerning only with `w:kern`? | open | |

## Queue

The corpus's discrepancies as of 2026-11-19 (master df3230ad, canvas measurer: 54 of 67 documents at Word's
page count, 51 with every page starting at Word's line), worked a round at a time, worst first:

| Documents | Word / ours | First divergence | Round |
|---|---|---|---|
| calibre family (calibre-demo x2, calibre-demo-docx, complex-tables, wikipedia-sample, microsoft-word-sample, sample-formatted) | 8 / 9 | p4: a row with a nested, vertically merged table | row rule done (canvas: all 8 pages start right, one extra at the end); next: embedded fonts (Ubuntu) |
| docx4j-getting-started | 45 / 46, 25 starts (was 23; canvas 18, was 16) | p10: a heading 14pt below Word's, then code paragraphs (the notes of pages 4, 10, 12 now Word's) | footnote block done; next: page 10 |
| docx4j-headers-footers | 31 / 34 (was 36) | p12, the attribute tables | next |
| test-long-document | 8 / 8, 7 starts (was 7 / 3) | widow control, not kerning | mostly done by the row round |
| docx4j-legal-contributor | 3 / 4 canvas, 3 / 3 ShaperText | p3 | |
| file-sample | 2 / 3 | | |
| docx-5mb | 4 / 4, all starts | - | done by the row round |
| docx4j-conventions | 9 / 9, 5 starts | p6 | |
| docx4j-repeating-section | 8 / 8, 7 starts | p8 | |
| docx4j-binding-simple | 1 / 1, 0 starts | p1's text: content-control data binding, not layout | |

### The showcase library

`tests/test-documents-showcase/` (made by `tools/showcase-docs/make.py`; Word's pages in its `word/`): designed
documents meant to look right, so every difference from Word is in plain view. The goal: each one matches Word
page for page and on sight (compare `tools/pdf-runner`'s PDF with Word's, page by page). First comparison
(2026-11-19): 7 of 8 at Word's page count, 5 with every page starting at Word's line; seen on the pages:

| Discrepancy (Word / ours) | Documents | Round |
|---|---|---|
| ~~A run's own font loses to a theme font set higher up~~ - not the engine: tools/pdf-runner gave the PDF writer five families only; it now gives the document's | annual-results, magazine-feature | done |
| Automatic hyphenation (`w:autoHyphenation`): Word hyphenates justified text ("flick-er", "conversa-tions"), we never do | magazine-feature | |
| A paragraph after a drop cap is not justified | magazine-feature | |
| Table column widths: Word follows the cells' widths (`w:tcW`), ours spreads the columns | annual-results, product-spec | |
| A table in a multi-column section spans the page instead of its column | field-guide (species cards) | |
| A paragraph's bottom border (a rule under text) drawn as a short stray dash | annual-results, field-guide, magazine | |
| Square wrap round a floating picture: text runs under it | diagrams, magazine-feature | |
| Top-and-bottom wrap round a floating text box: text runs under it | magazine-feature (pull quote) | |
| A callout shape's tail (`wedgeRectCallout`) not drawn | diagrams | |
| Row heights in a table with a vertically merged, vertically centred cell | annual-results (pipeline) | |
