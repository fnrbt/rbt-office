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
| A row breaks only where all its cells can: a paragraph between lines, a nested table between its rows, never inside a vertical-merge group | testing | calibre family: Word moves an outer row whose nested table has a vMerge, splits it without |
| Word drops the first paragraph's automatic space before | open | seen in Word, not yet probed |
| A list marker's font does not make its line taller (only the text does) | refuted | fixtures `list-marker-height`: Word's pitch is 15.44pt plain, with a Calibri bullet and numbered, but 16.03pt with Word's Symbol bullet and 21.05pt with a 16pt Symbol bullet - the marker counts. Ours: 16.07pt, already right |
| The footnote block ends at the bottom margin: its last line's bottom is the page height less the bottom margin, and the body may run down to it less the separator | testing | docx4j-getting-started (45 -> 46 pages): every paragraph of page 4 within 0.5pt of Word's, but our footnotes start 4.16-4.37pt above Word's on pages 4, 10 and 12 and end ~4.5pt above 720pt - so Word fits one more body line (`para-positions.vbs`, the footnotes' `Information(6)`) |
| A line with a footnote reference (superscript) is taller than the footnote text's single line | open | Word's two one-line footnotes on page 12 of docx4j-getting-started are 12.75pt apart; ours 12.21pt |
| Word fits ~1pt less per 468pt line than we do (test-long-document): kerning only with `w:kern`? | open | |
