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
  with `fetch-windows-fonts.sh`. The lines are `Word.getLayoutLines`' body lines.
- A page matches when its first text line is Word's (case, spaces and
  punctuation ignored, at most 40 characters compared). The table gives Word's
  pages, ours, how many of Word's pages start where ours do, and the first page
  that does not with both lines; `--detail` lists every page whose first or last
  line differs; `--lines <name>` prints our lines of one document (page, y,
  height, block, text) for finding where a break moved.

Run it before and after any layout change: a change must not make a document
that matched stop matching.
