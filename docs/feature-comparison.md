# rbt-office compared with other web docx / xlsx editors

As of 2026-11-11. Our side: `@rbt/office` 0.1.29 plus `master`, checked against the docs **and** `src/`. Competitor side: vendor docs, release notes, resellers and press found by web search in this session. Nothing was tested hands-on. **unverified** means no source was found this session; "(prior)" means general knowledge that wasn't re-checked.

**Historical comparison, not the current feature checklist.** The roadmap's [current capability matrix](continuation-roadmap.md#current-capability-matrix) is authoritative for our implementation, API/shell reachability and tested platforms. In particular, the old spreadsheet PDF/.NET-only claims below are superseded: the browser module and `ExportPdf` / `PrintPdf` exist; complete shell wiring and fresh release verification are separate. Pivot structural movement, Chromium full-height scrolling and shared-dialog coverage have also changed. Competitor claims and historical benchmark/oracle numbers have not been revalidated.

## 1. Summary

- **Where we stand out:** we are one of very few editors that are **fully client-side, page-accurate and keep what they don't model** for both docx and xlsx.
  - The ONLYOFFICE, Collabora, Univer, Syncfusion, Zoho and CKEditor families all need a server for the file, the conversion or both.
  - Apryse and Nutrient are client-side WASM, but commercial and lossy or limited in places.
  - SuperDoc and eigenpal are the nearest open competitors on docx; they are ProseMirror-based.
- **Layout fidelity and save safety are our strongest points:**
  - Word opened every save of 90 documents with the original's page count and text length (`docs/continuation-roadmap.md`, Cross-cutting).
  - Excel refused or repaired none of 333 saves of 111 workbooks.
  - Formula results agree with Excel on 50,480 of 50,496 cells (`docs/excel-parity-tracker.md`).
  - We render pattern shading, list indents, OMML math, fields/TOC, footnotes, chartEx waterfalls/funnels and slicers, and we export real PDF from Word in the browser.
- **What an evaluator will notice first:**
  - no real-time co-editing;
  - no spell check;
  - no touch support;
  - no screen-reader accessibility;
  - in the spreadsheet: charts can't be inserted or edited, pivots don't refresh, sparklines aren't drawn and there is no PDF/print in the browser;
  - in Word: no comment threads, charts and SmartArt shown as grey placeholders, no TOC insertion;
  - formulas: about 280 functions against Excel's ~500 and Univer's 528;
  - only Chromium has been tested;
  - no public licence or price yet.
- **Gaps already accepted for the alpha:** most of these are known to the roadmap. Its "Not in the alpha" list already excludes co-editing, macros in the browser, pivot refresh and chartEx maps/treemaps/sunbursts (`docs/continuation-roadmap.md`, "The initial release").

## 2a. docx matrix

Key: ✅ yes · ◐ partial · ❌ no · ? unverified. Column codes:

| Code | Product |
|---|---|
| **Us** | rbt-office |
| **MS** | Word for the web |
| **GD** | Google Docs |
| **OO** | ONLYOFFICE Docs |
| **CO** | Collabora Online |
| **ZW** | Zoho Writer |
| **SD** | SuperDoc v2 |
| **EP** | eigenpal docx-editor |
| **SF** | Syncfusion Document Editor |
| **NU** | Nutrient Document Authoring |
| **AP** | Apryse WebViewer DOCX |
| **DX** | DevExpress RichEdit |
| **CK** | CKEditor 5 + Word import/export |

| Feature | Us | MS | GD | OO | CO | ZW | SD | EP | SF | NU | AP | DX | CK |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| **Fidelity** | | | | | | | | | | | | | |
| True paginated layout | ✅¹ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ◐ | ❌ |
| Tables (merge, split across pages) | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ◐ | ◐ |
| Floating shapes / text boxes | ✅² | ✅ | ◐ | ✅ | ✅ | ? | ? | ◐ | ? | ◐ | ? | ❌ | ❌ |
| Headers / footers | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ◐ | ✅ | ✅ | ✅ | ✅ | ◐ |
| Footnotes / endnotes | ✅³ | ✅ | ✅ | ✅ | ✅ | ? | ? | ◐ | ✅ | ? | ? | ❌ | ◐ |
| Fields & TOC | ◐⁴ | ✅ | ✅ | ✅ | ✅ | ? | ? | ✅ | ✅ | ✅ | ❌ | ? | ❌ |
| Equations (OMML) | ◐⁵ | ✅ | ◐ | ✅ | ✅ | ? | ? | ◐ | ? | ? | ? | ❌ | ❌ |
| Charts | ❌⁶ | ◐ | ◐ | ✅ | ✅ | ? | ? | ◐ | ? | ? | ? | ❌ | ❌ |
| SmartArt | ❌⁶ | ◐ | ? | ✅ | ◐ | ? | ? | ◐ | ? | ? | ? | ❌ | ❌ |
| Content controls | ◐⁷ | ◐ | ? | ✅ | ✅ | ? | ? | ✅ | ✅ | ? | ? | ❌ | ❌ |
| **Editing** | | | | | | | | | | | | | |
| Track changes | ◐⁸ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ◐ᴾ | ✅ | ✅ | ✅ | ❌ | ◐ᴾ |
| Comments (threads, replies) | ◐⁹ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ◐ᴾ | ✅ | ? | ◐ | ❌ | ◐ᴾ |
| Styles | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ◐ |
| Lists | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Images | ◐¹⁰ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Tables | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Find / replace | ◐¹¹ | ✅ | ✅ | ✅ | ✅ | ✅ | ? | ✅ | ✅ | ? | ✅ | ✅ | ✅ |
| Undo / redo | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Spell check | ❌¹² | ✅ | ✅ | ✅ | ✅ | ✅ | ? | ? | ◐ˢ | ? | ? | ? | ◐ᴾ |
| **Round-trip** | | | | | | | | | | | | | |
| Save keeps what the editor doesn't model | ✅¹³ | ✅ | ◐ | ◐ | ◐ | ? | ? | ◐ | ❌ | ? | ? | ❌ | ❌ |
| **Collaboration** | | | | | | | | | | | | | |
| Real-time co-editing + presence | ❌ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ◐ᴾ | ◐ˢ | ❌ | ❌ | ❌ | ◐ᴾ |
| **Other** | | | | | | | | | | | | | |
| PDF export | ✅¹⁴ | ✅ | ✅ | ✅ | ✅ | ✅ | ? | ◐ᴾ | ◐ˢ | ✅ | ✅ | ✅ | ? |
| Print | ◐¹⁵ | ✅ | ✅ | ✅ | ✅ | ✅ | ? | ✅ | ✅ | ? | ? | ✅ | ✅ |
| Macros / VBA | ◐¹⁶ | ❌ | ❌ | ◐ JS | ◐ Basic | ? | ❌ | ❌ | ❌ | ❌ | ❌ | ❌ | ❌ |
| Mobile / touch | ❌¹⁷ | ✅ | ✅ | ✅ | ✅ | ✅ | ? | ? | ? | ? | ❌ | ? | ✅ |
| Accessibility (screen reader) | ❌¹⁸ | ✅ | ✅ | ? | ✅ | ? | ? | ? | ? | ? | ✅ | ? | ✅ |
| Large documents | ✅¹⁹ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ? | ? | ? | ? | ? | ◐ |
| Embed API / extensibility | ✅²⁰ | ◐ | ❌ | ✅ | ◐ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |

Superscript markers: ᴾ = paid or Pro tier only · ˢ = needs the vendor's or your own server.

Many of the competitor cells are (prior) knowledge: the research covered deployment and licensing more thoroughly than per-feature fidelity. Treat the competitor ✅ marks for ONLYOFFICE and Collabora fidelity as (prior).

**Footnotes (our cells):**

1. Our own layout engine on canvas, compared line by line with Word on a VM.
   - The round-trip set's page counts are not yet within one page of Word's everywhere. `layout/page-parity-2` is in progress, and some documents come out longer than Word's (`docs/continuation-roadmap.md`, Branches).
   - Line numbering (`w:lnNumType`) is not laid out.
2. Shapes, pictures and text boxes:
   - Done: preset and custom geometry, gradients, flips, wrap, `allowOverlap`, groups and drawing canvases (`src/Word/Shapes.fs`; roadmap Backlog/Word).
   - Group shapes have been checked on synthetic documents only.
   - Not edited: text inside a group's shapes, and a group's turn or flip.
3. Laid out from their blocks (0.1.29). Not yet: Word's numbering restarts and formats (`w:numStart`/`w:numFmt`), and footnotes set in columns (roadmap Backlog/Word).
4. Fields:
   - Computed and refreshed: PAGE, NUMPAGES, SECTIONPAGES, REF/PAGEREF, SEQ, STYLEREF, DATE/TIME and TOC (`refreshFields`, `docs/word-host-api.md`).
   - **There is no "insert table of contents" call** (no such export in `src/Word/Word.fs`).
   - AUTHOR and FILENAME are not computed (`docs/docx-editor-parity.md`).
5. OMML is laid out and drawn in every painter (`src/Word/Math.fs`). Insertion takes linear text only (`insertEquation`); there is no structural equation editor.
6. Charts, SmartArt and OLE objects in Word show as a **grey frame labelled "Chart" / "SmartArt"** (`src/Word/Paint.fs` ~1180). They are kept on save. The spreadsheet does draw charts.
7. Content controls:
   - Kept, listed and wrappable through the API (`setContentControl`, `getContentControlInfos`); checkboxes are drawn.
   - Dropdowns and date pickers can't be used interactively.
   - Block-level controls around rows, cells or repeating sections lose their binding on a save from the model (roadmap, alpha "Must be true" #1).
8. Track changes:
   - Typing and deleting are recorded (`setTrackChanges`), and changes can be accepted or rejected singly or all at once, and stepped through (`docs/word-host-api.md`, Comments and tracked changes).
   - Formatting revisions (`rPrChange`) are read (`DocxReader.fs:3305`).
   - Whether formatting changes and structural changes (paragraph marks, table rows) are *recorded* is **unverified**.
   - The review UI is the host's job: there is no balloon pane.
9. Comments:
   - Anchors are listed (`getCommentAnchors`), bodies are edited in place in a balloon (`enterComment`), and new comments are set on a selection (`setComment`).
   - **No replies or threads.** `commentsExtended` is only kept as a part (`DocxReader.fs:4220`); there is no reply API.
   - The comment pane is the host's.
10. Pictures are inserted, resized and moved. A picture on the clipboard (a screenshot) is not pasted, and drag-and-drop of text is missing (roadmap Backlog/Word, Editing).
11. Literal text with match case only (`findText`, `replaceAllText`): no whole word, wildcards, regex or formatting search.
12. No spell check. The hidden input proxy has `spellcheck="false"` (`src/Word/Input.fs:1893`). A "Check spelling as you type" checkbox exists in `WordUi.fs:665`, but nothing implements it.
13. Round-trip:
   - One save path (0.1.28, `docs/save-architecture.md`): an unchanged part is its original bytes, an edited paragraph or table is its source XML with the edit spliced in, and unknown elements are kept.
   - Known losses, all listed in the alpha must-haves:
     - a space lost between runs, in saves written from the model;
     - the binding of block-level content controls;
     - a bookmark renamed in a spliced paragraph keeps its old name.
   - `docs/docx-editor-parity.md` still says "🔴 regenerates XML". **That row is out of date.**
14. Real PDF with embedded fonts and our shaper. `src/Word.Pdf` is compiled to JS (`dist/word-pdf`), but the host must install it as `window.RedBatWordPdfExport` (`src/Word/WordHost.fs:171`).
15. `renderPage` paints any page for a print preview. There is no built-in print command; a host prints the PDF.
16. VBA/VBScript runs read-only layout queries only (`WordVba.fs`): no document edits, no `.docm` macros (roadmap, Macros).
17. Mouse events only. There is no `touch*` or `pointer*` listener anywhere in `src/` (grep), so no touch selection handles, pinch zoom or long-press. IME and virtual keyboards go through a hidden textarea (`Input.fs`).
18. The document is a canvas: there is no accessible text tree and no ARIA live region. The only ARIA is a label on the input proxy (`Input.fs:1889`) and `aria-pressed`/`aria-expanded` on toolbar buttons.
19. Long documents are laid out in slices (progressive open and layout). The 514-page harness document is part of the redbat gate, and the 600-page perf document takes 30 ms per keystroke in a footnote (roadmap).
20. Large host API, about 230 documented functions (`docs/word-host-api.md`), plus a separate small viewer module. It is only a Fable/JS module: no React or Vue wrappers, no events beyond `onChange`, and no i18n.

**Footnotes (competitor cells):**

- **Word for the web:** charts and SmartArt are shown, but authored only on desktop, and the same goes for content controls. Embedding needs WOPI/CSPP, and only storage vendors are accepted. Source: MS comparison page.
- **Google Docs:** Office editing mode keeps some unsupported content and warns about the rest; embedded charts become images.
- **ONLYOFFICE / Collabora:** both convert to an internal model on the server (x2t or LibreOffice), so the round-trip is ◐ (prior).
- **eigenpal:**
  - Its own fidelity matrix: charts, shapes and SmartArt render read-only and are kept on save; unknown XML is kept as generic nodes; there is no byte-for-byte guarantee.
  - Tracked changes, comments, collaboration and PDF are in the Pro licence.
  - The research says the core licence is **Apache-2.0**; our `docs/docx-editor-parity.md` says MIT. Re-check.
- **Syncfusion:** opening a DOCX needs its server web service (SFDT conversion), as do spell check and collaboration. A July 2026 "client-side" launch is reported, but **whether it removes the server is unverified**.
- **Apryse:** self-reported gaps are no TOC, no page-number fields, no co-editing and no mobile browsers.
- **DevExpress:** officially unsupported are track changes, comments, footnotes, floating objects and charts.
- **CKEditor:** converts DOCX to HTML, so it is not paginated; Word import is on the Professional tier and up. The real-time collaboration tier is (prior).

## 2b. xlsx matrix

Column codes:

| Code | Product |
|---|---|
| **Us** | rbt-office |
| **XW** | Excel for the web |
| **GS** | Google Sheets |
| **OO** | ONLYOFFICE |
| **CO** | Collabora Calc |
| **SJ** | SpreadJS |
| **UN** | Univer (Pro where noted) |
| **SF** | Syncfusion Spreadsheet |
| **HT** | Handsontable + HyperFormula |
| **FS** | FortuneSheet |
| **JS** | Jspreadsheet Pro |
| **AP** | Apryse Spreadsheet Editor |

x-spreadsheet is now a table component (`@wolf-table/table`) with no xlsx support, and Luckysheet is archived (Mar 2024). Both are left out.

| Feature | Us | XW | GS | OO | CO | SJ | UN | SF | HT | FS | JS | AP |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| **Fidelity** | | | | | | | | | | | | |
| Function count | ~280ᵃ | ~500 | ~512 | 400–450+ | ~500 (prior) | 500+ | 528 | ? | 423 | ? | 400–500 | 260+ |
| Dynamic arrays / LAMBDA | ✅ | ✅ | ✅ | ✅ (9.3) | ◐ | ✅ | ? | ? | ❌ | ❌ | ? | ❌ |
| Charts: shown | ✅ᵇ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ᴾ | ✅ | ❌ | ❌ | ✅ | ✅ |
| Charts: insert / edit | ❌ᵇ | ✅ | ✅ | ✅ | ✅ | ✅ | ◐ᴾ | ✅ | ❌ | ❌ | ✅ | ✅ |
| Excel 2016 charts (waterfall, funnel, treemap) | ◐ᶜ | ✅ | ◐ | ◐ open only | ◐ kept | ✅ | ? | ? | ❌ | ❌ | ? | ? |
| Conditional formatting (incl. bars, scales, icons) | ✅ᵈ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ◐ | ✅ | ? | ? |
| Pivot tables | ◐ᵉ | ✅ | ✅ | ✅ | ✅ | ✅ᴾ | ✅ᴾ | ◐ | ❌ | ❌ | ✅ᴾ | ? |
| Slicers / timelines | ◐ᶠ | ✅ | ✅ | ✅ | ? | ✅ᴾ | ✅ᴾ | ? | ❌ | ❌ | ? | ? |
| Sparklines | ❌ᵍ | ✅ | ◐ fn | ✅ | ✅ | ✅ | ✅ | ? | ❌ | ❌ | ? | ? |
| Table styles | ✅ʰ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ? | ❌ | ❌ | ? | ? |
| Data validation | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ? |
| Freeze panes | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ? |
| **Editing** | | | | | | | | | | | | |
| Comments (threaded) | ✅ⁱ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ◐ | ✅ | ✅ | ? |
| Find / replace | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ◐ | ✅ | ✅ | ✅ |
| Sort / filter / tables | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ◐ | ✅ | ? |
| Undo / redo | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| Named cell styles | ❌ʲ | ✅ | ❌ | ✅ | ✅ | ✅ | ? | ✅ | ❌ | ❌ | ? | ? |
| **Round-trip** | | | | | | | | | | | | |
| xlsx import + export client-side | ✅ | n/a | n/a | ❌ˢ | ❌ˢ | ✅ | ❌ˢ | ❌ˢ | ◐ export only | ◐ plugin | ✅ | ✅ |
| Save keeps what the editor doesn't model | ✅ᵏ | ✅ | ◐ | ? | ◐ | ? | ? | ? | ❌ | ❌ | ? | ? |
| **Collaboration** | | | | | | | | | | | | |
| Real-time co-editing + presence | ❌ | ✅ | ✅ | ✅ | ✅ | ◐ add-on | ✅ᴾ | ◐ DIY | ❌ | ◐ DIY | ◐ᴾ | ❌ |
| **Other** | | | | | | | | | | | | |
| PDF export | ◐ˡ | ✅ | ✅ | ✅ | ✅ | ✅ | ◐ᴾ | ◐ˢ | ❌ | ❌ | ✅ᴾ | ? |
| Print / page setup / page-break preview | ◐ˡ | ✅ | ✅ | ✅ | ✅ | ✅ | ◐ᴾ | ✅ | ❌ | ? | ? | ? |
| Macros / VBA | ◐ᵐ | ❌ (Office Scripts) | ❌ (Apps Script) | ◐ JS | ◐ Basic | ❌ (JS fns) | ❌ | ❌ | ❌ | ❌ | ❌ | ❌ |
| Mobile / touch | ❌ⁿ | ✅ | ✅ | ✅ | ✅ | ? | ◐ | ? | ✅ | ? | ? | ❌ |
| Accessibility | ◐ⁿ | ✅ | ✅ | ? | ✅ | ◐ | ? | ◐ ARIA | ✅ | ? | ? | ? |
| Large files | ✅ᵒ | ✅ | ✅ | ✅ | ◐ | ✅ | ✅ | ◐ | ✅ | ◐ | ✅ | ? |
| Embed API | ✅ | ◐ WOPI | ❌ | ✅ | ◐ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |

**Footnotes (our cells):**

- **a. Function count:**
  - Counted from the dispatch labels in `src/Spreadsheet/Formula.fs` (≈278 names) plus a few handled in `Recalc.fs`. That is a source count, not a published figure.
  - Present: XLOOKUP, LET, LAMBDA, MAP/REDUCE/SCAN, FILTER/SORT/UNIQUE, GROUPBY/PIVOTBY, XIRR/XNPV and the database functions.
  - Missing, for example:
    - most statistical distributions (BINOM.DIST, T.TEST, GAMMA, CONFIDENCE);
    - DDB, SYD and the bond functions (PRICE, YIELD);
    - EXPAND, WRAPROWS, REGEX*, IMAGE, FORMULATEXT and SHEET (as functions; SHEET and FORMULATEXT are only named in `Recalc.fs`).
  - Accuracy on what exists is high: 50,480 of 50,496 cells agree with Excel 2019 (`docs/excel-parity-tracker.md`).
- **b. Charts:**
  - Bar, column, line, pie, doughnut, area and scatter charts, combos and secondary axes are drawn with live values (SVG in the grid; `docs/spreadsheet-host-api.md`, "Charts and shapes of a file"). Radar, bubble, stock and 3-D-as-flat are drawn too (roadmap).
  - **Not editable, not movable, not insertable.** `InsertChart` is a menu entry with no implementation; the wisp host says "Insert chart not wired" (`spreadsheet-wisp/Program.fs`).
  - The portable canvas painter and the PDF still draw a labelled frame for some objects. Trend lines and error bars are not drawn.
- **c. chartEx:** waterfall and funnel are drawn and checked against Excel 2019. Map, sunburst, treemap, histogram, box-and-whisker and Pareto charts are placeholder boxes. All are kept byte for byte (roadmap Backlog/Spreadsheet).
- **d. Conditional formatting:** cellIs, expression, text, top/bottom, above-average, duplicates, dates, colour scales, data bars and icon sets, plus x14 rules, with a host API to add, change and reorder them (`docs/spreadsheet-host-api.md`, Conditional formats). Pivot-table CF and unknown x14 rule types are kept, not understood.
- **e. Pivot tables:**
  - A pivot of a file shows its cached cells, is kept on save and is copied with a duplicated sheet. GETPIVOTDATA works.
  - **No refresh from the cache.** A pivot doesn't move with row or column inserts (roadmap Backlog/Spreadsheet features).
  - New pivots are GROUPBY/PIVOTBY spilling formulas (`src/Spreadsheet/Pivot.fs`). Whether the builder dialog is wired in the browser component is **unverified**.
- **f. Slicers and timelines:** drawn and kept, within about 1 px of Excel. **Clicking doesn't filter yet** (roadmap).
- **g. Sparklines:** kept on save and moved with their cells, but **not drawn in the browser grid**. There is no sparkline code in `Spreadsheet.fs`, `Paint.fs` or `Charts.fs`; they are drawn only in the .NET PDF and native runner (`Spreadsheet.Pdf/Pdf.fs:1283`).
- **h. Table styles:** Excel's own built-in definitions plus custom ones, with tints matching Excel exactly for 250 colour/tint pairs. Known faults (roadmap):
  - black text in a white-text header shows white;
  - a style chosen through the API uses the Office theme's colours, not the workbook's;
  - on sheets with merges, only the lines above and below a table's cells are drawn.
- **i. Comments:** notes plus threaded comments with replies and the done mark, written to the threaded-comments part (`docs/spreadsheet-host-api.md`, Comments).
- **j. Named cell styles:** "Not available: the model has no named cell styles" (`docs/spreadsheet-host-api.md`). They are kept on export but can't be applied.
- **k. Round-trip:** an unchanged workbook is exported as the file that was opened, an unchanged sheet as its part, and unmodelled parts byte for byte. Excel refused or repaired none of 333 saves of 111 workbooks (roadmap Cross-cutting). Row-insert saves mostly fall back to the model writer, but are still accepted by Excel.
- **l. PDF and print:** pagination, print areas and titles, breaks and a page-break preview are in the browser bundle. **The PDF writer (`Rbt.Spreadsheet.Pdf`) is .NET only** ("only the PDF output itself is .NET only", `docs/spreadsheet-host-api.md` §Print). So there is no browser PDF or print for workbooks, even though we claim to be "fully client-side".
- **m. Macros:** VBA module source is read and written (`VbaProject.fs`). Macros run only in the native wisp host, setting Value/Formula only; **the browser component runs none** (roadmap, Macros).
- **n. Touch and accessibility:** no touch or pointer handlers. The grid is DOM cells, which is better for accessibility than a canvas, but it has no ARIA grid roles; `aria-*` appears only in `Toolbar.fs`. Accessibility is **unverified** with a screen reader.
- **o. Large files:** on the 846,000-cell sheet, a row insert takes ~100–160 ms and a sort ~0.65–1.0 s. The last rows can't be reached above ~125% zoom in Chromium, nor in Firefox (roadmap).

**Footnotes (competitor cells):**

- **ˢ** = the xlsx file goes through a server:
  - ONLYOFFICE: the Document Server.
  - Collabora: tile rendering on the server.
  - Univer: Univer Server's Exchange service, even with the Apache-2.0 core.
  - Syncfusion: XlsIO through openUrl/saveUrl.
- **Univer Freemium:** 1 pivot per workbook, line/column/bar charts only, a watermark, prints limited to 3 pages and 10 exports a month.
- **Google Sheets:** native Excel tables and table-based pivots have been kept on import only since Aug 2026. Macros and protection are not supported.
- **Collabora:** chartEx charts survive a save since 25.04.5.3, but upstream LibreOffice doesn't render waterfall or treemap charts.
- **Nutrient Document Authoring 2.0** (late 2026) also edits xlsx, but drops charts, images and CF on export and has only about six functions. Left out of the matrix as not competitive.

## 3. Deployment and licensing

| Product | Runs where | Self-host | Licence | Price (public) | Embed SDK |
|---|---|---|---|---|---|
| **rbt-office** | **Browser only, no server** (Word PDF in browser; xlsx PDF .NET only) | n/a (static files) | **Undecided**: currently private `@rbt/office` on the Forge registry; no licence field in `package.json`; licence and public npm are open decisions (roadmap) | none | JS modules (`Word.js`, spreadsheet component), host API docs; editor shell optional (open decision); native .NET hosts built but never run interactively |
| Word / Excel for the web | Microsoft cloud | ❌ | M365 | subscription | WOPI/CSPP, storage vendors only; Office.js add-ins |
| Google Docs / Sheets | Google cloud | ❌ | Workspace | subscription | none for editing (REST APIs, Apps Script) |
| ONLYOFFICE Docs | Browser plus Document Server | ✅ | AGPLv3 Community (20-connection cap; 9.4 reportedly lifted it, FAQ disagrees) / commercial | Enterprise $33 per connection per year (min 50); Developer quote only (reseller: $17.8k for 250) | Docs API (iframe), WOPI, plugins |
| Euro-Office (2026) | as ONLYOFFICE (fork of 9.3.1) | ✅ | AGPL, disputed by ONLYOFFICE | ? | as ONLYOFFICE |
| CryptPad | Browser (x2t compiled to WASM), E2E encrypted | ✅ | AGPL (prior) | free | app, not SDK |
| Collabora Online | Server (LibreOffice core renders tiles) | ✅ | MPL-2.0 (prior); CODE free for testing and home use | Business €3 per user per month; Enterprise quote | WOPI + PostMessage |
| Zoho Writer / Office Integrator | Zoho cloud | ❌ | SaaS | pay per document opened: 50 free a month; $50 / $150 for 2,000 a month | iframe API |
| SuperDoc (v2) | Browser; headless SDKs | ✅ | AGPLv3 or commercial | quote | JS, React, Vue; CLI, MCP |
| eigenpal docx-editor | Browser (ProseMirror + own painter) | ✅ | Apache-2.0 core (our doc says MIT, re-check) + Pro (tracked changes, comments, collaboration, PDF) | Pro quote | React, Vue |
| Syncfusion Document Editor / Spreadsheet | Browser UI, **server for DOCX/XLSX open and save** | ✅ (your server) | commercial; community licence for small firms | Spreadsheet SDK $1,199 per developer per year | JS, React, Angular, Vue, Blazor |
| Nutrient Document Authoring | Browser WASM, Node headless | ✅ | commercial | quote | JS |
| Apryse WebViewer (DOCX / spreadsheet) | Browser WASM, no server; not on mobile | ✅ | commercial add-on | quote | JS |
| DevExpress RichEdit | Client plus .NET backend | ✅ | commercial | subscription (prior) | ASP.NET Core / JS |
| Telerik / Kendo | HTML editor, server-side DOCX conversion | ✅ | commercial | subscription | no paginated web DOCX editor |
| TX Text Control | .NET backend; newer WASM build | ✅ | commercial | ? | JS / .NET |
| CKEditor 5 + Word converters | Browser HTML editor; conversion SaaS or on-prem Docker | ◐ (licence key) | GPL core / commercial | Word import on Professional and up | JS |
| Tiptap Conversion | Import on a SaaS/REST service, export in browser | ◐ | MIT core, paid Pro extensions | subscription | JS |
| SpreadJS | Browser only (collaboration server add-on) | ✅ | commercial | Developer $999; deployment $12,499 per domain; Pivot add-on $8,399 | JS, Designer, React/Angular/Vue |
| Univer | Canvas client, **server needed for xlsx**, collaboration, Pro | ✅ (Docker/K8s) | Apache-2.0 core; Pro closed | Freemium (limited); Business $99 per month / from $1,000 per year | Facade API, plugins |
| Handsontable + HyperFormula | Browser (DOM grid) | ✅ | commercial; HyperFormula GPLv3 or commercial | quote (reseller ~$1–1.9k per developer per year) | JS + wrappers |
| FortuneSheet | Browser (React) | ✅ | MIT | free | React; xlsx via third-party plugin; low activity since Apr 2025 |
| Jspreadsheet CE / Pro | Browser (DOM); Pro "Server" for collaboration | ✅ | CE MIT / Pro commercial | Enterprise $2,499 per year; Ultimate $4,499 per year | JS |
| Rows n Columns | Browser (React) | ✅ | commercial | pricing page (not checked) | React; Yjs/ShareDB hooks |
| IronCalc | Rust/WASM engine + web app | ✅ | MIT / Apache-2.0 | free | early stage |
| Quadratic | SaaS / Docker | ◐ | reportedly closed since Mar 2026 (unverified) | ? | app |

**Takeaway:** fully client-side xlsx with real import and export comes from SpreadJS, Apryse, Jspreadsheet Pro, Rows n Columns, IronCalc (early) and us. Fully client-side, paginated docx comes from Apryse, Nutrient, SuperDoc, eigenpal and us. Of these we are the only one covering **both formats in one package with byte-level preservation of unmodelled parts**. Apryse and Nutrient cover both, but are commercial, closed and partly lossy (Nutrient's xlsx export drops charts and CF).

## 4. Where we lead

1. **No server, both formats, one package.** Most competitors need a document or conversion server: ONLYOFFICE, Collabora, Univer for xlsx, Syncfusion, CKEditor and Tiptap import. Apryse is the only other client-side package with both formats, and it is closed, commercial and not on mobile browsers.
2. **Save safety, measured in real Office.**
   - Word: every save of 90 documents opens in real Word with the original's pages and text length.
   - Excel: none of 333 saves refused or repaired.
   - Untouched parts go back byte for byte, including chartEx, slicers, pivots, sparklines, VBA projects and custom XML.
   - eigenpal says "no byte-for-byte guarantee"; Nutrient's xlsx drops charts and CF; Syncfusion converts everything to SFDT or JSON.
   - Evidence: `docs/save-architecture.md`, roadmap Cross-cutting, `docs/spreadsheet-host-api.md` "XLSX: what is written how".
3. **Layout measured against Word itself.**
   - Our VBA-style layout queries (`getRangeInformation`, `computeRangeStatistics`) are compared line by line with Word on a VM, and we have a three-way harness against eigenpal.
   - We are closer to Word than eigenpal on paragraph spacing, list indents, pattern shading and top-and-bottom floats (`docs/docx-editor-parity.md`).
4. **Formula accuracy.** 50,480 of 50,496 cells agree with Excel. Dynamic arrays, LAMBDA, GROUPBY/PIVOTBY, structured references and iterative calculation are in. Apryse has 260+ functions and no XLOOKUP or IFS; Handsontable has no LAMBDA.
5. **Rendering of advanced Excel content without a server:**
   - waterfall and funnel charts, slicers and timelines within about 1 px of Excel;
   - Excel's exact table-style tints;
   - full CF including icon sets and data bars.
   Univer Freemium caps charts to three types, and ONLYOFFICE opens chartEx but can't insert it.
6. **True PDF from the browser for Word** (embedded fonts, our shaper), sharing one layout with the screen. Many competitors put PDF in a paid tier or on a server: eigenpal Pro (Node only), Syncfusion (server).
7. **Large-file performance without a server:** 846k-cell sheet edits in about 100–160 ms, and progressive layout of 500–600-page documents.
8. **Read-only VBA layout object model**, plus macros in the native host. No other embeddable SDK has VBA (ONLYOFFICE and Collabora use JS or Basic).
9. **Native hosts:** the same F# engine has .NET native and PDF targets. They are built but not yet run interactively, so don't market this yet.

## 5. Gaps that matter for an initial release

Ranked by how soon a first-time user or evaluator would notice the gap.

| # | Gap | Who notices, how fast | Evidence | Field status |
|---|---|---|---|---|
| 1 | **No licence, no public package, no demo page** | Every evaluator, before opening a file | `package.json` has no licence and is private-scoped. Roadmap "Open decisions" lists licence, npm, shell; alpha must-have #7 (demo page, changelog, honest feature list) | Every competitor has a published licence; most have a live demo |
| 2 | **Only Chromium is tested** (Firefox untested, Safari never run); last rows unreachable at over 125% zoom and in Firefox | Anyone on Safari/Firefox or a Mac, within seconds | Roadmap alpha #4 and Backlog ("Firefox is untested", "last rows are unreachable…") | All the SaaS and SDK competitors support every evergreen browser |
| 3 | **No spell check** in Word (red squiggles) | Every writer, within a minute of typing | `src/Word/Input.fs:1893` `spellcheck=false`; the `WordUi.fs:665` checkbox does nothing | Every suite has it; ONLYOFFICE added AI grammar in 9.2 |
| 4 | **Charts can't be inserted or edited in the spreadsheet**; sparklines are not drawn | Any spreadsheet evaluator trying Insert › Chart, or opening a dashboard with sparklines | `InsertChart` unimplemented ("Insert chart not wired", `spreadsheet-wisp/Program.fs`); `docs/spreadsheet-host-api.md` "Nothing of it can be edited"; no sparkline code in `Spreadsheet.fs`/`Paint.fs`/`Charts.fs` | Every spreadsheet competitor except Handsontable and FortuneSheet inserts charts |
| 5 | **Charts and SmartArt in Word show as grey "Chart" / "SmartArt" boxes** | Anyone opening a business report | `src/Word/Paint.fs` ~1180; `docs/docx-editor-parity.md` | eigenpal renders them read-only from saved values; Word web shows them |
| 6 | **No real-time co-editing or presence** | Evaluators comparing with Google/M365/ONLYOFFICE; it's the first question about "a Google Docs alternative" | No collaboration code in `src/` (grep: no collab/CRDT/Yjs); roadmap "Not in the alpha"; `docx-editor-parity.md` architecture note (we'd need our own OT/CRDT) | SuperDoc, ONLYOFFICE, Collabora, Univer Pro, SpreadJS add-on, eigenpal Pro all have it |
| 7 | **No touch support** (phones and tablets) | Anyone opening the demo on a phone or iPad | No `touch*`/`pointer*` listeners in `src/` (grep); Word input uses `MouseEvent` (`Input.fs:1433`) | The suites have mobile UIs; Univer added a mobile UI in 2026; Apryse is also not on mobile |
| 8 | **Spreadsheet PDF and print don't work in the browser** (.NET only) | Anyone pressing Print or "Download as PDF" on a workbook; also contradicts the "no server" message | `docs/spreadsheet-host-api.md` §Print: "only the PDF output itself is .NET only"; `package.json` doesn't compile `Spreadsheet.Pdf` | Universal in the suites; SpreadJS `savePDF` client-side |
| 9 | **Word comments: no replies or threads; track-changes review UI left to the host** | Any reviewer or legal user; on a Word doc with threaded comments, the replies aren't shown as a thread | No reply API (`docs/word-host-api.md` Comments and tracked changes); `commentsExtended` only kept (`DocxReader.fs:4220`). The xlsx side does have threads | Universal in the suites; SuperDoc threaded; eigenpal Pro |
| 10 | **Pivot tables don't refresh; slicers don't filter** | Anyone who changes source data or clicks a slicer on a dashboard | Roadmap Backlog/Spreadsheet features ("refresh a pivot from its cache"; "Left: a click on a button filtering") | Excel web, Sheets, ONLYOFFICE, SpreadJS, Univer Pro |
| 11 | **Page counts not always Word's** (some documents run longer) | Anyone comparing a multi-page file with Word side by side | Roadmap alpha #2, branch `layout/page-parity-2`; e.g. `docx4j-headers-footers` 42 against Word's 31 on `layout/page-parity` | The server suites are generally close (prior); eigenpal and SuperDoc have their own drift |
| 12 | **Function coverage ~280 against ~500** (statistical distributions, DDB/SYD, bonds, REGEX, EXPAND/WRAPROWS) | Finance and stats users; shows as `#NAME?` in their own workbooks | Function list in `src/Spreadsheet/Formula.fs` (count in fn. a) | Univer 528, Sheets ~512, SpreadJS 500+, HyperFormula 423 |
| 13 | **Numbers of General alignment drawn left-aligned** in the web grid | Every spreadsheet user, at once (looks un-Excel-like) | Roadmap Backlog/Spreadsheet bugs; alpha #3 | none (a bug) |
| 14 | **Find and replace is literal only** in Word (no whole word, wildcards, formats); **no TOC insertion** | Power users | `findText` / `replaceAllText` signatures; no TOC insert export in `Word.fs` | Standard in the suites |
| 15 | **Screen-reader accessibility** (canvas Word with no accessible text; grid without ARIA roles) | Public sector and enterprise procurement (WCAG/508 checklists), not casual users | `Input.fs:1889` (only an aria-label); no `role=` in `src/` | Apryse WCAG 2.2 AA; Collabora certified; Microsoft and Google |
| 16 | **Pasted screenshot not pasted; no drag-and-drop of text; copy across table cells not a table** | Daily users, within a session | Roadmap Backlog/Word, Editing | Standard |
| 17 | **No framework wrappers, i18n or UI shell decision** | Integrators | `docs/docx-editor-parity.md` §E; roadmap "Open decisions" (shell) | eigenpal React/Vue, 9 locales; SpreadJS and Syncfusion all frameworks |

Not ranked as release gaps, but say them in the feature list:

- macros don't run in the browser;
- named cell styles can't be applied;
- line numbering isn't laid out;
- chartEx maps, treemaps and sunbursts are placeholders;
- Word watermarks: VML text-path watermarks are drawn as fitted text (`Document.fs` `FitText`), but this is unverified against Word.

## 6. Sources

**Ours** (paths relative to the repo):
- `docs/continuation-roadmap.md` (alpha section, Backlog, Cross-cutting)
- `docs/word-host-api.md`
- `docs/spreadsheet-host-api.md`
- `docs/docx-editor-parity.md` (June 2026; its round-trip row is out of date)
- `docs/excel-parity-tracker.md`
- `docs/excel-advanced-features-tracker.md` (June 2026; superseded in places)
- `docs/save-architecture.md`
- `docs/word-parity.md`, `docs/excel-parity.md` (process docs)
- Code: `src/Word/{Input,Paint,WordHost,WordUi,Document}.fs`, `src/Office.Formats/src/{DocxReader,FormulaNames}.fs`, `src/Spreadsheet/{Formula,Recalc,Spreadsheet,SpreadsheetUi,Toolbar}.fs`, `src/Spreadsheet.Pdf/Pdf.fs`, `spreadsheet-wisp/Program.fs`, `package.json`

`src/Spreadsheet/EXCEL-FEATURES-STATUS.md` and `UI-FEATURES-WIP.md` date from May 2026 and are stale; don't cite them.

**docx competitors:**
- Microsoft:
  - https://support.microsoft.com/en-us/office/word-features-comparison-web-vs-desktop-3e863ce3-e82c-4211-8f97-5b33c36c55f8
  - https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/online/overview
- Google:
  - https://support.google.com/docs/answer/6055139
  - https://support.google.com/drive/answer/9276408
- ONLYOFFICE:
  - https://www.onlyoffice.com/developer-edition-prices
  - https://onlyoffice.com/blog/2025/10/docs-9-1-released
  - https://onlyoffice.com/blog/2025/12/onlyoffice-docs-9-2
  - https://www.componentsource.com/product/onlyoffice-docs-developer/prices
- Euro-Office:
  - https://en.wikipedia.org/wiki/Euro-Office
  - https://onlyoffice.com/blog/2026/03/onlyoffice-flags-license-violations-in-euro-office-project-by-nextcloud-and-ionos
- CryptPad: https://awesome.ecosyste.ms/projects/github.com%2Fcryptpad%2Fonlyoffice-x2t-wasm
- Collabora:
  - https://www.collaboraonline.com/subscriptions/
  - https://www.collaboraonline.com/collabora-online-25-04-release-notes/
- Zoho: https://www.zoho.com/officeintegrator/pricing.html
- Apryse:
  - https://docs.apryse.com/web/guides/docx-editor
  - https://apryse.com/capabilities/docx-editor/comparison (a competitor's comparison: biased, partly outdated)
- Nutrient: https://www.nutrient.io/sdk/document-authoring
- Syncfusion:
  - https://help.syncfusion.com/document-processing/word/word-processor/javascript-es5/web-services-overview
  - https://www.syncfusion.com/docx-editor-sdk/document-elements
- DevExpress: https://docs.devexpress.com/AspNetCore/401837/rich-edit/unsupported-features
- Telerik:
  - https://www.telerik.com/aspnet-mvc/documentation/html-helpers/editors/editor/import-export
  - https://www.telerik.com/kendo-angular-ui/components/knowledge-base/editor-import-docx-files
- TX Text Control: https://www.textcontrol.com/blog/tag/webassembly/2026/
- CKEditor:
  - https://ckeditor.com/docs/ckeditor5/latest/features/export-word.html
  - https://ckeditor.com/docs/cs/latest/onpremises/docx-onpremises/overview.html
- SuperDoc:
  - https://github.com/superdoc-dev/superdoc
  - https://docs.superdoc.dev/editor/migrate-from-v1/overview/
- eigenpal:
  - https://github.com/eigenpal/docx-editor
  - https://www.docx-editor.dev/docs/2.x/docx-fidelity
- Tiptap: https://tiptap.dev/docs/conversion/getting-started/feature-support-matrix
- Others:
  - https://github.com/portone-io/docx-editor
  - https://next.platejs.org/docs/docx-io
  - https://trendshift.io/repositories/227638 (Docen)

**xlsx competitors:**
- Microsoft:
  - https://support.microsoft.com/en-us/office/differences-between-using-a-workbook-in-the-browser-and-in-excel-f0dc28ed-b85d-4e1d-be6d-5878005db3b6
  - https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/faq/file-sizes
  - https://learn.microsoft.com/en-us/office/dev/scripts/resources/vba-differences
- Google:
  - https://support.google.com/docs/answer/190718
  - https://workspaceupdates.googleblog.com/2026/09/doubled-cell-limits-in-google-sheets-now-generally-available.html
  - https://workspaceupdates.googleblog.com/2026/08/improved-file-import-google-sheets-tables.html
- ONLYOFFICE:
  - https://helpcenter.onlyoffice.com/docs/userguides/spreadsheet_editor/InsertChart.aspx
  - https://onlyoffice.com/blog/2026/02/onlyoffice-docs-9-3
  - https://www.it-connect.tech/onlyoffice-docs-9-4-drops-the-20-connection-limit-and-updates-its-open-source-license/
  - https://www.onlyoffice.com/blog/2026/10/onlyoffice-docs-10
- Collabora / LibreOffice:
  - https://www.collaboraonline.com/collabora-online-25-04-release-notes/
  - https://www.libreoffice.org/bugzilla/show_bug.cgi?id=120609
- SpreadJS:
  - https://developer.mescius.com/spreadjs
  - https://developer.mescius.com/spreadjs/pricing
  - https://developer.mescius.com/spreadjs/docs/v19/features/Charts/ChartTypes
- Univer:
  - https://univer.ai/pricing
  - https://docs.univer.ai/guides/sheets/features/import-export
  - https://docs.univer.ai/guides/sheets/features/core/formula
  - https://docs.univer.ai/guides/sheets/pro-features/server/docker
- Syncfusion:
  - https://ej2.syncfusion.com/documentation/spreadsheet/open-save
  - https://www.syncfusion.com/feedback/18431/client-side-import-export-for-load-save-in-spreadsheet
  - https://www.syncfusion.com/sales/teamlicense
- Handsontable:
  - https://handsontable.com/pricing
  - https://handsontable.com/docs/hyperformula/guide/built-in-functions
  - https://handsontable.com/docs/vue-data-grid/changelog-17/
- FortuneSheet / Luckysheet:
  - https://github.com/ruilisi/fortune-sheet
  - https://github.com/dream-num/Luckysheet
  - https://www.npmjs.com/package/@corbe30/fortune-excel
- Jspreadsheet:
  - https://github.com/jspreadsheet/ce
  - https://jspreadsheet.com/pricing
- x-spreadsheet successor: https://github.com/wolf-table/table
- Apryse:
  - https://docs.apryse.com/web/guides/spreadsheet-editor
  - https://apryse.com/blog/spreadsheet-editor-chart-support
- Nutrient: https://www.nutrient.io/guides/document-authoring/spreadsheets/
- Rows n Columns: https://docs.rowsncolumns.app/getting-started/excel-google-sheet-compatibility
- IronCalc: https://github.com/ironcalc/IronCalc
- Quadratic: https://docs.quadratichq.com/import-data/import-excel-files
- EtherCalc: https://newreleases.io/project/github/audreyt/ethercalc/release/0.20260710.0
- Grist: https://github.com/gristlabs/grist-core
