# UI audit: rbt-office (`@rbt/office` 0.1.29 / master `52f54a15`) and redbat's host

2026-11-11. This is a read-only audit: nothing was edited in rbt-office or redbat-test. Its 71 screenshots were taken in a session scratch folder and are not kept in the repo:

**Historical snapshot:** use the [current capability matrix](continuation-roadmap.md#current-capability-matrix) for today's implementation/API/shell/platform status. Redbat and competitor observations here were not rerun during the correctness review; implementation follow-through is separate from the original screenshots.
- `redbat-*.png`: redbat's editors, from the adoption harness build (made from `redbat-office` at `8d62b0d9e`). Its `SpreadsheetEditor.tsx`, `WordEditor.tsx` and `SheetDialogs.tsx` are byte-identical to redbat-test's.
- `comp-*.png`: rbt-office's own browser UI from `dist/` (built 2026-11-10), through a scratch vite page: `WordShell.mount`, `createSpreadsheet` with `createSpreadsheetShell` and `createFormulaBar`, and the dialog specs through `OfficeUiDialog.renderDialog`.

**Implementation follow-through (2026-11-11):** the C5–C6 fixes now render all eleven existing field kinds and tabs in `OfficeUiDialog`, preserve tab edits, and provide accessible names plus keyboard/focus behavior. The renderer's seven Chromium dialog scenarios and real Word Cancel/undo workflow passed; see `systematic-review-2026-11-11.md` and the contract in `sharing-architecture.md`. This does not wire the unused advanced dialog specs into the browser shells or merge the parked command registry. The tables below remain the original audit snapshot.

**Formatting/name follow-through:** the browser shells now include Word font and paragraph borders/pattern-shading dialogs and Spreadsheet conditional-rule/scoped-name managers. The old five-kind conditional dialog was replaced, including its bare-widget entry point. The tables below remain the original audit, not current feature availability; see the roadmap and host API contracts for supported authoring and preservation limits.

**Sheet/list follow-through:** Spreadsheet now exposes sheet-tab operations, row/column sizing, autofit and hide/unhide; Word exposes list kind/marker/level, restart and continue. Actual desktop/390px workflows and targeted save/reopen regressions passed. Formatting-only Word edits retain displayed ordinals using standard numbering segments; reopening keeps those segments distinct. Native dialog parity and full Office/browser certification remain unverified.

**Function/style follow-through:** Spreadsheet provides a searchable 21-function insertion/help catalogue and guarded Tab completion; broader evaluator coverage remains open. Word provides a searchable existing paragraph/character styles pane, not style-definition authoring. Imported direct-versus-inherited typography, caret styles, undo, save/reopen and desktop/390px surfaces were exercised; detailed evidence is in the systematic review.

**Selection/page-setup follow-through:** Spreadsheet now shows cached live selection statistics with AutoFilter exclusions. Word's Page setup applies to the section containing the selection's focus, with preserved untouched geometry and imported section metadata. Desktop/390px controls, selective DOCX save/reopen, header positioning, undo and keyboard focus were exercised; the public legacy page-setup API still addresses the final section.

**Find/table-properties follow-through:** Spreadsheet's bare component and shell now share one per-grid search bar/session, including values/formulas, scope/range, case/whole-cell/regex, navigation and replacement. Word's table dialog now covers preferred width, alignment/indent, shading/borders, defaults and current-row minimum height/splitting/header behavior. Desktop/390px forms, keyboard search/replace/undo, selection guards and imported metadata preservation were exercised. Native Spreadsheet retains its separate dialog; Word exact-height and fixed/autofit layout behavior remain incomplete. See the systematic review for final build and round-trip evidence.

**Count/function-coverage follow-through:** Word now shows live body and active-selection word counts alongside action/error status, with explicit model/whitespace semantics and loaded-prefix labels. Spreadsheet's single function-help catalogue now covers 66 functions; repeated argument groups cycle through their range/criterion or test/result slots. Actual keyboard entry, undo, nested/foreign-story counts, dynamic-array and criteria formulas, XLSX reopen and desktop/390px controls were exercised. This is not complete evaluator help or Microsoft Word word-count parity; details and limits are in the host contracts and systematic review.

**Context/chart follow-through:** Word now has a shell-owned text/table context
menu with selection-safe right-click, caret-anchored keyboard entry and guarded
real clipboard actions. Spreadsheet's Insert Chart dialog authors column,
horizontal bar, line and single-series pie charts from bounded ranges, with
explicit header/category interpretation. Desktop/390px surfaces, undo,
clipboard permissions/stale state, actual chart download/reopen and imported
part preservation were exercised. Chart move/resize/delete and configuration
editing remain unimplemented; details and certification limits are in the
current roadmap, host contracts and systematic review.



## 0. The headline: where the UI lives

**redbat builds its own UI.** It uses none of rbt-office's chrome: no toolbar, menu, dialog, panel or spec.
- It imports only the engines: `@rbt/office/spreadsheet/Spreadsheet.js` (`createSpreadsheet`) in `mega/frontend/src/lib/spreadsheet/editor.ts:20`, and `@rbt/office/word/Word.js` (the `API` object) in `src/lib/word/editor.ts:11`.
- It drives them through their host APIs. Everything a user sees around the canvas is Solid code in redbat:
  - `src/components/office/OfficeEditorFrame.tsx`: the title, Save / Save and close, the dirty state, and the unsaved-changes guard.
  - `SpreadsheetEditor.tsx`: the toolbar, the Insert / Data / View menus, every context menu, the formula bar, the sheet tabs, the status bar and zoom.
  - `sheet/SheetDialogs.tsx`: the spreadsheet's 12 dialogs and panels.
  - `WordEditor.tsx`: the Word toolbar, its popovers, the find bar and the table tools.
- Nothing in redbat imports `WordShell`, `WordUi`, `WordDialogs`, `Toolbar`, `SpreadsheetUi`, `OfficeUiDialog` or `Rbt.Office.Ui`. (`grep -r @rbt/office mega/frontend/src` lists only the engine modules, the viewers, the PDF exporter and CSS.)

**rbt-office ships three separate, unconnected UI layers, and none of them is used by a real host:**
1. **Backend-neutral specs.**
   - `src/OfficeUi/Specs.fs` holds the types.
   - `src/Spreadsheet/SpreadsheetUi.fs` has 13 dialog specs.
   - `src/Word/WordUi.fs` has menus, toolbar rows and 12 dialog specs.
   - They are consumed only by the Wisp/native hosts and test runners: `spreadsheet-wisp/Program.fs` uses filter, insert table, macro, pivot, insert / rename sheet, open and save; `tools/excel-native-runner` and `tools/word-shell-runner` use them for render tests.
   - `formatCellsDialog`, `formatCellsAdvancedDialog`, `paragraphDialog`, `characterDialog`, both `preferencesDialog`s and both `aboutDialog`s are called by **nothing** that applies their result.
2. **The Fable DOM shells.**
   - `src/Word/WordShell.fs` with `WordHost.fs` and `WordDialogs.fs`: a menu bar, a toolbar, about 17 DOM dialogs, a comment layer and a bookmark / content-control navigator, mounted by `Shell.mount selector`.
   - `src/Spreadsheet/Toolbar.fs`: `createSpreadsheetShell` (a menu bar, one toolbar row, a find bar) and `createFormulaBar`.
   - Both are in `dist/`, and nothing in the repo or in redbat mounts them. The only pages that did (`src/Grid/test-*.html`) import from a path that no longer exists.
3. **Legacy DOM dialogs inside the spreadsheet component itself.**
   - `showFindReplaceDialog` (`Spreadsheet.fs:5393`) and the conditional-formatting dialog (`Spreadsheet.fs:5670`), with inline styles.
   - Ctrl+F in the grid still opens the first one (`Spreadsheet.fs:3266`), **in redbat too** (see gap 4).

The roadmap's open decision (*The initial release (alpha)*: whether the alpha ships the shell or only the canvas component) is therefore the deciding question.
- Today, every user-facing UI fix for redbat users goes into **redbat** (`SpreadsheetEditor.tsx`, `SheetDialogs.tsx`, `WordEditor.tsx`).
- A fix goes into **@rbt/office** only where the host API lacks the command: e.g. a Word paragraph-indent setter, chart insertion, a spreadsheet PDF.
- For the alpha's "complete UI in the browser component" (roadmap item 7), redbat's UI is the mature one and rbt-office's shells are prototypes. Two ways forward:
  - (a) bring a framework-free UI of redbat's quality into `@rbt/office`, which would replace `WordShell` / `Toolbar.fs` / `OfficeUiDialog`;
  - (b) declare the host-builds-the-UI model and ship redbat's UI as the reference host / demo.

## 1. Spreadsheet

Legend:
- **Def**: defined in rbt-office.
- **Wired**: connected to a working action (engine API or shell).
- **Comp**: reachable by a user of `@rbt/office` alone, meaning the component plus its shipped `SpreadsheetShell`.
- **Redbat**: reachable in redbat's sheet editor.
- y / n / part.

| Item | Def | Wired | Comp | Redbat | Where (file:line) and notes |
|---|---|---|---|---|---|
| **Format cells dialog** | part | **n** | **n** | **n** | **Simple spec:** `SpreadsheetUi.fs:672` `formatCellsDialog` (number format select, hex text and fill colours, a sample line). It works when rendered (`comp-dialog-formatCells.png`) but is called by nothing.<br>**Tabbed spec:** `SpreadsheetUi.fs:800` `formatCellsAdvancedDialog` (Number / Alignment / Font / Border / Fill, initial values hard-coded, result unapplied). It renders **empty**: `OfficeUiDialog.renderDialog` ignores `Tabs` (`OfficeUiDialog.fs:153-176` renders only `spec.Fields`). See `comp-dialog-formatCellsAdvanced.png`.<br>**Redbat:** no "Format cells…" anywhere (toolbar, menus, cell context menu `SpreadsheetEditor.tsx:520-551`). Ctrl+1 does nothing (checked). |
| Number formats (toolbar) | y | y | part | part | **Redbat:** 5 presets, General / Number / Currency ($) / Percent / Date, plus a hidden "Custom" label (`SpreadsheetEditor.tsx:55-61, 1060-1072`). No custom code entry, no date or time variants, no accounting, fraction, scientific or text, no preview.<br>**Shell:** "$" and "%" buttons send `formatCurrency` / `formatPercent`, which `ApplyFormat` rejects: `ApplyFormat: unknown command "formatCurrency"` (checked in the browser; `Toolbar.fs:350`, `Spreadsheet.fs:8398`). |
| Decimal places +/- | y | y | y | y | `IncreaseDecimals` / `DecreaseDecimals`; redbat `SpreadsheetEditor.tsx:1073-1078`. |
| Font family / size | y | y | part | y | Redbat selects (`:916-948`). With a mixed selection the size box shows a truncated "Mi" (`redbat-sheet-06-context-tab.png`). The shell has only A+ / A-. |
| Bold / italic / underline | y | y | y | y (buttons) | Strikethrough and double underline exist in the API (`ApplyFormat "strikethrough"`) but not in redbat's UI. |
| Ctrl+B / Ctrl+I / Ctrl+U in the grid | **n** | **n** | **n** | **n** | The component's key handler has no b / i / u case (`Spreadsheet.fs:3258-3280`). Redbat doesn't bind them either, yet its tooltips say "Bold (Ctrl+B)" (`SpreadsheetEditor.tsx:951`), and the shell's Format menu shows Ctrl+B. Checked: the selection's bold is unchanged after Ctrl+B. |
| Text colour / fill | y | y | part | part | Redbat: 10 preset swatches each, no custom colour or theme palette (`:72-73`). The shell's "A" / "Fill" buttons send `textColor` / `bgColor`, which are not `ApplyFormat` commands. |
| Borders | y | y | n | part | Redbat palette: 8 edge presets × 6 line styles, **no border colour**, no diagonal (`:79-105`, `redbat-sheet-03-palette-border.png`). API `ApplyBorder`. |
| Alignment, vertical alignment, wrap | y | y | part | y | Redbat `:1023-1046`. |
| Indent, rotation, shrink to fit | y (API) | y | n | **n** | `SetIndent`, `SetTextRotation`, `ApplyFormat "shrinkToFit"` exist; no UI anywhere. |
| Merge / unmerge | y | y | n | y | Redbat toggle button `:1048`. Merge across / merge-and-centre variants: none. |
| Cell protection, sheet protection | **n** | n | n | n | No API (no `Protect*` member), no UI. |
| Clear contents / formats / all | y | y | n | y | Redbat clear palette `:1080-1088`. |
| Format painter | y | y | n | y | Redbat `:905`. |
| Conditional formatting rules manager | y | y | part | **y** | **Redbat:** `SheetDialogs.tsx:1066` `ConditionalFormats`: list, new / edit / delete / up / down, rule types, icon sets, samples (`redbat-sheet-08-…png`).<br>**Component:** legacy `ShowConditionalFormattingDialog` (`Spreadsheet.fs:5670`): one rule, greater / less / between / equal / contains, inline-styled blue / yellow / grey buttons (`comp-sheet-04-cf-dialog.png`). |
| Data validation | y | y | n | y | `SheetDialogs.tsx:1469`; API `SetValidation`, in-cell list dropdown, circle invalid. |
| Sort dialog | y | y | part | y | `SheetDialogs.tsx:454` (keys, order, header row); quick A→Z / Z→A in the Data menu and context menu. |
| Filter UI | y | y | part | y | **Redbat:** `FilterPanel` (`SheetDialogs.tsx:561`): sort, search, value list, condition, clear, OK (`redbat-sheet-18-filter-panel.png`).<br>**Bare component:** draws a ▼ in **every** column header even without a filter (`comp-sheet-01-shell.png`). That is the old "filter dropdown" path (`Spreadsheet.fs:5166`). |
| Insert chart / edit chart | **n** | n | n | n | Charts are read and drawn only (`Charts.fs`); no API to insert or edit. Alpha decision: to close. |
| Insert picture | y | y | n | y | Redbat Insert ▸ Picture (`:1104`), picture context menu (delete). |
| Named ranges | y | y | n | y | `NamesDialog` `SheetDialogs.tsx:1590`; the name box goes to names. |
| Tables (ListObjects) | y | y | n | y | `TablesDialog` `:1822` with a style picker. |
| Pivot table | part | part | n | n | `SpreadsheetUi.pivotTableDialog` (`:609`) is wired only in `spreadsheet-wisp`. Redbat: none. |
| Find / replace | y | y | y (legacy) | y, **but** | Redbat `FindPanel` (`SheetDialogs.tsx:1983`): match case, whole cell, all sheets, in formulas; opened by the toolbar search button.<br>**Ctrl+F in the grid opens the component's legacy modal "Find and Replace"** instead (`Spreadsheet.fs:3266`). It has a different look (blue / green / yellow / grey buttons) and sits over redbat's panel if that is open. See `redbat-sheet-16-ctrl-f.png` and `redbat-sheet-20-ctrl-f-fresh.png`; checked on a fresh page. |
| Go to (name box) | y | y | y | y | Redbat name box `:1187`. |
| Freeze panes | y | y | n | y | View menu: at selection, top row, first column, unfreeze (`:1131-1134`). |
| Sheet tabs menu | y | y | n | y | **Redbat:** insert, duplicate, rename…, delete, move left / right, hide / unhide, **6 fixed tab colours** (`:593-602`, `redbat-sheet-06-context-tab.png`). Double-click renames; + adds a sheet. No drag to reorder, no custom colour, no "move to position".<br>**Spec dialogs:** `insertSheetDialog` / `renameSheetDialog` are wisp-only. |
| Row / column size, hide / unhide, auto-fit | y | y | n | y | Header context menus (`:557-583`, `redbat-sheet-05-context-row-header.png`); `LineSize` dialog `SheetDialogs.tsx:165`. Drag-resize is in the component. |
| Insert / delete rows and columns | y | y | y (menu) | y | Context menus and the Insert menu. Insert cells with shift right / down: none. |
| Group / outline | y | y | n | y | Data menu, View menu. |
| Insert function / formula help | **n** | n | n | n | No function list, no autocomplete in the edit box or formula bar, no argument hints (`Spreadsheet.fs:6164` sets `autocomplete=off`; nothing else). |
| Comments | y | y | n | y | Redbat `CommentDialog` (`SheetDialogs.tsx:240`) with threads; the context menu has new / edit / delete comment. No comments pane listing them all. |
| Hyperlinks | y | y | n | y | `HyperlinkDialog` `:392`; context menu. |
| Page setup | y | y | n | y | Redbat `PageSetupDialog` (`:746`): orientation, paper, scale / fit, titles, print area, grid lines, headings, centring, page order, page breaks, printed-page list (`redbat-sheet-12-…png`). Header / footer text: not editable there (`raw` read only). Spec `pageSetupDialog` (`SpreadsheetUi.fs:710`) is unused. |
| Print / PDF (corrected) | y | API | host callback | not rechecked | Browser `ExportPdf` / `PrintPdf` and the packaged writer exist. The shell's menu delegates `printPdf` to the host; source availability is not proof of a wired or verified end-user print workflow. See the current matrix. |
| Zoom | y | y | n | y | Status bar −/%/+ (`:1389-1397`). |
| Status bar sum / average / count | y | y | n | y | `GetStatus` → Average, Count, Sum (`:1378-1386`). No min / max / numerical count choice. The average shows raw decimals ("12,923,077.0754"). |
| View options | y | y | n | y | Grid lines, headers, formulas, page break preview (View menu). |
| Undo / redo | y | y | y | y | Toolbar, Ctrl+Z / Ctrl+Y. |
| Cut / copy / paste | y | y | y | y | Paste special (values / formats): none. |
| Fill down / right, autofill | y | y | y | y | Ctrl+D / Ctrl+R, Insert menu; fill handle. |
| Cell context menu | API only | — | n | y | The component has no menu of its own (`Spreadsheet.fs:605-612`, `OnContextMenu`). Without a host, right-click shows the browser's menu. Redbat: cut, copy, paste, insert / delete rows and columns, clear, sort, filter, comment, link (`redbat-sheet-04-context-cell.png`). **No "Format cells…"**. |
| Keyboard: arrows, Ctrl+arrows, Home / End, PgUp / PgDn, F2, Enter / Tab, Del, Ctrl+A, Alt+Down, Shift+F10 | y | y | y | y | `Spreadsheet.fs:3260-3340`. Missing: Ctrl+B / I / U, Ctrl+1, Ctrl+; (date), Ctrl+Shift+$ / % (formats), Ctrl+Space / Shift+Space (select column / row), F4 (absolute reference). |
| Macro dialog | part | part | n | n | `macroDialog` is wisp-only (out of the alpha). |

## 2. Word

| Item | Def | Wired | Comp (WordShell) | Redbat | Where (file:line) and notes |
|---|---|---|---|---|---|
| **Font dialog** | part | **n** | **n** | **n** | `WordUi.characterDialog` (`WordUi.fs:780`) is unused. Rendering it **throws** `Match failure: Rbt.Office.Ui.DialogFieldKind`, because `OfficeUiDialog.renderField` (`OfficeUiDialog.fs:60`) handles only text / number / checkbox / select / status / computed, not section, colour, slider, text area or radio (`comp-dialog-character.png` is blank).<br>Redbat: the toolbar has font, size, B / I / U, colour, highlight and clear (`WordEditor.tsx:791-877`). No strikethrough, superscript / subscript, caps, character spacing or underline styles in the UI; the engine has all of them (`toggleStrike`, `setVertAlignFor`, `toggleAllCaps` / `SmallCaps`, keys Ctrl+= / Ctrl+Shift+=). |
| **Paragraph dialog** | part | part | y (as "Paragraph style…") | **n** | `WordUi.paragraphDialog` (`:722`): unused, and it throws the same match failure.<br>WordShell's own `WordDialogs.paragraphStyleDialog` (`WordDialogs.fs:341`): alignment, before / after, line spacing, background, left / right indent, special indent, size. It works but is unstyled (`comp-word-03-dialog-paragraph-style.png`).<br>Redbat: a spacing popover only, with line spacing presets and before / after pt (`redbat-word-02-menu-spacing.png`). No indents (first-line, hanging, left / right), no keep-with-next / page-break-before, no tabs. The API has no direct indent setter, only `increaseIndent` / `decreaseIndent`. |
| Styles gallery / pane | part | y | part | part | Redbat: a style `<select>` (`:791`) listing the document's styles. No gallery with previews, no styles pane, no modify / new style. WordShell: a style select plus Format ▸ Heading 1-3. |
| Lists and numbering | y | y | y | part | Redbat: bullet / numbered toggles, list level ± (`:973-984`). No bullet or number style choice, no restart / continue numbering, no multilevel lists. API `setNumberFormat`, `setListLevel` exist. |
| Borders and shading (paragraph) | **n** | n | part | n | The paragraph style dialog has a "Background" colour; there are no paragraph borders in the API. |
| Table: insert | y | y | y | y | Redbat grid popover with rows × columns (`:988-1003`); WordShell Table menu and Insert table… dialog. |
| Table: rows / columns / merge / split | y | y | part | y | Redbat contextual tools appear in a table (`:1100-1128`, `redbat-word-07-table-tools.png`). |
| Table properties (borders, shading, widths, alignment, cell margins) | part | part | part | **n** | WordShell has "Cell style…" (`WordDialogs.tableCellStyleDialog`) and "Cell margins…". The API has `setSelectedTableCellShading`, `setCurrentTableCellMargins`. Redbat: none. |
| Insert picture | y | y | y | y | Redbat image button (`:1005`); resize by handles. Alt text, wrap or position UI: none. |
| Insert link | y | y | y | y | Redbat link popover and a link bubble (edit / remove) (`:1021`, `:1311`). Ctrl+K: **nothing** in redbat (checked). |
| Insert symbol / special character | **n** | n | n | n | — |
| Page break | y | y | n | y | Redbat header-footer menu ▸ Page break, Ctrl+Enter. |
| Section break | y | y | y | **n** | WordShell Insert ▸ Section break (next page only); API `setSectionBreak`. |
| Header / footer | y | y | y | y | Redbat: Header, Footer, link to previous, first / odd-even shown (`:1064-1088`, `:1150-1192`). |
| Page numbers | y | y | y | y | Redbat inserts a PAGE field. No format, position or "Page X of Y" choice. |
| Table of contents | **n** | n | n | n | No insertion API (TOC fields are kept and refreshed only). |
| Footnote / endnote | y | y | y | **n** | WordShell Insert ▸ Footnote / endnote… (annotation dialog); API `setNoteFor`, `enterNote`. Redbat: none. |
| Text box | y | y | y | y | Redbat header-footer menu ▸ Text box. |
| Caption, cross-reference, bookmark, field, equation, content control | y | y | y | **n** | WordShell has dialogs for all of them (`WordDialogs.fs:605-719`) and a navigator for bookmarks and content controls. Redbat: none. |
| Page setup | y | y | y | **n** | `WordDialogs.pageSetupDialog` (`:749`) works in WordShell: size, margins, header / footer distance, different first / odd-even (`comp-word-03-dialog-page-setup.png`). `WordUi.pageSetupDialog` (orientation + margins). API `setPageSetup`. **Redbat has no page setup**: no margins, orientation or size. |
| Find / replace | y | y | y | y, **but** | Redbat find bar (`:1205-1245`): find next, replace, replace all, match case. **Ctrl+F does not open it** (checked: nothing happens); only the toolbar button does. WordShell: Edit ▸ Find and replace. |
| Track changes | y | y | y (toolbar spec) | **n** | API `toggleTrackChanges`, `accept/rejectSelectedRevision`, `accept/rejectAllRevisions`, `selectNext/PreviousRevision`, `getRevisionInfos`. `WordUi.toolbar` has the buttons, but WordShell's visible `toolbarRows` don't include them (only the full `toolbar` list does). Redbat: nothing. Revisions in a file are drawn, but can't be turned on, accepted or rejected. |
| Comments / review pane | y | y | part | **n** | WordShell: Insert ▸ Comment…, a comment layer with connectors (`word-comment-layer`). There is no review pane listing comments or revisions anywhere. Replies: none (the roadmap's gap). Redbat: no way to add, see the list of, or resolve comments. |
| Navigation pane (headings outline) | n | n | n | n | WordShell's navigator lists bookmarks and content controls only. |
| Zoom | y | y | n | y | Redbat toolbar −/%/+ (`:1133-1149`). |
| Word count / status bar | API only | — | n | **n** | `computeRangeStatistics` exists. Redbat shows only "N pages" under the file name. WordShell's status line is a message line. |
| Spell check | n | n | n | n | The input proxy sets `spellcheck="false"` (`src/Word/Input.fs:1893`). Coming on `word/spell-check`. |
| Print / PDF | y | y | y | **n** | WordShell File ▸ Print PDF goes through `window.RedBatWordPdfExport`. **Redbat registers that exporter (`src/index.tsx:36-39`) but its Word editor has no Print / PDF button.** |
| Right-click / context menu | n | — | n | **n** | No context menu at all. A right-click on the canvas does nothing in redbat (checked; `redbat-word-04-right-click.png`). Cut / copy / paste / link / table actions on right-click: none. |
| Keyboard | y | y | y | y | Engine (`src/Word/Input.fs:1575-1700`): Ctrl+B / I / U, Ctrl+1/2/3 (headings), Ctrl+E / L / R / J, Ctrl+Shift+L / N, Ctrl+M, Ctrl+[ / ], Ctrl+= / Shift+=, Ctrl+Space, Shift+F3, Ctrl+Enter, Shift+Enter, Tab in tables, clipboard, undo / redo. Missing in redbat: Ctrl+F, Ctrl+H, Ctrl+K, Ctrl+P. |

## 3. Missing or broken items, by how fast a first-time user hits them

**The first minute** (formatting a cell or text, a right-click, the usual shortcuts):
1. **Spreadsheet: no Format Cells.** There is no number format beyond 5 presets: no custom code, no date or time styles, no preview. There is no strikethrough, rotation, indent, shrink, border colour or protection. Ctrl+1 does nothing and there is no context-menu entry. The engine already has `SetNumberFormat`, `SetTextRotation`, `SetIndent`, `ApplyFormat "strikethrough"/"shrinkToFit"` and `ApplyBorder`; only the UI is missing.
2. **Spreadsheet: Ctrl+B / Ctrl+I / Ctrl+U do nothing in the grid**, while the tooltips advertise them.
3. **Word: right-click does nothing** (no context menu at all).
4. **Spreadsheet Ctrl+F opens the component's old, differently styled modal Find dialog** instead of redbat's find panel. **Word Ctrl+F does nothing.** Ctrl+K (link) and Ctrl+P (print) do nothing in Word.
5. **Word: no Font dialog and no strikethrough, superscript / subscript or caps buttons.** Only B / I / U are on the toolbar.

**The first document or workbook of a real task:**

6. **Word: no paragraph indents** (first-line, hanging, left / right) and no paragraph dialog. There is spacing only.
7. **Word: no page setup** (margins, orientation, paper size) in redbat.
8. **Print / PDF shell reachability:** the original redbat audit found no button in either editor. Correction: the spreadsheet does have browser PDF/print APIs; current redbat wiring was not rechecked.
9. **Spreadsheet: charts cannot be inserted or edited.**
10. **Word: comments and tracked changes can't be added, listed, accepted or rejected.** There is no review pane; the engine API has all of it.
11. **Word: lists can't change bullet or number style or restart numbering.** No table properties (borders, shading, widths) and no paragraph borders.
12. **Word: no status bar** (word count, page X of Y).
13. **Spreadsheet: no insert function, autocomplete or argument help.**
14. **Word: no footnote, section break, TOC, symbol, caption or bookmark in redbat.** The engine and WordShell have most of them (no TOC, no symbol).
15. **Spell check in Word** (branch `word/spell-check`).

**Later:**

16. A styles pane or gallery, and modify-style. A navigation pane.
17. Spreadsheet: sheet protection; drag to reorder tabs, custom tab or font colours; paste special; insert cells with shift; pivot creation; the status bar's choice of statistics and its number formatting ("12,923,077.0754").
18. Word: picture wrap, position and alt text UI; page number formats.

**@rbt/office's own UI is broken in ways only an embedder sees.** This matters if the alpha ships the shell.
- The Fable `OfficeUiDialog` renderer ignores `Tabs`: Format Cells (advanced) and Preferences render empty.
- It throws `Match failure` on section, colour, slider, text-area and radio fields, so `paragraphDialog` and `characterDialog` crash.
- `SpreadsheetShell`'s actions (`formatCurrency`, `formatPercent`, `formatNumber`, `textColor`, `bgColor`, `printPdf`, …) match nothing in the component (`ApplyFormat` throws on them), and its buttons are text labels ("B", "A+", ".0+", "Fill").
- `WordShell` ships **no CSS at all**: unstyled HTML buttons stacked vertically, dialogs appended inline at the page bottom (`comp-word-01-shell.png`, `comp-word-03-*.png`).
- `WordShell`'s visible toolbar row leaves out the track-changes and many insert buttons that `WordUi.toolbar` defines.
- The bare spreadsheet component paints a ▼ filter arrow on every column header.
- `word-wisp`'s actions are almost all `ignore` (`word-wisp/Program.fs:69-89`). Native hosts are out of the alpha.

## 4. Current visual design

**Redbat (what users see).** A consistent, quiet, modern look, shared by both editors through `OfficeEditorFrame` and the `rb-office-*` CSS classes.
- **Layout:** a white modal frame with the file name and "No changes / N pages", Save / Save and close at the top right, then one 40 px toolbar row.
- **Toolbar:** monochrome stroke icons (Lucide-like, about 16 px) on ghost buttons, with a light grey hover and a pressed tint for active toggles. Thin vertical separators group the buttons. Native `<select>`s are used for style, font, size and number format.
- **Menus:** text with a ⌄ caret ("Insert ⌄ Data ⌄ View ⌄"), opening rounded white popovers with a soft shadow, separators, right-aligned shortcut hints and check marks.
- **Dialogs:** centred white cards with a 12 px radius on a dimmed backdrop. A title, a muted one-line explanation, compact labelled fields, and a black primary button with ghost secondaries (`redbat-sheet-08/12-*.png`). The filter panel and find panel are anchored popovers in the same style.
- **The two editors are consistent with each other:** the same frame, icon set, button styles, popover style and zoom control. The Word toolbar is a subset of the sheet's pattern; Word has no menus, only icon buttons and popovers, so it has fewer entry points.
- **What breaks the consistency:**
  - the spreadsheet component's legacy Find (Ctrl+F) and conditional-format dialogs: bright blue / green / yellow / grey full-width buttons and inline styles (`redbat-sheet-16-ctrl-f.png`, `comp-sheet-04-cf-dialog.png`);
  - the size select truncating "Mixed" to "Mi";
  - the sheet has menus while Word has none, so the same function, Find, is a menu-less icon in both, but Insert is a menu in the sheet and scattered icons in Word.

**@rbt/office's own shells.**
- **SpreadsheetShell:** a flat Google-Docs-like menu bar (File / Edit / Insert / Format / Data) and a toolbar of **text** buttons ("B I U | A+ A- | A Fill | L C R | $ % .0+ .0- | Cut Copy Paste | Undo Redo"), white with hairline borders (`comp-sheet-01-shell.png`). It is usable-looking but sparse and unwired.
- **WordShell:** unstyled browser buttons, one per line (`comp-word-01-shell.png`). The CSS classes exist (`word-toolbar`, `rb-word-tool-button…`) but no stylesheet is shipped.
- **Spec dialogs (`OfficeUiDialog.defaultStylesheet`):** a clean white card, an 8 px radius, a blue (#1A73E8) primary button and Material-like inputs (`comp-dialog-formatCells.png`). They are close in spirit to redbat but a different accent (blue rather than black) and radius.
- The `UI-FEATURES-WIP.md` and `src/Grid/*.md` documents describe an older inline-DOM design and are stale.
