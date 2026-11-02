# Continuation roadmap

Written 2026-10-12, when work was halted to move to another machine; next steps revised 2026-11-02. Picks up the fix / release / adopt loop on rbt-office (Word + Spreadsheet) and its adoption in redbat. Everything needed is on the remotes; nothing below depends on the old machine's temp directories.

## Branches (all on `origin` of rbt-office)

| Branch | What it is | State |
|---|---|---|
| `master`, tag `v0.1.22` | Last release, published as `@rbt/office` 0.1.22: 0.1.21 + saves Excel refused after a row or column insert (four causes, below), `xml:space` inherited, no-break and en spaces kept. **0.1.21 and earlier can write such files.** (`v0.1.17` published nothing) |
| `review-fixes` | The integration branch: what the next release is gated and cut from. Equal to `master` right after a release | nothing unreleased |
| `fix/sheet19` | Spreadsheet pass nine as it was handed off | Its item 1 is in `review-fixes`, rebased, without the handoff folder. Items 2-7 are not started: `tools/handoff-sheet19/HANDOFF.md` on `origin/fix/sheet19` keeps their findings, and `bigsheet.mjs`, `render.fsx` and `shoot.mjs` beside it |

`origin/fix/word19` is the pass as it was handed off (its `tools/handoff-word19/HANDOFF.md` has the host notes); the pass itself is in 0.1.18, rebased, without the handoff folder. The `tools/handoff-*` folders are temporary: delete them (or drop those commits) before merging a branch into `review-fixes`.

redbat (`cardillion-mega` remote): `master` has the 0.1.22 adoption (harness 1104/1104 with `word-new18`, `sheet-new18`; the host no longer draws its own Word shape frame, nor handles header text box double clicks or Ctrl+Enter; the name box goes where GoTo goes; View has a page break preview). On a new machine branch from `master`.

## Next steps, in order

1. **Saves that Word and Excel take as they should.** Fix what saves written from the model lose (the spaces between runs, drop caps, bound content controls: see Cross-cutting); make `opaque-objects`, `preserved-parts`, `reader-structure` valid documents. Keep the workbooks of the spreadsheet harness and open them in Excel as `tools/word-parity/open-all.vbs` does for Word. Compare formatting too, not only text length and pages.
2. **Real Word and Excel as a routine check.** A single-sheet Excel export so that the big dashboards can be compared; a scripted comparison of rendering, pages and positions over a fixed set of documents and workbooks, run before a release; fix what it finds (scatter axis limits first). Make the six font-dependent `tools/word-browser` checks independent of the machine's fonts.
3. **Macros (VBA) worth the name, checked against Excel and Word.** See *Macros* in the backlog for where the engine stands. In order:
   - **Excel's values as the formula oracle:** an `open-all`-style VBS that recalculates every corpus workbook fully (`CalculateFull`) and dumps every cell's `Value2` and `Text`; compare with our engine's values per cell. Cheapest, covers the most.
   - **Macros side by side:** a corpus of `.bas` modules run on the same workbook in real Excel and in our engine (`Vba.runOnModel`), every resulting cell compared. The first modules exercise what real macros use: loops over `Range`/`Cells`, `Worksheets("X")`, `Format`, dates, `IIf`, `Err`, `Scripting.Dictionary`, `WorksheetFunction`, formatting through `Font`/`Interior`, row inserts.
   - **Fill what they show missing,** in `Rbt.Vb` (publish, bump the pin) and in `SpreadsheetVba.fs`: the built-ins, then a writable object model over all sheets (`Worksheets`, `Font`, `Interior`, `NumberFormat`, `Rows/Columns.Insert/Delete`, `Copy`), then events (`Worksheet_SelectionChange`, `Workbook_BeforeSave`) with errors reported instead of swallowed.
   - **Macros in the browser component:** read the project from an opened `.xlsm` (the browser keeps `vbaProject.bin` but sets `Macros = None`), and give the host API `GetMacros`/`RunMacro`/events, documented in `docs/spreadsheet-host-api.md`; redbat decides whether to show them.
   - **Real `.xlsm` / `.docm`:** fix `tools/vba/validate-real-xlsm.sh` (old project names, `wordvm`) for `winvm`; check that macros Excel wrote survive our save and still run in Excel, and that a project we write opens and runs there.
   - **Word:** the object model only answers layout queries; a writable `Range` (`Text`, `InsertAfter`, `Font`, `Paragraphs.Add`) and `.docm` loading, then the same side-by-side check against Word. Meanwhile run `Range.Information` / `ComputeStatistics` over the round-trip set, comparing line and page positions, not only page counts.
4. **Hosts never run:** run the native hosts interactively, and Firefox, and list what breaks.
5. **Missing features, by what real files use:** Word flips, `allowOverlap`, text in group shapes and drawing canvases; spreadsheet chartEx (waterfall, funnel, maps), slicers, custom table styles (items 3 and 6 of the `fix/sheet19` handoff).
6. **Long sheets and source patching:** row and column inserts on long sheets (~1 s at 846,000 cells), sort there unmeasured; more edits patched into the source instead of written from the model (item 7).
7. Continue with the backlog.

## Backlog

### Spreadsheet — bugs in 0.1.16 as released
- On the 846,000-cell sheet: insert or delete a row or insert a column takes ~0.9-1 s. Sort there is unmeasured (`bigsheet.mjs` used a wrong key name).
- "Products sold by Quantity" in the e-commerce dashboard still wraps under its doughnut: the box is narrower in its scaled group. Whether Excel wraps it too is unchecked.

### Spreadsheet — features
- Charts and shapes print through `Paint.drawingOverlays` (`tools/spreadsheet-pdf-runner`). Each page holds what reaches it, cut at the page's edge, as Excel does (chart-zoo compared). Excel's scatter axis stops at 6 where ours stops at 5.5. The portable canvas still draws frames; native and wisp painters do not handle `VecPath`.
- chartEx waterfall/funnel/maps and others; slicers and timelines. Radar, bubble, stock and 3-D-as-flat are drawn; bubble and stock have been looked at on synthetic charts only, no workbook in the corpus has them.
- Shapes: custom geometry, preset polygons, rotation and flips are checked on synthetic shapes only (the corpus has `rect` and `roundRect`, no `rot` or flips). A group's own rotation is ignored. Unwritten: `normAutofit`, `spcBef/spcAft/lnSpc`, `vert`.
- Print and PDF use the sheet's default row and column sizes. Excel agrees on pages and scale for the lyric-multifamily sheets. Preview: grey outside the print area, frames over frozen panes.
- Custom table styles; built-in table style palettes.
- Source patching: row insert (2 of 110 corpus workbooks patched in place, 105 written from the model), table add (29 / 80 / 1), note add (22 from the model); 24 workbooks never patch.
- The last rows are unreachable above ~125% zoom in Chromium, and in Firefox.
- A formula's result-format inheritance follows my reading of Excel's rules, not a check against Excel.

### Word
- Not started in pass nineteen: flips (`flipH`/`flipV` are not read at all; file/line pointers in the handoff), `allowOverlap` in layout, text in group shapes and drawing canvases.
- Layout: a list number wider than its hanging indent overlaps the text; `w:lnNumType`; exact line heights stored as a multiple of the largest text size.
- Save: headers, footers and notes written from the model balance range marks but do not carry the ones the model lacks; comment/permission ranges that start mid-paragraph widen to the paragraph; line and group VML fallbacks are still dropped on a move.
- Tables: a row or column insert writes many paragraphs from the model (547 / 1552 across the round-trip pass).
- View anchor: `setAllowLinkedImages` still drops the staged layout; a measurer change far down a fully laid out document lays everything out at once with no anchor.
- Shape selection gaps: keys act on a shape only while no story is edited; the cursor over a handle ignores rotation; a text box entered by enterTextBox or a double click (a header's) is not selected - the host selects one it inserted itself.
- A header's text box with more text than its height (resume-template, typing at its end) does not show the overflow; unchecked against Word.
- The `■` third-level bullet in `test-lists.docx` is the fallback face's glyph, not an engine fault; mapping U+25A0 to U+25AA is an option, unverified against Word.
- Never exercised: bold typed into a multi-run field result.

### Macros (VBA), as of 0.1.22
- The language, `Rbt.Vb` 0.1.2 (Forge; ~4,400 lines: lexer, parser, bytecode compiler, stack VM, bytecode save/load): Sub/Function/Property, classes, Select Case, For/For Each/Do/While, With, On Error, GoTo/GoSub, ReDim Preserve, ByRef, Optional, ParamArray. Enum, Type, Declare, Implements and events parse; whether they run is unchecked. About 70 built-ins; missing `Format`, `IIf`, `Val`, every date function, `Debug.Print`, `Scripting.Dictionary`, `WorksheetFunction`.
- Excel (`src/Spreadsheet/SpreadsheetVba.fs`): a read-only `Range`/`Worksheet`/`Application` for layout queries (the 78-fixture layout harness); a writable one that sets only `Value`/`Value2`/`Formula` of the active sheet. `Workbook_Open` and `Worksheet_Change` fire by sub name; a macro's error leaves the model as it was and is not reported. `VbaProject.fs` reads and writes the module source of `vbaProject.bin` (source only, no p-code). Macros run only in `spreadsheet-wisp` (Run Macro dialog, `Worksheet_Change`); the browser component runs none and the host API has no call for them (`runForJs` is unused).
- Word (`src/Word/WordVba.fs`): read-only layout queries (`Range.Information`, `ComputeStatistics`, `Bounds`, `Words`/`Characters`/`Paragraphs`, `Selection`); no edits, no `.docm` macros, no events.
- Tests: `stage19-vba` (~20 checks), `stage20-vbabin`. `tools/vba/validate-real-xlsm.sh` is stale.

### Cross-cutting
- `tools/word-browser` depends on the machine's fonts: viewport, caret-movement, progressive-layout, progressive-read, drawing-order and mass-format have thresholds tuned on DejaVu and fail with `fonts-liberation` installed (Google Chrome's package pulls it in). Until they are made font-independent, run them with a `FONTCONFIG_FILE` that rejects `*/liberation/*`.
- **Workbooks in real Excel** (`run.mjs --keep DIR`, `tools/excel-parity/open-all.vbs`): 109 workbooks, saved as they were, with a typed cell, and with a row and a column inserted. Before 0.1.22 nine real workbooks' inserted saves were refused: a column moved past 16384, a shared formula reading another sheet, `#CYCLE` written as a cell error, a table with two columns of one name and a query table's fields not following its columns. Those rules are now in `packageProblems`, run on every export. On 0.1.22 Excel refuses none of the three saves of a workbook whose original it opens. `run.mjs --keep` still reports, of those saves: pivot parts left without an owner (`xl/pivotTables/pivotTable1.xml`, its cache: find which workbook, and whether the pivot table is lost) and a drawing relationship to `#Sheet1!A1` that `packageProblems` takes for a part. Still open: seven `synthetic-*` originals Excel refuses (drawing-extras, dynamic-arrays, fmt5, formatting, lambda, modern-functions, sparklines); `phpsp-32chartreadwrite` differs from its original in the sum of numbers in the plain save (100467 to 105537), `synthetic-groupby` gains formulas (2 to 26), `conditional-formatting-tutorial` loses 2 characters on insert. `CELL("width",A:A)` in column A is a cycle here, not in Excel. A row break's `max` does not follow an inserted column.
- Inline content controls are flattened in saves written from the model (`w:sdt` in a paragraph: `docx4j-binding-simple`, `docx4j-invoice*`), so Word no longer shows their bound values; the page counts of `docx-5mb`, `docx4j-repeating-section`, `file-sample` differ in those saves.
- **Round-trip saves in real Word** (`tools/word-parity/open-all.vbs`, 0.1.21, Word 16.0.19127, 381 files): every `.pass1` and `.pass2` save of a document Word can open opens with the original's page count and text length (180 saves, of 90 documents). Three synthetic fixtures cannot be opened even as originals (`opaque-objects`: error 5121; `preserved-parts`, `reader-structure`: "corrupted"), so they show nothing about saves, and `preserved-parts-edited` is rejected the same way; make them valid. Of the saves written from the model alone, 8 differ in text and 3 in pages: `docx4j-getting-started` -53 characters, `resume-template` -119, `docx4j-invoice-2013` -63, `docx4j-binding-simple` -12, `docx4j-hyperlink-binding` +12, `docx4j-legacy-forms` -5, `docx4j-invoice` and `docx4j-opendope-invoice` -1; pages `docx-5mb` 4 to 3, `docx4j-repeating-section` 8 to 7, `file-sample` 2 to 3. These are the known model-save faults below. Only text length and pages are compared; formatting is not. Workbooks went through Excel the same way on 0.1.22 (above).
- **LibreOffice is the one outside reader of the round-trip saves.** Real Word and Excel exist for this: the Windows VM, now `winvm` in the SSH config (`WORD_VM=winvm WORD_VM_HOST=winvm WORD_VM_DIR=C:/Users/olivia/RedBatExcelParity`; Excel 16.0.19127; `excel-print-pdf.sh` exports every sheet, so the big dashboards time out; it was formerly `wordvm`, `docs/word-parity.md`, `docs/excel-parity.md`, `tools/word-parity`, `tools/excel-parity`) gave the earlier parity references. Excel's PDF of chart-zoo has been compared with ours. `tools/libreoffice-check/check.mjs DIR` opens the documents `roundtrip.mjs --out DIR` keeps and compares the text LibreOffice reads from each save with the original's. On 0.1.18: 263 saves read as their originals, none failed. It is not part of the round trip yet (run it after one; write `--out` outside `/tmp`, the 381 documents are 75 MB), it compares text only, and the spreadsheet harness keeps no workbooks for it.
- What it found in the 13 saves written from the model alone (`.model.docx`, listed, not counted as failures): a space lost between runs (`atdeploymenttime`, `gravida,pellentesque`, `next.When` in docx4j-getting-started, resume-template, file-sample); a drop cap read as a separate word (the calibre demo and its four copies); the content of bound content controls replaced by their placeholder or value (`docx4j-binding-simple`, `-hyperlink-binding`, `-invoice-2013`); `mc:AlternateContent` written as its fallback (`reader-structure`). LibreOffice hangs on `test-complex-tables.docx` itself, writing text without end (3 GB in 25 s; the check kills the run by its process group and by what it has written), and does not open `docx4j-unmarshall-template.docx`, the originals.
- Native hosts are built, and the Word painter has a headless pixel test, but no native host has been run interactively. Firefox is untested.

## Procedures

### Layout a new machine needs
- Nothing refers to a neighbouring checkout. Every library of another Rbt repository comes from the Forge package registry at the version pinned in `Directory.Build.props`: `Rbt.Vb`, `Rbt.Text.*`, `Rbt.Pdf` for everything, `Rbt.Renderer`, `Rbt.Widgets` and `Rbt.Wisp` for the native hosts and their runners. A NuGet source for the registry with a reader token is the one thing a machine needs. To try a change in one of those libraries here, publish it (below) and bump the pin.
- `pnpm install` in each fresh checkout or worktree.
- A Playwright with a matching Chromium. On the old machine the repo's own had none, so the harnesses were run with `PLAYWRIGHT=<path to another project's node_modules/playwright/index.mjs>`; install browsers (`npx playwright install chromium`) or point `PLAYWRIGHT` likewise. On Ubuntu 26.04 the install needs `PLAYWRIGHT_HOST_PLATFORM_OVERRIDE=ubuntu24.04-x64`, and Chromium needs `libatk1.0-0t64 libatk-bridge2.0-0t64 libatspi2.0-0t64 libxcomposite1 libxdamage1`.

### Gates (run on the merged tree before every release, whatever an agent reported)
```
bash tests/run-all.sh                 # must end: OFFICE TESTS: ALL SUITES PASSED   (~20 min)
dotnet build word-native/Rbt.WordWidgets.fsproj
dotnet build word-wisp/Rbt.WordWisp.fsproj
dotnet build spreadsheet-wisp/Rbt.SpreadsheetWisp.fsproj
dotnet build spreadsheet-native/Rbt.SpreadsheetWidgets.fsproj      # each: 0 Error(s)
pnpm run build                        # 7 x "Fable compilation finished"
```
Browser harnesses, run per branch before merging: `tools/word-browser/{run,perf}.mjs`, `tools/docx-roundtrip/roundtrip.mjs`, `tools/xlsx-browser/{run,perf}.mjs`.

### Release
1. Bump `version` in `package.json` (must equal the tag), commit `Release X.Y.Z`.
2. Fast-forward `master` to `review-fixes`, `git tag vX.Y.Z master`, push `master` and the tag. The `v*` tag publishes `@rbt/office` through the repo's Forge action (`.forge/actions.lua`, `.forge/publish-npm`).
3. Wait for the publish to succeed before pinning anything to it (it has taken up to ~40 minutes and two attempts). `tools/redbat-adoption/forge-status.mjs <repo> <full sha> refs/tags/vX.Y.Z publish-npm` polls `GET /api/forge/repos/<repo>/actions/status?commit=&ref=&action=` on a redbat instance with a dev login; `status: missing` means no action was queued.
4. If another Rbt library changed: tag and publish it first, then bump its version in `Directory.Build.props`; a stale pin fails the restore. Restore with `RestoreNoHttpCache=true` after publishing one. The chain, in publish order: `rbt-text` → `rbt-renderer` (`Rbt.Renderer`, `Rbt.Renderer.Core`) and `rbt-pdf`; `rbt-event-loop` (`Rbt.EventLoop`) and `rbt-renderer` → `rbt-widgets` (`Rbt.Widgets.Platform`, `Rbt.Widgets`) → `rbt-wisp` (`Rbt.Wisp`) → here. A repository publishes on a `vX.Y.Z` tag only if it has a `refs/tags/v*` action policy allowing `package` (Forge, the repository's settings), and an action may declare one publish entry per registry: a repository with two packages has one action for each.

### Redbat adoption
1. In a redbat worktree branched from `master`: set `"@rbt/office": "X.Y.Z"` in `mega/frontend/package.json`.
2. `pnpm install` with the workspace root as cwd, with an npmrc mapping `@rbt:registry` to `https://artifacts.cardillion.com/api/packages/npm/41021897/` plus a reader token for that host. Check `mega/frontend/node_modules/@rbt/office/package.json` shows the new version.
3. Commit `Move to @rbt/office X.Y.Z`, then run the harness in `tools/redbat-adoption/` (see its README): regression (`runall.sh`), host adaptation, save-fidelity tables, performance against the previous release, exploratory flows with screenshots actually looked at.
4. Hold the merge to redbat `master` if a release has a file-corruption bug (0.1.14 was held for one); otherwise fast-forward.

### Working rules that held up
- At most two build-heavy jobs at once on a 2026 workstation; a third swaps it to death.
- One topic per commit, each commit gated on its own; a test seen failing on the old code before the fix.
- Agents: one heavy process each, no sub-agents, never kill processes by name (others share the names), never build through a copy of the tree or symlinked siblings, clean tree and nothing running at the end.
- Fable traps met so far: `=` on a `Map` recurses per entry and overflows at tens of thousands (`Model.sameMap`); `Recalc.diffCells` reads the fields of fable-library's `Map` tree, so check it on a Fable upgrade; `Math.Round` ignores `MidpointRounding`; `Double.ToString` custom formats and `Regex` `Group.Index` unsupported; `String.CompareOrdinal` with offsets broken; `StringBuilder.Append(char, count)` / `Insert` missing; mutable function bindings fail; functions are re-curried across the JS boundary; optional arguments arrive `undefined`; generic record comparison is slow; very long generated JS functions are never optimised by V8 (moving hot paths out of `evalFunction` and `Recalc.run` gave 40%).
