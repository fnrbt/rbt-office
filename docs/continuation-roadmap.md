# Continuation roadmap

Written 2026-10-12, when work was halted to move to another machine. Picks up the fix / release / adopt loop on rbt-office (Word + Spreadsheet) and its adoption in redbat. Everything needed is on the remotes; nothing below depends on the old machine's temp directories.

## Branches (all on `origin` of rbt-office)

| Branch | What it is | State |
|---|---|---|
| `master`, tag `v0.1.19` | Last release, published as `@rbt/office` 0.1.19: 0.1.18 + what adopting it in redbat found - a click on the frame of any shape selects it (one behind the text could not be selected once the host's own frame was gone), the listeners hear of the undo step a shape drag makes, a turned shape's handles from getShapeHandles, GoTo's growth of the grid let go of again. (`v0.1.17` published nothing: the registry answered its upload with 404, and Forge does not run a failed action again for the same tag and commit.) | released |
| `review-fixes` | The integration branch: what the next release is gated and cut from. Equal to `master` right after a release | Unreleased: item 1 of spreadsheet pass nine (chart layout, radar / bubble / stock, shape geometry, linked text) |
| `fix/sheet19` | Spreadsheet pass nine as it was handed off | Its item 1 is in `review-fixes`, rebased, without the handoff folder. Items 2-7 are not started: `tools/handoff-sheet19/HANDOFF.md` on `origin/fix/sheet19` keeps their findings, and `bigsheet.mjs`, `render.fsx` and `shoot.mjs` beside it |

`origin/fix/word19` is the pass as it was handed off (its `tools/handoff-word19/HANDOFF.md` has the host notes); the pass itself is in 0.1.18, rebased, without the handoff folder. The `tools/handoff-*` folders are temporary: delete them (or drop those commits) before merging a branch into `review-fixes`.

redbat (`cardillion-mega` remote): `master` has the 0.1.19 adoption (harness 1104/1104 with `word-new18`, `sheet-new18`; the host no longer draws its own Word shape frame, nor handles header text box double clicks or Ctrl+Enter; the name box goes where GoTo goes; View has a page break preview). On a new machine branch from `master`.

## Next steps, in order

1. **Finish `fix/sheet19`'s list** (its handoff, on `origin/fix/sheet19`): item 2 charts in print, PDF and canvas, item 4 print default sizes, item 5 large-sheet typing and the stack overflows, then 3, 6, 7.
2. Continue with the backlog.

## Backlog

### Spreadsheet — bugs in 0.1.16 as released
- **`DeleteRows`, `InsertColumns` and `SetFilter` throw "Maximum call stack size exceeded" on an 846,000-cell sheet.** Not investigated beyond the observation; sort on that sheet is unmeasured.
- Typing a number into a table cell on that sheet takes ~190–250 ms: `commitInput → asTableEntry → changeTables → runBookStep → recalculateEverything` (no hint) → `Recalc.diffCells` enumerates both whole cell maps. Insert row there: ~1 s. Full table of 27 operations in the sheet handoff.
- "Products sold by Quantity" in the e-commerce dashboard still wraps under its doughnut: the box is narrower in its scaled group. Whether Excel wraps it too is unchecked.

### Spreadsheet — features
- Charts in print/PDF and the canvas painter (plan in the handoff: a `Scene → PaintItem list` adapter; `VecGroup` drops text in both painters today); native and wisp painters do not handle `VecPath`.
- chartEx waterfall/funnel/maps and others; slicers and timelines. Radar, bubble, stock and 3-D-as-flat are drawn; bubble and stock have been looked at on synthetic charts only, no workbook in the corpus has them.
- Shapes: custom geometry, preset polygons, rotation and flips are checked on synthetic shapes only (the corpus has `rect` and `roundRect`, no `rot` or flips). A group's own rotation is ignored. Unwritten: `normAutofit`, `spcBef/spcAft/lnSpc`, `vert`.
- Print and PDF use 25 / 100 px for unsized rows and columns instead of the sheet's defaults (`Pagination.withPrintSizes`, `Pdf.fs` `buildPrintGeometry`, `withPrintColumnWidths`, `withPrintRowHeights`); the page-break preview must agree with print. Preview: grey outside the print area, frames over frozen panes.
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

### Cross-cutting
- **LibreOffice is the one outside reader of the round-trip saves.** Real Word and Excel exist for this: the Windows VM `wordvm` (`DESKTOP-OP2QS0N`, `docs/word-parity.md`, `docs/excel-parity.md`, `tools/word-parity`, `tools/excel-parity`) gave the earlier parity references. The tools were last changed in June 2026 and the VM is not reachable from the new machine: `wordvm` resolves to nothing here. Nothing since (the saves, the chart layout) has been compared with it. `tools/libreoffice-check/check.mjs DIR` opens the documents `roundtrip.mjs --out DIR` keeps and compares the text LibreOffice reads from each save with the original's. On 0.1.18: 263 saves read as their originals, none failed. It is not part of the round trip yet (run it after one; write `--out` outside `/tmp`, the 381 documents are 75 MB), it compares text only, and the spreadsheet harness keeps no workbooks for it.
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
- Fable traps met so far: `Math.Round` ignores `MidpointRounding`; `Double.ToString` custom formats and `Regex` `Group.Index` unsupported; `String.CompareOrdinal` with offsets broken; `StringBuilder.Append(char, count)` / `Insert` missing; mutable function bindings fail; functions are re-curried across the JS boundary; optional arguments arrive `undefined`; generic record comparison is slow; very long generated JS functions are never optimised by V8 (moving hot paths out of `evalFunction` and `Recalc.run` gave 40%).
