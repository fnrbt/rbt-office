# Continuation roadmap

Written 2026-10-12, when work was halted to move to another machine. Picks up the fix / release / adopt loop on rbt-office (Word + Spreadsheet) and its adoption in redbat. Everything needed is on the remotes; nothing below depends on the old machine's temp directories.

## Branches (all on `origin` of rbt-office)

| Branch | What it is | State |
|---|---|---|
| `master`, tag `v0.1.18` | Last release, published as `@rbt/office` 0.1.18 (the `v0.1.17` tag published nothing: the registry answered its upload with 404, and Forge does not run a failed action again for the same tag and commit): 0.1.16 + spreadsheet pass eight + Word pass nineteen + the registry-only build | released |
| `review-fixes` | The integration branch: what the next release is gated and cut from. Equal to `master` right after a release | nothing unreleased |
| `fix/sheet19` | Spreadsheet pass nine, branched from `review-fixes` as it was at 0.1.16 + pass eight | 1 `WIP (ungated)` chart commit that **does not compile**, 1 handoff commit with the uncompiled drafts. Read `tools/handoff-sheet19/HANDOFF.md` on the branch |

`origin/fix/word19` is the pass as it was handed off (its `tools/handoff-word19/HANDOFF.md` has the host notes); the pass itself is in 0.1.18, rebased, without the handoff folder. The `tools/handoff-*` folders are temporary: delete them (or drop those commits) before merging a branch into `review-fixes`.

redbat (`cardillion-mega` remote): `master` already contains the 0.1.16 adoption (1094/1094 harness assertions). The local `office-viewers` branch on the old machine equals it; on a new machine just branch from `master`.

## Next steps, in order

1. **Adopt 0.1.18 in redbat.** Host changes to make:
   - Word: remove the host's own Ctrl+Enter handler, its double-click handling for header text boxes, and its shape overlay / handle hit-test / drag / Delete / Escape handling — the engine now does all of these (details in the Word handoff, "What the host should remove or adapt"). Keep honouring `ScrollTop` notices. Tab outside lists and tables now types a tab.
   - Spreadsheet: charts are not present when `ImportXlsx` / `SwitchSheet` returns (`div.spreadsheet-object-pending`, filled a frame later); add a page-break-preview toggle (`SetViewMode("pageBreak")`, `GetViewMode`, `MovePageBreak`); `GoTo` reaches beyond the grid, so the name box's fallback to `SelectCell` can go.
2. **Resume `fix/sheet19`** from its handoff: rebase on `review-fixes`, splice the draft layout into `Charts.fs`, get it compiling, rewrite the two `stage52-charts` assertions that say radar and 3-D charts are not read. Then the rest of its list.
3. Continue with the backlog.

## Backlog

### Spreadsheet — bugs in 0.1.16 as released
- **`DeleteRows`, `InsertColumns` and `SetFilter` throw "Maximum call stack size exceeded" on an 846,000-cell sheet.** Not investigated beyond the observation; sort on that sheet is unmeasured.
- Typing a number into a table cell on that sheet takes ~190–250 ms: `commitInput → asTableEntry → changeTables → runBookStep → recalculateEverything` (no hint) → `Recalc.diffCells` enumerates both whole cell maps. Insert row there: ~1 s. Full table of 27 operations in the sheet handoff.
- Dashboard charts ignore manual `c:layout` (legend lands on a neighbouring shape, a fifth bar is lost), `c:dispUnits`, bold in the text-width estimate; KPI shapes with `textlink` draw the cached text, not the cell.

### Spreadsheet — features
- Charts in print/PDF and the canvas painter (plan in the handoff: a `Scene → PaintItem list` adapter; `VecGroup` drops text in both painters today); native and wisp painters do not handle `VecPath`.
- Radar, bubble, stock (parsed in the WIP, drawn only in the draft); chartEx waterfall/funnel/others; slicers and timelines.
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
- Shape selection gaps: keys act on a shape only while no story is edited; the cursor over a handle ignores rotation.
- The `■` third-level bullet in `test-lists.docx` is the fallback face's glyph, not an engine fault; mapping U+25A0 to U+25AA is an option, unverified against Word.
- Never exercised: bold typed into a multi-run field result.

### Cross-cutting
- **LibreOffice is the one outside reader so far; nothing has been opened in Word or Excel.** `tools/libreoffice-check/check.mjs DIR` opens the documents `roundtrip.mjs --out DIR` keeps and compares the text LibreOffice reads from each save with the original's. On 0.1.18: 263 saves read as their originals, none failed. It is not part of the round trip yet (run it after one; write `--out` outside `/tmp`, the 381 documents are 75 MB), it compares text only, and the spreadsheet harness keeps no workbooks for it.
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
