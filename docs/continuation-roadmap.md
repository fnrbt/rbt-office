# Continuation roadmap

Written 2026-10-12, when work was halted to move to another machine. Picks up the fix / release / adopt loop on rbt-office (Word + Spreadsheet) and its adoption in redbat. Everything needed is on the remotes; nothing below depends on the old machine's temp directories.

## Branches (all on `origin` of rbt-office)

| Branch | What it is | State |
|---|---|---|
| `master`, tag `v0.1.16` | Last release, published as `@rbt/office` 0.1.16 | released |
| `review-fixes` | The integration branch. `v0.1.16` + spreadsheet pass eight (9 commits) + this roadmap + `tools/redbat-adoption/` | full gates passed on the spreadsheet commits; **not released** (becomes 0.1.17) |
| `fix/word19` | Word pass nineteen, branched from `v0.1.16` | 4 gated commits, 1 `WIP (ungated)` save commit, 1 handoff commit. Read `tools/handoff-word19/HANDOFF.md` on the branch |
| `fix/sheet19` | Spreadsheet pass nine, branched from `review-fixes` | 1 `WIP (ungated)` chart commit that **does not compile**, 1 handoff commit with the uncompiled drafts. Read `tools/handoff-sheet19/HANDOFF.md` on the branch |

`package.json` says `0.1.16` on every branch. The `tools/handoff-*` folders are temporary: delete them (or drop those commits) before merging a branch into `review-fixes`.

redbat (`cardillion-mega` remote): `master` already contains the 0.1.16 adoption (1094/1094 harness assertions). The local `office-viewers` branch on the old machine equals it; on a new machine just branch from `master`.

## Next steps, in order

1. **Finish `fix/word19`.** Rebase on `review-fixes`. On the WIP save commit run `tools/docx-roundtrip/roundtrip.mjs` and `tools/word-browser/perf.mjs` (everything before them had passed); if green, reword the subject to drop "WIP (ungated): ". Remove `tools/handoff-word19/`. Merge `--ff-only` into `review-fixes`.
2. **Gate the merged tree, release 0.1.17** (procedures below).
3. **Adopt 0.1.17 in redbat.** Host changes to make:
   - Word: remove the host's own Ctrl+Enter handler, its double-click handling for header text boxes, and its shape overlay / handle hit-test / drag / Delete / Escape handling — the engine now does all of these (details in the Word handoff, "What the host should remove or adapt"). Keep honouring `ScrollTop` notices. Tab outside lists and tables now types a tab.
   - Spreadsheet: charts are not present when `ImportXlsx` / `SwitchSheet` returns (`div.spreadsheet-object-pending`, filled a frame later); add a page-break-preview toggle (`SetViewMode("pageBreak")`, `GetViewMode`, `MovePageBreak`); `GoTo` reaches beyond the grid, so the name box's fallback to `SelectCell` can go.
4. **Resume `fix/sheet19`** from its handoff: splice the draft layout into `Charts.fs`, get it compiling, rewrite the two `stage52-charts` assertions that say radar and 3-D charts are not read. Then the rest of its list.
5. Continue with the backlog.

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
- **Nothing has ever been opened in Word, Excel or LibreOffice.** Every fidelity claim rests on structural validation and re-reading with this reader. Installing LibreOffice and adding a headless open/convert check to the round-trip tools is the most valuable missing verification.
- Native hosts are built, and the Word painter has a headless pixel test, but no native host has been run interactively. Firefox is untested.

## Procedures

### Layout a new machine needs
- Sibling checkouts next to this repo for local builds: `../rbt-vb`, `../rbt-text`, `../rbt-pdf`, `../rbt-renderer` (Forge builds use published packages instead; versions pinned in `Directory.Build.props`).
- `pnpm install` in each fresh checkout or worktree.
- A Playwright with a matching Chromium. On the old machine the repo's own had none, so the harnesses were run with `PLAYWRIGHT=<path to another project's node_modules/playwright/index.mjs>`; install browsers (`npx playwright install chromium`) or point `PLAYWRIGHT` likewise.

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
4. If a sibling (`rbt-vb`, `rbt-text`, `rbt-pdf`) changed: tag and publish it first, then bump its version in `Directory.Build.props`. A stale pin only shows up in the Forge build. Restore with `RestoreNoHttpCache=true` after publishing a sibling.

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
