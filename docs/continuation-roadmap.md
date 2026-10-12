# Continuation roadmap

Written 2026-10-12. Picks up the autonomous fix/release/adopt loop on rbt-office (Word + Spreadsheet) and its adoption in redbat.

## Where things stand

| Thing | State |
|---|---|
| Latest release | `v0.1.16` (published to the Forge npm registry) |
| `review-fixes` (= working branch) | `v0.1.16` + spreadsheet pass eight (9 commits), gated, **not released** |
| `package.json` version | `0.1.16` — next release is `0.1.17` |
| redbat `master` (local only, never pushed) | on `@rbt/office` 0.1.16, 1094/1094 harness assertions |
| redbat worktree `/home/dave/Dev/redbat-office-viewers` (`office-viewers`) | same as redbat master |

### In flight when this was written

Two background agents were running. Their work is in their worktrees; anything committed there is safe.

| Worktree / branch | Pass | Known state |
|---|---|---|
| `/home/dave/Dev/rbt-office-word19` / `fix/word19` | Word pass nineteen | 1 commit (view anchor kept when a font face arrives mid-layout); standard-keys work was uncommitted |
| `/home/dave/Dev/rbt-office-sheet19` / `fix/sheet19` | Spreadsheet pass nine | just started |

If a worktree has uncommitted changes and no agent running, read the diff, gate it, and either commit or discard.

## Next steps, in order

1. **Land Word pass nineteen.** Check the worktree is clean, `git -C <worktree> rebase review-fixes`, `git merge --ff-only`, remove worktree and branch.
2. **Gate the merged tree** (see "Gates"), then **release 0.1.17** (see "Release").
3. **Adopt 0.1.17 in redbat** on the `office-viewers` worktree (see "Redbat adoption"). Host changes to expect:
   - Spreadsheet: charts are not present when `ImportXlsx`/`SwitchSheet` returns (`div.spreadsheet-object-pending`, filled a frame later); new `SetViewMode("pageBreak")`, `GetViewMode`, `MovePageBreak` — add a page-break-preview toggle; `GoTo` now reaches beyond the grid, so the name-box fallback to `SelectCell` can go.
   - Word: if pass nineteen moved shape selection, header-box double-click and Ctrl+Enter into the engine, remove the host's own versions (its report says exactly what).
4. **Land Spreadsheet pass nine**, release, adopt again.
5. Launch the next passes from the backlog below. Max two build-heavy jobs at once (agents and my own gates both count).

## Backlog

### Word
- Whatever pass nineteen reports as not done from: header/footer text boxes reachable by the engine's own pointer handling; shapes in headers/footers for move/resize/delete; engine-side selected-shape state with painted handles; rotated frames, flips, `allowOverlap`; text in group shapes and canvases.
- Layout: list number wider than its hanging indent; `w:lnNumType`; exact line heights stored as exact; `■` symbol bullet draws as a tall block.
- Save: do not prune unreferenced parts or add `styles.xml`/rels the source lacks; patch the VML fallback of a moved shape instead of dropping it; headers/footers/notes written from the model should carry range marks, not only balance them; comment/permission ranges starting mid-paragraph widen to the paragraph.
- Tables: row/column insert writes many paragraphs from the model (547 / 1552 in the round-trip pass) — patch structurally in source XML.
- Not verified anywhere: bold typed into a multi-run field result; rotated shape frames in the host.

### Spreadsheet
- Whatever pass nine reports as not done from: chart cosmetics on the dashboards, manual `c:layout`, data labels, axes; charts in print/PDF/canvas; radar, bubble, stock, chartEx kinds, slicers; print using the sheet's default row/column sizes; typing on the 846,000-cell sheet (~173 ms); custom table styles; source patching for row insert and table add (row insert: 2 of 110 corpus workbooks patched).
- Page-break preview: grey outside the print area, frames over frozen panes.
- Browser limit: last rows unreachable above ~125% zoom in Chromium and in Firefox.
- A formula's result-format inheritance rules are my reading of Excel, not checked against Excel.

### Cross-cutting
- **Nothing has ever been opened in real Word, Excel or LibreOffice** (`soffice` is not installed). All fidelity claims rest on structural validation and re-reading with this reader. Installing LibreOffice and adding a headless open/convert check to the round-trip tools is the single most valuable verification gap.
- Native hosts (`word-native`, `spreadsheet-native`, wisp) are built and the Word painter has a pixel test; no native host has been run interactively.
- Firefox is untested.

## Procedures

`R=/home/dave/Dev/rbt-office`, `W=/home/dave/Dev/redbat-office-viewers`, `S=` this session's scratchpad (harness lives in `$S/adopt11/`; copy to `adopt12/` and repoint paths for the next run).

### Gates
```
( cd $R; bash tests/run-all.sh )            # must end: OFFICE TESTS: ALL SUITES PASSED
dotnet build  word-native/Rbt.WordWidgets.fsproj, word-wisp/Rbt.WordWisp.fsproj,
              spreadsheet-wisp/Rbt.SpreadsheetWisp.fsproj, spreadsheet-native/Rbt.SpreadsheetWidgets.fsproj   # 0 Error(s)
pnpm run build                              # 7 x "Fable compilation finished"
```
Browser harnesses (`tools/word-browser/{run,perf}.mjs`, `tools/docx-roundtrip/roundtrip.mjs`, `tools/xlsx-browser/{run,perf}.mjs`) need
`PLAYWRIGHT=/home/dave/Dev/translation/new/translation/node_modules/playwright/index.mjs`.
Do not merge into `$R` while gates run there.

### Release
1. Bump `version` in `package.json`, commit `Release X`.
2. `git -C $R fetch -q . review-fixes:master; git -C $R tag vX master; git -C $R push -q origin master vX`.
3. Poll `node $S/status.mjs rbt-office <sha> refs/tags/vX publish-npm` until `"status":"succeeded"` (up to ~40 min). **Do not pin redbat before it has succeeded.**

### Redbat adoption
1. `git -C $W rebase master`; set `"@rbt/office": "X"` in `mega/frontend/package.json`.
2. `pnpm install` in `$W` with a temporary npmrc (`@rbt:registry=https://artifacts.cardillion.com/api/packages/npm/41021897/` and the token from `~/.nuget/NuGet/NuGet.Config`); delete the npmrc; check `node_modules/@rbt/office/package.json`.
3. Commit `Move to @rbt/office X`; run an adoption agent (regression via `runall.sh`, host adaptation, fidelity tables, perf vs previous, exploratory).
4. If no corruption-class package bug: `git -C $W rebase master`, `git -C /home/dave/Dev/redbat merge --ff-only office-viewers`, `pnpm install --frozen-lockfile` there.

### Standing rules
- Never push or deploy redbat master; never touch `/home/dave/Dev/redbat-test`.
- No approval prompts, no questions: absolute paths, `git -C`, subshells; no `cd x && ...`.
- Agents: one heavy process each, no subagents, never kill by process name, fails-first tests, every commit gated on its own, clean tree and no background processes at the end.
- Re-run the full gates on the merged tree before every release regardless of what an agent reports.
