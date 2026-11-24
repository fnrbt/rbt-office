# Handoff: Excel 2019 parity work (2026-11-22)

**Superseded snapshot.** Both branches below are now integrated. Resume on `origin/wip/excel-trailing-fill`; the [current branch handoff](../excel-parity-tracker.md#spreadsheet-branch-handoff) distinguishes shipped work from WIP.

State when the work stopped. Master is at the dashboards round (1580bef4, pushed). Every gate passes there except the Word
browser check `table-insert-xml`, which was already failing. Two rounds are unmerged, each with its own handoff on its
branch (both pushed to origin):

| Round | Branch (origin) | Handoff on that branch | State |
|---|---|---|---|
| Charts and slicers, round 7 | `wip/excel-charts-slicers-round-7` (326cc38a) | `docs/handoff/2026-11-22-excel-charts-slicers-round-7.md` | All 5 items done. Gates not fully confirmed (see below). |
| Print, round 3 | `wip/excel-print-round-3` (af3a9fbf) | `docs/handoff/2026-11-22-excel-print-round-3.md` | WIP: tools only; AutoFit width model not found; items 2-5 not started. |

Read one with `git show origin/wip/excel-print-round-3:docs/handoff/2026-11-22-excel-print-round-3.md`.

## Next steps, in order

1. **Merge the charts and slicers round.** It covers:
   - slicer and caption cuts ("Centr...");
   - Lato in linked pictures;
   - the Pareto inset;
   - sunburst label breaks;
   - filled maps from `cx:geoCache`, with a new `Inflate.fs`.

   Its build and its xlsx-browser checks pass. The one exception is two expectations it updated (`lazy-sheets`,
   `sheet-pdf-default-fonts`), which were rerun only on their own. Its `tests/run-all.sh --core` run had not finished.
   Merge it into master and run the full gates; push if they pass.

   It needs Lato in `artifacts/fonts/windows-extra`. This checkout has it (copied from the VM's Office cloud-font cache);
   a fresh checkout must copy it from there.
2. **Continue print round 3** from its handoff:
   - AutoFit columns: Excel re-fits `bestFit` number columns when it lays out pages. The open question is the font that
     measures formatted numbers; test Small Fonts / MS Sans Serif first.
   - Then icons and sparklines in the painter, hairline borders, Segoe UI and the blank last page, and removing the data
     validation arrow from print (Paint.fs "Pass 6").

   Its scratchpad data is copied to the git-ignored `artifacts/handoff/print-round-3/` (bf3, bf4, pq). The paths in that
   handoff point at the scratchpad copy, which may be gone; use this one.
3. **Remaining differences** are listed per round in `docs/excel-parity-tracker.md` (the sections dated 2026-11-22). The
   largest open ones:
   - the zoomed first screen fitted on open;
   - text off by 1 px at zooms other than 100 %;
   - map layout constants fitted to one picture;
   - text shadow, glow and reflection;
   - IRR / XIRR last digits;
   - subnormal numbers after `*`, `/` and `^`.

## Method and rules (unchanged)

- **Parity loop:** take a discrepancy, state the simplest Excel rule, check generated synthetic fixtures in Excel 2019 on
  the VM, implement only what is confirmed, then add a tracker row and a regression test. Generators and probes are in
  `tools/excel-parity/`.
- **Gates:**
  - `pnpm run build:fable`;
  - `tests/run-all.sh --core` and `--native`;
  - the Word and xlsx browser suites;
  - Word corpus page parity.

  Timing checks (grid-scroll-frames, keystroke-long, macro-callbacks, word-layout perf) flake under load; rerun them on
  their own.
- **Private templates:** the templates in `artifacts/real-estate/` are licensed for private use. Never commit them, their
  pictures or their content; tests use synthetic fixtures only.
- **VM:** quit only the Excel instance you started; never kill EXCEL.EXE by name. A probe whose open fails leaves its Excel
  running; fix that in the probe. At the stop, five EXCEL.EXE processes were running on the VM, and neither agent claimed
  them.
- **Commits:** author fn <fn@cardillion.com>, with no credit lines. Pushing to origin mirrors to GitHub. At most two
  worktree agents at once.
