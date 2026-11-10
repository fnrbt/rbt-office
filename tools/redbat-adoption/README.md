# Redbat adoption harness

The scripts used to check a new `@rbt/office` release inside the host application redbat: regression suites, save-fidelity tables, performance runs and repro scripts for package bugs found there. Last full run: 0.1.29, 1104/1104 (the long document is 514 pages since 0.1.29 lays notes out at their real height: `LONG_PAGES` in `lib.mjs`).

This is a snapshot of a working directory, not a polished tool. Paths are hard-coded:

| Path in the scripts | What it is |
|---|---|
| `/tmp/redbat-adoption` | The working directory. Copy this folder there (`cp -r tools/redbat-adoption /tmp/redbat-adoption`) and run from it; builds, logs and screenshots land in it. |
| `/home/dave/Dev/redbat-office-viewers` | A redbat worktree on the branch being tested (Playwright is resolved from its `mega/frontend`). |
| `/home/dave/Dev/rbt-office/tests` | This repo's fixtures. |

On another machine, either create the same paths or `sed` them.

## Use

1. Copy `harness-keep/office-harness.html` → `mega/frontend/`, `office-harness.tsx` → `mega/frontend/src/`, `vite.harness.config.ts` → `mega/frontend/`. These are temporary and must never be committed to redbat.
2. `( cd mega/frontend; npx vite build --config vite.harness.config.ts --outDir /tmp/redbat-adoption/build )`
3. `node static.mjs /tmp/redbat-adoption/build 16139 &`
4. `bash runall.sh <tag>` — prints pass/fail per suite and a total. Suites are the `word-*.mjs`, `sheet-*.mjs`, `ui-*.mjs`, `preview.mjs`, `chat.mjs`, `files.mjs`, `shots.mjs`, `fidelity.mjs` files; `lib.mjs` has the helpers.
5. Remove the three harness files from the redbat tree.

Other scripts: `fid16-*.mjs` / `validate16.py` (which package parts an edit changes; structural validation of saved packages), `perf*.mjs` and `vers-server.mjs` (timings, two builds side by side), `repro-*.mjs` and `probe-bm.mjs` (package bug repros), `explore*.mjs` (exploratory flows with screenshots).

`forge-status.mjs <repo> <sha> <ref> <action>` polls a Forge action run (used to wait for `publish-npm` after tagging a release). It reads a dev cookie from a local redbat instance; adjust for your setup.
