# rbt-office

Start from `docs/continuation-roadmap.md`. It is the one place the state of the work is kept:
the branches and what is on each, the next steps, the backlog (known bugs and missing features),
and the procedures - the release gate, releasing, adopting a release in redbat, setting up a machine.

- Work in progress that was handed off lives on its branch in `tools/handoff-*/HANDOFF.md`
  (e.g. `fix/sheet19`). The roadmap's branches table says which branch has one. Those folders are
  temporary: they are deleted before the branch is merged into `review-fixes`.
- The redbat adoption harness and its last result: `tools/redbat-adoption/README.md`.
- Other Rbt libraries come only from the Forge package registry, never from a neighbouring checkout.
- Keep the roadmap current: when a step is done, a branch is handed off or merged, or a bug is found
  or fixed, update it in the same change.
