# Redbat review-harness handoff — @rbt/office 0.1.31

## Internal release versus fidelity approval

Publish the automated-tested internal package first so Redbat can adopt it and build this harness. Human fidelity review is **not** a prerequisite for internal publication. The checklist below describes work to do in that harness, not a claim that Office parity is complete.

## Reference evidence and locations

The completed Office run below belongs to candidate `52d305e2`. Keep that provenance distinct from the final publication commit and registry tarball; do not silently relabel old evidence. The release changes publishing policy/documentation, not editor runtime code.

- Commit: `52d305e242bc62f9a5bb949835dcc3bbb5970137`
- Package SHA256: `cc669f87bbda1d0cd6ff835ee2581695453e6cf8647ce739701af43701634311`
- Manifest SHA256: `9cb13b6ad5ccdb20ea52aa1f4a59cd769c14a82bd520106673761a5064384535`
- Linux bundle: `/home/dave/Dev/rbt-office-releases/0.1.31-52d305e2/`
- winvm corpus: `C:\Users\olivia\RbtOfficeRelease-52d305e2\`
- Documents: `word/docs/` and `excel/docs/` beneath either location.
- Findings: `review-notes.md`. Detailed differences: `office-comparison.json` and `excel-comparison.json`.

Use copies for manual edits. Do not overwrite the manifest-bound originals or generated outputs.

## Redbat setup and data contract

1. After the registry confirms publication, pin `"@rbt/office": "0.1.31"` in Redbat's `mega/frontend/package.json` and run `pnpm install` from its workspace root. Use the existing reader credentials and `@rbt:registry=https://artifacts.cardillion.com/api/packages/npm/41021897/`; do not embed tokens in the harness.
2. Record the installed version and registry `dist.integrity`. `npm view @rbt/office@0.1.31 version dist.integrity --registry=https://artifacts.cardillion.com/api/packages/npm/41021897/` identifies the actual published artifact. Publication status is established by the registry, not merely by a tag or this document.
3. Transfer `/home/dave/Dev/rbt-office-releases/redbat-review-0.1.31.zip` to the Redbat workstation, extract it outside the source tree, and select the extracted directory in the harness. Its `SHA256SUMS` binds the reference corpus and tarball; `review-data-sha256.json` binds the included review files. The archive is a portable copy, not an npm dependency and not a public upload.
4. Use the existing package adapters in `tools/redbat-adoption/harness-keep/office-harness.tsx` as the integration reference. Package imports use `@rbt/office/word/...`, `@rbt/office/spreadsheet/...` and `@rbt/office/office-formats/src/Main.js`; see `docs/word-host-api.md` and `docs/spreadsheet-host-api.md`. The existing regression harness is documented in `tools/redbat-adoption/README.md`; it is not the new human review UI.
5. Keep the review loop local: file/directory selection, rendering, editing, dialogs, comparisons and downloading results need no HTTP/controller. Loading the application/dependency assets is separate. Backend storage is an optional host hook, not a requirement for basic review.

### Inputs to load

| Relative path | Meaning / parsing |
|---|---|
| `commit`, `version`, `SHA256SUMS` | Reference candidate identity and per-file SHA-256 bindings. |
| `word/docs/*.docx` | Originals, first/second saves, model-only reconstructions and named edit outputs. |
| `excel/docs/*.xlsx` | Originals, first/second saves, cell edits and row/column insertions. |
| `office-comparison.json` | Object with `word` and `excel` comparison arrays, `openFailures`, `capabilityErrors` and `operationErrors`. Use each entry's `original`/`reference` and `candidate` names to associate records with files. |
| `excel-comparison.json`, `excel-summary.json` | Additional classified spreadsheet results and counts. |
| `word/opened.tsv`, `excel/opened.tsv` | Aggregate open results, terminated by `DONE`; opening is not fidelity approval. |
| `word/opened-detail.tsv` | ASCII, five tab-separated fields: escaped filename, kind, index, field, escaped value. |
| `excel/evidence/*.tsv` | UTF-16 records, first field is the record kind; includes expected-insertion sidecars without corresponding XLSX files. |
| `static-interaction.zip`, `prepare.log`, `review-notes.md` | Offline screenshots/workflow evidence, per-fixture export diagnostics and known limits. |
| `office-tools.zip` | The exact Office probes and runner/collection helpers used for the reference evidence. |

Comparison entries have `removedCount`, `addedCount`, `removed`, `added`; unavailable Excel baselines instead have `comparisonUnavailable`. Do not interpret an unavailable comparison as an empty diff.

Word diff records are `[kind, index, field, value]`; Excel records retain their kind-specific field order. For example, CELL is `[CELL, sheet, address, type, value, formula]`; SHAPE is `[SHAPE, sheet, index, name, type, topLeft, bottomRight, left, top, width, height]`, with physical dimensions in points. Preserve strings for comparison rather than coercing all values to numbers.

The probe escape layer remains inside the parsed JSON strings: decode `\\`, `\t`, `\r`, `\n`, and, for Word, `\uXXXX`, in a single left-to-right pass. Do not double-decode literal backslashes. Ordinary JSON decoding alone does not remove this probe layer.

### Review UI and saved results

- Show a case list grouped by Word/Excel, operation, differing/matching/unavailable status and known capability limits. Keep model-only reconstruction separate from ordinary saves.
- Show original and candidate side by side, the intended operation, and raw record differences. Offer download/open-in-Office links; COM records are not screenshots. Human Office screenshots/observations can be attached to the case.
- Allow normal local edits, undo/redo and save/reopen in the package-backed view. Offer the same workflow through Redbat's browser and desktop host adapters; record which host was exercised.
- Load document bytes only for the selected case; do not mount or parse the entire corpus at once. Large comparison files can be indexed during harness import, with raw evidence retained separately.
- Record a verdict per case: pass, defect, accepted limitation or not verified; include explanation, exact filenames, operation, app/Office versions, reviewer, date and optional screenshots.
- Export review results as JSON plus attachments. Bind results to the installed package version/integrity and the reference corpus manifest hash separately. Do not overwrite the reference comparison JSON, suppress capability errors or automatically produce a human attestation.
- A pass requires only intended changes and preserved unrelated content. Opening a file or getting an empty record diff must not auto-approve its visual appearance.

The fixture priorities below are the initial review queue. They are not a request to visually inspect all 2,168 files before Redbat can adopt the internal package.

## Already completed — do not repeat merely to establish that files open

- Full release preparation: Fable builds, behavioral core tests, installed-package checks and export/package checks.
- Exact installed package running from `file://`: open, type, undo/redo, clipboard, dialog cancel/apply, save, two fresh reopens; zero HTTP attempts or page errors.
- Actual Office on winvm opened all **1,608 DOCX and 560 XLSX** inputs, with complete recorded coverage.

These establish execution and opening, not complete visual or editing fidelity. The remaining work is comparison, unsupported-capability testing and explicit release decisions.

## 1. Word: review suspected editing losses first

Open each original beside its saved or edited counterpart in real Word. Check body text, comments, controls, tables, fields, headers/footers, page breaks and print preview. For edited outputs, distinguish the intended edit from collateral changes; the generated edit descriptions and per-fixture diagnostics are in `prepare.log`.

| Priority | Files / variants | What to check | Acceptable result |
|---|---|---|---|
| Blocking investigation | `patched-paragraphs.original.docx` versus `patched-paragraphs.bold-chosen-at-a-caret-{1,2,5}-character[s]-typed.docx` (actual names below) | Diagnostics report removed comment-range markers. Inspect comment text, attachment and selected range after typing. | No unintended orphaned, deleted or relocated comment. An intentional change must be explained. |
| Blocking investigation | Paragraph-join outputs: suffix `.joined-to-the-one-before-Backspace.docx`, especially `docx4j-invoice-2013` | Diagnostics report lost content-control wrappers and paragraph properties. Check control boundaries, mapped values, spacing and indentation around the join. | Only the requested merge changes content/formatting; unrelated controls and bindings survive. |
| High | `.split-in-two-Enter.docx` outputs and other affected fixtures listed in `prepare.log` | Some operations regenerate neighboring paragraphs. Compare adjacent text, formatting, pagination, headers and footers. | No unintended changes outside the edited range. |
| High | `docx4j-invoice-2013` and `reader-structure`: `.original`, `.pass1`, `.pass2` | Four ordinary-save comparisons differ only in reported control IDs. Check control type, text, title/tag and XML mapping—not just appearance. | Controls remain usable and mapped correctly; assess whether any integration depends on their IDs. |
| High | Table row insertion/deletion, column insertion and cell-shading variants of `reader-structure` and `docx4j-invoice-2013` | Inspect edited table geometry, text, borders, controls and pagination. | Intended table change only; no displaced or missing content. |
| Scope decision | All 42 differing `.model.docx` comparisons, identified in `office-comparison.json` | Model-only reconstruction discards retained source-package state. Review and classify differences separately from normal saves. | Do not advertise this path as lossless. Decide whether the observed losses are acceptable for its exposed use. |

Exact comment-regression filenames:

- `patched-paragraphs.bold-chosen-at-a-caret-1-character-typed.docx`
- `patched-paragraphs.bold-chosen-at-a-caret-2-characters-typed.docx`
- `patched-paragraphs.bold-chosen-at-a-caret-5-characters-typed.docx`

The ordinary-save comparisons otherwise match captured records. They do not replace visual checks of layout and formatting. Record which originals and variants the human actually reviewed; do not imply every edited file was visually inspected.

## 2. Excel: compare structural edits and their intended semantics

For each fixture below, compare `.original.xlsx`, `.save.xlsx`, `.save2.xlsx`, `.edit.xlsx` and `.insert.xlsx` as relevant. The generated insertion targets **row 2 and column B on the imported active sheet**, not necessarily the first worksheet.

For a visual Excel-generated baseline, open a copy of the original and perform those insertions on its active sheet. The `*.expected-insert.xlsx.tsv` files are recorded comparisons, **not saved baseline workbooks**.

| Fixtures | Review / test | Acceptable result or required decision |
|---|---|---|
| `conditional-formatting-tutorial` | Check the cached whitespace results and their displayed behavior after opening and recalculation. | Spaces must not disappear or turn into different values. The recorded insertion comparison now matches. |
| `jumeirah-hotel`, `saopaulo-condos`, `milk-street-office`, `lyric-multifamily`, `synthetic-cell-comments-print` | Inspect note text, owning cells, box sizes and placement; check inserted columns remain visible. | No lost notes, accidental resizing or hidden zero-width inserted columns. Explicitly assess geometry differences from Excel's own insertion. |
| `poi-49156` | Compare the absolute picture before and after insertion. | Original physical position and size remain unchanged. The exact candidate's recorded geometry satisfies this; Excel's own insertion normalizes it slightly differently. |
| `dashboard-ecommerce-sales`, `poi-47504` | Inspect drawing positions/sizes and dashboard layout after insertion. | No clipping, overlap or unusable layout; classify remaining geometry differences rather than calling them exact parity. |
| `dcf-stock-analysis`, `pandas-testmultiindex` | Confirm edits happened on the active sheet, with unrelated sheets preserved. | Correct target sheet and intended changes only. Recorded insertion comparisons now match. |
| `lyric-multifamily`, `poi-47804` | Compare differing formula strings and recalculate. | Equivalent results and references; do not assume changed spelling is harmless without checking semantics. |
| `pandas-test_decimal` | Inspect the inserted table column and its generated header. | Header is unique and references remain correct; decide whether the naming difference from Excel is acceptable. |
| `synthetic-formatting` | Inspect the note exposed after export. Its original package has a comments part without the associated VML drawing. | Existing note content is preserved; document the source-metadata normalization. |
| `phpsp-32chartreadwrite` | Check formulas and chart references, allowing for RANDBETWEEN recalculation. | Volatile values may differ; formulas, references and nonvolatile content must not change unexpectedly. |

**Product-policy decision:** insertion currently uses model-default row/column dimensions, rather than Excel's inherited neighbor dimensions. This explains some geometry differences but is not automatic approval. Accept and document that behavior, or request an implementation change and a new candidate.

### Rejected insertion baselines

- `synthetic-slicers`: Excel rejects column-B insertion through the active PivotTable.
- `synthetic-table-styles`: Excel rejects the requested column insertion.

Inspect candidate table/PivotTable integrity and decide whether the application should reject the operation or intentionally support different semantics. Do not switch worksheets to manufacture a passing baseline. These two cases remain without equivalent Excel-produced insertion comparisons.

## 3. Tests needing a more capable Excel installation

winvm runs **Office 2019 Professional2019Retail**, Click-to-Run `16.0.19127.20860`. A recent build number does not make it Microsoft 365.

- [ ] On an Excel edition/build supporting the relevant functions, compare originals and exports for `synthetic-dynamic-arrays`, `synthetic-groupby` and `synthetic-lambda`. Recalculate and inspect formulas, spill dimensions, values and neighboring cells after save and edits. Original files also report unsupported-function errors on winvm; its results cannot certify modern-function behavior.
- [ ] On an installation exposing threaded comments, inspect threads, authors, text and replies before/after save and structural edits. The current probe recorded 1,023 unavailable CommentsThreaded API results; it did not verify those records.
- [ ] Where chart-series formulas are accessible, inspect the series sources that produced 24 Series.Formula capability errors. Include the dashboard's affected charts. Confirm intended series and rendered data rather than treating missing COM evidence as a pass.

Record unavailable tests as **not verified**, not passed. If these capabilities are outside the release's supported scope, that limitation needs an explicit release decision.

## 4. Visible rendering limits requiring acceptance

The exact offline screenshots show:

- Word retains the imported chart package parts but displays a chart placeholder.
- The spreadsheet's map chart is explicitly not drawn.

Review `static-interaction.zip`. Decide whether these limits are acceptable for this release and describe them accurately. Preserved package bytes do not establish that the static page or desktop UI renders the feature.

## 5. Record fidelity-review results

For each reviewed case record:

- Exact filenames, application edition/build and operation.
- Expected versus observed behavior; screenshots where appearance matters.
- Result: pass, defect requiring a fix, explicitly accepted limitation, or not verified.
- Reviewer and date. Keep results with this bundle.

Unintended loss of supported document content, broken controls/comments, incorrect formula results or invalid table/PivotTable edits must be resolved before approval. Accepted limitations must be explicit; an open count or a zero-exit test is not a substitute.

After actual human Word **and** Excel comparison, a reviewer can create `oracle.json` in the reviewed bundle as a separate fidelity attestation. It is not an internal-publish gate. Required fields, as defined by `scripts/verify-release-oracle.mjs`:

- `commit`, `packageSha256`, `manifestSha256`: exact values at the top of this file.
- `reviewer`, `reviewedAt` (ISO date), `wordVersion`, `excelVersion`.
- `comparison`: substantive review results, coverage and accepted/unverified limits.
- `wordAndExcelCompared`: `true` only after that comparison actually happened.

Do not create a placeholder or automatically approve the attestation. Any implementation change requires appropriately refreshed evidence; retained historical comparisons must keep their original bindings. Internal publication proceeds through the automated release gate independently of this later human review.
