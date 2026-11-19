# Word Rendering Parity Work

Current goal: iteratively improve rbt-office's Word editor render/layout parity against Microsoft Word for real DOCX files and focused generated repros.

## Confirmed Repo Map

- Browser/Fable Word path: `src/Word/`
  - Layout core: `src/Word/Layout.fs`
  - Browser canvas measurement/render bridge: `src/Word/Renderer.fs`
  - Browser API exports include `getLayoutSegments`, `importFromDocxPackage`, `exportDocumentPackage`: `src/Word/Word.fs`
  - Shell/host bridge and PDF export hook: `src/Word/WordShell.fs`, `src/Word/WordHost.fs`
- Native path: `word-wisp/`
  - Shared renderer-only adapter: `office-wisp/word-render/WordRender.fs`
  - Canonical layout settings: `src/Word/Layout.fs`; fonts/images: `office-wisp/native/`
  - Wisp chrome uses `office-wisp/OfficeShell.fs`; the Widgets host is retired.
- Print/PDF path: `src/Word.Pdf/`
  - PDF layout/paint conversion: `src/Word.Pdf/Pdf.fs`
- DOCX read/write path: `src/Office.Formats/`
  - DOCX reader: `src/Office.Formats/src/DocxReader.fs`
  - DOCX writer: `src/Office.Formats/src/WordPackageWriterModel.fs`, `WordPackageWriterSplice.fs`, `WordPackageWriterAssembly.fs`
  - Focused DOCX creation helper: `src/Office.Formats/src/DocumentBuilder.fs`

## Fixtures

The extraction contains many DOCX fixtures under:

- `tests/test-documents/`
- `tests/test-documents-stress/`
- `tests/test-documents-edge-cases/`
- `tests/test-documents-business/`
- `tests/test-documents-wild/`
- `tests/test-documents-features/`
- `tests/test-documents-online/`

Good initial candidates:

- `tests/test-documents/sample-formatted.docx`
- `tests/test-documents/calibre-demo.docx`
- `tests/test-documents-business/docx4j-tables.docx`
- `tests/test-documents-edge-cases/complex-tables.docx`
- `tests/test-documents-features/test-lists.docx`

## Comparison Loop

1. Pick a real DOCX fixture or generate a focused repro with the DOCX writer.
2. Query Microsoft Word for page/paragraph/line geometry.
3. Query our browser path via `getLayoutSegments` and visual screenshot checks.
4. Query native path via `word-native`/`../widgets`.
5. Query print path via `src/Word.Pdf`.
6. Compare paragraph and line top/left/width/height positions first, then inspect screenshots/PDF output.
7. Fix the smallest layout/import/export discrepancy that explains the measured delta.
8. Keep the repro fixture and comparison output.

For each identified bug, prefer this stricter loop:

1. Add or generate a focused DOCX repro under `tests/repros/word/`.
2. Capture the old wrong rendering before editing code. Use a temporary git worktree at the pre-fix commit and write PNG/JSON outputs under `artifacts/baseline/`; these artifacts are for records and are not committed.
3. Fix the importer/layout/rendering bug.
4. Re-run Word/browser/native/PDF as appropriate.
5. Commit the repro and the code fix together, or commit repro/tooling separately when retrofitting an already-fixed bug.

## Current Machine State

- Use `winvm` for Office automation. `ssh winvm "hostname && whoami"` returns `DESKTOP-NNIBN4S` and `desktop-nnibn4s\olivia`.
- Word automation should live in checked-in `.vbs` files and be invoked with `cscript`; avoid one-off giant SSH/PowerShell command lines.
- The Linux wrappers default to `WORD_VM_RUN_MODE=task`, which copies the checked-in VBS to the VM and runs it through an interactive-user scheduled task. This mode is verified with `tests/test-documents-features/test-formatting.docx`.
- `WORD_VM_RUN_MODE=direct` runs `cscript` directly over SSH. Direct SSH copied files to the VM, but Windows blocked `CreateObject` in the SSH-launched VBS process with `Permission denied` here, so keep task mode as the default unless the VM setup changes.
- Windows/Word fonts are copied from the VM into uncommitted local artifacts with `tools/word-parity/fetch-windows-fonts.sh`. Browser, native, and PDF runners use `RB_WORD_FONT_DIR` when set, otherwise they look for `artifacts/fonts/windows`.
- No old `.vbs`, `.vba`, `.ps1`, `.cmd`, or comparison scripts were present in the extracted repo.
- `dotnet restore Office.slnx` completed.
- `dotnet build Office.slnx --no-restore` succeeds after adapting `src/Word.Pdf/Pdf.fs` to the updated `../pdf` subset API and restoring `onText` support in `../widgets`.

## Added Tooling

- `tools/word-parity/extract-word-layout.vbs` extracts Microsoft Word paragraph/line geometry to TSV, including table row/column context and Word point positions.
- `docs/word-vba-layout-api.md` tracks the Microsoft Word VBA layout-query API surface and which parts our browser API mirrors.
- `tools/word-parity/print-word-pdf.vbs` opens a DOCX in Word and writes a PDF.
- `tools/word-parity/page-parity.mjs` (with `page-lines.vbs`) compares our page count and every page's first and last line with Word's for the whole corpus; Word's side is committed in `tests/word/word-pages/`. See `tools/word-parity/README.md`.
- `tools/word-parity/run-word-vbs-on-vm.sh` copies a VBS script and DOCX to `winvm`, runs `cscript`, and copies the output back.
- `tools/word-parity/open-all.vbs` retains aggregate `opened.tsv` and adds escaped `opened-detail.tsv`: exact body/story/table/control text, XML bindings, field types/codes/results, TOC counts, hyperlink targets, drawing structure and pagination. It opens read-only with macros/link updates disabled and `OpenAndRepair=False`, closes its documents and quits its own application without process-name kills. A document window is required: Word 16.0.19127 returns `E_FAIL` for `Shape.Type` when `Documents.Open` uses `Visible=False`; the focused original/model probe passes with `Visible=True`. Word exposes no authoritative repair-result property, so repair status is explicitly unknown.
- `tools/word-parity/run-word-vbs-task.ps1` runs the copied VBS through Windows Task Scheduler with an interactive-user token.
- `tools/word-parity/word-layout-tsv.sh` wraps the layout extractor.
- `tools/word-parity/word-print-pdf.sh` wraps the PDF generator.
- `tools/word-parity/create-focused-repros.py` writes deterministic minimal DOCX repro files.
- `tools/word-parity/capture-browser-docx.mjs` captures a browser canvas PNG and optional layout JSON for a DOCX. Set `OFFICE_ROOT=/path/to/worktree` to capture from an old commit.
- `tools/word-parity/compare-word-browser-layouts.mjs` runs Word TSV extraction plus browser line capture for one or more DOCX files/directories, then writes per-document layout diffs and a ranked summary.
- `tools/word-parity/compare-vba-layout-results.mjs` runs Word TSV extraction plus browser `getRangeInformation`/`computeRangeStatistics` calls for matched line ranges. Use it to verify that our VBA-style layout query results match Microsoft Word separately from rendered screenshot comparison.
- `tools/word-parity/fetch-windows-fonts.sh` copies Calibri, Arial, Times New Roman, and related Windows fonts from `winvm` into `artifacts/fonts/windows`.
- `tools/browser-runner/` is a standalone Fable/Vite browser harness for the canvas renderer. It loads DOCX files through `Office.Formats`, imports them into `Rbt.Word.API`, exposes `window.rbWordRunner`, and can write segment JSON, line-box JSON, plus canvas screenshots through Playwright.
- `window.rbWordRunner.getVbaLayoutLines()`, `getRangeInformation(...)`, `getSelectionInformation()`, `getRangeBounds(...)`, and `computeRangeStatistics(...)` expose VBA-style layout probes on our browser path.
- `tools/native-runner/` is a native DOCX-to-PNG snapshot harness. It loads DOCX bytes with the native `DocxReader`, maps the typed parse tree into `Rbt.Word.Document`, renders `word-native` through `../widgets` + `../renderer`, and writes a PNG.
- `tools/pdf-runner/` is the project PDF backend harness. It reuses the native typed DOCX import, calls `Rbt.Word.Pdf.toBytes`, and writes the resulting PDF.

Run on the Windows VM from an interactive session:

```cmd
cscript //nologo tools\word-parity\extract-word-layout.vbs C:\path\input.docx C:\path\word-layout.tsv
```

Or from Linux, using the VM wrapper:

```sh
tools/word-parity/word-layout-tsv.sh tests/test-documents/sample-formatted.docx artifacts/word/sample-formatted.tsv
tools/word-parity/word-print-pdf.sh tests/test-documents/sample-formatted.docx artifacts/word/sample-formatted.pdf
```

Smoke-tested commands:

```sh
tools/word-parity/word-layout-tsv.sh tests/test-documents-features/test-formatting.docx artifacts/word/test-formatting.tsv
tools/word-parity/word-print-pdf.sh tests/test-documents-features/test-formatting.docx artifacts/word/test-formatting-word.pdf
```

Font setup:

```sh
tools/word-parity/fetch-windows-fonts.sh
```

This writes uncommitted files under `artifacts/fonts/windows`. To use another font directory:

```sh
RB_WORD_FONT_DIR=/path/to/windows/fonts dotnet run --project tools/pdf-runner/PdfRunner.fsproj -- input.docx output.pdf
```

Browser runner commands:

```sh
cd tools/browser-runner
npm install
npm run build:fable
npm run dev
```

Then open:

```text
http://127.0.0.1:5173/?doc=/@fs//home/dave/Dev/browser/office/tests/test-documents-features/test-formatting.docx
```

Automated smoke check:

```sh
cd tools/browser-runner
npm run check
```

The check currently writes:

- `artifacts/browser/test-formatting-layout.json`
- `artifacts/browser/test-formatting.png`

Native runner command:

```sh
dotnet run --project tools/native-runner/NativeRunner.fsproj -- \
  tests/test-documents-features/test-formatting.docx \
  artifacts/native/test-formatting.png \
  900 1600
```

The native runner currently writes a full-widget raster snapshot. It preserves typed paragraph/table blocks directly from the native DOCX parse tree, avoiding the browser-only Fable object bridge.

Smoke-tested output:

- `artifacts/native/test-formatting.png`

Project PDF runner command:

```sh
dotnet run --project tools/pdf-runner/PdfRunner.fsproj -- \
  tests/test-documents-features/test-formatting.docx \
  artifacts/pdf/test-formatting.pdf
```

Smoke-tested outputs:

- `artifacts/pdf/test-formatting.pdf`
- `artifacts/pdf/test-complex-tables.pdf`

Focused repro commands:

```sh
tools/word-parity/create-focused-repros.py tests/repros/word
node tools/word-parity/capture-browser-docx.mjs \
  tests/repros/word/table-import-preserves-structure.docx \
  artifacts/browser/table-import-current.png \
  artifacts/browser/table-import-current.json
```

For line-level browser geometry instead of segment-level geometry:

```sh
BROWSER_CAPTURE_LAYOUT=lines node tools/word-parity/capture-browser-docx.mjs \
  tests/repros/word/table-import-preserves-structure.docx \
  artifacts/browser/table-import-current.png \
  artifacts/browser/table-import-lines.json
```

Word/browser numeric comparison:

```sh
node tools/word-parity/compare-word-browser-layouts.mjs \
  --out artifacts/layout-compare-current \
  --max-lines 180 \
  tests/repros/word \
  tests/test-documents-features/test-complex-tables.docx \
  tests/test-documents-features/test-formatting.docx
```

This writes:

- `artifacts/layout-compare-current/summary.tsv`: ranked per-document deltas above the configured threshold.
- `artifacts/layout-compare-current/<doc>/word-layout.tsv`: raw Word line geometry in points.
- `artifacts/layout-compare-current/<doc>/browser-lines.json`: browser line boxes in CSS pixels.
- `artifacts/layout-compare-current/<doc>/layout-diff.tsv`: matched Word/browser line pairs with `dx`, `dy`, and consecutive-line `gapDy`.

Word/browser VBA-style result comparison:

```sh
node tools/word-parity/compare-vba-layout-results.mjs \
  --out artifacts/vba-layout-compare-current \
  tests/repros/word/table-valign-row-height.docx \
  tests/repros/word/list-default-indent.docx
```

This writes:

- `artifacts/vba-layout-compare-current/summary.tsv`: ranked mismatches in VBA-style call results.
- `artifacts/vba-layout-compare-current/<doc>/word-layout.tsv`: raw Word `Range.Information`/`ComputeStatistics` values emitted by VBS.
- `artifacts/vba-layout-compare-current/<doc>/browser-range-info.json`: browser line ranges with direct `getRangeInformation`/`computeRangeStatistics` results.
- `artifacts/vba-layout-compare-current/<doc>/vba-diff.tsv`: matched Word/browser call-result pairs, including page, text-boundary, table, list-column, and statistics fields.

Old-render capture example:

```sh
git worktree add /tmp/office-before-fix <commit>
OFFICE_ROOT=/tmp/office-before-fix \
  node tools/word-parity/capture-browser-docx.mjs \
  tests/repros/word/table-import-preserves-structure.docx \
  artifacts/baseline/table-import-before-fix.png \
  artifacts/baseline/table-import-before-fix.json
```

The TSV columns are:

```text
doc	page	paragraph	line	start	end	x	y	lastX	lastY	inTable	table	row	column	style	text
```

Word positions are reported in points relative to the page. The comparator converts them to CSS pixels with `96 / 72` before comparing against browser line boxes. `lastX`/`lastY` are the last non-marker character position Word exposes for the line; use them as hints, not as exact line width.

Browser line JSON records include:

```text
Kind	BlockId	BlockIndex	LineIndex	Page	X	Y	PageY	Width	Height	LineWidth	StartOffset	EndOffset	BlockY	BlockHeight	AvailableWidth	Indent	Text	Segments
```

`X` and `PageY` are the comparable page-relative CSS-pixel coordinates. `Y` is absolute canvas Y including page gaps.

## Next Setup Tasks

- Add a native geometry export alongside the native PNG runner if we need numeric native line/paragraph positions, not just raster snapshots.
- Extend the comparison driver to optionally include native geometry once the native runner exports line boxes.
- Resume parity fixes on tables/floats using the existing real-world fixtures and focused DOCX writer repros.
