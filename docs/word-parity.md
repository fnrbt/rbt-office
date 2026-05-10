# Word Rendering Parity Work

Current goal: iteratively improve RedBat Word render/layout parity against Microsoft Word for real DOCX files and focused generated repros.

## Confirmed Repo Map

- Browser/Fable Word path: `src/Word/`
  - Layout core: `src/Word/Layout.fs`
  - Browser canvas measurement/render bridge: `src/Word/Renderer.fs`
  - Browser API exports include `getLayoutSegments`, `importFromDocxPackage`, `exportDocumentPackage`: `src/Word/Word.fs`
  - Shell/host bridge and PDF export hook: `src/Word/WordShell.fs`, `src/Word/WordHost.fs`
- Native path: `word-native/`
  - Native layout settings and renderer integration: `word-native/WordEditor.fs`
  - Uses `../widgets`.
- Print/PDF path: `src/Word.Pdf/`
  - PDF layout/paint conversion: `src/Word.Pdf/Pdf.fs`
- DOCX read/write path: `src/Office.Formats/`
  - DOCX reader: `src/Office.Formats/src/DocxReader.fs`
  - DOCX writer: `src/Office.Formats/src/DocumentWriter.fs`
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

- `ssh wordvm hostname` works and returns `DESKTOP-OP2QS0N`.
- Word automation should live in checked-in `.vbs` files and be invoked with `cscript`; avoid one-off giant SSH/PowerShell command lines.
- The Linux wrappers default to `WORD_VM_RUN_MODE=task`, which copies the checked-in VBS to the VM and runs it through an interactive-user scheduled task. This mode is verified with `tests/test-documents-features/test-formatting.docx`.
- `WORD_VM_RUN_MODE=direct` runs `cscript` directly over SSH. Direct SSH copied files to the VM, but Windows blocked `CreateObject` in the SSH-launched VBS process with `Permission denied` here, so keep task mode as the default unless the VM setup changes.
- No old `.vbs`, `.vba`, `.ps1`, `.cmd`, or comparison scripts were present in the extracted repo.
- `dotnet restore Office.slnx` completed.
- `dotnet build Office.slnx --no-restore` succeeds after adapting `src/Word.Pdf/Pdf.fs` to the updated `../pdf` subset API and restoring `onText` support in `../widgets`.

## Added Tooling

- `tools/word-parity/extract-word-layout.vbs` extracts Microsoft Word paragraph/line geometry to TSV.
- `tools/word-parity/print-word-pdf.vbs` opens a DOCX in Word and writes a PDF.
- `tools/word-parity/run-word-vbs-on-vm.sh` copies a VBS script and DOCX to `wordvm`, runs `cscript`, and copies the output back.
- `tools/word-parity/run-word-vbs-task.ps1` runs the copied VBS through Windows Task Scheduler with an interactive-user token.
- `tools/word-parity/word-layout-tsv.sh` wraps the layout extractor.
- `tools/word-parity/word-print-pdf.sh` wraps the PDF generator.
- `tools/word-parity/create-focused-repros.py` writes deterministic minimal DOCX repro files.
- `tools/word-parity/capture-browser-docx.mjs` captures a browser canvas PNG and optional layout JSON for a DOCX. Set `OFFICE_ROOT=/path/to/worktree` to capture from an old commit.
- `tools/browser-runner/` is a standalone Fable/Vite browser harness for the canvas renderer. It loads DOCX files through `Office.Formats`, imports them into `RedBat.Word.API`, exposes `window.rbWordRunner`, and can write layout JSON plus canvas screenshots through Playwright.
- `tools/native-runner/` is a native DOCX-to-PNG snapshot harness. It loads DOCX bytes with the native `DocxReader`, maps the typed parse tree into `RedBat.Word.Document`, renders `word-native` through `../widgets` + `../renderer`, and writes a PNG.
- `tools/pdf-runner/` is the project PDF backend harness. It reuses the native typed DOCX import, calls `RedBat.Word.Pdf.toBytes`, and writes the resulting PDF.

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
doc	page	paragraph	line	start	end	x	y	text
```

Positions are reported in Word points relative to the page.

## Next Setup Tasks

- Add a native geometry export alongside the native PNG runner if we need numeric native line/paragraph positions, not just raster snapshots.
- Add a small comparison driver that runs Word TSV/PDF, browser JSON/PNG, native PNG, and project PDF for the same fixture into a single artifact folder.
- Resume parity fixes on tables/floats using the existing real-world fixtures and focused DOCX writer repros.
