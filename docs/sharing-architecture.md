# Sharing audit: one code base, three ways to run it

2026-11-11, `master` at `e32373f7`. Read-only: nothing in the repo was edited, built or run. Line counts are `wc -l` of the source files (comments included). Ranges inside a file such as "drawCell 3416-3978" come from reading where its definitions start and end, so they are approximate.

**Native spreadsheet follow-through (2026-11-11):** the S4–S6 correctness changes retain complete `Workbook` snapshots rather than sheet-array snapshots and use the shared `Recalc` engine. `Workbook.updateSheetModelWithTables` keeps model-level table changes and the formula registry consistent. The Wisp host accepts imported context with `createWithWorkbook` / `replaceWorkbook`; `exportWorkbook` writes current names, name metadata, visibility, active sheet and iteration, remapping stable local-name scopes only at the XLSX boundary. Sheet insertion, rename and deletion use workbook reference-rewriting operations and undo as complete transactions. The audit tables and line counts below remain the original snapshot; a portable controller extraction is still outstanding, not claimed by these correctness fixes. Verification is recorded in `systematic-review-2026-11-11.md`.

For Widgets, use `createStateFromWorkbook` / `replaceWorkbook` when opening imported context; `createState` is for fresh model arrays. Treat both hosts' sheet/name/index signals as UI projections and make content changes through their transaction functions, not direct signal assignments. The native gate now includes both host regression suites and the Wisp executable smoke.

**Browser dialog contract (C5–C6 follow-through):** `OfficeUiDialog` renders all eleven current field kinds and tabs without changing `DialogSpec`. Mount `renderDialog state` once and retain its mount disposer. With tabs, `Fields` is ignored as the current spec states; values from every tab survive switching and reach the button callback. Colour uses the browser-native picker; sliders have range and numeric inputs. Computed fields read the same value map.

Field edits remain dialog-local. `DialogButton.Action` owns validation and the editor transaction: validate before mutating, return `false` to retain the dialog and its edits, and invoke one undoable editor command on acceptance. Returning `false` is not a rollback of mutations made by the callback. Cancel must not apply those local values. A callback that opens another dialog does not have its replacement closed by the previous callback's return value.

The renderer names the dialog and controls, focuses on opening, contains focus, and restores the invoker on close/unmount. Tab headers support arrows/Home/End. Unmodified Enter activates the primary action except in multiline/select/radio/checkbox/colour controls or focused buttons, which retain their native behavior; composition is not submission. Escape activates the standard `dlg-cancel` button, or closes without invoking an arbitrary secondary action. This completes the renderer, not browser-shell command wiring or the parked foundation's expanded validation/spec API. The audit tables below remain historical.

The direction (`docs/continuation-roadmap.md`, *Three ways to run it*):
1. An embedded widget with no UI (redbat today).
2. The browser with our own UI and no server.
3. A .NET desktop app.

The widget should be shared close to 100%. The UI may differ per platform, but its commands, specs, state and validation should be shared, with thin renderers.

## 0. Summary

- **What is shared already:**
  - The **models** (Word `Document`, sheet `SpreadsheetModel` / `Workbook`).
  - The **layout** (Word `Layout.fs` / `WordLayout.fs`, sheet `Layout.fs`).
  - A **paint-command language** per editor (Word `Paint.PaintCommand`, sheet `Paint.PaintItem`).
  - The **formula engines**.
  - Most **editing operations as pure functions** (Word `Document.fs` / `WordStyles.fs`, sheet `SheetOps.fs` / `Input.fs`).
  - The **xlsx reader and writer**.
  - **Word PDF**.

  About 26k lines of Word and about 45k lines of the spreadsheet sit in netstandard projects that also build on .NET. The few `#if FABLE_COMPILER` branches in them are platform primitives: Map tree internals, WeakMap versus ConditionalWeakTable, `String.fromCodePoint`, and a per-thread table-plan scope.
- **What is not shared** is the controller: the code between the input events and the model, and between the model and the surface.
  - In the browser it is one large closure per editor:
    - Spreadsheet: the `Spreadsheet` class, 10,344 lines.
    - Word: `Input.create` plus `Renderer.fs` plus `Word.fs`, about 9,700 lines.
  - That code is browser-only:
    - In Word it sits behind `#if FABLE_COMPILER`.
    - In the spreadsheet it is written with `[<Emit>]` / `jsNative`. It compiles on .NET but throws there at run time.
  - Each .NET host writes a small controller of its own again, and there are four such hosts on two toolkits.
- **Word import and save are not shared either.**
  - The browser imports and saves through JS-object packages (`Docx.fs`, and `DocumentWriter.fs` reading `obj` with `?`).
  - .NET has a separate importer (`src/Word.Import`) and **no Word save at all**.
- **The spreadsheet paints cells twice.**
  - The browser uses DOM nodes with CSS (`Spreadsheet.fs` `drawCell`).
  - .NET and PDF use `Paint.fs`.
  - They already disagree. Example: numbers of General alignment are at the right in `Paint.generalAlign` but at the left in the DOM, which is the roadmap's open alpha bug.

## 1. Layers per editor

Legend:
- **S**: shared, compiled for both and working on both.
- **B**: browser-only, whether behind `#if FABLE_COMPILER`, dynamic JS (`Emit` / `jsNative` / `?`), or only ever used there.
- **N**: .NET-only.
- **D**: the same logic implemented more than once.

The .NET hosts come in pairs on two toolkits:
- `word-native` and `spreadsheet-native` are on **Rbt.Widgets** ("Composit"). The test runners use them.
- `word-wisp`, `spreadsheet-wisp` and `office-wisp` are on **Rbt.Wisp**, its replacement (see the header of `office-wisp/OfficeShell.fs`).

### 1.1 Word

| Layer | Shared | Browser-only | .NET-only | Duplicated |
|---|---|---|---|---|
| Model | `Rope` 429, `TextBuffer` 2,039, `Document` 6,077, `Shapes` 280, `WordStyles` 892, `Math` 427 = **10,144** | – | – | – |
| Import | Office.Formats `DocxReader` 7,190 and `DocumentProcessors` 764 (typed) | `Docx.fs` 3,468: imports the *JS objects* that `Main.fs` (1,931, the Fable API) makes of the processor's output, and exports back to JS objects | `src/Word.Import/WordImport.fs` 1,228 (net10.0): imports the typed `DocxReader` output | **D**: two importers from the same reader output (3,468 vs 1,228); the .NET one lags (no progressive load, no source splicing) |
| Save | – | `DocumentWriter.fs` 6,380 reads its input as `obj` (`item?Field` about 200 times) and returns `JS.Promise`; `writeFromEditorPackage` | **none**: no .NET host can save a docx (`word-wisp/Program.fs`: `SaveFile = ignore`) | – |
| Layout | `Layout` 7,630, `WordLayout` 2,088, `BorderGrid` 271, `LayoutFingerprint` 182 = **10,171** | Progressive layout, layout cache, stale-layout reuse and scheduling (`Renderer.fs` ~477-1230) | – | **D**: `layoutSettings` written three times (`Renderer.layoutSettingsFor`, `WordEditor.layoutSettings`, `WordRender.layoutSettings`) with **different values**. The native hosts scale the page to the widget width and use `LongTrailingSpaceFitSafety` 0.25 where the browser uses 3.5, so lines and pages break differently on .NET. |
| Painting: commands | `Paint.fs` 4,540: `PaintCommand` (text, underline, image, rects, lines, shapes, groups, mirror) and `collectDocumentCommands` / `Chunks` | – | – | – |
| Painting: backend | – | `Renderer.renderPaintCommand` and the canvas 2D state, about 250 lines (1549-1790), plus frame / band painting about 400 more | `WordEditor.toRendererCommand` (~90, Rbt.Widgets), `WordRender` 141 (Wisp) | **D** (justified: one per surface), but there are two .NET ones |
| Frames, dirty bands, page desk, viewport | – | `Renderer.fs` 1790-2600: `PaintFrame`, `computeDirtyRange`, bands, page sheets, `setViewport`; **about 85% of `Renderer.fs` does not touch the canvas** (the `ctx.Ctx` uses cluster in 1550-1790 and 2300-2430) | Native: no pages, no viewport; the whole document is laid out and collected every paint (`WordEditor.commands`) | – |
| Text measuring and shaping | `ShaperText` 258 (Rbt.Text.Engine and Rbt.OpenType, both netstandard and Fable): a `TextMeasurer` from font bytes; Word PDF uses it | `Renderer.measureTextStyled` (canvas `measureText` with per-family metric scales) unless the host installs measurers | Native hosts measure with `Rbt.Renderer.Text` and load TTFs from their own font folders | **D**: three measurers, so three sets of widths and line breaks. The font-folder and family logic is copied in `WordEditor.fs` (~70), `WordRender.fs` (~60), `SpreadsheetGrid.fs` (~210) and `GridRender.fs` (~60). |
| Input: keyboard | – | `Input.create` (981-2346, one closure): keys ~1450-1890 | `WordEditor.handleKey` (~70), `word-wisp handleKey` (~40), `office-wisp/KeyMap.fs` 85 (keycode to key) | **D** three times; the browser one is the only complete one (stories, text boxes, shapes, PageUp/Down by screen, column keeping) |
| Input: IME, text | – | Hidden proxy textarea, `compositionstart` / `end`, `beforeinput` (`Input.fs` ~1890-2030); `spellcheck=false` | Native: `CompositionStart / Update / End` ignored ("hosts own native IME"; `WordEditor.fs:1028`); commits only | – |
| Input: mouse | – | `mousedown` / move / up, drag of images and shapes, resize handles, double / triple click (~2055-2340); no touch | `handlePointer` in both natives (click count, image move) | **D** |
| Clipboard | `Document.copySelection` / `pasteClip` / `clipOfHtml` (model side) | `Clipboard.fs` 19, `Input.htmlClip` (DOM HTML parse), copy / cut / paste events | `Platform.clipboard` text only | Partly **D** (the natives copy text only) |
| Selection and caret | Selection in `Document`, `Paint.selectionCommands*` | `Renderer.caretRect` / `hitTest` / `shapeAt` / `noteAt` | `WordEditor.caretRect` / `hitTest` / word and paragraph ranges / visual-line moves; the same again in `word-wisp/EditorHost` | **D**: word boundaries in `Input.isWordChar` / `wordAroundCaret` vs `WordEditor.isWordChar` / `wordRangeAt`, with different rules (the browser counts `'` and `’` as part of a word) |
| Editing commands | `Document.fs` (`insertText`, `deleteBackward`, `splitBlock`, track variants, tables, fields, …), `WordStyles` toggles | The orchestration around them (pending format, typing runs, stories, text boxes, headers and footers) in `Input.fs` and `Word.fs` | `WordEditor.splitBlockPreservingStyles`, `insertStyledText`, `autoListAtCursor`, `deleteWord` (its own versions) | **D**: the native hosts reimplement Enter and auto-list |
| Undo | `Commands.fs` 105 (`CommandHistory`, `ICommand`) | Typing runs, story commands, `ExtendedCommand` (`Input.fs` 545-700) | Lists of `Document` snapshots (`WordEditor.undo`, `word-wisp EditorState.Undo`) | **D** |
| Scrolling / virtualisation | – | The host scrolls a container and calls `setViewport`; only the visible band is painted | `scrollWidget` (Rbt.Widgets), the Wisp scroll view; everything is collected | – |
| PDF | `Word.Pdf` 1,142: in the Fable build (`dist/word-pdf`) **and** .NET (`tools/pdf-runner`) | – | – | – |
| Host API | – | `Word.fs` 4,315 (`API` module; ~700 lines are JS-object conversion, the rest model and layout queries such as `getSelectionInformation`, `getRangeBounds`, page numbering and story handling) | – | Its queries are needed on .NET too (VBA `Range.Information`), but they are written against `Ren.*` / `Inp.*` |
| VBA object model | – | `WordVba.fs` 8-340 (through `API.Editor`) | `WordVba.fs` 340-683 | **D**: two halves of one file, 332 + 343 |
| UI specs | `OfficeUi/Specs.fs` 282, `WordUi.fs` 859 (menus, toolbar rows, 12 dialogs) | – | – | – |
| UI renderers | – | `FableUI` 705 (`Reactive` 149, `Dom` 165, `UI` 103, `OfficeUiDialog` 232: no `Tabs`, no section / colour / slider / text-area / radio fields) | `office-wisp/OfficeShell.fs` 198 (every field kind, tabs) | Two renderers (justified), but the DOM one is incomplete |
| Shell | – | `WordShell` 654 (uses `WordUi.menus` / `toolbarRows`), `WordHost` 784, `WordDialogs` 950 (**17 hand-written DOM dialogs**; find / replace, insert table and page setup also exist as specs), `WordViewer` 375; nothing mounts them (`docs/ui-audit.md`) | `word-wisp/Program.fs` 132 | **D**: `WordDialogs` vs the `WordUi` specs |

**Word totals, approximately:**
- Shared and working on both: model, layout, paint, specs, PDF, ShaperText ≈ **27k**.
- Browser-only: Renderer 3,023, Input ~2,390, Word.fs 4,315, Docx 3,468, WordVba ~332, Shell, Host, Dialogs and Viewer 2,763, FableUI 705 ≈ **17k**, plus `DocumentWriter` 6,380 that works only with JS objects.
- .NET-only: Word.Import 1,228, WordVba ~343, word-native 1,159, word-wisp 507 ≈ **3.2k**.

### 1.2 Spreadsheet

| Layer | Shared | Browser-only | .NET-only | Duplicated |
|---|---|---|---|---|
| Model | `Model` 5,121, `Workbook` 1,166, `NumberText` 344, `CellText` 1,570, `TableStyles` 252, `Pivot` 696, `ConditionalFormatting` 1,161 | – | – | – |
| Formulas / recalc | `Formula` 14,391, `Recalc` 2,379, `Calc` 726 | – | – | **D**: the browser uses `Recalc` (incremental, sort-aware); **both .NET hosts use the old `Calc.trySingleCellEdit` engine** (`SpreadsheetEditor.recalcEditedCell`, `spreadsheet-wisp/EditorHost`) |
| Import / save | `XlsxBridge` 3,080 (typed, both), Office.Formats `XlsxReader` 6,588 and `XlsxWriter` 5,294 | `XlsxImporter` 163 (buffers and promises around `XlsxBridge`: thin, as it should be) | `spreadsheet-wisp` writes with `XlsxWriter.writeWorkbookPackageBytes` | – |
| Editing commands | `SheetOps` 3,534, `Input` 655 (parse input, edit text, fill series) | **112 distinct `SheetOps` functions are called only from `Spreadsheet.fs`**; the orchestration around them (recalc after edit, book steps, sheet add / move / delete, tables, filters, CF rules, comments, links) is in the class, ~2,500 lines (514-3040) | The .NET hosts call **no** `SheetOps` function | **D**: native and Wisp write `writeCell`, `applyCellEdit`, `copySelection`, `pasteFromClipboard`, `resolveNavTarget`, `pageRowJump` and `ensureCellVisible` twice, nearly line for line (`SpreadsheetEditor.fs` vs `spreadsheet-wisp/EditorHost.fs`), and the browser a third time |
| Layout | `Layout.fs` 712: metrics, prefix sums, `hitTest`, `visibleCells`, freeze zones, `textOverflowSpan`, edges | – | – | Hit testing is mostly shared already: the browser uses `indexAtOffset` / `hitTest` / `metricsOf` |
| Painting: cells | `Paint.fs` 1,384: `buildCellPaintItems` (~710: fills, borders, table styles, CF, alignment with `generalAlign`, wrap via `wrapLines`, rotation, overflow), headers, `selectionPaintItems`, freeze dividers, drawings, chart scene items | **`drawCell` 3416-3978 (~560)**: the same cell look as **CSS on a DOM node per cell** (text-align, flex for vertical alignment, `white-space` for wrap, a legacy hard-coded table look `#4472C4` when no table look applies); headers 3978-4044; `render` 4044-4264 (node pools per pane); overlays (selection, active cell, fill handle) 4896-5112; outline 4704-4811; page breaks 4570-4632 | `SpreadsheetGrid.toRendererCommands` / `appendPaintItem` (~280, Rbt.Widgets); `GridRender` 259 (Wisp) | **D, the largest one**: about 1,100 lines of DOM cell, header and selection drawing reimplement what `Paint.fs` does. A painted path exists in the browser already, unused: `renderPortablePaintToCanvas` (1259-1420), `PortableCellGeometry`, `UseSharedLayoutGeometry`. |
| Painting: charts, slicers | `Charts.fs` 3,031 makes a `Scene` of primitives; `Slicers.fs` 441 | `Charts.sceneSvg` (SVG nodes); `renderDrawings` 4380-4570 | `Paint.scenePaintItems`; `tools/excel-native-runner/ChartRender.fs` 855 (a third chart painter, in a test tool) | **D** (scene to SVG vs scene to paint items; small) |
| Text measuring | `Layout.TextMeasurer` (a function) | Canvas `browserTextWidth` / `cellTextWidth` (Emit), `fitMeasure` for auto-fit | `SpreadsheetGrid.measureText` (Rbt.Renderer.Text, CJK fallback chain), `GridRender.measureText` | **D**; there is no shaper on the sheet side (`ShaperText` takes Word's `TextStyle`) |
| Input: keyboard | – | `handleKeyDown` 3173-3349 (~176) | `SpreadsheetEditor.handleKey` 476-640 (~164), `EditorHost.handleKey` 436-581 (~145) | **D** three times, each with different gaps (`ui-audit.md`: the browser has no Ctrl+B/I/U) |
| Input: IME, edit box | – | A DOM `<input>` over the active cell (`EditInput`) for typing and IME, plus the formula bar | Wisp / Widgets text field | – |
| Input: mouse | `Layout.hitTest` | Pointer handlers in `Mount` (6072-6745): drag select, fill handle, resize, picture drag, filter buttons | `handlePointer` ×2 | **D** (gestures) |
| Clipboard | `tsvOfRows` / `rowsOfTsv` are inside `Spreadsheet.fs` (private) | System clipboard, HTML table parse, `selectionHtml` | `Platform.clipboard` text, in-process | **D** (TSV logic is not shared) |
| Selection | `Model.Selection`, `Paint.selectionPaintItems` | DOM selection rectangles per pane | Painted | **D** (drawing) |
| Undo | `Model.Command` / `History` (functions `Undo` / `Redo`) | `pushValueNeutral`, `bookSteps` / `runBookStep` | Snapshots of the sheet array (`UndoStack: (SpreadsheetModel[] * int) list`) ×2 | **D** |
| Scrolling / virtualisation | `Layout.visibleCells` / `visibleRange` | A native scroll container, node pools recycled in `render`, frozen panes as sticky wrappers | `viewportWidget` with scroll offsets into `Paint.buildViewportPaintItems` | – |
| PDF / print | `Pagination` 287, `PrintHeaderFooter` 330 | **none** (alpha gap) | `Spreadsheet.Pdf` 1,831 (net10.0; not in `build:fable`, though it uses only `Paint` / `Layout` and `Rbt.Pdf`, which ships Fable sources); `Png.fs` 161 | – |
| Host API | – | Members 6745-10343 (~3,600, JS-shaped `obj` arguments), `HostShapes` 1,308 (JS object conversion), `Tabs` 184 | – | – |
| VBA | `SpreadsheetVba` 1,137 (partly Emit) | – | `spreadsheet-wisp` runs it | – |
| UI specs | `SpreadsheetUi.fs` 1,033 (menus, toolbar rows, formula bar, sheet tabs, 13 dialogs) | – | – | – |
| UI renderers / shell | – | `Toolbar.fs` 621 (**its own spec types** `MenuItem` / `Menu` / `ToolbarItem`, not `Rbt.Office.Ui`); legacy dialogs inside the component: find / replace 5391-5668 (~277), conditional formatting 5668-6049 (~381), filter dropdown 5166-5391 (~225), validation list | `spreadsheet-wisp/Program.fs` 480 with `OfficeShell` | **D**: `Toolbar.fs` types vs `OfficeUi`; the legacy dialogs vs `SpreadsheetUi.findReplaceDialog` and filter specs |
| `src/Grid` | `Dom.fs` 161 is the only part used (DOM helpers) | `Grid.fs` 473, `Hooks` 194 and `DataSource` 103 are compiled but used by nothing; `Types`, `Rendering`, `VirtualScroll`, `GridLogic` and `Editing` (557) are not in the fsproj | – | Dead code |

**Spreadsheet totals, approximately:**
- Shared and working on both ≈ **45k**.
- Browser-only: `Spreadsheet.fs` 10,344, `HostShapes` 1,308, `Toolbar` 621, `Tabs` 184, `XlsxImporter` 163, `Grid` 931 compiled ≈ **13.5k**.
- .NET-only: native 1,502, Wisp 1,353, `Spreadsheet.Pdf` 1,831, `Png` 161 ≈ **4.8k**.

`Spreadsheet.fs` by part (line ranges):

| Part | Lines | What it is | Portable? |
|---|---|---|---|
| 1-513 | ~510 | Emit helpers, clipboard TSV / HTML, node types | TSV yes; the rest no |
| 514-3040 | ~2,500 | State, recalc orchestration, paste, fill, format, sheets and book undo, filters, autofit | **Yes, almost all** (model logic behind DOM-typed signatures) |
| 3040-3416 | ~375 | Edit start and finish, key handler, undo / redo | Yes (with an input event type) |
| 3416-4044 | ~630 | `drawCell`, headers | No: replaced by `Paint.fs` |
| 4044-5117 | ~1,070 | Render passes, drawings, pages, outline, overlays | Mostly replaced by `Paint.fs`; drawings partly |
| 5117-6049 | ~930 | Validation list, filter dropdown, find / replace and CF dialogs | Become specs and a UI layer |
| 6049-6745 | ~700 | Scroll, `Mount` (DOM construction, event wiring) | Browser adapter |
| 6745-10343 | ~3,600 | Host API members | Typed core plus a thin JS facade |

## 2. The external Rbt libraries

| Package (version pinned) | What it gives | Fable? | Used by |
|---|---|---|---|
| **Rbt.Text.Engine** 0.1.3, **Rbt.OpenType.Core** 0.1.3 (rbt-text; plus Rbt.Shaper.Core, Rbt.Bidi.Core, Rbt.Unicode.Core transitively) | Shaping, advances, `measureRange`, OpenType parsing (cmap, glyf, CFF, GDEF / GSUB, COLR…), glyph outlines | **Yes**: `fable/` sources and netstandard2.0 | `src/ShaperText` (Word only), so also Word.Pdf and the browser Word host when it installs ShaperText measurers |
| **Rbt.Renderer** 0.1.4 | The `Cmd` scene language (FillRect, paths, groups, clip, glyph runs, images…), its rasteriser and GPU backends, `Text.FontData` | Its `Rbt.Renderer.Core` (Cmd, Types, Text, Pcf) has `fable/` sources and netstandard; **rbt-office references only the net10.0 `Rbt.Renderer`** | The native and Wisp hosts' `toRendererCommand` / `appendPaintItem` / `GridRender` |
| **Rbt.Widgets** 0.1.2 (+ `Rbt.Widgets.Platform`) | The "Composit" retained widget toolkit: `canvas`, scroll widgets, `Platform.clipboard`, key / pointer events | No (net10.0) | `word-native`, `spreadsheet-native`, the native, PDF and Excel test runners |
| **Rbt.Wisp** 0.1.0 | A Solid-style reactive UI on Rbt.Renderer (`signal`, `refSignal`, `dyn`, layout, the Atrium theme, modal, menu, segmented, slider…) | No (net10.0) | `word-wisp`, `spreadsheet-wisp`, `office-wisp` (the spec renderer) |
| **Rbt.Pdf** 0.1.4 | A PDF writer (Doc, Draw, Color, Font, Image) | **Yes**: `fable/Rbt.Pdf.fs` | `Word.Pdf` (both platforms), `Spreadsheet.Pdf` (.NET only by build choice) |
| **Rbt.Vb** 0.1.2 | VBA / VBScript interpreter | Yes | `WordVba`, `SpreadsheetVba` |

Note: `FableUI.Reactive` (browser) and `Wisp.Reactive` (.NET) are the same model (signals, effects, computations) written twice in two repositories.

## 3. What a host of each mode has to write today

**Mode 1, embedded with no UI (redbat):**
- **Spreadsheet:**
  - The host calls `createSpreadsheet(container)`. The component builds its own DOM grid, scroll container, edit input and legacy dialogs. Its Ctrl+F dialog even opens over redbat's own (`ui-audit.md`).
  - The host builds every menu, toolbar, context menu and dialog itself in Solid (`SpreadsheetEditor.tsx`, `SheetDialogs.tsx`) and drives the ~3,600-line JS host API.
- **Word:**
  - The host gives a canvas and a scroll container, and calls `setViewport` on every scroll or resize.
  - It sizes a spacer from `getDocumentSize` and loads fonts.
  - It installs ShaperText measurers and painters (`setTextMeasurersForFaces`, `setTextPathPainterFactory`) and decodes images (`setImageCache`).
  - It reads the docx with Office.Formats and calls `importFromDocxPackage`, and exports with `exportDocumentPackage` and `writeFromEditorPackage`.
  - It builds the whole UI.
- **The cost:** redbat has no command or state source of ours to bind its UI to. Every enabled / checked state is asked of the host API one by one.

**Mode 2, the browser with our UI:**
- `WordShell.mount` and `createSpreadsheetShell` / `createFormulaBar` exist, but nothing mounts them.
- They are incomplete:
  - The DOM spec renderer throws on 5 field kinds and ignores tabs.
  - The shell's "$" / "%" / colour buttons send commands `ApplyFormat` rejects.
- A host would write little, but what it gets is prototype UI.

**Mode 3, .NET desktop:** a host has to write the whole controller. `word-wisp` is 507 lines plus `office-wisp` 283. With it you get:
- One continuous page, laid out with other settings than the browser's.
- No page frames, no virtualisation, no headers / footers / notes editing, no shapes, no tracked changes.
- Snapshot undo, text-only clipboard and no IME preedit.
- **No docx save.**

`spreadsheet-wisp` is 1,353 lines, with:
- The old `Calc` engine and none of the 112 `SheetOps` operations the browser wires.
- No filters, CF rules, comments, sheet operations or charts editing beyond what `Program.fs` wires by hand.

Everything these hosts do, the browser also does separately and more completely:
- Key maps, pointer gestures, word boundaries, visual-line moves, caret geometry and Enter / auto-list.
- Undo, clipboard, font loading, layout settings and the paint-command converter.

## 4. The biggest divergences, ranked by how much code a fix unifies

1. **Spreadsheet controller and DOM grid vs the painted grid and the two .NET controllers.**
   - What it covers:
     - The ~2,900 portable lines of `Spreadsheet.fs`: state, orchestration, keys and edit.
     - The ~1,700 lines of DOM drawing that `Paint.fs` already covers: cells, headers, overlays, outline, pages.
     - The ~2,000 lines of native and Wisp controllers (`SpreadsheetEditor.fs` + `EditorHost.fs`), which also use the old `Calc` engine.
   - About **6-7k lines** collapse into one portable `SheetView` controller plus a canvas painter (~300).
   - It fixes General alignment and the legacy table look for free, and gives the .NET host every feature of the browser.
2. **The Word view and input controller (`Renderer.fs` + `Input.fs` + the query half of `Word.fs`).**
   - About 85% of `Renderer.fs` (~2,500), most of `Input.create` (~1,600 of 2,390) and ~2,000 lines of `Word.fs` queries are platform-neutral logic locked behind `#if FABLE_COMPILER`.
   - That logic covers progressive layout, dirty bands, hit tests, caret, shapes, stories, typing runs, pending format and page numbering.
   - The .NET hosts reimplement subsets of it (1,159 + 507) with diverging layout settings.
   - About **6k lines** to one `WordView` / `WordInput`, and the same pages on both platforms.
3. **Word import and save through JS objects.**
   - `Docx.fs` (3,468) plus the `obj`-reading front of `DocumentWriter` (~1,500 of 6,380 touch `?`) vs `WordImport.fs` (1,228), and no .NET save.
   - A typed handover (the typed `EditorProcessor` output to `Document`, and `Document` to a typed writer package) unifies about **4-5k lines**.
   - It gives .NET a docx save, and it removes a second importer that lags.
4. **Text measuring.**
   - There are four measurers: browser canvas for both editors, `Rbt.Renderer.Text` in four .NET hosts, ShaperText (Word, PDF), and Spreadsheet.Pdf's font metrics.
   - The font-folder and family-key code is copied four times (~400 lines).
   - Few lines unify (about **0.6-1k**), but it decides whether the same document breaks the same way on both platforms, which is the precondition for "shared close to 100%". So it is ranked above smaller code wins.
5. **The UI layers.**
   - `WordDialogs` (950, hand-written DOM dialogs, three of which exist as specs too).
   - `Toolbar.fs` (621, its own spec types).
   - The legacy in-component dialogs (~880: find / replace, CF, filter dropdown).
   - The incomplete DOM spec renderer (232) vs the complete Wisp one (198).
   - About **2.5k lines** to specs plus two renderers, and a command registry that mode 1 hosts can bind to.

Smaller ones:
- Two .NET toolkits with near-copy hosts: Widgets `word-native` / `spreadsheet-native` vs Wisp. About 2.6k lines of host code exist twice; retiring one halves the native surface.
- `WordVba` in two halves (332 / 343).
- Two undo designs per editor (commands in the browser, snapshots on .NET).
- Chart scene to SVG vs scene to paint items, plus a third painter in `excel-native-runner/ChartRender.fs` (855).
- `src/Grid` dead code (770 compiled, 557 not).
- `Spreadsheet.Pdf` kept out of the Fable build.

## 5. Target architecture

```
        +---------------------------------------------------------------+
        | Host UI (per platform, thin)                                  |
        |  browser: FableUI DOM renderer | .NET: Wisp renderer |        |
        |  or the embedding host's own (redbat Solid)                   |
        +----------------------^----------------------------------------+
                               | command registry, specs, view state (signals)
        +----------------------+----------------------------------------+
        | UI core (netstandard, Fable-safe)  src/OfficeUi + WordUi/SheetUi
        |  Command { Id; Label; Shortcut; Enabled; Checked; Run }        |
        |  Menu/Toolbar/Dialog specs referencing command ids,           |
        |  dialog state + validation (Map<string,string> -> errors)     |
        +----------------------^----------------------------------------+
                               | typed editor API (what Word.fs/Spreadsheet members wrap today)
        +----------------------+----------------------------------------+
        | Editor core (netstandard, Fable-safe)                          |
        |  Model | Formats (typed in/out) | Layout | Paint commands      |
        |  Text (FontProvider -> shaper/measurer, Rbt.Text.Engine)       |
        |  Controller: WordView / SheetView                              |
        |    in : InputEvent (keys, text/composition, pointer, wheel,    |
        |         focus, clipboard payloads), viewport size/scroll/zoom  |
        |    out: model changes + undo, Scene (layers, dirty bands),     |
        |         caret / IME rect, cursor shape, a11y snapshot,         |
        |         clipboard payloads, notifications                      |
        +----------------------^----------------------------------------+
                               | the platform boundary
        +----------------------+----------------------------------------+
        | Platform adapters (thin)                                       |
        |  browser: Canvas2D painter, hidden textarea (IME, paste),      |
        |    scroll container or own scrollbars, a11y mirror DOM,        |
        |    rAF/idle scheduler, Clipboard API, FontFace loading,        |
        |    JS facade (HostShapes, Word.fs API) over the typed API      |
        |  .NET: Rbt.Renderer Cmd converter, Wisp/Platform events,       |
        |    timers, clipboard, font files                               |
        +---------------------------------------------------------------+
```

**Where the boundary sits.** A host provides exactly the following:
- `ISurface`:
  - `size`, `devicePixelRatio`.
  - `paint(band, commands)`, from `Paint.PaintCommand` / `PaintItem` (the existing languages: no new one).
  - `decodeImage(bytes) -> handle`.
  - `invalidate()`.
- `InputEvent` delivery. The union:
  - `Key(key, code, mods, repeat)`.
  - `TextInput(text)` and `Composition(Start | Update text | End text)`.
  - `Pointer(Down | Move | Up, x, y, buttons, clicks, pointerType)`.
  - `Wheel` and `Focus`.
  - `Paste(text, html, imageBytes)`.
  - `ContextMenu(x, y)`.
- `IPlatform`:
  - `now`, `scheduleIdle(budget)` (progressive layout already slices by budget).
  - `clipboardWrite(text, html)`.
  - `fontBytes(family, bold, italic)`.
  - `textInputAt(rect)`, which places the IME / proxy input at the caret.

**What stays per platform:**
- The painter: ~250-300 lines each, of which the browser one exists (`renderPaintCommand`, `renderPortablePaintToCanvas`) and the .NET ones exist (the `toRendererCommand` family).
- The event translation (~150), the IME element, the scroll container and the a11y mirror.
- On .NET: the toolkit host window.

**The UI.**
- One command registry per editor, built from the typed editor API: ids such as `format.bold` and `sheet.insertRow`, with enabled / checked computed from the controller's state.
- `WordUi` / `SpreadsheetUi` menus, toolbar rows and dialogs refer to commands instead of closures where they can.
- Dialog specs gain `Tabs` (already in the type), `Validate: Map -> (fieldId * message) list`, and an `Initial` read from the model.
- Two renderers of the same specs:
  - FableUI DOM: complete `OfficeUiDialog` with tabs and all eleven field kinds.
  - Wisp: `OfficeShell`, already complete.
- Mode 1 hosts (redbat) get the registry and state as plain JS (`getCommands()`, `onCommandStateChange`) to bind their own Solid UI. They may also use our specs' data (the option lists, the validation) without our renderers.
- Bespoke surfaces (the filter panel, CF rule manager, comment threads) are allowed as per-platform widgets, but their state and validation live in the UI core.

**Migration rules:**
- Pure logic is never written inside `#if FABLE_COMPILER`. A branch is allowed only around a platform primitive of a few lines (as today in `Model.fs`, `TextBuffer.fs`, `Layout.fs`).
- `Emit` / `jsNative` appear only in adapter modules.

## 6. Migration plan

Each step is shippable on its own:
- It keeps the JS host API (redbat) unchanged unless the step says so.
- It keeps `tools/word-browser` and `tools/xlsx-browser` green.
- It adds a .NET check where one is cheap.

Sizes are rough: S under ~300 lines changed, M ~300-1,500, L ~1,500-4,000, XL more.

1. **One cell look for both sheet painters (S-M).** Done on branch `sheet/shared-cell-look` (2026-11-11): `Paint.cellLook` / `CellLook`, the shared edges `Layout.edgeBetween` / `edgeRuns` / `tableEntryAt`; `drawCell` maps the look to CSS; xlsx-browser `cell-look-shared` compares every cell of its fixture with the look.
   - Extract the per-cell resolved look into `Paint.fs`: fill, font, colour, horizontal alignment including `generalAlign`, vertical alignment, wrap, rotation, indent, the table look, CF fill / font.
   - Make `drawCell` read it and map it to CSS.
   - This fixes General numbers drawn at the left (an alpha item) and removes the legacy `#4472C4` table look branch.
   - Harness: the existing cell-style checks plus a new "number of General alignment at the right" check.
2. **Spreadsheet PDF in the browser (S-M; an alpha item).** Done on branch `sheet/pdf-in-browser` (2026-11-11): `dist/spreadsheet-pdf`, `ExportPdf` / `PrintPdf` / `SetPdfFonts` (`docs/spreadsheet-host-api.md` §PDF and printing). What the audit missed: Rbt.Pdf's Fable sources need `Rbt.OpenType.Core` (its `Inflate`); compiled from the net10.0 project, three `TryParse` calls of `DocxReader` were ambiguous (span overloads) and `Char.ConvertToUtf32` in `Pdf.fs` is not in Fable; Rbt.Pdf's browser build cannot decode PNG (the component has the browser turn PNG pictures into JPEG); and a Fable bundle of a project carries its own copy of its references (a second Model, Paint, fable-library ... of 2.5 MB loaded beside the grid's), which `scripts/share-bundle-modules.mjs` points back at `dist/spreadsheet`.
   - Add `src/Spreadsheet.Pdf` to `build:fable`. It uses only `Paint` / `Layout` / `Pagination` and Rbt.Pdf, which has Fable sources.
   - Replace `Png.fs` use with a browser image path, and expose `ExportPdf` on the component.
   - It proves the "paint items to a second backend" path in the browser before the grid depends on it.
3. **A painted grid in the browser, behind a flag (M-L).**
   - Paint the body cells, headers, grid lines, selection and freeze dividers on a canvas from `Paint.buildViewportPaintItems` and `selectionPaintItems`. The code exists as `renderPortablePaintToCanvas`.
   - Keep these as DOM for now:
     - The scroll container (a spacer for extent).
     - The edit input, comment box, list and filter buttons, pictures and charts.
   - Give the harness a geometry API in place of the DOM queries: `PortableCellGeometry` and `GetCellEdges` exist. About 112 lines of `tools/xlsx-browser/*.js` query `.spreadsheet-cell` / `-col-header` / `-row-header` / `-edge` / `-cf-*`.
   - Run the harness in both modes, and gate on `perf.mjs` / `bigsheet.mjs` / `grid-scroll-frames`.
   - Shippable with the flag off. Redbat is unaffected until it opts in.
4. **One text stack (M).**
   - Generalise `src/ShaperText` into a `FontProvider` (font bytes by family / bold / italic, with fallbacks and the CJK chain from `SpreadsheetGrid.fs`) and measurers for both Word's `TextStyle` and the sheet's `TextMeasurer`.
   - Use it in the four .NET hosts, deleting four font loaders.
   - Make it the default measurer of the browser spreadsheet's painted mode and of Word when the host gives font bytes.
   - Unify `layoutSettings` (one function in `Rbt.Word.Layout`; the .NET hosts stop scaling the page).
   - Check: the same document gives the same `LayoutFingerprint` in the browser runner and a .NET runner with the same fonts.
5. **Extract `SheetView` from `Spreadsheet.fs` (L).**
   - Move parts 514-3416 (state, orchestration, undo, keys, edit) into a portable controller over `InputEvent`, with typed arguments. The class keeps `Mount`, the DOM adapter and the JS member facade, which now call it.
   - Then re-host `spreadsheet-wisp` on `SheetView`. That deletes its `EditorHost` (614) and the old `Calc` use, and it gets every `SheetOps` feature.
   - The JS API stays byte-compatible. The xlsx-browser `host-api.js` / `sheet-api.js` suites are the guard.
6. **The painted grid as the default; delete the DOM cell paths (M).**
   - Remove `drawCell`, the header nodes, DOM overlays, outline and page-break nodes (~1,700 lines), and `src/Grid` except `Dom.fs`.
   - Draw charts through `scenePaintItems` on the canvas, keeping SVG only for export if wanted.
   - Prerequisite: step 3's perf parity and redbat's adoption harness (1104/1104) on the painted grid.
7. **Lift `WordView` out of `Renderer.fs` (L).**
   - Move everything except the canvas calls out of `#if FABLE_COMPILER` into a portable module:
     - Progressive layout, caches, frames, dirty bands, viewport mapping.
     - Hit tests, caret, shapes, notes, page frames.
   - `Renderer.fs` keeps `renderPaintCommand`, canvas sizing and images (~400).
   - Expect .NET compile errors in code that was never type-checked there (see risks).
   - Re-host `word-wisp` on it. It gets pages, virtualisation and the same layout.
8. **Split `Input.fs` into `WordInput` plus a browser proxy (L).**
   - Portable: the key-to-command map, typing runs, pending format, stories, text box and shape gestures, word boundaries.
   - Browser: the textarea, composition events, clipboard events and HTML parsing.
   - Delete the native key handlers.
   - Move the queries of `Word.fs` (`getSelectionInformation`, `getRangeBounds`, page numbers) to the typed API, so `WordVba` needs one half only.
9. **Typed Word import and save (L-XL).**
   - Make `Docx.fs` import from the typed `EditorProcessor` output, the one `WordImport.fs` reads. Keep one importer; the JS package is made from the typed one only at the API edge.
   - Give `DocumentWriter` a typed input record and a byte-array result, wrapped in a Promise only in `Main.fs`.
   - Guards: `tools/docx-roundtrip`, the open-all Word checks, `save-architecture.md`'s splice tests. Then add docx save to the .NET host.
   - This is the riskiest step for "no save loses anything", so it is done after the alpha gates or behind the existing one-save work.
10. **A command registry and complete spec renderers (M-L).**
    - `Command` in `OfficeUi`; `WordUi` / `SpreadsheetUi` build from it.
    - `OfficeUiDialog` gets tabs and all field kinds.
    - Port `WordDialogs` and the legacy sheet dialogs to specs, and make `Toolbar.fs` use `Rbt.Office.Ui` types.
    - Export `getCommands` / state to JS for mode 1 hosts.
    - This can run in parallel with steps 5-9, and it is the backbone of alpha item 7 (a complete UI).
11. **One .NET toolkit (M).**
    - Move the test runners (`native-runner`, `pdf-runner`, `excel-native-runner`, `spreadsheet-pdf-runner`) from the Rbt.Widgets hosts to headless `Cmd` rendering from the shared controllers.
    - Delete `word-native` and `spreadsheet-native` (2,661 lines).

Order rationale:
- Steps 1-3 are cheap, fix alpha bugs and de-risk the biggest divergence while the browser keeps its DOM path.
- Step 4 is a precondition for "the same pages on .NET".
- Steps 5-8 are the controller moves.
- Step 9 is format work with its own gates.
- Steps 10-11 can interleave.

## 7. Risks

- **A painted grid's performance in the browser.**
  - The DOM grid reuses nodes and lets the compositor scroll; a canvas repaints. On a scroll frame of the 846,000-cell sheet that is the whole visible body.
  - `Paint.fs` builds F# lists of records per cell. In Fable that means allocation and GC (cf. the `moveRowBlocks` and Map-node costs noted in the roadmap).
  - Mitigations:
    - Arrays in place of lists on the hot path.
    - A per-pane tile or band cache that blits on scroll and paints only the exposed strip (Word's `computeDirtyRange` / band painting is the model).
    - Cached text widths per (string, font).
  - Gate on `grid-scroll-frames` and `perf.mjs`, run alone, not under load.
- **Text.**
  - Canvas text has no subpixel layout like CSS's. Text that runs over neighbours, rotated text, and rich runs in a cell must come out of `Paint.wrapLines` and `textOverflowSpan` exactly as they do now in CSS.
  - Fonts must be loaded before painting (FontFace API); otherwise widths change after the first frame.
  - With ShaperText, shaping cost per cell in JS needs a cache.
- **Accessibility.**
  - Today there is none. `aria-` / `role` appear 0 times in `Spreadsheet.fs`, once in `Input.fs`, twice in `WordShell`, and the alpha lists no screen-reader support. Painting does not make that worse, but it closes the cheap door of DOM cells.
  - The plan should include an a11y snapshot from the controllers: the active cell and its neighbours, the current paragraph and line. A browser adapter mirrors it as an ARIA grid of the visible window and an `aria-live` region.
  - Browser find-in-page (Ctrl+F) and text selection of cell text also go away with canvas.
- **IME, text input and spell check.**
  - Composition must stay in a real focused element at the caret: Word's proxy textarea, the sheet's edit input. Both exist and should survive as adapters, not move into the core.
  - Browser spell check needs an editable element with the text, or our own checker (the alpha decided on our own squiggles, which fits the painted model).
  - On .NET, Rbt.Widgets / Wisp deliver only committed text (`WordEditor.fs:1028`), so preedit needs work in the Platform package (another repo).
  - Touch (an alpha item) belongs in the shared `InputEvent` and gesture code from the start, rather than per adapter.
- **Fable traps when lifting code out of `#if FABLE_COMPILER`.**
  - Code there was never type-checked on .NET. Expect errors and behaviour differences, such as `int` overflow and `char` / code-point handling, and `Dictionary` of records or tuples, which hashes structurally and slowly in Fable.
  - `Emit` / `jsNative` code compiles on .NET and throws at run time (`Docx.fs`, `DocumentWriter`, `Spreadsheet.fs`, `HostShapes`). A .NET build passing proves nothing; .NET tests must exercise the paths.
  - Reference-equality caches (`obj.ReferenceEquals` on models, WeakMap vs ConditionalWeakTable) are load-bearing for performance on both. Keep them as primitives behind small `#if`.
  - `Model.fs` reads fable-library's Map tree internals. A Fable upgrade can break it silently; keep its .NET twin tested.
  - `[<AttachMembers>]` classes define the JS API shape. Moving logic must not change member names, overloads or optional `obj` arguments that redbat calls.
  - Threads: .NET hosts lay out on several threads (`TablePlanScopeHolder`). Portable controllers must not add global mutable state; `Paint.fs` has `drawingSink` and `SpreadsheetGrid.fs` has module-level caches today.
- **Coupling of hosts and harnesses to the DOM.**
  - About 112 lines of the xlsx harness and redbat's CSS address `.spreadsheet-*` nodes.
  - Steps 3 and 6 need a geometry and test API first, and a redbat adoption run.
- **Scope pressure.** Steps 5, 7, 8 and 9 are large refactors of the files with the most traffic (`Spreadsheet.fs`, `Input.fs`, `Renderer.fs`, `DocumentWriter.fs`), during alpha work.
  - Do them as moves with no behaviour change, one region per branch, and rebase often.
  - Use the slot script for every heavy run (memory: several agents at once swap the machine).
