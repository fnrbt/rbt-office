# Word fidelity parity tracker — vs. eigenpal/docx-editor

**Yardstick:** [eigenpal/docx-editor](https://github.com/eigenpal/docx-editor)
(~1.6k★, very active, React/Vue/Nuxt, MIT). Goal: **match or exceed on every
relevant dimension.** This is a living doc — append findings as we discover
them. Keep claims grounded with `file:line` evidence on both sides.

Reference clone (for source spelunking): `/tmp/docx-editor-cmp` (shallow).
Re-clone with `git clone --depth 1 https://github.com/eigenpal/docx-editor`.

## How to use this doc
- One row per dimension. When you verify something, set the status + cite code.
- Prefer **real comparisons** (render the same fixture, or read both
  implementations) over marketing claims — see [[feedback-parity-comparison-like-for-like]].
- When a 🔴/🟡 gets closed, flip it to 🟢 and link the commit.

## Status key (from *our* perspective)
- 🟢 **match/exceed** — we're at or past them
- 🟡 **partial / behind** — we have something, but weaker
- 🔴 **missing** — they have it, we don't
- 🔵 **we lead** — we have it, they don't
- ⬜ **unverified** — claimed, not yet checked in code

---

## Scorecard

### A. On-screen rendering & layout fidelity
| Dimension | Status | Us | Them | Notes |
|---|---|---|---|---|
| Text measurement technique | 🟢 | hidden-canvas `measureText` + own line-break (`Paint.fs`/`Layout.fs`) | hidden-canvas `ctx.measureText` + own line-break (`layout-painter/renderParagraph/line.ts:241`) | **Same core technique.** Neither uses contenteditable for visuals. |
| Render targets | 🔵 | canvas + **true PDF (font embed + shaper)** + native widgets ([[project-dual-render-paths]]) | DOM only | Our PDF/native paths are structural advantages they can't match from DOM. |
| **Justified text** | 🟢 | **FIXED** — `Layout.fs justifyLineToEdge` distributes inter-word slack so all-but-last lines end flush at the right margin (handles first-line-indent; skips tabs/images/floats). In shared Layout.fs ⟹ canvas+PDF+native. Validated 3-way on footnote-placement. | computes real justify spacing | Was the cheapest win; now at parity. |
| Tables: vmerge / split-across-page / cut borders | 🟢 | yes | `TableFragment.fromRow/toRow`, `data-vmerge-continuation`, cut-border rule | Both mature here. valign+rowspan verified 3-way = tie. |
| Table cell **pattern shading** (pct5/pct10/stripe) | 🔵 | renders pct/stripe fills close to Word | **renders shaded cells as plain white — no pattern support** | Verified 3-way; big visible win. |
| Table **gridspan bottom-fit** | 🟢 | renders correctly (verified tall render: gridspan A/B/C/D at page-2 bottom) | renders correctly | **Earlier "blank" was a harness artifact** — 1600px canvas clipped this multi-page (A4) doc. Harness now auto-sizes. NOT a renderer bug. |
| Table styled-grid cell sizing | 🟢 | **matches Word** (measured: row pitch 31px both, cols 200/201px both) | matches Word | Earlier "over-sized" was an eyeball miscall; pixel-measured = match. |
| Default line / paragraph spacing | 🔵 | **matches Word (measured)** | runs **tighter** than Word | **Pixel-measured** on running-header: para gap Word=29px, ours=29px, docx-editor=18px (~38% tight). their `fix-default-line-spacing`. |
| Footnotes / endnotes | 🟢 | at parity ([[project-footnote-rendering-gap]]) | `layout-painter/renderPage/footnotes.ts` | |
| Multi-column sections | 🟢 | real column flow ([[project-multi-column-sections]]) | `layout-engine/section-breaks.ts` | Ours deferred: balancing, table-borders/floats in columns. |
| List default indent (no per-level cascade) | 🔵 | matches Word — all levels at same default indent (validated 3-way) | **over-indents** each nested level (diverges from Word); their roadmap `list-operations-fidelity` | We're closer to Word on `list-default-indent.docx`. |
| **Floating image wrap** (square / top-bottom) | 🟢 | 3-way: top-bottom **matches Word**; square competitive | float-zone pipeline (`measureBlocksWithFloats`), but **mispositions** float on `float-image-wrap-top-bottom` | Revised from earlier 🔴 — rendered evidence shows ours ≥ docx-editor on these repros. Re-audit tight/through. |
| Tab stops (left/center/right/decimal/multiple) | 🟢 | **FIXED** — all stop types + decimal alignment (`Layout.tabWidth`) | yes | verified `tab-stops.docx`. |
| Underline styles (double/dotted/dash/wavy/thick + colour) | 🟢 | **FIXED** — distinct styles + colour (canvas+PDF) | yes (CSS text-decoration) | verified `underline-styles.docx`. |
| Hidden text (`w:vanish`) | 🟢 | **FIXED** — hidden text not laid out/rendered (Layout skip) | shows it grayed | `hidden-text.docx`. |
| Run effects (`w:dstrike`/`w:position`/`w:em`) | 🟢 | **FIXED** — double-strike, baseline raise/lower, emphasis dots (canvas+PDF) | rendered | `run-effects.docx`. `w:sym` deferred (needs symbol font); `w:kern`/`w:effect` low-value. |
| Math (OMML) render | 🟢 | **FIXED** — real 2D math layout (`Math.fs`): self-contained OMML parser + box layout (fractions, sub/superscripts, radicals, n-ary ∑/∫/∏ with limits, delimiters, functions, accents, bars, matrices) → glyph/rule primitives → emitted as DrawText/FillRect, so **all render paths** draw it. Atomic `MathSegment` in the line model. | renders equation (`MathExtension`) | `omml-equations.docx`, verified canvas+PDF (quadratic formula, etc.). Refinement: inline display-math line spacing is a touch tight on canvas. |
| Fields (PAGE/NUMPAGES/DATE/TIME/TOC/SEQ/REF/PAGEREF/STYLEREF) | 🟢 | **computed** in `resultForField` (body) + `runningFieldResult` (header/footer); STYLEREF = nearest-heading text, DATE/TIME with `\@` picture | their `fieldParser` handles | `fields-styleref-date-time.docx`. Remaining 🔴: AUTHOR/FILENAME (need core.xml→editor-Document plumbing; `Creator` is parsed in DocxReader but not threaded); header STYLEREF (needs per-page heading context). |
| SDT control types + dataBinding | 🟢 | **checkbox renders as real ☒/☐ glyph** from `w14:checkbox`/`w14:checked` (not the symbol-font char); dropdown/combo/date/plain-text surface their content; id/tag/alias preserved | full types + interactive editing | `sdt-controls.docx` verified. Interactive editing (open dropdown, toggle) is still their lead. |
| Preset-geometry shapes (`prstGeom` fill/gradient/stroke) | 🟢 | **FIXED** — `a:prstGeom` (60+ presets: ellipse/roundRect/triangle/diamond/polygons/stars/arrows/chevron/flowchart…) + `a:custGeom` freeform paths + `a:gradFill` linear/radial gradients, parsed (DocxReader) → `ShapeGeometry`/`GradientSpec` (Document) → `DrawShape` paint command → **all three render paths** (canvas full gradient, PDF gradient≈midpoint-solid, native `FillPath` full gradient). New `Shapes.fs` path library. | `shapeParser` | `autoshapes-gallery.docx`, verified canvas+PDF. Unmapped presets fall back to rect; arcTo in custGeom dropped (approx). |
| RTL / bidi paragraphs | 🟡 | **canvas FIXED** — `w:bidi` → right-align + token reorder (browser handles in-token glyph order). PDF/native glyph path still needs `../text/bidi` wired. | right-aligns | engine `../text/bidi` exists; canvas done, PDF follow-up. |
| Paragraph/table borders + shading | 🟡 | borders + pct-shading | active `openspec/changes/paragraph-border-rendering`, `table-rendering-fidelity` | Verify edge cases head-to-head. |
| Watermarks (VML) | 🔴 | ✗ | `docx/vmlWatermarkParser.ts` + `serializer/vmlWatermarkSerializer.ts` + `watermarkApi.ts` | |
| Equations (OMML) render | 🟢 | **laid out + rendered** (`Math.fs`), not just XML-preserved — see the Math (OMML) row above | `prosemirror/extensions/nodes/MathExtension.ts` (rendered) | |

### B. Document model & round-trip
| Dimension | Status | Us | Them | Notes |
|---|---|---|---|---|
| Model freedom for overlapping ranges | 🔵 | run carries `CommentId`+`BookmarkStartId`+`RevisionId`+`NoteId` independently (`TextBuffer`/`Document.fs`) | PM tree+marks → needs **marker nodes** (`injectReplyRangeMarkers.ts`, `commentIdAllocator.ts`) | Our schema-free model is structurally cleaner; they pay the PM tax. |
| **Selective round-trip save** (preserve unknown XML) | 🔴 | regenerates XML (`DocumentWriter.fs`, 1900 LOC) → drops unsupported constructs | `docx/selectiveSave.ts` + `selectiveXmlPatch.ts`: patch changed nodes, preserve the rest | Biggest *round-trip* gap. Math/charts/SmartArt/CC/watermarks survive their save, not ours. |
| DOCX export coverage | 🟡 | paragraphs/runs/tables/numbering/section/borders/shading (`DocumentWriter.fs`) | full serializer suite (`docx/serializer/*`) | Ours solid for core; theirs broader + lossless. |
| Parser breadth (OOXML read) | 🟢 | `DocxReader.fs` 4291 LOC: styles, numbering, sections, tables, footnotes, fields, bookmarks, comments(refs), drawings, theme | comparable `docx/*Parser.ts` suite | Roughly even; spot-check specifics. |

### C. Editing, review & collaboration
| Dimension | Status | Us | Them | Notes |
|---|---|---|---|---|
| Tracked changes | 🟡 | display (insert underline / delete strike); not accept/reject | real OOXML revisions, suggesting mode, structural (para-break, table row/cell) | |
| Comments | 🟡 | render refs + callout | **threads + replies**, sidebar (`commentParser.ts`, reply range markers) | |
| Content controls (SDT) | 🟡 | parsed | parsed + serialized + UI (data-binding/repeating = their roadmap) | |
| Real-time collaboration | 🔴 | ✗ | yes (rides PM transaction/position-mapping) | The thing PM buys them — see rationale below. |
| Template variables | 🔴 | ✗ | variable-insertion system | |
| Agent SDK / MCP for editing | 🟡 | VBA/VBScript scripting bridge (`WordVba.fs`/`SpreadsheetVba.fs`) | `packages/agents` SDK + MCP | Different shape; both "agent-ready". |

### D. Where we lead (🔵 keep pressing)
| Dimension | Notes |
|---|---|
| True PDF export | embedded fonts + shaper = real print fidelity; DOM can't match. |
| Native render target | widgets path (WPF/GTK) in addition to web. |
| VBA/VBScript scripting | `runVba`/`runVbs` host bridge (commit `abde676`). They have no macro story. |
| Unified Office suite | spreadsheet + word share the engine; they're Word-only. |
| Real-Word parity harness | VBS automation + PNG/PDF diff vs actual Word. |

### E. Product maturity / packaging
| Dimension | Status | Notes |
|---|---|---|
| Frameworks | 🟡 | them: React + Vue + Nuxt adapters. us: Fable web + native. |
| i18n | 🟡 | them: 9 locales. us: ? |
| Release cadence / community | 🟡 | them: 98 releases, 1.6k★. |

---

## Three-way render harness
`tools/docx-editor-compare/` renders a `.docx` three ways for like-for-like comparison:
- `capture-ours.mjs <in.docx> <out.png>` — our browser-runner canvas (reuses built Fable JS).
- `capture-docx-editor.mjs <in.docx> <out.png> [--page N]` — eigenpal/docx-editor via npm
  `@eigenpal/docx-editor-react` `renderAsync` (readOnly), screenshots the painter's pages.
- MS Word ground truth = committed `docs/images/word/*-word.png` (captured on a Windows VM; can't
  regenerate from Linux).
- `align-3way.sh <name>` — crops each render to its page sheet and scales all three to the **same
  pixels-per-inch / page-aligned** (same zoom + pan) with a 1-inch grid, so positions are directly
  comparable across panels (don't naive-`-trim`+montage — that gives each panel a different zoom/pan).
- Gallery of results: `docs/docx-editor-3way-gallery.html`.
- **Fonts (REQUIRED for a fair run):** all three must use the same Windows faces. Mirror them first with
  `tools/word-parity/fetch-windows-fonts.sh` → `artifacts/fonts/windows/` ([[feedback-windows-fonts]]).
  `capture-ours.mjs` passes `fontRoot` (runner injects @font-face); `capture-docx-editor.mjs` injects the
  *identical* @font-face set (mirror of `browser-runner/Program.fs`) — pass `--no-fonts` only to disable.
  Do NOT compare with system-font fallback; metrics diverge and contaminate every delta.
- Gotchas baked in: docx-editor's `renderAsync` promise only resolves on first `onChange` (never in
  readOnly) → poll the DOM for `.layout-page` instead; multi-page docs exceed Chromium's screenshot
  cap → use `--page N`.

## Verified findings log
Newest first. Each entry: what, evidence both sides, date.

- **2026-06-06** — ✅ **FIXED 4 run-formatting gaps** (TextStyle gains `Hidden`/`DoubleStrike`/`Position`/
  `EmphasisMark`, threaded `DocxReader → Docx → Paint`; canvas+PDF): `w:vanish` (hidden text now omitted
  from layout), `w:dstrike` (two strike lines), `w:position` (baseline raise/lower in pt), `w:em` (dot above
  each char). Verified 2-way on `hidden-text.docx`/`run-effects.docx` — now matches docx-editor.
  Deferred: `w:sym` (parsed shape exists but needs a symbol font like Wingdings to render),
  `w:kern`/`w:effect` (low value). Larger remaining: fields beyond PAGE/NUMPAGES (TOC/STYLEREF/SEQ/PAGEREF
  need doc-wide counters + bookmark/heading→page maps) and non-rect preset-geometry shapes + gradients.

- **2026-06-06** — 🔻 **Source comparison: gaps where docx-editor leads (fixtures in
  `tests/repros/word-gaps/`, rendered 2-way).** Confirmed by reading their parsers (`runParser.ts`,
  `sdtProperties.ts`, `fieldParser.ts`, `shapeParser.ts`) that docx-editor handles and we **don't**:
  - **`w:vanish` (hidden text)** 🔴 — we render it as normal black text (a correctness bug — should be
    invisible); docx-editor recognizes it (faint/hidden). `hidden-text.docx`.
  - **`w:dstrike`** 🔴 — we render plain (no strike at all); docx-editor strikes through. `run-effects.docx`.
  - **`w:position` (baseline raise/lower)** 🔴 — we ignore; docx-editor raises/lowers. `run-effects.docx`.
  - **`w:em` (emphasis marks)** 🔴 — we ignore; docx-editor draws the dots. `run-effects.docx`.
  - **`w:sym` / `w:kern` / `w:effect`** 🔴 — in their `runParser`, not in ours (source-confirmed).
  - **OMML math rendering** 🔴 — we show the literal `[Equation]`; docx-editor renders the equation
    content (`MathExtension`). `math-omml.docx`.
  - **Fields**: TOC / STYLEREF / SEQ / PAGEREF / AUTHOR / FILENAME — their `fieldParser` handles; we
    compute only PAGE/NUMPAGES (source-confirmed).
  - **SDT types + dataBinding** (checkbox/dropdown/date/picture) — their `sdtProperties` distinguishes
    types; we keep only id/tag/alias. Note: *static render* of `sdt-controls.docx` is a tie (both show the
    sdtContent); their edge is making them interactive.
  - **Preset-geometry shapes** (`prstGeom` + solidFill/gradient/stroke) — their `shapeParser`; we do VML
    images + basic textboxes only.
  - Inconclusive: VML **watermark** in-page render (neither shows it; their parser feeds a watermark API).

- **2026-06-06** — 🚀 **PERF BUG FIXED → `../text`-measure on canvas IS viable (conclusion reversed).**
  The 80ms/word was a Fable bug, not shaping cost: `RuntimeHelpers.GetHashCode(font)` (cache key in
  `Engine.shapeText` + `Shaper.toPipelineFont` + `featureCache`) compiles to `identityHash`, which calls the
  record's structural `GetHashCode` → **deep-hashes the entire font (raw bytes + cmap + GSUB/GPOS) on every
  word.** Fix: `RuntimeHelpers.GetHashCode` → `LanguagePrimitives.PhysicalHash` (true O(1) reference hash) at
  the 4 sites. Result: **83ms → 51.7µs/word (~1600×)**, width identical. A page ≈ 10–25ms total. So canvas
  CAN switch to `../text`-measure for layout==print at acceptable perf — the original strategy is achievable.
- **2026-06-06** — (superseded by the perf fix above) `../text` → Fable → canvas: built, wired, benchmarked. (1) The *entire* `../text` stack (engine/shaper/opentype/bidi/unicode)
  **already Fable-compiles** — a probe emitted 101 JS files clean; the `#if FABLE_COMPILER` groundwork
  was already there. (2) Wired pluggably: browser-runner fetches `ucd.dat` + face bytes, builds a
  `ShaperText` font set, installs it via `WordApi.setTextMeasurers` behind `?measure=text` (TFM bridged by
  multi-targeting the `../text` chain `netstandard2.0;net10.0` + guarding 2 net-only `open`s in
  Office.Formats). (3) Correct — `measureWidth(Calibri,11,"Hello world") = 51.4px` (spot-on). (4) **But
  ~80ms per distinct word-measure (cache hit ~40ms)** under Fable — ~10⁴× slower than browser
  `measureText`; a page = tens of seconds. So canvas stays on browser-measure; `../text` measurement stays
  where `../text` also renders (PDF/native). The toggle + Fable-compat remain for a future optimized path
  (F# `Map` hotspots in the shaper / worker-offload) or one-off exact-print export. Confirms the
  two-renderer strategy empirically.

- **2026-06-06** — ✅ **FIXED tab stops.** `DocxReader.allTabStops` now parses every `w:tab` (pos/val/
  leader) into `BlockMetadata.TabStops: TabStop list`; `Layout.tabWidth` dispatches left/center/right/
  decimal (decimal aligns the pre-`.` prefix on the stop). Verified: `tab-stops.docx` now decimal-aligns
  the numbers at 3", centers the center-tab line, and honors the 2" left tab (was: all tabs → tiny default).
- **2026-06-06** — ✅ **FIXED underline styles.** `TextStyle` now carries `UnderlineStyle`+`UnderlineColor`,
  threaded through `DocxReader → Docx.fs → Paint`; `DrawUnderline` carries the line style; canvas
  (`Renderer.fs`) + PDF (`Pdf.fs`) draw single/double/thick/dotted/dash/dotDash/dotDotDash/wave/wavyDouble
  + a distinct underline colour. Verified on canvas: all 8 styles render distinctly.
- **2026-06-06** — 🔎 **Adversarial self-audit (code-driven) found 3 genuine gaps** (fixtures in
  `tests/repros/word-audit/`, rendered ours-vs-docx-editor):
  1. **Tab stops** 🔴 — parser only reads `w:tab val="right"` (`DocxReader.rightTabStop`); decimal/center/
     left/multiple tab stops are dropped and the model carries a single `RightTabPosition`. **Demo
     dramatic:** on `tab-stops.docx` docx-editor decimal-aligns numbers at 3", centers, and honors the 2"
     left tab; **ours collapses every tab to a tiny default advance** (nothing aligned). Strongest gap found.
  2. **Underline styles** 🔴 — `TextStyle.Underline` is a **bool** (`DocxReader: onOff "u"`), so
     `w:u val=double|thick|dotted|dash|dotDash|wavy|wavyDouble` and `w:u color` all render as a plain single
     black underline (`underline-styles.docx`).
  3. **RTL/bidi** 🟡 — canvas left-aligns `w:bidi` paragraphs (docx-editor right-aligns, matching Word).
     BUT **we own a full UBA bidi engine in `../text/bidi` (`Bidi.Algorithm`, conformance-tested) — it is
     wired into NOTHING.** The `../text` shaper (ligatures/kerning) IS wired into the **PDF path**
     (`Word.Pdf` → `ShaperText.shapeText`), not the canvas path. So RTL+shaping are *integration* gaps, not
     capability gaps: utilize `../text/bidi` (reorder + paragraph direction/alignment) in Layout/Paint.
  Note: tab-stops + underline are genuine model/parser gaps (no `../text` involved); fix is self-contained.

- **2026-06-06** — **12-fixture font-fair, page-aligned sweep.** Net: ours closer to Word on more
  dimensions. **Ours-closer:** table-cell pattern shading (docx-editor renders shaded cells **white** —
  no pattern support), list-default-indent, two-column density, nested-table row heights,
  running-header body spacing, float-image-wrap-top-bottom (docx-editor mispositions the float).
  **Ties:** table-valign, table-vmerge, alignment-inherits, heading-cambria, footnote (ours edges
  line-rhythm). **docx-editor-closer:** `table-grid-style-borders` (ours over-sizes rows/width).
  **Recurring docx-editor weakness:** default line/para spacing runs *tighter* than Word (≥4 fixtures) —
  their `fix-default-line-spacing`. Gallery: `docs/docx-editor-3way-gallery.html`.
- **2026-06-06** — ✅ **RETRACTED the gridspan "bug" + the grid-borders "loss" after measuring.**
  (1) `table-gridspan-bottom-fit` renders **correctly** — the "blank page" was the comparison harness
  clipping a multi-page (A4) doc at the default 1600px canvas; a tall render shows the gridspan A/B/C/D
  table at the page-2 bottom (matches Word's bottom-fit). Fixed `capture-ours.mjs` to **auto-size the
  canvas** to content height (single-page docs unchanged). (2) `table-grid-style-borders` is **not**
  over-sized — pixel-measured at matched scale, Word and ours both have 31px row pitch and 200/201px
  columns. Lesson: subtle sizing calls need pixel measurement, not eyeballed triptychs.
- **2026-06-06** — **Re-ran 3-way font-fair** (Windows Calibri/Cambria/Arial/Times across all three).
  Verdicts unchanged but now contamination-free: footnote-placement near-identical wrap on all three;
  list-default-indent still ours-closer (cascade is font-independent); table-valign + alignment ties.
- **2026-06-06** — **Justify FIX validated 3-way.** On `footnote-placement.docx`, our body text now
  renders flush-both-margins like Word and docx-editor (was ragged-left before). `Layout.fs justifyLineToEdge`.
- **2026-06-06** — **We beat docx-editor on `list-default-indent`.** Word keeps all list levels at the
  same default indent; ours matches, docx-editor cascades nested levels rightward. 3-way triptych.
- **2026-06-06** — `table-valign-row-height` + `alignment-inherits-default-spacing`: both renderers match
  Word (cell valign; left/center/right). Subtle inter-paragraph spacing diffs only.
- **2026-06-06** — (superseded) Justify was unimplemented on our side. `src/Word/Layout.fs:674`
  returns `0.0` offset for `Justify`; no inter-word distribution in
  `Layout.fs`/`Paint.fs`; `src/Word/Document.fs:2801 alignJustify` only sets
  metadata. Theirs distributes (`renderParagraph/line.ts`). → row A/justified.
- **2026-06-06** — docx-editor is dual-render: hidden PM + custom canvas-measured
  painter to static DOM (`docs/ARCHITECTURE.md`, `renderParagraph/line.ts:241`).
  Visual fidelity is **not** contenteditable-bound. → row A/measurement.
- **2026-06-06** — Their round-trip uses `selectiveSave`/`selectiveXmlPatch` to
  preserve unmodified XML; ours regenerates from model. → row B/selective-save.

## To investigate (⬜ open questions)
- ✅ DONE: headless docx-editor render harness exists (`tools/docx-editor-compare/`). renderAsync runs fine via npm package.
- Install real Windows Calibri/Cambria into the headless env so all three renders use Word's actual faces (current runs fall back) — then re-shoot for a fair pixel diff.
- Broaden the 3-way set beyond 4 repros: tables (gridspan/borders), floats-with-wrap, headers/footers, real-world docs.
- Their line-break algorithm vs ours on long justified/kerned paragraphs — who's closer to Word's break points?
- Tab-leader fidelity depth (both have active work).
- List numbering edge cases (`openspec/changes/list-operations-fidelity`).
- Do they handle charts (`c:chart`) visually, or detect-only like us?
- Their float-wrap pipeline depth (tight vs square vs through) — worth a deep read before we build ours.

---

## Architecture rationale: why they "hack around" ProseMirror — and what it buys them

It looks like a crazy amount of work (dual render + marker nodes + selective
save + React/Vue parity). It's a **rational build-vs-buy trade**, not waste:

**What PM buys them (the gain):** the entire rich-text *editing engine* for free
— selection model, cursor/IME handling, transaction-based undo/redo, schema
validity invariant, decorations, plugin ecosystem, and — the crown jewel —
**collaborative editing**. PM's `Step`/`Mapping`/`rebase` machinery solves
position-mapping through concurrent edits, which is the single hardest problem in
real-time collab and in "agent-ready" programmatic edits. Their headline
features (real-time collaboration, suggesting-mode tracked changes, agent SDK)
*are* PM's transaction/mapping system. They wanted to build the **DOCX layer**
(their moat) and rent the editor.

**Why the "hacking" is inherent, not gratuitous:**
- *Dual render is forced:* contenteditable can't paginate, can't do per-page
  headers/footers, and won't match Word's line-breaking. So "paged WYSIWYG that
  looks like Word" + "PM for editing" ⟹ you *must* render yourself. Hence a
  separate painter kept in sync with hidden PM.
- *Marker nodes are forced:* PM's model is a tree of nodes + non-overlapping
  marks; OOXML range constructs (`bookmark`/`commentRange`/`moveFrom`/revisions)
  overlap arbitrarily and cross blocks ⟹ shim with paired marker nodes +
  `selectiveSave` to preserve what the model can't hold.

**The cost they live with:** every visual feature exists twice (PM schema/`toDOM`
for editing semantics *and* the painter for pixels), kept in sync, **× React and
Vue adapters** (their CLAUDE.md: change layout/measure/paint in *both* or one
"regresses silently").

**Why we don't pay this tax:** we're a fidelity-first *renderer* (canvas/PDF/
native) + editor + spreadsheet, not a collab-editor-first product. We render
straight from our own schema-free model to three targets — no PM impedance
mismatch, no dual-render sync, no marker-node gymnastics, single codebase. The
bill we've deferred is exactly PM's crown jewel: **if/when we want real-time
collaboration, we'd have to build our own transaction/position-mapping (OT/CRDT)
layer** — the thing PM would have handed us. That's the one dimension where their
bet clearly out-positions ours; on rendering fidelity and Office breadth, ours
does.
