# Word and Spreadsheet: systematic review and build-out plan

Date: 2026-11-11. Inspected package version: `0.1.30`.

## Scope and evidence

The initial review was read-only: source, tests, configuration and existing documentation were not edited, and no builds, formatters, dependency installation or full test suites were run. The implementation follow-through below records subsequent authorized changes separately; the findings remain a historical snapshot.

Reviewed the project graph, packaging/release scripts, test runners, shared dialog specification/renderer, Word editing/import/save/layout paths, Spreadsheet editing/recalculation/structural operations, and browser/native host boundaries. This is a systematic, sampled review, not an exhaustive proof of correctness or a security certification.

Evidence labels:
- **Observed:** executed during this review.
- **Source-proven:** behavior or inconsistency follows directly from the cited implementation; not reproduced end to end here.
- **Risk / [INFERENCE]:** a plausible failure or architectural consequence requiring the stated check.
- **Documented:** existing repository evidence, not independently rerun.

Priorities: **P1** data integrity, misleading success, or a required product path blocked; **P2** significant usability, parity or maintenance defect; **P3** lower-impact delivery/documentation weakness. No P0 emergency is asserted.

## Implementation follow-through

The first correctness batch implements **W2** (PDF font identity) and **S1** (cross-sheet conditional-format and validation references). Word PDF now embeds the bytes belonging to the actual shaping font, including external shapers and colliding face names. Spreadsheet structural edits rewrite foreign rule formulas without moving their application ranges; undo restores untouched formula operands while retaining subsequent unrelated rule edits. A rule-only change does not scan the foreign sheet's entire cell map. Other findings remain open unless separately recorded in `continuation-roadmap.md`.

Verification: targeted .NET builds succeeded; `dotnet fsi --quiet tests/word/pdf.fsx` passed **20/20**, and `dotnet fsi --quiet tests/formula/stage21-structure.fsx` passed **188/188**. New regressions cover missing styled fonts, external font ownership, duplicate face names, foreign conditional/validation formulas across row/column insertion/deletion, unchanged application ranges, and selective undo. Browser/Fable execution, saved XLSX reopening in Excel, and the full release gate were not run in this batch.

The parked `ui/foundation` work was not duplicated. Coding used Terra and Sol without nested agents or build processes. Builds/tests were serialized through `/tmp/rbt-office-heavy.lock`, with single-node MSBuild, node reuse disabled and a 4 GiB managed-heap limit per process. Memory checks showed approximately 20 GiB available and zero swap use. An initial compiler-process exit 134 in unchanged OfficeUi did not recur in its isolated diagnostic build or the subsequent target builds; its cause was not established. Existing FSharp.Core downgrade warnings and Word layout warnings remain; no dependency pins were changed.

The next batch fixes **W1** by removing the lossy editor save entry points and the duplicate paragraph writer, migrating all repository callers and the document builder to the package writer. This is a breaking API removal, with migration instructions in `word-host-api.md`; external consumers must adopt the package contract before upgrading. The model's internal paragraph conversion remains where the package exporter uses it, not as a public document-save route.

Fresh Word and Office.Formats Fable builds passed. Targeted Chromium checks passed **126 assertions across seven checks**, including the temporary generator/preservation smoke, save-undo-boundary, revision, header-footer-xml, notes-xml, explicit-format and progressive-failure. The save-boundary regression now serializes the captured package after subsequent typing and reopens the saved file to verify snapshot content. The temporary smoke verified builder formatting/order, fresh note/comment/image parts, edited-document reopen and byte-identical preservation of the fixture's header/footer/note/style/chart/theme parts. It was removed after passing. These checks do not replace real Word interoperation, external host adoption or the full release gate.

The chart-reference batch addresses **S2**: imported chart formulas follow the workbook's structural-edit history and current sheet identities in both rendering and XLSX export. ChartEx names resolve against current workbook names. Binding changes invalidate parsed/rendered charts independently of cell changes; ordinary unrelated cell edits retain the existing dependency-based rendering behavior. The obsolete requirement that affected chart XML remain byte-identical after row insertion was removed.

Verification: fresh Office.Formats and Spreadsheet Fable builds passed after correcting integration compile errors. Three targeted Chromium checks passed **53 assertions**: `chart-reference-edits` (14), `charts-in-grid` (31), and `row-insert-keeps-drawing-part` (8). They exercise plotted values after row/column insertion/deletion, rename, undo/redo, moved-cell edits, expanded ranges beyond original cache counts, chartEx names-only edits, and save/reopen. Existing dashboard checks retain deferred-data rendering, pivot total exclusion and no redraw for unrelated edits. Exported packages passed the harness's integrity checks. The initial new regression accidentally selected background rectangles after chart resizing; its selector now identifies the data-series fill. Real Excel, Firefox/Safari, PDF output and the full release gate were not exercised. S3 remains next.

The clipboard batch fixes **S3** in the browser session: internal cells retain a stable source sheet ID across sheet switches; cross-sheet cuts use workbook move-reference rewriting for carried formulas, dependents and defined names. Paste undo restores overwritten target cells and workbook references; the earlier source Cut remains its own undo step. Workbook replacement/source deletion clears the clipboard, and pending reads are guarded against workbook replacement/disposal. Public API signatures are unchanged; semantics are documented in `spreadsheet-host-api.md`.

Verification: fresh Office.Formats and Spreadsheet Fable builds passed. Five Chromium checks passed **45 assertions**: `grid-clipboard` (15), `grid-clipboard-without-events` (11), `grid-clipboard-cross-sheet-copy` (10), `grid-clipboard-cross-sheet-cut` (8), and `grid-clipboard-deleted-source` (1). Coverage includes clipboard events, matching system reads, denied/unavailable access, unrelated external content precedence, formula shifting versus moving, formatting, foreign dependents/names, separate source/target undo and redo, exact saved/reopened formulas/results, import invalidation and deleted-source rejection. Two exported workbooks passed package-integrity checks. Builds/browser execution were serialized; the memory check showed 20 GiB available and no swap use. Real Excel, native hosts, other browsers and the full release gate were not run.

Delivery graph changes for **C1/C2** remove all eight sibling-repository entries from `Office.slnx` and delete the unused `Spreadsheet.Core` project and solution entry. No project consumed that nominal Core split; the supported browser and native hosts continue to use the main Spreadsheet project with registry-pinned external dependencies. This removes the misleading duplicate compile list rather than claiming a platform-neutral extraction has been completed.

**C3/C4** now fail closed. The publisher requires a clean exact-commit release bundle, runs the core gate, installs the single packed tarball as a consumer and runs the Word/Spreadsheet browser harnesses against its modules before publishing those same bytes. The bundle retains Word/Excel save corpora, complete open logs and a reviewer attestation bound to commit/package/corpus hashes. `--prepare` does not publish; Forge must be provisioned with the reviewed bundle via `FORGE_RELEASE_BUNDLE`. The revised release procedure is in `continuation-roadmap.md`; no bundle transport or Office execution is claimed here. Test runners have explicit core/native profiles, mandatory suite/fixture preflight, an explicit office-formats suite inventory, and propagated `--skip-build` freshness opt-out. Native tests no longer depend on whether a DLL happened to exist.

Verification: all **15 supported solution projects** restored/built in an isolated `/tmp` copy without sibling repositories or copied `bin`/`obj` outputs, using registry-pinned packages from the machine's NuGet cache. The initial compiler process exited 134 without a source diagnostic; a narrow format build and then the solution succeeded with `DOTNET_PROCESSOR_COUNT=2`, single-node MSBuild, node reuse disabled and a 4 GiB managed-heap cap. Existing downgrade and incomplete-match warnings remain. The explicit native gate built its targets and passed **25 painter assertions** plus the Wisp editor-host smoke.

Negative checks removed a Word suite, a formula stage, a chartEx fixture and the Wisp host suite from the isolated test copy; every invocation failed in preflight. Both runners rejected unknown options. A throwaway command recorder confirmed root `--skip-build` reached the formula runner, invoked no builds, and propagated intentional suite failure (not a claim that the core suites passed). The publication failure regression proved a failed core gate never reaches `npm publish`; synthetic oracle-validator inputs were accepted and rejected after tarball tampering. A local pack/install of the existing `dist` matched byte-for-byte and passed the installed-consumer `grid-clipboard` check (**15 assertions**); this exercises package consumption, not a freshly built release candidate. Shell syntax checks passed. No actual publication, full core/browser release run, or real Word/Excel oracle was performed. Temporary copies/packages were removed.

The Word correctness batch implements **W3–W8**. Revision identities are indexed by weak content-object keys rather than the most recent 400 observations, preserving retained undo/redo checkpoints without retaining abandoned documents. Progressive-load waiters receive their own terminal result on completion, failure, cancellation, replacement or disposal. Input listeners are owned and removed together; delayed callbacks cannot mutate a disposed editor or reused canvas. Internal rich paste requires explicit, unchanged clipboard provenance instead of matching plain text. Browser cuts commit only after successful transfer and a still-valid captured edit; failed writes are reported without deletion. The explicit synchronous host `cutSelection` contract is retained and documented separately. Physical and virtual paragraph insertion share list/table/tracked-change behavior.

Verification: fresh Office.Formats and Word Fable builds passed after fixing canvas binding and promise-call integration errors. Nine targeted Chromium checks passed **253 assertions**: revision (34), progressive-cancel (44), progressive-failure (34), save-undo-boundary (4), editing (79), ime (45), clipboard-keys (4), typing (5), track-changes (4). Coverage includes two 410-character coalesced runs, saved checkpoints and byte-identical unchanged export/reopen; outgoing load promises and shared viewers; retained canvas/proxy events and pending callbacks after disposal; external same-text HTML/plain content, links and multiple editors; denied/unavailable/pending clipboard writes; ordinary/list/table/tracked virtual Enter, duplicate-event suppression and undo. The empty-list regression setup was corrected to retain an actual empty list paragraph instead of an empty document. Builds and browser execution were serialized with bounded .NET processor/heap settings; the initial memory check showed 19 GiB available and no swap. Native hosts, physical mobile keyboards, Firefox/Safari, real Word and the full release gate were not exercised.

The native spreadsheet batch implements **S4–S6**. Wisp and Widgets now retain authoritative `Workbook` state and complete workbook history, using the existing shared `Recalc` engine for value edits, range clearing, clipboard changes and cross-sheet dependencies/spills. Typed edits calculate before `Worksheet_Change`, then calculate macro writes within the same undo transaction. Model-level table commands share `Workbook.updateSheetModelWithTables`. Sheet operations use stable, non-recycled IDs and workbook reference rewriting. Wisp imports the complete context and exports current names, name metadata, tables, visibility, active sheet and iterative settings, remapping local-name scopes to output positions only at save.

Verification: both native projects built; the Wisp regression suite passed **50 assertions** and the Widgets suite **52 assertions**, including keyboard edits/deletion, cross-sheet chains and moves, spill growth/shrink, names/table metadata, complete sheet history, save/reopen, replacement cancellation and macro-event sequencing. The actual Wisp executable passed `--selftest`. Its first launch exposed an FSharp.Core runtime mismatch: the executable now selects 10.1.401, as required by the pinned Wisp package. Shared library builds retain SDK-compatible Core for `dotnet fsi`; Wisp scripts resolve their native dependencies from the library output rather than the executable's newer runtime directory. Existing library downgrade and incomplete-match warnings remain.

The native gate now requires/builds Widgets, runs both spreadsheet host suites and exercises the Wisp executable; its shell syntax check passed. All heavyweight checks were serialized with a 4 GiB managed-heap cap and two logical processors. The final memory snapshot showed **19 GiB available, no swap use**. No interactive native window/OS clipboard, browser/Fable, real Excel or full release/native gate was run in this batch. These fixes share workbook operations and calculation, not a completed portable-controller extraction.

The spreadsheet geometry batch implements **S7**. Large browser grids now cap their physical layout extent at 8,000,000 pixels and map scrollbars to logical sheet coordinates. Cells, frozen panes, headers, selection/edit overlays, comments, validation lists, outlines and drawings use the rebased viewport origin. Zoom preserves the logical viewport; wheel input travels in screen pixels rather than jumping through the compressed scrollbar range.

Pivot layouts now move bounds, header/data offsets and row/column items with structural edits. Deleted repeated-label parents do not lend stale labels to surviving children; deleting a grand-total column no longer substitutes its remaining member. GETPIVOTDATA dependencies include pivot label/value cells, and metadata changes rewire affected lookup formulas. Export patches surviving source XML spans rather than recreating pivot items, preserving data-field indexes and unknown attributes/children; inserted gaps become blank items and wholly deleted pivots are removed. Pivot-cache refresh remains a separate, unimplemented authoring capability.

Verification: fresh Spreadsheet .NET and Fable builds passed, as did the Office.Formats Fable build. Eight targeted Chromium checks passed **276 assertions**: grid-extent-scroll-and-stacking (27), grid-full-height-zoom (46), grid-selection-and-keyboard (40), api-view (83), chrome-outline (48), chrome-frozen-drawings (18), comment-box-in-view (7), and pivot-structure-roundtrip (7). Seven exported packages passed integrity checks. A separate actual-surface smoke visibly reached row 1,048,576 at 400% with a frozen row/column, hidden penultimate row and custom height, committed typed text, and retained the last row at 100%. The temporary fixture/server/build output was removed.

The pivot suite passed **74 assertions**, structural suite **259**, and incremental recalculation suite **10,616**. Structural regressions cover insertion/deletion, relative offsets, deleted totals, cross-sheet dependencies, selective invalidation, undo/redo, save/reopen, qualified location tags and preservation of unknown pivot-item XML. The structural run exposed and fixed both a duplicate FSI binding and the singleton-member/deleted-total lookup defect; the affected suites were rerun successfully. The wider recalculation run preceded that final lookup correction. All heavy work was serialized with two logical processors and a 4 GiB managed-heap cap; the memory snapshot showed **19 GiB available and no swap use**. Existing build warnings remain. No Firefox/Safari, real Excel, native interaction, full release gate or pivot-cache refresh is claimed. Next review work remains browser renderer C5–C6, without duplicating the parked UI branch.

The shared browser dialog batch implements **C5–C6** in `OfficeUiDialog`. All eleven current field kinds render, tab values persist, computed previews share that state, and non-first select defaults stay synchronized with the visible control. Labels/title/description are associated; tabs have keyboard navigation; modal focus is initialized, contained and restored. Enter respects multiline/native controls and composition, Escape cancels safely, and callback-driven replacement is not closed by the outgoing callback. Deferred focus and document listeners are disposed. This adapts the relevant parked `ui/foundation` ideas without merging its command registry, expanded spec/validation API or Format Cells work. Existing callback validation remains authoritative: `false` retains edits but does not roll back mutations made by the callback.

Verification: the runner's Fable compilation and **seven Chromium dialog cases** passed, plus keyboard, focus, ARIA, retained-tab/validation, replacement and disposal checks. A genuine Word hyperlink dialog left document/history unchanged on Cancel and applied exactly one undoable command on acceptance. A separate visual browser smoke inspected both tabs, edited values, observed the live preview and submitted once with Enter. The initial runner-local Fable 5 manifest was removed so it inherits repository Fable 4.28; Vite rewrites the private build entry before module discovery. The initial select-keyboard check incorrectly opened the native popup before testing collapsed-select navigation; it now focuses the select and tests native Home/ArrowDown behavior. Existing Word compiler warnings remain. No full release gate, Firefox/Safari, native dialog run, physical IME or complete browser-shell wiring is claimed. Next review item: **C7**, the current capability matrix and historical-document separation.

The documentation batch implements **C7**. The roadmap now owns one [current capability matrix](continuation-roadmap.md#current-capability-matrix), based on code commit `adb77f90`, separating import, preservation, rendering, editing, API exposure, shell reachability and platform evidence. Architecture/UI/competitor audits are explicitly historical. Incorrect Spreadsheet PDF “.NET only” claims and obsolete roadmap tasks are corrected: the browser module and APIs exist, while menu/native-shell wiring and fresh PDF verification remain separate. Source inspection also confirms that WordShell Save and native Word Open/Save/PDF are not completed product workflows. No new runtime behavior or release certification is claimed by this documentation-only batch.

Documentation verification: **18 local links/anchors** resolved, and the matrix has **ten capability rows with all eight status columns**. No builds or runtime suites were rerun for this documentation-only change; earlier results remain dated evidence.

The browser-shell IO batch connects **Word File → Save / Ctrl/Cmd+S** to the existing source-preserving package writer and **Spreadsheet File → Print PDF** to the mounted grid's existing PDF API. Word captures the saved revision, coalesces overlapping requests, preserves later edits as dirty, and rejects stale completion after import/disposal/removal. Its fixed-name download reports browser handoff, not confirmed disk persistence. Spreadsheet reports pending/requested/error state and suppresses overlapping clicks and stale UI updates. Its factory now requires the grid as the first argument; repository callers and host contracts are migrated. PDF font bytes remain host-provided.

Verification: fresh Office.Formats, Word, Spreadsheet and Spreadsheet.Pdf Fable builds passed. Chromium regressions passed **28 assertions**: `shell-save` (16), `shell-print-pdf` (6), and `shell-print-pdf-failure` (6). Word checks cover actual writer bytes/reopen, bold formatting, retained opaque chart XML, saved-revision undo/redo, async edits, writer rejection, replacement/disposal, and shortcut scope/composition. Spreadsheet checks compare actual edited-workbook PDF text/pages, require a real loaded PDF iframe and print dispatch, and exercise failed module fetch, retry and detached UI. A separate actual-surface run downloaded DOCX from File → Save and from real Ctrl+S after typing, and visibly reached Spreadsheet's print-request status.

The first Word run exposed a shared `FableUI/Dom.fs` cleanup defect: Fable hoisted returned-function adapters out of attribute/child loops, so nested editor cleanup was skipped. Per-item callback scopes now preserve the correct disposers; the regression checks both released host ownership and removed input proxy. Real keyboard input also exposed the proxy's placement outside the shell subtree; the owned document listener now scopes events to that editor's proxy or shell rather than intercepting unrelated inputs.

The original Playwright headless-shell runtime lacked the native PDF viewer and never emitted the print iframe's load event. PDF checks now use installed full Chromium's new headless mode; strict print-dispatch assertions remain, with only the platform print function intercepted in the automated check. Physical printer/dialog acceptance, filesystem persistence, Firefox/Safari, native apps, Office interoperability and the full release gate were not certified. Builds/browser suites were serialized; the memory snapshot showed **18 GiB available and zero swap use**. Existing compiler warnings remain.

The browser file-session batch adds actual DOCX/XLSX Open, opened-filename Save, visible dirty state, replacement confirmation and unload protection to both shells. Failed/cancelled/stale reads preserve the current session. Spreadsheet Save commits a pending cell before revision capture, refuses rejected validation without losing the buffer, and snapshots source/iteration metadata across asynchronous archive reads. Open and unload also protect active cell entries. Shared import logic now exposes a guarded installation boundary; host API documentation records those methods and the shell callback/signature changes.

Word PDF now captures the whole live document and lazily invokes the existing packaged renderer, with host-provided font bytes, browser image decoding, popup reservation and stale/error cleanup. The browser shaper's Unicode data was not previously initialized: the package build now extracts `ucd.dat` from the **resolved pinned NuGet assembly**, shares Word modules, and loads the asset beside `Browser.js`. No sibling checkout or unrelated Unicode version is used. Missing fonts fail explicitly; PDF export does not mark DOCX edits saved.

Verification: fresh Office.Formats, Word, Word.Pdf and Spreadsheet Fable builds passed. **72 Chromium assertions** passed: Word `shell-files` (18), `shell-save` (16), `shell-pdf` (6); Spreadsheet `shell-file-session` (12), `shell-file-races` (5), `shell-file-shortcuts` (7), `shell-file-edit-buffer` (8). Six exported workbooks passed package integrity. Regressions include actual package save/reopen, read/platform errors, overlapping operations, edit/undo/redo during reads, disposal, unload/shortcut scope, pending input and validation rejection. Initial title assertions were corrected to await reactive rendering/capture the true pre-disposal state; the initial PDF run exposed the missing Unicode initialization rather than an unavailable backend.

Actual-surface proof uploaded and edited `sample-formatted.docx`, saved its original filename, and opened a generated **166,882-byte PDF** whose extracted text contains the typed `FILE PDF PROOF` and whose image list contains four JPEG placements. The spreadsheet surface opened `milk-street-office.xlsx`, typed into the active cell editor and downloaded the same filename with Ctrl+S, clearing its dirty indicator. One early visual attempt typed before DOCX installation completed; the stale-open guard correctly cancelled that replacement. This is workflow evidence, not layout/font parity or Office certification. All heavy work was serialized; final memory was **18 GiB available**, with **76 KiB swap already present at the initial snapshot**. Existing package-downgrade/Fable warnings remain; native apps, other browsers, physical printing and the full release gate were not run.

The shared-command batch adapts the parked foundation's registry into `OfficeUi/Commands.fs` without replacing the current spec schema. Both browser shells route their formatting/find controls through it. Word no longer selects the whole document for toolbar formatting; effective selection state, mixed toggles, font values, paragraph properties and undo/import refresh drive the chrome. Paragraph Apply patches only edited fields, preserving unrelated mixed metadata in one undo step. Find recomputes matches and uses the normalized selection end, including backwards ranges; the obsolete cached navigation fields are removed.

Spreadsheet now owns formatting, history and active-sheet find/replace rather than requiring host callbacks. Its factory is `createSpreadsheetShell(grid, onAction)`; clipboard, structure and data actions remain delegated. The browser-facing command facade exposes callable `Run`, `State` and `OnChange`. The shared renderer hosts five supported Format Cells tabs; unchanged mixed fields are omitted, chosen edits form one patch/undo step, and invalid pending input is retained. Native and bare-widget command integration, advanced Format Cells properties and a common visual shell remain outside this batch.

Verification: fresh Office.Formats, Word and Spreadsheet Fable builds passed. **198 Chromium assertions** passed across Word `shell-editing` (30), `shell-files` (18), `shell-save` (16), `selection-format` (24), `pending-format` (29); Spreadsheet `shell-selection-commands` (31), `shell-find-selection` (18), `shell-file-session` (12), `shell-file-races` (5), `shell-file-shortcuts` (7), `shell-file-edit-buffer` (8). Eight exported workbooks passed package integrity. A temporary .NET script passed seven portable-registry checks. Integration exposed and corrected incomplete Word record/type changes, a raw F# registry without callable JavaScript methods, and paragraph selects that displayed their first option instead of the selected value. Browser fixtures were corrected for portal placement, asynchronous reactive rendering and the file status region's specific identity.

Actual-surface proof selected and bolded only `alpha ` in Word, selected `beta ` and observed it remained unbolded, and opened paragraph properties showing the chosen center alignment. Real Ctrl+F navigated to the second `alpha`; Replace all produced `omega beta omega`. Spreadsheet's real Format Cells dialog changed A1 to 18 points while neighboring cells retained their formats; Find next selected A2 and Replace all changed both `alpha` cells while preserving `beta`. Screenshots showed both dialog surfaces. Heavy jobs were serialized and .NET processors/heap capped; an observed memory snapshot showed **18 GiB available and 88 KiB swap** (76 KiB at the start, with other sessions active). Existing compiler warnings remain. No native, Firefox/Safari, Office interoperability, physical-print or full release gate was run.

The shared-chrome batch adds `FableUI/Chrome.fs`: document-owned, idempotent styles scoped to office shell/dialog roots, with shared theme tokens and editor-specific layout. Word's toolbar wraps and its logical-width workspace scrolls locally; Spreadsheet's toolbar/search wrap with bounded menu panels. Controls have descriptive names, toggle state and visible focus. Legacy Word dialogs now have title associations and cleaned-up focus timers; both shells restore visible menu triggers when dismissing menu-launched dialogs. Opening a spreadsheet color picker no longer mutates the selected cells.

Verification: fresh Word and Spreadsheet Fable builds passed with existing warnings. Actual Chromium desktop (1200px) and narrow (390px) surfaces were inspected, including screenshots and accessibility snapshots. All five menus in each editor remained within the narrow viewport; page width stayed 390px. Word selected and bolded only `alpha `, wrapped Tab inside paragraph properties and returned focus to Format after Escape. Spreadsheet bolded A1 without changing B1, applied 18-point text through Format Cells, found A2, and left formatting unchanged when opening/cancelling its text-color picker. Styled modal bounds stayed within the narrow viewport. Visual verification exposed and fixed the missing spreadsheet dialog stylesheet and menu-dialog focus loss; both corrected surfaces were checked again.

Two independently compiled shells on one page installed one shared stylesheet; disposing Spreadsheet left Word styled, while the host's Georgia button retained its native minimum height. Builds were serialized under a file lock with two .NET processors and a 4 GiB heap cap. An observed memory snapshot showed **16 GiB available and 96 KiB swap**, versus 19 GiB/88 KiB initially with other sessions active. This is browser-chrome evidence, not screen-reader/document-grid accessibility certification, native acceptance, Firefox/Safari coverage or a full release gate.

The data/link/table batch connects Spreadsheet's Data menu to shared commands and real grid operations: validation editing/clearing, quick and two-key custom sorting, and supported value/condition/top-bottom/color/average filters. `SetAutoFilterAndFilter` validates the range and criterion before one undoable change. Migrated data callbacks were removed from the demo and usage documentation. Word's menu/toolbar/scoped Ctrl+K share hyperlink insertion, whole-link editing and removal; table operations respect caret context, prefill cell/table properties and reject stale document snapshots.

Word's partial cell transaction preserves unrelated imported shading, padding, borders and origin metadata. Deliberate width changes set an explicit twip preferred width; border edits update the imported edge overrides rather than being hidden behind them. Shared reactive DOM properties no longer rewrite an unchanged browser value, preserving transient negative-number input. Imported non-default dropdown choices are rendered as selected, including retained shading patterns not in the normal choice list.

Verification: final Office.Formats, Word and Spreadsheet Fable builds passed with existing warnings. Actual Chromium menu/keyboard/forms produced **46 successful targeted assertions**, covering link validation/edit/remove/undo and Ctrl+K scope; table dimension errors, imported-property prefill, no-op Apply, negative input rejection, partial edits, cancellation and separate undo; validation types/custom formula prefill, scoped clearing with preserved relative-reference anchors; sorting headers/rows/formulas; filter cancellation, invalid-column rejection, hidden-row visibility and atomic undo. An additional direct API check rejected invalid filter criteria without changing range or revision. DOCX XML retained edited width/borders and original shading; XLSX export/reimport retained filters and custom validation. Host edits invalidate an open data dialog, and read-only data commands refuse execution.

Screenshots covered desktop and 390px forms. Integration exposed and fixed F# indentation/type-inference failures, incorrect initial dropdown values, negative signs erased by reactive writeback, competing imported width/border representations and custom-formula prefill. Fixture numeric seeds were corrected to use the grid's string input contract; a clearing assertion was corrected after inspecting the retained validation anchor rather than assuming rewritten formula text. Builds were serialized with two .NET processors and a 4 GiB heap cap; memory snapshots ranged from **16–18 GiB available**, ending at **17 GiB**, with **96 KiB swap unchanged**. No native, Firefox/Safari, screen-reader, full release or real Office interoperability gate was run.

The formatting/name batch adds Spreadsheet rule and scoped-name managers through `data.conditionalFormatting` and `data.names`. `Managers.fs` sits behind an editor-specific callback port, so shell and bare-widget entry points use one lazily mounted shared renderer; the old five-kind conditional dialog is removed. Rule forms separate type choice, operands and differential formatting. Names distinguish workbook/sheet scopes, retain hidden metadata, reject duplicate creation, and keep existing identity immutable rather than emulating rename with delete/create.

Word's Font and Paragraph borders/shading dialogs patch only changed fields. Mixed runs/paragraphs keep their individual untouched properties; caret font changes remain pending until typing. Border sides are independent, and paragraph pattern/foreground shading now crosses the import, model, painter and DOCX writer. Each accepted selection edit is one undo operation. Browser dialogs retain validation errors, reject changed document/selection contexts and restore their invoker. Native menu actions are optional rather than new visible no-ops.

Targeted verification passed isolated Office.Formats, Word and Spreadsheet Fable builds, native Spreadsheet and `word-wisp` builds, **93 DOCX-reader checks**, **43 XLSX-rule checks** and **221 XLSX-writer checks**, including source-backed differential-style cloning and opaque-rule preservation regressions. Existing build warnings remain. Actual Chromium surfaces were inspected at desktop and 390px widths; **74 successful targeted assertions** covered mixed/no-op edits, validation retention, atomic undo, caret formatting, stale contexts, read-only commands, scoped-name recalculation, duplicate creation, built-in-name rejection and bare-widget invocation.

Imported-package checks exercised edited shared DXFs without losing font metrics, number formats or borders; untouched rules retained their original differential fill. Edited x14 data bars retained automatic thresholds and advanced settings. Opaque x14 rules, pivot-group metadata, hidden names and untouched comment whitespace survived save/reopen/re-save. Word border sides and paragraph shading survived DOCX export/reimport. Integration also fixed native module guards, readonly-menu refresh, compact dialog layout and F# fixture/type-inference errors.

Builds/tests used one shared lock, two .NET processors and a 4 GiB managed-heap cap. Available-memory snapshots ranged from 14–17 GiB, ending at 15 GiB with 138 MiB swap in use; concurrent sessions prevent attributing machine-wide swap changes to this batch. Owned browser tabs, preview services and scratch builds were removed. Native compilation is not native UI verification; no full release gate, Firefox/Safari, screen-reader or real Office interoperability certification was run.

The sheet/list batch adds visible sheet tabs with owned menu/keyboard navigation, stable-identity targeting and create/rename/duplicate/delete/move/hide/unhide commands. Ordinary Unhide excludes very-hidden sheets; the last visible sheet cannot be hidden or deleted. Shared row-height/column-width dialogs preserve mixed values unless explicitly changed; autofit, zero-size hiding and size-restoring unhide use existing grid transactions. `OnReadOnlyChange` immediately refreshes mounted shell controls without a workbook edit.

Word's Bullets and numbering dialog applies partial marker/kind/level changes and restart/continue as one undoable transaction. Formatting an existing numbered item preserves its ordinal, imported template punctuation and unselected nested parent counters. Shared counter transitions materialize affected suffix items as standard OOXML numbering instances with per-level start overrides, not private XML metadata. Reopened instances remain distinct; arbitrary multilevel-template authoring is not claimed. Restart includes the contiguous same-instance tail; Continue can cross ordinary paragraphs to the nearest preceding same-level numbered list in the same story.

Verification passed final Office.Formats/Spreadsheet/Word Fable builds, native Spreadsheet and Word-wisp builds, **89 sheet-structure** and **50 Word-list native assertions**. Browser regressions passed **33 shell-selection**, **91 sheet API** and **14 DOCX numbering assertions**; **25 XLSX exports** had no new package-integrity problems. Actual desktop/390px surfaces produced **74 successful targeted assertions**, including sheet keyboard/context-menu ownership, cancellation, pending invalid entries, stale dialogs, read-only mode, geometry and XLSX reopen; Word mixed/no-op changes, repeated partial formatting, nested counters, restart validation, combined format/restart, continue, list removal, one-step undo, stale/read-only guards and DOCX reopen. No final page errors were reported.

Integration fixed an apostrophe-sensitive sheet-reference prefilter that broke rename undo, stale read-only shell controls, list-format edits that reset later ordinals, imported numbering-template loss and combined format/restart rejection. Test setup was corrected where an unfocused F2 never began editing, a preview omitted corpus fixture routes, and manually modified DOCX packages incorrectly retained `Unchanged=true`. These setup failures are not counted as product passes.

Resource incident: the root filesystem filled during a Word source write and truncated the final section/header-footer functions in `Document.fs`. That tail was recovered from HEAD; its rebuilt JavaScript matched the saved pre-failure compiled tail exactly after excluding six intentional copies of the new `ListStartOverrides` field. This verifies the recovered executable behavior, not byte-identical source comments/whitespace. Space was subsequently freed by another process; this session deleted no unrelated files. Builds/tests remained serialized under the shared lock with two .NET processors and a 4 GiB heap cap. Native compilation is not native UI proof; no full release gate, Firefox/Safari, screen-reader or real Office interoperability certification was run.

Final machine snapshot: **14 GiB available RAM, 2.8 GiB swap in use and 98 GiB free on the root filesystem**. Concurrent sessions prevent attributing machine-wide swap changes to this batch. The owned preview tabs and server were closed.

The function/style batch adds a searchable Spreadsheet Insert Function chooser and shared lexical argument context for grid and standalone formula-bar inputs. The catalogue contains 21 documented functions, not every evaluator function. Insertion captures cell/edit/workbook state and changes only the pending buffer; normal entry owns validation, calculation and undo. Tab completes exact names or unambiguous prefixes, but does not rewrite structured-reference or array-constant text. Context handling distinguishes nested calls, quoted/escaped text, quoted sheet names, bracketed references and array commas; it is not a second formula evaluator.

Word's styles pane searches and applies existing paragraph/character styles, reports current/mixed state, reveals hidden styles deliberately and refreshes on document replacement. It opens above the workspace with search focused; it does not author style definitions. Named character styles now work for future caret typing. Actual imported-package checks exposed inherited font/size values becoming unintended direct overrides on style application. Reader/native-import provenance plus the shared style application path now distinguish genuine direct fonts, including values equal to the old style, from inherited typography. Direct bold/font/size overrides survive; inherited typography follows the new style and survives save/reopen.

Verification: final Office.Formats, Word and Spreadsheet Fable builds passed, as did native Word-wisp/Spreadsheet builds and **93 DOCX-reader assertions**. Targeted Chromium regressions passed **20 explicit-format**, **36 pending-format** (including real keyboard input), **35 shell-editing** and **26 function-help assertions**; one checked XLSX export had no package-integrity problems. Actual desktop/390px workflows produced **50 successful targeted assertions**, covering fresh/imported style application, mixed/hidden states, read-only guards, undo, caret typing, identical retained style-definition XML, document replacement, function insertion/calculation, nested/quoted formulas, cancellation, keyboard focus containment, readable descriptions, bounded hints and XLSX reopen.

Integration also fixed Fable's dropped case-insensitive `IndexOf` comparison in chooser search, array commas advancing the wrong argument, completion inside structured references, flex-shrunk chooser descriptions and hint padding overflowing compact viewports. Test setup needed the original synthetic DOCX fixture route and the pending-format harness's real-input driver; a server-restart fetch failure was infrastructure, not an application regression. Focus restoration correctly returns to the grid when the toolbar preserves grid focus. The attempted bottom-edge scenario auto-scrolled the edited cell into view, so the observed claim is bounded visible hints, not a separately forced bottom-edge placement test.

Builds/tests remained serialized under the shared lock with two .NET processors and a 4 GiB managed-heap cap. Available-memory snapshots ranged from 18 to 13 GiB, swap from 2.7 to 2.8 GiB, with other sessions active; root free space stayed at least 96 GiB. No native-window, Firefox/Safari, screen-reader, full-release or real Word/Excel certification was run.

The selection/page-setup batch adds a live Spreadsheet status bar backed by the existing cached `GetStatus()` contract, not a second aggregation implementation. It shows populated/numeric counts and numeric sum, average, minimum and maximum; undefined aggregates are hidden. Selection changes, recalculation, undo/redo, filtering, sheet switches and import refresh the display. Pending cell edits remain uncommitted. Read-only viewing, shell isolation and disposal/remount were exercised, including sparse whole-column selection and compact wrapping.

Word's Page setup now explicitly targets the section containing the selection's focus rather than silently editing the final section. It offers paper presets/custom dimensions, orientation, margins, header/footer distances in inches and different-first-page behavior. Unchanged fields retain imported precision; invalid, stale and read-only applications do not mutate the document. Geometry applies in one undo step. The existing public `setPageSetup`/`getPageSetup` APIs intentionally retain their documented final-section scope.

The writer now merges changed section properties into their source XML, retaining unmodeled attributes and children. Targeted saves preserved `rsidSect`/`rsidR`, `docGrid`, paper codes, paper source, gutters and precise untouched margins. An actual two-section edit left the unselected final section and inherited header XML byte-identical. Per-section header/footer positioning is used for paint and editable stories; the regression checks that changing one section's header distance moves its caret but not the following section's.

Verification: Office.Formats, Spreadsheet and Word Fable builds passed, along with the native Word-wisp build and **33 native section assertions**. Chromium passed **19 selection-status** and **13 page-setup regression assertions**, plus **17 successful targeted surface assertions**. One XLSX export passed package-integrity checks. Actual keyboard entry/selection/undo, DOCX upload, selective section save/reopen, custom-size cancellation, desktop/390px controls, dialog scrolling, focus containment and Escape focus restoration were exercised. Both final browser surfaces had no recorded errors. Browser verification exposed an orientation select that reopened as Portrait despite landscape dimensions; using the existing controlled-option helper fixed it and a reopen regression now covers it. Initial F# indentation errors and a test fixture with no selected paragraph were corrected before the successful runs.

Builds/tests were serialized under the shared lock with two .NET processors and a 4 GiB managed-heap cap. Available RAM was 15 GiB initially and 17 GiB after cleanup; swap remained 2.8 GiB in those snapshots. Root free space was 95/93 GiB respectively, with concurrent sessions active. Owned browser tabs, preview service and scratch build/fixture files were removed. No full release gate, real Word/Excel oracle, native-window or Firefox/Safari certification was run.

The Find/table-properties batch replaces Spreadsheet's legacy browser Find dialog with one grid-owned search session and bar. Bare-grid Ctrl+F/H, public APIs, shell commands and shell controls share query/options/results. It supports sheet/workbook and optional-range scope, values/formulas, case/whole-cell/regex matching, next/previous, replace current/all, live refresh and undo. Shell detach returns the bar to its grid; disposal removes listeners and DOM. The native Find descriptor is retained: the Wisp host still consumes it.

Word's existing table-defaults dialog now edits preferred width, alignment, left/start indent, table shading/background, grid borders, default margins/spacing and current-row minimum height, splitting and repeating-header state. Table and row edits form one undo transaction; untouched imported metadata and unrelated tables survive. Only consecutive leading header rows repeat. Exact row-height and fixed/autofit layout behavior are not implemented by this workflow; untouched imported values remain preserved, and editing height selects the minimum-height rule.

Integration corrected F# inference/indentation, the border-width label (pixels, not points), explicit allowed row splitting (`w:cantSplit w:val="0"`) and percentage-width conversion. Editor packages now consistently carry actual percentages; conversion to/from DOCX fiftieths occurs only at the format boundary. The previous magnitude heuristic reopened 1% as 50% and would normalize percentages above 100 twice. The round-trip metadata comparator now includes table indent.

Verification: targeted Chromium regressions passed **30 shared-Find**, **18 selection-Find**, **26 component lifecycle** and **11 table-property assertions** (**85 total**); two XLSX exports passed package-integrity checks. Actual desktop/390px shell and bare-grid workflows passed **24 targeted surface assertions**, including malformed regex rejection, replacement/undo, focus restoration, imported table preservation and fresh 50% table geometry. Save/reopen preserved explicit row splitting, 1% and 125% table widths, and a 0.5% cell width. Final actual surfaces recorded no browser errors. Native table-column tests passed **24 assertions**, with **3 additional native import/indent geometry assertions**.

Office.Formats, Word and Spreadsheet Fable compilation passed; final Word-wisp and Spreadsheet-wisp native builds passed with existing package-downgrade/pattern warnings. Native verification caught the still-used Spreadsheet dialog descriptor being removed; it was restored before the successful build. A combined build command timed out behind an orphaned MSBuild worker retaining the lock; the worker exited, and queued native checks completed. Subsequent lock commands use `flock --close` so workers cannot inherit the lock descriptor.

Heavy work remained serialized with two .NET processors and a 4 GiB managed-heap cap. Available-memory snapshots were 16 GiB initially and 17 GiB after preview cleanup; swap stayed at 2.7 GiB, and root free space was 93/91 GiB, with other sessions active. Owned browser tabs, preview service and scratch files were removed. No full release gate, native-window, Firefox/Safari or real Word/Excel oracle was run.

The count/function-assistance batch adds live Word body totals and active-story selection counts without replacing action/error messages. Totals include nested table cells and body text boxes once; headers, footers, notes and comments are excluded from body totals but their active selections are counted. Counts reuse Unicode-whitespace range semantics and walk rope leaves without flattening text. Immutable body-block identity caches totals across caret-only changes. Partial-word fragments, hidden/tracked text and object/reference placeholders follow the documented model semantics, not Microsoft Word linguistic/revision-view parity.

Spreadsheet's single help catalogue now covers **66 functions**, adding 45 common criteria, logical, lookup, dynamic-array, date/time, text and numeric functions. Chooser insertion, Tab completion and both inline/formula-bar hints share it. Repeating argument groups have explicit typed start indexes rather than parsing presentation labels on each hint. Exact-name completion wins; ambiguous prefixes retain ordinary Tab behavior without guessing a function. Descriptions document evaluator limits, including INDEX's single-cell result, UNIQUE without exactly-once, approximate lookup defaults and fixed weekend days.

Verification: fresh Office.Formats, Word and Spreadsheet Fable compilation and both native Wisp builds passed. Chromium regressions passed **17 shell word-count**, **35 nested-table editing/count** and **42 function-help assertions** (**94 total**). **22 targeted surface assertions** covered actual keyboard typing/selection/undo/redo, nested-table and text-box imports, header-story counting, date/criteria/dynamic-array formula entry, compact chooser/hint placement and XLSX reopen. Three XLSX outputs passed package-integrity checks. A controlled progressive package retained its ten-block prefix when counted (20 words), updated to 1,000 words on completion, and displayed a cancelled prefix as “Body (loaded)”. Other existing shell queries can complete progressive packages; the count API itself does not.

Actual controls were inspected at 1200px and 390px. Final browser surfaces recorded no errors. Verification setup required a result-marker hook and mounting the preview grid; neither was an application defect. Singular word labels were corrected, and count tests now assert displayed numbers rather than pinning incidental wording. No new evaluator behavior, full release gate, native-window, Firefox/Safari or real Word/Excel oracle was exercised.

Builds and browser checks remained serialized; .NET used two processors and a 4 GiB managed-heap cap, with non-inherited lock descriptors. Available RAM was 16 GiB initially and 20 GiB after cleanup; swap remained 2.7 GiB. Root free space was 89/87 GiB with concurrent sessions active. Owned preview tabs, service and temporary compilation files were removed.

The Word context-menu batch adds a shell-owned canvas adapter rather than
changing bare-widget browser menus. Right-click preserves a selected range
when the hit is inside it and otherwise targets the clicked caret. Keyboard
entry, arrow/Home/End traversal, dismissal/focus restoration, viewport clamping
and scrolling share the shell's existing command actions. Clipboard commands
use the real browser API; cut waits for transfer success, and delayed cut/paste
cannot mutate a changed document or selection. Read-only and composition
boundaries are enforced, with permission failures reported rather than hidden.

Word Fable compilation and the native Word-wisp build passed with existing
warnings. The new Chromium context regression passed **20 assertions**,
including failed/delayed clipboard writes, stale paste, undo, link removal,
read-only/composition isolation and disposal. **10 actual-surface assertions**
passed for real secondary-click targeting, native clipboard copy/paste, font
dialog focus, table commands, 390px menu bounds and independent mounted shells.
Inserting a row through the menu changed the table from two to three rows.
Desktop and narrow screenshots were inspected; the final surface reported no
browser errors. No Firefox/Safari, native-window or real Word certification.

Spreadsheet chart insertion now supports classic clustered columns, horizontal
bars, lines and single-series pies through the grid API and browser Insert
menu. A bounded contiguous source has explicit header/category flags, columns
as series, an optional title and a 20,000-value limit. Multi-series pie requests
are rejected, not truncated. Authored parts belong to worksheet history and
use existing drawing rendering/reference rewriting and source-aware export,
without an export/reimport workaround. Existing picture manipulation APIs
remain picture-only; chart selection/move/resize/delete and configuration
editing are still separate work.

Integration fixed F# indentation/type inference, qualified numeric-value cases,
and line-chart OOXML ordering (series precede marker and axis declarations).
A live deletion scenario exposed insertion-time caches resurrecting removed
chart values/categories; changed authored references now invalidate those
caches, and export refreshes them from current cells.

Final Office.Formats/Spreadsheet Fable compilation and native Spreadsheet-wisp
compilation passed. Chart regressions passed **30 API**, **6 shell**, and
**46 imported-chart/reference/preservation assertions** (seven informational
log lines excluded). Together with Word, this batch passed **102 assertions**.
The chart checks validated **16 exported packages**, with no new integrity
problems. Ten additional chart surface assertions brought the batch's actual
UI total to **20**: source/title/category rendering, formula-driven bar heights,
undo/redo, visible pie validation/correction, 390px dialog fit, and real File
Save/Open. The downloaded workbook reopened with both column and pie drawings
and passed an additional package-integrity check. Final surfaces reported no
browser errors. This is Chromium evidence, not real Excel or cross-browser
certification; no full release gate or native-window workflow was run.

Heavy jobs used the shared non-inherited file lock, two .NET processors and a
4 GiB managed-heap cap. Available RAM snapshots were 20 GiB initially, 19 GiB
before chart verification and 15 GiB after it; swap remained 2.7 GiB. Root free
space was 87/83 GiB with other sessions active.
Owned preview tabs, services, temporary compilation output and the downloaded
verification workbook were removed after the checks.

## Recommendation

Keep the existing models, format parsers, layout engines and corpus. Do not rewrite either editor. Build one platform-neutral editing session per editor, one source-preserving save route, and thin platform adapters. Finish trustworthy open/edit/undo/save workflows before adding more feature flags or UI specifications.

The principal structural weakness is not merely file size: browser closures own transactions, selection, invalidation and host integration, while desktop hosts implement different subsets of those policies. A shared model does not make two independently orchestrated editors behaviorally shared.

Deliver three explicit products from the same engines: an embeddable widget without chrome, a complete browser application shell, and a .NET desktop application. Browser alpha can precede desktop, as the current roadmap specifies, but extraction must preserve the desktop destination rather than produce a second implementation later.

## Shared architecture and delivery findings

### C1 — P1: the solution contradicts the package-only build contract

**Source-proven portability defect.** `Office.slnx:51-58` includes eight projects from sibling repositories. `AGENT.md:11` and `Directory.Build.props:2-6` explicitly require registry packages, not neighbouring checkouts. Both spreadsheet native hosts correctly reference the in-repository Spreadsheet project (`spreadsheet-native/Rbt.SpreadsheetWidgets.fsproj:11`, `spreadsheet-wisp/Rbt.SpreadsheetWisp.fsproj:17`).

A fresh checkout without these siblings cannot resolve the solution graph. A developer with them can accidentally build unrelated source alongside pinned packages. The sibling paths all exist on this workstation: a read-only existence probe returned eight `true` values, explaining why local use can conceal the portability problem. No clean-machine build was attempted.

**Plan:** remove sibling solution entries, retaining the pinned package references. **Acceptance:** restore/build the supported solution and hosts in a checkout containing only this repository and registry credentials.

### C2 — P1: Spreadsheet.Core is a stale, non-self-contained project

**Source-proven project-graph defect; compiler result not run.** `src/Spreadsheet.Core/Rbt.Spreadsheet.Core.fsproj:8-18` links Model and Paint but omits TableStyles, Charts and Slicers. Model uses `TableStyles.TableStyleElement` at `src/Spreadsheet/Model.fs:1021` and `TableStyles.builtIn` at `3224`; Paint directly uses `Rbt.Spreadsheet.Charts.Scene` at `src/Spreadsheet/Paint.fs:1044`. The project has no references providing these modules. It is nevertheless included in `Office.slnx:18-20`.

The main Spreadsheet project includes these dependencies in compile order; native hosts use that main project rather than Core. Thus the advertised split is not the actual dependency boundary, and the ordinary test gate does not exercise this solution project.

**Plan:** either make Core the real platform-neutral dependency used by all hosts, or remove the obsolete project when establishing the real split. Do not maintain duplicate compile lists for nominal architecture. **Acceptance:** every project in the supported solution builds and has an identified consumer or explicit tooling purpose.

### C3 — P1: publishing is not coupled to regression evidence

**Source-proven release risk.** `.forge/actions.lua:2-6` selects publishing for version tags; `.forge/publish-npm:24-32` restores tools, installs dependencies, builds, packs and publishes. It does not execute tests, browser harnesses or a check for a previously validated commit. The roadmap prescribes a manual gate, but the publish action does not enforce it.

This does not establish that existing releases were untested. It establishes that a compiling, behaviorally broken tag can be published through the checked-in action.

**Plan:** require a successful gate for the exact commit being packaged; retain external Word/Excel oracle evidence as a release artifact where automation needs a Windows runner. **Acceptance:** a deliberate regression prevents publication, and the published tarball is the one exercised by the consumer smoke test.

### C4 — P2: test runners can report success after required coverage disappears

**Source-proven.** `tests/run-all.sh:36` skips missing explicitly named suites; `tests/formula/run-all.sh:81` similarly skips missing stages. Both can subsequently print an all-passed message. `tests/run-all.sh:56-62` makes native painter execution depend on a pre-existing DLL, and the root script never selects `tests/spreadsheet/wisp-editorhost.fsx`.

The issue is coverage integrity, not a claim that current listed files are missing. Another smaller mismatch: root `--skip-build` is not propagated to the formula runner, which always builds (`tests/formula/run-all.sh:9-16`).

**Plan:** fail on missing mandatory suites; define explicit browser/core/native gates and report optional exclusions separately. Native verification should build the intended output rather than infer freshness from file existence. **Acceptance:** removing a mandatory fixture/suite fails the gate; a clean checkout's reported coverage matches a developer checkout's.

### C5 — P1 for the full browser shell: shared dialog specs exceed renderer capability

**Source-proven; previously documented in `docs/ui-audit.md:53,102-103`.** `src/OfficeUi/Specs.fs:176-204` defines section, colour, slider, text-area and radio fields. `src/FableUI/OfficeUiDialog.fs:60-133` handles none of those variants. `openDialog` initializes only `spec.Fields` (`45-49`) and rendering uses only those fields (`153-176`), despite the `Tabs` contract at `Specs.fs:234-239`.

Consequences: advanced dialogs can be empty or throw; adding a spec is not shipping a feature. This is not an engine limitation and should not be fixed independently in every dialog.

**Plan:** complete one renderer and one state/validation path, then connect dialogs to real selection values and commands. **Acceptance:** every supported field kind and tab renders; switching tabs preserves edits; OK applies one undoable transaction; Cancel applies nothing; validation errors keep the dialog open.

### C6 — P2: shared dialogs lack basic keyboard and accessible-name behavior

**Observed in Chromium against the existing `dist/fable-ui` modules**, using an in-memory fixture with one text field and a primary OK button. No page or test file was added.

- Opening left focus on `BODY`.
- Focusing the text field and pressing Enter produced `submitted: 0`; the dialog remained visible.
- Escape left the dialog present.
- The input's `labels.length` was `0`; browser observation reported a textbox with an empty accessible name.

`src/OfficeUi/Specs.fs:220` promises Enter activation for the primary button. `OfficeUiDialog.fs:135-149` wires only mouse click; `153-176` provides neither focus/keyboard handling nor an accessible title association. Text labels are siblings without `for`/input IDs (`62-71`). This is a real shipped-renderer problem, not just a missing end-user menu.

**Plan:** centralize initial focus, focus containment/restoration, Enter/Escape semantics, label associations and dialog title association in the renderer. Preserve multiline Enter behavior. **Acceptance:** keyboard-only open/edit/submit/cancel, focus restoration to the editor, and named controls in the accessibility tree.

### C7 — P2: architecture/status documents mix current facts and historical snapshots

**Source-proven documentation inconsistency.** `docs/sharing-architecture.md:103` and `docs/ui-audit.md:87` describe Spreadsheet PDF as .NET-only. The current `package.json:13` compiles `src/Spreadsheet.Pdf`, and `src/Spreadsheet/Spreadsheet.fs:8025-8089` exposes `ExportPdf` and `PrintPdf` using a lazy-loaded browser module (`164-165`). Those are no longer absent APIs; end-to-end quality was not exercised here.

The roadmap also mixes completed work and old next steps. Treat historical benchmark/parity numbers as dated evidence, not current release certification.

**Plan:** maintain one current capability matrix with links to dated evidence. Retain historical audits, clearly labelled snapshots. **Acceptance:** each claimed capability says separately whether it is parsed, preserved, rendered, editable, exposed through an API, reachable through the shipped shell, and tested on each platform.

## Word findings

### W1 — P1: the documented paragraph-only save route loses package content

**Source-proven API hazard.** `docs/word-host-api.md:443` recommends `exportToDocx` with `WordWriter.writeFromEditor`. `src/Word/Word.fs:4047-4053` returns paragraph objects; `src/Office.Formats/src/DocumentWriter.fs:5878-5911` creates a fresh package without a preservation input or header/footer parts. The preservation-aware route exists separately at `Word.fs:4063-4073` and `DocumentWriter.fs:5918-5929`.

**Trigger:** a host follows that documented route to save an imported document with headers, footers or other package content. A readable output file can silently omit that content. The paragraph-only route also does not end the typing run, unlike `exportDocumentPackage:4072`.

**Plan:** make source-preserving package export the sole document-save contract and migrate callers. If paragraph conversion is genuinely required, label it explicitly as lossy conversion, never document save. **Check:** exercise every advertised save entry point with headers, footers, styles, notes and opaque parts; compare saved content and save-mid-typing undo boundaries. Existing `checks/revision.js` and `checks/save-undo-boundary.js` cover the package route, not this alternative.

### W2 — P1: PDF shaping and embedding can choose different fonts

**Source-proven.** `src/ShaperText/ShaperText.fs:177-190` prefers the requested family's regular face before another family. `src/Word.Pdf/Pdf.fs:236-247,271-281` instead falls back from the requested exact style to the default family's exact style. At `Pdf.fs:613-645`, the returned shaping font is ignored (`Some (_, shapedText)`), while its glyph IDs are written using the independently selected PDF font.

**Trigger:** provide family A regular and family B bold, use B as default, and request A bold. Glyph IDs shaped with A can be interpreted in B, producing wrong/missing glyphs and positions.

**Plan:** carry the resolved shaping face and embedding identity through the text pipeline once. **Check:** mixed-family missing-bold/italic fixtures and external shaper input; compare actual glyphs, extracted text and embedded font identity. `tests/word/pdf.fsx:61-73` uses a single regular face and does not cover this fallback conflict.

### W3 — P2: undo can restore content without restoring its saved revision identity

**Source-proven.** `src/Word/Word.fs:100-121` caps `KnownRevisions` at 400 observed states. `src/Word/Input.fs:612-624` coalesces consecutive characters into one history command retaining the original undo target. The history's command limit does not therefore protect that original revision from eviction.

**Trigger:** type more than 400 content changes in one continuous run, then undo to the saved/imported document. The old content can receive a new revision, leaving it marked modified and disabling the unchanged-byte export path (`Word.fs:4069`).

**Plan:** bind revision identity to retained history/content states, preserving saved/opened identities while relevant. **Check:** long coalesced typing, undo, redo, saved checkpoints, modified flag and byte-identical unchanged export. Current short revision checks do not cover eviction.

### W4 — P2: progressive-load waiters do not have reliable cancellation settlement

**Source-proven.** `Word.fs:581-588` registers `whenLoadComplete` waiters that read mutable `LoadFailure`. Disposal clears waiters without invoking them (`286-297`). Replacement calls `stopLoading` (`216-245`) before setting the new failure state.

**Trigger:** await completion, then dispose or replace the still-loading document. Disposal leaves the promise pending; replacement can resolve the outgoing cancelled load as successful.

**Plan:** give each load generation a terminal result and settle every waiter exactly once on success, failure, replacement or disposal. **Check:** retain and await outgoing promises through all transitions, including a shared package displayed by another viewer. `checks/progressive-cancel.js` covers explicit cancellation but not these outgoing-waiter cases.

### W5 — P2: editor disposal leaves canvas listeners alive

**Source-proven.** `Input.fs:1881,2049,2055,2194` attaches keydown, focus, mousedown and mousemove listeners. Disposal (`2329-2337`) removes the window mouseup listener and input proxy, not those canvas listeners; keydown has no disposed guard (`1481-1502`).

**Trigger:** dispose while retaining the canvas, or mount a replacement editor on it. Old handlers can retain and act on the old document; repeated mounts accumulate handlers.

**Plan:** own every listener and pending input callback in the editor lifecycle. **Check:** events after disposal produce no edits or paints; repeated remounts yield one response per event. Existing IME checks verify proxy removal, not canvas reuse.

### W6 — P2: plain-text equality is incorrectly treated as rich clipboard provenance

**Source-proven.** `Input.fs:319-320` has a module-global `lastClip`; copying stores rich content (`1381-1386`), and paste chooses it ahead of supplied HTML if normalized plain text matches (`1393-1404`).

**Trigger:** copy bold “Hello” in the editor, then italic “Hello” from another application, then paste. The old rich content wins over the external HTML. The cache spans editor instances.

**Plan:** use explicit provenance for internal rich payloads; text equality is not ownership. **Check:** identical text with different formatting/links, multiple editors, and text-only external overwrites. Existing paste checks use different external text or deliberately exercise the cached internal path.

### W7 — P2: asynchronous cut mutates before clipboard success

**Source-proven ordering; failure-path impact is conditional.** `Input.fs:1385-1386` deletes during `copyToClipboard true`. The canvas fallback (`1682-1691`) then calls `Clipboard.copyText` and ignores its promise; `src/Word/Clipboard.fs:7-9` calls `navigator.clipboard.writeText`.

**Trigger:** that fallback is used and permission/embedding policy rejects the write. Content disappears without reaching the clipboard; undo can recover it.

**Plan:** commit the captured cut only after successful transfer, checking that intervening edits have not invalidated the captured operation; report failure. **Check:** denied/unavailable clipboard and edits while the write is pending. Successful proxy clipboard checks do not exercise this failure path.

### W8 — P2: virtual paragraph input does not share physical Enter semantics

**Source-proven input-path gap; precise browser outcome is [INFERENCE].** `Input.fs:1991-2015` handles text, deletion and `insertLineBreak` but not `insertParagraph`. The textarea fallback (`2038-2045`) inserts text; physical Enter invokes paragraph/list operations (`1819-1846`).

**Trigger:** an input method emits `insertParagraph` without the handled keyboard event. It can bypass paragraph/list splitting and instead insert a newline as text.

**Plan:** route both event sources to the same edit intent. **Check:** supported virtual keyboards and controlled beforeinput cases in ordinary paragraphs, lists, tables and tracked changes, including undo. Existing IME coverage omits paragraph insertion.

### W9 — P1 for desktop delivery: the Word desktop shell is not an operational full editor

**Source-proven feature gap, not a browser regression.** `word-wisp/Program.fs:69-89` wires Open/Save/PDF and numerous dialog/formatting actions to `ignore`; selection formatting is hardcoded (`65-68`). The browser and desktop also have different history implementations (`word-native/WordEditor.fs:506-515,700-717`; `word-wisp/EditorHost.fs:134-149`).

**Plan:** wire the selected desktop shell to the shared session and typed source-preserving IO, not another set of host-local commands. **Check:** real file dialogs, save/reopen, unsaved-change prompts, system clipboard, actual selection state and interactive IME. A successful build or screenshot cannot establish these behaviors.

## Spreadsheet findings

### S1 — P1: structural edits leave foreign-sheet rule formulas pointing at old cells

**Source-proven.** `src/Spreadsheet/Workbook.fs:570-582` uses `rewriteFormulasForEdit` for other sheets. That function visits formula cells only (`src/Spreadsheet/Model.fs:3454-3469`). Names/tables receive separate handling, but conditional formats and validations on other sheets do not. Undo records foreign cells, tables and names only (`Workbook.fs:605-649`).

**Trigger:** a validation list on Other references `Data!$A$2:$A$5`; insert before Data row 2. Values move, but the rule still points at the old range. Foreign conditional expressions and formula thresholds have the same ownership gap.

**Plan:** reuse a shared formula-owner traversal, updating references without moving the foreign rules' application ranges; include the changes in undo. A broader walker already exists at `Model.fs:3474-3511`. **Check:** row/column insert/delete, undo/redo and save/reopen with foreign validation/custom/conditional/threshold formulas. Stage21 tests local rule movement and foreign cell formulas/names, not this combination.

### S2 — P1: chart anchors move, but chart bindings retain old references

**Source-proven.** `Spreadsheet.fs:958-999` reads drawing parts from the imported package and caches parsed charts by path. ChartEx names come from original `source.Workbook.NamedRanges` (`991-997`), while data lookup uses current sheet names and requested coordinates (`1015-1023`). Structural changes move anchors (`Model.fs:4034`), not series bindings.

**Trigger:** insert above data referenced by `Data!$B$2:$B$5`, or rename Data. The chart can read the wrong coordinates or fall back to cached imported values even while its object anchor moves.

**Coverage problem:** `tools/xlsx-browser/component.js:657-669` expects chart parts to remain byte-identical after structure edits. That is only correct for unaffected parts; it is not a valid requirement for references that must change. Existing live-cell-refresh checks do not establish structural/rename correctness.

**Plan:** include series, categories, chart names and chartEx names in workbook reference ownership; patch affected XML and invalidate bindings. **Check:** plotted values and saved/reopened references after insert/delete/rename/name edits and undo/redo—not just SVG existence or position.

### S3 — P1 for cut / P2 for copy: internal clipboard state belongs to a sheet, not the workbook session

**Source-proven.** Copy/cut stores `model.Clipboard` (`Spreadsheet.fs:1581,3220`); switching sheets replaces the active model (`6894-6900`). Internal paste requires that target model's clipboard (`2316-2321`), and denied clipboard access falls back to it (`3403-3410`). Internal cut repointing supplies the active sheet as both source and target (`1613-1619,1645-1651`).

**Trigger:** copy/cut a formula/formatted range, switch sheets and paste. The internal rich move is lost; external clipboard parsing may substitute different semantics, dependent references do not follow a cross-sheet move, or no paste is available without clipboard access.

**Plan:** session-owned clipboard with source sheet identity; use workbook move operations and workbook-aware undo. **Check:** formulas, formatting, dependents and names across sheets, permission allowed/denied, undo/redo and reload. `stage24-moves.fsx:149-152` checks the helper, not this browser clipboard lifecycle.

### S4 — P1 for desktop: native clear and multi-sheet editing do not perform correct recalculation

**Source-proven.** Multi-sheet edits use `Formula.recalcAndSpill None None edited` without workbook context (`spreadsheet-wisp/EditorHost.fs:273-289`; `spreadsheet-native/SpreadsheetEditor.fs:303-319`). Wisp Delete clears cells through `applyChange` (`567-573`), whose implementation only invalidates calculation and installs the model (`86-94`). Widgets has the same omission (`SpreadsheetEditor.fs:92-101,626-628`) and clears only the focused cell rather than the selected range.

**Trigger:** A1=2, B1=`=A1*5`; delete A1. B1's cached result is not recalculated by this path. Cross-sheet edits also lack the required workbook propagation.

**Plan:** route every value-changing command through the shared workbook mutation/recalculation transaction. **Check:** deleting precedents, cross-sheet chains, spills, cut/paste, range deletion, undo/redo and exported caches. Native self-tests check deletion of a formula, not a precedent with dependents.

### S5 — P1 for desktop save: Wisp reconstructs an incomplete workbook and exports stale names

**Source-proven.** `spreadsheet-wisp/EditorHost.fs:100-115` reconstructs workbook state with empty Tables, Names and NameInfo, and no Iteration. Structural operations retain only changed sheet models (`118-137`). `spreadsheet-wisp/Program.fs:99-101` writes imported `source.BaselineNames` on save.

**Trigger:** open a defined name referring to `Data!$A$2`, insert before row 2, save. The cell moves while the name remains at its imported address. Not having a names dialog does not make names immutable under structural edits.

**Plan:** retain and mutate the complete imported workbook context, then export current metadata. **Check:** global/local names, tables and iterative settings across native open/edit/undo/save/reopen; compare with browser semantics.

### S6 — P1 for desktop integrity: sheet operations bypass both reference rewriting and complete undo

**Source-proven.** Wisp insert/rename/delete directly mutate sheet/name arrays (`Program.fs:310-328,371-380`). History stores only `(SpreadsheetModel array * int)` (`EditorHost.fs:33-34,139-157`), not sheet names or other workbook metadata.

**Trigger:** edit a cell, add/delete a sheet, then undo. Old sheet models can be restored with the separately modified name array. Renames leave formula names stale; sheet deletion does not coherently rewrite references; these sheet actions themselves are not normal undoable workbook transactions.

**Plan:** use stable sheet identity and workbook-level sheet operations; history covers the complete transaction. **Check:** interleave cell edits and sheet insert/rename/delete, undo/redo the whole sequence, and verify identities, names, order, active sheet, formulas and saved content.

### S7 — P2: known full-height zoom and pivot-geometry limitations remain

**Documented and source-supported, not newly reproduced.** `Spreadsheet.fs:1891-1921` applies zoom and allocates a full-sheet-height element without compensation for browser layout extent limits. Exact bottom-reachability thresholds depend on browser/runtime. `Model.fs:4027-4068` updates structural metadata without moving `Pivots`; `Formula.fs:12071` uses their stored bounds for GETPIVOTDATA.

**Plan/check:** test maximum-height navigation and selection together with zoom, frozen panes and hidden/custom-sized rows; use segmented/logical scrolling if browser extent limits require it. Move pivot layouts with structural edits and invalidate dependent calculations. Treat pivot-cache refresh as a separate authoring capability rather than claiming that displayed imported pivots are fully editable.

## Structural critique and target boundaries

### Retain

- Typed document/workbook models and functional edit operations.
- Source-aware OOXML preservation rather than reconstructing every imported file from a lossy model.
- Layout and paint-command seams, with surface-specific adapters.
- Real-file corpus, formula oracle, browser harnesses and external Word/Excel comparisons.
- Registry-pinned external Rbt dependencies.
- Existing ZIP resource controls: `src/Office.Formats/src/ZipInterop.fs:76-103` defines entry/count/total budgets. This review does **not** claim ZIP decompression is unbounded.

### Change

1. **Own editing state in one session per editor.** The session owns document/workbook, selection/story or active sheet, undo transactions, revision/dirty state, edit coalescing, recalculation and layout invalidation. Browser/native adapters translate events and schedule work; they do not invent edit semantics.
2. **Keep format IO separate from UI.** A typed import/save representation should carry source spans, relationships and opaque parts. JavaScript object conversion belongs at the public boundary, not throughout the writer. Save must consume a coherent revision without mutating the active editing session.
3. **Share semantic rendering, not necessarily the rendering technology.** DOM spreadsheet cells can remain if needed for interaction/accessibility/performance. Resolve alignment, styles, overflow and layout geometry once; do not independently implement the same decisions in CSS and the paint backend. Require parity examples before replacing the DOM grid wholesale.
4. **One command vocabulary, optional chrome.** A command has availability, current/mixed state and execution semantics. Menus, toolbar, shortcuts, context menus and host APIs reach the same operations. An embedded widget requests host-owned dialogs instead of unconditionally opening legacy component modals.
5. **One UI specification/state path per platform renderer.** Reuse `OfficeUi` where it actually models the interaction; do not force document editing into a generic form framework. Remove obsolete shell/dialog variants only after their commands and behavior have migrated.
6. **One supported desktop shell.** Evaluate the existing Wisp direction against required input/IME/accessibility behavior. Select it explicitly rather than maintaining two independent desktop editor implementations indefinitely. Keep renderer test tools separate from product hosts.
7. **Split by responsibility, not arbitrary line counts.** Extract session/history, clipboard, source IO, layout scheduling, view geometry, command dispatch and shell composition incrementally. Keep editor-specific models separate; avoid a universal Word/Spreadsheet document abstraction.

Desired dependency direction:

```text
Browser shell / native shell / external host
                 |
        Command and query facade
                 |
       WordSession / WorkbookSession
         /          |            \
   editing model  format IO   layout / render semantics
                                  |
                        DOM / canvas / native / PDF
```

Platform services provide clipboard, font bytes/measurement, scheduling, file access and surface input. The core never calls browser globals to perform an editing transaction.

## Common delivery sequence

1. **Trust and reproducibility:** repair project graph and mandatory gates; define supported browsers and capability matrix; establish clean-checkout and packed-package consumer checks. Address data-loss and false-success findings before new authoring features.
2. **Extract sessions through existing behavior:** migrate a complete open/type/format/undo/save flow first, then structural edits and rich content. Preserve current public behavior during the implementation, but remove obsolete internal paths at each completed cutover rather than leave permanent dual implementations.
3. **Complete browser applications:** a real demo mounts each shipped shell; all visible commands work, reflect selection state and are undoable. Implement keyboard/accessibility behavior while wiring the shell, not after feature completion.
4. **Cross-platform parity:** connect the selected desktop host to the same sessions and IO. Run interactive desktop and browser scenarios with fixed font assets; .NET compilation alone is not an acceptance criterion.
5. **Broaden features using evidence:** rank requested workflows and actual corpus features. Do not chase a larger function/dialog count while common operations still lose information.

Release acceptance should include open → edit → undo/redo → save → reload, interoperation with real Word/Excel, and DOM/native/PDF comparisons for the supported subset. Unsupported features must remain preserved where promised and be visibly identified when not editable. Performance budgets must name fixture, operation, machine/browser, font assets and measurement method; establish controlled baselines before choosing thresholds.

## Word build-out plan

### W-A — Make embedded editing and saving trustworthy

**Dependencies:** shared gate/project cleanup; W1-W7. Fix the save contract first, then revision/lifecycle and clipboard defects; fix font identity before claiming PDF fidelity.

**Deliver:** a stable no-chrome editor contract for open, edit, selection, undo/redo, save, load completion and disposal. Start extracting `WordSession` around these operations rather than exposing browser-owned history as the shared interface.

**Exit criteria:** long typing/undo returns to a clean saved state; cancelled/disposed loads settle; remount does not duplicate input; failed clipboard operations do not delete content; all advertised save routes preserve supported and opaque package content; PDF uses the shaping font's glyphs. Test exact semantic/XML preservation, not only file validity or text length.

### W-B — Consolidate layout, text and rich editing behavior

**Deliver:** shared story/selection/pending-format state, paragraph and list intents, rich clipboard provenance, one typed source-preservation contract on browser and .NET, and consistent resolved font identities for measurement/render/PDF. Keep the established layout algorithms; separate scheduling and viewport policy from the browser renderer.

**Exit criteria:** the same edit-intent traces yield equivalent documents, selections and undo boundaries on both runtimes. Compare page starts, line wraps, table continuation, note placement and caret/hit-test geometry under fixed font assets. Page counts alone are insufficient. Address documented content-control, cross-cell copy and screenshot-paste gaps as explicit workflows, not silent fidelity promises.

### W-C — Ship a complete browser writing application

**Deliver in workflow order:** open/save and dirty/error state; keyboard/IME and find/replace; font/paragraph/styles and lists; tables/pictures/links; page setup/sections/headers/footers; notes/comments/review; print/PDF. Build real selection-driven dialogs, not hardcoded demonstrations.

**Exit criteria:** each visible action works, reports capability/availability, updates actual selection state, participates in undo and survives reload. Dialog focus returns to the correct story/caret. Chromium and Firefox pass agreed workflows; touch/virtual keyboard and accessibility have explicit supported behavior.

### W-D — Complete the selected desktop application

**Deliver:** the same `WordSession` and preservation-aware IO behind native event/render/file/clipboard adapters. Replace `ignore` callbacks and hardcoded formatting indicators; retire duplicate native edit/history behavior after cutover.

**Exit criteria:** interactively open a real DOCX, edit text/tables/stories, undo, save/reopen and export PDF; exercise IME, system clipboard, viewport scrolling and unsaved-change/error handling. Native and browser fixture traces agree. Macros, collaboration and advanced authoring should not displace these foundations.

## Spreadsheet build-out plan

### S-A — Establish one authoritative workbook transaction boundary

**Dependencies:** C2 and S1-S6. Retain workbook context, stable sheet IDs, complete history, revision state and session-owned clipboard. Use the existing workbook/recalculation primitives rather than another native-specific calculator.

**Deliver:** all edits pass through `WorkbookSession`, including value/clear/paste/fill/sort, structure changes and sheet lifecycle. A reference-owner inventory covers cell formulas, names, tables, validation, conditional formats, charts and pivot geometry. Distinguish changing a reference from moving the range to which a rule applies.

**Exit criteria:** insert/delete/move/rename and undo/redo update every affected owner; cross-sheet cut preserves formula semantics; desktop deletion recalculates dependents; save/reopen retains current metadata. Compare formulas/results/references as well as unchanged source bytes.

### S-B — Finish dependable grid and host behavior

**Deliver:** consistent cell-look resolution across DOM/native/PDF, full-height logical navigation, correct zoom/freeze/hit testing, formula bar and workbook-aware clipboard. Keep DOM virtualization unless measurement proves replacing it is necessary.

**Exit criteria:** reach and edit the last row/column at 10%, 100%, 125%, 200% and 400%, with frozen panes, custom sizes and hidden/filtered rows. Test Chromium and Firefox, not just small sheets at default zoom. Establish controlled budgets for large-sheet scroll, edit/recalc, row/column insertion and sorting; compare against the existing 846,000-cell fixture without concurrent heavy jobs.

### S-C — Build the complete spreadsheet browser shell

**Deliver in workflow order:** formula entry/navigation and formatting; format cells/custom number formats; sort/filter and tables; conditional-format and validation editors; named ranges and sheet-tab actions; find/replace; print/page setup/PDF. Then formula assistance and chart insertion/editing on top of correct binding updates.

**Exit criteria:** toolbar/menu/shortcut/context-menu actions share commands, availability and mixed-selection state. No phantom commands or legacy modal popping over a host-owned panel. Every editing dialog supports edit → undo/redo → save/reopen. Browser PDF already exists: validate and expose it, do not rebuild it as a supposedly missing engine.

Chart and pivot scope must be explicit: imported rendering, editable chart series, cached pivot display, pivot refresh and formula PIVOTBY are different capabilities. Add authoring only when its mutation/save/reference lifecycle is complete.

### S-D — Complete desktop using the same workbook session

**Deliver:** native rendering/input/clipboard/file adapters around the same workbook context and recalculation/history. Make zoom affect geometry, not only a toolbar signal. Replace direct array mutations and stale-baseline export; remove superseded native command implementations.

**Exit criteria:** browser/native command conformance for cell edits, cross-sheet formulas, structure/sheet operations and undo; native interactive proof for focus, IME, scroll, copy/cut/paste, drawings and PDF. A real edited workbook reopens correctly in Excel. Do not advertise desktop feature parity based on compilation.

## First implementation queue

1. Close the dangerous save routes and data-reference holes: W1, S1-S3; repair PDF font identity W2.
2. Make the build/release graph and coverage truthful: C1-C4. Browser release can proceed separately from desktop, but must state that scope.
3. Fix Word state/lifecycle/clipboard invariants W3-W8 while extracting its session.
4. Establish the complete workbook session and migrate native operations S4-S6 before further desktop feature work.
5. Complete the shared dialog renderer C5-C6, then wire the two real browser shells in the workflow order above.
6. Address layout/scroll/parity gaps with controlled fixtures; complete desktop adapters and then expand advanced authoring.

The order prioritizes silent semantic loss over cosmetic parity. Workstreams may overlap once command/session contracts are fixed, but a second UI implementation must not become a second source of editing semantics.

## Verification performed and limits

- Read-only source/project/script inspection, including the referenced conflicting contracts.
- Browser execution of the existing shared dialog renderer with a temporary in-memory fixture: keyboard and accessibility results recorded under C6. The tab and temporary static server were closed.
- Read-only existence probe for the eight sibling projects: all exist locally; portability failure is inferred for an isolated checkout, not reproduced here.
- No fresh compilation of `dist`; browser results apply to the existing generated artifacts and align with the cited source paths.
- No full regression, Office VM oracle, performance run, native interactive session or Firefox/Safari session. Historical results are not represented as results of this review.
