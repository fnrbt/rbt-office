#!/usr/bin/env node
// Browser checks for the Word editor's browser-only code (Input.fs, Word.fs,
// Renderer.fs), run in headless Chromium against Fable output. The .NET suites
// cannot reach this code (it is #if FABLE_COMPILER / jsNative).
//
// Usage:
//   dotnet fable src/Office.Formats --outDir $DIST/office-formats --lang JavaScript --noCache
//   dotnet fable src/Word           --outDir $DIST/word           --lang JavaScript --noCache
//   node tools/word-browser/run.mjs --dist $DIST [check ...]
//
// Checks (tools/word-browser/checks/<name>.js, all by default):
//   explicit-format  toggle bold off under a bold style, save, undo/redo
//   page-frames      page frames / page count / hit-testing follow section page setup
//   ime              composition (CJK, dead keys), beforeinput/input, undo steps, dispose
//   typing           real mouse click and keyboard typing through the input proxy
//   track-changes    tracked typing and Enter through the keyboard, saved as w:ins and a tracked mark
//   linked-images    a linked picture is not requested until the host allows it; unsafe links are inactive
//   opaque-click     clicking, dragging and typing next to a chart placeholder keeps the chart on export
//   selection-format getSelectionFormat (value / "mixed"), canUndo / canRedo, onChange and unsubscribing
//   viewport         setViewport (zoom / devicePixelRatio / scroll): pixels and hit tests against whole-canvas mode, renderPage timing
//   placeholder-and-frames  the chart placeholder's label is drawn at its own size, centred; content-control frames only around the caret's control
//   caret-movement   Up / Down by visual lines keeping the column, Home / End, Ctrl+arrows, PageUp / PageDown, Ctrl+Home / End, Shift variants, double / triple click
//   revision         getRevision / markSaved / isModified across edits, undo and redo
//   viewer-measurers a viewer with the editor's measurers has the editor's pages; renderPage uses the text painter made for its canvas
//   page-centering   narrower pages are centred through a viewport (X in getPageFrames, pixels, hit tests); nothing is painted between pages
//   image-handles    a picture found in the painted pixels is resized by its corner handle, selected by a click, in whole-canvas and viewport modes
//   commands         point sizes, colours, fonts, styles, spacing, page break, table rows / columns / merge / split, hyperlinks
//   painter-frames   a plain-function text path painter / measurers: the paint frame is built when one is installed or cleared, never while scrolling
//   viewer-entry     word/WordViewer.js imports no editor module (import graph of the Fable output) and is the whole viewer
//   keystroke-long   the time of a keystroke in a 75-page and a 600-page document (see also perf.mjs)
//   table-tools      insertTable (borders, caret rectangle, Tab / Shift+Tab between cells), getCaretContext, no internal style ids, spacing in the selection format, markSaved(revision)
//   caret-frames     a caret move / selection change builds no paint frame (getRenderStats().FrameBuilds); the pixels are those of a new frame
//   viewer-first-page  the viewer's getRenderStats; the first renderPage collects the commands of the page's paragraphs only; pages are the same pixels in any order
//   required-fonts   getRequiredFonts(package | editor), viewer.getRequiredFonts(), the viewer module's function: the faces text is measured and painted in
//   progressive-layout  a long document through a viewport: the first screens at once, the rest in slices (isLayoutComplete / whenLayoutComplete / onChange), the whole as laid out at once; the viewer's { progressive: true }
//   header-footer-edit  a header / footer edited in place: double click, typing, undo across header and body, a new footer with fields, a table, link to previous, save
//   header-footer-variants  undo / redo of link to previous stays in the header; PAGE / NUMPAGES / SECTIONPAGES from the pages shown (the page of the endnotes too), while editing, saved, and once a progressive layout is complete; header text at the size the size box says; a blank title / even page header until something is typed; first and even footers; whole-canvas mode
//   merged-cell-split  a vertically merged cell whose rows cross a page break is split with them (lines, pixels, a real click and typing on the part on the next page); a table row kept with the next
//   table-columns    a table in a two-column section goes from column to column and page to page row by row, the header row repeated (lines, pixels, a real click and typing in the second column)
//   anchored-objects  anchored pictures placed by their anchor (page / margin / column / paragraph, offsets and alignments) and text around them: square, tight along a wrap polygon (lines, pixels, a real click inside the picture's box), top and bottom, none; the polygon saved
//   section-pages    a section's text at its own left margin (pixels, caret, hit test); layout queries find a page by the pages' frames
//   table-borders    table borders in pixels: double and thin-thick lines, corners and crossings, the wider of two conflicting borders
//   table-break-borders  where a page break parts two rows each page has the border between them: the last row's bottom, the first row's top, a repeated header row's (pixels)
//   body-page-field  a PAGE field in the body's text shows the page number in its section's number format (roman, letters), the field's own format first
//   nested-table-edit  a table in a table cell: a real click, typing, Enter / Backspace, Tab between its cells, undo, a save; Enter / Backspace in a cell of the outer table
//   edited-paragraph-xml  an edited paragraph is saved as its own XML with the change of text made in it: bookmark, complex field, hyperlink anchor, content control properties, rsids, w14:paraId kept
//   list-format-numbering  a list's number format changed: numbering.xml keeps the other lists and definitions as they were read; a definition shared with another list is copied for the list
//   media-pruning    picture files and relationships nothing refers to are not written: a picture's paragraph written from the model over and over, a picture deleted
//   progressive-read  a DOCX opened front first (openProgressive): the package read to its end is the package forEditorPackage makes, for every test document; an editor and a viewer show the first screens before the document is read, typing and scrolling meanwhile, what needs the whole document reads the rest
//   progressive-failure  a DOCX opened front first whose body is not well formed at paragraph 7000 of 9000: the load ends failed (package, editor, viewer), what was read is laid out, getText / find / page count answer for it, saving refuses
//   progressive-cancel  the reading of a DOCX opened front first is cancelled: Progressive.cancel, cancelLoad, an editor / viewer done with the package (another document, dispose), a package shown twice; no slice afterwards
//   fonts-pending    setFontAvailability / getPendingFonts / FontsPending; onFontsDiscovered is called before the blocks are given to an editor, also by getPageCount at open: a host with the bytes at hand has the final page count the first time
//   formatted-paragraph-xml  a paragraph made bold / italic / struck through / underlined as a whole, or aligned, is saved as its own XML with the properties set in it: bookmark, complex field, hyperlink anchor, content control properties, rsids, w14:paraId kept
//   drawing-order    the layers of a page in pixels (a header's drawings and watermark below the body, drawings behind the text, the text, drawings in front by relativeHeight); text boxes placed by their anchor, their text in their insets and clipped, text around them; a text box saved as the element it was read from, a changed one from the model in choice and fallback
//   mass-format      everything selected and made bold through a viewport: the change returns with what is in view laid out, the rest follows in slices (isLayoutComplete, onChange, whenLayoutComplete), the layout at the end is that of the document laid out at once; undo; typing meanwhile; half way down the document the blocks above the view stay as they were laid out (getRenderStats().StaleBlocks) until the layout from the start has come as far, undo and redo meanwhile, scrolling up into them, and the view stays on its block (the ScrollTop notice)
//   underline-style  an underline read from a file and one made in the editor are one format in the model, both saved as w:u single; a double underline stays double
//   reference-marks  the mark of a note / a comment has the formatting of its run when read; made bold and written from the model it is bold, the editor's display colour of a comment's mark is not in the file
//   hyperlink-relationships  a paragraph with a hyperlink saved after a change: the link keeps the relationship id the file had, no relationship is left without a referrer, a link taken away takes its relationship; no empty w:pPr
//   patched-paragraph-xml  a part of a paragraph made bold / given a font, size, colour; its spacing, alignment, list: saved as its own XML with runs split and properties set, w14 content, proofing marks, smart tag, field code, comment range, hyperlink kept; in a table in the table's XML, a paragraph that cannot be patched alone from the model
//   split-join-xml   a paragraph split in two (Enter) is saved as two paragraphs made of the XML that was read, the run split, the second half without the paragraph's id; paragraphs joined (Backspace) are the first with the second's content; in a table cell too; a split joined again is the paragraph that was read
//   text-box-navigation  moving through the body does not go into a text box: Ctrl+End is the end of the body's last paragraph (with a caret, and typing goes there), the arrows and Ctrl+A stay in the body; setSelection on what findText found in a box enters the box
//   list-level-indent  a numbered paragraph has the indents of its numbering's level over those of its paragraph style (file-sample.docx: bullets at 0.75 in), its own over both; saved without an indent of its own
//   pending-format   a character format chosen with nothing selected (toggleBold, Ctrl+B, a font, a size, a colour) is, with the caret inside a word, that word's (the caret stays); elsewhere the format of what is typed next at the caret: getSelectionFormat says it, onChange tells of it (FormatChanged), the document is unchanged until text is typed, and a caret that moves drops it
//   api-positions    setSelection with an offset beyond the paragraph's end is the end (and the next keystroke works), below 0 the start, an unknown block is refused; the range queries take offsets the same way; insertPageBreak in a text box, a header or a footer and insertTextBox in a text box are refused
//   rotated-text-box the frame, the selection and the caret of a turned text box that is edited are turned as its text is, and a point is looked up in its text as it is before the turn (13.96 degrees and a quarter turn)
//   table-insert-xml a row or a column inserted into a table is saved like its neighbour (w:tcPr, w:trPr, the paragraph's properties), and w:tblGrid has the columns the table has after a column is inserted or deleted
//   wrap-band        no text in a band beside a floating picture that is narrower than an inch (float-image-wrap-square: the 45.6 px left of the picture), and text on the side(s) wrapText says
//   api-units        what the API takes in points, pixels and multiples is what the file states: setSpacingBefore / After (10 pt was saved as 13.35 pt, and a file's 12 pt read back as 9), font sizes, line spacing, the page, a table's cell margins - set, saved, read again
//   link-underline   a link is underlined unless its text says it is not: the underline of a link is switched off (and on), saved as <w:u w:val="none"/> and read again
//   block-ids        no document under tests/ has a block id twice (getDuplicateBlockIds), nor after edits that make blocks in and around a table in a table
//   text-box-exact-line  a DrawingML text box that states no line and no fill has none, and a line of an exact height lower than its text keeps the text inside it (resume-template's header)
//   section-starts   a section that starts on an odd / even page has a blank page before it where the next page is of the other kind; w:vAlign (center, bottom, both); a page break in columns starts a new page, a column break the next column (and is saved as one); setSectionBreak says how the section after the break starts, saved in that section's properties
//   list-levels      a list item moved to another level (Tab / Shift+Tab, increaseIndent / decreaseIndent / setListLevel) has that level's number ("I.a)", a legal "1.1.3"), indents and bullet at once - what the file read with it at that level gives; saved with its w:ilvl, no indent of its own and the numbering part as read
//   paste-and-boxes  pasted plain text (pasteText, Ctrl+V) has the format chosen at the caret on every line; typing and pasting over a selection of the body keep the text box of the paragraph the selection starts in and never change a box's text, replaceAllText replaces in boxes as findText finds there; a text box inserted into a header is saved in the header's part
//   bookmark-integrity  bold chosen at a caret and 1 / 2 / 5 characters typed into a title with three bookmarks: one bold run, the paragraph changed in its own XML with its bookmarks, outline level and rsids; written from the model every bookmark once and the ranges balanced; one of several bookmarks of a run removed or renamed
//   header-text-box  the text box of a header (resume-template.docx) is one of getShapeRects (Story, PageIndex), is entered with enterTextBox, typed into and saved in the header's part
//   range-marks      the marks of ranges the model does not hold - a bookmark between two blocks, in a paragraph without runs, ending before the first run of the next paragraph, the range of a permission and of a comment - are written where the file had them when their paragraph is written from the model (a table's too); a block that is deleted takes the end of the bookmark it opened; a header written from the model opens and closes each range once
//   field-results    text typed in another format into the result of a field is of the field, every character of it, and is saved in the paragraph's own XML - the result run split, the field's marks once - and read again as it was; a result of which parts are formatted differently is read as its parts; an entry of a table of contents (a field over several paragraphs) is changed in its own XML, and one written from the model leaves the other entries the text they were
//   comment-mark     the mark of a comment ("[Comment 5]") is one thing: setSelection into it goes to its nearer end, the arrow keys step over it, Shift+arrow takes all of it, text is typed before or after it, Backspace after it removes it whole (one undo step), and the file has one w:commentReference
//   native-fixture   the fixture the native painter is tested with (tests/word/native-painter.fsx) is laid out as tests/word/native-painter/expected.json says: the page, the lines of text, the shape's rectangle
//   shape-edit       moveShape / resizeShape / deleteShape (each one undo step), getShapeHandles / hitTestShapeHandle; a shape that was read and is only moved or resized is saved as its own element with the anchor's offsets and extent changed; an aligned shape that is moved is by its offset from the page; a shape aligned in one direction and offset in the other is read with its offset
//   measure-mid-layout  the measurers change (a face arrives: setTextMeasurersForFaces, at once or after some slices; setTextMeasurers), the host zooms or gets another width while most of a long document is laid out again with the view far down it: the blocks above the view stay as they were laid out, the document never ends above the view, the view stays on its block (the ScrollTop notice), the layout at the end is that of the new measurers
//   keys             Ctrl+Enter is a page break, Ctrl+Shift+Enter a column break (saved as one), Shift+Enter a line break, Ctrl+Backspace / Ctrl+Delete a word, Tab a tab character outside lists and tables (Ctrl+Tab in a table); one undo step each; insertColumnBreak / insertLineBreak; nothing in a text box
//   header-box-pointer  the text box of a header with real clicks: a double click on it from the body edits its text, a click on it while the header is edited does, a click elsewhere on the header goes back to the header, a double click on the body to the body; moveShape / resizeShape / deleteShape / getShapeHandles for a header's shape, a moved one saved in its part
//   shape-select     the selected shape of the editor itself: selectShape / getSelectedShape / the ShapeSelectionChanged notice, frame and handles painted (turned with a turned shape), a click into a box selects it, Escape, the arrow keys, Delete, a drag of the frame and of a handle with the real pointer - one undo step each, also by the north-west handle; resizeShapeBy
//   shape-pointer    a click on the frame of a shape that is not selected (behind the text, without text) selects it and a drag from there moves it; the listeners hear of the undo step when a drag ends; a turned shape's handles (getShapeHandles, hitTestShapeHandle) where they are painted, getShapeRects' Rotation
//   editing          Enter in the paragraph that ends a section keeps the sections (pages, saved w:sectPr); keys typed one after the other are one undo step, so are Backspaces / Deletes; a caret move, Enter, a format, an undo start a new one; copy / cut / paste through the browser's events on the input proxy: text and HTML written, a copy pasted with its formatting, other text as text, HTML of another program with its paragraphs, list and bold; copySelection / cutSelection / pasteClipboard; Ctrl+E / L / R / J, Ctrl+M / Ctrl+Shift+M, Ctrl+Shift+> / <, Ctrl+] / [, Ctrl+= / Ctrl+Shift+=, Ctrl+Space (on a selection and at a caret); Shift+F3 (Each Word / UPPER / lower, on a selection or the word at the caret), Ctrl+Shift+L (bullets), Ctrl+Shift+N (Normal)
//   clipboard-keys   the real Ctrl+C / Ctrl+V (Playwright): what is pasted has the formatting it was copied with
//   default-tab-stop  w:defaultTabStop of the settings part is the distance of the default tab stops (708, 1440 twips; 720 when none is stated), and the saved settings part keeps it
//   xml-scan         the DOCX reader's XML scanner gives the nodes the browser's parser gives, for every part of the test documents
//
// Environment (as tools/docx-roundtrip/roundtrip.mjs): PLAYWRIGHT, JSZIP_DIR, PORT (default 16022).
// Exit code is non-zero when any check fails.
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import { syntheticFixtures } from '../docx-roundtrip/fixtures.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.resolve(here, '../..');
const args = process.argv.slice(2);
const opt = (name) => { const i = args.indexOf(name); if (i < 0) return undefined; const v = args[i + 1]; args.splice(i, 2); return v; };
const dist = path.resolve(opt('--dist') ?? path.join(repo, 'dist'));
const port = Number(process.env.PORT ?? 16022);
const all = ['explicit-format', 'page-frames', 'ime', 'typing', 'track-changes', 'linked-images', 'opaque-click', 'viewport', 'selection-format',
    'placeholder-and-frames', 'caret-movement', 'revision', 'viewer-measurers', 'page-centering', 'commands', 'image-handles', 'painter-frames', 'viewer-entry', 'keystroke-long', 'table-tools', 'caret-frames', 'viewer-first-page', 'required-fonts', 'progressive-layout', 'xml-scan', 'header-footer-edit', 'header-footer-variants', 'merged-cell-split', 'table-columns', 'anchored-objects', 'section-pages', 'table-borders', 'table-break-borders', 'body-page-field', 'nested-table-edit', 'edited-paragraph-xml', 'list-format-numbering', 'media-pruning', 'progressive-read', 'progressive-failure', 'progressive-cancel', 'fonts-pending', 'formatted-paragraph-xml', 'drawing-order', 'mass-format', 'underline-style', 'reference-marks', 'hyperlink-relationships', 'patched-paragraph-xml', 'split-join-xml', 'text-box-navigation', 'list-level-indent', 'pending-format', 'api-positions', 'rotated-text-box', 'table-insert-xml', 'wrap-band', 'api-units', 'link-underline', 'block-ids', 'text-box-exact-line', 'section-starts', 'list-levels', 'paste-and-boxes', 'bookmark-integrity', 'header-text-box', 'range-marks', 'field-results', 'comment-mark', 'native-fixture', 'shape-edit', 'measure-mid-layout', 'keys', 'header-box-pointer', 'shape-select', 'shape-pointer', 'editing', 'clipboard-keys', 'default-tab-stop'];
const checks = args.length ? args : all;

const jszipDir = [process.env.JSZIP_DIR, path.join(repo, 'node_modules/jszip'), path.resolve(repo, '../rbt-office/node_modules/jszip')].find((p) => p && fs.existsSync(p));
if (!jszipDir) { console.error('jszip not found; set JSZIP_DIR'); process.exit(2); }
const JSZip = createRequire(import.meta.url)(path.join(jszipDir, 'dist/jszip.min.js'));
const fixtures = new Map((await syntheticFixtures(JSZip)).map((f) => [f.name, f.data]));

const page = (check) => `<!doctype html><meta charset="utf-8"><body style="margin:0">
<canvas id="c" width="1400" height="4000" style="width:1400px;height:4000px"></canvas>
<script src="/jszip.min.js"></script><script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script>
<script type="module" src="/checks/${check}.js"></script>`;
const mime = { '.js': 'text/javascript', '.mjs': 'text/javascript' };
let externalProbes = 0;
const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://x');
    const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
    if (url.pathname === '/') return send(200, 'text/html', page(url.searchParams.get('check')));
    if (url.pathname === '/jszip.min.js') return send(200, 'text/javascript', fs.readFileSync(path.join(jszipDir, 'dist/jszip.min.js')));
    if (url.pathname === '/jszip-shim.js') return send(200, 'text/javascript', 'export default window.JSZip;');
    // Stands in for an address a document names as the source of a linked picture.
    if (url.pathname === '/external-probe.png') { externalProbes++; return send(200, 'image/png', Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64')); }
    if (url.pathname === '/external-probe-count') return send(200, 'text/plain', String(externalProbes));
    // The DOCX files under tests/ (paths relative to the repository).
    if (url.pathname === '/corpus-list') {
        const found = [];
        const walk = (dir) => { for (const entry of fs.readdirSync(dir, { withFileTypes: true })) { const full = path.join(dir, entry.name); if (entry.isDirectory()) walk(full); else if (entry.name.endsWith('.docx')) found.push(path.relative(repo, full)); } };
        for (const entry of fs.readdirSync(path.join(repo, 'tests'), { withFileTypes: true })) if (entry.isDirectory() && entry.name.startsWith('test-documents')) walk(path.join(repo, 'tests', entry.name));
        return send(200, 'application/json', JSON.stringify(found.sort()));
    }
    if (url.pathname === '/fixture') { const data = fixtures.get(url.searchParams.get('name')); return data ? send(200, 'application/octet-stream', data) : send(404, 'text/plain', 'no fixture'); }
    for (const [prefix, root] of [['/checks/', path.join(here, 'checks')], ['/repo/', repo], ['/', dist]]) {
        if (!url.pathname.startsWith(prefix)) continue;
        const file = path.join(root, path.normalize(url.pathname.slice(prefix.length)));
        if (file.startsWith(root) && fs.existsSync(file) && fs.statSync(file).isFile())
            return send(200, mime[path.extname(file)] ?? 'application/octet-stream', fs.readFileSync(file));
    }
    send(404, 'text/plain', 'not found');
});

for (const required of ['office-formats/src/Main.js', 'word/Word.js'])
    if (!fs.existsSync(path.join(dist, required))) { console.error(`missing ${required} under ${dist}; Fable-compile first (see header)`); process.exit(2); }
await new Promise((resolve) => server.listen(port, '127.0.0.1', resolve));
let playwright;
try { playwright = await import(process.env.PLAYWRIGHT ? pathToFileURL(process.env.PLAYWRIGHT).href : 'playwright'); }
catch { console.error('playwright not found; set PLAYWRIGHT=/path/to/playwright/index.mjs'); server.close(); process.exit(2); }
const browser = await (playwright.chromium ?? playwright.default.chromium).launch();
let failedChecks = 0;
try {
    const runCheck = async (check) => {
        externalProbes = 0;
        const tab = await browser.newPage({ viewport: { width: 1100, height: 900 } });
        const errors = [];
        tab.on('pageerror', (e) => errors.push('page error: ' + e.message));
        await tab.goto(`http://127.0.0.1:${port}/?check=${encodeURIComponent(check)}`);
        await tab.waitForFunction('window.result !== undefined || window.drive !== undefined', null, { timeout: 240000 });
        if (await tab.evaluate('window.drive !== undefined')) {
            // The check wants real input: [["click" | "dblclick" | "tripleclick" | "shiftclick", x, y] | ["type", text] | ["press", key] | ["down", x, y] | ["move", x, y] | ["up"] | ["hover", x, y] | ["wheel", dx, dy] | ["eval", js] | ["then", js returning more actions]] ...
            const perform = async (actions) => {
                for (const [action, a, b] of actions) {
                    if (action === 'click') await tab.mouse.click(a, b);
                    else if (action === 'dblclick') await tab.mouse.dblclick(a, b);
                    else if (action === 'tripleclick') await tab.mouse.click(a, b, { clickCount: 3 });
                    else if (action === 'shiftclick') { await tab.keyboard.down('Shift'); await tab.mouse.click(a, b); await tab.keyboard.up('Shift'); }
                    else if (action === 'type') await tab.keyboard.type(a);
                    else if (action === 'press') await tab.keyboard.press(a);
                    else if (action === 'down') { await tab.mouse.move(a, b); await tab.mouse.down(); }
                    else if (action === 'move') await tab.mouse.move(a, b, { steps: 4 });
                    else if (action === 'hover') await tab.mouse.move(a, b);
                    else if (action === 'up') await tab.mouse.up();
                    else if (action === 'wheel') await tab.mouse.wheel(a, b);
                    else if (action === 'eval') await tab.evaluate(a);
                    // The next actions depend on what the page shows by now: the expression returns them.
                    else if (action === 'then') await perform(await tab.evaluate(a));
                }
            };
            await perform(await tab.evaluate('window.drive'));
            await tab.evaluate('window.finish()');
            await tab.waitForFunction('window.result !== undefined', null, { timeout: 240000 });
        }
        const result = await tab.evaluate('window.result');
        await tab.close();
        return { result, errors };
    };
    for (const check of checks) {
        let { result, errors } = await runCheck(check);
        // A module request that the page could not fetch (the browser was busy starting up): once more.
        if (result.log.some((line) => line.includes('Failed to fetch dynamically imported module'))) ({ result, errors } = await runCheck(check));
        const problems = [...errors, ...result.log.filter((line) => process.env.VERBOSE || !line.startsWith("ok"))];
        console.log(`${result.failed === 0 && errors.length === 0 ? 'ok  ' : 'FAIL'} ${check}  (${result.passed} passed, ${result.failed} failed)`);
        for (const line of problems) console.log('     ' + line);
        if (result.failed || errors.length) failedChecks++;
    }
} finally {
    await browser.close();
    server.close();
}
console.log(failedChecks === 0 ? '\nWORD BROWSER CHECKS: ALL PASSED' : `\nWORD BROWSER CHECKS: ${failedChecks} check(s) FAILED`);
process.exit(failedChecks === 0 ? 0 : 1);
