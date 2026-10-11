#!/usr/bin/env node
// Timings of the Word editor on a long document in headless Chromium, against
// Fable output: the cost of a keystroke (typing in a plain paragraph in the
// middle of the document), of a caret move (through the API and by the arrow
// keys), of Enter, of a Backspace that merges two paragraphs, of a short
// string inserted and of a scroll frame.
//
// The document is tests/test-documents-wild/docx4j-getting-started.docx with
// its body repeated (--repeat N, default 8: some 500 pages).
//
// Usage:
//   node tools/word-browser/perf.mjs --dist $DIST [--progressive] [--repeat N] [--keys N] [--profile keys] [--callers FUNCTION] [--painter] [--measure-cost MICROSECONDS]
// --profile select: the time under each function while everything is selected and the host's queries are answered;
// --profile keys prints the functions with the most self time while typing, and the time under each
// function (--profile enter: the time under each function while Enter and Backspace alternate;
// --profile load: where opening the document spends its time); --callers FUNCTION adds the time under that function by its chain of callers.
// --painter installs a (trivial) text path painter and measurers first.
// --measure-cost N (implies --painter) makes the measurers cost what a host's shaped text stack costs: a key is
// made of the style and the text for every call, widths are kept in a map of 5000 entries that is emptied when
// full (as the host's caches are), and a run that is not in it takes N microseconds to shape. With N = 300 the
// 600-page document is laid out in the time the host's real text stack takes (some 2.3 s).
// Environment: PLAYWRIGHT, JSZIP_DIR, PORT (default 16125).
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.resolve(here, '../..');
const args = process.argv.slice(2);
const opt = (name) => { const i = args.indexOf(name); if (i < 0) return undefined; const v = args[i + 1]; args.splice(i, 2); return v; };
const flag = (name) => { const i = args.indexOf(name); if (i < 0) return false; args.splice(i, 1); return true; };
const dist = path.resolve(opt('--dist') ?? path.join(repo, 'dist'));
const repeat = Number(opt('--repeat') ?? 8);
const keys = Number(opt('--keys') ?? 40);
const profile = opt('--profile');
const measureCost = Number(opt('--measure-cost') ?? 0);
const painter = flag('--painter') || measureCost > 0;
// --one-shot: the rest of the document is laid out at once after the first paint (completeLayout), not in idle slices.
const oneShot = flag('--one-shot');
// --progressive: the document is opened front first (officeFormats.word.openProgressive): shown when its first blocks are read.
const progressive = flag('--progressive');
const pathsOf = opt('--callers');
// --under FUNCTION (with --profile format): the self times listed are those of what runs under that function.
const under = opt('--under');
const port = Number(process.env.PORT ?? 16125);
const jszipDir = [process.env.JSZIP_DIR, path.join(repo, 'node_modules/jszip'), path.resolve(repo, '../rbt-office/node_modules/jszip')].find((p) => p && fs.existsSync(p));
if (!jszipDir) { console.error('jszip not found; set JSZIP_DIR'); process.exit(2); }
const JSZip = createRequire(import.meta.url)(path.join(jszipDir, 'dist/jszip.min.js'));

// The long document: the body's blocks repeated (the final section properties stay last).
export const longDocument = async (times) => {
    const zip = await JSZip.loadAsync(fs.readFileSync(path.join(repo, 'tests/test-documents-wild/docx4j-getting-started.docx')));
    const xml = await zip.file('word/document.xml').async('string');
    const open = xml.indexOf('>', xml.indexOf('<w:body')) + 1, close = xml.lastIndexOf('</w:body>');
    const sect = xml.lastIndexOf('<w:sectPr', close);
    const blocks = xml.slice(open, sect > open ? sect : close), tail = xml.slice(sect > open ? sect : close, close);
    zip.file('word/document.xml', xml.slice(0, open) + blocks.repeat(times) + tail + xml.slice(close));
    return zip.generateAsync({ type: 'nodebuffer' });
};
const docx = await longDocument(repeat);

const page = `<!doctype html><meta charset="utf-8"><body style="margin:0"><canvas id="c"></canvas>
<script src="/jszip.min.js"></script><script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script><script type="module">
try {
    const main = await import('/office-formats/src/Main.js');
    const OF = typeof main.default === 'function' ? main.default() : main.default;
    const Word = (await import('/word/Word.js')).default;
    const editor = Word.initializeCanvas(document.getElementById('c'));
    const stats = (times) => { const sorted = times.slice().sort((a, b) => a - b); return { mean: times.reduce((a, b) => a + b, 0) / times.length, median: sorted[Math.floor(sorted.length / 2)], max: sorted[sorted.length - 1] }; };
    window.perf = {
        async load(usePainter, measureCost, oneShot, progressive) {
            const bytes = new Uint8Array(await (await fetch('/long.docx')).arrayBuffer());
            const t0 = performance.now();
            // (Front first: the package is handed over when the first blocks are read; the rest is read in slices.)
            const editorPackage = progressive ? await OF.word.openProgressive(bytes) : await OF.processors.forEditorPackage(bytes);
            const read = performance.now() - t0;
            const firstItems = editorPackage.Paragraphs.length;
            // A host asks which font faces the package needs before it loads it.
            let t2 = performance.now();
            const faces = Word.getRequiredFonts ? Word.getRequiredFonts(editorPackage).length : 0;
            const fonts = performance.now() - t2;
            // As a host does: the measurers and the painter first, then the document (laid out once).
            let t1 = performance.now();
            if (usePainter && measureCost > 0) {
                // A host's shaped text stack, in what it costs (see the header).
                const styleKey = (style) => [style?.FontFamily ?? '', style?.FontSize ?? '', style?.Bold ? 'b' : '', style?.Italic ? 'i' : '', style?.SmallCaps ? 'sc' : '', style?.AllCaps ? 'ac' : ''].join('|');
                const texts = new Map(), ranges = new Map(), shapes = new Map();
                const keep = (cache, key, value) => { if (cache.size >= 5000) cache.clear(); cache.set(key, value); return value; };
                const shape = (text, style) => {
                    const key = styleKey(style) + '|' + text;
                    if (shapes.has(key)) return shapes.get(key);
                    window.perfMeasure.shaped += 1;
                    const until = performance.now() + measureCost / 1000;
                    while (performance.now() < until) { /* shaping */ }
                    return keep(shapes, key, text.length * style.FontSize * 0.5);
                };
                window.perfMeasure = { shaped: 0, calls: 0 };
                Word.setTextMeasurers(editor, (text, style) => {
                    window.perfMeasure.calls += 1;
                    const key = styleKey(style) + '|' + text;
                    if (texts.has(key)) return texts.get(key);
                    return keep(texts, key, shape(text, style));
                }, (text, style, start, length) => {
                    window.perfMeasure.calls += 1;
                    const key = styleKey(style) + '|' + start + '|' + length + '|' + text;
                    if (ranges.has(key)) return ranges.get(key);
                    shape(text, style);
                    return keep(ranges, key, Math.max(0, Math.min(length, text.length - start)) * style.FontSize * 0.5);
                });
                Word.setTextPathPainter(editor, (text, style, x, baseline, width) => false);
            } else if (usePainter) {
                Word.setTextMeasurers(editor, (text, style) => text.length * style.FontSize * 0.5, (text, style, start, length) => Math.max(0, Math.min(length, text.length - start)) * style.FontSize * 0.5);
                Word.setTextPathPainter(editor, (text, style, x, baseline, width) => false);
            }
            const installed = performance.now() - t1;
            // (A host shows the document through a viewport: what is painted first is one screen.)
            Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 1000, height: 800 });
            t1 = performance.now();
            Word.importFromDocxPackage(editor, editorPackage);
            const imported = performance.now() - t1;
            // What is on screen now is the first paint: the first screens are laid out, the rest follows in idle slices.
            const firstPaint = performance.now() - t0;
            const firstSize = Word.getDocumentSize(editor);
            const firstBlocks = Word.getRenderStats(editor).LaidOutLimit;
            const pixels = document.getElementById('c').getContext('2d').getImageData(0, 0, 1000, 800).data;
            let ink = 0;
            for (let i = 0; i < pixels.length; i += 4) if (pixels[i] < 100 && pixels[i + 1] < 100 && pixels[i + 2] < 100) ink += 1;
            let notices = 0, longest = 0, last = performance.now();
            const off = Word.onChange ? Word.onChange(editor, (change) => { if (change.LayoutChanged) notices += 1; }) : () => {};
            // (How long the page could not answer while the rest was laid out: the longest gap between two timer ticks.)
            let longestAt = null;
            const ticker = setInterval(() => { const now = performance.now(); if (now - last > longest) { longest = now - last; const stats = Word.getRenderStats(editor); longestAt = { reading: Word.isLoadComplete ? !Word.isLoadComplete(editor) : false, layoutSlices: stats.LayoutSlices ?? 0, readSlices: editorPackage.Progressive?.Slices ?? 0, laidOut: stats.LaidOutLimit }; } last = now; }, 4);
            t1 = performance.now();
            // (A document that is still being read: read to its end, in slices, then laid out.)
            let readAll = 0, longestRead = 0;
            if (progressive && Word.whenLoadComplete) { await Word.whenLoadComplete(editor); readAll = performance.now() - t0; longestRead = longest; }
            if (oneShot && Word.completeLayout) Word.completeLayout(editor);
            if (Word.whenLayoutComplete) await Word.whenLayoutComplete(editor);
            const rest = performance.now() - t1;
            const sliceStats = { ...Word.getRenderStats(editor) };
            clearInterval(ticker); off();
            const loaded = performance.now() - t0;
            const size = Word.getDocumentSize(editor);
            // The caret in a plain paragraph near the middle of the document.
            t1 = performance.now();
            const frames = Word.getPageFrames(editor);
            const framesTime = performance.now() - t1;
            const target = frames[Math.floor(frames.length / 2)];
            Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: target.Top, scrollLeft: 0, width: 1000, height: 800 });
            return { loaded, read, imported, installed, fonts, faces, frames: framesTime, progressive: !!progressive, firstItems, readAll, longestRead, atOnce: !!(editorPackage.Progressive && editorPackage.Progressive.AtOnce), slices: editorPackage.Progressive?.Slices ?? 0, longestSlice: editorPackage.Progressive?.LongestSliceMs ?? 0, longestSliceAt: editorPackage.Progressive?.LongestSlice ?? 0, longestSliceNotify: editorPackage.Progressive?.LongestSliceNotifyMs ?? 0, longestSliceItems: editorPackage.Progressive?.LongestSliceItems ?? 0, longestSlicePass: editorPackage.Progressive?.LongestSlicePassMs ?? 0, pages: size.PageCount, top: target.Top, measured: window.perfMeasure ?? null,
                firstPaint, firstPages: firstSize.PageCount, firstBlocks, firstComplete: firstSize.LayoutComplete, ink, rest, notices, longest, passes: Word.getRenderStats(editor).LayoutPasses, sliceStats, longestAt };
        },
        // Put the caret into a paragraph of plain text on the page in view (by clicking at a point of the canvas).
        caretAt(x, y) { return Word.getSelectionInformation ? JSON.stringify(Word.getCaretRect(editor)) : null; },
        keys(count) {
            const times = [];
            for (let i = 0; i < count; i++) { const t0 = performance.now(); Word.insertText(editor, 'x'); times.push(performance.now() - t0); }
            return stats(times);
        },
        moves(count) {
            const times = [];
            for (let i = 0; i < count; i++) { const t0 = performance.now(); Word.moveLeft ? Word.moveLeft(editor) : null; times.push(performance.now() - t0); }
            return stats(times);
        },
        scroll(count, top) {
            const times = [];
            for (let i = 0; i < count; i++) { const t0 = performance.now(); Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: top + i * 40, scrollLeft: 0, width: 1000, height: 800 }); times.push(performance.now() - t0); }
            return stats(times);
        },
        // Caret moves through the API: the caret is put at other offsets of its paragraph.
        carets(count) {
            const sel = Word.getSelection(editor), times = [];
            const builds = Word.getRenderStats(editor).FrameBuilds;
            for (let i = 0; i < count; i++) { const at = 1 + (i % 5); const t0 = performance.now(); Word.setSelection(editor, sel.FocusBlockId, at, sel.FocusBlockId, at); times.push(performance.now() - t0); }
            Word.setSelection(editor, sel.AnchorBlockId, sel.AnchorOffset, sel.FocusBlockId, sel.FocusOffset);
            return { ...stats(times), builds: Word.getRenderStats(editor).FrameBuilds - builds };
        },
        // Keys through the editor's input element, as the keyboard sends them (each a list: pressed in turn, timed apart).
        pressed(names, count) {
            const input = Word.getInputElement(editor), times = names.map(() => []);
            const builds = Word.getRenderStats(editor).FrameBuilds;
            for (let i = 0; i < count; i++) names.forEach((name, k) => {
                const shift = name.startsWith('Shift+'), key = shift ? name.slice(6) : name;
                const t0 = performance.now();
                input.dispatchEvent(new KeyboardEvent('keydown', { key, shiftKey: shift, bubbles: true, cancelable: true }));
                times[k].push(performance.now() - t0);
            });
            const after = Word.getRenderStats(editor);
            return { times: times.map(stats), builds: after.FrameBuilds - builds, laidOut: after.LaidOutBlocks, collected: after.CollectedBlocks, blocks: after.Blocks };
        },
        pastes(count, text) {
            const times = [];
            for (let i = 0; i < count; i++) { const t0 = performance.now(); Word.insertText(editor, text); times.push(performance.now() - t0); }
            return stats(times);
        },
        // The viewer (a preview): load, page frames, the first page painted, the pages after it.
        async viewer(pages) {
            const createViewer = (await import('/word/WordViewer.js')).default;
            const bytes = new Uint8Array(await (await fetch('/long.docx')).arrayBuffer());
            let t0 = performance.now();
            const editorPackage = await OF.processors.forEditorPackage(bytes);
            const read = performance.now() - t0;
            t0 = performance.now();
            const viewer = createViewer(document.createElement('canvas'));
            viewer.load(editorPackage);
            const frames = viewer.getPageFrames();
            const laidOut = performance.now() - t0;
            const target = document.createElement('canvas');
            t0 = performance.now();
            viewer.renderPage(target, 0, 1);
            const first = performance.now() - t0;
            const firstStats = viewer.getRenderStats ? viewer.getRenderStats() : null;
            const times = [];
            for (let i = 1; i <= pages && i < frames.length; i++) { const t = performance.now(); viewer.renderPage(target, i, 1); times.push(performance.now() - t); }
            return { read, laidOut, first, later: stats(times), pages: frames.length, firstStats, lastStats: viewer.getRenderStats ? viewer.getRenderStats() : null };
        },
        // Everything selected: the frame after it, scroll frames with it, then a caret placed again.
        selectAll(top) {
            const before = Word.getSelection(editor);
            let t0 = performance.now(); Word.selectAll(editor); const select = performance.now() - t0;
            // What a host reads after the change: the selection's format and the caret's context (asked twice, as a host with two listeners does).
            const asked = [];
            for (let i = 0; i < 3; i++) { const t = performance.now(); Word.getSelectionFormat(editor); const m = performance.now(); Word.getCaretContext(editor); asked.push(m - t, performance.now() - m); }
            // (And once more, the caret placed in between: the first time includes what the engine's code costs when it first runs.)
            Word.setSelection(editor, before.FocusBlockId, before.FocusOffset, before.FocusBlockId, before.FocusOffset); Word.selectAll(editor);
            { const t = performance.now(); Word.getSelectionFormat(editor); asked.push(performance.now() - t); }
            const times = [];
            for (let i = 0; i < 30; i++) { const t = performance.now(); Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: top + i * 400, scrollLeft: 0, width: 1000, height: 800 }); times.push(performance.now() - t); }
            t0 = performance.now(); Word.setSelection(editor, before.FocusBlockId, before.FocusOffset, before.FocusBlockId, before.FocusOffset); const collapse = performance.now() - t0;
            Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: top, scrollLeft: 0, width: 1000, height: 800 });
            return { select, asked, scroll: stats(times), collapse };
        },
        // The whole document made bold, the step undone, redone and undone again: one step each
        // (step: 'select' | 'bold' | 'undo' | 'redo' | 'collapse'), so that a profile is of one of them.
        formatAll(step) {
            const passesBefore = Word.getRenderStats(editor).LayoutPasses;
            const t = performance.now();
            if (step === 'select') { window.perfCaret = Word.getSelection(editor); Word.selectAll(editor); }
            else if (step === 'bold') Word.toggleBold(editor);
            else if (step === 'undo') Word.undo(editor);
            else if (step === 'redo') Word.redo(editor);
            else if (step === 'collapse') { const c = window.perfCaret; Word.setSelection(editor, c.FocusBlockId, c.FocusOffset, c.FocusBlockId, c.FocusOffset); }
            // (With what a host reads after every change.)
            const command = performance.now() - t;
            Word.getSelectionFormat(editor); Word.getDocumentSize(editor);
            const took = performance.now() - t;
            const stats = Word.getRenderStats(editor);
            return { took, command, passes: stats.LayoutPasses - passesBefore, stale: stats.StaleBlocks ?? 0, laidOut: stats.LaidOutBlocks, collected: stats.CollectedBlocks, blocks: stats.Blocks, pages: Word.getDocumentSize(editor).PageCount, complete: Word.isLayoutComplete ? Word.isLayoutComplete(editor) : true };
        },
        // What the host reads after every change, and vertical caret moves.
        queries(count) {
            const context = [], format = [];
            for (let i = 0; i < count; i++) { let t = performance.now(); Word.getCaretContext(editor); context.push(performance.now() - t); t = performance.now(); Word.getSelectionFormat(editor); format.push(performance.now() - t); }
            return { context: stats(context), format: stats(format) };
        },
        text() { return Word.getText(editor).length; },
        renderStats() { return Word.getRenderStats ? Word.getRenderStats(editor) : null; },
    };
} catch (e) { window.loadError = String(e && e.stack || e); }
</script>`;
const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://x');
    const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
    if (url.pathname === '/') return send(200, 'text/html', page);
    if (url.pathname === '/long.docx') return send(200, 'application/octet-stream', docx);
    if (url.pathname === '/jszip.min.js') return send(200, 'text/javascript', fs.readFileSync(path.join(jszipDir, 'dist/jszip.min.js')));
    if (url.pathname === '/jszip-shim.js') return send(200, 'text/javascript', 'export default window.JSZip;');
    const file = path.join(dist, path.normalize(url.pathname));
    if (file.startsWith(dist) && fs.existsSync(file) && fs.statSync(file).isFile()) return send(200, 'text/javascript', fs.readFileSync(file));
    send(404, 'text/plain', 'not found');
});
if (import.meta.url === pathToFileURL(process.argv[1]).href) {
    await new Promise((resolve) => server.listen(port, '127.0.0.1', resolve));
    const playwright = await import(process.env.PLAYWRIGHT ? pathToFileURL(process.env.PLAYWRIGHT).href : 'playwright');
    const browser = await (playwright.chromium ?? playwright.default.chromium).launch();
    const ms = (t) => t.toFixed(1) + ' ms';
    const line = (s) => `mean ${ms(s.mean)}, median ${ms(s.median)}, max ${ms(s.max)}`;
    try {
        const tab = await browser.newPage({ viewport: { width: 1100, height: 900 } });
        tab.on('pageerror', (e) => console.error('page error:', e.message));
        await tab.goto(`http://127.0.0.1:${port}/`);
        await tab.waitForFunction('window.perf !== undefined || window.loadError !== undefined', null, { timeout: 240000 });
        const loadError = await tab.evaluate('window.loadError');
        if (loadError) throw new Error('page failed to load modules: ' + loadError);
        const cdpLoad = await tab.context().newCDPSession(tab);
        if (profile === 'load') { await cdpLoad.send('Profiler.enable'); await cdpLoad.send('Profiler.setSamplingInterval', { interval: 500 }); await cdpLoad.send('Profiler.start'); }
        const loaded = await tab.evaluate(([usePainter, cost, atOnce, frontFirst]) => window.perf.load(usePainter, cost, atOnce, frontFirst), [painter, measureCost, oneShot, progressive]);
        if (profile === 'load') {
            // Where opening the document spends its time: self time, and time under each function.
            const { profile: cpu } = await cdpLoad.send('Profiler.stop');
            const interval = cpu.timeDeltas.reduce((a, b) => a + b, 0) / Math.max(1, cpu.samples.length) / 1000;
            const byId = new Map(cpu.nodes.map((node) => [node.id, node]));
            const selfById = new Map();
            for (const id of cpu.samples) selfById.set(id, (selfById.get(id) ?? 0) + interval);
            const keyOf = (frame) => `${frame.functionName || '(anonymous)'}  ${frame.url.split('/').slice(-2).join('/')}:${frame.lineNumber + 1}`;
            const self = new Map();
            for (const [id, time] of selfById) { const key = keyOf(byId.get(id).callFrame); self.set(key, (self.get(key) ?? 0) + time); }
            console.log('  -- opening: self time --');
            for (const [key, time] of [...self.entries()].sort((a, b) => b[1] - a[1]).slice(0, 28)) console.log(`  ${time.toFixed(0).padStart(7)} ms  ${key}`);
            const total = new Map();
            const walk = (node, open) => {
                const key = keyOf(node.callFrame);
                let sum = selfById.get(node.id) ?? 0;
                const inside = open.has(key);
                if (!inside) open.add(key);
                for (const child of node.children ?? []) sum += walk(byId.get(child), open);
                if (!inside) { open.delete(key); total.set(key, (total.get(key) ?? 0) + sum); }
                return sum;
            };
            walk(cpu.nodes[0], new Set());
            console.log('  -- opening: time under a function --');
            for (const [key, time] of [...total.entries()].filter(([key]) => (process.env.PROFILE_FILTER ? new RegExp(process.env.PROFILE_FILTER) : /word\/|office-formats\//).test(key)).sort((a, b) => b[1] - a[1]).slice(0, 70)) console.log(`  ${time.toFixed(0).padStart(7)} ms  ${key}`);
        }
        console.log(`  getRequiredFonts(package): ${loaded.fonts.toFixed(0)} ms (${loaded.faces} faces), before the import`);
        console.log(`  opening: read the package ${(loaded.read / 1000).toFixed(2)} s, import and first paint ${(loaded.imported / 1000).toFixed(2)} s${painter ? `, measurers and painter installed ${(loaded.installed / 1000).toFixed(2)} s` : ''}, page frames ${loaded.frames.toFixed(0)} ms`);
        console.log(`  first paint ${(loaded.firstPaint / 1000).toFixed(2)} s after the bytes (${loaded.firstBlocks} blocks, ${loaded.firstPages} pages laid out, ${loaded.ink} dark pixels on the first screen${loaded.firstComplete ? '; the whole document' : ''}); the rest of the layout ${(loaded.rest / 1000).toFixed(2)} s in ${loaded.passes} passes, ${loaded.notices} notices to onChange, the longest time without an answer ${loaded.longest.toFixed(0)} ms`);
        if (loaded.sliceStats && loaded.sliceStats.LayoutSlices) console.log(`  the layout's slices: ${loaded.sliceStats.LayoutSlices}, the longest (slice ${loaded.sliceStats.LongestLayoutSlice}, from ${loaded.sliceStats.LongestLayoutSliceFrom} blocks) ${loaded.sliceStats.LongestLayoutSliceMs.toFixed(1)} ms: ${loaded.sliceStats.LongestLayoutSlicePasses} passes in ${loaded.sliceStats.LongestLayoutSliceLayoutMs.toFixed(1)} ms, the rest what followed it (repaint, listeners)`);
        if (loaded.longestAt) console.log(`  the longest time without an answer ended ${loaded.longestAt.reading ? 'while the document was read' : 'while it was laid out'}: after ${loaded.longestAt.readSlices} slices of reading and ${loaded.longestAt.layoutSlices} of layout (${loaded.longestAt.laidOut} blocks laid out)`);
        if (loaded.progressive) console.log(`  opened front first: ${loaded.firstItems} blocks in the package when it was handed over${loaded.atOnce ? ' (the main part was read at once)' : ''}; read to its end ${(loaded.readAll / 1000).toFixed(2)} s after the bytes, in ${loaded.slices} slices of which the longest (slice ${loaded.longestSliceAt}) took ${loaded.longestSlice.toFixed(1)} ms - ${loaded.longestSliceItems} items, its longest pass ${loaded.longestSlicePass.toFixed(1)} ms, the listeners ${loaded.longestSliceNotify.toFixed(1)} ms - (the longest time between two timer ticks while it was read: ${loaded.longestRead.toFixed(0)} ms)`);
        if (loaded.measured) console.log(`  measurers: ${loaded.measured.calls} calls, ${loaded.measured.shaped} runs shaped at ${measureCost} microseconds`);
        console.log(`docx4j-getting-started x ${repeat}: ${loaded.pages} pages, loaded in ${(loaded.loaded / 1000).toFixed(2)} s${painter ? ' (text path painter and measurers installed)' : ''}`);
        // A click into the text of the page in view puts the caret there.
        await tab.mouse.click(400, 300);
        await tab.evaluate(() => window.perf.keys(3));
        const cdp = await tab.context().newCDPSession(tab);
        if (profile === 'keys') { await cdp.send('Profiler.enable'); await cdp.send('Profiler.setSamplingInterval', { interval: 200 }); await cdp.send('Profiler.start'); }
        const typed = await tab.evaluate((n) => window.perf.keys(n), keys);
        if (profile === 'keys') {
            const { profile: cpu } = await cdp.send('Profiler.stop');
            const self = new Map();
            const interval = cpu.timeDeltas.reduce((a, b) => a + b, 0) / Math.max(1, cpu.samples.length) / 1000;
            const byId = new Map(cpu.nodes.map((node) => [node.id, node]));
            for (const id of cpu.samples) {
                const frame = byId.get(id).callFrame;
                const key = `${frame.functionName || '(anonymous)'}  ${frame.url.split('/').slice(-2).join('/')}:${frame.lineNumber + 1}`;
                self.set(key, (self.get(key) ?? 0) + interval);
            }
            console.log(`  -- self time per keystroke (of ${keys}) --`);
            for (const [key, time] of [...self.entries()].sort((a, b) => b[1] - a[1]).slice(0, 25)) console.log(`  ${(time / keys).toFixed(2).padStart(7)} ms  ${key}`);
            // Time under each function (itself and what it calls), counted once per call path.
            const selfById = new Map();
            for (const id of cpu.samples) selfById.set(id, (selfById.get(id) ?? 0) + interval);
            const total = new Map();
            const walk = (node, open) => {
                const frame = node.callFrame;
                const key = `${frame.functionName || '(anonymous)'}  ${frame.url.split('/').slice(-2).join('/')}:${frame.lineNumber + 1}`;
                let sum = selfById.get(node.id) ?? 0;
                const inside = open.has(key);
                if (!inside) open.add(key);
                for (const child of node.children ?? []) sum += walk(byId.get(child), open);
                if (!inside) { open.delete(key); total.set(key, (total.get(key) ?? 0) + sum); }
                return sum;
            };
            walk(cpu.nodes[0], new Set());
            if (pathsOf) {
                // Who calls a function: time under it by its chain of callers.
                const parent = new Map();
                for (const node of cpu.nodes) for (const child of node.children ?? []) parent.set(child, node.id);
                const inclusive = (node) => (selfById.get(node.id) ?? 0) + (node.children ?? []).reduce((n, child) => n + inclusive(byId.get(child)), 0);
                const chains = new Map();
                for (const node of cpu.nodes) {
                    if (node.callFrame.functionName !== pathsOf) continue;
                    const names = [];
                    for (let id = parent.get(node.id); id !== undefined && names.length < 7; id = parent.get(id)) names.push(byId.get(id).callFrame.functionName || '(anonymous)');
                    if (names.includes(pathsOf)) continue;
                    const key = names.join(' < ');
                    chains.set(key, (chains.get(key) ?? 0) + inclusive(node));
                }
                console.log(`  -- callers of ${pathsOf}, per keystroke --`);
                for (const [key, time] of [...chains.entries()].sort((a, b) => b[1] - a[1]).slice(0, 12)) console.log(`  ${(time / keys).toFixed(2).padStart(7)} ms  ${key}`);
            }
            console.log(`  -- time under a function per keystroke --`);
            for (const [key, time] of [...total.entries()].filter(([key]) => /word\//.test(key)).sort((a, b) => b[1] - a[1]).slice(0, 60)) console.log(`  ${(time / keys).toFixed(2).padStart(7)} ms  ${key}`);
        }
        console.log(`  a keystroke (insertText in a plain paragraph, page ${Math.floor(loaded.pages / 2) + 1}): ${line(typed)}`);
        const carets = await tab.evaluate((n) => window.perf.carets(n), 40);
        console.log(`  a caret move (setSelection):  ${line(carets)}; frame builds ${carets.builds} in 40`);
        const arrows = await tab.evaluate(() => window.perf.pressed(['ArrowRight', 'ArrowLeft'], 20));
        console.log(`  ArrowRight:  ${line(arrows.times[0])}; ArrowLeft: ${line(arrows.times[1])}; frame builds ${arrows.builds} in 40`);
        const shifted = await tab.evaluate(() => window.perf.pressed(['Shift+ArrowRight', 'Shift+ArrowLeft'], 20));
        console.log(`  Shift+ArrowRight:  ${line(shifted.times[0])}; frame builds ${shifted.builds} in 40`);
        if (profile === 'enter') { await cdp.send('Profiler.enable'); await cdp.send('Profiler.setSamplingInterval', { interval: 200 }); await cdp.send('Profiler.start'); }
        const split = await tab.evaluate(() => window.perf.pressed(['Enter', 'Backspace'], 15));
        if (profile === 'enter') {
            // Time under each function while Enter and Backspace alternate (30 key presses).
            const { profile: cpu } = await cdp.send('Profiler.stop');
            const interval = cpu.timeDeltas.reduce((a, b) => a + b, 0) / Math.max(1, cpu.samples.length) / 1000;
            const byId = new Map(cpu.nodes.map((node) => [node.id, node]));
            const selfById = new Map();
            for (const id of cpu.samples) selfById.set(id, (selfById.get(id) ?? 0) + interval);
            const total = new Map();
            const walk = (node, open) => {
                const frame = node.callFrame;
                const key = `${frame.functionName || '(anonymous)'}  ${frame.url.split('/').slice(-2).join('/')}:${frame.lineNumber + 1}`;
                let sum = selfById.get(node.id) ?? 0;
                const inside = open.has(key);
                if (!inside) open.add(key);
                for (const child of node.children ?? []) sum += walk(byId.get(child), open);
                if (!inside) { open.delete(key); total.set(key, (total.get(key) ?? 0) + sum); }
                return sum;
            };
            walk(cpu.nodes[0], new Set());
            console.log(`  -- time under a function per key press (Enter / Backspace) --`);
            for (const [key, time] of [...total.entries()].filter(([key]) => /word\//.test(key)).sort((a, b) => b[1] - a[1]).slice(0, 60)) console.log(`  ${(time / 30).toFixed(2).padStart(7)} ms  ${key}`);
        }
        console.log(`  Enter:  ${line(split.times[0])}`);
        console.log(`  Backspace (merging the paragraphs again):  ${line(split.times[1])}; the last one laid out ${split.laidOut} and collected ${split.collected} of ${split.blocks} blocks`);
        const pasted = await tab.evaluate(() => window.perf.pastes(15, 'a short string '));
        console.log(`  a short string inserted (15 characters):  ${line(pasted)}`);
        const vertical = await tab.evaluate(() => window.perf.pressed(['ArrowDown', 'ArrowUp'], 20));
        console.log(`  ArrowDown:  ${line(vertical.times[0])}; ArrowUp: ${line(vertical.times[1])}`);
        const asked = await tab.evaluate(() => window.perf.queries(40));
        console.log(`  getCaretContext:  ${line(asked.context)}; getSelectionFormat: ${line(asked.format)}`);
        if (profile === 'select') { await cdp.send('Profiler.enable'); await cdp.send('Profiler.setSamplingInterval', { interval: 100 }); await cdp.send('Profiler.start'); }
        const all = await tab.evaluate((top) => window.perf.selectAll(top), loaded.top);
        if (profile === 'select') {
            // Where selecting everything and the queries after it spend their time (time under each function).
            const { profile: cpu } = await cdp.send('Profiler.stop');
            const interval = cpu.timeDeltas.reduce((a, b) => a + b, 0) / Math.max(1, cpu.samples.length) / 1000;
            const byId = new Map(cpu.nodes.map((node) => [node.id, node]));
            const selfById = new Map();
            for (const id of cpu.samples) selfById.set(id, (selfById.get(id) ?? 0) + interval);
            const keyOf = (frame) => `${frame.functionName || '(anonymous)'}  ${frame.url.split('/').slice(-2).join('/')}:${frame.lineNumber + 1}`;
            const total = new Map();
            const walk = (node, open) => {
                const key = keyOf(node.callFrame);
                let sum = selfById.get(node.id) ?? 0;
                const inside = open.has(key);
                if (!inside) open.add(key);
                for (const child of node.children ?? []) sum += walk(byId.get(child), open);
                if (!inside) { open.delete(key); total.set(key, (total.get(key) ?? 0) + sum); }
                return sum;
            };
            walk(cpu.nodes[0], new Set());
            console.log('  -- select all and the queries after it: time under a function --');
            for (const [key, time] of [...total.entries()].filter(([key]) => /word\//.test(key)).sort((a, b) => b[1] - a[1]).slice(0, 40)) console.log(`  ${time.toFixed(1).padStart(7)} ms  ${key}`);
        }
        console.log(`  select all:  ${ms(all.select)}; getSelectionFormat / getCaretContext with everything selected (asked three times): ${all.asked.slice(0, 6).map(ms).join(', ')}; getSelectionFormat after selecting everything a second time ${ms(all.asked[6])}; a scroll frame (400 px) with everything selected: ${line(all.scroll)}; the caret placed again ${ms(all.collapse)}`);
        {
            // Bold on everything, and undo: --profile format prints the time under each function for the step named by --step (bold | undo | redo).
            const profiled = opt('--step') ?? 'undo';
            const steps = ['select', 'bold', 'undo', 'redo', 'undo', 'collapse'];
            const results = [];
            let done = false;
            for (const step of steps) {
                const wanted = profile === 'format' && step === profiled && !done;
                if (wanted) { await cdp.send('Profiler.enable'); await cdp.send('Profiler.setSamplingInterval', { interval: 200 }); await cdp.send('Profiler.start'); }
                results.push(await tab.evaluate((name) => window.perf.formatAll(name), step));
                if (wanted) {
                    done = true;
                    const { profile: cpu } = await cdp.send('Profiler.stop');
                    const interval = cpu.timeDeltas.reduce((a, b) => a + b, 0) / Math.max(1, cpu.samples.length) / 1000;
                    const byId = new Map(cpu.nodes.map((node) => [node.id, node]));
                    const selfById = new Map();
                    for (const id of cpu.samples) selfById.set(id, (selfById.get(id) ?? 0) + interval);
                    const keyOf = (frame) => `${frame.functionName || '(anonymous)'}  ${frame.url.split('/').slice(-2).join('/')}:${frame.lineNumber + 1}`;
                    const total = new Map(), self = new Map();
                    // (--under FUNCTION: the self times are of what runs under that function only.)
                    let underDepth = under ? 0 : 1;
                    const walk = (node, open) => {
                        const key = keyOf(node.callFrame);
                        const entered = under && node.callFrame.functionName === under;
                        if (entered) underDepth++;
                        try { return walkNode(node, open, key); } finally { if (entered) underDepth--; }
                    };
                    const walkNode = (node, open, key) => {
                        let sum = selfById.get(node.id) ?? 0;
                        if (underDepth > 0) self.set(key, (self.get(key) ?? 0) + sum);
                        const inside = open.has(key);
                        if (!inside) open.add(key);
                        for (const child of node.children ?? []) sum += walk(byId.get(child), open);
                        if (!inside) { open.delete(key); total.set(key, (total.get(key) ?? 0) + sum); }
                        return sum;
                    };
                    walk(cpu.nodes[0], new Set());
                    console.log(`  -- ${step} of the whole document: time under a function --`);
                    for (const [key, time] of [...total.entries()].sort((a, b) => b[1] - a[1]).slice(0, 45)) console.log(`  ${time.toFixed(1).padStart(7)} ms  ${key}`);
                    console.log(`  -- ${step} of the whole document: self time --`);
                    for (const [key, time] of [...self.entries()].sort((a, b) => b[1] - a[1]).slice(0, 25)) console.log(`  ${time.toFixed(1).padStart(7)} ms  ${key}`);
                }
            }
            const [, bold, undo, redo, undoAgain] = results;
            console.log(`  the whole document made bold: ${ms(bold.took)} (the command ${ms(bold.command)}, the host's queries after it ${ms(bold.took - bold.command)}; ${bold.laidOut} of ${bold.blocks} blocks laid out in the last of ${bold.passes} passes, ${bold.stale} above the view kept as they were, ${bold.pages} pages${bold.complete ? '' : ', the layout not complete'}); undo ${ms(undo.took)} (${undo.laidOut} laid out, ${undo.pages} pages); redo ${ms(redo.took)}; undo again ${ms(undoAgain.took)}`);
        }
        const viewed = await tab.evaluate(() => window.perf.viewer(40));
        console.log(`  the viewer: package read ${(viewed.read / 1000).toFixed(2)} s, load and page frames ${(viewed.laidOut / 1000).toFixed(2)} s (${viewed.pages} pages), the first renderPage ${ms(viewed.first)}, the next 40: ${line(viewed.later)}`
            + (viewed.firstStats ? `; commands collected for the first page: ${viewed.firstStats.CollectedOnPaint} paragraphs of ${viewed.firstStats.Blocks} blocks, after 41 pages ${viewed.lastStats.CollectedOnPaint}` : ''));
        const scrolled = await tab.evaluate(([n, top]) => window.perf.scroll(n, top), [60, loaded.top]);
        console.log(`  a scroll frame (40 px):  ${line(scrolled)}`);
        console.log(`  render stats: ${JSON.stringify(await tab.evaluate(() => window.perf.renderStats()))}`);
    } finally { await browser.close(); server.close(); }
}
