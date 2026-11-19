#!/usr/bin/env node
// What the shell's navigation pane and outline view cost on a long document, in headless Chromium,
// against Fable output: keystrokes in a plain paragraph in the middle of the document (the shell
// mounted, as a user has it) with the navigation pane closed and open, the longest time the page
// could not answer meanwhile (the pane draws itself again a while after the typing, not at each key),
// the outline computed once, the pane opened, outline view entered, and a promote in outline view.
//
// The document is perf.mjs's: tests/test-documents-wild/docx4j-getting-started.docx with its body
// repeated (--repeat N, default 8; 366 pages on the reference machine).
//
// The keystrokes are timed in runs of --keys, --gap ms apart, each run into another plain paragraph of
// the middle of the document, the pane closed and open in turn. Keystrokes cost more from run to run
// whatever the pane (the paragraphs are further down): PANE=0 keeps the pane closed in every run, the
// control to compare with.
//
// Usage: node tools/word-browser/perf-outline.mjs --dist $DIST [--repeat N] [--keys N] [--gap MS]
// Environment: PLAYWRIGHT, JSZIP_DIR, PORT (default 16126).
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.resolve(here, '../..');
const args = process.argv.slice(2);
const opt = (name) => { const i = args.indexOf(name); if (i < 0) return undefined; const v = args[i + 1]; args.splice(i, 2); return v; };
const dist = path.resolve(opt('--dist') ?? path.join(repo, 'dist'));
const repeat = Number(opt('--repeat') ?? 8);
const keys = Number(opt('--keys') ?? 60);
const gap = Number(opt('--gap') ?? 40);
const port = Number(process.env.PORT ?? 16126);
const jszipDir = [process.env.JSZIP_DIR, path.join(repo, 'node_modules/jszip')].find((p) => p && fs.existsSync(p));
if (!jszipDir) { console.error('jszip not found; set JSZIP_DIR'); process.exit(2); }
const JSZip = createRequire(import.meta.url)(path.join(jszipDir, 'dist/jszip.min.js'));

const zip = await JSZip.loadAsync(fs.readFileSync(path.join(repo, 'tests/test-documents-wild/docx4j-getting-started.docx')));
const xml = await zip.file('word/document.xml').async('string');
const open = xml.indexOf('>', xml.indexOf('<w:body')) + 1, close = xml.lastIndexOf('</w:body>');
const sect = xml.lastIndexOf('<w:sectPr', close);
const blocks = xml.slice(open, sect > open ? sect : close), tail = xml.slice(sect > open ? sect : close, close);
zip.file('word/document.xml', xml.slice(0, open) + blocks.repeat(repeat) + tail + xml.slice(close));
const docx = await zip.generateAsync({ type: 'nodebuffer' });

const page = `<!doctype html><meta charset="utf-8"><body style="margin:0"><div id="shell" style="height:880px"></div>
<script src="/jszip.min.js"></script><script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script><script type="module">
try {
    const main = await import('/office-formats/src/Main.js');
    const OF = typeof main.default === 'function' ? main.default() : main.default;
    const Word = (await import('/word/Word.js')).default;
    const Host = await import('/word/WordHost.js'), Shell = await import('/word/WordShell.js'), Dom = await import('/word/FableUI/Dom.js');
    const host = Host.create('', 'word-shell');
    Dom.mountApp('#shell', Shell.view(host));
    const editor = host.Editor;
    const control = id => document.querySelector('[data-testid="' + id + '"]');
    const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
    const stats = (times) => { const sorted = times.slice().sort((a, b) => a - b); return { mean: times.reduce((a, b) => a + b, 0) / times.length, median: sorted[Math.floor(sorted.length / 2)], p90: sorted[Math.floor(sorted.length * 0.9)], max: sorted[sorted.length - 1] }; };
    window.perf = {
        async load() {
            const bytes = new Uint8Array(await (await fetch('/long.docx')).arrayBuffer());
            const pkg = await OF.processors.forEditorPackage(bytes);
            Host.importDocumentPackage(host, pkg, 'long.docx');
            await Word.whenLayoutComplete(editor);
            // Plain paragraphs near the middle: each run of keystrokes types into another (a paragraph
            // typed into grows, and its keystrokes cost more: runs compared are of like paragraphs).
            const lines = Word.getOutlineLines(editor, false);
            const length = lines[Math.floor(lines.length / 2)].Text.length;
            window.targets = [];
            for (let at = Math.floor(lines.length / 2); at < lines.length && window.targets.length < 12; at++)
                if (lines[at].Kind === 'paragraph' && lines[at].Level === 10 && lines[at].Text.length > 60 && lines[at].Text.length < 400) window.targets.push(lines[at].BlockId);
            Word.setSelection(editor, window.targets[0], 5, window.targets[0], 5);
            return { pages: Word.getPageCount(editor), blocks: lines.length, headings: Word.getOutline(editor).length };
        },
        async typing(count, gap, target) {
            const id = window.targets[target % window.targets.length];
            Word.setSelection(editor, id, 5, id, 5);
            for (let i = 0; i < 3; i++) Word.insertText(editor, 'x');
            await sleep(700);
            const times = [];
            let longest = 0, last = performance.now();
            const ticker = setInterval(() => { const now = performance.now(); if (now - last > longest) longest = now - last; last = now; }, 4);
            for (let i = 0; i < count; i++) {
                const t0 = performance.now(); Word.insertText(editor, 'x'); times.push(performance.now() - t0);
                await sleep(gap);
            }
            await sleep(700);
            clearInterval(ticker);
            return { ...stats(times), longest };
        },
        outlineOnce() {
            Word.insertText(editor, 'y');
            const t0 = performance.now(); const n = Word.getOutline(editor).length; const outline = performance.now() - t0;
            const t1 = performance.now(); Word.getOutline(editor); const cached = performance.now() - t1;
            return { outline, cached, n };
        },
        async togglePane() {
            const t0 = performance.now();
            control('word-view-menu').click(); control('word-menu-view-navigation').click();
            const opened = performance.now() - t0;
            await sleep(50);
            return { opened, items: document.querySelectorAll('[data-testid="word-nav-item"]').length };
        },
        async caretMoves(count) {
            const sel = Word.getSelection(editor), times = [];
            for (let i = 0; i < count; i++) { const at = 1 + (i % 5); const t0 = performance.now(); Word.setSelection(editor, sel.FocusBlockId, at, sel.FocusBlockId, at); times.push(performance.now() - t0); }
            return stats(times);
        },
        async outlineView() {
            let t0 = performance.now();
            control('word-view-menu').click(); control('word-menu-view-outline').click();
            const count = () => document.querySelectorAll('[data-testid="word-outline-line"]').length;
            while (count() === 0 && performance.now() - t0 < 10000) await new Promise(resolve => queueMicrotask(resolve));
            const entered = performance.now() - t0;
            const lines = count();
            await sleep(100);
            t0 = performance.now();
            control('word-outline-promote').click();
            const promoted = performance.now() - t0;
            await sleep(100);
            t0 = performance.now();
            Word.undo(editor);
            const undone = performance.now() - t0;
            await sleep(100);
            t0 = performance.now();
            control('word-outline-close').click();
            const closed = performance.now() - t0;
            return { entered, lines, promoted, undone, closed };
        },
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
await new Promise((resolve) => server.listen(port, '127.0.0.1', resolve));
const playwright = await import(process.env.PLAYWRIGHT ? pathToFileURL(process.env.PLAYWRIGHT).href : 'playwright');
const browser = await (playwright.chromium ?? playwright.default.chromium).launch();
const ms = (t) => t.toFixed(1) + ' ms';
const line = (s) => `mean ${ms(s.mean)}, median ${ms(s.median)}, p90 ${ms(s.p90)}, max ${ms(s.max)}`;
try {
    const tab = await browser.newPage({ viewport: { width: 1300, height: 900 } });
    tab.on('pageerror', (e) => console.error('page error:', e.message));
    await tab.goto(`http://127.0.0.1:${port}/`);
    await tab.waitForFunction('window.perf !== undefined || window.loadError !== undefined', null, { timeout: 240000 });
    const loadError = await tab.evaluate('window.loadError');
    if (loadError) throw new Error(loadError);
    const loaded = await tab.evaluate(() => window.perf.load());
    console.log(`Chromium ${browser.version()}; docx4j-getting-started x ${repeat}: ${loaded.pages} pages, ${loaded.blocks} blocks, ${loaded.headings} headings`);
    // Runs of keystrokes with the pane closed and open, alternating, each into another paragraph.
    const runs = { closed: [], open: [] };
    let pane = null;
    for (let round = 0; round < 3; round++) {
        runs.closed.push(await tab.evaluate(([n, g, t]) => window.perf.typing(n, g, t), [keys, gap, round * 2]));
        if (process.env.PANE !== "0") pane = await tab.evaluate(() => window.perf.togglePane());
        runs.open.push(await tab.evaluate(([n, g, t]) => window.perf.typing(n, g, t), [keys, gap, round * 2 + 1]));
        if (process.env.PANE !== "0") await tab.evaluate(() => window.perf.togglePane());
    }
    const merged = (list) => { const medians = list.map(r => r.median), all = { mean: list.reduce((a, r) => a + r.mean, 0) / list.length, median: medians.slice().sort((a, b) => a - b)[1], p90: Math.max(...list.map(r => r.p90)), max: Math.max(...list.map(r => r.max)) }; return [all, Math.max(...list.map(r => r.longest)), medians]; };
    const [closedAll, closedLongest, closedMedians] = merged(runs.closed), [openAll, openLongest, openMedians] = merged(runs.open);
    if (pane) console.log(`navigation pane opened: ${ms(pane.opened)}, ${pane.items} items`);
    console.log(`keystroke, pane closed: ${line(closedAll)}; longest unresponsive ${ms(closedLongest)} (run medians ${closedMedians.map(ms).join(', ')})`);
    console.log(`keystroke, pane open:   ${line(openAll)}; longest unresponsive ${ms(openLongest)} (run medians ${openMedians.map(ms).join(', ')})`);
    const movesClosed = await tab.evaluate(() => window.perf.caretMoves(40));
    await tab.evaluate(() => window.perf.togglePane());
    const movesOpen = await tab.evaluate(() => window.perf.caretMoves(40));
    console.log(`caret move, pane closed: ${line(movesClosed)}`);
    console.log(`caret move, pane open:   ${line(movesOpen)}`);
    const once = await tab.evaluate(() => window.perf.outlineOnce());
    console.log(`the outline computed after an edit: ${ms(once.outline)} (${once.n} headings); asked again: ${ms(once.cached)}`);
    const view = await tab.evaluate(() => window.perf.outlineView());
    console.log(`outline view entered: ${ms(view.entered)} (first ${view.lines} lines; the rest in slices); promote there: ${ms(view.promoted)}; its undo: ${ms(view.undone)}; closed: ${ms(view.closed)}`);
    const after = await tab.evaluate(([n, g, t]) => window.perf.typing(n, g, t), [keys, gap, 7]);
    console.log(`keystroke, pane open, after outline view: ${line(after)}; longest unresponsive ${ms(after.longest)}`);
} finally {
    await browser.close();
    server.close();
}
