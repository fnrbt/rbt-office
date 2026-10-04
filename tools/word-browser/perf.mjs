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
//   node tools/word-browser/perf.mjs --dist $DIST [--repeat N] [--keys N] [--profile keys] [--callers FUNCTION] [--painter]
// --profile keys prints the functions with the most self time while typing, and the time under each
// function (--profile enter: the time under each function while Enter and Backspace alternate); --callers FUNCTION adds the time under that function by its chain of callers.
// --painter installs a (trivial) text path painter and measurers first.
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
const painter = flag('--painter');
const pathsOf = opt('--callers');
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
        async load(usePainter) {
            const bytes = new Uint8Array(await (await fetch('/long.docx')).arrayBuffer());
            const t0 = performance.now();
            Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
            const loaded = performance.now() - t0;
            if (usePainter) {
                Word.setTextMeasurers(editor, (text, style) => text.length * style.FontSize * 0.5, (text, style, start, length) => Math.max(0, Math.min(length, text.length - start)) * style.FontSize * 0.5);
                Word.setTextPathPainter(editor, (text, style, x, baseline, width) => false);
            }
            const size = Word.getDocumentSize(editor);
            // The caret in a plain paragraph near the middle of the document.
            const frames = Word.getPageFrames(editor);
            const target = frames[Math.floor(frames.length / 2)];
            Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: target.Top, scrollLeft: 0, width: 1000, height: 800 });
            return { loaded, pages: size.PageCount, top: target.Top };
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
        const loaded = await tab.evaluate((usePainter) => window.perf.load(usePainter), painter);
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
        const scrolled = await tab.evaluate(([n, top]) => window.perf.scroll(n, top), [60, loaded.top]);
        console.log(`  a scroll frame (40 px):  ${line(scrolled)}`);
        console.log(`  render stats: ${JSON.stringify(await tab.evaluate(() => window.perf.renderStats()))}`);
    } finally { await browser.close(); server.close(); }
}
