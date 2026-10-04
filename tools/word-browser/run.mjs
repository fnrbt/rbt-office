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
//   commands         point sizes, colours, fonts, styles, spacing, page break, table rows / columns / merge / split, hyperlinks, image resize by its handles
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
    'placeholder-and-frames', 'caret-movement', 'revision', 'viewer-measurers', 'page-centering', 'commands'];
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
            // The check wants real input: [["click" | "dblclick" | "tripleclick" | "shiftclick", x, y] | ["type", text] | ["press", key] | ["down", x, y] | ["move", x, y] | ["up"] | ["wheel", dx, dy] | ["eval", js]] ...
            for (const [action, a, b] of await tab.evaluate('window.drive')) {
                if (action === 'click') await tab.mouse.click(a, b);
                else if (action === 'dblclick') await tab.mouse.dblclick(a, b);
                else if (action === 'tripleclick') await tab.mouse.click(a, b, { clickCount: 3 });
                else if (action === 'shiftclick') { await tab.keyboard.down('Shift'); await tab.mouse.click(a, b); await tab.keyboard.up('Shift'); }
                else if (action === 'type') await tab.keyboard.type(a);
                else if (action === 'press') await tab.keyboard.press(a);
                else if (action === 'down') { await tab.mouse.move(a, b); await tab.mouse.down(); }
                else if (action === 'move') await tab.mouse.move(a, b, { steps: 4 });
                else if (action === 'up') await tab.mouse.up();
                else if (action === 'wheel') await tab.mouse.wheel(a, b);
                else if (action === 'eval') await tab.evaluate(a);
            }
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
