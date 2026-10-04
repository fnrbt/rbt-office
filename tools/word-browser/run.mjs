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
const all = ['explicit-format', 'page-frames', 'ime', 'typing', 'track-changes'];
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
const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://x');
    const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
    if (url.pathname === '/') return send(200, 'text/html', page(url.searchParams.get('check')));
    if (url.pathname === '/jszip.min.js') return send(200, 'text/javascript', fs.readFileSync(path.join(jszipDir, 'dist/jszip.min.js')));
    if (url.pathname === '/jszip-shim.js') return send(200, 'text/javascript', 'export default window.JSZip;');
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
    for (const check of checks) {
        const tab = await browser.newPage({ viewport: { width: 1100, height: 900 } });
        const errors = [];
        tab.on('pageerror', (e) => errors.push('page error: ' + e.message));
        await tab.goto(`http://127.0.0.1:${port}/?check=${encodeURIComponent(check)}`);
        await tab.waitForFunction('window.result !== undefined || window.drive !== undefined', null, { timeout: 60000 });
        if (await tab.evaluate('window.drive !== undefined')) {
            // The check wants real input: [["click", x, y] | ["type", text] | ["press", key]] ...
            for (const [action, a, b] of await tab.evaluate('window.drive')) {
                if (action === 'click') await tab.mouse.click(a, b);
                else if (action === 'type') await tab.keyboard.type(a);
                else if (action === 'press') await tab.keyboard.press(a);
            }
            await tab.evaluate('window.finish()');
            await tab.waitForFunction('window.result !== undefined', null, { timeout: 60000 });
        }
        const result = await tab.evaluate('window.result');
        const problems = [...errors, ...result.log.filter((line) => !line.startsWith('ok'))];
        console.log(`${result.failed === 0 && errors.length === 0 ? 'ok  ' : 'FAIL'} ${check}  (${result.passed} passed, ${result.failed} failed)`);
        for (const line of problems) console.log('     ' + line);
        if (result.failed || errors.length) failedChecks++;
        await tab.close();
    }
} finally {
    await browser.close();
    server.close();
}
console.log(failedChecks === 0 ? '\nWORD BROWSER CHECKS: ALL PASSED' : `\nWORD BROWSER CHECKS: ${failedChecks} check(s) FAILED`);
process.exit(failedChecks === 0 ? 0 : 1);
