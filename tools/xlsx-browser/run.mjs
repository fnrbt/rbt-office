#!/usr/bin/env node
// XLSX reader / writer check in a headless browser, against Fable output.
//
// The .NET suites cover reader -> writer -> reader natively; this drives the
// same thing through the browser path (JSZip, promises, Uint8Array part data):
// every workbook is opened with excel.load, written with excel.writeWorkbook,
// opened again, and its pictures, kept drawing objects (charts), threaded
// comments and kept parts are compared.
//
// It then runs the checks of the Spreadsheet web component (component.js):
// change events, dispose, sheet access and sizing, and its XLSX export.
//
// Usage:
//   dotnet fable src/Office.Formats --outDir $DIST/office-formats --lang JavaScript --noCache
//   dotnet fable src/Spreadsheet    --outDir $DIST/spreadsheet    --lang JavaScript --noCache
//   node tools/xlsx-browser/run.mjs --dist $DIST [--only NAME] [file.xlsx ...]
// (--only runs a single component check, or "roundtrip" for the reader / writer part.)
// The Spreadsheet bundle carries its own compiled copy of the XLSX reader and
// writer (Office.Formats is a project reference of it); the checks read the
// exported bytes back with the Office.Formats bundle.
// With no files it takes every workbook under tests/real-world and
// tests/test-documents-excel that has a part under xl/media, xl/charts or
// xl/drawings. Environment: PLAYWRIGHT, JSZIP_DIR, PORT (default 16023).
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
const verbose = (() => { const i = args.indexOf('--verbose'); if (i < 0) return false; args.splice(i, 1); return true; })();
const port = Number(process.env.PORT ?? 16023);
const only = opt('--only');
const jszipDir = [process.env.JSZIP_DIR, path.join(repo, 'node_modules/jszip'), path.resolve(repo, '../rbt-office/node_modules/jszip')].find((p) => p && fs.existsSync(p));
if (!jszipDir) { console.error('jszip not found; set JSZIP_DIR'); process.exit(2); }
const JSZip = createRequire(import.meta.url)(path.join(jszipDir, 'dist/jszip.min.js'));

let files = args;
if (files.length === 0) {
    for (const dir of ['tests/real-world', 'tests/test-documents-excel'])
        for (const name of fs.readdirSync(path.join(repo, dir)).sort()) {
            if (!/\.xls[xm]$/.test(name)) continue;
            const full = path.join(repo, dir, name);
            try {
                const zip = await JSZip.loadAsync(fs.readFileSync(full));
                if (Object.keys(zip.files).some((entry) => /^xl\/(media|charts|drawings\/drawing)/.test(entry))) files.push(full);
            } catch { /* not a zip: not for this check */ }
        }
}
const fixtures = new Map(files.map((file) => [path.basename(file), fs.readFileSync(file)]));
// Workbooks the component checks open, by path relative to the repository.
const componentFixture = (name) => { const file = path.join(repo, path.normalize(name)); return file.startsWith(repo) && fs.existsSync(file) ? fs.readFileSync(file) : null; };
const componentPage = `<!doctype html><meta charset="utf-8"><body style="margin:0"><script src="/jszip.min.js"></script>
<script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script><script type="module">
try { await import('/component.js'); } catch (e) { window.loadError = String(e && e.stack || e); }
</script>`;

const page = `<!doctype html><meta charset="utf-8"><script src="/jszip.min.js"></script>
<script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script><script type="module">
try {
    const main = await import('/office-formats/src/Main.js');
    const OF = typeof main.default === 'function' ? main.default() : main.default;
    const sum = (bytes) => { let s = 0; for (const b of bytes) s = (s + b) | 0; return s; };
    const describe = (wb) => OF.excel.sheetNames(wb).map((name) => {
        const sheet = OF.excel.sheet(name, wb);
        return { name,
            images: OF.excel.images(sheet).map((image) => ({ anchor: image.anchor, name: image.name ?? null, type: image.contentType, bytes: image.data.length, sum: sum(image.data) })),
            objects: OF.excel.drawingObjects(sheet).map((item) => ({ anchor: item.anchor, kind: item.kind, name: item.name ?? null })),
            threads: OF.excel.threadedComments(sheet) };
    });
    window.check = async (name) => {
        const original = new Uint8Array(await (await fetch('/fixture?name=' + encodeURIComponent(name))).arrayBuffer());
        const first = await OF.excel.load(original);
        const blob = await OF.excel.writeWorkbook(first);
        const saved = new Uint8Array(await blob.arrayBuffer());
        const second = await OF.excel.load(saved);
        const a = describe(first), b = describe(second);
        // Chart parts must come back byte for byte.
        const zip0 = await window.JSZip.loadAsync(original), zip1 = await window.JSZip.loadAsync(saved);
        const changed = [];
        for (const entry of Object.keys(zip0.files)) {
            if (!/^xl\\/charts\\/[^/]+\\.xml$/.test(entry)) continue;
            const before = await zip0.files[entry].async('base64');
            const after = zip1.files[entry] ? await zip1.files[entry].async('base64') : null;
            if (before !== after) changed.push(entry + (after === null ? ' (missing)' : ' (changed)'));
        }
        const count = (list, key) => list.reduce((n, sheet) => n + sheet[key].length, 0);
        return { same: JSON.stringify(a) === JSON.stringify(b), changed, images: count(a, 'images'), objects: count(a, 'objects'),
                 charts: a.reduce((n, sheet) => n + sheet.objects.filter((item) => item.kind === 'chart').length, 0), threads: count(a, 'threads'), bytes: saved.length };
    };
} catch (e) { window.loadError = String(e && e.stack || e); }
</script>`;
const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://x');
    const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
    if (url.pathname === '/') return send(200, 'text/html', page);
    if (url.pathname === '/jszip.min.js') return send(200, 'text/javascript', fs.readFileSync(path.join(jszipDir, 'dist/jszip.min.js')));
    if (url.pathname === '/jszip-shim.js') return send(200, 'text/javascript', 'export default window.JSZip;');
    if (url.pathname === '/component') return send(200, 'text/html', componentPage);
    if (url.pathname === '/component.js') return send(200, 'text/javascript', fs.readFileSync(path.join(here, 'component.js')));
    if (url.pathname === '/fixture') { const name = url.searchParams.get('name'); const data = fixtures.get(name) ?? componentFixture(name); return data ? send(200, 'application/octet-stream', data) : send(404, 'text/plain', 'no fixture'); }
    const file = path.join(dist, path.normalize(url.pathname));
    if (file.startsWith(dist) && fs.existsSync(file) && fs.statSync(file).isFile()) return send(200, 'text/javascript', fs.readFileSync(file));
    send(404, 'text/plain', 'not found');
});
if (!fs.existsSync(path.join(dist, 'office-formats/src/Main.js'))) { console.error(`missing office-formats/src/Main.js under ${dist}; Fable-compile first (see header)`); process.exit(2); }
await new Promise((resolve) => server.listen(port, '127.0.0.1', resolve));
const playwright = await import(process.env.PLAYWRIGHT ? pathToFileURL(process.env.PLAYWRIGHT).href : 'playwright');
const browser = await (playwright.chromium ?? playwright.default.chromium).launch();
let failed = 0, pictures = 0, charts = 0, componentChecks = 0;
try {
    const tab = await browser.newPage();
    tab.on('pageerror', (e) => console.error('page error:', e.message));
    await tab.goto(`http://127.0.0.1:${port}/`);
    await tab.waitForFunction('window.check !== undefined || window.loadError !== undefined');
    const loadError = await tab.evaluate('window.loadError');
    if (loadError) throw new Error('page failed to load modules: ' + loadError);
    for (const name of (only && only !== 'roundtrip' ? [] : fixtures.keys())) {
        let result;
        try { result = await tab.evaluate((n) => window.check(n), name); }
        catch (e) { failed++; console.log(`FAIL ${name}  threw: ${String(e.message).split('\n')[0]}`); continue; }
        const good = result.same && result.changed.length === 0;
        pictures += result.images; charts += result.charts;
        console.log(`${good ? 'ok  ' : 'FAIL'} ${name}  (${result.images} pictures, ${result.objects} kept objects of which ${result.charts} charts, ${result.threads} comment threads, ${result.bytes} bytes saved)`);
        if (!result.same) console.log('     pictures / objects / threaded comments differ after a save');
        for (const entry of result.changed) console.log('     ' + entry);
        if (!good) failed++;
    }
    // ---- the Spreadsheet web component ----
    if (only !== 'roundtrip') {
        if (!fs.existsSync(path.join(dist, 'spreadsheet/Spreadsheet.js'))) throw new Error(`missing spreadsheet/Spreadsheet.js under ${dist}; Fable-compile src/Spreadsheet first (see header)`);
        const componentTab = await browser.newPage({ viewport: { width: 1200, height: 900 } });
        const pageErrors = [];
        componentTab.on('pageerror', (e) => pageErrors.push(e.message));
        await componentTab.goto(`http://127.0.0.1:${port}/component`);
        await componentTab.waitForFunction('window.componentChecks !== undefined || window.loadError !== undefined', null, { timeout: 120000 });
        const componentLoadError = await componentTab.evaluate('window.loadError');
        if (componentLoadError) throw new Error('component page failed to load modules: ' + componentLoadError);
        for (const name of await componentTab.evaluate('Object.keys(window.componentChecks)')) {
            if (only && only !== name) continue;
            componentChecks++;
            let lines;
            try { lines = await componentTab.evaluate((n) => window.componentChecks[n](), name); }
            catch (e) { lines = ['FAIL threw: ' + String(e.message).split('\n')[0]]; }
            lines.push(...pageErrors.splice(0).map((message) => 'FAIL page error: ' + message));
            const bad = lines.filter((line) => !line.startsWith('ok'));
            console.log(`${bad.length === 0 ? 'ok  ' : 'FAIL'} ${name}  (${lines.length - bad.length} passed, ${bad.length} failed)`);
            for (const line of (verbose ? lines : bad)) console.log('     ' + line);
            if (bad.length) failed++;
        }
    }
} finally { await browser.close(); server.close(); }
console.log(failed === 0 ? `\nXLSX BROWSER CHECKS: ALL PASSED (${only && only !== 'roundtrip' ? 0 : fixtures.size} workbooks round-tripped with ${pictures} pictures and ${charts} charts, ${componentChecks} component checks)` : `\nXLSX BROWSER CHECKS: ${failed} FAILED`);
process.exit(failed === 0 ? 0 : 1);
