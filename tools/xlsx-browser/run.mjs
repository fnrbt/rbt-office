#!/usr/bin/env node
// XLSX reader / writer check in a headless browser, against Fable output.
//
// The .NET suites cover reader -> writer -> reader natively; this drives the
// same thing through the browser path (JSZip, promises, Uint8Array part data):
// every workbook is opened with excel.load, written with excel.writeWorkbook,
// opened again, and its pictures, kept drawing objects (charts), threaded
// comments and kept parts are compared.
//
// It then runs the checks of the Spreadsheet web component (component.js,
// host-api.js, sheet-api.js, features.js, chrome.js): change events, dispose, sheet access and
// sizing, its XLSX export, what hosts rely on (borders, text overflow, row /
// column counts, revisions, package integrity of exported files, lazy sheets,
// incremental recalculation), and the host API ("api-..." checks: reading
// cells, formats, ...).
//
// Usage:
//   dotnet fable src/Office.Formats --outDir $DIST/office-formats --lang JavaScript --noCache
//   dotnet fable src/Spreadsheet    --outDir $DIST/spreadsheet    --lang JavaScript --noCache
//   node tools/xlsx-browser/run.mjs --dist $DIST [--only NAME] [file.xlsx ...]
// (--only runs the named component checks (comma separated), or "roundtrip" for the reader / writer part.)
// The Spreadsheet bundle carries its own compiled copy of the XLSX reader and
// writer (Office.Formats is a project reference of it); the checks read the
// exported bytes back with the Office.Formats bundle.
// With no files it takes every workbook under tests/real-world and
// tests/test-documents-excel that has a part under xl/media, xl/charts or
// xl/drawings. Environment: PLAYWRIGHT, JSZIP_DIR, PORT (default 16023); for the sheet-pdf check RB_WORD_FONT_DIR (fonts),
// SHEET_PDF_REFERENCE (the .NET PDF to compare with) and SHEET_PDF_OUT (where the PDFs it made are written).
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
// Each check gets this long (the page is closed when one takes longer).
const checkTimeout = Number(process.env.CHECK_TIMEOUT ?? 240000);
// --keep DIR: saves of every workbook of the tests into DIR instead of the checks: as it was opened
// (.save.xlsx), a cell typed (.edit.xlsx), a row and a column inserted (.insert.xlsx), beside a copy of
// the original (.original.xlsx): for tools/excel-parity/open-all.vbs.
const keepDir = opt('--keep');
const only = keepDir ? 'keep-saves' : opt('--only');
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
const fontDir = path.resolve(process.env.RB_WORD_FONT_DIR ?? path.join(repo, 'artifacts/fonts/windows'));
const fixtures = new Map(files.map((file) => [path.basename(file), fs.readFileSync(file)]));
// Workbooks the component checks open, by path relative to the repository.
const componentFixture = (name) => { const file = path.join(repo, path.normalize(name)); return file.startsWith(repo) && fs.existsSync(file) ? fs.readFileSync(file) : null; };
const componentPage = `<!doctype html><meta charset="utf-8"><body style="margin:0"><script src="/jszip.min.js"></script>
<script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script><script type="module">
try { await import('/component.js'); await import('/host-api.js'); await import('/sheet-api.js'); await import('/features.js'); await import('/chrome.js'); window.componentReady = true; } catch (e) { window.loadError = String(e && e.stack || e); }
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
    if (/^\/[\w-]+\.js$/.test(url.pathname) && fs.existsSync(path.join(here, url.pathname.slice(1)))) return send(200, 'text/javascript', fs.readFileSync(path.join(here, url.pathname.slice(1))));
    // The TrueType fonts the PDF checks set text in: RB_WORD_FONT_DIR (as the .NET PDF runners), else artifacts/fonts/windows.
    if (url.pathname === '/font') { const name = path.basename(url.searchParams.get('name') ?? ''); const file = path.join(fontDir, name); return name && fs.existsSync(file) ? send(200, 'application/octet-stream', fs.readFileSync(file)) : send(404, 'text/plain', 'no font'); }
    // The .NET PDF of the sheet the sheet-pdf check exports (SHEET_PDF_REFERENCE: tools/spreadsheet-pdf-runner's output), to compare with.
    if (url.pathname === '/sheet-pdf-reference') return process.env.SHEET_PDF_REFERENCE && fs.existsSync(process.env.SHEET_PDF_REFERENCE) ? send(200, 'application/pdf', fs.readFileSync(process.env.SHEET_PDF_REFERENCE)) : send(404, 'text/plain', 'no reference');
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
        if (keepDir) {
            const walk = (dir) => fs.readdirSync(dir, { withFileTypes: true }).flatMap((entry) => entry.isDirectory() ? walk(path.join(dir, entry.name)) : [path.join(dir, entry.name)]);
            const list = walk(path.join(repo, 'tests')).filter((file) => /\.xlsx$/.test(file) && fs.statSync(file).size < 6000000 && (!process.env.KEEP_MATCH || new RegExp(process.env.KEEP_MATCH).test(file))).map((file) => path.relative(repo, file));
            fs.mkdirSync(keepDir, { recursive: true });
            for (const file of list) fs.copyFileSync(path.join(repo, file), path.join(keepDir, path.basename(file, '.xlsx') + '.original.xlsx'));
            await componentTab.exposeFunction('keepSave', (name, base64) => fs.writeFileSync(path.join(keepDir, name), Buffer.from(base64, 'base64')));
            await componentTab.addInitScript((names) => { window.keepList = names; }, list);
        }
        await componentTab.goto(`http://127.0.0.1:${port}/component`);
        await componentTab.waitForFunction('window.componentReady !== undefined || window.loadError !== undefined', null, { timeout: 120000 });
        const componentLoadError = await componentTab.evaluate('window.loadError');
        if (componentLoadError) throw new Error('component page failed to load modules: ' + componentLoadError);
        const debuggerSession = await componentTab.context().newCDPSession(componentTab);
        await debuggerSession.send('Debugger.enable');
        // For the checks that a disposed grid can be collected: a full garbage collection on request.
        await componentTab.exposeFunction('collectGarbage', async () => { await debuggerSession.send('HeapProfiler.enable'); await debuggerSession.send('HeapProfiler.collectGarbage'); });
        // For the pixel checks of the DOM grid: a screenshot of a region of the page, as a PNG in base64.
        await componentTab.exposeFunction('screenshotRegion', async (x, y, width, height) => {
            const png = await componentTab.screenshot({ clip: { x, y, width, height } });
            // (SCREENSHOT_DIR keeps what the pixel checks saw, for a look when one fails.)
            if (process.env.SCREENSHOT_DIR) fs.writeFileSync(path.join(process.env.SCREENSHOT_DIR, `shot-${Date.now()}.png`), png);
            return png.toString('base64');
        });
        // For the checks of what the user does: the browser's own mouse and keyboard (trusted events,
        // hit-tested by the browser). realMouse('click' | 'move' | 'down' | 'up' | 'dblclick', x, y, options), realKey('Alt+ArrowDown').
        await componentTab.exposeFunction('realMouse', async (type, x, y, options = {}) => {
            if (type === 'click') await componentTab.mouse.click(x, y, options);
            else if (type === 'dblclick') await componentTab.mouse.dblclick(x, y, options);
            else if (type === 'move') await componentTab.mouse.move(x, y, options);
            else if (type === 'down') await componentTab.mouse.down(options);
            else if (type === 'up') await componentTab.mouse.up(options);
            else throw new Error('realMouse: ' + type);
        });
        // SHEET_PDF_OUT: a directory the PDF checks write the PDFs they made into (keepPdf(name, base64)), for a look.
        await componentTab.exposeFunction('keepPdf', (name, base64) => { if (process.env.SHEET_PDF_OUT) { fs.mkdirSync(process.env.SHEET_PDF_OUT, { recursive: true }); fs.writeFileSync(path.join(process.env.SHEET_PDF_OUT, path.basename(name)), Buffer.from(base64, 'base64')); } });
        await componentTab.exposeFunction('realKey', async (key) => { await componentTab.keyboard.press(key); });
        await componentTab.exposeFunction('realType', async (text) => { await componentTab.keyboard.type(text); });
        for (const name of await componentTab.evaluate('Object.keys(window.componentChecks)')) {
            if (only && !only.split(',').includes(name)) continue;
            componentChecks++;
            let lines;
            try { lines = await Promise.race([componentTab.evaluate((n) => window.componentChecks[n](), name), new Promise((_, reject) => setTimeout(() => reject(new Error(`no result after ${checkTimeout} ms`)), checkTimeout))]); }
            catch (e) {
                lines = ['FAIL threw: ' + String(e.message).split('\n')[0]];
                if (/no result after/.test(String(e.message))) {
                    // Where the page is stuck: pause it and read the stack.
                    try {
                        const paused = new Promise((resolve) => debuggerSession.once('Debugger.paused', resolve));
                        await debuggerSession.send('Debugger.pause');
                        const { callFrames } = await Promise.race([paused, new Promise((_, reject) => setTimeout(() => reject(new Error('not paused')), 10000))]);
                        for (const frame of callFrames.slice(0, 14)) lines.push(`FAIL   at ${frame.functionName || '(anonymous)'} ${frame.url.split('/').slice(-2).join('/')}:${frame.location.lineNumber + 1}`);
                    } catch { /* the page is gone */ }
                }
            }
            lines.push(...pageErrors.splice(0).map((message) => 'FAIL page error: ' + message));
            // Every workbook the check exported went through the package-integrity check (component.js).
            try { lines.push(...[...new Set(await componentTab.evaluate('window.exportProblems.splice(0)'))].map((message) => 'FAIL exported package: ' + message)); } catch { /* the page is gone */ }
            const bad = lines.filter((line) => !line.startsWith('ok'));
            console.log(`${bad.length === 0 ? 'ok  ' : 'FAIL'} ${name}  (${lines.length - bad.length} passed, ${bad.length} failed)`);
            for (const line of (verbose ? lines : bad)) console.log('     ' + line);
            if (bad.length) failed++;
        }
            try { console.log(`(${await componentTab.evaluate('window.exportsChecked')} exported workbooks checked for package integrity)`); } catch { /* the page is gone */ }
    }
} finally { await browser.close(); server.close(); }
console.log(failed === 0 ? `\nXLSX BROWSER CHECKS: ALL PASSED (${only && only !== 'roundtrip' ? 0 : fixtures.size} workbooks round-tripped with ${pictures} pictures and ${charts} charts, ${componentChecks} component checks)` : `\nXLSX BROWSER CHECKS: ${failed} FAILED`);
process.exit(failed === 0 ? 0 : 1);
