#!/usr/bin/env node
// DOCX save round-trip check, run in a headless browser against Fable output.
//
// DocumentWriter is JS-only, so the save path cannot be exercised from the
// .NET suites. This drives the real browser pipeline instead:
//
//   docx -> Main.processors.forEditorPackage -> Docx.importFromEditorPackage
//        -> Docx.exportDocumentPackage -> DocumentWriter.writeFromEditorPackage
//        -> docx' -> (read again) -> docx'' -> (read again)
//
// and asserts that
//   * what the reader sees after the first save equals what it saw in the
//     original (fidelity: sizes, indents, spacing, lists, tables, tabs, ...),
//   * a second save reproduces the first one exactly (fixed point),
//   * every written XML part is well formed and the children of the
//     elements this writer produces are in schema (xsd:sequence) order.
//
// Usage:
//   dotnet fable src/Office.Formats --outDir $DIST/office-formats --lang JavaScript --noCache
//   dotnet fable src/Word           --outDir $DIST/word           --lang JavaScript --noCache
//   node tools/docx-roundtrip/roundtrip.mjs --dist $DIST [--out DIR] [--verbose] [fixture.docx ...]
//
// For every document it also checks that import followed by export with no
// edit yields the same number of blocks (empty paragraphs at the end included).
//
// With no fixture arguments it runs the built-in synthetic documents (one per
// feature the writer has to preserve) plus the small corpus files listed in
// DEFAULT_CORPUS. --out keeps the saved documents (pass 1 and 2) so the .NET
// reader can be pointed at them: dotnet fsi tools/docx-roundtrip/native-check.fsx DIR
//
// Environment:
//   PLAYWRIGHT  path to playwright's index.mjs (default: resolve "playwright")
//   JSZIP_DIR   path to a jszip package dir   (default: node_modules/jszip here
//               or in a sibling checkout)
//   PORT        port to serve on (default 16021)
// Exit code is non-zero when any check fails.

import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import { syntheticFixtures } from './fixtures.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.resolve(here, '../..');
const args = process.argv.slice(2);
const opt = (name) => { const i = args.indexOf(name); if (i < 0) return undefined; const v = args[i + 1]; args.splice(i, 2); return v; };
const flag = (name) => { const i = args.indexOf(name); if (i < 0) return false; args.splice(i, 1); return true; };
const dist = path.resolve(opt('--dist') ?? path.join(repo, 'dist'));
const outDir = opt('--out');
const verbose = flag('--verbose');
const expectFailures = flag('--report-only');
const port = Number(process.env.PORT ?? 16021);

const firstExisting = (candidates) => candidates.find((p) => p && fs.existsSync(p));
const jszipDir = firstExisting([
    process.env.JSZIP_DIR,
    path.join(repo, 'node_modules/jszip'),
    path.resolve(repo, '../rbt-office/node_modules/jszip'),
]);
if (!jszipDir) { console.error('jszip not found; set JSZIP_DIR'); process.exit(2); }
const JSZip = createRequire(import.meta.url)(path.join(jszipDir, 'dist/jszip.min.js'));

const DEFAULT_CORPUS = [
    'tests/test-documents/sample-formatted.docx',
    'tests/test-documents-features/test-lists.docx',
    'tests/test-documents-features/test-formatting.docx',
    'tests/test-documents-features/test-complex-tables.docx',
    'tests/test-documents-features/test-header-footer.docx',
    'tests/test-documents-features/test-multi-headers.docx',
];

const mime = { '.js': 'text/javascript', '.mjs': 'text/javascript', '.html': 'text/html', '.json': 'application/json' };
const fixtures = new Map(); // name -> Buffer

const page = `<!doctype html><meta charset="utf-8">
<script src="/jszip.min.js"></script>
<script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script>
<script type="module" src="/page.js"></script>`;

const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://x');
    const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
    if (url.pathname === '/') return send(200, 'text/html', page);
    if (url.pathname === '/jszip.min.js') return send(200, 'text/javascript', fs.readFileSync(path.join(jszipDir, 'dist/jszip.min.js')));
    if (url.pathname === '/jszip-shim.js') return send(200, 'text/javascript', 'export default window.JSZip;');
    if (url.pathname === '/page.js') return send(200, 'text/javascript', fs.readFileSync(path.join(here, 'page.js')));
    if (url.pathname === '/fixture') {
        const data = fixtures.get(url.searchParams.get('name'));
        return data ? send(200, 'application/octet-stream', data) : send(404, 'text/plain', 'no fixture');
    }
    const file = path.join(dist, path.normalize(url.pathname));
    if (file.startsWith(dist) && fs.existsSync(file) && fs.statSync(file).isFile())
        return send(200, mime[path.extname(file)] ?? 'application/octet-stream', fs.readFileSync(file));
    send(404, 'text/plain', 'not found');
});

async function main() {
    for (const required of ['office-formats/src/Main.js', 'word/Docx.js'])
        if (!fs.existsSync(path.join(dist, required))) { console.error(`missing ${required} under ${dist}; Fable-compile first (see header)`); process.exit(2); }

    const expectations = new Map();
    if (args.length === 0) {
        for (const f of await syntheticFixtures(JSZip)) { fixtures.set(f.name, f.data); expectations.set(f.name, f.expect ?? []); }
        for (const rel of DEFAULT_CORPUS) {
            const p = path.join(repo, rel);
            if (fs.existsSync(p)) fixtures.set(rel, fs.readFileSync(p));
        }
    } else {
        for (const p of args) fixtures.set(p, fs.readFileSync(p));
    }

    await new Promise((resolve) => server.listen(port, '127.0.0.1', resolve));
    let playwright;
    try {
        playwright = await import(process.env.PLAYWRIGHT ? pathToFileURL(process.env.PLAYWRIGHT).href : 'playwright');
    } catch (e) {
        console.error('playwright not found; set PLAYWRIGHT=/path/to/playwright/index.mjs'); server.close(); process.exit(2);
    }
    const browser = await (playwright.chromium ?? playwright.default.chromium).launch();
    let failed = 0;
    try {
        const tab = await browser.newPage();
        tab.on('pageerror', (e) => console.error('page error:', e.message));
        if (verbose) tab.on('console', (m) => console.log('  [page]', m.text()));
        await tab.goto(`http://127.0.0.1:${port}/`);
        await tab.waitForFunction('window.roundTrip !== undefined || window.loadError !== undefined');
        const loadError = await tab.evaluate('window.loadError');
        if (loadError) throw new Error('page failed to load modules: ' + loadError);

        if (outDir) fs.mkdirSync(outDir, { recursive: true });
        for (const name of fixtures.keys()) {
            const result = await tab.evaluate(async ([n, expect]) => window.roundTrip(n, expect), [name, expectations.get(name) ?? []]);
            const problems = result.problems;
            console.log(`${problems.length === 0 ? 'ok  ' : 'FAIL'} ${name}  (${result.stats})`);
            const limit = verbose ? 200 : 12;
            for (const p of problems.slice(0, limit)) console.log('     ' + p);
            if (problems.length > limit) console.log(`     ... ${problems.length - limit} more`);
            if (verbose && result.xml) console.log(result.xml);
            if (problems.length) failed++;
            if (outDir && result.saved) {
                const base = path.basename(name).replace(/\.docx$/, '');
                result.saved.forEach((b64, i) => fs.writeFileSync(path.join(outDir, `${base}.pass${i + 1}.docx`), Buffer.from(b64, 'base64')));
            }
        }
    } finally {
        await browser.close();
        server.close();
    }
    console.log(failed === 0 ? '\nDOCX ROUND TRIP: ALL PASSED' : `\nDOCX ROUND TRIP: ${failed} fixture(s) FAILED`);
    process.exit(failed === 0 || expectFailures ? 0 : 1);
}

main().catch((e) => { console.error(e); server.close(); process.exit(2); });
