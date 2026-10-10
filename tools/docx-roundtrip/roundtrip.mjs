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
// Every document is saved in two ways, and both are checked:
//
//   * as the editor saves it. What was not changed goes back as the text it
//     was read from, so for a document nobody edited word/document.xml, the
//     numbering, the notes, the comments, the styles, the settings and every
//     header / footer are the original's, character for character (a byte
//     order mark aside);
//   * from the model alone (package.RegenerateAll): what a block, a header or
//     a note that was changed is written as.
//
// For each of the two it asserts that
//   * what the reader sees after the first save equals what it saw in the
//     original (fidelity: sizes, indents, spacing, lists, tables, tabs, ...),
//   * the saved document is laid out as the original: the same number of
//     pages, the same laid-out blocks, each on the same page at the same y
//     with the same height (within half a pixel), line count and line
//     heights, and the same resolved formatting of every paragraph and run
//     (Word.getLayoutFingerprint),
//   * a second save reproduces the first one exactly (fixed point),
//   * every written XML part is well formed and the children of the
//     elements this writer produces are in schema (xsd:sequence) order,
//   * every relationship target is in the package, no relationship id is
//     used twice, and every r:id the XML uses is one the part has.
//
// Package integrity is asserted of every package the writer makes here - the
// unedited save, the save from the model alone, and each of the edited ones
// below: every relationship id any part's XML uses (r:id, r:embed, r:link,
// ..., o:relid) is one that part's relationships part has, every relationship
// target is in the package, every part has a content type and
// [Content_Types].xml names no part that is not there (what the original has
// wrong already is left out). An unedited save adds no part (no numbering
// part for a document without one), writes a header or footer under the name
// it had, and the summary says how many relationship parts and
// [Content_Types].xml files came back as the text they were read from.
//
// The document is also opened front first (officeFormats.word.openProgressive)
// and read to its end while an editor shows it, step by step and in the
// package's own slices: the model the editor exports, the layout fingerprint
// and every part of an unedited save are, byte for byte, those of the
// document read at once.
//
// Then the document is edited with real key events:
//   * one character typed and deleted again: the saved document reads back
//     and is laid out as the original;
//   * one character typed into a paragraph outside the tables and kept: the
//     saved document has it, its parts are well formed, and every other
//     paragraph of document.xml is still the text it was read from.
//
// And a character is typed at the start of every tenth paragraph of the
// body's text (real key events), and the document saved: it reads back with
// the text and is laid out as the editor that saved it, and document.xml
// differs from the original in those paragraphs only, in each by the
// character alone (one character more in a w:t, or one run more that holds
// it). The summary says how many paragraphs were written from the model
// instead (the XML could not be matched with the text the editor had); at
// least 95% must not be, and no paragraph nobody typed in may change (but in
// the documents of tenthOthersKnown below). --no-tenth skips it.
//
// Usage:
//   dotnet fable src/Office.Formats --outDir $DIST/office-formats --lang JavaScript --noCache
//   dotnet fable src/Word           --outDir $DIST/word           --lang JavaScript --noCache
//   node tools/docx-roundtrip/roundtrip.mjs --dist $DIST [--out DIR] [--verbose] [--synthetic] [--xml-summary] [--report-only] [fixture.docx ...]
//
// With no fixture arguments it runs the built-in synthetic documents (one per
// feature the writer has to preserve, fixtures.mjs) and every DOCX under
// tests/test-documents* (--synthetic: the synthetic ones only).
// --xml-summary prints how many documents come back with the document.xml
// they were read from, and, for the save from the model alone, which elements
// and attributes it has more or fewer of than the originals.
// --out keeps the documents: <name>.original.docx, .pass1.docx / .pass2.docx
// (the editor's first and second save) and .model.docx (from the model
// alone), so the .NET reader and layout can be pointed at them:
//   dotnet fsi tests/word/roundtrip-layout.fsx DIR     (laid out as the originals)
//   dotnet fsi tools/docx-roundtrip/native-check.fsx DIR   (pass 1 and 2 hold the same text)
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
const xmlSummary = flag('--xml-summary');
const syntheticOnly = flag('--synthetic');
const skipTenth = flag('--no-tenth');
// Every tenth paragraph (see the header): the share of the paragraphs typed in that must differ by the character alone,
// and the documents in which a paragraph nobody typed in may change (name fragments).
const tenthConfinedShare = 0.95;
const tenthOthersKnown = [
    'test-header-footer.docx', // w:sectPr is the first child of w:body: nothing of this body is written back as it was read (planBody)
];
const port = Number(process.env.PORT ?? 16021);

const firstExisting = (candidates) => candidates.find((p) => p && fs.existsSync(p));
const jszipDir = firstExisting([
    process.env.JSZIP_DIR,
    path.join(repo, 'node_modules/jszip'),
    path.resolve(repo, '../rbt-office/node_modules/jszip'),
]);
if (!jszipDir) { console.error('jszip not found; set JSZIP_DIR'); process.exit(2); }
const JSZip = createRequire(import.meta.url)(path.join(jszipDir, 'dist/jszip.min.js'));

// Every DOCX under tests/test-documents* (paths relative to the repository).
const corpusFiles = () => {
    const found = [];
    const walk = (dir) => { for (const entry of fs.readdirSync(dir, { withFileTypes: true })) { const full = path.join(dir, entry.name); if (entry.isDirectory()) walk(full); else if (entry.name.endsWith('.docx')) found.push(path.relative(repo, full)); } };
    for (const entry of fs.readdirSync(path.join(repo, 'tests'), { withFileTypes: true })) if (entry.isDirectory() && entry.name.startsWith('test-documents')) walk(path.join(repo, 'tests', entry.name));
    return found.sort();
};

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
    for (const required of ['office-formats/src/Main.js', 'word/Docx.js', 'word/Word.js'])
        if (!fs.existsSync(path.join(dist, required))) { console.error(`missing ${required} under ${dist}; Fable-compile first (see header)`); process.exit(2); }

    const expectations = new Map();
    if (args.length === 0) {
        for (const f of await syntheticFixtures(JSZip)) { fixtures.set(f.name, f.data); expectations.set(f.name, f.expect ?? []); }
        if (!syntheticOnly) for (const rel of corpusFiles()) fixtures.set(rel, fs.readFileSync(path.join(repo, rel)));
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
    let failed = 0, identicalXml = 0, compared = 0;
    const notIdentical = [];
    const frontFirst = { editors: 0, documents: 0 }; // documents an editor showed before they were read to their end
    const asRead = { total: 0, same: 0, notes: [] }; // relationship parts and [Content_Types].xml of unedited saves
    const tenthTotals = { edited: 0, confined: 0, rewritten: 0, others: 0 };
    const tenthNotes = [];
    const xmlKinds = new Map(); // element / attribute kind -> how often it was in the originals and in the saved documents
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
            // The trivial edit: one character typed and deleted with real key events.
            if (await tab.evaluate('window.editOpen()')) {
                await tab.keyboard.type('x');
                await tab.keyboard.press('Backspace');
                problems.push(...await tab.evaluate('window.editFinish()'));
            }
            // A real edit: one character typed into a paragraph outside the tables.
            if (await tab.evaluate('window.editOpen(true)')) {
                await tab.keyboard.type('Q');
                problems.push(...await tab.evaluate(([n, expect]) => window.editKeptFinish(n, expect), [name, expectations.get(name) ?? []]));
            }
            // A character typed at the start of every tenth paragraph.
            if (!skipTenth) {
                const count = await tab.evaluate((n) => window.tenthOpen(n), name);
                for (let i = 0; i < count; i++) { await tab.evaluate((index) => window.tenthSelect(index), i); await tab.keyboard.type('Q'); }
                if (count > 0) {
                    const tenth = await tab.evaluate('window.tenthFinish()');
                    problems.push(...tenth.problems);
                    tenthTotals.edited += tenth.edited; tenthTotals.confined += tenth.confined; tenthTotals.rewritten += tenth.rewritten; tenthTotals.others += tenth.others;
                    if (tenth.rewritten || tenth.others) {
                        tenthNotes.push(`${name}: ${tenth.edited} typed in, ${tenth.confined} changed by the character alone, ${tenth.rewritten} written from the model, ${tenth.others} others changed`);
                        if (verbose || process.env.TENTH_WHY) for (const line of tenth.why) tenthNotes.push('      ' + line);
                    }
                    // Paragraphs nobody typed in stay the text they were read from - but for the documents listed in the header.
                    if (tenth.others > 0 && !tenthOthersKnown.some((known) => name.includes(known))) problems.push(`every tenth paragraph: ${tenth.others} paragraphs nobody typed in are not the text they were read from`);
                }
            }
            for (const [kind, [before, after]] of Object.entries(result.xmlDiff ?? {})) {
                const entry = xmlKinds.get(kind) ?? { files: 0, before: 0, after: 0 };
                entry.files++; entry.before += before; entry.after += after; xmlKinds.set(kind, entry);
            }
            if (result.xmlIdentical) identicalXml++; else notIdentical.push(name);
            frontFirst.editors += result.frontFirst ?? 0; if (result.frontFirst) frontFirst.documents++;
            asRead.total += result.partsAsRead?.total ?? 0; asRead.same += result.partsAsRead?.same ?? 0;
            if (result.partsAsRead?.changed.length) asRead.notes.push(`${name}: ${result.partsAsRead.changed.join(', ')}`);
            compared++;
            // An unedited save writes the parts it read back as they were.
            if (!expectations.get(name)?.some((e) => e.mutate) && !result.problems.some((p) => p.startsWith('threw'))) {
                for (const part of result.changedParts ?? []) problems.push(`unedited save: ${part} is not the text that was read`);
            }
            console.log(`${problems.length === 0 ? 'ok  ' : 'FAIL'} ${name}  (${result.stats})`);
            const limit = verbose ? 200 : 12;
            for (const p of problems.slice(0, limit)) console.log('     ' + p);
            if (problems.length > limit) console.log(`     ... ${problems.length - limit} more`);
            if (verbose && result.xml) console.log(result.xml);
            if (problems.length) failed++;
            if (outDir && result.saved) {
                const base = path.basename(name).replace(/\.docx$/, '');
                result.saved.forEach((b64, i) => fs.writeFileSync(path.join(outDir, `${base}.${['pass1', 'pass2', 'model'][i]}.docx`), Buffer.from(b64, 'base64')));
                if (fixtures.get(name).length && !(expectations.get(name) ?? []).some((e) => e.mutate)) fs.writeFileSync(path.join(outDir, `${base}.original.docx`), fixtures.get(name));
            }
        }
    } finally {
        await browser.close();
        server.close();
    }
    if (xmlSummary) {
        console.log(`\nword/document.xml of an unedited save: identical to the original in ${identicalXml} of ${compared} documents.`);
        if (notIdentical.length) console.log('  not identical: ' + notIdentical.join(', '));
        console.log('Written from the model alone (what a changed block is written as), against the original:');
        const rows = [...xmlKinds.entries()].sort((a, b) => b[1].files - a[1].files || a[0].localeCompare(b[0]));
        if (rows.length) console.log('Elements and attributes whose number changed (kind: documents, occurrences before -> after):');
        for (const [kind, e] of rows) console.log(`  ${kind}: ${e.files} documents, ${e.before} -> ${e.after}`);
    }
    console.log(`\nOpened front first: ${frontFirst.documents} of ${compared} documents were shown by an editor before they were read to their end (the others are read at once: short, or a main part that cannot be read front first); each is, read to its end, the document read at once.`);
    if (asRead.total) {
        console.log(`\nRelationship parts and [Content_Types].xml of an unedited save: ${asRead.same} of ${asRead.total} are the text that was read.`);
        if (xmlSummary || verbose) for (const line of asRead.notes) console.log('  not as read: ' + line);
    }
    if (!skipTenth) {
        console.log(`\nA character typed at the start of every tenth paragraph: ${tenthTotals.edited} paragraphs; ${tenthTotals.confined} differ from the original by the character alone, ${tenthTotals.rewritten} were written from the model; ${tenthTotals.others} other paragraphs changed.`);
        for (const line of tenthNotes) console.log('  ' + line);
        if (tenthTotals.confined < tenthTotals.edited * tenthConfinedShare) { console.log(`  FAIL: fewer than ${tenthConfinedShare * 100}% by the character alone`); failed++; }
    }
    console.log(failed === 0 ? '\nDOCX ROUND TRIP: ALL PASSED' : `\nDOCX ROUND TRIP: ${failed} fixture(s) FAILED`);
    process.exit(failed === 0 || expectFailures ? 0 : 1);
}

main().catch((e) => { console.error(e); server.close(); process.exit(2); });
