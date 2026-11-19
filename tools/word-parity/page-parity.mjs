#!/usr/bin/env node
// Page parity against Microsoft Word: for each corpus document, our page count
// and the first and last line of every page against Word's.
//
//   node tools/word-parity/page-parity.mjs --dist dist [--refresh-word] [--json out.json] [--detail] [name-filter ...]
//   node tools/word-parity/page-parity.mjs --dist dist --dir <folder> [--refresh-word]   (the .docx of a folder, Word's TSVs in <folder>/word)
//   node tools/word-parity/page-parity.mjs --dist dist --embedded-fonts ...   (also the fonts a document embeds, as a host that loads them would)
//   node tools/word-parity/page-parity.mjs --dist dist --lines <name>   (our lines of one document: page, y on the page in px, height, block, text)
//
// Word's side is tests/word/word-pages/<dir>__<name>.tsv (committed), written by
// tools/word-parity/page-lines.vbs on the Windows VM: Word 16 / 2019 opens each
// document, repaginates, and records ComputeStatistics(wdStatisticPages) and, per
// page, the character position it starts at and the text of its first and last
// line, and of its first and last line with a letter or number (what is compared). --refresh-word stages tests/test-documents*/*.docx, runs that script on
// $WORD_VM (default winvm) in $WORD_VM_DIR/pageparity (direct cscript over ssh),
// fetches the TSVs back and also Word's PDFs into artifacts/word-pages/.
//
// Our side is the browser layout of the Fable output in $DIST (Chromium through
// $PLAYWRIGHT), with Word's own fonts registered as web fonts: every .ttf / .ttc
// of $RB_WORD_FONT_DIR (default artifacts/fonts/windows) and of
// $RB_WORD_EXTRA_FONT_DIR (default windows-extra beside it: Consolas, Courier New, Tahoma, Georgia ... -
// tools/word-parity/fetch-windows-fonts.sh and the README), each under the family
// and style its name table gives. Lines are Word.getLayoutLines' body lines (not
// headers, footers or notes), in document order, per page; the first and last
// with a letter or number are compared.
//
// A page "matches" when its first line is Word's first line of that page (the
// texts compared with case, spaces and punctuation taken out, on the shorter's
// length, at most 40 characters). The table: Word's pages, ours, the pages of
// Word whose first line is ours on the same page, and the first page that
// differs with both first lines. --detail prints every page.
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.resolve(here, '../..');
const args = process.argv.slice(2);
const opt = (name) => { const i = args.indexOf(name); if (i < 0) return undefined; const v = args[i + 1]; args.splice(i, 2); return v; };
const flag = (name) => { const i = args.indexOf(name); if (i < 0) return false; args.splice(i, 1); return true; };
const dist = path.resolve(opt('--dist') ?? path.join(repo, 'dist'));
const jsonOut = opt('--json');
const refresh = flag('--refresh-word');
const detail = flag('--detail');
const linesFor = opt('--lines');
const docDir = opt('--dir');
const embedFonts = flag('--embedded-fonts');
const filters = args;
const port = Number(process.env.PORT ?? 16031);
const refDir = docDir ? path.join(path.resolve(docDir), 'word') : path.join(repo, 'tests/word/word-pages');
const fontDir = path.resolve(process.env.RB_WORD_FONT_DIR ?? path.join(repo, 'artifacts/fonts/windows'));
const fontDirs = [fontDir, path.resolve(process.env.RB_WORD_EXTRA_FONT_DIR ?? path.join(path.dirname(fontDir), 'windows-extra'))]
    .filter((d) => fs.existsSync(d));

// The corpus: every .docx of tests/test-documents*/, keyed <dir>__<name> (with --dir, that folder's, keyed <name>).
const corpus = [];
if (docDir)
    for (const file of fs.readdirSync(docDir).filter((f) => f.endsWith('.docx')).sort())
        corpus.push({ key: file.slice(0, -5), file: path.resolve(docDir, file) });
else
    for (const dir of fs.readdirSync(path.join(repo, 'tests')).filter((d) => d.startsWith('test-documents')).sort())
        for (const file of fs.readdirSync(path.join(repo, 'tests', dir)).filter((f) => f.endsWith('.docx')).sort())
            corpus.push({ key: `${dir}__${file.slice(0, -5)}`, file: path.join(repo, 'tests', dir, file) });
const selected = corpus.filter((d) => !filters.length || filters.some((f) => d.key.includes(f)));

if (refresh) {
    const vm = process.env.WORD_VM ?? 'winvm';
    const remote = `${process.env.WORD_VM_DIR ?? 'C:/Users/olivia/RbtExcelParity'}/${docDir ? 'pageparity-dir' : 'pageparity'}`;
    const stage = fs.mkdtempSync(path.join(process.env.TMPDIR ?? '/tmp', 'page-parity-'));
    for (const d of selected) fs.copyFileSync(d.file, path.join(stage, d.key + '.docx'));
    const run = (cmd, a) => { const r = spawnSync(cmd, a, { stdio: ['ignore', 'inherit', 'inherit'] }); if (r.status !== 0) throw new Error(`${cmd} ${a.join(' ')} failed`); };
    run('ssh', [vm, `powershell -NoProfile -Command "Remove-Item -Recurse -Force -ErrorAction SilentlyContinue '${remote}'; New-Item -ItemType Directory -Force -Path '${remote}/docs' | Out-Null"`]);
    run('scp', ['-q', path.join(here, 'page-lines.vbs'), `${vm}:${remote}/`]);
    run('scp', ['-q', ...selected.map((d) => path.join(stage, d.key + '.docx')), `${vm}:${remote}/docs/`]);
    run('ssh', [vm, `cscript //nologo ${remote}/page-lines.vbs ${remote}/docs ${remote}/out pdf`]);
    const pdfDir = docDir ? refDir : path.join(repo, 'artifacts/word-pages');
    fs.mkdirSync(pdfDir, { recursive: true });
    run('scp', ['-q', `${vm}:${remote}/out/*`, pdfDir]);
    fs.mkdirSync(refDir, { recursive: true });
    for (const d of selected) {
        const tsv = path.join(pdfDir, d.key + '.tsv');
        if (!fs.existsSync(tsv)) { console.error(`no TSV from Word for ${d.key}`); continue; }
        const raw = fs.readFileSync(tsv);
        const text = raw[0] === 0xff && raw[1] === 0xfe ? raw.subarray(2).toString('utf16le') : raw.toString('utf8');
        fs.writeFileSync(path.join(refDir, d.key + '.tsv'), text.replace(/\r\n/g, '\n'));
    }
    fs.rmSync(stage, { recursive: true, force: true });
}

const readWord = (key) => {
    const file = path.join(refDir, key + '.tsv');
    if (!fs.existsSync(file)) return null;
    const pages = [];
    let count = 0;
    for (const line of fs.readFileSync(file, 'utf8').split('\n')) {
        const cells = line.split('\t');
        if (cells[0] === 'pages') count = Number(cells[1]);
        else if (/^\d+$/.test(cells[0])) pages[Number(cells[0]) - 1] = { start: Number(cells[1]), first: cells[4] ?? cells[2] ?? '', last: cells[5] ?? cells[3] ?? '' };
    }
    return { count, pages };
};

// Family, weight and style of a TrueType / OpenType file (its first face) from its name and OS/2 tables.
const faceOf = (file) => {
    const b = fs.readFileSync(file);
    let base = 0;
    if (b.toString('latin1', 0, 4) === 'ttcf') base = b.readUInt32BE(12);
    const numTables = b.readUInt16BE(base + 4);
    const tables = {};
    for (let i = 0; i < numTables; i++) { const o = base + 12 + i * 16; tables[b.toString('latin1', o, o + 4)] = b.readUInt32BE(o + 8); }
    const names = {};
    const n = tables.name;
    if (n === undefined) return null;
    const count = b.readUInt16BE(n + 2), strings = n + b.readUInt16BE(n + 4);
    for (let i = 0; i < count; i++) {
        const r = n + 6 + i * 12;
        const platform = b.readUInt16BE(r), lang = b.readUInt16BE(r + 4), id = b.readUInt16BE(r + 6), len = b.readUInt16BE(r + 8), off = b.readUInt16BE(r + 10);
        if (platform === 3 && (lang === 0x409 || names[id] === undefined)) {
            const s = b.subarray(strings + off, strings + off + len);
            let t = ''; for (let k = 0; k + 1 < s.length; k += 2) t += String.fromCharCode(s.readUInt16BE(k));
            names[id] = t;
        }
    }
    const os2 = tables['OS/2'];
    const weight = os2 !== undefined ? b.readUInt16BE(os2 + 4) : 400;
    const fsSelection = os2 !== undefined ? b.readUInt16BE(os2 + 62) : 0;
    const sub = (names[2] ?? '').toLowerCase();
    return { family: names[1], weight: weight >= 600 || sub.includes('bold') ? 'bold' : 'normal', style: (fsSelection & 1) || sub.includes('italic') ? 'italic' : 'normal' };
};
const faces = [];
for (const dir of fontDirs)
    for (const file of fs.readdirSync(dir).filter((f) => /\.(ttf|ttc|otf)$/i.test(f)).sort()) {
        try { const f = faceOf(path.join(dir, file)); if (f?.family) faces.push({ ...f, url: `/font/${encodeURIComponent(path.basename(dir))}/${encodeURIComponent(file)}` }); } catch { }
    }

if (!faces.length) { console.error(`no fonts in ${fontDir} (RB_WORD_FONT_DIR; tools/word-parity/fetch-windows-fonts.sh)`); process.exit(2); }
console.error(`${faces.length} faces from ${fontDirs.join(', ')}`);
const pageHtml = `<!doctype html><meta charset="utf-8"><body style="margin:0">
<script src="/jszip.min.js"></script><script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script>
<script type="module">
for (const f of ${JSON.stringify(faces)}) {
    try { const face = new FontFace(f.family, 'url(' + f.url + ')', { weight: f.weight, style: f.style }); await face.load(); document.fonts.add(face); } catch { }
}
const Word = (await import('/word/Word.js')).default;
const main = await import('/office-formats/src/Main.js');
const OF = typeof main.default === 'function' ? main.default() : main.default;
// The fonts a document embeds (word/fonts/*.odttf, obfuscated with the key fontTable.xml gives), as web fonts.
const embedded = async (bytes) => {
    if (!window.embedFonts) return [];
    const zip = await window.JSZip.loadAsync(bytes);
    const table = await zip.file('word/fontTable.xml')?.async('string');
    const rels = await zip.file('word/_rels/fontTable.xml.rels')?.async('string');
    if (!table || !rels) return [];
    const targets = {};
    for (const m of rels.matchAll(/<Relationship [^>]*>/g)) { const id = /Id="([^"]+)"/.exec(m[0])?.[1], t = /Target="([^"]+)"/.exec(m[0])?.[1]; if (id && t) targets[id] = t; }
    const added = [];
    for (const font of table.matchAll(/<w:font w:name="([^"]+)"[^>]*>([\\s\\S]*?)<\\/w:font>/g)) {
        for (const e of font[2].matchAll(/<w:embed(Regular|Bold|Italic|BoldItalic) r:id="([^"]+)" w:fontKey="\\{([^}]+)\\}"/g)) {
            const file = zip.file('word/' + targets[e[2]]);
            if (!file) continue;
            const data = new Uint8Array(await file.async('uint8array'));
            const hex = e[3].replace(/-/g, '');
            const key = []; for (let i = 0; i < 32; i += 2) key.push(parseInt(hex.slice(i, i + 2), 16));
            key.reverse();
            for (let i = 0; i < 32; i++) data[i] ^= key[i % 16];
            const face = new FontFace(font[1], data, { weight: e[1].startsWith('Bold') ? 'bold' : 'normal', style: e[1].endsWith('Italic') ? 'italic' : 'normal' });
            try { await face.load(); document.fonts.add(face); added.push(face); } catch { }
        }
    }
    return added;
};
const editorOf = async (url) => {
    const bytes = new Uint8Array(await (await fetch(url)).arrayBuffer());
    const added = await embedded(bytes);
    const editor = Word.initializeCanvas(document.createElement('canvas'));
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
    Word.completeLayout(editor);
    return { editor, done: () => { for (const f of added) document.fonts.delete(f); } };
};
window.layoutOf = async (url) => {
    const { editor, done } = await editorOf(url);
    const count = Word.getPageCount(editor);
    const pages = [];
    for (const line of Word.getLayoutLines(editor)) {
        if (String(line.BlockId).startsWith('__')) continue;
        const p = (pages[line.Page - 1] ??= { lines: 0, first: '', last: '' });
        p.lines++;
        if (!/[\\p{L}\\p{N}]/u.test(line.Text)) continue;
        if (p.first === '') p.first = line.Text;
        p.last = line.Text;
    }
    done();
    return { count, pages: Array.from({ length: Math.max(count, pages.length) }, (_, i) => pages[i] ?? { lines: 0, first: '', last: '' }) };
};
window.linesOf = async (url) => {
    const { editor } = await editorOf(url);
    return Word.getLayoutLines(editor).map((l) => [l.Page, Math.round(l.PageY * 100) / 100, Math.round(l.Height * 100) / 100, l.BlockId + "@" + Math.round(l.BlockHeight * 75) / 100, l.Text, (l.Segments ?? []).map((s) => s.Text + "=" + Math.round(s.Width * 75) / 100).join(" ").slice(0, 120)]);
};
window.embedFonts = ${embedFonts};
window.ready = true;
</script>`;

const jszipDir = [process.env.JSZIP_DIR, path.join(repo, 'node_modules/jszip')].find((p) => p && fs.existsSync(p));
if (!jszipDir) { console.error('jszip not found (pnpm install, or JSZIP_DIR)'); process.exit(2); }
if (!fs.existsSync(path.join(dist, 'word/Word.js'))) { console.error(`no word/Word.js under ${dist} (pnpm run build:fable)`); process.exit(2); }
const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://x');
    const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
    if (url.pathname === '/') return send(200, 'text/html', pageHtml);
    if (url.pathname === '/jszip.min.js') return send(200, 'text/javascript', fs.readFileSync(path.join(jszipDir, 'dist/jszip.min.js')));
    if (url.pathname === '/jszip-shim.js') return send(200, 'text/javascript', 'export default window.JSZip;');
    const fontMatch = /^\/font\/([^/]+)\/([^/]+)$/.exec(url.pathname);
    if (fontMatch) {
        const dir = fontDirs.find((d) => path.basename(d) === decodeURIComponent(fontMatch[1]));
        const file = dir && path.join(dir, path.basename(decodeURIComponent(fontMatch[2])));
        return file && fs.existsSync(file) ? send(200, 'font/ttf', fs.readFileSync(file)) : send(404, 'text/plain', 'no font');
    }
    if (url.pathname === '/doc') { const d = corpus[Number(url.searchParams.get('i'))]; return d ? send(200, 'application/octet-stream', fs.readFileSync(d.file)) : send(404, 'text/plain', 'no document'); }
    for (const [prefix, root] of [['/', dist]]) {
        if (!url.pathname.startsWith(prefix)) continue;
        const file = path.join(root, path.normalize(decodeURIComponent(url.pathname.slice(prefix.length))));
        if (file.startsWith(root) && fs.existsSync(file) && fs.statSync(file).isFile())
            return send(200, /\.m?js$/.test(file) ? 'text/javascript' : 'application/octet-stream', fs.readFileSync(file));
    }
    send(404, 'text/plain', 'not found');
});
await new Promise((resolve) => server.listen(port, '127.0.0.1', resolve));
const playwright = await import(process.env.PLAYWRIGHT ? pathToFileURL(process.env.PLAYWRIGHT).href : 'playwright');
const browser = await (playwright.chromium ?? playwright.default.chromium).launch();

const norm = (s) => String(s ?? '').toLowerCase().normalize('NFKD').replace(/[^\p{L}\p{N}]/gu, '');
const same = (a, b) => { const x = norm(a), y = norm(b); const n = Math.min(40, x.length, y.length); return n === 0 ? x.length === y.length : x.slice(0, n) === y.slice(0, n); };
const cut = (s, n = 48) => { s = String(s ?? ''); return s.length > n ? s.slice(0, n - 1) + '…' : s; };

const rows = [];
try {
    const tab = await browser.newPage();
    tab.on('pageerror', (e) => console.error('page error:', e.message));
    await tab.goto(`http://127.0.0.1:${port}/`);
    await tab.waitForFunction('window.ready === true', null, { timeout: 120000 });
    if (linesFor) {
        const d = corpus.find((c) => c.key.includes(linesFor));
        for (const l of await tab.evaluate((u) => window.linesOf(u), '/doc?i=' + corpus.indexOf(d))) console.log(l.join('\t'));
        process.exit(0);
    }
    for (const d of selected) {
        const word = readWord(d.key);
        if (!word) { console.error(`${d.key}: no Word reference (--refresh-word)`); continue; }
        let ours;
        try { ours = await tab.evaluate((u) => window.layoutOf(u), '/doc?i=' + corpus.indexOf(d)); }
        catch (e) { rows.push({ key: d.key, word: word.count, ours: null, error: String(e.message).split('\n')[0] }); continue; }
        const pages = [];
        let matched = 0, firstDiff = null;
        for (let i = 0; i < word.count; i++) {
            const w = word.pages[i] ?? { first: '', last: '' }, o = ours.pages[i] ?? { first: '', last: '' };
            const ok = i < ours.count && same(w.first, o.first);
            if (ok) matched++; else if (firstDiff === null) firstDiff = { page: i + 1, word: w.first, ours: o.first };
            pages.push({ page: i + 1, ok, lastOk: same(w.last, o.last), wordFirst: w.first, oursFirst: o.first, wordLast: w.last, oursLast: o.last });
        }
        rows.push({ key: d.key, word: word.count, ours: ours.count, matched, firstDiff, pages });
    }
} finally { await browser.close(); server.close(); }

const short = (k) => k.replace(/^test-documents-?/, '').replace('__', '/');
console.log(`${'document'.padEnd(52)} ${'Word'.padStart(5)} ${'ours'.padStart(5)} ${'starts'.padStart(9)}  first page that differs`);
let samePages = 0, allStarts = 0;
for (const r of rows) {
    if (r.ours === null) { console.log(`${short(r.key).padEnd(52)} ${String(r.word).padStart(5)}  error ${r.error}`); continue; }
    if (r.ours === r.word) samePages++;
    if (r.ours === r.word && r.matched === r.word) allStarts++;
    const diff = r.firstDiff ? `p${r.firstDiff.page}: Word "${cut(r.firstDiff.word, 30)}" / ours "${cut(r.firstDiff.ours, 30)}"` : '';
    console.log(`${short(r.key).padEnd(52)} ${String(r.word).padStart(5)} ${String(r.ours).padStart(5)} ${`${r.matched}/${r.word}`.padStart(9)}  ${r.ours === r.word ? '' : '*'}${diff}`);
    if (detail) for (const p of r.pages) if (!p.ok || !p.lastOk)
        console.log(`    p${p.page} first ${p.ok ? '=' : '≠'} W "${cut(p.wordFirst)}" O "${cut(p.oursFirst)}" | last ${p.lastOk ? '=' : '≠'} W "${cut(p.wordLast)}" O "${cut(p.oursLast)}"`);
}
console.log(`\n${samePages} of ${rows.length} documents at Word's page count; ${allStarts} with every page starting at Word's line.`);
if (jsonOut) fs.writeFileSync(jsonOut, JSON.stringify(rows, null, 1));
