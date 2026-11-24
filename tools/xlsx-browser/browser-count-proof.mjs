#!/usr/bin/env node
// Browser Fable proof: export a workbook's sheets with the Spreadsheet component's ExportPdf and
// write the PDF bytes plus a sheet->file map that measure-print-count-transitions.py understands.
//
//   node tools/xlsx-browser/browser-count-proof.mjs --dist DIR --fixture BOOK.xlsx --out DIR \
//        --fonts DIR [--label NAME] [--sheets 0,1]
//
// The page mounts the raw component (`/spreadsheet/Spreadsheet.js`), imports the workbook, hands it
// the PDF faces fetched from --fonts (family/weight/style names as their .ttf files), and exports
// each requested sheet. No product code is imported from the checkout; everything comes from DIST.
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const args = process.argv.slice(2);
const opt = (name, fallback) => { const i = args.indexOf(name); if (i < 0) return fallback; const v = args[i + 1]; args.splice(i, 2); return v; };
const dist = path.resolve(opt('--dist'));
const fixture = path.resolve(opt('--fixture'));
const outDir = path.resolve(opt('--out'));
const fontDir = path.resolve(opt('--fonts', path.join(here, '../../artifacts/fonts/windows')));
const label = opt('--label', 'browser');
const sheets = (opt('--sheets', '') || '').split(',').filter((s) => s !== '').map(Number);
if (!dist || !fixture || !outDir) { console.error('usage: --dist DIR --fixture BOOK.xlsx --out DIR [--fonts DIR] [--label NAME] [--sheets 0,1]'); process.exit(2); }
fs.mkdirSync(outDir, { recursive: true });

const FACE_FILES = [
    ['Calibri', false, false, 'calibri.ttf'], ['Calibri', true, false, 'calibrib.ttf'],
    ['Calibri', false, true, 'calibrii.ttf'], ['Calibri', true, true, 'calibriz.ttf'],
    ['Arial', false, false, 'arial.ttf'], ['Arial', true, false, 'arialbd.ttf'],
    ['Arial', false, true, 'ariali.ttf'], ['Arial', true, true, 'arialbi.ttf'],
    ['Verdana', false, false, 'verdana.ttf'], ['Verdana', true, false, 'verdanab.ttf'],
    ['Verdana', false, true, 'verdanai.ttf'], ['Verdana', true, true, 'verdanaz.ttf'],
    ['Georgia', false, false, 'georgia.ttf'], ['Georgia', true, false, 'georgiab.ttf'],
    ['Georgia', false, true, 'georgiai.ttf'], ['Georgia', true, true, 'georgiaz.ttf'],
    // Faces of the full-number cutoff anchors/holdout and the older overflow/accounting fixtures.
    ['Cambria', false, false, 'cambria.ttf'], ['Cambria', true, false, 'cambriab.ttf'],
    ['Cambria', false, true, 'cambriai.ttf'], ['Cambria', true, true, 'cambriaz.ttf'],
    ['Trebuchet MS', false, false, 'trebuc.ttf'], ['Trebuchet MS', true, false, 'trebucbd.ttf'],
    ['Trebuchet MS', false, true, 'trebucit.ttf'], ['Trebuchet MS', true, true, 'trebucbi.ttf'],
    ['Aptos Narrow', false, false, 'aptosnarrow-regular.ttf'],
    ['Aptos Narrow', true, false, 'aptosnarrow-bold.ttf'],
    ['Aptos Narrow', false, true, 'aptosnarrow-italic.ttf'],
];
const fonts = FACE_FILES.filter(([, , , file]) => fs.existsSync(path.join(fontDir, file)));
const fontByName = new Map(fonts.map((f) => [f[3], f]));

const fixtureBytes = fs.readFileSync(fixture);
const jszip = path.join(here, '../../node_modules/jszip/dist/jszip.min.js');
if (!fs.existsSync(jszip)) { console.error('jszip not found at ' + jszip); process.exit(2); }
const pageHtml = `<!doctype html><meta charset="utf-8"><body style="margin:0">
<script src="/jszip.min.js"></script>
<script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script>
<script type="module">
try {
    const { default: createSpreadsheet } = await import('/spreadsheet/Spreadsheet.js');
    const host = document.createElement('div');
    host.style.cssText = 'position:absolute;left:0;top:0;width:1100px;height:600px';
    document.body.appendChild(host);
    const grid = createSpreadsheet(host);
    grid.Mount();
    const bytes = new Uint8Array(await (await fetch('/fixture')).arrayBuffer());
    await grid.ImportXlsx(bytes);
    const faces = ${JSON.stringify(fonts.map(([family, bold, italic, file]) => ({ family, bold, italic, file })))};
    const pdfFonts = [];
    for (const face of faces) {
        const response = await fetch('/font?name=' + encodeURIComponent(face.file));
        if (!response.ok) continue;
        pdfFonts.push({ family: face.family, bold: face.bold, italic: face.italic,
                        data: new Uint8Array(await response.arrayBuffer()) });
    }
    grid.SetPdfFonts(pdfFonts);
    window.proof = {
        names: grid.GetSheetNames(),
        fonts: pdfFonts.length,
        exportSheet: async (index) => {
            const pdf = await grid.ExportPdf({ what: 'sheet', sheetIndex: index, fileName: 'browser-proof' });
            let binary = '';
            for (const b of pdf) binary += String.fromCharCode(b);
            return btoa(binary);
        },
    };
} catch (e) { window.loadError = String(e && e.stack || e); }
</script>`;

const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://x');
    const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
    if (url.pathname === '/') return send(200, 'text/html', pageHtml);
    if (url.pathname === '/fixture') return send(200, 'application/octet-stream', fixtureBytes);
    if (url.pathname === '/jszip.min.js') return send(200, 'text/javascript', fs.readFileSync(jszip));
    if (url.pathname === '/jszip-shim.js') return send(200, 'text/javascript', 'export default window.JSZip;');
    if (url.pathname === '/font') {
        const entry = fontByName.get(path.basename(url.searchParams.get('name') ?? ''));
        return entry ? send(200, 'application/octet-stream', fs.readFileSync(path.join(fontDir, entry[3]))) : send(404, 'text/plain', 'no font');
    }
    const file = path.join(dist, path.normalize(url.pathname));
    if (file.startsWith(dist) && fs.existsSync(file) && fs.statSync(file).isFile()) return send(200, 'text/javascript', fs.readFileSync(file));
    send(404, 'text/plain', 'not found');
});
await new Promise((resolve) => server.listen(Number(process.env.PORT ?? 16321), '127.0.0.1', resolve));
const playwright = await import(process.env.PLAYWRIGHT ? pathToFileURL(process.env.PLAYWRIGHT).href : 'playwright');
const browser = await playwright.chromium.launch();
let result;
try {
    const tab = await browser.newPage({ viewport: { width: 1200, height: 800 } });
    const errors = [];
    tab.on('pageerror', (e) => errors.push(e.message));
    await tab.goto(`http://127.0.0.1:${server.address().port}/`);
    await tab.waitForFunction('window.proof !== undefined || window.loadError !== undefined', null, { timeout: 180000 });
    const loadError = await tab.evaluate('window.loadError');
    if (loadError) throw new Error('page failed to load: ' + loadError);
    const names = await tab.evaluate('window.proof.names');
    const chosen = sheets.length ? sheets : names.map((_, i) => i);
    const made = [];
    for (const index of chosen) {
        const base64 = await tab.evaluate((i) => window.proof.exportSheet(i), index);
        const file = `${label}-sheet${index}.pdf`;
        fs.writeFileSync(path.join(outDir, file), Buffer.from(base64, 'base64'));
        made.push({ sheet: names[index], index, file });
    }
    result = { names, fonts: await tab.evaluate('window.proof.fonts'), made, errors };
} finally { await browser.close(); server.close(); }
const tsv = path.join(outDir, `${label}.tsv`);
fs.writeFileSync(tsv, result.made.map((m) => `${m.sheet}\t${m.index}\tpdf\t${m.file}`).join('\n') + '\n');
console.log(JSON.stringify({ dist, fixture, fonts: result.fonts, names: result.names, made: result.made, pageErrors: result.errors }, null, 2));
