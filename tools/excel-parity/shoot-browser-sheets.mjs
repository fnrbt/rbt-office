// Screenshots of the browser grid, for comparing with Excel's (shoot-sheets.ps1 or export-ranges.vbs):
//   node tools/excel-parity/shoot-browser-sheets.mjs --dist dist OUTDIR 'tests/x.xlsx|Sheet|A1' ...
// Each workbook is imported into the web component (as tools/xlsx-browser/run.mjs serves it), the sheet
// shown and scrolled so that the cell is at the top left, every drawing waited for, and the grid's
// content (without its headers) written to OUTDIR/<workbook>-<sheet>.png. PORT, PLAYWRIGHT as run.mjs.
import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { facesIn } from './windows-faces.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.resolve(here, '../..');
const harness = path.join(repo, 'tools/xlsx-browser');
const args = process.argv.slice(2);
const opt = (name) => { const i = args.indexOf(name); if (i < 0) return undefined; const v = args[i + 1]; args.splice(i, 2); return v; };
const dist = path.resolve(opt('--dist') ?? path.join(repo, 'dist'));
const width = Number(opt('--width') ?? 1600), height = Number(opt('--height') ?? 1000);
const [outDir, ...shots] = args;
if (!outDir || shots.length === 0) { console.error('usage: shoot-browser-sheets.mjs [--dist DIR] OUTDIR "workbook|sheet|A1" ...'); process.exit(2); }
fs.mkdirSync(outDir, { recursive: true });
const jszip = path.join(repo, 'node_modules/jszip/dist/jszip.min.js');
const port = Number(process.env.PORT ?? 16201);
// The Windows fonts, registered as web fonts before the workbook is shown (RB_WORD_FONT_DIR, RB_WORD_EXTRA_FONT_DIR):
// without them the grid measures and draws Calibri, Tahoma ... in the browser's fallback.
const fontDir = path.resolve(process.env.RB_WORD_FONT_DIR ?? path.join(repo, 'artifacts/fonts/windows'));
const fontDirs = [fontDir, path.resolve(process.env.RB_WORD_EXTRA_FONT_DIR ?? path.join(path.dirname(fontDir), 'windows-extra'))];
const faces = facesIn(fontDirs).map((f, i) => ({ family: f.family, weight: f.weight, style: f.style, url: `/font/${i}`, path: path.join(f.dir, f.file) }));
console.error(`${faces.length} faces from ${fontDirs.join(', ')}`);
const page = `<!doctype html><meta charset="utf-8"><body style="margin:0"><script src="/jszip.min.js"></script>
<script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script><script type="module">
for (const f of ${JSON.stringify(faces.map(({ path: _, ...f }) => f))}) {
    try { const face = new FontFace(f.family, 'url(' + f.url + ')', { weight: f.weight, style: f.style }); await face.load(); document.fonts.add(face); } catch { }
}
try {
    const { createSpreadsheet, container, frame, everyObjectSeen, bytesOf } = await import('/component.js');
    window.show = async (file, sheet, cell) => {
        document.body.innerHTML = '';
        const host = container('position:absolute;left:0;top:0;width:${width}px;height:${height}px');
        const grid = createSpreadsheet(host);
        grid.Mount();
        await grid.ImportXlsx((await bytesOf(file)).slice());
        grid.SwitchSheet(grid.GetSheetNames().indexOf(sheet)); await frame();
        await everyObjectSeen(host);
        const scroller = host.querySelector('.spreadsheet-container');
        scroller.scrollTop = 0; scroller.scrollLeft = 0; await frame();
        const m = /^([A-Z]+)([0-9]+)$/.exec(cell);
        const col = [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1, row = Number(m[2]) - 1;
        // (The cell is selected, which scrolls it into view, and the grid scrolled on until it is at the top
        // left of the scrolling pane - below the column headers, right of the row headers. An empty cell has no
        // element: the active-cell frame gives its place.)
        const visible = (sel) => [...host.querySelectorAll(sel)].map((e) => e.getBoundingClientRect()).find((r) => r.width > 0 && r.height > 0);
        grid.SelectCell(row, col); await frame();
        // (At a zoom the grid's scroll offsets are not screen pixels: each step learns how far the cell moves.)
        let kTop = 1, kLeft = 1;
        for (let i = 0; i < 8; i++) {
            const active = visible('.spreadsheet-active-cell');
            if (!active) break;
            const box = scroller.getBoundingClientRect();
            const colHeader = visible('.spreadsheet-col-header'), rowHeader = visible('.spreadsheet-row-header');
            const top = colHeader ? colHeader.bottom : box.top + scroller.clientTop, left = rowHeader ? rowHeader.right : box.left + scroller.clientLeft;
            const dy = active.top - top, dx = active.left - left;
            if (Math.abs(dy) < 0.5 && Math.abs(dx) < 0.5) break;
            const y0 = scroller.scrollTop, x0 = scroller.scrollLeft;
            scroller.scrollTop += dy * kTop; scroller.scrollLeft += dx * kLeft;
            await frame();
            const moved = visible('.spreadsheet-active-cell');
            if (!moved) break;
            if (Math.abs(active.top - moved.top) > 0.5) kTop = (scroller.scrollTop - y0) / (active.top - moved.top);
            if (Math.abs(active.left - moved.left) > 0.5) kLeft = (scroller.scrollLeft - x0) / (active.left - moved.left);
        }
        // (Selecting a cell out of sight would scroll the grid away - and, far down, not back to the same rows:
        // the selection is hidden instead.)
        const hide = document.createElement('style');
        hide.textContent = '.spreadsheet-selection,.spreadsheet-active-cell,.spreadsheet-fill-handle{display:none !important}';
        host.appendChild(hide);
        // (Objects in view are laid out as they come into it; everyObjectSeen would scroll the grid back to its top.)
        const view = scroller.getBoundingClientRect();
        const pendingInView = () => [...host.querySelectorAll('.spreadsheet-object-pending')].map((e) => e.getBoundingClientRect())
            .some((r) => r.bottom > view.top && r.top < view.bottom && r.right > view.left && r.left < view.right);
        for (let i = 0; i < 60 && pendingInView(); i++) await frame();
        await frame();
        const box = host.querySelector('.spreadsheet-container').getBoundingClientRect();
        return { x: box.left, y: box.top, width: box.width, height: box.height };
    };
    window.ready = true;
} catch (e) { window.loadError = String(e && e.stack || e); }
</script>`;
const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://x');
    const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
    if (url.pathname === '/') return send(200, 'text/html', page);
    if (url.pathname === '/jszip.min.js') return send(200, 'text/javascript', fs.readFileSync(jszip));
    if (url.pathname === '/jszip-shim.js') return send(200, 'text/javascript', 'export default window.JSZip;');
    if (/^\/[\w-]+\.js$/.test(url.pathname) && fs.existsSync(path.join(harness, url.pathname.slice(1)))) return send(200, 'text/javascript', fs.readFileSync(path.join(harness, url.pathname.slice(1))));
    if (url.pathname.startsWith('/font/')) { const f = faces[Number(url.pathname.slice(6))]; return f ? send(200, 'font/ttf', fs.readFileSync(f.path)) : send(404, 'text/plain', 'no font'); }
    if (url.pathname === '/fixture') { const file = path.join(repo, path.normalize(url.searchParams.get('name'))); return file.startsWith(repo) && fs.existsSync(file) ? send(200, 'application/octet-stream', fs.readFileSync(file)) : send(404, 'text/plain', 'no fixture'); }
    const file = path.join(dist, path.normalize(url.pathname));
    if (file.startsWith(dist) && fs.existsSync(file) && fs.statSync(file).isFile()) return send(200, 'text/javascript', fs.readFileSync(file));
    send(404, 'text/plain', 'not found');
});
await new Promise((resolve) => server.listen(port, '127.0.0.1', resolve));
const playwright = await import(process.env.PLAYWRIGHT ? pathToFileURL(process.env.PLAYWRIGHT).href : 'playwright');
const browser = await (playwright.chromium ?? playwright.default.chromium).launch();
try {
    const tab = await browser.newPage({ viewport: { width, height } });
    tab.on('pageerror', (e) => console.error('page error:', e.message));
    await tab.goto(`http://127.0.0.1:${port}/`);
    await tab.waitForFunction('window.ready !== undefined || window.loadError !== undefined');
    const loadError = await tab.evaluate('window.loadError');
    if (loadError) throw new Error(loadError);
    for (const shot of shots) {
        const [file, sheet, cell = 'A1'] = shot.split('|');
        const clip = await tab.evaluate(([f, s, c]) => window.show(f, s, c), [file, sheet, cell]);
        const out = path.join(outDir, `${path.basename(file).replace(/\.xlsx$/, '')}-${sheet.replace(/\W+/g, '_')}-${cell}.png`);
        fs.writeFileSync(out, await tab.screenshot({ clip }));
        console.log(out);
    }
} finally { await browser.close(); server.close(); }
