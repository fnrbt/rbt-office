// Screenshots of the browser grid, for comparing with Excel's (shoot-sheets.ps1 or export-ranges.vbs):
//   node tools/excel-parity/shoot-browser-sheets.mjs --dist dist OUTDIR 'tests/x.xlsx|Sheet|A1' ...
// Each workbook is imported into the web component (as tools/xlsx-browser/run.mjs serves it), the sheet
// shown and scrolled so that the cell is at the top left, every drawing waited for, and the grid's
// content (without its headers) written to OUTDIR/<workbook>-<sheet>.png. PORT, PLAYWRIGHT as run.mjs.
import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

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
const page = `<!doctype html><meta charset="utf-8"><body style="margin:0"><script src="/jszip.min.js"></script>
<script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script><script type="module">
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
        const at = host.querySelector('.spreadsheet-content').getBoundingClientRect();
        const target = host.querySelector('.spreadsheet-cell[data-row="' + row + '"][data-col="' + col + '"]');
        if (target) { const r = target.getBoundingClientRect(); scroller.scrollTop += r.top - at.top; scroller.scrollLeft += r.left - at.left; }
        grid.SelectCell(1048575, 16383);
        await frame(); await everyObjectSeen(host); await frame();
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
        const out = path.join(outDir, `${path.basename(file, '.xlsx')}-${sheet.replace(/\W+/g, '_')}-${cell}.png`);
        fs.writeFileSync(out, await tab.screenshot({ clip }));
        console.log(out);
    }
} finally { await browser.close(); server.close(); }
