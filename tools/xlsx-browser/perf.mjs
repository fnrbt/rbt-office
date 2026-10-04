#!/usr/bin/env node
// Timings of the Spreadsheet web component in a headless browser, against
// Fable output: opening a workbook (ImportXlsx), writing it (ExportXlsx),
// reading the written file again, and frame times while scrolling a sheet.
//
// Usage:
//   node tools/xlsx-browser/perf.mjs --dist $DIST [--profile load|export|reread|scroll] [--sheet NAME] [--frames N] [--no-scroll] [--no-jump] [file.xlsx ...]
// (--no-jump leaves out the jumps to the far end: 0.1.4 took minutes for one frame there.)
// With no files: tests/real-world/dashboard-sales-analysis.xlsx and milk-street-office.xlsx.
// --profile prints the functions with the most self time in that phase (CPU profile).
// --sheet picks the sheet to scroll (default: the one with the most cells).
// Environment: PLAYWRIGHT, JSZIP_DIR, PORT (default 16101).
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.resolve(here, '../..');
const args = process.argv.slice(2);
const opt = (name) => { const i = args.indexOf(name); if (i < 0) return undefined; const v = args[i + 1]; args.splice(i, 2); return v; };
const dist = path.resolve(opt('--dist') ?? path.join(repo, 'dist'));
const profile = opt('--profile');
const sheetName = opt('--sheet');
const frames = Number(opt('--frames') ?? 120);
const noJump = (() => { const i = args.indexOf('--no-jump'); if (i < 0) return false; args.splice(i, 1); return true; })();
const noScroll = (() => { const i = args.indexOf('--no-scroll'); if (i < 0) return false; args.splice(i, 1); return true; })();
const port = Number(process.env.PORT ?? 16101);
const jszipDir = [process.env.JSZIP_DIR, path.join(repo, 'node_modules/jszip'), path.resolve(repo, '../rbt-office/node_modules/jszip')].find((p) => p && fs.existsSync(p));
if (!jszipDir) { console.error('jszip not found; set JSZIP_DIR'); process.exit(2); }
const files = args.length ? args : ['tests/real-world/dashboard-sales-analysis.xlsx', 'tests/real-world/milk-street-office.xlsx'].map((f) => path.join(repo, f));

const page = `<!doctype html><meta charset="utf-8"><body style="margin:0"><div id="host" style="position:absolute;left:0;top:0;width:1200px;height:800px"></div>
<script src="/jszip.min.js"></script><script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script><script type="module">
try {
    const { default: createSpreadsheet } = await import('/spreadsheet/Spreadsheet.js');
    const importer = await import('/spreadsheet/XlsxImporter.js');
    const host = document.getElementById('host');
    let grid = null, written = null;
    window.perf = {
        async load(name) {
            const bytes = new Uint8Array(await (await fetch('/file?name=' + encodeURIComponent(name))).arrayBuffer());
            if (grid) grid.Dispose();
            grid = createSpreadsheet(host); grid.Mount();
            const t0 = performance.now();
            await grid.ImportXlsx(bytes);
            return performance.now() - t0;
        },
        async export() {
            // An edit first, as a host saves after one.
            grid.SetCell(0, 0, 'edited');
            const t0 = performance.now();
            written = await grid.ExportXlsx();
            return [performance.now() - t0, written.length];
        },
        async reread() {
            const t0 = performance.now();
            const names = await importer.getSheetNamesPromise(written);
            return [performance.now() - t0, Array.from(names).length];
        },
        sheets() { return grid.GetSheetNames().map((name, index) => ({ name, index, cells: grid.GetSheets()[index].Cells.size ?? 0, rows: grid.GetSheets()[index].RowCount })); },
        // Scroll the sheet a screenful at a time, one step per frame; the time of a
        // frame is from one requestAnimationFrame callback to the next.
        async scroll(index, frames, jump) {
            if (grid.GetActiveSheetIndex() !== index) grid.SwitchSheet(index);
            const scroller = host.querySelector('.spreadsheet-container');
            const content = host.querySelector('.spreadsheet-content');
            const model = grid.GetModel();
            // The extent the sheet needs (0.1.4 does not size it; the host did).
            const layout = await import('/spreadsheet/Layout.js');
            const extent = 30 + layout.rowTop(model, model.RowCount);
            if (content && content.getBoundingClientRect().height < extent - 1) content.style.height = extent + 'px';
            const next = () => new Promise((resolve) => requestAnimationFrame(resolve));
            scroller.scrollTop = 0; await next(); await next();
            const times = [];
            // A screenful per frame from the top (what holding PageDown does) ...
            const step = Math.max(40, scroller.clientHeight - 30);
            let last = await next();
            for (let i = 0; i < frames; i++) {
                scroller.scrollTop = Math.min(scroller.scrollHeight, (i + 1) * step);
                const now = await next();
                times.push(now - last); last = now;
            }
            // ... then one jump to the far end of the sheet and back to the middle.
            const jumps = [];
            for (const target of (jump ? [scroller.scrollHeight, Math.floor(scroller.scrollHeight / 2)] : [])) {
                scroller.scrollTop = target;
                const now = await next();
                jumps.push(now - last); last = now;
            }
            times.sort((a, b) => a - b);
            const at = (q) => times[Math.min(times.length - 1, Math.floor(q * times.length))];
            return { frames: times.length, mean: times.reduce((a, b) => a + b, 0) / times.length, median: at(0.5), p95: at(0.95), max: times[times.length - 1], over33: times.filter((t) => t > 33.4).length, rendered: host.querySelectorAll('.spreadsheet-content > *').length, jumps };
        },
    };
} catch (e) { window.loadError = String(e && e.stack || e); }
</script>`;
const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://x');
    const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
    if (url.pathname === '/') return send(200, 'text/html', page);
    if (url.pathname === '/jszip.min.js') return send(200, 'text/javascript', fs.readFileSync(path.join(jszipDir, 'dist/jszip.min.js')));
    if (url.pathname === '/jszip-shim.js') return send(200, 'text/javascript', 'export default window.JSZip;');
    if (url.pathname === '/file') { const name = url.searchParams.get('name'); return files.includes(name) ? send(200, 'application/octet-stream', fs.readFileSync(name)) : send(404, 'text/plain', 'no file'); }
    const file = path.join(dist, path.normalize(url.pathname));
    if (file.startsWith(dist) && fs.existsSync(file) && fs.statSync(file).isFile()) return send(200, 'text/javascript', fs.readFileSync(file));
    send(404, 'text/plain', 'not found');
});
await new Promise((resolve) => server.listen(port, '127.0.0.1', resolve));
const playwright = await import(process.env.PLAYWRIGHT ? pathToFileURL(process.env.PLAYWRIGHT).href : 'playwright');
const browser = await (playwright.chromium ?? playwright.default.chromium).launch();
const ms = (t) => (t >= 1000 ? (t / 1000).toFixed(2) + ' s' : t.toFixed(0) + ' ms');
try {
    for (const file of files) {
        const tab = await browser.newPage({ viewport: { width: 1200, height: 800 } });
        tab.on('pageerror', (e) => console.error('page error:', e.message));
        await tab.goto(`http://127.0.0.1:${port}/`);
        await tab.waitForFunction('window.perf !== undefined || window.loadError !== undefined', null, { timeout: 120000 });
        const loadError = await tab.evaluate('window.loadError');
        if (loadError) throw new Error('page failed to load modules: ' + loadError);
        const cdp = await tab.context().newCDPSession(tab);
        const phase = async (name, run) => {
            if (profile !== name) return run();
            await cdp.send('Profiler.enable'); await cdp.send('Profiler.setSamplingInterval', { interval: 500 }); await cdp.send('Profiler.start');
            const result = await run();
            const { profile: cpu } = await cdp.send('Profiler.stop');
            const self = new Map();
            const interval = cpu.timeDeltas.reduce((a, b) => a + b, 0) / Math.max(1, cpu.samples.length) / 1000;
            const byId = new Map(cpu.nodes.map((node) => [node.id, node]));
            for (const id of cpu.samples) {
                const node = byId.get(id), frame = node.callFrame;
                const key = `${frame.functionName || '(anonymous)'}  ${frame.url.split('/').slice(-2).join('/')}:${frame.lineNumber + 1}`;
                self.set(key, (self.get(key) ?? 0) + interval);
            }
            console.log(`  -- self time, ${name} --`);
            for (const [key, time] of [...self.entries()].sort((a, b) => b[1] - a[1]).slice(0, 40)) console.log(`  ${time.toFixed(0).padStart(7)} ms  ${key}`);
            return result;
        };
        console.log(path.basename(file) + `  (${(fs.statSync(file).size / 1e6).toFixed(1)} MB)`);
        const load = await phase('load', () => tab.evaluate((name) => window.perf.load(name), file));
        console.log(`  load (ImportXlsx)        ${ms(load)}`);
        const [exportTime, size] = await phase('export', () => tab.evaluate(() => window.perf.export()));
        console.log(`  export (ExportXlsx)      ${ms(exportTime)}  (${(size / 1e6).toFixed(1)} MB written)`);
        const [reread] = await phase('reread', () => tab.evaluate(() => window.perf.reread()));
        console.log(`  re-read (sheet names)    ${ms(reread)}`);
        const sheets = await tab.evaluate(() => window.perf.sheets());
        const target = sheetName ? sheets.find((s) => s.name === sheetName) : sheets.slice().sort((a, b) => b.rows - a.rows)[0];
        if (target && !noScroll) {
            const s = await phase('scroll', () => tab.evaluate(([index, n, jump]) => window.perf.scroll(index, n, jump), [target.index, frames, !noJump]));
            console.log(`  scroll "${target.name}" (${target.rows} rows): ${s.frames} frames, mean ${s.mean.toFixed(1)} ms, median ${s.median.toFixed(1)} ms, p95 ${s.p95.toFixed(1)} ms, max ${s.max.toFixed(1)} ms, ${s.over33} over 33 ms; ${s.jumps.length ? `jump to the end ${s.jumps[0].toFixed(0)} ms, to the middle ${s.jumps[1].toFixed(0)} ms` : 'no jumps'}  (${s.rendered} nodes)`);
        }
        await tab.close();
    }
} finally { await browser.close(); server.close(); }
