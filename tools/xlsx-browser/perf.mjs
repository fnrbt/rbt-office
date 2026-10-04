#!/usr/bin/env node
// Timings of the Spreadsheet web component in a headless browser, against
// Fable output: opening a workbook (ImportXlsx), writing it (ExportXlsx),
// reading the written file again, and frame times while scrolling a sheet.
//
// Usage:
//   node tools/xlsx-browser/perf.mjs --dist $DIST [--profile load|export|reread|scroll|edits|first] [--sheet NAME] [--frames N] [--steps 120,0] [--no-scroll] [--no-jump] [--edits] [file.xlsx ...]
// (--edits also times edits of a large block: a format on 10 000 cells, a paste of 1 000 x 20
// cells, a typed entry and their undo, on the first sheet and on the sheet with the most cells.)
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
// Scroll steps to measure, in pixels per frame (0 = a screenful): a mouse wheel notch is about 100 px.
const steps = (opt('--steps') ?? '120,0').split(',').map(Number);
const noJump = (() => { const i = args.indexOf('--no-jump'); if (i < 0) return false; args.splice(i, 1); return true; })();
const noScroll = (() => { const i = args.indexOf('--no-scroll'); if (i < 0) return false; args.splice(i, 1); return true; })();
const edits = (() => { const i = args.indexOf('--edits'); if (i < 0) return false; args.splice(i, 1); return true; })();
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
        // Edits of a large block on a sheet: a format applied to 10 000 cells, a paste of
        // 1 000 x 20 cells, one typed entry, and their undo. Each time is the call itself
        // (model, recalculation, render) and the frame that follows it.
        // The first edit of a workbook that was just opened: the sheets with formulas are
        // converted and the dependency graph is built (it is built in idle time otherwise).
        async firstEdit(name) {
            const bytes = new Uint8Array(await (await fetch('/file?name=' + encodeURIComponent(name))).arrayBuffer());
            if (grid) grid.Dispose();
            grid = createSpreadsheet(host); grid.Mount();
            await grid.ImportXlsx(bytes);
            const t0 = performance.now();
            grid.EnterText(7, 3, '54321');
            const first = performance.now() - t0;
            const s = grid.GetRecalcStats ? grid.GetRecalcStats() : null;
            const t1 = performance.now();
            grid.EnterText(7, 3, '12345');
            return { first, second: performance.now() - t1, kind: s ? s.Kind + (s.Reason ? ': ' + s.Reason : '') + ', ' + s.Formulas + ' formulas in the graph' : '' };
        },
        async edits(index) {
            if (grid.GetActiveSheetIndex() !== index) grid.SwitchSheet(index);
            const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
            await frame();
            // [the call, the call and the frame after it, what the recalculation was (GetRecalcStats)]
            const recalc = () => { const s = grid.GetRecalcStats ? grid.GetRecalcStats() : null; return s ? (s.Kind === 'incremental' ? 'incremental: ' + s.ChangedCells + ' cells changed, ' + s.EvaluatedFormulas + ' of ' + s.Formulas + ' formulas evaluated' : s.Kind + ' recalculation' + (s.Reason ? ': ' + s.Reason : '')) : ''; };
            const timed = async (run) => { const t0 = performance.now(); run(); const call = performance.now() - t0; const kind = recalc(); await new Promise((resolve) => requestAnimationFrame(resolve)); return [call, performance.now() - t0, kind]; };
            const out = {};
            // The first edit of the workbook builds the dependency graph (unless it was built in idle time).
            out.first = await timed(() => grid.EnterText(7, 3, '54321'));
            await timed(() => grid.Undo());
            grid.SelectRange(0, 0, 999, 9);
            out.format = await timed(() => grid.ApplyFormat('bold'));
            out.formatUndo = await timed(() => grid.Undo());
            out.color = await timed(() => grid.ApplyBgColor('#ffeeaa'));
            await timed(() => grid.Undo());
            out.border = await timed(() => grid.ApplyBorder('all', 'thin', '#000000'));
            await timed(() => grid.Undo());
            const block = Array.from({ length: 1000 }, (_, r) => Array.from({ length: 20 }, (_, c) => (c % 3 === 0 ? 'item ' + (r * 20 + c) : String(r * 20 + c + 0.5))).join('\\t')).join('\\n');
            grid.SelectCell(5, 2);
            out.paste = await timed(() => grid.PasteText(block, ''));
            out.pasteUndo = await timed(() => grid.Undo());
            grid.SelectCell(7, 3);
            out.type = await timed(() => grid.EnterText(7, 3, '12345'));
            out.typeText = await timed(() => grid.EnterText(8, 3, 'some text'));
            await timed(() => grid.Undo()); await timed(() => grid.Undo());
            grid.SelectRange(0, 0, 999, 9);
            out.clear = await timed(() => host.querySelector('.spreadsheet-content').dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', bubbles: true, cancelable: true })));
            out.clearUndo = await timed(() => grid.Undo());
            return out;
        },
        // Scroll the sheet a screenful at a time, one step per frame; the time of a
        // frame is from one requestAnimationFrame callback to the next.
        async scroll(index, frames, jump, stepSize) {
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
            // stepSize pixels per frame from the top; 0 is a screenful per frame
            // (every cell on screen is new in every frame) ...
            const step = stepSize > 0 ? stepSize : Math.max(40, scroller.clientHeight - 30);
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
        if (edits) {
            const f = await phase('first', () => tab.evaluate((name) => window.perf.firstEdit(name), file));
            console.log(`  the first edit right after opening (a typed number): ${ms(f.first)} [${f.kind}]; the same cell again ${ms(f.second)}`);
            for (const sheet of (sheetName ? sheets.filter((s) => s.name === sheetName) : [sheets[0], sheets.slice().sort((a, b) => b.cells - a.cells)[0]].filter((s, i, all) => all.indexOf(s) === i))) {
                const e = await phase('edits', () => tab.evaluate((index) => window.perf.edits(index), sheet.index));
                const both = ([call, total, kind]) => `${ms(call)} (${ms(total)} with the frame)${kind ? ' [' + kind + ']' : ''}`;
                console.log(`  edits on "${sheet.name}" (${sheet.cells} cells):`);
                console.log(`    the first typed number                ${both(e.first)}`);
                console.log(`    ApplyFormat('bold') on 10 000 cells   ${both(e.format)}; undo ${both(e.formatUndo)}`);
                console.log(`    ApplyBgColor on 10 000 cells          ${both(e.color)}`);
                console.log(`    ApplyBorder('all') on 10 000 cells    ${both(e.border)}`);
                console.log(`    paste of 1 000 x 20 cells             ${both(e.paste)}; undo ${both(e.pasteUndo)}`);
                console.log(`    one typed number                      ${both(e.type)}; one typed text ${both(e.typeText)}`);
                console.log(`    Delete on 10 000 cells                ${both(e.clear)}; undo ${both(e.clearUndo)}`);
            }
        }
        const target = sheetName ? sheets.find((s) => s.name === sheetName) : sheets.slice().sort((a, b) => b.rows - a.rows)[0];
        if (target && !noScroll) {
          for (const stepSize of steps) {
            const s = await phase('scroll', () => tab.evaluate(([index, n, jump, stepSize]) => window.perf.scroll(index, n, jump, stepSize), [target.index, frames, !noJump && stepSize === 0, stepSize]));
            console.log(`  scroll "${target.name}" (${target.rows} rows) ${stepSize > 0 ? stepSize + ' px' : 'a screenful'} per frame: ${s.frames} frames, mean ${s.mean.toFixed(1)} ms, median ${s.median.toFixed(1)} ms, p95 ${s.p95.toFixed(1)} ms, max ${s.max.toFixed(1)} ms, ${s.over33} over 33 ms; ${s.jumps.length ? `jump to the end ${s.jumps[0].toFixed(0)} ms, to the middle ${s.jumps[1].toFixed(0)} ms` : 'no jumps'}  (${s.rendered} nodes)`);
          }
        }
        await tab.close();
    }
} finally { await browser.close(); server.close(); }
