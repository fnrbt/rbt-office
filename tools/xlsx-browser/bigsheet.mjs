#!/usr/bin/env node
// Operations on a large sheet (dashboard-sales-analysis.xlsx "Data", 846 000 cells), each timed:
// the median of 5 runs of the call, and of the call with the frame after it.
// usage: node tools/xlsx-browser/bigsheet.mjs [--dist DIR] [--profile NAME] [--only a,b] [--runs N]
// Environment: PLAYWRIGHT, PORT (default 16177).
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const args = process.argv.slice(2);
const opt = (name) => { const i = args.indexOf(name); if (i < 0) return undefined; const v = args[i + 1]; args.splice(i, 2); return v; };
const dist = path.resolve(opt('--dist') ?? path.join(repo, 'dist'));
const profile = opt('--profile');
const only = opt('--only');
const runs = Number(opt('--runs') ?? 5);
const port = Number(process.env.PORT ?? 16177);
const file = path.join(repo, 'tests/real-world/dashboard-sales-analysis.xlsx');
const jszipDir = path.join(repo, 'node_modules/jszip');

const page = `<!doctype html><meta charset="utf-8"><body style="margin:0"><div id="host" style="position:absolute;left:0;top:0;width:1200px;height:800px"></div>
<script src="/jszip.min.js"></script><script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script><script type="module">
try {
    const { default: createSpreadsheet } = await import('/spreadsheet/Spreadsheet.js');
    const host = document.getElementById('host');
    let grid = null;
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const key = (k, init = {}) => host.querySelector('.spreadsheet-content').dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init }));
    const block = Array.from({ length: 1000 }, (_, r) => Array.from({ length: 11 }, (_, c) => String(r * 20 + c + 0.5)).join('\\t')).join('\\n');
    const ops = {
        'type a number': { run: () => grid.EnterText(7, 3, String(Math.round(Math.random() * 1e6))), after: () => grid.Undo() },
        'type a text': { run: () => grid.EnterText(8, 3, 'some text ' + Math.random()), after: () => grid.Undo() },
        'type a formula': { run: () => grid.EnterText(9, 14, '=D8+' + Math.round(Math.random() * 100)), after: () => grid.Undo() },
        'undo a typed number': { before: () => grid.EnterText(7, 3, String(Math.round(Math.random() * 1e6))), run: () => grid.Undo() },
        'redo a typed number': { before: () => { grid.EnterText(7, 3, String(Math.round(Math.random() * 1e6))); grid.Undo(); }, run: () => grid.Redo(), after: () => grid.Undo() },
        'select a cell': { run: () => grid.SelectCell(20 + Math.round(Math.random() * 10), 4) },
        'arrow down': { before: () => grid.SelectCell(12, 4), run: () => key('ArrowDown') },
        'select a range of 10 000 cells': { run: () => grid.SelectRange(0, 0, 999, 9) },
        'select a column (header click)': { run: () => grid.SelectRange(0, 3, grid.GetModel().RowCount - 1, 3), after: () => grid.SelectCell(5, 5) },
        'select all': { run: () => grid.SelectRange(0, 0, grid.GetModel().RowCount - 1, grid.GetModel().ColCount - 1), after: () => grid.SelectCell(5, 5) },
        'scroll a screenful': { run: () => { const s = host.querySelector('.spreadsheet-container'); s.scrollTop = (s.scrollTop + 700) % 200000; } },
        'set a column width': { run: () => grid.SetColumnWidth(3, 90 + Math.round(Math.random() * 40)), after: () => grid.Undo() },
        'set a row height': { run: () => grid.SetRowHeight(10, 20 + Math.round(Math.random() * 20)), after: () => grid.Undo() },
        'bold on one cell': { before: () => grid.SelectCell(7, 3), run: () => grid.ApplyFormat('bold'), after: () => grid.Undo() },
        'bold on 10 000 cells': { before: () => grid.SelectRange(0, 0, 999, 9), run: () => grid.ApplyFormat('bold'), after: () => grid.Undo() },
        'copy 1 000 x 11 cells': { before: () => grid.SelectRange(1, 0, 1000, 10), run: () => grid.Copy(), wait: true },
        'paste 1 000 x 11 cells': { before: () => grid.SelectCell(5, 0), run: () => grid.PasteText(block, ''), after: () => grid.Undo() },
        'undo a paste of 1 000 x 11': { before: () => { grid.SelectCell(5, 0); grid.PasteText(block, ''); }, run: () => grid.Undo() },
        'delete 10 000 cells': { before: () => grid.SelectRange(0, 0, 999, 9), run: () => key('Delete'), after: () => grid.Undo() },
        'insert a row': { run: () => grid.InsertRows(30, 1), after: () => grid.Undo() },
        'delete a row': { run: () => grid.DeleteRows(60, 1), after: () => grid.Undo() },
        'undo an inserted row': { before: () => grid.InsertRows(30, 1), run: () => grid.Undo() },
        'redo an inserted row': { before: () => { grid.InsertRows(30, 1); grid.Undo(); }, run: () => grid.Redo(), after: () => grid.Undo() },
        'insert 100 rows': { run: () => grid.InsertRows(30, 100), after: () => grid.Undo() },
        'delete 100 rows': { run: () => grid.DeleteRows(60, 100), after: () => grid.Undo() },
        'insert a column': { run: () => grid.InsertColumns(3, 1), after: () => grid.Undo() },
        'delete a column': { run: () => grid.DeleteColumns(3, 1), after: () => grid.Undo() },
        'sort by a column': { run: () => grid.SortRange('A1:K76931', [{ column: 3, descending: Math.random() > 0.5 }], { hasHeader: true }), after: () => grid.Undo() },
        'set an auto filter': { run: () => grid.SetAutoFilter('A1:K76931'), after: () => grid.ClearAutoFilter() },
        'filter a column': { before: () => grid.SetAutoFilter('A1:K76931'), run: () => grid.SetFilter(3, { conditions: [{ operator: 'greaterThan', value: 100 }] }), after: () => grid.ClearAutoFilter() },
        'find a text': { run: () => grid.FindAll ? grid.FindAll('zzzz-not-there') : null },
        'GetCell x 1000': { run: () => { for (let i = 0; i < 1000; i++) grid.GetCell(i, 3); } },
    };
    window.perf = {
        names: () => Object.keys(ops),
        async open(name) {
            const bytes = new Uint8Array(await (await fetch('/file')).arrayBuffer());
            grid = createSpreadsheet(host); grid.Mount();
            await grid.ImportXlsx(bytes);
            grid.SwitchSheet(grid.GetSheetNames().indexOf(name));
            await frame();
            grid.EnterText(7, 3, '54321'); grid.Undo(); await frame();
            return grid.GetSheets()[grid.GetActiveSheetIndex()].Cells.size;
        },
        async prepare(name) { const op = ops[name]; try { if (op.before) await op.before(); } catch (e) { return String(e && e.message || e); } await frame(); return ''; },
        async run(name) {
            const op = ops[name];
            try {
                const t0 = performance.now();
                const result = op.run();
                if (op.wait) await result;
                const call = performance.now() - t0;
                await new Promise((resolve) => requestAnimationFrame(resolve));
                const total = performance.now() - t0;
                return [call, total, ''];
            } catch (e) { return [0, 0, String(e && e.message || e)]; }
        },
        async finish(name) { const op = ops[name]; try { if (op.after) await op.after(); } catch (e) { return String(e && e.message || e); } await frame(); return ''; },
    };
} catch (e) { window.loadError = String(e && e.stack || e); }
</script>`;
const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://x');
    const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
    if (url.pathname === '/') return send(200, 'text/html', page);
    if (url.pathname === '/jszip.min.js') return send(200, 'text/javascript', fs.readFileSync(path.join(jszipDir, 'dist/jszip.min.js')));
    if (url.pathname === '/jszip-shim.js') return send(200, 'text/javascript', 'export default window.JSZip;');
    if (url.pathname === '/file') return send(200, 'application/octet-stream', fs.readFileSync(file));
    const target = path.join(dist, path.normalize(url.pathname));
    if (target.startsWith(dist) && fs.existsSync(target) && fs.statSync(target).isFile()) return send(200, 'text/javascript', fs.readFileSync(target));
    send(404, 'text/plain', 'not found');
});
await new Promise((resolve) => server.listen(port, '127.0.0.1', resolve));
const playwright = await import(pathToFileURL(process.env.PLAYWRIGHT ?? path.join(repo, 'node_modules/playwright/index.mjs')).href);
const browser = await (playwright.chromium ?? playwright.default.chromium).launch();
try {
    const tab = await browser.newPage({ viewport: { width: 1200, height: 800 } });
    tab.on('pageerror', (e) => console.error('page error:', e.message));
    await tab.goto(`http://127.0.0.1:${port}/`);
    await tab.waitForFunction('window.perf !== undefined || window.loadError !== undefined', null, { timeout: 120000 });
    const loadError = await tab.evaluate('window.loadError');
    if (loadError) throw new Error(loadError);
    const cells = await tab.evaluate(() => window.perf.open('Data'));
    console.log(`"Data": ${cells} cells`);
    const cdp = await tab.context().newCDPSession(tab);
    for (const name of await tab.evaluate(() => window.perf.names())) {
        if (only && !only.split(',').includes(name)) continue;
        const calls = [], totals = [];
        let problem = '';
        const profiled = profile === name;
        const selfTime = new Map(), totalTime = new Map();
        for (let i = 0; i < runs; i++) {
            problem ||= await tab.evaluate((n) => window.perf.prepare(n), name);
            if (profiled) { await cdp.send('Profiler.enable'); await cdp.send('Profiler.setSamplingInterval', { interval: 200 }); await cdp.send('Profiler.start'); }
            const [call, total, error] = await tab.evaluate((n) => window.perf.run(n), name);
            if (profiled) {
                const { profile: cpu } = await cdp.send('Profiler.stop');
                const interval = cpu.timeDeltas.reduce((a, b) => a + b, 0) / Math.max(1, cpu.samples.length) / 1000;
                const byId = new Map(cpu.nodes.map((node) => [node.id, node]));
                const keyOf = (node) => `${node.callFrame.functionName || '(anonymous)'}  ${node.callFrame.url.split('/').slice(-2).join('/')}:${node.callFrame.lineNumber + 1}`;
                const selfOf = new Map();
                for (const id of cpu.samples) { selfOf.set(id, (selfOf.get(id) ?? 0) + interval); const key = keyOf(byId.get(id)); selfTime.set(key, (selfTime.get(key) ?? 0) + interval); }
                const walk = (node, open) => {
                    let sum = selfOf.get(node.id) ?? 0;
                    const key = keyOf(node), mine = !open.has(key);
                    if (mine) open.add(key);
                    for (const child of node.children ?? []) sum += walk(byId.get(child), open);
                    if (mine) { open.delete(key); totalTime.set(key, (totalTime.get(key) ?? 0) + sum); }
                    return sum;
                };
                walk(cpu.nodes[0], new Set());
            }
            problem ||= error;
            problem ||= await tab.evaluate((n) => window.perf.finish(n), name);
            calls.push(call); totals.push(total);
        }
        const median = (list) => list.slice().sort((a, b) => a - b)[Math.floor(list.length / 2)];
        console.log(`  ${name.padEnd(34)} ${median(calls).toFixed(1).padStart(8)} ms   ${median(totals).toFixed(1).padStart(8)} ms with the frame${problem ? '   [' + problem + ']' : ''}`);
        if (profiled) {
            console.log(`   -- time under (${runs} runs) --`);
            for (const [key, time] of [...totalTime.entries()].sort((a, b) => b[1] - a[1]).slice(0, 40)) console.log(`   ${time.toFixed(0).padStart(7)} ms  ${key}`);
            console.log(`   -- self time (${runs} runs) --`);
            for (const [key, time] of [...selfTime.entries()].sort((a, b) => b[1] - a[1]).slice(0, 25)) console.log(`   ${time.toFixed(0).padStart(7)} ms  ${key}`);
        }
    }
} finally { await browser.close(); server.close(); }
