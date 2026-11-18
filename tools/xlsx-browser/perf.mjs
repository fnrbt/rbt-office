#!/usr/bin/env node
// Timings of the Spreadsheet web component in a headless browser, against
// Fable output: opening a workbook (ImportXlsx), writing it (ExportXlsx),
// reading the written file again, and frame times while scrolling a sheet.
//
// Usage:
//   node tools/xlsx-browser/perf.mjs --dist $DIST [--profile load|export|reread|scroll|edits|first|repeat|host|more] [--first-runs N] [--only-first] [--host] [--more] [--sheet NAME] [--frames N] [--steps 120,0] [--no-scroll] [--no-jump] [--edits] [file.xlsx ...]
// (--edits also times edits of a large block: a format on 10 000 cells, a paste of 1 000 x 20
// cells, a typed entry and their undo, on the first sheet and on the sheet with the most cells.)
// (--no-jump leaves out the jumps to the far end: 0.1.4 took minutes for one frame there.)
// With no files: tests/real-world/dashboard-sales-analysis.xlsx and milk-street-office.xlsx.
// --profile prints the functions with the most self time in that phase (CPU profile);
// with PROFILE_TOTAL=1 in the environment also the time under each function.
// --sheet picks the sheet to scroll (default: the one with the most cells).
// --first-cell "Sheet!B7" (or "Sheet!B7=0.07,0.05": the two numbers typed) is the cell the first edit types into (default D8 of the first sheet; the stats of the
// recalculation are printed with it); --find-input lists the constant cells that the most formulas depend on
// ("model inputs": what a host's user edits), to pick one.
// PROFILE_LINES="eval$,run" adds the samples of those functions by line of the generated JS;
// JS_FLAGS are V8's ("--no-turbo-inlining" shows the time of functions the optimiser would fold into their callers).
// --host types the same edit as a host application does: see hostEdit.
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
const firstRuns = Number(opt('--first-runs') ?? 5);
const firstCell = opt('--first-cell');
const findInput = (() => { const i = args.indexOf('--find-input'); if (i < 0) return false; args.splice(i, 1); return true; })();
const edits = (() => { const i = args.indexOf('--edits'); if (i < 0) return false; args.splice(i, 1); return true; })();
const more = (() => { const i = args.indexOf('--more'); if (i < 0) return false; args.splice(i, 1); return true; })();
const hostLike = (() => { const i = args.indexOf('--host'); if (i < 0) return false; args.splice(i, 1); return true; })();
const onlyFirst = (() => { const i = args.indexOf('--only-first'); if (i < 0) return false; args.splice(i, 1); return true; })();
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
        // (In two calls, so that a profile of the edit does not hold the opening.)
        async openFresh(name) {
            const bytes = new Uint8Array(await (await fetch('/file?name=' + encodeURIComponent(name))).arrayBuffer());
            if (grid) grid.Dispose();
            grid = createSpreadsheet(host); grid.Mount();
            await grid.ImportXlsx(bytes);
        },
        async firstEdit(cell) {
            let row = 7, col = 3, value = '54321', again = '12345';
            if (cell) {
                const [place, values] = cell.split('=');
                const [sheet, ref] = place.split('!');
                grid.SwitchSheet(grid.GetSheetNames().indexOf(sheet));
                const m = /^([A-Z]+)([0-9]+)$/.exec(ref);
                col = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1; row = Number(m[2]) - 1;
                // (Releases before 0.1.8 have no GetCell: fixed numbers there.)
                const now = grid.GetCell ? grid.GetCell(row, col).value : null;
                if (values) [value, again] = values.split(',');
                else if (typeof now === 'number') { value = String(now * 1.1 + 1); again = String(now); } else { value = '0.07'; again = '0.05'; }
                await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
            }
            const t0 = performance.now();
            grid.EnterText(row, col, value);
            const first = performance.now() - t0;
            const s = grid.GetRecalcStats ? grid.GetRecalcStats() : null;
            const t1 = performance.now();
            grid.EnterText(row, col, again);
            window.perf.lastEdit = { row, col, value, again };
            return { first, second: performance.now() - t1, kind: s ? s.Kind + (s.Reason ? ': ' + s.Reason : '') + ', ' + s.Formulas + ' formulas in the graph' + (cell ? ', ' + s.ReachedFormulas + ' reached, ' + s.EvaluatedFormulas + ' evaluated' : '') : '' };
        },
        // The cell of firstEdit typed into again, several times (the graph is built, the sheets converted): the median.
        repeatEdit(times) {
            const { row, col, value, again } = window.perf.lastEdit;
            const all = [];
            for (let i = 0; i < times; i++) { const t0 = performance.now(); grid.EnterText(row, col, i % 2 ? again : value); all.push(performance.now() - t0); }
            return all.sort((a, b) => a - b)[Math.floor(all.length / 2)];
        },
        // The same edit as a host application makes it (--host): the entry typed into the cell's
        // edit box and committed with Enter, with subscribers that read what a toolbar, a formula
        // bar and a status bar show after every change (a host's editor does: selection, format,
        // status, the active cell, view options, filters, outline, undo state). Per edit:
        // [the keystroke's handler, that and the frame after it, the time inside the subscribers].
        async hostEdit(times) {
            const { row, col, value, again } = window.perf.lastEdit;
            let inSubscribers = 0, calls = 0;
            if (!window.perf.hostSubscribed) {
                window.perf.hostSubscribed = true;
                let structureKey = '';
                const refresh = () => {
                    const t0 = performance.now();
                    calls++;
                    grid.GetSelectionRef(); const active = grid.GetActiveCell();
                    grid.IsModified(); grid.CanUndo(); grid.CanRedo(); grid.GetActiveSheetIndex();
                    grid.GetSelectionFormat(); grid.GetStatus();
                    const cell = grid.GetCell(active.Row, active.Col);
                    if (cell.spill) grid.GetGhostFormula();
                    grid.GetSheetKind(); grid.GetViewOptions();
                    if (!grid.IsEditing()) grid.GetCurrentCellFormula();
                    const key = grid.GetActiveSheetIndex() + ':' + grid.GetRevision();
                    if (key !== structureKey) { structureKey = key; grid.GetFilters(); grid.GetOutline(); grid.GetShowOutline(); }
                    inSubscribers += performance.now() - t0;
                };
                grid.OnChange((change) => { if (change.StructureChanged || change.WorkbookChanged) { grid.GetSheetInfos(); grid.GetDefinedNames(); } refresh(); });
                grid.OnSelectionChange(() => refresh());
                grid.OnEditStateChange(() => { const t0 = performance.now(); grid.GetCurrentCellFormula(); inSubscribers += performance.now() - t0; });
                window.perf.hostCounters = () => { const out = [inSubscribers, calls]; inSubscribers = 0; calls = 0; return out; };
            }
            const frame = () => new Promise((resolve) => requestAnimationFrame(resolve));
            const all = [];
            for (let i = 0; i < times; i++) {
                grid.SelectCell(row, col);
                await frame(); await frame();
                window.perf.hostCounters();
                const t0 = performance.now();
                grid.StartEdit(i % 2 ? again : value);
                const box = document.activeElement;
                box.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
                const call = performance.now() - t0;
                await frame();
                const total = performance.now() - t0;
                const [inside, count] = window.perf.hostCounters();
                all.push([call, total, inside, count]);
            }
            const median = (k) => all.map((t) => t[k]).sort((a, b) => a - b)[Math.floor(all.length / 2)];
            return { call: median(0), total: median(1), subscribers: median(2), refreshes: median(3), moved: grid.GetActiveCell().Row !== row };
        },
        // The constant number cells of every sheet that the most formulas depend on.
        async findInputs() {
            const found = [];
            const sheets = grid.GetSheetNames().map((name) => ({ name }));
            for (let index = 0; index < sheets.length; index++) {
                grid.SwitchSheet(index);
                const used = grid.GetUsedRange();
                if (!used) continue;
                let tried = 0;
                for (let row = used.firstRow; row <= Math.min(used.lastRow, used.firstRow + 400) && tried < 400; row++)
                    for (let col = used.firstCol; col <= Math.min(used.lastCol, used.firstCol + 30) && tried < 400; col++) {
                        const cell = grid.GetCell(row, col);
                        if (cell.formula || typeof cell.value !== 'number') continue;
                        tried++;
                        grid.EnterText(row, col, String(cell.value + 1));
                        const stats = grid.GetRecalcStats();
                        grid.Undo();
                        if (stats.ReachedFormulas > 0) found.push({ cell: sheets[index].name + '!' + cell.address, reached: stats.ReachedFormulas, evaluated: stats.EvaluatedFormulas });
                    }
            }
            return found.sort((a, b) => b.reached - a.reached).slice(0, 8);
        },
        // Edits that are more than cells (--more): a paste whose formulas close a circular reference
        // (milk-street: Annual!C21:N40 pasted at C101), rows inserted and deleted, a defined name,
        // a filter, sheets renamed / added / moved / deleted. Each: [the call in ms, what the recalculation was].
        async more(name) {
            await window.perf.openFresh(name);
            const recalc = () => { const s = grid.GetRecalcStats ? grid.GetRecalcStats() : null; return s ? (s.Kind === 'incremental' ? 'incremental: ' + s.ChangedCells + ' cells changed, ' + s.EvaluatedFormulas + ' of ' + s.Formulas + ' formulas evaluated' : s.Kind + ' recalculation' + (s.Reason ? ': ' + s.Reason : '')) : ''; };
            const timed = async (run) => { const t0 = performance.now(); try { await run(); } catch (e) { return [performance.now() - t0, 'threw: ' + String(e && e.message || e)]; } return [performance.now() - t0, recalc()]; };
            const out = {};
            grid.EnterText(7, 3, '54321'); grid.Undo();   // (the graph is built)
            grid.SelectRange(20, 2, 39, 13); await grid.Copy();
            grid.SelectCell(100, 2);
            out.pasteCycle = await timed(() => grid.Paste());
            out.pasteCycleUndo = await timed(() => grid.Undo());
            out.pasteCycleRedo = await timed(() => grid.Redo());
            out.typedBesideCycle = await timed(() => grid.EnterText(7, 3, '777'));
            grid.Undo(); grid.Undo();
            grid.SelectCell(30, 2);
            out.insertRow = await timed(() => grid.InsertRows(30, 1));
            out.insertRowUndo = await timed(() => grid.Undo());
            out.deleteRow = await timed(() => grid.DeleteRows(60, 2));
            out.deleteRowUndo = await timed(() => grid.Undo());
            out.insertColumn = await timed(() => grid.InsertColumns(3, 1));
            await timed(() => grid.Undo());
            out.defineName = await timed(() => grid.SetDefinedName('PerfRate', '=' + "'" + grid.GetSheetNames()[0].replace(/'/g, "''") + "'" + '!$D$8'));
            out.deleteName = await timed(() => grid.DeleteDefinedName('PerfRate'));
            const used = grid.GetUsedRange();
            out.setFilter = await timed(() => { grid.SetAutoFilter('A' + (used.firstRow + 1) + ':' + 'H' + (used.lastRow + 1)); grid.SetFilter(2, { conditions: [{ operator: 'greaterThan', value: 0 }] }); });
            out.clearFilter = await timed(() => grid.ClearAutoFilter());
            const names = grid.GetSheetNames();
            out.renameSheet = await timed(() => grid.RenameSheet(0, names[0] + ' x'));
            await timed(() => grid.Undo());
            out.addSheet = await timed(() => grid.AddSheet('Perf sheet'));
            out.moveSheet = await timed(() => grid.MoveSheet(names.length, 0));
            out.deleteSheet = await timed(() => grid.DeleteSheet(0));
            return out;
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
        // A sheet put on display for the first time (its cells converted, its charts and shapes drawn):
        // [the call, the call and the frame after it, objects that are drawn / that are labelled frames].
        async show(index) {
            const other = grid.GetSheetNames().findIndex((_, i) => i !== index);
            if (grid.GetActiveSheetIndex() === index && other >= 0) { grid.SwitchSheet(other); await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))); }
            const t0 = performance.now();
            grid.SwitchSheet(index);
            const call = performance.now() - t0;
            await new Promise((resolve) => requestAnimationFrame(resolve));
            return [call, performance.now() - t0, host.querySelectorAll('.spreadsheet-object-drawn').length, host.querySelectorAll('.spreadsheet-object-label').length];
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
const browser = await (playwright.chromium ?? playwright.default.chromium).launch(process.env.JS_FLAGS ? { args: ['--js-flags=' + process.env.JS_FLAGS] } : {});
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
            const { profile: cpu } = await cdp.send('Profiler.stop'); if (process.env.PROFILE_OUT) fs.writeFileSync(process.env.PROFILE_OUT, JSON.stringify(cpu));
            const self = new Map();
            const interval = cpu.timeDeltas.reduce((a, b) => a + b, 0) / Math.max(1, cpu.samples.length) / 1000;
            const byId = new Map(cpu.nodes.map((node) => [node.id, node]));
            for (const id of cpu.samples) {
                const node = byId.get(id), frame = node.callFrame;
                const key = `${frame.functionName || '(anonymous)'}  ${frame.url.split('/').slice(-2).join('/')}:${frame.lineNumber + 1}`;
                self.set(key, (self.get(key) ?? 0) + interval);
            }
            if (process.env.PROFILE_TOTAL) {
                // Time under each function (itself and what it calls; a recursive one counts once per chain).
                const selfOf = new Map(); for (const id of cpu.samples) selfOf.set(id, (selfOf.get(id) ?? 0) + interval);
                const total = new Map();
                const keyOf = (node) => `${node.callFrame.functionName || '(anonymous)'}  ${node.callFrame.url.split('/').slice(-2).join('/')}:${node.callFrame.lineNumber + 1}`;
                const walk = (node, open) => {
                    let sum = selfOf.get(node.id) ?? 0;
                    const key = keyOf(node), mine = !open.has(key);
                    if (mine) open.add(key);
                    for (const child of node.children ?? []) sum += walk(byId.get(child), open);
                    if (mine) { open.delete(key); total.set(key, (total.get(key) ?? 0) + sum); }
                    return sum;
                };
                walk(cpu.nodes[0], new Set());
                console.log(`  -- time under, ${name} --`);
                for (const [key, time] of [...total.entries()].sort((a, b) => b[1] - a[1]).slice(0, 45)) console.log(`  ${time.toFixed(0).padStart(7)} ms  ${key}`);
            }
            if (process.env.PROFILE_LINES) {
                // Samples by source line (of the generated JS), for the functions named in PROFILE_LINES ("eval$,run").
                const wanted = new Set(process.env.PROFILE_LINES.split(','));
                const lines = new Map();
                for (const node of cpu.nodes) {
                    if (!wanted.has(node.callFrame.functionName)) continue;
                    for (const tick of node.positionTicks ?? []) { const key = `${node.callFrame.functionName} ${node.callFrame.url.split('/').slice(-1)[0]}:${tick.line}`; lines.set(key, (lines.get(key) ?? 0) + tick.ticks * interval); }
                }
                console.log(`  -- self time by line, ${name} --`);
                for (const [key, time] of [...lines.entries()].sort((a, b) => b[1] - a[1]).slice(0, 60)) console.log(`  ${time.toFixed(0).padStart(7)} ms  ${key}`);
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
        if (findInput) {
            console.log('  constant cells the most formulas depend on:');
            for (const input of await tab.evaluate(() => window.perf.findInputs())) console.log(`    ${input.cell}  ${input.reached} formulas reached, ${input.evaluated} evaluated`);
        }
        if (edits) {
            // The median of several openings (--first-runs N, default 5): one run is off by tens of milliseconds.
            const firsts = [];
            let f = null;
            for (let run = 0; run < firstRuns; run++) {
                await tab.evaluate((name) => window.perf.openFresh(name), file);
                f = await (run === firstRuns - 1 ? phase('first', () => tab.evaluate((cell) => window.perf.firstEdit(cell), firstCell ?? null)) : tab.evaluate((cell) => window.perf.firstEdit(cell), firstCell ?? null));
                if (run === firstRuns - 1 && profile === 'first') continue; // (the profiler's start gives the idle-time build the time to run first)
                firsts.push(f.first);
            }
            const sortedFirsts = firsts.slice().sort((a, b) => a - b);
            console.log(`  the first edit right after opening (a typed number${firstCell ? ' in ' + firstCell : ''}): median ${ms(sortedFirsts[Math.floor(sortedFirsts.length / 2)])} of ${firsts.map((t) => t.toFixed(0)).join(', ')} ms [${f.kind}]; the same cell again ${ms(f.second)}`);
            const repeated = await phase('repeat', () => tab.evaluate(() => window.perf.repeatEdit(7)));
            console.log(`  that cell typed into seven more times: median ${ms(repeated)}`);
            if (hostLike) {
                const h = await phase('host', () => tab.evaluate(() => window.perf.hostEdit(7)));
                console.log(`  the same edit as a host makes it (typed into the edit box, Enter, subscribers that read toolbar, bar and status), seven times: median ${ms(h.call)} for the keystroke, ${ms(h.total)} with the frame after it; ${ms(h.subscribers)} of it in the subscribers (${h.refreshes} refreshes)${h.moved ? '' : '  [the selection did not move: the entry was not committed]'}`);
            }
            for (const sheet of (onlyFirst ? [] : sheetName ? sheets.filter((s) => s.name === sheetName) : [sheets[0], sheets.slice().sort((a, b) => b.cells - a.cells)[0]].filter((s, i, all) => all.indexOf(s) === i))) {
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
        if (more) {
            const m = await phase('more', () => tab.evaluate((name) => window.perf.more(name), file));
            console.log('  edits that are more than cells:');
            for (const [key, [time, kind]] of Object.entries(m)) console.log(`    ${key.padEnd(18)} ${ms(time).padStart(8)}  [${kind}]`);
        }
        const target = sheetName ? sheets.find((s) => s.name === sheetName) : sheets.slice().sort((a, b) => b.rows - a.rows)[0];
        if (target && sheetName) {
            // (The sheet named with --sheet: what it costs to put it on display, a second time too.)
            const shown = [];
            for (let run = 0; run < 3; run++) shown.push(await tab.evaluate((index) => window.perf.show(index), target.index));
            console.log(`  "${target.name}" put on display: ${ms(shown[0][0])} (${ms(shown[0][1])} with the frame); again ${ms(Math.min(shown[1][0], shown[2][0]))} (${ms(Math.min(shown[1][1], shown[2][1]))})  [${shown[0][2]} objects drawn, ${shown[0][3]} labelled frames]`);
        }
        if (target && !noScroll) {
          for (const stepSize of steps) {
            const s = await phase('scroll', () => tab.evaluate(([index, n, jump, stepSize]) => window.perf.scroll(index, n, jump, stepSize), [target.index, frames, !noJump && stepSize === 0, stepSize]));
            console.log(`  scroll "${target.name}" (${target.rows} rows) ${stepSize > 0 ? stepSize + ' px' : 'a screenful'} per frame: ${s.frames} frames, mean ${s.mean.toFixed(1)} ms, median ${s.median.toFixed(1)} ms, p95 ${s.p95.toFixed(1)} ms, max ${s.max.toFixed(1)} ms, ${s.over33} over 33 ms; ${s.jumps.length ? `jump to the end ${s.jumps[0].toFixed(0)} ms, to the middle ${s.jumps[1].toFixed(0)} ms` : 'no jumps'}  (${s.rendered} nodes)`);
          }
        }
        await tab.close();
    }
} finally { await browser.close(); server.close(); }
