#!/usr/bin/env node
// The showcase macros (tests/macros/showcase) on the grid in a headless browser: each on a blank grid,
// put in with SetMacroModule, run with RunMacro (the browser path: Fable's VBA engine and MacroHost),
// then a PNG of the grid, for a look at what they painted.
//
//   pnpm run build:fable
//   PLAYWRIGHT=<repo>/tools/browser-runner/node_modules/playwright/index.mjs \
//     node tools/xlsx-browser/showcase.mjs --dist dist --out <dir> [name ...]
//
// Header comments of a macro (as tests/macros/showcase.fsx reads them): `' Entry: Main`, `' Input: text`, `' Before: Seed`;
// the PNG shows the grid as the macro left it (1280 x 720, its top left).
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.resolve(here, '../..');
const args = process.argv.slice(2);
const opt = (name, fallback) => { const i = args.indexOf(name); if (i < 0) return fallback; const v = args[i + 1]; args.splice(i, 2); return v; };
const dist = path.resolve(opt('--dist', path.join(repo, 'dist')));
const out = path.resolve(opt('--out', path.join(repo, 'artifacts/showcase')));
const timeout = Number(opt('--timeout', '180000'));
// --every MS: a PNG every MS milliseconds while the macro runs (<name>-001.png ...); --video: a WebM of each run.
// --size 1920x1080: the page (and video) size (1280x720); --hold MS: stay on what it left (for a video).
// --play SEC: a macro that goes on after its entry ends (Application.OnTime: a game's loop) is watched for
// SEC seconds; with `' Autopilot: snake` the keys are pressed to play it (towards the apple), as a player would.
const every = Number(opt('--every', '0'));
const [width, height] = opt('--size', '1280x720').split('x').map(Number);
const hold = Number(opt('--hold', '0'));
const play = Number(opt('--play', '0'));
const video = (() => { const i = args.indexOf('--video'); if (i < 0) return false; args.splice(i, 1); return true; })();
const port = Number(process.env.PORT ?? 16031);
const folder = path.join(repo, 'tests/macros/showcase');
fs.mkdirSync(out, { recursive: true });

const header = (key, source) => source.split(/\r?\n/).map((line) => line.trim()).filter((line) => line.toLowerCase().startsWith(`' ${key.toLowerCase()}:`)).map((line) => line.slice(key.length + 3).trim());
const groups = new Map();
for (const file of fs.readdirSync(folder).sort()) {
    if (!/\.(bas|cls)$/.test(file) || file.startsWith('_')) continue;
    const name = file.split('.')[0];
    if (args.length && !args.includes(name)) continue;
    const parts = file.replace(/\.(bas|cls)$/, '').split('.');
    const moduleName = parts.length > 1 ? parts[parts.length - 1] : 'Module1';
    if (!groups.has(name)) groups.set(name, []);
    groups.get(name).push({ module: moduleName, cls: file.endsWith('.cls'), source: fs.readFileSync(path.join(folder, file), 'utf8') });
}

const jszipDir = [process.env.JSZIP_DIR, path.join(repo, 'node_modules/jszip')].find((p) => p && fs.existsSync(p));
if (!jszipDir) { console.error('jszip not found; set JSZIP_DIR'); process.exit(2); }
const page = `<!doctype html><meta charset="utf-8"><body style="margin:0"><script src="/jszip.min.js"></script>
<script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script>
<script type="module">
try {
    const { default: createSpreadsheet } = await import('/spreadsheet/Spreadsheet.js');
    window.runShowcase = async (modules, entry, inputs, before) => {
        document.body.innerHTML = '';
        const host = document.createElement('div');
        host.id = 'grid';
        host.setAttribute('style', 'position:absolute;left:0;top:0;width:100vw;height:100vh');
        document.body.appendChild(host);
        const grid = createSpreadsheet(host);
        grid.Mount();
        const output = [];
        let asked = 0;
        grid.SetMacroHooks({ output: (text) => output.push(text), messageBox: (prompt) => { output.push('MsgBox: ' + prompt); return 1; },
                             inputBox: () => inputs.length ? inputs[Math.min(asked++, inputs.length - 1)] : null });
        for (const m of modules) {
            // (A class module is told by its Attribute lines; one found without them gets them.)
            const code = m.cls && !/Attribute VB_/i.test(m.source) ? 'Attribute VB_Name = "' + m.module + '"\\nAttribute VB_PredeclaredId = False\\nAttribute VB_Exposed = False\\n' + m.source : m.source;
            await grid.SetMacroModule(m.module, code);
        }
        // (Checked once every module is in: one alone may use what another declares.)
        const problems = grid.GetMacroProblems();
        if (problems && problems.length) return { ok: false, compile: problems.map((p) => \`\${p.module} \${p.line}: \${p.description}\`).slice(0, 3) };
        await grid.EnableMacros();
        if (before) { const set = await grid.RunMacro(before); if (!set.ok) return { ...set, description: before + ': ' + set.description }; }
        const started = performance.now();
        const result = await grid.RunMacro(entry);
        const ms = performance.now() - started;
        window.lastGrid = grid;
        return { ...result, ms, output: output.slice(0, 5) };
    };
    window.ready = true;
} catch (e) { window.loadError = String(e && e.stack || e); }
</script>`;

// Plays tests/macros/showcase/snake (its board B2:Z27, coloured cells): the head steps towards the apple,
// never into the wall or itself, by the arrow keys the game binds with Application.OnKey.
const snakePilot = () => {
    const grid = window.lastGrid;
    const cells = grid.GetRange('B2:Z27');
    const fill = (r, c) => (r < 0 || c < 0 || r >= cells.length || c >= cells[0].length) ? 'wall' : String(cells[r][c]?.format?.fill ?? cells[r][c]?.fill ?? '').toUpperCase();
    let head = null, apple = null;
    for (let r = 0; r < cells.length; r++) for (let c = 0; c < cells[0].length; c++) {
        const f = fill(r, c);
        if (f === '#EB6B6F') head = [r, c];
        if (f === '#7C3F58') apple = [r, c];
    }
    if (!head) return 'no head';
    const free = (r, c) => { const f = fill(r, c); return f === '#FFF6D3' || f === '#7C3F58'; };
    const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    const last = window.pilotHeading ?? 'ArrowUp';
    const distance = ([r, c]) => apple ? Math.abs(r - apple[0]) + Math.abs(c - apple[1]) : 0;
    const options = Object.entries(moves)
        .map(([key, [dr, dc]]) => ({ key, at: [head[0] + dr, head[1] + dc] }))
        .filter((o) => free(o.at[0], o.at[1]))
        .sort((a, b) => distance(a.at) - distance(b.at) || (a.key === last ? -1 : b.key === last ? 1 : 0));
    if (!options.length) return 'trapped';
    const choice = options[0].key;
    if (choice !== last || distance(head) > 0) {
        window.pilotHeading = choice;
        document.querySelector('.spreadsheet-content').dispatchEvent(new KeyboardEvent('keydown', { key: choice, bubbles: true, cancelable: true }));
    }
    return choice;
};

const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://x');
    const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
    if (url.pathname === '/') return send(200, 'text/html', page);
    if (url.pathname === '/jszip.min.js') return send(200, 'text/javascript', fs.readFileSync(path.join(jszipDir, 'dist/jszip.min.js')));
    if (url.pathname === '/jszip-shim.js') return send(200, 'text/javascript', 'export default window.JSZip;');
    const file = path.join(dist, path.normalize(url.pathname));
    if (file.startsWith(dist) && fs.existsSync(file) && fs.statSync(file).isFile()) return send(200, file.endsWith('.ttf') ? 'font/ttf' : 'text/javascript', fs.readFileSync(file));
    send(404, 'text/plain', 'not found');
});
await new Promise((resolve) => server.listen(port, '127.0.0.1', resolve));
const playwright = await import(process.env.PLAYWRIGHT ? pathToFileURL(process.env.PLAYWRIGHT).href : 'playwright');
const browser = await (playwright.chromium ?? playwright.default.chromium).launch();
let failed = 0;
try {
    for (const [name, modules] of groups) {
        const all = modules.map((m) => m.source).join('\n');
        const entry = header('Entry', all)[0] ?? 'Main';
        const inputs = header('Input', all);
        const before = header('Before', all)[0] ?? null;
        for (const file of fs.readdirSync(out)) if (new RegExp(`^${name}-\\d+\\.png$`).test(file)) fs.rmSync(path.join(out, file));
        const context = await browser.newContext({ viewport: { width, height }, ...(video ? { recordVideo: { dir: out, size: { width, height } } } : {}) });
        const tab = await context.newPage();
        const errors = [];
        tab.on('pageerror', (e) => errors.push(e.message));
        await tab.goto(`http://127.0.0.1:${port}/`);
        await tab.waitForFunction('window.ready || window.loadError', null, { timeout: 60000 });
        const loadError = await tab.evaluate('window.loadError');
        if (loadError) throw new Error(loadError);
        let result;
        // (The grid runs the macro a slice at a time and shows what it does: frames can be taken meanwhile.)
        const running = tab.evaluate(([m, e, i, b]) => window.runShowcase(m, e, i, b), [modules, entry, inputs, before]);
        let done = false;
        running.then(() => { done = true; }, () => { done = true; });
        const frames = (async () => {
            if (!every) return;
            for (let k = 1; !done && k < 1000; k++) {
                await new Promise((resolve) => setTimeout(resolve, every));
                if (done) break;
                try { await tab.screenshot({ path: path.join(out, `${name}-${String(k).padStart(3, '0')}.png`), timeout: 10000 }); } catch { /* the page is busy */ }
            }
        })();
        try {
            result = await Promise.race([running,
                new Promise((_, reject) => setTimeout(() => reject(new Error(`timed out after ${timeout} ms`)), timeout))]);
        } catch (e) { result = { ok: false, description: String(e.message).split('\n')[0] }; }
        done = true;
        await frames;
        // (A game goes on on its timers: watched - and played - for a while.)
        if (play > 0 && result.ok) {
            const pilot = header('Autopilot', all)[0] ?? '';
            const until = Date.now() + play * 1000;
            while (Date.now() < until) {
                await new Promise((resolve) => setTimeout(resolve, 350));
                if (pilot === 'snake') await tab.evaluate(snakePilot);
            }
        }
        if (hold > 0) await new Promise((resolve) => setTimeout(resolve, hold));
        const shot = path.join(out, `${name}.png`);
        try {
            await tab.evaluate(() => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))));
            await tab.screenshot({ path: shot, timeout: 30000 });
        } catch (e) { result.shot = String(e.message).split('\n')[0]; }
        const status = result.ok ? 'ok  ' : 'FAIL';
        if (!result.ok) failed++;
        const detail = result.ok ? '' : (result.compile ? 'compile: ' + result.compile.join('; ') : `error ${result.number ?? ''} ${result.description ?? ''} (${result.module ?? ''}.${result.procedure ?? ''} line ${result.line ?? ''})`);
        console.log(`${status} ${name.padEnd(22)} ${result.ms ? (result.ms / 1000).toFixed(2) + 's' : ''}  ${detail} ${result.output?.length ? '| ' + result.output.join(' / ') : ''} ${errors.length ? '| page: ' + errors[0] : ''}`);
        await tab.close();
        if (video) { const file = await tab.video()?.path(); await context.close(); if (file) fs.renameSync(file, path.join(out, `${name}.webm`)); }
        else await context.close();
    }
} finally {
    await browser.close();
    server.close();
}
console.log(`screenshots in ${out}`);
process.exit(failed ? 1 : 0);
