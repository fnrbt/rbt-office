// A screen recording of writing a macro in the spreadsheet's VBA editor and running it: the shell (menus,
// formula bar) over the grid, Data > VBA editor, the macro typed at a typing pace, Apply, Run - and the
// grid showing what the macro does as it runs. For a product video.
//
//   pnpm run build:fable
//   PLAYWRIGHT=<repo>/tools/browser-runner/node_modules/playwright/index.mjs \
//     node tools/xlsx-browser/vba-demo.mjs --dist dist --out <dir> [--size 1920x1080] [--shots] [--macro FILE]
//
// --macro FILE: the macro to type (a .bas body; default a colour sweep); --shots: PNGs of each step too.
import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { spawnSync } from 'node:child_process';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const args = process.argv.slice(2);
const opt = (name, fallback) => { const i = args.indexOf(name); if (i < 0) return fallback; const v = args[i + 1]; args.splice(i, 2); return v; };
const flag = (name) => { const i = args.indexOf(name); if (i < 0) return false; args.splice(i, 1); return true; };
const dist = path.resolve(opt('--dist', path.join(repo, 'dist')));
const out = path.resolve(opt('--out', 'vba-demo'));
const [width, height] = opt('--size', '1920x1080').split('x').map(Number);
// (--scale: the page is laid out that much smaller and drawn that much finer - larger on the video.)
const scale = Number(opt('--scale', '1.5'));
const macroFile = opt('--macro', null);
const shots = flag('--shots');
const port = Number(process.env.PORT ?? 16032);
fs.mkdirSync(out, { recursive: true });

const macro = macroFile ? fs.readFileSync(macroFile, 'utf8').trim() : [
    'Sub Sweep()',
    '    Dim r As Long, c As Long',
    '    For r = 1 To 40',
    '        For c = 1 To 26',
    '            Cells(r, c).Interior.Color = RGB(255 - r * 5, 60 + c * 7, 120 + r * 3)',
    '            Cells(r, c).Value = r * c',
    '        Next',
    '    Next',
    'End Sub'].join('\n');

const jszipDir = [process.env.JSZIP_DIR, path.join(repo, 'node_modules/jszip')].find((p) => p && fs.existsSync(p));
const page = `<!doctype html><meta charset="utf-8">
<style>html,body{margin:0;height:100%;overflow:hidden;font-family:Segoe UI,Carlito,Arial,sans-serif}
#app{display:flex;flex-direction:column;height:100vh}#shell{flex:none}#grid{position:relative;flex:1;min-height:0}</style>
<body><div id="app"><div id="shell"></div><div id="grid"></div></div>
<script src="/jszip.min.js"></script>
<script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script>
<script type="module">
try {
    const { default: createSpreadsheet } = await import('/spreadsheet/Spreadsheet.js');
    const { createSpreadsheetShell } = await import('/spreadsheet/Toolbar.js');
    const grid = createSpreadsheet(document.getElementById('grid'));
    grid.Mount();
    const shell = createSpreadsheetShell(grid, () => {});
    shell.Mount(document.getElementById('shell'));
    grid.SetMacroHooks({ output: () => {}, messageBox: () => 1 });
    window.grid = grid;
    window.ready = true;
} catch (e) { window.loadError = String(e && e.stack || e); }
</script>`;

const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://x');
    const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
    if (url.pathname === '/') return send(200, 'text/html', page);
    if (url.pathname === '/jszip.min.js') return send(200, 'text/javascript', fs.readFileSync(path.join(jszipDir, 'dist/jszip.min.js')));
    if (url.pathname === '/jszip-shim.js') return send(200, 'text/javascript', 'export default window.JSZip;');
    const file = path.join(dist, path.normalize(url.pathname));
    const type = file.endsWith('.css') ? 'text/css' : file.endsWith('.ttf') ? 'font/ttf' : 'text/javascript';
    if (file.startsWith(dist) && fs.existsSync(file) && fs.statSync(file).isFile()) return send(200, type, fs.readFileSync(file));
    send(404, 'text/plain', 'not found');
});
await new Promise((resolve) => server.listen(port, '127.0.0.1', resolve));
const playwright = await import(process.env.PLAYWRIGHT ? pathToFileURL(process.env.PLAYWRIGHT).href : 'playwright');
const browser = await (playwright.chromium ?? playwright.default.chromium).launch();

// The screen as the browser draws it, at device pixels (Playwright's own recording is at CSS size): Chrome's
// screencast, a JPEG a changed frame, put together by ffmpeg with each frame's own duration.
const startScreencast = async (tab, dir) => {
    fs.mkdirSync(dir, { recursive: true });
    const cdp = await tab.context().newCDPSession(tab);
    const frames = [];
    cdp.on('Page.screencastFrame', (f) => {
        const file = path.join(dir, `${String(frames.length).padStart(6, '0')}.jpg`);
        fs.writeFileSync(file, Buffer.from(f.data, 'base64'));
        frames.push({ file, t: f.metadata.timestamp });
        cdp.send('Page.screencastFrameAck', { sessionId: f.sessionId }).catch(() => {});
    });
    await cdp.send('Page.startScreencast', { format: 'jpeg', quality: 92, maxWidth: width, maxHeight: height, everyNthFrame: 1 });
    return async (output) => {
        // (The screen unchanged at the end is the last frame until now.)
        const stoppedAt = Date.now() / 1000;
        await cdp.send('Page.stopScreencast').catch(() => {});
        if (!frames.length) return;
        const list = frames.map((f, i) => `file '${f.file}'\nduration ${Math.max(0.001, (i + 1 < frames.length ? frames[i + 1].t : Math.max(f.t + 0.5, stoppedAt)) - f.t).toFixed(4)}`).join('\n') + `\nfile '${frames[frames.length - 1].file}'\n`;
        fs.writeFileSync(path.join(dir, 'list.txt'), list);
        const run = spawnSync('ffmpeg', ['-v', 'error', '-y', '-f', 'concat', '-safe', '0', '-i', path.join(dir, 'list.txt'), '-vf', `scale=${width}:${height}:force_original_aspect_ratio=decrease,pad=${width}:${height}:0:0,fps=30`, '-pix_fmt', 'yuv420p', '-c:v', 'libx264', '-crf', '14', output], { stdio: 'inherit' });
        if (run.status === 0) fs.rmSync(dir, { recursive: true, force: true });
    };
};
const context = await browser.newContext({ viewport: { width: Math.round(width / scale), height: Math.round(height / scale) }, deviceScaleFactor: scale });
const tab = await context.newPage();
const errors = [];
tab.on('pageerror', (e) => errors.push(e.message));
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
let step = 0;
const shot = async (label) => { if (shots) await tab.screenshot({ path: path.join(out, `${String(++step).padStart(2, '0')}-${label}.png`) }); };
// (What happened when, for the edit: captions and cuts.)
let started = Date.now();
let stopScreencast = null;
const events = [];
const mark = (label) => events.push({ t: (Date.now() - started) / 1000, label });
try {
    await tab.goto(`http://127.0.0.1:${port}/`);
    await tab.waitForFunction('window.ready || window.loadError', null, { timeout: 60000 });
    const loadError = await tab.evaluate('window.loadError');
    if (loadError) throw new Error(loadError);
    stopScreencast = await startScreencast(tab, path.join(out, 'frames'));
    started = Date.now();
    await wait(1200);
    await shot('blank');
    mark('menu');
    await tab.click('[data-sheet-menu="data"]');
    await wait(500);
    await tab.click('[data-sheet-menu-item="vba-editor"]');
    await tab.waitForSelector('[data-testid="dlg-field-source"]');
    await wait(700);
    await shot('editor');
    mark('type');
    await tab.click('[data-testid="dlg-field-source"]');
    await tab.keyboard.type(macro, { delay: 28 });
    await wait(600);
    await shot('typed');
    mark('apply');
    await tab.click('[data-testid="dlg-vba-apply"]');
    await wait(900);
    await shot('applied');
    mark('run');
    await tab.click('[data-testid="dlg-vba-run"]');
    await wait(300);
    // (The dialog closes for the run to be seen, as a user would close it.)
    const closer = await tab.$('[data-testid="dlg-vba-close"]');
    await tab.waitForFunction('!window.grid.MacroRunning', null, { timeout: 120000 });
    mark('done');
    await shot('ran');
    if (closer) { await closer.click().catch(() => {}); await wait(400); }
    await wait(2500);
    await shot('end');
} finally {
    if (stopScreencast) await stopScreencast(path.join(out, 'vba-editor.mp4'));
    await tab.close();
    await context.close();
    await browser.close();
    server.close();
    fs.writeFileSync(path.join(out, 'vba-editor.events.json'), JSON.stringify(events, null, 1));
}
console.log(`${errors.length ? 'page errors: ' + errors.join(' / ') + '\n' : ''}${path.join(out, 'vba-editor.mp4')}`);
