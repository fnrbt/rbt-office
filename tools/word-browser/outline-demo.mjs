// A screen recording of the Word editor's outline features: a document with headings, View > Navigation
// pane (go to headings, collapse, move a heading with its content from the pane's menu), View > Outline
// (Show level, a heading moved and promoted, collapse and expand), back to print layout. For a product
// video (tools/word-browser/vba-demo.mjs is the VBA editor's).
//
//   pnpm run build:fable
//   PLAYWRIGHT=<repo>/tools/browser-runner/node_modules/playwright/index.mjs RB_WORD_FONT_DIR=<fonts> \
//     node tools/word-browser/outline-demo.mjs --dist dist --doc FILE.docx --out <dir> [--size 1920x1080] [--shots]
//
// --doc: a document with Heading 1-3 (office-hype's docs/make_outline.py makes one); its headings are
// found by their text (the --pane, --move, --chapter, --promote options name them; defaults fit that document).
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
const out = path.resolve(opt('--out', 'word-outline-demo'));
const [width, height] = opt('--size', '1920x1080').split('x').map(Number);
// (--scale: the page is laid out that much smaller and drawn that much finer - larger on the video.)
const scale = Number(opt('--scale', '1.5'));
const docFile = path.resolve(opt('--doc', 'field-handbook.docx'));
const paneVisits = opt('--pane', 'Inspections|Faults and incidents|On site').split('|');
const moveHeading = opt('--move', 'Inspections');
const outlineLine = opt('--promote', 'Emergencies');
const chapterLine = opt('--chapter', 'Faults and incidents');
const shots = flag('--shots');
const port = Number(process.env.PORT ?? 16034);
const fontDir = path.resolve(process.env.RB_WORD_FONT_DIR ?? path.join(repo, 'artifacts/fonts/windows'));
fs.mkdirSync(out, { recursive: true });

const jszipDir = [process.env.JSZIP_DIR, path.join(repo, 'node_modules/jszip')].find((p) => p && fs.existsSync(p));
const faces = [['Calibri', 'calibri.ttf', 'normal', 'normal'], ['Calibri', 'calibrib.ttf', 'bold', 'normal'],
    ['Calibri', 'calibrii.ttf', 'normal', 'italic'], ['Calibri', 'calibriz.ttf', 'bold', 'italic'],
    ['Calibri Light', 'calibril.ttf', 'normal', 'normal'], ['Cambria', 'cambria.ttc', 'normal', 'normal'], ['Arial', 'arial.ttf', 'normal', 'normal']];
const page = `<!doctype html><meta charset="utf-8">
<style>html,body{margin:0;height:100%;overflow:hidden}#app{height:100vh}</style>
<body><div id="app"></div>
<script src="/jszip.min.js"></script><script type="importmap">{"imports":{"jszip":"/jszip-shim.js"}}</script>
<script type="module">
try {
    for (const [family, file, weight, style] of ${JSON.stringify(faces)}) {
        try { const face = new FontFace(family, 'url(/font?name=' + file + ')', { weight, style }); await face.load(); document.fonts.add(face); } catch {}
    }
    const Word = (await import('/word/Word.js')).default;
    const Host = await import('/word/WordHost.js');
    const Shell = await import('/word/WordShell.js');
    const Dom = await import('/word/FableUI/Dom.js');
    const host = Host.create('', 'word-shell');
    Dom.mountApp('#app', Shell.view(host));
    const bytes = await (await fetch('/doc.docx')).arrayBuffer();
    Host.openFile(host, new File([bytes], ${JSON.stringify(path.basename(docFile))}));
    for (let i = 0; i < 100 && /^Opening/.test(host.Status.Value); i++) await new Promise((r) => setTimeout(r, 100));
    window.Word = Word; window.editor = host.Editor;
    window.ready = true;
} catch (e) { window.loadError = String(e && e.stack || e); }
</script>`;

const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://x');
    const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
    if (url.pathname === '/') return send(200, 'text/html', page);
    if (url.pathname === '/doc.docx') return send(200, 'application/octet-stream', fs.readFileSync(docFile));
    if (url.pathname === '/jszip.min.js') return send(200, 'text/javascript', fs.readFileSync(path.join(jszipDir, 'dist/jszip.min.js')));
    if (url.pathname === '/jszip-shim.js') return send(200, 'text/javascript', 'export default window.JSZip;');
    if (url.pathname === '/font') { const name = path.basename(url.searchParams.get('name') ?? ''); const file = path.join(fontDir, name); return name && fs.existsSync(file) ? send(200, 'application/octet-stream', fs.readFileSync(file)) : send(404, 'text/plain', 'no font'); }
    const file = path.join(dist, path.normalize(url.pathname));
    const type = file.endsWith('.css') ? 'text/css' : file.endsWith('.js') || file.endsWith('.mjs') ? 'text/javascript' : 'application/octet-stream';
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
tab.on('console', (m) => { if (m.type() === 'error' || process.env.DEMO_DEBUG) console.error('console:', m.type(), m.text()); });
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
let step = 0;
const shot = async (label) => { if (shots) await tab.screenshot({ path: path.join(out, `${String(++step).padStart(2, '0')}-${label}.png`) }); };
let started = Date.now();
let stopScreencast = null;
const events = [];
const testid = (id) => `[data-testid="${id}"]`;
const navItem = (text) => tab.locator(testid('word-nav-item'), { hasText: text }).first();
const outlineRow = (text) => tab.locator(testid('word-outline-line'), { hasText: text }).first();
// (The pointer glides to what it clicks, as a person's would.)
const glide = async (locator) => { const b = await locator.boundingBox(); if (!b) throw new Error('not on screen'); await tab.mouse.move(b.x + Math.min(40, b.width / 2), b.y + b.height / 2, { steps: 18 }); return b; };
const press = async (locator) => { await glide(locator); await wait(180); await locator.click({ position: undefined }); };
const mark = (label) => events.push({ t: (Date.now() - started) / 1000, label });
try {
    await tab.goto(`http://127.0.0.1:${port}/`);
    await tab.waitForFunction('window.ready || window.loadError', null, { timeout: 60000 });
    const loadError = await tab.evaluate('window.loadError');
    if (loadError) throw new Error(loadError);
    await wait(1500);
    stopScreencast = await startScreencast(tab, path.join(out, 'frames'));
    started = Date.now();
    await tab.mouse.move(width / scale / 2, height / scale / 2);
    await wait(1500);
    await shot('document');
    // The navigation pane.
    mark('pane');
    await press(tab.locator(testid('word-view-menu')));
    await wait(500);
    await press(tab.locator(testid('word-menu-view-navigation')));
    await tab.waitForSelector(testid('word-nav-item'));
    await wait(1200);
    await shot('pane');
    mark('go');
    for (const heading of paneVisits) { await press(navItem(heading)); await wait(1300); }
    await shot('went');
    mark('collapse');
    for (const heading of ['Before the visit', 'On site']) {
        await press(navItem(heading).locator(testid('word-nav-toggle')));
        await wait(700);
    }
    await wait(600);
    await shot('collapsed');
    mark('move');
    const item = navItem(moveHeading);
    await glide(item);
    await wait(300);
    await item.click({ button: 'right' });
    await wait(700);
    await press(tab.locator(testid('word-nav-menu-move-up')));
    await wait(1800);
    await shot('moved');
    for (const heading of ['Before the visit', 'On site']) {
        await press(navItem(heading).locator(testid('word-nav-toggle')));
        await wait(500);
    }
    await wait(800);
    // Outline view.
    mark('outline');
    await press(tab.locator(testid('word-view-menu')));
    await wait(500);
    await press(tab.locator(testid('word-menu-view-outline')));
    await tab.waitForSelector(testid('word-outline-line'));
    await wait(1600);
    await shot('outline');
    mark('levels');
    const show = tab.locator(testid('word-outline-show-level'));
    await glide(show);
    await show.selectOption('1');
    await wait(1500);
    await shot('level-1');
    // A chapter moved, with everything under it.
    mark('move-chapter');
    await press(outlineRow(chapterLine));
    await wait(600);
    await tab.keyboard.press('Alt+Shift+ArrowUp');
    await wait(1600);
    await shot('chapter-moved');
    mark('promote');
    await glide(show);
    await show.selectOption('2');
    await wait(1400);
    await press(outlineRow(outlineLine));
    await wait(600);
    await press(tab.locator(testid('word-outline-promote')));
    await wait(1800);
    await shot('promoted');
    mark('all');
    await glide(show);
    await show.selectOption({ label: 'All levels' });
    await wait(1800);
    await shot('all');
    mark('print');
    await press(tab.locator(testid('word-outline-close')));
    await wait(2500);
    mark('done');
    await shot('end');
} finally {
    if (stopScreencast) await stopScreencast(path.join(out, 'word-outline.mp4'));
    await tab.close();
    await context.close();
    await browser.close();
    server.close();
    fs.writeFileSync(path.join(out, 'word-outline.events.json'), JSON.stringify(events, null, 1));
}
console.log(`${errors.length ? 'page errors: ' + errors.join(' / ') + '\n' : ''}${path.join(out, 'word-outline.mp4')}`);
