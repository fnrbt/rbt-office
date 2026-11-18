// A screen recording of the Word editor's VBA editor: the Word shell with a document, Insert > Macros
// (VBA), a macro typed at a typing pace, Apply, Run - the document as the macro leaves it. For a product
// video (tools/xlsx-browser/vba-demo.mjs is the spreadsheet's).
//
//   pnpm run build:fable
//   PLAYWRIGHT=<repo>/tools/browser-runner/node_modules/playwright/index.mjs RB_WORD_FONT_DIR=<fonts> \
//     node tools/word-browser/vba-demo.mjs --dist dist --out <dir> [--size 1920x1080] [--shots] [--macro FILE] [--text FILE]
//
// --macro FILE: the macro to type (default: every character a colour of a rainbow, the title larger);
// --text FILE: the document's paragraphs, one a line; --shots: PNGs of each step.
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
const out = path.resolve(opt('--out', 'word-vba-demo'));
const [width, height] = opt('--size', '1920x1080').split('x').map(Number);
// (--scale: the page is laid out that much smaller and drawn that much finer - larger on the video.)
const scale = Number(opt('--scale', '1.5'));
const macroFile = opt('--macro', null);
const textFile = opt('--text', null);
const shots = flag('--shots');
const port = Number(process.env.PORT ?? 16033);
const fontDir = path.resolve(process.env.RB_WORD_FONT_DIR ?? path.join(repo, 'artifacts/fonts/windows'));
fs.mkdirSync(out, { recursive: true });

const macro = macroFile ? fs.readFileSync(macroFile, 'utf8').trim() : [
    'Sub Rainbow()',
    '    Dim i As Long, k As Double',
    '    For i = 1 To ActiveDocument.Characters.Count',
    '        k = i / 4',
    '        ActiveDocument.Characters(i).Font.Color = RGB(128 + 127 * Sin(k), 128 + 127 * Sin(k + 2.1), 128 + 127 * Sin(k + 4.2))',
    '    Next',
    '    ActiveDocument.Paragraphs(1).Range.Font.Size = 40',
    '    ActiveDocument.Paragraphs(1).Range.Font.Bold = True',
    'End Sub'].join('\n');
const paragraphs = textFile ? fs.readFileSync(textFile, 'utf8').split(/\r?\n/).filter((line) => line.trim()) : [
    'Macros, in the browser',
    'This document is open in a web page. The macro that colours it is VBA - the language of Word\'s own macros - run by an engine written in F#, against an object model of the document: Characters, Paragraphs, Range, Font.',
    'Nothing is installed and nothing is sent to a server.'];

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
    Word.setParagraphs(host.Editor, ${JSON.stringify(paragraphs)});
    Word.setMacroHooks(host.Editor, { output: () => {}, messageBox: () => 1 });
    window.Word = Word; window.editor = host.Editor;
    window.ready = true;
} catch (e) { window.loadError = String(e && e.stack || e); }
</script>`;

const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://x');
    const send = (code, type, body) => { res.writeHead(code, { 'content-type': type, 'cache-control': 'no-store' }); res.end(body); };
    if (url.pathname === '/') return send(200, 'text/html', page);
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
const mark = (label) => events.push({ t: (Date.now() - started) / 1000, label });
try {
    await tab.goto(`http://127.0.0.1:${port}/`);
    await tab.waitForFunction('window.ready || window.loadError', null, { timeout: 60000 });
    const loadError = await tab.evaluate('window.loadError');
    if (loadError) throw new Error(loadError);
    // (The editor edits once its text engine is loaded: until then its commands, the VBA editor's too, wait.)
    await tab.waitForFunction(() => !document.body.innerText.includes('Loading shaped text engine'), null, { timeout: 60000 }).catch(() => {});
    stopScreencast = await startScreencast(tab, path.join(out, 'frames'));
    started = Date.now();
    await wait(1500);
    await shot('document');
    mark('menu');
    await tab.click('[data-testid="word-insert-menu"]');
    await wait(500);
    await tab.click('[data-testid="word-menu-vba"]');
    await tab.waitForSelector('[data-testid="word-vba-code"]', { timeout: 8000 }).catch(async (e) => {
        const ids = await tab.evaluate(() => [...document.querySelectorAll('[data-testid]')].map((n) => n.getAttribute('data-testid')).filter((t) => /vba|menu-v|dialog/.test(t)));
        console.error('no editor; ids:', ids.join(' '), '| errors:', errors.join(' / '));
        throw e;
    });
    await wait(700);
    await shot('editor');
    mark('type');
    // (The editor's module may start with code: what is typed replaces it.)
    await tab.click('[data-testid="word-vba-code"]');
    await tab.keyboard.press('Control+A');
    await tab.keyboard.type(macro, { delay: 26 });
    await wait(600);
    await shot('typed');
    mark('apply');
    await tab.click('[data-testid="word-vba-apply"]');
    await wait(900);
    await shot('applied');
    mark('run');
    await tab.click('[data-testid="word-vba-run"]');
    await wait(1500);
    mark('done');
    await shot('ran');
    const close = await tab.$('[data-testid="word-vba-close"]');
    if (close) { await close.click().catch(() => {}); await wait(400); }
    await wait(3000);
    await shot('end');
} finally {
    if (stopScreencast) await stopScreencast(path.join(out, 'word-vba-editor.mp4'));
    await tab.close();
    await context.close();
    await browser.close();
    server.close();
    fs.writeFileSync(path.join(out, 'word-vba-editor.events.json'), JSON.stringify(events, null, 1));
}
console.log(`${errors.length ? 'page errors: ' + errors.join(' / ') + '\n' : ''}${path.join(out, 'word-vba-editor.mp4')}`);
