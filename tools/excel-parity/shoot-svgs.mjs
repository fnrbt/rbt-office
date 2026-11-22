// node tools/excel-parity/shoot-svgs.mjs DIR: every DIR/*.svg (render-charts.fsx) drawn in Chromium with the Windows
// fonts (as shoot-browser-sheets.mjs registers them) and written beside it as a PNG of its own size, for comparing
// with Excel's pictures (Chart.Export). PLAYWRIGHT as tools/xlsx-browser/run.mjs.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { facesIn } from './windows-faces.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.resolve(here, '../..');
const dir = path.resolve(process.argv[2]);
const playwright = await import(pathToFileURL(path.resolve(process.env.PLAYWRIGHT ?? path.join(repo, 'tools/browser-runner/node_modules/playwright/index.mjs'))).href);
const fontDir = path.resolve(process.env.RB_WORD_FONT_DIR ?? path.join(repo, 'artifacts/fonts/windows'));
const faces = facesIn([fontDir, path.join(path.dirname(fontDir), 'windows-extra')]);
const css = faces.map((f) => `@font-face{font-family:'${f.family}';font-weight:${f.weight};font-style:${f.style};src:url('${pathToFileURL(path.join(f.dir, f.file)).href}')}`).join('\n');
const browser = await playwright.chromium.launch();
const page = await browser.newPage({ deviceScaleFactor: 1 });
for (const file of fs.readdirSync(dir).filter((f) => f.endsWith('.svg')).sort()) {
    const svg = fs.readFileSync(path.join(dir, file), 'utf8');
    const w = Number(/width="([0-9.]+)"/.exec(svg)[1]), h = Number(/height="([0-9.]+)"/.exec(svg)[1]);
    await page.setViewportSize({ width: Math.ceil(w), height: Math.ceil(h) });
    const html = path.join(dir, '.shoot.html');
    fs.writeFileSync(html, `<!doctype html><meta charset="utf-8"><style>${css}\nbody{margin:0;background:#fff}</style>${svg}`);
    await page.goto(pathToFileURL(html).href);
    await page.evaluate(() => document.fonts.ready);
    await page.screenshot({ path: path.join(dir, file.replace(/\.svg$/, '.png')), clip: { x: 0, y: 0, width: Math.ceil(w), height: Math.ceil(h) } });
}
await browser.close();
fs.rmSync(path.join(dir, '.shoot.html'), { force: true });
