#!/usr/bin/env node
// Render a .docx with eigenpal/docx-editor in headless Chromium and screenshot
// the painted pages. Mirrors tools/word-parity/capture-browser-docx.mjs so the
// two outputs are comparable (same viewport, same Windows fonts).
//
// Usage:
//   node capture-docx-editor.mjs <input.docx> <output.png> [--page N] [--zoom N] [--no-fonts]
import { spawn } from 'node:child_process';
import path from 'node:path';
import fs from 'node:fs/promises';
import { createRequire } from 'node:module';

const here = path.dirname(new URL(import.meta.url).pathname);
const officeRoot = path.resolve(process.env.OFFICE_ROOT ?? path.join(here, '../..'));
const args = process.argv.slice(2);
const positional = args.filter((a) => !a.startsWith('--'));
const fixture = path.resolve(positional[0] ?? '');
const outputPng = path.resolve(positional[1] ?? '');
const pageIndex = args.includes('--page') ? Number(args[args.indexOf('--page') + 1]) : null;
const zoom = args.includes('--zoom') ? Number(args[args.indexOf('--zoom') + 1]) : 1;
const useFonts = !args.includes('--no-fonts');
const fontDir = path.resolve(process.env.RB_WORD_FONT_DIR ?? path.join(officeRoot, 'artifacts/fonts/windows'));
const port = Number(process.env.DOCX_EDITOR_PORT ?? 5188);
const viewportWidth = Number(process.env.VIEWPORT_WIDTH ?? 1000);
const viewportHeight = Number(process.env.VIEWPORT_HEIGHT ?? 1600);

if (!positional[0] || !positional[1]) {
  console.error('Usage: node capture-docx-editor.mjs <input.docx> <output.png> [--page N] [--zoom N] [--no-fonts]');
  process.exit(2);
}

function run(command, cmdArgs, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, cmdArgs, { cwd: here, stdio: 'inherit', ...options });
    child.on('exit', (code) => (code === 0 ? resolve() : reject(new Error(`${command} exited ${code}`))));
  });
}
async function exists(p) { try { await fs.access(p); return true; } catch { return false; } }

await fs.mkdir(path.dirname(outputPng), { recursive: true });

if (!(await exists(path.join(here, 'node_modules')))) {
  await run('npm', ['install']);
}

const req = createRequire(path.join(here, 'package.json'));
const { chromium } = req('playwright');

const viteBin = path.join(here, 'node_modules/.bin/vite');
const server = spawn(viteBin, ['--host', '127.0.0.1', '--port', String(port)], {
  cwd: here, stdio: ['ignore', 'pipe', 'pipe'],
});
let serverReady = false;
server.stdout.on('data', (c) => { const t = c.toString(); process.stdout.write(t); if (t.includes('Local:')) serverReady = true; });
server.stderr.on('data', (c) => process.stderr.write(c));

const t0 = Date.now();
while (!serverReady && Date.now() - t0 < 60000) await new Promise((r) => setTimeout(r, 200));
if (!serverReady) { server.kill(); throw new Error('vite did not start'); }

let exitCode = 0;
try {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: viewportWidth, height: viewportHeight }, deviceScaleFactor: 2 });
  page.on('console', (m) => { if (m.type() === 'error' || m.type() === 'warning') console.log(`[browser:${m.type()}] ${m.text()}`); });
  page.on('pageerror', (e) => console.error(`[browser:pageerror] ${e?.stack || e}`));

  const fontsParam = useFonts && (await exists(fontDir)) ? `&fonts=/@fs/${fontDir}` : '';
  const url = `http://127.0.0.1:${port}/?doc=/@fs/${fixture}${fontsParam}&zoom=${zoom}`;
  console.log('loading', url);
  await page.goto(url);
  await page.waitForFunction(() => window.__docxEditorReady || window.__docxEditorError, null, { timeout: 60000 });
  const err = await page.evaluate(() => window.__docxEditorError);
  if (err) throw new Error('docx-editor render failed: ' + err);

  // Find the painted pages container; fall back to the whole container.
  const selectors = ['.paged-editor__pages', '.paged-editor', '#container'];
  let target = null;
  for (const sel of selectors) { if (await page.locator(sel).count()) { target = sel; break; } }
  if (pageIndex != null) {
    const pages = page.locator('.layout-page');
    if (await pages.count() > pageIndex) {
      await pages.nth(pageIndex).screenshot({ path: outputPng });
      console.log(`docx-editor capture ok (page ${pageIndex}): ${fixture} -> ${outputPng}`);
    } else { throw new Error(`page ${pageIndex} not found (have ${await pages.count()})`); }
  } else {
    await page.locator(target).first().screenshot({ path: outputPng });
    console.log(`docx-editor capture ok (${target}): ${fixture} -> ${outputPng}`);
  }
  await browser.close();
} catch (e) {
  console.error(e);
  exitCode = 1;
} finally {
  server.kill();
}
process.exit(exitCode);
