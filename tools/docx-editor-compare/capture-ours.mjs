#!/usr/bin/env node
// Render a .docx with OUR Word renderer (browser-runner canvas) and screenshot
// #word-canvas. Reuses the already-Fable-built JS under tools/browser-runner/src
// (does NOT re-run Fable). Mirrors the load flow of check-browser-runner.mjs.
//
// Usage: node capture-ours.mjs <input.docx> <output.png>
import { spawn } from 'node:child_process';
import path from 'node:path';
import fs from 'node:fs/promises';
import { createRequire } from 'node:module';

const here = path.dirname(new URL(import.meta.url).pathname);
const officeRoot = path.resolve(here, '../..');
const runnerRoot = path.join(officeRoot, 'tools/browser-runner');
const fixture = path.resolve(process.argv[2] ?? '');
const outputPng = path.resolve(process.argv[3] ?? '');
const port = Number(process.env.OURS_PORT ?? 5179);
const viewportWidth = Number(process.env.VIEWPORT_WIDTH ?? 1100);
const viewportHeight = Number(process.env.VIEWPORT_HEIGHT ?? 1800);
const fontDir = path.resolve(process.env.RB_WORD_FONT_DIR ?? path.join(officeRoot, 'artifacts/fonts/windows'));

if (!process.argv[2] || !process.argv[3]) {
  console.error('Usage: node capture-ours.mjs <input.docx> <output.png>');
  process.exit(2);
}
await fs.mkdir(path.dirname(outputPng), { recursive: true });
async function fileExists(p) { try { await fs.access(p); return true; } catch { return false; } }

const req = createRequire(path.join(runnerRoot, 'package.json'));
const { chromium } = req('playwright');
const viteBin = path.join(runnerRoot, 'node_modules/.bin/vite');
const server = spawn(viteBin, ['--host', '127.0.0.1', '--port', String(port)], { cwd: runnerRoot, stdio: ['ignore', 'pipe', 'pipe'] });
let ready = false;
server.stdout.on('data', (c) => { const t = c.toString(); if (t.includes('Local:')) ready = true; });
server.stderr.on('data', (c) => process.stderr.write(c));
const t0 = Date.now();
while (!ready && Date.now() - t0 < 60000) await new Promise((r) => setTimeout(r, 150));
if (!ready) { server.kill(); throw new Error('vite did not start'); }

let exitCode = 0;
try {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: viewportWidth, height: viewportHeight }, deviceScaleFactor: 2 });
  page.on('pageerror', (e) => console.error(`[pageerror] ${e?.stack || e}`));
  const params = new URLSearchParams({ doc: `/@fs/${fixture}` });
  // Canvas defaults to 900x1600; multi-page docs need a taller canvas or content
  // below 1600px is clipped. Override with CANVAS_HEIGHT/CANVAS_WIDTH.
  if (process.env.CANVAS_HEIGHT) params.set('canvasHeight', process.env.CANVAS_HEIGHT);
  if (process.env.CANVAS_WIDTH) params.set('canvasWidth', process.env.CANVAS_WIDTH);
  // Use the Windows fonts so our render matches MS Word (and docx-editor) — the
  // runner injects @font-face for Calibri/Cambria/Arial/Times from fontRoot.
  if (await fileExists(path.join(fontDir, 'calibri.ttf'))) {
    params.set('fontRoot', `/@fs/${fontDir}`);
  }
  // MEASURE=text → use the ../text shaping engine for layout measurement
  // (canvas layout == print) instead of browser measureText.
  if (process.env.MEASURE === 'text') {
    const ucd = path.resolve(officeRoot, '../text/unicode/src/Unicode/ucd.dat');
    if (await fileExists(ucd)) {
      params.set('measure', 'text');
      params.set('ucd', `/@fs/${ucd}`);
    } else { console.error('ucd.dat not found at', ucd); }
  }
  const load = async () => {
    await page.goto(`http://127.0.0.1:${port}/?${params.toString()}`);
    await page.waitForFunction(() => {
      const s = document.querySelector('#status')?.textContent || '';
      return window.rbWordRunner && (s.startsWith('Loaded ') || s.startsWith('Load failed:'));
    }, null, { timeout: 60000 });
    const status = await page.locator('#status').textContent();
    if (!status?.startsWith('Loaded ')) throw new Error('load failed: ' + status);
  };
  await load();
  // Auto-size: the canvas defaults to 1600px tall; multi-page docs have content
  // below that which would be silently clipped. Measure the content bottom and,
  // if it exceeds the canvas, reload with a canvas tall enough to show it all.
  if (!process.env.CANVAS_HEIGHT) {
    const contentBottom = await page.evaluate(() => {
      const segs = JSON.parse(window.rbWordRunner.getLayoutJson());
      return segs.reduce((m, s) => Math.max(m, (s.Y || 0) + (s.Height || 0)), 0);
    });
    const needed = Math.ceil(contentBottom + 160);
    if (needed > 1600) {
      params.set('canvasHeight', String(needed));
      await load();
    }
  }
  await page.evaluate(() => window.rbWordRunner.clearSelection());
  await new Promise((r) => setTimeout(r, 300));
  await page.locator('#word-canvas').screenshot({ path: outputPng });
  console.log(`ours capture ok: ${fixture} -> ${outputPng}`);
  await browser.close();
} catch (e) { console.error(e); exitCode = 1; } finally { server.kill(); }
process.exit(exitCode);
