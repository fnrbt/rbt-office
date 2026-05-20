#!/usr/bin/env node
import { spawn } from 'node:child_process';
import path from 'node:path';
import fs from 'node:fs/promises';
import { createRequire } from 'node:module';

const officeRoot = path.resolve(process.env.OFFICE_ROOT ?? process.cwd());
const runnerRoot = path.join(officeRoot, 'tools/browser-runner');
const args = process.argv.slice(2);
let outputJsonArg = null;
if (args[2] === '--segments') {
  outputJsonArg = args[3] ?? null;
} else {
  outputJsonArg = args[2] ?? null;
}
const fixture = path.resolve(args[0] ?? '');
const outputPng = path.resolve(args[1] ?? '');
const outputJson = outputJsonArg ? path.resolve(outputJsonArg) : null;
const port = Number(process.env.BROWSER_CAPTURE_PORT ?? 5181);
const fontDir = path.resolve(process.env.RB_WORD_FONT_DIR ?? path.join(officeRoot, 'artifacts/fonts/windows'));
const layoutMode = process.env.BROWSER_CAPTURE_LAYOUT ?? 'segments';
const viewportWidth = Number(process.env.BROWSER_CAPTURE_VIEWPORT_WIDTH ?? 1100);
const viewportHeight = Number(process.env.BROWSER_CAPTURE_VIEWPORT_HEIGHT ?? 1800);
const canvasWidth = Number(process.env.BROWSER_CAPTURE_CANVAS_WIDTH ?? 900);
const canvasHeight = Number(process.env.BROWSER_CAPTURE_CANVAS_HEIGHT ?? 1600);
const loadTimeoutMs = Number(process.env.BROWSER_CAPTURE_LOAD_TIMEOUT_MS ?? 30000);
const serverSettleMs = Number(process.env.BROWSER_CAPTURE_SERVER_SETTLE_MS ?? 500);
const requestedCaptureAttempts = Number(process.env.BROWSER_CAPTURE_ATTEMPTS ?? 2);
const captureAttempts = Number.isFinite(requestedCaptureAttempts) && requestedCaptureAttempts > 0
  ? Math.floor(requestedCaptureAttempts)
  : 1;

if (!args[0] || !args[1] || (args[2] === '--segments' && !args[3])) {
  console.error('Usage: OFFICE_ROOT=/path/to/office node tools/word-parity/capture-browser-docx.mjs input.docx output.png [layout.json|--segments layout.json]');
  process.exit(2);
}

function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd: options.cwd ?? runnerRoot, stdio: 'inherit', ...options });
    child.on('exit', code => code === 0 ? resolve() : reject(new Error(`${command} exited ${code}`)));
  });
}

async function fileExists(file) {
  try {
    await fs.access(file);
    return true;
  } catch {
    return false;
  }
}

await fs.mkdir(path.dirname(outputPng), { recursive: true });
if (outputJson) await fs.mkdir(path.dirname(outputJson), { recursive: true });
const servedDir = path.join(runnerRoot, '.capture-docx');
await fs.mkdir(servedDir, { recursive: true });
const servedFixture = path.join(servedDir, path.basename(fixture));
await fs.copyFile(fixture, servedFixture);

try {
  await fs.access(path.join(runnerRoot, 'node_modules'));
} catch {
  await run('npm', ['install'], { cwd: runnerRoot });
}

const runnerRequire = createRequire(path.join(runnerRoot, 'package.json'));
const { chromium } = runnerRequire('playwright');

await run('dotnet', ['tool', 'restore'], { cwd: runnerRoot });
await run('dotnet', ['fable', '.', '--outDir', 'src', '--lang', 'JavaScript'], { cwd: runnerRoot });

const viteBin = path.join(runnerRoot, 'node_modules/.bin/vite');
const server = spawn(viteBin, ['--host', '127.0.0.1', '--port', String(port)], {
  cwd: runnerRoot,
  stdio: ['ignore', 'pipe', 'pipe'],
});

let serverReady = false;
server.stdout.on('data', chunk => {
  const text = chunk.toString();
  process.stdout.write(text);
  if (text.includes('Local:')) serverReady = true;
});
server.stderr.on('data', chunk => process.stderr.write(chunk));

for (let i = 0; i < 120 && !serverReady; i++) {
  await new Promise(resolve => setTimeout(resolve, 250));
}
if (!serverReady) {
  server.kill('SIGKILL');
  throw new Error('vite server did not start');
}
await new Promise(resolve => setTimeout(resolve, serverSettleMs));

async function captureOnce() {
  let browser = null;
  try {
    browser = await chromium.launch({
      args: ['--disable-dev-shm-usage', '--disable-gpu'],
    });
    const page = await browser.newPage({ viewport: { width: viewportWidth, height: viewportHeight } });
    const docUrl = `/@fs/${servedFixture}`;
    const params = new URLSearchParams({
      doc: docUrl,
      canvasWidth: String(canvasWidth),
      canvasHeight: String(canvasHeight),
    });
    if (await fileExists(path.join(fontDir, 'calibri.ttf'))) {
      params.set('fontRoot', `/@fs/${fontDir}`);
    }
    await page.goto(`http://127.0.0.1:${port}/?${params.toString()}`);
    await page.waitForFunction(
      () => window.rbWordRunner && document.querySelector('#status')?.textContent?.startsWith('Loaded '),
      null,
      { timeout: loadTimeoutMs }
    );
    await page.evaluate(() => window.rbWordRunner.clearSelection());
    if (outputJson) {
      const layout = layoutMode === 'range-info' || layoutMode === 'vba-range-info'
        ? await page.evaluate(() =>
            window.rbWordRunner.getVbaLayoutLines().map(line => ({
              ...line,
              Information: window.rbWordRunner.getRangeInformation(line.BlockId, line.StartOffset, line.EndOffset),
              Statistics: window.rbWordRunner.computeRangeStatistics(line.BlockId, line.StartOffset, line.EndOffset),
              Bounds: window.rbWordRunner.getRangeBounds(line.BlockId, line.StartOffset, line.EndOffset),
            }))
          )
        : layoutMode === 'table-geometry'
          ? await page.evaluate(() => window.rbWordRunner.getTableGeometry())
        : layoutMode === 'vba' || layoutMode === 'vba-lines'
        ? await page.evaluate(() => window.rbWordRunner.getVbaLayoutLines())
        : layoutMode === 'lines'
          ? await page.evaluate(() => window.rbWordRunner.getLayoutLines())
          : await page.evaluate(() => window.rbWordRunner.getLayoutSegments());
      await fs.writeFile(outputJson, JSON.stringify(layout, null, 2));
    }
    await page.locator('#word-canvas').screenshot({ path: outputPng });
    console.log(`browser capture ok: ${fixture} -> ${outputPng}`);
  } finally {
    if (browser) {
      await browser.close().catch(() => {});
    }
  }
}

try {
  let lastError = null;
  for (let attempt = 1; attempt <= captureAttempts; attempt += 1) {
    try {
      await captureOnce();
      lastError = null;
      break;
    } catch (error) {
      lastError = error;
      if (attempt < captureAttempts) {
        console.warn(`browser capture attempt ${attempt} failed: ${error.message}`);
        await new Promise(resolve => setTimeout(resolve, 750));
      }
    }
  }
  if (lastError) throw lastError;
} finally {
  server.kill('SIGKILL');
}
