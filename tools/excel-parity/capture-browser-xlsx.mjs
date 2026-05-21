#!/usr/bin/env node
import { spawn } from 'node:child_process';
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';

const officeRoot = path.resolve(process.env.OFFICE_ROOT ?? process.cwd());
const runnerRoot = path.join(officeRoot, 'tools/browser-runner');
const fixture = path.resolve(process.argv[2] ?? '');
const outputPng = path.resolve(process.argv[3] ?? '');
const port = Number(process.env.EXCEL_BROWSER_CAPTURE_PORT ?? 5191);
const viewportWidth = Number(process.env.EXCEL_BROWSER_CAPTURE_VIEWPORT_WIDTH ?? 980);
const viewportHeight = Number(process.env.EXCEL_BROWSER_CAPTURE_VIEWPORT_HEIGHT ?? 760);
const canvasWidth = Number(process.env.EXCEL_BROWSER_CAPTURE_CANVAS_WIDTH ?? 900);
const canvasHeight = Number(process.env.EXCEL_BROWSER_CAPTURE_CANVAS_HEIGHT ?? 620);
const captureTmp = process.env.TMPDIR ?? path.join(officeRoot, 'tmp/browser-capture');
process.env.TMPDIR = captureTmp;

if (!process.argv[2] || !process.argv[3]) {
  console.error('Usage: node tools/excel-parity/capture-browser-xlsx.mjs input.xlsx output.png');
  process.exit(2);
}

function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd: options.cwd ?? officeRoot, stdio: 'inherit', ...options });
    child.on('exit', code => code === 0 ? resolve() : reject(new Error(`${command} exited ${code}`)));
  });
}

await fs.mkdir(path.dirname(outputPng), { recursive: true });
await fs.mkdir(path.join(runnerRoot, '.capture-xlsx'), { recursive: true });
await fs.mkdir(captureTmp, { recursive: true });

if (process.env.EXCEL_BROWSER_CAPTURE_BUILD !== '0') {
  await run('dotnet', ['tool', 'restore'], { cwd: officeRoot });
  await run('dotnet', ['fable', 'src/Office.Formats', '--outDir', 'dist/office-formats', '--lang', 'JavaScript', '--noCache'], { cwd: officeRoot });
  await run('dotnet', ['fable', 'src/Spreadsheet', '--outDir', 'dist/spreadsheet', '--lang', 'JavaScript', '--noCache'], { cwd: officeRoot });
}

const distNodeModules = path.join(officeRoot, 'dist/node_modules');
try {
  await fs.lstat(distNodeModules);
} catch {
  await fs.symlink(path.join(runnerRoot, 'node_modules'), distNodeModules, 'dir');
}

const htmlPath = path.join(runnerRoot, '.capture-xlsx', 'index.html');
const modulePath = `${officeRoot}/dist/spreadsheet/Spreadsheet.js`;
const fixturePath = fixture;

await fs.writeFile(htmlPath, `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Spreadsheet Browser Capture</title>
  <style>
    html, body { margin: 0; background: #fff; font-family: system-ui, sans-serif; }
    #host { position: absolute; left: -10000px; top: 0; width: ${canvasWidth}px; height: ${canvasHeight}px; }
    #snapshot { display: block; width: ${canvasWidth}px; height: ${canvasHeight}px; background: #fff; }
  </style>
</head>
<body>
  <div id="host"></div>
  <canvas id="snapshot" width="${canvasWidth}" height="${canvasHeight}"></canvas>
  <script type="module">
    function describeError(error) {
      if (!error) return 'unknown error';
      if (error.stack) return error.stack;
      if (error.message) return error.message;
      try { return JSON.stringify(error); } catch { return String(error); }
    }

    try {
      const module = await import('/@fs/${modulePath}');
      const createSpreadsheet = module.default;
      const response = await fetch('/@fs/${fixturePath}');
      if (!response.ok) throw new Error('fixture fetch failed: ' + response.status);
      const buffer = await response.arrayBuffer();
      const host = document.getElementById('host');
      const canvas = document.getElementById('snapshot');
      const spreadsheet = createSpreadsheet(host);
      spreadsheet.Mount();
      spreadsheet.SetReadOnly(true);
      await spreadsheet.ImportXlsxFromBuffer(buffer, null);
      spreadsheet.RenderPortableCanvas(canvas);
      window.rbExcelCapture = spreadsheet;
      document.body.dataset.ready = '1';
    } catch (error) {
      const message = describeError(error);
      console.error('[capture-error] ' + message);
      document.body.dataset.error = message;
    }
  </script>
</body>
</html>
`);

const runnerRequire = createRequire(path.join(runnerRoot, 'package.json'));
const { chromium } = runnerRequire('playwright');
const viteBin = path.join(runnerRoot, 'node_modules/.bin/vite');
const server = spawn(viteBin, ['--host', '127.0.0.1', '--port', String(port)], {
  cwd: runnerRoot,
  stdio: ['ignore', 'pipe', 'pipe'],
});

let serverReady = false;
let serverOrigin = '';
server.stdout.on('data', chunk => {
  const text = chunk.toString();
  process.stdout.write(text);
  const match = text.match(/Local:\s+(https?:\/\/127\.0\.0\.1:\d+)\//);
  if (match) {
    serverOrigin = match[1];
    serverReady = true;
  }
});
server.stderr.on('data', chunk => process.stderr.write(chunk));

for (let i = 0; i < 120 && !serverReady; i++) {
  await new Promise(resolve => setTimeout(resolve, 250));
}
if (!serverReady) {
  server.kill('SIGKILL');
  throw new Error('vite server did not start');
}

try {
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage({ viewport: { width: viewportWidth, height: viewportHeight } });
    page.on('console', msg => {
      if (msg.type() !== 'debug') console.log(`[browser:${msg.type()}] ${msg.text()}`);
    });
    page.on('pageerror', error => {
      const message = error?.stack || error?.message || String(error) || 'unknown page error';
      console.error(`[browser:pageerror] ${message}`);
    });
    page.on('requestfailed', request => {
      console.error(`[browser:requestfailed] ${request.url()} ${request.failure()?.errorText ?? ''}`);
    });
    await page.goto(`${serverOrigin}/.capture-xlsx/index.html`);
    await page.waitForFunction(
      () => document.body.dataset.ready === '1' || !!document.body.dataset.error,
      null,
      { timeout: 30000 });
    const pageError = await page.evaluate(() => document.body.dataset.error || '');
    if (pageError) throw new Error(`browser xlsx capture failed: ${pageError}`);
    const modelInfo = await page.evaluate(() => {
      const model = window.rbExcelCapture.GetModel();
      return {
        rows: model.RowCount,
        cols: model.ColCount,
        cells: model.Cells?.size ?? model.Cells?.Count ?? null,
        richRuns: model.RichRuns?.size ?? model.RichRuns?.Count ?? null,
        paintItems: window.rbExcelCapture.PortablePaintItemCount(),
      };
    });
    console.log(`browser model: rows=${modelInfo.rows} cols=${modelInfo.cols} cells=${modelInfo.cells} richRuns=${modelInfo.richRuns} paintItems=${modelInfo.paintItems}`);
    await page.locator('#snapshot').screenshot({ path: outputPng });
    console.log(`browser xlsx capture ok: ${fixture} -> ${outputPng}`);
  } finally {
    await browser.close();
  }
} finally {
  server.kill('SIGKILL');
}
