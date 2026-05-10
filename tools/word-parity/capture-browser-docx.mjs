#!/usr/bin/env node
import { spawn } from 'node:child_process';
import path from 'node:path';
import fs from 'node:fs/promises';
import { createRequire } from 'node:module';

const officeRoot = path.resolve(process.env.OFFICE_ROOT ?? process.cwd());
const runnerRoot = path.join(officeRoot, 'tools/browser-runner');
const fixture = path.resolve(process.argv[2] ?? '');
const outputPng = path.resolve(process.argv[3] ?? '');
const outputJson = process.argv[4] ? path.resolve(process.argv[4]) : null;
const port = Number(process.env.BROWSER_CAPTURE_PORT ?? 5181);
const fontDir = path.resolve(process.env.RB_WORD_FONT_DIR ?? path.join(officeRoot, 'artifacts/fonts/windows'));
const layoutMode = process.env.BROWSER_CAPTURE_LAYOUT ?? 'segments';

if (!process.argv[2] || !process.argv[3]) {
  console.error('Usage: OFFICE_ROOT=/path/to/office node tools/word-parity/capture-browser-docx.mjs input.docx output.png [layout.json]');
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

try {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1100, height: 1800 } });
  const docUrl = `/@fs/${servedFixture}`;
  const params = new URLSearchParams({ doc: docUrl });
  if (await fileExists(path.join(fontDir, 'calibri.ttf'))) {
    params.set('fontRoot', `/@fs/${fontDir}`);
  }
  await page.goto(`http://127.0.0.1:${port}/?${params.toString()}`);
  await page.waitForFunction(
    () => window.rbWordRunner && document.querySelector('#status')?.textContent?.startsWith('Loaded '),
    null,
    { timeout: 30000 }
  );
  await page.evaluate(() => window.rbWordRunner.clearSelection());
  if (outputJson) {
    const layout = layoutMode === 'lines'
      ? await page.evaluate(() => window.rbWordRunner.getLayoutLines())
      : await page.evaluate(() => window.rbWordRunner.getLayoutSegments());
    await fs.writeFile(outputJson, JSON.stringify(layout, null, 2));
  }
  await page.locator('#word-canvas').screenshot({ path: outputPng });
  await browser.close();
  console.log(`browser capture ok: ${fixture} -> ${outputPng}`);
} finally {
  server.kill('SIGKILL');
}
