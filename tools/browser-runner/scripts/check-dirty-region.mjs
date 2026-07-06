import { chromium } from 'playwright';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import fs from 'node:fs/promises';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const officeRoot = path.resolve(root, '../..');
const skipBuild = process.argv.includes('--skip-build');
const port = Number(process.env.RB_DIRTY_REGION_PORT ?? 5184);
const wordFixture = path.resolve(officeRoot, 'tests/test-documents-business/docx4j-invoice.docx');
const fontDir = path.resolve(process.env.RB_WORD_FONT_DIR ?? path.join(officeRoot, 'artifacts/fonts/windows'));

function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd: options.cwd ?? root, stdio: 'inherit', env: process.env });
    child.on('exit', code => code === 0 ? resolve() : reject(new Error(`${command} exited ${code}`)));
  });
}

function fsUrl(file) {
  return `/@fs${file}`;
}

async function fileExists(file) {
  try {
    await fs.access(file);
    return true;
  } catch {
    return false;
  }
}

async function startVite() {
  const viteBin = path.join(root, 'node_modules/.bin/vite');
  const server = spawn(viteBin, ['--host', '127.0.0.1', '--port', String(port)], {
    cwd: root,
    stdio: ['ignore', 'pipe', 'pipe']
  });

  let serverReady = false;
  let startupOutput = '';
  server.stdout.on('data', chunk => {
    const text = chunk.toString();
    startupOutput += text;
    process.stdout.write(text);
    if (text.includes('Local:')) serverReady = true;
  });
  server.stderr.on('data', chunk => {
    const text = chunk.toString();
    startupOutput += text;
    process.stderr.write(text);
  });

  for (let i = 0; i < 100 && !serverReady; i++) {
    await new Promise(resolve => setTimeout(resolve, 250));
  }
  if (!serverReady) {
    server.kill('SIGKILL');
    throw new Error(`vite server did not start:\n${startupOutput}`);
  }
  return server;
}

async function waitForRunner(page) {
  await page.waitForFunction(
    () => {
      const status = document.querySelector('#status')?.textContent || '';
      return window.rbWordRunner && (status === 'Ready' || status.startsWith('Loaded ') || status.startsWith('Load failed:'));
    },
    null,
    { timeout: 120000 }
  );
}

async function waitForLoaded(page) {
  await page.waitForFunction(
    () => {
      const status = document.querySelector('#status')?.textContent || '';
      return window.rbWordRunner && (status.startsWith('Loaded ') || status.startsWith('Load failed:'));
    },
    null,
    { timeout: 120000 }
  );
  const statusText = await page.locator('#status').textContent();
  if (!statusText?.startsWith('Loaded ')) {
    throw new Error(`Word fixture did not load: ${statusText}`);
  }
}

async function expectEquivalent(page, label, mutate) {
  await page.evaluate(mutate);
  const dirty = await page.evaluate(() => window.rbWordRunner.canvasToPngDataUrl());
  const full = await page.evaluate(() => {
    window.rbWordRunner.forceFullRender();
    return window.rbWordRunner.canvasToPngDataUrl();
  });
  if (dirty !== full) {
    throw new Error(`dirty-region PNG diverged after ${label}`);
  }
  console.log(`dirty-region ok: ${label}`);
}

async function installPageHelpers(page) {
  await page.evaluate(() => {
    window.rbDirtySelectLine = (position, collapsed, minSpan = 4) => {
      const lines = window.rbWordRunner.getVbaLayoutLines()
        .filter(line => line.BlockId && line.EndOffset >= line.StartOffset);
      if (!lines.length) throw new Error('no layout lines available');
      const index =
        position === 'top' ? 0 :
        position === 'bottom' ? lines.length - 1 :
        Math.floor(lines.length / 2);
      const line = lines[index];
      const start = line.StartOffset;
      const end = collapsed ? start : Math.max(start, Math.min(line.EndOffset, start + minSpan));
      window.rbWordRunner.selectFindMatch({ BlockId: line.BlockId, Start: start, End: end, Text: '' });
    };
  });
}


async function loadFixture(page) {
  const params = new URLSearchParams({
    doc: fsUrl(wordFixture),
    canvasWidth: '980',
    canvasHeight: '2200'
  });
  if (await fileExists(path.join(fontDir, 'calibri.ttf'))) {
    params.set('fontRoot', fsUrl(fontDir));
  }
  await page.goto(`http://127.0.0.1:${port}/?${params.toString()}`);
  await waitForLoaded(page);
  await installPageHelpers(page);
  await page.waitForTimeout(500);
  await page.evaluate(() => window.rbWordRunner.forceFullRender());
}

async function loadSynthetic(page) {
  await page.goto(`http://127.0.0.1:${port}/?canvasWidth=980&canvasHeight=2200`);
  await waitForRunner(page);
  await installPageHelpers(page);
  await page.evaluate(() => {
    const paragraphs = Array.from({ length: 180 }, (_, i) =>
      `${i}: alpha beta gamma delta epsilon invoice total paragraph text ${i % 9}`
    );
    window.rbWordRunner.setParagraphs(paragraphs);
    window.rbWordRunner.forceFullRender();
  });
}

if (!skipBuild) {
  await run('dotnet', ['tool', 'restore'], { cwd: officeRoot });
  await run('dotnet', ['fable', 'src/Word', '--outDir', 'dist/word', '--lang', 'JavaScript'], { cwd: officeRoot });
  await run('dotnet', ['fable', '.', '--outDir', 'src', '--lang', 'JavaScript'], { cwd: root });
}

const server = await startVite();
let browser;
try {
  browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1200, height: 2400 } });
  page.on('console', msg => {
    if (msg.type() !== 'debug') console.log(`[browser:${msg.type()}] ${msg.text()}`);
  });
  page.on('pageerror', error => {
    console.error(`[browser:pageerror] ${error?.stack || error?.message || String(error)}`);
  });

  await loadFixture(page);
  await expectEquivalent(page, 'docx top insert', () => {
    window.rbDirtySelectLine('top', true);
    window.rbWordRunner.insertTextAt('[top-dirty] ');
  });
  await expectEquivalent(page, 'docx middle insert', () => {
    window.rbDirtySelectLine('middle', true);
    window.rbWordRunner.insertTextAt('[middle-dirty] ');
  });
  await expectEquivalent(page, 'docx selection change only', () => {
    window.rbDirtySelectLine('middle', false, 8);
  });
  await expectEquivalent(page, 'docx bottom insert', () => {
    window.rbDirtySelectLine('bottom', true);
    window.rbWordRunner.insertTextAt(' [bottom-dirty]');
  });

  await loadSynthetic(page);
  await expectEquivalent(page, 'synthetic top insert', () => {
    window.rbDirtySelectLine('top', true);
    window.rbWordRunner.insertTextAt('top ');
  });
  await expectEquivalent(page, 'synthetic selection change only', () => {
    window.rbDirtySelectLine('middle', false, 12);
  });
  await expectEquivalent(page, 'synthetic wrapping insert', () => {
    window.rbDirtySelectLine('middle', true);
    window.rbWordRunner.insertTextAt(' wrap '.repeat(90));
  });
  await expectEquivalent(page, 'synthetic bottom insert', () => {
    window.rbDirtySelectLine('bottom', true);
    window.rbWordRunner.insertTextAt(' bottom');
  });
  await expectEquivalent(page, 'synthetic replace all', () => {
    const count = window.rbWordRunner.replaceAllText('gamma', 'GAMMA-dirty', false);
    if (count <= 0) throw new Error('replaceAllText did not replace synthetic text');
  });

  console.log('dirty-region pixel equivalence passed');
} finally {
  if (browser) await browser.close();
  server.kill('SIGKILL');
}
