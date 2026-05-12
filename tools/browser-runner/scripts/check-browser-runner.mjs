import { chromium } from 'playwright';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import fs from 'node:fs/promises';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const officeRoot = path.resolve(root, '../..');
const fixtureArg = process.argv[2] ?? '../../tests/test-documents-features/test-formatting.docx';
const outputStem = process.argv[3] ?? path.basename(fixtureArg, path.extname(fixtureArg));
const fixture = path.resolve(root, fixtureArg);
const outDir = path.resolve(root, '../../artifacts/browser');
const fontDir = path.resolve(process.env.RB_WORD_FONT_DIR ?? path.join(officeRoot, 'artifacts/fonts/windows'));
await fs.mkdir(outDir, { recursive: true });

function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd: root, stdio: 'inherit', ...options });
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

await run('dotnet', ['tool', 'restore']);
await run('dotnet', ['fable', '.', '--outDir', 'src', '--lang', 'JavaScript']);

const viteBin = path.join(root, 'node_modules/.bin/vite');
const server = spawn(viteBin, ['--host', '127.0.0.1', '--port', '5179'], {
  cwd: root,
  stdio: ['ignore', 'pipe', 'pipe']
});

let serverReady = false;
server.stdout.on('data', chunk => {
  const text = chunk.toString();
  process.stdout.write(text);
  if (text.includes('Local:')) serverReady = true;
});
server.stderr.on('data', chunk => process.stderr.write(chunk));

for (let i = 0; i < 80 && !serverReady; i++) {
  await new Promise(resolve => setTimeout(resolve, 250));
}
if (!serverReady) {
  server.kill();
  throw new Error('vite server did not start');
}

try {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1100, height: 1800 } });
  const docUrl = `/@fs/${fixture}`;
  const params = new URLSearchParams({ doc: docUrl });
  if (await fileExists(path.join(fontDir, 'calibri.ttf'))) {
    params.set('fontRoot', `/@fs/${fontDir}`);
  }
  await page.goto(`http://127.0.0.1:5179/?${params.toString()}`);
  await page.waitForFunction(
    () => window.rbWordRunner && document.querySelector('#status')?.textContent?.startsWith('Loaded '),
    null,
    { timeout: 30000 }
  );
  await page.evaluate(() => window.rbWordRunner.clearSelection());
  const segments = await page.evaluate(() => window.rbWordRunner.getLayoutSegments());
  const vbaLines = await page.evaluate(() => window.rbWordRunner.getVbaLayoutLines());
  if (!vbaLines.length || typeof vbaLines[0].HorizontalPositionRelativeToPage !== 'number') {
    throw new Error('VBA-style layout lines are missing Word-compatible position fields');
  }
  const info = await page.evaluate(line =>
    window.rbWordRunner.getRangeInformation(line.BlockId, line.StartOffset, line.EndOffset),
    vbaLines[0]
  );
  if (!info || typeof info.wdActiveEndPageNumber !== 'number') {
    throw new Error('Range.Information-compatible query did not return page information');
  }
  await fs.writeFile(path.join(outDir, `${outputStem}-layout.json`), JSON.stringify(segments, null, 2));
  await page.locator('#word-canvas').screenshot({ path: path.join(outDir, `${outputStem}.png`) });
  await browser.close();
  console.log(`browser runner ok: ${segments.length} layout segments, ${vbaLines.length} VBA-style lines`);
} finally {
  server.kill('SIGKILL');
}
