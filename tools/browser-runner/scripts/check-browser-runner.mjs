import { chromium } from 'playwright';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import fs from 'node:fs/promises';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const fixtureArg = process.argv[2] ?? '../../tests/test-documents-features/test-formatting.docx';
const outputStem = process.argv[3] ?? path.basename(fixtureArg, path.extname(fixtureArg));
const fixture = path.resolve(root, fixtureArg);
const outDir = path.resolve(root, '../../artifacts/browser');
await fs.mkdir(outDir, { recursive: true });

function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd: root, stdio: 'inherit', ...options });
    child.on('exit', code => code === 0 ? resolve() : reject(new Error(`${command} exited ${code}`)));
  });
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
  await page.goto(`http://127.0.0.1:5179/?doc=${encodeURIComponent(docUrl)}`);
  await page.waitForFunction(
    () => window.rbWordRunner && document.querySelector('#status')?.textContent?.startsWith('Loaded '),
    null,
    { timeout: 30000 }
  );
  const segments = await page.evaluate(() => window.rbWordRunner.getLayoutSegments());
  await fs.writeFile(path.join(outDir, `${outputStem}-layout.json`), JSON.stringify(segments, null, 2));
  await page.locator('#word-canvas').screenshot({ path: path.join(outDir, `${outputStem}.png`) });
  await browser.close();
  console.log(`browser runner ok: ${segments.length} layout segments`);
} finally {
  server.kill('SIGKILL');
}
