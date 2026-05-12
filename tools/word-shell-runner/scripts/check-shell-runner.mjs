// End-to-end test for the OfficeUi dialog framework in the browser.
//
//   1. dotnet tool restore + dotnet fable -> compile Program.fs -> Program.js
//   2. vite dev server on 127.0.0.1:5180
//   3. playwright (chromium): open page, exercise each dialog, snapshot
//
// On success exits 0 and writes per-dialog screenshots + a summary JSON to
// ../../artifacts/shell-runner/.

import { chromium } from 'playwright';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import fs from 'node:fs/promises';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const outDir = path.resolve(root, '../../artifacts/shell-runner');
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
const server = spawn(viteBin, ['--host', '127.0.0.1', '--port', '5180'], {
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

const cases = [];

function assertEqual(actual, expected, label) {
    if (actual !== expected) {
        throw new Error(`${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
    }
}

try {
    const browser = await chromium.launch();
    const page = await browser.newPage({ viewport: { width: 800, height: 700 } });
    await page.goto('http://127.0.0.1:5180/');
    await page.waitForFunction(() => typeof window.rbDialog === 'object', null, { timeout: 15000 });

    async function openAndSnap(testId, snapshotName, fillFn, expectedResult) {
        await page.locator(`[data-testid="${testId}"]`).click();
        await page.waitForSelector(`.rb-dlg-card`, { timeout: 2000 });
        if (fillFn) await fillFn(page);
        await page.screenshot({ path: path.join(outDir, `${snapshotName}.png`), fullPage: true });
        // Click the primary button.
        await page.locator(`.rb-dlg-button-primary`).click();
        // Dialog should close.
        await page.waitForFunction(() => document.querySelector('.rb-dlg-card') === null, null, { timeout: 2000 });
        const result = await page.locator('[data-testid="result"]').textContent();
        cases.push({ testId, snapshotName, result });
        if (expectedResult !== undefined) assertEqual(result, expectedResult, snapshotName);
    }

    // Find / Replace — fills both fields, clicks Replace all (primary).
    await openAndSnap(
        'btn-find',
        'find-replace',
        async (p) => {
            await p.locator('[data-testid="dlg-field-find"]').fill('hello');
            await p.locator('[data-testid="dlg-field-replace"]').fill('world');
            await p.locator('[data-testid="dlg-field-case"]').check();
        },
        'replace:hello->world'
    );

    // Insert Table — defaults 2 x 3, click Insert.
    await openAndSnap(
        'btn-table',
        'insert-table',
        async (p) => {
            await p.locator('[data-testid="dlg-field-rows"]').fill('4');
            await p.locator('[data-testid="dlg-field-cols"]').fill('5');
        },
        'table:4x5'
    );

    // Insert Link — fill URL.
    await openAndSnap(
        'btn-link',
        'insert-link',
        async (p) => {
            await p.locator('[data-testid="dlg-field-url"]').fill('https://example.com');
        },
        'link:https://example.com'
    );

    // Page Setup — change orientation via the <select>, leave margins as-is.
    await openAndSnap(
        'btn-page-setup',
        'page-setup',
        async (p) => {
            await p.locator('[data-testid="dlg-field-orientation"]').selectOption('landscape');
        },
        'page:landscape ml=72 mr=72 mt=72 mb=72'
    );

    // Insert Image — paste a URL.
    await openAndSnap(
        'btn-image',
        'insert-image',
        async (p) => {
            await p.locator('[data-testid="dlg-field-src"]').fill('/img/cover.png');
        },
        'image:/img/cover.png'
    );

    await fs.writeFile(
        path.join(outDir, 'summary.json'),
        JSON.stringify({ ok: true, cases }, null, 2)
    );
    await browser.close();
    console.log(`shell runner ok: ${cases.length} dialog cases`);
} catch (err) {
    console.error('shell runner FAILED:', err);
    await fs.writeFile(
        path.join(outDir, 'summary.json'),
        JSON.stringify({ ok: false, cases, error: String(err) }, null, 2)
    );
    process.exitCode = 1;
} finally {
    server.kill('SIGKILL');
}
