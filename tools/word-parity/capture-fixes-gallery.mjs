#!/usr/bin/env node
import { createRequire } from 'node:module';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

const officeRoot = path.resolve(process.env.OFFICE_ROOT ?? process.cwd());
const runnerRequire = createRequire(path.join(officeRoot, 'tools/browser-runner/package.json'));
const { chromium } = runnerRequire('playwright');

const output = path.resolve(process.argv[2] ?? path.join(officeRoot, 'artifacts/word-parity-fixes-page-full.png'));
const html = path.resolve(process.argv[3] ?? path.join(officeRoot, 'docs/word-parity-fixes.html'));

async function capture(fullPage) {
  const browser = await chromium.launch({ headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1900, height: 1400 }, deviceScaleFactor: 1 });
    await page.goto(pathToFileURL(html).href, { waitUntil: 'load' });
    await page.screenshot({ path: output, fullPage });
  } finally {
    await browser.close().catch(() => {});
  }
}

try {
  await capture(true);
} catch (error) {
  const reason = String(error?.message ?? error).split('\n')[0];
  console.warn(`full-page screenshot failed, falling back to viewport: ${reason}`);
  await capture(false);
}

console.log(`gallery screenshot ok: ${html} -> ${output}`);
