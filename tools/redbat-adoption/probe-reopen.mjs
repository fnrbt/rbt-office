import { launch, newPage, open } from "./lib.mjs";
const { browser, context } = await launch({ viewport: { width: 1100, height: 520 } });
const page = await newPage(context);
await open(page, "word-editor", process.argv[2]);
await page.waitForSelector('[data-testid="word-toolbar"]'); await page.waitForTimeout(1500);
await page.evaluate((t) => { const w = window.__word, e = window.__editors[0]; const f = w.findText(e, t, true); if (f.length) w.selectFindMatch(e, f[0]); }, process.argv[3]);
await page.waitForTimeout(500);
await page.screenshot({ path: process.argv[4] });
await browser.close();
