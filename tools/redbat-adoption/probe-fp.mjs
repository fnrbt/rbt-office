import { launch, newPage, open, T, S } from "./lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "word-editor", process.argv[2]);
await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors.at(-1)), null, { timeout: 300000 });
const r = await page.evaluate(() => { const m = window.__wordModule, e = window.__editors.at(-1); const fp = m.getLayoutFingerprint(e); const pkg = m.exportDocumentPackage(e); const blocks = pkg.Blocks ?? pkg.blocks; const keys = Object.keys(pkg); const b0 = blocks?.find((b) => JSON.stringify(b).includes("Rows")) ; return { fpLen: fp.length, fp0: fp.slice(0, 3), fpKeys: Object.keys(fp[0] ?? {}), keys, nblocks: blocks?.length, blockKeys: blocks ? Object.keys(blocks[0]) : null, tableSample: b0 ? JSON.stringify(b0).slice(0, 1500) : null, tg: (() => { try { return JSON.stringify(m.getTableGeometry(e)).slice(0, 600); } catch (x) { return String(x); } })() }; });
console.log(JSON.stringify(r, null, 1).slice(0, 5000));
await browser.close();
