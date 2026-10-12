// Production build: which chunks a preview requests.
import { launch, newPage, open, T, SHOTS, check, summary } from "./lib.mjs";
import fs from "node:fs";
const { browser } = await launch();
const js = (page) => page.requests.filter((r) => r.startsWith("/assets/") && r.endsWith(".js")).map((r) => r.slice(8));
const sizeOf = (list) => list.reduce((n, f) => n + fs.statSync(`${process.env.BUILD}/assets/${f}`).size, 0);
for (const [component, label] of [["word-preview", "preview component"], ["chat-viewer", "chat attachment viewer"]]) {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });
  const page = await newPage(ctx);
  const t = Date.now();
  await open(page, component, `${T}/test-documents-wild/docx4j-getting-started.docx`, { params: { mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } });
  await page.waitForSelector(".rb-file-viewer-canvas", { timeout: 120000 });
  await page.waitForTimeout(800);
  const list = js(page);
  const editorChunks = list.filter((f) => /^(Word-|WordEditor-|editor-|OfficeEditorFrame-)/.test(f));
  check(`${label} (production chunks): no editor chunk requested`, editorChunks.length === 0 && list.some((f) => f.startsWith("WordViewer-")) && list.some((f) => f.startsWith("browser-")), { editorChunks, office: list.filter((f) => /Word|browser|viewer|Main|zip/i.test(f)) });
  const painted = await page.evaluate(() => [...document.querySelectorAll(".rb-file-viewer-canvas")].filter((c) => c.width >= 800).length);
  check(`${label}: pages painted with the shaped text stack`, painted >= 1 && (await page.evaluate(() => window.RedBatWordTextStack.loadedFaceFiles().length)) > 3, painted);
  console.log(`  ${label}: ${Date.now() - t} ms; JS requested ${list.length} files, ${(sizeOf(list) / 1000).toFixed(0)} kB`);
  if (component === "chat-viewer") await page.screenshot({ path: `${SHOTS}/prod-word-preview.png` });
  await ctx.close();
}
{
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });
  const page = await newPage(ctx);
  await open(page, "word-editor", `${T}/test-documents-wild/docx4j-getting-started.docx`);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 120000 });
  await page.waitForTimeout(500);
  const list = js(page);
  const r = await page.evaluate(() => ({ pages: window.__word.getPageCount(window.__editors[0]), stats: window.__wordModule.getRenderStats(window.__editors[0]) }));
  check("editor (production chunks): loads the editor chunk and the text stack's editor part, 54 pages", list.some((f) => f.startsWith("Word-")) && list.some((f) => f.startsWith("editor-")) && r.pages === 54 && r.stats.PainterVersion >= 1, { r, office: list.filter((f) => /Word|browser|editor|viewer|Main|zip/i.test(f)) });
  const before = r.stats.FrameBuilds;
  await page.evaluate(() => { const sc = document.querySelector('[data-testid="word-scroll"]'); for (let i = 0; i < 30; i++) { sc.scrollTop += 900; sc.dispatchEvent(new Event("scroll")); } });
  check("editor (production build): scrolling builds no frame", (await page.evaluate(() => window.__wordModule.getRenderStats(window.__editors[0]).FrameBuilds)) === before);
  await ctx.close();
}
{
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });
  const page = await newPage(ctx);
  await open(page, "sheet-preview", `${T}/real-world/milk-street-office.xlsx`);
  await page.waitForFunction(() => document.querySelector('[data-testid="spreadsheet-preview"]')?.style.visibility === "visible", null, { timeout: 120000 });
  const list = js(page);
  check("spreadsheet preview (production chunks): no editor chunk", !list.some((f) => /^(SpreadsheetEditor-|OfficeEditorFrame-|Word)/.test(f)), list.filter((f) => /sheet|zip|Word/i.test(f)));
  await ctx.close();
}
summary();
await browser.close();
