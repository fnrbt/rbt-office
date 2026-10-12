import { launch, newPage, open, T, SHOTS, check, summary, fontStats } from "./lib.mjs";
const S = "/tmp/redbat-adoption";
const docs = [`${T}/test-documents/sample-formatted.docx`, `${T}/test-documents-wild/docx4j-getting-started.docx`, `${S}/long-generated.docx`, `${T}/repros/word/landscape-section.docx`, `${T}/test-documents-wild/docx4j-chart.docx`];
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
// Each document in a fresh browser context (cold font cache per context is not guaranteed: the HTTP cache is per context).
for (const file of docs) {
  const name = file.split("/").pop();
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });
  const page = await newPage(ctx);
  fontStats.bytes = 0; fontStats.files = [];
  let t = Date.now();
  await open(page, "word-preview", file);
  await page.waitForFunction(() => window.__wordPages > 0, null, { timeout: 300000 });
  const previewMs = Date.now() - t;
  const previewBytes = fontStats.bytes, previewFiles = fontStats.files.length;
  const preview = await ev(page, () => ({ pages: window.__wordPages, shaped: window.__wordLoaded.shapedText, frames: window.__wordLoaded.pages.map((p) => [Math.round(p.top), Math.round(p.width), Math.round(p.height)]), families: window.RedBatWordTextStack.loadedFamilyNames() }));
  await page.waitForTimeout(500);
  await page.screenshot({ path: `${SHOTS}/word-preview-${name}.png` });
  // a second preview in the same session fetches no font again
  fontStats.bytes = 0; fontStats.files = [];
  await ev(page, () => window.__setOpen(false));
  await ev(page, () => { window.__wordPages = 0; window.__setOpen(true); });
  await page.waitForFunction(() => window.__wordPages > 0, null, { timeout: 300000 });
  const again = fontStats.bytes;
  // the editor in the same session
  fontStats.bytes = 0; fontStats.files = [];
  const page2 = await newPage(ctx);
  t = Date.now();
  await open(page2, "word-editor", file);
  await page2.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  const editorMs = Date.now() - t;
  // (0.1.10: the package is read behind the editor, and the faces its later blocks need are loaded when it is read.)
  await ev(page2, async () => { await window.__word.whenLoadComplete(window.__editors[0]); await new Promise((r) => { const f = () => (window.__word.getPendingFonts(window.__editors[0]).length === 0 ? r() : setTimeout(f, 20)); f(); }); });
  const editor = await ev(page2, () => { const w = window.__word, e = window.__editors[0]; return { pages: w.getPageCount(e), frames: w.getPageFrames(e).map((p) => [Math.round(p.Top), Math.round(p.Width), Math.round(p.Height)]), x: w.getPageFrames(e).map((p) => Math.round(p.X)), families: window.RedBatWordTextStack.loadedFamilyNames() }; });
  check(`${name}: preview and editor have the same page count`, preview.pages === editor.pages && preview.shaped, { preview: preview.pages, editor: editor.pages });
  check(`${name}: every page starts at the same place and has the same size`, JSON.stringify(preview.frames) === JSON.stringify(editor.frames), { firstDiff: preview.frames.findIndex((f, i) => JSON.stringify(f) !== JSON.stringify(editor.frames[i])) });
  check(`${name}: the second preview fetched no font`, again === 0, again);
  console.log(`  ${name}: pages ${preview.pages}; preview ${previewMs} ms, fonts ${(previewBytes / 1048576).toFixed(1)} MB in ${previewFiles} files (${preview.families.join(", ")}); editor ${editorMs} ms`);
  if (name.startsWith("landscape")) {
    const xs = [...new Set(editor.x)];
    check("landscape-section: the engine centres the narrower pages (X)", xs.length > 1 && Math.min(...xs) === 0, { x: editor.x, widths: editor.frames.map((f) => f[1]) });
    // centred on screen: the portrait page's middle is the landscape page's middle
    await page2.screenshot({ path: `${SHOTS}/word-editor-landscape.png` });
    await page2.evaluate(() => { const s = document.querySelector('[data-testid="word-scroll"]'); s.scrollTop = s.scrollHeight / 2 - 300; });
    await page2.waitForTimeout(300);
    await page2.screenshot({ path: `${SHOTS}/word-editor-landscape-mid.png` });
  }
  if (name.includes("chart")) await page2.screenshot({ path: `${SHOTS}/word-editor-chart.png` });
  await ctx.close();
}
// browser-metrics pagination for reference (what the preview did before)
{
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });
  const page = await newPage(ctx);
  await open(page, "word-preview", docs[2]);
  await page.waitForFunction(() => window.__wordPages > 0, null, { timeout: 300000 });
  const pages = await ev(page, async () => { const v = window.__wordLoaded.viewer; v.clearTextMeasurer(); return v.getPageCount(); });
  console.log(`  long-generated.docx with the browser's own metrics (the old preview): ${pages} pages`);
  await ctx.close();
}
summary();
await browser.close();
