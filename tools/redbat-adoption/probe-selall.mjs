import { launch, newPage, open, S } from "./lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
const ev = (fn, arg) => page.evaluate(fn, arg);
await open(page, "word-editor", `${S}/long-generated.docx`);
await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 });
await page.waitForTimeout(500);
await ev(() => window.__word.getInputElement(window.__editors[0]).focus());
const raf = () => ev(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
for (let i = 0; i < 3; i++) {
  await page.keyboard.press("Control+Home"); await raf();
  let t = Date.now(); await page.keyboard.press("Control+a"); await raf(); const wall = Date.now() - t;
  const parts = await ev(() => { const w = window.__word, e = window.__editors[0]; const time = (fn) => { const t = performance.now(); fn(); return Math.round(performance.now() - t); }; return { getSelectionFormat: time(() => w.getSelectionFormat(e)), getCaretContext: time(() => w.getCaretContext(e)), getCaretRect: time(() => w.getCaretRect(e)), getDocumentSize: time(() => w.getDocumentSize(e)), repaint: time(() => w.setViewport(e, w.getViewport(e))) }; });
  await page.keyboard.press("Control+Home"); await raf();
  const engine = await ev(() => { const w = window.__word, e = window.__editors[0]; const t = performance.now(); w.selectAll(e); return Math.round(performance.now() - t); });
  console.log(`select all: Ctrl+A to painted ${wall} ms; engine selectAll (with the host's change listener) ${engine} ms; with the whole document selected ${JSON.stringify(parts)}`);
}
await browser.close();
