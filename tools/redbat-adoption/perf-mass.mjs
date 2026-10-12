// Select all + bold / undo on the long document, the view at the top and in the middle; what the host makes of the deferred layout.
import { launch, newPage, open, T, S, SHOTS } from "./lib.mjs";
const { browser } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const file = `${S}/long-generated.docx`;
const state = (page) => ev(page, () => { const w = window.__word, e = window.__editors[0]; const s = w.getDocumentSize(e); return { complete: w.isLayoutComplete(e), pages: s.PageCount, sub: document.querySelector('[data-testid="office-editor-status"]')?.parentElement?.innerText.replace(/\n/g, " | ") }; });
// one key through the browser: time until the keystroke's task and the two frames after it are over
const timed = async (page, act) => {
  await ev(page, () => { window.__lt = []; window.__lto?.disconnect(); window.__lto = new PerformanceObserver((l) => { for (const e of l.getEntries()) window.__lt.push(Math.round(e.duration)); }); window.__lto.observe({ entryTypes: ["longtask"] }); });
  const t = Date.now(); await act(); await ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
  const ms = Date.now() - t; const st = await state(page);
  await page.waitForTimeout(60);
  return `${ms} ms; longest task ${Math.max(0, ...(await ev(page, () => window.__lt)))} ms; right after: layoutComplete=${st.complete} pages=${st.pages} "${st.sub}"`;
};
for (const where of ["top", "middle"]) {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });
  const page = await newPage(ctx);
  await open(page, "word-editor", file);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await ev(page, () => window.__word.whenLayoutComplete(window.__editors[0]));
  await page.waitForFunction(() => window.__word.isLoadComplete(window.__editors[0]) !== false, null, { timeout: 300000 });
  await ev(page, () => window.__word.whenLayoutComplete(window.__editors[0]));
  await page.waitForTimeout(500);
  if (where === "middle") { await ev(page, () => { const s = document.querySelector('[data-testid="word-scroll"]'); s.scrollTop = s.scrollHeight / 2; }); await page.waitForTimeout(500); }
  await ev(page, () => window.__word.getInputElement(window.__editors[0]).focus());
  console.log(`${where}: before`, JSON.stringify(await state(page)));
  console.log(`${where}: select all (Ctrl+A): ${await timed(page, () => page.keyboard.press("Control+a"))}`);
  const top0 = await ev(page, () => document.querySelector('[data-testid="word-scroll"]').scrollTop);
  console.log(`${where}: bold (Ctrl+B): ${await timed(page, () => page.keyboard.press("Control+b"))}`);
  const t = Date.now(); await ev(page, () => window.__word.whenLayoutComplete(window.__editors[0])); console.log(`${where}:   layout complete ${Date.now() - t} ms later`, JSON.stringify(await state(page)), "scrollTop moved by", (await ev(page, () => document.querySelector('[data-testid="word-scroll"]').scrollTop)) - top0);
  console.log(`${where}: undo (Ctrl+Z): ${await timed(page, () => page.keyboard.press("Control+z"))}`);
  await ev(page, () => window.__word.whenLayoutComplete(window.__editors[0]));
  await page.keyboard.press("Control+a");
  console.log(`${where}: bold (toolbar button): ${await timed(page, () => sel(page, "word-bold").click())}`);
  await ev(page, () => window.__word.whenLayoutComplete(window.__editors[0]));
  console.log(`${where}: undo (toolbar button): ${await timed(page, () => sel(page, "word-undo").click())}`);
  await ev(page, () => window.__word.whenLayoutComplete(window.__editors[0]));
  console.log(`${where}: after all`, JSON.stringify(await state(page)));
  await ctx.close();
}
await browser.close();
