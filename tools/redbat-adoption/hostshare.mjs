// What the host's own calls cost per keystroke / caret move (522 pages, real text stack).
import { launch, newPage, open, S } from "./lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
const t0 = Date.now();
await open(page, "word-editor", `${S}/long-generated.docx`);
await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
console.log("open ms", Date.now() - t0);
await page.waitForTimeout(500);
await page.evaluate(() => {
  const w = window.__word, e = window.__editors[0]; const lines = w.getLayoutLines(e); const l = lines[lines.length >> 1]; w.setSelection(e, l.BlockId, 1, l.BlockId, 1); w.getInputElement(e).focus();
  window.__acc = {};
  for (const name of ["setViewport", "getDocumentSize", "getCaretRect", "getSelectionFormat", "getCaretContext", "getHyperlink"]) {
    const f = w[name]; w[name] = (...a) => { const t = performance.now(); try { return f(...a); } finally { const x = (window.__acc[name] ??= { n: 0, ms: 0 }); x.n++; x.ms += performance.now() - t; } };
  }
  // painter calls: count fills on the editor canvas context
  const c = document.querySelector('[data-testid="word-canvas"]').getContext("2d"); const fill = c.fill.bind(c); window.__fills = 0; c.fill = (...a) => { window.__fills++; return fill(...a); };
  window.__keys = []; window.__eto = new PerformanceObserver((l) => { for (const en of l.getEntries()) if (en.name === "keydown") window.__keys.push(Math.round(en.processingEnd - en.processingStart)); }); window.__eto.observe({ type: "event", durationThreshold: 16 });
});
await page.waitForTimeout(500);
const run = async (label, fn, n) => {
  await page.evaluate(() => { window.__acc = {}; window.__fills = 0; window.__keys = []; });
  const t = Date.now(); await fn(); await page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))); const ms = Date.now() - t;
  await page.waitForTimeout(300);
  const r = await page.evaluate(() => ({ acc: window.__acc, fills: window.__fills, keys: window.__keys }));
  const host = Object.entries(r.acc).map(([k, v]) => `${k} ${v.n}x ${(v.ms / n).toFixed(2)} ms/key`).join(", ");
  const keys = r.keys.sort((a, b) => a - b);
  console.log(`${label}: ${(ms / n).toFixed(1)} ms/key wall; keydown handler (>=16 ms only, ${keys.length}/${n}) median ${keys[keys.length >> 1] ?? "<16"} ms; host calls per key: ${host}; text runs painted per key ${(r.fills / n).toFixed(0)}`);
};
await run("typing 40 characters", () => page.keyboard.type("The quick brown fox jumps over the lazy "), 40);
await run("30 x ArrowRight", async () => { for (let i = 0; i < 30; i++) await page.keyboard.press("ArrowRight"); }, 30);
await run("30 x ArrowDown", async () => { for (let i = 0; i < 30; i++) await page.keyboard.press("ArrowDown"); }, 30);
await run("30 x ArrowUp", async () => { for (let i = 0; i < 30; i++) await page.keyboard.press("ArrowUp"); }, 30);
await browser.close();
