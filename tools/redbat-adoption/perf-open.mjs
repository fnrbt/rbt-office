// Open-to-first-paint in the host, measured in the page: from the document's bytes having arrived
// (the fixture's responseEnd) to ink on the canvas; when the rest is read and laid out; longest task.
// usage: node perf-open.mjs <file> [runs]
import { launch, newPage, open, S } from "./lib.mjs";
const file = process.argv[2] || `${S}/long-generated.docx`;
const runs = Number(process.argv[3] ?? 3);
const { browser } = await launch();
const probe = () => {
  window.__ink = null; window.__lt = [];
  try { new PerformanceObserver((l) => { for (const e of l.getEntries()) window.__lt.push([Math.round(e.startTime), Math.round(e.duration)]); }).observe({ entryTypes: ["longtask"] }); } catch {}
  const tick = () => {
    if (window.__ink == null) {
      const c = document.querySelector('[data-testid="word-canvas"]') ?? [...document.querySelectorAll(".rb-file-viewer-canvas")].find((c) => c.width >= 600);
      if (c && c.width > 0 && c.height > 0) {
        try {
          const d = c.getContext("2d").getImageData(0, 0, c.width, Math.min(c.height, 400)).data; let n = 0;
          for (let i = 0; i < d.length; i += 16) if (d[i + 3] > 0 && d[i] < 90 && d[i + 1] < 90 && d[i + 2] < 90) n++;
          if (n > 100) window.__ink = performance.now();
        } catch {}
      }
    }
    if (window.__usable == null && document.querySelector('[data-testid="word-toolbar"]') && !document.querySelector(".rb-office-editor-cover")) window.__usable = performance.now();
    requestAnimationFrame(tick);
  };
  requestAnimationFrame(tick);
};
const med = (a) => a.slice().sort((x, y) => x - y)[a.length >> 1];
for (const component of ["word-editor", "chat-viewer"]) {
  const rows = [];
  for (let run = 0; run < runs; run++) {
    const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } }); const page = await newPage(ctx);
    await page.addInitScript(probe);
    await open(page, component, file, component === "chat-viewer" ? { params: { mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } } : { params: { hookWord: "1" } });
    await page.waitForFunction((editor) => window.__ink != null && (!editor || window.__usable != null), component === "word-editor", { timeout: 300000, polling: 20 });
    const first = await page.evaluate(() => { const r = performance.getEntriesByType("resource").find((e) => e.name.includes("/api/fixture/")); return { bytes: r.responseEnd, ink: Math.max(window.__ink, window.__usable ?? 0), canvasInk: window.__ink, usable: window.__usable ?? null, loadComplete: window.__editors?.[0] ? window.__word.isLoadComplete?.(window.__editors[0]) : null, sub: (document.querySelector('[data-testid="office-editor-status"]') ?? document.querySelector(".rb-file-viewer-header-copy .rb-app-subtitle"))?.innerText?.replace(/\n/g, " | ") }; });
    // the whole document: laid out
    if (component === "word-editor") await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000, polling: 20 });
    else await page.waitForFunction(() => /^\d+ pages?$/.test(document.querySelector(".rb-file-viewer-header-copy .rb-app-subtitle")?.textContent.trim() ?? ""), null, { timeout: 300000, polling: 20 });
    const done = await page.evaluate(() => ({ now: performance.now(), lt: window.__lt, sub: (document.querySelector('[data-testid="office-editor-status"]') ?? document.querySelector(".rb-file-viewer-header-copy .rb-app-subtitle"))?.innerText?.replace(/\n/g, " | ") }));
    const after = done.lt.filter(([s]) => s >= first.bytes);
    const row = { paint: Math.round(first.ink - first.bytes), canvas: Math.round(first.canvasInk - first.bytes), whole: Math.round(done.now - first.bytes), longest: Math.max(0, ...after.map(([, d]) => d)), longestAfterPaint: Math.max(0, ...after.filter(([s]) => s >= first.ink).map(([, d]) => d)) };
    rows.push(row);
    console.log(`${component} run ${run + 1}: bytes -> shown ${row.paint} ms (canvas painted ${row.canvas} ms) ("${first.sub}", load complete at paint: ${first.loadComplete}); bytes -> whole document laid out ${row.whole} ms ("${done.sub}"); longest task after the bytes ${row.longest} ms, after first paint ${row.longestAfterPaint} ms`);
    await ctx.close();
  }
  console.log(`${component} median of ${runs}: first paint ${med(rows.map((r) => r.paint))} ms, whole ${med(rows.map((r) => r.whole))} ms, longest task ${med(rows.map((r) => r.longest))} ms, longest after paint ${med(rows.map((r) => r.longestAfterPaint))} ms`);
}
await browser.close();
