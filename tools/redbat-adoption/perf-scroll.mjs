// Scroll frames on the Data sheet of the dashboard (76,931 rows): wheel steps, one per frame. usage: PORT=.. node perf-scroll.mjs
import { launch, newPage, open } from "./lib.mjs";
import { g, sel, ready, frames } from "./ui-lib.mjs";
const T = "/home/dave/Dev/rbt-office/tests";
const { browser, context } = await launch();
for (let run = 1; run <= 2; run++) {
  const page = await newPage(context);
  await open(page, "sheet-editor", `${T}/real-world/dashboard-ecommerce-sales.xlsx`); await ready(page);
  const idx = await g(page, () => window.__grids[0].GetSheetNames().indexOf("Data"));
  await sel(page, `sheet-tab-${idx}`).click(); await frames(page); await page.waitForTimeout(500);
  const r = await g(page, async () => {
    const sc = document.querySelector(".spreadsheet-container");
    const long = []; new PerformanceObserver((l) => { for (const e of l.getEntries()) long.push(Math.round(e.duration)); }).observe({ entryTypes: ["longtask"] });
    const work = [], gaps = []; let last = performance.now();
    for (let i = 0; i < 240; i++) {
      const t = performance.now();
      sc.scrollTop += i < 120 ? 380 : 4200; sc.dispatchEvent(new Event("scroll"));
      work.push(performance.now() - t);
      await new Promise((r) => requestAnimationFrame(r));
      const now = performance.now(); gaps.push(now - last); last = now;
    }
    const q = (a, p) => { const s = [...a].sort((x, y) => x - y); return Math.round(s[Math.min(s.length - 1, Math.floor(s.length * p))] * 10) / 10; };
    const row = Number([...document.querySelectorAll(".spreadsheet-row-header")].map((e) => e.dataset.row).filter(Boolean)[3] ?? -1);
    return { steps: 240, frameGap: { median: q(gaps, 0.5), p95: q(gaps, 0.95), worst: q(gaps, 1) }, syncWork: { median: q(work, 0.5), p95: q(work, 0.95), worst: q(work, 1) }, over33ms: gaps.filter((x) => x > 33.4).length, longTasks: long, scrollTop: sc.scrollTop, rowInView: row, rows: window.__grids[0].GetModel().RowCount };
  });
  console.log(`scroll run ${run}: ${JSON.stringify(r)}`);
  await page.close();
}
await browser.close();
