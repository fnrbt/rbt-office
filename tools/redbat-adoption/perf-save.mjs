// The host's save of the dashboard workbook after one edit: from the click to the upload request leaving
// the page (the host's work + ExportXlsx), against ExportXlsx called alone on the same state.
import { launch, newPage, open, T, waitUploads } from "./lib.mjs";
const { browser } = await launch();
const med = (a) => a.slice().sort((x, y) => x - y)[a.length >> 1];
const rows = [];
for (let run = 0; run < 3; run++) {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } }); const page = await newPage(ctx);
  await page.addInitScript(() => { const of = window.fetch; window.fetch = (...a) => { if (String(a[0]).includes("/api/upload")) window.__uploadAt = performance.now(); return of(...a); }; });
  await open(page, "sheet-editor", `${T}/real-world/dashboard-sales-analysis.xlsx`);
  await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 180000 }); await page.waitForTimeout(400);
  const opened = await page.evaluate(() => { const r = performance.getEntriesByType("resource").find((e) => e.name.includes("/api/fixture/")); return r.responseEnd; });
  await page.evaluate(() => window.__grids[0].EnterText(60, 30, "edit"));
  // ExportXlsx alone first (nothing is cached between exports of an edited workbook? measured both ways below)
  const alone1 = await page.evaluate(async () => { const t = performance.now(); const b = await window.__grids[0].ExportXlsx(); return [performance.now() - t, b.byteLength]; });
  await page.evaluate(() => { window.__clickAt = performance.now(); document.querySelector('[data-testid="office-editor-save"]').click(); });
  await page.waitForFunction(() => window.__uploadAt != null, null, { timeout: 180000 });
  const host = await page.evaluate(() => window.__uploadAt - window.__clickAt);
  await waitUploads(page, 1, 180000);
  await page.evaluate(() => window.__grids[0].EnterText(61, 30, "edit2"));
  const alone2 = await page.evaluate(async () => { const t = performance.now(); const b = await window.__grids[0].ExportXlsx(); return [performance.now() - t, b.byteLength]; });
  const verify = await page.evaluate(async () => { const b = await window.__grids[0].ExportXlsx(); const m = await import("/node_modules/.vite/deps/@rbt_office_spreadsheet_XlsxImporter__js.js").catch(() => null); return m ? null : null; });
  rows.push({ host, alone1: alone1[0], alone2: alone2[0] });
  console.log(`run ${run + 1}: click -> upload request ${host.toFixed(0)} ms; ExportXlsx alone before ${alone1[0].toFixed(0)} ms, after ${alone2[0].toFixed(0)} ms (${alone1[1]} bytes)`);
  await ctx.close();
}
console.log(`median: host save (to the upload request) ${med(rows.map((r) => r.host)).toFixed(0)} ms; ExportXlsx alone ${med(rows.flatMap((r) => [r.alone1, r.alone2])).toFixed(0)} ms`);
await browser.close();
