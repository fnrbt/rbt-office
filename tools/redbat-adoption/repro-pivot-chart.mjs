// A pivot chart is drawn with the pivot table's Grand Total row once its sheet has been shown.
// dashboard-ecommerce-sales.xlsx, chart "Sales and Profit Analysis" on the Dashboard sheet: its categories are
// 'Combo Chart'!$A$4:$A$16 (13 cells, the last is "Grand Total"), its cache has ptCount 12 (Jan..Dec) - what Excel draws.
// usage: node repro-pivot-chart.mjs
import { launch, newPage, open, T, S } from "./lib.mjs";
import { g, sel, ready, frames } from "./ui-lib.mjs";
const { browser, context } = await launch({ viewport: { width: 1400, height: 820 } });
const page = await newPage(context);
await open(page, "sheet-editor", `${T}/real-world/dashboard-ecommerce-sales.xlsx`); await ready(page);
const names = await g(page, () => window.__grids.at(-1).GetSheetNames());
const show = async (name) => { await sel(page, `sheet-tab-${names.indexOf(name)}`).click(); await frames(page); await page.waitForTimeout(500); };
const combo = () => g(page, () => { const o = [...document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-object-drawn')].map((o) => [...o.querySelectorAll("text")].map((t) => t.textContent)).find((t) => t.includes("Sum of Profit")); return o ? o.filter((t) => /^[A-Z][a-z]{2}$|Total/.test(t)).join(" ") + " | axis: " + o.filter((t) => /^\$ /.test(t)).join(" ") : null; });
await show("Dashboard");
console.log("Dashboard shown first (the pivot's sheet not yet shown), categories:", await combo(), "| loaded:", await g(page, () => { const gr = window.__grids.at(-1); return gr.GetSheetNames().map((n, i) => (gr.IsSheetLoaded?.(i) ? 1 : 0)).join(""); }));
await page.screenshot({ path: `${S}/explore/repro-pivot-chart-1-before.png` });
await show("Combo Chart"); await show("Dashboard");
console.log("after the Combo Chart sheet was shown,                     categories:", await combo(), "| loaded:", await g(page, () => { const gr = window.__grids.at(-1); return gr.GetSheetNames().map((n, i) => (gr.IsSheetLoaded?.(i) ? 1 : 0)).join(""); }));
await page.screenshot({ path: `${S}/explore/repro-pivot-chart-2-after.png` });
await browser.close();
