import { launch, newPage, open, T } from "./lib.mjs";
import { g, ready } from "./ui-lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "sheet-editor", `${T}/real-world/dashboard-ecommerce-sales.xlsx`); await ready(page);
console.log(JSON.stringify(await g(page, () => { const gr = window.__grids[0]; const hs = [...document.querySelectorAll(".spreadsheet-col-header")]; return { n: hs.length, sample: hs.slice(0, 5).map((h) => [h.dataset.col, h.className, Math.round(h.getBoundingClientRect().width)]), headings: gr.GetShowHeadings?.(), used: gr.GetUsedRange(), sheet: gr.GetSheetNames()[gr.GetActiveSheetIndex()], classes: [...new Set([...document.querySelectorAll('[data-testid="sheet-grid"] [class*="header"]')].map((e) => e.className))].slice(0, 10) }; })));
await browser.close();
