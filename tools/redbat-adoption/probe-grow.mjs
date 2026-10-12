import { launch, newPage, open, T } from "./lib.mjs";
import { g, sel, ready, frames, select, cellPoint } from "./ui-lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`); await ready(page);
const st = (label) => g(page, (label) => { const gr = window.__grids.at(-1); const m = gr.GetModel(); const u = gr.GetUsedRange(); return label + " " + JSON.stringify({ rows: m.RowCount, used: u.range, rev: gr.GetRevision(), mod: gr.IsModified(), c400: gr.GetCell(400, 1)?.value ?? null, far: m.RowCount > 998 ? (gr.GetCell(998, 2)?.value ?? null) : "(outside)", undoDisabled: document.querySelector('[data-testid="sheet-undo"]').disabled }); }, label);
console.log(await st("open"));
await select(page, 400, 1); await frames(page);
const p = await cellPoint(page, 400, 1); await page.mouse.click(p.x, p.y);
await page.keyboard.type("below the end"); await page.keyboard.press("Enter"); await frames(page);
console.log(await st("typed"));
await g(page, () => window.__grids.at(-1).SetCell(998, 2, "42")); await frames(page);
console.log(await st("SetCell(998)"));
await sel(page, "sheet-undo").click(); await frames(page);
console.log(await st("undo 1"));
await g(page, () => window.__grids.at(-1).SetCell(450, 2, "7")); await frames(page);
console.log(await st("SetCell inside the grid (450)"), await g(page, () => window.__grids.at(-1).GetCell(450, 2)?.value));
await sel(page, "sheet-undo").click(); await frames(page);
console.log(await st("undo of it"), await g(page, () => window.__grids.at(-1).GetCell(450, 2)?.value ?? null));
// typed beyond via keyboard: go to last row of grid and type
const rows = await g(page, () => window.__grids.at(-1).GetModel().RowCount);
await g(page, (r) => { const gr = window.__grids.at(-1); gr.ScrollToCell(r, 1); gr.SelectCell(r, 1); }, rows - 1); await frames(page);
console.log(await st("selected last grid row " + rows));
await page.keyboard.type("x"); await page.keyboard.press("Enter"); await frames(page);
console.log(await st("typed in last grid row"), await g(page, () => window.__grids.at(-1).GetSelectionRef()));
console.log("cell typed:", await g(page, (r) => window.__grids.at(-1).GetCell(r, 1)?.value ?? null, rows - 1));
await sel(page, "sheet-undo").click(); await frames(page);
console.log(await st("undo"), "cell typed:", await g(page, (r) => window.__grids.at(-1).GetCell(r, 1)?.value ?? null, rows - 1));
// EnterText beyond
const rows2 = await g(page, () => window.__grids.at(-1).GetModel().RowCount);
await g(page, (r) => { const gr = window.__grids.at(-1); gr.SetCell(r + 300, 3, "9"); }, rows2); await frames(page);
console.log(await st("SetCell beyond again, as the only edit since"), await g(page, (r) => window.__grids.at(-1).GetCell(r + 300, 3)?.value ?? null, rows2));
const dis = await g(page, () => document.querySelector('[data-testid="sheet-undo"]').disabled);
if (!dis) { await sel(page, "sheet-undo").click(); await frames(page); console.log(await st("undo"), await g(page, (r) => window.__grids.at(-1).GetCell(r + 300, 3)?.value ?? null, rows2)); }
await browser.close();
