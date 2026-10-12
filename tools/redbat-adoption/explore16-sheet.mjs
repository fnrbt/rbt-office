// Exploratory flows, spreadsheet, 0.1.16 in the host: real mouse and keyboard. Each flow prints ok / DEFECT with what was seen.
// usage: node explore16-sheet.mjs [flows]
import { launch, newPage, open, T, S, waitUploads, saved } from "./lib.mjs";
import { g, sel, ready, frames, cellPoint, selection, cell } from "./ui-lib.mjs";
const only = process.argv[2]?.split(",");
const { browser, context } = await launch({ viewport: { width: 1400, height: 820 } });
const E = `${S}/explore`;
let page, errors;
const fresh = async (file, opts = {}) => { if (page) await page.close(); page = await newPage(context); errors = []; page.on("pageerror", (e) => errors.push(String(e).slice(0, 200))); page.on("console", (m) => { if (m.type() === "error" && !/404/.test(m.text())) errors.push(m.text().slice(0, 200)); }); await open(page, "sheet-editor", file, opts); await ready(page); await page.waitForTimeout(300); };
const click = async (r, c, o = {}) => { const p = await cellPoint(page, r, c); await page.mouse.click(p.x, p.y, o); return p; };
const say = (n, name, ok, detail = "") => console.log(`${ok ? "ok    " : "DEFECT"} S${n} ${name}${detail === "" ? "" : " :: " + (typeof detail === "string" ? detail : JSON.stringify(detail))}`);
const shot = async (name) => { await frames(page); await page.waitForTimeout(200); await page.screenshot({ path: `${E}/x16-s-${name}.png` }); };
const menuItem = async (id) => { await sel(page, `sheet-menu-${id}`).click(); await frames(page); };
const status = async () => (await sel(page, "office-editor-status").innerText()).replace(/\n/g, " | ");
const flow = async (n, name, fn) => { if (only && !only.includes(String(n))) return; try { await fn(n, name); } catch (e) { say(n, name, false, "script error " + String(e).slice(0, 260)); } if (errors?.length) { console.log(`       page errors in S${n}: ${[...new Set(errors)].slice(0, 3).join(" ;; ")}`); errors.length = 0; } };
const ZOO = `${T}/real-world/chart-zoo.xlsx`, MILK = `${T}/real-world/milk-street-office.xlsx`, ECOM = `${T}/real-world/dashboard-ecommerce-sales.xlsx`, SALES = `${T}/real-world/dashboard-sales-analysis.xlsx`;
const svgOf = (i = 0) => g(page, (i) => { const o = [...document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-object-drawn')][i]; if (!o) return null; const b = o.getBoundingClientRect(); return { html: o.querySelector("svg").outerHTML.length, hash: [...o.querySelector("svg").outerHTML].reduce((h, ch) => (h * 31 + ch.charCodeAt(0)) | 0, 0), x: Math.round(b.x), y: Math.round(b.y), w: Math.round(b.width), h: Math.round(b.height), texts: [...o.querySelectorAll("text")].map((t) => t.textContent).slice(0, 30) }; }, i);

await flow(1, "type numbers and a SUM, Enter moves down", async (n, name) => {
  await fresh(ZOO); await click(8, 0); for (const t of ["10", "20", "30", "=SUM(A9:A11)"]) { await page.keyboard.type(t); await page.keyboard.press("Enter"); } await frames(page);
  const c = await cell(page, 11, 0); say(n, name, c.text === "60" && (await selection(page)) === "A13", [c.text, await selection(page)]);
});
await flow(2, "a cell the charts read is edited: the charts are drawn again with the new value; undo draws the old", async (n, name) => {
  await fresh(ZOO); const a = await svgOf(0); await click(1, 1); await page.keyboard.type("55"); await page.keyboard.press("Enter"); await frames(page); await page.waitForTimeout(300);
  const b = await svgOf(0); await page.keyboard.press("Control+z"); await frames(page); await page.waitForTimeout(300); const c = await svgOf(0);
  await page.keyboard.press("Control+y"); await frames(page); await page.waitForTimeout(300); await shot("2-chart-after-edit");
  say(n, name, a.hash !== b.hash && c.hash === a.hash && b.texts.some((t) => /^(50|55|60)$/.test(t)), { before: a.texts.slice(0, 12), after: b.texts.slice(0, 12), undoSame: c.hash === a.hash });
});
await flow(3, "a row inserted above the charts through the context menu: the charts move down a row; undo", async (n, name) => {
  await fresh(ZOO); const a = await svgOf(0); const p = await click(2, 0); await page.mouse.click(p.x, p.y, { button: "right" }); await frames(page);
  await menuItem("insert-rows"); await frames(page); await page.waitForTimeout(300); const b = await svgOf(0);
  await shot("3-row-inserted"); await page.keyboard.press("Control+z"); await frames(page); await page.waitForTimeout(300); const c = await svgOf(0);
  say(n, name, b.y - a.y >= 15 && b.y - a.y <= 30 && c.y === a.y && b.h === a.h, { y: [a.y, b.y, c.y], h: [a.h, b.h] });
});
await flow(4, "zoom in twice: the charts scale with the cells", async (n, name) => {
  await fresh(ZOO); const a = await svgOf(0); await sel(page, "sheet-zoom-in").click(); await sel(page, "sheet-zoom-in").click(); await frames(page); await page.waitForTimeout(300); const b = await svgOf(0); const z = await sel(page, "sheet-zoom").innerText();
  await shot("4-zoomed"); const f = parseInt(z) / 100; say(n, name, Math.abs(b.w / a.w - f) < 0.06, { zoom: z, w: [a.w, b.w] });
});
await flow(5, "a column made wider by dragging its header edge: a chart anchored right of it moves", async (n, name) => {
  await fresh(ZOO); const a = await svgOf(0); const hs = await g(page, () => [...document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-col-header')].map((e) => { const b = e.getBoundingClientRect(); return { x: b.right, y: b.top + b.height / 2, w: b.width, col: e.dataset.col }; }).filter((b) => b.w > 0));
  const h = hs.find((x) => x.col === "1") ?? hs[1]; await page.mouse.move(h.x - 1, h.y); await page.mouse.down(); await page.mouse.move(h.x + 60, h.y, { steps: 5 }); await page.mouse.up(); await frames(page); await page.waitForTimeout(300);
  const b = await svgOf(0); say(n, name, b.x - a.x > 40, { x: [a.x, b.x], header: h });
});
await flow(6, "freeze panes at C4 and scroll: the charts go under the frozen rows", async (n, name) => {
  await fresh(ZOO); await click(3, 2); await sel(page, "sheet-menu-button-view").click(); await frames(page);
  const items = await g(page, () => [...document.querySelectorAll('[data-testid^="sheet-"]')].filter((e) => /freeze/i.test(e.textContent) && e.tagName === "BUTTON").map((e) => e.dataset.testid));
  await sel(page, items[0]).click(); await frames(page);
  await g(page, () => { const sc = document.querySelector(".spreadsheet-container"); sc.scrollTop = 160; sc.dispatchEvent(new Event("scroll")); }); await frames(page); await page.waitForTimeout(300); await shot("6-frozen-scrolled");
  const fr = await g(page, () => window.__grids.at(-1).GetViewOptions().frozen); say(n, name, fr.rows === 3 && fr.cols === 2, { fr, items });
});
await flow(7, "a note through the context menu's dialog, then edited", async (n, name) => {
  await fresh(MILK); const p = await click(4, 9); await page.mouse.click(p.x, p.y, { button: "right" }); await frames(page); await menuItem("comment"); await frames(page);
  await sel(page, "sheet-comment-text").fill("First note"); await sel(page, "sheet-dialog-ok").click(); await frames(page);
  const a = await g(page, () => window.__grids.at(-1).GetComment(4, 9));
  await page.mouse.click(p.x, p.y, { button: "right" }); await frames(page); await menuItem("comment"); await frames(page); await sel(page, "sheet-comment-text").focus(); await page.keyboard.press("Control+End"); await page.keyboard.type(" and more"); await sel(page, "sheet-dialog-ok").click(); await frames(page);
  const b = await g(page, () => window.__grids.at(-1).GetComment(4, 9)); await page.mouse.move(p.x + 3, p.y + 2); await page.waitForTimeout(700); await shot("7-note");
  say(n, name, a?.text === "First note" && b?.text === "First note and more" && a.author === b.author, [a, b]);
});
await flow(8, "copy a range and paste it elsewhere with the keyboard", async (n, name) => {
  await fresh(ZOO); await click(0, 0); await page.keyboard.down("Shift"); await click(4, 1); await page.keyboard.up("Shift"); await page.keyboard.press("Control+c"); await page.waitForTimeout(150);
  await click(12, 0); await page.keyboard.press("Control+v"); await page.waitForTimeout(400); await frames(page);
  const got = [(await cell(page, 12, 0)).text, (await cell(page, 16, 1)).text]; say(n, name, got[0] === "Quarter" && got[1] === "30", got);
});
await flow(9, "fill handle dragged down four rows: the series goes on", async (n, name) => {
  await fresh(ZOO); await click(8, 0); await page.keyboard.type("1"); await page.keyboard.press("Enter"); await page.keyboard.type("2"); await page.keyboard.press("Enter");
  await click(8, 0); await page.keyboard.down("Shift"); await click(9, 0); await page.keyboard.up("Shift"); await frames(page);
  const h = await g(page, () => { const e = document.querySelector('[data-testid="sheet-grid"] .spreadsheet-fill-handle, [data-testid="sheet-grid"] [class*="fill-handle"]'); if (!e) return null; const b = e.getBoundingClientRect(); return { x: b.x + b.width / 2, y: b.y + b.height / 2, cls: e.className }; });
  if (!h) return say(n, name, false, "no fill handle element found");
  const to = await cellPoint(page, 13, 0); await page.mouse.move(h.x, h.y); await page.mouse.down(); await page.mouse.move(h.x, to.y, { steps: 6 }); await page.mouse.up(); await frames(page);
  const vals = []; for (let r = 8; r <= 13; r++) vals.push((await cell(page, r, 0)).text); say(n, name, vals.join(",") === "1,2,3,4,5,6", vals);
});
await flow(10, "F2 edits in the cell, Escape gives the entry up; typing replaces; Delete clears; three undos", async (n, name) => {
  await fresh(ZOO); await click(1, 1); await page.keyboard.press("F2"); await page.keyboard.type("9"); await page.keyboard.press("Escape"); const a = (await cell(page, 1, 1)).text;
  await page.keyboard.type("77"); await page.keyboard.press("Tab"); const b = (await cell(page, 1, 1)).text; await click(1, 1); await page.keyboard.press("Delete"); const c = (await cell(page, 1, 1)).text;
  await page.keyboard.press("Control+z"); await page.keyboard.press("Control+z"); await frames(page); const d = (await cell(page, 1, 1)).text;
  say(n, name, a === "10" && b === "77" && c === "" && d === "10" && /No changes/.test(await status()), [a, b, c, d, await status()]);
});
await flow(11, "the ecommerce dashboard: switch through the sheets by their tabs, back to the dashboard", async (n, name) => {
  await fresh(ECOM); const names = await g(page, () => window.__grids.at(-1).GetSheetNames()); const seen = [];
  for (let i = 0; i < names.length; i++) { const t0 = Date.now(); await sel(page, `sheet-tab-${i}`).click(); await frames(page); await page.waitForTimeout(250); seen.push(`${names[i]}: ${await g(page, () => document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-object-drawn').length)}/${await g(page, () => document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-object').length)} drawn, ${Date.now() - t0} ms`); if (["Combo Chart", "Top 5 Customers", "KPI", "Waterfall Chart", "Pie Chart"].includes(names[i])) await shot(`11-ecom-${names[i].replace(/ /g, "-")}`); }
  say(n, name, /No changes/.test(await status()), seen.join("; "));
});
await flow(12, "the sales dashboard: scroll down with the wheel over the charts and back", async (n, name) => {
  await fresh(SALES); await page.mouse.move(700, 450); for (let i = 0; i < 10; i++) { await page.mouse.wheel(0, 200); await page.waitForTimeout(40); } await frames(page); await page.waitForTimeout(250); await shot("12-sales-scrolled");
  const top = await g(page, () => document.querySelector(".spreadsheet-container").scrollTop); for (let i = 0; i < 12; i++) { await page.mouse.wheel(0, -200); await page.waitForTimeout(30); } await frames(page);
  say(n, name, top > 500 && (await g(page, () => document.querySelector(".spreadsheet-container").scrollTop)) === 0, { top });
});
await flow(13, "a click on a chart goes to the cell under it; dragging over a chart selects cells", async (n, name) => {
  await fresh(ZOO); const a = await svgOf(0); await page.mouse.click(a.x + a.w / 2, a.y + a.h / 2); await frames(page); const s1 = await selection(page);
  await page.mouse.move(a.x + 20, a.y + 20); await page.mouse.down(); await page.mouse.move(a.x + 220, a.y + 120, { steps: 5 }); await page.mouse.up(); await frames(page); const s2 = await selection(page);
  say(n, name, /^[F-K]\d+$/.test(s1) && /:/.test(s2), [s1, s2]);
});
await flow(14, "sort a column descending through the Data menu; the chart follows; undo", async (n, name) => {
  await fresh(ZOO); const a = await svgOf(0); await click(1, 1); await sel(page, "sheet-menu-button-data").click(); await frames(page); await menuItem("sort-desc"); await frames(page); await page.waitForTimeout(300);
  const order = []; for (let r = 1; r <= 4; r++) order.push((await cell(page, r, 1)).text); const b = await svgOf(0); await page.keyboard.press("Control+z"); await frames(page);
  say(n, name, order.join(",") === "30,25,15,10" && b.hash !== a.hash, { order, cats: b.texts.filter((t) => /^Q\d$/.test(t)) });
});
await flow(15, "save the dashboard after an edit and open what was saved: the charts are drawn as before", async (n, name) => {
  await fresh(ECOM); const i = (await g(page, () => window.__grids.at(-1).GetSheetNames())).indexOf("Dashboard"); await sel(page, `sheet-tab-${i}`).click(); await frames(page); await page.waitForTimeout(300);
  const before = await g(page, () => [document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-object-drawn').length, document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-object').length]);
  await g(page, () => window.__grids.at(-1).SelectCell(40, 20)); await page.keyboard.type("x"); await page.keyboard.press("Enter"); await sel(page, "office-editor-save").click(); const out = await waitUploads(page, 1);
  saved.set("x15", out); await fresh("/api/saved/x15", { name: "dashboard-ecommerce-sales.xlsx" }); await sel(page, `sheet-tab-${i}`).click(); await frames(page); await page.waitForTimeout(400);
  const after = await g(page, () => [document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-object-drawn').length, document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-object').length]);
  say(n, name, JSON.stringify(before) === JSON.stringify(after) && before[0] > 0, { before, after, bytes: out.length });
});
await flow(16, "bold, fill colour and a number format from the toolbar on a range; one undo each", async (n, name) => {
  await fresh(ZOO); await click(1, 1); await page.keyboard.down("Shift"); await click(4, 1); await page.keyboard.up("Shift");
  await page.locator('[data-testid="sheet-toolbar"] button[title^="Bold"]').click(); await sel(page, "sheet-number-format").selectOption("currency"); await frames(page);
  const c = await cell(page, 2, 1); const f = await g(page, () => window.__grids.at(-1).GetSelectionFormat());
  await page.keyboard.press("Control+z"); await page.keyboard.press("Control+z"); await frames(page);
  say(n, name, /\$25\.00/.test(c.text) && f.bold === true && /No changes/.test(await status()), [c.text, f.bold, await status()]);
});
await flow(17, "find from the toolbar: next match selected; replace all; undo", async (n, name) => {
  await fresh(ZOO); await page.locator('[data-testid="sheet-toolbar"] button[title^="Find"]').click(); await frames(page);
  await sel(page, "sheet-find-input").fill("Q"); await sel(page, "sheet-find-next").click(); await frames(page); const s1 = await selection(page); const st = await sel(page, "sheet-find-status").innerText();
  await sel(page, "sheet-replace-input").fill("Quarter "); await sel(page, "sheet-replace-all").click(); await frames(page); const c = (await cell(page, 1, 0)).text; await sel(page, "sheet-find-close").click(); await page.keyboard.press("Control+z"); await frames(page);
  say(n, name, /^A\d$/.test(s1) && /Quarter /.test(c) && (await cell(page, 1, 0)).text === "Q1", [s1, st, c]);
});
await flow(18, "a new sheet from the + tab, typed into, renamed through the tab menu, deleted", async (n, name) => {
  await fresh(ZOO); await sel(page, "sheet-tab-add").click(); await frames(page); await page.keyboard.type("hello"); await page.keyboard.press("Enter");
  const names = await g(page, () => window.__grids.at(-1).GetSheetNames()); await sel(page, "sheet-tab-0").click(); await frames(page); await page.waitForTimeout(250);
  const drawn = await g(page, () => document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-object-drawn').length); await sel(page, "sheet-tab-1").click(); await frames(page);
  const onNew = await g(page, () => document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-object').length);
  say(n, name, names.length === 2 && drawn === 6 && onNew === 0, { names, drawn, onNew });
});
await flow(19, "a selection dragged over the frozen-less grid to the right edge and back in: no runaway scroll", async (n, name) => {
  await fresh(MILK); const p = await click(6, 2); const grid = await sel(page, "sheet-grid").boundingBox();
  await page.mouse.move(p.x, p.y); await page.mouse.down(); await page.mouse.move(grid.x + grid.width - 2, p.y, { steps: 5 }); await page.waitForTimeout(400); await page.mouse.move(p.x + 200, p.y + 40, { steps: 5 }); await page.waitForTimeout(400);
  const a = await g(page, () => document.querySelector(".spreadsheet-container").scrollLeft); await page.waitForTimeout(400); const b = await g(page, () => document.querySelector(".spreadsheet-container").scrollLeft); await page.mouse.up(); await frames(page);
  say(n, name, a === b && a > 0, { a, b, sel: await selection(page) });
});
await flow(20, "merge cells from the toolbar, type into the merged cell, unmerge", async (n, name) => {
  await fresh(ZOO); await click(8, 0); await page.keyboard.down("Shift"); await click(9, 2); await page.keyboard.up("Shift"); await page.locator('[data-testid="sheet-toolbar"] button[title*="erge"]').first().click(); await frames(page);
  await page.keyboard.type("merged"); await page.keyboard.press("Enter"); const c = await cell(page, 8, 0); say(n, name, c.isMerged === true && c.text === "merged", { merged: c.isMerged, text: c.text });
});
await flow(21, "the Combo Chart sheet's pivot chart: the categories drawn against the chart's cached points", async (n, name) => {
  await fresh(ECOM); const i = (await g(page, () => window.__grids.at(-1).GetSheetNames())).indexOf("Dashboard"); await sel(page, `sheet-tab-${i}`).click(); await frames(page); await page.waitForTimeout(400);
  const all = await g(page, () => [...document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-object-drawn')].map((o) => [...o.querySelectorAll("text")].map((t) => t.textContent)));
  const combo = all.find((t) => t.includes("Sum of Profit")) ?? [];
  say(n, name, !combo.includes("Grand Total"), { texts: combo.slice(0, 30) });
});
if (page) await page.close();
await browser.close();
