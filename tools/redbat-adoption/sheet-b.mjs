import { launch, newPage, open, T, SHOTS, check, summary, saved, unzip, waitUploads } from "./lib.mjs";
import fs from "node:fs";
const { browser, context } = await launch();
const g = (page, fn, arg) => page.evaluate(fn, arg);
const ready = async (page) => { await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 180000 }); await page.waitForTimeout(300); };
const status = (page) => page.locator('[data-testid="office-editor-status"]').innerText();
const file = `${T}/real-world/dashboard-sales-analysis.xlsx`;
const page = await newPage(context);
let t = Date.now();
await open(page, "sheet-editor", file);
await ready(page);
console.log("  editor load ms", Date.now() - t);
await page.screenshot({ path: `${SHOTS}/sheet-dashboard.png` });
// ---- lazy sheets: only what is shown is converted
const names = await g(page, () => window.__grids[0].GetSheetNames());
const dataIndex = names.indexOf("Data");
const loadedNow = () => g(page, () => { const gr = window.__grids[0]; return gr.GetSheetNames().map((n, i) => gr.IsSheetLoaded(i)); });
let loaded = await loadedNow();
check("opening converts the active sheet only; the big Data sheet is not loaded", loaded.filter(Boolean).length === 1 && loaded[dataIndex] === false && dataIndex >= 0, { names, loaded });
check("tabs are there for every visible sheet without loading them", (await page.locator('[data-testid="sheet-tabs"] [role="tab"]').count()) === (await g(page, () => window.__grids[0].SheetVisibility().filter(Boolean).length)));
// an edit on the open sheet, saved without ever showing Data
await g(page, () => { const gr = window.__grids[0]; gr.EnterText(60, 30, "lazy-edit"); });
const afterEdit = await loadedNow();
console.log("  loaded after an edit on the Dashboard sheet (recalculation reads the sheets its formulas refer to):", JSON.stringify(afterEdit));
t = Date.now();
await page.click('[data-testid="office-editor-save"]');
const lazyBytes = await waitUploads(page, 1, 180000);
console.log("  save ms with the Data sheet not loaded", Date.now() - t);
await page.waitForTimeout(200);
loaded = await loadedNow();
check("saving loads no sheet: the same ones as before it, and not Data", JSON.stringify(loaded) === JSON.stringify(afterEdit) && loaded[dataIndex] === false, loaded);
check("saved (lazy)", (await status(page)).includes("Saved") && (await page.locator('[data-testid="office-editor-error"]').count()) === 0, await status(page));
{
  const orig = await unzip(fs.readFileSync(file)), out = await unzip(lazyBytes);
  const wb = await out.get("xl/workbook.xml").async("string");
  const rels = await out.get("xl/_rels/workbook.xml.rels").async("string");
  const part = (z, wbXml, relsXml, name) => { const rid = new RegExp(`<sheet [^>]*name="${name}"[^>]*r:id="([^"]+)"`).exec(wbXml)?.[1]; const target = new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${rid}"`).exec(relsXml); return "xl/" + (target[1] ?? target[2]).replace(/^\/xl\//, ""); };
  const owb = await orig.get("xl/workbook.xml").async("string"), orels = await orig.get("xl/_rels/workbook.xml.rels").async("string");
  const a = await orig.get(part(orig, owb, orels, "Data")).async("string"), b = await out.get(part(out, wb, rels, "Data")).async("string");
  const rows = (x) => (x.match(/<row /g) || []).length, cells = (x) => (x.match(/<c /g) || []).length;
  check("the Data sheet nobody looked at is written with every row and cell", rows(a) === rows(b) && cells(a) === cells(b) && rows(a) > 70000, { rows: [rows(a), rows(b)], cells: [cells(a), cells(b)], identical: a === b });
  saved.set("dash-lazy", lazyBytes);
}
t = Date.now();
await page.locator('[data-testid="sheet-tabs"], [data-testid="spreadsheet-preview-tabs"]').getByText("Data", { exact: true }).click();
console.log("  showing the Data sheet ms", Date.now() - t);
await page.waitForTimeout(300);
check("the Data sheet is loaded once its tab is shown", (await loadedNow())[dataIndex] === true);
const info = await g(page, () => { const m = window.__grids[0].GetModel(); const sc = document.querySelector(".spreadsheet-container"); return { rows: m.RowCount, scrollHeight: sc.scrollHeight, clientHeight: sc.clientHeight }; });
check("scroll extent follows the sheet", info.scrollHeight > 76000 * 15, info);
// Row 5,000 by scroll (wheel).
const box = await page.locator('[data-testid="sheet-grid"]').boundingBox();
await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
const rowVisible = (n) => g(page, (n) => { const sc = document.querySelector(".spreadsheet-container").getBoundingClientRect(); return [...document.querySelectorAll(".spreadsheet-row-header")].some((h) => h.textContent.trim() === String(n) && h.getBoundingClientRect().top >= sc.top - 1 && h.getBoundingClientRect().bottom <= sc.bottom + 1); }, n);
let wheels = 0;
while (!(await rowVisible(5000)) && wheels < 1000) { await page.mouse.wheel(0, 240); wheels++; await page.waitForTimeout(10); }
await page.waitForTimeout(100);
check("row 5,000 reachable by wheel scrolling", await rowVisible(5000), { wheels });
const cellText = await g(page, () => [...document.querySelectorAll('.spreadsheet-cell[data-row="4999"]')].map((c) => c.textContent).filter(Boolean).slice(0, 3));
check("row 5,000 cells are drawn", cellText.length > 0, cellText);
await page.screenshot({ path: `${SHOTS}/sheet-row5000.png` });
// By keyboard: back to the top, then PageDown until row 5,000.
await g(page, () => window.__grids[0].SelectCell(0, 0));
await page.locator(".spreadsheet-content").first().focus();
let presses = 0;
while ((await g(page, () => window.__grids[0].GetActiveCell().Row)) < 4999 && presses < 1000) { await page.keyboard.press("PageDown"); presses++; }
const active = await g(page, () => window.__grids[0].GetActiveCell());
check("row 5,000 reachable by keyboard (PageDown)", active.Row >= 4999, { presses, active });
check("the active cell is revealed", await rowVisible(active.Row + 1), await g(page, (n) => { const sc = document.querySelector(".spreadsheet-container"); const r = sc.getBoundingClientRect(); const h = [...document.querySelectorAll(".spreadsheet-row-header")].find((h) => h.textContent.trim() === String(n)); const a = document.querySelector(".spreadsheet-active-cell")?.getBoundingClientRect(); return { sc: [r.top, r.bottom, sc.clientHeight, sc.scrollTop], h: h ? [h.getBoundingClientRect().top, h.getBoundingClientRect().bottom] : null, a: a ? [a.top, a.bottom] : null }; }, active.Row + 1));
await page.screenshot({ path: `${SHOTS}/sheet-kbd5000.png` });
await page.keyboard.press("Control+ArrowDown");
const last = await g(page, () => window.__grids[0].GetActiveCell());
check("Ctrl+Down jumps to the last data row and shows it", last.Row > 70000 && (await rowVisible(last.Row + 1)), last);
check("cell reference follows the keyboard", (await page.locator('[data-testid="sheet-cell-ref"]').inputValue()) === "A" + (last.Row + 1));
// edit there and save
await page.keyboard.type("edited-last");
await page.keyboard.press("Enter");
check("dirty", (await status(page)).includes("Unsaved"));
t = Date.now();
await page.click('[data-testid="office-editor-save"]');
const bytes = await waitUploads(page, 2, 180000);
console.log("  save ms (export + verify)", Date.now() - t);
await page.waitForTimeout(200);
check("saved with no warning or error", (await status(page)).includes("Saved") && (await page.locator('[data-testid="office-editor-error"]').count()) === 0 && (await page.locator('[data-testid="sheet-notice"]').count()) === 0, await status(page));
const orig = await unzip(fs.readFileSync(file));
const out = await unzip(bytes);
const count = (z, re) => [...z.keys()].filter((n) => re.test(n)).length;
for (const [label, re] of [["pivot tables", /^xl\/pivotTables\/pivotTable\d+\.xml$/], ["pivot caches", /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/], ["pivot cache records", /^xl\/pivotCache\/pivotCacheRecords\d+\.xml$/], ["slicers", /^xl\/slicers\/slicer\d+\.xml$/], ["slicer caches", /^xl\/slicerCaches\//], ["charts", /^xl\/charts\/chart\d+\.xml$/], ["media", /^xl\/media\//]]) {
  check(`saved file keeps ${label}`, count(out, re) === count(orig, re) && (count(orig, re) > 0 || /media|records/.test(label)), { orig: count(orig, re), out: count(out, re) });
}
const anchors = async (z) => { let n = 0; for (const [name, f] of z) if (/^xl\/drawings\/[^/]+\.xml$/.test(name)) n += ((await f.async("string")).match(/<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)[ >]/g) || []).length; return n; };
check("saved file keeps every drawing anchor (charts, shapes, slicer frames)", (await anchors(out)) === (await anchors(orig)) && (await anchors(orig)) === 15, { orig: await anchors(orig), out: await anchors(out) });
const missing = [...orig.keys()].filter((n) => !out.has(n) && !/^xl\/drawings\//.test(n));
console.log("  parts of the original not in the saved file (drawings are renamed):", missing.join(", "));
saved.set("dash", bytes);
await page.close();
// re-open
const page2 = await newPage(context);
await open(page2, "sheet-editor", "/api/saved/dash", { name: "dashboard-sales-analysis.xlsx" });
await ready(page2);
const r = await g(page2, (row) => { const gr = window.__grids[0]; const names = gr.GetSheetNames(); gr.SwitchSheet(names.indexOf("Data")); gr.SelectCell(row, 0); return { names, text: gr.GetSelectionText().trim() }; }, last.Row);
check("re-open: edit on the last row is there", r.text === "edited-last", r);
check("re-open: same sheets", r.names.length === 8, r.names);
await page2.close();
// the lazily saved file: opens, the edit is there, Data loads when shown and has its rows
{
  const page3 = await newPage(context);
  await open(page3, "sheet-editor", "/api/saved/dash-lazy", { name: "dashboard-sales-analysis.xlsx" });
  await ready(page3);
  const r3 = await g(page3, () => { const gr = window.__grids[0]; const names = gr.GetSheetNames(); const di = names.indexOf("Data"); gr.SelectCell(60, 30); const edit = gr.GetSelectionText().trim(); const before = gr.IsSheetLoaded(di); gr.SwitchSheet(di); gr.SelectCell(0, 0); const head = gr.GetSelectionText().trim(); return { edit, before, after: gr.IsSheetLoaded(di), rows: gr.GetModel().RowCount, head }; });
  check("re-open of the lazily saved file: the edit, and the Data sheet with its rows", r3.edit === "lazy-edit" && r3.before === false && r3.after === true && r3.rows > 70000 && r3.head.length > 0, r3);
  await page3.close();
}
// ---- preview: the same laziness
{
  const page4 = await newPage(context);
  t = Date.now();
  await open(page4, "sheet-preview", file);
  await page4.waitForFunction(() => document.querySelector('[data-testid="spreadsheet-preview"]')?.style.visibility === "visible", null, { timeout: 180000 });
  console.log("  preview load ms", Date.now() - t);
  const l4 = await g(page4, () => { const gr = window.__grids[0]; return gr.GetSheetNames().map((n, i) => gr.IsSheetLoaded(i)); });
  check("preview: only the shown sheet is converted", l4.filter(Boolean).length === 1, l4);
  await page4.close();
}
summary();
await browser.close();
