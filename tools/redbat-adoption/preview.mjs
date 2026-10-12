import { launch, newPage, open, T, SHOTS, check, summary } from "./lib.mjs";
const S = "/tmp/redbat-adoption";
const { browser, context } = await launch();
const g = (page, fn, arg) => page.evaluate(fn, arg);
const loaded = (page) => page.waitForFunction(() => document.querySelector('[data-testid="spreadsheet-preview"]')?.style.visibility === "visible", null, { timeout: 180000 });
const shot = (page, name) => page.screenshot({ path: `${SHOTS}/${name}.png` });
const cellsWith = (page, prop) => g(page, (prop) => [...document.querySelectorAll(".spreadsheet-cell")].filter((c) => { const s = getComputedStyle(c); return prop === "bg" ? s.backgroundColor !== "rgb(255, 255, 255)" && s.backgroundColor !== "rgba(0, 0, 0, 0)" : prop === "border" ? ["Top", "Bottom", "Left", "Right"].some((d) => s["border" + d + "Style"] !== "none" && parseFloat(s["border" + d + "Width"]) > 0 && !/rgba\(0, 0, 0, 0\)|224, 224, 224/.test(s["border" + d + "Color"])) : false; }).length, prop);

{ // milk-street: fills, borders, number formats, read-only
  const page = await newPage(context);
  await open(page, "sheet-preview", `${T}/real-world/milk-street-office.xlsx`);
  await loaded(page);
  await page.waitForTimeout(300);
  await shot(page, "preview-milk");
  check("preview is read-only", await g(page, () => window.__grids[0].IsReadOnly()));
  check("fills shown", (await cellsWith(page, "bg")) > 10, await cellsWith(page, "bg"));
  const bordered = await g(page, () => { const gr = window.__grids[0]; let n = 0; for (let r = 0; r < 40; r++) for (let c = 0; c < 14; c++) { const e = gr.GetCellEdges(r, c); if ([e.Top, e.Right, e.Bottom, e.Left].some((x) => x && x !== "grid")) n++; } return { cells: n, drawn: document.querySelectorAll(".spreadsheet-cell .spreadsheet-edge").length }; });
  check("cell borders shown (GetCellEdges and edge elements)", bordered.cells > 5 && bordered.drawn > 5, bordered);
  const title = await g(page, () => { const c = document.querySelector('.spreadsheet-cell[data-row="1"][data-col="1"]'); const r = c.getBoundingClientRect(); const range = document.createRange(); range.selectNodeContents(c); const text = range.getBoundingClientRect(); const clip = /inset\(0px (-?[\d.]+)px/.exec(c.style.clipPath); return { text: c.textContent, cellWidth: r.width, textWidth: text.width, overflow: getComputedStyle(c).overflow, allowed: r.width + (clip ? -parseFloat(clip[1]) : 0) }; });
  check("B2's long title overflows into the empty cells to its right and is fully visible", title.text === "45 Milk Street - Investment Analysis" && title.textWidth > title.cellWidth * 5 && title.overflow === "visible" && title.textWidth <= title.allowed, title);
  // pixels: bold dark text well to the right of B2's own 18 px
  const titlePixels = await page.evaluate(() => { const c = document.querySelector('.spreadsheet-cell[data-row="1"][data-col="1"]').getBoundingClientRect(); return { x: c.right + 150, y: c.top, w: 100, h: c.height }; });
  const png = await page.screenshot({ clip: { x: titlePixels.x, y: titlePixels.y, width: titlePixels.w, height: titlePixels.h } });
  const dark = await page.evaluate(async (b64) => { const img = new Image(); img.src = "data:image/png;base64," + b64; await img.decode(); const cv = document.createElement("canvas"); cv.width = img.width; cv.height = img.height; const cx = cv.getContext("2d"); cx.drawImage(img, 0, 0); const d = cx.getImageData(0, 0, cv.width, cv.height).data; let n = 0; for (let i = 0; i < d.length; i += 4) if (d[i] < 100 && d[i + 1] < 100 && d[i + 2] < 100) n++; return n; }, png.toString("base64"));
  check("pixels: the title's text is drawn 150-250 px right of its own cell", dark > 60, dark);
  check("number formats shown", await g(page, () => [...document.querySelectorAll(".spreadsheet-cell")].some((c) => c.textContent.includes("$ 20,500,000"))));
  check("no edit affordances (filter buttons, resize handles, fill handle)", await g(page, () => [...document.querySelectorAll(".col-filter-btn, .col-resize-handle, .row-resize-handle, .spreadsheet-fill-handle")].every((e) => getComputedStyle(e).display === "none")));
  // typing does nothing; selection and copy work
  const cell = page.locator('.spreadsheet-cell[data-row="6"][data-col="4"]').first();
  await cell.click();
  await page.keyboard.type("zzz");
  check("typing does not start an entry", !(await g(page, () => window.__grids[0].IsEditing())));
  await cell.dblclick();
  check("double click does not start an entry", !(await g(page, () => window.__grids[0].IsEditing())));
  await page.keyboard.press("Control+c");
  await page.waitForTimeout(150);
  check("copy from a preview", (await g(page, () => navigator.clipboard.readText())).trim() === "45 Milk Street", await g(page, () => navigator.clipboard.readText()));
  await page.keyboard.press("Delete");
  check("Delete changes nothing", (await g(page, () => window.__grids[0].GetSelectionText().trim())) === "45 Milk Street");
  const tabs = await page.locator('[data-testid="spreadsheet-preview-tabs"] button').allInnerTexts();
  check("sheet tabs", tabs.join(",") === "Annual,Monthly", tabs);
  await page.getByRole("button", { name: "Monthly", exact: true }).click();
  check("tab switches the sheet", (await g(page, () => window.__grids[0].GetActiveSheetIndex())) === 1);
  await page.waitForTimeout(200);
  await shot(page, "preview-milk-monthly");
  await page.close();
}
{ // merges + conditional formatting
  const page = await newPage(context);
  await open(page, "sheet-preview", `${T}/test-documents-excel/synthetic-condformat.xlsx`);
  await loaded(page);
  await page.waitForTimeout(300);
  await shot(page, "preview-condformat");
  const cf = await g(page, () => { const cells = [...document.querySelectorAll(".spreadsheet-cell")]; return { bars: cells.filter((c) => c.hasAttribute("data-cf-bar")).length, colored: cells.filter((c) => { const b = getComputedStyle(c).backgroundColor; return b !== "rgb(255, 255, 255)" && b !== "rgba(0, 0, 0, 0)"; }).length, icons: document.querySelectorAll(".spreadsheet-cell [class*=icon], .spreadsheet-cell svg").length }; });
  check("conditional formatting shown (coloured cells / data bars)", cf.colored + cf.bars > 0, cf);
  check("no sheet tabs for a single sheet", (await page.locator('[data-testid="spreadsheet-preview-tabs"]').count()) === 0);
  await page.close();
}
{ // merges
  const page = await newPage(context);
  await open(page, "sheet-preview", `${T}/real-world/dcf-stock-analysis.xlsx`);
  await loaded(page);
  // (0.1.12 opens a workbook on the sheet it was saved on; the merges looked at are the first sheet's.)
  const openedOn = await g(page, () => { const gr = window.__grids[0]; const at = gr.GetActiveSheetIndex(); gr.SwitchSheet(0); return at; });
  console.log("  dcf-stock-analysis opens on sheet", openedOn);
  await page.waitForTimeout(300);
  await shot(page, "preview-dcf-merges");
  const merged = await g(page, () => { const gr = window.__grids[0]; const m = gr.GetModel().MergedCells; const arr = Array.isArray(m) ? m : (() => { const a = []; for (let l = m; l && l.tail != null; l = l.tail) a.push(l.head); return a; })(); return [arr.slice(0, 3).map((x) => [x.StartRow, x.StartCol, x.EndRow, x.EndCol])]; });
  console.log("  merges per sheet (first 3):", JSON.stringify(merged));
  const first = merged[0].find((m) => m[1] !== m[3]) ?? merged[0][0];
  if (first) {
    const geo = await g(page, ([r, c, r2, c2]) => { window.__grids[0].ScrollToCell(r, c); const cell = document.querySelector(`.spreadsheet-cell[data-row="${r}"][data-col="${c}"]`); const inside = document.querySelector(`.spreadsheet-cell[data-row="${r2}"][data-col="${c2}"]`); const single = document.querySelector(`.spreadsheet-col-header[data-col="${c}"]`); return { w: cell?.getBoundingClientRect().width, colW: single?.getBoundingClientRect().width, coveredCellDrawn: !!inside && (r !== r2 || c !== c2) && getComputedStyle(inside).display !== "none" }; }, first);
    check("a merged cell is one cell spanning its columns", geo.w > geo.colW + 1 || first[1] === first[3], geo);
    check("cells covered by the merge are not drawn", !geo.coveredCellDrawn, geo);
  } else check("fixture has merged cells", false);
  await page.close();
}
{ // pictures and chart frames
  const page = await newPage(context);
  await open(page, "sheet-preview", `${T}/test-documents-excel/synthetic-drawing-extras.xlsx`);
  await loaded(page);
  await page.waitForTimeout(500);
  await shot(page, "preview-drawing-extras");
  const d = await g(page, () => ({ imgs: [...document.querySelectorAll(".spreadsheet-container img")].filter((i) => i.naturalWidth > 0).length, drawings: window.__grids[0].GetDrawings().map((x) => x.Kind ?? x.kind ?? Object.keys(x).join("|")), frames: [...document.querySelectorAll(".spreadsheet-container [class*=drawing]")].map((e) => e.className + ":" + e.textContent.slice(0, 30)) }));
  check("picture placed and decoded (the fixture holds a real PNG now)", d.frames.some((f) => /picture/.test(f)) && d.imgs >= 1, d);
  check("chart frame shown", d.frames.some((f) => /chart/i.test(f)), d.frames);
  await page.close();
}
{
  const page = await newPage(context);
  await open(page, "sheet-preview", `${T}/real-world/chart-zoo.xlsx`);
  await loaded(page);
  await page.waitForTimeout(500);
  await shot(page, "preview-chart-zoo");
  const frames = await g(page, () => [...document.querySelectorAll(".spreadsheet-container [class*=drawing]")].length);
  check("chart-zoo: chart frames shown", frames >= 6, frames);
  await page.close();
}
{ // a real picture
  const page = await newPage(context);
  await open(page, "sheet-preview", `${T}/test-documents-excel/phpsp-27template.xlsx`);
  await loaded(page);
  await page.waitForTimeout(700);
  await shot(page, "preview-pictures");
  const imgs = await g(page, () => [...document.querySelectorAll(".spreadsheet-container img")].map((i) => [i.naturalWidth, Math.round(i.getBoundingClientRect().width)]));
  check("pictures decoded and shown", imgs.some(([n, w]) => n > 0 && w > 0), imgs);
  await page.close();
}
{ // hidden sheets
  const page = await newPage(context);
  await open(page, "sheet-preview", `${S}/hidden-sheet.xlsx`);
  await loaded(page);
  const tabs = await page.locator('[data-testid="spreadsheet-preview-tabs"] button').allInnerTexts();
  const names = await g(page, () => window.__grids[0].GetSheetNames());
  check("hidden sheet has no tab", names.length === 2 && tabs.length === 0 && !(await page.getByRole("button", { name: "Monthly-CF" }).count()), { names, tabs });
  await page.close();
}
{ // CSV
  const page = await newPage(context);
  await open(page, "sheet-preview", `${S}/sample.csv`);
  await loaded(page);
  await page.waitForTimeout(200);
  await shot(page, "preview-csv");
  const r = await g(page, () => { const gr = window.__grids[0]; gr.SelectRange(0, 0, 3, 2); return { text: gr.GetSelectionText(), ro: gr.IsReadOnly() }; });
  check("CSV preview through the same component", r.ro && r.text.replace(/\r/g, "").trim() === "name\tqty\tprice\nwidget\t3\t1.5\nquoted, comma\t4\t2.25\nlast\t5\t10", JSON.stringify(r.text));
  await page.close();
}
{ // not a workbook
  const page = await newPage(context);
  await open(page, "sheet-preview", `${S}/sample.csv`, { name: "broken.xlsx" });
  await page.waitForSelector(".rb-file-viewer-error-copy", { timeout: 60000 });
  check("unreadable workbook shows the error state", (await page.locator(".rb-app-subtitle").innerText()).includes("Not a readable"), await page.locator(".rb-app-subtitle").innerText());
  await page.close();
}
summary();
await browser.close();
