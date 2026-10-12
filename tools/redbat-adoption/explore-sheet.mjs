// Exploratory pass, spreadsheets: open several workbooks, exercise the editor, take screenshots into explore/.
import { launch, newPage, open, T, S } from "./lib.mjs";
import { g, sel, ready, frames, status } from "./ui-lib.mjs";
const { browser, context } = await launch();
const E = `${S}/explore`;
const books = { charts: `${T}/real-world/chart-zoo.xlsx`, cf: `${T}/real-world/conditional-formatting-tutorial.xlsx`, sales: `${T}/real-world/dashboard-sales-analysis.xlsx`, lyric: `${T}/real-world/lyric-multifamily.xlsx`, condos: `${T}/real-world/saopaulo-condos.xlsx`, dcf: `${T}/real-world/dcf-stock-analysis.xlsx` };
const only = process.argv[2]?.split(",");
for (const [key, file] of Object.entries(books)) {
  if (only && !only.includes(key)) continue;
  const page = await newPage(context);
  const errors = [];
  page.on("pageerror", (e) => errors.push(String(e).slice(0, 200)));
  page.on("console", (m) => { if (m.type() === "error" && !/404/.test(m.text())) errors.push(m.text().slice(0, 200)); });
  await open(page, "sheet-editor", file); await ready(page); await page.waitForTimeout(900);
  const info = await g(page, () => { const gr = window.__grids[0]; const m = gr.GetModel(); const infos = gr.GetSheetInfos(); return { sheets: infos.map((i) => `${i.name}${i.kind !== "worksheet" ? "(" + i.kind + ")" : ""}${i.visibility !== "visible" ? "[" + i.visibility + "]" : ""}`), active: gr.GetActiveSheetIndex(), used: gr.GetUsedRange().range, frozen: [m.FrozenRows, m.FrozenCols], drawings: (gr.GetDrawings?.() ?? []).length, cf: gr.GetConditionalFormats().length, tables: (gr.GetTables?.() ?? []).length, filters: !!gr.GetFilters() }; });
  console.log(key, JSON.stringify(info));
  await page.screenshot({ path: `${E}/s-${key}-1-open.png` });
  // each visible worksheet: a look
  const tabs = await g(page, () => window.__grids[0].GetSheetInfos().filter((i) => i.visibility === "visible").map((i) => i.index));
  let shot = 2;
  for (const i of tabs.filter((i) => i !== info.active).slice(0, 2)) { await sel(page, `sheet-tab-${i}`).click(); await page.waitForTimeout(900); await page.screenshot({ path: `${E}/s-${key}-${shot++}-tab${i}.png` }); }
  await sel(page, `sheet-tab-${info.active}`).click(); await page.waitForTimeout(500);
  // typing, copy / paste, undo / redo
  const u = await g(page, () => window.__grids[0].GetUsedRange());
  const spot = await g(page, () => { const gr = window.__grids[0]; const u = gr.GetUsedRange(); for (let r = u.firstRow; r <= Math.min(u.lastRow, u.firstRow + 40); r++) for (let c = u.firstCol; c <= Math.min(u.lastCol, u.firstCol + 12); c++) { const x = gr.GetCell(r, c); if (x.valueType !== "empty" && !x.isMerged && !x.formula && gr.GetCell(r + 1, c).valueType !== "empty" && !gr.GetCell(r + 1, c).isMerged) return [r, c]; } return [u.firstRow, u.firstCol]; });
  await g(page, ([r, c]) => window.__grids[0].SelectRange(r, c, r + 1, c), spot); await page.locator('[data-testid="sheet-grid"] .spreadsheet-content').focus();
  await page.keyboard.press("Control+c");
  const dest = [Math.min(u.lastRow + 2, 1048570), u.firstCol];
  let pasted = "n/a";
  try { await g(page, ([r, c]) => { const gr = window.__grids[0]; gr.SelectCell(r, c); gr.ScrollToCell(r, c); }, dest); await page.keyboard.press("Control+v"); await frames(page); await page.keyboard.press("ArrowDown"); await page.keyboard.press("ArrowDown"); await page.keyboard.type("typed 123"); await page.keyboard.press("Enter"); await frames(page);
    pasted = await g(page, ([[r, c], [sr, sc]]) => { const gr = window.__grids[0]; return [gr.GetCell(r, c).text === gr.GetCell(sr, sc).text, gr.GetCell(r + 1, c).text === gr.GetCell(sr + 1, sc).text, gr.GetCell(r + 2, c).text].join(); }, [dest, spot]);
  } catch (e) { pasted = "error " + String(e).slice(0, 120); }
  await page.screenshot({ path: `${E}/s-${key}-${shot++}-pasted-typed.png` });
  let n = 0; while (!(await sel(page, "sheet-undo").isDisabled()) && n++ < 20) await sel(page, "sheet-undo").click();
  const clean = /No changes/.test(await status(page));
  for (let i = 0; i < n; i++) await sel(page, "sheet-redo").click();
  console.log(`  ${key}: copy ${spot} x2 rows -> paste at ${dest} + typed: ${pasted}; ${n} undo steps back to "No changes": ${clean}; after redo: ${(await status(page)).replace(/\n/g, " ")}`);
  // find
  await sel(page, "sheet-find").click(); await sel(page, "sheet-find-input").fill("a"); await sel(page, "sheet-find-next").click(); await sel(page, "sheet-find-next").click(); await frames(page);
  console.log(`  ${key}: find "a": ${await sel(page, "sheet-find-status").innerText()}`);
  await sel(page, "sheet-find-close").click();
  // zoom, the top of the sheet
  await g(page, () => window.__grids[0].ScrollToCell(0, 0)); await sel(page, "sheet-zoom-in").click(); await sel(page, "sheet-zoom-in").click(); await page.waitForTimeout(700);
  await page.screenshot({ path: `${E}/s-${key}-${shot++}-zoom.png` });
  await sel(page, "sheet-zoom-out").click(); await sel(page, "sheet-zoom-out").click(); await sel(page, "sheet-zoom-out").click();
  // the end of the used range, in a small window
  await g(page, () => { const gr = window.__grids[0]; const u = gr.GetUsedRange(); gr.ScrollToCell(u.lastRow, u.lastCol); }); await page.setViewportSize({ width: 820, height: 620 }); await page.waitForTimeout(700);
  await page.screenshot({ path: `${E}/s-${key}-${shot++}-end-small.png` });
  await page.setViewportSize({ width: 1366, height: 768 });
  console.log(`  ${key}: errors ${JSON.stringify(errors)}`);
  await page.close();
}
await browser.close();
