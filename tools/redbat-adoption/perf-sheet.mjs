// Spreadsheet timings: first edit and paste on milk-street, open / save of the dashboard workbook.
import { launch, newPage, open, T, waitUploads } from "./lib.mjs";
const { browser } = await launch();
const g = (page, fn, arg) => page.evaluate(fn, arg);
const ready = async (page) => { await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 180000 }); };
const frames = (page) => g(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
for (let run = 0; run < 2; run++) {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 }, permissions: ["clipboard-read", "clipboard-write"] });
  const page = await newPage(ctx);
  let t = Date.now();
  await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`);
  await ready(page);
  console.log(`milk-street run ${run + 1}: open ${Date.now() - t} ms`);
  await page.waitForTimeout(400);
  const timed = (label, fn, arg) => g(page, ([src, arg]) => { const gr = window.__grids[0]; const f = new Function("gr", "arg", `return (${src})(gr, arg)`); const t = performance.now(); f(gr, arg); const ms = performance.now() - t; const s = gr.GetRecalcStats ? gr.GetRecalcStats() : null; return { ms: Math.round(ms * 10) / 10, recalc: s ? `${s.Kind}${s.Reason ? " (" + s.Reason + ")" : ""}, changed ${s.ChangedCells}, evaluated ${s.EvaluatedFormulas} of ${s.Formulas}` : "n/a" }; }, [fn.toString(), arg]).then((r) => console.log(`  ${label}: ${r.ms} ms; recalculation ${r.recalc}`));
  // E21 is the acquisition price: an input most of the model depends on
  await timed("first edit: an input the model depends on (E21 acquisition price -> 21,000,000)", (gr) => gr.EnterText(20, 4, "21000000"));
  await timed("second edit of the same input (back to 20,500,000)", (gr) => gr.EnterText(20, 4, "20500000"));
  await timed("edit of an empty cell nothing depends on (U61)", (gr) => gr.EnterText(60, 20, "x"));
  await timed("a new formula over the model (=SUM(E21:E60) in U62)", (gr) => gr.EnterText(61, 20, "=SUM(E21:E60)"));
  // by keyboard, as a user: wall time from the first key to the frame after Enter
  await g(page, () => window.__grids[0].SelectCell(20, 4));
  await page.locator(".spreadsheet-content").first().focus();
  t = Date.now(); await page.keyboard.type("21500000"); await page.keyboard.press("Enter"); await frames(page);
  console.log(`  typed into E21 and Enter (9 keys): ${Date.now() - t} ms wall`);
  // paste: copy 20 x 12 model cells (formulas) and paste them into an empty area, via the toolbar (system clipboard)
  await g(page, () => window.__grids[0].SelectRange(20, 2, 39, 13));
  await page.click('[data-testid="sheet-copy"]'); await page.waitForTimeout(300);
  await g(page, () => { const gr = window.__grids[0]; gr.SelectCell(100, 2); window.__pasteDone = null; const stop = gr.OnChange(() => { window.__pasteDone = performance.now(); stop(); }); window.__pasteStart = performance.now(); });
  t = Date.now(); await page.click('[data-testid="sheet-paste"]');
  await page.waitForFunction(() => window.__pasteDone != null, null, { timeout: 60000 }); await frames(page);
  const stats = await g(page, () => { const gr = window.__grids[0]; const s = gr.GetRecalcStats ? gr.GetRecalcStats() : null; return s ? `${s.Kind}${s.Reason ? " (" + s.Reason + ")" : ""}, changed ${s.ChangedCells}, evaluated ${s.EvaluatedFormulas} of ${s.Formulas}` : "n/a"; });
  console.log(`  paste of 20 x 12 cells (240, with formulas): ${Date.now() - t} ms wall click to painted; recalculation ${stats}`);
  await g(page, () => { window.__grids[0].SelectRange(100, 2, 100, 3); });
  await timed("undo of the paste", (gr) => gr.Undo());
  await timed("define a name (Price = Annual!$E$21)", (gr) => gr.SetDefinedName("PerfPrice", gr.GetSheetNames()[gr.GetActiveSheetIndex()].replace(/^(.*\W.*)$/, "'$1'") + "!$E$21", null));
  await timed("a formula that reads the name", (gr) => gr.EnterText(62, 20, "=PerfPrice*2"));
  await timed("delete the name", (gr) => gr.DeleteDefinedName("PerfPrice", null));
  await timed("undo of that", (gr) => gr.Undo());
  await timed("set an AutoFilter on the used range", (gr) => { const u = gr.GetUsedRange(); const col = (n) => { let s = ""; for (n = n + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; }; gr.SetAutoFilter(`${col(u.firstCol)}20:${col(u.lastCol)}${u.lastRow + 1}`); });
  await timed("filter a column (E greater than 0)", (gr) => gr.SetFilter(4, { conditions: [{ operator: "greaterThan", value: 0 }] }));
  await timed("clear the filter", (gr) => gr.ClearAutoFilter());
  await g(page, () => window.__grids[0].SelectCell(30, 2));
  await timed("insert a row at 31", (gr) => gr.InsertRows());
  await timed("undo of the row insert", (gr) => gr.Undo());
  await timed("delete row 31", (gr) => gr.DeleteRows());
  await timed("undo of the row delete", (gr) => gr.Undo());
  await timed("insert a column at C", (gr) => gr.InsertColumns());
  await timed("undo of the column insert", (gr) => gr.Undo());
  t = Date.now(); await page.click('[data-testid="office-editor-save"]'); await waitUploads(page, 1); console.log(`  save: ${Date.now() - t} ms`);
  await ctx.close();
}
for (let run = 0; run < 2; run++) {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });
  const page = await newPage(ctx);
  let t = Date.now();
  await open(page, "sheet-editor", `${T}/real-world/dashboard-sales-analysis.xlsx`);
  await ready(page);
  console.log(`dashboard run ${run + 1}: open ${Date.now() - t} ms`);
  await page.waitForTimeout(400);
  const loaded = () => g(page, () => { const gr = window.__grids[0]; return gr.GetSheetNames().map((n, i) => (gr.IsSheetLoaded(i) ? 1 : 0)).join(""); });
  const r = await g(page, () => { const gr = window.__grids[0]; const t = performance.now(); gr.EnterText(60, 30, "edit"); const ms = performance.now() - t; const s = gr.GetRecalcStats ? gr.GetRecalcStats() : null; return { ms: Math.round(ms), s: s ? `${s.Kind}${s.Reason ? " (" + s.Reason + ")" : ""}, evaluated ${s.EvaluatedFormulas} of ${s.Formulas}` : "n/a" }; });
  console.log(`  first edit on the Dashboard sheet: ${r.ms} ms; recalculation ${r.s}; sheets loaded ${await loaded()}`);
  t = Date.now(); await page.click('[data-testid="office-editor-save"]'); await waitUploads(page, 1, 180000); console.log(`  save with Data not loaded: ${Date.now() - t} ms (${page.uploads[0].length} bytes); sheets loaded ${await loaded()}`);
  t = Date.now(); await page.getByRole("button", { name: "Data", exact: true }).click(); await frames(page); console.log(`  showing the Data sheet (76,931 rows): ${Date.now() - t} ms`);
  const r2 = await g(page, () => { const gr = window.__grids[0]; const t = performance.now(); gr.EnterText(5, 20, "edit2"); const ms = performance.now() - t; const s = gr.GetRecalcStats ? gr.GetRecalcStats() : null; return { ms: Math.round(ms), s: s ? `${s.Kind}${s.Reason ? " (" + s.Reason + ")" : ""}, evaluated ${s.EvaluatedFormulas} of ${s.Formulas}` : "n/a" }; });
  console.log(`  edit on the Data sheet: ${r2.ms} ms; recalculation ${r2.s}`);
  t = Date.now(); await page.click('[data-testid="office-editor-save"]'); await waitUploads(page, 2, 180000); console.log(`  save with Data loaded and edited: ${Date.now() - t} ms`);
  await ctx.close();
  const ctx2 = await browser.newContext({ viewport: { width: 1366, height: 768 } });
  const p2 = await newPage(ctx2);
  t = Date.now(); await open(p2, "sheet-preview", `${T}/real-world/dashboard-sales-analysis.xlsx`);
  await p2.waitForFunction(() => document.querySelector('[data-testid="spreadsheet-preview"]')?.style.visibility === "visible", null, { timeout: 180000 });
  console.log(`  preview open: ${Date.now() - t} ms`);
  await ctx2.close();
}
await browser.close();
