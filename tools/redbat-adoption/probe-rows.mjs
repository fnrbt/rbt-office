import { launch, newPage, open, T, S } from "./lib.mjs";
import { g, sel, ready, frames, status } from "./ui-lib.mjs";
const { browser, context } = await launch();
for (const f of ["lyric-multifamily", "saopaulo-condos", "milk-street-office", "chart-zoo"]) {
  const page = await newPage(context);
  await open(page, "sheet-editor", `${T}/real-world/${f}.xlsx`); await ready(page);
  const a = await g(page, () => { const gr = window.__grids[0]; const m = gr.GetModel(); const u = gr.GetUsedRange(); gr.SelectCell(u.lastRow, u.firstCol); gr.ScrollToCell(u.lastRow, u.firstCol); return { rows: m.RowCount, cols: m.ColCount, used: u.range }; });
  await page.locator('[data-testid="sheet-grid"] .spreadsheet-content').focus();
  for (let i = 0; i < 5; i++) await page.keyboard.press("ArrowDown");
  await page.keyboard.type("below"); await page.keyboard.press("Enter"); await frames(page);
  const b = await g(page, () => { const gr = window.__grids[0]; return { rows: gr.GetModel().RowCount, sel: gr.GetSelectionRef(), used: gr.GetUsedRange().range }; });
  console.log(f, JSON.stringify(a), "after 5x ArrowDown + typing:", JSON.stringify(b));
  if (f === "lyric-multifamily") await page.screenshot({ path: `${S}/explore/s-lyric-last-row.png` });
  await page.close();
}
await browser.close();
