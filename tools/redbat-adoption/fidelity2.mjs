// One edit, then save: which parts of the workbook change.
import { launch, newPage, open, T, S, waitUploads, unzip } from "./lib.mjs";
import fs from "node:fs";
const { browser, context } = await launch();
for (const file of [`${T}/real-world/dashboard-ecommerce-sales.xlsx`, `${T}/real-world/milk-street-office.xlsx`, `${T}/real-world/chart-zoo.xlsx`]) {
  const name = file.split("/").pop();
  const original = await unzip(fs.readFileSync(file));
  const page = await newPage(context);
  await open(page, "sheet-editor", file);
  await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 300000 }); await page.waitForTimeout(300);
  const where = await page.evaluate(() => { const g = window.__grids[0]; const u = g.GetUsedRange(); const r = (u?.lastRow ?? 0) + 3; g.EnterText(r, 0, "edited"); return { sheet: g.GetSheetNames()[g.GetActiveSheetIndex()], row: r }; });
  await page.click('[data-testid="office-editor-save"]');
  const data = await waitUploads(page, 1, 300000);
  fs.writeFileSync(`${S}/out/edited-${name}`, data);
  const zip = await unzip(data);
  const dropped = [], changed = [], added = [];
  for (const n of original.keys()) { if (!zip.has(n)) { dropped.push(n); continue; } const x = Buffer.from(await original.get(n).async("uint8array")), y = Buffer.from(await zip.get(n).async("uint8array")); if (!x.equals(y)) changed.push(`${n} (${x.length} -> ${y.length})`); }
  for (const n of zip.keys()) if (!original.has(n)) added.push(n);
  console.log(`${name}: "edited" typed into A${where.row + 1} of ${where.sheet}; ${original.size - dropped.length - changed.length}/${original.size} parts byte-identical\n  dropped: ${JSON.stringify(dropped)}\n  changed: ${JSON.stringify(changed)}\n  added: ${JSON.stringify(added)}`);
  await page.close();
}
await browser.close();
