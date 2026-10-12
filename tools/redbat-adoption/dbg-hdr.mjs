import { launch, newPage, open, T, SHOTS } from "./lib.mjs";
import { g, sel, ready, frames, cell, select, selection, cellPoint, revision } from "./ui-lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`);
await ready(page);
const R = 100, C = 17;
const data = [["Name", "Qty", "Price"], ["pear", "3", "1.5"], ["apple", "10", "2"]];
await g(page, ([R, C, data]) => { const gr = window.__grids[0]; data.forEach((row, i) => row.forEach((v, j) => gr.SetCell(R + i, C + j, v))); gr.ScrollToCell(R + 14, C + 6); gr.ScrollToCell(R - 2, C - 3); gr.SelectCell(R + 1, C); gr.DefineTable("R101:T103", "Fruit", true); gr.EnterText(R, C + 5, "=SUM(Fruit[Qty])"); }, [R, C, data]);
await frames(page);
{
  const p = await cellPoint(page, R + 3, C);
  await page.mouse.click(p.x, p.y);
  await page.keyboard.type("kiwi"); await page.keyboard.press("Tab"); await page.keyboard.type("5"); await page.keyboard.press("Enter");
  await frames(page);
  console.log("grown", await g(page, () => window.__grids[0].GetTables()[0].range));
  const variant = process.argv[2];
  if (variant === "button") { await sel(page, "sheet-undo").click(); await sel(page, "sheet-undo").click(); }
  else if (variant === "api") await g(page, () => { window.__grids[0].Undo(); window.__grids[0].Undo(); });
  else if (variant === "one") await g(page, () => { window.__grids[0].Undo(); });
  await frames(page);
  console.log("undone", await g(page, () => JSON.stringify(window.__grids[0].GetTables()[0])), "sel", await selection(page));
}
for (const how of ["mouse", "api"]) {
  if (how === "mouse") { const p = await cellPoint(page, R, C + 1); await page.mouse.click(p.x, p.y); } else await select(page, R, C + 1);
  await frames(page);
  console.log(how, "selection", await selection(page), "focus", await g(page, () => document.activeElement?.className));
  await page.keyboard.type(how === "mouse" ? "Count" : "Amount");
  console.log(" editing", await g(page, () => window.__grids[0].IsEditing()), await sel(page, "sheet-formula").inputValue());
  await page.keyboard.press("Enter");
  await frames(page);
  console.log(" cell", (await cell(page, R, C + 1)).text, "columns", await g(page, () => window.__grids[0].GetTables()[0].columns), "formula", (await cell(page, R, C + 5)).formula, "notice", await sel(page, "sheet-notice").innerText().catch(() => "-"));
}
await g(page, ([R, C]) => window.__grids[0].EnterText(R, C + 2, "Cost"), [R, C]);
console.log("EnterText: columns", await g(page, () => window.__grids[0].GetTables()[0].columns));
await page.screenshot({ path: `${SHOTS}/dbg-hdr.png` });
await browser.close();
