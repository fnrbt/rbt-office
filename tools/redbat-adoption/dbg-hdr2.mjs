import { launch, newPage, open, T, SHOTS } from "./lib.mjs";
import { g, sel, ready, frames, cell, select, selection, cellPoint, revision } from "./ui-lib.mjs";
const { browser, context } = await launch();
const steps = { rename: (gr) => gr.RenameTable("Fruit", "Fruit2"), style: (gr) => gr.SetTableStyle("Fruit", { name: "TableStyleDark3", showRowStripes: true }), totalOn: (gr) => gr.SetTableTotalRow("Fruit", true), totalOnOff: (gr) => { gr.SetTableTotalRow("Fruit", true); gr.SetTableTotalRow("Fruit", false); }, none: () => {} };
for (const variant of process.argv.slice(2)) {
  const page = await newPage(context);
  await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`);
  await ready(page);
  const R = 100, C = 17;
  const data = [["Name", "Qty", "Price"], ["pear", "3", "1.5"], ["apple", "10", "2"]];
  await g(page, ([R, C, data]) => { const gr = window.__grids[0]; data.forEach((row, i) => row.forEach((v, j) => gr.SetCell(R + i, C + j, v))); gr.ScrollToCell(R + 14, C + 6); gr.ScrollToCell(R - 2, C - 3); gr.SelectCell(R + 1, C); gr.DefineTable("R101:T103", "Fruit", true); }, [R, C, data]);
  await page.evaluate(`(${steps[variant].toString()})(window.__grids[0])`);
  await g(page, ([R, C]) => window.__grids[0].EnterText(R, C + 1, "Count"), [R, C]);
  await frames(page);
  console.log(variant, "cell", (await cell(page, R, C + 1)).text, "table", await g(page, () => JSON.stringify(window.__grids[0].GetTables().map((t) => [t.name, t.range, t.columns, t.hasTotals]))));
  await page.close();
}
await browser.close();
