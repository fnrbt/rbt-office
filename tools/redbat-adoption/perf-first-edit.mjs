// First edit on milk-street: the input measured before (E21) and the input with the most dependents.
import { launch, newPage, open, T } from "./lib.mjs";
const { browser } = await launch();
const file = `${T}/real-world/milk-street-office.xlsx`;
const fresh = async () => { const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } }); const page = await newPage(ctx); await open(page, "sheet-editor", file); await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 180000 }); await page.waitForTimeout(500); return { ctx, page }; };
// which numeric input cell reaches the most formulas
let best;
{ const { ctx, page } = await fresh();
  const found = await page.evaluate(() => { const gr = window.__grids[0]; const u = gr.GetUsedRange(); const out = []; for (let r = 0; r <= Math.min(u.lastRow, 120); r++) for (let c = 0; c <= Math.min(u.lastCol, 14); c++) { const cell = gr.GetCell(r, c); if (cell.valueType !== "number" || cell.formula) continue; const v = cell.value; gr.EnterText(r, c, String(v * 1.01)); const s = gr.GetRecalcStats(); out.push({ r, c, address: cell.address, reached: s.ReachedFormulas, evaluated: s.EvaluatedFormulas, value: v }); gr.Undo(); } out.sort((a, b) => b.evaluated - a.evaluated); return out.slice(0, 5); });
  console.log("inputs with the most dependents:", JSON.stringify(found.map((f) => `${f.address} (${f.evaluated} formulas)`)));
  best = found[0]; await ctx.close(); }
const med = (a) => a.slice().sort((x, y) => x - y)[a.length >> 1];
for (const target of [{ r: 20, c: 4, address: "E21", text: "21000000", back: "20500000" }, { ...best, text: String(best.value * 1.02), back: String(best.value) }]) {
  const first = [], second = [], keys = []; let stats;
  for (let run = 0; run < 5; run++) {
    const { ctx, page } = await fresh();
    const r = await page.evaluate(([r, c, text, back]) => { const gr = window.__grids[0]; let t = performance.now(); gr.EnterText(r, c, text); const a = performance.now() - t; const s = gr.GetRecalcStats(); t = performance.now(); gr.EnterText(r, c, back); const b = performance.now() - t; return { a, b, s: `${s.Kind}, reached ${s.ReachedFormulas}, evaluated ${s.EvaluatedFormulas} of ${s.Formulas}, evaluations ${s.Evaluations}` }; }, [target.r, target.c, target.text, target.back]);
    first.push(r.a); second.push(r.b); stats = r.s;
    await ctx.close();
  }
  console.log(`${target.address}: first edit ${first.map((x) => x.toFixed(0)).join(" / ")} ms (median ${med(first).toFixed(0)}); second edit ${second.map((x) => x.toFixed(0)).join(" / ")} ms (median ${med(second).toFixed(0)}); ${stats}`);
}
await browser.close();
