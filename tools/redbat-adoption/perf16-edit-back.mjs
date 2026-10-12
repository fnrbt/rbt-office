// E12 edited away from the value the file has and back to it, in the host: is the edit back to the saved value dearer?
import { launch, newPage, open, T } from "./lib.mjs";
const { browser } = await launch();
const med = (a) => a.slice().sort((x, y) => x - y)[a.length >> 1];
const rows = [];
for (let run = 0; run < 5; run++) {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } }); const page = await newPage(ctx);
  await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`); await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 180000 }); await page.waitForTimeout(500);
  rows.push(await page.evaluate(() => { const gr = window.__grids[0]; const v = gr.GetCell(11, 4).value; const time = (text, stats) => { const t = performance.now(); gr.EnterText(11, 4, text); const ms = performance.now() - t; if (stats) gr.GetRecalcStats(); return ms; };
    const a = String(v * 1.02), b = String(v * 1.03), o = String(v);
    return { v, t: [time(a, true), time(o, false), time(a), time(o), time(a), time(b), time(a), time(o)], mod: gr.IsModified() }; }));
  await ctx.close();
}
const names = ["away (first)", "BACK to the file's value", "away", "BACK", "away", "to another value", "to the first", "BACK"];
console.log(`E12 = ${rows[0].v}; medians of 5: ` + names.map((n, i) => `${n} ${med(rows.map((r) => r.t[i])).toFixed(0)} ms`).join("; "));
console.log(rows.map((r) => r.t.map((x) => x.toFixed(0)).join("/")).join("   "));
await browser.close();
