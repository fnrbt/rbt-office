// Is the host's "second edit of E12" (199 ms on 0.1.14, 202 ms on 0.1.16) the host's doing? The same edits in the host,
// with and without two edits of E21 before them (as the package's own measurement has them), and with the host's listeners counted.
import { launch, newPage, open, T } from "./lib.mjs";
const { browser } = await launch();
const med = (a) => a.slice().sort((x, y) => x - y)[a.length >> 1];
for (const warm of [false, true]) {
  const rows = [];
  for (let run = 0; run < 5; run++) {
    const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } }); const page = await newPage(ctx);
    await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`); await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 180000 }); await page.waitForTimeout(500);
    rows.push(await page.evaluate((warm) => { const gr = window.__grids[0]; const time = (r, c, text) => { const t = performance.now(); gr.EnterText(r, c, text); return performance.now() - t; }; if (warm) { time(20, 4, "21000000"); time(20, 4, "20500000"); } return [time(11, 4, "0.07"), time(11, 4, "0.065"), time(11, 4, "0.07"), time(11, 4, "0.065")]; }, warm));
    await ctx.close();
  }
  console.log(`in the host, ${warm ? "after two edits of E21" : "E12 as the first edit of the session"}: E12 first ${med(rows.map((r) => r[0])).toFixed(0)} ms, second ${med(rows.map((r) => r[1])).toFixed(0)} ms, third ${med(rows.map((r) => r[2])).toFixed(0)} ms, fourth ${med(rows.map((r) => r[3])).toFixed(0)} ms   (${rows.map((r) => r.map((x) => x.toFixed(0)).join("/")).join("  ")})`);
}
await browser.close();
