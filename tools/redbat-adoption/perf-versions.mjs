// The package alone, version against version, on this machine now: ImportXlsx, first edit, ExportXlsx.
// usage: node perf-versions.mjs 0.1.7 0.1.9 0.1.10
import { launch, newPage, BASE, T, fixtureUrl } from "./lib.mjs";
const versions = process.argv.slice(2);
const { browser } = await launch();
const med = (a) => a.slice().sort((x, y) => x - y)[a.length >> 1];
const fmt = (a) => `${a.map((x) => x.toFixed(0)).join(" / ")} (median ${med(a).toFixed(0)})`;
const run = async (version, file, steps) => {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } }); const page = await newPage(ctx);
  await page.goto(`http://127.0.0.1:16142/`);
  await page.waitForFunction(() => window.__mounted === true, null, { timeout: 120000 });
  const out = await page.evaluate(async ([version, url, steps]) => {
    const mod = await import(/* @vite-ignore */ `/pnpm/@rbt+office@${version}/node_modules/@rbt/office/dist/spreadsheet/Spreadsheet.js`);
    const host = document.createElement("div"); host.style.cssText = "position:fixed;left:0;top:0;width:1300px;height:600px;background:#fff"; document.body.appendChild(host);
    const data = await (await fetch(url)).arrayBuffer();
    const gr = mod.default(host); gr.Mount();
    await new Promise((r) => setTimeout(r, 300));
    const out = {}; let t = performance.now();
    await gr.ImportXlsx(data); out.import = performance.now() - t;
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    out.importPainted = performance.now() - t;
    await new Promise((r) => setTimeout(r, 300));
    for (const [name, r, c, text] of steps) { t = performance.now(); gr.EnterText(r, c, text); out[name] = performance.now() - t; }
    t = performance.now(); const bytes = await gr.ExportXlsx(); out.export = performance.now() - t; out.bytes = bytes.byteLength;
    return out;
  }, [version, fixtureUrl(file), steps]);
  await ctx.close();
  return out;
};
for (const [label, file, steps] of [["milk-street-office.xlsx", `${T}/real-world/milk-street-office.xlsx`, [["first edit E21", 20, 4, "21000000"], ["second edit E21", 20, 4, "20500000"]]], ["dashboard-sales-analysis.xlsx", `${T}/real-world/dashboard-sales-analysis.xlsx`, [["edit", 60, 30, "edit"]]]]) {
  console.log(label);
  for (const version of versions) {
    await run(version, file, steps); // modules compiled and cached by the dev server
    const rows = []; for (let i = 0; i < 3; i++) rows.push(await run(version, file, steps));
    const keys = Object.keys(rows[0]).filter((k) => k !== "bytes");
    console.log(`  ${version}: ` + keys.map((k) => `${k} ${fmt(rows.map((r) => r[k]))}`).join("; ") + `; ${rows[0].bytes} bytes`);
  }
}
await browser.close();
