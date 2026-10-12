// The package alone, 0.1.14 against 0.1.16, on this machine now (same load): import + painted, edits, wheel scroll over the
// dashboard (charts drawn in 0.1.16), export. usage: node vers-server.mjs & ; node perf16-vers.mjs 0.1.14 0.1.16
import { launch, newPage, T, fixtureUrl } from "./lib.mjs";
const versions = process.argv.slice(2);
const { browser } = await launch();
const med = (a) => a.slice().sort((x, y) => x - y)[a.length >> 1];
const fmt = (a) => `${med(a).toFixed(0)} (${a.map((x) => x.toFixed(0)).join("/")})`;
const run = async (version, file, steps, sheetName) => {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } }); const page = await newPage(ctx);
  await page.goto(`http://127.0.0.1:16142/`);
  await page.waitForFunction(() => window.__mounted === true, null, { timeout: 120000 });
  const out = await page.evaluate(async ([version, url, steps, sheetName]) => {
    const mod = await import(/* @vite-ignore */ `/pnpm/@rbt+office@${version}/node_modules/@rbt/office/dist/spreadsheet/Spreadsheet.js`);
    const host = document.createElement("div"); host.style.cssText = "position:fixed;left:0;top:0;width:1300px;height:600px;background:#fff"; document.body.appendChild(host);
    const data = await (await fetch(url)).arrayBuffer();
    const gr = mod.default(host); gr.Mount();
    await new Promise((r) => setTimeout(r, 300));
    const raf2 = () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    const out = {}; let t = performance.now();
    await gr.ImportXlsx(data); out["import"] = performance.now() - t;
    await raf2(); out["import+painted"] = performance.now() - t;
    await new Promise((r) => setTimeout(r, 300));
    if (sheetName) { const i = gr.GetSheetNames().indexOf(sheetName); if (i >= 0 && i !== gr.GetActiveSheetIndex()) { t = performance.now(); gr.SwitchSheet(i); await raf2(); out[`show ${sheetName}`] = performance.now() - t; await new Promise((r) => setTimeout(r, 200)); } }
    out.drawn = host.querySelectorAll(".spreadsheet-object-drawn").length; out.objects = host.querySelectorAll(".spreadsheet-object").length;
    for (const [name, r, c, text] of steps) { t = performance.now(); gr.EnterText(r, c, text); out[name] = performance.now() - t; }
    // wheel scroll: 60 steps of 120 px down, one per frame
    const sc = host.querySelector(".spreadsheet-container"); const work = [], gaps = []; let last = performance.now();
    for (let i = 0; i < 60; i++) { const t0 = performance.now(); sc.scrollTop += 120; sc.dispatchEvent(new Event("scroll")); work.push(performance.now() - t0); await new Promise((r) => requestAnimationFrame(r)); const now = performance.now(); gaps.push(now - last); last = now; }
    const q = (a, p) => { const s = [...a].sort((x, y) => x - y); return s[Math.min(s.length - 1, Math.floor(s.length * p))]; };
    out["scroll work median"] = q(work, 0.5); out["scroll work p95"] = q(work, 0.95); out["scroll frame median"] = q(gaps, 0.5); out["scroll frame p95"] = q(gaps, 0.95);
    sc.scrollTop = 0; sc.dispatchEvent(new Event("scroll")); await raf2();
    t = performance.now(); const bytes = await gr.ExportXlsx(); out["export"] = performance.now() - t;
    return out;
  }, [version, fixtureUrl(file), steps, sheetName]);
  await ctx.close();
  return out;
};
const cases = [
  ["milk-street-office.xlsx", `${T}/real-world/milk-street-office.xlsx`, [["E21 first", 20, 4, "21000000"], ["E21 second", 20, 4, "20500000"], ["E12 first", 11, 4, "0.07"], ["E12 second", 11, 4, "0.065"], ["E12 third", 11, 4, "0.07"]], null],
  ["dashboard-ecommerce-sales.xlsx (Dashboard sheet)", `${T}/real-world/dashboard-ecommerce-sales.xlsx`, [["edit", 60, 30, "edit"]], "Dashboard"],
  ["dashboard-sales-analysis.xlsx", `${T}/real-world/dashboard-sales-analysis.xlsx`, [["edit", 60, 30, "edit"]], "Dashboard"],
  ["chart-zoo.xlsx", `${T}/real-world/chart-zoo.xlsx`, [["edit B2 (the charts read it)", 1, 1, "12"]], null],
];
for (const [label, file, steps, sheet] of cases) {
  console.log(label);
  for (const version of versions) {
    await run(version, file, steps, sheet);
    const rows = []; for (let i = 0; i < 3; i++) rows.push(await run(version, file, steps, sheet));
    console.log(`  ${version}: ` + Object.keys(rows[0]).map((k) => (k === "drawn" || k === "objects" ? `${k} ${rows[0][k]}` : `${k} ${fmt(rows.map((r) => r[k]))}`)).join("; "));
  }
}
await browser.close();
