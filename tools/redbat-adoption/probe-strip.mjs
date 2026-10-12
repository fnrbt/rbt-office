import { launch, newPage, open, T, S, SHOTS } from "./lib.mjs";
import { g, ready, frames, sel } from "./ui-lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, process.argv[2] ?? "sheet-editor", `${S}/hidden-sheet.xlsx`);
if ((process.argv[2] ?? "sheet-editor") === "sheet-editor") await ready(page); else await page.waitForFunction(() => window.__grids.length > 0 && document.querySelector(".spreadsheet-cell"), null, { timeout: 60000 });
await page.waitForTimeout(500);
await g(page, () => { const gr = window.__grids[0]; gr.AddSheet?.(); gr.SwitchSheet(gr.GetSheetNames().length - 1); gr.GroupColumns("C:D"); gr.GroupRows("4:6"); if (window.location.search.includes("x")) {} });
const steps = [["summary above/left", (gr) => gr.SetOutlineSummary(false, false)], ["nested row group", (gr) => gr.GroupRows("5:5")], ["ungroup outer", (gr) => gr.UngroupRows("4:6")]];
const texts = () => g(page, () => { const out = []; const strip = [...document.querySelectorAll(".spreadsheet-outline-strip")]; const walk = (n) => { for (const c of n.childNodes) { if (c.nodeType === 3 && c.textContent.trim()) { const e = c.parentElement; const q = e.getBoundingClientRect(); out.push(`"${c.textContent.trim()}" in ${e.className} ${Math.round(q.left)},${Math.round(q.top)} ${Math.round(q.width)}x${Math.round(q.height)} font=${getComputedStyle(e).fontSize}`); } else if (c.nodeType === 1) walk(c); } }; strip.forEach(walk); const cont = document.querySelector(".spreadsheet-container"); const all = [...cont.querySelectorAll("*")].filter((e) => e.childElementCount === 0 && /^\d$/.test(e.textContent.trim()) && !/cell|header|outline-level/.test(e.className)).map((e) => { const q = e.getBoundingClientRect(); return `${e.tagName}.${e.className} "${e.textContent}" ${Math.round(q.left)},${Math.round(q.top)} ${Math.round(q.width)}x${Math.round(q.height)} font=${getComputedStyle(e).fontSize}`; }); return { out, all }; });
for (const [name, fn] of steps) { await g(page, (src) => new Function("gr", `return (${src})(gr)`)(window.__grids[0]), fn.toString()); await frames(page); console.log("after", name, JSON.stringify(await texts())); }
await frames(page);
const dump = () => g(page, () => [...document.querySelectorAll('[class*="spreadsheet-outline"]')].map((e) => { const q = e.getBoundingClientRect(); return `${e.className} [${e.dataset.axis ?? ""} ${e.dataset.level ?? ""} ${e.dataset.first ?? ""}-${e.dataset.last ?? ""}] "${e.textContent}" ${Math.round(q.left)},${Math.round(q.top)} ${Math.round(q.width)}x${Math.round(q.height)} font=${getComputedStyle(e).fontSize} overflow=${getComputedStyle(e).overflow}`; }));
console.log((await dump()).join("\n"));

await page.screenshot({ path: `${SHOTS}/probe-strip.png`, clip: { x: 0, y: 180, width: 600, height: 260 } }); console.log(JSON.stringify(await g(page, () => [...document.querySelectorAll(".spreadsheet-outline-level, .spreadsheet-outline-button")].map((e) => getComputedStyle(e).fontSize + e.textContent))));
await browser.close();
