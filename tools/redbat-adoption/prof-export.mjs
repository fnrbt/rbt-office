import { launch, newPage, open, T } from "./lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
const cdp = await context.newCDPSession(page);
await cdp.send("Profiler.enable"); await cdp.send("Profiler.setSamplingInterval", { interval: 300 });
await open(page, "sheet-editor", `${T}/real-world/dashboard-sales-analysis.xlsx`);
await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 180000 }); await page.waitForTimeout(400);
await page.evaluate(() => window.__grids[0].EnterText(60, 30, "edit"));
await cdp.send("Profiler.start");
const ms = await page.evaluate(async () => { const t = performance.now(); await window.__grids[0].ExportXlsx(); return performance.now() - t; });
const p = (await cdp.send("Profiler.stop")).profile;
const nodes = new Map(p.nodes.map((x) => [x.id, x])); const self = new Map(); p.samples.forEach((id, i) => self.set(id, (self.get(id) ?? 0) + p.timeDeltas[i] / 1000));
const parent = new Map(); for (const n of p.nodes) for (const c of n.children ?? []) parent.set(c, n.id);
const label = (n) => `${n.callFrame.functionName || "(anonymous)"} ${n.callFrame.url.split("/").pop().split("?")[0]}:${n.callFrame.lineNumber + 1}`;
const incl = new Map(); for (const [id, m] of self) { const seen = new Set(); for (let x = id; x != null; x = parent.get(x)) { const k = label(nodes.get(x)); if (!seen.has(k)) { seen.add(k); incl.set(k, (incl.get(k) ?? 0) + m); } } }
console.log(`ExportXlsx ${ms.toFixed(0)} ms`);
for (const [k, m] of [...incl].filter(([k]) => !/^\((root|program|idle)\)/.test(k) && /[A-Z][a-z]+_|Xlsx|export|Export|write|Write|generate|deflate|garbage|scan|kept|sheet|Sheet|string|String/.test(k)).sort((a, b) => b[1] - a[1]).slice(0, 40)) console.log(`  ${m.toFixed(0).padStart(6)} ms  ${k}`);
await browser.close();
