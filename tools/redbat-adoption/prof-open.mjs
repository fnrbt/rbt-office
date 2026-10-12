import { launch, newPage, open, S } from "./lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await page.addInitScript(() => { const marks = (window.__marks = []); const t0 = performance.now(); const of = window.fetch; window.fetch = async (...a) => { const r = await of(...a); marks.push([Math.round(performance.now() - t0), "fetched " + String(a[0]).split("/").pop().slice(0, 30)]); return r; }; });
const cdp = await context.newCDPSession(page);
await cdp.send("Profiler.enable"); await cdp.send("Profiler.setSamplingInterval", { interval: 200 });
await cdp.send("Profiler.start");
await open(page, "chat-viewer", `${S}/long-generated.docx`, { params: { mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } });
await page.waitForFunction(() => [...document.querySelectorAll(".rb-file-viewer-canvas")].some((c) => c.width >= 800), null, { timeout: 300000, polling: 30 });
const p = (await cdp.send("Profiler.stop")).profile;
const nodes = new Map(p.nodes.map((x) => [x.id, x])); const self = new Map(); p.samples.forEach((id, i) => self.set(id, (self.get(id) ?? 0) + p.timeDeltas[i] / 1000));
const parent = new Map(); for (const n of p.nodes) for (const c of n.children ?? []) parent.set(c, n.id);
const inc = (re) => { let ms = 0; for (const [id, m] of self) { for (let x = id; x != null; x = parent.get(x)) { const n = nodes.get(x); if (re.test(`${n.callFrame.functionName} ${n.callFrame.url}`)) { ms += m; break; } } } return Math.round(ms); };
let idle = 0, total = 0; for (const [id, m] of self) { total += m; if (nodes.get(id).callFrame.functionName === "(idle)") idle += m; }
console.log(`until the first page is painted: busy ${Math.round(total - idle)} ms, idle ${Math.round(idle)} ms; reader (readEditorPackage) ${inc(/^readEditorPackage /)}, collectImageSources ${inc(/^collectImageSources /)}, getRequiredFonts ${inc(/^getRequiredFonts /)}, font set (rebuildFontSet) ${inc(/^rebuildFontSet /)}, viewer.load ${inc(/^(state\.load|load) .*WordViewer/)}, importFromEditorPackage ${inc(/^importFromEditorPackage /)}, layout slices (continueLayout/scheduleLayout) ${inc(/^(continueLayout|scheduleLayout|layoutSlice)/)}, renderPage ${inc(/^renderPage /)}, GC ${inc(/^\(garbage collector\)/)}, module evaluation/compile ${inc(/^\(anonymous\) .*\.vite\/deps\/[^:]*$/)}`);
console.log(JSON.stringify(await page.evaluate(() => window.__marks)));
await browser.close();
