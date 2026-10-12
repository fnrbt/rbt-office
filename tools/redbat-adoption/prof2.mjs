// CPU profile of the host from the document's bytes having arrived to first paint (+ extra ms).
// usage: node prof2.mjs <word-editor|chat-viewer|sheet-editor> [file] [extraMs] [until: ink|toolbar]
import { launch, newPage, open, S } from "./lib.mjs";
const component = process.argv[2] ?? "word-editor";
const file = process.argv[3] || `${S}/long-generated.docx`;
const extra = Number(process.argv[4] ?? 0);
const { browser, context } = await launch();
const page = await newPage(context);
const cdp = await context.newCDPSession(page);
await cdp.send("Profiler.enable"); await cdp.send("Profiler.setSamplingInterval", { interval: 150 });
await page.exposeFunction("__startProfile", async () => { await cdp.send("Profiler.start"); });
await page.addInitScript(() => {
  const marks = (window.__marks = []); const of = window.fetch;
  window.fetch = async (...a) => { const r = await of(...a); const name = String(a[0]); if (name.includes("/api/fixture/")) { const b = await r.arrayBuffer(); await window.__startProfile(); window.__t0 = performance.now(); return new Response(b, { status: r.status }); } marks.push([Math.round(performance.now() - (window.__t0 ?? 0)), "fetched " + name.split("/").pop().slice(0, 40), window.__t0 == null ? "(before bytes)" : ""]); return r; };
  window.__ink = null;
  const tick = () => {
    if (window.__ink == null) {
      const c = document.querySelector('[data-testid="word-canvas"]') ?? [...document.querySelectorAll(".rb-file-viewer-canvas")].find((c) => c.width >= 600);
      if (c && c.width > 0 && c.height > 0) { try { const d = c.getContext("2d").getImageData(0, 0, c.width, Math.min(c.height, 400)).data; let n = 0; for (let i = 0; i < d.length; i += 16) if (d[i + 3] > 0 && d[i] < 90 && d[i + 1] < 90 && d[i + 2] < 90) n++; if (n > 100) window.__ink = performance.now(); } catch {} }
      if (!c && document.querySelector('[data-testid="sheet-toolbar"]')) window.__ink = performance.now();
    }
    requestAnimationFrame(tick);
  };
  requestAnimationFrame(tick);
});
const mime = /\.docx$/.test(file) ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
await open(page, component, file, component === "chat-viewer" ? { params: { mime } } : {});
await page.waitForFunction(() => window.__ink != null, null, { timeout: 300000, polling: 10 });
if (extra) await page.waitForTimeout(extra);
const p = (await cdp.send("Profiler.stop")).profile;
const ink = await page.evaluate(() => Math.round(window.__ink - window.__t0));
const nodes = new Map(p.nodes.map((x) => [x.id, x])); const self = new Map(); p.samples.forEach((id, i) => self.set(id, (self.get(id) ?? 0) + p.timeDeltas[i] / 1000));
const parent = new Map(); for (const n of p.nodes) for (const c of n.children ?? []) parent.set(c, n.id);
const label = (n) => `${n.callFrame.functionName || "(anonymous)"} ${n.callFrame.url.split("/").pop().split("?")[0]}:${n.callFrame.lineNumber + 1}`;
let idle = 0, total = 0; for (const [id, m] of self) { total += m; if (/^\((idle|program)\)$/.test(nodes.get(id).callFrame.functionName)) idle += m; }
console.log(`${component}: bytes -> first paint ${ink} ms; profile ${Math.round(total)} ms, busy ${Math.round(total - idle)} ms`);
const bySelf = new Map(); for (const [id, m] of self) { const k = label(nodes.get(id)); bySelf.set(k, (bySelf.get(k) ?? 0) + m); }
console.log("top self time:"); for (const [k, m] of [...bySelf].sort((a, b) => b[1] - a[1]).slice(0, 25)) console.log(`  ${m.toFixed(0).padStart(6)} ms  ${k}`);
// inclusive time by function (counted once per stack)
const incl = new Map(); for (const [id, m] of self) { const seen = new Set(); for (let x = id; x != null; x = parent.get(x)) { const k = label(nodes.get(x)); if (!seen.has(k)) { seen.add(k); incl.set(k, (incl.get(k) ?? 0) + m); } } }
const want = new RegExp(process.argv[5] ?? ".");
console.log("top inclusive time:"); for (const [k, m] of [...incl].filter(([k]) => want.test(k) && !/^\((root|program|idle)\)/.test(k)).sort((a, b) => b[1] - a[1]).slice(0, 45)) console.log(`  ${m.toFixed(0).padStart(6)} ms  ${k}`);
console.log(JSON.stringify(await page.evaluate(() => window.__marks)));
await browser.close();
