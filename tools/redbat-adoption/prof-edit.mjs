// Where the host's time goes in a repeated edit of the input with the most dependents (milk-street E12):
// a CPU profile of the second / third EnterText in the host, split by who called what.
// usage: node prof-edit.mjs
import { launch, newPage, open, T } from "./lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`);
await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 180000 }); await page.waitForTimeout(600);
const cdp = await context.newCDPSession(page);
await cdp.send("Profiler.enable"); await cdp.send("Profiler.setSamplingInterval", { interval: 100 });
// time the public methods of the grid as the host calls them, by wrapping them on the instance
const wall = await page.evaluate(() => {
  const gr = window.__grids[0]; const spent = (window.__spent = {}); let depth = 0;
  for (const name of ["GetSelectionFormat", "GetStatus", "GetCell", "GetViewOptions", "GetFilters", "GetOutline", "GetShowOutline", "IsModified", "CanUndo", "CanRedo", "GetActiveSheetIndex", "GetSelectionRef", "GetSelection", "GetGhostFormula", "GetSheetKind", "IsEditing", "GetRevision", "GetSheetInfos", "GetDefinedNames", "GetCellText", "GetFormula"]) {
    const f = gr[name]; if (typeof f !== "function") continue;
    gr[name] = function (...a) { if (window.__inEdit !== true || depth > 0) return f.apply(this, a); depth++; const t = performance.now(); try { return f.apply(this, a); } finally { depth--; const s = (spent[name] ??= { ms: 0, calls: 0 }); s.ms += performance.now() - t; s.calls++; } };
  }
  gr.EnterText(11, 4, "0.07");   // the first edit builds the graph
  return null;
});
await page.waitForTimeout(300);
await cdp.send("Profiler.start");
const timing = await page.evaluate(() => {
  const gr = window.__grids[0]; const out = [];
  for (const text of ["0.065", "0.07", "0.065", "0.07"]) { for (const k of Object.keys(window.__spent)) delete window.__spent[k]; window.__inEdit = true; const t = performance.now(); gr.EnterText(11, 4, text); const ms = performance.now() - t; window.__inEdit = false; out.push({ ms: Math.round(ms * 10) / 10, host: JSON.parse(JSON.stringify(window.__spent)) }); }
  return out;
});
const { profile } = await cdp.send("Profiler.stop");
for (const t of timing) console.log(`EnterText(E12) in the host: ${t.ms} ms; of it in grid methods the host's listeners called: ${Object.entries(t.host).map(([k, v]) => `${k} ${v.ms.toFixed(1)} ms (${v.calls}x)`).join(", ")} = ${Object.values(t.host).reduce((n, v) => n + v.ms, 0).toFixed(1)} ms`);
// profile: self time by function, and inclusive time below the first frame of a host chunk
const nodes = new Map(profile.nodes.map((n) => [n.id, n])); const parent = new Map();
for (const n of profile.nodes) for (const c of n.children ?? []) parent.set(c, n.id);
const dt = profile.timeDeltas; const self = new Map(); let total = 0;
profile.samples.forEach((id, i) => { self.set(id, (self.get(id) ?? 0) + dt[i] / 1000); total += dt[i] / 1000; });
const chunk = (n) => (n.callFrame.url.split("/").pop() || "(native)").replace(/-[A-Za-z0-9_]{8}\.js$/, ".js");
const stackOf = (id) => { const s = []; for (let x = id; x != null; x = parent.get(x)) s.push(nodes.get(x)); return s.reverse(); };
const byWho = new Map(), byFn = new Map(), hostCalls = new Map();
for (const [id, ms] of self) {
  const st = stackOf(id); const leaf = st.at(-1);
  if (!st.some((n) => n.callFrame.functionName === "EnterText")) continue; // only the edits
  const hostAt = st.findIndex((n) => /SpreadsheetEditor|office-harness|OfficeEditorFrame|web|solid|index/.test(chunk(n)) && n.callFrame.url);
  const who = hostAt >= 0 ? "under the host's listeners" : "the engine's own";
  byWho.set(who, (byWho.get(who) ?? 0) + ms);
  if (hostAt >= 0) { const next = st.slice(hostAt).find((n) => /spreadsheet-engine|Spreadsheet/.test(chunk(n)) && !/SpreadsheetEditor/.test(chunk(n))); const key = next ? `${next.callFrame.functionName || "(anonymous)"} [engine, called by the host]` : `${leaf.callFrame.functionName || "(anonymous)"} [${chunk(leaf)}]`; hostCalls.set(key, (hostCalls.get(key) ?? 0) + ms); }
  const k = `${leaf.callFrame.functionName || "(anonymous)"} [${chunk(leaf)}:${leaf.callFrame.lineNumber}:${leaf.callFrame.columnNumber}]`; byFn.set(k, (byFn.get(k) ?? 0) + ms);
}
const top = (m, n) => [...m.entries()].sort((a, b) => b[1] - a[1]).slice(0, n).map(([k, v]) => `  ${v.toFixed(1)} ms  ${k}`).join("\n");
console.log(`profile of the 4 edits: ${total.toFixed(0)} ms sampled in all`);
console.log(top(byWho, 5));
console.log("under the host's listeners, by the engine method the host called:\n" + top(hostCalls, 12));
console.log("self time, top functions:\n" + top(byFn, 14));
await browser.close();
