// Long tasks while a long document is read behind the editor, and what the engine told meanwhile.
import { launch, newPage, open, S } from "./lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await page.addInitScript(() => { window.__lt = []; new PerformanceObserver((l) => { for (const e of l.getEntries()) window.__lt.push([Math.round(e.startTime), Math.round(e.duration)]); }).observe({ entryTypes: ["longtask"] }); });
await open(page, process.argv[2] ?? "word-editor", `${S}/long-generated.docx`, { params: { hookWord: "1" } });
await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 60000 }); console.log("toolbar at", await page.evaluate(() => [Math.round(performance.now()), window.__word.isLoadComplete(window.__editors[0])]));
await page.evaluate(() => { const w = window.__word, e = window.__editors[0]; window.__notes = []; const t = () => Math.round(performance.now()); let last = null;
  w.onChange(e, (c) => { const k = `${c.DocumentChanged ? "D" : ""}${c.SelectionChanged ? "S" : ""}${c.LayoutChanged ? "L" : ""}${c.LoadChanged ? "R" : ""}${c.LoadComplete ? "!" : ""}${c.LayoutComplete ? "#" : ""}`; if (k !== last) { window.__notes.push([t(), k]); last = k; } }); });
await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 });
await page.waitForTimeout(300);
console.log(JSON.stringify(await page.evaluate(() => ({ bytes: Math.round(performance.getEntriesByType("resource").find((e) => e.name.includes("/api/fixture/")).responseEnd), lt: window.__lt.filter(([, d]) => d > 60), notes: window.__notes }))));
await browser.close();
