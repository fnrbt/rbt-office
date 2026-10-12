// Word editor on the 522-page document with the real text stack: open, caret moves, keystrokes.
import { launch, newPage, open, T, S, SHOTS } from "./lib.mjs";
const { browser } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const file = `${S}/long-generated.docx`;
const stats = (page) => ev(page, () => window.__wordModule.getRenderStats(window.__editors[0]));
const diff = (a, b) => Object.fromEntries(Object.keys(b).filter((k) => typeof b[k] === "number").map((k) => [k, b[k] - a[k]]).filter(([k, v]) => v !== 0 && k !== "Blocks"));
for (let run = 0; run < 2; run++) {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });
  const page = await newPage(ctx);
  let t = Date.now();
  await open(page, "word-editor", file);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  console.log(`editor open (run ${run + 1}, ${run ? "warm http cache: no" : "cold"}): ${Date.now() - t} ms`, JSON.stringify(await stats(page)));
  if (run === 0) {
    await page.waitForTimeout(500);
    await ev(page, () => window.__word.getInputElement(window.__editors[0]).focus());
    for (const where of ["start", "middle", "end"]) {
      if (where === "start") await page.keyboard.press("Control+Home");
      if (where === "end") await page.keyboard.press("Control+End");
      if (where === "middle") await ev(page, () => { const w = window.__word, e = window.__editors[0]; const lines = w.getLayoutLines(e); const l = lines[lines.length >> 1]; w.setSelection(e, l.BlockId, 1, l.BlockId, 1); });
      await page.waitForTimeout(400);
      // real key events through the browser: wall time per key, and the longest main-thread task
      const longTasks = () => ev(page, () => { window.__lt = []; window.__lto?.disconnect(); window.__lto = new PerformanceObserver((l) => { for (const e of l.getEntries()) window.__lt.push(Math.round(e.duration)); }); window.__lto.observe({ entryTypes: ["longtask"] }); });
      const eventTimes = (types) => ev(page, (types) => { window.__et = []; window.__eto?.disconnect(); window.__eto = new PerformanceObserver((l) => { for (const e of l.getEntries()) if (types.includes(e.name)) window.__et.push([e.name, Math.round(e.processingEnd - e.processingStart), Math.round(e.duration)]); }); window.__eto.observe({ type: "event", durationThreshold: 16, buffered: false }); }, types);
      await longTasks(); await eventTimes(["keydown", "input", "beforeinput", "keypress"]);
      let s0 = await stats(page); t = Date.now();
      for (let i = 0; i < 20; i++) await page.keyboard.press(i < 10 ? "ArrowRight" : "ArrowDown");
      await ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
      let ms = Date.now() - t; let s1 = await stats(page);
      console.log(`  ${where}: 20 caret moves ${ms} ms (${(ms / 20).toFixed(1)} ms/key wall incl. automation); stats delta ${JSON.stringify(diff(s0, s1))}; long tasks ${JSON.stringify(await ev(page, () => window.__lt))}; slow events ${JSON.stringify(await ev(page, () => window.__et))}`);
      await longTasks(); await eventTimes(["keydown", "input", "beforeinput", "keypress"]);
      s0 = await stats(page); t = Date.now();
      await page.keyboard.type("The quick brown fox ");
      await ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
      ms = Date.now() - t; s1 = await stats(page);
      console.log(`  ${where}: 20 keystrokes ${ms} ms (${(ms / 20).toFixed(1)} ms/key wall); stats delta ${JSON.stringify(diff(s0, s1))}; long tasks ${JSON.stringify(await ev(page, () => window.__lt))}; slow events ${JSON.stringify(await ev(page, () => window.__et))}`);
      // the engine alone, synchronously, and the host's share
      const inner = await ev(page, () => {
        const w = window.__word, e = window.__editors[0]; const sel = w.getSelection(e); const id = sel.FocusBlockId; const o = sel.FocusOffset;
        const time = (n, fn) => { const t = performance.now(); for (let i = 0; i < n; i++) fn(i); return Math.round(((performance.now() - t) / n) * 100) / 100; };
        return {
          "setSelection (engine + host change listener)": time(20, (i) => w.setSelection(e, id, Math.max(0, o - 1 - (i % 5)), id, Math.max(0, o - 1 - (i % 5)))),
          "insertText 1 char (engine + host change listener)": time(20, () => w.insertText(e, "x")),
          "host: getSelectionFormat": time(50, () => w.getSelectionFormat(e)),
          "host: getCaretContext": time(50, () => w.getCaretContext(e)),
          "host: getDocumentSize": time(50, () => w.getDocumentSize(e)),
          "host: getCaretRect": time(50, () => w.getCaretRect(e)),
          "host: setViewport (repaint)": time(20, () => w.setViewport(e, w.getViewport(e))),
        };
      });
      console.log(`  ${where}: per call ms ${JSON.stringify(inner)}`);
      // Enter (a new paragraph) and Backspace over it
      s0 = await stats(page); t = Date.now();
      await page.keyboard.press("Enter"); await page.keyboard.press("Backspace");
      await ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
      console.log(`  ${where}: Enter + Backspace ${Date.now() - t} ms; stats delta ${JSON.stringify(diff(s0, await stats(page)))}`);
    }
    // select all + bold (a whole-document change), undo
    t = Date.now(); await page.keyboard.press("Control+a"); await ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))); console.log(`  select all: ${Date.now() - t} ms`);
    t = Date.now(); await page.keyboard.press("Control+b"); await ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))); console.log(`  bold on the whole document: ${Date.now() - t} ms`);
    t = Date.now(); await page.keyboard.press("Control+z"); await ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)))); console.log(`  undo of that: ${Date.now() - t} ms`);
    // save
    t = Date.now(); await page.keyboard.press("Control+s");
    while (page.uploads.length < 1) await new Promise((r) => setTimeout(r, 50));
    console.log(`  save (export + write + verify re-open + upload): ${Date.now() - t} ms, ${page.uploads[0].length} bytes`);
  }
  await ctx.close();
}
for (let run = 0; run < 2; run++) {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });
  const page = await newPage(ctx);
  const t = Date.now();
  await open(page, "word-preview", file);
  await page.waitForFunction(() => window.__wordPages > 0, null, { timeout: 300000 });
  const ready = Date.now() - t;
  await page.waitForFunction(() => [...document.querySelectorAll(".rb-file-viewer-canvas")].some((c) => c.width >= 800), null, { timeout: 60000 });
  console.log(`preview open (run ${run + 1}): pages known ${ready} ms, first page painted ${Date.now() - t} ms`, JSON.stringify(await ev(page, () => window.__wordLoaded.viewer.getRenderStats?.())));
  await ctx.close();
}
await browser.close();
