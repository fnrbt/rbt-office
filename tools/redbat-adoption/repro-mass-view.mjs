// Mid-document select all + bold on a long document: what the engine says its size is at every notice,
// and where the host's scroller ends up. usage: node repro-mass-view.mjs
import { launch, newPage, open, T, S } from "./lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "word-editor", `${S}/long-generated.docx`);
await page.waitForSelector('[data-testid="word-toolbar"]');
await page.waitForFunction(() => window.__word.isLoadComplete(window.__editors[0]) !== false && window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 });
await page.waitForTimeout(600);
const frame = () => page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
await page.evaluate((keep) => {
  window.__keep = keep; window.__log = []; const s = document.querySelector('[data-testid="word-scroll"]'); const w = window.__word, e = window.__editors[0];
  const snap = (what, n) => { const d = w.getDocumentSize(e); window.__log.push(`${what} t=${Math.round(performance.now())} ${n ? `doc=${n.DocumentChanged} sel=${n.SelectionChanged} layout=${n.LayoutChanged} complete=${n.LayoutComplete} ScrollTop=${n.ScrollTop}` : ""} size.Height=${Math.round(d.Height)} pages=${d.PageCount} stale=${w.getRenderStats(e).StaleBlocks} engineTop=${Math.round(w.getViewport(e).scrollTop)} hostTop=${Math.round(s.scrollTop)} spacer=${document.querySelector(".rb-word-editor-spacer").style.height}`); };
  window.__snap = snap;
  s.addEventListener("scroll", () => snap("SCROLL"));
  w.onChange(e, (n) => snap("notice", n));
  s.scrollTop = Math.round((s.scrollHeight - s.clientHeight) * 0.5);
  if (window.__keep) { const st = document.createElement("style"); st.textContent = `.rb-word-editor-spacer{min-height:${s.scrollHeight}px}`; document.head.append(st); w.onChange(e, (n) => { if (typeof n.ScrollTop === "number") s.scrollTop = n.ScrollTop; }); }
}, process.argv[2] === "keep");
await frame(); await page.waitForTimeout(300);
await page.evaluate(() => window.__word.getInputElement(window.__editors[0]).focus());
const top = () => page.evaluate(() => { for (let y = 10; y < 600; y += 10) { const p = window.__word.hitTestViewport(window.__editors[0], 500, y); if (p) return `${p.BlockId}:${p.Offset} at y=${y}`; } return null; });
console.log("text at y=100 before:", await top());
await page.keyboard.press("Control+a"); await frame();
await page.evaluate(() => window.__snap("BOLD-KEY"));
await page.keyboard.press("Control+b");
await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 }); await frame(); await page.waitForTimeout(300);
console.log((await page.evaluate(() => window.__log)).join("\n"));
console.log("text at y=100 after:", await top());
await browser.close();
