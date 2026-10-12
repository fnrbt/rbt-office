// Progressive layout: first screens at once, pages added without moving what is on screen, completion where it is needed.
import { launch, newPage, open, T, S, SHOTS, check, summary, waitUploads, unzip } from "./lib.mjs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const file = `${S}/long-generated.docx`;
const status = (page) => page.locator('[data-testid="office-editor-status"]').innerText();
{ // editor
  const page = await newPage(context);
  await open(page, "word-editor", file);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  const first = await ev(page, () => { const w = window.__word, e = window.__editors[0]; const sc = document.querySelector('[data-testid="word-scroll"]'); return { complete: w.isLayoutComplete(e), size: w.getDocumentSize(e), spacer: sc.scrollHeight }; });
  const sub = await status(page);
  check("editor is usable before the whole document is laid out", first.complete === false && first.size.LayoutComplete === false && first.size.PageCount > 0 && first.size.PageCount < 522, first);
  check("page indicator says there is more to come", /\d+\+ pages…/.test(sub), sub);
  await page.screenshot({ path: `${SHOTS}/word-progressive-first.png` });
  // scroll a little and type while the rest is laid out
  await ev(page, () => { const sc = document.querySelector('[data-testid="word-scroll"]'); sc.scrollTop = 300; sc.dispatchEvent(new Event("scroll")); window.__word.getInputElement(window.__editors[0]).focus(); });
  const before = await ev(page, () => { const c = document.querySelector('[data-testid="word-canvas"]'); return c.toDataURL().length; });
  await ev(page, () => new Promise((resolve) => { const w = window.__word, e = window.__editors[0]; window.__tops = []; const sc = document.querySelector('[data-testid="word-scroll"]'); const stop = w.onChange(e, (c) => { if (c.LayoutChanged) window.__tops.push([sc.scrollTop, w.getDocumentSize(e).PageCount, c.DocumentChanged]); }); w.whenLayoutComplete(e).then(() => { stop(); resolve(); }); }));
  await page.waitForTimeout(300);
  const after = await ev(page, () => { const w = window.__word, e = window.__editors[0]; const sc = document.querySelector('[data-testid="word-scroll"]'); return { complete: w.isLayoutComplete(e), pages: w.getDocumentSize(e).PageCount, scrollTop: sc.scrollTop, spacer: sc.scrollHeight, notices: window.__tops.length, tops: [...new Set(window.__tops.map((t) => t[0]))], counts: window.__tops.map((t) => t[1]), docChanged: window.__tops.some((t) => t[2]) }; });
  check("the layout completes in the background: 522 pages, progress notices with growing page counts, none a document change", after.complete && after.pages === 522 && after.notices >= 2 && after.counts.every((n, i) => i === 0 || n >= after.counts[i - 1]) && !after.docChanged, { notices: after.notices, counts: after.counts.slice(0, 8) });
  check("the scroll area grew and the view did not move", after.spacer > first.spacer * 3 && after.scrollTop === 300 && after.tops.join() === "300", { first: first.spacer, after: after.spacer, tops: after.tops });
  check("status clean and complete", /No changes\s+522 pages$/.test((await status(page)).trim()), await status(page));
  await page.close();
}
{ // editing, find and save before completion
  const page = await newPage(context);
  await open(page, "word-editor", file);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await ev(page, () => window.__word.getInputElement(window.__editors[0]).focus());
  await page.keyboard.type("EARLY ");
  const r = await ev(page, () => { const w = window.__word, e = window.__editors[0]; return { modified: w.isModified(e), text: w.getText(e).slice(0, 30) }; });
  check("typing right after the first paint goes into the document", r.modified && r.text.includes("EARLY "), r);
  await page.keyboard.press("Control+s");
  const bytes = await waitUploads(page, 1);
  await page.waitForTimeout(200);
  const xml = await (await unzip(bytes)).get("word/document.xml").async("string");
  const orig = await (await unzip((await import("node:fs")).readFileSync(file))).get("word/document.xml").async("string");
  const paras = (x) => (x.match(/<w:p[ >]/g) || []).length;
  check("a save right away writes the whole document (every paragraph) with the edit", xml.includes("EARLY") && paras(xml) === paras(orig) && (await ev(page, () => window.__word.isLayoutComplete(window.__editors[0]))), { saved: paras(xml), orig: paras(orig) });
  check("saved and clean", (await status(page)).includes("Saved"), await status(page));
  await page.close();
  const page2 = await newPage(context);
  await open(page2, "word-editor", file);
  await page2.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  const lastWords = await ev(page2, () => { const pkg = window.__word.exportDocumentPackage(window.__editors[0]); const ps = pkg.Paragraphs.filter((p) => p.Runs?.some((r) => (r.Text ?? "").trim().length > 12)); const last = ps.at(-1).Runs.map((r) => r.Text ?? "").join(""); return last.trim().split(/\s+/).slice(-3).join(" "); });
  await page2.click('[data-testid="word-find"]');
  await page2.fill('[data-testid="word-find-input"]', lastWords);
  await page2.click('[data-testid="word-find-next"]');
  const found = await ev(page2, () => { const w = window.__word, e = window.__editors[0]; return { status: document.querySelector('[data-testid="word-find-status"]').textContent, complete: w.isLayoutComplete(e), caretY: w.getCaretRect(e)?.Y, height: w.getDocumentSize(e).Height }; });
  check("find for text at the end of the document, asked before the layout is complete, finds it, beyond the pages that were laid out", /of \d+/.test(found.status) && found.complete && found.caretY > 21000, { lastWords, found });
  await page2.close();
}
{ // preview
  const page = await newPage(context);
  await open(page, "chat-viewer", file, { params: { mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } });
  await page.waitForFunction(() => [...document.querySelectorAll(".rb-file-viewer-canvas")].some((c) => c.width >= 800), null, { timeout: 300000, polling: 20 });
  const first = await ev(page, () => ({ sub: document.querySelector('[data-testid="word-viewer-pages"]').textContent.trim(), canvases: document.querySelectorAll(".rb-file-viewer-canvas").length, firstCanvas: document.querySelector(".rb-file-viewer-canvas") }).sub + "|" + document.querySelectorAll(".rb-file-viewer-canvas").length);
  await ev(page, () => { window.__c0 = document.querySelector(".rb-file-viewer-canvas"); const sc = document.querySelector(".rb-file-viewer-stage-scroll"); sc.scrollTop = 400; });
  await page.screenshot({ path: `${SHOTS}/word-preview-progressive.png` });
  await page.waitForFunction(() => document.querySelector('[data-testid="word-viewer-pages"]').textContent.trim() === "522 pages", null, { timeout: 300000 });
  const after = await ev(page, () => { const sc = document.querySelector(".rb-file-viewer-stage-scroll"); const cs = [...document.querySelectorAll(".rb-file-viewer-canvas")]; return { canvases: cs.length, sameFirstCanvas: cs[0] === window.__c0, firstStillPainted: cs[0].width >= 800, scrollTop: sc.scrollTop, painted: cs.filter((c) => c.width >= 800).length }; });
  check("preview shows its first pages before the layout is complete", /^\d+\+ pages…$/.test(first.split("|")[0]) && Number(first.split("|")[1]) < 522, first);
  check("pages are appended: 522 canvases, the first ones kept (no repaint), the view where it was, only pages near the view painted", after.canvases === 522 && after.sameFirstCanvas && after.firstStillPainted && after.scrollTop === 400 && after.painted < 8, after);
  // a page far down paints when scrolled to
  await ev(page, () => { const sc = document.querySelector(".rb-file-viewer-stage-scroll"); sc.scrollTop = sc.scrollHeight * 0.9; });
  await page.waitForTimeout(500);
  const far = await ev(page, () => { const cs = [...document.querySelectorAll(".rb-file-viewer-canvas")]; const i = cs.findIndex((c, k) => k > 400 && c.width >= 800); if (i < 0) return null; const d = cs[i].getContext("2d").getImageData(0, 0, cs[i].width, cs[i].height).data; let n = 0; for (let k = 0; k < d.length; k += 16) if (d[k] < 90) n++; return { i, ink: n }; });
  check("a page near the end paints when scrolled to", far && far.ink > 300, far);
  await page.close();
}
summary();
await browser.close();
