// Render stats while scrolling, save protocol with a delayed / failed upload, caret context in the toolbar.
import { launch, newPage, open, T, S, SHOTS, check, summary, unzip, waitUploads, LONG_PAGES } from "./lib.mjs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const ready = async (page) => { await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 }); await page.waitForTimeout(400); };
const status = (page) => page.locator('[data-testid="office-editor-status"]').innerText();
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const focusDoc = (page) => ev(page, () => window.__word.getInputElement(window.__editors.at(-1)).focus());
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const stats = (page) => ev(page, () => window.__wordModule.getRenderStats(window.__editors.at(-1)));
const docText = async (bytes) => (await (await unzip(bytes)).get("word/document.xml").async("string"));

// ---------------- long document: no frame builds while scrolling ----------------
{
  const page = await newPage(context);
  let t = Date.now();
  await open(page, "word-editor", `${S}/long-generated.docx`);
  await ready(page);
  console.log("  long document editor load ms", Date.now() - t);
  await ev(page, () => window.__word.whenLayoutComplete(window.__editors[0]));
  await page.waitForTimeout(300);
  const s0 = await stats(page);
  check("the stack's painter and measurers are installed (plain functions, no runtime-internal import)", s0.PainterVersion >= 1 && s0.MeasurerVersion >= 1, s0);
  const scrolled = await ev(page, async () => {
    const sc = document.querySelector('[data-testid="word-scroll"]');
    const max = sc.scrollHeight - sc.clientHeight; const times = [];
    for (let i = 1; i <= 80; i++) {
      const t = performance.now();
      sc.scrollTop = Math.round((max * i) / 80);
      sc.dispatchEvent(new Event("scroll"));
      times.push(performance.now() - t);
      if (i % 10 === 0) await new Promise((r) => requestAnimationFrame(r));
    }
    // and small steps, as a wheel does
    for (let i = 0; i < 40; i++) { const t = performance.now(); sc.scrollTop -= 120; sc.dispatchEvent(new Event("scroll")); times.push(performance.now() - t); }
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    times.sort((a, b) => a - b);
    return { max, pages: window.__word.getPageCount(window.__editors[0]), median: times[times.length >> 1], p95: times[Math.floor(times.length * 0.95)], worst: times.at(-1) };
  });
  const s1 = await stats(page);
  check("scrolling the long document end to end builds no frame", s1.FrameBuilds === s0.FrameBuilds && s1.PainterVersion === s0.PainterVersion && s1.MeasurerVersion === s0.MeasurerVersion && scrolled.pages === LONG_PAGES, { before: s0, after: s1 });
  console.log(`  scroll repaint ms: median ${scrolled.median.toFixed(1)}, p95 ${scrolled.p95.toFixed(1)}, worst ${scrolled.worst.toFixed(1)} (120 scroll positions)`);
  check("a scroll repaint is a few ms, not a frame build", scrolled.p95 < 40, scrolled);
  // pixels: the page in view is painted with glyph outlines (not blank)
  const ink = await ev(page, () => { const c = document.querySelector('[data-testid="word-canvas"]'); const d = c.getContext("2d").getImageData(0, 0, c.width, c.height).data; let dark = 0; for (let i = 0; i < d.length; i += 16) if (d[i] < 90 && d[i + 1] < 90 && d[i + 2] < 90) dark++; return dark; });
  check("the scrolled-to page is painted (dark text pixels)", ink > 500, ink);
  await page.screenshot({ path: `${SHOTS}/word-long-scrolled.png` });
  // caret moves and a keystroke
  await ev(page, () => { const sc = document.querySelector('[data-testid="word-scroll"]'); sc.scrollTop = 0; sc.dispatchEvent(new Event("scroll")); });
  await focusDoc(page);
  await page.keyboard.press("Control+Home");
  const s2 = await stats(page);
  let t2 = Date.now();
  for (let i = 0; i < 10; i++) await page.keyboard.press("ArrowRight");
  for (let i = 0; i < 5; i++) await page.keyboard.press("ArrowDown");
  await frame(page);
  const s3 = await stats(page);
  check("caret moves build no frame, lay out and collect nothing", s3.FrameBuilds === s2.FrameBuilds && s3.LaidOutBlocks === s2.LaidOutBlocks && s3.CollectedBlocks === s2.CollectedBlocks, { s2, s3 });
  console.log(`  15 caret moves at page 1 of ${LONG_PAGES}: ${Date.now() - t2} ms wall`);
  t2 = Date.now();
  await page.keyboard.type("a");
  const s4 = await stats(page);
  check("a keystroke lays out and collects the paragraph it is in, not the document (9008 blocks)", s4.Blocks === 9008 && s4.LaidOutBlocks <= 3 && s4.CollectedBlocks <= 3 && s4.FrameBuilds === s3.FrameBuilds + 1, s4);
  await page.keyboard.type("bcde");
  await frame(page);
  console.log(`  5 keystrokes at page 1 of ${LONG_PAGES}: ${Date.now() - t2} ms wall`);
  await page.close();
}

// ---------------- preview of the long document: page paints are cheap ----------------
{
  const page = await newPage(context);
  await open(page, "word-preview", `${S}/long-generated.docx`);
  await page.waitForFunction(() => window.__wordPages > 0, null, { timeout: 300000 });
  console.log("  viewer stats once the preview shows its first pages:", JSON.stringify(await ev(page, () => window.__wordLoaded.viewer.getRenderStats())));
  const r = await ev(page, () => { const v = window.__wordLoaded.viewer; const c = document.createElement("canvas"); const times = []; for (const i of [0, 50, 100, 200, 300, 400, 500, 529, 1, 2, 3, 4, 5, 6, 7, 8]) { const t = performance.now(); v.renderPage(c, i, 1); times.push(performance.now() - t); } times.sort((a, b) => a - b); return { first: times.at(-1), median: times[times.length >> 1], w: c.width }; });
  console.log(`  preview renderPage on ${LONG_PAGES} pages: median ${r.median.toFixed(1)} ms, slowest (first, builds the frame) ${r.first.toFixed(1)} ms`);
  check("preview: painting a page of a long document reuses the frame", r.median < 30 && r.w > 700, r);
  await page.close();
}

// ---------------- save protocol, guard, failure ----------------
{
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/test-documents/sample-formatted.docx`);
  await ready(page);
  await focusDoc(page);
  await page.keyboard.press("Control+Home");
  await page.keyboard.type("FIRST ");
  let release; page.upload.hold = new Promise((r) => (release = r));
  await page.keyboard.press("Control+s");
  while ((page.uploadsStarted ?? 0) < 1) await new Promise((r) => setTimeout(r, 30));
  check("Saving... while the upload runs", (await status(page)).includes("Saving"), await status(page));
  await page.keyboard.type("SECOND ");
  release(); page.upload.hold = null;
  const one = await waitUploads(page, 1);
  await page.waitForTimeout(200);
  check("typing during the upload leaves the document dirty", (await status(page)).includes("Unsaved") && (await ev(page, () => window.__word.isModified(window.__editors[0]))) && !(await sel(page, "office-editor-save").isDisabled()), await status(page));
  const x1 = await docText(one);
  check("the first upload holds what was there when the save began", x1.includes("FIRST") && !x1.includes("SECOND"));
  // (What was typed after the save began is one undo step - keys typed one after the other are one step since
  // @rbt/office 0.1.26, and an export ends the run since 0.1.29.)
  await page.keyboard.press("Control+z");
  const afterUndo = { status: await status(page), modified: await ev(page, () => window.__word.isModified(window.__editors[0])) };
  console.log("  after undoing what was typed during the upload:", JSON.stringify(afterUndo));
  check("undoing what was typed during the upload gives the saved state", afterUndo.status.includes("Saved") && !afterUndo.modified, afterUndo);
  await page.keyboard.press("Control+y");
  await page.keyboard.press("Control+s");
  const two = await waitUploads(page, 2);
  await page.waitForTimeout(200);
  check("a second save uploads the newer content and is clean", (await docText(two)).includes("SECOND") && (await status(page)).includes("Saved") && !(await ev(page, () => window.__word.isModified(window.__editors[0]))), await status(page));
  // failed upload
  await page.keyboard.type("THIRD ");
  page.upload.fail = 503;
  await sel(page, "office-editor-save").click();
  await page.waitForSelector('[data-testid="office-editor-error"]', { timeout: 60000 });
  check("failed upload: error shown, still dirty", (await sel(page, "office-editor-error").innerText()).includes("Not saved") && (await ev(page, () => window.__word.isModified(window.__editors[0]))) && page.uploads.length === 2, await sel(page, "office-editor-error").innerText());
  page.upload.fail = 0;
  await sel(page, "office-editor-save").click();
  const three = await waitUploads(page, 3);
  await page.waitForTimeout(200);
  check("saving again after the failure works", (await docText(three)).includes("THIRD") && (await status(page)).includes("Saved"));
  // guard
  await focusDoc(page);
  await page.keyboard.type("x");
  await sel(page, "office-editor-close").click();
  check("closing with unsaved changes asks", (await sel(page, "office-editor-confirm").count()) === 1);
  await page.keyboard.type("yy");
  check("nothing typed while the question is up reaches the document", !(await ev(page, () => window.__word.getText(window.__editors[0]).includes("xyy"))));
  await sel(page, "office-editor-keep").click();
  check("Keep editing gives the keyboard back to the document", await ev(page, () => document.activeElement === window.__word.getInputElement(window.__editors[0])));
  await page.keyboard.press("Control+z");
  check("undo back to the saved content is clean", (await status(page)).includes("Saved"), await status(page));
  await sel(page, "office-editor-close").click();
  await page.waitForFunction(() => window.__closed === 1);
  check("then closing does not ask", true);
  await page.close();
}

// ---------------- caret context: list state, hyperlink bubble ----------------
{
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/test-documents/sample-formatted.docx`);
  await ready(page);
  const lines = await ev(page, () => { const w = window.__word, e = window.__editors[0]; const pkg = w.exportDocumentPackage(e); const ids = w.getLayoutLines(e); const find = (pred) => ids.find(pred); const text = (x) => x.Text ?? ""; return { body: find((x) => text(x).includes("This document demonstrates"))?.BlockId, bullets: pkg.Paragraphs.filter((p) => p.ListStyle === "bullet").length, numbered: pkg.Paragraphs.filter((p) => p.ListStyle === "numbered").length }; });
  // put the caret in a bulleted and in a numbered paragraph via the document's own lists
  const listBlocks = await ev(page, () => { const w = window.__word, e = window.__editors[0]; const out = {}; const all = w.getLayoutLines(e); const seen = new Set(); for (const l of all) { if (seen.has(l.BlockId)) continue; seen.add(l.BlockId); w.setSelection(e, l.BlockId, 0, l.BlockId, 0); const k = w.getSelectionFormat(e).Paragraph.ListKind; const c = w.getCaretContext(e); if (k === "bullet" && c.InList && !out.bullet) out.bullet = l.BlockId; if (k === "numbered" && c.InList && !out.numbered) out.numbered = l.BlockId; if (out.bullet && out.numbered) break; } return out; });
  check("fixture has a bulleted and a numbered paragraph", !!listBlocks.bullet && !!listBlocks.numbered, { lines, listBlocks });
  const at = async (id) => { await ev(page, (id) => window.__word.setSelection(window.__editors[0], id, 0, id, 0), id); await frame(page); };
  await at(listBlocks.bullet);
  check("in a bulleted paragraph: bullets pressed, indent buttons speak of list levels", (await sel(page, "word-bullets").getAttribute("aria-pressed")) === "true" && (await sel(page, "word-numbers").getAttribute("aria-pressed")) === "false" && (await sel(page, "word-indent").getAttribute("title")) === "Increase list level");
  await at(listBlocks.numbered);
  check("in a numbered paragraph: numbers pressed", (await sel(page, "word-numbers").getAttribute("aria-pressed")) === "true" && (await sel(page, "word-bullets").getAttribute("aria-pressed")) === "false");
  await at(lines.body);
  check("outside a list: neither pressed, plain indent", (await sel(page, "word-numbers").getAttribute("aria-pressed")) === "false" && (await sel(page, "word-bullets").getAttribute("aria-pressed")) === "false" && (await sel(page, "word-indent").getAttribute("title")) === "Increase indent");
  // hyperlink bubble
  await sel(page, "word-link").click();
  await sel(page, "word-link-target").fill("https://example.com/ctx");
  await sel(page, "word-link-text").fill("LINK");
  await sel(page, "word-link-apply").click();
  await ev(page, (id) => window.__word.setSelection(window.__editors[0], id, 2, id, 2), lines.body);
  await frame(page);
  check("caret in the link: context says so and the bubble shows", (await ev(page, () => window.__word.getCaretContext(window.__editors[0]).InHyperlink)) && (await sel(page, "word-link-bubble").count()) === 1 && (await sel(page, "word-link").getAttribute("title")) === "Edit link");
  await ev(page, (id) => window.__word.setSelection(window.__editors[0], id, 30, id, 30), lines.body);
  await frame(page);
  check("caret outside the link: no bubble", !(await ev(page, () => window.__word.getCaretContext(window.__editors[0]).InHyperlink)) && (await sel(page, "word-link-bubble").count()) === 0);
  await page.close();
}
summary();
await browser.close();
