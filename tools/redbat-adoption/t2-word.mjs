// Task 2 Word probes (a, c, d, f, g + API cost). usage: node t2-word.mjs [keys]
import { launch, newPage, open, T, S, check, summary } from "./lib.mjs";
const only = process.argv[2]?.split(",");
const want = (k) => !only || only.includes(k);
const { browser, context } = await launch({ viewport: { width: 1366, height: 768 } });
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const E = `${S}/explore`;
const ready = async (page, complete = true) => {
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  if (complete) await page.waitForFunction(() => window.__word.isLoadComplete(window.__editors[0]) !== false && window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 });
  await page.waitForTimeout(600);
};
const focus = (page) => ev(page, () => window.__word.getInputElement(window.__editors[0]).focus({ preventScroll: true }));
// The text at a point of the view: the block and offset there, and the line's text.
const textAt = (page, y) => ev(page, (y) => {
  const w = window.__word, e = window.__editors[0]; const c = document.querySelector('[data-testid="word-canvas"]');
  const out = [];
  for (const yy of [y - 250, y - 200, y - 150, y - 100, y - 60]) { const p = w.hitTestViewport(e, 500, yy); out.push(p ? `${p.BlockId}:${p.Offset}` : null); }
  return { hits: out, engineTop: w.getViewport(e).scrollTop, hostTop: document.querySelector('[data-testid="word-scroll"]').scrollTop, complete: w.isLayoutComplete(e), stale: w.getRenderStats(e).StaleBlocks };
}, y);

if (want("a")) {
  const page = await newPage(context);
  const notices = [];
  await open(page, "word-editor", `${S}/long-generated.docx`);
  await ready(page);
  await ev(page, () => { window.__notices = []; window.__word.onChange(window.__editors[0], (n) => { if (typeof n.ScrollTop === "number") window.__notices.push({ ScrollTop: n.ScrollTop, LayoutChanged: n.LayoutChanged, host: document.querySelector('[data-testid="word-scroll"]').scrollTop }); }); });
  // scroll to the middle
  await ev(page, () => { const s = document.querySelector('[data-testid="word-scroll"]'); s.scrollTop = Math.round((s.scrollHeight - s.clientHeight) * 0.5); });
  await frame(page); await page.waitForTimeout(300);
  await page.mouse.click(500, 400); await frame(page);
  const before = await textAt(page, 300);
  await page.screenshot({ path: `${E}/t2-a-1-before.png` });
  await page.keyboard.press("Control+a"); await frame(page);
  const t0 = Date.now();
  await page.keyboard.press("Control+b");
  await frame(page);
  const right = await textAt(page, 300);
  await page.screenshot({ path: `${E}/t2-a-2-after-bold.png` });
  await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 });
  const done = Date.now() - t0;
  await frame(page); await page.waitForTimeout(300);
  const after = await textAt(page, 300);
  await page.screenshot({ path: `${E}/t2-a-3-complete.png` });
  const got = await ev(page, () => window.__notices);
  console.log("  a: before", JSON.stringify(before), "\n     right after", JSON.stringify(right), "\n     complete", JSON.stringify(after), `(${done} ms)`, "\n     notices", JSON.stringify(got));
  check("a: a ScrollTop notice arrived", got.length > 0, got.length);
  check("a: right after the bold the view shows the same text", JSON.stringify(right.hits) === JSON.stringify(before.hits), [before.hits, right.hits]);
  check("a: when the layout is complete the view shows the same text", JSON.stringify(after.hits) === JSON.stringify(before.hits), [before.hits, after.hits]);
  check("a: host scroller and engine viewport agree", Math.abs(after.engineTop - after.hostTop) < 1.5, [after.engineTop, after.hostTop]);
  // a wheel scroll afterwards does not jump
  await page.mouse.move(500, 400); await page.mouse.wheel(0, 40); await frame(page); await page.waitForTimeout(200);
  const wheel = await textAt(page, 260);
  check("a: after a 40 px wheel scroll the same text is 40 px higher", JSON.stringify(wheel.hits) === JSON.stringify(before.hits), [before.hits, wheel.hits]);
  // undo
  await focus(page); await page.keyboard.press("Control+z"); await frame(page);
  await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 });
  await frame(page); await page.waitForTimeout(300);
  const undone = await textAt(page, 260);
  check("a: after undo the view shows the same text", JSON.stringify(undone.hits) === JSON.stringify(before.hits), [before.hits, undone.hits]);
  check("a: after undo host scroller and engine viewport agree", Math.abs(undone.engineTop - undone.hostTop) < 1.5, [undone.engineTop, undone.hostTop]);
  await page.close();
}

if (want("c")) {
  // odd / even sections add blank pages to getPageFrames; the host's page count and scroll extent follow
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/test-documents/sample-formatted.docx`);
  await ready(page);
  const r = await ev(page, async () => {
    const w = window.__word, m = window.__wordModule, e = window.__editors[0];
    const pages = () => w.getDocumentSize(e).PageCount;
    const out = { before: pages() };
    const lines = w.getLayoutLines(e); const first = lines.find((l) => l.Page === 1 && l.Text.length > 5);
    m.setSelection(e, first.BlockId, 0, first.BlockId, 0);
    for (const kind of ["nextPage", "oddPage", "evenPage", "continuous", "none"]) {
      w.setSectionBreak(e, kind); w.completeLayout(e);
      const fr = Array.from(w.getPageFrames(e));
      const ls = w.getLayoutLines(e); const firstOfNext = ls.find((l) => l.BlockId !== first.BlockId && l.Y > ls.find((x) => x.BlockId === first.BlockId).Y);
      out[kind] = { pages: pages(), frames: fr.length, nextStartsOnPage: firstOfNext?.Page, blank: fr.map((f, i) => (ls.some((l) => l.Page === i + 1) ? "" : i + 1)).filter((x) => x !== "").join(",") };
    }
    out.keys = Object.keys(Array.from(w.getPageFrames(e))[0]);
    return out;
  });
  console.log("  c:", JSON.stringify(r));
  check("c: nextPage adds a page, the text after starts on page 2", r.nextPage.nextStartsOnPage === 2, r.nextPage);
  check("c: oddPage: a blank page 2, the text after starts on page 3", r.oddPage.nextStartsOnPage === 3 && r.oddPage.blank === "2", r.oddPage);
  check("c: evenPage: the text after starts on page 2, no blank page", r.evenPage.nextStartsOnPage === 2 && r.evenPage.blank === "", r.evenPage);
  check("c: continuous: the text after stays on page 1", r.continuous.nextStartsOnPage === 1, r.continuous);
  check("c: none: as before", r.none.pages === r.before, [r.before, r.none]);
  // host subtitle follows
  await ev(page, () => { const w = window.__word, e = window.__editors[0]; w.setSectionBreak(e, "oddPage"); });
  await frame(page); await page.waitForTimeout(400);
  const sub = await page.locator('[data-testid="office-editor-status"]').innerText().catch(() => "");
  const title = await ev(page, () => document.querySelector(".rb-office-editor-subtitle, [data-testid='office-editor-subtitle']")?.textContent ?? document.body.innerText.match(/\d+\+? pages?/)?.[0]);
  const n = await ev(page, () => window.__word.getDocumentSize(window.__editors[0]).PageCount);
  check("c: the host's page count says the pages with the blank one", String(title).includes(String(n)), [title, n, sub]);
  await ev(page, () => { const s = document.querySelector('[data-testid="word-scroll"]'); s.scrollTop = 900; });
  await sel(page, "word-zoom-out").click(); await sel(page, "word-zoom-out").click(); await frame(page); await page.waitForTimeout(300);
  await page.screenshot({ path: `${E}/t2-c-oddpage.png` });
  await page.close();
}

if (want("d")) {
  // header text boxes: listed, entered by API, by click, by double click
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/test-documents-features/resume-template.docx`);
  await ready(page);
  const shapes = await ev(page, () => { const w = window.__word, e = window.__editors[0]; return Array.from(w.getShapeRects(e)).map((r) => ({ Id: r.Id, Story: r.Story, PageIndex: r.PageIndex, X: r.X, Y: r.Y, W: r.Width, H: r.Height, Text: r.Text.slice(0, 30) })); });
  console.log("  d: shapes", JSON.stringify(shapes));
  const hdr = shapes.find((s) => s.Story !== "body");
  check("d: resume-template has a header/footer text box with Story and PageIndex", !!hdr && typeof hdr.PageIndex === "number", hdr);
  if (hdr) {
    const pt = await ev(page, (r) => { const w = window.__word, e = window.__editors[0]; const fr = Array.from(w.getPageFrames(e))[r.PageIndex]; const v = w.documentToViewport(e, (fr.X ?? 0) + r.X + r.W / 2, r.Y + r.H / 2); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: c.left + v.X, y: c.top + v.Y, frameX: fr.X }; }, hdr);
    await page.screenshot({ path: `${E}/t2-d-0.png` });
    await page.mouse.click(pt.x, pt.y); await frame(page); await page.waitForTimeout(200);
    const c1 = await ev(page, () => { const c = window.__word.getCaretContext(window.__editors[0]); return { Story: c.Story, TextBoxId: c.TextBoxId, hf: !!c.HeaderFooter }; });
    console.log("  d: after one click", JSON.stringify(c1), JSON.stringify(pt));
    await page.mouse.dblclick(pt.x, pt.y); await frame(page); await page.waitForTimeout(300);
    const c2 = await ev(page, () => { const c = window.__word.getCaretContext(window.__editors[0]); return { Story: c.Story, TextBoxId: c.TextBoxId, hf: !!c.HeaderFooter }; });
    console.log("  d: after a double click", JSON.stringify(c2));
    await page.screenshot({ path: `${E}/t2-d-1-dblclick.png` });
    check("d: a double click on a header text box enters the box", c2.Story === "textbox" && c2.TextBoxId === hdr.Id, c2);
    await page.keyboard.type("ZZ"); await frame(page); await page.waitForTimeout(200);
    const t = await ev(page, (id) => Array.from(window.__word.getShapeRects(window.__editors[0])).find((r) => r.Id === id)?.Text.slice(0, 40), hdr.Id);
    check("d: typing after it goes into the box", /ZZ/.test(t ?? ""), t);
    await page.screenshot({ path: `${E}/t2-d-2-typed.png` });
    const bar = await page.locator('[data-testid="word-text-box-bar"]').count();
    check("d: the text box bar is shown", bar === 1, bar);
    await page.keyboard.press("Escape"); await frame(page);
    const c3 = await ev(page, () => window.__word.getCaretContext(window.__editors[0]).Story);
    console.log("  d: after Escape", c3);
  }
  await page.close();
}

if (want("f")) {
  // caret into a [Comment n] mark snaps to its end
  const page = await newPage(context);
  const file = `${T}/test-documents-business/docx4j-comments-one.docx`;
  console.log("  f: file", file);
  await open(page, "word-editor", file);
  await ready(page);
  const r = await ev(page, () => {
    const w = window.__word, m = window.__wordModule, e = window.__editors[0];
    const line = w.getLayoutLines(e).find((l) => /\[Comment \d+\]/.test(l.Text));
    if (!line) return null;
    const full = w.getLayoutLines(e).filter((l) => l.BlockId === line.BlockId).map((l) => l.Text).join("");
    const at = full.indexOf("[Comment"); const len = full.slice(at).indexOf("]") + 1;
    m.setSelection(e, line.BlockId, at + 3, line.BlockId, at + 3);
    const s1 = m.getSelection(e);
    m.setSelection(e, line.BlockId, at + len - 2, line.BlockId, at + len - 2);
    const s2 = m.getSelection(e);
    return { block: line.BlockId, at, len, inner3: s1.FocusOffset, innerEnd: s2.FocusOffset, anchors: w.getCommentAnchors(e).length, full: full.slice(Math.max(0, at - 10), at + len + 10) };
  });
  console.log("  f:", JSON.stringify(r));
  if (r) {
    check("f: a position 3 into the mark is its start", r.inner3 === r.at, r);
    check("f: a position near its end is its end", r.innerEnd === r.at + r.len, r);
    // click into the middle of the mark
    const pt = await ev(page, (r) => { const w = window.__word, e = window.__editors[0]; const b = w.getRangeBounds(e, r.block, r.at + 2, r.at + r.len - 2); const one = Array.isArray(b) ? b[0] : b; const fr = Array.from(w.getPageFrames(e)); const v = w.documentToViewport(e, one.X + one.Width / 2, one.Y + (one.Height ?? 10) / 2); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: c.left + v.X, y: c.top + v.Y, b: JSON.stringify(one).slice(0, 300) }; }, r).catch((e) => ({ err: String(e) }));
    console.log("  f: bounds", JSON.stringify(pt));
    if (pt.x) {
      await ev(page, (r) => { const w = window.__word, m = window.__wordModule, e = window.__editors[0]; m.setSelection(e, r.block, 0, r.block, 0); }, r);
      await page.mouse.click(pt.x, pt.y); await frame(page);
      const s = await ev(page, () => window.__wordModule.getSelection(window.__editors[0]));
      check("f: a click in the mark puts the caret at one of its ends", s.FocusBlockId === r.block && (s.FocusOffset === r.at || s.FocusOffset === r.at + r.len), s);
      await page.keyboard.type("Q"); await frame(page);
      const after = await ev(page, (r) => window.__word.getLayoutLines(window.__editors[0]).filter((l) => l.BlockId === r.block).map((l) => l.Text).join(""), r);
      check("f: text typed there stands beside the whole mark", /\[Comment \d+\]Q|Q\[Comment \d+\]/.test(after), after.slice(Math.max(0, r.at - 5), r.at + r.len + 8));
      await page.keyboard.press("Backspace"); await page.keyboard.press("ArrowRight"); await page.keyboard.press("ArrowLeft"); await frame(page);
      const s3 = await ev(page, () => window.__wordModule.getSelection(window.__editors[0]).FocusOffset);
      check("f: arrow keys step over the mark", s3 === r.at || s3 === r.at + r.len, [s3, r.at, r.len]);
      await page.screenshot({ path: `${E}/t2-f-comment.png` });
    }
  }
  await page.close();
}

if (want("g")) {
  // what the host reads: selection format / context with the caret in a multi-run field; hyperlink bubble on a link made of several runs
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/test-documents/sample-formatted.docx`);
  await ready(page);
  const r = await ev(page, () => {
    const w = window.__word, m = window.__wordModule, e = window.__editors[0];
    const exp = w.exportDocumentPackage(e).Paragraphs; let fields = 0, multi = 0; const codes = [];
    const walk = (o, d = 0) => { if (!o || typeof o !== "object" || d > 8) return; if (Array.isArray(o)) { let prev = null; for (const x of o) { if (x && typeof x === "object" && x.FieldCode) { fields++; if (prev === x.FieldCode) multi++; prev = x.FieldCode; if (codes.length < 5) codes.push(String(x.FieldCode).slice(0, 40)); } else prev = null; walk(x, d + 1); } return; } for (const k of Object.keys(o)) walk(o[k], d + 1); };
    walk(exp);
    return { fields, multi, codes };
  });
  console.log("  g:", JSON.stringify(r));
  await page.close();
}

if (want("cost")) {
  const page = await newPage(context);
  await open(page, "word-editor", `${S}/long-generated.docx`);
  await ready(page);
  const r = await ev(page, () => {
    const w = window.__word, e = window.__editors[0];
    const time = (f, n = 20) => { const t = performance.now(); for (let i = 0; i < n; i++) f(); return Math.round(((performance.now() - t) / n) * 1000) / 1000; };
    const out = { shapeRects0: time(() => w.getShapeRects(e)), pageFrames: time(() => w.getPageFrames(e)) };
    w.insertTextBox(e, "Box");
    const id = Array.from(w.getShapeRects(e))[0]?.Id;
    out.shapeRects1 = time(() => w.getShapeRects(e)); out.handles = time(() => w.getShapeHandles(e, id)); out.hit = time(() => w.hitTestShapeHandle(e, id, 100, 100)); out.vAlign = w.getSectionVerticalAlignment(e);
    out.pages = w.getDocumentSize(e).PageCount;
    return out;
  });
  console.log("  cost (ms per call, long document):", JSON.stringify(r));
  await page.close();
}
await browser.close();
summary();
