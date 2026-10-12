// Exploratory pass on what 0.1.14 changed, Word. usage: PORT=.. node explore14-word.mjs v14|v13 [keys]
import { launch, newPage, open, T, S } from "./lib.mjs";
const tag = process.argv[2] ?? "v14";
const only = process.argv[3]?.split(",");
const { browser, context } = await launch({ viewport: { width: 1500, height: 1000 } });
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const E = `/tmp/redbat-adoption/explore`;
const docs = {
  phpword: `${T}/test-documents-wild/phpword-sample-11-readword2007.docx`, three: `${T}/word/roundtrip-saved/three-sections.original.docx`, landscape: `${T}/repros/word/landscape-section.docx`,
  headers: `${T}/test-documents-business/docx4j-headers.docx`, hfmulti: `${T}/test-documents-business/docx4j-hf-sections-multiple.docx`, twohdr: `${T}/repros/word/two-header-sections.docx`, balanced: `${T}/repros/word/balanced-columns.docx`,
  filesample: `${T}/test-documents-online/file-sample.docx`, wps: `${T}/test-documents-wild/docx4j-drawingml-wps.docx`, shapes: `${T}/repros/word-shapes/autoshapes-gallery.docx`, resume: `${T}/test-documents-features/resume-template.docx`, float: `${T}/repros/word/float-image-wrap-square.docx`, started: `${T}/test-documents-wild/docx4j-getting-started.docx`,
};
for (const [key, file] of Object.entries(docs)) {
  if (only && !only.includes(key)) continue;
  const page = await newPage(context);
  const errors = [];
  page.on("pageerror", (e) => errors.push(String(e).slice(0, 200)));
  page.on("console", (m) => { if (m.type() === "error" && !/404/.test(m.text())) errors.push(m.text().slice(0, 200)); });
  await open(page, "word-editor", file);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await page.waitForFunction(() => window.__word.isLoadComplete(window.__editors[0]) !== false && window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 });
  await page.waitForTimeout(900);
  const info = await ev(page, () => { const w = window.__word, e = window.__editors[0]; const fr = Array.from(w.getPageFrames(e)); return { pages: w.getDocumentSize(e).PageCount, frames: fr.slice(0, 12).map((p) => `${Math.round(p.Width)}x${Math.round(p.Height)}`).join(" "), shapes: w.getShapeRects(e).length }; });
  console.log(tag, key, JSON.stringify(info));
  const shot = async (name) => { await frame(page); await page.waitForTimeout(250); await page.screenshot({ path: `${E}/${tag}-w-${key}-${name}.png` }); };
  const zoomOut = async (n) => { for (let i = 0; i < n; i++) await sel(page, "word-zoom-out").click(); await frame(page); };
  const scrollTo = (f) => ev(page, (f) => { const s = document.querySelector('[data-testid="word-scroll"]'); s.scrollTop = (s.scrollHeight - s.clientHeight) * f; s.dispatchEvent(new Event("scroll")); }, f);
  const focus = () => ev(page, () => window.__word.getInputElement(window.__editors[0]).focus({ preventScroll: true }));
  const boxPoint = (r) => ev(page, (r) => { const w = window.__word, e = window.__editors[0]; const v = w.documentToViewport(e, r.X + r.Width / 2, r.Y + r.Height / 2); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: c.left + v.X, y: c.top + v.Y }; }, r);
  if (["phpword", "three", "landscape", "headers", "hfmulti", "twohdr", "balanced", "started"].includes(key)) {
    await zoomOut(2); await shot("1-overview-50");
    if (info.pages > 2) { await scrollTo(0.5); await shot("2-middle-50"); }
    if (key === "phpword") { await sel(page, "word-zoom-in").click(); await sel(page, "word-zoom-in").click(); await scrollTo(0); await shot("3-top-100"); await scrollTo(1); await shot("4-end-100"); }
  } else if (key === "filesample") {
    await shot("1-open");
    // the list: Tab on its second item, then typing at a caret with Bold
    const at = await ev(page, () => { const w = window.__word, m = window.__wordModule, e = window.__editors[0]; const f = w.findText(e, "Lists ", true); if (!f.length) return null; w.selectFindMatch(e, f[0]); const s = m.getSelection(e); m.setSelection(e, s.AnchorBlockId, 0, s.AnchorBlockId, 0); w.getInputElement(e).focus({ preventScroll: true }); return w.getSelectionFormat(e).Paragraph; });
    await page.keyboard.press("Tab"); await frame(page);
    const after = await ev(page, () => window.__word.getSelectionFormat(window.__editors[0]).Paragraph);
    console.log(`  Tab on "Lists": level ${at?.ListLevel} indent ${at?.IndentLeft} -> level ${after.ListLevel} indent ${after.IndentLeft}`);
    await page.keyboard.press("End"); await sel(page, "word-bold").click(); await page.keyboard.type("bold typed"); await sel(page, "word-bold").click(); await page.keyboard.type(" plain typed");
    await shot("2-list-tab-and-pending-bold");
    await page.keyboard.press("Enter"); await page.keyboard.type("new item"); await page.keyboard.press("Shift+Tab"); await shot("3-enter-shift-tab");
    await scrollTo(0.35); await shot("4-tables"); await scrollTo(1); await shot("5-end-columns");
  } else if (key === "wps" || key === "shapes" || key === "resume") {
    await shot("1-open");
    const list = await ev(page, () => window.__word.getShapeRects(window.__editors[0]).map((r) => ({ Id: r.Id, X: r.X, Y: r.Y, Width: r.Width, Height: r.Height, Text: r.Text, Behind: r.BehindText })));
    console.log("  boxes:", list.slice(0, 10).map((r) => `${JSON.stringify(r.Text.slice(0, 18))}@${Math.round(r.X)},${Math.round(r.Y)} ${Math.round(r.Width)}x${Math.round(r.Height)}`).join("  "));
    await focus(); await page.keyboard.press("Control+End"); await page.keyboard.type(" END-TYPED"); await frame(page);
    console.log("  after Ctrl+End + typing: story", await ev(page, () => window.__word.getCaretContext(window.__editors[0]).Story), "boxes changed:", JSON.stringify(list.map((r) => r.Text)) !== JSON.stringify(await ev(page, () => window.__word.getShapeRects(window.__editors[0]).map((r) => r.Text))));
    await page.keyboard.press("Control+a"); await sel(page, "word-italic").click(); await frame(page);
    await shot("2-ctrl-end-typed-select-all-italic");
    const box = list.find((r) => r.Text && !r.Behind && r.Width > 40);
    if (box) {
      await ev(page, (y) => { const s = document.querySelector('[data-testid="word-scroll"]'); s.scrollTop = Math.max(0, y - 200); s.dispatchEvent(new Event("scroll")); }, box.Y); await frame(page);
      const p = await boxPoint(box); await page.mouse.click(p.x, p.y); await frame(page);
      const c = await ev(page, () => { const w = window.__word, e = window.__editors[0]; return { story: w.getCaretContext(e).Story, caret: w.getCaretRect(e) }; });
      await page.keyboard.type("IN-BOX "); await page.keyboard.press("Control+a"); await frame(page);
      console.log(`  click in box ${JSON.stringify(box.Text.slice(0, 20))}: story ${c.story}, caret rotation ${c.caret?.Rotation}`);
      await shot("3-in-box-typed-selected");
      await page.keyboard.press("Escape"); await frame(page);
    }
  } else if (key === "float") { await shot("1-open"); await zoomOut(2); await shot("2-overview-50"); }
  if (errors.length) console.log("  ERRORS", key, [...new Set(errors)].slice(0, 4));
  await page.close();
}
await browser.close();
