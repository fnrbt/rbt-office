// Exploratory pass, Word: open several documents, exercise the editor, take screenshots into explore/.
import { launch, newPage, open, T, S } from "./lib.mjs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const E = `${S}/explore`;
const docs = { float: `${T}/repros/word/float-image-wrap-square.docx`, behind: `${T}/repros/word/float-image-behind-doc.docx`, shapes: `${T}/repros/word-shapes/autoshapes-gallery.docx`, watermark: `${T}/repros/word-gaps/watermark.docx`, resume: `${T}/test-documents-features/resume-template.docx`, sample: `${T}/test-documents/sample-formatted.docx`, phpword: `${T}/test-documents-wild/phpword-sample-11-readword2007.docx`, filesample: `${T}/test-documents-online/file-sample.docx` };
const only = process.argv[2]?.split(",");
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
  const info = await ev(page, () => { const w = window.__word, e = window.__editors[0]; return { pages: w.getDocumentSize(e).PageCount, shapes: w.getShapeRects(e).map((r) => `${r.Text.slice(0, 16)}|${r.BehindText ? "behind" : "front"}|${r.Wrap}|${Math.round(r.X)},${Math.round(r.Y)} ${Math.round(r.Width)}x${Math.round(r.Height)}`).slice(0, 14) }; });
  console.log(key, JSON.stringify(info));
  await page.screenshot({ path: `${E}/w-${key}-1-open.png` });
  const scroll = (f) => ev(page, (f) => { const s = document.querySelector('[data-testid="word-scroll"]'); s.scrollTop = (s.scrollHeight - s.clientHeight) * f; }, f);
  if (key === "sample" || key === "phpword" || key === "filesample") {
    // typing, selection, copy / paste, undo / redo
    await ev(page, () => window.__word.getInputElement(window.__editors[0]).focus());
    await page.keyboard.press("Control+Home"); await page.keyboard.press("ArrowDown"); await page.keyboard.press("Home");
    await page.keyboard.type("Typed here. ");
    await page.keyboard.press("Shift+Home"); await page.keyboard.press("Control+c"); await page.keyboard.press("End"); await page.keyboard.press("Control+v"); await page.keyboard.press("Control+v");
    await frame(page);
    for (let i = 0; i < 5; i++) await page.keyboard.press("Shift+ArrowDown");
    await frame(page);
    await page.screenshot({ path: `${E}/w-${key}-2-typed-pasted-selected.png` });
    await page.keyboard.press("ArrowLeft");
    const rev = () => ev(page, () => window.__word.getRevision(window.__editors[0]));
    let n = 0; while ((await ev(page, () => window.__word.canUndo(window.__editors[0]))) && n++ < 60) await page.keyboard.press("Control+z");
    const clean = await ev(page, () => !window.__word.isModified(window.__editors[0]));
    for (let i = 0; i < n; i++) await page.keyboard.press("Control+y");
    console.log(`  ${key}: ${n} undo steps back to unmodified: ${clean}; after redo modified: ${await ev(page, () => window.__word.isModified(window.__editors[0]))}`);
    // find / replace
    await sel(page, "word-find").click(); await sel(page, "word-find-input").fill("the"); await sel(page, "word-find-next").click(); await sel(page, "word-find-next").click(); await frame(page);
    const found = await sel(page, "word-find-status").innerText();
    await page.screenshot({ path: `${E}/w-${key}-3-find.png` });
    await sel(page, "word-replace-input").fill("THE"); await sel(page, "word-replace-all").click(); await frame(page);
    console.log(`  ${key}: find "the": ${found}; replace all: ${await sel(page, "word-find-status").innerText()}`);
    await page.keyboard.press("Escape"); await page.waitForTimeout(300);
    // zoom, scroll, resize
    await sel(page, "word-zoom-in").click(); await sel(page, "word-zoom-in").click(); await scroll(0.35); await page.waitForTimeout(500);
    await page.screenshot({ path: `${E}/w-${key}-4-zoom-mid.png` });
    await sel(page, "word-zoom").click(); await scroll(1); await page.waitForTimeout(500);
    await page.screenshot({ path: `${E}/w-${key}-5-end.png` });
    await page.setViewportSize({ width: 820, height: 620 }); await scroll(0.5); await page.waitForTimeout(600);
    await page.screenshot({ path: `${E}/w-${key}-6-small-window.png` });
    await page.setViewportSize({ width: 1366, height: 768 });
  } else if (info.pages > 1) { await scroll(1); await page.waitForTimeout(400); await page.screenshot({ path: `${E}/w-${key}-2-end.png` }); }
  console.log(`  ${key}: errors ${JSON.stringify(errors)}`);
  await page.close();
}
await browser.close();
