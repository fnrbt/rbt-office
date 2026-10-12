import { launch, newPage, open, T, S, SHOTS } from "./lib.mjs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const st = (page) => ev(page, () => { const w = window.__word, e = window.__editors[0]; const f = w.getSelectionFormat(e); const c = w.getCaretContext(e); return { story: c.Story, tb: c.TextBoxId, bold: f.Bold, caret: w.getCaretRect(e) && [Math.round(w.getCaretRect(e).X), Math.round(w.getCaretRect(e).Y)], boxes: (w.getShapeRects ? w.getShapeRects(e) : []).map((r) => r.Text), text: w.getDocumentText ? w.getDocumentText(e).slice(0, 200) : null }; });
const page = await newPage(context);
await open(page, "word-editor", process.argv[2] ?? `${T}/test-documents-wild/docx4j-drawingml-wps.docx`);
await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
await page.waitForTimeout(800);
console.log("start", JSON.stringify(await st(page)));
await sel(page, "word-bold").click(); await frame(page);
console.log("pending at start", JSON.stringify(await st(page)));
await page.keyboard.type("xx"); await frame(page);
console.log("typed at start", JSON.stringify(await st(page)));
await page.keyboard.press("Control+End"); await frame(page);
console.log("ctrl+end", JSON.stringify(await st(page)));
await page.keyboard.type("yy"); await frame(page);
console.log("typed", JSON.stringify(await st(page)));
for (let i = 0; i < 6; i++) { await page.keyboard.press("ArrowLeft"); }
await frame(page); console.log("6 left", JSON.stringify(await st(page)));
await page.keyboard.press("Control+Home"); await frame(page);
for (let i = 0; i < 4; i++) { await page.keyboard.press("ArrowDown"); await frame(page); console.log("down", i, JSON.stringify((await st(page)).caret), (await st(page)).story); }
await page.screenshot({ path: `${SHOTS}/probe-tb-4.png` });
await browser.close();
