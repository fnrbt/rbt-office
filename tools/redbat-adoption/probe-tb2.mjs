import { launch, newPage, open, T, S, SHOTS } from "./lib.mjs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const ctx = (page) => ev(page, () => window.__word.getCaretContext(window.__editors[0]));
const fmt = (page) => ev(page, () => { const f = window.__word.getSelectionFormat(window.__editors[0]); return { Bold: f.Bold, I: f.Italic, size: f.FontSizePt, fam: f.FontFamily, al: f.Paragraph.Alignment, col: f.IsCollapsed, color: f.Color }; });
const page = await newPage(context);
await open(page, "word-editor", `${T}/test-documents-wild/docx4j-drawingml-wps.docx`);
await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
await page.waitForTimeout(800);
// body: pending bold
await page.keyboard.press("Control+End");
await sel(page, "word-bold").click(); await frame(page);
console.log("body pending", JSON.stringify(await fmt(page)));
await page.keyboard.type("bb"); await frame(page);
console.log("body after typing", JSON.stringify(await fmt(page)));
const id = await ev(page, () => window.__word.getShapeRects(window.__editors[0])[0].Id);
console.log("enter", await ev(page, (id) => window.__word.enterTextBox(window.__editors[0], id), id));
await frame(page);
console.log("tb ctx", (await ctx(page)).Story, JSON.stringify(await fmt(page)));
await sel(page, "word-bold").click(); await frame(page);
console.log("tb pending", JSON.stringify(await fmt(page)), "focus in doc", await ev(page, () => document.activeElement === window.__word.getInputElement(window.__editors[0])));
await page.keyboard.type("bb"); await frame(page);
console.log("tb after typing", JSON.stringify(await fmt(page)), await ev(page, () => window.__word.getShapeRects(window.__editors[0])[0].Text));
// selection formatting
await page.keyboard.press("Control+a"); await frame(page);
console.log("tb select all", JSON.stringify(await fmt(page)), (await ctx(page)).Story);
await sel(page, "word-bold").click(); await frame(page);
console.log("tb sel bold", JSON.stringify(await fmt(page)));
await sel(page, "word-italic").click(); await sel(page, "word-align-right").click(); await frame(page);
console.log("tb sel italic+right", JSON.stringify(await fmt(page)));
await page.locator('[data-testid="word-font-size"]').selectOption("14").catch((e) => console.log("size select:", String(e).slice(0, 150)));
await frame(page);
console.log("tb size", JSON.stringify(await fmt(page)), (await ctx(page)).Story);
await page.screenshot({ path: `${SHOTS}/probe-tb-3.png` });
for (let i = 0; i < 3; i++) { await page.keyboard.press("Control+z"); await frame(page); console.log("undo", i, JSON.stringify(await fmt(page)), (await ctx(page)).Story); }
// find from inside the box
await page.keyboard.press("Control+f").catch(() => {});
await browser.close();
