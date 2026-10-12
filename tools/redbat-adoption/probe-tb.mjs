import { launch, newPage, open, T, S, SHOTS } from "./lib.mjs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const ctx = (page) => ev(page, () => window.__word.getCaretContext(window.__editors[0]));
for (const f of [process.argv[2] ?? `${T}/test-documents-wild/docx4j-drawingml-wps.docx`]) {
  const page = await newPage(context);
  await open(page, "word-editor", f);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await page.waitForTimeout(800);
  const rects = await ev(page, () => window.__word.getShapeRects(window.__editors[0]));
  console.log(JSON.stringify(rects));
  await page.screenshot({ path: `${SHOTS}/probe-tb-0.png` });
  const front = rects.find((r) => !r.BehindText && r.Text);
  if (!front) { console.log("no front box"); continue; }
  const pt = await ev(page, (r) => { const w = window.__word, e = window.__editors[0]; const v = w.documentToViewport(e, r.X + r.Width / 2, r.Y + r.Height / 2); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: c.left + v.X, y: c.top + v.Y }; }, front);
  await page.mouse.click(pt.x, pt.y);
  await frame(page);
  console.log("ctx after click", JSON.stringify(await ctx(page)));
  await page.keyboard.type("ZZ");
  await sel(page, "word-bold").click().catch((e) => console.log("no bold btn", String(e).slice(0, 100)));
  await page.keyboard.type("bb");
  await frame(page);
  console.log("fmt", JSON.stringify(await ev(page, () => { const f = window.__word.getSelectionFormat(window.__editors[0]); return { Bold: f.Bold, size: f.FontSizePt, fam: f.FontFamily, al: f.Paragraph.Alignment }; })));
  console.log("rect text", (await ev(page, () => window.__word.getShapeRects(window.__editors[0]).map((r) => r.Text))).join(" | "));
  console.log("caret", JSON.stringify(await ev(page, () => window.__word.getCaretRect(window.__editors[0]))));
  await page.screenshot({ path: `${SHOTS}/probe-tb-1.png` });
  // commands
  for (const [name, fn] of Object.entries({
    table: () => window.__word.insertTable(window.__editors[0], 2, 2),
    pagebreak: () => window.__word.insertPageBreak(window.__editors[0]),
    textbox: () => window.__word.insertTextBox(window.__editors[0], "inner"),
    header: () => window.__word.enterHeaderFooter(window.__editors[0], 0, "header"),
  })) {
    let r; try { r = await ev(page, fn); } catch (e) { r = "THROW " + String(e).slice(0, 200); }
    await frame(page);
    console.log(name, "->", JSON.stringify(r), "story", (await ctx(page)).Story, "boxes", (await ev(page, () => window.__word.getShapeRects(window.__editors[0]).map((r) => r.Id + ":" + r.Text.slice(0, 30)))).join(" | "));
    if (name !== "header") { await page.keyboard.press("Control+z"); await frame(page); console.log("  after undo story", (await ctx(page)).Story); }
  }
  await page.screenshot({ path: `${SHOTS}/probe-tb-2.png` });
  await page.keyboard.press("Escape"); await frame(page);
  console.log("after esc", (await ctx(page)).Story, "editor open", await sel(page, "office-editor").count(), "confirm", await sel(page, "office-editor-confirm").count());
  await page.keyboard.press("Escape"); await frame(page);
  console.log("after esc2", (await ctx(page)).Story, "editor open", await sel(page, "office-editor").count(), "confirm", await sel(page, "office-editor-confirm").count());
  await page.close();
}
await browser.close();
