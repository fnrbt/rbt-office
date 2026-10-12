import { launch, newPage, open, T, S } from "./lib.mjs";
const { browser, context } = await launch();
const ready = async (page) => { await page.waitForSelector('[data-testid="word-toolbar"]'); await page.waitForFunction(() => window.__word.isLoadComplete(window.__editors[0]) !== false && window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 }); await page.waitForTimeout(500); };
{
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/test-documents-features/resume-template.docx`); await ready(page);
  console.log("header box:", JSON.stringify(await page.evaluate(() => { const w = window.__word, e = window.__editors[0]; const r = Array.from(w.getShapeRects(e))[0]; const out = {}; out.enterHF = w.enterHeaderFooter(e, 0, "header"); out.ctx1 = w.getCaretContext(e).Story; out.enterBox = w.enterTextBox(e, r.Id); const c = w.getCaretContext(e); out.ctx2 = [c.Story, c.TextBoxId, !!c.HeaderFooter]; out.rectsWhile = Array.from(w.getShapeRects(e)).map((x) => [x.Id, x.Story, x.PageIndex]); out.exit = w.exitTextBox(e); out.ctx3 = w.getCaretContext(e).Story; return out; })));
  await page.close();
}
{
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/repros/word/landscape-section.docx`); await ready(page);
  const r = await page.evaluate(() => { const w = window.__word, m = window.__wordModule, e = window.__editors[0]; const fr = Array.from(w.getPageFrames(e)); const out = { frames: fr.map((f) => [f.X, f.Top, f.Width, f.Height]) };
    const lines = w.getLayoutLines(e); const narrow = fr.findIndex((f) => f.X > 0); const line = lines.find((l) => l.Page === narrow + 1 && l.Text.trim().length > 3) ?? lines[0];
    out.line = [line.Page, line.X, line.Y, line.Text.slice(0, 20)];
    m.setSelection(e, line.BlockId, 0, line.BlockId, 0); out.caret = w.getCaretRect(e);
    out.ins = w.insertTextBox(e, "Box"); const s = Array.from(w.getShapeRects(e)); out.shape = s.map((x) => [x.Id, x.X, x.Y, x.Width, x.Height, x.Story, x.PageIndex]);
    const id = s[0].Id; out.handles = Array.from(w.getShapeHandles(e, id)).slice(0, 1);
    out.enter = w.enterTextBox(e, id); out.rectsWhile = Array.from(w.getShapeRects(e)).length; out.handlesWhile = w.getShapeHandles(e, id).length; out.ctx = w.getCaretContext(e).TextBoxId; w.exitTextBox(e);
    out.vp = w.documentToViewport(e, s[0].X, s[0].Y); out.narrow = narrow; return out; });
  console.log("landscape:", JSON.stringify(r));
  await page.screenshot({ path: `${S}/explore/dbg-landscape-box.png` });
  await page.close();
}
await browser.close();
