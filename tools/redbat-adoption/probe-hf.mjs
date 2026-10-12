import { launch, newPage, open, T } from "./lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "word-editor", `${T}/repros/word/landscape-section.docx`);
await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
await page.waitForTimeout(400);
console.log(await page.evaluate(() => {
  const w = window.__word, e = window.__editors[0]; const out = [];
  const snap = (label) => { const c = w.getCaretContext(e); out.push(`${label}: story ${c.Story}, hf ${JSON.stringify(c.HeaderFooter)}, modified ${w.isModified(e)}, canUndo ${w.canUndo(e)}, rev ${w.getRevision(e)}`); };
  snap("start");
  out.push("enter -> " + w.enterHeaderFooter(e, 2, "header")); snap("entered page 3 header");
  out.push("link(true) -> " + w.setHeaderFooterLinkToPrevious(e, true)); snap("linked");
  w.undo(e); snap("after undo 1");
  w.undo(e); snap("after undo 2");
  return out.join("\n");
}));
await browser.close();
