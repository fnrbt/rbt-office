import { launch, newPage, open, T, S } from "./lib.mjs";
import { g, sel, ready, frames, cellPoint, cell } from "./ui-lib.mjs";
const { browser, context } = await launch({ viewport: { width: 1400, height: 820 } });
{ // S2: chart after undo
  const page = await newPage(context); await open(page, "sheet-editor", `${T}/real-world/chart-zoo.xlsx`); await ready(page);
  const texts = () => g(page, () => { const o = document.querySelector('[data-testid="sheet-grid"] .spreadsheet-object-drawn'); return [...o.querySelectorAll("text")].map((t) => t.textContent).slice(0, 10).join(" ") + " #rects " + o.querySelectorAll("rect").length + " len " + o.querySelector("svg").outerHTML.length; });
  console.log("S2 before:", await texts()); const p = await cellPoint(page, 1, 1); await page.mouse.click(p.x, p.y); await page.keyboard.type("55"); await page.keyboard.press("Enter"); await frames(page); await page.waitForTimeout(300);
  console.log("S2 edited:", await texts()); await page.keyboard.press("Control+z"); await frames(page); await page.waitForTimeout(500); console.log("S2 undone:", await texts(), "| B2 =", (await cell(page, 1, 1)).text);
  const bars = () => g(page, () => [...document.querySelector('[data-testid="sheet-grid"] .spreadsheet-object-drawn').querySelectorAll("rect")].map((r) => Math.round(Number(r.getAttribute("height")))).filter((h) => h > 0 && h < 260).join(","));
  console.log("S2 bar heights after undo:", await bars());
  // S5: widen column B by its header edge
  const w0 = await g(page, () => window.__grids.at(-1).GetColumnWidth?.(1));
  const hs = await g(page, () => [...document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-col-header')].map((e) => { const b = e.getBoundingClientRect(); return { r: b.right, y: b.top + b.height / 2, w: b.width, col: e.dataset.col, cur: getComputedStyle(e).cursor }; }).filter((b) => b.w > 0).slice(0, 4));
  console.log("S5 headers:", JSON.stringify(hs));
  const h = hs.find((x) => x.col === "1"); const x0 = await g(page, () => Math.round(document.querySelector('[data-testid="sheet-grid"] .spreadsheet-object-drawn').getBoundingClientRect().x));
  for (const off of [0, -2, 2]) { await page.mouse.move(h.r + off, h.y); await page.mouse.down(); await page.mouse.move(h.r + off + 60, h.y, { steps: 6 }); await page.mouse.up(); await frames(page); await page.waitForTimeout(200); const hw = await g(page, () => Math.round([...document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-col-header')].find((e) => e.dataset.col === "1").getBoundingClientRect().width)); const x1 = await g(page, () => Math.round(document.querySelector('[data-testid="sheet-grid"] .spreadsheet-object-drawn').getBoundingClientRect().x)); console.log(`S5 drag from edge${off >= 0 ? "+" : ""}${off}: column B width ${hw}, chart x ${x0} -> ${x1}`); if (hw > 80) break; }
  await page.close();
}
{ // W14: Ctrl+Enter
  const page = await newPage(context); await open(page, "word-editor", `${T}/test-documents/sample-formatted.docx`); await page.waitForSelector('[data-testid="word-toolbar"]'); await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors[0])); await page.waitForTimeout(600);
  const info = () => page.evaluate(() => { const w = window.__word, e = window.__editors[0]; w.completeLayout(e); const ls = w.getLayoutLines(e); return { pages: w.getDocumentSize(e).PageCount, blocks: new Set(ls.map((l) => l.BlockId)).size, sel: window.__wordModule.getSelection(e).FocusBlockId + ":" + window.__wordModule.getSelection(e).FocusOffset, page2first: ls.find((l) => l.Page === 2)?.Text.slice(0, 30) }; });
  await page.evaluate(() => { const w = window.__word, m = window.__wordModule, e = window.__editors[0]; const l = w.getLayoutLines(e).find((x) => x.Text.length > 40 && !x.InTable); m.setSelection(e, l.BlockId, 10, l.BlockId, 10); w.getInputElement(e).focus({ preventScroll: true }); });
  console.log("W14 before:", JSON.stringify(await info())); await page.keyboard.press("Control+Enter"); await page.waitForTimeout(200); console.log("W14 after Ctrl+Enter:", JSON.stringify(await info()), "modified", await page.evaluate(() => window.__word.isModified(window.__editors[0])));
  await page.evaluate(() => window.__word.insertPageBreak(window.__editors[0])); console.log("W14 after insertPageBreak():", JSON.stringify(await info()));
  await page.close();
}
await browser.close();
