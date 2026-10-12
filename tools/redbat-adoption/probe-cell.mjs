import { launch, newPage, open, T, S, SHOTS } from "./lib.mjs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const page = await newPage(context);
await open(page, "word-editor", process.argv[2]);
await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors.at(-1)), null, { timeout: 300000 });
await page.waitForTimeout(300);
const r = await ev(page, () => { const w = window.__word, e = window.__editors.at(-1), m = window.__wordModule; const out = [];
  const cells = m.getTableGeometry(e).filter((g) => g.Kind === "table-cell" && g.Width > 60 && g.Height > 12);
  for (const cell of cells.slice(0, 6)) { const pos = m.hitTestPosition(e, cell.X + Math.min(14, cell.Width / 3), cell.Y + Math.min(12, cell.Height / 2)); if (!pos) { out.push("nopos"); continue; }
    m.setSelection(e, pos.BlockId, 0, pos.BlockId, 0); const c = w.getCaretContext(e); m.setSelection(e, pos.BlockId, 0, pos.BlockId, 100000); const s = m.getSelection(e); const s1 = JSON.stringify(s); m.setSelection(e, pos.BlockId, 3, pos.BlockId, 3); const before = w.getRevision(e); w.insertTextAt ? 0 : 0; out.push({ cell: [cell.Table, cell.Row, cell.Cell, Math.round(cell.X), Math.round(cell.Y), cell.Page], pos, inTable: c.InTable, s1, caret: w.getCaretRect(e) && Math.round(w.getCaretRect(e).Y), sel3: JSON.stringify(m.getSelection(e)) }); }
  w.getInputElement(e).focus({ preventScroll: true });
  return out; });
for (const x of r) console.log(JSON.stringify(x));
await page.keyboard.type("x");
console.log("modified after typing:", await ev(page, () => window.__word.isModified(window.__editors.at(-1))), JSON.stringify(await ev(page, () => window.__wordModule.getSelection(window.__editors.at(-1)))));
await page.screenshot({ path: `${SHOTS}/probe-cell.png` });
await browser.close();
