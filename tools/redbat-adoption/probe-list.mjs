import { launch, newPage, open, T } from "./lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "word-editor", process.argv[2]);
await page.waitForSelector('[data-testid="word-toolbar"]'); await page.waitForTimeout(1500);
const r = await page.evaluate(() => { const w = window.__word, m = window.__wordModule, e = window.__editors[0]; const out = []; const blocks = m.getLayoutFingerprint(e); for (let i = 0; i < blocks.length; i++) { const t = blocks[i].Text ?? ""; if (t.length < 3) continue; const f = w.findText(e, t.slice(0, 40), true); if (!f.length) continue; w.selectFindMatch(e, f[0]); const s = m.getSelection(e); m.setSelection(e, s.AnchorBlockId, 0, s.AnchorBlockId, 0); const c = w.getCaretContext(e); const p = w.getSelectionFormat(e).Paragraph; if (c.InList || p.ListKind !== "none") out.push([i, t.slice(0, 25), f.length, Math.min(s.AnchorOffset, s.FocusOffset), c.InList, c.InTable, p.ListKind, p.ListLevel, p.IndentLeft]); } return out.slice(0, 12); });
for (const x of r) console.log(JSON.stringify(x));
await browser.close();
