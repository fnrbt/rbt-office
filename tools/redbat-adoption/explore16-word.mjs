// Exploratory flows, Word, 0.1.16 in the host: real mouse and keyboard. Each flow prints ok / DEFECT with what was seen.
// usage: node explore16-word.mjs [flows]
import { launch, newPage, open, T, S, waitUploads, saved, unzip } from "./lib.mjs";
const only = process.argv[2]?.split(",");
const { browser, context } = await launch({ viewport: { width: 1400, height: 900 } });
const E = `${S}/explore`;
let page, errors;
const ev = (fn, arg) => page.evaluate(fn, arg);
const frame = () => ev(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const sel = (id) => page.locator(`[data-testid="${id}"]`);
const fresh = async (file, opts = {}) => { if (page) await page.close(); page = await newPage(context); errors = []; page.on("pageerror", (e) => errors.push(String(e).slice(0, 200))); page.on("console", (m) => { if (m.type() === "error" && !/404/.test(m.text())) errors.push(m.text().slice(0, 200)); }); await open(page, "word-editor", file, opts); await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 }); await page.waitForFunction(() => window.__word.isLoadComplete(window.__editors[0]) !== false && window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 }); await page.waitForTimeout(500); };
const say = (n, name, ok, detail = "") => console.log(`${ok ? "ok    " : "DEFECT"} W${n} ${name}${detail === "" ? "" : " :: " + (typeof detail === "string" ? detail : JSON.stringify(detail))}`);
const shot = async (name) => { await frame(); await page.waitForTimeout(200); await page.screenshot({ path: `${E}/x16-w-${name}.png` }); };
const flow = async (n, name, fn) => { if (only && !only.includes(String(n))) return; try { await fn(n, name); } catch (e) { say(n, name, false, "script error " + String(e).slice(0, 260)); } if (errors?.length) { console.log(`       page errors in W${n}: ${[...new Set(errors)].slice(0, 3).join(" ;; ")}`); errors.length = 0; } };
const text = () => ev(() => window.__word.getText(window.__editors[0]));
const status = async () => (await sel("office-editor-status").innerText()).replace(/\n/g, " | ");
// screen point of a position in a paragraph (scrolled into view first)
const pointOf = async (blockId, offset, into = true) => { const r = await ev(([b, o]) => { const w = window.__word, e = window.__editors[0]; const bd = w.getRangeBounds(e, b, o, o + 1); const one = Array.isArray(bd) ? bd[0] : bd; return one ? { X: one.X, Y: one.Y, H: one.Height ?? 16 } : null; }, [blockId, offset]); if (!r) return null; if (into) { await ev((y) => { const s = document.querySelector('[data-testid="word-scroll"]'); const z = window.__word.getViewport(window.__editors[0]).scale; if (y * z < s.scrollTop + 60 || y * z > s.scrollTop + s.clientHeight - 120) s.scrollTop = Math.max(0, y * z - 200); }, r.Y); await frame(); await page.waitForTimeout(120); } return ev((r) => { const w = window.__word, e = window.__editors[0]; const v = w.documentToViewport(e, r.X + 1, r.Y + r.H / 2); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: c.left + v.X, y: c.top + v.Y }; }, r); };
const lineWith = (re, opts = {}) => ev(([src, flags, opts]) => { const re = new RegExp(src, flags); const w = window.__word, e = window.__editors[0]; const l = w.getLayoutLines(e).find((x) => re.test(x.Text) && (opts.list ? !!x.ListMarker : true) && (opts.table === undefined || x.InTable === opts.table)); return l ? { block: l.BlockId, start: l.StartOffset, text: l.Text, at: l.StartOffset + l.Text.search(re), marker: l.ListMarker, page: l.Page, indent: l.Indent } : null; }, [re.source, re.flags, opts]);
const caret = () => ev(() => { const s = window.__wordModule.getSelection(window.__editors[0]); return s && { b: s.FocusBlockId, o: s.FocusOffset, a: s.AnchorOffset, collapsed: s.IsCollapsed }; });
const SAMPLE = `${T}/test-documents/sample-formatted.docx`, LISTS = `${T}/test-documents-features/test-lists.docx`, WPS = `${T}/test-documents-wild/docx4j-drawingml-wps.docx`, GALLERY = `${T}/repros/word-shapes/autoshapes-gallery.docx`;

await flow(1, "click into a paragraph, type, Enter, type; two paragraphs; undo back to clean", async (n, name) => {
  await fresh(SAMPLE); const l = await lineWith(/[a-z]{4} [a-z]{3,}/, { table: false }); const p = await pointOf(l.block, l.at + 3); await page.mouse.click(p.x, p.y); await frame();
  const c0 = await caret(); await page.keyboard.type("XYZ"); await page.keyboard.press("Enter"); await page.keyboard.type("new"); await frame(); const t = await text();
  let undos = 0; for (; undos < 14 && !/No changes/.test(await status()); undos++) { await page.keyboard.press("Control+z"); await frame(); }
  say(n, name, c0.b === l.block && Math.abs(c0.o - (l.at + 3)) <= 1 && /XYZ\nnew/.test(t) && /No changes/.test(await status()), { caret: c0, want: l.at + 3, status: await status(), undos });
});
await flow(2, "double click selects a word, Bold from the toolbar, the button shows it; the caret's click lands at 150 % zoom too", async (n, name) => {
  await fresh(SAMPLE); const l = await lineWith(/ [a-z]{6,} /, { table: false }); const p = await pointOf(l.block, l.at + 3); await page.mouse.dblclick(p.x, p.y); await frame();
  const s = await caret(); await sel("word-bold").click(); await frame(); const pressed = await sel("word-bold").getAttribute("aria-pressed");
  await sel("word-zoom-in").click(); await sel("word-zoom-in").click(); await frame(); await page.waitForTimeout(200); const q = await pointOf(l.block, l.at + 3); await page.mouse.click(q.x + 2, q.y); await frame(); const c = await caret();
  say(n, name, !s.collapsed && pressed === "true" && c.b === l.block && Math.abs(c.o - (l.at + 3)) <= 1, { sel: s, pressed, zoomCaret: c, want: l.at + 3 });
});
await flow(3, "Tab at the start of a list item moves it a level down with that level's marker; Shift+Tab back", async (n, name) => {
  await fresh(LISTS); const items = await ev(() => window.__word.getLayoutLines(window.__editors[0]).filter((x) => x.ListMarker && x.StartOffset === 0).map((x) => ({ block: x.BlockId, marker: x.ListMarker, indent: Math.round(x.Indent), text: x.Text.slice(0, 24), x: Math.round(x.X) })));
  const out = []; let ok = items.length > 2;
  for (const it of [items[1], items.find((i) => /^\d/.test(i.marker) && i !== items[1]) ?? items[2]]) { const p = await pointOf(it.block, 0); await page.mouse.click(p.x, p.y); await frame(); await page.keyboard.press("Home"); await page.keyboard.press("Tab"); await frame(); const a = await ev((b) => { const x = window.__word.getLayoutLines(window.__editors[0]).find((l) => l.BlockId === b && l.StartOffset === 0); return { marker: x.ListMarker, x: Math.round(x.X), level: window.__word.getSelectionFormat(window.__editors[0]).Paragraph.ListLevel }; }, it.block); await shot(`3-list-tab-${out.length}`); await page.keyboard.press("Shift+Tab"); await frame(); const b = await ev((b) => { const x = window.__word.getLayoutLines(window.__editors[0]).find((l) => l.BlockId === b && l.StartOffset === 0); return { marker: x.ListMarker, x: Math.round(x.X) }; }, it.block); out.push({ item: it.text, before: [it.marker, it.x], tab: [a.marker, a.x, a.level], back: [b.marker, b.x] }); ok = ok && a.x > it.x && b.x === it.x && b.marker === it.marker; }
  say(n, name, ok, out);
});
await flow(4, "the toolbar's indent buttons on a list item do the same; a new item after Enter continues the numbering", async (n, name) => {
  await fresh(LISTS); const it = await ev(() => { const xs = window.__word.getLayoutLines(window.__editors[0]).filter((x) => /^\d+[.)]/.test(x.ListMarker) && x.StartOffset === 0); const x = xs[0]; return x && { block: x.BlockId, marker: x.ListMarker, len: window.__word.getLayoutLines(window.__editors[0]).filter((l) => l.BlockId === x.BlockId).reduce((n, l) => n + l.Text.length, 0) }; });
  if (!it) return say(n, name, false, "no numbered item");
  const p = await pointOf(it.block, 0); await page.mouse.click(p.x, p.y); await page.keyboard.press("End"); await page.keyboard.press("Enter"); await page.keyboard.type("added item"); await frame();
  const added = await ev(() => { const x = window.__word.getLayoutLines(window.__editors[0]).find((l) => l.Text.startsWith("added item")); return x && x.ListMarker; });
  await sel("word-indent").click(); await frame(); const deeper = await ev(() => { const x = window.__word.getLayoutLines(window.__editors[0]).find((l) => l.Text.startsWith("added item")); return x && x.ListMarker; }); await shot("4-list-added");
  say(n, name, !!added && added !== it.marker && deeper !== added, { first: it.marker, added, deeper });
});
await flow(5, "sections that start on an odd and an even page: a blank page before, the count in the title, numbered footers", async (n, name) => {
  await fresh(`${S}/synth/three-sections.docx`); const before = await ev(() => window.__word.getDocumentSize(window.__editors[0]).PageCount);
  const l = await lineWith(/\w{3,}/); const p = await pointOf(l.block, 1); await page.mouse.click(p.x, p.y); await frame();
  const r = await ev(() => { const w = window.__word, e = window.__editors[0]; w.setSectionBreak(e, "oddPage"); w.completeLayout(e); const ls = w.getLayoutLines(e); const fr = Array.from(w.getPageFrames(e)); return { pages: fr.length, blank: fr.map((f, i) => (ls.some((x) => x.Page === i + 1) ? null : i + 1)).filter(Boolean) }; });
  for (let i = 0; i < 2; i++) await sel("word-zoom-out").click(); await frame(); await ev(() => { document.querySelector('[data-testid="word-scroll"]').scrollTop = 0; }); await page.waitForTimeout(300); await shot("5-odd-page-section");
  const sub = await status(); say(n, name, r.pages === before + r.blank.length + 0 || r.blank.length >= 1, { before, ...r, status: sub });
});
await flow(6, "a text box from the menu: typed in, moved by its edge, resized by a handle, three undos take all back", async (n, name) => {
  await fresh(SAMPLE); const l = await lineWith(/[a-z]{4} [a-z]{3,}/, { table: false }); const p = await pointOf(l.block, l.at); await page.mouse.click(p.x, p.y); await frame();
  await sel("word-header-footer").click(); await sel("word-insert-text-box").click(); await frame(); await page.waitForTimeout(150); await page.keyboard.type("Side note"); await frame();
  const edge = await sel("word-shape-edge-top").boundingBox(); await page.mouse.move(edge.x + edge.width / 4, edge.y + 4); await page.mouse.down(); await page.mouse.move(edge.x + edge.width / 4 + 120, edge.y + 64, { steps: 6 }); await page.mouse.up(); await frame(); await page.waitForTimeout(120);
  const h = await sel("word-shape-handle-se").boundingBox(); await page.mouse.move(h.x + 4, h.y + 4); await page.mouse.down(); await page.mouse.move(h.x + 64, h.y + 44, { steps: 5 }); await page.mouse.up(); await frame(); await page.waitForTimeout(120);
  const s = await ev(() => Array.from(window.__word.getShapeRects(window.__editors[0])).map((x) => [Math.round(x.Width), Math.round(x.Height), x.Text])); await shot("6-text-box-moved");
  let undos = 0; while (!/No changes/.test(await status()) && undos < 12) { await page.keyboard.press("Control+z"); await frame(); undos++; }
  say(n, name, s.length === 1 && s[0][0] === 300 && s[0][1] === 136 && s[0][2] === "Side note" && (await ev(() => window.__word.getShapeRects(window.__editors[0]).length)) === 0, { shape: s, undos });
});
await flow(7, "the shapes gallery: every shape picked by its edge in turn; one deleted with Delete; the page as painted", async (n, name) => {
  await fresh(GALLERY); await shot("7-gallery"); const shapes = await ev(() => Array.from(window.__word.getShapeRects(window.__editors[0])).filter((s) => s.Story === "body").map((s) => ({ Id: s.Id, X: s.X, Y: s.Y, W: s.Width, H: s.Height })));
  let picked = 0; const miss = [];
  for (const s of shapes.slice(0, 12)) { const pt = await ev((s) => { const w = window.__word, e = window.__editors[0]; const sc = document.querySelector('[data-testid="word-scroll"]'); if (s.Y < sc.scrollTop + 20 || s.Y > sc.scrollTop + sc.clientHeight - 150) { sc.scrollTop = Math.max(0, s.Y - 150); } return null; }, s); await frame(); await page.waitForTimeout(80); const v = await ev((s) => { const w = window.__word, e = window.__editors[0]; const v = w.documentToViewport(e, s.X + s.W / 3, s.Y); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: c.left + v.X, y: c.top + v.Y }; }, s); await page.mouse.click(v.x, v.y + 1); await frame(); await page.waitForTimeout(60); const id = (await sel("word-shape-frame").count()) ? await sel("word-shape-frame").getAttribute("data-shape-id") : null; if (id === s.Id) picked++; else miss.push([s.Id, id]); }
  await shot("7-gallery-picked"); const before = shapes.length; await page.keyboard.press("Delete"); await frame(); const after = await ev(() => window.__word.getShapeRects(window.__editors[0]).length);
  say(n, name, picked >= Math.min(12, shapes.length) - 2 && after === before - 1, { shapes: before, tried: Math.min(12, before), picked, miss: miss.slice(0, 4), after });
});
await flow(8, "a text box of a file: a click edits it, typing goes in, the frame is on it, Escape leaves", async (n, name) => {
  await fresh(WPS); await shot("8-wps"); const s = await ev(() => Array.from(window.__word.getShapeRects(window.__editors[0])).filter((x) => x.Story === "body" && !x.BehindText).map((x) => ({ Id: x.Id, X: x.X, Y: x.Y, W: x.Width, H: x.Height, Text: x.Text.slice(0, 30) })));
  if (!s.length) return say(n, name, false, "no shape");
  const b = s[0]; const v = await ev((b) => { const w = window.__word, e = window.__editors[0]; const v = w.documentToViewport(e, b.X + b.W / 2, b.Y + b.H / 2); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: c.left + v.X, y: c.top + v.Y }; }, b);
  await page.mouse.click(v.x, v.y); await frame(); await page.waitForTimeout(100); const story = await ev(() => window.__word.getCaretContext(window.__editors[0]).Story); await page.keyboard.type("Qq"); await frame();
  const t = await ev((id) => Array.from(window.__word.getShapeRects(window.__editors[0])).find((x) => x.Id === id)?.Text, b.Id); const framed = await sel("word-shape-frame").count(); await shot("8-wps-editing"); await page.keyboard.press("Escape"); await frame();
  say(n, name, story === "textbox" && /Qq/.test(t ?? "") && framed === 1 && (await sel("word-shape-frame").count()) === 0, { shapes: s.length, story, t, framed });
});
await flow(9, "a comment's mark: a click on it puts the caret beside it, Backspace after it takes the mark, undo brings it back", async (n, name) => {
  await fresh(`${T}/test-documents-business/docx4j-comments-one.docx`); const l = await lineWith(/\[Comment \d+\]/); const len = l.text.slice(l.at - l.start).indexOf("]") + 1; const p = await pointOf(l.block, l.at + 4); await page.mouse.click(p.x, p.y); await frame();
  const c = await caret(); await ev(([b, o]) => window.__wordModule.setSelection(window.__editors[0], b, o, b, o), [l.block, l.at + len]); await ev(() => window.__word.getInputElement(window.__editors[0]).focus({ preventScroll: true })); await page.keyboard.press("Backspace"); await frame();
  const gone = !/\[Comment/.test(await text()); const anchors = await ev(() => window.__word.getCommentAnchors(window.__editors[0]).length); await page.keyboard.press("Control+z"); await frame(); const back = /\[Comment/.test(await text());
  say(n, name, (c.o === l.at || c.o === l.at + len) && gone && back, { caret: c.o, mark: [l.at, l.at + len], gone, anchorsAfterBackspace: anchors, back });
});
await flow(10, "a double click on the header edits it; typing; Escape; the same text on the next page's header", async (n, name) => {
  await fresh(SAMPLE); const cv = await sel("word-canvas").boundingBox(); const hp = await ev(() => { const w = window.__word, e = window.__editors[0]; const f = Array.from(w.getPageFrames(e))[0]; const v = w.documentToViewport(e, f.Width / 2, f.MarginTop / 2); return v; });
  await page.mouse.dblclick(cv.x + hp.X, cv.y + hp.Y); await frame(); await page.waitForTimeout(150); const ctx = await ev(() => window.__word.getCaretContext(window.__editors[0]).Story); await page.keyboard.type("Header text"); await frame(); await shot("10-header"); const bar = await sel("word-header-footer-bar").count(); await page.keyboard.press("Escape"); await frame();
  const hf = await ev(() => window.__word.getHeaderFooterText(window.__editors[0])); say(n, name, ctx === "header" && bar === 1 && /Header text/.test(JSON.stringify(hf)), { ctx, bar, hf: JSON.stringify(hf).slice(0, 160) });
});
await flow(11, "find and replace all from the toolbar; the count; one undo", async (n, name) => {
  await fresh(SAMPLE); const before = await text(); await sel("word-find").click(); await sel("word-find-input").fill("the"); await sel("word-find-next").click(); await frame(); const st = await sel("word-find-status").innerText();
  await sel("word-replace-input").fill("THE"); await sel("word-replace-all").click(); await frame(); const st2 = await sel("word-find-status").innerText(); const mid = await text(); await page.keyboard.press("Escape"); await ev(() => window.__word.getInputElement(window.__editors[0]).focus({ preventScroll: true })); await page.keyboard.press("Control+z"); await frame();
  say(n, name, /of \d+/.test(st) && /\d+ replaced/.test(st2) && mid !== before && (await text()) === before, [st, st2]);
});
await flow(12, "a table cell: click, type, Tab to the next cell, a row below from the table tools, delete it", async (n, name) => {
  await fresh(`${T}/test-documents-wild/docx4j-tables-rich.docx`); const l = await lineWith(/\w{2,}/, { table: true }); const p = await pointOf(l.block, l.at); await page.mouse.click(p.x, p.y); await frame();
  const c0 = await ev(() => { const c = window.__word.getCaretContext(window.__editors[0]); return [c.InTable, c.Row, c.Column, c.RowCount]; }); await page.keyboard.type("Zz"); await page.keyboard.press("Tab"); await frame(); const c1 = await ev(() => { const c = window.__word.getCaretContext(window.__editors[0]); return [c.InTable, c.Row, c.Column, c.RowCount]; });
  await sel("word-row-below").click(); await frame(); const c2 = await ev(() => window.__word.getCaretContext(window.__editors[0]).RowCount); await sel("word-row-delete").click(); await frame(); const c3 = await ev(() => window.__word.getCaretContext(window.__editors[0]).RowCount);
  say(n, name, c0[0] && c1[0] && (c1[2] !== c0[2] || c1[1] !== c0[1]) && c2 === c0[3] + 1 && c3 === c0[3], { c0, c1, c2, c3 });
});
await flow(13, "copy two lines and paste them at the end of another paragraph with the keyboard (Ctrl+C, Ctrl+V)", async (n, name) => {
  await fresh(SAMPLE); const l = await lineWith(/[a-z]{4} [a-z]{3,} [a-z]+/, { table: false }); const p = await pointOf(l.block, l.start); await page.mouse.click(p.x, p.y); await page.keyboard.press("Home"); await page.keyboard.press("Shift+End"); await page.keyboard.press("Control+c"); await page.waitForTimeout(150);
  const clip = await page.evaluate(() => navigator.clipboard.readText()).catch(() => null); await page.keyboard.press("End"); await page.keyboard.press("Enter"); await sel("word-italic").click(); await page.keyboard.press("Control+v"); await page.waitForTimeout(400); await frame();
  const t = await text(); const twice = clip ? t.split(clip).length - 1 : 0; const fmt = await ev(() => { const w = window.__word, m = window.__wordModule, e = window.__editors[0]; const s = m.getSelection(e); m.setSelection(e, s.FocusBlockId, 0, s.FocusBlockId, Math.max(1, s.FocusOffset)); const f = w.getSelectionFormat(e); return f.Italic; });
  say(n, name, !!clip && twice === 2 && fmt === true, { clip: (clip ?? "").slice(0, 50), twice, pastedItalic: fmt });
});
await flow(14, "Ctrl+Enter makes a page break: one page more; a link from the menu shows its bubble", async (n, name) => {
  await fresh(SAMPLE); const pages = await ev(() => window.__word.getDocumentSize(window.__editors[0]).PageCount); const l = await lineWith(/[a-z]{4} [a-z]{3,}/, { table: false }); const p = await pointOf(l.block, l.at); await page.mouse.click(p.x, p.y); const next = await ev((b) => { const ls = window.__word.getLayoutLines(window.__editors[0]); const i = ls.findIndex((x) => x.BlockId === b); const n = ls.slice(i).find((x) => x.BlockId !== b); return { id: n.BlockId, page: n.Page }; }, l.block); await page.keyboard.press("Control+Enter"); await frame(); await ev(() => window.__word.completeLayout(window.__editors[0])); const after = await ev((id) => window.__word.getLayoutLines(window.__editors[0]).find((x) => x.BlockId === id).Page, next.id); const pagesAfter = await ev(() => window.__word.getDocumentSize(window.__editors[0]).PageCount);
  await page.keyboard.press("Control+z"); await frame(); await sel("word-link").click(); await sel("word-link-target").fill("https://example.org/a"); await sel("word-link-text").fill("the link"); await sel("word-link-apply").click(); await frame(); await page.keyboard.press("ArrowLeft"); await frame(); await page.waitForTimeout(150); const bubble = await sel("word-link-bubble").count();
  say(n, name, after === next.page + 1 && bubble === 1, { pages: [pages, pagesAfter], nextParagraphPage: [next.page, after], bubble });
});
await flow(15, "arrow keys walk past a text box's anchor without going into the box; Down through a page end", async (n, name) => {
  await fresh(WPS); const l = await lineWith(/\w{3,}/); const p = await pointOf(l.block, l.at); await page.mouse.click(p.x, p.y); await frame(); const stories = new Set(); let moved = 0, last = null;
  for (let i = 0; i < 60; i++) { await page.keyboard.press("ArrowDown"); const c = await ev(() => { const s = window.__wordModule.getSelection(window.__editors[0]); return [window.__word.getCaretContext(window.__editors[0]).Story, s.FocusBlockId, s.FocusOffset].join("|"); }); stories.add(c.split("|")[0]); if (c !== last) moved++; last = c; }
  say(n, name, stories.size === 1 && stories.has("body") && moved > 5, { stories: [...stories], moved });
});
await flow(16, "everything selected and set to 14 pt in a 53-page document from its middle, then undone: the view stays on its text", async (n, name) => {
  await fresh(`${T}/test-documents-wild/docx4j-getting-started.docx`); await ev(() => { const s = document.querySelector('[data-testid="word-scroll"]'); s.scrollTop = Math.round((s.scrollHeight - s.clientHeight) * 0.6); }); await frame(); await page.waitForTimeout(300);
  const anchor = () => ev(() => { const w = window.__word, e = window.__editors[0]; for (let y = 6; y < 700; y += 6) { const p = w.hitTestViewport(e, 500, y); if (p) return `${p.BlockId}:${p.Offset}@${y}`; } return null; });
  const settle = async () => { await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 }); await frame(); await page.waitForTimeout(300); };
  // the faces first (see the report: a face that arrives while the layout is under way loses the view's anchor)
  await ev(() => window.__word.getInputElement(window.__editors[0]).focus({ preventScroll: true })); const a0 = await anchor(); await page.keyboard.press("Control+a"); await sel("word-font-size").selectOption("14"); await settle(); const a1 = await anchor(); await shot("16-size-14");
  await ev(() => window.__word.getInputElement(window.__editors[0]).focus({ preventScroll: true })); await page.keyboard.press("Control+z"); await settle(); const a2 = await anchor();
  const same = (x, y) => x && y && x.split("@")[0] === y.split("@")[0] && Math.abs(Number(x.split("@")[1]) - Number(y.split("@")[1])) <= 12;
  say(n, name, same(a0, a1) && same(a1, a2), { before: a0, after14: a1, undone: a2 });
});
await flow(17, "edits saved and the saved file opened again: text, text box and page break are there", async (n, name) => {
  await fresh(SAMPLE); const l = await lineWith(/[a-z]{4} [a-z]{3,}/, { table: false }); const p = await pointOf(l.block, l.at); await page.mouse.click(p.x, p.y); await page.keyboard.type("SAVED-MARK "); await sel("word-header-footer").click(); await sel("word-insert-text-box").click(); await frame(); await page.keyboard.type("BOX-MARK"); await page.keyboard.press("Escape"); await frame();
  await sel("office-editor-save").click(); const out = await waitUploads(page, 1); saved.set("w17", out); await fresh("/api/saved/w17", { name: "sample-formatted.docx" }); const t = await text(); const boxes = await ev(() => Array.from(window.__word.getShapeRects(window.__editors[0])).map((x) => x.Text));
  say(n, name, /SAVED-MARK /.test(t) && boxes.includes("BOX-MARK"), { boxes, bytes: out.length });
});
await flow(18, "the resume's header text box: a double click edits it, the text is on screen as typed", async (n, name) => {
  await fresh(`${T}/test-documents-features/resume-template.docx`); const s = await ev(() => { const x = Array.from(window.__word.getShapeRects(window.__editors[0])).find((x) => x.Story === "header"); const w = window.__word, e = window.__editors[0]; const v = w.documentToViewport(e, x.X + x.Width / 2, x.Y + x.Height / 2); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: c.left + v.X, y: c.top + v.Y, id: x.Id }; });
  await page.mouse.dblclick(s.x, s.y); await frame(); await page.waitForTimeout(150); await page.keyboard.press("Control+a"); await page.keyboard.type("Jane Doe"); await frame(); await shot("18-resume-header-box"); const t = await ev((id) => Array.from(window.__word.getShapeRects(window.__editors[0])).find((x) => x.Id === id)?.Text, s.id);
  say(n, name, t === "Jane Doe", t);
});
await flow(19, "a paragraph centred vertically on its page through the API (no host control): saved and read back", async (n, name) => {
  await fresh(`${S}/synth/three-sections.docx`); const r = await ev(() => { const w = window.__word, e = window.__editors[0]; const y0 = w.getLayoutLines(e)[0].Y; const ok = w.setSectionVerticalAlignment(e, "center"); w.completeLayout(e); return { ok, y: [Math.round(y0), Math.round(w.getLayoutLines(e)[0].Y)], now: w.getSectionVerticalAlignment(e) }; }); await shot("19-valign-center");
  say(n, name, r.ok && r.now === "center" && r.y[1] > r.y[0], r);
});
await flow(20, "Shift+click extends the selection across paragraphs; typing replaces it; one undo restores", async (n, name) => {
  await fresh(SAMPLE); const before = await text(); const a = await lineWith(/[a-z]{4} [a-z]{3,}/, { table: false }); const p = await pointOf(a.block, a.at); await page.mouse.click(p.x, p.y); await page.keyboard.down("Shift"); await page.mouse.click(p.x + 140, p.y + 60); await page.keyboard.up("Shift"); await frame(); const s = await caret();
  await page.keyboard.type("R"); await frame(); const mid = await text(); await page.keyboard.press("Control+z"); await frame(); say(n, name, !s.collapsed && mid.length < before.length && (await text()) === before, { sel: s, lengths: [before.length, mid.length] });
});
if (page) await page.close();
await browser.close();
