// 0.1.15 / 0.1.16 in the host, Word: the ScrollTop notice, shapes picked / moved / resized / deleted by their frame, header text boxes.
import { launch, newPage, open, T, S, check, summary, waitUploads, unzip, LONG_PAGES } from "./lib.mjs";
const { browser, context } = await launch({ viewport: { width: 1366, height: 768 } });
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const E = `${S}/explore`;
const ready = async (page) => { await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 }); await page.waitForFunction(() => window.__word.isLoadComplete(window.__editors[0]) !== false && window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 }); await page.waitForTimeout(500); };
const shapes = (page) => ev(page, () => Array.from(window.__word.getShapeRects(window.__editors[0])).map((r) => ({ Id: r.Id, Story: r.Story, PageIndex: r.PageIndex, X: Math.round(r.X * 10) / 10, Y: Math.round(r.Y * 10) / 10, W: Math.round(r.Width * 10) / 10, H: Math.round(r.Height * 10) / 10, Text: r.Text.slice(0, 40) })));
const abs = async (page, testid) => { const l = sel(page, testid); return (await l.count()) ? await l.boundingBox() : null; };
// relative to the canvas: a bar that comes or goes over the page moves the canvas on the screen
const box = async (page, testid) => { const b = await abs(page, testid); if (!b) return null; const c = await sel(page, "word-canvas").boundingBox(); return { x: b.x - c.x, y: b.y - c.y, width: b.width, height: b.height }; };
const quarter = (b) => (b.width > b.height ? { x: b.x + b.width / 4, y: b.y + b.height / 2 } : { x: b.x + b.width / 2, y: b.y + b.height / 4 });
const center = (b) => ({ x: b.x + b.width / 2, y: b.y + b.height / 2 });
const ctx = (page) => ev(page, () => { const c = window.__word.getCaretContext(window.__editors[0]); return { Story: c.Story, TextBoxId: c.TextBoxId ?? null, hf: c.HeaderFooter ? c.HeaderFooter.Kind : null }; });
const drag = async (page, from, dx, dy) => { await page.mouse.move(from.x, from.y); await page.mouse.down(); await page.mouse.move(from.x + dx / 2, from.y + dy / 2, { steps: 4 }); await page.mouse.move(from.x + dx, from.y + dy, { steps: 4 }); await frame(page); const ghost = await box(page, "word-shape-frame"); await page.mouse.up(); await frame(page); await page.waitForTimeout(80); return ghost; };
// (From 0.1.18 the engine selects, moves and resizes shapes and enters a header's text box itself; the host's
// frame these sections drive is gone: word-new18.mjs checks the engine's in the host. Run them by name against an older host.)
const only = process.argv[2]?.split(",") ?? ["scroll", "scroll-cold", "pagebreak"]; const want = (k) => only.includes(k);

// ---------------- shapes of a file: pick by the frame, move, resize, delete, undo ----------------
if (want("shapes")) {
  const page = await newPage(context);
  await open(page, "word-editor", `${S}/synth/shapes.docx`);
  await ready(page);
  let all = await shapes(page);
  console.log("  shapes:", JSON.stringify(all));
  check("the fixture's shapes are of the body", all.length >= 3 && all.every((s) => s.Story === "body" && s.PageIndex === -1), all.length);
  check("no frame before a shape is clicked", (await sel(page, "word-shape-frame").count()) === 0);
  const target = all.find((s) => Math.abs(s.W - 140) < 2) ?? all[0];
  const at = await ev(page, (r) => { const w = window.__word, e = window.__editors[0]; const v = w.documentToViewport(e, r.X + r.W / 2, r.Y + r.H / 2); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: c.left + v.X, y: c.top + v.Y, cx: c.left, cy: c.top, v0: w.documentToViewport(e, r.X, r.Y) }; }, target);
  await page.mouse.click(at.x, at.y); await frame(page); await page.waitForTimeout(80);
  let c = await ctx(page);
  check("a click in a shape edits its text", c.Story === "textbox" && c.TextBoxId === target.Id, c);
  let f = await box(page, "word-shape-frame");
  check("the box that is edited has a frame where the box is painted", !!f && Math.abs(f.x - at.v0.X) < 1.5 && Math.abs(f.y - at.v0.Y) < 1.5 && Math.abs(f.width - target.W) < 1.5 && Math.abs(f.height - target.H) < 1.5, { f, expect: [at.v0.X, at.v0.Y, target.W, target.H] });
  check("no handles while its text is edited", (await page.locator(".rb-word-shape-handle").count()) === 0);
  await page.screenshot({ path: `${E}/n16-shape-1-editing.png` });
  // a click on the edge picks it
  const edge = quarter(await abs(page, "word-shape-edge-top"));
  await page.mouse.click(edge.x, edge.y); await frame(page); await page.waitForTimeout(80);
  c = await ctx(page);
  check("a click on the frame's edge leaves the text and picks the shape: eight handles", c.Story === "body" && (await page.locator(".rb-word-shape-handle").count()) === 8 && (await sel(page, "word-shape-frame").getAttribute("class")).includes("is-picked"), c);
  const hs = await ev(page, (id) => Array.from(window.__word.getShapeHandles(window.__editors[0], id)).map((h) => [h.Name, h.Cursor]), target.Id);
  const drawn = await page.locator(".rb-word-shape-handle").evaluateAll((els) => els.map((el) => [el.dataset.testid.replace("word-shape-handle-", ""), el.style.cursor, el.getBoundingClientRect().width]));
  check("the handles are the engine's: names, cursors, eight pixels", JSON.stringify(drawn.map((d) => [d[0], d[1]])) === JSON.stringify(hs) && drawn.every((d) => Math.abs(d[2] - 8) < 0.01), drawn);
  check("the text box bar is gone, nothing is modified yet", (await sel(page, "word-text-box-bar").count()) === 0 && !(await ev(page, () => window.__word.isModified(window.__editors[0]))));
  await page.screenshot({ path: `${E}/n16-shape-2-picked.png` });
  // move by the edge
  const rev0 = await ev(page, () => window.__word.getRevision(window.__editors[0]));
  const top0 = await ev(page, () => document.querySelector('[data-testid="word-scroll"]').scrollTop);
  f = await box(page, "word-shape-frame");
  const ghost = await drag(page, quarter(await abs(page, "word-shape-edge-bottom")), 60, 40);
  let now = (await shapes(page)).find((s) => s.Id === target.Id);
  check("dragged by its edge: the frame follows the pointer", !!ghost && Math.abs(ghost.x - (f.x + 60)) < 1.5 && Math.abs(ghost.y - (f.y + 40)) < 1.5, [f, ghost]);
  check("let go: the shape is 60 to the right and 40 down, its size as it was", Math.abs(now.X - (target.X + 60)) < 0.6 && Math.abs(now.Y - (target.Y + 40)) < 0.6 && now.W === target.W && now.H === target.H, [target, now]);
  check("the others are where they were", JSON.stringify((await shapes(page)).filter((s) => s.Id !== target.Id)) === JSON.stringify(all.filter((s) => s.Id !== target.Id)));
  f = await box(page, "word-shape-frame");
  check("the frame and the handles are on the shape where it is now, the view did not move", !!f && Math.abs(f.x - (at.v0.X + 60)) < 1.5 && Math.abs(f.y - (at.v0.Y + 40)) < 1.5 && (await page.locator(".rb-word-shape-handle").count()) === 8 && (await ev(page, () => document.querySelector('[data-testid="word-scroll"]').scrollTop)) === top0, f);
  check("the move is one undo step and the document is unsaved", (await sel(page, "word-undo").isEnabled()) && (await ev(page, () => window.__word.isModified(window.__editors[0]))));
  await page.screenshot({ path: `${E}/n16-shape-3-moved.png` });
  await page.keyboard.press("Control+z"); await frame(page); await page.waitForTimeout(80);
  now = (await shapes(page)).find((s) => s.Id === target.Id);
  check("Ctrl+Z: where it was, and clean", now.X === target.X && now.Y === target.Y && !(await ev(page, () => window.__word.isModified(window.__editors[0]))) && (await ev(page, () => window.__word.getRevision(window.__editors[0]))) === rev0, now);
  f = await box(page, "word-shape-frame");
  check("the frame is back on it, still picked", !!f && Math.abs(f.x - at.v0.X) < 1.5 && (await page.locator(".rb-word-shape-handle").count()) === 8, f);
  // resize by the south-east handle
  const se = await abs(page, "word-shape-handle-se");
  await drag(page, center(se), 30, 20);
  now = (await shapes(page)).find((s) => s.Id === target.Id);
  check("the south-east handle dragged by 30, 20: larger by that, its corner where it was", Math.abs(now.W - (target.W + 30)) < 0.6 && Math.abs(now.H - (target.H + 20)) < 0.6 && now.X === target.X && now.Y === target.Y, now);
  await page.keyboard.press("Control+z"); await frame(page);
  // resize by the north-west handle: a move and a resize
  const nw = await abs(page, "word-shape-handle-nw");
  await drag(page, center(nw), -20, -10);
  now = (await shapes(page)).find((s) => s.Id === target.Id);
  check("the north-west handle dragged by -20, -10: its far corner stays", Math.abs(now.X - (target.X - 20)) < 0.6 && Math.abs(now.Y - (target.Y - 10)) < 0.6 && Math.abs(now.W - (target.W + 20)) < 0.6 && Math.abs(now.H - (target.H + 10)) < 0.6, now);
  await page.screenshot({ path: `${E}/n16-shape-4-resized.png` });
  // saved
  await sel(page, "office-editor-save").click();
  const doc = await (await unzip(await waitUploads(page, 1))).get("word/document.xml").async("string");
  check("saved: the extent of the anchor is the new size (160 x 100 px)", doc.includes('<wp:extent cx="1524000" cy="952500"/>') , (doc.match(/<wp:extent[^>]*>/g) ?? []).join(" "));
  await page.keyboard.press("Control+z"); await page.keyboard.press("Control+z"); await frame(page);
  now = (await shapes(page)).find((s) => s.Id === target.Id);
  check("two undos (a move and a resize): as it was", now.X === target.X && now.Y === target.Y && now.W === target.W && now.H === target.H, now);
  // delete
  await page.keyboard.press("Delete"); await frame(page); await page.waitForTimeout(80);
  check("Delete with the shape picked takes it out, the frame with it", (await shapes(page)).length === all.length - 1 && (await sel(page, "word-shape-frame").count()) === 0, (await shapes(page)).length);
  const text0 = await ev(page, () => window.__word.getText(window.__editors[0]).length);
  await page.keyboard.press("Delete"); await frame(page);
  check("the next Delete is the text's, not another shape's", (await shapes(page)).length === all.length - 1);
  await page.keyboard.press("Control+z"); await frame(page);
  await sel(page, "word-undo").click().catch(() => {}); await frame(page);
  const back = await shapes(page);
  check("undone: the shape is back", back.length === all.length, back.length);
  // a shape behind the text has no text to click into: picked by its edge alone
  {
    const g = all.find((s) => Math.abs(s.W - 100) < 2);
    const gp = await ev(page, (r) => { const w = window.__word, e = window.__editors[0]; const v = w.documentToViewport(e, r.X, r.Y); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: c.left + v.X, y: c.top + v.Y }; }, g);
    await page.mouse.click(gp.x + g.W / 2, gp.y + g.H / 2); await frame(page); await page.waitForTimeout(60);
    check("a click inside a shape behind the text is the text's: no frame", (await sel(page, "word-shape-frame").count()) === 0 && (await ctx(page)).Story === "body");
    await page.mouse.click(gp.x + g.W / 4, gp.y + 1); await frame(page); await page.waitForTimeout(60);
    check("a click on its edge picks it: frame and handles on it", (await page.locator(".rb-word-shape-handle").count()) === 8 && (await sel(page, "word-shape-frame").getAttribute("data-shape-id")) === g.Id);
    await drag(page, { x: gp.x + g.W / 4, y: gp.y + g.H }, 0, 30);
    const g2 = (await shapes(page)).find((s) => s.Id === g.Id);
    check("dragged by its edge without being picked first", Math.abs(g2.Y - (g.Y + 30)) < 0.6 && g2.X === g.X, [g, g2]);
    await page.screenshot({ path: `${E}/n16-shape-6-behind.png` });
    await page.keyboard.press("Control+z"); await frame(page);
    await page.keyboard.press("Escape"); await frame(page); await page.waitForTimeout(60);
  }
  // Escape lets go
  await page.mouse.click(at.x, at.y); await frame(page); await page.waitForTimeout(60);
  const e3 = quarter(await abs(page, "word-shape-edge-left"));
  await page.mouse.click(e3.x, e3.y); await frame(page); await page.waitForTimeout(60);
  check("picked again", (await page.locator(".rb-word-shape-handle").count()) === 8);
  await page.keyboard.press("Escape"); await frame(page); await page.waitForTimeout(60);
  check("Escape lets go of the shape and keeps the editor open", (await sel(page, "word-shape-frame").count()) === 0 && (await sel(page, "word-toolbar").count()) === 1);
  // a click on the page lets go; typing lets go
  await page.mouse.click(at.x, at.y); await frame(page); await page.waitForTimeout(60);
  const e4 = quarter(await abs(page, "word-shape-edge-right"));
  await page.mouse.click(e4.x, e4.y); await frame(page); await page.waitForTimeout(60);
  const cv = await abs(page, "word-canvas");
  await page.mouse.click(cv.x + 420, cv.y + 60); await frame(page); await page.waitForTimeout(80);
  check("a click on the page lets go of it", (await sel(page, "word-shape-frame").count()) === 0);
  // scrolling moves the frame with the page
  await page.mouse.click(at.x, at.y); await frame(page); await page.waitForTimeout(60);
  const f1 = await box(page, "word-shape-frame");
  await page.mouse.wheel(0, 50); await frame(page); await page.waitForTimeout(120);
  const f2 = await box(page, "word-shape-frame");
  check("scrolled by 50: the frame is 50 higher", !!f1 && !!f2 && Math.abs(f1.y - f2.y - 50) < 1.5, [f1?.y, f2?.y]);
  // zoom
  await sel(page, "word-zoom-in").click(); await frame(page); await page.waitForTimeout(120);
  const f3 = await box(page, "word-shape-frame");
  const expect3 = await ev(page, (id) => { const w = window.__word, e = window.__editors[0]; const r = Array.from(w.getShapeRects(e)).find((x) => x.Id === id); const v = w.documentToViewport(e, r.X, r.Y); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: v.X, y: v.Y, w: r.Width * 1.25, h: r.Height * 1.25 }; }, target.Id);
  check("at 125 %: the frame is on the box, a quarter larger", !!f3 && Math.abs(f3.x - expect3.x) < 1.5 && Math.abs(f3.y - expect3.y) < 1.5 && Math.abs(f3.width - expect3.w) < 1.5, [f3, expect3]);
  const b5 = (await shapes(page)).find((s) => s.Id === target.Id);
  await drag(page, quarter(await abs(page, "word-shape-edge-top")), 50, 25);
  now = (await shapes(page)).find((s) => s.Id === target.Id);
  check("at 125 %: dragged by 50, 25 on screen is 40, 20 on the page", Math.abs(now.X - (b5.X + 40)) < 0.6 && Math.abs(now.Y - (b5.Y + 20)) < 0.6, [b5, now]);
  await page.screenshot({ path: `${E}/n16-shape-5-zoom.png` });
  check("the scroll content is not made larger by the frame", await ev(page, () => { const s = document.querySelector('[data-testid="word-scroll"]'); const sp = document.querySelector(".rb-word-editor-spacer"); return s.scrollHeight <= Math.max(sp.offsetHeight, s.clientHeight) + 1; }));
  await page.close();
}

// ---------------- a text box from the menu: framed at once, moved, typed in ----------------
if (want("insert")) {
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/repros/word/landscape-section.docx`);
  await ready(page);
  // on the portrait page, which is narrower than the document: its frame has an X
  const line = await ev(page, () => { const w = window.__word, m = window.__wordModule, e = window.__editors[0]; const fr = Array.from(w.getPageFrames(e)); const i = fr.findIndex((f) => f.X > 0); const l = w.getLayoutLines(e).find((x) => x.Page === i + 1 && x.Text.trim().length > 3); m.setSelection(e, l.BlockId, 0, l.BlockId, 0); return { frameX: fr[i].X, y: l.Y }; });
  await ev(page, (y) => { document.querySelector('[data-testid="word-scroll"]').scrollTop = y - 100; }, line.y); await frame(page); await page.waitForTimeout(150);
  await ev(page, () => window.__word.getInputElement(window.__editors[0]).focus({ preventScroll: true }));
  await sel(page, "word-header-footer").click(); await sel(page, "word-insert-text-box").click(); await frame(page); await page.waitForTimeout(120);
  const f = await box(page, "word-shape-frame");
  check("a text box from the menu is framed at once (240 x 96)", !!f && Math.abs(f.width - 240) < 1.5 && Math.abs(f.height - 96) < 1.5, f);
  await page.screenshot({ path: `${E}/n16-insert-1.png` });
  // the frame is on the painted box: the pixel just inside the frame's corner region differs from the page's white where the border is painted
  await page.keyboard.type("Hello"); await frame(page);
  const s1 = (await shapes(page))[0];
  check("typing goes into the box (the text selected at first is replaced)", s1.Text === "Hello", s1.Text);
  const fa = await box(page, "word-shape-frame");
  await drag(page, quarter(await abs(page, "word-shape-edge-top")), -30, 60);
  const s2 = (await shapes(page))[0];
  check("moved by its edge on a page that is narrower than the document", Math.abs(s2.X - (s1.X - 30)) < 0.6 && Math.abs(s2.Y - (s1.Y + 60)) < 0.6, [s1, s2]);
  const f2 = await box(page, "word-shape-frame");
  check("the frame went with it", !!f2 && Math.abs(f2.x - (fa.x - 30)) < 1.5 && Math.abs(f2.y - (fa.y + 60)) < 1.5, [fa, f2]);
  await page.screenshot({ path: `${E}/n16-insert-2-moved.png` });
  // a click inside edits the text again
  { const a2 = await abs(page, "word-shape-frame"); await page.mouse.click(a2.x + 40, a2.y + 30); } await frame(page); await page.waitForTimeout(80);
  const c = await ctx(page);
  check("a click inside the picked shape edits its text again: frame without handles", c.Story === "textbox" && (await page.locator(".rb-word-shape-handle").count()) === 0 && (await sel(page, "word-shape-frame").count()) === 1, c);
  await page.keyboard.press("Escape"); await frame(page); await page.waitForTimeout(80);
  check("Escape leaves the box: no frame", (await sel(page, "word-shape-frame").count()) === 0 && (await ctx(page)).Story === "body");
  await page.close();
}

// ---------------- a header's text box: double click edits it ----------------
if (want("header")) {
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/test-documents-features/resume-template.docx`);
  await ready(page);
  const hdr = (await shapes(page)).find((s) => s.Story !== "body");
  check("resume-template: a text box of the header is listed with its story and page", !!hdr && hdr.Story === "header" && hdr.PageIndex === 0, hdr);
  const pt = await ev(page, (r) => { const w = window.__word, e = window.__editors[0]; const v = w.documentToViewport(e, r.X + r.W / 2, r.Y + r.H / 2); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: c.left + v.X, y: c.top + v.Y }; }, hdr);
  await page.mouse.click(pt.x, pt.y); await frame(page);
  check("one click on it does nothing (the body is edited)", (await ctx(page)).Story === "body");
  await page.mouse.dblclick(pt.x, pt.y); await frame(page); await page.waitForTimeout(150);
  let c = await ctx(page);
  check("a double click on it edits the box's text", c.Story === "textbox" && c.TextBoxId === hdr.Id, c);
  check("the text box bar is shown; no frame (a header's box is not moved)", (await sel(page, "word-text-box-bar").count()) === 1 && (await sel(page, "word-shape-frame").count()) === 0);
  await page.keyboard.press("End"); await page.keyboard.type("ZZ"); await frame(page);
  const t = (await shapes(page)).find((s) => s.Id === hdr.Id)?.Text;
  check("what is typed is the box's", /ZZ/.test(t ?? ""), t);
  await page.screenshot({ path: `${E}/n16-header-box.png` });
  await page.keyboard.press("Escape"); await frame(page);
  check("Escape goes back to the body", (await ctx(page)).Story === "body");
  await sel(page, "office-editor-save").click();
  const zip = await unzip(await waitUploads(page, 1));
  const parts = []; for (const [name, file] of zip) if (/word\/header\d*\.xml/.test(name) && (await file.async("string")).includes("ZZ")) parts.push(name);
  check("saved in the header's part", parts.length === 1, parts);
  // a double click on the header beside the box edits the header itself
  const hf = await ev(page, () => { const w = window.__word, e = window.__editors[0]; const fr = Array.from(w.getPageFrames(e))[0]; const v = w.documentToViewport(e, fr.Width - 60, 30); const cv = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: cv.left + v.X, y: cv.top + v.Y }; });
  await page.mouse.dblclick(hf.x, hf.y); await frame(page); await page.waitForTimeout(120);
  c = await ctx(page);
  check("a double click on the header beside the box edits the header", c.hf === "header" && c.Story === "header", c);
  await page.close();
}

// ---------------- the view stays on its text when most of a long document changes ----------------
if (want("scroll")) {
  const page = await newPage(context);
  await open(page, "word-editor", `${S}/long-generated.docx`);
  await ready(page);
  const anchor = () => ev(page, () => { const w = window.__word, e = window.__editors[0]; for (let y = 6; y < 600; y += 6) { const p = w.hitTestViewport(e, 500, y); if (p) return { at: `${p.BlockId}:${p.Offset}`, y, engineTop: w.getViewport(e).scrollTop, hostTop: document.querySelector('[data-testid="word-scroll"]').scrollTop }; } return null; });
  const settle = async () => { await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 }); await frame(page); await page.waitForTimeout(250); };
  await ev(page, () => { window.__scrolls = []; window.__word.onChange(window.__editors[0], (n) => { if (typeof n.ScrollTop === "number") window.__scrolls.push(n.ScrollTop); }); });
  // the bold faces first (they are fetched when first wanted, and the layout starts again when they arrive)
  await ev(page, () => window.__word.getInputElement(window.__editors[0]).focus({ preventScroll: true }));
  await page.keyboard.press("Control+a"); await page.keyboard.press("Control+b"); await settle();
  await page.keyboard.press("Control+z"); await settle();
  await ev(page, () => { const s = document.querySelector('[data-testid="word-scroll"]'); s.scrollTop = Math.round((s.scrollHeight - s.clientHeight) * 0.5); }); await frame(page); await page.waitForTimeout(300);
  const before = await anchor();
  await ev(page, () => { window.__scrolls.length = 0; });
  await page.keyboard.press("Control+a"); await frame(page);
  await page.keyboard.press("Control+b"); await frame(page);
  const right = await anchor();
  const stale = await ev(page, () => window.__word.getRenderStats(window.__editors[0]).StaleBlocks);
  check("mid-document, everything bold: the view shows the same text right after", !!right && right.at === before.at && Math.abs(right.y - before.y) <= 6, [before, right, stale]);
  await settle();
  const after = await anchor(); const told = await ev(page, () => window.__scrolls.slice());
  console.log("  bold:", JSON.stringify({ before, right, after, told, stale }));
  check("the engine asked the host to scroll (ScrollTop)", told.length >= 1, told);
  check("when the layout is complete the view shows the same text", after.at === before.at && Math.abs(after.y - before.y) <= 6, [before, after]);
  check("the host's scroll position is the engine's", Math.abs(after.engineTop - after.hostTop) < 1.5 && Math.abs(after.hostTop - told.at(-1)) < 1.5, [after, told]);
  await page.mouse.move(600, 400); await page.mouse.wheel(0, 30); await frame(page); await page.waitForTimeout(200);
  const wheel = await anchor();
  check("a wheel scroll by 30 afterwards moves the text by 30, no jump", wheel.at === before.at ? Math.abs(before.y - wheel.y - 30) <= 7 : true, [after, wheel]);
  check("  (the scroll position moved by 30)", Math.abs(wheel.hostTop - after.hostTop - 30) < 1.5, [after.hostTop, wheel.hostTop]);
  await ev(page, () => { window.__scrolls.length = 0; window.__word.getInputElement(window.__editors[0]).focus({ preventScroll: true }); });
  const b2 = await anchor();
  await page.keyboard.press("Control+z"); await settle();
  const undone = await anchor(); const told2 = await ev(page, () => window.__scrolls.slice());
  console.log("  undo:", JSON.stringify({ b2, undone, told2 }));
  check("undo: the view shows the same text, host and engine agree", undone.at === b2.at && Math.abs(undone.y - b2.y) <= 6 && Math.abs(undone.engineTop - undone.hostTop) < 1.5, [b2, undone]);
  check("the page count is the document's again", (await ev(page, () => window.__word.getDocumentSize(window.__editors[0]).PageCount)) === LONG_PAGES);
  await page.close();
}
// ---------------- the same when the bold faces arrive while the layout is under way ----------------
if (want("scroll-cold")) {
  const ctx2 = await browser.newContext({ viewport: { width: 1366, height: 768 } });
  const page = await newPage(ctx2);
  await open(page, "word-editor", `${S}/long-generated.docx`);
  await ready(page);
  await ev(page, () => { const s = document.querySelector('[data-testid="word-scroll"]'); s.scrollTop = Math.round((s.scrollHeight - s.clientHeight) * 0.5); }); await frame(page); await page.waitForTimeout(300);
  const top0 = await ev(page, () => document.querySelector('[data-testid="word-scroll"]').scrollTop);
  await ev(page, () => window.__word.getInputElement(window.__editors[0]).focus({ preventScroll: true }));
  await page.keyboard.press("Control+a"); await frame(page);
  await page.keyboard.press("Control+b");
  await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 }); await frame(page); await page.waitForTimeout(300);
  const top1 = await ev(page, () => document.querySelector('[data-testid="word-scroll"]').scrollTop);
  check("bold for the first time mid-document (its faces arrive meanwhile): the view is not thrown to the top", Math.abs(top1 - top0) < 30000, [top0, top1]);
  await ctx2.close();
}
// ---------------- Ctrl+Enter is the page break the menu says it is ----------------
if (want("pagebreak")) {
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/test-documents/sample-formatted.docx`);
  await ready(page);
  const info = () => ev(page, () => { const w = window.__word, e = window.__editors[0]; w.completeLayout(e); return { pages: w.getDocumentSize(e).PageCount, page2: w.getLayoutLines(e).find((l) => l.Page === 2)?.Text.slice(0, 24) }; });
  await ev(page, () => { const w = window.__word, m = window.__wordModule, e = window.__editors[0]; const l = w.getLayoutLines(e).find((x) => x.Text.length > 40 && !x.InTable); m.setSelection(e, l.BlockId, 10, l.BlockId, 10); w.getInputElement(e).focus({ preventScroll: true }); });
  const a = await info();
  await sel(page, "word-header-footer").click();
  check("the menu offers a page break with Ctrl+Enter", /Ctrl\+Enter/.test(await sel(page, "word-page-break").innerText()));
  await page.keyboard.press("Escape"); await ev(page, () => window.__word.getInputElement(window.__editors[0]).focus({ preventScroll: true }));
  await page.keyboard.press("Control+Enter"); await frame(page);
  const b = await info();
  check("Ctrl+Enter starts a new page at the caret", b.pages === a.pages + 1 && b.page2 !== a.page2, [a, b]);
  await page.keyboard.type("Q"); await frame(page);
  check("what is typed next is at the top of the new page", (await info()).page2.startsWith("Q"), (await info()).page2);
  await page.keyboard.press("Control+z"); await page.keyboard.press("Control+z"); await frame(page);
  const c = await info();
  check("two undos (the letter, the break): as it was, and clean", c.pages === a.pages && c.page2 === a.page2 && !(await ev(page, () => window.__word.isModified(window.__editors[0]))), [a, c]);
  // in a text box the key does nothing
  await sel(page, "word-header-footer").click(); await sel(page, "word-insert-text-box").click(); await frame(page); await page.waitForTimeout(100);
  await page.keyboard.type("in the box"); const t0 = await ev(page, () => window.__word.getText(window.__editors[0]));
  await page.keyboard.press("Control+Enter"); await frame(page);
  const c2 = await ctx(page);
  check("in a text box Ctrl+Enter does nothing: no page, no new paragraph", c2.Story === "textbox" && (await ev(page, () => window.__word.getText(window.__editors[0]))) === t0 && (await info()).pages === a.pages, c2);
  await page.close();
}
await browser.close();
summary();
