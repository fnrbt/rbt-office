// 0.1.18 in the host, Word: the engine selects, moves, resizes and deletes shapes itself (the host draws no frame),
// enters a header's text box on a double click, and makes Ctrl+Enter one page break.
import { launch, newPage, open, T, S, check, summary, waitUploads, unzip } from "./lib.mjs";
const { browser, context } = await launch({ viewport: { width: 1366, height: 768 } });
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const E = `${S}/explore`;
const ready = async (page) => { await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 }); await page.waitForFunction(() => window.__word.isLoadComplete(window.__editors[0]) !== false && window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 }); await page.waitForTimeout(500); };
const shapes = (page) => ev(page, () => Array.from(window.__word.getShapeRects(window.__editors[0])).map((r) => ({ Id: r.Id, Story: r.Story, PageIndex: r.PageIndex, X: Math.round(r.X * 10) / 10, Y: Math.round(r.Y * 10) / 10, W: Math.round(r.Width * 10) / 10, H: Math.round(r.Height * 10) / 10, Text: r.Text.slice(0, 40) })));
const ctx = (page) => ev(page, () => { const c = window.__word.getCaretContext(window.__editors[0]); return { Story: c.Story, TextBoxId: c.TextBoxId ?? null, hf: c.HeaderFooter ? c.HeaderFooter.Kind : null }; });
const selected = (page) => ev(page, () => window.__word.getSelectedShape(window.__editors[0]));
// a point of the page (layout pixels) on the screen
const onScreen = (page, x, y) => ev(page, ([x, y]) => { const w = window.__word, e = window.__editors[0]; const v = w.documentToViewport(e, x, y); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: c.left + v.X, y: c.top + v.Y }; }, [x, y]);
const handle = async (page, id, name) => { const h = await ev(page, ([id, name]) => Array.from(window.__word.getShapeHandles(window.__editors[0], id)).find((h) => h.Name === name), [id, name]); return h ? onScreen(page, h.X + h.Width / 2, h.Y + h.Height / 2) : null; };
const drag = async (page, from, dx, dy) => { await page.mouse.move(from.x, from.y); await page.mouse.down(); await page.mouse.move(from.x + dx / 2, from.y + dy / 2, { steps: 4 }); await page.mouse.move(from.x + dx, from.y + dy, { steps: 4 }); await frame(page); await page.mouse.up(); await frame(page); await page.waitForTimeout(80); };
const scrollTop = (page) => ev(page, () => document.querySelector('[data-testid="word-scroll"]').scrollTop);
const modified = (page) => ev(page, () => window.__word.isModified(window.__editors[0]));
const noOverlay = async (page) => (await page.locator(".rb-word-shape-layer, .rb-word-shape-frame, .rb-word-shape-handle").count()) === 0;
const only = process.argv[2]?.split(","); const want = (k) => !only || only.includes(k);

// ---------------- shapes of a file: selected by a click, moved, resized, deleted by the engine ----------------
if (want("shapes")) {
  const page = await newPage(context);
  await open(page, "word-editor", `${S}/synth/shapes.docx`);
  await ready(page);
  const all = await shapes(page);
  check("the fixture's shapes are of the body", all.length >= 3 && all.every((s) => s.Story === "body"), all.length);
  check("nothing is selected at first", (await selected(page)) === null);
  const target = all.find((s) => Math.abs(s.W - 140) < 2) ?? all[0];
  const mid = await onScreen(page, target.X + target.W / 2, target.Y + target.H / 2);
  await page.mouse.click(mid.x, mid.y); await frame(page); await page.waitForTimeout(80);
  let c = await ctx(page);
  check("a click in a shape edits its text and selects it", c.Story === "textbox" && c.TextBoxId === target.Id && (await selected(page)) === target.Id, [c, await selected(page)]);
  check("the host draws nothing over the canvas", await noOverlay(page));
  await page.screenshot({ path: `${E}/n18-shape-1-editing.png` });
  await page.keyboard.press("Escape"); await frame(page); await page.waitForTimeout(60);
  c = await ctx(page);
  check("Escape: back to the body, the box still selected, the editor open", c.Story === "body" && (await selected(page)) === target.Id && (await sel(page, "word-toolbar").count()) === 1, [c, await selected(page)]);
  await page.screenshot({ path: `${E}/n18-shape-2-selected.png` });
  // (The oval is turned: getShapeHandles gives its handles unturned, the engine paints and hits them turned.
  // Sizes are checked on the rectangle behind the text, which is not.)
  const rect = all.find((s) => s.W === 100 && s.H === 50);
  const onEdge = await onScreen(page, rect.X + rect.W / 4, rect.Y);
  await page.mouse.click(onEdge.x, onEdge.y); await frame(page); await page.waitForTimeout(60);
  const clicked = await selected(page);
  check("a click on the frame of the shape behind the text selects it (0.1.19)", clicked === rect.Id && (await ctx(page)).Story === "body", [clicked, await ctx(page)]);
  if (clicked !== rect.Id) await ev(page, (id) => window.__word.selectShape(window.__editors[0], id), rect.Id);
  await page.screenshot({ path: `${E}/n18-shape-3-rect.png` });
  const top0 = await scrollTop(page);
  const rev0 = await ev(page, () => window.__word.getRevision(window.__editors[0]));
  await drag(page, await handle(page, rect.Id, "se"), 30, 20);
  let now = (await shapes(page)).find((s) => s.Id === rect.Id);
  check("the south-east handle dragged by 30, 20: larger by that, its corner where it was", Math.abs(now.W - (rect.W + 30)) < 0.6 && Math.abs(now.H - (rect.H + 20)) < 0.6 && now.X === rect.X && now.Y === rect.Y, [rect, now]);
  check("the view did not move, the shape is still selected, the document is unsaved", (await scrollTop(page)) === top0 && (await selected(page)) === rect.Id && (await modified(page)), [top0, await scrollTop(page), await selected(page)]);
  check("the host's undo button follows", await sel(page, "word-undo").isEnabled());
  await page.screenshot({ path: `${E}/n18-shape-4-resized.png` });
  await page.keyboard.press("Control+z"); await frame(page);
  now = (await shapes(page)).find((s) => s.Id === rect.Id);
  check("one Ctrl+Z: as it was, and clean", now.W === rect.W && now.H === rect.H && !(await modified(page)) && (await ev(page, () => window.__word.getRevision(window.__editors[0]))) === rev0, now);
  await drag(page, await handle(page, rect.Id, "nw"), -20, -10);
  now = (await shapes(page)).find((s) => s.Id === rect.Id);
  check("the north-west handle dragged by -20, -10: its far corner stays", Math.abs(now.X - (rect.X - 20)) < 0.6 && Math.abs(now.Y - (rect.Y - 10)) < 0.6 && Math.abs(now.W - (rect.W + 20)) < 0.6 && Math.abs(now.H - (rect.H + 10)) < 0.6, now);
  await page.keyboard.press("Control+z"); await frame(page);
  now = (await shapes(page)).find((s) => s.Id === rect.Id);
  check("  and one Ctrl+Z takes it back", now.X === rect.X && now.Y === rect.Y && now.W === rect.W, now);
  await page.keyboard.press("ArrowRight"); await page.keyboard.press("Shift+ArrowDown"); await frame(page);
  now = (await shapes(page)).find((s) => s.Id === rect.Id);
  check("Right moves it by a pixel, Shift+Down by ten", Math.abs(now.X - (rect.X + 1)) < 0.6 && Math.abs(now.Y - (rect.Y + 10)) < 0.6 && now.W === rect.W, [rect, now]);
  check("  the view stays where it is", (await scrollTop(page)) === top0);
  await page.keyboard.press("Control+z"); await page.keyboard.press("Control+z"); await frame(page);
  await drag(page, await onScreen(page, rect.X + rect.W - 8, rect.Y + rect.H - 8), 60, 40);
  now = (await shapes(page)).find((s) => s.Id === rect.Id);
  check("dragged from inside: 60 to the right and 40 down, its size as it was", Math.abs(now.X - (rect.X + 60)) < 0.6 && Math.abs(now.Y - (rect.Y + 40)) < 0.6 && now.W === rect.W && now.H === rect.H, [rect, now]);
  check("the others are where they were", JSON.stringify((await shapes(page)).filter((s) => s.Id !== rect.Id)) === JSON.stringify(all.filter((s) => s.Id !== rect.Id)));
  check("  the view stays where it is", (await scrollTop(page)) === top0, [top0, await scrollTop(page)]);
  await page.screenshot({ path: `${E}/n18-shape-5-moved.png` });
  await page.keyboard.press("Control+z"); await frame(page);
  // the turned oval moved by its frame (the edge between two handles, turned with it)
  {
    await page.mouse.click(mid.x, mid.y); await frame(page); await page.keyboard.press("Escape"); await frame(page);
    const o0 = (await shapes(page)).find((s) => s.Id === target.Id);
    await drag(page, mid, 30, 15);
    const o1 = (await shapes(page)).find((s) => s.Id === target.Id);
    check("the turned oval, selected, dragged from its middle: moved by 30, 15", Math.abs(o1.X - (o0.X + 30)) < 0.6 && Math.abs(o1.Y - (o0.Y + 15)) < 0.6 && o1.W === o0.W, [o0, o1]);
    await page.keyboard.press("Control+z"); await frame(page);
    await ev(page, (id) => window.__word.selectShape(window.__editors[0], id), rect.Id);
  }
  // saved: a resize writes the new extent
  await drag(page, await handle(page, rect.Id, "se"), 20, 10);
  await sel(page, "office-editor-save").click();
  const doc = await (await unzip(await waitUploads(page, 1))).get("word/document.xml").async("string");
  check("saved: the extent of the anchor is the new size (120 x 60 px)", doc.includes('<wp:extent cx="1143000" cy="571500"/>'), (doc.match(/<wp:extent[^>]*>/g) ?? []).join(" "));
  await page.keyboard.press("Control+z"); await frame(page);
  const textLength = await ev(page, () => window.__word.getText(window.__editors[0]).length);
  await page.keyboard.press("Delete"); await frame(page); await page.waitForTimeout(80);
  check("Delete takes the selected shape out (one Delete, one shape, no text)", (await shapes(page)).length === all.length - 1 && (await selected(page)) === null && (await ev(page, () => window.__word.getText(window.__editors[0]).length)) === textLength, (await shapes(page)).length);
  await page.keyboard.press("Control+z"); await frame(page);
  check("undone: the shape is back", (await shapes(page)).length === all.length);
  // Escape lets go of a selected shape
  await page.mouse.click(mid.x, mid.y); await frame(page); await page.keyboard.press("Escape"); await frame(page);
  check("selected again (click, Escape)", (await selected(page)) === target.Id);
  await page.keyboard.press("Escape"); await frame(page); await page.waitForTimeout(60);
  check("Escape lets go of it and keeps the editor open", (await selected(page)) === null && (await sel(page, "word-toolbar").count()) === 1 && (await sel(page, "office-editor-confirm").count()) === 0);
  // a click on the page lets go
  await page.mouse.click(mid.x, mid.y); await frame(page); await page.keyboard.press("Escape"); await frame(page);
  const cv = await sel(page, "word-canvas").boundingBox();
  await page.mouse.click(cv.x + 420, cv.y + 60); await frame(page); await page.waitForTimeout(80);
  check("a click on the page lets go of it", (await selected(page)) === null && (await ctx(page)).Story === "body");
  // at 125 %
  await sel(page, "word-zoom-in").click(); await frame(page); await page.waitForTimeout(120);
  const b5 = (await shapes(page)).find((s) => s.Id === rect.Id);
  await ev(page, (id) => window.__word.selectShape(window.__editors[0], id), rect.Id);
  await drag(page, await handle(page, rect.Id, "se"), 50, 25);
  const z = (await shapes(page)).find((s) => s.Id === rect.Id);
  check("at 125 %: the handle dragged by 50, 25 on screen is 40, 20 on the page", Math.abs(z.W - (b5.W + 40)) < 0.6 && Math.abs(z.H - (b5.H + 20)) < 0.6, [b5, z]);
  await page.screenshot({ path: `${E}/n18-shape-6-zoom.png` });
  check("still nothing drawn by the host", await noOverlay(page));
  await page.close();
}

// ---------------- a text box from the menu: selected, typed in, moved ----------------
if (want("insert")) {
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/test-documents/sample-formatted.docx`);
  await ready(page);
  await ev(page, () => window.__word.getInputElement(window.__editors[0]).focus({ preventScroll: true }));
  await sel(page, "word-header-footer").click(); await sel(page, "word-insert-text-box").click(); await frame(page); await page.waitForTimeout(120);
  const s0 = (await shapes(page))[0];
  check("a text box from the menu: its text is edited, the box selected", (await ctx(page)).Story === "textbox" && (await selected(page)) === s0?.Id, [await ctx(page), await selected(page), s0]);
  await page.keyboard.type("Hello"); await frame(page);
  const s1 = (await shapes(page))[0];
  check("typing goes into the box (the text selected at first is replaced)", s1.Text === "Hello", s1.Text);
  await page.keyboard.press("Escape"); await frame(page);
  check("Escape: the body, the box selected", (await ctx(page)).Story === "body" && (await selected(page)) === s1.Id);
  await drag(page, await onScreen(page, s1.X + s1.W / 4, s1.Y + 1), -30, 60);
  const s2 = (await shapes(page))[0];
  check("moved by its frame", Math.abs(s2.X - (s1.X - 30)) < 0.6 && Math.abs(s2.Y - (s1.Y + 60)) < 0.6, [s1, s2]);
  await page.screenshot({ path: `${E}/n18-insert-moved.png` });
  await page.close();
}

// ---------------- a header's text box: a double click edits it ----------------
if (want("header")) {
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/test-documents-features/resume-template.docx`);
  await ready(page);
  const hdr = (await shapes(page)).find((s) => s.Story !== "body");
  check("resume-template: a text box of the header is listed", !!hdr && hdr.Story === "header", hdr);
  const pt = await onScreen(page, hdr.X + hdr.W / 2, hdr.Y + hdr.H / 2);
  await page.mouse.click(pt.x, pt.y); await frame(page);
  check("one click on it does nothing (the body is edited)", (await ctx(page)).Story === "body");
  await page.mouse.dblclick(pt.x, pt.y); await frame(page); await page.waitForTimeout(150);
  let c = await ctx(page);
  check("a double click on it edits the box's text", c.Story === "textbox" && c.TextBoxId === hdr.Id, c);
  await page.keyboard.press("End"); await page.keyboard.type("ZZ"); await frame(page);
  const t = (await shapes(page)).find((s) => s.Id === hdr.Id)?.Text;
  check("what is typed is the box's", /ZZ/.test(t ?? ""), t);
  await page.screenshot({ path: `${E}/n18-header-box.png` });
  await page.keyboard.press("Escape"); await frame(page); await page.keyboard.press("Escape"); await frame(page);
  c = await ctx(page);
  check("Escape, Escape: back in the body, the editor open", c.Story === "body" && (await sel(page, "word-toolbar").count()) === 1, c);
  await sel(page, "office-editor-save").click();
  const zip = await unzip(await waitUploads(page, 1));
  const parts = []; for (const [name, file] of zip) if (/word\/header\d*\.xml/.test(name) && (await file.async("string")).includes("ZZ")) parts.push(name);
  check("saved in the header's part", parts.length === 1, parts);
  await page.close();
}

// ---------------- Ctrl+Enter: one page break, one undo step ----------------
if (want("pagebreak")) {
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/test-documents/sample-formatted.docx`);
  await ready(page);
  const pages = () => ev(page, () => { const w = window.__word, e = window.__editors[0]; w.completeLayout(e); return w.getDocumentSize(e).PageCount; });
  await ev(page, () => { const w = window.__word, m = window.__wordModule, e = window.__editors[0]; const l = w.getLayoutLines(e).find((x) => x.Text.length > 40 && !x.InTable); m.setSelection(e, l.BlockId, 10, l.BlockId, 10); w.getInputElement(e).focus({ preventScroll: true }); });
  const p0 = await pages(); const t0 = await ev(page, () => window.__word.getText(window.__editors[0]));
  await page.keyboard.press("Control+Enter"); await frame(page);
  check("Ctrl+Enter: one page more", (await pages()) === p0 + 1, [p0, await pages()]);
  await page.keyboard.press("Control+z"); await frame(page);
  check("one Ctrl+Z: as it was, and clean", (await pages()) === p0 && (await ev(page, () => window.__word.getText(window.__editors[0]))) === t0 && !(await modified(page)));
  await page.close();
}
await browser.close();
summary();
