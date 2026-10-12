import { launch, newPage, open, T, SHOTS, check, summary, saved, unzip, listenerProbe, fontStats, waitUploads } from "./lib.mjs";
import fs from "node:fs";
import zlib from "node:zlib";
const S = "/tmp/redbat-adoption";
const { browser, context } = await launch();
await context.addInitScript(listenerProbe);
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const ready = async (page) => { await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 }); await page.waitForTimeout(400); };
const status = (page) => page.locator('[data-testid="office-editor-status"]').innerText();
const fmt = (page) => ev(page, () => window.__word.getSelectionFormat(window.__editors.at(-1)));
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const pages = (page) => ev(page, () => window.__word.getPageCount(window.__editors.at(-1)));
const focusDoc = (page) => ev(page, () => window.__word.getInputElement(window.__editors.at(-1)).focus());
// a small PNG (solid colour) for the image tests
function png(w, h) {
  const crc = (buf) => { let c, t = []; for (let n = 0; n < 256; n++) { c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } let x = 0xffffffff; for (const b of buf) x = t[(x ^ b) & 255] ^ (x >>> 8); return (x ^ 0xffffffff) >>> 0; };
  const chunk = (type, data) => { const len = Buffer.alloc(4); len.writeUInt32BE(data.length); const td = Buffer.concat([Buffer.from(type), data]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([len, td, c]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  const raw = Buffer.alloc((w * 3 + 1) * h); for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { const o = y * (w * 3 + 1) + 1 + x * 3; raw[o] = 200; raw[o + 1] = 30; raw[o + 2] = 30; }
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk("IHDR", ihdr), chunk("IDAT", zlib.deflateSync(raw)), chunk("IEND", Buffer.alloc(0))]);
}
fs.writeFileSync(`${S}/red.png`, png(200, 120));

// ---------------- editing: dirty state, toolbar, table, link, image, keys ----------------
{
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/test-documents/sample-formatted.docx`);
  await ready(page);
  check("clean after load", (await status(page)).includes("No changes"), await status(page));
  check("Save disabled while clean", await page.locator('[data-testid="office-editor-save"]').isDisabled());
  await focusDoc(page);

  // --- dirty follows the engine's saved revision
  await page.keyboard.type("x");
  check("typing makes it dirty", (await status(page)).includes("Unsaved"));
  await page.keyboard.press("Control+z");
  check("undo back to the loaded content is clean", (await status(page)).includes("No changes") && !(await ev(page, () => window.__word.isModified(window.__editors[0]))), await status(page));
  await page.keyboard.press("Control+y");
  check("redo is dirty again", (await status(page)).includes("Unsaved"));
  await page.keyboard.press("Control+s");
  await waitUploads(page, 1);
  await page.waitForTimeout(200);
  check("clean after save", (await status(page)).includes("Saved") && !(await ev(page, () => window.__word.isModified(window.__editors[0]))), await status(page));
  await page.keyboard.type("y");
  check("dirty after a change to the saved content", (await status(page)).includes("Unsaved"));
  await page.keyboard.press("Control+z");
  check("undo back to the saved content is clean", (await status(page)).includes("Saved"), await status(page));
  await page.keyboard.press("Control+z");
  check("undo past the saved content is dirty", (await status(page)).includes("Unsaved"), await status(page));
  await page.keyboard.press("Control+y");
  check("redo to the saved content is clean", (await status(page)).includes("Saved"), await status(page));

  // --- toolbar controls: each changes the document and shows the selection's state
  // put the caret in the first body paragraph and select a word
  const body = await ev(page, () => { const w = window.__word, e = window.__editors[0]; const lines = w.getLayoutLines(e); const l = lines.find((x) => (x.Text ?? x.text ?? "").includes("This document demonstrates")); return l ? { id: l.BlockId ?? l.blockId, keys: Object.keys(l) } : { keys: Object.keys(lines[0] ?? {}) }; });
  check("found the first body paragraph", !!body.id, body.keys);
  const select = (a, b) => ev(page, ([id, a, b]) => { const w = window.__word, e = window.__editors[0]; w.setSelection(e, id, a, id, b); }, [body.id, a, b]);
  await select(5, 13); // "document"
  await frame(page);
  let f = await fmt(page);
  check("selection reported", f.HasSelection && !f.IsCollapsed);
  const sel = (id) => page.locator(`[data-testid="${id}"]`);

  await sel("word-font-size").selectOption("18");
  f = await fmt(page);
  check("font size: set in points and shown", f.FontSizePt === 18 && Math.abs(f.FontSize - 24) < 0.01 && (await sel("word-font-size").inputValue()) === "18", { pt: f.FontSizePt, px: f.FontSize, box: await sel("word-font-size").inputValue() });

  const familyOptions = await sel("word-font-family").locator("option").allInnerTexts();
  check("font box lists the text stack's families", ["Arial", "Georgia", "Times New Roman", "Verdana"].every((n) => familyOptions.includes(n)), familyOptions.join("|"));
  fontStats.bytes = 0; fontStats.files = [];
  await sel("word-font-family").selectOption("Georgia");
  f = await fmt(page);
  check("font family: set and shown", f.FontFamily === "Georgia" && (await sel("word-font-family").inputValue()) === "Georgia", f.FontFamily);
  await page.waitForFunction(() => window.RedBatWordTextStack.loadedFamilyNames().includes("Georgia"), null, { timeout: 30000 });
  check("the picked family's face is fetched on demand (the one face the text needs)", fontStats.files.join() === "georgia.ttf", fontStats.files);

  await sel("word-color").click();
  await page.click('[data-testid="word-color-palette"] [data-color="#DC2626"]');
  f = await fmt(page);
  check("text colour: set and shown", (f.Color || "").toUpperCase() === "#DC2626" && (await sel("word-color-current").evaluate((el) => getComputedStyle(el).backgroundColor)) === "rgb(220, 38, 38)", f.Color);
  await sel("word-highlight").click();
  await page.click('[data-testid="word-highlight-palette"] [data-color="yellow"]');
  f = await fmt(page);
  check("highlight: set and shown", f.Highlight === "yellow" && (await sel("word-highlight-current").evaluate((el) => getComputedStyle(el).backgroundColor)) === "rgb(255, 255, 0)", f.Highlight);
  await sel("word-bold").click();
  f = await fmt(page);
  check("bold reflected", f.Bold === true && (await sel("word-bold").getAttribute("aria-pressed")) === "true");
  await page.screenshot({ path: `${SHOTS}/word-format.png` });
  await sel("word-clear-format").click();
  f = await fmt(page);
  check("clear formatting", f.Bold === false && !f.Highlight && (f.Color || "").toUpperCase() !== "#DC2626" && f.FontFamily !== "Georgia" && (await sel("word-bold").getAttribute("aria-pressed")) === "false", { b: f.Bold, h: f.Highlight, c: f.Color, fam: f.FontFamily, pt: f.FontSizePt });

  // paragraph style
  const styleOptions = await sel("word-style").locator("option").evaluateAll((os) => os.filter((o) => !o.hidden).map((o) => o.value + "=" + o.textContent));
  console.log("  style box:", styleOptions.join(" | "));
  const allStyles = await ev(page, () => window.__word.getStyles(window.__editors[0]).map((s) => s.Id));
  check("no internal style id in the style list or in the box (the host filters nothing)", !allStyles.some((id) => id.startsWith("__")) && !styleOptions.some((o) => o.includes("__")) && !String((await fmt(page)).Paragraph.StyleId ?? "").startsWith("__"), allStyles.filter((id) => id.startsWith("__")));
  const before = await fmt(page);
  await sel("word-style").selectOption("style:Heading2");
  f = await fmt(page);
  check("paragraph style: applied and shown", f.Paragraph.StyleId === "Heading2" && (await sel("word-style").inputValue()) === "style:Heading2", { was: before.Paragraph.StyleId, now: f.Paragraph.StyleId });
  await page.keyboard.press("Control+z");
  await frame(page);
  f = await fmt(page);
  check("style change is one undo step", f.Paragraph.StyleId === before.Paragraph.StyleId && (await sel("word-style").inputValue()) === (await ev(page, (id) => id ? "style:" + id : "style:Normal", before.Paragraph.StyleId)), { now: f.Paragraph.StyleId, box: await sel("word-style").inputValue() });

  // line and paragraph spacing: shown from the selection, applied, shown again
  const nextParaY = () => ev(page, () => { const w = window.__word, e = window.__editors[0]; const l = w.getLayoutLines(e).find((x) => (x.Text ?? "").includes("There is support for images")); return l ? Math.round((l.Y ?? l.y) * 10) / 10 : null; });
  const spacingShown = async () => { await sel("word-spacing").click(); const r = await ev(page, () => ({ before: document.querySelector('[data-testid="word-space-before"]').value, after: document.querySelector('[data-testid="word-space-after"]').value, pressed: [...document.querySelectorAll('[data-testid^="word-line-spacing-"][aria-pressed="true"]')].map((b) => b.textContent.trim()), note: document.querySelector('[data-testid="word-line-spacing-current"]')?.textContent ?? null, apply: document.querySelector('[data-testid="word-spacing-apply"]').disabled })); await page.keyboard.press("Escape"); return r; };
  const y0 = await nextParaY();
  f = await fmt(page);
  let shown = await spacingShown();
  console.log("  selection spacing:", JSON.stringify({ LineSpacing: f.Paragraph.LineSpacing, Rule: f.Paragraph.LineSpacingRule, Before: f.Paragraph.SpaceBefore, After: f.Paragraph.SpaceAfter }), "popover:", JSON.stringify(shown));
  check("spacing popover shows the paragraph's own spacing (line 1.15, before 0, after 0), Apply disabled until something is typed", shown.before === String(f.Paragraph.SpaceBefore) && shown.after === String(f.Paragraph.SpaceAfter) && shown.pressed.join() === "1.15" && f.Paragraph.LineSpacing === 1.15 && shown.apply === true, shown);
  await sel("word-spacing").click();
  await page.click('[data-testid="word-line-spacing-2"]');
  const y1 = await nextParaY();
  check("line spacing 2 pushes the next paragraph down", y1 > y0 + 20, { y0, y1 });
  shown = await spacingShown();
  check("line spacing 2 is shown as the pressed choice", shown.pressed.join() === "2" && (await fmt(page)).Paragraph.LineSpacing === 2, shown);
  await sel("word-spacing").click();
  await sel("word-space-before").fill("24");
  await sel("word-space-after").fill("36");
  await sel("word-spacing-apply").click();
  const y2 = await nextParaY();
  check("space after 36 pt pushes the next paragraph further", y2 > y1 + 30, { y1, y2 });
  check("keyboard back in the document after the spacing popover", await ev(page, () => document.activeElement === window.__word.getInputElement(window.__editors[0])));
  f = await fmt(page);
  shown = await spacingShown();
  check("the spacing fields show what the paragraph now has (24 / 36 pt)", shown.before === "24" && shown.after === "36" && f.Paragraph.SpaceBefore === 24 && f.Paragraph.SpaceAfter === 36, { shown, p: f.Paragraph });
  // only one field typed: the other keeps its value
  await sel("word-spacing").click();
  await sel("word-space-after").fill("12");
  await sel("word-spacing-apply").click();
  f = await fmt(page);
  check("typing one field leaves the other as it was", f.Paragraph.SpaceBefore === 24 && f.Paragraph.SpaceAfter === 12, f.Paragraph);
  // a selection over paragraphs with different spacing
  const nextId = await ev(page, () => { const w = window.__word, e = window.__editors[0]; return w.getLayoutLines(e).find((x) => (x.Text ?? "").includes("There is support for images")).BlockId; });
  await ev(page, ([a, b]) => window.__word.setSelection(window.__editors[0], a, 2, b, 5), [body.id, nextId]);
  f = await fmt(page);
  shown = await spacingShown();
  check("mixed spacing over two paragraphs: no value, no pressed choice, 'Mixed' said", f.Paragraph.SpaceAfter === "mixed" && shown.after === "" && shown.pressed.length === 0 && /Mixed/.test(shown.note ?? ""), { shown, p: f.Paragraph });
  await select(5, 13);
  await sel("word-spacing").click();
  await sel("word-spacing-reset").click();
  const y3 = await nextParaY();
  check("spacing reset goes back to the style's (the file's own direct spacing is gone too)", y3 < y1 && Math.abs(y3 - y0) < 20, { y0, y1, y3 });
  f = await fmt(page);
  shown = await spacingShown();
  check("after the reset the popover shows the style's spacing", shown.before === String(f.Paragraph.SpaceBefore) && shown.after === String(f.Paragraph.SpaceAfter), { shown, p: f.Paragraph });

  // page break
  const p0 = await pages(page);
  await select(20, 20);
  await sel("word-header-footer").click();
  await sel("word-page-break").click();
  const p1 = await pages(page);
  check("page break adds a page", p1 === p0 + 1, { p0, p1 });
  await frame(page);
  check("page count shown in the header", (await status(page)).includes(`${p1} pages`), await status(page));
  await page.keyboard.press("Control+z");
  check("page break undone", (await pages(page)) === p0);

  // hyperlink: insert with text, bubble under the caret, edit, remove
  await select(0, 0);
  await sel("word-link").click();
  check("link popover asks for text at a caret", (await sel("word-link-text").count()) === 1);
  await sel("word-link-target").fill("https://example.com/a");
  await sel("word-link-text").fill("LINK");
  await sel("word-link-apply").click();
  await select(2, 2);
  await frame(page);
  let link = await ev(page, () => window.__word.getHyperlink(window.__editors[0]));
  check("hyperlink inserted with its text", link && link.Target === "https://example.com/a" && link.Text === "LINK", link);
  f = await fmt(page);
  check("selection format reports the link", f.Hyperlink === "https://example.com/a");
  await page.waitForSelector('[data-testid="word-link-bubble"]', { timeout: 5000 });
  const bubble = await sel("word-link-bubble").boundingBox();
  const caret = await ev(page, () => { const w = window.__word, e = window.__editors[0]; const c = w.getCaretRect(e); const r = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: r.left + c.ViewportX, y: r.top + c.ViewportY + c.ViewportHeight }; });
  check("link bubble sits under the caret and shows the target", (await sel("word-link-bubble").innerText()).includes("https://example.com/a") && Math.abs(bubble.y - caret.y) < 12 && Math.abs(bubble.x - caret.x) < 12, { bubble, caret });
  await page.screenshot({ path: `${SHOTS}/word-link-bubble.png` });
  await sel("word-link-bubble-edit").click();
  check("Edit opens the popover with the target", (await sel("word-link-target").inputValue()) === "https://example.com/a" && (await sel("word-link-text").count()) === 0);
  await sel("word-link-target").fill("https://example.com/b");
  await sel("word-link-apply").click();
  link = await ev(page, () => window.__word.getHyperlink(window.__editors[0]));
  check("link target changed, text kept", link && link.Target === "https://example.com/b" && link.Text === "LINK", link);
  await select(40, 40);
  await frame(page);
  check("no bubble outside a link", (await sel("word-link-bubble").count()) === 0);
  await select(2, 2);
  await frame(page);
  await sel("word-link-bubble-remove").click();
  link = await ev(page, () => window.__word.getHyperlink(window.__editors[0]));
  const textAfter = await ev(page, () => window.__word.getText(window.__editors[0]).includes("LINKThis document"));
  check("Remove takes the link off and keeps the text", link == null && textAfter, { link, textAfter });
  // link on a selection
  await select(9, 17);
  await sel("word-link").click();
  check("link popover for a selection asks for no text", (await sel("word-link-text").count()) === 0);
  await sel("word-link-target").fill("https://example.com/sel");
  await page.keyboard.press("Enter");
  await select(11, 11);
  link = await ev(page, () => window.__word.getHyperlink(window.__editors[0]));
  check("selected text became the link", link && link.Target === "https://example.com/sel" && link.Text.length === 8, link);

  // --- caret keys (the engine's; the view follows the caret)
  const scrollTop = () => ev(page, () => document.querySelector('[data-testid="word-scroll"]').scrollTop);
  const caretY = () => ev(page, () => window.__word.getCaretRect(window.__editors[0])?.Y);
  const caretOnScreen = () => ev(page, () => { const c = window.__word.getCaretRect(window.__editors[0]); const h = document.querySelector('[data-testid="word-scroll"]').clientHeight; return c && c.ViewportY >= 0 && c.ViewportY + c.ViewportHeight <= h; });
  await focusDoc(page);
  await page.keyboard.press("Control+End");
  await frame(page); await frame(page);
  const endY = await caretY();
  const docH = await ev(page, () => window.__word.getDocumentSize(window.__editors[0]).Height);
  const lastLine = await ev(page, () => Math.max(...window.__word.getLayoutLines(window.__editors[0]).map((l) => l.Y)));
  check("Ctrl+End puts the caret on the last line of the text and scrolls there", endY >= lastLine - 100 && (await scrollTop()) > 1000 && (await caretOnScreen()), { endY, lastLine, docH, top: await scrollTop() });
  await page.keyboard.press("Control+Home");
  await frame(page); await frame(page);
  check("Ctrl+Home goes back to the start", (await caretY()) < 400 && (await scrollTop()) < 200 && (await caretOnScreen()), { y: await caretY(), top: await scrollTop() });
  const ya = await caretY();
  await page.keyboard.press("PageDown");
  await frame(page); await frame(page);
  const yb = await caretY();
  check("PageDown moves the caret about a screen and keeps it in view", yb > ya + 200 && (await caretOnScreen()), { ya, yb, top: await scrollTop() });
  await page.keyboard.press("PageUp");
  await frame(page); await frame(page);
  check("PageUp comes back", Math.abs((await caretY()) - ya) < 60 && (await caretOnScreen()), { now: await caretY(), ya });
  await select(30, 30);
  const x0 = await ev(page, () => window.__word.getCaretRect(window.__editors[0]).X);
  await page.keyboard.press("End");
  const xe = await ev(page, () => window.__word.getCaretRect(window.__editors[0]).X);
  await page.keyboard.press("Home");
  const xh = await ev(page, () => window.__word.getCaretRect(window.__editors[0]).X);
  check("Home / End move along the visual line", xe > x0 && xh < x0, { xh, x0, xe });
  const yl = await caretY();
  await page.keyboard.press("ArrowDown");
  check("Down moves one visual line", (await caretY()) > yl && (await caretY()) < yl + 60);
  await page.keyboard.press("ArrowUp");
  await page.keyboard.press("Control+Shift+ArrowRight");
  let s2 = await ev(page, () => window.__word.getSelection(window.__editors[0]));
  check("Ctrl+Shift+Right selects a word", !s2.IsCollapsed && s2.FocusOffset - s2.AnchorOffset >= 2, s2);
  // double / triple click
  const pt = await ev(page, ([id]) => { const w = window.__word, e = window.__editors[0]; w.setSelection(e, id, 45, id, 45); const c = w.getCaretRect(e); const r = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: r.left + c.ViewportX + 3, y: r.top + c.ViewportY + c.ViewportHeight / 2 }; }, [body.id]);
  await page.mouse.dblclick(pt.x, pt.y);
  s2 = await ev(page, () => window.__word.getSelection(window.__editors[0]));
  const wordLen = Math.abs(s2.FocusOffset - s2.AnchorOffset);
  check("double click selects a word", !s2.IsCollapsed && wordLen >= 2 && wordLen < 20, s2);
  await page.mouse.click(pt.x, pt.y, { clickCount: 3 });
  s2 = await ev(page, () => window.__word.getSelection(window.__editors[0]));
  check("triple click selects the paragraph", Math.abs(s2.FocusOffset - s2.AnchorOffset) > 100 || s2.AnchorBlockId !== s2.FocusBlockId, s2);
  await page.close();
}
summary();
await browser.close();
