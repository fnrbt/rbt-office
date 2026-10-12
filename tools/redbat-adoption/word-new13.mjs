// What @rbt/office 0.1.13 adds, in the Word editor and preview: text boxes edited in place, changes of a whole long document, a document closed while it is read.
import { launch, newPage, open, T, S, SHOTS, check, summary, waitUploads, unzip, saved } from "./lib.mjs";
import fs from "node:fs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const ctx = (page) => ev(page, () => window.__word.getCaretContext(window.__editors[0]));
const fmt = (page) => ev(page, () => { const f = window.__word.getSelectionFormat(window.__editors[0]); return { bold: f.Bold, italic: f.Italic, size: f.FontSizePt, align: f.Paragraph.Alignment, collapsed: f.IsCollapsed, color: f.Color }; });
const boxes = (page) => ev(page, () => window.__word.getShapeRects(window.__editors[0]));
const status = (page) => sel(page, "office-editor-status").innerText();
const boxPoint = (page, r) => ev(page, (r) => { const w = window.__word, e = window.__editors[0]; const v = w.documentToViewport(e, r.X + r.Width / 2, r.Y + r.Height / 2); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: c.left + v.X, y: c.top + v.Y }; }, r);

// ---- a text box: entered by a click, the bar, formatting, Escape, save
{
  const file = `${T}/test-documents-wild/docx4j-drawingml-wps.docx`;
  const page = await newPage(context);
  await open(page, "word-editor", file);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await page.waitForTimeout(600);
  const list = await boxes(page);
  const box = list.find((r) => !r.BehindText && r.Text);
  check("getShapeRects lists the document's text box", !!box && box.Width > 20 && box.Height > 20 && typeof box.AnchorBlockId === "string", list.map((r) => ({ id: r.Id, text: r.Text, wrap: r.Wrap, behind: r.BehindText })));
  check("in the body there is no text box bar", (await sel(page, "word-text-box-bar").count()) === 0 && (await ctx(page)).Story === "body");
  const p = await boxPoint(page, box);
  await page.mouse.click(p.x, p.y); await frame(page);
  let c = await ctx(page);
  check("a click in the box edits its text: the engine says so and the bar shows it", c.Story === "textbox" && c.TextBoxId === box.Id && c.InHeaderFooter === false && (await sel(page, "word-text-box-bar").count()) === 1 && (await sel(page, "word-header-footer-bar").count()) === 0 && (await sel(page, "word-text-box-name").innerText()) === "Text box", c);
  check("the keyboard is in the document", await ev(page, () => document.activeElement === window.__word.getInputElement(window.__editors[0])));
  await page.keyboard.type("ZZ"); await frame(page);
  check("typing goes into the box, and the document is dirty", (await boxes(page))[0].Text.includes("ZZ") && (await status(page)).includes("Unsaved"), (await boxes(page))[0].Text);
  // formatting commands of the toolbar, on the box's text
  await page.keyboard.press("Control+a"); await frame(page);
  let f = await fmt(page);
  check("Ctrl+A selects the box's text (the caret stays in the box)", !f.collapsed && (await ctx(page)).Story === "textbox", f);
  await sel(page, "word-bold").click(); await frame(page);
  await sel(page, "word-italic").click(); await frame(page);
  await sel(page, "word-align-right").click(); await frame(page);
  await sel(page, "word-font-size").selectOption("14"); await frame(page);
  f = await fmt(page); c = await ctx(page);
  check("Bold, Italic, Align right and the size box work on it, and the toolbar shows them", f.bold === true && f.italic === true && f.align === "right" && f.size === 14 && c.Story === "textbox" && (await sel(page, "word-bold").getAttribute("aria-pressed")) === "true" && (await sel(page, "word-align-right").getAttribute("aria-pressed")) === "true" && (await sel(page, "word-font-size").inputValue()) === "14", f);
  await page.screenshot({ path: `${SHOTS}/word-text-box-bar.png` });
  await sel(page, "word-undo").click(); await frame(page);
  f = await fmt(page);
  check("undo takes the last of them back, still in the box", f.size !== 14 && f.bold === true && (await ctx(page)).Story === "textbox", f);
  await sel(page, "word-redo").click(); await frame(page);
  // the insert menu: no page break in a box
  await sel(page, "word-header-footer").click();
  check("the menu offers no page break in a text box", await sel(page, "word-page-break").isDisabled());
  await page.keyboard.press("Escape"); await frame(page);
  check("Escape closes the menu first", (await sel(page, "word-header-footer-menu").count()) === 0 && (await ctx(page)).Story === "textbox");
  await page.keyboard.press("Escape"); await frame(page);
  check("Escape leaves the text box and not the editor", (await ctx(page)).Story === "body" && (await sel(page, "word-text-box-bar").count()) === 0 && (await sel(page, "office-editor").count()) === 1 && (await sel(page, "office-editor-confirm").count()) === 0);
  await sel(page, "word-header-footer").click();
  check("in the body the page break is offered again", !(await sel(page, "word-page-break").isDisabled()));
  await page.keyboard.press("Escape"); await frame(page);
  // in again through the API, out through the bar's button
  check("enterTextBox goes in", (await ev(page, (id) => window.__word.enterTextBox(window.__editors[0], id), box.Id)) === true);
  await frame(page);
  check("... and the bar follows a change the host did not make", (await sel(page, "word-text-box-bar").count()) === 1);
  await sel(page, "word-text-box-close").click(); await frame(page);
  check("Close text box goes back to the body, the keyboard in the document", (await ctx(page)).Story === "body" && (await sel(page, "word-text-box-bar").count()) === 0 && (await ev(page, () => document.activeElement === window.__word.getInputElement(window.__editors[0]))));
  // a click outside the box leaves it
  await page.mouse.click(p.x, p.y); await frame(page);
  const inAgain = (await ctx(page)).Story;
  const frames0 = await ev(page, () => { const w = window.__word, e = window.__editors[0]; const fr = w.getPageFrames(e)[0]; const v = w.documentToViewport(e, fr.X + fr.Width / 2, fr.Top + 130); const c = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: c.left + v.X, y: c.top + v.Y }; });
  await page.mouse.click(frames0.x, frames0.y); await frame(page);
  check("a click outside the box goes back to the body, and the bar goes", inAgain === "textbox" && (await ctx(page)).Story === "body" && (await sel(page, "word-text-box-bar").count()) === 0);
  // save and read again
  await sel(page, "office-editor-save").click();
  const bytes = await waitUploads(page, 1);
  fs.writeFileSync(`${S}/out/textbox13-docx4j-drawingml-wps.docx`, bytes);
  const report = await ev(page, async () => (await window.__formats()).default().word.getLastSaveReport());
  console.log("  save report:", JSON.stringify(report));
  check("saved; getLastSaveReport tells what the save made of the body", (await status(page)).includes("Saved") && report && typeof report.PatchedParagraphs === "number" && Array.isArray(report.Reasons), report);
  saved.set("tb13.docx", bytes);
  const again = await newPage(context);
  await open(again, "word-editor", "/api/saved/tb13.docx", { name: "docx4j-drawingml-wps.docx" });
  await again.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await again.waitForTimeout(500);
  const back = await boxes(again);
  await ev(again, (id) => { const w = window.__word, e = window.__editors[0]; w.enterTextBox(e, id); }, back[0]?.Id);
  await again.keyboard.press("Control+a"); await frame(again);
  const f2 = await fmt(again);
  check("read again: the box has the text and its formatting", back.length === list.length && back[0].Text.includes("ZZ") && f2.bold === true && f2.italic === true && f2.size === 14 && f2.align === "right", { text: back[0]?.Text, f2 });
  await again.close(); await page.close();
}

// ---- a change of a whole long document: the view stays, the layout follows in slices
{
  const file = `${S}/long-generated.docx`;
  const page = await newPage(context);
  await open(page, "word-editor", file);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await page.waitForFunction(() => window.__word.isLoadComplete(window.__editors[0]) !== false, null, { timeout: 300000 });
  await ev(page, () => window.__word.whenLayoutComplete(window.__editors[0]));
  await page.waitForTimeout(500);
  const scrollTop = () => ev(page, () => Math.round(document.querySelector('[data-testid="word-scroll"]').scrollTop));
  const subtitle = () => ev(page, () => document.querySelector('[data-testid="office-editor-status"]').parentElement.innerText.replace(/\n/g, " | "));
  const layout = () => ev(page, () => { const w = window.__word, e = window.__editors[0]; return { complete: w.isLayoutComplete(e), pages: w.getDocumentSize(e).PageCount }; });
  const pages0 = (await layout()).pages;
  await ev(page, () => window.__word.getInputElement(window.__editors[0]).focus());
  await page.keyboard.press("Control+a"); await frame(page); await page.waitForTimeout(100);
  check("Ctrl+A selects everything and the view stays at the top", (await scrollTop()) === 0 && (await fmt(page)).collapsed === false, await scrollTop());
  for (const how of ["key", "button"]) {
    let t = Date.now();
    if (how === "key") await page.keyboard.press("Control+b"); else await sel(page, "word-bold").click();
    await frame(page);
    const ms = Date.now() - t;
    const after = await layout();
    const sub = await subtitle();
    check(`bold on all ${pages0} pages (${how}): the host leaves the rest of the layout to the engine's slices (isLayoutComplete false after the change and its frame)`, after.complete === false && after.pages < pages0, { ms, ...after });
    check(`... the page count says there are more to come, the view has not moved, the toolbar shows bold`, /\d+\+ pages…/.test(sub) && (await scrollTop()) === 0 && (await sel(page, "word-bold").getAttribute("aria-pressed")) === "true", { sub, top: await scrollTop() });
    console.log(`  bold (${how}) returned in ${ms} ms with ${after.pages} pages laid out`);
    if (how === "key") await page.screenshot({ path: `${SHOTS}/word-mass-bold.png` });
    t = Date.now();
    await ev(page, () => window.__word.whenLayoutComplete(window.__editors[0])); await frame(page);
    const done = await layout();
    check(`... and when the slices are through the count is the whole document's`, done.complete && done.pages > pages0 && new RegExp(`\\| ${done.pages} pages$`).test(await subtitle()) && (await scrollTop()) === 0, { ms: Date.now() - t, ...done, sub: await subtitle() });
    t = Date.now();
    if (how === "key") await page.keyboard.press("Control+z"); else await sel(page, "word-undo").click();
    await frame(page);
    const undone = await layout();
    check(`undo of it (${how}): in slices too, the view where it was`, undone.complete === false && (await scrollTop()) === 0, { ms: Date.now() - t, ...undone });
    await ev(page, () => window.__word.whenLayoutComplete(window.__editors[0])); await frame(page);
    check(`... to the ${pages0} pages there were, unmodified`, (await layout()).pages === pages0 && /No changes/.test(await status(page)), { ...(await layout()), status: await status(page) });
  }
  // the view still follows what the keyboard does
  await page.keyboard.press("ArrowLeft"); await frame(page);
  await page.keyboard.press("Control+End"); await frame(page); await page.waitForTimeout(150);
  const atEnd = await scrollTop();
  check("Ctrl+End still takes the view to the caret", atEnd > 100000, atEnd);
  await page.keyboard.press("Control+Home"); await frame(page); await page.waitForTimeout(150);
  check("Ctrl+Home back to the top", (await scrollTop()) < 200, await scrollTop());
  for (let i = 0; i < 40; i++) await page.keyboard.press("Shift+ArrowDown");
  await frame(page); await page.waitForTimeout(150);
  check("a selection extended with the keyboard is followed", (await scrollTop()) > 300, await scrollTop());
  const top1 = await scrollTop();
  await sel(page, "word-italic").click(); await frame(page); await page.waitForTimeout(100);
  check("formatting a range leaves the view where it is", (await scrollTop()) === top1 && (await fmt(page)).italic === true);
  await page.keyboard.press("Control+Home"); await frame(page);
  await page.keyboard.press("Control+a"); await page.keyboard.type("x"); await frame(page); await page.waitForTimeout(100);
  check("typing over everything: the view follows the caret (the document is one line)", (await scrollTop()) < 200 && (await fmt(page)).collapsed === true && (await layout()).pages === 1, { top: await scrollTop(), ...(await layout()) });
  await page.close();
}

// ---- closed while the document is read: nobody is told of an error
for (const [component, when] of [["word-editor", "shown"], ["word-editor", "early"], ["word-preview", "shown"], ["word-preview", "early"]]) {
  const page = await newPage(context);
  const errors = [];
  page.on("console", (m) => { if (m.type() === "error" && !/404/.test(m.text())) errors.push(m.text().slice(0, 200)); });
  page.on("pageerror", (e) => errors.push("pageerror: " + String(e).slice(0, 200)));
  await ev(page, () => 0).catch(() => {});
  await page.addInitScript(() => { window.addEventListener("unhandledrejection", (e) => console.error("unhandled rejection: " + String(e.reason))); });
  await open(page, component, `${S}/long-generated.docx`, component === "word-preview" ? { params: { hookWord: "1" } } : {});
  let at;
  if (when === "shown") {
    if (component === "word-editor") await page.waitForFunction(() => window.__editors.length > 0 && document.querySelector('[data-testid="word-toolbar"]'), null, { timeout: 300000 });
    else await page.waitForFunction(() => window.__wordLoaded != null, null, { timeout: 300000 });
    at = await ev(page, (component) => { const s = component === "word-editor" ? { load: window.__word.isLoadComplete(window.__editors[0]) } : { load: window.__wordLoaded.viewer.isLoadComplete() }; window.__setOpen(false); return s; }, component);
  } else {
    // as soon as the engine exists, or after a moment when it does not come that fast
    await page.waitForFunction(() => window.__editors.length > 0, null, { timeout: 1500 }).catch(() => {});
    at = await ev(page, () => { window.__setOpen(false); return { load: "n/a" }; });
  }
  await page.waitForTimeout(2500);
  const state = await ev(page, (component) => { const w = window.__word, e = window.__editors[0]; const v = window.__wordLoaded?.viewer; return { editor: e && component === "word-editor" ? { complete: w.isLoadComplete(e), failed: w.isLoadFailed(e), error: w.getLoadError(e) } : null, viewer: v ? { complete: v.isLoadComplete(), failed: v.isLoadFailed(), error: v.getLoadError(), told: window.__wordLoaded.loadFailed } : null, alerts: [...document.querySelectorAll('[role="alert"], .rb-office-notice, .rb-file-viewer-error-copy')].map((a) => a.textContent), body: document.body.innerText.slice(0, 100) }; }, component);
  check(`${component} closed ${when === "shown" ? "while the rest is read" : "before it is shown"}: no error in the console, none on the page`, errors.length === 0 && state.alerts.length === 0 && !/cancel|could not/i.test(state.body), { at, errors, state });
  if (when === "shown" && component === "word-preview") check("  (the engine does say the reading was cancelled; the host's document does not pass it on as a failure)", at.load === false ? state.viewer.failed === true && /cancel/i.test(state.viewer.error ?? "") && state.viewer.told === false : true, { at, viewer: state.viewer });
  if (when === "shown" && component === "word-editor") console.log("  editor after dispose:", JSON.stringify(state.editor), "at close:", JSON.stringify(at));
  await page.close();
}
await browser.close();
process.exit(summary() ? 1 : 0);
