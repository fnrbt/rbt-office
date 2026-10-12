// What @rbt/office 0.1.14 changes, in the Word editor: a format chosen at a caret, spacing in points, the text box command, page break refused, sections.
import { launch, newPage, open, T, S, SHOTS, check, summary, waitUploads, unzip } from "./lib.mjs";
import fs from "node:fs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const ctx = (page) => ev(page, () => window.__word.getCaretContext(window.__editors[0]));
const fmt = (page) => ev(page, () => { const f = window.__word.getSelectionFormat(window.__editors[0]); return { bold: f.Bold, italic: f.Italic, size: f.FontSizePt, collapsed: f.IsCollapsed, before: f.Paragraph.SpaceBefore, after: f.Paragraph.SpaceAfter }; });
const boxes = (page) => ev(page, () => window.__word.getShapeRects(window.__editors[0]));
const status = (page) => sel(page, "office-editor-status").innerText();
const pressed = (page, id) => sel(page, id).getAttribute("aria-pressed");
const ready = async (page) => { await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 }); await page.waitForFunction(() => window.__word.isLoadComplete(window.__editors.at(-1)) !== false && window.__word.isLayoutComplete(window.__editors.at(-1)), null, { timeout: 300000 }); await page.waitForTimeout(400); };
const caretAt = (page, text, offset = 3) => ev(page, ([text, offset]) => { const w = window.__word, m = window.__wordModule, e = window.__editors[0]; const found = w.findText(e, text, true); if (!found.length) return null; w.selectFindMatch(e, found[0]); const s = m.getSelection(e); const o = Math.min(s.AnchorOffset, s.FocusOffset) + offset; const ok = m.setSelection(e, s.AnchorBlockId, o, s.AnchorBlockId, o); w.getInputElement(e).focus({ preventScroll: true }); return { id: s.AnchorBlockId, o, ok }; }, [text, offset]);
const paraText = (page, id) => ev(page, (id) => { const m = window.__wordModule, e = window.__editors[0]; return (m.getLayoutFingerprint(e).find((b) => b.Id === id) ?? {}).Text ?? null; }, id);

// ---- a character format chosen at a caret
{
  const file = `${T}/test-documents/sample-formatted.docx`;
  const page = await newPage(context);
  await open(page, "word-editor", file); await ready(page);
  const notices = await ev(page, () => { window.__notices = []; window.__word.onChange(window.__editors[0], (c) => window.__notices.push({ d: c.DocumentChanged, s: c.SelectionChanged, f: c.FormatChanged, m: c.Modified, r: c.Revision })); return true; });
  const fp = await ev(page, () => window.__wordModule.getLayoutFingerprint(window.__editors[0]).filter((b) => (b.Text ?? "").length > 30).slice(0, 3).map((b) => ({ text: b.Text.slice(0, 30), runs: b.Runs })));
  console.log("  paragraphs", JSON.stringify(fp).slice(0, 500));
  const at = await caretAt(page, fp[0].text.slice(0, 20), 4); await frame(page);
  const f0 = await fmt(page);
  check("a caret in plain text: Bold is off in the engine and in the toolbar", at && at.ok === true && f0.collapsed === true && f0.bold === false && (await pressed(page, "word-bold")) === "false", { at, f0 });
  await ev(page, () => (window.__notices.length = 0));
  const scroll0 = await ev(page, () => document.querySelector('[data-testid="word-scroll"]').scrollTop);
  await sel(page, "word-bold").click(); await frame(page);
  const f1 = await fmt(page); const n1 = await ev(page, () => window.__notices.slice());
  check("Bold at the caret: the engine says it of what is typed next, the toolbar button is pressed", f1.bold === true && f1.collapsed === true && (await pressed(page, "word-bold")) === "true", f1);
  check("the notice is FormatChanged with SelectionChanged, not DocumentChanged, not Modified", n1.length >= 1 && n1.every((n) => n.d === false && n.m === false) && n1.some((n) => n.f === true && n.s === true), n1);
  check("the document is not dirty, nothing was uploaded, undo is not offered for it", /No changes/.test(await status(page)) && page.uploads.length === 0 && (await sel(page, "word-undo").isDisabled()), await status(page));
  await page.keyboard.type("QQ"); await frame(page);
  const blk = await ev(page, (t) => { const b = window.__wordModule.getLayoutFingerprint(window.__editors[0]).find((b) => (b.Text ?? "").startsWith(t)); return b ? { text: b.Text.slice(0, 20), runs: b.Runs } : null; }, fp[0].text.slice(0, 4) + "QQ");
  const parts = (blk?.runs ?? "").split(" | ");
  check("what is typed is a run of its own (bold), between two that are not", !!blk && parts.length === 3 && parts[1].startsWith("2*") && parts[1] !== parts[0].replace(/^\d+/, "2"), blk);
  check("typing made it dirty", /Unsaved/.test(await status(page)));
  // keyboard: Ctrl+I at a caret, the toolbar follows a frame later
  await page.keyboard.press("Control+i"); await frame(page); await frame(page);
  check("Ctrl+I at the caret: the Italic button is pressed without anything typed", (await pressed(page, "word-italic")) === "true" && (await fmt(page)).italic === true);
  await page.keyboard.press("ArrowLeft"); await frame(page); await frame(page);
  await page.keyboard.press("ArrowLeft"); await frame(page); await frame(page);
  await page.keyboard.press("ArrowLeft"); await frame(page); await frame(page);
  check("moving the caret drops it: the toolbar shows the text's own format again", (await pressed(page, "word-italic")) === "false" && (await fmt(page)).italic === false, await fmt(page));
  // the size box at a caret
  await sel(page, "word-font-size").selectOption("20"); await frame(page);
  check("a size chosen at the caret is shown in the size box", (await sel(page, "word-font-size").inputValue()) === "20" && (await fmt(page)).size === 20, await fmt(page));
  await page.keyboard.type("Z"); await frame(page);
  // save: the typed text is bold in the file
  await sel(page, "office-editor-save").click();
  const body = await waitUploads(page, 1);
  const doc = await (await unzip(body)).get("word/document.xml").async("string");
  const boldQ = (doc.match(/<w:r><w:rPr><w:b\/>(?:(?!<\/w:r>)[\s\S])*?<w:t[^>]*>Q+<\/w:t><\/w:r>/g) ?? []);
  const run = boldQ.join("");
  const runZ = /<w:r[ >](?:(?!<\/w:r>)[\s\S])*?<w:t[^>]*>Z<\/w:t><\/w:r>/.exec(doc)?.[0] ?? "";
  check("saved: QQ is bold, Z a 20 pt run", boldQ.map((r) => /Q+/.exec(r.replace(/<[^>]*>/g, ""))[0]).join("") === "QQ" && /w:sz w:val="40"/.test(runZ), { run: run.slice(-260), runZ: runZ.slice(-260) });
  // a toolbar click with the caret out of view leaves the view where it is
  await ev(page, () => { const sc = document.querySelector('[data-testid="word-scroll"]'); sc.scrollTop = sc.scrollHeight; sc.dispatchEvent(new Event("scroll")); });
  await frame(page);
  const s1 = await ev(page, () => document.querySelector('[data-testid="word-scroll"]').scrollTop);
  await sel(page, "word-underline").click(); await frame(page); await frame(page);
  const s2 = await ev(page, () => document.querySelector('[data-testid="word-scroll"]').scrollTop);
  check("a format chosen at a caret that is out of view does not move the view", s1 > 0 ? s2 === s1 : true, { s1, s2 });
  await page.close();
}

// ---- paragraph spacing is in points
{
  const file = `${T}/test-documents-wild/docx4j-getting-started.docx`;
  const zip = await unzip(fs.readFileSync(file));
  const xml = await zip.get("word/document.xml").async("string");
  const paras = (xml.slice(xml.indexOf("<w:body")).match(/<w:p[ >](?:(?!<w:p[ >])[\s\S])*?<\/w:p>/g) ?? []);
  const withSpacing = paras.map((p) => ({ p, before: /<w:pPr>(?:(?!<\/w:pPr>)[\s\S])*?<w:spacing [^>]*w:before="(\d+)"/.exec(p)?.[1], after: /<w:pPr>(?:(?!<\/w:pPr>)[\s\S])*?<w:spacing [^>]*w:after="(\d+)"/.exec(p)?.[1], text: (p.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) ?? []).map((t) => t.replace(/<[^>]*>/g, "")).join("") })).filter((x) => (x.before || x.after) && x.text.length > 25 && !/[&]/.test(x.text));
  console.log("  paragraphs that state spacing:", withSpacing.length, withSpacing.slice(0, 3).map((x) => [x.before, x.after, x.text.slice(0, 30)]));
  const page = await newPage(context);
  await open(page, "word-editor", file); await ready(page);
  let done = false;
  for (const target of withSpacing.slice(0, 8)) {
    const at = await caretAt(page, target.text.slice(0, 25), 2); await frame(page);
    if (!at) continue;
    const f = await fmt(page);
    const wantBefore = target.before ? Number(target.before) / 20 : null, wantAfter = target.after ? Number(target.after) / 20 : null;
    check("getSelectionFormat gives the file's points (w:before / 20, w:after / 20)", (wantBefore === null || Math.abs(f.before - wantBefore) < 0.01) && (wantAfter === null || Math.abs(f.after - wantAfter) < 0.01), { file: [target.before, target.after], engine: [f.before, f.after] });
    await sel(page, "word-spacing").click(); await frame(page);
    const shown = [await sel(page, "word-space-before").inputValue(), await sel(page, "word-space-after").inputValue()];
    check("the spacing popover shows those points", (wantBefore === null || Number(shown[0]) === wantBefore) && (wantAfter === null || Number(shown[1]) === wantAfter), shown);
    await page.screenshot({ path: `${SHOTS}/word14-spacing.png` });
    await sel(page, "word-space-before").fill("10"); await sel(page, "word-space-after").fill("7.5");
    await sel(page, "word-spacing-apply").click(); await frame(page);
    const f2 = await fmt(page);
    check("10 and 7.5 typed in the popover are 10 and 7.5 in the engine", f2.before === 10 && f2.after === 7.5, f2);
    await sel(page, "office-editor-save").click();
    const body = await waitUploads(page, 1);
    const doc = await (await unzip(body)).get("word/document.xml").async("string");
    const p = (doc.slice(doc.indexOf("<w:body")).match(/<w:p[ >](?:(?!<w:p[ >])[\s\S])*?<\/w:p>/g) ?? []).find((p) => p.replace(/<[^>]*>/g, "").startsWith(target.text.slice(0, 25)));
    const sp = /<w:spacing [^>]*>/.exec(p ?? "")?.[0];
    check("saved: w:before=200 and w:after=150", /w:before="200"/.test(sp ?? "") && /w:after="150"/.test(sp ?? ""), sp);
    done = true; break;
  }
  check("a paragraph with stated spacing was found and tried", done);
  await page.close();
}

// ---- the text box command; page break and text box refused where the engine refuses them
{
  const file = `${T}/test-documents/sample-formatted.docx`;
  const page = await newPage(context);
  await open(page, "word-editor", file); await ready(page);
  const fp = await ev(page, () => window.__wordModule.getLayoutFingerprint(window.__editors[0]).filter((b) => (b.Text ?? "").length > 30).slice(0, 3).map((b) => ({ id: b.Id, text: b.Text.slice(0, 30) })));
  const at = await caretAt(page, fp[1].text.slice(0, 20), 4); await frame(page);
  const n0 = (await boxes(page)).length;
  await sel(page, "word-header-footer").click();
  check("the insert menu offers a text box in the body", (await sel(page, "word-insert-text-box").count()) === 1 && !(await sel(page, "word-insert-text-box").isDisabled()));
  await page.screenshot({ path: `${SHOTS}/word14-insert-menu.png` });
  await sel(page, "word-insert-text-box").click(); await frame(page); await frame(page);
  let list = await boxes(page); let c = await ctx(page);
  const made = list.find((b) => b.Text === "Text box");
  check("Text box: one more shape, anchored in the caret's paragraph", list.length === n0 + 1 && !!made && made.AnchorBlockId === at.id, list.map((b) => ({ id: b.Id, text: b.Text, anchor: b.AnchorBlockId, x: Math.round(b.X), y: Math.round(b.Y), w: b.Width, h: b.Height })));
  check("its text is being edited (the text box bar is shown) with all of it selected", c.Story === "textbox" && c.TextBoxId === made?.Id && (await sel(page, "word-text-box-bar").count()) === 1 && (await fmt(page)).collapsed === false, c);
  check("the keyboard is in the document", await ev(page, () => document.activeElement === window.__word.getInputElement(window.__editors[0])));
  await page.keyboard.type("Note to self"); await frame(page);
  list = await boxes(page);
  check("typing replaces the placeholder", list.find((b) => b.Id === made.Id)?.Text === "Note to self", list.map((b) => b.Text));
  await page.screenshot({ path: `${SHOTS}/word14-text-box-made.png` });
  await sel(page, "word-header-footer").click();
  check("in a text box neither a text box nor a page break is offered", (await sel(page, "word-insert-text-box").isDisabled()) && (await sel(page, "word-page-break").isDisabled()));
  await page.keyboard.press("Escape"); await frame(page);
  const refused = await ev(page, () => { const w = window.__word, e = window.__editors[0]; const r0 = w.getShapeRects(e).length; const a = w.insertTextBox(e, "x"); const b = w.insertPageBreak(e); return { a, b, same: w.getShapeRects(e).length === r0 }; });
  check("the engine itself refuses both in a text box (false, nothing changes)", refused.a === false && refused.b === false && refused.same, refused);
  await page.keyboard.press("Control+Enter"); await frame(page);
  check("Ctrl+Enter in a text box starts no page", (await ctx(page)).Story === "textbox");
  await page.keyboard.press("Escape"); await frame(page);
  check("Escape leaves the box", (await ctx(page)).Story === "body");
  // save and read back
  await sel(page, "office-editor-save").click();
  const body = await waitUploads(page, 1);
  const doc = await (await unzip(body)).get("word/document.xml").async("string");
  check("saved: the document has a text box with the text", /<w:txbxContent>(?:(?!<\/w:txbxContent>)[\s\S])*Note to self/.test(doc), (doc.match(/<w:txbxContent>/g) ?? []).length);
  fs.writeFileSync(`${S}/out/textbox14-sample-formatted.docx`, body);
  // undo: the typing, then the box
  for (let i = 0; i < 40 && (await boxes(page)).length > n0; i++) { await sel(page, "word-undo").click(); await frame(page); }
  check("undo takes the box away again", (await boxes(page)).length === n0 && (await ctx(page)).Story === "body");
  // in a header: no page break, no text box
  await sel(page, "word-header-footer").click(); await sel(page, "word-edit-header").click(); await frame(page);
  c = await ctx(page);
  await sel(page, "word-header-footer").click();
  check("in a header neither is offered", c.Story === "header" && (await sel(page, "word-page-break").isDisabled()) && (await sel(page, "word-insert-text-box").isDisabled()), c.Story);
  await page.keyboard.press("Escape"); await frame(page);
  const inHeader = await ev(page, () => window.__word.insertPageBreak(window.__editors[0]));
  check("insertPageBreak says false in a header", inHeader === false);
  await page.close();
}

// ---- body navigation stays out of text boxes; select all leaves their text alone
{
  const file = `${T}/test-documents-wild/docx4j-drawingml-wps.docx`;
  const page = await newPage(context);
  await open(page, "word-editor", file); await ready(page);
  const before = (await boxes(page)).map((b) => b.Text).join("|");
  await ev(page, () => window.__word.getInputElement(window.__editors[0]).focus({ preventScroll: true }));
  await page.keyboard.press("Control+End"); await frame(page);
  check("Ctrl+End is in the body", (await ctx(page)).Story === "body" && (await sel(page, "word-text-box-bar").count()) === 0, await ctx(page));
  await page.keyboard.type("END"); await frame(page);
  check("what is typed there does not go into a box", (await boxes(page)).map((b) => b.Text).join("|") === before);
  await page.keyboard.press("Control+a"); await frame(page);
  await sel(page, "word-bold").click(); await frame(page);
  const rot = await ev(page, () => { const w = window.__word, e = window.__editors[0]; const b = w.getShapeRects(e).find((r) => r.Text); w.enterTextBox(e, b.Id); const c = w.getCaretRect(e); return c ? { rotation: c.Rotation, keys: Object.keys(c) } : null; });
  check("getCaretRect in a text box has Rotation (a number)", rot && typeof rot.rotation === "number", rot);
  await page.screenshot({ path: `${SHOTS}/word14-wps-box-caret.png` });
  const dup = await ev(page, () => window.__word.getDuplicateBlockIds(window.__editors[0]));
  check("getDuplicateBlockIds: none", Array.isArray(dup) ? dup.length === 0 : dup != null && Array.from(dup).length === 0, dup);
  const clamp = await ev(page, () => { const m = window.__wordModule, e = window.__editors[0]; window.__word.exitTextBox(e); const id = m.getSelection(e).AnchorBlockId; return { far: m.setSelection(e, id, 99999, id, 99999), none: m.setSelection(e, "no-such-block", 0, "no-such-block", 0) }; });
  check("setSelection clamps an offset (true) and refuses an unknown block (false)", clamp.far === true && clamp.none === false, clamp);
  await page.close();
}

// ---- sections: the pages the editor and the preview count
{
  for (const [file, want] of [[`${T}/word/roundtrip-saved/three-sections.original.docx`, 3], [`${T}/test-documents-wild/phpword-sample-11-readword2007.docx`, null], [`${T}/repros/word/landscape-section.docx`, null], [`${T}/test-documents-business/docx4j-headers.docx`, null]]) {
    const page = await newPage(context);
    await open(page, "word-editor", file); await ready(page);
    const n = await ev(page, () => window.__word.getPageCount(window.__editors[0]));
    const st = (await status(page)).replace(/\n/g, " | ");
    check(`${file.split("/").pop()}: the status shows the engine's page count${want ? ` (${want})` : ""}`, new RegExp(`\\b${n} pages?$`).test(st.trim()) && (want === null || n === want), { pages: n, status: st });
    await page.close();
  }
}
await browser.close();
process.exit(summary() ? 1 : 0);
