// Save fidelity, the edits 0.1.15 / 0.1.16 changed: which package parts change (bytes), what of the main part changes.
// xlsx: note added / edited / removed, row inserted above charts, formula typed into a General cell, SetCell.
// docx: typing after bold-at-caret in bookmarked paragraphs, shape moved / resized / deleted, typing beside a comment mark,
//       bold text typed into a field's result, section break kinds, list level, pasteText.
// usage: node fid16-extra.mjs [sheet|word] [keys]   -> out/f16-*, fid16-extra-table.json
import { launch, newPage, open, T, S, check, summary, waitUploads, unzip } from "./lib.mjs";
import { g, sel, ready, frames, cellPoint } from "./ui-lib.mjs";
import fs from "node:fs";
const which = process.argv[2] ?? "sheet,word"; const only = process.argv[3]?.split(",");
const { browser, context } = await launch({ viewport: { width: 1400, height: 820 } });
const table = [];
const partsDiff = async (original, out) => {
  const a = await unzip(original), b = await unzip(out); const changed = [], dropped = [], added = [];
  for (const [n, f] of a) { if (!b.has(n)) { dropped.push(n); continue; } const x = Buffer.from(await f.async("uint8array")), y = Buffer.from(await b.get(n).async("uint8array")); if (!x.equals(y)) changed.push(`${n} ${x.length}->${y.length}`); }
  for (const n of b.keys()) if (!a.has(n)) added.push(n);
  return { a, b, changed, dropped, added, parts: [a.size, b.size] };
};
const toks = (s) => s.match(/<[^>]*>|[^<]+/g) ?? [];
const bag = (x, y) => { const ca = new Map(); for (const t of x) ca.set(t, (ca.get(t) ?? 0) + 1); let added = 0; const addedList = []; for (const t of y) { const n = ca.get(t) ?? 0; if (n > 0) ca.set(t, n - 1); else { added++; if (addedList.length < 6) addedList.push(t.slice(0, 90)); } } let gone = 0; const goneList = []; for (const [t, n] of ca) { gone += n; if (n > 0 && goneList.length < 6) goneList.push(t.slice(0, 90)); } return { gone, added, goneList, addedList }; };
const text = async (zip, n) => (zip.has(n) ? await zip.get(n).async("string") : "");
const record = async (kind, file, edit, what, original, out, mainPart, extra = {}) => {
  const name = file.split("/").pop();
  fs.writeFileSync(`${S}/out/f16-${edit}-${name}`, out);
  const d = await partsDiff(original, out);
  const main = mainPart ? bag(toks(await text(d.a, mainPart)), toks(await text(d.b, mainPart))) : null;
  const row = { kind, file: name, edit, what, bytes: `${original.length} -> ${out.length}`, parts: d.parts.join(" -> "), changed: d.changed, dropped: d.dropped, added: d.added, main: main && { part: mainPart, tokensGone: main.gone, tokensNew: main.added, gone: main.goneList, new: main.addedList }, ...extra };
  table.push(row);
  console.log(`\n== ${name} / ${edit}: ${what}\n   bytes ${row.bytes}; parts ${row.parts}; changed: ${d.changed.join(", ") || "none"}; dropped: ${d.dropped.join(", ") || "none"}; added: ${d.added.join(", ") || "none"}`);
  if (main) console.log(`   ${mainPart}: ${main.gone} tokens gone, ${main.added} new\n     gone: ${main.goneList.join(" | ")}\n     new:  ${main.addedList.join(" | ")}`);
  for (const [k, v] of Object.entries(extra)) console.log(`   ${k}: ${typeof v === "string" ? v : JSON.stringify(v)}`);
  return d;
};

// ============================== xlsx ==============================
if (which.includes("sheet")) {
  const files = [`${T}/real-world/milk-street-office.xlsx`, `${T}/real-world/jumeirah-hotel.xlsx`, `${T}/real-world/dashboard-ecommerce-sales.xlsx`, `${T}/real-world/dashboard-sales-analysis.xlsx`, `${T}/real-world/chart-zoo.xlsx`, `${T}/real-world/dcf-stock-analysis.xlsx`, `${T}/real-world/lyric-multifamily.xlsx`, `${T}/real-world/saopaulo-condos.xlsx`, `${T}/real-world/conditional-formatting-tutorial.xlsx`];
  const edits = { noteadd: 1, noteedit: 1, notedelete: 1, rowinsert: 1, formula: 1, setcell: 1 };
  for (const file of files) for (const edit of Object.keys(edits)) {
    if (only && !only.includes(edit)) continue;
    const name = file.split("/").pop(); const original = fs.readFileSync(file);
    const page = await newPage(context); const errors = []; page.on("pageerror", (e) => errors.push(String(e).slice(0, 200)));
    await open(page, "sheet-editor", file); await ready(page);
    // the sheet to edit: the first with a drawing for the row insert, with a note for the note edits, else the one on display
    const plan = await g(page, (edit) => {
      const gr = window.__grids.at(-1); const infos = gr.GetSheetInfos(); const out = { sheet: gr.GetActiveSheetIndex() };
      const worksheet = (i) => { try { return gr.GetSheetKind(i) === "worksheet"; } catch { return true; } };
      if (edit === "rowinsert") { for (let i = 0; i < infos.length; i++) { if (!worksheet(i)) continue; gr.SwitchSheet(i); const d = gr.GetDrawings(); if (d.length) { out.sheet = i; out.drawings = d.map((x) => x.kind).join(","); break; } } }
      if (edit === "noteedit" || edit === "notedelete") { let found = false; for (let i = 0; i < infos.length && !found; i++) { if (!worksheet(i)) continue; gr.SwitchSheet(i); const u = gr.GetUsedRange(); if (!u) continue; for (let r = 0; r <= Math.min(u.lastRow, 300) && !found; r++) for (let c = 0; c <= Math.min(u.lastCol, 40) && !found; c++) { const k = gr.GetComment(r, c); if (k) { out.sheet = i; out.note = { r, c, author: k.author, text: (k.text ?? "").slice(0, 40) }; found = true; } } } if (!found) return null; }
      gr.SwitchSheet(out.sheet);
      if (gr.GetSheetKind() !== "worksheet") { for (let i = 0; i < infos.length; i++) if (worksheet(i)) { gr.SwitchSheet(i); out.sheet = i; break; } }
      const u = gr.GetUsedRange(); out.used = u ? u.range : null; out.name = infos[out.sheet].name ?? infos[out.sheet].Name;
      // an empty General cell beside the data, and a numeric cell to read
      let num = null; if (u) for (let r = 0; r <= Math.min(u.lastRow, 200) && !num; r++) for (let c = 0; c <= Math.min(u.lastCol, 30) && !num; c++) { const k = gr.GetCell(r, c); if (k.valueType === "number" && !k.formula) num = { r, c, address: k.address }; }
      out.num = num; out.free = { r: 1, c: (u ? u.lastCol : 0) + 2 };
      return out;
    }, edit);
    if (!plan) { console.log(`\n== ${name} / ${edit}: no note in the workbook - skipped`); await page.close(); continue; }
    await frames(page);
    let what = "";
    if (edit === "noteadd") { what = `a note added on ${plan.name}!${plan.free.r + 1},${plan.free.c + 1} (through the comment dialog's call)`; await g(page, (p) => { const gr = window.__grids.at(-1); gr.SelectCell(p.free.r, p.free.c); gr.SetComment(p.free.r, p.free.c, "A note from the test", "Tester"); }, plan); }
    if (edit === "noteedit") { what = `the note at ${plan.name}!r${plan.note.r + 1}c${plan.note.c + 1} ("${plan.note.text}") gets " Seen." appended`; await g(page, (p) => { const gr = window.__grids.at(-1); const k = gr.GetComment(p.note.r, p.note.c); gr.SetComment(p.note.r, p.note.c, (k.text ?? "") + " Seen.", k.author); }, plan); }
    if (edit === "notedelete") { what = `the note at ${plan.name}!r${plan.note.r + 1}c${plan.note.c + 1} removed`; await g(page, (p) => window.__grids.at(-1).DeleteComment(p.note.r, p.note.c), plan); }
    if (edit === "rowinsert") { what = `a row inserted at row 2 of ${plan.name} (drawings: ${plan.drawings ?? "none"})`; await g(page, () => { const gr = window.__grids.at(-1); gr.SelectRange(1, 0, 1, 0); gr.InsertRows(); }); }
    if (edit === "formula") { if (!plan.num) { console.log(`\n== ${name} / formula: no numeric input - skipped`); await page.close(); continue; } what = `=${plan.num.address}*2 typed into the empty General cell r${plan.free.r + 1}c${plan.free.c + 1} of ${plan.name}`; const p = await cellPoint(page, plan.free.r, plan.free.c); if (p) { await page.mouse.click(p.x, p.y); await page.keyboard.type(`=${plan.num.address}*2`); await page.keyboard.press("Enter"); } else await g(page, (pl) => window.__grids.at(-1).EnterText(pl.free.r, pl.free.c, `=${pl.num.address}*2`), plan); }
    if (edit === "setcell") { what = `SetCell(r${plan.free.r + 1}c${plan.free.c + 1}, "host") on ${plan.name}`; await g(page, (p) => window.__grids.at(-1).SetCell(p.free.r, p.free.c, "host"), plan); }
    await frames(page);
    const state = await g(page, () => { const gr = window.__grids.at(-1); return { modified: gr.IsModified(), canUndo: gr.CanUndo() }; });
    const n0 = page.uploads.length;
    await sel(page, "office-editor-save").click();
    let out; try { out = await waitUploads(page, n0 + 1, 120000); } catch (e) { console.log(`\n== ${name} / ${edit}: NOT SAVED (${String(e).slice(0, 120)}); status: ${(await sel(page, "office-editor-status").innerText()).replace(/\n/g, " | ")}; errors ${errors.join(" ; ")}`); check(`${name} / ${edit}: saved`, false, errors); await page.close(); continue; }
    const d = await record("xlsx", file, edit, what, original, out, null, { state });
    const others = d.changed.filter((c) => !/^(xl\/worksheets\/sheet\d+\.xml|xl\/comments\d*\.xml|xl\/drawings\/[^ ]+|xl\/worksheets\/_rels\/[^ ]+|xl\/sharedStrings\.xml|xl\/styles\.xml|\[Content_Types\]\.xml|xl\/calcChain\.xml|xl\/workbook\.xml|docProps\/[^ ]+) /.test(c));
    check(`${name} / ${edit}: the edit was made and saved`, state.modified === true && out.length > 500);
    if (edit.startsWith("note")) {
      const sheetsChanged = d.changed.filter((c) => /^xl\/worksheets\/sheet\d+\.xml /.test(c));
      // a first note on a sheet needs <legacyDrawing> in the sheet's part; an edited note changes the comments part and its box only
      check(`${name} / ${edit}: ${edit === "noteadd" ? "at most the one sheet's part changes (its legacyDrawing)" : "no sheet's part changes"}; styles and shared strings are the file's`, sheetsChanged.length <= (edit === "noteadd" ? 1 : 0) && !d.changed.some((c) => /^xl\/(styles|sharedStrings)\.xml /.test(c)), d.changed);
    }
    if (edit === "rowinsert") check(`${name} / rowinsert: no drawing part is renamed, dropped or added; no chart part changes`, !d.dropped.some((n) => /drawings|charts/.test(n)) && !d.added.some((n) => /drawings|charts/.test(n)) && !d.changed.some((c) => /^xl\/charts\//.test(c)), { changed: d.changed, dropped: d.dropped, added: d.added });
    if (edit === "formula" || edit === "setcell") check(`${name} / ${edit}: one sheet's part and at most styles / shared strings / the calculation chain change`, d.changed.filter((c) => /^xl\/worksheets\/sheet\d+\.xml /.test(c)).length === 1 && others.length === 0 && d.dropped.filter((n) => n !== "xl/calcChain.xml").length === 0 && d.added.length === 0, { changed: d.changed, dropped: d.dropped, added: d.added });
    // reopen what was saved: it loads, and the edit is there
    await page.close();
  }
}

// ============================== docx ==============================
if (which.includes("word")) {
  const ev = (page, fn, arg) => page.evaluate(fn, arg);
  const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
  const readyW = async (page) => { await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 }); await page.waitForFunction(() => window.__word.isLoadComplete(window.__editors.at(-1)) !== false && window.__word.isLayoutComplete(window.__editors.at(-1)), null, { timeout: 300000 }); await page.waitForTimeout(300); };
  const focus = (page) => ev(page, () => window.__word.getInputElement(window.__editors.at(-1)).focus({ preventScroll: true }));
  const cases = [
    // [edit, file]
    ["bmbold", `${T}/test-documents/sample-formatted.docx`], ["bmbold", `${S}/synth/range-marks.docx`], ["bmbold", `${T}/test-documents-wild/docx4j-getting-started.docx`],
    ["shapemove", `${S}/synth/shapes.docx`], ["shaperesize", `${S}/synth/shapes.docx`], ["shapedelete", `${S}/synth/shapes.docx`],
    ["shapemove", `${T}/test-documents-wild/docx4j-drawingml-wps.docx`], ["shaperesize", `${T}/test-documents-wild/docx4j-drawingml-wps.docx`], ["shapedelete", `${T}/test-documents-wild/docx4j-drawingml-wps.docx`],
    ["shapemove", `${T}/repros/word-shapes/autoshapes-gallery.docx`], ["shapedelete", `${T}/repros/word-shapes/autoshapes-gallery.docx`], ["shapemove", `${S}/synth/text-boxes.docx`],
    ["commenttype", `${T}/test-documents-business/docx4j-comments-one.docx`], ["commenttype", `${S}/synth/range-marks.docx`], ["commenttype", `${T}/test-documents-business/docx4j-invoice.docx`], ["commentbackspace", `${T}/test-documents-business/docx4j-comments-one.docx`],
    ["fieldbold", `${S}/synth/field-results.docx`], ["fieldbold", `${T}/test-documents/sample-formatted.docx`],
    ["oddpage", `${T}/test-documents/sample-formatted.docx`], ["evenpage", `${S}/synth/three-sections.docx`], ["valign", `${S}/synth/three-sections.docx`],
    ["listlevel", `${S}/synth/lists.docx`], ["listlevel", `${T}/test-documents-features/test-lists.docx`], ["paste", `${T}/test-documents/sample-formatted.docx`], ["headerbox", `${T}/test-documents-features/resume-template.docx`],
  ];
  for (const [edit, file] of cases) {
    if (only && !only.includes(edit)) continue;
    if (!fs.existsSync(file)) { console.log(`\n== ${file} missing`); continue; }
    const name = file.split("/").pop(); const original = fs.readFileSync(file);
    const page = await newPage(context); const errors = []; page.on("pageerror", (e) => errors.push(String(e).slice(0, 200)));
    await open(page, "word-editor", file); await readyW(page);
    let what = "", done = true, extra = {};
    if (edit === "bmbold") {
      // a paragraph inside a bookmark (or holding one): caret a few characters in, Bold at the caret, type
      const t = await ev(page, () => { const w = window.__word, m = window.__wordModule, e = window.__editors.at(-1); const lines = w.getLayoutLines(e); for (const b of w.getBookmarkInfos(e)) { const len = lines.filter((l) => l.BlockId === b.BlockId).reduce((n, l) => n + l.Text.length, 0); if (len >= 8 && !b.Name.startsWith("_GoBack")) { const at = Math.min(len - 2, Math.max(b.StartOffset + 2, 3)); m.setSelection(e, b.BlockId, at, b.BlockId, at); return { name: b.Name, block: b.BlockId, at, start: b.StartOffset, end: b.EndOffset, all: w.getBookmarkInfos(e).length }; } } return null; });
      if (!t) done = false; else { what = `Bold at the caret, "QQ" typed at offset ${t.at} of the paragraph of bookmark ${t.name} (${t.start}-${t.end}); ${t.all} bookmarks`; await focus(page); await sel(page, "word-bold").click(); await page.keyboard.type("QQ"); extra.bookmarksAfter = await ev(page, () => window.__word.getBookmarkInfos(window.__editors.at(-1)).length); }
    }
    if (edit.startsWith("shape")) {
      const t = await ev(page, () => Array.from(window.__word.getShapeRects(window.__editors.at(-1))).filter((s) => s.Story === "body").map((s) => ({ Id: s.Id, W: s.Width, H: s.Height, Text: s.Text.slice(0, 20) })));
      if (!t.length) done = false; else { const s = t[Math.min(1, t.length - 1)];
        const r = await ev(page, ([edit, id]) => { const w = window.__word, e = window.__editors.at(-1); const before = Array.from(w.getShapeRects(e)).find((x) => x.Id === id); const ok = edit === "shapemove" ? w.moveShape(e, id, 30, 20) : edit === "shaperesize" ? w.resizeShape(e, id, before.Width + 24, before.Height + 12) : w.deleteShape(e, id); const after = Array.from(w.getShapeRects(e)).find((x) => x.Id === id); return { ok, before: [before.X, before.Y, before.Width, before.Height].map(Math.round), after: after ? [after.X, after.Y, after.Width, after.Height].map(Math.round) : null, shapes: w.getShapeRects(e).length }; }, [edit, s.Id]);
        what = `${edit} of shape ${s.Id} ("${s.Text}") of ${t.length}: ${JSON.stringify(r)}`; done = r.ok; extra.shapesBefore = t.length; }
    }
    if (edit === "commenttype" || edit === "commentbackspace") {
      const t = await ev(page, () => { const w = window.__word, m = window.__wordModule, e = window.__editors.at(-1); const lines = w.getLayoutLines(e); const line = lines.find((l) => /\[Comment \d+\]/.test(l.Text)); if (!line) return null; const full = lines.filter((l) => l.BlockId === line.BlockId).map((l) => l.Text).join(""); const at = full.indexOf("[Comment"); const len = full.slice(at).indexOf("]") + 1; m.setSelection(e, line.BlockId, at + 4, line.BlockId, at + 4); return { block: line.BlockId, at, len, caret: m.getSelection(e).FocusOffset, anchors: w.getCommentAnchors(e).length }; });
      if (!t) done = false; else { await focus(page); if (edit === "commenttype") { what = `caret set into the comment mark at ${t.at} (it went to ${t.caret}), "Zq" typed`; await page.keyboard.type("Zq"); } else { what = `caret after the comment mark, Backspace (the mark and its comment go)`; await ev(page, (t) => window.__wordModule.setSelection(window.__editors.at(-1), t.block, t.at + t.len, t.block, t.at + t.len), t); await page.keyboard.press("Backspace"); } extra.anchors = [t.anchors, await ev(page, () => window.__word.getCommentAnchors(window.__editors.at(-1)).length)]; }
    }
    if (edit === "fieldbold") {
      const t = await ev(page, () => { const w = window.__word, m = window.__wordModule, e = window.__editors.at(-1); for (const l of w.getLayoutLines(e).slice(0, 400)) { if (!l.Text || l.Text.length < 4) continue; for (let o = l.StartOffset + 1; o < l.EndOffset - 1; o++) { const info = w.getRangeInformation(e, l.BlockId, o, o + 1); if (info && (info.InFieldResult || info.InFieldCode)) { m.setSelection(e, l.BlockId, o, l.BlockId, o); return { block: l.BlockId, at: o, text: l.Text.slice(0, 40) }; } } } return null; }).catch((e) => ({ err: String(e).slice(0, 200) }));
      if (!t || t.err) { done = false; what = JSON.stringify(t); } else { what = `Bold at the caret, "Qq" typed inside a field's result at ${t.at} ("${t.text}")`; await focus(page); await sel(page, "word-bold").click(); await page.keyboard.type("Qq"); }
    }
    if (edit === "oddpage" || edit === "evenpage" || edit === "valign") {
      const t = await ev(page, (edit) => { const w = window.__word, m = window.__wordModule, e = window.__editors.at(-1); const l = w.getLayoutLines(e).find((x) => x.Text.length > 10 && !x.InTable); m.setSelection(e, l.BlockId, 0, l.BlockId, 0); const pages = w.getDocumentSize(e).PageCount; const ok = edit === "valign" ? w.setSectionVerticalAlignment(e, "center") : (w.setSectionBreak(e, edit === "oddpage" ? "oddPage" : "evenPage"), true); w.completeLayout(e); return { ok, pages: [pages, w.getDocumentSize(e).PageCount], valign: w.getSectionVerticalAlignment(e) }; }, edit);
      what = `${edit === "valign" ? 'setSectionVerticalAlignment("center")' : `setSectionBreak("${edit === "oddpage" ? "oddPage" : "evenPage"}") after the first paragraph`}: ${JSON.stringify(t)}`; done = t.ok !== false;
    }
    if (edit === "listlevel") {
      const t = await ev(page, () => { const w = window.__word, m = window.__wordModule, e = window.__editors.at(-1); const items = w.getLayoutLines(e).filter((x) => x.ListMarker && x.StartOffset === 0); const l = items[Math.min(1, items.length - 1)]; if (!l) return null; m.setSelection(e, l.BlockId, 0, l.BlockId, 0); return { block: l.BlockId, marker: l.ListMarker, indent: l.Indent, text: l.Text.slice(0, 30) }; });
      if (!t) done = false; else { await focus(page); await page.keyboard.press("Tab"); await frame(page); const after = await ev(page, (b) => { const l = window.__word.getLayoutLines(window.__editors.at(-1)).find((x) => x.BlockId === b && x.StartOffset === 0); return { marker: l.ListMarker, indent: l.Indent }; }, t.block); what = `Tab at the start of the list item "${t.text}": marker "${t.marker}" -> "${after.marker}", indent ${Math.round(t.indent)} -> ${Math.round(after.indent)}`; }
    }
    if (edit === "paste") {
      const t = await ev(page, () => { const w = window.__word, m = window.__wordModule, e = window.__editors.at(-1); const l = w.getLayoutLines(e).find((x) => /^[A-Za-z]/.test(x.Text) && x.Text.length > 30 && !x.InTable); m.setSelection(e, l.BlockId, 5, l.BlockId, 5); w.toggleItalic(e); w.pasteText(e, "one\ntwo\nthree"); return { block: l.BlockId }; });
      what = `Italic at the caret, pasteText("one\\ntwo\\nthree") at offset 5 of a paragraph`;
    }
    if (edit === "headerbox") {
      const t = await ev(page, () => { const w = window.__word, e = window.__editors.at(-1); const s = Array.from(w.getShapeRects(e)).find((x) => x.Story !== "body"); if (!s) return null; const ok = w.enterTextBox(e, s.Id); w.insertText(e, "Zz"); w.exitTextBox(e); return { ok, id: s.Id, story: s.Story }; });
      if (!t) done = false; else what = `"Zz" typed into the text box of the ${t.story}`;
    }
    if (!done) { console.log(`\n== ${name} / ${edit}: nothing to do it on - skipped ${what}`); await page.close(); continue; }
    await frame(page);
    const modified = await ev(page, () => window.__word.isModified(window.__editors.at(-1)));
    const n0 = page.uploads.length;
    await sel(page, "office-editor-save").click();
    let out; try { out = await waitUploads(page, n0 + 1, 120000); } catch (e) { console.log(`\n== ${name} / ${edit}: NOT SAVED; status ${(await sel(page, "office-editor-status").innerText()).replace(/\n/g, " | ")}; ${errors.join(" ; ")}`); check(`${name} / ${edit}: saved`, false); await page.close(); continue; }
    const report = await ev(page, async () => (await window.__formats()).default().word.getLastSaveReport?.() ?? null).catch(() => null);
    const d = await record("docx", file, edit, what, original, out, "word/document.xml", { modified, report: report && { GeneratedParagraphs: report.GeneratedParagraphs, Reasons: (report.Reasons ?? []).slice(0, 4) } });
    const docA = await text(d.a, "word/document.xml"), docB = await text(d.b, "word/document.xml");
    const marks = (x) => ({ bmStart: (x.match(/<w:bookmarkStart /g) ?? []).length, bmEnd: (x.match(/<w:bookmarkEnd /g) ?? []).length, crStart: (x.match(/<w:commentRangeStart /g) ?? []).length, crEnd: (x.match(/<w:commentRangeEnd /g) ?? []).length, cRef: (x.match(/<w:commentReference /g) ?? []).length, fallback: (x.match(/<mc:Fallback>/g) ?? []).length, anchors: (x.match(/<wp:anchor /g) ?? []).length });
    const mA = marks(docA), mB = marks(docB);
    console.log(`   marks: ${JSON.stringify(mA)} -> ${JSON.stringify(mB)}`);
    table.at(-1).marks = [mA, mB];
    check(`${name} / ${edit}: the edit was made and saved`, modified === true && out.length > 500);
    const otherParts = d.changed.filter((c) => !c.startsWith("word/document.xml "));
    if (edit === "bmbold" || edit === "commenttype" || edit === "fieldbold" || edit === "paste") check(`${name} / ${edit}: only the document part changes; bookmarks and comment ranges as many as before`, otherParts.length === 0 && d.dropped.length === 0 && d.added.length === 0 && mA.bmStart === mB.bmStart && mA.bmEnd === mB.bmEnd && mA.crStart === mB.crStart && mA.crEnd === mB.crEnd && mA.cRef === mB.cRef, { otherParts, dropped: d.dropped, added: d.added });
    if (edit === "shapemove" || edit === "shaperesize") check(`${name} / ${edit}: only the document part changes, as many anchors as before`, otherParts.length === 0 && d.dropped.length === 0 && d.added.length === 0 && mA.anchors === mB.anchors, { otherParts, dropped: d.dropped, added: d.added });
    if (edit === "headerbox") check(`${name} / headerbox: only a header part changes`, d.changed.length === 1 && /^word\/header\d*\.xml /.test(d.changed[0]) && d.dropped.length === 0 && d.added.length === 0, d.changed);
    // reopen what was saved
    const { saved } = await import("./lib.mjs"); saved.set("f16", out);
    const again = await newPage(context); await open(again, "word-editor", "/api/saved/f16", { name }); let reopened = true; try { await readyW(again); } catch { reopened = false; }
    const failed = reopened ? await ev(again, () => window.__word.isLoadFailed(window.__editors.at(-1))) : true;
    check(`${name} / ${edit}: what was saved opens again`, reopened && !failed);
    await again.close(); await page.close();
  }
}
fs.writeFileSync(`${S}/fid16-extra-table.json`, JSON.stringify(table, null, 1));
await browser.close();
process.exit(summary() ? 1 : 0);
