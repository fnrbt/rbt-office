// Save fidelity after ONE edit through the UI and the Save button: docx. usage: node fid13-word.mjs [file|-] [edits,comma] [table]
import { launch, newPage, open, T, S, check, summary, waitUploads, unzip } from "./lib.mjs";
import fs from "node:fs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const bytesOf = async (zip, name) => Buffer.from(await zip.get(name).async("uint8array"));
const textOf = async (zip, name) => (zip.has(name) ? await zip.get(name).async("string") : "");
const count = (s, re) => (s.match(re) ?? []).length;
const toks = (s) => s.match(/<[^>]*>|[^<]+/g) ?? [];
const bagDiff = (x, y) => { const ca = new Map(); for (const t of x) ca.set(t, (ca.get(t) ?? 0) + 1); const added = []; for (const t of y) { const n = ca.get(t) ?? 0; if (n > 0) ca.set(t, n - 1); else added.push(t); } const gone = []; for (const [t, n] of ca) for (let i = 0; i < n; i++) gone.push(t); return { gone, added }; };
const shapes = (l) => { const c = new Map(); for (const t of l) { const k = t.startsWith("<") ? t.replace(/"[^"]*"/g, '"…"').slice(0, 60) : "(text)"; c.set(k, (c.get(k) ?? 0) + 1); } return [...c.entries()].sort((p, q) => q[1] - p[1]).slice(0, 8).map(([k, n]) => `${n}x ${k}`).join("  "); };
const clip = (s, n = 300) => (s.length > n ? s.slice(0, n * 0.65) + ` …(${s.length})… ` + s.slice(-n * 0.25) : s);
const paragraphs = (xml) => { const body = xml.slice(xml.indexOf("<w:body")); return body.match(/<w:p[ >](?:(?!<w:p[ >])[\s\S])*?<\/w:p>/g) ?? []; };
// relationships of the main part that something must refer to
const orphans = async (zip) => { const rels = await textOf(zip, "word/_rels/document.xml.rels"), doc = await textOf(zip, "word/document.xml"); const out = []; for (const m of rels.matchAll(/<Relationship [^>]*>/g)) { const id = /Id="([^"]+)"/.exec(m[0])?.[1], type = /Type="[^"]*\/([^"\/]+)"/.exec(m[0])?.[1]; if (/^(image|hyperlink|header|footer|chart|oleObject|package|diagramData|diagramLayout|diagramQuickStyle|diagramColors|control|aFChunk|subDocument|video|audio)$/.test(type) && !doc.includes(`"${id}"`)) out.push(`${id} (${type})`); } return out; };
const ids = (xml) => [...xml.matchAll(/<Relationship [^>]*Id="([^"]+)"/g)].map((m) => m[1]);
const docs = process.argv[2] && process.argv[2] !== "-" ? [[process.argv[2], process.argv[4] === "table"]] : [[`${T}/test-documents/sample-formatted.docx`, false], [`${T}/test-documents-wild/docx4j-getting-started.docx`, false], [`${T}/test-documents-wild/phpword-sample-11-readword2007.docx`, false], [`${T}/test-documents-wild/docx4j-tables-rich.docx`, true], [`${T}/test-documents-online/file-sample.docx`, false]];
const edits = (process.argv[3] ?? "type,bold,size,align,enter,row,link,pendbold,spacing,listtab,column").split(",");
const expected = { pendbold: [1, 1], spacing: [1, 1], listtab: [1, 1], type: [1, 1], bold: [1, 1], size: [2, 2], align: [1, 1], enter: [2, 1], link: [1, 1] };
const table = [];
for (const [file, inTable] of docs) {
  const name = file.split("/").pop();
  const original = fs.readFileSync(file);
  const a = await unzip(original);
  const docA = await textOf(a, "word/document.xml");
  const parasA = paragraphs(docA);
  const orphansA = await orphans(a);
  console.log(`\n##### ${name}${inTable ? " (every edit in a paragraph of a table cell)" : ""}: ${original.length} bytes, ${a.size} parts, document.xml ${docA.length} chars, ${parasA.length} paragraphs, ${count(docA, /<w:tbl>/g)} tables, ${count(docA, /<w:hyperlink /g)} hyperlinks; relationships nothing refers to in the file as it is: ${orphansA.join(", ") || "none"}`);
  for (const edit of edits) {
    const page = await newPage(context);
    await open(page, "word-editor", file);
    await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
    await page.waitForFunction(() => window.__word.isLoadComplete(window.__editors.at(-1)) !== false && window.__word.isLayoutComplete(window.__editors.at(-1)), null, { timeout: 300000 });
    await page.waitForTimeout(300);
    // ---- where: a caret at the start of the paragraph to edit
    const cellMode = inTable || edit === "row" || edit === "column";
    let target = await ev(page, ([edit, cellMode]) => {
      const w = window.__word, e = window.__editors.at(-1), m = window.__wordModule;
      if (cellMode && edit !== "link") {
        const cells = m.getTableGeometry(e).filter((g) => g.Kind === "table-cell" && g.Width > 60 && g.Height > 12); const out = [];
        for (const cell of cells) {
          const pos = m.hitTestPosition(e, cell.X + Math.min(14, cell.Width / 3), cell.Y + Math.min(12, cell.Height / 2));
          if (!pos) continue;
          m.setSelection(e, pos.BlockId, 0, pos.BlockId, 0);
          const c = w.getCaretContext(e);
          if (!c.InTable) continue;
          out.push({ id: pos.BlockId, length: 0, where: `table ${cell.Table}, row ${cell.Row}, cell ${cell.Cell} (${c.RowCount} rows x ${c.ColumnCount} columns)`, text: "" });
        }
        return { candidates: out.filter((x, i) => out.findIndex((y) => y.id === x.id) === i).slice(0, 80) };
        return null;
      }
      const blocks = m.getLayoutFingerprint(e);
      for (const strict of [true, false]) for (let i = 0; i < blocks.length; i++) {
        const block = blocks[i]; const text = block.Text ?? "";
        if (edit === "link" && !/,link,/.test(block.Runs ?? "")) continue;
        if (text.length < (edit === "listtab" ? 4 : 14) || text.length > 900) continue;
        if (edit !== "link" && edit !== "listtab" && (strict ? !/^[A-Za-z][A-Za-z ,.'-]{50,}/.test(text) : !/^[A-Za-z]+ [A-Za-z]+ [A-Za-z ,.'-]{4,}/.test(text))) continue;
        if (edit === "link" && !/^\S+ \S/.test(text)) continue;
        if (edit === "size" && !((blocks[i + 1]?.Text ?? "").length >= 6)) continue;
        const matches = w.findText(e, text.slice(0, 40), true);
        if (matches.length !== 1) continue;
        w.selectFindMatch(e, matches[0]);
        const s = m.getSelection(e);
        if (Math.min(s.AnchorOffset, s.FocusOffset) !== 0) continue;
        const id = s.AnchorBlockId;
        m.setSelection(e, id, 0, id, 0);
        if (w.getCaretContext(e).InTable) continue;
        if (edit === "listtab" && !(w.getCaretContext(e).InList)) continue;
        w.getInputElement(e).focus({ preventScroll: true });
        return { id, length: text.length, where: `body paragraph ${i + 1} of ${blocks.length} blocks`, text: text.slice(0, 50), runs: (block.Runs ?? "").split(" | ").length };
      }
      return null;
    }, [edit, cellMode]);
    if (target?.candidates) {
      // the paragraph's length: the caret's offset after End (real key)
      let found = null;
      for (const c of target.candidates) {
        await ev(page, (id) => { const e = window.__editors.at(-1); window.__wordModule.setSelection(e, id, 0, id, 0); window.__word.getInputElement(e).focus({ preventScroll: true }); }, c.id);
        await page.keyboard.press("End");
        const s = await ev(page, () => window.__wordModule.getSelection(window.__editors.at(-1)));
        if (s.FocusBlockId !== c.id) continue;
        c.length = s.FocusOffset;
        if (edit === "row" || edit === "column" || c.length >= 12) { found = c; break; }
      }
      if (found) await ev(page, (id) => window.__wordModule.setSelection(window.__editors.at(-1), id, 0, id, 0), found.id);
      target = found;
    }
    if (!target) { console.log(`\n=== ${name} / ${edit}: no place for it in this document`); table.push({ file: name, edit, what: "no place for it" }); await page.close(); continue; }
    await frame(page);
    const selInfo = () => ev(page, () => { const m = window.__wordModule, w = window.__word, e = window.__editors.at(-1); const s = m.getSelection(e); const f = w.getSelectionFormat(e); return { a: s.AnchorBlockId, ao: s.AnchorOffset, f: s.FocusBlockId, fo: s.FocusOffset, bold: f.Bold, size: f.FontSizePt, align: f.Paragraph.Alignment, link: !!f.Hyperlink, ctx: w.getCaretContext(e).InTable }; });
    const setSel = (ao, fo) => ev(page, ([id, ao, fo]) => window.__wordModule.setSelection(window.__editors.at(-1), id, ao, id, fo), [target.id, ao, fo]);
    let what = "";
    if (edit === "type") { await setSel(3, 3); await page.keyboard.type("x"); what = `"x" typed at offset 3`; }
    else if (edit === "bold" || edit === "link") {
      await page.keyboard.press("Control+Shift+ArrowRight"); await frame(page);
      const s = await selInfo();
      await sel(page, "word-bold").click(); await frame(page);
      const s2 = await selInfo();
      what = `the first word (offsets ${Math.min(s.ao, s.fo)}-${Math.max(s.ao, s.fo)}, Ctrl+Shift+Right) made bold with the toolbar button (bold ${JSON.stringify(s.bold)} -> ${JSON.stringify(s2.bold)})${edit === "link" ? `; the paragraph has a hyperlink (${target.runs} runs)` : ""}`;
    } else if (edit === "size") {
      await page.keyboard.press("End"); await setSel(Math.max(0, target.length - 5), Math.max(0, target.length - 5)); await frame(page);
      for (let i = 0; i < 11; i++) await page.keyboard.press("Shift+ArrowRight");
      await frame(page);
      const s = await selInfo();
      const to = s.size === 20 ? "24" : "20";
      await sel(page, "word-font-size").selectOption(to); await frame(page);
      const s2 = await selInfo();
      what = `font size of a selection ${s.a === s.f ? "that stayed in ONE paragraph (Shift+Right did not leave it)" : "across two paragraphs"} (last 5 characters of one, Shift+Right x11) ${JSON.stringify(s.size)} -> ${JSON.stringify(s2.size)} pt with the size box`;
    } else if (edit === "align") {
      await setSel(2, 2); await frame(page);
      const s = await selInfo(); const to = s.align === "center" ? "right" : "center";
      await sel(page, `word-align-${to}`).click(); await frame(page);
      what = `alignment ${s.align} -> ${(await selInfo()).align} with the toolbar button`;
    } else if (edit === "enter") {
      const at = Math.max(1, Math.floor(target.length / 2));
      await setSel(at, at); await page.keyboard.press("Enter");
      what = `Enter at offset ${at} of ${target.length}`;
    } else if (edit === "pendbold") {
      await setSel(3, 3); await frame(page);
      const s = await selInfo();
      await sel(page, "word-bold").click(); await frame(page);
      const s2 = await selInfo();
      const mod = await ev(page, () => window.__word.isModified(window.__editors.at(-1)));
      await page.keyboard.type("new"); await frame(page);
      what = `Bold toggled at a collapsed caret (offset 3; bold ${JSON.stringify(s.bold)} -> ${JSON.stringify(s2.bold)}, modified after the toggle alone: ${mod}), then "new" typed`;
    } else if (edit === "spacing") {
      await setSel(2, 2); await frame(page);
      const f0 = await ev(page, () => { const p = window.__word.getSelectionFormat(window.__editors.at(-1)).Paragraph; return [p.SpaceBefore, p.SpaceAfter]; });
      await sel(page, "word-spacing").click(); await frame(page);
      await sel(page, "word-space-before").fill("10"); await sel(page, "word-space-after").fill("7.5");
      await sel(page, "word-spacing-apply").click(); await frame(page);
      const f1 = await ev(page, () => { const p = window.__word.getSelectionFormat(window.__editors.at(-1)).Paragraph; return [p.SpaceBefore, p.SpaceAfter]; });
      what = `paragraph spacing before/after ${JSON.stringify(f0)} -> ${JSON.stringify(f1)} pt (spacing popover: 10 and 7.5)`;
    } else if (edit === "listtab") {
      await setSel(0, 0); await frame(page);
      const l0 = await ev(page, () => { const p = window.__word.getSelectionFormat(window.__editors.at(-1)).Paragraph; return [p.ListKind, p.ListLevel, p.IndentLeft]; });
      await page.keyboard.press("Tab"); await frame(page);
      const l1 = await ev(page, () => { const p = window.__word.getSelectionFormat(window.__editors.at(-1)).Paragraph; return [p.ListKind, p.ListLevel, p.IndentLeft]; });
      what = `Tab at the start of a list item: kind/level/indent ${JSON.stringify(l0)} -> ${JSON.stringify(l1)}`;
    } else if (edit === "column") {
      const c0 = await ev(page, () => { const c = window.__word.getCaretContext(window.__editors.at(-1)); return [c.RowCount, c.ColumnCount]; });
      await sel(page, "word-column-right").click(); await frame(page);
      const c1 = await ev(page, () => { const c = window.__word.getCaretContext(window.__editors.at(-1)); return [c.RowCount, c.ColumnCount]; });
      what = `a column inserted to the right (toolbar): rows x columns ${c0.join(" x ")} -> ${c1.join(" x ")}`;
    } else if (edit === "row") {
      await sel(page, "word-row-below").click(); await frame(page);
      what = `a row inserted below (toolbar's Insert row below)`;
    }
    await frame(page);
    const modified = await ev(page, () => window.__word.isModified(window.__editors.at(-1)));
    const t0 = Date.now();
    await sel(page, "office-editor-save").click();
    const out = await waitUploads(page, 1, 180000);
    const ms = Date.now() - t0;
    const report = await ev(page, async () => (await window.__formats()).default().word.getLastSaveReport());
    fs.writeFileSync(`${S}/out/f14-${edit}${inTable ? "-cell" : ""}-${name}`, out);
    const b = await unzip(out);
    console.log(`\n=== ${name} / ${edit}: ${what}; in ${target.where}${target.text ? ` "${target.text}..."` : ""}; modified ${modified}; saved in ${ms} ms, ${original.length} -> ${out.length} bytes`);
    console.log(`  getLastSaveReport: ${JSON.stringify(report)}`);
    const changed = [], dropped = [], added = [];
    for (const n of a.keys()) { if (!b.has(n)) { dropped.push(n); continue; } const x = await bytesOf(a, n), y = await bytesOf(b, n); if (!x.equals(y)) changed.push({ n, from: x.length, to: y.length, t: /\.(xml|rels)$/.test(n) ? bagDiff(toks(x.toString("utf8")), toks(y.toString("utf8"))) : null }); }
    for (const n of b.keys()) if (!a.has(n)) added.push(`${n} (${(await bytesOf(b, n)).length} bytes)`);
    console.log(`  parts: ${a.size} in the file, ${a.size - changed.length - dropped.length} byte-identical, ${changed.length} changed, ${dropped.length} dropped, ${added.length} added`);
    for (const c of changed) { console.log(`    changed ${c.n}: ${c.from} -> ${c.to} bytes${c.t ? `; ${c.t.gone.length} tags/texts gone, ${c.t.added.length} new` : ""}`); if (c.t && c.n !== "word/document.xml") { if (c.t.gone.length) console.log(`       gone: ${shapes(c.t.gone)}`); if (c.t.added.length) console.log(`       new:  ${shapes(c.t.added)}`); } }
    for (const d of dropped) console.log(`    dropped ${d}`);
    for (const x of added) console.log(`    added   ${x}`);
    const docB = await textOf(b, "word/document.xml");
    const parasB = paragraphs(docB);
    const pd = bagDiff(parasA, parasB);
    // outside the paragraphs: the document with every paragraph taken out
    const strip = (xml, ps) => { let s = xml; return toks(xml.replace(/<w:p[ >](?:(?!<w:p[ >])[\s\S])*?<\/w:p>/g, "")); };
    const od = bagDiff(strip(docA), strip(docB));
    console.log(`  word/document.xml paragraph by paragraph: ${parasA.length} -> ${parasB.length} paragraphs; ${pd.added.length} of the saved ones are not XML of the file, ${pd.gone.length} of the file's are not in the saved one; outside the paragraphs (tables' and sections' own tags) ${od.gone.length} tags gone, ${od.added.length} new${od.gone.length + od.added.length ? `: gone ${shapes(od.gone)} | new ${shapes(od.added)}` : ""}`);
    const n = Math.min(3, Math.max(pd.gone.length, pd.added.length));
    for (let i = 0; i < n; i++) {
      const was = pd.gone[i] ?? "", is = pd.added[i] ?? "";
      const td = bagDiff(toks(was), toks(is));
      console.log(`      paragraph ${i + 1}: ${was.length} -> ${is.length} chars; tags/texts gone: ${td.gone.length ? td.gone.map((t) => clip(t, 90)).slice(0, 8).join(" ") : "none"} | new: ${td.added.length ? td.added.map((t) => clip(t, 90)).slice(0, 10).join(" ") : "none"}`);
    }
    if (Math.max(pd.gone.length, pd.added.length) > 3) console.log(`      ... and ${Math.max(pd.gone.length, pd.added.length) - 3} more; the new ones by shape: ${shapes(pd.added.flatMap(toks))}`);
    const orphansB = (await orphans(b)).filter((o) => !orphansA.includes(o));
    const relIds = ids(await textOf(b, "word/_rels/document.xml.rels"));
    const dupIds = relIds.filter((id, i) => relIds.indexOf(id) !== i);
    const exp = expected[edit];
    const flags = [];
    if (exp && (pd.added.length > exp[0] || pd.gone.length > exp[1])) flags.push(`more paragraphs changed than the edit touches (${pd.added.length} new / ${pd.gone.length} gone, expected at most ${exp[0]} / ${exp[1]})`);
    if (edit === "row" && pd.gone.length > 0) flags.push(`${pd.gone.length} existing paragraphs written anew by a row insert`);
    if (edit === "column" && pd.gone.length > 0) flags.push(`${pd.gone.length} existing paragraphs written anew by a column insert`);
    const marks = (xml) => { const names = [...xml.matchAll(/<w:bookmarkStart [^>]*w:name="([^"]*)"/g)].map((m) => m[1]); return { starts: names.length, dupNames: names.filter((n, i) => names.indexOf(n) !== i).length, ends: count(xml, /<w:bookmarkEnd /g) }; };
    const mA = marks(docA), mB = marks(docB);
    if (mB.dupNames > mA.dupNames || mB.starts - mB.ends !== mA.starts - mA.ends) flags.push(`bookmarks: starts/ends/duplicate names ${mA.starts}/${mA.ends}/${mA.dupNames} -> ${mB.starts}/${mB.ends}/${mB.dupNames}`);
    if (edit === "column") { const g = (x) => (x.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/g) ?? []).map((t) => count(t, /<w:gridCol/g)).join(","); const bare = count(docB, /<w:tc><w:p\/><\/w:tc>|<w:tc><w:p><\/w:p><\/w:tc>/g) - count(docA, /<w:tc><w:p\/><\/w:tc>|<w:tc><w:p><\/w:p><\/w:tc>/g); console.log(`  tables' grids (columns per table): ${g(docA)} -> ${g(docB)}; cells ${count(docA, /<w:tc>/g)} -> ${count(docB, /<w:tc>/g)}; bare cells (<w:tc><w:p/></w:tc>) added: ${bare}; new cells with properties: ${count(docB, /<w:tcPr>/g) - count(docA, /<w:tcPr>/g)}`); }
    if (edit !== "row" && edit !== "column" && edit !== "enter" && report.GeneratedParagraphs > 0) flags.push(`${report.GeneratedParagraphs} paragraphs written from the model (${report.Reasons.join("; ")})`);
    if (edit !== "row" && edit !== "column" && (od.gone.length || od.added.length)) flags.push(`tags outside the paragraphs changed (${od.gone.length} gone, ${od.added.length} new)`);
    if (changed.some((c) => c.n !== "word/document.xml") || dropped.length || added.length) flags.push(`other parts touched: ${[...changed.filter((c) => c.n !== "word/document.xml").map((c) => c.n), ...dropped.map((d) => "-" + d), ...added.map((x) => "+" + x)].join(", ")}`);
    if (orphansB.length) flags.push(`relationships without a referrer: ${orphansB.join(", ")}`);
    if (dupIds.length) flags.push(`duplicate relationship ids: ${dupIds.join(", ")}`);
    console.log(`  relationships without a referrer (new): ${orphansB.join(", ") || "none"}; duplicate ids: ${dupIds.join(", ") || "none"}`);
    console.log(`  FLAGS: ${flags.length ? flags.join(" || ") : "none"}`);
    check(`${name} / ${edit}: the edit was made and saved`, modified === true && out.length > 1000);
    check(`${name} / ${edit}: no relationship without a referrer, no duplicate id`, orphansB.length === 0 && dupIds.length === 0, { orphansB, dupIds });
    table.push({ file: name, edit, cell: cellMode, what, ms, report, parts: { changed: changed.map((c) => `${c.n} ${c.from}->${c.to}`), dropped, added }, paragraphs: [parasA.length, parasB.length, pd.added.length, pd.gone.length], outside: [od.gone.length, od.added.length], flags });
    await page.close();
  }
}
fs.writeFileSync(`${S}/fid14-word-table-${process.argv[2] && process.argv[2] !== "-" ? "one" : "all"}.json`, JSON.stringify(table, null, 1));
await browser.close();
process.exit(summary() ? 1 : 0);
