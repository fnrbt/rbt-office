// Save fidelity after ONE edit through the UI and the Save button: xlsx. usage: node fid13-sheet.mjs [file] [edits,comma]
import { launch, newPage, open, T, S, check, summary, waitUploads, unzip } from "./lib.mjs";
import { g, sel, ready, frames, select, cellPoint, menu } from "./ui-lib.mjs";
import fs from "node:fs";
const { browser, context } = await launch();
const bytesOf = async (zip, name) => Buffer.from(await zip.get(name).async("uint8array"));
const textOf = async (zip, name) => (zip.has(name) ? await zip.get(name).async("string") : "");
const count = (s, re) => (s.match(re) ?? []).length;
const one = (s, re) => (re.exec(s) ?? [""])[0];
const span = (a, b) => { let pre = 0; const m = Math.min(a.length, b.length); while (pre < m && a[pre] === b[pre]) pre++; let suf = 0; while (suf < m - pre && a[a.length - 1 - suf] === b[b.length - 1 - suf]) suf++; return { pre, suf, from: a.length - pre - suf, to: b.length - pre - suf, was: a.slice(pre, a.length - suf), is: b.slice(pre, b.length - suf) }; };
const toks = (s) => s.match(/<[^>]*>|[^<]+/g) ?? [];
const bagDiff = (x, y) => { const ca = new Map(); for (const t of toks(x)) ca.set(t, (ca.get(t) ?? 0) + 1); const added = []; for (const t of toks(y)) { const n = ca.get(t) ?? 0; if (n > 0) ca.set(t, n - 1); else added.push(t); } const gone = []; for (const [t, n] of ca) for (let i = 0; i < n; i++) gone.push(t); return { gone, added, total: toks(x).length }; };
const shapes = (l) => { const c = new Map(); for (const t of l) { const k = t.startsWith("<") ? t.replace(/"[^"]*"/g, '"…"').slice(0, 80) : "(text)"; c.set(k, (c.get(k) ?? 0) + 1); } return [...c.entries()].sort((p, q) => q[1] - p[1]).slice(0, 7).map(([k, n]) => `${n}x ${k}`).join("  "); };
const clip = (s, n = 260) => (s.length > n ? s.slice(0, n * 0.6) + ` …(${s.length} chars)… ` + s.slice(-n * 0.3) : s);
const sheetFacts = (x) => ({
  sheetPr: one(x, /<sheetPr[^>]*\/>|<sheetPr[\s\S]*?<\/sheetPr>/), codeName: one(x, /codeName="[^"]*"/), ignoredErrors: one(x, /<ignoredErrors>[\s\S]*?<\/ignoredErrors>/),
  rowStyles: count(x, /<row [^>]*\bs="\d+"[^>]*>/g), rowCustomFormat: count(x, /<row [^>]*customFormat="1"/g), rowHeights: count(x, /<row [^>]*\bht="/g), rows: count(x, /<row[ >]/g),
  sharedFormulas: count(x, /<f [^>]*t="shared"/g), sharedMasters: count(x, /<f [^>]*t="shared"[^>]*ref="/g), formulas: count(x, /<f[ >]/g), arrayFormulas: count(x, /<f [^>]*t="array"/g),
  selection: (x.match(/<selection[^>]*\/>/g) ?? []).join(""), sheetView: one(x, /<sheetView [^>]*>/), cols: one(x, /<cols>[\s\S]*?<\/cols>/), merges: one(x, /<mergeCells[^>]*>/), cf: count(x, /<conditionalFormatting/g), dv: one(x, /<dataValidations[^>]*>/),
  pageSetup: one(x, /<pageSetup[^>]*\/>/), pageMargins: one(x, /<pageMargins[^>]*\/>/), headerFooter: one(x, /<headerFooter[\s\S]*?(\/>|<\/headerFooter>)/), ext: count(x, /<ext /g), drawing: one(x, /<drawing[^>]*\/>/), legacyDrawing: one(x, /<legacyDrawing[^>]*\/>/), tableParts: one(x, /<tableParts[\s\S]*?(\/>|<\/tableParts>)/), cells: count(x, /<c /g), strCells: count(x, /<c [^>]*t="str"/g), sCells: count(x, /<c [^>]*t="s"/g),
});
const files = process.argv[2] && process.argv[2] !== "-" ? [process.argv[2]] : [`${T}/real-world/dashboard-ecommerce-sales.xlsx`, `${T}/real-world/milk-street-office.xlsx`];
const edits = (process.argv[3] ?? "type,below,fill,width,rowheight,row,delrow,comment,editcomment,sort,hidecol").split(",");
const table = [];
for (const file of files) {
  const name = file.split("/").pop();
  const original = fs.readFileSync(file);
  const a = await unzip(original);
  for (const edit of edits) {
    const page = await newPage(context);
    await open(page, "sheet-editor", file); await ready(page); await page.waitForTimeout(300);
    let where = await g(page, (edit) => {
      const gr = window.__grids[0]; const u = gr.GetUsedRange(); const sheet = gr.GetActiveSheetIndex();
      let r, c, extra = null;
      const plain = (i, j) => { const x = gr.GetCell(i, j); return x.valueType !== "empty" && !x.isMerged && !x.formula && !x.comment; };
      if (edit === "below") { r = u.lastRow + 3; c = u.firstCol; }
      else if (edit === "editcomment") {
        const names = gr.GetSheetNames();
        for (const i of [sheet, ...names.map((_, i) => i).filter((i) => i !== sheet)]) { let list = []; try { list = gr.GetComments(i); } catch {} const note = list.find((n) => !n.thread); if (note) return { other: i, row: note.row, col: note.col, author: note.author, text: note.text, notes: list.length }; }
        return null;
      }
      else if (edit === "sort") {
        // four constants in a column, not all the same
        outer: for (let j = u.firstCol; j <= Math.min(u.lastCol, u.firstCol + 30); j++) for (let i = u.firstRow; i + 3 <= Math.min(u.lastRow, u.firstRow + 200); i++) {
          if ([0, 1, 2, 3].every((k) => plain(i + k, j))) { const v = [0, 1, 2, 3].map((k) => gr.GetCell(i + k, j).text); if (new Set(v).size >= 3) { r = i; c = j; extra = v; break outer; } }
        }
        if (r === undefined) return "other-sheet";
      } else {
        r = u.firstRow; c = u.firstCol;
        outer: for (let i = u.firstRow; i <= Math.min(u.lastRow, u.firstRow + 60); i++) for (let j = u.firstCol; j <= Math.min(u.lastCol, u.firstCol + 20); j++) if (plain(i, j)) { r = i; c = j; break outer; }
      }
      gr.SelectCell(r, c);
      return { sheet, name: gr.GetSheetNames()[sheet], row: r, col: c, ref: gr.GetSelectionRef(), before: gr.GetCell(r, c).text, extra, width: gr.GetColumnWidth(c) };
    }, edit);
    let switched = "";
    if (where && where.other !== undefined) {
      const note = where;
      if ((await g(page, () => window.__grids[0].GetActiveSheetIndex())) !== note.other) { await sel(page, `sheet-tab-${note.other}`).click(); await frames(page); await page.waitForTimeout(300); }
      where = await g(page, ([r, c]) => { const gr = window.__grids[0]; gr.SelectCell(r, c); const sheet = gr.GetActiveSheetIndex(); return { sheet, name: gr.GetSheetNames()[sheet], row: r, col: c, ref: gr.GetSelectionRef(), before: gr.GetCell(r, c).text, extra: null, width: 0 }; }, [note.row, note.col]);
      where.note = note;
    }
    if (where === "other-sheet") {
      // the sheet the file opens on has no four constants in a column: on the first sheet that has
      const n = await g(page, () => window.__grids[0].GetSheetInfos().filter((i) => i.kind === "worksheet" && i.visibility === "visible").map((i) => i.index));
      for (const i of n) {
        await sel(page, `sheet-tab-${i}`).click(); await frames(page); await page.waitForTimeout(300);
        where = await g(page, () => { const gr = window.__grids[0]; const u = gr.GetUsedRange(); const sheet = gr.GetActiveSheetIndex(); const plain = (i, j) => { const x = gr.GetCell(i, j); return x.valueType !== "empty" && !x.isMerged && !x.formula && !x.comment; };
          for (let j = u.firstCol; j <= Math.min(u.lastCol, u.firstCol + 30); j++) for (let i = u.firstRow + 1; i + 3 <= Math.min(u.lastRow, u.firstRow + 200); i++) if ([0, 1, 2, 3].every((k) => plain(i + k, j))) { const v = [0, 1, 2, 3].map((k) => gr.GetCell(i + k, j).text); if (new Set(v).size >= 3) { gr.SelectCell(i, j); return { sheet, name: gr.GetSheetNames()[sheet], row: i, col: j, ref: gr.GetSelectionRef(), before: v[0], extra: v, width: 0 }; } } return null; });
        if (where) { switched = ` (after a click on the tab of ${where.name}: the sheet the file opens on has no such cells)`; break; }
      }
    }
    if (!where) { console.log(`${name} / ${edit}: no place for it found`); await page.close(); continue; }
    await frames(page);
    const colLetter = where.ref.replace(/\d+/g, "");
    let what = ""; let row14 = {};
    const headerPoint = async (axis, index) => g(page, ([axis, index]) => { const sc = document.querySelector('[data-testid="sheet-grid"]').getBoundingClientRect(); const q = [...document.querySelectorAll(axis === "row" ? `.spreadsheet-row-header[data-row="${index}"]` : `.spreadsheet-col-header[data-col="${index}"]`)].map((x) => x.getBoundingClientRect()).find((q) => q.width > 0 && q.height > 0 && q.bottom <= sc.bottom - 2 && q.right <= sc.right - 2 && q.top >= sc.top && q.left >= sc.left); return q ? { x: q.left + q.width / 2, y: q.top + q.height / 2 } : null; }, [axis, index]);
    if (edit === "type" || edit === "below") {
      const p = await cellPoint(page, where.row, where.col);
      if (p) await page.mouse.click(p.x, p.y); else await page.locator('[data-testid="sheet-grid"] .spreadsheet-content').focus();
      await page.keyboard.type("edited"); await page.keyboard.press("Enter");
      const u = await g(page, () => { const gr = window.__grids[0]; return { used: gr.GetUsedRange().range, rows: gr.GetModel().RowCount }; });
      what = edit === "type" ? `"edited" typed over ${where.name}!${where.ref} ("${where.before}")` : `"edited" typed into ${where.name}!${where.ref}, three rows BELOW the last used row (${p ? "clicked" : "selected"}; used range now ${u.used}, grid ${u.rows} rows)`;
    } else if (edit === "rowheight") {
      await cellPoint(page, where.row, where.col);
      const p = await headerPoint("row", where.row);
      const h0 = await g(page, (r) => window.__grids[0].GetRowHeight(r), where.row);
      const h = Math.round(h0) + 13;
      if (p) { await page.mouse.click(p.x, p.y, { button: "right" }); await sel(page, "sheet-context-menu").waitFor(); await sel(page, "sheet-menu-size").click(); await sel(page, "sheet-size-input").fill(String(h)); await sel(page, "sheet-dialog-ok").click(); }
      else await g(page, ([r, h]) => window.__grids[0].SetRowHeight(`${r + 1}:${r + 1}`, h), [where.row, h]);
      what = `height of row ${where.row + 1} of ${where.name} ${Math.round(h0)} -> ${h} px (${p ? "row header menu, dialog" : "the sheet shows no headings: SetRowHeight, the call the dialog makes"})`;
    } else if (edit === "delrow") {
      const p = await cellPoint(page, where.row, where.col);
      await page.mouse.click(p.x, p.y, { button: "right" }); await sel(page, "sheet-context-menu").waitFor(); await sel(page, "sheet-menu-delete-rows").click();
      what = `row ${where.row + 1} of ${where.name} deleted (cell menu)`;
    } else if (edit === "hidecol") {
      await cellPoint(page, where.row, where.col);
      const p = await headerPoint("col", where.col);
      if (p) { await page.mouse.click(p.x, p.y, { button: "right" }); await sel(page, "sheet-context-menu").waitFor(); await sel(page, "sheet-menu-hide").click(); }
      else await g(page, (l) => window.__grids[0].HideColumns(`${l}:${l}`), colLetter);
      what = `column ${colLetter} of ${where.name} hidden (${p ? "column header menu" : "the sheet shows no headings: HideColumns, the call the menu makes"})`;
    } else if (edit === "editcomment") {
      const p = await cellPoint(page, where.row, where.col);
      await page.mouse.click(p.x, p.y, { button: "right" }); await sel(page, "sheet-context-menu").waitFor(); await sel(page, "sheet-menu-comment").click();
      await sel(page, "sheet-comment-text").focus(); await page.keyboard.press("Control+End"); await page.keyboard.type(" Seen."); await sel(page, "sheet-dialog-ok").click();
      const now = await g(page, ([r, c]) => window.__grids[0].GetComment(r, c), [where.row, where.col]);
      what = `the note of ${where.name}!${where.ref} (one of ${where.note.notes}, author ${JSON.stringify(where.note.author)}) got " Seen." at its end (cell menu, dialog); its author now ${JSON.stringify(now.author)}`;
    } else if (edit === "fill") {
      await sel(page, "sheet-fill-color").click();
      await page.locator('[data-testid="sheet-fill-palette"] [data-color="#FEF08A"]').click();
      what = `fill of ${where.name}!${where.ref} ("${where.before}") set`;
    } else if (edit === "width") {
      await cellPoint(page, where.row, where.col);
      const p = await headerPoint("col", where.col);
      const w = Math.round(where.width) + 37;
      if (p) { await page.mouse.click(p.x, p.y, { button: "right" }); await sel(page, "sheet-context-menu").waitFor(); await sel(page, "sheet-menu-size").click(); await sel(page, "sheet-size-input").fill(String(w)); await sel(page, "sheet-dialog-ok").click(); }
      else await g(page, ([l, w]) => window.__grids[0].SetColumnWidth(`${l}:${l}`, w), [colLetter, w]);
      what = `width of column ${colLetter} of ${where.name} ${Math.round(where.width)} -> ${w} px (${p ? "column header menu, dialog" : "the sheet shows no headings, so no header menu: SetColumnWidth, the call the dialog makes"})`;
    } else if (edit === "row") {
      const p = await cellPoint(page, where.row, where.col);
      await page.mouse.click(p.x, p.y, { button: "right" }); await sel(page, "sheet-context-menu").waitFor(); await sel(page, "sheet-menu-insert-rows").click();
      what = `a row inserted above row ${where.row + 1} of ${where.name} (cell menu)`;
    } else if (edit === "comment") {
      const p = await cellPoint(page, where.row, where.col);
      await page.mouse.click(p.x, p.y, { button: "right" }); await sel(page, "sheet-context-menu").waitFor(); await sel(page, "sheet-menu-comment").click();
      await sel(page, "sheet-comment-text").fill("Check this figure"); await sel(page, "sheet-dialog-ok").click();
      what = `a comment added to ${where.name}!${where.ref} (cell menu, dialog)`;
    } else if (edit === "sort") {
      await menu(page, "data", "sort"); await sel(page, "sheet-sort-dialog").waitFor();
      const range = `${colLetter}${where.row + 1}:${colLetter}${where.row + 4}`;
      await sel(page, "sheet-sort-range").fill(range);
      if (await sel(page, "sheet-sort-header").isChecked()) await sel(page, "sheet-sort-header").uncheck();
      const sorted = [...where.extra].sort((x, y) => (isNaN(Number(x.replace(/[,$%]/g, ""))) || isNaN(Number(y.replace(/[,$%]/g, ""))) ? x.localeCompare(y) : Number(x.replace(/[,$%]/g, "")) - Number(y.replace(/[,$%]/g, ""))));
      const order = sorted.join("|") === where.extra.join("|") ? "desc" : "asc";
      await sel(page, "sheet-sort-order-0").selectOption(order);
      await sel(page, "sheet-dialog-ok").click();
      await frames(page);
      const after = await g(page, ([r, c]) => [0, 1, 2, 3].map((k) => window.__grids[0].GetCell(r + k, c).text), [where.row, where.col]);
      what = `${where.name}!${range} sorted ${order} (Data > Sort)${switched}: ${where.extra.join(", ")} -> ${after.join(", ")}`;
    }
    await frames(page);
    const dirty = await sel(page, "office-editor-status").innerText();
    const t0 = Date.now();
    await sel(page, "office-editor-save").click();
    const out = await waitUploads(page, 1, 180000);
    const ms = Date.now() - t0;
    fs.writeFileSync(`${S}/out/f14-${edit}-${name}`, out);
    const b = await unzip(out);
    console.log(`\n=== ${name} / ${edit}: ${what}; status before save "${dirty.replace(/\n/g, " ")}"; saved in ${ms} ms, ${original.length} -> ${out.length} bytes`);
    const changed = [], dropped = [], added = [];
    for (const n of a.keys()) { if (!b.has(n)) { dropped.push(n); continue; } const x = await bytesOf(a, n), y = await bytesOf(b, n); if (!x.equals(y)) changed.push({ n, from: x.length, to: y.length, d: span(x.toString("utf8"), y.toString("utf8")), t: /\.(xml|rels|vml)$/.test(n) ? bagDiff(x.toString("utf8"), y.toString("utf8")) : null }); }
    for (const n of b.keys()) if (!a.has(n)) added.push(`${n} (${(await bytesOf(b, n)).length} bytes)`);
    console.log(`  parts: ${a.size} in the file, ${a.size - changed.length - dropped.length} byte-identical, ${changed.length} changed, ${dropped.length} dropped, ${added.length} added`);
    for (const c of changed) { console.log(`    changed ${c.n}: ${c.from} -> ${c.to} bytes${c.t ? `; of ${c.t.total} tags and texts ${c.t.gone.length} are not in the saved part, ${c.t.added.length} are new in it` : ""}`); if (c.t) { if (c.t.gone.length) console.log(`       gone: ${shapes(c.t.gone)}${c.t.gone.length <= 4 ? "\n             " + c.t.gone.map((t) => clip(t, 200)).join("\n             ") : ""}`); if (c.t.added.length) console.log(`       new:  ${shapes(c.t.added)}${c.t.added.length <= 4 ? "\n             " + c.t.added.map((t) => clip(t, 200)).join("\n             ") : ""}`); } }
    for (const d of dropped) console.log(`    dropped ${d}`);
    for (const x of added) console.log(`    added   ${x}`);
    const part = `xl/worksheets/sheet${where.sheet + 1}.xml`;
    // which sheet part is the active sheet's: by the workbook's relationships
    const wb = await textOf(a, "xl/workbook.xml"), rels = await textOf(a, "xl/_rels/workbook.xml.rels");
    const rid = [...wb.matchAll(/<sheet [^>]*r:id="([^"]+)"/g)][where.sheet]?.[1];
    const target = rid ? new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${rid}"`).exec(rels) : null;
    const sheetPart = target ? "xl/" + (target[1] ?? target[2]).replace(/^\/xl\//, "") : part;
    const fa = sheetFacts(await textOf(a, sheetPart)), fb = sheetFacts(await textOf(b, sheetPart));
    const differ = Object.keys(fa).filter((k) => JSON.stringify(fa[k]) !== JSON.stringify(fb[k]));
    console.log(`  edited sheet ${sheetPart}: facts that differ: ${differ.length === 0 ? "none" : differ.map((k) => `${k}: ${JSON.stringify(clip(String(fa[k]), 120))} -> ${JSON.stringify(clip(String(fb[k]), 120))}`).join("; ")}`);
    console.log(`  kept: codeName ${fa.codeName || "(none)"} | ignoredErrors ${fa.ignoredErrors.length} chars | rows with a style ${fa.rowStyles} | shared formulas ${fa.sharedFormulas} (${fa.sharedMasters} masters) | selection ${JSON.stringify(fa.selection)}`);
    // styles, shared strings, comments
    {
      const sa = await textOf(a, "xl/styles.xml"), sb = await textOf(b, "xl/styles.xml");
      const xfs = (x) => { const m = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(x); return m ? count(m[1], /<xf /g) : 0; };
      const ssa = await textOf(a, "xl/sharedStrings.xml"), ssb = await textOf(b, "xl/sharedStrings.xml");
      let uses = 0, usesA = 0;
      for (const n of b.keys()) if (/^xl\/worksheets\/[^/]+\.xml$/.test(n)) uses += count(await textOf(b, n), /<c [^>]*t="s"/g);
      for (const n of a.keys()) if (/^xl\/worksheets\/[^/]+\.xml$/.test(n)) usesA += count(await textOf(a, n), /<c [^>]*t="s"/g);
      const sstHead = (x) => one(x, /<sst [^>]*>/).replace(/xmlns(:\w+)?="[^"]*"/g, "").replace(/\s+/g, " ");
      console.log(`  styles.xml ${sa.length} -> ${sb.length} chars; cellXfs ${xfs(sa)} -> ${xfs(sb)}; fonts ${count(sa, /<font>|<font /g)} -> ${count(sb, /<font>|<font /g)}; fills ${count(sa, /<fill>/g)} -> ${count(sb, /<fill>/g)}; borders ${count(sa, /<border>|<border /g)} -> ${count(sb, /<border>|<border /g)}; dxfs ${count(sa, /<dxf>/g)} -> ${count(sb, /<dxf>/g)}`);
      console.log(`  sharedStrings: file says ${sstHead(ssa)} (cells use ${usesA}, ${count(ssa, /<si>|<si /g)} entries) -> saved says ${sstHead(ssb)} (cells use ${uses}, ${count(ssb, /<si>|<si /g)} entries)`);
      const sstCount = Number(/ count="(\d+)"/.exec(one(ssb, /<sst [^>]*>/))?.[1] ?? -1);
      row14 = { styles: `${sa.length} -> ${sb.length}`, cellXfs: `${xfs(sa)} -> ${xfs(sb)}`, sst: `count ${sstCount}, uses ${uses}, entries ${count(ssb, /<si>|<si /g)}` };
      for (const n of b.keys()) if (/^xl\/comments\d*\.xml$/.test(n)) {
        const ca = await textOf(a, n), cb = await textOf(b, n);
        const attrs = (x) => { const s = new Map(); for (const m of x.matchAll(/<comment ([^>]*)>/g)) { const k = [...m[1].matchAll(/([\w:]+)="/g)].map((q) => q[1]).join(","); s.set(k, (s.get(k) ?? 0) + 1); } return [...s.entries()].map(([k, v]) => `${v}x{${k}}`).join(" "); };
        if (ca !== cb) console.log(`  ${n}: ${ca ? "" : "(new part) "}root ${JSON.stringify(clip(one(ca, /<comments[^>]*>/), 160))} -> ${JSON.stringify(clip(one(cb, /<comments[^>]*>/), 160))}; authors ${JSON.stringify((ca.match(/<author>[^<]*<\/author>|<author\/>/g) ?? []).join(""))} -> ${JSON.stringify((cb.match(/<author>[^<]*<\/author>|<author\/>/g) ?? []).join(""))}; comment attributes ${attrs(ca) || "-"} -> ${attrs(cb)}; comments ${count(ca, /<comment /g)} -> ${count(cb, /<comment /g)}`);
      }
    }
    const names = changed.map((c) => c.n);
    const otherSheets = names.filter((n) => /^xl\/worksheets\/sheet/.test(n) && n !== sheetPart);
    const allowed = { type: [sheetPart, "xl/sharedStrings.xml"], below: [sheetPart, "xl/sharedStrings.xml"], fill: [sheetPart, "xl/styles.xml"], width: [sheetPart], rowheight: [sheetPart], hidecol: [sheetPart] }[edit];
    // (calcChain goes when formulas moved: Excel builds it again. A drawing whose anchors moved is written under another name.)
    const reallyDropped = dropped.filter((d) => d !== "xl/calcChain.xml" && !((edit === "row" || edit === "delrow") && /^xl\/drawings\/(_rels\/)?drawing\d+\.xml(\.rels)?$/.test(d) && added.length > 0));
    check(`${name} / ${edit}: nothing dropped${dropped.length ? ` (but ${dropped.join(", ")})` : ""}`, reallyDropped.length === 0, reallyDropped);
    // another sheet may change in its formulas' text (references into the sheet that got a row) or in its view (the tab that was active)
    const otherReal = changed.filter((c) => otherSheets.includes(c.n) && !(c.t.gone.every((t) => !t.startsWith("<") || /^<sheetView /.test(t) || /^<c r="[A-Z]+\d+"( s="\d+")?( t="\w+")?>$/.test(t)) && c.t.gone.length === c.t.added.length)).map((c) => c.n);
    check(`${name} / ${edit}: no other sheet changes${otherSheets.length ? ` but in formula text / its view (${otherSheets.join(", ")})` : ""}`, otherReal.length === 0, otherReal);
    if (allowed) {
      const cachedOnly = changed.filter((c) => otherSheets.includes(c.n) && c.t.gone.every((t) => !t.startsWith("<")) && c.t.added.every((t) => !t.startsWith("<"))).map((c) => c.n);
      check(`${name} / ${edit}: only ${allowed.join(" + ")} change${cachedOnly.length ? ` (and the cached value of a formula in ${cachedOnly.join(", ")} that reads the cell)` : ""}, nothing is added`, names.every((n) => allowed.includes(n) || cachedOnly.includes(n)) && added.length === 0, { names, added });
      const keep = ["sheetPr", "codeName", "ignoredErrors", "rowStyles", "rowCustomFormat", ...(edit === "rowheight" ? [] : ["rowHeights"]), "sharedFormulas", "sharedMasters", "formulas", "arrayFormulas", "merges", "cf", "dv", "pageSetup", "pageMargins", "headerFooter", "ext", "drawing", "legacyDrawing", "tableParts", ...(edit === "type" || edit === "below" ? ["cols"] : edit === "fill" ? ["cols", "rows", "cells", "selection"] : ["rows", "cells", "selection", "sCells", "strCells"])];
      check(`${name} / ${edit}: the sheet is patched: sheetPr / codeName, ignoredErrors, row styles, shared formulas${edit !== "type" && edit !== "below" ? ", selection" : ""} and the rest are the file's`, keep.every((k) => !differ.includes(k)), differ.filter((k) => keep.includes(k)));
      const sd = changed.find((c) => c.n === sheetPart)?.d;
      const st = changed.find((c) => c.n === sheetPart)?.t;
      if (st) check(`${name} / ${edit}: the sheet's XML differs in a few elements only`, st.gone.length <= 6 && st.added.length <= 12, { gone: st.gone.length, added: st.added.length });
    }
    table.push({ file: name, edit, what, ms, ...row14, differ, bytes: `${original.length} -> ${out.length}`, changed: changed.map((c) => `${c.n} ${c.from}->${c.to}`), dropped, added });
    await page.close();
  }
}
fs.writeFileSync(`${S}/fid14-sheet-table.json`, JSON.stringify(table, null, 1));
await browser.close();
process.exit(summary() ? 1 : 0);
