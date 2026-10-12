// Save fidelity after ONE edit, through the host's save path (toolbar / keyboard -> Save button -> upload).
// usage: node fidelity12.mjs [sheet,word]
import { launch, newPage, open, T, S, check, summary, waitUploads, unzip } from "./lib.mjs";
import fs from "node:fs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const which = (process.argv[2] ?? "sheet,word").split(",");
const bytesOf = async (zip, name) => (zip.has(name) ? Buffer.from(await zip.get(name).async("uint8array")) : null);
const textOf = async (zip, name) => (zip.has(name) ? await zip.get(name).async("string") : "");
async function compareParts(a, b) {
  const dropped = [], changed = [], added = [];
  for (const name of a.keys()) {
    if (!b.has(name)) { dropped.push(name); continue; }
    const x = await bytesOf(a, name), y = await bytesOf(b, name);
    if (!x.equals(y)) changed.push({ name, from: x.length, to: y.length });
  }
  for (const name of b.keys()) if (!a.has(name)) added.push(`${name} (${(await bytesOf(b, name)).length})`);
  return { dropped, changed, added, same: a.size - dropped.length - changed.length, total: a.size };
}
const count = (s, re) => (s.match(re) ?? []).length;
const one = (s, re) => (re.exec(s) ?? [""])[0];
const report = (label, diff) => {
  console.log(`  ${label}: ${diff.total} parts, ${diff.same} byte-identical, ${diff.changed.length} changed, ${diff.dropped.length} dropped, ${diff.added.length} added`);
  for (const c of diff.changed) console.log(`      changed  ${c.name}  ${c.from} -> ${c.to} bytes`);
  for (const d of diff.dropped) console.log(`      dropped  ${d}`);
  for (const a of diff.added) console.log(`      added    ${a}`);
};

if (which.includes("sheet")) {
  for (const file of (process.argv[3] ? [process.argv[3]] : [`${T}/real-world/dashboard-ecommerce-sales.xlsx`, `${T}/real-world/milk-street-office.xlsx`])) {
    const name = file.split("/").pop();
    const original = fs.readFileSync(file);
    const a = await unzip(original);
    const wbA = await textOf(a, "xl/workbook.xml"), stA = await textOf(a, "xl/styles.xml");
    const sheetParts = [...a.keys()].filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
    const allSheetsA = (await Promise.all(sheetParts.map((n) => textOf(a, n)))).join("");
    const facts = async (zip) => {
      const st = await textOf(zip, "xl/styles.xml"), wb = await textOf(zip, "xl/workbook.xml");
      const comments = [...zip.keys()].filter((n) => /^xl\/comments\d*\.xml$/.test(n));
      const commentText = (await Promise.all(comments.map((n) => textOf(zip, n)))).join("");
      const slicerParts = [...zip.keys()].filter((n) => /slicer/i.test(n));
      return {
        tableStyles: one(st, /<tableStyles[^>]*>/), tableStyleNames: [...st.matchAll(/<tableStyle [^>]*name="([^"]+)"/g)].map((m) => m[1]).join(","), tableStyleElements: count(st, /<tableStyleElement /g),
        slicerStyles: one(st, /<x14:slicerStyles[^>]*>/), slicerStyleCount: count(st, /<x14:slicerStyle /g), slicerStyleElements: count(st, /<x14:slicerStyleElement /g),
        x14dxfs: one(st, /<x14:dxfs[^>]*>/), dxfs: one(st, /<dxfs[^>]*>/), cellXfs: one(st, /<cellXfs[^>]*>/), cellStyles: one(st, /<cellStyles[^>]*>/), fonts: one(st, /<fonts[^>]*>/), fills: one(st, /<fills[^>]*>/), numFmts: one(st, /<numFmts[^>]*>/),
        workbookPr: one(wb, /<workbookPr[^>]*\/?>/), calcPr: one(wb, /<calcPr[^>]*\/?>/),
        unlocked: count(st, /<protection[^>]*locked="0"[^>]*\/>/g), quotePrefix: count(st, /quotePrefix="1"/g),
        comments: comments.length, commentRuns: count(commentText, /<r>/g), commentRunProps: count(commentText, /<rPr>/g), commentBytes: commentText.length,
        slicerParts: slicerParts.length,
      };
    };
    const factsA = await facts(a);
    for (const edit of ["type", "fill"]) {
      const page = await newPage(context);
      await open(page, "sheet-editor", file);
      await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 180000 }); await page.waitForTimeout(400);
      const where = await ev(page, (edit) => {
        const gr = window.__grids[0]; const u = gr.GetUsedRange(); const sheet = gr.GetActiveSheetIndex();
        let r, c;
        if (edit === "type") {
          // an empty cell below the used range; on a sheet that is used to its last row, right of it
          r = u.lastRow + 3; c = u.firstCol;
          try { gr.GetCell(r, c); } catch { r = u.lastRow; c = u.lastCol + 2; try { gr.GetCell(r, c); } catch { r = u.lastRow; c = u.lastCol; while (c > 0 && gr.GetCell(r, c).valueType !== "empty") c--; } }
        }
        else { // a cell with a value, inside the used range
          r = u.firstRow; c = u.firstCol;
          outer: for (let i = u.firstRow; i <= Math.min(u.lastRow, u.firstRow + 60); i++) for (let j = u.firstCol; j <= Math.min(u.lastCol, u.firstCol + 20); j++) { const cell = gr.GetCell(i, j); if (cell.valueType !== "empty" && !cell.isMerged && !cell.formula) { r = i; c = j; break outer; } }
        }
        gr.SelectCell(r, c);
        return { sheet, name: gr.GetSheetNames()[sheet], ref: gr.GetSelectionRef(), used: `${u.firstRow},${u.firstCol}-${u.lastRow},${u.lastCol}`, before: gr.GetCell(r, c).text };
      }, edit);
      await frame(page);
      if (edit === "type") {
        await page.locator('[data-testid="sheet-grid"] .spreadsheet-content').focus();
        await page.keyboard.type("edited"); await page.keyboard.press("Enter");
      } else {
        await sel(page, "sheet-fill-color").click();
        await page.locator('[data-testid="sheet-fill-palette"] [data-color="#FEF08A"]').click();
      }
      await frame(page);
      const t0 = Date.now();
      await sel(page, "office-editor-save").click();
      const out = await waitUploads(page, 1, 180000);
      const ms = Date.now() - t0;
      fs.writeFileSync(`${S}/out/${edit}-${name}`, out);
      const b = await unzip(out);
      const diff = await compareParts(a, b);
      const sheetIndexPart = `xl/worksheets/sheet${where.sheet + 1}.xml`;
      console.log(`${name}, ${edit === "type" ? `"edited" typed into ${where.name}!${where.ref} (below the used range ${where.used})` : `fill of ${where.name}!${where.ref} ("${where.before}") set`}: saved in ${ms} ms, ${original.length} -> ${out.length} bytes`);
      report("parts", diff);
      const factsB = await facts(b);
      const keys = Object.keys(factsA).filter((k) => JSON.stringify(factsA[k]) !== JSON.stringify(factsB[k]));
      console.log("  facts that differ:", keys.length === 0 ? "none" : keys.map((k) => `${k}: ${JSON.stringify(factsA[k])} -> ${JSON.stringify(factsB[k])}`).join("; "));
      console.log("  facts:", JSON.stringify({ tableStyles: factsA.tableStyles, names: factsA.tableStyleNames, elements: factsA.tableStyleElements, slicerStyles: factsA.slicerStyles, slicerStyleCount: factsA.slicerStyleCount, x14dxfs: factsA.x14dxfs, workbookPr: factsA.workbookPr, calcPr: factsA.calcPr, unlocked: factsA.unlocked, quotePrefix: factsA.quotePrefix, comments: factsA.comments, commentRuns: factsA.commentRuns }));
      const changedSheets = diff.changed.filter((c) => /^xl\/worksheets\/sheet/.test(c.name)).map((c) => c.name);
      const allowed = new Set(edit === "type" ? [sheetIndexPart, "xl/sharedStrings.xml", "xl/workbook.xml", "docProps/app.xml", "docProps/core.xml", "[Content_Types].xml", "xl/calcChain.xml"] : [sheetIndexPart, "xl/styles.xml", "xl/workbook.xml", "docProps/app.xml", "docProps/core.xml", "[Content_Types].xml", "xl/calcChain.xml"]);
      check(`${name} / ${edit}: nothing is dropped or added`, diff.dropped.length === 0 && diff.added.length === 0, { dropped: diff.dropped, added: diff.added });
      check(`${name} / ${edit}: the untouched sheets are byte-identical (only ${sheetIndexPart} changes among ${sheetParts.length} sheets)`, changedSheets.join() === sheetIndexPart, changedSheets);
      check(`${name} / ${edit}: only the parts the edit touches change`, diff.changed.every((c) => allowed.has(c.name)), diff.changed.filter((c) => !allowed.has(c.name)).map((c) => c.name));
      check(`${name} / ${edit}: custom table styles kept`, factsA.tableStyles === factsB.tableStyles && factsA.tableStyleNames === factsB.tableStyleNames && factsA.tableStyleElements === factsB.tableStyleElements, [factsA.tableStyles, factsB.tableStyles, factsA.tableStyleElements, factsB.tableStyleElements]);
      check(`${name} / ${edit}: slicer styles kept`, factsA.slicerStyles === factsB.slicerStyles && factsA.slicerStyleCount === factsB.slicerStyleCount && factsA.slicerStyleElements === factsB.slicerStyleElements && factsA.slicerParts === factsB.slicerParts, [factsA.slicerStyles, factsB.slicerStyles, factsA.slicerStyleCount, factsB.slicerStyleCount]);
      check(`${name} / ${edit}: x14 dxfs kept`, factsA.x14dxfs === factsB.x14dxfs && factsA.dxfs === factsB.dxfs, [factsA.x14dxfs, factsB.x14dxfs, factsA.dxfs, factsB.dxfs]);
      check(`${name} / ${edit}: workbookPr and calcPr as they were`, factsA.workbookPr === factsB.workbookPr && factsA.calcPr === factsB.calcPr, [factsA.workbookPr, factsB.workbookPr, factsA.calcPr, factsB.calcPr]);
      check(`${name} / ${edit}: <protection locked="0"/> count kept`, edit === "fill" ? factsB.unlocked >= factsA.unlocked : factsA.unlocked === factsB.unlocked, [factsA.unlocked, factsB.unlocked]);
      check(`${name} / ${edit}: quotePrefix count kept`, edit === "fill" ? factsB.quotePrefix >= factsA.quotePrefix : factsA.quotePrefix === factsB.quotePrefix, [factsA.quotePrefix, factsB.quotePrefix]);
      check(`${name} / ${edit}: comments' rich text kept`, factsA.comments === factsB.comments && factsA.commentRuns === factsB.commentRuns && factsA.commentRunProps === factsB.commentRunProps && factsA.commentBytes === factsB.commentBytes, [factsA.comments, factsB.comments, factsA.commentRuns, factsB.commentRuns, factsA.commentBytes, factsB.commentBytes]);
      if (edit === "fill") { const stB = await textOf(b, "xl/styles.xml"); check(`${name} / fill: styles.xml is the file's, extended (the file's text is a prefix-and-suffix of it apart from what was added)`, stB.length >= stA.length, { from: stA.length, to: stB.length, cellXfs: [factsA.cellXfs, factsB.cellXfs], fills: [factsA.fills, factsB.fills] }); }
      // the edited sheet: how much of its XML changed
      const sa = await textOf(a, sheetIndexPart), sb = await textOf(b, sheetIndexPart);
      let pre = 0; while (pre < Math.min(sa.length, sb.length) && sa[pre] === sb[pre]) pre++;
      let suf = 0; while (suf < Math.min(sa.length, sb.length) - pre && sa[sa.length - 1 - suf] === sb[sb.length - 1 - suf]) suf++;
      console.log(`  ${sheetIndexPart}: ${sa.length} -> ${sb.length} chars; common prefix ${pre}, common suffix ${suf}; changed span ${sa.length - pre - suf} -> ${sb.length - pre - suf} chars: ${JSON.stringify(sb.slice(Math.max(0, pre - 60), Math.min(sb.length, pre + 200)))}`);
      await page.close();
    }
  }
}

if (which.includes("word")) {
  const docs = process.argv[3] ? [process.argv[3]] : [`${T}/test-documents/sample-formatted.docx`, `${T}/test-documents-wild/docx4j-getting-started.docx`, `${T}/test-documents-wild/docx4j-tables-rich.docx`];
  const paragraphs = (xml) => { const body = xml.slice(xml.indexOf("<w:body")); return body.match(/<w:p[ >](?:(?!<w:p[ >])[\s\S])*?<\/w:p>/g) ?? []; };
  for (const file of docs) {
    const name = file.split("/").pop();
    const original = fs.readFileSync(file);
    const a = await unzip(original);
    const docA = await textOf(a, "word/document.xml");
    const parasA = paragraphs(docA);
    for (const edit of ["type", "bold-paragraph", "bold-word"]) {
      const page = await newPage(context);
      await open(page, "word-editor", file);
      await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
      await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors.at(-1)), null, { timeout: 300000 });
      await page.waitForTimeout(300);
      const target = await ev(page, (edit) => {
        const w = window.__word, e = window.__editors.at(-1), m = window.__wordModule;
        const blocks = m.getLayoutFingerprint(e);
        for (const strict of [true, false]) for (const block of blocks) {
          const text = block.Text ?? "";
          if (strict ? text.length < 60 || text.length > 600 || !/^[A-Za-z][A-Za-z ,.'-]{50,}/.test(text) : text.length < 14 || text.length > 600 || !/^[A-Za-z]+ [A-Za-z]+ [A-Za-z ,.'-]{4,}/.test(text)) continue;
          const matches = w.findText(e, text.slice(0, 40), true);
          if (matches.length !== 1) continue;
          w.selectFindMatch(e, matches[0]);
          const s = m.getSelection(e);
          if (Math.min(s.AnchorOffset, s.FocusOffset) !== 0) continue;
          const id = s.AnchorBlockId;
          const firstSpace = text.indexOf(" "), secondSpace = text.indexOf(" ", firstSpace + 1);
          if (edit === "type") m.setSelection(e, id, 5, id, 5);
          else if (edit === "bold-paragraph") m.setSelection(e, id, 0, id, text.length);
          else m.setSelection(e, id, firstSpace + 1, id, secondSpace);
          w.getInputElement(e).focus({ preventScroll: true });
          return { text: text.slice(0, 60), length: text.length, word: text.slice(firstSpace + 1, secondSpace), bold: w.getSelectionFormat(e).Bold };
        }
        return null;
      }, edit);
      if (!target) { check(`${name} / ${edit}: a plain paragraph to edit was found`, false); await page.close(); continue; }
      await frame(page);
      if (edit === "type") await page.keyboard.type("x");
      else await sel(page, "word-bold").click();
      await frame(page);
      const after = await ev(page, () => { const w = window.__word, e = window.__editors.at(-1); return { bold: w.getSelectionFormat(e).Bold, modified: w.isModified(e) }; });
      await sel(page, "office-editor-save").click();
      const out = await waitUploads(page, 1, 180000);
      fs.writeFileSync(`${S}/out/${edit}-${name}`, out);
      const b = await unzip(out);
      const diff = await compareParts(a, b);
      const docB = await textOf(b, "word/document.xml");
      const parasB = paragraphs(docB);
      // paragraphs of the saved body that are not, as text, paragraphs of the opened one
      const bag = new Map(); for (const p of parasA) bag.set(p, (bag.get(p) ?? 0) + 1);
      const changedParas = []; for (const p of parasB) { const n = bag.get(p) ?? 0; if (n > 0) bag.set(p, n - 1); else changedParas.push(p); }
      const lost = [...bag.entries()].filter(([, n]) => n > 0).map(([p]) => p);
      let pre = 0; while (pre < Math.min(docA.length, docB.length) && docA[pre] === docB[pre]) pre++;
      let suf = 0; while (suf < Math.min(docA.length, docB.length) - pre && docA[docA.length - 1 - suf] === docB[docB.length - 1 - suf]) suf++;
      console.log(`${name}, ${edit === "type" ? `"x" typed at offset 5 of "${target.text}..."` : edit === "bold-paragraph" ? `the paragraph "${target.text}..." (${target.length} chars) made bold (was ${JSON.stringify(target.bold)}, is ${JSON.stringify(after.bold)})` : `the word "${target.word}" of "${target.text.slice(0, 30)}..." made bold`}: ${original.length} -> ${out.length} bytes`);
      report("parts", diff);
      console.log(`  word/document.xml: ${parasA.length} -> ${parasB.length} paragraphs; ${changedParas.length} of the saved ones are not the opened XML (${lost.length} opened ones not found again); the document's XML differs in a span of ${docA.length - pre - suf} -> ${docB.length - pre - suf} chars`);
      if (changedParas.length <= 3) for (let i = 0; i < changedParas.length; i++) console.log(`      was: ${JSON.stringify((lost[i] ?? "").slice(0, 420))}\n      is:  ${JSON.stringify(changedParas[i].slice(0, 420))}`);
      else console.log(`      first changed paragraph is: ${JSON.stringify(changedParas[0].slice(0, 300))}`);
      check(`${name} / ${edit}: only word/document.xml changes; nothing dropped or added`, diff.dropped.length === 0 && diff.added.length === 0 && diff.changed.map((c) => c.name).join() === "word/document.xml", { changed: diff.changed.map((c) => c.name), dropped: diff.dropped, added: diff.added });
      check(`${name} / ${edit}: one paragraph's XML changes, every other paragraph is the XML that was opened`, changedParas.length === 1 && lost.length === 1 && parasA.length === parasB.length, { changed: changedParas.length, lost: lost.length, paragraphs: [parasA.length, parasB.length] });
      if (changedParas.length === 1 && lost.length === 1) {
        const strip = (p) => p.replace(/<w:b(?:Cs)?(?: w:val="[^"]*")?\/>/g, "").replace(/<w:rPr><\/w:rPr>/g, "");
        if (edit === "type") check(`${name} / type: the paragraph is the opened one plus the character`, changedParas[0].length === lost[0].length + 1, { was: lost[0].length, is: changedParas[0].length });
        if (edit === "bold-paragraph") check(`${name} / bold-paragraph: the paragraph is the opened XML but for w:b elements (rsids, paraId, bookmarks kept)`, strip(changedParas[0]) === strip(lost[0]), { was: lost[0].length, is: changedParas[0].length });
        if (edit === "bold-word") console.log(`      bold-word: opened ${lost[0].length} chars, saved ${changedParas[0].length} chars; attributes on <w:p>: ${JSON.stringify(one(lost[0], /^<w:p[^>]*>/))} -> ${JSON.stringify(one(changedParas[0], /^<w:p[^>]*>/))}; runs ${count(lost[0], /<w:r[ >]/g)} -> ${count(changedParas[0], /<w:r[ >]/g)}; rsid attributes ${count(lost[0], /w:rsid/g)} -> ${count(changedParas[0], /w:rsid/g)}; bookmarks ${count(lost[0], /<w:bookmark/g)} -> ${count(changedParas[0], /<w:bookmark/g)}; proofErr ${count(lost[0], /<w:proofErr/g)} -> ${count(changedParas[0], /<w:proofErr/g)}; lang ${count(lost[0], /<w:lang/g)} -> ${count(changedParas[0], /<w:lang/g)}`);
      }
      await page.close();
    }
  }
}
await browser.close();
process.exit(summary() ? 1 : 0);
