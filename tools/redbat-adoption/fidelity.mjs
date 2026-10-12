// Save fidelity through the host's real save path: OfficeEditorFrame -> serialize -> (mocked) upload.
// An unedited document cannot be saved from the UI (Save is disabled while clean), so the test
// enables the frame's own Save button and clicks it: the path from the click on is the product's.
import { launch, newPage, open, T, S, check, summary, waitUploads, unzip, saved } from "./lib.mjs";
import fs from "node:fs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const status = (page) => sel(page, "office-editor-status").innerText();
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const forceSave = async (page) => { await ev(page, () => { const b = document.querySelector('[data-testid="office-editor-save"]'); b.disabled = false; b.removeAttribute("disabled"); b.click(); }); };
const which = (process.argv[2] ?? "word,sheet").split(",");

const wordReady = async (page) => {
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors.at(-1)), null, { timeout: 300000 });
  await page.waitForTimeout(300);
};
const pages = (page) => ev(page, () => window.__word.getPageCount(window.__editors.at(-1)));
const text = async (zip, name) => (zip.has(name) ? await zip.get(name).async("string") : null);
const bytes = async (zip, name) => (zip.has(name) ? Buffer.from(await zip.get(name).async("uint8array")) : null);

/** The parts of `a` that `b` lacks or has with other bytes; and the ones only `b` has. */
async function compareParts(a, b) {
  const dropped = [], changed = [], added = [];
  for (const name of a.keys()) {
    if (!b.has(name)) { dropped.push(name); continue; }
    const x = await bytes(a, name), y = await bytes(b, name);
    if (!x.equals(y)) changed.push(`${name} (${x.length} -> ${y.length})`);
  }
  for (const name of b.keys()) if (!a.has(name)) added.push(name);
  return { dropped, changed, added, same: a.size - dropped.length - changed.length };
}

/** Where two texts differ, widened to the paragraphs that hold the difference. */
function paragraphDiff(a, b) {
  let pre = 0;
  const max = Math.min(a.length, b.length);
  while (pre < max && a[pre] === b[pre]) pre++;
  let suf = 0;
  while (suf < max - pre && a[a.length - 1 - suf] === b[b.length - 1 - suf]) suf++;
  const widen = (s) => {
    const from = Math.max(s.lastIndexOf("<w:p>", pre), s.lastIndexOf("<w:p ", pre));
    const endAt = Math.max(pre, s.length - suf);
    const close = s.indexOf("</w:p>", Math.max(from, endAt - 6));
    const to = close < 0 ? s.length : close + 6;
    const piece = s.slice(from, to);
    return { from, to, piece, paragraphs: (piece.match(/<w:p[ >]/g) ?? []).length };
  };
  return { pre, suf, a: widen(a), b: widen(b), changedA: a.length - pre - suf, changedB: b.length - pre - suf };
}

if (which.includes("word")) {
  const docs = [
    `${T}/test-documents-wild/docx4j-getting-started.docx`,
    `${T}/test-documents/sample-formatted.docx`,
    `${T}/test-documents-wild/docx4j-tables-rich.docx`,
    `${T}/test-documents-wild/phpword-sample-11-readword2007.docx`,
    `${T}/repros/word/landscape-section.docx`,
  ];
  for (const file of docs) {
    const name = file.split("/").pop();
    console.log(`--- ${name}`);
    const original = await unzip(fs.readFileSync(file));
    const page = await newPage(context);
    await open(page, "word-editor", file);
    await wordReady(page);
    const pages0 = await pages(page);
    check(`${name}: clean after load`, (await status(page)).includes("No changes"), await status(page));

    // 1. an unedited save
    await forceSave(page);
    const first = await waitUploads(page, 1);
    fs.writeFileSync(`${S}/out/unedited-${name}`, first);
    const zip1 = await unzip(first);
    const doc0 = await text(original, "word/document.xml");
    const doc1 = await text(zip1, "word/document.xml");
    check(`${name}: an unedited save uploads the original word/document.xml`, doc0 === doc1, doc0 === doc1 ? `${doc0.length} chars` : (() => { const d = paragraphDiff(doc0, doc1); return { lengths: [doc0.length, doc1.length], firstDifferenceAt: d.pre, was: doc0.slice(d.pre - 60, d.pre + 120), is: doc1.slice(d.pre - 60, d.pre + 120) }; })());
    const parts1 = await compareParts(original, zip1);
    console.log(`  unedited save, other parts: ${parts1.same}/${original.size} byte-identical; dropped ${JSON.stringify(parts1.dropped)}; changed ${JSON.stringify(parts1.changed)}; added ${JSON.stringify(parts1.added)}`);
    check(`${name}: an unedited save drops no part`, parts1.dropped.length === 0, parts1.dropped);
    check(`${name}: still clean after the unedited save`, !(await status(page)).includes("Unsaved"), await status(page));

    // 2. one character
    await ev(page, () => window.__word.getInputElement(window.__editors.at(-1)).focus());
    await page.keyboard.press("Control+Home");
    await page.keyboard.type("Q");
    await frame(page);
    await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors.at(-1)), null, { timeout: 300000 });
    const pagesTyped = await pages(page);
    check(`${name}: dirty after one character`, (await status(page)).includes("Unsaved"), await status(page));
    await page.keyboard.press("Control+s");
    const second = await waitUploads(page, 2);
    fs.writeFileSync(`${S}/out/typed-${name}`, second);
    const zip2 = await unzip(second);
    const doc2 = await text(zip2, "word/document.xml");
    // The root element is looked at on its own: the writer may declare a namespace more on it.
    const body = (s) => s.slice(s.indexOf("<w:body"));
    const root = (s) => s.slice(0, s.indexOf("<w:body"));
    check(`${name}: after typing one character the root element is as it was`, root(doc0) === root(doc2), root(doc0) === root(doc2) ? "" : { was: root(doc0).slice(-300), is: root(doc2).slice(-300) });
    const d = paragraphDiff(body(doc0), body(doc2));
    const outside = body(doc0).slice(0, d.a.from) === body(doc2).slice(0, d.b.from) && body(doc0).slice(d.a.to) === body(doc2).slice(d.b.to);
    check(`${name}: after typing one character the difference is inside one paragraph`, doc0 !== doc2 && d.a.paragraphs === 1 && d.b.paragraphs === 1 && outside && /Q/.test(d.b.piece), { paragraphsOld: d.a.paragraphs, paragraphsNew: d.b.paragraphs, outsideSame: outside, changedCharsOld: d.changedA, changedCharsNew: d.changedB, paragraphOld: d.a.piece.length, paragraphNew: d.b.piece.length, of: doc0.length });
    if (!(d.a.paragraphs === 1 && d.b.paragraphs === 1)) console.log("  old:", d.a.piece.slice(0, 700), "\n  new:", d.b.piece.slice(0, 700));
    const parts2 = await compareParts(original, zip2);
    console.log(`  one-character save, parts: ${parts2.same}/${original.size} byte-identical; dropped ${JSON.stringify(parts2.dropped)}; changed ${JSON.stringify(parts2.changed)}; added ${JSON.stringify(parts2.added)}`);

    // 3. re-open both
    for (const [label, data, expected] of [["unedited", first, pages0], ["one character", second, pagesTyped]]) {
      const id = `${name}-${label.replace(/ /g, "-")}`;
      saved.set(id, data);
      const again = await newPage(context);
      await open(again, "word-editor", `/api/saved/${id}`, { name });
      await wordReady(again);
      const n = await pages(again);
      check(`${name}: the ${label} save re-opens with the page count it had before the save`, n === expected, { before: expected, reopened: n });
      await again.close();
    }
    await page.close();
  }
}

if (which.includes("sheet")) {
  const books = [`${T}/real-world/milk-street-office.xlsx`, `${T}/real-world/dashboard-ecommerce-sales.xlsx`, `${T}/real-world/chart-zoo.xlsx`, `${T}/real-world/conditional-formatting-tutorial.xlsx`];
  for (const file of books) {
    const name = file.split("/").pop();
    console.log(`--- ${name}`);
    const original = await unzip(fs.readFileSync(file));
    const page = await newPage(context);
    await open(page, "sheet-editor", file);
    await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 300000 });
    await page.waitForTimeout(400);
    check(`${name}: clean after load`, (await status(page)).includes("No changes"), await status(page));
    await forceSave(page);
    const data = await waitUploads(page, 1, 300000);
    fs.writeFileSync(`${S}/out/unedited-${name}`, data);
    const zip = await unzip(data);
    const parts = await compareParts(original, zip);
    console.log(`  unedited save: ${parts.same}/${original.size} parts byte-identical\n  dropped: ${JSON.stringify(parts.dropped)}\n  changed: ${JSON.stringify(parts.changed)}\n  added: ${JSON.stringify(parts.added)}`);
    check(`${name}: an unedited save keeps every part the original had`, parts.dropped.length === 0, parts.dropped);
    check(`${name}: still clean after the unedited save`, !(await status(page)).includes("Unsaved"), await status(page));
    // the saved workbook opens with the same sheets and used ranges
    const shape = (p) => ev(p, () => { const g = window.__grids.at(-1); return { sheets: g.GetSheetInfos().map((s) => [s.name, s.visibility]), used: g.GetUsedRange(), tables: g.GetTables().map((t) => [t.name, t.range]), names: g.GetDefinedNames().length, cf: g.GetConditionalFormats().length, drawings: g.GetDrawings().length }; });
    const before = await shape(page);
    saved.set(name, data);
    const again = await newPage(context);
    await open(again, "sheet-editor", `/api/saved/${name}`, { name });
    await again.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 300000 });
    await again.waitForTimeout(400);
    const after = await shape(again);
    check(`${name}: the saved workbook re-opens with the same sheets, used range, tables, names, rules and drawings`, JSON.stringify(before) === JSON.stringify(after), JSON.stringify(before) === JSON.stringify(after) ? { sheets: before.sheets.length, tables: before.tables.length, names: before.names, cf: before.cf, drawings: before.drawings } : { before, after });
    await again.close();
    await page.close();
  }
}
await browser.close();
process.exit(summary() ? 1 : 0);
