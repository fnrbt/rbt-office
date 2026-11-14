import { launch, newPage, open, T, SHOTS, check, summary, saved, unzip, listenerProbe, waitUploads } from "./lib.mjs";
import fs from "node:fs";
const S = "/tmp/redbat-adoption";
const { browser, context } = await launch();
await context.addInitScript(listenerProbe);
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const ready = async (page) => { await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 }); await page.waitForTimeout(400); };
const status = (page) => page.locator('[data-testid="office-editor-status"]').innerText();
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const focusDoc = (page) => ev(page, () => window.__word.getInputElement(window.__editors.at(-1)).focus());
const tables = (page) => ev(page, () => { const pkg = window.__word.exportDocumentPackage(window.__editors.at(-1)); return pkg.Paragraphs.filter((p) => p.Type === "Table").map((t) => t.Rows.map((r) => r.Cells.map((c) => (c.ColSpan > 1 ? "S" + c.ColSpan : "") + c.Paragraphs.map((p) => (p.Runs ?? []).map((r) => r.Text ?? "").join("")).join("/")))); });
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);

// ---------------- chart document: table tools, image, save round trip ----------------
{
  const file = `${T}/test-documents-wild/docx4j-chart.docx`;
  const page = await newPage(context);
  await open(page, "word-editor", file);
  await ready(page);
  await focusDoc(page);
  check("table tools disabled outside a table", await sel(page, "word-row-above").isDisabled());
  const toolbarOutside = await ev(page, () => document.querySelector('[data-testid="word-toolbar"]').offsetHeight);
  await page.keyboard.press("Control+End");
  await page.keyboard.type("after the chart");
  await page.keyboard.press("Enter");
  // table
  await sel(page, "word-table").click();
  await sel(page, "word-table-insert").click();
  await frame(page);
  let t = await tables(page);
  check("table inserted (3 x 3)", t.length === 1 && t[0].length === 3 && t[0][0].length === 3, t);
  // right after the insert, before anything is typed
  const ctx0 = await ev(page, () => window.__word.getCaretContext(window.__editors[0]));
  check("the caret is in the new table's first cell and the table tools are shown at once", ctx0.InTable && ctx0.Row === 0 && ctx0.Column === 0 && ctx0.RowCount === 3 && ctx0.ColumnCount === 3 && (await sel(page, "word-row-above").isEnabled()), ctx0);
  { const h = await ev(page, () => document.querySelector('[data-testid="word-toolbar"]').offsetHeight); check("the toolbar keeps its height when the caret goes into a table", h === toolbarOutside, { outside: toolbarOutside, inside: h }); }
  check("the table tools say where the caret is", (await sel(page, "word-table-position").getAttribute("title")) === "Row 1 of 3, column 1 of 3", await sel(page, "word-table-position").getAttribute("title"));
  await page.keyboard.type("a1");
  await frame(page);
  const cellA = (await ev(page, () => window.__word.getSelection(window.__editors[0]))).FocusBlockId;
  await page.keyboard.press("Tab");
  const ctx1 = await ev(page, () => window.__word.getCaretContext(window.__editors[0]));
  check("Tab moves to the next cell", ctx1.InTable && ctx1.Row === 0 && ctx1.Column === 1, ctx1);
  await frame(page);
  check("the position follows the caret (a frame later)", (await sel(page, "word-table-position").getAttribute("title")) === "Row 1 of 3, column 2 of 3");
  const cellB = (await ev(page, () => window.__word.getSelection(window.__editors[0]))).FocusBlockId;
  await page.keyboard.type("b1");
  await page.keyboard.press("Tab");
  await page.keyboard.press("Tab");
  const ctx2 = await ev(page, () => window.__word.getCaretContext(window.__editors[0]));
  check("Tab at the end of a row goes to the next row", ctx2.Row === 1 && ctx2.Column === 0, ctx2);
  await page.keyboard.press("Shift+Tab");
  await page.keyboard.press("Shift+Tab");
  const ctx3 = await ev(page, () => window.__word.getCaretContext(window.__editors[0]));
  check("Shift+Tab goes back", ctx3.Row === 0 && ctx3.Column === 1, ctx3);
  await frame(page);
  t = await tables(page);
  check("typing fills the cells Tab went to", t[0][0][0] === "a1" && t[0][0][1] === "b1", t[0][0]);
  await page.screenshot({ path: `${SHOTS}/word-table-tools-1366.png` });
  await sel(page, "word-row-above").click();
  t = await tables(page);
  check("row above", t[0].length === 4 && t[0][1][0] === "a1", t[0].map((r) => r.join(",")));
  await sel(page, "word-row-below").click();
  t = await tables(page);
  check("row below", t[0].length === 5 && t[0][1][0] === "a1" && t[0][2].join("") === "", t[0].map((r) => r.join(",")));
  await sel(page, "word-column-left").click();
  t = await tables(page);
  check("column to the left", t[0][1].length === 4 && t[0][1][0] === "a1" && t[0][1][2] === "b1", t[0][1]);
  await sel(page, "word-column-right").click();
  t = await tables(page);
  check("column to the right", t[0][1].length === 5 && t[0][1][2] === "b1" && t[0][1][3] === "", t[0][1]);
  await sel(page, "word-column-delete").click();
  t = await tables(page);
  check("delete column", t[0][1].length === 4 && !t[0][1].includes("b1"), t[0][1]);
  await page.keyboard.press("Control+z");
  t = await tables(page);
  check("delete column is one undo step", t[0][1].length === 5 && t[0][1][2] === "b1", t[0][1]);
  await sel(page, "word-row-delete").click();
  t = await tables(page);
  check("delete row", t[0].length === 4 && !t[0].some((r) => r.includes("b1")), t[0].map((r) => r.join(",")));
  await page.keyboard.press("Control+z");
  t = await tables(page);
  check("delete row undone", t[0].length === 5 && t[0][1][2] === "b1");
  // merge: select from the b1 cell to the next cell with Shift+Right across the cell boundary
  check("merge and split disabled with the caret in one plain cell", (await sel(page, "word-cells-merge").isDisabled()) && (await sel(page, "word-cell-split").isDisabled()));
  await ev(page, ([a, b]) => window.__word.setSelection(window.__editors[0], a, 0, b, 1), [cellA, cellB]);
  await frame(page);
  const canMerge = await ev(page, () => window.__word.canMergeTableCells(window.__editors[0]));
  check("merge enabled for a selection over two cells", canMerge && !(await sel(page, "word-cells-merge").isDisabled()), { canMerge, s: await ev(page, () => window.__word.getSelection(window.__editors[0])) });
  await sel(page, "word-cells-merge").click();
  t = await tables(page);
  check("cells merged (the two cells and the one between them)", t[0][1].length === 3 && t[0][1][0].startsWith("S3"), t[0][1]);
  await ev(page, (a) => window.__word.setSelection(window.__editors[0], a, 0, a, 0), cellA);
  await frame(page);
  check("split enabled in the merged cell", !(await sel(page, "word-cell-split").isDisabled()) && (await ev(page, () => window.__word.getCaretContext(window.__editors[0]).CanSplitCell)));
  await sel(page, "word-cell-split").click();
  t = await tables(page);
  check("cell split", t[0][1].length === 5 && !t[0][1].some((c) => c.startsWith("S")), t[0][1]);

  await page.keyboard.press("Control+Home");
  await frame(page);
  check("table tools disabled when the caret leaves the table", (await sel(page, "word-row-above").isDisabled()) && !(await ev(page, () => window.__word.getCaretContext(window.__editors[0]).InTable)));
  // image: insert, click selects it, a handle resizes it, undo restores
  await page.keyboard.press("Control+Home");
  await sel(page, "word-image-input").setInputFiles(`${S}/red.png`);
  await page.waitForTimeout(600);
  const findImage = () => ev(page, () => { const w = window.__word, e = window.__editors[0]; const size = w.getDocumentSize(e); for (let y = 60; y < Math.min(size.Height, 1100); y += 20) for (let x = 60; x < size.Width; x += 20) { const hit = w.getImageHitTest(e, x, y); if (hit && Math.abs(hit.Width - 200) < 1) { const v = w.documentToViewport(e, hit.X, hit.Y); const r = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); const z = w.getViewport(e).scale; return { ...hit, vx: r.left + v.X, vy: r.top + v.Y, vw: hit.Width * z, vh: hit.Height * z }; } } return null; });
  let img = await findImage();
  check("image inserted at its size", img && Math.round(img.Width) === 200 && Math.round(img.Height) === 120, img);
  await page.evaluate(() => document.querySelector('[data-testid="word-scroll"]').scrollTop = 0);
  await frame(page); await page.waitForTimeout(200); await frame(page);
  img = await findImage();
  await page.mouse.click(img.vx + img.vw / 2, img.vy + img.vh / 2);
  await frame(page);
  const s1 = await ev(page, () => window.__word.getSelection(window.__editors[0]));
  check("a click on the picture selects it", !s1.IsCollapsed && s1.AnchorBlockId === img.BlockId && Math.abs(s1.FocusOffset - s1.AnchorOffset) === 1, s1);
  await page.screenshot({ path: `${SHOTS}/word-image-selected.png` });
  await frame(page); await page.waitForTimeout(150);
  img = await findImage();
  // handle pixels: something that is not the picture's red and not white at its bottom-right corner
  await page.mouse.move(img.vx + img.vw, img.vy + img.vh);
  await page.mouse.down();
  await page.mouse.move(img.vx + img.vw + 30, img.vy + img.vh + 18, { steps: 6 });
  await page.mouse.up();
  await frame(page);
  const resized = await ev(page, (id) => { const w = window.__word, e = window.__editors[0]; const size = w.getDocumentSize(e); for (let y = 60; y < 1100; y += 20) for (let x = 60; x < size.Width; x += 20) { const hit = w.getImageHitTest(e, x, y); if (hit && hit.BlockId === id) return { w: Math.round(hit.Width), h: Math.round(hit.Height) }; } return null; }, img.BlockId);
  check("dragging the corner handle resizes the picture", resized && resized.w > 215 && resized.h > 125, resized);
  await page.keyboard.press("Control+z");
  const undone = await ev(page, (id) => { const w = window.__word, e = window.__editors[0]; const size = w.getDocumentSize(e); for (let y = 60; y < 1100; y += 20) for (let x = 60; x < size.Width; x += 20) { const hit = w.getImageHitTest(e, x, y); if (hit && hit.BlockId === id) return { w: Math.round(hit.Width), h: Math.round(hit.Height) }; } return null; }, img.BlockId);
  check("the resize is one undo step", undone && undone.w === 200 && undone.h === 120, undone);

  // save: chart parts byte-identical
  await sel(page, "office-editor-save").click();
  const bytes = await waitUploads(page, 1);
  await page.waitForTimeout(200);
  check("saved", (await status(page)).includes("Saved") && (await sel(page, "office-editor-error").count()) === 0, await status(page));
  const orig = await unzip(fs.readFileSync(file)), out = await unzip(bytes);
  const chartParts = [...orig.keys()].filter((n) => /chart|embeddings/i.test(n) && !n.endsWith("/"));
  let identical = 0; const different = [];
  for (const n of chartParts) { const a = await orig.get(n).async("nodebuffer"); const b = out.has(n) ? await out.get(n).async("nodebuffer") : null; if (b && a.equals(b)) identical++; else different.push(n); }
  check("chart parts are byte-identical in the saved file", chartParts.length > 0 && different.length === 0, { parts: chartParts, different });
  const docXml = await out.get("word/document.xml").async("string");
  check("saved document has the chart reference, the table, the picture and the text", /<c:chart /.test(docXml) && /<w:tbl>/.test(docXml) && /after the chart/.test(docXml) && [...out.keys()].some((n) => /^word\/media\//.test(n)), { chart: /<c:chart /.test(docXml), tbl: /<w:tbl>/.test(docXml) });
  saved.set("chart", bytes);
  await page.screenshot({ path: `${SHOTS}/word-chart-saved.png` });

  // dispose: a second open + close leaves nothing more behind
  await sel(page, "office-editor-close").click();
  await page.waitForFunction(() => window.__closed === 1);
  await page.waitForTimeout(300);
  const longLived = (l) => JSON.stringify(Object.fromEntries(Object.entries(l).filter(([k]) => !k.startsWith("detached:"))));
  const l1 = await ev(page, () => window.__listeners());
  await ev(page, () => window.__setOpen(true));
  await ready(page);
  await focusDoc(page);
  await page.keyboard.type("z");
  await sel(page, "office-editor-close").click();
  await sel(page, "office-editor-discard").click();
  await page.waitForFunction(() => window.__closed === 2);
  await page.waitForTimeout(300);
  const l2 = await ev(page, () => window.__listeners());
  check("unsaved-changes prompt shown on close, and a second open+close leaves no extra live listener", longLived(l1) === longLived(l2), { l1: longLived(l1), l2: longLived(l2) });
  await page.close();

  const page2 = await newPage(context);
  await open(page2, "word-editor", "/api/saved/chart", { name: "docx4j-chart.docx" });
  await ready(page2);
  const t2 = await tables(page2);
  const text = await ev(page2, () => window.__word.getText(window.__editors[0]));
  check("re-open: table, text and clean state", t2.length === 1 && t2[0].length === 5 && t2[0][1].some((c) => c.includes("a1")) && text.includes("after the chart") && (await status(page2)).includes("No changes"), { t: t2[0]?.map((r) => r.join(",")), status: await status(page2) });
  await page2.screenshot({ path: `${SHOTS}/word-chart-reopened.png` });
  await page2.close();
}
summary();
await browser.close();
