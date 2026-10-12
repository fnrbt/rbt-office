// What @rbt/office 0.1.14 changes, in the spreadsheet editor and preview: sheets that grow, the 1904 date system, notes without the host naming the author again, object labels.
import { launch, newPage, open, T, S, SHOTS, check, summary, unzip, waitUploads } from "./lib.mjs";
import { g, sel, ready, status, frames, select, cellPoint, undo, revision, selection } from "./ui-lib.mjs";
import fs from "node:fs";
const { browser, context } = await launch();
const size = (page) => g(page, () => { const gr = window.__grids.at(-1); const m = gr.GetModel(); const u = gr.GetUsedRange(); return { rows: m.RowCount, cols: m.ColCount, used: u && u.range, lastRow: u ? u.lastRow : -1, lastCol: u ? u.lastCol : -1, modified: gr.IsModified(), revision: gr.GetRevision(), canUndo: gr.CanUndo ? gr.CanUndo() : null }; });

// ---- a sheet grows as the caret nears its end: not an edit
{
  const file = `${T}/real-world/milk-street-office.xlsx`;
  const page = await newPage(context);
  await open(page, "sheet-editor", file); await ready(page);
  const s0 = await size(page);
  console.log("  at open", s0);
  let changes = 0;
  await g(page, () => { window.__changes = 0; window.__grids.at(-1).OnChange(() => window.__changes++); });
  await select(page, s0.lastRow, 1); await frames(page);
  await page.locator('[data-testid="sheet-grid"] .spreadsheet-container').first().focus().catch(() => {});
  const p = await cellPoint(page, s0.lastRow, 1);
  if (p) await page.mouse.click(p.x, p.y);
  for (let i = 0; i < 30; i++) await page.keyboard.press("ArrowDown");
  await frames(page);
  const s1 = await size(page);
  const active = await selection(page);
  check("the arrow keys go below the last used row", new RegExp(`^B${s0.lastRow + 31}$`).test(active), { active, lastUsedRow: s0.lastRow + 1 });
  check("the grid has rows to go into (RowCount grew or was large enough), the used range is what it was", s1.rows >= s0.lastRow + 31 + 50 && s1.used === s0.used, { before: s0, after: s1 });
  changes = await g(page, () => window.__changes);
  check("growing is not a change: no change event, same revision, not modified, status clean, no upload, undo not offered", changes === 0 && s1.revision === s0.revision && s1.modified === false && /No changes/.test(await status(page)) && page.uploads.length === 0 && (await sel(page, "sheet-undo").isDisabled()), { changes, status: await status(page) });
  // far right
  await select(page, 5, s0.lastCol); await frames(page);
  const q = await cellPoint(page, 5, s0.lastCol);
  if (q) await page.mouse.click(q.x, q.y);
  for (let i = 0; i < 12; i++) await page.keyboard.press("ArrowRight");
  await frames(page);
  const s2 = await size(page);
  check("and to the right of the last used column, likewise", s2.cols >= s0.lastCol + 13 && s2.modified === false && /No changes/.test(await status(page)) && (await g(page, () => window.__changes)) === 0, { cols: [s0.cols, s2.cols], lastUsedCol: s0.lastCol });
  // closing a workbook that only grew asks nothing
  await sel(page, "office-editor-close").click(); await page.waitForTimeout(300);
  check("closing after only moving about asks nothing", (await sel(page, "office-editor-confirm").count()) === 0 && (await g(page, () => window.__closed)) === 1);
  await page.close();
}

// ---- typing below the last used row, a paste and a fill past the end; saved
{
  const file = `${T}/real-world/milk-street-office.xlsx`;
  const page = await newPage(context);
  await open(page, "sheet-editor", file); await ready(page);
  const s0 = await size(page);
  const row = s0.lastRow + 3;
  await select(page, row, 1); await frames(page);
  const p = await cellPoint(page, row, 1);
  check("a cell below the last used row is on screen to click", !!p, p);
  await page.mouse.click(p.x, p.y);
  await page.keyboard.type("below the end"); await page.keyboard.press("Enter"); await frames(page);
  const s1 = await size(page);
  check("typed there: the cell has it, the used range reaches it, the document is dirty", (await g(page, (r) => window.__grids.at(-1).GetCell(r, 1), row))?.value === "below the end" || JSON.stringify(await g(page, (r) => window.__grids.at(-1).GetCell(r, 1), row)).includes("below the end") && s1.lastRow === row && /Unsaved/.test(await status(page)), { cell: await g(page, (r) => window.__grids.at(-1).GetCell(r, 1), row), used: s1.used });
  // SetCell far beyond the grid
  const far = s1.rows + 500;
  await g(page, (r) => window.__grids.at(-1).SetCell(r, 2, "42"), far); await frames(page);
  const s2 = await size(page);
  check("SetCell beyond the grid enlarges the sheet", s2.rows > far && s2.lastRow === far, { rows: [s1.rows, s2.rows], far });
  // a paste that reaches beyond the end
  await select(page, 0, 0, 5, 3); await frames(page);
  const a = await cellPoint(page, 0, 0); await page.mouse.click(a.x, a.y); await select(page, 0, 0, 5, 3);
  await page.keyboard.press("Control+c"); await page.waitForTimeout(150);
  const s4 = await size(page);
  await g(page, (r) => { const gr = window.__grids.at(-1); gr.ScrollToCell(r, 0); gr.SelectCell(r, 0); }, s4.rows - 2); await frames(page);
  const b = await cellPoint(page, s4.rows - 2, 0);
  if (b) await page.mouse.click(b.x, b.y);
  await page.keyboard.press("Control+v"); await page.waitForTimeout(400); await frames(page);
  const s5 = await size(page);
  check("a paste of six rows two rows above the grid's end makes the sheet large enough", s5.rows >= s4.rows + 4 && s5.lastRow >= s4.rows + 3, { rows: [s4.rows, s5.rows], lastRow: s5.lastRow, b: !!b });
  await page.screenshot({ path: `${SHOTS}/sheet14-paste-past-end.png` });
  await sel(page, "office-editor-save").click();
  const out = await waitUploads(page, 1, 180000);
  fs.writeFileSync(`${S}/out/below14-milk-street-office.xlsx`, out);
  const zip = await unzip(out);
  const wb = await zip.get("xl/workbook.xml").async("string");
  const rel = await zip.get("xl/_rels/workbook.xml.rels").async("string");
  const active = await g(page, () => window.__grids.at(-1).GetActiveSheetIndex());
  const rid = [...wb.matchAll(/<sheet [^>]*r:id="([^"]+)"/g)][active][1];
  const part = "xl/" + new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${rid}"`).exec(rel).slice(1).find(Boolean).replace(/^\/xl\//, "");
  const xml = await zip.get(part).async("string");
  const dim = /<dimension ref="([^"]+)"/.exec(xml)?.[1];
  const rowsIn = [...xml.matchAll(/<row [^>]*\br="(\d+)"/g)].map((m) => Number(m[1]));
  check("saved: the typed row is in the sheet's part, rows in order, the dimension reaches the last row written", xml.includes(`<row r="${row + 1}"`) && rowsIn.every((r, i) => i === 0 || r > rowsIn[i - 1]) && Number(/(\d+)$/.exec(dim)[1]) === rowsIn.at(-1), { part, dim, lastRows: rowsIn.slice(-4) });
  await page.close();
}

// ---- the preview: scrolled to the end it has room below, and stays a preview
{
  const file = `${T}/real-world/milk-street-office.xlsx`;
  const page = await newPage(context);
  await open(page, "sheet-preview", file);
  await page.waitForSelector(".spreadsheet-container", { timeout: 180000 }); await page.waitForTimeout(500);
  const s0 = await size(page);
  await g(page, (r) => window.__grids.at(-1).ScrollToCell(r, 0), s0.lastRow); await frames(page); await page.waitForTimeout(200);
  await page.screenshot({ path: `${SHOTS}/sheet14-preview-end.png` });
  const s1 = await size(page);
  check("the preview scrolled to its last used row: nothing modified", s1.modified === false && s1.used === s0.used, { rows: [s0.rows, s1.rows] });
  await page.close();
}

// ---- a workbook in the 1904 date system
{
  for (const name of ["pandas-times_1904.xlsx", "poi-47504.xlsx"]) {
    const file = `${T}/test-documents-excel/${name}`;
    const zip0 = await unzip(fs.readFileSync(file));
    const page = await newPage(context);
    await open(page, "sheet-editor", file); await ready(page);
    const sys = await g(page, () => window.__grids.at(-1).GetDateSystem());
    check(`${name}: GetDateSystem is 1904`, sys === 1904, sys);
    const cells = await g(page, () => { const gr = window.__grids.at(-1); const u = gr.GetUsedRange() ?? { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }; const out = []; for (let r = u.firstRow; r <= Math.min(u.lastRow, u.firstRow + 6); r++) for (let c = u.firstCol; c <= Math.min(u.lastCol, u.firstCol + 4); c++) { const cell = gr.GetCell(r, c); if (cell && (cell.value ?? cell.text ?? "") !== "") out.push({ r, c, value: cell.value, text: cell.text ?? cell.display ?? cell.formatted, raw: cell.raw }); } return out; });
    console.log("  cells", JSON.stringify(cells).slice(0, 700));
    await page.screenshot({ path: `${SHOTS}/sheet14-1904-${name}.png` });
    // an edit and a save: the workbook still says 1904
    const u = await size(page);
    await g(page, (r) => window.__grids.at(-1).SetCell(r, 0, "note"), u.lastRow + 2); await frames(page);
    await sel(page, "office-editor-save").click();
    const out = await waitUploads(page, 1, 180000);
    fs.writeFileSync(`${S}/out/d1904-${name}`, out);
    const wb = await (await unzip(out)).get("xl/workbook.xml").async("string");
    check(`${name}: saved with date1904`, /date1904="(1|true)"/.test(wb), /<workbookPr[^>]*>/.exec(wb)?.[0]);
    await page.close();
  }
  const page = await newPage(context);
  await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`); await ready(page);
  check("an ordinary workbook is 1900", (await g(page, () => window.__grids.at(-1).GetDateSystem())) === 1900);
  await page.close();
}

// ---- notes: the author through the dialog, without the host naming it again; data-comment
{
  const file = `${T}/real-world/jumeirah-hotel.xlsx`;
  const page = await newPage(context);
  page.profile = { id: 7, username: "dana", email: "dana@example.com", displayName: "Dana Host" };
  await open(page, "sheet-editor", file); await ready(page);
  const found = await g(page, () => { const gr = window.__grids[0]; const names = gr.GetSheetNames(); for (let i = 0; i < names.length; i++) { let list; try { list = gr.GetComments(i); } catch { continue; } const own = list.filter((c) => c.author && !c.thread); if (own.length) return { sheet: i, notes: own.map((c) => ({ row: c.row, col: c.col, author: c.author, text: c.text })) }; } return null; });
  if ((await g(page, () => window.__grids[0].GetActiveSheetIndex())) !== found.sheet) { await sel(page, `sheet-tab-${found.sheet}`).click(); await frames(page); }
  const note = found.notes[0];
  const rightClick = async (r, c, item) => { const q = await cellPoint(page, r, c); await page.mouse.click(q.x, q.y, { button: "right" }); await sel(page, "sheet-context-menu").waitFor(); await sel(page, `sheet-menu-${item}`).click(); };
  const comment = (r, c) => g(page, ([r, c]) => window.__grids[0].GetComment(r, c), [r, c]);
  const calls = await g(page, () => { const gr = window.__grids[0]; window.__authors = []; for (const k of ["SetComment", "SetCommentRuns"]) { const f = gr[k]; gr[k] = function (r, c, t, author) { window.__authors.push([k, author === undefined ? "(none)" : author]); return f.apply(this, arguments); }; } return true; });
  await rightClick(note.row, note.col, "comment");
  await sel(page, "sheet-comment-text").focus(); await page.keyboard.press("Control+End"); await page.keyboard.type(" Seen.");
  await sel(page, "sheet-dialog-ok").click(); await frames(page);
  const edited = await comment(note.row, note.col);
  const said = await g(page, () => window.__authors.slice());
  check("an edited note keeps its author, and the host did not say it", edited.author === note.author && edited.text.endsWith(" Seen.") && said.length === 1 && said[0][1] === "(none)", { author: [note.author, edited.author], said });
  const attr = await g(page, ([r, c]) => [...document.querySelectorAll(`.spreadsheet-cell[data-row="${r}"][data-col="${c}"]`)].map((e) => e.getAttribute("data-comment")).find(Boolean), [note.row, note.col]);
  check("data-comment says the author once", attr && attr.split(note.author).length - 1 === 1, attr);
  // a new one, some cells away
  const free = await g(page, ([r, c]) => { const gr = window.__grids[0]; for (let d = 1; d < 30; d++) if (!gr.GetComment(r + d, c)) return [r + d, c]; return null; }, [note.row, note.col]);
  await rightClick(free[0], free[1], "comment");
  await sel(page, "sheet-comment-text").fill("Mine");
  await sel(page, "sheet-dialog-ok").click(); await frames(page);
  const mine = await comment(free[0], free[1]);
  check("a new note is the signed-in user's", mine.author === "Dana Host" && (await g(page, () => window.__authors.at(-1)[1])) === "Dana Host", mine);
  // comment box at the grid's right edge / last rows in view stays inside
  const edge = await g(page, () => { const gr = window.__grids[0]; const sc = document.querySelector('[data-testid="sheet-grid"]').getBoundingClientRect(); const cells = [...document.querySelectorAll(".spreadsheet-cell")].map((e) => ({ e, q: e.getBoundingClientRect() })).filter((x) => x.q.width > 0 && x.q.right < sc.right - 2 && x.q.bottom < sc.bottom - 30 && x.q.top > sc.top + 30); const right = cells.sort((a, b) => b.q.right - a.q.right || b.q.bottom - a.q.bottom)[0]; const r = Number(right.e.dataset.row), c = Number(right.e.dataset.col); gr.SetComment(r, c, "A note at the edge of the view that is long enough to need room of its own beside the cell.", "Edge"); return { r, c }; });
  await frames(page);
  const ep = await g(page, ([r, c]) => { const q = [...document.querySelectorAll(`.spreadsheet-cell[data-row="${r}"][data-col="${c}"]`)].map((e) => e.getBoundingClientRect()).find((q) => q.width > 0); return { x: q.left + q.width / 2, y: q.top + q.height / 2 }; }, [edge.r, edge.c]);
  await page.mouse.move(ep.x - 3, ep.y); await page.mouse.move(ep.x, ep.y); await page.waitForTimeout(700);
  const box = await g(page, () => { const sc = document.querySelector('[data-testid="sheet-grid"]').getBoundingClientRect(); return [...document.querySelectorAll(".spreadsheet-comment-box")].filter((b) => b.getBoundingClientRect().width > 0).map((b) => { const r = b.getBoundingClientRect(); return { rect: [Math.round(r.left), Math.round(r.top), Math.round(r.right), Math.round(r.bottom)], grid: [Math.round(sc.left), Math.round(sc.top), Math.round(sc.right), Math.round(sc.bottom)], inside: r.left >= sc.left - 1 && r.right <= sc.right + 1 && r.top >= sc.top - 1 && r.bottom <= sc.bottom + 1 }; }); });
  check("the comment box of a cell at the right edge and the last rows in view is inside the grid", box.length === 1 && box[0].inside, box);
  await page.screenshot({ path: `${SHOTS}/sheet14-comment-edge.png` });
  await page.close();
}

// ---- object frames carry a label that is cut, in the editor and the preview
{
  const file = `${T}/test-documents-excel/chart-zoo.xlsx`;
  for (const c of fs.existsSync(file) ? ["sheet-editor", "sheet-preview"] : []) {
    const page = await newPage(context);
    await open(page, c, file);
    await page.waitForSelector(".spreadsheet-container", { timeout: 180000 }); await page.waitForTimeout(800);
    const labels = await g(page, () => [...document.querySelectorAll(".spreadsheet-object")].filter((o) => o.getBoundingClientRect().width > 0).map((o) => { const l = o.querySelector(".spreadsheet-object-label"); const or = o.getBoundingClientRect(); const lr = l && l.getBoundingClientRect(); return { label: l ? l.textContent.slice(0, 40) : null, inside: !lr || (lr.left >= or.left - 1 && lr.right <= or.right + 1 && lr.top >= or.top - 1 && lr.bottom <= or.bottom + 1), overflow: l && getComputedStyle(l).textOverflow, lines: lr ? Math.round(lr.height) : 0 }; }));
    console.log(`  ${c}: ${labels.length} object frames`, JSON.stringify(labels.slice(0, 4)));
    check(`${c}: every object frame's label is one line inside its frame`, labels.every((l) => l.label === null || (l.inside && l.lines < 30)), labels.filter((l) => l.label !== null && !(l.inside && l.lines < 30)).slice(0, 3));
    await page.screenshot({ path: `${SHOTS}/sheet14-objects-${c}.png` });
    await page.close();
  }
}
await browser.close();
process.exit(summary() ? 1 : 0);
