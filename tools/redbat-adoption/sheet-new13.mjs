// What @rbt/office 0.1.13 adds, in the spreadsheet editor and preview: the comment box, notes in runs.
import { launch, newPage, open, T, S, SHOTS, check, summary, saved, unzip, waitUploads } from "./lib.mjs";
import { g, sel, ready, status, frames, select, cellPoint, undo, revision } from "./ui-lib.mjs";
import fs from "node:fs";
const { browser, context } = await launch();
const file = `${T}/real-world/jumeirah-hotel.xlsx`;
const look = (runs) => (runs ?? []).map((r) => `${r.bold ? "B" : ""}${r.italic ? "I" : ""}${r.underline ? "U" : ""}[${r.fontName ?? ""} ${r.fontSize ?? ""} ${r.color ?? ""}]${JSON.stringify(r.text)}`).join(" + ");
const page = await newPage(context);
await open(page, "sheet-editor", file); await ready(page);
// a sheet with a note in runs
const found = await g(page, () => { const gr = window.__grids[0]; const names = gr.GetSheetNames(); for (let i = 0; i < names.length; i++) { let list; try { list = gr.GetComments(i); } catch { continue; } const rich = list.filter((c) => c.runs && c.runs.length > 1 && !c.thread); if (rich.length) return { sheet: i, name: names[i], all: list.length, rich: rich.map((c) => ({ row: c.row, col: c.col, address: c.address, text: c.text, runs: c.runs })) }; } return null; });
check("the file has notes in runs (GetComments gives them)", !!found && found.rich.length > 0, found && { sheet: found.name, notes: found.all, rich: found.rich.length });
const note = found.rich[0];
console.log("  note", note.address, look(note.runs));
if ((await g(page, () => window.__grids[0].GetActiveSheetIndex())) !== found.sheet) { await sel(page, `sheet-tab-${found.sheet}`).click(); await frames(page); }
await select(page, note.row, note.col); await frames(page);
let p = await cellPoint(page, note.row, note.col);
const cellAttrs = (r, c) => g(page, ([r, c]) => { const el = [...document.querySelectorAll(`.spreadsheet-cell[data-row="${r}"][data-col="${c}"]`)].find((e) => e.getBoundingClientRect().width > 0); return el ? { title: el.getAttribute("title"), comment: el.getAttribute("data-comment") } : null; }, [r, c]);
const attrs = await cellAttrs(note.row, note.col);
check("a commented cell has no title (no browser tooltip) and carries data-comment", attrs && attrs.title === null && attrs.comment.endsWith(note.text), attrs);
await page.mouse.move(p.x - 3, p.y); await page.mouse.move(p.x, p.y); await page.waitForTimeout(700);
const box = await g(page, () => [...document.querySelectorAll(".spreadsheet-comment-box")].filter((b) => b.getBoundingClientRect().width > 0).map((b) => { const r = b.getBoundingClientRect(); const cs = getComputedStyle(b); return { row: b.dataset.row, col: b.dataset.col, kind: b.dataset.kind, text: b.textContent, spans: [...b.querySelectorAll("span")].map((s) => getComputedStyle(s).fontWeight), rect: [Math.round(r.left), Math.round(r.top), Math.round(r.width), Math.round(r.height)], bg: cs.backgroundColor, color: cs.color, z: cs.zIndex, inView: r.right <= innerWidth && r.bottom <= innerHeight && r.left >= 0 && r.top >= 0 }; }));
check("hovering shows the grid's comment box, once, for that cell, the note in its runs", box.length === 1 && Number(box[0].row) === note.row && Number(box[0].col) === note.col && box[0].kind === "note" && box[0].text.includes(note.text.slice(0, 20)) && new Set(box[0].spans).size > 1, box);
check("the box is inside the window and readable (has a background of its own)", box[0]?.inView && box[0].rect[2] > 60 && box[0].rect[3] > 20 && box[0].bg !== "rgba(0, 0, 0, 0)", box[0]);
check("the host draws no popover of its own beside it", (await page.locator('[role="tooltip"], .rb-tooltip, .rb-sheet-comment-popover').count()) === 0);
await page.screenshot({ path: `${SHOTS}/sheet-comment-box.png` });
await page.mouse.move(5, 5); await page.waitForTimeout(300);
check("the box goes when the pointer leaves the cell", (await g(page, () => [...document.querySelectorAll(".spreadsheet-comment-box")].filter((b) => b.getBoundingClientRect().width > 0).length)) === 0);

// the dialog: a note in runs, left as it is
const rightClick = async (r, c, item) => { const q = await cellPoint(page, r, c); await page.mouse.click(q.x, q.y, { button: "right" }); await sel(page, "sheet-context-menu").waitFor(); await sel(page, `sheet-menu-${item}`).click(); };
const comment = (r = note.row, c = note.col) => g(page, ([r, c]) => window.__grids[0].GetComment(r, c), [r, c]);
const r0 = await revision(page);
await rightClick(note.row, note.col, "comment");
check("the dialog has the note's text, and shows it as it is formatted", (await sel(page, "sheet-comment-text").inputValue()) === note.text && (await sel(page, "sheet-comment-preview").count()) === 1 && (await sel(page, "sheet-comment-preview").locator("span").count()) >= 2, await sel(page, "sheet-comment-preview").innerText().catch(() => null));
check("no comment box under the dialog", (await g(page, () => [...document.querySelectorAll(".spreadsheet-comment-box")].filter((b) => b.getBoundingClientRect().width > 0).length)) === 0);
await page.screenshot({ path: `${SHOTS}/sheet-comment-dialog-rich.png` });
await sel(page, "sheet-dialog-ok").click(); await frames(page);
check("Save with nothing changed changes nothing (the note is not written again)", (await revision(page)) === r0 && /No changes/.test(await status(page)) && look((await comment()).runs) === look(note.runs), { revision: [r0, await revision(page)], status: await status(page) });
// text added at the end: the runs stay
await rightClick(note.row, note.col, "comment");
await sel(page, "sheet-comment-text").focus(); await page.keyboard.press("Control+End"); await page.keyboard.type(" Checked.");
await sel(page, "sheet-dialog-ok").click(); await frames(page);
let c1 = await comment();
const expectRuns = note.runs.map((r, i) => (i === note.runs.length - 1 ? { ...r, text: r.text + " Checked." } : r));
check("text typed at the end of a note in runs: the runs keep their formatting, the last one has the text", c1.text === note.text + " Checked." && look(c1.runs) === look(expectRuns), look(c1.runs));
await undo(page); await frames(page);
check("one undo brings the note back as it was", look((await comment()).runs) === look(note.runs) && (await revision(page)) === r0);
// a word replaced in the middle of the first run
await rightClick(note.row, note.col, "comment");
await sel(page, "sheet-comment-text").evaluate((el, n) => { el.focus(); el.setSelectionRange(0, n); }, Math.min(3, note.runs[0].text.length));
await page.keyboard.type("XY");
await sel(page, "sheet-dialog-ok").click(); await frames(page);
c1 = await comment();
check("text replaced at the start: it takes the look of the run it is in; the others are as they were", c1.runs && c1.runs.length === note.runs.length && c1.runs[0].text.startsWith("XY") && c1.runs[0].bold === note.runs[0].bold && look(c1.runs.slice(1)) === look(note.runs.slice(1)), look(c1.runs));
await undo(page); await frames(page);

// a new note, bold and italic from the dialog
const empty = await g(page, () => { const gr = window.__grids[0]; const u = gr.GetUsedRange(); for (let r = u.firstRow; r <= u.lastRow; r++) for (let c = u.firstCol; c <= u.lastCol; c++) { const x = gr.GetCell(r, c); if (!x.comment && !x.isMerged && x.valueType !== "empty") return [r, c]; } return [u.lastRow, u.firstCol]; });
await select(page, ...empty); await frames(page);
await rightClick(...empty, "comment");
await sel(page, "sheet-comment-text").fill("Check this figure");
check("a plain note shows no preview", (await sel(page, "sheet-comment-preview").count()) === 0 && (await sel(page, "sheet-comment-bold").getAttribute("aria-pressed")) === "false");
await sel(page, "sheet-comment-text").evaluate((el) => el.setSelectionRange(0, 5));
await sel(page, "sheet-comment-bold").click();
check("Bold with text selected: the button is pressed, the selection stays in the text, the preview shows it", (await sel(page, "sheet-comment-bold").getAttribute("aria-pressed")) === "true" && (await sel(page, "sheet-comment-text").evaluate((el) => document.activeElement === el && el.selectionStart === 0 && el.selectionEnd === 5)) && (await sel(page, "sheet-comment-preview").locator("span").first().evaluate((s) => s.textContent + "|" + s.style.fontWeight)) === "Check|700");
await sel(page, "sheet-comment-text").evaluate((el) => el.setSelectionRange(11, 17));
await page.keyboard.press("Control+i");
check("Ctrl+I in the text: Italic is pressed for the selection, Bold is not", (await sel(page, "sheet-comment-italic").getAttribute("aria-pressed")) === "true" && (await sel(page, "sheet-comment-bold").getAttribute("aria-pressed")) === "false" && /is-active/.test(await sel(page, "sheet-comment-italic").getAttribute("class")));
await page.screenshot({ path: `${SHOTS}/sheet-comment-dialog-format.png` });
await page.keyboard.press("Control+Enter"); await frames(page);
c1 = await comment(...empty);
check("the note is saved in runs: bold 'Check', plain ' this ', italic 'figure'", c1 && look(c1.runs) === 'B[  ]"Check" + [  ]" this " + I[  ]"figure"', c1 && look(c1.runs));
// taking the formatting away again makes it a plain note
await rightClick(...empty, "comment");
await sel(page, "sheet-comment-text").evaluate((el) => el.setSelectionRange(0, 5)); await sel(page, "sheet-comment-bold").click();
await sel(page, "sheet-comment-text").evaluate((el) => el.setSelectionRange(11, 17)); await sel(page, "sheet-comment-italic").click();
await sel(page, "sheet-dialog-ok").click(); await frames(page);
c1 = await comment(...empty);
check("with all formatting taken away it is a plain note (runs null)", c1.text === "Check this figure" && c1.runs === null, c1);
await undo(page); await frames(page);
check("undo: in runs again", look((await comment(...empty)).runs).startsWith("B["));
// hover the new one
p = await cellPoint(page, ...empty);
await page.mouse.move(p.x - 2, p.y); await page.mouse.move(p.x, p.y); await page.waitForTimeout(700);
const box2 = await g(page, () => [...document.querySelectorAll(".spreadsheet-comment-box")].filter((b) => b.getBoundingClientRect().width > 0).map((b) => ({ text: b.textContent, weights: [...b.querySelectorAll("span")].map((s) => `${s.textContent}:${getComputedStyle(s).fontWeight}:${getComputedStyle(s).fontStyle}`) })));
check("its box shows the runs", box2.length === 1 && box2[0].weights.some((w) => /^Check:(700|bold):normal/.test(w)) && box2[0].weights.some((w) => /^figure:400:italic/.test(w)), box2);
await page.screenshot({ path: `${SHOTS}/sheet-comment-box-new.png` });
await page.mouse.move(5, 5);

// save, read again
await sel(page, "office-editor-save").click();
const bytes = await waitUploads(page, 1, 180000);
fs.writeFileSync(`${S}/out/comments13-jumeirah-hotel.xlsx`, bytes);
saved.set("c13.xlsx", bytes);
const page2 = await newPage(context);
await open(page2, "sheet-editor", "/api/saved/c13.xlsx", { name: "jumeirah-hotel.xlsx" }); await ready(page2);
const back = await g(page2, ([sheet, e, n]) => { const gr = window.__grids[0]; const list = gr.GetComments(sheet); const at = (r, c) => list.find((x) => x.row === r && x.col === c); return { added: at(e[0], e[1]), old: at(n[0], n[1]), count: list.length }; }, [found.sheet, empty, [note.row, note.col]]);
check("saved and read again: the new note's runs and the file's own are there", back.added && look(back.added.runs) === 'B[  ]"Check" + [  ]" this " + I[  ]"figure"' && back.old && look(back.old.runs) === look(note.runs) && back.count === found.all + 1, { added: back.added && look(back.added.runs), old: back.old && look(back.old.runs), count: [found.all, back.count] });
await page2.close();
await page.close();

// a new note is the signed-in user's
{
  const pg = await newPage(context);
  pg.profile = { id: 7, username: "dana", email: "dana@example.com", displayName: "Dana Host" };
  await open(pg, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`); await ready(pg);
  const old = await g(pg, () => { const c = window.__grids[0].GetComments()[0]; return { row: c.row, col: c.col, author: c.author, text: c.text }; });
  const at = await g(pg, () => { const gr = window.__grids[0]; const u = gr.GetUsedRange(); for (let r = u.firstRow; r <= u.lastRow; r++) for (let c = u.firstCol; c <= u.lastCol; c++) { const x = gr.GetCell(r, c); if (!x.comment && !x.isMerged && x.valueType !== "empty") return [r, c]; } });
  const rc = async (r, c) => { const q = await cellPoint(pg, r, c); await pg.mouse.click(q.x, q.y, { button: "right" }); await sel(pg, "sheet-context-menu").waitFor(); await sel(pg, "sheet-menu-comment").click(); };
  await select(pg, ...at); await frames(pg);
  await rc(...at); await sel(pg, "sheet-comment-text").fill("Mine"); await pg.waitForTimeout(200); await sel(pg, "sheet-dialog-ok").click(); await frames(pg);
  const mine = await g(pg, ([r, c]) => window.__grids[0].GetComment(r, c), at);
  check("a new note has the signed-in user's name for its author", mine?.text === "Mine" && mine.author === "Dana Host", mine);
  await select(pg, old.row, old.col); await frames(pg);
  await rc(old.row, old.col); await sel(pg, "sheet-comment-text").focus(); await pg.keyboard.press("Control+End"); await pg.keyboard.type(" ok"); await sel(pg, "sheet-dialog-ok").click(); await frames(pg);
  const kept = await g(pg, ([r, c]) => window.__grids[0].GetComment(r, c), [old.row, old.col]);
  check("a note that was there keeps its author when it is edited", kept.text === old.text + " ok" && kept.author === old.author, { was: old.author, is: kept.author });
  await sel(pg, "office-editor-save").click();
  const out = await waitUploads(pg, 1, 180000);
  const xml = await (await unzip(out)).get("xl/comments1.xml").async("string");
  check("saved: the file lists both authors and each note has its own", /<author>Dana Host<\/author>/.test(xml) && new RegExp(`<author>${old.author}</author>`).test(xml) && new Set([...xml.matchAll(/<comment [^>]*authorId="(\d+)"/g)].map((m) => m[1])).size === 2, xml.slice(0, 400));
  await pg.close();
}

// the preview shows the box too
{
  const pv = await newPage(context);
  await open(pv, "sheet-preview", file);
  await pv.waitForFunction(() => window.__grids.length > 0 && document.querySelectorAll(".spreadsheet-cell").length > 20, null, { timeout: 120000 });
  await pv.waitForTimeout(500);
  if ((await pv.evaluate(() => window.__grids[0].GetActiveSheetIndex())) !== found.sheet) { await pv.evaluate((i) => window.__grids[0].SwitchSheet ? window.__grids[0].SwitchSheet(i) : window.__grids[0].SetActiveSheet(i), found.sheet); await pv.waitForTimeout(500); }
  await pv.evaluate(([r, c]) => window.__grids[0].ScrollToCell(r, c), [note.row, note.col]); await pv.waitForTimeout(300);
  const q = await pv.evaluate(([r, c]) => { const el = [...document.querySelectorAll(`.spreadsheet-cell[data-row="${r}"][data-col="${c}"]`)].map((e) => e.getBoundingClientRect()).find((x) => x.width > 0); return el ? { x: el.left + el.width / 2, y: el.top + el.height / 2 } : null; }, [note.row, note.col]);
  if (q) { await pv.mouse.move(q.x - 2, q.y); await pv.mouse.move(q.x, q.y); await pv.waitForTimeout(700); }
  const pb = await pv.evaluate(() => [...document.querySelectorAll(".spreadsheet-comment-box")].filter((b) => b.getBoundingClientRect().width > 0).map((b) => { const r = b.getBoundingClientRect(); return { kind: b.dataset.kind, text: b.textContent.slice(0, 40), inView: r.right <= innerWidth && r.bottom <= innerHeight, clippedBy: (() => { let e = b.parentElement; while (e) { const cs = getComputedStyle(e); if (/hidden|auto|scroll|clip/.test(cs.overflow)) { const pr = e.getBoundingClientRect(); if (r.right > pr.right + 1 || r.bottom > pr.bottom + 1 || r.left < pr.left - 1 || r.top < pr.top - 1) return e.className; } e = e.parentElement; } return null; })() }; }));
  check("preview: hovering a commented cell shows the box", !!q && pb.length === 1 && pb[0].kind === "note", { q, pb });
  console.log(`  NOTE preview, the cell at the bottom of the view (y=${Math.round(q.y)} of 768): box inside the window ${pb[0]?.inView}, clipped by ${pb[0]?.clippedBy}`);
  await pv.screenshot({ path: `${SHOTS}/sheet-comment-box-preview-edge.png` });
  await pv.close();
}
await browser.close();
process.exit(summary() ? 1 : 0);
