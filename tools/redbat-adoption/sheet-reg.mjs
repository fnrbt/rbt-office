import { launch, newPage, open, T, SHOTS, check, summary, saved, unzip, listenerProbe, waitUploads } from "./lib.mjs";
import fs from "node:fs";
const { browser, context } = await launch();
await context.addInitScript(listenerProbe);
const g = (page, fn, arg) => page.evaluate(fn, arg);
const ready = async (page) => { await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 180000 }); await page.waitForTimeout(300); };
const status = (page) => page.locator('[data-testid="office-editor-status"]').innerText();
// The four edges of a cell as the component reports them (GetCellEdges): the style name, "grid" or "-".
const edges = (page, r, c) => g(page, ([r, c]) => { const e = window.__grids.at(-1).GetCellEdges(r, c); const d = (x) => (x == null ? "-" : x === "grid" ? "grid" : `${x.Style}/${x.Width}/${x.Color}`); return [d(e.Top), d(e.Right), d(e.Bottom), d(e.Left)]; }, [r, c]);
const formula = (page) => page.locator('[data-testid="sheet-formula"]').inputValue();
// "One undo step": after the edit, a single Undo gives back exactly what `look` saw before it and Redo the edit again.
// (Undo lives in the workbook session since 0.1.32; the model's own History is no longer filled.)
const undoes = (page) => g(page, () => window.__grids.at(-1).Undo());
const redoes = (page) => g(page, () => window.__grids.at(-1).Redo());
const oneStep = async (page, look, before) => { const after = JSON.stringify(await look()); await undoes(page); const back = JSON.stringify(await look()); await redoes(page); const again = JSON.stringify(await look()); return { ok: after !== before && back === before && again === after, before, after, back, again }; };
const boldCount = (page) => g(page, () => (window.__grids.at(-1).GetSelectionHtml().match(/font-weight:\s*bold/g) || []).length);

// ---------- milk-street: formatting, formula bar, clipboard, structure, save ----------
{
  const page = await newPage(context);
  await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`);
  await ready(page);
  check("initial status clean", (await status(page)).includes("No changes"), await status(page));
  check("undo disabled initially", await page.locator('[data-testid="sheet-undo"]').isDisabled());

  // Range formatting: 20 x 12 = 240 cells (over the old 100-cell cap), one undo step.
  await g(page, () => window.__grids[0].SelectRange(30, 2, 49, 13));
  const before = await boldCount(page);
  await page.click('[data-testid="sheet-bold"]');
  const after = await boldCount(page);
  check("bold applied to the whole 240-cell range", after === 240, { before, after });
  check("no cap notice", (await page.locator('[data-testid="sheet-notice"]').count()) === 0);
  check("dirty after format", (await status(page)).includes("Unsaved"));
  await page.click('[data-testid="sheet-undo"]');
  check("one undo restores every cell", (await boldCount(page)) === before, { now: await boldCount(page), before });
  check("undo stack empty after a single undo", await page.locator('[data-testid="sheet-undo"]').isDisabled());
  await page.click('[data-testid="sheet-redo"]');
  check("redo re-applies", (await boldCount(page)) === 240);
  await page.click('[data-testid="sheet-undo"]');

  // Fill colour and border on a range, one undo step each.
  await g(page, () => window.__grids[0].SelectRange(60, 20, 62, 22));
  await page.click('[data-testid="sheet-fill-color"]');
  await page.click('[data-testid="sheet-fill-palette"] [data-color="#FEF08A"]');
  let html = await g(page, () => window.__grids[0].GetSelectionHtml());
  check("fill applied to 9 cells", (html.match(/background-color:\s*#FEF08A/gi) || []).length === 9, (html.match(/background-color/g) || []).length);
  await page.click('[data-testid="sheet-border"]');
  await page.click('[data-testid="sheet-border-line-medium"]');
  const lookBorder = async () => [await edges(page, 60, 20), await edges(page, 61, 21), await edges(page, 62, 22)];
  const beforeBorder = JSON.stringify(await lookBorder());
  await page.click('[data-testid="sheet-border-all"]');
  const borderStep = await oneStep(page, lookBorder, beforeBorder);
  check("border on the range is one undo step", borderStep.ok, borderStep);
  check("border menu closed after applying", (await page.locator('[data-testid="sheet-border-menu"]').count()) === 0);
  await g(page, () => window.__grids[0].SelectCell(0, 0));
  await page.locator('[data-testid="sheet-grid"]').screenshot({ path: `${SHOTS}/sheet-border.png` });

  // Formula bar <-> in-cell edit.
  await g(page, () => window.__grids[0].SelectCell(60, 24));
  await page.locator(".spreadsheet-content").first().focus();
  await page.keyboard.type("abc");
  check("bar mirrors in-cell typing", (await formula(page)) === "abc", await formula(page));
  check("status dirty while typing", (await status(page)).includes("Unsaved"));
  await page.keyboard.type("d");
  check("bar follows each keystroke", (await formula(page)) === "abcd");
  await page.keyboard.press("Escape");
  check("escape cancels and bar shows the cell again", (await formula(page)) === "" && !(await g(page, () => window.__grids[0].IsEditing())));
  check("editor still open after Escape in a cell", (await page.locator('[data-testid="office-editor"]').count()) === 1);
  await page.click('[data-testid="sheet-formula"]');
  check("clicking the bar starts the entry", await g(page, () => window.__grids[0].IsEditing()));
  check("bar keeps the keyboard", await g(page, () => document.activeElement?.getAttribute("data-testid") === "sheet-formula"));
  await page.keyboard.type("=1+2");
  const inCell = await g(page, () => document.querySelector(".spreadsheet-edit-input").value);
  check("cell edit box mirrors the bar", inCell === "=1+2", inCell);
  check("edit box visible on the cell", await g(page, () => getComputedStyle(document.querySelector(".spreadsheet-edit-input")).display !== "none"));
  await page.keyboard.press("Enter");
  const committed = await g(page, () => { const gr = window.__grids[0]; gr.SelectCell(60, 24); return { text: gr.GetSelectionText(), f: gr.GetCurrentCellFormula(), editing: gr.IsEditing() }; });
  check("Enter in the bar commits the formula", committed.text.trim() === "3" && committed.f === "=1+2" && !committed.editing, committed);
  check("bar shows the formula of the cell", (await formula(page)) === "=1+2");
  // type in cell, continue in bar, commit by clicking another cell
  await g(page, () => window.__grids[0].SelectCell(61, 24));
  await page.locator(".spreadsheet-content").first().focus();
  await page.keyboard.type("12");
  await page.click('[data-testid="sheet-formula"]');
  await page.keyboard.press("End");
  await page.keyboard.type("3");
  check("entry continues in the bar", (await formula(page)) === "123" && (await g(page, () => document.querySelector(".spreadsheet-edit-input").value)) === "123");
  await page.keyboard.press("Escape");
  check("Escape in the bar cancels", !(await g(page, () => window.__grids[0].IsEditing())) && (await formula(page)) === "");
  check("editor still open after Escape in the bar", (await page.locator('[data-testid="office-editor"]').count()) === 1);
  await page.keyboard.type("77");
  await page.click('[data-testid="sheet-bold"]');
  const c2 = await g(page, () => { const gr = window.__grids[0]; return { text: gr.GetSelectionText().trim(), html: gr.GetSelectionHtml() }; });
  check("a toolbar command commits the entry first", c2.text === "77" && /font-weight:\s*bold/.test(c2.html), c2.text);

  // Clipboard: copy -> system clipboard as TSV, paste within and across sheets.
  await g(page, () => { const gr = window.__grids[0]; gr.EnterText(70, 20, "a"); gr.EnterText(70, 21, "b"); gr.EnterText(71, 20, "1"); gr.EnterText(71, 21, "=U72*2"); gr.SelectRange(70, 20, 71, 21); });
  await page.click('[data-testid="sheet-copy"]');
  await page.waitForTimeout(200);
  const clip = await g(page, () => navigator.clipboard.readText());
  check("system clipboard holds TSV", clip.replace(/\r/g, "").trim() === "a\tb\n1\t2", JSON.stringify(clip));
  const clipHtml = await g(page, async () => { const items = await navigator.clipboard.read(); for (const it of items) if (it.types.includes("text/html")) return (await (await it.getType("text/html")).text()).slice(0, 200); return ""; });
  check("system clipboard holds an HTML table", /<table/i.test(clipHtml), clipHtml.slice(0, 80));
  await g(page, () => window.__grids[0].SelectCell(74, 20));
  await page.click('[data-testid="sheet-paste"]');
  await page.waitForTimeout(300);
  let pasted = await g(page, () => { const gr = window.__grids[0]; gr.SelectRange(74, 20, 75, 21); const t = gr.GetSelectionText(); gr.SelectCell(75, 21); return { t, f: gr.GetCurrentCellFormula() }; });
  check("paste within the sheet", pasted.t.replace(/\r/g, "").trim() === "a\tb\n1\t2", JSON.stringify(pasted));
  check("pasted formula is relative", pasted.f === "=U76*2", pasted.f);
  await page.locator('[data-testid="sheet-tabs"], [data-testid="spreadsheet-preview-tabs"]').getByText("Monthly", { exact: true }).click();
  check("sheet switched", (await g(page, () => window.__grids[0].GetActiveSheetIndex())) === 1);
  await g(page, () => window.__grids[0].SelectCell(200, 30));
  await page.click('[data-testid="sheet-paste"]');
  await page.waitForTimeout(300);
  pasted = await g(page, () => { const gr = window.__grids[0]; gr.SelectRange(200, 30, 201, 31); return gr.GetSelectionText(); });
  const cells = (t) => t.replace(/\r/g, "").trim().split(/[\t\n]/).map((c) => c.trim()).join(",");
  check("paste across sheets", cells(pasted) === "a,b,1,2", JSON.stringify(pasted));
  // keyboard copy / paste
  await g(page, () => window.__grids[0].SelectRange(200, 30, 200, 31));
  await page.locator(".spreadsheet-content").first().focus();
  await page.keyboard.press("Control+c");
  await page.waitForTimeout(150);
  check("Ctrl+C puts TSV on the clipboard", (await g(page, () => navigator.clipboard.readText())).replace(/\r/g, "").trim() === "a\tb");
  await g(page, () => window.__grids[0].SelectCell(204, 30));
  await page.locator(".spreadsheet-content").first().focus();
  await page.keyboard.press("Control+v");
  await page.waitForTimeout(300);
  check("Ctrl+V pastes", (await g(page, () => { const gr = window.__grids[0]; gr.SelectRange(204, 30, 204, 31); return gr.GetSelectionText(); })).trim() === "a\tb");
  // cut
  await g(page, () => window.__grids[0].SelectRange(204, 30, 204, 31));
  await page.click('[data-testid="sheet-cut"]');
  await page.waitForTimeout(200);
  check("cut clears the cells", (await g(page, () => window.__grids[0].GetSelectionText())).trim() === "");
  await page.locator('[data-testid="sheet-tabs"], [data-testid="spreadsheet-preview-tabs"]').getByText("Annual", { exact: true }).click();

  // Structure edits.
  await g(page, () => { const gr = window.__grids[0]; gr.EnterText(80, 20, "anchor"); gr.SelectCell(80, 20); });
  await page.click('[data-testid="sheet-menu-button-insert"]'); await page.click('[data-testid="sheet-insert-row"]');
  let at = await g(page, () => { const gr = window.__grids[0]; gr.SelectCell(81, 20); return gr.GetSelectionText().trim(); });
  check("insert row shifts cells down", at === "anchor", at);
  await g(page, () => window.__grids[0].SelectCell(80, 20));
  await page.click('[data-testid="sheet-menu-button-insert"]'); await page.click('[data-testid="sheet-delete-row"]');
  at = await g(page, () => { const gr = window.__grids[0]; gr.SelectCell(80, 20); return gr.GetSelectionText().trim(); });
  check("delete row shifts cells up", at === "anchor", at);
  await page.click('[data-testid="sheet-menu-button-insert"]'); await page.click('[data-testid="sheet-insert-column"]');
  at = await g(page, () => { const gr = window.__grids[0]; gr.SelectCell(80, 21); return gr.GetSelectionText().trim(); });
  check("insert column shifts cells right", at === "anchor", at);
  await g(page, () => window.__grids[0].SelectCell(80, 20));
  await page.click('[data-testid="sheet-menu-button-insert"]'); await page.click('[data-testid="sheet-delete-column"]');
  at = await g(page, () => { const gr = window.__grids[0]; gr.SelectCell(80, 20); return gr.GetSelectionText().trim(); });
  check("delete column shifts cells left", at === "anchor", at);
  // merge
  await g(page, () => window.__grids[0].SelectRange(84, 20, 85, 22));
  await page.click('[data-testid="sheet-merge"]');
  check("merge button shows the merged state", (await page.locator('[data-testid="sheet-merge"]').getAttribute("aria-pressed")) === "true");
  // number format reflects
  await g(page, () => { const gr = window.__grids[0]; gr.EnterText(88, 20, "0.5"); gr.SelectCell(88, 20); });
  await page.selectOption('[data-testid="sheet-number-format"]', "percent");
  check("number format applied and shown", (await g(page, () => window.__grids[0].GetSelectionText().trim())) === "50.00%" && (await page.locator('[data-testid="sheet-number-format"]').inputValue()) === "percent");
  await page.screenshot({ path: `${SHOTS}/sheet-editor-1366.png` });

  // Save while an entry is being typed.
  await g(page, () => window.__grids[0].SelectCell(90, 20));
  await page.locator(".spreadsheet-content").first().focus();
  await page.keyboard.type("typed-at-save");
  await page.keyboard.press("Control+s");
  const bytes = await waitUploads(page, 1);
  await page.waitForTimeout(200);
  check("status Saved, no error or warning", (await status(page)).includes("Saved") && (await page.locator('[data-testid="office-editor-error"]').count()) === 0, await status(page));
  const orig = await unzip(fs.readFileSync(`${T}/real-world/milk-street-office.xlsx`));
  const out = await unzip(bytes);
  const countDT = async (z) => { let n = 0; for (const [name, f] of z) if (/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) n += ((await f.async("string")).match(/t="dataTable"/g) || []).length; return n; };
  const dtOrig = await countDT(orig), dtOut = await countDT(out);
  check("data-table formulas kept", dtOrig > 0 && dtOut === dtOrig, { dtOrig, dtOut });
  saved.set("milk", bytes);
  const mediumBorders = async (z) => ((await z.get("xl/styles.xml").async("string")).match(/<top style="medium"/g) || []).length;
  check("saved styles carry the new medium border", (await mediumBorders(out)) > (await mediumBorders(orig)), { before: await mediumBorders(orig), after: await mediumBorders(out) });

  // Dispose leaves no listeners: a second open/close cycle adds none.
  await page.click('[data-testid="office-editor-close"]');
  await page.waitForFunction(() => window.__closed === 1);
  await page.waitForTimeout(300);
  const l1 = await g(page, () => window.__listeners());
  await g(page, () => window.__setOpen(true));
  await ready(page);
  await g(page, () => window.__grids.at(-1).SelectCell(3, 3));
  await page.locator(".spreadsheet-content").first().focus();
  await page.keyboard.type("x");
  await page.keyboard.press("Escape");
  await page.click('[data-testid="office-editor-close"]');
  await page.waitForFunction(() => window.__closed === 2);
  await page.waitForTimeout(300);
  const l2 = await g(page, () => window.__listeners());
  const longLived = (l) => JSON.stringify(Object.fromEntries(Object.entries(l).filter(([k]) => !k.startsWith("detached:"))));
  check("a second open+close leaves no extra listener on window / document / other live targets", longLived(l1) === longLived(l2), { l1: longLived(l1), l2: longLived(l2) });
  console.log("  live-target listeners after close:", longLived(l2));
  check("grid DOM removed", (await page.locator(".spreadsheet-container").count()) === 0);
  await page.close();
}

// ---------- re-open the saved milk-street ----------
{
  const page = await newPage(context);
  await open(page, "sheet-editor", "/api/saved/milk", { name: "milk-street-office.xlsx" });
  await ready(page);
  const r = await g(page, () => {
    const gr = window.__grids[0]; const out = {};
    gr.SelectCell(60, 24); out.formula = gr.GetCurrentCellFormula(); out.value = gr.GetSelectionText().trim();
    gr.SelectRange(60, 20, 62, 22); out.html = gr.GetSelectionHtml();
    gr.SelectCell(90, 20); out.typed = gr.GetSelectionText().trim();
    gr.SelectCell(88, 20); out.pct = gr.GetSelectionText().trim();
    gr.SelectRange(74, 20, 75, 21); out.pasted = gr.GetSelectionText();
    gr.SelectCell(61, 24); out.bold = gr.GetSelectionHtml();
    gr.SwitchSheet(1); gr.SelectRange(200, 30, 201, 31); out.sheet2 = gr.GetSelectionText(); gr.SwitchSheet(0);
    gr.SelectCell(84, 20);
    return out;
  });
  check("re-open: formula and value", r.formula === "=1+2" && r.value === "3", r.formula + " " + r.value);
  check("re-open: fill", (r.html.match(/background-color:\s*#FEF08A/gi) || []).length === 9);
  await g(page, () => window.__grids[0].ScrollToCell(60, 20));
  await page.locator('[data-testid="sheet-grid"]').screenshot({ path: `${SHOTS}/sheet-border-reopened.png` });
  check("re-open: entry typed at save time", r.typed === "typed-at-save", r.typed);
  check("re-open: number format", r.pct === "50.00%", r.pct);
  check("re-open: pasted cells", r.pasted.replace(/\r/g, "").trim() === "a\tb\n1\t2");
  check("re-open: bold", /font-weight:\s*bold/.test(r.bold));
  check("re-open: second sheet paste", r.sheet2.replace(/\r/g, "").trim().split(/[\t\n]/).map((c) => c.trim()).join(",") === "a,b,1,2", JSON.stringify(r.sheet2));
  check("re-open: merge", (await page.locator('[data-testid="sheet-merge"]').getAttribute("aria-pressed")) === "true");
  check("re-open: clean", (await status(page)).includes("No changes"));
  await page.close();
}
summary();
await browser.close();
