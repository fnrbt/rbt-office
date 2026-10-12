// What @rbt/office 0.1.11 / 0.1.12 add, in the spreadsheet editor and preview.
import { launch, newPage, open, T, S, SHOTS, check, summary, saved, unzip } from "./lib.mjs";
import { g, sel, ready, status, frames, select, selection, menu, undo, revision } from "./ui-lib.mjs";
import fs from "node:fs";
const { browser, context } = await launch();
const text = (page, id) => sel(page, id).innerText();

// ---- a workbook opens on the sheet it was saved on
{
  const file = `${T}/real-world/dcf-stock-analysis.xlsx`;
  const parts = await unzip(fs.readFileSync(file));
  const wb = await parts.get("xl/workbook.xml").async("string");
  const activeTab = Number(/activeTab="(\d+)"/.exec(wb)?.[1] ?? 0);
  const page = await newPage(context);
  await open(page, "sheet-editor", file); await ready(page);
  const at = await g(page, () => ({ index: window.__grids[0].GetActiveSheetIndex(), selected: [...document.querySelectorAll('[data-testid="sheet-tabs"] [role="tab"][aria-selected="true"]')].map((t) => t.getAttribute("data-testid")), infos: window.__grids[0].GetSheetInfos().map((i) => [i.index, i.isActive, i.isLoaded, i.kind]) }));
  check("editor: opens on the file's active sheet, and its tab is the selected one", activeTab > 0 && at.index === activeTab && at.selected.join() === `sheet-tab-${activeTab}`, { activeTab, at });
  check("editor: opening is not a change", /No changes/.test(await status(page)), await status(page));
  await page.close();
  const preview = await newPage(context);
  await open(preview, "sheet-preview", file);
  await preview.waitForSelector('[data-testid="spreadsheet-preview-tabs"]', { timeout: 120000 });
  const pv = await g(preview, () => ({ index: window.__grids[0].GetActiveSheetIndex(), names: window.__grids[0].GetSheetNames(), active: [...document.querySelectorAll('[data-testid="spreadsheet-preview-tabs"] .rb-btn')].filter((b) => /secondary/.test(b.className)).map((b) => b.textContent.trim()) }));
  check("preview: opens on the file's active sheet, and marks its tab", pv.index === activeTab && pv.active.join() === pv.names[activeTab], pv);
  await preview.close();
}

// ---- a chart sheet: marked among the tabs, takes no entry, the cell commands are out of reach
{
  const file = `${T}/test-documents-excel/pandas-chartsheet.xlsx`;
  const page = await newPage(context);
  await open(page, "sheet-editor", file); await ready(page);
  const infos = await g(page, () => window.__grids[0].GetSheetInfos().map((i) => ({ index: i.index, name: i.name, kind: i.kind, visible: i.visibility })));
  const chart = infos.find((i) => i.kind === "chartsheet"), sheet = infos.find((i) => i.kind === "worksheet");
  console.log("  sheets:", JSON.stringify(infos));
  const tab = sel(page, `sheet-tab-${chart.index}`);
  check("a chart sheet's tab is marked and says what it is", (await tab.getAttribute("data-kind")) === "chartsheet" && /is-chart/.test(await tab.getAttribute("class")) && /Chart sheet/.test(await tab.getAttribute("title")) && (await tab.locator("svg").count()) === 1 && (await sel(page, `sheet-tab-${sheet.index}`).locator("svg").count()) === 0);
  await sel(page, `sheet-tab-${sheet.index}`).click(); await frames(page);
  check("on a worksheet the cell commands are in reach", (await sel(page, "sheet-toolbar-cells").evaluate((el) => el.inert)) === false && (await sel(page, "sheet-kind-notice").count()) === 0);
  await tab.click(); await frames(page);
  const on = await g(page, () => ({ kind: window.__grids[0].GetSheetKind(), inert: document.querySelector('[data-testid="sheet-toolbar-cells"]').inert, bar: document.querySelector(".rb-sheet-editor-formula").inert, notice: document.querySelector('[data-testid="sheet-kind-notice"]')?.textContent ?? null, undoReach: !document.querySelector('[data-testid="sheet-undo"]').closest("[inert]") }));
  check("on the chart sheet: the cell commands and the bar are inert, a line says why", on.kind === "chartsheet" && on.inert === true && on.bar === true && /Chart sheet/.test(on.notice ?? "") && on.undoReach, on);
  await page.screenshot({ path: `${SHOTS}/new12-chartsheet.png` });
  const rev = await revision(page);
  await page.locator('[data-testid="sheet-grid"] .spreadsheet-content').focus();
  await page.keyboard.type("abc"); await page.keyboard.press("Enter"); await frames(page);
  const entered = await g(page, () => ({ accepted: window.__grids[0].EnterText(0, 0, "x"), a1: window.__grids[0].GetCell(0, 0).text, editing: window.__grids[0].IsEditing() }));
  check("typing on the chart sheet changes nothing", (await revision(page)) === rev && entered.accepted === false && entered.a1 === "" && entered.editing === false && /No changes/.test(await status(page)), { entered, status: await status(page) });
  await page.locator('[data-testid="sheet-grid"]').click({ button: "right", position: { x: 200, y: 150 } }); await frames(page);
  check("no cell menu on the chart sheet", (await page.locator(".rb-office-context-menu, [data-testid='sheet-context-menu'], [role='menu'][aria-label='Cells']").count()) === 0);
  await sel(page, `sheet-tab-${sheet.index}`).click(); await frames(page);
  check("back on the worksheet everything is in reach again", (await sel(page, "sheet-toolbar-cells").evaluate((el) => el.inert)) === false && (await sel(page, "sheet-kind-notice").count()) === 0);
  await page.close();
  const preview = await newPage(context);
  await open(preview, "sheet-preview", file);
  await preview.waitForSelector('[data-testid="spreadsheet-preview-tabs"]', { timeout: 120000 });
  const marked = await g(preview, () => [...document.querySelectorAll('[data-testid="spreadsheet-preview-tabs"] .rb-btn')].map((b) => [b.textContent.trim(), b.getAttribute("data-kind"), b.querySelectorAll("svg").length]));
  check("preview: the chart sheet's tab is marked too", marked.some((m) => m[1] === "chartsheet" && m[2] === 1) && marked.some((m) => m[1] === "worksheet" && m[2] === 0), marked);
  await preview.close();
}

// ---- outline symbols, spilled cells, page setup; then the saved file in a preview
{
  const page = await newPage(context);
  await open(page, "sheet-editor", `${S}/hidden-sheet.xlsx`); await ready(page);
  await sel(page, "sheet-tab-add").click(); await frames(page);
  const table = [["Name", "Qty", "Price"], ["pear", "3", "1.5"], ["apple", "10", "2"], ["fig", "7", "9"], ["apple", "2", "4"], ["kiwi", "5", "6"], ["plum", "8", "3"]];
  await g(page, (table) => { const gr = window.__grids[0]; table.forEach((row, i) => row.forEach((v, j) => gr.EnterText(i, j, v))); }, table);
  // groups: rows 2-6, and 3-4 inside them
  await select(page, 1, 0, 5, 0); await menu(page, "data", "group-rows"); await frames(page);
  await select(page, 2, 0, 3, 0); await menu(page, "data", "group-rows"); await frames(page);
  const strip = () => g(page, () => { const s = [...document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-outline-strip, [data-testid="sheet-grid"] .spreadsheet-outline-line')]; return { n: s.length, stray: s.flatMap((el) => [...el.childNodes].filter((n) => n.nodeType === 3 && n.textContent.trim() !== "").map((n) => n.textContent)), fontSizes: [...new Set(s.map((el) => getComputedStyle(el).fontSize))], buttons: document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-outline-button').length, levels: document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-outline-level').length }; });
  let st = await strip();
  check("nested groups: the strip has its brackets and buttons and no stray text (no CSS of ours hides any)", st.n > 0 && st.stray.length === 0 && !st.fontSizes.includes("0px") && st.buttons === 2 && st.levels === 3, st);
  // the sheet's own switch
  let rev = await revision(page);
  await menu(page, "view", "show-outline"); await frames(page);
  let o = await g(page, () => ({ symbols: window.__grids[0].GetShowOutlineSymbols(), outline: window.__grids[0].GetOutline().showSymbols, component: window.__grids[0].GetShowOutline(), strips: document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-outline-button').length, levels: window.__grids[0].GetOutline().rowLevels }));
  check("View > Outline off: the sheet's own switch (saved, a step), its strips go, the groups stay", o.symbols === false && o.outline === false && o.component !== false && o.strips === 0 && o.levels === 2 && (await revision(page)) !== rev, o);
  await sel(page, "sheet-menu-button-view").click();
  check("the menu shows it unchecked", (await sel(page, "sheet-show-outline").getAttribute("aria-checked")) === "false");
  await page.keyboard.press("Escape");
  await undo(page); await frames(page);
  o = await g(page, () => ({ symbols: window.__grids[0].GetShowOutlineSymbols(), strips: document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-outline-button').length }));
  check("undo brings the strips back", o.symbols === true && o.strips === 2, o);
  await menu(page, "view", "show-outline"); await frames(page);
  await select(page, 5, 0, 5, 0); await menu(page, "data", "group-rows"); await frames(page);
  o = await g(page, () => ({ symbols: window.__grids[0].GetShowOutlineSymbols(), strips: document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-outline-button').length }));
  check("grouping with the strips off switches them on (what was grouped is seen)", o.symbols === true && o.strips >= 2, o);
  // headers hidden: the strip stays
  await menu(page, "view", "show-headers"); await frames(page);
  st = await strip();
  check("headers hidden: the outline strip is still drawn", st.buttons >= 2 && st.levels === 3, st);
  await page.screenshot({ path: `${SHOTS}/new12-outline-no-headers.png` });
  await menu(page, "view", "show-headers"); await frames(page);

  // ---- a spilled array: the bar shows the anchor's formula greyed; an entry starts empty
  await g(page, () => { const gr = window.__grids[0]; gr.EnterText(0, 6, "=SEQUENCE(3,2)"); gr.SelectCell(1, 7); });
  await frames(page);
  let bar = await g(page, () => ({ value: document.querySelector('[data-testid="sheet-formula"]').value, spilled: document.querySelector('[data-testid="sheet-formula"]').classList.contains("is-spilled"), ghost: window.__grids[0].GetGhostFormula(), own: window.__grids[0].GetCurrentCellFormula(), cell: window.__grids[0].GetCell(1, 7).text, status: document.querySelector('[data-testid="sheet-status-spill"]')?.textContent }));
  check("a spilled cell: the bar shows the anchor's formula, greyed (GetGhostFormula); the cell has none of its own", bar.value === "=SEQUENCE(3,2)" && bar.spilled && bar.ghost === "=SEQUENCE(3,2)" && bar.own === "" && bar.cell === "4" && /Array: G1:H3/.test(bar.status ?? ""), bar);
  await page.locator('[data-testid="sheet-grid"] .spreadsheet-content').focus();
  await page.keyboard.press("F2"); await frames(page);
  bar = await g(page, () => ({ value: document.querySelector('[data-testid="sheet-formula"]').value, spilled: document.querySelector('[data-testid="sheet-formula"]').classList.contains("is-spilled"), editing: window.__grids[0].IsEditing(), box: document.querySelector(".spreadsheet-edit-input")?.value }));
  check("F2 on it starts an empty entry, in the cell and in the bar", bar.editing && bar.value === "" && bar.box === "" && !bar.spilled, bar);
  await page.keyboard.press("Escape"); await frames(page);
  await sel(page, "sheet-formula").click(); await frames(page);
  bar = await g(page, () => ({ value: document.querySelector('[data-testid="sheet-formula"]').value, editing: window.__grids[0].IsEditing() }));
  check("a click in the bar starts an empty entry too", bar.editing && bar.value === "", bar);
  await page.keyboard.press("Escape"); await frames(page);
  await g(page, () => window.__grids[0].SelectCell(0, 6)); await frames(page);
  bar = await g(page, () => ({ value: document.querySelector('[data-testid="sheet-formula"]').value, spilled: document.querySelector('[data-testid="sheet-formula"]').classList.contains("is-spilled"), ghost: window.__grids[0].GetGhostFormula() }));
  check("the anchor shows its own formula, not greyed", bar.value === "=SEQUENCE(3,2)" && !bar.spilled && bar.ghost == null, bar);

  // ---- a list validation whose source is a formula, through the editor's dialog
  await g(page, () => window.__grids[0].SelectCell(0, 9)); await frames(page);
  await menu(page, "data", "validation");
  await sel(page, "sheet-validation-dialog").waitFor(); await sel(page, "sheet-validation-type").selectOption("list");
  await sel(page, "sheet-validation-list").fill("=OFFSET($A$2,0,0,3,1)");
  await sel(page, "sheet-dialog-ok").click(); await frames(page);
  const list = await g(page, () => ({ list: window.__grids[0].GetValidationList(0, 9), rule: window.__grids[0].GetValidations?.().map((v) => v.source ?? v.values) }));
  check("a list validation whose source is a formula offers the cells the formula gives", (await sel(page, "sheet-validation-list").count()) === 0 && (list.list ?? []).join() === "pear,apple,fig", list);
  await g(page, () => window.__grids[0].EnterText(0, 9, "banana")); await frames(page);
  const refused = await g(page, () => window.__grids[0].GetCell(0, 9).text);
  await g(page, () => window.__grids[0].EnterText(0, 9, "fig")); await frames(page);
  check("... and entries are checked against it", refused === "" && (await g(page, () => window.__grids[0].GetCell(0, 9).text)) === "fig", { refused });

  // ---- page setup and the printed pages
  await g(page, () => window.__grids[0].SelectRange(0, 0, 3, 2)); await frames(page);
  rev = await revision(page);
  await menu(page, "view", "page-setup");
  await sel(page, "sheet-page-setup-dialog").waitFor();
  const pagesOf = () => g(page, () => [...document.querySelectorAll('[data-testid="sheet-page-list"] [data-range]')].map((r) => r.getAttribute("data-range")));
  const engine = () => g(page, () => { const p = window.__grids[0].GetPrintPages(); return { count: p.pageCount, areas: p.printAreas, ranges: p.pages.map((x) => x.range), scale: p.scale }; });
  let e = await engine();
  check("the dialog lists the pages the sheet prints on, and how many", (await text(page, "sheet-page-summary")).includes(`${e.count} ${e.count === 1 ? "page" : "pages"} at ${Math.round(e.scale)}%`) && (await pagesOf()).join() === e.ranges.join() && e.count >= 1, { summary: await text(page, "sheet-page-summary"), e });
  check("opening it changed nothing", (await revision(page)) === rev);
  await sel(page, "sheet-page-area-set").click();
  check("the selection becomes the print area", (await sel(page, "sheet-page-areas").inputValue()) === "A1:C4");
  await g(page, () => window.__grids[0].SelectRange(0, 6, 2, 7));
  await sel(page, "sheet-page-areas").fill("A1:C4, $G$1:$H$3");
  await sel(page, "sheet-page-orientation").selectOption("landscape");
  await sel(page, "sheet-page-apply").click(); await frames(page);
  e = await engine();
  const setup = await g(page, () => { const s = window.__grids[0].GetPageSetup(); return { areas: s.printAreas, area: s.printArea, orientation: s.orientation }; });
  check("Apply: a print area of two areas and landscape, as one step; each area is its own run of pages", setup.areas.join() === "A1:C4,G1:H3" && setup.orientation === "landscape" && e.count === 2 && e.ranges.join() === "A1:C4,G1:H3" && (await pagesOf()).join() === "A1:C4,G1:H3" && /2 pages at 100% - 2 print areas/.test(await text(page, "sheet-page-summary")) && /area 2/.test(await text(page, "sheet-page-1")), { setup, e, summary: await text(page, "sheet-page-summary") });
  await page.screenshot({ path: `${SHOTS}/new12-page-setup.png` });
  const afterApply = await revision(page);
  await sel(page, "sheet-page-apply").click(); await frames(page);
  check("Apply again changes nothing", (await revision(page)) === afterApply);
  await sel(page, "sheet-page-1").click(); await frames(page);
  check("a click on a page selects its cells", (await selection(page)) === "G1:H3", await selection(page));
  await sel(page, "sheet-page-areas").fill("A1:C4, nonsense");
  await sel(page, "sheet-dialog-ok").click(); await frames(page);
  check("a print area that is no range is refused, the dialog stays", (await sel(page, "sheet-page-setup-dialog").count()) === 1 && /not a range/.test(await text(page, "sheet-dialog-error")) && (await revision(page)) === afterApply, await text(page, "sheet-dialog-error").catch(() => ""));
  await sel(page, "sheet-page-area-clear").click();
  await sel(page, "sheet-page-scaling").selectOption("fit");
  await sel(page, "sheet-page-fit-width").fill("1");
  await sel(page, "sheet-page-fit-height").fill("1");
  await sel(page, "sheet-dialog-ok").click(); await frames(page);
  const fitted = await g(page, () => { const s = window.__grids[0].GetPageSetup(); const p = window.__grids[0].GetPrintPages(); return { areas: s.printAreas, scale: s.scale, w: s.fitToWidth, h: s.fitToHeight, count: p.pageCount, printed: p.scale }; });
  check("OK: no print area, fitted to one page - the whole sheet on one page", (await sel(page, "sheet-page-setup-dialog").count()) === 0 && fitted.areas.length === 0 && fitted.scale == null && fitted.w === 1 && fitted.h === 1 && fitted.count === 1, fitted);
  await undo(page); await frames(page);
  const undone = await g(page, () => { const s = window.__grids[0].GetPageSetup(); return { areas: s.printAreas, w: s.fitToWidth, scale: s.scale }; });
  check("one undo takes the whole OK back", undone.areas.join() === "A1:C4,G1:H3" && undone.scale === 100, undone);
  // a manual break
  await g(page, () => window.__grids[0].SelectCell(2, 0)); await frames(page);
  await menu(page, "view", "page-setup"); await sel(page, "sheet-page-setup-dialog").waitFor();
  const beforeBreak = (await pagesOf()).length;
  await sel(page, "sheet-page-break-row").click(); await frames(page);
  const broken = await g(page, () => ({ breaks: window.__grids[0].GetPageBreaks(), ranges: window.__grids[0].GetPrintPages().pages.map((p) => p.range) }));
  check("a break above the active row: each of the two areas breaks there, listed", broken.breaks.rows.join() === "2" && (await pagesOf()).length === beforeBreak + 2 && (await pagesOf()).join() === "A1:C2,A3:C4,G1:H2,G3:H3" && (await pagesOf()).join() === broken.ranges.join(), { broken, listed: await pagesOf() });
  await sel(page, "sheet-page-break-clear").click(); await frames(page);
  check("the break removed", (await pagesOf()).length === beforeBreak && (await g(page, () => window.__grids[0].GetPageBreaks().rows.length)) === 0);
  await page.keyboard.press("Escape"); await frames(page);
  check("Escape closes the dialog", (await sel(page, "sheet-page-setup-dialog").count()) === 0);

  // ---- the saved workbook: areas kept; in a preview its filter buttons and outline work as view state
  await g(page, () => { const gr = window.__grids[0]; gr.SetAutoFilter("A1:C7"); });
  const bytes = Buffer.from(await g(page, async () => Array.from(await window.__grids[0].ExportXlsx())));
  const parts = await unzip(bytes);
  const wbXml = await parts.get("xl/workbook.xml").async("string");
  check("saved: Print_Area names both areas", /_xlnm\.Print_Area[^>]*>[^<]*\$A\$1:\$C\$4,[^<]*\$G\$1:\$H\$3/.test(wbXml), (/<definedName name="_xlnm\.Print_Area"[^>]*>[^<]*/.exec(wbXml) ?? [""])[0]);
  saved.set("new12.xlsx", bytes);
  await page.close();

  const preview = await newPage(context);
  await open(preview, "sheet-preview", "/api/saved/new12.xlsx", { name: "new12.xlsx" });
  await preview.waitForSelector('[data-testid="spreadsheet-preview"] .spreadsheet-content', { timeout: 120000 });
  await preview.waitForTimeout(300);
  const pv = () => g(preview, () => { const gr = window.__grids[0]; return { sheet: gr.GetActiveSheetIndex(), names: gr.GetSheetNames().length, readOnly: gr.IsReadOnly(), buttons: gr.GetFilterButtonRects().length, drawn: document.querySelectorAll('[data-testid="spreadsheet-preview"] .spreadsheet-filter-button').length, rev: gr.GetRevision(), modified: gr.IsModified(), hidden: (gr.GetFilters()?.hiddenRows ?? []).join(), outline: document.querySelectorAll('[data-testid="spreadsheet-preview"] .spreadsheet-outline-button').length }; });
  let v = await pv();
  check("preview: opens on the sheet that was saved on, and draws its filter buttons and outline", v.readOnly && v.sheet === v.names - 1 && v.buttons === 3 && v.drawn >= 3 && v.outline >= 2, v);
  const rev0 = v.rev;
  await preview.locator('[data-testid="spreadsheet-preview"] .spreadsheet-filter-button[data-col="0"]').first().click();
  await sel(preview, "sheet-filter-panel").waitFor({ timeout: 30000 });
  const panel = await g(preview, () => { const p = document.querySelector('[data-testid="sheet-filter-panel"]'); const b = p.getBoundingClientRect(); const top = document.elementFromPoint(b.left + 20, b.top + 20); return { sort: document.querySelectorAll('[data-testid="sheet-filter-sort-asc"]').length, values: [...p.querySelectorAll("input[data-value]")].map((i) => i.getAttribute("data-value")), onTop: p.contains(top), styled: getComputedStyle(p).position, w: Math.round(b.width) }; });
  check("preview: a filter button opens the dropdown - on top, styled, without the sort commands", panel.sort === 0 && panel.values.join() === "apple,fig,kiwi,pear,plum" && panel.onTop && panel.styled === "fixed" && panel.w > 200, panel);
  await preview.screenshot({ path: `${SHOTS}/new12-preview-filter.png` });
  await preview.locator('[data-testid="sheet-filter-values"] input[data-value="apple"]').uncheck();
  await sel(preview, "sheet-filter-apply").click(); await preview.waitForTimeout(200);
  v = await pv();
  check("preview: filtering hides the rows as view state - no revision, nothing modified", v.hidden === "2,4" && v.rev === rev0 && v.modified === false && (await sel(preview, "sheet-filter-panel").count()) === 0, v);
  const same = Buffer.from(await g(preview, async () => Array.from(await window.__grids[0].ExportXlsx())));
  check("preview: an export is still the file that was opened", same.equals(bytes), { bytes: same.length, opened: bytes.length });
  await preview.locator('[data-testid="spreadsheet-preview"] .spreadsheet-filter-button[data-col="0"]').first().click();
  await sel(preview, "sheet-filter-panel").waitFor();
  await preview.keyboard.press("Escape"); await preview.waitForTimeout(100);
  check("preview: Escape closes the dropdown only", (await sel(preview, "sheet-filter-panel").count()) === 0 && (await g(preview, () => window.__closed)) === 0 && (await preview.locator('[data-testid="spreadsheet-preview"]').count()) === 1);
  await preview.locator('[data-testid="spreadsheet-preview"] .spreadsheet-filter-button[data-col="0"]').first().click();
  await sel(preview, "sheet-filter-panel").waitFor();
  await sel(preview, "sheet-filter-clear-column").click(); await preview.waitForTimeout(200);
  v = await pv();
  check("preview: the filter cleared again", v.hidden === "" && v.rev === rev0, v);
  // outline buttons in the viewer
  const rowsHidden = () => g(preview, () => { const gr = window.__grids[0]; const out = []; for (let r = 0; r < 8; r++) if (gr.GetRowHeight(r) === 0 || gr.IsRowHidden?.(r)) out.push(r); return out.join(); });
  const before = await rowsHidden();
  await preview.locator('[data-testid="spreadsheet-preview"] .spreadsheet-outline-level[data-axis="rows"][data-level="1"]').click(); await preview.waitForTimeout(200);
  const collapsed = await rowsHidden();
  v = await pv();
  check("preview: the outline's level button collapses the groups as view state", collapsed !== before && collapsed.length > 0 && v.rev === rev0 && v.modified === false, { before, collapsed });
  await preview.close();
}

await browser.close();
process.exit(summary() ? 1 : 0);
