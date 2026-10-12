// Spreadsheet UI, part 1: toolbar, status bar, name box, tabs, context menus, freeze, picture.
import { launch, newPage, open, T, S, SHOTS, check, summary, waitUploads, saved } from "./lib.mjs";
import { g, sel, ready, status, frames, cell, select, selection, menu, undo, redo, cellPoint, gridFocused, feature, revision } from "./ui-lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`);
await ready(page);
// a small table in an empty part of the sheet: R101:T105
const R = 100, C = 17;
await g(page, ([R, C]) => { const gr = window.__grids[0]; const rows = [["Name", "Qty", "Price"], ["pear", "3", "1.5"], ["apple", "10", "2"], ["fig", "7", "9"], ["apple", "2", "4"]]; rows.forEach((row, i) => row.forEach((v, j) => gr.EnterText(R + i, C + j, v))); gr.ScrollToCell(R + 12, C + 6); gr.ScrollToCell(R - 2, C - 3); }, [R, C]);
await frames(page);
const fmt = async (r, c) => (await cell(page, r, c)).format;

// ---- font family and size
await select(page, R + 1, C, R + 4, C);
await feature(page, "font family (toolbar box)", { read: async () => (await fmt(R + 1, C)).fontFamily + "|" + (await fmt(R + 4, C)).fontFamily, act: () => sel(page, "sheet-font-family").selectOption("Georgia"), expect: (v) => v === "Georgia|Georgia" });
check("the font box shows the selection's font", (await sel(page, "sheet-font-family").inputValue()) === "Georgia");
await feature(page, "font size (toolbar box)", { read: async () => (await fmt(R + 1, C)).fontSize, act: () => sel(page, "sheet-font-size").selectOption("14"), expect: (v) => v === 14 });
check("the size box shows it, and the keyboard is back in the grid", (await sel(page, "sheet-font-size").inputValue()) === "14" && (await gridFocused(page)));
// ---- alignment
await feature(page, "vertical alignment (menu)", { read: async () => (await fmt(R + 1, C)).verticalAlign, act: async () => { await sel(page, "sheet-valign").click(); await sel(page, "sheet-valign-top").click(); }, expect: (v) => v === "top" });
await feature(page, "horizontal alignment", { read: async () => (await fmt(R + 1, C)).horizontalAlign, act: () => sel(page, "sheet-align-center").click(), expect: (v) => v === "center" });
check("alignment buttons reflect the selection", (await sel(page, "sheet-align-center").getAttribute("aria-pressed")) === "true" && (await sel(page, "sheet-align-left").getAttribute("aria-pressed")) === "false");
await feature(page, "wrap text", { read: async () => (await fmt(R + 1, C)).wrap, act: () => sel(page, "sheet-wrap").click(), expect: (v) => v === true });
check("wrap button pressed", (await sel(page, "sheet-wrap").getAttribute("aria-pressed")) === "true");
// ---- mixed state
await select(page, R + 1, C, R + 1, C);
await sel(page, "sheet-bold").click();
await select(page, R + 1, C, R + 2, C);
await frames(page);
check("a toggle that is on for part of the selection shows mixed", (await sel(page, "sheet-bold").getAttribute("aria-pressed")) === "mixed");
// ---- decimals
await select(page, R + 1, C + 2, R + 4, C + 2);
await feature(page, "more decimals", { read: async () => (await cell(page, R + 1, C + 2)).text + "|" + (await fmt(R + 1, C + 2)).numberFormat, act: () => sel(page, "sheet-decimal-increase").click(), expect: (v) => v === "1.50|0.00" });
await feature(page, "fewer decimals", { read: async () => (await cell(page, R + 3, C + 2)).text, act: async () => { await sel(page, "sheet-decimal-decrease").click(); await sel(page, "sheet-decimal-decrease").click(); }, expect: (v) => v === "9", steps: 2 });
await undo(page); await undo(page);
// ---- merge
await select(page, R + 7, C, R + 7, C + 2);
await feature(page, "merge cells", { read: async () => (await cell(page, R + 7, C)).mergeRange, act: () => sel(page, "sheet-merge").click(), expect: (v) => v === "R108:T108" });
// ---- format painter with a real drag
await select(page, R + 1, C);
await sel(page, "sheet-fill-color").click();
await page.click('[data-testid="sheet-fill-palette"] [data-color="#FEF08A"]');
await sel(page, "sheet-format-painter").click();
check("the painter is armed", (await sel(page, "sheet-format-painter").getAttribute("aria-pressed")) === "true");
const from = await cellPoint(page, R + 2, C + 1), to = await cellPoint(page, R + 3, C + 2);
await feature(page, "format painter (drag over the target cells)", {
  read: async () => [await fmt(R + 2, C + 1), await fmt(R + 3, C + 2)].map((f) => `${f.fill}/${f.bold}/${f.fontFamily}/${f.fontSize}`).join(" "),
  act: async () => { await page.mouse.move(from.x, from.y); await page.mouse.down(); await page.mouse.move(to.x, to.y, { steps: 4 }); await page.mouse.up(); },
  expect: (v) => v === "#FEF08A/true/Georgia/14 #FEF08A/true/Georgia/14",
});
check("the painter is off again and the values are untouched", (await sel(page, "sheet-format-painter").getAttribute("aria-pressed")) === "false" && (await cell(page, R + 2, C + 1)).value === 10);
// ---- clear
await select(page, R + 3, C + 1, R + 3, C + 2);
await feature(page, "clear formats", { read: async () => (await fmt(R + 3, C + 2)).fill + "|" + (await cell(page, R + 3, C + 2)).value, act: async () => { await sel(page, "sheet-clear").click(); await sel(page, "sheet-clear-formats").click(); }, expect: (v) => v === "null|9" });
await select(page, R + 8, C);
await g(page, ([r, c]) => window.__grids[0].EnterText(r, c, "gone"), [R + 8, C]);
await feature(page, "clear contents", { read: async () => (await cell(page, R + 8, C)).value, act: async () => { await sel(page, "sheet-clear").click(); await sel(page, "sheet-clear-contents").click(); }, expect: (v) => v === null });
await page.screenshot({ path: `${SHOTS}/ui-toolbar.png` });

// ---- status bar and zoom
await select(page, R + 1, C + 1, R + 4, C + 1);
await frames(page);
check("status bar: sum, average and count of the selection", (await sel(page, "sheet-status-sum").innerText()) === "Sum: 22" && (await sel(page, "sheet-status-average").innerText()) === "Average: 5.5" && (await sel(page, "sheet-status-count").innerText()) === "Count: 4", await sel(page, "sheet-status").innerText());
await select(page, R + 1, C);
await frames(page);
check("status bar: nothing for one cell", (await sel(page, "sheet-status-sum").count()) === 0);
const rev = await revision(page);
await sel(page, "sheet-zoom-in").click();
check("zoom in: the sheet's zoom changes, the label follows, and it is not an edit", (await g(page, () => window.__grids[0].GetZoom())) === 1.1 && (await sel(page, "sheet-zoom").innerText()) === "110%" && (await revision(page)) === rev);
await sel(page, "sheet-zoom").click();
check("zoom reset", (await g(page, () => window.__grids[0].GetZoom())) === 1);

// ---- name box
await sel(page, "sheet-cell-ref").click();
await page.keyboard.type("U105:V106");
await page.keyboard.press("Enter");
check("name box: a typed range is selected and the keyboard is in the grid", (await selection(page)) === "U105:V106" && (await gridFocused(page)) && (await sel(page, "sheet-cell-ref").inputValue()) === "U105:V106");
await sel(page, "sheet-cell-ref").click();
await page.keyboard.type("Monthly!B3");
await page.keyboard.press("Enter");
check("name box: another sheet's cell switches the sheet", (await g(page, () => window.__grids[0].GetActiveSheetIndex())) === 1 && (await selection(page)) === "B3");
await sel(page, "sheet-cell-ref").click();
await page.keyboard.type("nonsense here");
await page.keyboard.press("Enter");
check("name box: what is not a place says so", (await sel(page, "sheet-notice").count()) === 1, await sel(page, "sheet-notice").innerText().catch(() => ""));
await sel(page, "sheet-tab-0").click();

// ---- sheet tabs
const infos = () => g(page, () => window.__grids[0].GetSheetInfos().map((s) => `${s.name}${s.visibility === "visible" ? "" : "(" + s.visibility + ")"}${s.tabColor ? s.tabColor : ""}`).join(","));
await feature(page, "add a sheet (+)", { read: infos, act: () => sel(page, "sheet-tab-add").click(), expect: (v) => v === "Annual,Monthly,Sheet1" });
check("the new sheet is shown and has a tab", (await g(page, () => window.__grids[0].GetActiveSheetIndex())) === 2 && (await sel(page, "sheet-tab-2").innerText()) === "Sheet1");
const tabMenu = async (index, item) => { await sel(page, `sheet-tab-${index}`).click({ button: "right" }); await sel(page, "sheet-context-menu").waitFor(); if (item) await sel(page, `sheet-menu-${item}`).click(); };
await tabMenu(2);
await page.screenshot({ path: `${SHOTS}/ui-tab-menu.png` });
await page.keyboard.press("Escape");
check("Escape closes the tab menu and not the editor", (await sel(page, "sheet-context-menu").count()) === 0 && (await sel(page, "office-editor").count()) === 1);
await feature(page, "rename a sheet (tab menu, dialog)", { read: infos, act: async () => { await tabMenu(2, "tab-rename"); await sel(page, "sheet-rename-input").fill("Notes"); await page.keyboard.press("Enter"); }, expect: (v) => v === "Annual,Monthly,Notes" });
await tabMenu(2, "tab-rename");
await sel(page, "sheet-rename-input").fill("Annual");
await page.keyboard.press("Enter");
check("a name that is taken is refused in the dialog", (await sel(page, "sheet-dialog-error").count()) === 1 && (await infos()) === "Annual,Monthly,Notes", await sel(page, "sheet-dialog-error").innerText().catch(() => ""));
await page.screenshot({ path: `${SHOTS}/ui-rename-dialog.png` });
await page.keyboard.press("Escape");
check("Escape closes the dialog; the grid has the keyboard", (await sel(page, "sheet-rename-dialog").count()) === 0 && (await gridFocused(page)));
await feature(page, "move a sheet left (tab menu)", { read: infos, act: () => tabMenu(2, "tab-move-left"), expect: (v) => v === "Annual,Notes,Monthly" });
await feature(page, "tab colour (tab menu)", { read: infos, act: () => tabMenu(1, "tab-color-2563EB"), expect: (v) => v === "Annual,Notes#2563EB,Monthly" });
check("the tab shows its colour", (await sel(page, "sheet-tab-1").evaluate((el) => getComputedStyle(el).boxShadow)).includes("rgb(37, 99, 235)"));
await feature(page, "duplicate a sheet (tab menu)", { read: infos, act: () => tabMenu(2, "tab-duplicate"), expect: (v) => v === "Annual,Notes#2563EB,Monthly,Monthly (2)" });
await feature(page, "hide a sheet (tab menu)", { read: infos, act: () => tabMenu(3, "tab-hide"), expect: (v) => v === "Annual,Notes#2563EB,Monthly,Monthly (2)(hidden)" });
check("a hidden sheet has no tab", (await sel(page, "sheet-tab-3").count()) === 0 && (await page.locator('[data-testid="sheet-tabs"] [role="tab"]').count()) === 3);
await feature(page, "unhide a sheet (tab menu)", { read: infos, act: () => tabMenu(0, "tab-unhide-3"), expect: (v) => v === "Annual,Notes#2563EB,Monthly,Monthly (2)" });
await feature(page, "delete a sheet (tab menu)", { read: infos, act: () => tabMenu(3, "tab-delete"), expect: (v) => v === "Annual,Notes#2563EB,Monthly" });
// keyboard: the menu key on a tab
await sel(page, "sheet-tab-0").focus();
await page.keyboard.press("Shift+F10");
await sel(page, "sheet-context-menu").waitFor();
await page.keyboard.press("ArrowDown");
check("menu key on a tab opens its menu; arrows move in it", await g(page, () => document.activeElement?.getAttribute("data-testid") === "sheet-menu-tab-duplicate"));
await page.keyboard.press("Escape");
await sel(page, "sheet-tab-0").click();

// ---- context menus on cells and headers
const p = await cellPoint(page, R + 2, C);
await page.mouse.click(p.x, p.y, { button: "right" });
await sel(page, "sheet-context-menu").waitFor();
check("right click on a cell: it is selected and the menu opens at the pointer", (await selection(page)) === "R103" && Math.abs((await sel(page, "sheet-context-menu").boundingBox()).x - p.x) < 4);
await page.screenshot({ path: `${SHOTS}/ui-cell-menu.png` });
await page.keyboard.press("Escape");
const rowsOf = async () => [await cell(page, R + 2, C), await cell(page, R + 3, C)].map((c) => c.value).join(",");
await feature(page, "insert a row (cell menu)", { read: rowsOf, act: async () => { await page.mouse.click(p.x, p.y, { button: "right" }); await sel(page, "sheet-menu-insert-rows").click(); }, expect: (v) => v === ",apple" });
await feature(page, "delete a row (cell menu)", { read: rowsOf, act: async () => { const p = await cellPoint(page, R + 2, C); await page.mouse.click(p.x, p.y, { button: "right" }); await sel(page, "sheet-menu-delete-rows").click(); }, expect: (v) => v === "apple,fig" });
// keyboard on a cell
await select(page, R + 9, C);
await page.locator(".spreadsheet-content").first().focus();
await page.keyboard.press("Shift+F10");
await sel(page, "sheet-context-menu").waitFor();
check("Shift+F10 on the active cell opens the menu with the first entry focused", await g(page, () => document.activeElement?.getAttribute("data-testid") === "sheet-menu-cut"));
await page.keyboard.press("Escape");
check("after Escape the keyboard is in the grid and the selection is as it was", (await gridFocused(page)) && (await selection(page)) === "R110");
// (the component ignores contextmenu events for 400 ms after a keyboard-opened menu: the browser's own one)
await page.waitForTimeout(450);
// row header
const headerPoint = async (axis, index) => {
  const find = () => g(page, ([axis, index]) => { const sc = document.querySelector('[data-testid="sheet-grid"]').getBoundingClientRect(); const q = [...document.querySelectorAll(axis === "row" ? `.spreadsheet-row-header[data-row="${index}"]` : `.spreadsheet-col-header[data-col="${index}"]`)].map((x) => x.getBoundingClientRect()).find((q) => q.width > 0 && q.height > 0 && q.bottom <= sc.bottom - 2 && q.right <= sc.right - 2 && q.top >= sc.top && q.left >= sc.left); return q ? { x: q.left + q.width / 2, y: q.top + q.height / 2 } : null; }, [axis, index]);
  let p = await find();
  if (!p) { await g(page, ([axis, index]) => { const gr = window.__grids[0]; const a = gr.GetActiveCell(); if (axis === "row") { gr.ScrollToCell(Math.max(0, index - 4), a.Col); gr.ScrollToCell(index + 4, a.Col); } else { gr.ScrollToCell(a.Row, Math.max(0, index - 2)); gr.ScrollToCell(a.Row, index + 1); } }, [axis, index]); await frames(page); p = await find(); }
  if (!p) throw new Error(`no ${axis} header ${index} on screen`);
  return p;
};
const rowHeader = (row) => headerPoint("row", row);
const colHeader = (col) => headerPoint("col", col);
const headerMenu = async (point, item) => { await page.mouse.click(point.x, point.y, { button: "right" }); await sel(page, "sheet-context-menu").waitFor(); if (item) await sel(page, `sheet-menu-${item}`).click(); };
let rh = await rowHeader(R + 6);
await headerMenu(rh);
check("right click on a row header selects the row", /^A107:[A-Z]+107$/.test(await selection(page)), await selection(page));
await page.screenshot({ path: `${SHOTS}/ui-row-menu.png` });
await page.keyboard.press("Escape");
await feature(page, "hide a row (row header menu)", { read: () => g(page, () => window.__grids[0].GetHiddenRows().join()), act: async () => headerMenu(await rowHeader(R + 6), "hide"), expect: (v) => v.split(",").includes("107:107") });
// unhide through the rows around it
const a = await rowHeader(R + 5), b = await rowHeader(R + 7);
await page.mouse.move(a.x, a.y); await page.mouse.down(); await page.mouse.move(b.x, b.y, { steps: 3 }); await page.mouse.up();
await feature(page, "unhide rows (row header menu, rows around the hidden one selected)", { read: () => g(page, () => window.__grids[0].GetHiddenRows().join()), act: () => headerMenu(a, "unhide"), expect: (v) => !v.split(",").includes("107:107") });
rh = await rowHeader(R + 6);
await feature(page, "row height (dialog)", { read: () => g(page, (r) => window.__grids[0].GetRowHeight(r), R + 6), act: async () => { await headerMenu(await rowHeader(R + 6), "size"); await sel(page, "sheet-size-input").fill("44"); await page.keyboard.press("Enter"); }, expect: (v) => v === 44 });
let ch = await colHeader(C);
await feature(page, "column width (dialog)", { read: () => g(page, (c) => window.__grids[0].GetColumnWidth(c), C), act: async () => { await headerMenu(await colHeader(C), "size"); await sel(page, "sheet-size-input").fill("180"); await sel(page, "sheet-dialog-ok").click(); }, expect: (v) => v === 180 });
ch = await colHeader(C);
await feature(page, "auto-fit a column (column header menu)", { read: () => g(page, (c) => window.__grids[0].GetColumnWidth(c), C), act: async () => headerMenu(await colHeader(C), "auto-fit"), expect: (v) => v > 20 && v < 120 });
ch = await colHeader(C + 4);
await feature(page, "hide a column (column header menu)", { read: () => g(page, () => window.__grids[0].GetHiddenColumns().join()), act: async () => headerMenu(await colHeader(C + 4), "hide"), expect: (v) => v === "V:V" });

// ---- freeze panes and view options
await select(page, 4, 2);
await feature(page, "freeze panes at the selection (View menu)", { read: () => g(page, () => JSON.stringify(window.__grids[0].GetFrozen())), act: () => menu(page, "view", "freeze-here"), expect: (v) => v === '{"rows":4,"cols":2}' });
await sel(page, "sheet-menu-button-view").click();
await page.screenshot({ path: `${SHOTS}/ui-view-menu.png` });
check("the View menu shows the state", !(await sel(page, "sheet-unfreeze").isDisabled()) && (await sel(page, "sheet-show-headers").getAttribute("aria-checked")) === "true");
await sel(page, "sheet-unfreeze").click();
check("unfreeze", (await g(page, () => JSON.stringify(window.__grids[0].GetFrozen()))) === '{"rows":0,"cols":0}');
await undo(page);
await feature(page, "show formulas (View menu)", { read: () => g(page, () => window.__grids[0].GetShowFormulas()), act: () => menu(page, "view", "show-formulas"), expect: (v) => v === true });
await menu(page, "view", "show-formulas");

// ---- picture
await select(page, R + 12, C);
await feature(page, "insert a picture from a file", { read: () => g(page, () => window.__grids[0].GetPictures().map((p) => `${p.name}:${p.width}x${p.height}@${p.row},${p.col}`).join()), act: async () => { await sel(page, "sheet-picture-input").setInputFiles(`${S}/red.png`); await page.waitForFunction(() => window.__grids[0].GetPictures().length === 1); }, expect: (v) => v === `red:200x120@${R + 12},${C}` });
await g(page, ([r, c]) => { const gr = window.__grids[0]; gr.ScrollToCell(r + 10, c); gr.ScrollToCell(r - 4, c - 3); }, [R + 12, C]);
await page.waitForTimeout(300);
check("the picture is drawn", await g(page, () => [...document.querySelectorAll(".spreadsheet-picture")].some((img) => img.getBoundingClientRect().width > 150)));
await page.screenshot({ path: `${SHOTS}/ui-picture.png` });

// ---- save, re-open
check("dirty", (await status(page)).includes("Unsaved"));
await sel(page, "office-editor-save").click();
const bytes = await waitUploads(page, 1);
await page.waitForTimeout(200);
check("saved", (await status(page)).includes("Saved") && (await sel(page, "office-editor-error").count()) === 0, await status(page));
saved.set("ui1", bytes);
const expectState = await g(page, ([R, C]) => { const gr = window.__grids[0]; const f = (r, c) => { const x = gr.GetCell(r, c).format; return [x.fontFamily, x.fontSize, x.verticalAlign, x.horizontalAlign, x.wrap, x.bold, x.fill, x.numberFormat].join("/"); }; return { a: f(R + 1, C), painted: f(R + 2, C + 1), price: gr.GetCell(R + 1, C + 2).text, merge: gr.GetCell(R + 7, C).mergeRange, sheets: gr.GetSheetInfos().map((s) => s.name + (s.tabColor ?? "")).join(), frozen: JSON.stringify(gr.GetFrozen()), rowH: gr.GetRowHeight(R + 6), hiddenCols: gr.GetHiddenColumns().join(), pictures: gr.GetPictures().map((p) => `${p.width}x${p.height}@${p.row},${p.col}`).join() }; }, [R, C]);
await page.close();
const page2 = await newPage(context);
await open(page2, "sheet-editor", "/api/saved/ui1", { name: "milk-street-office.xlsx" });
await ready(page2);
const got = await g(page2, ([R, C]) => { const gr = window.__grids[0]; const f = (r, c) => { const x = gr.GetCell(r, c).format; return [x.fontFamily, x.fontSize, x.verticalAlign, x.horizontalAlign, x.wrap, x.bold, x.fill, x.numberFormat].join("/"); }; return { a: f(R + 1, C), painted: f(R + 2, C + 1), price: gr.GetCell(R + 1, C + 2).text, merge: gr.GetCell(R + 7, C).mergeRange, sheets: gr.GetSheetInfos().map((s) => s.name + (s.tabColor ?? "")).join(), frozen: JSON.stringify(gr.GetFrozen()), rowH: gr.GetRowHeight(R + 6), hiddenCols: gr.GetHiddenColumns().join(), pictures: gr.GetPictures().map((p) => `${p.width}x${p.height}@${p.row},${p.col}`).join() }; }, [R, C]);
for (const key of Object.keys(expectState)) check(`re-open: ${key} survives the save`, JSON.stringify(got[key]) === JSON.stringify(expectState[key]), { saved: expectState[key], reopened: got[key] });
check("re-open: clean, tabs with the colour", (await status(page2)).includes("No changes") && (await sel(page2, "sheet-tab-1").evaluate((el) => getComputedStyle(el).boxShadow)).includes("rgb(37, 99, 235)"));
await page2.close();
summary();
await browser.close();
