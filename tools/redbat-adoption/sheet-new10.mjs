// What @rbt/office 0.1.10 adds, in the spreadsheet editor: the grid's filter buttons with the editor's
// dropdown, the outline strip with the editor's group / view commands, spilled arrays in the bar and status.
import { launch, newPage, open, T, S, SHOTS, check, summary, waitUploads, saved } from "./lib.mjs";
import { g, sel, ready, status, frames, cell, select, selection, menu, undo, redo, cellPoint, gridFocused, feature, revision } from "./ui-lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "sheet-editor", `${S}/hidden-sheet.xlsx`);
await ready(page);
// a small table of data on a fresh sheet
await sel(page, "sheet-tab-add").click(); await frames(page);
const R = 2, C = 1;
const table = [["Name", "Qty", "Price"], ["pear", "3", "1.5"], ["apple", "10", "2"], ["fig", "7", "9"], ["apple", "2", "4"]];
await g(page, ([R, C, table]) => { const gr = window.__grids[0]; table.forEach((row, i) => row.forEach((v, j) => gr.EnterText(R + i, C + j, v))); gr.SelectCell(R + 1, C); }, [R, C, table]);
await frames(page);
const column = (j) => g(page, ([R, C, j]) => { const gr = window.__grids[0]; const out = []; for (let i = 1; i <= 4; i++) out.push(gr.GetCell(R + i, C + j).text); return out.join(","); }, [R, C, j]);
const fbtn = (col) => page.locator(`[data-testid="sheet-grid"] .spreadsheet-filter-button[data-col="${col}"]`).first();
const filters = () => g(page, () => { const f = window.__grids[0].GetFilters(); return f ? `${f.range} hidden:${f.hiddenRows.join("/")} ${f.columns.map((c) => c.letter + JSON.stringify(c.criteria)).join(" ")}` : "none"; });
const panelBox = () => sel(page, "sheet-filter-panel").boundingBox();
const rects = () => g(page, () => window.__grids[0].GetFilterButtonRects());
const shot = (name) => page.screenshot({ path: `${SHOTS}/new10-${name}.png` });

// ---- filter buttons: the grid's, with the editor's dropdown anchored at the button
check("no filter: no buttons", (await rects()).length === 0 && (await page.locator(".spreadsheet-filter-button").count()) === 0);
await menu(page, "data", "filter"); await frames(page);
let buttons = await rects();
check("Data > Filter: the grid draws a button per header cell, and nothing else does", buttons.length === 3 && buttons.map((b) => `${b.row},${b.column}`).join(" ") === `${R},${C} ${R},${C + 1} ${R},${C + 2}` && (await g(page, () => document.querySelectorAll('.rb-sheet-filter-button, [data-testid^="sheet-filter-button-"]').length)) === 0 && (await g(page, () => [...document.querySelectorAll(".col-filter-btn")].filter((b) => b.getBoundingClientRect().width > 0).length)) === 0, buttons);
const before = await selection(page);
await fbtn(C + 1).click();
await sel(page, "sheet-filter-panel").waitFor();
let box = await panelBox();
let b = buttons[1].clientRect;
check("a click on a button opens the editor's dropdown under it, and selects nothing", Math.abs(box.x - b.left) <= 1 && Math.abs(box.y - (b.top + b.height + 2)) <= 1 && (await selection(page)) === before && (await g(page, () => document.querySelectorAll(".spreadsheet-filter-menu, .spreadsheet-filter-checklist, .filter-dropdown").length)) === 0, { box, button: b });
check("the dropdown is the column's", (await sel(page, "sheet-filter-panel").getAttribute("aria-label")) === "Sort and filter column C" && (await g(page, () => [...document.querySelectorAll('[data-testid="sheet-filter-values"] input[data-value]')].map((i) => i.getAttribute("data-value")).join())) === "2,3,7,10");
await shot("filter-dropdown");
await sel(page, "sheet-filter-operator").selectOption("greaterThan");
await sel(page, "sheet-filter-operand").fill("5");
await feature(page, "a condition set from the dropdown", { read: filters, act: () => page.keyboard.press("Enter"), expect: (v) => /C\{"conditions":\[\{"operator":"greaterThan","value":5\}\]/.test(v) });
await frames(page);
buttons = await rects();
check("the filtered column's button is active (the grid marks it)", buttons.find((x) => x.column === C + 1)?.active === true && (await fbtn(C + 1).getAttribute("data-active")) === "1" && buttons.filter((x) => x.active).length === 1, buttons.map((x) => [x.column, x.active]));
// Alt+Down on a header cell
await select(page, R, C);
await page.locator('[data-testid="sheet-grid"] .spreadsheet-content').first().focus();
await page.keyboard.press("Alt+ArrowDown");
await sel(page, "sheet-filter-panel").waitFor();
box = await panelBox(); b = (await rects()).find((x) => x.column === C).clientRect;
check("Alt+Down on a header cell opens the dropdown of its button", Math.abs(box.x - b.left) <= 1 && Math.abs(box.y - (b.top + b.height + 2)) <= 1 && (await sel(page, "sheet-filter-panel").getAttribute("aria-label")) === "Sort and filter column B", { box, b });
await page.keyboard.press("Escape");
check("Escape closes it and the grid has the keyboard", (await sel(page, "sheet-filter-panel").count()) === 0 && (await gridFocused(page)));
// sort from the dropdown: the button shows the direction
await menu(page, "data", "filter-clear");
await fbtn(C + 1).click();
await feature(page, "sort from the dropdown", { read: () => column(1), act: () => sel(page, "sheet-filter-sort-desc").click(), expect: (v) => v === "10,7,3,2" });
await frames(page);
check("the sorted column's button says so", (await rects()).find((x) => x.column === C + 1)?.sorted === "desc", (await rects()).map((x) => [x.column, x.sorted]));
// zoomed and scrolled: the dropdown follows the button
await sel(page, "sheet-zoom-in").click(); await sel(page, "sheet-zoom-in").click(); await frames(page);
b = (await rects()).find((x) => x.column === C + 2).clientRect;
await fbtn(C + 2).click();
await sel(page, "sheet-filter-panel").waitFor();
box = await panelBox();
check("zoomed to 125 %: the button is larger and the dropdown is still under it", b.width === 20 && Math.abs(box.x - b.left) <= 1 && Math.abs(box.y - (b.top + b.height + 2)) <= 1, { box, b });
await page.keyboard.press("Escape");
await sel(page, "sheet-zoom").click(); await frames(page);
// a table's own button, on a sheet without an AutoFilter
await menu(page, "data", "filter"); await frames(page);
check("Data > Filter again removes the AutoFilter and its buttons", (await filters()) === "none" && (await rects()).length === 0);
await g(page, ([R, C]) => { const gr = window.__grids[0]; const col = (n) => String.fromCharCode(65 + n); gr.DefineTable(`${col(C)}${R + 1}:${col(C + 2)}${R + 5}`, "Fruit", true, "TableStyleMedium2"); }, [R, C]).catch((e) => console.log("  DefineTable:", String(e).slice(0, 200)));
await frames(page);
buttons = await rects();
if (buttons.length === 3) {
  await fbtn(C).click();
  await sel(page, "sheet-filter-panel").waitFor();
  await feature(page, "a table's button: sort from the dropdown without a sheet AutoFilter", { read: () => column(0), act: () => sel(page, "sheet-filter-sort-asc").click(), expect: (v) => v === "apple,apple,fig,pear" });
  await fbtn(C).click();
  await sel(page, "sheet-filter-search").fill("ap");
  await feature(page, "a table's button: filter by value", { read: filters, act: () => sel(page, "sheet-filter-apply").click(), expect: (v) => /B\{"values":\["apple"\]/.test(v) });
  await menu(page, "data", "filter-clear");
} else console.log("  (no table buttons: ", JSON.stringify(buttons), ")");

// ---- outline: the grid's strip, the editor's commands
const outline = () => g(page, () => { const o = window.__grids[0].GetOutline(); return { rows: o.rows.map((x) => `${x.range}@${x.level}${x.collapsed ? "c" : ""}`).join(" "), columns: o.columns.map((x) => `${x.range}@${x.level}${x.collapsed ? "c" : ""}`).join(" "), below: o.summaryBelow, right: o.summaryRight }; });
const corner = () => g(page, () => { const q = document.querySelector('[data-testid="sheet-grid"] .spreadsheet-corner').getBoundingClientRect(); return [Math.round(q.width), Math.round(q.height)]; });
const strip = () => g(page, () => ({ buttons: [...document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-outline-button')].filter((e) => e.getBoundingClientRect().width > 0).length, levels: [...document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-outline-level')].filter((e) => e.getBoundingClientRect().width > 0).length }));
const corner0 = await corner();
check("a sheet without groups: the header area is 50 x 30 and there is no strip", corner0.join() === "50,30" && (await strip()).buttons === 0, corner0);
await select(page, R + 1, C, R + 3, C);
await feature(page, "Data > Group rows", { read: outline, act: () => menu(page, "data", "group-rows"), expect: (v) => v.rows === `${R + 2}:${R + 4}@1` });
await frames(page);
const corner1 = await corner();
check("the grid draws its strip (a +/- button, the level buttons) and the headers move over", corner1[0] > corner0[0] && corner1[1] === corner0[1] && (await strip()).buttons === 1 && (await strip()).levels === 2, { corner1, strip: await strip() });
await select(page, R + 1, C + 1, R + 1, C + 2);
await feature(page, "Data > Group columns", { read: outline, act: () => menu(page, "data", "group-columns"), expect: (v) => v.columns === "C:D@1" });
await frames(page);
check("a strip above the column letters too", (await corner())[1] > corner0[1], await corner());
await shot("outline");
// the grid's own buttons: collapse, levels; the editor's chrome follows
await page.locator('[data-testid="sheet-grid"] .spreadsheet-outline-button[data-axis="rows"]').first().click(); await frames(page);
check("the strip's - button collapses the group (one undo step, shown by the editor)", (await outline()).rows.endsWith("c") && /Unsaved/.test(await status(page)) && !(await sel(page, "sheet-undo").isDisabled()), await outline());
await page.locator('[data-testid="sheet-grid"] .spreadsheet-outline-level[data-axis="rows"][data-level="2"]').first().click(); await frames(page);
check("the level button 2 shows everything again", !(await outline()).rows.endsWith("c"), await outline());
// the filter dropdown still lands under its button with the strips there
await select(page, R + 1, C);
await menu(page, "data", "filter"); await frames(page);
b = (await rects()).find((x) => x.column === C)?.clientRect;
await fbtn(C).click(); await sel(page, "sheet-filter-panel").waitFor();
box = await panelBox();
check("with the strips the dropdown is still under its button", !!b && Math.abs(box.x - b.left) <= 1 && Math.abs(box.y - (b.top + b.height + 2)) <= 1, { box, b });
await page.keyboard.press("Escape");
await menu(page, "data", "filter"); await frames(page);
// View: strips off / on (a setting, not an edit), summary side (the sheet's)
const rev = await revision(page);
await menu(page, "view", "show-outline"); await frames(page);
check("View > Outline hides the strips (0.1.12: the sheet's own switch, saved with it): the header area is 50 x 30 again", (await corner()).join() === "50,30" && (await strip()).buttons === 0 && (await revision(page)) !== rev && (await g(page, () => window.__grids[0].GetShowOutlineSymbols())) === false && (await g(page, () => window.__grids[0].GetShowOutline())) === true, await corner());
await sel(page, "sheet-menu-button-view").click();
check("the menu shows it unchecked", (await sel(page, "sheet-show-outline").getAttribute("aria-checked")) === "false" && (await sel(page, "sheet-outline-summary-below").getAttribute("aria-checked")) === "true");
await sel(page, "sheet-show-outline").click(); await frames(page);
check("and shows them again", (await strip()).buttons >= 2 && (await g(page, () => window.__grids[0].GetShowOutlineSymbols())) === true, await strip());
await menu(page, "view", "outline-summary-below"); await frames(page);
check("View > Summary rows below, unticked: the sheet's summaries are above", (await outline()).below === false && (await outline()).right === true, await outline());
await sel(page, "sheet-menu-button-view").click();
check("the menu follows", (await sel(page, "sheet-outline-summary-below").getAttribute("aria-checked")) === "false" && (await sel(page, "sheet-outline-summary-right").getAttribute("aria-checked")) === "true");
await sel(page, "sheet-outline-summary-right").click(); await frames(page);
check("View > Summary columns right, unticked", (await outline()).right === false, await outline());
// the row header's menu
await select(page, R + 2, C);
const hdr = await g(page, ([r]) => { const e = [...document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-row-header')].find((e) => e.textContent.trim() === String(r + 1) && e.getBoundingClientRect().height > 0); const q = e?.getBoundingClientRect(); return q ? { x: q.left + q.width / 2, y: q.top + q.height / 2 } : null; }, [R + 2]);
if (hdr) {
  await page.mouse.click(hdr.x, hdr.y, { button: "right" });
  await page.locator('[data-testid="sheet-menu-group"]').waitFor({ timeout: 5000 }).catch(() => {});
  const has = await page.locator('[data-testid="sheet-menu-group"]').count();
  if (has) { await feature(page, "row header menu > Group row", { read: outline, act: () => page.locator('[data-testid="sheet-menu-group"]').click(), expect: (v) => v.rows.includes("@2") }); }
  else check("row header menu has Group", false, await g(page, () => [...document.querySelectorAll('[data-testid^="sheet-menu-"]')].map((e) => e.dataset.testid)));
} else check("row header found", false);
await select(page, R + 1, C, R + 3, C);
await feature(page, "Data > Ungroup rows", { read: outline, act: () => menu(page, "data", "ungroup-rows"), expect: (v) => !v.rows.includes(`${R + 2}:${R + 4}@1`) });

// ---- a dynamic array: the bar and the status
const SR = 12, SC = 6;
await g(page, ([r, c]) => { const gr = window.__grids[0]; gr.EnterText(r, c, "=SEQUENCE(3,2)"); gr.SelectCell(r, c); }, [SR, SC]); await frames(page);
const bar = async () => ({ value: await sel(page, "sheet-formula").inputValue(), spilled: (await sel(page, "sheet-formula").getAttribute("class")).includes("is-spilled"), title: await sel(page, "sheet-formula").getAttribute("title"), status: (await sel(page, "sheet-status-spill").count()) ? await sel(page, "sheet-status-spill").innerText() : null });
let now = await bar();
check("on the cell that spills: its formula, and the array's range in the status bar", now.value === "=SEQUENCE(3,2)" && !now.spilled && now.status === "Array: G13:H15", now);
await select(page, SR + 1, SC + 1); await frames(page);
now = await bar();
check("on a cell the array spilled into: the formula it comes from, greyed, and the range", now.value === "=SEQUENCE(3,2)" && now.spilled && /Spilled from G13/.test(now.title ?? "") && now.status === "Array: G13:H15" && (await cell(page, SR + 1, SC + 1)).formula == null, now);
check("the grid outlines the array of the active cell", (await g(page, () => [...document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-spill-outline')].map((e) => e.dataset.range).join())).includes("G13:H15"));
await shot("spill");
await select(page, SR + 5, SC); await frames(page);
now = await bar();
check("outside the array: nothing of it", now.value === "" && !now.spilled && now.status === null, now);
// typing into a spilled cell from the bar is an entry of its own; Escape gives the grey formula back
await select(page, SR + 2, SC); await frames(page);
await sel(page, "sheet-formula").click();
now = await bar();
check("the bar on a spilled cell starts an empty entry (not the other cell's formula)", now.value === "" && !now.spilled, now);
await page.keyboard.press("Enter"); await frames(page);
await select(page, SR + 2, SC); await frames(page);
now = await bar();
check("entered as it is (empty) nothing changed: the array is intact", now.spilled && (await cell(page, SR + 2, SC)).text === "5" && (await cell(page, SR, SC)).text === "1", { now, anchor: (await cell(page, SR, SC)).text });
await sel(page, "sheet-formula").click();
await page.keyboard.press("Escape"); await frames(page);
now = await bar();
check("Escape: the cell is still the array's, the bar shows the grey formula again", now.value === "=SEQUENCE(3,2)" && now.spilled && (await cell(page, SR + 2, SC)).text === "5", now);
await sel(page, "sheet-formula").click(); await page.keyboard.type("x"); await page.keyboard.press("Enter"); await frames(page);
await select(page, SR, SC); await frames(page);
now = await bar();
check("a value entered into the array's way: the formula cell says #SPILL! and no array is shown", (await cell(page, SR, SC)).text === "#SPILL!" && now.status === null && now.value === "=SEQUENCE(3,2)", { now, text: (await cell(page, SR, SC)).text });
await undo(page); await frames(page);
await select(page, SR + 2, SC); await frames(page);
now = await bar();
check("undo: the array is back", now.spilled && now.status === "Array: G13:H15", now);

// ---- saved and opened again: groups, summary side, the array
await page.keyboard.press("Control+s");
const bytes = await waitUploads(page, 1);
saved.set("new10.xlsx", bytes);
const page2 = await newPage(context);
await open(page2, "sheet-editor", "/api/saved/new10.xlsx", { name: "new10.xlsx" });
await ready(page2);
const tabs = await page2.locator('[data-testid^="sheet-tab-"]:not([data-testid="sheet-tab-add"])').count();
await page2.locator('[data-testid^="sheet-tab-"]:not([data-testid="sheet-tab-add"])').nth(tabs - 1).click(); await frames(page2);
const again = await g(page2, ([r, c]) => { const gr = window.__grids.at(-1); const o = gr.GetOutline(); return { rows: o.rows.length, columns: o.columns.map((x) => x.range).join(), below: o.summaryBelow, right: o.summaryRight, spill: gr.GetCell(r + 1, c + 1).spill, strip: document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-outline-button').length }; }, [SR, SC]);
check("re-opened: the groups, where their summaries are, the array and the strip", again.rows >= 1 && again.columns === "C:D" && again.below === false && again.right === false && again.spill?.range === "G13:H15" && again.strip >= 2, again);
await browser.close();
process.exit(summary() ? 1 : 0);
