// What 0.1.9 adds to the host: quick sort by the component, tables, data-bar and icon-set options.
import { launch, newPage, open, T, S, SHOTS, check, summary, waitUploads, saved } from "./lib.mjs";
import { g, sel, ready, status, frames, cell, select, selection, menu, undo, redo, cellPoint, gridFocused, feature, revision } from "./ui-lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`);
await ready(page);
const R = 100, C = 17;
const data = [["Name", "Qty", "Price"], ["pear", "3", "1.5"], ["apple", "10", "2"], ["fig", "7", "9"], ["apple", "2", "4"]];
await g(page, ([R, C, data]) => { const gr = window.__grids[0]; data.forEach((row, i) => row.forEach((v, j) => gr.SetCell(R + i, C + j, v))); gr.ScrollToCell(R + 14, C + 6); gr.ScrollToCell(R - 2, C - 3); gr.SelectCell(R + 1, C); }, [R, C, data]);
await frames(page);
const block = (rows = 5) => g(page, ([R, C, rows]) => { const gr = window.__grids[0]; const out = []; for (let i = 0; i < rows; i++) out.push([0, 1, 2].map((j) => gr.GetCell(R + i, C + j).text).join(",")); return out.join("|"); }, [R, C, rows]);
const tables = () => g(page, () => window.__grids[0].GetTables());
const shot = async (name) => { await page.mouse.move(5, 5); await page.waitForTimeout(350); await page.screenshot({ path: `${SHOTS}/new-${name}.png` }); };
const notice = () => sel(page, "sheet-notice").innerText().catch(() => "");

// ---------------- quick sort: the component's (current region, header kept)
await select(page, R + 2, C + 1);
await feature(page, "Sort A to Z (Data menu): a block away from row 1 is sorted by the cell's column, its header row stays", { read: () => block(), act: () => menu(page, "data", "sort-asc"), expect: (v) => v === "Name,Qty,Price|apple,2,4|pear,3,1.5|fig,7,9|apple,10,2" });
await undo(page);
await select(page, R + 3, C);
await feature(page, "Sort Z to A (Data menu) by a text column", { read: () => block(), act: () => menu(page, "data", "sort-desc"), expect: (v) => v === "Name,Qty,Price|pear,3,1.5|fig,7,9|apple,10,2|apple,2,4" });
await undo(page);
// a multi-cell selection: still the region around the active cell (not the selected cells alone)
await select(page, R + 1, C + 1, R + 2, C + 1);
await feature(page, "Sort with two cells selected sorts the whole block (rows stay together)", { read: () => block(), act: () => menu(page, "data", "sort-asc"), expect: (v) => v === "Name,Qty,Price|apple,2,4|pear,3,1.5|fig,7,9|apple,10,2" });
await undo(page);
// the cell context menu
{
  await select(page, R + 2, C + 2);
  const p = await cellPoint(page, R + 2, C + 2);
  await page.mouse.click(p.x, p.y, { button: "right" });
  await page.locator('[role="menu"] [data-testid$="sort-desc"], [role="menu"] >> text="Sort Z to A"').first().click();
  await frames(page);
  check("Sort Z to A from the cell's context menu", (await block()) === "Name,Qty,Price|fig,7,9|apple,2,4|apple,10,2|pear,3,1.5", await block());
  await undo(page);
}
{
  await select(page, R + 40, C + 8);
  const r0 = await revision(page);
  await menu(page, "data", "sort-asc");
  await frames(page);
  check("Sort in an empty area: nothing changes and a notice says why", (await revision(page)) === r0 && /no data around/.test(await notice()), await notice());
}

// ---------------- tables
await select(page, R + 1, C);
await menu(page, "insert", "insert-table");
await sel(page, "sheet-tables-dialog").waitFor();
check("Insert > Table...: the data around the cell is proposed, with headers and Medium 2", (await sel(page, "sheet-table-range").inputValue()) === "R101:T105" && (await sel(page, "sheet-table-headers").isChecked()) && (await page.locator('.rb-sheet-table-style.is-active').getAttribute("data-style")) === "TableStyleMedium2" && (await page.locator(".rb-sheet-table-style").count()) === 28, { range: await sel(page, "sheet-table-range").inputValue(), styles: await page.locator(".rb-sheet-table-style").count() });
await shot("table-new");
const headerBackground = () => g(page, ([R, C]) => { const el = [...document.querySelectorAll(`.spreadsheet-cell[data-row="${R}"][data-col="${C}"]`)].find((e) => e.getBoundingClientRect().width > 0); return el ? getComputedStyle(el).backgroundColor : null; }, [R, C]);
const plainHeader = await headerBackground();
const r1 = await revision(page);
await sel(page, "sheet-dialog-ok").click();
await frames(page);
let t = await tables();
check("Create table: the table is in the workbook (name, range, columns from the header row, style)", t.length === 1 && t[0].name === "Table1" && t[0].range === "R101:T105" && t[0].hasHeaders && t[0].columns.join() === "Name,Qty,Price" && t[0].style?.name === "TableStyleMedium2" && t[0].style.showRowStripes, t);
check("the dialog now shows that table (Apply, its cells, Total row)", (await sel(page, "sheet-dialog-ok").innerText()) === "Apply" && (await sel(page, "sheet-table-name").inputValue()) === "Table1" && (await sel(page, "sheet-table-cells").innerText()) === "R101:T105" && (await sel(page, "sheet-table-totals").count()) === 1 && (await page.locator('[data-testid="sheet-tables-list"] [aria-selected="true"]').getAttribute("data-table")) === "Table1");
check("the grid draws the style (the header cell is filled)", (await headerBackground()) !== plainHeader, { before: plainHeader, after: await headerBackground() });
await shot("table-created");
// another style: Light 9, no bands
await sel(page, "sheet-table-family-Light").click();
check("Light has 21 styles, Dark 11", (await page.locator(".rb-sheet-table-style").count()) === 21);
await shot("table-styles-light");
await page.locator('[data-style="TableStyleLight9"]').click();
await sel(page, "sheet-table-banded").uncheck();
await sel(page, "sheet-dialog-ok").click();
await frames(page);
t = await tables();
check("Apply: style Light 9 without banded rows", t[0].style?.name === "TableStyleLight9" && t[0].style.showRowStripes === false, t[0].style);
await sel(page, "sheet-table-family-Dark").click();
check("Dark has 11 styles", (await page.locator(".rb-sheet-table-style").count()) === 11);
await shot("table-styles-dark");
await page.locator('[data-style="TableStyleDark3"]').click();
await sel(page, "sheet-table-banded").check();
// rename and total row in one Apply
await sel(page, "sheet-table-name").fill("Fruit");
await sel(page, "sheet-table-totals").check();
await sel(page, "sheet-dialog-ok").click();
await frames(page);
t = await tables();
check("Apply: renamed, Dark 3, total row shown (the range grew by it)", t[0].name === "Fruit" && t[0].style?.name === "TableStyleDark3" && t[0].hasTotals && t[0].range === "R101:T106" && (await cell(page, R + 5, C)).text === "Total", { t: t[0], total: (await cell(page, R + 5, C)).text, sum: (await cell(page, R + 5, C + 2)).text });
check("the total row sums the last column", (await cell(page, R + 5, C + 2)).text === "16.5" && /SUBTOTAL\(109,Fruit\[Price\]\)/.test((await cell(page, R + 5, C + 2)).formula), await cell(page, R + 5, C + 2));
// a taken / invalid name is refused in the dialog
await sel(page, "sheet-table-name").fill("A1");
await sel(page, "sheet-dialog-ok").click();
check("an invalid name is refused and shown in the dialog", (await sel(page, "sheet-dialog-error").count()) === 1 && (await tables())[0].name === "Fruit", await sel(page, "sheet-dialog-error").innerText().catch(() => "-"));
await shot("table-error");
await sel(page, "sheet-table-name").fill("Fruit");
await sel(page, "sheet-table-totals").uncheck();
await sel(page, "sheet-dialog-ok").click();
await frames(page);
check("Total row off again", !(await tables())[0].hasTotals && (await tables())[0].range === "R101:T105" && (await cell(page, R + 5, C)).text === "");
await sel(page, "sheet-dialog-cancel").click();
await frames(page);
check("the dialog closed and the grid has the keyboard", (await sel(page, "sheet-tables-dialog").count()) === 0 && (await gridFocused(page)));
await shot("table-dark-grid");
// a structured reference, then typing under the table (0.1.9: the table follows)
await g(page, ([R, C]) => window.__grids[0].EnterText(R, C + 5, "=SUM(Fruit[Qty])"), [R, C]);
check("a structured reference reads the table", (await cell(page, R, C + 5)).text === "22", await cell(page, R, C + 5));
{
  const p = await cellPoint(page, R + 5, C);
  await page.mouse.click(p.x, p.y);
  const r0 = await revision(page);
  await page.keyboard.type("kiwi");
  await page.keyboard.press("Tab");
  await page.keyboard.type("5");
  await page.keyboard.press("Enter");
  await frames(page);
  t = await tables();
  check("typing under the table extends it; the formula follows", t[0].range === "R101:T106" && (await cell(page, R, C + 5)).text === "27", { range: t[0].range, sum: (await cell(page, R, C + 5)).text });
  await undo(page); await undo(page);
  await frames(page);
  check("two undos (one per entry) take the row out of the table again", (await tables())[0].range === "R101:T105" && (await cell(page, R, C + 5)).text === "22" && (await revision(page)) === r0, { range: (await tables())[0].range, revision: [r0, await revision(page)] });
}
// renaming a column by typing in the header cell
{
  await g(page, ([R, C]) => { const gr = window.__grids[0]; gr.ScrollToCell(R + 14, C + 6); gr.ScrollToCell(R - 2, C - 3); }, [R, C]);
  await frames(page);
  const p = await cellPoint(page, R, C + 1);
  await page.mouse.click(p.x, p.y);
  await page.keyboard.type("Count");
  await page.keyboard.press("Enter");
  await frames(page);
  check("typing in a header cell renames the column; the structured reference follows", (await tables())[0].columns.join() === "Name,Count,Price" && (await cell(page, R, C + 5)).formula.includes("Fruit[Count]"), { columns: (await tables())[0].columns, formula: (await cell(page, R, C + 5)).formula });
}
// the Data menu opens the same dialog on the table under the cell; convert to range
await select(page, R + 2, C + 1);
await menu(page, "data", "tables");
await sel(page, "sheet-tables-dialog").waitFor();
check("Data > Tables... shows the table the cell is in", (await sel(page, "sheet-table-name").inputValue()) === "Fruit" && (await page.locator('.rb-sheet-table-style.is-active').getAttribute("data-style")) === "TableStyleDark3");
await shot("table-edit");
await sel(page, "sheet-tables-convert").click();
await frames(page);
check("Convert to range: no table, the cells stay, the formula names the cells", (await tables()).length === 0 && (await block()).startsWith("Name,Count,Price|pear,3,1.5") && /SUM\(\$S\$102:\$S\$105\)/.test((await cell(page, R, C + 5)).formula) && (await sel(page, "sheet-dialog-ok").innerText()) === "Create table", { tables: await tables(), formula: (await cell(page, R, C + 5)).formula });
await page.keyboard.press("Escape");
await frames(page);
check("Escape closes the dialog only", (await sel(page, "sheet-tables-dialog").count()) === 0 && (await sel(page, "office-editor").count()) === 1 && (await gridFocused(page)));
await undo(page);
await frames(page);
check("Undo brings the table back", (await tables()).length === 1 && (await tables())[0].name === "Fruit" && (await cell(page, R, C + 5)).formula.includes("Fruit[Count]"), await tables());
// no headers
await g(page, ([R, C]) => { const gr = window.__grids[0]; [[1, 2], [3, 4]].forEach((row, i) => row.forEach((v, j) => gr.SetCell(R + 12 + i, C + j, String(v)))); gr.SelectRange(R + 12, C, R + 13, C + 1); }, [R, C]);
await menu(page, "insert", "insert-table");
await sel(page, "sheet-tables-dialog").waitFor();
check("with cells selected the selection is proposed", (await sel(page, "sheet-table-range").inputValue()) === "R113:S114");
await sel(page, "sheet-table-headers").uncheck();
await sel(page, "sheet-table-name").fill("Plain");
await sel(page, "sheet-table-family-none").click();
await sel(page, "sheet-dialog-ok").click();
await frames(page);
t = (await tables()).find((x) => x.name === "Plain");
check("a table without headers and without a style", t && !t.hasHeaders && t.columns.join() === "Column1,Column2" && t.style == null && t.range === "R113:S114", t);
await sel(page, "sheet-dialog-cancel").click();

// ---------------- conditional formats: data-bar and icon-set options
const rules = () => g(page, () => window.__grids[0].GetConditionalFormats().map((r) => r.rule));
await g(page, ([R, C]) => { const gr = window.__grids[0]; [-4, 2, 6, 9].forEach((v, i) => { gr.SetCell(R + 20 + i, C, String(v)); gr.SetCell(R + 20 + i, C + 1, String(v)); gr.SetCell(R + 20 + i, C + 2, String(v)); gr.SetCell(R + 20 + i, C + 3, String(v)); }); gr.ScrollToCell(R + 34, C + 6); gr.ScrollToCell(R + 14, C - 3); gr.SelectRange(R + 20, C, R + 23, C); }, [R, C]);
const before = (await rules()).length;
await menu(page, "data", "conditional-formats");
await sel(page, "sheet-cf-dialog").waitFor();
const addRule = async (fill) => { await sel(page, "sheet-cf-add").click(); await sel(page, "sheet-cf-rule-dialog").waitFor(); await fill(); await sel(page, "sheet-dialog-ok").click(); await sel(page, "sheet-cf-dialog").waitFor({ timeout: 4000 }).catch(async () => console.log("  rule not accepted:", await sel(page, "sheet-dialog-error").innerText().catch(() => "?"))); };
await addRule(async () => { await sel(page, "sheet-cf-kind").selectOption("dataBar"); });
let all = await rules();
check("a data bar with no option touched is the plain (2007) one", all.length === before + 1 && all.at(-1).type === "dataBar" && !("gradient" in all.at(-1)), all.at(-1));
await addRule(async () => {
  await sel(page, "sheet-cf-range").fill("S121:S124");
  await sel(page, "sheet-cf-kind").selectOption("dataBar");
  check("the plain bar's form: solid, no own negative colour, bars from the edge, value shown", (await sel(page, "sheet-cf-bar-fill").inputValue()) === "solid" && !(await sel(page, "sheet-cf-bar-negative-use").isChecked()) && (await sel(page, "sheet-cf-bar-axis").inputValue()) === "none" && (await sel(page, "sheet-cf-bar-show-value").isChecked()));
  await sel(page, "sheet-cf-bar-fill").selectOption("gradient");
  await sel(page, "sheet-cf-bar-negative-use").check();
  await sel(page, "sheet-cf-bar-axis").selectOption("automatic");
  await sel(page, "sheet-cf-bar-direction").selectOption("rightToLeft");
  await sel(page, "sheet-cf-bar-show-value").uncheck();
  await shot("cf-data-bar");
});
all = await rules();
let bar = all.find((r) => r.type === "dataBar" && r.gradient !== undefined);
check("data bar options are in the rule (gradient, negative colour, axis, direction, bar only)", bar && bar.gradient === true && bar.negativeColor === "#FF0000" && bar.axisPosition === "automatic" && bar.direction === "rightToLeft" && bar.showValue === false, bar);
check("the list says what the bar is", (await page.locator('[data-testid="sheet-cf-list"] [role="option"]', { hasText: "Data bar, gradient, bar only" }).count()) === 1, await sel(page, "sheet-cf-list").innerText());
// edit it: the form shows the options; one change keeps the rest
await page.locator('[data-testid="sheet-cf-list"] [role="option"]', { hasText: "Data bar, gradient, bar only" }).click();
await sel(page, "sheet-cf-edit").click();
await sel(page, "sheet-cf-rule-dialog").waitFor();
check("Edit shows the bar's options", (await sel(page, "sheet-cf-bar-fill").inputValue()) === "gradient" && (await sel(page, "sheet-cf-bar-negative-use").isChecked()) && (await sel(page, "sheet-cf-bar-negative").inputValue()).toUpperCase() === "#FF0000" && (await sel(page, "sheet-cf-bar-axis").inputValue()) === "automatic" && (await sel(page, "sheet-cf-bar-direction").inputValue()) === "rightToLeft" && !(await sel(page, "sheet-cf-bar-show-value").isChecked()));
await sel(page, "sheet-cf-bar-axis").selectOption("middle");
await sel(page, "sheet-cf-bar-fill").selectOption("solid");
await sel(page, "sheet-dialog-ok").click();
await sel(page, "sheet-cf-dialog").waitFor();
const bar2 = (await rules()).find((r) => r.type === "dataBar" && r.gradient !== undefined);
check("one option changed, the others kept (also the ones the form has no field for)", bar2.axisPosition === "middle" && bar2.gradient === false && bar2.direction === "rightToLeft" && bar2.showValue === false && bar2.negativeColor === "#FF0000" && bar2.minLength === bar.minLength && bar2.maxLength === bar.maxLength && JSON.stringify(bar2.min) === JSON.stringify(bar.min), bar2);
// the plain bar: a new colour leaves it plain
await page.locator('[data-testid="sheet-cf-list"] [role="option"]', { hasText: /^.*Data bar\s*R121:R124/ }).click();
await sel(page, "sheet-cf-edit").click();
await sel(page, "sheet-cf-rule-dialog").waitFor();
await sel(page, "sheet-cf-bar").fill("#16a34a");
await sel(page, "sheet-dialog-ok").click();
await sel(page, "sheet-cf-dialog").waitFor();
const plain = (await rules()).find((r) => r.type === "dataBar" && r.gradient === undefined);
check("a plain bar edited for its colour stays plain", plain && plain.color.toUpperCase() === "#16A34A" && !("axisPosition" in plain), plain);
// icon sets
await addRule(async () => {
  await sel(page, "sheet-cf-range").fill("T121:T124");
  await sel(page, "sheet-cf-kind").selectOption("iconSet");
  await sel(page, "sheet-cf-icons").selectOption("3Stars");
  await sel(page, "sheet-cf-icons-show-value").uncheck();
});
let icons = (await rules()).find((r) => r.type === "iconSet" && r.iconSet === "3Stars");
check("icon set 3Stars, icon only", icons && icons.showValue === false && icons.thresholds.length === 3 && !icons.icons, icons);
await addRule(async () => {
  await sel(page, "sheet-cf-range").fill("U121:U124");
  await sel(page, "sheet-cf-kind").selectOption("iconSet");
  await sel(page, "sheet-cf-icons").selectOption("3Arrows");
  await sel(page, "sheet-cf-icons-reverse").check();
  await sel(page, "sheet-cf-icons-custom").check();
  check("custom icons: one box per step, starting as the family's own; Reverse is off and disabled", (await page.locator('[data-testid="sheet-cf-icons-steps"] select').count()) === 3 && (await sel(page, "sheet-cf-icon-0").inputValue()) === "3Arrows:0" && (await sel(page, "sheet-cf-icon-2").inputValue()) === "3Arrows:2" && (await sel(page, "sheet-cf-icons-reverse").isDisabled()) && !(await sel(page, "sheet-cf-icons-reverse").isChecked()));
  await sel(page, "sheet-cf-icon-0").selectOption("NoIcons:0");
  await sel(page, "sheet-cf-icon-2").selectOption("3Flags:2");
  await shot("cf-icon-set");
});
icons = (await rules()).find((r) => r.type === "iconSet" && Array.isArray(r.icons));
check("custom icons are in the rule (no icon for the lowest step, a flag for the highest)", icons && JSON.stringify(icons.icons) === JSON.stringify([{ iconSet: "NoIcons", index: 0 }, { iconSet: "3Arrows", index: 1 }, { iconSet: "3Flags", index: 2 }]) && icons.reverse === false, icons);
await page.locator('[data-testid="sheet-cf-list"] [role="option"]', { hasText: "custom icons" }).click();
await sel(page, "sheet-cf-edit").click();
await sel(page, "sheet-cf-rule-dialog").waitFor();
check("Edit shows the custom icons", (await sel(page, "sheet-cf-icons-custom").isChecked()) && (await sel(page, "sheet-cf-icon-0").inputValue()) === "NoIcons:0" && (await sel(page, "sheet-cf-icon-2").inputValue()) === "3Flags:2");
await sel(page, "sheet-cf-icons").selectOption("5Boxes");
check("a family of five steps: five boxes, the three chosen icons kept", (await page.locator('[data-testid="sheet-cf-icons-steps"] select').count()) === 5 && (await sel(page, "sheet-cf-icon-2").inputValue()) === "3Flags:2" && (await sel(page, "sheet-cf-icon-4").inputValue()) === "5Boxes:4");
await shot("cf-icon-set-5");
await sel(page, "sheet-cf-icons-custom").uncheck();
await sel(page, "sheet-dialog-ok").click();
await sel(page, "sheet-cf-dialog").waitFor();
icons = (await rules()).find((r) => r.type === "iconSet" && r.iconSet === "5Boxes");
check("the family's own icons again: 5Boxes with five thresholds, no custom icons", icons && icons.thresholds.length === 5 && !icons.icons, icons);
await shot("cf-manager");
await sel(page, "sheet-dialog-cancel").click();
await frames(page);
const drawn = await g(page, () => ({ icons: [...document.querySelectorAll("[data-cf-icon-set]")].map((e) => e.getAttribute("data-cf-icon-set")).filter((v, i, a) => a.indexOf(v) === i) }));
check("the grid draws the icons of the new families", drawn.icons.includes("3Stars") && drawn.icons.includes("5Boxes"), drawn);
await select(page, R + 26, C);
await shot("cf-grid");
// save and re-open: the rules and the tables are in the file
const expected = JSON.stringify(await rules());
const expectedTables = JSON.stringify(await tables());
await page.keyboard.press("Control+s");
const bytes = await waitUploads(page, 1);
saved.set("cf", bytes);
const again = await newPage(context);
await open(again, "sheet-editor", "/api/saved/cf", { name: "saved.xlsx" });
await ready(again);
const reopened = await g(again, () => ({ rules: JSON.stringify(window.__grids[0].GetConditionalFormats().map((r) => r.rule)), tables: JSON.stringify(window.__grids[0].GetTables()) }));
check("saved and re-opened: the same rules (x14 forms written and read)", reopened.rules === expected, reopened.rules === expected ? `${JSON.parse(expected).length} rules` : { expected, got: reopened.rules });
check("saved and re-opened: the same tables", reopened.tables === expectedTables, reopened.tables === expectedTables ? expectedTables : { expectedTables, got: reopened.tables });
await g(again, ([R, C]) => { const gr = window.__grids[0]; gr.ScrollToCell(R + 34, C + 6); gr.ScrollToCell(R - 2, C - 3); gr.SelectCell(R + 30, C + 6); }, [R, C]);
await frames(again);
await again.screenshot({ path: `${SHOTS}/new-reopened.png` });
await browser.close();
process.exit(summary() ? 1 : 0);
