// Spreadsheet UI, part 2: sort, AutoFilter, conditional formats, validation, names, find and replace, comments, links.
import { launch, newPage, open, T, S, SHOTS, check, summary, waitUploads, saved } from "./lib.mjs";
import { g, sel, ready, status, frames, cell, select, selection, menu, undo, redo, cellPoint, gridFocused, feature, revision } from "./ui-lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
page.on("console", (m) => { if (m.text().startsWith("TRACE")) console.log(m.text()); });
await page.addInitScript(() => { window.__opened = []; window.open = (...a) => { window.__opened.push(a); return null; }; });
await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`);
await ready(page);
const R = 100, C = 17;
const table = [["Name", "Qty", "Price"], ["pear", "3", "1.5"], ["apple", "10", "2"], ["fig", "7", "9"], ["apple", "2", "4"]];
await g(page, ([R, C, table]) => { const gr = window.__grids[0]; table.forEach((row, i) => row.forEach((v, j) => gr.EnterText(R + i, C + j, v))); gr.ScrollToCell(R + 14, C + 6); gr.ScrollToCell(R - 2, C - 3); gr.SelectCell(R + 1, C); }, [R, C, table]);
await frames(page);
const column = (j) => g(page, ([R, C, j]) => { const gr = window.__grids[0]; const out = []; for (let i = 1; i <= 4; i++) out.push(gr.GetCell(R + i, C + j).text); return out.join(","); }, [R, C, j]);
const shot = (name) => page.screenshot({ path: `${SHOTS}/ui-${name}.png` });

// ---- sort dialog
await select(page, R + 2, C + 1);
await menu(page, "data", "sort");
await sel(page, "sheet-sort-dialog").waitFor();
check("sort dialog: the data around the cell is the range, its first row taken for headers, the cell's column the key", (await sel(page, "sheet-sort-range").inputValue()) === "R101:T105" && (await sel(page, "sheet-sort-header").isChecked()) && (await sel(page, "sheet-sort-column-0").evaluate((el) => el.selectedOptions[0].textContent)) === "Qty (S)", await sel(page, "sheet-sort-range").inputValue());
await sel(page, "sheet-sort-column-0").selectOption({ label: "Name (R)" });
await sel(page, "sheet-sort-add").click();
await sel(page, "sheet-sort-column-1").selectOption({ label: "Qty (S)" });
await sel(page, "sheet-sort-order-1").selectOption("desc");
await shot("sort-dialog");
await feature(page, "sort by two keys (dialog)", { read: async () => (await column(0)) + "|" + (await column(1)), act: () => sel(page, "sheet-dialog-ok").click(), expect: (v) => v === "apple,apple,fig,pear|10,2,7,3" });
check("the dialog closed and the grid has the keyboard", (await sel(page, "sheet-sort-dialog").count()) === 0 && (await gridFocused(page)));
await select(page, R + 2, C + 2);
await feature(page, "sort Z to A (Data menu)", { read: () => column(2), act: () => menu(page, "data", "sort-desc"), expect: (v) => v === "9,4,2,1.5" });
await undo(page);

// ---- AutoFilter: the grid draws the header buttons (0.1.10), the dropdown is the editor's
const fbtn = (page, col) => page.locator(`[data-testid="sheet-grid"] .spreadsheet-filter-button[data-col="${col}"]`).first();
await select(page, R + 2, C);
const filters = () => g(page, () => { const f = window.__grids[0].GetFilters(); return f ? `${f.range} hidden:${f.hiddenRows.join("/")} ${f.columns.map((c) => c.letter + JSON.stringify(c.criteria)).join(" ")}` : "none"; });
await feature(page, "AutoFilter on the data around the cell (Data menu)", { read: filters, act: () => menu(page, "data", "filter"), expect: (v) => v.startsWith("R101:T105 hidden: ") });
await frames(page);
const buttons = await g(page, ([R, C]) => [...document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-filter-button')].filter((b) => b.getBoundingClientRect().width > 0).map((b) => { const q = b.getBoundingClientRect(); const h = [...document.querySelectorAll(`.spreadsheet-cell[data-row="${R}"][data-col="${b.dataset.col}"]`)].map((e) => e.getBoundingClientRect()).find((x) => x.width > 0); return { col: Number(b.dataset.col), inside: !!h && q.left >= h.left && q.right <= h.right + 1 && q.top >= h.top && q.bottom <= h.bottom + 1 }; }), [R, C]);
{ const dup = await g(page, () => ({ own: document.querySelectorAll('[data-testid^="sheet-filter-button-"], .rb-sheet-filter-button').length, old: [...document.querySelectorAll(".col-filter-btn")].filter((b) => b.getBoundingClientRect().width > 0).length })); check("the editor draws no filter buttons of its own, and the grid's older ones are hidden", dup.own === 0 && dup.old === 0, dup); }
check("a dropdown button sits in each header cell of the filter range", buttons.length === 3 && buttons.every((b) => b.inside) && buttons.map((b) => b.col).join() === `${C},${C + 1},${C + 2}`, buttons);
await fbtn(page, C).click();
await sel(page, "sheet-filter-panel").waitFor();
check("the dropdown lists the column's values", (await g(page, () => [...document.querySelectorAll('[data-testid="sheet-filter-values"] input[data-value]')].map((i) => i.getAttribute("data-value")).join())) === "apple,fig,pear");
await sel(page, "sheet-filter-search").fill("ap");
check("search narrows the list", (await g(page, () => [...document.querySelectorAll('[data-testid="sheet-filter-values"] input[data-value]')].map((i) => i.getAttribute("data-value")).join())) === "apple");
await shot("filter-panel");
await feature(page, "filter by value (header dropdown, search + OK)", { read: filters, act: () => sel(page, "sheet-filter-apply").click(), expect: (v) => v === `R101:T105 hidden:${R + 3}/${R + 4} R{"values":["apple"],"blanks":false}` });
await frames(page);
check("the filtered column's button is marked, the rows are gone from the grid", (await fbtn(page, C).getAttribute("data-active")) === "1" && (await g(page, ([r, c]) => ![...document.querySelectorAll(`.spreadsheet-cell[data-row="${r}"][data-col="${c}"]`)].some((e) => e.getBoundingClientRect().height > 0), [R + 3, C])));
await shot("filter-applied");
// a condition on another column
await fbtn(page, C + 1).click();
await sel(page, "sheet-filter-operator").selectOption("greaterThan");
await sel(page, "sheet-filter-operand").fill("5");
await feature(page, "filter by condition (header dropdown)", { read: filters, act: () => page.keyboard.press("Enter"), expect: (v) => /S\{"conditions":\[\{"operator":"greaterThan","value":"?5"?\}\]/.test(v) && v.includes(`hidden:${R + 2}/${R + 3}/${R + 4}`) });
// clear one column's filter from its dropdown, then all
await fbtn(page, C + 1).click();
await sel(page, "sheet-filter-clear-column").click();
check("Clear filter in the dropdown takes that column's filter off", (await filters()) === `R101:T105 hidden:${R + 3}/${R + 4} R{"values":["apple"],"blanks":false}`, await filters());
await menu(page, "data", "filter-clear");
check("Clear filters shows every row; the AutoFilter stays", (await filters()) === "R101:T105 hidden: ", await filters());
// sort from the dropdown
await fbtn(page, C + 1).click();
await feature(page, "sort from the header dropdown", { read: () => column(1), act: () => sel(page, "sheet-filter-sort-asc").click(), expect: (v) => v === "2,3,7,10" });
// Escape closes the dropdown only
await fbtn(page, C).click();
await page.keyboard.press("Escape");
check("Escape closes the dropdown, not the editor", (await sel(page, "sheet-filter-panel").count()) === 0 && (await sel(page, "office-editor").count()) === 1 && (await gridFocused(page)));

// ---- conditional formats
await select(page, R + 1, C + 1, R + 4, C + 1);
await menu(page, "data", "conditional-formats");
await sel(page, "sheet-cf-dialog").waitFor();
const before = await g(page, () => window.__grids[0].GetConditionalFormats().length);
const rules = () => g(page, () => window.__grids[0].GetConditionalFormats().map((r) => `${r.range}:${r.rule.type}${r.rule.operator ? ":" + r.rule.operator + ":" + r.rule.formula1 : ""}${r.rule.format?.fill ?? ""}${r.rule.color ?? ""}${r.rule.text ?? ""}${r.rule.rank ?? ""}${r.rule.iconSet ?? ""}${r.rule.stops ? r.rule.stops.length : ""}`).join(" | "));
const addRule = async (fill) => { await sel(page, "sheet-cf-add").click(); await sel(page, "sheet-cf-rule-dialog").waitFor(); if (fill) await fill(); await sel(page, "sheet-dialog-ok").click(); await sel(page, "sheet-cf-dialog").waitFor(); };
await sel(page, "sheet-cf-add").click();
await sel(page, "sheet-cf-rule-dialog").waitFor({ timeout: 3000 }).catch(async () => { await page.screenshot({ path: "shots/dbg-cf.png" }); console.log("DBG dialogs", await g(page, () => [...document.querySelectorAll(".rb-office-dialog, .rb-sheet-filter-panel")].map((d) => d.getAttribute("data-testid")).join()), "error", await sel(page, "sheet-dialog-error").innerText().catch(() => "-")); });
check("a new rule starts with the selection as its range", (await sel(page, "sheet-cf-range").inputValue()) === "S102:S105");
await page.click('[data-testid="sheet-cf-fills"] [data-color="#C6EFCE"]');
await sel(page, "sheet-cf-value1").fill("5");
await shot("cf-rule-dialog");
await page.keyboard.press("Enter");
await sel(page, "sheet-cf-dialog").waitFor();
check("cell-value rule added (Enter saves): in the model and in the list", (await rules()).startsWith("S102:S105:cellIs:greaterThan:5#C6EFCE") && (await page.locator('[data-testid="sheet-cf-list"] [role="option"]').count()) === before + 1, await rules());
await addRule(async () => { await sel(page, "sheet-cf-kind").selectOption("dataBar"); });
await addRule(async () => { await sel(page, "sheet-cf-kind").selectOption("colorScale"); await sel(page, "sheet-cf-range").fill("T102:T105"); });
await addRule(async () => { await sel(page, "sheet-cf-kind").selectOption("iconSet"); await sel(page, "sheet-cf-icons").selectOption("3Arrows"); await sel(page, "sheet-cf-range").fill("T102:T105"); });
await addRule(async () => { await sel(page, "sheet-cf-kind").selectOption("containsText"); await sel(page, "sheet-cf-range").fill("R102:R105"); await sel(page, "sheet-cf-text").fill("app"); });
await addRule(async () => { await sel(page, "sheet-cf-kind").selectOption("top"); await sel(page, "sheet-cf-rank").fill("1"); await sel(page, "sheet-cf-range").fill("T102:T105"); await page.click('[data-testid="sheet-cf-fills"] [data-color="#FFEB9C"]'); });
await addRule(async () => { await sel(page, "sheet-cf-kind").selectOption("duplicateValues"); await sel(page, "sheet-cf-range").fill("R102:R105"); await page.click('[data-testid="sheet-cf-fills"] [data-color="#BDD7EE"]'); });
let r = await rules();
check("one rule of each kind: duplicates, top, text contains, icon set, colour scale, data bar, cell value (newest first)", r.split(" | ").slice(0, 7).map((x) => x.split(":").slice(1).join(":").replace(/^[A-Z]+\d+:/, "")).join(" ; ") === "duplicateValues#BDD7EE ; top#FFEB9C1 ; containsText#FFC7CEapp ; iconSet3Arrows ; colorScale2 ; dataBar#638EC6 ; cellIs:greaterThan:5#C6EFCE", r);
await shot("cf-manager");
// reorder: the newest (first) one down
await sel(page, "sheet-cf-down").click();
r = await rules();
check("Down moves the selected rule after the next one", r.split(" | ")[1].includes("duplicateValues") && r.split(" | ")[0].includes("top"), r.split(" | ").slice(0, 2));
// edit the cell-value rule
await page.locator('[data-testid="sheet-cf-list"] [role="option"]', { hasText: "Cell value greater than 5" }).click();
await sel(page, "sheet-cf-edit").click();
check("Edit opens the rule with its values", (await sel(page, "sheet-cf-value1").inputValue()) === "5" && (await sel(page, "sheet-cf-operator").inputValue()) === "greaterThan" && (await sel(page, "sheet-cf-range").inputValue()) === "S102:S105");
await sel(page, "sheet-cf-operator").selectOption("between");
await sel(page, "sheet-cf-value1").fill("3");
await sel(page, "sheet-cf-value2").fill("8");
await sel(page, "sheet-dialog-ok").click();
await sel(page, "sheet-cf-dialog").waitFor();
check("the edited rule is updated, not added", (await g(page, () => window.__grids[0].GetConditionalFormats().length)) === before + 7 && (await g(page, () => { const x = window.__grids[0].GetConditionalFormats().find((r) => r.rule.type === "cellIs" && r.range === "S102:S105").rule; return `${x.operator}:${x.formula1}:${x.formula2}`; })) === "between:3:8");
// an invalid rule is refused in the dialog
await sel(page, "sheet-cf-add").click();
await sel(page, "sheet-cf-range").fill("not a range");
await sel(page, "sheet-dialog-ok").click();
check("a rule the component refuses stays open with the reason", (await sel(page, "sheet-dialog-error").count()) === 1 && (await sel(page, "sheet-cf-rule-dialog").count()) === 1, await sel(page, "sheet-dialog-error").innerText().catch(() => ""));
await sel(page, "sheet-dialog-cancel").click();
// delete the data bar
await page.locator('[data-testid="sheet-cf-list"] [role="option"]', { hasText: "Data bar" }).click();
await sel(page, "sheet-cf-delete").click();
check("Delete removes the selected rule", (await g(page, () => window.__grids[0].GetConditionalFormats().length)) === before + 6 && !(await rules()).includes("dataBar"));
await page.keyboard.press("Escape");
check("Escape closes the manager; the grid has the keyboard", (await sel(page, "sheet-cf-dialog").count()) === 0 && (await gridFocused(page)));
await frames(page);
// what the grid shows: 7 is between 3 and 8 -> green fill
const fills = await g(page, ([R, C]) => { const out = {}; for (let i = 1; i <= 4; i++) { const gr = window.__grids[0]; const v = gr.GetCell(R + i, C + 1).value; const el = [...document.querySelectorAll(`.spreadsheet-cell[data-row="${R + i}"][data-col="${C + 1}"]`)].find((e) => e.getBoundingClientRect().width > 0); out[v] = el ? getComputedStyle(el).backgroundColor : null; } return out; }, [R, C]);
check("the grid draws the rule: 3 and 7 are filled green, 2 and 10 are not", fills[7] === "rgb(198, 239, 206)" && fills[3] === "rgb(198, 239, 206)" && fills[10] !== "rgb(198, 239, 206)" && fills[2] !== "rgb(198, 239, 206)", fills);
await shot("cf-in-grid");
// undo: each change of the manager was one step
const count = () => g(page, () => window.__grids[0].GetConditionalFormats().length);
await undo(page);
check("undo brings the deleted rule back", (await count()) === before + 7 && (await rules()).includes("dataBar"));
await undo(page);
check("undo takes the edit back", (await rules()).includes("cellIs:greaterThan:5"));
await redo(page); await redo(page);

// ---- data validation
await select(page, R + 1, C + 4, R + 4, C + 4);
const validation = () => g(page, ([r, c]) => { const v = window.__grids[0].GetValidationAt(r, c); return v ? `${v.range}:${v.type}:${(v.values ?? []).join("/")}${v.operator ?? ""}${v.formula1 ?? ""}-${v.formula2 ?? ""}:${v.errorStyle}:${v.errorMessage ?? ""}:${v.inputMessage ?? ""}` : "none"; }, [R + 1, C + 4]);
await feature(page, "list validation (dialog)", {
  read: validation,
  act: async () => { await menu(page, "data", "validation"); await sel(page, "sheet-validation-dialog").waitFor(); await sel(page, "sheet-validation-type").selectOption("list"); await sel(page, "sheet-validation-list").fill("Yes, No, Maybe"); await sel(page, "sheet-validation-input-message").fill("Pick one"); await sel(page, "sheet-validation-error-message").fill("Yes, No or Maybe"); await shot("validation-dialog"); await sel(page, "sheet-dialog-ok").click(); },
  expect: (v) => v === "V102:V105:list:Yes/No/Maybe-:stop:Yes, No or Maybe:Pick one",
});
// the rule refuses an entry typed in the grid
await select(page, R + 1, C + 4);
await page.locator(".spreadsheet-content").first().focus();
await page.keyboard.type("Never");
await page.keyboard.press("Enter");
check("an entry that is not in the list is refused, with the rule's message", (await cell(page, R + 1, C + 4)).value === null && (await sel(page, "sheet-notice").innerText()).includes("Yes, No or Maybe"), await sel(page, "sheet-notice").innerText().catch(() => ""));
// the dropdown in the cell
await select(page, R + 1, C + 4);
await frames(page);
await page.locator(".spreadsheet-list-button").click();
await page.locator(".validation-list-item", { hasText: "No" }).first().click();
check("the cell's dropdown enters a value of the list", (await cell(page, R + 1, C + 4)).value === "No", (await cell(page, R + 1, C + 4)).value);
// a number rule on other cells, edited through the dialog again
await select(page, R + 1, C + 5, R + 4, C + 5);
await menu(page, "data", "validation");
await sel(page, "sheet-validation-type").selectOption("whole");
await sel(page, "sheet-validation-value1").fill("1");
await sel(page, "sheet-validation-value2").fill("10");
await sel(page, "sheet-validation-style").selectOption("warning");
await page.keyboard.press("Enter");
check("whole-number rule between 1 and 10 as a warning", (await g(page, ([r, c]) => { const v = window.__grids[0].GetValidationAt(r, c); return `${v.type}:${v.operator}:${v.formula1}:${v.formula2}:${v.errorStyle}`; }, [R + 2, C + 5])) === "whole:between:1:10:warning");
await select(page, R + 2, C + 5);
await menu(page, "data", "validation");
check("the dialog shows the rule of the cell", (await sel(page, "sheet-validation-type").inputValue()) === "whole" && (await sel(page, "sheet-validation-value2").inputValue()) === "10" && (await sel(page, "sheet-validation-range").inputValue()) === "W102:W105");
await feature(page, "clear validation (dialog)", { read: () => g(page, ([r, c]) => !!window.__grids[0].GetValidationAt(r, c), [R + 2, C + 5]), act: () => sel(page, "sheet-validation-clear").click(), expect: (v) => v === false });
await undo(page);

// ---- defined names and the name box
await select(page, R + 1, C + 1, R + 4, C + 1);
const names = () => g(page, () => window.__grids[0].GetDefinedNames().filter((n) => !n.isBuiltIn && !n.hidden).map((n) => `${n.name}=${n.refersTo}@${n.scopeName ?? "wb"}`).join(" | "));
const namesBefore = await names();
await menu(page, "data", "names");
await sel(page, "sheet-names-dialog").waitFor();
check("a new name offers the selection, absolute and with its sheet", (await sel(page, "sheet-names-refers").inputValue()) === "Annual!$S$102:$S$105", await sel(page, "sheet-names-refers").inputValue());
await sel(page, "sheet-names-name").fill("Quantities");
await shot("names-dialog");
await page.keyboard.press("Enter");
check("name added (Enter): in the model and in the list", (await names()).includes("Quantities=Annual!$S$102:$S$105@wb") && (await page.locator('[data-testid="sheet-names-list"] [data-name="Quantities"]').count()) === 1, await names());
await sel(page, "sheet-names-name").fill("1bad");
await sel(page, "sheet-dialog-ok").click();
check("a name the component refuses is reported in the dialog", (await sel(page, "sheet-dialog-error").count()) === 1, await sel(page, "sheet-dialog-error").innerText().catch(() => ""));
await page.keyboard.press("Escape");
await undo(page);
check("undo removes the name", (await names()) === namesBefore);
await redo(page);
// used in a formula, and reached from the name box
await g(page, ([r, c]) => window.__grids[0].EnterText(r, c, "=SUM(Quantities)"), [R + 6, C + 1]);
check("a formula can use the name", (await cell(page, R + 6, C + 1)).value === 22);
await sel(page, "sheet-cell-ref").click();
await page.keyboard.type("Quantities");
await page.keyboard.press("Enter");
check("the name box goes to the name", (await selection(page)) === "S102:S105" && (await gridFocused(page)));
check("the name box offers the workbook's names", (await g(page, () => [...document.querySelectorAll("#rb-sheet-names option")].map((o) => o.value))).includes("Quantities"));
// edit and delete in the manager
await menu(page, "data", "names");
await page.locator('[data-testid="sheet-names-list"] [data-name="Quantities"]').click();
await sel(page, "sheet-names-refers").fill("Annual!$S$102:$S$104");
await sel(page, "sheet-dialog-ok").click();
check("changing what a name refers to recalculates the formula", (await names()).includes("Quantities=Annual!$S$102:$S$104") && (await cell(page, R + 6, C + 1)).value === 12, (await cell(page, R + 6, C + 1)).value);
await sel(page, "sheet-names-goto").click();
check("Go to selects the name's range and closes the dialog", (await selection(page)) === "S102:S104" && (await sel(page, "sheet-names-dialog").count()) === 0);

// ---- find and replace
await select(page, 0, 0);
await sel(page, "sheet-find").click();
await sel(page, "sheet-find-panel").waitFor();
await sel(page, "sheet-find-input").fill("apple");
await page.keyboard.press("Enter");
check("find: Enter goes to the first match and says how many there are", (await selection(page)) === "R102" && /2 found/.test(await sel(page, "sheet-find-status").innerText()), [await selection(page), await sel(page, "sheet-find-status").innerText()]);
await page.keyboard.press("Enter");
check("find next", (await selection(page)) === "R105", await selection(page));
await sel(page, "sheet-find-previous").click();
check("find previous", (await selection(page)) === "R102");
await sel(page, "sheet-replace-input").fill("plum");
await shot("find-panel");
await feature(page, "replace one match", { read: () => column(0), act: () => sel(page, "sheet-replace-one").click(), expect: (v) => v === "plum,pear,fig,apple" });
await feature(page, "replace all", { read: () => column(0), act: () => sel(page, "sheet-replace-all").click(), expect: (v) => v === "plum,pear,fig,plum" });
check("replace all says how many cells changed", /1 cell changed/.test(await sel(page, "sheet-find-status").innerText()), await sel(page, "sheet-find-status").innerText());
// all sheets
await sel(page, "sheet-find-input").fill("Milk Street");
await sel(page, "sheet-find-workbook").check();
await sel(page, "sheet-find-next").click();
const wb = await sel(page, "sheet-find-status").innerText();
console.log("  workbook search status:", wb);
check("a search of all sheets reports a match with its sheet and the number found", /^(Annual|Monthly)![A-Z]+\d+ - \d+ found$/.test(wb), wb);
await page.keyboard.press("Escape");
check("Escape closes the find panel; the grid has the keyboard", (await sel(page, "sheet-find-panel").count()) === 0 && (await gridFocused(page)) && (await sel(page, "office-editor").count()) === 1);
await sel(page, "sheet-tab-0").click();

// ---- comments
await g(page, ([R, C]) => { const gr = window.__grids[0]; gr.ScrollToCell(R + 14, C + 6); gr.ScrollToCell(R - 2, C - 3); }, [R, C]);
const comment = () => g(page, ([r, c]) => window.__grids[0].GetComment(r, c)?.text ?? null, [R + 3, C]);
const rightClick = async (r, c, item) => { const p = await cellPoint(page, r, c); await page.mouse.click(p.x, p.y, { button: "right" }); await sel(page, "sheet-context-menu").waitFor(); await sel(page, `sheet-menu-${item}`).click(); };
await feature(page, "new comment (cell menu, dialog)", { read: comment, act: async () => { await rightClick(R + 3, C, "comment"); await sel(page, "sheet-comment-text").fill("Check this figure"); await shot("comment-dialog"); await sel(page, "sheet-dialog-ok").click(); }, expect: (v) => v === "Check this figure" });
await frames(page);
const hover = await g(page, ([r, c]) => { const el = [...document.querySelectorAll(`.spreadsheet-cell[data-row="${r}"][data-col="${c}"]`)].find((e) => e.getBoundingClientRect().width > 0); return el?.getAttribute("data-comment"); }, [R + 3, C]);
check("the cell carries the comment (data-comment; the grid shows it in a box on hover)", hover != null && hover.includes("Check this figure"), hover);
await rightClick(R + 3, C, "comment");
check("the menu says Edit for a cell with a comment, and the dialog has its text", (await sel(page, "sheet-comment-text").inputValue()) === "Check this figure");
await sel(page, "sheet-comment-text").fill("Checked");
await page.keyboard.press("Control+Enter");
check("edited (Ctrl+Enter saves)", (await comment()) === "Checked" && (await sel(page, "sheet-comment-dialog").count()) === 0);
await feature(page, "delete comment (cell menu)", { read: comment, act: () => rightClick(R + 3, C, "delete-comment"), expect: (v) => v === null });
await undo(page);

// ---- hyperlinks
const link = (r, c) => g(page, ([r, c]) => { const h = window.__grids[0].GetHyperlink(r, c); return h ? `${h.target}|${h.display}|${h.tooltip}` : null; }, [r, c]);
await select(page, R + 8, C);
await feature(page, "insert a link (Insert menu, dialog)", { read: () => link(R + 8, C), act: async () => { await menu(page, "insert", "insert-hyperlink"); await sel(page, "sheet-hyperlink-target").fill("https://example.com/report"); await sel(page, "sheet-hyperlink-display").fill("Report"); await sel(page, "sheet-hyperlink-tooltip").fill("Open the report"); await shot("hyperlink-dialog"); await page.keyboard.press("Enter"); }, expect: (v) => v === "https://example.com/report|Report|Open the report" });
check("the cell shows the link's text", (await cell(page, R + 8, C)).text === "Report");
let p = await cellPoint(page, R + 8, C);
await page.keyboard.down("Control"); await page.mouse.click(p.x, p.y); await page.keyboard.up("Control");
check("Ctrl+click on an http link: the editor opens it in a new tab with noopener", JSON.stringify(await g(page, () => window.__opened)) === '[["https://example.com/report","_blank","noopener"]]', await g(page, () => window.__opened));
await g(page, ([r, c]) => window.__grids[0].SetHyperlink(`R${r + 1}`, "javascript:alert(1)", "bad"), [R + 9, C]);
p = await cellPoint(page, R + 9, C);
await page.keyboard.down("Control"); await page.mouse.click(p.x, p.y); await page.keyboard.up("Control");
check("a javascript: link is not opened; a notice says so", (await g(page, () => window.__opened.length)) === 1 && (await sel(page, "sheet-notice").innerText()).includes("was not opened"), await sel(page, "sheet-notice").innerText().catch(() => ""));
await g(page, ([r]) => window.__grids[0].SetHyperlink(`S${r + 1}`, "mailto:someone@example.com", "mail"), [R + 9]);
p = await cellPoint(page, R + 9, C + 1);
await page.keyboard.down("Control"); await page.mouse.click(p.x, p.y); await page.keyboard.up("Control");
check("a mailto: link is opened", (await g(page, () => window.__opened.at(-1)?.[0])) === "mailto:someone@example.com");
await rightClick(R + 8, C, "hyperlink");
check("Edit link shows the link", (await sel(page, "sheet-hyperlink-target").inputValue()) === "https://example.com/report");
await sel(page, "sheet-hyperlink-target").fill("Monthly!B3");
await sel(page, "sheet-dialog-ok").click();
check("the link was changed to a place in the workbook", (await link(R + 8, C)).startsWith("Monthly!B3|"));
await feature(page, "remove a link (cell menu)", { read: () => link(R + 9, C), act: () => rightClick(R + 9, C, "remove-hyperlink"), expect: (v) => v === null });

// ---- save and re-open
const snapshot = (pg) => g(pg, ([R, C]) => { const gr = window.__grids[0]; const f = gr.GetFilters(); return {
  names: gr.GetDefinedNames().filter((n) => !n.isBuiltIn && !n.hidden).map((n) => `${n.name}=${n.refersTo}`).sort().join("|"),
  cf: gr.GetConditionalFormats().map((r) => `${r.range}:${r.rule.type}:${r.rule.operator ?? ""}:${r.rule.formula1 ?? ""}:${r.rule.formula2 ?? ""}:${r.rule.format?.fill ?? ""}:${r.rule.text ?? ""}:${r.rule.iconSet ?? ""}`).join("|"),
  validation: [gr.GetValidationAt(R + 1, C + 4), gr.GetValidationAt(R + 2, C + 5)].map((v) => (v ? `${v.range}:${v.type}:${(v.values ?? []).join("/")}:${v.operator ?? ""}:${v.formula1 ?? ""}:${v.formula2 ?? ""}:${v.errorStyle}:${v.errorMessage ?? ""}:${v.inputMessage ?? ""}` : "none")).join("|"),
  filter: f ? f.range : "none",
  comment: gr.GetComment(R + 3, C)?.text ?? null,
  links: [gr.GetHyperlink(R + 8, C), gr.GetHyperlink(R + 9, C + 1)].map((h) => (h ? `${h.target}|${h.display}|${h.tooltip}` : "none")).join(" ; "),
  cells: [0, 1, 2].map((j) => [1, 2, 3, 4].map((i) => gr.GetCell(R + i, C + j).text).join(",")).join(" / "),
  picked: gr.GetCell(R + 1, C + 4).value,
  sum: gr.GetCell(R + 6, C + 1).value,
}; }, [R, C]);
const want = await snapshot(page);
await sel(page, "office-editor-save").click();
const bytes = await waitUploads(page, 1);
await page.waitForTimeout(200);
check("saved", (await status(page)).includes("Saved") && (await sel(page, "office-editor-error").count()) === 0, await status(page));
saved.set("ui2", bytes);
await page.close();
const page2 = await newPage(context);
await open(page2, "sheet-editor", "/api/saved/ui2", { name: "milk-street-office.xlsx" });
await ready(page2);
const got = await snapshot(page2);
for (const key of Object.keys(want)) check(`re-open: ${key} survives the save`, JSON.stringify(got[key]) === JSON.stringify(want[key]), { saved: String(want[key]).slice(0, 300), reopened: String(got[key]).slice(0, 300) });
await g(page2, ([R, C]) => { const gr = window.__grids[0]; gr.ScrollToCell(R + 14, C + 6); gr.ScrollToCell(R - 2, C - 3); }, [R, C]);
await page2.waitForTimeout(300);
check("re-open: clean, and the filter buttons are drawn again", (await status(page2)).includes("No changes") && (await page2.evaluate(() => window.__grids.at(-1).GetFilterButtonRects().length)) === 3);
await page2.screenshot({ path: `${SHOTS}/ui-reopened.png` });
await page2.close();
summary();
await browser.close();
