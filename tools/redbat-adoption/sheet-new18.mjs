// 0.1.18 in the host, spreadsheet: the View menu's page break preview, the name box going wherever GoTo goes.
import { launch, newPage, open, T, S, check, summary } from "./lib.mjs";
import { g, sel, ready, frames, selection } from "./ui-lib.mjs";
const { browser, context } = await launch({ viewport: { width: 1400, height: 820 } });
const E = `${S}/explore`;
const only = process.argv[2]?.split(","); const want = (k) => !only || only.includes(k);
const size = (page) => g(page, () => { const m = window.__grids.at(-1).GetModel(); return { rows: m.RowCount, cols: m.ColCount }; });
const errorsOf = (page) => { const errors = []; page.on("pageerror", (e) => errors.push(String(e).slice(0, 300))); page.on("console", (m) => { if (m.type() === "error" && !/404/.test(m.text())) errors.push(m.text().slice(0, 300)); }); return errors; };
const viewMenu = async (page) => { await sel(page, "sheet-menu-button-view").click(); await sel(page, "sheet-menu-view").waitFor(); };
const nameBox = async (page, text) => { await sel(page, "sheet-cell-ref").click(); await page.keyboard.press("Control+a"); await page.keyboard.type(text); await page.keyboard.press("Enter"); await frames(page); await page.waitForTimeout(150); };

if (want("preview")) {
  const page = await newPage(context); const errors = errorsOf(page);
  await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`); await ready(page);
  const pages = () => page.locator("div.spreadsheet-page").count();
  await viewMenu(page);
  const item = sel(page, "sheet-page-break-preview");
  check("the View menu offers a page break preview, not checked", (await item.count()) === 1 && (await item.getAttribute("aria-checked")) !== "true", await item.getAttribute("aria-checked"));
  check("no page frames in the normal view", (await pages()) === 0);
  await item.click(); await frames(page); await page.waitForTimeout(200);
  const n = await pages();
  const printPages = await g(page, () => window.__grids.at(-1).GetPrintPages().length);
  check("chosen: the grid is in page break view, a frame per printed page", (await g(page, () => window.__grids.at(-1).GetViewMode())) === "pageBreak" && n >= 1, [n, printPages]);
  const label = await page.locator("div.spreadsheet-page").first().getAttribute("data-page");
  check("the frames are numbered", label === "1", label);
  await page.screenshot({ path: `${E}/n18-sheet-preview.png` });
  await viewMenu(page);
  check("the menu shows it checked", (await sel(page, "sheet-page-break-preview").getAttribute("aria-checked")) === "true", await sel(page, "sheet-page-break-preview").getAttribute("aria-checked"));
  await page.keyboard.press("Escape");
  // editing in the preview
  await g(page, () => window.__grids.at(-1).SelectCell(2, 1)); await page.keyboard.type("in preview"); await page.keyboard.press("Enter"); await frames(page);
  check("cells are edited as before", (await g(page, () => window.__grids.at(-1).GetCell(2, 1).text)) === "in preview");
  check("the view mode is no edit: the undo has the entry only", /Unsaved/.test(await sel(page, "office-editor-status").innerText()));
  await sel(page, "sheet-undo").click(); await frames(page);
  check("undone: nothing to save (choosing the view changed nothing)", /No changes/.test(await sel(page, "office-editor-status").innerText()), await sel(page, "office-editor-status").innerText());
  // another sheet keeps the mode
  const sheets = await g(page, () => window.__grids.at(-1).GetSheetNames().length);
  if (sheets > 1) { await g(page, () => window.__grids.at(-1).SwitchSheet(1)); await frames(page); await page.waitForTimeout(200); await viewMenu(page); check("another sheet: still the preview, and the menu says so", (await sel(page, "sheet-page-break-preview").getAttribute("aria-checked")) === "true"); await page.keyboard.press("Escape"); await g(page, () => window.__grids.at(-1).SwitchSheet(0)); await frames(page); }
  await viewMenu(page); await sel(page, "sheet-page-break-preview").click(); await frames(page); await page.waitForTimeout(150);
  check("chosen again: the normal view, no frames", (await g(page, () => window.__grids.at(-1).GetViewMode())) === "normal" && (await pages()) === 0);
  check("no errors", errors.length === 0, errors.slice(0, 3));
  await page.close();
}

if (want("namebox")) {
  const page = await newPage(context); const errors = errorsOf(page);
  await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`); await ready(page);
  const s0 = await size(page);
  await nameBox(page, "AB3000");
  check("the name box goes beyond the grid: the grid grows to hold it", (await selection(page)) === "AB3000" && (await size(page)).rows >= 3000, [await selection(page), await size(page)]);
  await nameBox(page, "XFD1048576");
  check("to the last cell a sheet has", (await selection(page)) === "XFD1048576", await selection(page));
  await nameBox(page, "A1");
  await frames(page); await page.waitForTimeout(200);
  const back = await size(page);
  check("back at A1 the grid is as small as it was", back.rows === s0.rows && back.cols === s0.cols, [s0, back]);
  await nameBox(page, "XFE1");
  check("a cell no sheet has: a notice, the selection stays", (await selection(page)) === "A1" && /not a reference/.test(await sel(page, "sheet-notice").innerText().catch(() => "")), await sel(page, "sheet-notice").innerText().catch(() => "-"));
  await nameBox(page, "NoSuchName");
  check("a name nobody defined: a notice", /not a reference, a defined name or a table/.test(await sel(page, "sheet-notice").innerText().catch(() => "")));
  check("nothing is modified", /No changes/.test(await sel(page, "office-editor-status").innerText()));
  check("no errors", errors.length === 0, errors.slice(0, 3));
  await page.close();
}
await browser.close();
summary();
