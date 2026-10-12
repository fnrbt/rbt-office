import { launch, newPage, open, S } from "./lib.mjs";
import { g, ready, frames } from "./ui-lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "sheet-editor", `${S}/hidden-sheet.xlsx`);
await ready(page);
console.log(JSON.stringify(await g(page, () => { const gr = window.__grids[0]; gr.AddSheet(); gr.SwitchSheet(gr.GetSheetNames().length - 1); gr.EnterText(1, 1, "=SEQUENCE(3)"); gr.SelectCell(2, 1); const shown = gr.GetCurrentCellFormula(); gr.StartEdit(); const seeded = gr.GetEditState ? gr.GetEditState() : null; const ok = gr.CommitEdit(); return { barTextOfSpilledCell: shown, editing: seeded, committed: ok, anchorAfter: gr.GetCell(1, 1).text, b3: gr.GetCell(2, 1), canUndo: gr.CanUndo() }; })).slice(0, 900));
// by keyboard: F2 on a spilled cell, Enter
await g(page, () => { const gr = window.__grids[0]; gr.Undo(); gr.SelectCell(3, 1); });
console.log("before F2:", JSON.stringify(await g(page, () => ({ anchor: window.__grids[0].GetCell(1, 1).text, b4: window.__grids[0].GetCell(3, 1).text }))));
await page.locator('[data-testid="sheet-grid"] .spreadsheet-content').first().focus();
await page.keyboard.press("F2"); await frames(page);
console.log("edit box after F2:", JSON.stringify(await page.evaluate(() => document.querySelector(".spreadsheet-edit-input")?.value)));
await page.keyboard.press("Enter"); await frames(page);
console.log("after F2, Enter:", JSON.stringify(await g(page, () => ({ anchor: window.__grids[0].GetCell(1, 1).text, b4: window.__grids[0].GetCell(3, 1), modified: window.__grids[0].IsModified() }))).slice(0, 600));
await browser.close();
