// Helpers for the spreadsheet UI tests.
import { check } from "./lib.mjs";
export const g = (page, fn, arg) => page.evaluate(fn, arg);
export const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
export const ready = async (page) => { await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 180000 }); await page.waitForTimeout(300); };
export const status = (page) => sel(page, "office-editor-status").innerText();
export const frames = (page) => g(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
export const cell = (page, r, c) => g(page, ([r, c]) => window.__grids.at(-1).GetCell(r, c), [r, c]);
export const select = (page, r1, c1, r2 = r1, c2 = c1) => g(page, (a) => { const gr = window.__grids.at(-1); if (a[0] === a[2] && a[1] === a[3]) gr.SelectCell(a[0], a[1]); else gr.SelectRange(...a); }, [r1, c1, r2, c2]);
export const selection = (page) => g(page, () => window.__grids.at(-1).GetSelectionRef());
export const revision = (page) => g(page, () => window.__grids.at(-1).GetRevision());
export const menu = async (page, which, item) => { await sel(page, `sheet-menu-button-${which}`).click(); await sel(page, `sheet-${item}`).click(); };
export const undo = async (page) => { await sel(page, "sheet-undo").click(); };
export const redo = async (page) => { await sel(page, "sheet-redo").click(); };
/** The middle of a cell on screen (scrolled into view first). */
export const cellPoint = async (page, r, c) => {
  const find = () => g(page, ([r, c]) => { const q = [...document.querySelectorAll(`.spreadsheet-cell[data-row="${r}"][data-col="${c}"]`)].map((e) => e.getBoundingClientRect()).find((q) => q.width > 0 && q.height > 0); const sc = document.querySelector('[data-testid="sheet-grid"]').getBoundingClientRect(); return q && q.bottom < sc.bottom - 4 && q.right < sc.right - 4 ? { x: q.left + q.width / 2, y: q.top + q.height / 2, left: q.left, top: q.top, right: q.right, bottom: q.bottom } : null; }, [r, c]);
  let p = await find();
  if (!p) { await g(page, ([r, c]) => { const gr = window.__grids.at(-1); gr.ScrollToCell(Math.min(r + 8, gr.GetModel().RowCount - 1), Math.min(c + 3, gr.GetModel().ColCount - 1)); gr.ScrollToCell(Math.max(0, r - 3), Math.max(0, c - 2)); }, [r, c]); await frames(page); p = await find(); }
  return p;
};
export const gridFocused = (page) => g(page, () => !!document.activeElement?.closest('[data-testid="sheet-grid"]'));
/**
 * One feature: `act` through the UI, then `read` the model. Asserts the model
 * changed to what `expect` accepts, that one Undo brings back what was read
 * before, and that Redo applies it again.
 */
export async function feature(page, name, { read, act, expect, steps = 1 }) {
  const before = JSON.stringify(await read());
  const r0 = await revision(page);
  await act();
  await frames(page);
  const after = await read();
  const ok = check(`${name}: the model changed`, JSON.stringify(after) !== before && (await expect(after)), after);
  for (let i = 0; i < steps; i++) await undo(page);
  await frames(page);
  const undone = JSON.stringify(await read());
  check(`${name}: undo reverts it`, undone === before && (await revision(page)) === r0, { undone: undone.slice(0, 200), revision: [r0, await revision(page)] });
  for (let i = 0; i < steps; i++) await redo(page);
  await frames(page);
  check(`${name}: redo applies it again`, JSON.stringify(await read()) === JSON.stringify(after));
  return ok;
}
