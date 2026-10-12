// GoTo to a cell beyond the grid's present size: the docs (spreadsheet-host-api.md, "The end of the sheet is Excel's") say
// SelectCell(1048575, 0) and GoTo reach row 1 048 576. usage: node repro-goto.mjs
import { launch, newPage, open, T } from "./lib.mjs";
import { g, ready } from "./ui-lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`); await ready(page);
console.log(JSON.stringify(await g(page, () => {
  const gr = window.__grids.at(-1); const out = {}; const size = () => [gr.GetModel().RowCount, gr.GetModel().ColCount];
  out.sizeAtOpen = size();
  const attempt = (name, f) => { try { f(); out[name] = `ok -> ${gr.GetSelectionRef()} grid ${size()}`; } catch (e) { out[name] = `THROWS ${String(e.message ?? e)}`; } gr.SelectCell(0, 0); };
  attempt('GoTo("AB600")', () => gr.GoTo("AB600"));
  attempt('GoTo("A1048576")', () => gr.GoTo("A1048576"));
  attempt('GoTo("XFD1")', () => gr.GoTo("XFD1"));
  attempt('GoTo("A600:B700")', () => gr.GoTo("A600:B700"));
  attempt("SelectCell(599, 27)", () => gr.SelectCell(599, 27));
  attempt("SelectCell(1048575, 0)", () => gr.SelectCell(1048575, 0));
  attempt("SelectRange(599, 0, 699, 1)", () => gr.SelectRange(599, 0, 699, 1));
  return out;
}), null, 1));
await browser.close();
