import { launch, newPage, open, T, SHOTS } from "./lib.mjs";
import { g, ready, frames } from "./ui-lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`);
await ready(page);
await page.screenshot({ path: `${SHOTS}/probe-line.png`, clip: { x: 0, y: 180, width: 700, height: 200 } });
console.log(JSON.stringify(await g(page, () => { const gr = window.__grids[0]; const at = document.elementsFromPoint(166, 345).slice(0, 6).map((e) => `${e.tagName}.${String(e.className).slice(0, 60)}[${e.dataset.row ?? ""},${e.dataset.col ?? ""}] ${Math.round(e.getBoundingClientRect().left)}-${Math.round(e.getBoundingClientRect().right)} bl=${getComputedStyle(e).borderLeft} bg=${getComputedStyle(e).backgroundColor}`); return { at, frozen: gr.GetViewOptions().frozen, b5: gr.GetCell(4, 1), c5: gr.GetCell(4, 2) }; })).slice(0, 2500));
await browser.close();
