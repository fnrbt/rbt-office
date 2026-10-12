// Exploratory pass on what 0.1.14 changed, spreadsheet. usage: PORT=.. node explore14-sheet.mjs v14|v13 [keys]
import { launch, newPage, open, T } from "./lib.mjs";
import { g, sel, ready, frames, select, cellPoint, selection, status } from "./ui-lib.mjs";
const tag = process.argv[2] ?? "v14"; const only = process.argv[3]?.split(",");
const { browser, context } = await launch({ viewport: { width: 1400, height: 820 } });
const E = `/tmp/redbat-adoption/explore`;
const size = (page) => g(page, () => { const gr = window.__grids.at(-1); const m = gr.GetModel(); const u = gr.GetUsedRange(); return `grid ${m.RowCount}x${m.ColCount}, used ${u ? u.range : "-"}, sel ${gr.GetSelectionRef()}, modified ${gr.IsModified()}`; });
const run = async (key, file, fn) => { if (only && !only.includes(key)) return; const page = await newPage(context); const errors = []; page.on("pageerror", (e) => errors.push(String(e).slice(0, 300))); await open(page, "sheet-editor", file); await ready(page); try { await fn(page, async (name) => { await frames(page); await page.waitForTimeout(250); await page.screenshot({ path: `${E}/${tag}-s-${key}-${name}.png` }); }); } catch (e) { console.log(`  ${key}: script error ${String(e).slice(0, 300)}`); } if (errors.length) console.log("  PAGE ERRORS", key, [...new Set(errors)].slice(0, 3)); await page.close(); };
const click = async (page, r, c) => { const p = await cellPoint(page, r, c); if (p) await page.mouse.click(p.x, p.y); return p; };

await run("milk-end", `${T}/real-world/milk-street-office.xlsx`, async (page, shot) => {
  console.log(tag, "milk at open:", await size(page));
  await click(page, 5, 2);
  await page.keyboard.press("Control+End"); await frames(page); console.log("  Ctrl+End:", await size(page));
  await page.keyboard.press("Control+ArrowDown"); await frames(page); console.log("  Ctrl+Down from there:", await size(page));
  await page.keyboard.press("Control+ArrowDown"); await frames(page); console.log("  Ctrl+Down again:", await size(page));
  await page.keyboard.press("Control+Home"); await frames(page);
  await page.keyboard.press("Control+ArrowRight"); await page.keyboard.press("Control+ArrowRight"); await page.keyboard.press("Control+ArrowRight"); await page.keyboard.press("Control+ArrowRight"); await frames(page); console.log("  Ctrl+Right x4 from A1:", await size(page));
  await shot("1-ctrl-right-end");
  await page.keyboard.press("Control+Home"); await frames(page);
  // wheel to the very end and beyond
  const sc = page.locator(".spreadsheet-container").first(); const box = await sc.boundingBox(); await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  for (let i = 0; i < 60; i++) await page.mouse.wheel(0, 2000);
  await frames(page); await page.waitForTimeout(300);
  console.log("  wheel to the end:", await size(page), await g(page, () => { const s = document.querySelector(".spreadsheet-container"); return `scrollTop ${Math.round(s.scrollTop)} of ${s.scrollHeight - s.clientHeight}`; }));
  for (let i = 0; i < 60; i++) await page.mouse.wheel(0, 2000);
  await frames(page); await page.waitForTimeout(300);
  console.log("  60 more wheels:", await size(page), "status:", (await status(page)).replace(/\n/g, " | "));
  await shot("2-wheeled-past-end");
});
await run("fill", `${T}/real-world/milk-street-office.xlsx`, async (page, shot) => {
  // a fill dragged past the end of the grid's rows in view: type two numbers below the data, drag the fill handle down
  const u = await g(page, () => window.__grids[0].GetUsedRange());
  const r = u.lastRow + 2;
  await select(page, r, 2); await click(page, r, 2);
  await page.keyboard.type("1"); await page.keyboard.press("Enter"); await page.keyboard.type("2"); await page.keyboard.press("Enter");
  await select(page, r, 2, r + 1, 2); await frames(page);
  const h = await g(page, () => { const e = [...document.querySelectorAll('[class*="fill-handle"]')].map((x) => ({ c: x.className, q: x.getBoundingClientRect() })).find((x) => x.q.width > 0); return e ? { cls: String(e.c), x: e.q.left + e.q.width / 2, y: e.q.top + e.q.height / 2 } : null; });
  console.log(tag, "fill handle:", JSON.stringify(h), await size(page));
  if (h) {
    const grid = await page.locator('[data-testid="sheet-grid"]').boundingBox();
    await page.mouse.move(h.x, h.y); await page.mouse.down();
    for (let y = h.y; y < grid.y + grid.height - 6; y += 30) await page.mouse.move(h.x, y, { steps: 2 });
    for (let i = 0; i < 40; i++) { await page.mouse.move(h.x + (i % 2), grid.y + grid.height - 3); await page.waitForTimeout(60); }
    await shot("1-dragging");
    await page.mouse.up(); await frames(page);
    console.log("  after the drag:", await size(page), "last values:", await g(page, () => { const gr = window.__grids[0]; const u = gr.GetUsedRange(); return [u.lastRow - 2, u.lastRow - 1, u.lastRow].map((r) => `C${r + 1}=${gr.GetCell(r, 2).text}`).join(" "); }));
    await shot("2-filled");
    await sel(page, "sheet-undo").click(); await frames(page); console.log("  after one undo:", await size(page));
  }
});
await run("d1904", `${T}/test-documents-excel/poi-47813.xlsx`, async (page, shot) => {
  const info = await g(page, () => { const gr = window.__grids[0]; const u = gr.GetUsedRange(); const out = []; if (u) for (let r = u.firstRow; r <= Math.min(u.lastRow, u.firstRow + 12); r++) for (let c = u.firstCol; c <= Math.min(u.lastCol, u.firstCol + 6); c++) { const x = gr.GetCell(r, c); if (x.valueType !== "empty") out.push(`${String.fromCharCode(65 + c)}${r + 1}=${JSON.stringify(x.value)}→${JSON.stringify(x.text)}${x.formula ? ` (=${x.formula})` : ""}`); } return { sys: gr.GetDateSystem ? gr.GetDateSystem() : "n/a", out: out.slice(0, 24) }; });
  console.log(tag, "poi-47813 date system", info.sys, info.out.join("  "));
  await shot("1-open");
});
await run("d1904b", `${T}/test-documents-excel/poi-47804.xlsx`, async (page, shot) => {
  const info = await g(page, () => { const gr = window.__grids[0]; const u = gr.GetUsedRange(); const out = []; if (u) for (let r = u.firstRow; r <= Math.min(u.lastRow, u.firstRow + 12); r++) for (let c = u.firstCol; c <= Math.min(u.lastCol, u.firstCol + 6); c++) { const x = gr.GetCell(r, c); if (x.valueType !== "empty") out.push(`${String.fromCharCode(65 + c)}${r + 1}=${JSON.stringify(x.value)}→${JSON.stringify(x.text)}${x.formula ? ` (=${x.formula})` : ""}`); } return { sys: gr.GetDateSystem ? gr.GetDateSystem() : "n/a", out: out.slice(0, 24) }; });
  console.log(tag, "poi-47804 date system", info.sys, info.out.join("  "));
  // a date typed into a 1904 workbook
  const u = await g(page, () => window.__grids[0].GetUsedRange() ?? { lastRow: 0 });
  await select(page, u.lastRow + 2, 0); await click(page, u.lastRow + 2, 0);
  await page.keyboard.type("2024-03-15"); await page.keyboard.press("Enter"); await page.keyboard.type("=A" + (u.lastRow + 3) + "+1"); await page.keyboard.press("Enter"); await page.keyboard.type("=YEAR(A" + (u.lastRow + 3) + ")"); await page.keyboard.press("Enter"); await frames(page);
  console.log("  typed 2024-03-15, =that+1, =YEAR(that):", await g(page, (r) => [0, 1, 2].map((k) => { const x = window.__grids[0].GetCell(r + k, 0); return `${JSON.stringify(x.value)}→${JSON.stringify(x.text)}`; }).join("  "), u.lastRow + 2));
  await shot("1-typed-date");
});
await run("charts", `${T}/real-world/chart-zoo.xlsx`, async (page, shot) => {
  const labels = await g(page, () => [...document.querySelectorAll(".spreadsheet-object")].filter((o) => o.getBoundingClientRect().width > 0).map((o) => { const l = o.querySelector(".spreadsheet-object-label"); const or = o.getBoundingClientRect(); const lr = l && l.getBoundingClientRect(); return `${(l ? l.textContent : o.textContent).slice(0, 30)}|label ${l ? "span" : "none"}|${lr ? (lr.left >= or.left - 1 && lr.right <= or.right + 1 ? "inside" : "OUTSIDE") : ""}|${Math.round(or.width)}x${Math.round(or.height)}`; }));
  console.log(tag, "chart-zoo objects:", labels.length, labels.slice(0, 6).join("  "));
  await shot("1-open");
});
await run("dash", `${T}/real-world/dashboard-ecommerce-sales.xlsx`, async (page, shot) => {
  const labels = await g(page, () => [...document.querySelectorAll(".spreadsheet-object")].filter((o) => o.getBoundingClientRect().width > 0).map((o) => { const l = o.querySelector(".spreadsheet-object-label"); const or = o.getBoundingClientRect(); const lr = l && l.getBoundingClientRect(); return `${(l ? l.textContent : o.textContent).slice(0, 24)}|${lr ? (lr.left >= or.left - 1 && lr.right <= or.right + 1 && lr.height < 30 ? "inside, one line" : "OUTSIDE or wrapped " + Math.round(lr.height)) : "no label span"}|${Math.round(or.width)}x${Math.round(or.height)}`; }));
  console.log(tag, "dashboard objects:", labels.length, labels.slice(0, 12).join("  "));
  await shot("1-open");
});
await browser.close();
