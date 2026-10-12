// 0.1.15 / 0.1.16 in the host, spreadsheet: SetCell is an undo step, the sheet's size shrinks again, Ctrl+arrows reach the limit,
// charts and shapes are drawn, formulas take a number format, a drag at the edge scrolls.
import { launch, newPage, open, T, S, check, summary, waitUploads, unzip } from "./lib.mjs";
import { g, sel, ready, frames, select, cellPoint, selection, cell, undo, redo, revision } from "./ui-lib.mjs";
const { browser, context } = await launch({ viewport: { width: 1400, height: 820 } });
const E = `${S}/explore`;
const only = process.argv[2]?.split(","); const want = (k) => !only || only.includes(k);
const size = (page) => g(page, () => { const gr = window.__grids.at(-1); const m = gr.GetModel(); return { rows: m.RowCount, cols: m.ColCount }; });
const click = async (page, r, c) => { const p = await cellPoint(page, r, c); if (p) await page.mouse.click(p.x, p.y); return p; };
const errorsOf = (page) => { const errors = []; page.on("pageerror", (e) => errors.push(String(e).slice(0, 300))); page.on("console", (m) => { if (m.type() === "error" && !/404/.test(m.text())) errors.push(m.text().slice(0, 300)); }); return errors; };

// ---------------- (i) SetCell is one undo step; the host's undo button and dirty state follow ----------------
if (want("i")) {
  const page = await newPage(context);
  await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`); await ready(page);
  const before = await cell(page, 3, 1); const r0 = await revision(page);
  check("i: at open nothing to undo", !(await sel(page, "sheet-undo").isEnabled()));
  await g(page, () => window.__grids.at(-1).SetCell(3, 1, "host write")); await frames(page);
  check("i: after SetCell the undo button is enabled and the document is unsaved", (await sel(page, "sheet-undo").isEnabled()) && /Unsaved/.test(await sel(page, "office-editor-status").innerText()), await sel(page, "office-editor-status").innerText());
  check("i: the revision moved by one", (await revision(page)) !== r0);
  await undo(page); await frames(page);
  const after = await cell(page, 3, 1);
  check("i: one undo takes it back, the document is clean again", JSON.stringify(after.value ?? after.text) === JSON.stringify(before.value ?? before.text) && (await revision(page)) === r0 && /No changes/.test(await sel(page, "office-editor-status").innerText()), [before.text, after.text, await sel(page, "office-editor-status").innerText()]);
  await redo(page); await frames(page);
  check("i: redo applies it again", (await cell(page, 3, 1)).text === "host write");
  await page.close();
}

// ---------------- (j) RowCount / ColCount shrink; Ctrl+arrows reach the sheet's limit; the chrome follows ----------------
if (want("j")) {
  const page = await newPage(context); const errors = errorsOf(page);
  await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`); await ready(page);
  const s0 = await size(page);
  await click(page, 5, 2); await frames(page);
  await page.keyboard.press("Control+End"); await frames(page);
  const end = await selection(page);
  await page.keyboard.press("Control+ArrowDown"); await frames(page); await page.waitForTimeout(150);
  const down = await selection(page); const s1 = await size(page);
  check("j: Ctrl+Down with nothing below goes to row 1048576", /1048576$/.test(down), [end, down]);
  const ref = await sel(page, "sheet-cell-ref").inputValue().catch(async () => sel(page, "sheet-cell-ref").innerText());
  check("j: the name box shows that cell, in full", ref === down, [ref, down]);
  const refBox = await sel(page, "sheet-cell-ref").evaluate((el) => ({ scroll: el.scrollWidth, client: el.clientWidth }));
  check("j: the name box has room for it", refBox.scroll <= refBox.client + 1, refBox);
  check("j: the grid is that long now", s1.rows === 1048576, s1);
  await page.screenshot({ path: `${E}/n16-sheet-j-1-bottom.png` });
  await page.keyboard.press("Control+ArrowRight"); await frames(page); await page.waitForTimeout(150);
  const right = await selection(page); const s2 = await size(page);
  check("j: Ctrl+Right with nothing ahead goes to column XFD", /^XFD/.test(right) && s2.cols === 16384, [right, s2]);
  check("j: the name box shows XFD1048576", (await sel(page, "sheet-cell-ref").inputValue().catch(() => "")) === right, right);
  await page.screenshot({ path: `${E}/n16-sheet-j-2-corner.png` });
  check("j: nothing is modified by going there", /No changes/.test(await sel(page, "office-editor-status").innerText()) && !(await sel(page, "sheet-undo").isEnabled()));
  await page.keyboard.press("Control+Home"); await frames(page); await page.waitForTimeout(200);
  const s3 = await size(page);
  check("j: back at A1 the sheet is as long as it was", s3.rows === s0.rows && s3.cols === s0.cols, [s0, s3]);
  const sc = await g(page, () => { const c = document.querySelector('[data-testid="sheet-grid"] .spreadsheet-content') ?? document.querySelector(".spreadsheet-content"); const all = [...document.querySelectorAll('[data-testid="sheet-grid"] *')].filter((e) => e.scrollHeight > e.clientHeight + 50 && getComputedStyle(e).overflowY !== "visible"); return all.map((e) => ({ cls: e.className.slice(0, 40), scrollHeight: e.scrollHeight, scrollTop: e.scrollTop })); });
  check("j: the scroll extent shrank with it (no million-row scroll bar left)", sc.every((x) => x.scrollHeight < 200000), sc);
  check("j: the name box says A1", (await sel(page, "sheet-cell-ref").inputValue().catch(() => "")) === "A1");
  // name box navigation to a far cell and back
  await sel(page, "sheet-cell-ref").click(); await page.keyboard.press("Control+a"); await page.keyboard.type("ZZ5000"); await page.keyboard.press("Enter"); await frames(page); await page.waitForTimeout(150);
  const far = await size(page);
  check("j: the name box goes to ZZ5000: the grid grows to hold it", (await selection(page)) === "ZZ5000" && far.rows >= 5000 && far.cols >= 702, [await selection(page), far]);
  await page.keyboard.type("x"); await page.keyboard.press("Enter"); await frames(page);
  await page.keyboard.press("Control+Home"); await frames(page); await page.waitForTimeout(150);
  const kept = await size(page);
  check("j: a cell written out there keeps the sheet that large", kept.rows >= 5000 && kept.cols >= 702, kept);
  await sel(page, "sheet-cell-ref").click(); await page.keyboard.press("Control+a"); await page.keyboard.type("b600:c700"); await page.keyboard.press("Enter"); await frames(page);
  check("j: the name box selects a range out there", (await selection(page)) === "B600:C700", await selection(page));
  await sel(page, "sheet-cell-ref").click(); await page.keyboard.press("Control+a"); await page.keyboard.type("XFE1"); await page.keyboard.press("Enter"); await frames(page);
  check("j: a cell no sheet has is refused with a notice", (await selection(page)) === "B600:C700" && /not a reference|outside the sheet/.test(await sel(page, "sheet-notice").innerText().catch(() => "")), await sel(page, "sheet-notice").innerText().catch(() => "-"));
  await page.keyboard.press("Control+Home"); await frames(page);
  await undo(page); await frames(page); await page.waitForTimeout(200);
  const shrunk = await size(page);
  check("j: undone: the sheet is as it was again", shrunk.rows === s0.rows && shrunk.cols === s0.cols, [s0, shrunk]);
  // status bar: a selection of a whole column counts cells without hanging
  await click(page, 6, 2);
  for (let i = 0; i < 400 && !/1048576$/.test(await selection(page)); i++) await page.keyboard.press("Control+Shift+ArrowDown");
  await frames(page); await page.waitForTimeout(200);
  const st = await sel(page, "sheet-status").innerText(); const selTo = await selection(page);
  check("j: Ctrl+Shift+Down to the end of the sheet: the selection reaches row 1048576 and the status bar counts", /^C7:C1048576$/.test(selTo) && /Count: \d+/.test(st), [selTo, st.replace(/\n/g, " | ")]);
  await page.keyboard.press("Control+Home"); await frames(page); await page.waitForTimeout(200);
  check("j: and back: the sheet is as it was", JSON.stringify(await size(page)) === JSON.stringify(s0), await size(page));
  check("j: no errors", errors.length === 0, errors.slice(0, 3));
  await page.close();
}

// ---------------- (k) charts and shapes are drawn; OnError "drawing"; the host's CSS leaves them alone ----------------
if (want("k")) {
  for (const [key, file] of [["ecommerce", `${T}/real-world/dashboard-ecommerce-sales.xlsx`], ["sales", `${T}/real-world/dashboard-sales-analysis.xlsx`], ["zoo", `${T}/real-world/chart-zoo.xlsx`]]) {
    const page = await newPage(context); const errors = errorsOf(page);
    await open(page, "sheet-editor", file); await ready(page); await page.waitForTimeout(500);
    const sheets = await g(page, () => window.__grids.at(-1).GetSheetInfos().map((s) => s.name ?? s.Name));
    let drawnTotal = 0, labelTotal = 0, objectTotal = 0;
    for (let i = 0; i < Math.min(sheets.length, 8); i++) {
      await g(page, (i) => window.__grids.at(-1).SwitchSheet(i), i); await frames(page); await page.waitForTimeout(350);
      const info = await g(page, () => { const objs = [...document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-object')]; return { objects: objs.length, drawn: objs.filter((o) => o.classList.contains("spreadsheet-object-drawn")).length, labels: document.querySelectorAll('[data-testid="sheet-grid"] .spreadsheet-object-label').length, svgs: objs.filter((o) => o.querySelector("svg")).length, pointer: [...new Set(objs.map((o) => getComputedStyle(o).pointerEvents))], bg: [...new Set(objs.filter((o) => o.classList.contains("spreadsheet-object-drawn")).map((o) => getComputedStyle(o).backgroundColor))], border: [...new Set(objs.filter((o) => o.classList.contains("spreadsheet-object-drawn")).map((o) => getComputedStyle(o).borderTopWidth))], drawings: window.__grids.at(-1).GetDrawings().map((d) => d.kind).join(",") }; });
      if (info.objects) console.log(`  k ${key} sheet ${i} "${sheets[i]}":`, JSON.stringify(info));
      drawnTotal += info.drawn; labelTotal += info.labels; objectTotal += info.objects;
      if (info.objects) await page.screenshot({ path: `${E}/n16-sheet-k-${key}-${i}.png` });
    }
    check(`k ${key}: objects are drawn as SVG`, drawnTotal > 0, { objectTotal, drawnTotal, labelTotal });
    check(`k ${key}: no error reported while drawing`, errors.length === 0, errors.slice(0, 4));
    check(`k ${key}: drawing did not modify the workbook`, /No changes/.test(await sel(page, "office-editor-status").innerText()));
    await page.close();
  }
  // dark theme: the grid is a white sheet in either theme, what is drawn on it is the file's
  const page = await newPage(context);
  await page.emulateMedia({ colorScheme: "dark" });
  await page.addInitScript(() => { try { localStorage.setItem("theme", "dark"); localStorage.setItem("rb-theme", "dark"); } catch {} });
  await open(page, "sheet-editor", `${T}/real-world/dashboard-ecommerce-sales.xlsx`); await ready(page); await page.waitForTimeout(600);
  await g(page, () => { document.documentElement.classList.add("dark"); document.documentElement.dataset.theme = "dark"; }); await frames(page);
  const dark = await g(page, () => { const o = document.querySelector('[data-testid="sheet-grid"] .spreadsheet-object-drawn'); const grid = document.querySelector('[data-testid="sheet-grid"]'); return { isDark: document.documentElement.classList.contains("dark"), gridBg: getComputedStyle(grid).backgroundColor, svgColor: o ? getComputedStyle(o.querySelector("svg")).color : null, objBg: o ? getComputedStyle(o).backgroundColor : null }; });
  console.log("  k dark:", JSON.stringify(dark));
  await page.screenshot({ path: `${E}/n16-sheet-k-dark.png` });
  check("k: in the dark theme the grid stays a white sheet", /255, 255, 255/.test(dark.gridBg), dark);
  await page.close();
}

// ---------------- (l) a formula typed in a General cell gets a number format in the same undo step ----------------
if (want("l")) {
  const page = await newPage(context);
  await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`); await ready(page);
  // an empty area far right
  const fmt = (r, c) => g(page, ([r, c]) => { const gr = window.__grids.at(-1); gr.SelectCell(r, c); return gr.GetSelectionFormat().numberFormat ?? null; }, [r, c]);
  const R = 2, C = 40;
  await g(page, ([R, C]) => { const gr = window.__grids.at(-1); gr.ScrollToCell(R, C); gr.SelectCell(R, C); }, [R, C]); await frames(page);
  await click(page, R, C);
  await page.keyboard.type("1/15/2026"); await page.keyboard.press("Enter");
  const colLetter = await g(page, ([R, C]) => { window.__grids.at(-1).SelectCell(R, C); return window.__grids.at(-1).GetSelectionRef(); }, [R, C]);
  console.log("  l: first cell is", colLetter);
  // type the formula against whatever that column is
  const col = colLetter.replace(/\d+/g, "");
  await click(page, R + 1, C); await page.keyboard.type(`=${col}${R + 1}+30`); await page.keyboard.press("Enter"); await frames(page);
  const r1 = await revision(page);
  const f1 = await fmt(R + 1, C); const c1 = await cell(page, R + 1, C);
  check("l: a date plus a number, typed in a General cell, is shown as a date", !!f1 && /[dmy]/i.test(f1) && /2026|26/.test(c1.text ?? ""), [f1, c1.text]);
  const box = await sel(page, "sheet-number-format").inputValue().catch(() => null);
  check("l: the toolbar's number format box shows a date format", box !== null && !/general/i.test(box), box);
  await undo(page); await frames(page);
  const c2 = await cell(page, R + 1, C); const f2 = await fmt(R + 1, C);
  check("l: one undo takes the formula and the format away together", (c2.text ?? "") === "" && (f2 === null || /general/i.test(f2)), [c2.text, f2]);
  await redo(page); await frames(page);
  check("l: redo brings both back", (await fmt(R + 1, C)) === f1 && (await revision(page)) === r1);
  // TODAY()
  await click(page, R + 3, C); await page.keyboard.type("=TODAY()"); await page.keyboard.press("Enter"); await frames(page);
  const ft = await fmt(R + 3, C); const ct = await cell(page, R + 3, C);
  check("l: =TODAY() is shown as a date", !!ft && /[dmy]/i.test(ft) && /\d+[\/.-]\d+/.test(ct.text ?? ""), [ft, ct.text]);
  // through the formula bar
  await click(page, R + 5, C); await sel(page, "sheet-formula").click(); await page.keyboard.type(`=${col}${R + 1}+7`); await page.keyboard.press("Enter"); await frames(page);
  const fb = await fmt(R + 5, C);
  check("l: entered through the formula bar: the same", !!fb && /[dmy]/i.test(fb), fb);
  // a cell that has a format of its own keeps it
  await click(page, R + 7, C); await sel(page, "sheet-number-format").selectOption("number").catch(() => {}); await frames(page);
  const own = await fmt(R + 7, C);
  await click(page, R + 7, C); await page.keyboard.type(`=${col}${R + 1}+1`); await page.keyboard.press("Enter"); await frames(page);
  check("l: a cell with a number format of its own keeps it", (await fmt(R + 7, C)) === own && own !== null, [own, await fmt(R + 7, C), (await cell(page, R + 7, C)).text]);
  await page.screenshot({ path: `${E}/n16-sheet-l.png` });
  await page.close();
}

// ---------------- (m) a drag held at the edge of the grid scrolls it ----------------
if (want("m")) {
  const page = await newPage(context);
  await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`); await ready(page);
  const p = await click(page, 6, 2);
  const grid = await sel(page, "sheet-grid").boundingBox();
  const top = () => g(page, () => { const gr = window.__grids.at(-1); const all = [...document.querySelectorAll('[data-testid="sheet-grid"] *')].filter((e) => e.scrollHeight > e.clientHeight + 50 && getComputedStyle(e).overflowY !== "visible"); return { top: Math.max(0, ...all.map((e) => e.scrollTop)), left: Math.max(0, ...all.map((e) => e.scrollLeft)), sel: gr.GetSelectionRef() }; });
  const t0 = await top();
  await page.mouse.move(p.x, p.y); await page.mouse.down();
  await page.mouse.move(p.x + 30, grid.y + grid.height - 2, { steps: 6 });
  await page.waitForTimeout(700);
  const t1 = await top();
  await page.mouse.move(p.x + 30, grid.y + grid.height + 40, { steps: 2 });
  await page.waitForTimeout(500);
  const t2 = await top();
  await page.mouse.up(); await frames(page);
  const t3 = await top();
  await page.waitForTimeout(400);
  const t4 = await top();
  console.log("  m:", JSON.stringify({ t0, t1, t2, t3, t4 }));
  check("m: held at the bottom edge while dragging: the sheet scrolls and the selection grows", t1.top > t0.top && t1.sel !== t0.sel, [t0, t1]);
  check("m: further out it goes on (faster)", t2.top > t1.top, [t1, t2]);
  check("m: let go: the scrolling stops", t4.top === t3.top, [t3, t4]);
  check("m: the selection runs from the cell the drag began in", /^C7:/.test(t3.sel), t3.sel);
  await page.screenshot({ path: `${E}/n16-sheet-m.png` });
  // the chrome follows: the name box shows the range's start, the status bar counts
  const st = await sel(page, "sheet-status").innerText();
  check("m: the status bar counts the dragged selection", /Count: \d+/.test(st), st);
  // to the right edge
  await page.keyboard.press("Control+Home"); await frames(page);
  const q = await click(page, 4, 1);
  await page.mouse.move(q.x, q.y); await page.mouse.down();
  await page.mouse.move(grid.x + grid.width - 2, q.y + 5, { steps: 6 }); await page.waitForTimeout(700);
  const r1 = await top(); await page.mouse.up();
  check("m: held at the right edge: it scrolls to the right", r1.left > 0, r1);
  // the host's own overlays do not swallow the drag: a mouseup outside the grid ends it
  await page.keyboard.press("Control+Home"); await frames(page);
  const z = await click(page, 6, 2);
  await page.mouse.move(z.x, z.y); await page.mouse.down(); await page.mouse.move(z.x, grid.y + grid.height + 30, { steps: 4 }); await page.waitForTimeout(300); await page.mouse.up();
  const a = await top(); await page.waitForTimeout(400); const b = await top();
  check("m: let go outside the grid (over the tabs / status bar): the scrolling stops there too", a.top === b.top, [a, b]);
  await page.close();
}
await browser.close();
summary();
