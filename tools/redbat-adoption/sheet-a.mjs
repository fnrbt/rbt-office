// Borders, row/column operations, dirty state, delayed and failed uploads.
import { launch, newPage, open, T, SHOTS, check, summary, saved, unzip, waitUploads } from "./lib.mjs";
import fs from "node:fs";
const { browser, context } = await launch();
const g = (page, fn, arg) => page.evaluate(fn, arg);
const ready = async (page) => { await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 180000 }); await page.waitForTimeout(300); };
const status = (page) => page.locator('[data-testid="office-editor-status"]').innerText();
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const undoDepth = (page) => g(page, () => { let n = 0; for (let l = window.__grids.at(-1).GetModel().History.Past; l && l.tail != null; l = l.tail) n++; return n; });
// The four edges of a cell as the component reports them (GetCellEdges): the style name, "grid" or "-".
const edges = (page, r, c) => g(page, ([r, c]) => { const e = window.__grids.at(-1).GetCellEdges(r, c); const d = (x) => (x == null ? "-" : x === "grid" ? "grid" : `${x.Style}/${x.Width}/${x.Color}`); return [d(e.Top), d(e.Right), d(e.Bottom), d(e.Left)]; }, [r, c]);
const applyBorder = async (page, range, line, side) => {
  await g(page, (r) => window.__grids.at(-1).SelectRange(...r), range);
  await sel(page, "sheet-border").click();
  if (line) await sel(page, `sheet-border-line-${line}`).click();
  await sel(page, `sheet-border-${side}`).click();
};
const widths = { thin: 1, medium: 2, thick: 3, dashed: 1, dotted: 1, double: 3 };
// Blocks of 3 x 3 cells in an empty part of the sheet, one per case.
// (the sheet has 26 columns)
const R0 = 60, C0 = 17;
const block = (i) => { const top = R0 + Math.floor(i / 2) * 4, left = C0 + (i % 2) * 4; return [top, left, top + 2, left + 2]; };
const cases = [
  ["thin", "all"], ["medium", "all"], ["thick", "all"], ["dashed", "all"], ["dotted", "all"], ["double", "all"],
  ["medium", "outside"], ["dashed", "inside"], ["thick", "top"], ["double", "bottom"], ["dotted", "left"], ["medium", "right"],
];
const expectSides = (side, r, c, [top, left, bottom, right]) => {
  // which of [top, right, bottom, left] of cell (r, c) the side sets
  switch (side) {
    case "all": return [true, true, true, true];
    case "outside": return [r === top, c === right, r === bottom, c === left];
    case "inside": return [r > top, c < right, r < bottom, c > left];
    case "top": return [true, false, false, false];
    case "bottom": return [false, false, true, false];
    case "left": return [false, false, false, true];
    case "right": return [false, true, false, false];
  }
};
// Whether the edge on side k (0 top, 1 right, 2 bottom, 3 left) of cell (r, c) is set: by the cell
// itself or, a shared edge being one line, by its neighbour in the block.
const edgeSet = (side, r, c, b, k) => {
  const inBlock = (r, c) => r >= b[0] && r <= b[2] && c >= b[1] && c <= b[3];
  const own = expectSides(side, r, c, b)[k];
  const [nr, nc, nk] = k === 0 ? [r - 1, c, 2] : k === 1 ? [r, c + 1, 3] : k === 2 ? [r + 1, c, 0] : [r, c - 1, 1];
  return own || (inBlock(nr, nc) && expectSides(side, nr, nc, b)[nk]);
};
// Pixels of the grid: for each edge, how it is drawn across (dark pixels in a cross-section) and along.
const measureEdges = async (page) => {
  const png = await sel(page, "sheet-grid").screenshot();
  const rects = [];
  for (let i = 0; i < cases.length; i++) { const b = block(i); for (let r = b[0]; r <= b[2]; r++) for (let c = b[1]; c <= b[3]; c++) rects.push([i, r, c]); }
  return g(page, async ([b64, rects]) => {
    const img = new Image(); img.src = "data:image/png;base64," + b64; await img.decode();
    const cv = document.createElement("canvas"); cv.width = img.width; cv.height = img.height; const cx = cv.getContext("2d"); cx.drawImage(img, 0, 0);
    const data = cx.getImageData(0, 0, cv.width, cv.height).data;
    const dark = (x, y) => { if (x < 0 || y < 0 || x >= cv.width || y >= cv.height) return false; const o = (y * cv.width + x) * 4; return data[o] + data[o + 1] + data[o + 2] < 330; };
    const origin = document.querySelector('[data-testid="sheet-grid"]').getBoundingClientRect();
    const out = {};
    for (const [i, r, c] of rects) {
      const el = document.querySelector(`.spreadsheet-cell[data-row="${r}"][data-col="${c}"]`); if (!el) { out[`${r},${c}`] = null; continue; }
      const q = el.getBoundingClientRect(); const L = Math.round(q.left - origin.left), T = Math.round(q.top - origin.top), R = Math.round(q.right - origin.left), B = Math.round(q.bottom - origin.top);
      const side = (k) => {
        const horizontal = k === 0 || k === 2; const at = k === 0 ? T : k === 2 ? B : k === 1 ? R : L;
        const from = horizontal ? L : T, to = horizontal ? R : B; const a = Math.round(from + (to - from) * 0.2), z = Math.round(from + (to - from) * 0.8);
        let covered = 0, most = 0, pattern = "", transitions = 0, last = null;
        for (let p = a; p < z; p++) {
          let n = 0, pat = "";
          for (let d = -4; d <= 4; d++) { const on = horizontal ? dark(p, at + d) : dark(at + d, p); n += on ? 1 : 0; pat += on ? "1" : "0"; }
          if (n > 0) covered++;
          if (n > most) { most = n; pattern = pat.replace(/^0+|0+$/g, ""); }
          if (last !== null && (n > 0) !== last) transitions++;
          last = n > 0;
        }
        return { coverage: Math.round((covered / (z - a)) * 100) / 100, most, pattern, transitions };
      };
      out[`${r},${c}`] = [side(0), side(1), side(2), side(3)];
    }
    return out;
  }, [png.toString("base64"), rects]);
};
const verifyBorders = async (page, label) => {
  // tall enough for all six rows of blocks
  await page.setViewportSize({ width: 1366, height: 1000 });
  await page.waitForTimeout(400);
  const px = await measureEdges(page);
  await sel(page, "sheet-grid").screenshot({ path: `${SHOTS}/borders-${label}.png` });
  const seen = {};
  for (let i = 0; i < cases.length; i++) {
    const [line, side] = cases[i]; const b = block(i); const bad = [];
    for (let r = b[0]; r <= b[2]; r++) for (let c = b[1]; c <= b[3]; c++) {
      const got = await edges(page, r, c); const m = px[`${r},${c}`];
      ["top", "right", "bottom", "left"].forEach((d, k) => {
        const want = edgeSet(side, r, c, b, k);
        // the component's own account of the edge
        if (want && got[k] !== `${line}/${widths[line]}/#000000`) bad.push(`edge ${r},${c} ${d}=${got[k]}`);
        if (!want && got[k] !== "-") bad.push(`edge ${r},${c} ${d} set: ${got[k]}`);
        // and what is on screen
        if (!m) { bad.push(`no element ${r},${c}`); return; }
        const e = m[k];
        if (!want) { if (e.coverage > 0) bad.push(`pixels ${r},${c} ${d} drawn ${JSON.stringify(e)}`); return; }
        (seen[line] ??= []).push(e);
        const solid = e.coverage >= 0.98;
        const ok = line === "thin" ? solid && e.pattern === "1" : line === "medium" ? solid && e.pattern === "11" : line === "thick" ? solid && e.pattern === "111" : line === "double" ? solid && e.pattern === "101"
          : /* dashed, dotted */ e.pattern === "1" && e.coverage > 0.2 && e.coverage < 0.95 && e.transitions >= 4;
        if (!ok) bad.push(`pixels ${r},${c} ${d} ${JSON.stringify(e)}`);
      });
    }
    check(`${label}: ${line} ${side} border - GetCellEdges and the pixels of all 36 edges of the 3x3 block`, bad.length === 0, bad.slice(0, 4));
  }
  const avg = (line) => seen[line].reduce((n, e) => n + e.transitions, 0) / seen[line].length;
  check(`${label}: dotted lines have more, shorter marks than dashed ones`, avg("dotted") > avg("dashed") * 1.5, { dashed: avg("dashed"), dotted: avg("dotted") });
  // one line per shared edge: the inner lines of "medium all" are as wide as its outline
  const inner = px[`${block(1)[0] + 1},${block(1)[1] + 1}`], outer = px[`${block(1)[0]},${block(1)[1]}`];
  check(`${label}: a shared edge is one line - inner and outer lines of "medium all" are both 2 px`, inner.every((e) => e.pattern === "11") && outer[0].pattern === "11" && outer[3].pattern === "11", { inner: inner.map((e) => e.pattern), outer: outer.map((e) => e.pattern) });
  await page.setViewportSize({ width: 1366, height: 768 });
  await page.waitForTimeout(300);
};

let borderShot;
{
  const page = await newPage(context);
  await open(page, "sheet-editor", `${T}/real-world/milk-street-office.xlsx`);
  await ready(page);
  check("clean after load", (await status(page)).includes("No changes") && !(await g(page, () => window.__grids[0].IsModified())), await status(page));

  // ---- the border menu
  await sel(page, "sheet-border").click();
  const menu = await g(page, () => ({ sides: [...document.querySelectorAll('[data-testid^="sheet-border-"]')].map((e) => e.getAttribute("data-testid").slice(13)).filter((n) => !n.startsWith("line") && n !== "menu"), lines: [...document.querySelectorAll('[data-testid^="sheet-border-line-"]')].map((e) => e.getAttribute("data-testid").slice(18)) }));
  check("border menu offers all / outside / inside / each side / none and six line styles", menu.sides.join() === "all,outside,inside,top,bottom,left,right,none" && menu.lines.join() === "thin,medium,thick,dashed,dotted,double", menu);
  await page.screenshot({ path: `${SHOTS}/sheet-border-menu.png` });
  await page.keyboard.press("Escape");
  for (let i = 0; i < cases.length; i++) {
    const d0 = await undoDepth(page);
    await applyBorder(page, block(i), cases[i][0], cases[i][1]);
    if (i === 0 || i === 6) check(`${cases[i][1]} border on a 3x3 block is one undo step`, (await undoDepth(page)) === d0 + 1);
  }
  await g(page, ([r, c]) => { const gr = window.__grids[0]; gr.SelectCell(r - 2, c - 1); gr.ScrollToCell(r + 24, c + 7); gr.ScrollToCell(r - 1, c - 3); }, [R0, C0]);
  await page.waitForTimeout(300);
  await verifyBorders(page, "editor");
  await sel(page, "sheet-grid").screenshot({ path: `${SHOTS}/sheet-borders.png` });
  // clear: "No border" on the first block, and line "none" via the side buttons is not offered; a second click of "none" changes nothing
  await g(page, (r) => window.__grids[0].SelectRange(...r), block(0));
  await sel(page, "sheet-border").click();
  await sel(page, "sheet-border-none").click();
  const cleared = await edges(page, block(0)[0] + 1, block(0)[1] + 1);
  check("No border clears all four edges", cleared.join() === "-,-,-,-", cleared);
  await sel(page, "sheet-undo").click();
  check("clearing is one undo step", (await edges(page, block(0)[0] + 1, block(0)[1] + 1)).join() === "thin/1/#000000,thin/1/#000000,thin/1/#000000,thin/1/#000000");

  // ---- rows and columns: the selection's count, one undo step, formulas follow
  await g(page, () => { const gr = window.__grids[0]; for (let i = 0; i < 8; i++) gr.EnterText(100 + i, 20, String(i + 1)); gr.EnterText(110, 20, "=SUM(U101:U108)"); gr.EnterText(100, 23, "=U101+U108"); });
  const read = (r, c) => g(page, ([r, c]) => { const gr = window.__grids[0]; gr.SelectCell(r, c); return { f: gr.GetCurrentCellFormula(), v: gr.GetSelectionText().trim() }; }, [r, c]);
  check("formula before", (await read(110, 20)).v === "36");
  await g(page, () => window.__grids[0].SelectRange(102, 20, 104, 22)); // rows 103..105
  await page.click('[data-testid="sheet-menu-button-insert"]');
  check("the menu says how many", (await sel(page, "sheet-insert-row").innerText()) === "Insert 3 rows above" && (await sel(page, "sheet-delete-column").innerText()) === "Delete 3 columns", await sel(page, "sheet-insert-row").innerText());
  await page.keyboard.press("Escape");
  let d = await undoDepth(page);
  await page.click('[data-testid="sheet-menu-button-insert"]'); await sel(page, "sheet-insert-row").click();
  check("insert rows: 3 rows for a 3-row selection, one undo step", (await undoDepth(page)) === d + 1, { depth: (await undoDepth(page)) - d });
  let f = await read(113, 20);
  check("insert rows: the formula moved down and its range grew", f.f === "=SUM(U101:U111)" && f.v === "36", f);
  check("insert rows: cells below moved by 3", (await read(105, 20)).v === "3" && (await read(102, 20)).v === "", [await read(105, 20), await read(102, 20)]);
  await sel(page, "sheet-undo").click();
  check("insert rows: one undo restores", (await read(110, 20)).f === "=SUM(U101:U108)" && (await read(102, 20)).v === "3");
  await g(page, () => window.__grids[0].SelectRange(102, 20, 104, 22));
  d = await undoDepth(page);
  await page.click('[data-testid="sheet-menu-button-insert"]'); await sel(page, "sheet-delete-row").click();
  f = await read(107, 20);
  check("delete rows: 3 rows, one undo step, the formula follows", (await undoDepth(page)) === d + 1 && f.f === "=SUM(U101:U105)" && f.v === "24", { depth: (await undoDepth(page)) - d, f });
  check("delete rows: other formula keeps both ends", (await read(100, 23)).f === "=U101+U105", await read(100, 23));
  await sel(page, "sheet-undo").click();
  check("delete rows: one undo restores", (await read(110, 20)).f === "=SUM(U101:U108)" && (await read(110, 20)).v === "36" && (await read(103, 20)).v === "4");
  // columns
  await g(page, () => window.__grids[0].SelectRange(100, 21, 100, 22)); // 2 columns between U and X
  d = await undoDepth(page);
  await page.click('[data-testid="sheet-menu-button-insert"]'); await sel(page, "sheet-insert-column").click();
  f = await read(100, 25);
  check("insert columns: 2 columns, one undo step, the formula moved right and still reads U", (await undoDepth(page)) === d + 1 && f.f === "=U101+U108" && f.v === "9", { depth: (await undoDepth(page)) - d, f });
  await g(page, () => window.__grids[0].SelectRange(100, 18, 100, 19)); // 2 columns left of U
  d = await undoDepth(page);
  await page.click('[data-testid="sheet-menu-button-insert"]'); await sel(page, "sheet-delete-column").click();
  f = await read(100, 23);
  check("delete columns: 2 columns, one undo step, references shift left", (await undoDepth(page)) === d + 1 && f.f === "=S101+S108" && f.v === "9", { depth: (await undoDepth(page)) - d, f });
  await sel(page, "sheet-undo").click();
  await sel(page, "sheet-undo").click();
  check("column edits undone in two steps", (await read(100, 23)).f === "=U101+U108");
  // over the old 50-row cap
  await g(page, () => window.__grids[0].SelectRange(200, 0, 259, 0));
  d = await undoDepth(page);
  await page.click('[data-testid="sheet-menu-button-insert"]'); await sel(page, "sheet-delete-row").click();
  check("60 rows deleted in one step, no cap notice", (await undoDepth(page)) === d + 1 && (await sel(page, "sheet-notice").count()) === 0);
  await sel(page, "sheet-undo").click();

  // ---- dirty state follows the component's saved revision
  check("dirty after the edits", (await status(page)).includes("Unsaved") && (await g(page, () => window.__grids[0].IsModified())));
  await sel(page, "office-editor-save").click();
  const first = await waitUploads(page, 1);
  await page.waitForTimeout(200);
  check("clean after save", (await status(page)).includes("Saved") && !(await g(page, () => window.__grids[0].IsModified())) && (await sel(page, "office-editor-save").isDisabled()), await status(page));
  saved.set("borders", first);
  await g(page, () => { const gr = window.__grids[0]; gr.EnterText(120, 20, "after-save"); });
  check("dirty after a change to the saved content", (await status(page)).includes("Unsaved"));
  await sel(page, "sheet-undo").click();
  check("undo back to the saved content is clean", (await status(page)).includes("Saved") && (await sel(page, "office-editor-save").isDisabled()), await status(page));
  await sel(page, "sheet-undo").click();
  check("undo past the saved content is dirty", (await status(page)).includes("Unsaved"), await status(page));
  await sel(page, "sheet-redo").click();
  check("redo to the saved content is clean", (await status(page)).includes("Saved"), await status(page));

  // ---- typing during a slow upload
  let release; page.upload.hold = new Promise((r) => (release = r));
  await g(page, () => window.__grids[0].EnterText(121, 20, "before-upload"));
  await sel(page, "office-editor-save").click();
  while ((page.uploadsStarted ?? 0) < 2) await new Promise((r) => setTimeout(r, 30));
  check("status says Saving while the upload runs", (await status(page)).includes("Saving"), await status(page));
  await g(page, () => window.__grids[0].SelectCell(122, 20));
  await page.locator(".spreadsheet-content").first().focus();
  await page.keyboard.type("during-upload");
  await page.keyboard.press("Enter");
  release(); page.upload.hold = null;
  const second = await waitUploads(page, 2);
  await page.waitForTimeout(200);
  check("typing during the upload leaves the editor dirty", (await status(page)).includes("Unsaved") && (await g(page, () => window.__grids[0].IsModified())) && !(await sel(page, "office-editor-save").isDisabled()), await status(page));
  const strings = async (bytes) => { const z = await unzip(bytes); let s = z.has("xl/sharedStrings.xml") ? await z.get("xl/sharedStrings.xml").async("string") : ""; for (const [n, f] of z) if (/^xl\/worksheets\/sheet\d+\.xml$/.test(n)) s += await f.async("string"); return s; };
  const s2 = await strings(second);
  check("the upload that was under way holds what was there when it began, not what was typed since", s2.includes("before-upload") && !s2.includes("during-upload"));
  await sel(page, "sheet-undo").click();
  check("undoing what was typed during the upload is clean again", (await status(page)).includes("Saved"), await status(page));
  await sel(page, "sheet-redo").click();
  await sel(page, "office-editor-save").click();
  const third = await waitUploads(page, 3);
  await page.waitForTimeout(200);
  check("a second save uploads the newer content and is clean", (await strings(third)).includes("during-upload") && (await status(page)).includes("Saved"), await status(page));

  // ---- Save and close while typing goes on: stays open
  page.upload.hold = new Promise((r) => (release = r));
  await g(page, () => window.__grids[0].EnterText(123, 20, "x1"));
  await sel(page, "office-editor-save-close").click();
  while (page.uploadsStarted < 4) await new Promise((r) => setTimeout(r, 30));
  await g(page, () => window.__grids[0].EnterText(124, 20, "x2"));
  release(); page.upload.hold = null;
  await waitUploads(page, 4);
  await page.waitForTimeout(300);
  check("Save and close does not close over what was typed during the save", (await g(page, () => window.__closed)) === 0 && (await status(page)).includes("Unsaved"), await status(page));

  // ---- failed upload
  page.upload.fail = 500;
  await sel(page, "office-editor-save").click();
  await page.waitForSelector('[data-testid="office-editor-error"]', { timeout: 60000 });
  check("failed upload: error shown, still dirty, nothing marked saved", (await sel(page, "office-editor-error").innerText()).includes("Not saved") && (await status(page)).includes("Not saved") && (await g(page, () => window.__grids[0].IsModified())) && page.uploads.length === 4, await sel(page, "office-editor-error").innerText());
  page.upload.fail = 0;
  await sel(page, "office-editor-save").click();
  await waitUploads(page, 5);
  await page.waitForTimeout(200);
  check("saving again after the failure works", (await status(page)).includes("Saved") && (await sel(page, "office-editor-error").count()) === 0, await status(page));

  // ---- unsaved-changes guard
  await g(page, () => window.__grids[0].EnterText(125, 20, "unsaved"));
  await sel(page, "office-editor-close").click();
  check("closing with unsaved changes asks", (await sel(page, "office-editor-confirm").count()) === 1 && (await g(page, () => window.__closed)) === 0);
  await sel(page, "office-editor-keep").click();
  check("Keep editing stays", (await sel(page, "office-editor-confirm").count()) === 0 && (await sel(page, "office-editor").count()) === 1);
  await page.keyboard.press("Escape");
  check("Escape asks too", (await sel(page, "office-editor-confirm").count()) === 1);
  await sel(page, "office-editor-discard").click();
  await page.waitForFunction(() => window.__closed === 1);
  check("Discard closes", true);
  await page.close();
}
// ---- re-open what was saved after the border / row / column edits
{
  const page = await newPage(context);
  await open(page, "sheet-editor", "/api/saved/borders", { name: "milk-street-office.xlsx" });
  await ready(page);
  await g(page, ([r, c]) => { const gr = window.__grids[0]; gr.SelectCell(r - 2, c - 1); gr.ScrollToCell(r + 24, c + 7); gr.ScrollToCell(r - 1, c - 3); }, [R0, C0]);
  await page.waitForTimeout(300);
  await verifyBorders(page, "re-opened");
  await sel(page, "sheet-grid").screenshot({ path: `${SHOTS}/sheet-borders-reopened.png` });
  const z = await unzip(saved.get("borders"));
  const styles = await z.get("xl/styles.xml").async("string");
  const orig = await (await unzip(fs.readFileSync(`${T}/real-world/milk-street-office.xlsx`))).get("xl/styles.xml").async("string");
  const count = (s, style) => (s.match(new RegExp(`style="${style}"`, "g")) || []).length;
  check("saved styles.xml has each new line style", ["medium", "thick", "dashed", "dotted", "double"].every((st) => count(styles, st) > count(orig, st)), Object.fromEntries(["thin", "medium", "thick", "dashed", "dotted", "double"].map((st) => [st, `${count(orig, st)}->${count(styles, st)}`])));
  const r = await g(page, () => { const gr = window.__grids[0]; gr.SelectCell(110, 20); return { f: gr.GetCurrentCellFormula(), v: gr.GetSelectionText().trim() }; });
  check("re-open: the formula and its value", r.f === "=SUM(U101:U108)" && r.v === "36", r);
  check("re-open: clean", (await status(page)).includes("No changes"));
  await page.close();
}
summary();
await browser.close();
