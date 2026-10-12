import { launch, newPage, open, T, S, SHOTS, check, summary, waitUploads } from "./lib.mjs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
{ // spreadsheet editor
  const page = await newPage(context);
  await open(page, "sheet-editor", `${T}/real-world/saopaulo-condos.xlsx`);
  await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 180000 });
  await page.waitForTimeout(500);
  await page.screenshot({ path: `${SHOTS}/final-sheet-editor.png` });
  await ev(page, () => window.__grids[0].SelectRange(3, 1, 6, 3));
  await page.click('[data-testid="sheet-border"]');
  await page.click('[data-testid="sheet-border-line-dashed"]');
  await page.waitForTimeout(150);
  await page.screenshot({ path: `${SHOTS}/final-sheet-editor-border-menu.png` });
  const tb = await ev(page, () => { const t = document.querySelector('[data-testid="sheet-toolbar"]'); const m = document.querySelector('[data-testid="sheet-border-menu"]').getBoundingClientRect(); const shell = document.querySelector(".rb-office-editor-shell").getBoundingClientRect(); return { h: t.offsetHeight, clipped: [...t.children].filter((c) => c.getBoundingClientRect().right > t.getBoundingClientRect().right + 0.5).length, menuInside: m.right <= shell.right && m.left >= shell.left }; });
  check("sheet toolbar on one row at 1366, border menu inside the editor", tb.h < 60 && tb.clipped === 0 && tb.menuInside, tb);
  await page.close();
}
{ // CSV: clean on open, edit, save
  const page = await newPage(context);
  await open(page, "sheet-editor", `${S}/sample.csv`);
  await page.waitForSelector('[data-testid="sheet-toolbar"]', { timeout: 180000 });
  await page.waitForTimeout(300);
  const st = () => page.locator('[data-testid="office-editor-status"]').innerText();
  check("CSV opens clean (installing its workbook is not an edit)", (await st()).includes("No changes") && !(await ev(page, () => window.__grids[0].IsModified())), await st());
  await ev(page, () => window.__grids[0].EnterText(1, 1, "42"));
  check("CSV dirty after an edit", (await st()).includes("Unsaved"));
  await page.click('[data-testid="sheet-undo"]');
  check("CSV clean after undo", (await st()).includes("No changes"), await st());
  await page.click('[data-testid="sheet-redo"]');
  await page.click('[data-testid="office-editor-save"]');
  const bytes = await waitUploads(page, 1);
  await page.waitForTimeout(200);
  check("CSV saved as text with the edit", bytes.toString("utf8").split(/\r?\n/)[1] === "widget,42,1.5" && (await st()).includes("Saved"), bytes.toString("utf8").slice(0, 80));
  await page.screenshot({ path: `${SHOTS}/final-csv-editor.png` });
  await page.close();
}
{ // word editor
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/test-documents/sample-formatted.docx`);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 180000 });
  await page.waitForTimeout(500);
  await page.screenshot({ path: `${SHOTS}/final-word-editor.png` });
  await page.click('[data-testid="word-spacing"]');
  await page.waitForTimeout(100);
  await page.screenshot({ path: `${SHOTS}/final-word-editor-spacing.png` });
  await page.keyboard.press("Escape");
  // a new table at the end, caret in it: table tools
  await ev(page, () => window.__word.getInputElement(window.__editors[0]).focus());
  await page.keyboard.press("Control+End");
  await page.keyboard.press("Enter");
  await page.click('[data-testid="word-table"]');
  await page.click('[data-testid="word-table-insert"]');
  await page.keyboard.type("first cell");
  await page.keyboard.press("Tab");
  await page.keyboard.type("second");
  await page.waitForTimeout(400);
  const tb = await ev(page, () => { const t = document.querySelector('[data-testid="word-toolbar"]'); return { h: t.offsetHeight, tools: !!document.querySelector('[data-testid="word-table-tools"]'), clipped: [...t.querySelectorAll("button, select")].filter((c) => c.getBoundingClientRect().right > t.getBoundingClientRect().right + 0.5 || c.getBoundingClientRect().bottom > t.getBoundingClientRect().bottom + 0.5).length }; });
  check("word toolbar with the table tools: nothing clipped at 1366", tb.tools && tb.clipped === 0, tb);
  await page.screenshot({ path: `${SHOTS}/final-word-editor-table.png` });
  await page.close();
}
{ // previews
  const page = await newPage(context);
  await open(page, "chat-viewer", `${T}/test-documents-wild/docx4j-getting-started.docx`, { params: { mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } });
  await page.waitForSelector(".rb-file-viewer-canvas", { timeout: 120000 });
  await page.waitForTimeout(800);
  await page.screenshot({ path: `${SHOTS}/final-word-preview.png` });
  await page.close();
  const page2 = await newPage(context);
  await open(page2, "chat-viewer", `${T}/real-world/saopaulo-condos.xlsx`, { params: { mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } });
  await page2.waitForFunction(() => document.querySelector('[data-testid="spreadsheet-preview"]')?.style.visibility === "visible", null, { timeout: 120000 });
  await page2.waitForTimeout(500);
  await page2.screenshot({ path: `${SHOTS}/final-sheet-preview.png` });
  await page2.close();
  for (const f of ["real-world/chart-zoo.xlsx", "test-documents-excel/synthetic-condformat.xlsx", "test-documents-excel/synthetic-drawing-extras.xlsx", "real-world/milk-street-office.xlsx"]) {
    const p = await newPage(context);
    await open(p, "chat-viewer", `${T}/${f}`, { params: { mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } });
    await p.waitForFunction(() => document.querySelector('[data-testid="spreadsheet-preview"]')?.style.visibility === "visible", null, { timeout: 120000 });
    await p.waitForTimeout(500);
    await p.screenshot({ path: `${SHOTS}/final-preview-${f.split("/").pop()}.png` });
    await p.close();
  }
}
summary();
await browser.close();
