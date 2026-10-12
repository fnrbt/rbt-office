import { launch, newPage, open, T, SHOTS, check, summary, listenerProbe } from "./lib.mjs";
const { browser, context } = await launch();
await context.addInitScript(listenerProbe);
const ev = (page, fn, arg) => page.evaluate(fn, arg);
{
  const page = await newPage(context);
  await open(page, "chat-viewer", `${T}/test-documents/sample-formatted.docx`, { params: { mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } });
  await page.waitForSelector(".rb-file-viewer-canvas", { timeout: 120000 });
  await page.waitForTimeout(600);
  const info = await ev(page, () => ({ sub: document.querySelector(".rb-file-viewer-header-copy .rb-app-subtitle").textContent.trim(), canvases: document.querySelectorAll(".rb-file-viewer-canvas").length, painted: [...document.querySelectorAll(".rb-file-viewer-canvas")].filter((c) => c.width >= 800).length, stack: window.RedBatWordTextStack?.loadedFamilyNames?.() }));
  check("chat attachment: Word document shown with the shared preview (shaped pages)", info.sub === "10 pages" && info.canvases === 10 && info.painted >= 1 && info.painted < 10 && info.stack?.includes("Calibri"), info);
  await page.screenshot({ path: `${SHOTS}/chat-word.png` });
  await page.keyboard.press("+");
  await page.waitForTimeout(200);
  check("zoom", (await page.locator(".rb-file-viewer-zoom-label").innerText()) === "120%");
  await page.keyboard.press("Escape");
  await page.waitForFunction(() => window.__closed === 1);
  check("Escape closes", true);
  await page.close();
}
{
  const page = await newPage(context);
  await open(page, "chat-viewer", `${T}/real-world/milk-street-office.xlsx`, { params: { mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } });
  await page.waitForFunction(() => document.querySelector('[data-testid="spreadsheet-preview"]')?.style.visibility === "visible", null, { timeout: 120000 });
  await page.waitForTimeout(400);
  const info = await ev(page, () => ({ ro: window.__grids[0].IsReadOnly(), tabs: [...document.querySelectorAll('[data-testid="spreadsheet-preview-tabs"] button')].map((b) => b.textContent), fills: [...document.querySelectorAll(".spreadsheet-cell")].filter((c) => getComputedStyle(c).backgroundColor === "rgb(255, 255, 153)").length, h: document.querySelector(".spreadsheet-container").clientHeight }));
  check("chat attachment: workbook shown with the shared preview (engine grid, read-only, tabs, fills)", info.ro && info.tabs.join() === "Annual,Monthly" && info.fills > 5 && info.h > 300, info);
  await page.screenshot({ path: `${SHOTS}/chat-sheet.png` });
  const l0 = await ev(page, () => window.__listeners());
  await page.keyboard.press("Escape");
  await page.waitForFunction(() => window.__closed === 1);
  await page.waitForTimeout(200);
  check("grid removed on close", (await page.locator(".spreadsheet-container").count()) === 0);
  await page.close();
}
{
  const page = await newPage(context);
  await open(page, "chat-viewer", `/tmp/redbat-adoption/sample.csv`, { params: { mime: "text/csv" } });
  await page.waitForFunction(() => document.querySelector('[data-testid="spreadsheet-preview"]')?.style.visibility === "visible" || document.querySelector("pre, textarea, .rb-text-viewer"), null, { timeout: 60000 });
  await page.waitForTimeout(300);
  console.log("  csv attachment viewer:", await ev(page, () => (document.querySelector('[data-testid="spreadsheet-preview"]') ? "spreadsheet preview" : "text viewer")));
  await page.screenshot({ path: `${SHOTS}/chat-csv.png` });
  await page.close();
}
{ // unreadable docx: error inside the modal with its chrome
  const page = await newPage(context);
  await open(page, "chat-viewer", `/tmp/redbat-adoption/sample.csv`, { name: "broken.docx", params: { mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } });
  await page.waitForSelector(".rb-file-viewer-error-copy", { timeout: 60000 });
  check("unreadable document: error state with Download, inside the viewer", (await page.locator(".rb-state-card").innerText()).includes("Not a readable DOCX") && (await page.getByRole("button", { name: "Download" }).count()) >= 1, await page.locator(".rb-state-card").innerText());
  await page.close();
}
summary();
await browser.close();
