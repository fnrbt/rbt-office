import { launch, newPage, open, T, SHOTS, check, summary, attested, BASE } from "./lib.mjs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
for (const [file, mime] of [[`${T}/test-documents/sample-formatted.docx`, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"], [`${T}/real-world/milk-street-office.xlsx`, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]]) {
  const page = await newPage(context);
  attested.path = file;
  const name = file.split("/").pop();
  await page.goto(`${BASE}/office-harness.html?${new URLSearchParams({ c: "files-viewer", name, mime, url: "x", hookWord: "1" })}`);
  await page.waitForFunction(() => window.__mounted === true, null, { timeout: 120000 });
  await page.waitForTimeout(500);
  const idError = await ev(page, () => window.__idError);
  if (idError) { console.log("  cannot build ids for the Files viewer:", idError); await page.screenshot({ path: `${SHOTS}/files-${name}-iderror.png` }); await page.close(); continue; }
  if (name.endsWith(".docx")) await page.waitForSelector(".rb-file-viewer-canvas", { timeout: 120000 });
  else await page.waitForFunction(() => document.querySelector('[data-testid="spreadsheet-preview"]')?.style.visibility === "visible", null, { timeout: 120000 });
  await page.waitForTimeout(700);
  await page.screenshot({ path: `${SHOTS}/files-${name}.png` });
  const sub = await page.locator(".rb-file-viewer-header-copy .rb-app-subtitle").innerText();
  if (name.endsWith(".docx")) check("Files viewer: Word preview with the page count in the header", /9 pages/.test(sub), sub);
  else check("Files viewer: spreadsheet preview fills the stage, with tabs", await ev(page, () => document.querySelector(".spreadsheet-container").clientHeight > 300 && document.querySelectorAll('[data-testid="spreadsheet-preview-tabs"] button').length === 2), await ev(page, () => document.querySelector(".spreadsheet-container").clientHeight));
  await page.click('[data-testid="file-viewer-edit-office"]');
  await page.waitForSelector(name.endsWith(".docx") ? '[data-testid="word-toolbar"]' : '[data-testid="sheet-toolbar"]', { timeout: 120000 });
  await page.waitForTimeout(500);
  check(`Files viewer: Edit opens the ${name.endsWith(".docx") ? "Word" : "spreadsheet"} editor over the preview`, (await page.locator('[data-testid="office-editor"]').count()) === 1);
  await page.screenshot({ path: `${SHOTS}/files-${name}-editor.png` });
  if (name.endsWith(".docx")) {
    const same = await ev(page, () => window.__word.getPageCount(window.__editors[0]));
    const previewRequestsBefore = page.requests.length;
    check("editor opened from the preview has the preview's page count", same === 9, same);
  }
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  check("Escape closes the editor and leaves the preview open", (await page.locator('[data-testid="office-editor"]').count()) === 0 && (await ev(page, () => window.__closed)) === 0 && (await page.locator(".rb-file-viewer-shell").count()) === 1);
  await page.close();
}
summary();
await browser.close();
