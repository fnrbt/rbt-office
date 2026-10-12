import { launch, newPage, open, S, waitUploads } from "./lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "word-editor", `${S}/broken-late.docx`);
await page.waitForFunction(() => document.querySelector('[data-testid="word-toolbar"]') || document.querySelector('[data-testid="office-editor-load-error"]'), null, { timeout: 60000 });
console.log("load error:", await page.locator('[data-testid="office-editor-load-error"]').innerText().catch(() => null));
if (await page.locator('[data-testid="word-toolbar"]').count()) {
  await page.waitForTimeout(4000);
  console.log(JSON.stringify(await page.evaluate(() => { const w = window.__word, e = window.__editors[0]; let text = null, err = null; try { text = w.getText(e); } catch (x) { err = String(x); } return { load: w.isLoadComplete(e), notice: document.querySelector(".rb-office-notice")?.textContent, status: document.querySelector('[data-testid="office-editor-status"]').innerText, chars: text?.length, has6999: text?.includes("Filler paragraph 6999."), has7001: text?.includes("Filler paragraph 7001."), err }; })));
  await page.keyboard.type("X");
  await page.keyboard.press("Control+s");
  await page.waitForTimeout(3000);
  console.log("uploads:", page.uploads.length, "error shown:", await page.locator('[data-testid="office-editor-error"]').innerText().catch(() => null), "| status:", await page.locator('[data-testid="office-editor-status"]').innerText());
}
const whole = await page.evaluate(async (url) => { const of = (await window.__formats()).default(); const b = await (await fetch(url)).arrayBuffer(); try { const p = await of.processors.forEditorPackage(b); return { blocks: p.Paragraphs.length, setup: p.PageSetup != null }; } catch (e) { return { err: String(e) }; } }, "/api/fixture/" + encodeURIComponent(`${S}/broken-late.docx`));
console.log("forEditorPackage of the same file:", JSON.stringify(whole));
await browser.close();
