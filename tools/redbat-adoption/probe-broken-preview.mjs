import { launch, newPage, open, S, SHOTS } from "./lib.mjs";
const { browser, context } = await launch();
const page = await newPage(context);
await open(page, "chat-viewer", `${S}/broken-late.docx`, { params: { mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } });
await page.waitForFunction(() => document.querySelector(".rb-file-viewer-canvas") || document.querySelector(".rb-file-viewer-error-copy"), null, { timeout: 60000 });
await page.waitForTimeout(4000);
console.log(JSON.stringify(await page.evaluate(() => ({ notice: document.querySelector('[data-testid="word-preview-truncated"]')?.textContent.trim(), canvases: document.querySelectorAll(".rb-file-viewer-canvas").length, sub: document.querySelector(".rb-file-viewer-header-copy .rb-app-subtitle")?.textContent }))));
await page.screenshot({ path: `${SHOTS}/broken-preview.png` });
await browser.close();
