// First paint and full layout of the 522-page document, editor and preview.
import { launch, newPage, open, S } from "./lib.mjs";
const { browser } = await launch();
const file = `${S}/long-generated.docx`;
for (let run = 0; run < 2; run++) {
  { const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } }); const page = await newPage(ctx);
    const t = Date.now();
    await open(page, "word-editor", file);
    await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
    await page.waitForFunction(() => { const c = document.querySelector('[data-testid="word-canvas"]'); if (!c || !c.width) return false; const d = c.getContext("2d").getImageData(0, 0, c.width, Math.min(c.height, 500)).data; let n = 0; for (let i = 0; i < d.length; i += 16) if (d[i] < 90 && d[i + 1] < 90 && d[i + 2] < 90) n++; return n > 200; }, null, { timeout: 300000, polling: 30 });
    const first = Date.now() - t; const sub1 = await page.locator('[data-testid="office-editor-status"]').innerText();
    await page.waitForFunction(() => (window.__word.isLayoutComplete ? window.__word.isLayoutComplete(window.__editors[0]) : true), null, { timeout: 300000, polling: 30 });
    const done = Date.now() - t; await page.waitForTimeout(300);
    console.log(`editor run ${run + 1}: first screen painted and usable ${first} ms ("${sub1.replace(/\n/g, " | ")}"); whole layout ${done} ms ("${(await page.locator('[data-testid="office-editor-status"]').innerText()).replace(/\n/g, " | ")}")`);
    await ctx.close(); }
  { const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } }); const page = await newPage(ctx);
    const t = Date.now();
    await open(page, "chat-viewer", file, { params: { mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } });
    await page.waitForFunction(() => [...document.querySelectorAll(".rb-file-viewer-canvas")].some((c) => c.width >= 800), null, { timeout: 300000, polling: 30 });
    const first = Date.now() - t; const sub1 = await page.locator(".rb-file-viewer-header-copy .rb-app-subtitle").innerText();
    await page.waitForFunction(() => /^522 pages$/.test(document.querySelector(".rb-file-viewer-header-copy .rb-app-subtitle").textContent.trim()), null, { timeout: 300000, polling: 30 });
    console.log(`preview run ${run + 1}: first page painted ${first} ms ("${sub1}"); whole layout ${Date.now() - t} ms`);
    await ctx.close(); }
}
await browser.close();
