import { launch, newPage, open, S } from "./lib.mjs";
const { browser } = await launch();
for (let run = 0; run < 3; run++) {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });
  const page = await newPage(ctx);
  const t = Date.now();
  await open(page, "word-preview", `${S}/long-generated.docx`);
  await page.waitForFunction(() => window.__wordPages > 0, null, { timeout: 300000 });
  const ready = Date.now() - t;
  await page.waitForFunction(() => [...document.querySelectorAll(".rb-file-viewer-canvas")].some((c) => c.width >= 800), null, { timeout: 60000 });
  console.log(`preview open (run ${run + 1}): pages known ${ready} ms, first page painted ${Date.now() - t} ms`);
  await ctx.close();
}
await browser.close();
