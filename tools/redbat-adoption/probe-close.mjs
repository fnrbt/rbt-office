import { launch, newPage, open, S } from "./lib.mjs";
const { browser, context } = await launch();
const busyProbe = () => new Promise((resolve) => { let last = performance.now(), kept = 0, slices = 0; const start = last; const id = setInterval(() => { const now = performance.now(); if (now - last > 6) { kept += now - last; slices++; } last = now; if (now - start > 1200) { clearInterval(id); resolve({ kept: Math.round(kept), slices }); } }, 1); });
for (let i = 0; i < 3; i++) {
  const page = await newPage(context);
  await open(page, "word-preview", `${S}/long-generated.docx`);
  await page.waitForFunction(() => window.__wordLoaded != null, null, { timeout: 120000 });
  await page.evaluate(() => window.__setOpen(false));
  console.log("preview closed at open: page kept busy for", JSON.stringify(await page.evaluate(busyProbe)));
  await page.close();
}
await browser.close();
