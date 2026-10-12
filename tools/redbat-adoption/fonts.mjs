// Which faces each document fetches, and whether the scan needed a second layout.
import { launch, newPage, open, T, S, fontStats, BASE, fixtureUrl } from "./lib.mjs";
const docs = [`${T}/test-documents/sample-formatted.docx`, `${T}/test-documents-wild/docx4j-getting-started.docx`, `${S}/long-generated.docx`, `${T}/repros/word/landscape-section.docx`, `${T}/test-documents-wild/docx4j-chart.docx`];
const { browser } = await launch();
for (const file of docs) {
  const ctx = await browser.newContext({ viewport: { width: 1366, height: 768 } });
  const page = await newPage(ctx, { verbose: false });
  const asked = [];
  page.on("console", (m) => { if (m.text().startsWith("word-text-stack:")) asked.push(m.text().slice(17)); });
  fontStats.bytes = 0; fontStats.files = [];
  const t = Date.now();
  const q = new URLSearchParams({ c: "word-preview", url: fixtureUrl(file), name: file.split("/").pop(), rbWordTextDebug: "1" });
  await page.goto(`${BASE}/office-harness.html?${q}`);
  await page.waitForFunction(() => window.__wordPages > 0, null, { timeout: 300000 });
  console.log(`${file.split("/").pop()}: ${await page.evaluate(() => window.__wordPages)} pages, ${Date.now() - t} ms, ${(fontStats.bytes / 1048576).toFixed(1)} MB in ${fontStats.files.length} files: ${fontStats.files.join(" ")}${asked.length ? " | second layout:" + asked.join(";") : ""} | editor modules requested: ${page.requests.filter((r) => /@rbt\/office\/dist\/word\/(Word|Input|WordHost|WordShell|WordUi|WordDialogs)\.js$/.test(r)).map((r) => r.split("/").pop()).join(",") || "none"}; office modules ${page.requests.filter((r) => /@rbt\/office\/dist\//.test(r)).length}`);
  await ctx.close();
}
await browser.close();
