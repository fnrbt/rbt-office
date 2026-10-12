// A bookmark around a run that is split by typing: how often is it written? usage: PORT=.. node probe-bm.mjs
import { launch, newPage, open, T, waitUploads, unzip } from "./lib.mjs";
const { browser, context } = await launch();
for (const mode of (process.argv[2] ?? "A,B,C,D,E").split(",")) {
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/test-documents/sample-formatted.docx`);
  await page.waitForSelector('[data-testid="word-toolbar"]'); await page.waitForTimeout(1500);
  await page.evaluate(() => { const w = window.__word, m = window.__wordModule, e = window.__editors[0]; w.selectFindMatch(e, w.findText(e, "Demonstration of DOCX", true)[0]); const s = m.getSelection(e); m.setSelection(e, s.AnchorBlockId, 4, s.AnchorBlockId, 4); w.getInputElement(e).focus({ preventScroll: true }); });
  const bold = () => page.locator('[data-testid="word-bold"]').click();
  const left = async (n) => { for (let i = 0; i < n; i++) await page.keyboard.press("ArrowLeft"); };
  if (mode === "A") { await bold(); await page.keyboard.type("QQ"); }                                   // Bold at the caret, two characters
  if (mode === "F") { await page.keyboard.type("QQ"); await page.keyboard.press("Shift+ArrowLeft"); await page.keyboard.press("Shift+ArrowLeft"); await bold(); } // two characters typed, selected, made bold
  if (mode === "G") { await bold(); await page.keyboard.type("Q"); await page.waitForTimeout(1200); await page.keyboard.type("Q"); } // as A, slowly
  if (mode === "B") { await page.keyboard.type("Q"); await left(3); await page.keyboard.type("Z"); }     // two plain insertions at two places
  if (mode === "C") { await bold(); await page.keyboard.type("Q"); await left(3); await page.keyboard.type("Z"); } // a bold character, then a plain one elsewhere
  if (mode === "D") { for (let i = 0; i < 3; i++) await page.keyboard.press("Shift+ArrowRight"); await bold(); await page.keyboard.press("Home"); for (let i = 0; i < 2; i++) await page.keyboard.press("Shift+ArrowRight"); await page.locator('[data-testid="word-italic"]').click(); } // two ranges formatted
  if (mode === "E") { for (let i = 0; i < 3; i++) await page.keyboard.press("Shift+ArrowRight"); await bold(); await page.keyboard.press("End"); await page.keyboard.type("Z"); } // a range made bold, then typing at the end
  await page.waitForTimeout(200);
  await page.locator('[data-testid="office-editor-save"]').click();
  const doc = await (await unzip(await waitUploads(page, 1))).get("word/document.xml").async("string");
  console.log("    report", JSON.stringify(await page.evaluate(async () => (await window.__formats()).default().word.getLastSaveReport())));
  const p = doc.slice(doc.indexOf("<w:body>"), doc.indexOf("</w:p>"));
  const starts = [...doc.matchAll(/<w:bookmarkStart [^>]*w:name="([^"]*)"[^>]*>/g)].map((m) => m[0]);
  console.log(`${mode}: bookmarkStart in the document ${starts.length}, OLE_LINK1 ${starts.filter((s) => s.includes("OLE_LINK1")).length} times, bookmarkEnd ${(doc.match(/<w:bookmarkEnd /g) ?? []).length}`);
  console.log("   ", p.replace(/<w:rPr>.*?<\/w:rPr>/g, "").slice(0, 600));
  await page.close();
}
await browser.close();
