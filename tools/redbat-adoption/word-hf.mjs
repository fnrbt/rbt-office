// Header and footer editing: contextual bar, menu, link to previous, page number fields, save round trip.
import { launch, newPage, open, T, S, SHOTS, check, summary, waitUploads, unzip, saved } from "./lib.mjs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const ctx = (page) => ev(page, () => window.__word.getCaretContext(window.__editors[0]));
const status = (page) => sel(page, "office-editor-status").innerText();
const partText = (page, key) => ev(page, (key) => { const pkg = window.__word.exportDocumentPackage(window.__editors[0]); const text = (v) => (typeof v === "string" ? "" : Array.isArray(v) ? v.map(text).join("") : v && typeof v === "object" ? (typeof v.Text === "string" ? (v.FieldCode ? `{${v.FieldCode.trim()}}` : v.Text) : "") + Object.entries(v).filter(([k]) => k !== "Text").map(([, x]) => text(x)).join("") : ""); return (pkg[key] ?? []).map(text); }, key);
const headerPoint = (page, pageIndex, kind) => ev(page, ([i, kind]) => { const w = window.__word, e = window.__editors[0]; const f = w.getPageFrames(e)[i]; const y = kind === "header" ? f.Top + f.MarginTop / 2 : f.Top + f.Height - f.MarginBottom / 2; const v = w.documentToViewport(e, f.X + f.Width / 2, y); const r = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: r.left + v.X, y: r.top + v.Y }; }, [pageIndex, kind]);
{
  const file = `${T}/test-documents/sample-formatted.docx`;
  const page = await newPage(context);
  await open(page, "word-editor", file);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await page.waitForTimeout(500);
  check("no header/footer bar in the body", (await sel(page, "word-header-footer-bar").count()) === 0 && (await ctx(page)).Story === "body");
  // a real double click in the header area of page 1
  const p = await headerPoint(page, 0, "header");
  await page.mouse.dblclick(p.x, p.y);
  await frame(page);
  let c = await ctx(page);
  check("double click in the header: the engine edits it and the bar shows which", c.InHeaderFooter && c.Story === "header" && c.HeaderFooter?.Kind === "header" && (await sel(page, "word-header-footer-bar").count()) === 1 && /^Header.* - section 1$/.test(await sel(page, "word-header-footer-name").innerText()), { hf: c.HeaderFooter, name: await sel(page, "word-header-footer-name").innerText().catch(() => null) });
  check("link to previous is disabled in the first section", await sel(page, "word-header-footer-link").isDisabled());
  const before = (await partText(page, "HeaderParts")).join("|");
  await page.keyboard.type("HDR page ");
  await sel(page, "word-header-footer-page-number").click();
  await page.keyboard.type(" of ");
  await sel(page, "word-header-footer-page-count").click();
  await frame(page);
  let parts = (await partText(page, "HeaderParts")).join("|");
  check("typing and the Page number / Page count buttons write into the header", /HDR page \{PAGE[^}]*\} of \{NUMPAGES[^}]*\}/.test(parts), parts);
  check("keyboard still in the document after the buttons, and dirty", (await ev(page, () => document.activeElement === window.__word.getInputElement(window.__editors[0]))) && (await status(page)).includes("Unsaved"));
  await page.screenshot({ path: `${SHOTS}/word-header-bar.png` });
  await page.keyboard.press("Control+z");
  await frame(page);
  const afterUndo = (await partText(page, "HeaderParts")).join("|");
  check("undo takes the page count field out again", !/NUMPAGES/.test(afterUndo) && /PAGE/.test(afterUndo), afterUndo);
  await page.keyboard.press("Control+y");
  await sel(page, "word-header-footer-close").click();
  await frame(page);
  c = await ctx(page);
  check("Close goes back to the body", !c.InHeaderFooter && c.Story === "body" && (await sel(page, "word-header-footer-bar").count()) === 0, c);
  // the menu: footer
  await sel(page, "word-header-footer").click();
  await page.screenshot({ path: `${SHOTS}/word-header-footer-menu.png` });
  await sel(page, "word-edit-footer").click();
  await frame(page);
  c = await ctx(page);
  check("menu > Footer edits the footer of the caret's page", c.Story === "footer" && /^Footer/.test(await sel(page, "word-header-footer-name").innerText()), c.HeaderFooter);
  await page.keyboard.type("FTR");
  await page.keyboard.press("Escape");
  await frame(page);
  check("Escape leaves the footer and not the editor", (await ctx(page)).Story === "body" && (await sel(page, "office-editor").count()) === 1 && (await sel(page, "office-editor-confirm").count()) === 0 && (await sel(page, "word-header-footer-bar").count()) === 0);
  check("the footer has the text", (await partText(page, "FooterParts")).join("|").includes("FTR"));
  // save and re-open
  await sel(page, "office-editor-save").click();
  const bytes = await waitUploads(page, 1);
  await page.waitForTimeout(200);
  check("saved", (await status(page)).includes("Saved"), await status(page));
  const z = await unzip(bytes);
  let hdr = "", ftr = "";
  for (const [n, f] of z) { if (/^word\/header\d*\.xml$/.test(n)) hdr += await f.async("string"); if (/^word\/footer\d*\.xml$/.test(n)) ftr += await f.async("string"); }
  check("saved file: header part with the text and PAGE / NUMPAGES fields, footer part with its text", /HDR page/.test(hdr) && /PAGE/.test(hdr) && /NUMPAGES/.test(hdr) && /FTR/.test(ftr), { hdr: hdr.length, ftr: ftr.length });
  saved.set("hf", bytes);
  await page.close();
  const page2 = await newPage(context);
  await open(page2, "word-editor", "/api/saved/hf", { name: "sample-formatted.docx" });
  await page2.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await page2.waitForTimeout(400);
  parts = (await partText(page2, "HeaderParts")).join("|");
  check("re-open: the header and footer are there, clean", /HDR page \{PAGE[^}]*\} of \{NUMPAGES[^}]*\}/.test(parts) && (await partText(page2, "FooterParts")).join("|").includes("FTR") && (await status(page2)).includes("No changes"), parts);
  // pixels: the header text is painted on page 1
  const hp = await headerPoint(page2, 0, "header");
  const ink = await ev(page2, ([y]) => { const c = document.querySelector('[data-testid="word-canvas"]'); const r = c.getBoundingClientRect(); const k = c.width / r.width; const d = c.getContext("2d").getImageData(0, Math.max(0, Math.round((y - r.top - 30) * k)), c.width, Math.round(60 * k)).data; let n = 0; for (let i = 0; i < d.length; i += 4) if (d[i] < 110 && d[i + 1] < 110 && d[i + 2] < 110) n++; return n; }, [hp.y]);
  check("re-open: the header is painted on the page", ink > 80, ink);
  await page2.screenshot({ path: `${SHOTS}/word-header-reopened.png` });
  // page number from the body: goes into the footer
  await ev(page2, () => window.__word.getInputElement(window.__editors[0]).focus());
  await sel(page2, "word-header-footer").click();
  await sel(page2, "word-insert-page-number").click();
  await frame(page2);
  const c2 = await ctx(page2);
  const f2 = (await partText(page2, "FooterParts")).join("|");
  check("menu > Page number from the body: the footer is entered and gets a PAGE field", c2.Story === "footer" && /\{PAGE/.test(f2), { story: c2.Story, f2 });
  await page2.keyboard.press("Control+z");
  await frame(page2);
  check("undo removes the field", !/\{PAGE/.test((await partText(page2, "FooterParts")).join("|")));
  await page2.close();
}
{ // sections: link to previous
  const page = await newPage(context);
  await open(page, "word-editor", `${T}/repros/word/landscape-section.docx`);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await page.waitForTimeout(400);
  const pages = await ev(page, () => window.__word.getPageCount(window.__editors[0]));
  let found = null;
  for (let i = pages - 1; i >= 1 && !found; i--) { await ev(page, (i) => window.__word.enterHeaderFooter(window.__editors[0], i, "header"), i); const c = await ctx(page); if (c.HeaderFooter?.CanLinkToPrevious) found = { i, hf: c.HeaderFooter }; else await ev(page, () => window.__word.exitHeaderFooter(window.__editors[0])); }
  if (!found) console.log("  fixture has no second section whose header can link to the previous one - link toggle not exercised");
  else {
    await frame(page);
    const was = found.hf.LinkedToPrevious;
    check("in a later section the link box is enabled and shows the state", !(await sel(page, "word-header-footer-link").isDisabled()) && (await sel(page, "word-header-footer-link").isChecked()) === was && new RegExp(`section ${found.hf.Section + 1}$`).test(await sel(page, "word-header-footer-name").innerText()), found);
    await sel(page, "word-header-footer-link").click();
    await frame(page);
    let c = await ctx(page);
    check("clicking it changes Link to Previous in the model and the box", c.HeaderFooter.LinkedToPrevious === !was && (await sel(page, "word-header-footer-link").isChecked()) === !was && (await ev(page, () => window.__word.isModified(window.__editors[0]))), c.HeaderFooter);
    await page.screenshot({ path: `${SHOTS}/word-header-link.png` });
    await page.keyboard.press("Control+z");
    await frame(page);
    c = await ctx(page);
    if (c.Story === "body") console.log("  [package] undo of Link to Previous leaves the header (story is body again); re-entering to read the state");
    await ev(page, (i) => window.__word.enterHeaderFooter(window.__editors[0], i, "header"), found.i);
    await frame(page);
    c = await ctx(page);
    check("undo reverts it", c.HeaderFooter?.LinkedToPrevious === was && (await sel(page, "word-header-footer-link").isChecked()) === was, c.HeaderFooter);
  }
  await page.close();
}
summary();
await browser.close();
