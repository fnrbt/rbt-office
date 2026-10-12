// 0.1.9 in the Word editor's header / footer bar: entering and leaving is not a change; "Pages in section".
import { launch, newPage, open, T, S, SHOTS, check, summary, waitUploads, unzip } from "./lib.mjs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const frame = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const sel = (page, id) => page.locator(`[data-testid="${id}"]`);
const ctx = (page) => ev(page, () => window.__word.getCaretContext(window.__editors[0]));
const status = (page) => sel(page, "office-editor-status").innerText();
const rev = (page) => ev(page, () => window.__word.getRevision(window.__editors[0]));
const fingerprint = (page) => ev(page, () => window.__word.getLayoutFingerprint(window.__editors[0]));
const partText = (page, key) => ev(page, (key) => { const pkg = window.__word.exportDocumentPackage(window.__editors[0]); const text = (v) => (typeof v === "string" ? "" : Array.isArray(v) ? v.map(text).join("") : v && typeof v === "object" ? (typeof v.Text === "string" ? (v.FieldCode ? `{${v.FieldCode.trim()}}` : v.Text) : "") + Object.entries(v).filter(([k]) => k !== "Text").map(([, x]) => text(x)).join("") : ""); return (pkg[key] ?? []).map(text); }, key);
const headerPoint = (page, pageIndex, kind) => ev(page, ([i, kind]) => { const w = window.__word, e = window.__editors[0]; const f = w.getPageFrames(e)[i]; const y = kind === "header" ? f.Top + f.MarginTop / 2 : f.Top + f.Height - f.MarginBottom / 2; const v = w.documentToViewport(e, f.X + f.Width / 2, y); const r = document.querySelector('[data-testid="word-canvas"]').getBoundingClientRect(); return { x: r.left + v.X, y: r.top + v.Y }; }, [pageIndex, kind]);
for (const file of [`${T}/repros/word/landscape-section.docx`, `${T}/test-documents/sample-formatted.docx`, `${T}/test-documents-wild/phpword-sample-11-readword2007.docx`]) {
  const name = file.split("/").pop();
  const page = await newPage(context);
  await open(page, "word-editor", file);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await page.waitForFunction(() => window.__word.isLayoutComplete(window.__editors[0]), null, { timeout: 300000 });
  await page.waitForTimeout(400);
  const r0 = await rev(page), parts0 = JSON.stringify([await partText(page, "HeaderParts"), await partText(page, "FooterParts")]), f0 = JSON.stringify(await fingerprint(page));
  // double click into the header, leave with the bar's button
  const p = await headerPoint(page, 0, "header");
  await page.mouse.dblclick(p.x, p.y);
  await frame(page);
  check(`${name}: in the header after a double click`, (await ctx(page)).Story === "header" && (await sel(page, "word-header-footer-bar").count()) === 1);
  check(`${name}: entering a header is not a change (clean, same revision, Save disabled)`, (await status(page)).includes("No changes") && (await rev(page)) === r0 && (await sel(page, "office-editor-save").isDisabled()), { status: await status(page), revision: [r0, await rev(page)] });
  await sel(page, "word-header-footer-close").click();
  await frame(page);
  check(`${name}: opened and closed, the document is clean and as it was (parts, layout fingerprint)`, (await ctx(page)).Story === "body" && (await status(page)).includes("No changes") && (await rev(page)) === r0 && JSON.stringify([await partText(page, "HeaderParts"), await partText(page, "FooterParts")]) === parts0 && JSON.stringify(await fingerprint(page)) === f0, { status: await status(page), revision: [r0, await rev(page)], sameLayout: JSON.stringify(await fingerprint(page)) === f0 });
  // the footer through the menu, left with Escape; the frame's close then does not ask
  await sel(page, "word-header-footer").click();
  await sel(page, "word-edit-footer").click();
  await frame(page);
  check(`${name}: menu > Footer`, (await ctx(page)).Story === "footer");
  await page.keyboard.press("Escape");
  await frame(page);
  check(`${name}: Escape leaves the footer; still clean`, (await ctx(page)).Story === "body" && (await status(page)).includes("No changes") && (await rev(page)) === r0, await status(page));
  // Pages in section
  await sel(page, "word-header-footer").click();
  await sel(page, "word-edit-header").click();
  await frame(page);
  await page.keyboard.press("Control+End");
  await page.keyboard.type(" Sec ");
  await sel(page, "word-header-footer-section-pages").click();
  await frame(page);
  const parts = (await partText(page, "HeaderParts")).join("|");
  check(`${name}: "Pages in section" writes a SECTIONPAGES field into the header; dirty; keyboard in the document`, /Sec \{SECTIONPAGES[^}]*\}/.test(parts) && (await status(page)).includes("Unsaved") && (await ev(page, () => document.activeElement === window.__word.getInputElement(window.__editors[0]))), parts.slice(0, 300));
  await page.mouse.move(5, 5); await page.waitForTimeout(300);
  await page.screenshot({ path: `${SHOTS}/new-word-header-bar-${name.replace(/\W+/g, "-")}.png` });
  // save: the field is in the file
  await page.keyboard.press("Control+s");
  const zip = await unzip(await waitUploads(page, 1));
  let found = false;
  for (const [n, f] of zip) if (/word\/header\d*\.xml$/.test(n) && /SECTIONPAGES/.test(await f.async("string"))) found = true;
  check(`${name}: saved with the field in a header part`, found, [...zip.keys()].filter((n) => /header|footer/.test(n)));
  // undo everything: clean again? (saved at this revision, so undo makes it dirty; redo clean)
  await sel(page, "word-header-footer-close").click();
  await frame(page);
  check(`${name}: back in the body, saved`, (await ctx(page)).Story === "body" && (await status(page)).includes("Saved"), await status(page));
  await page.close();
}
await browser.close();
process.exit(summary() ? 1 : 0);
