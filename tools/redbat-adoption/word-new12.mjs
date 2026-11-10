// What 0.1.11 / 0.1.12 add to the Word host: cancel / dispose, fonts as they arrive, Unchanged saves.
import fs from "node:fs";
import { launch, newPage, open, S, T, check, summary, waitUploads, LONG_PAGES } from "./lib.mjs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const long = `${S}/long-generated.docx`, late = `${S}/late.docx`;
const status = (page) => page.locator('[data-testid="office-editor-status"]').innerText();
const raf = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ---- a preview that closes ends the reading and the layout of its document
// (How busy the page is afterwards: the time timers were kept waiting, summed over 1.2 s. Reading on
// takes about a second of 10 ms slices.)
const busyProbe = () => new Promise((resolve) => { let last = performance.now(), kept = 0, slices = 0; const start = last; const id = setInterval(() => { const now = performance.now(); if (now - last > 6) { kept += now - last; slices++; } last = now; if (now - start > 1200) { clearInterval(id); resolve({ kept: Math.round(kept), slices }); } }, 1); });
{
  const page = await newPage(context);
  await open(page, "word-preview", long);
  await page.waitForFunction(() => window.__wordLoaded != null, null, { timeout: 120000 });
  const atClose = await ev(page, () => { const v = window.__wordLoaded.viewer; const s = { blocks: v.getRenderStats().Blocks, load: v.isLoadComplete(), layout: v.isLayoutComplete(), pages: v.getPageFrames().length }; window.__setOpen(false); return s; });
  const busy = await ev(page, busyProbe);
  const later = await ev(page, () => { const v = window.__wordLoaded.viewer; return { blocks: v.getRenderStats().Blocks, slices: v.getRenderStats().LayoutSlices }; });
  await sleep(500);
  const last = await ev(page, () => { const v = window.__wordLoaded.viewer; return { blocks: v.getRenderStats().Blocks, slices: v.getRenderStats().LayoutSlices }; });
  console.log("  after the preview closed:", JSON.stringify({ atClose, busy, later, last }));
  check("preview closed while its document is read: it was still being read", atClose.load === false && atClose.blocks < 9008, atClose);
  check("... and nothing more of it is read or laid out afterwards (the page is idle)", later.blocks === atClose.blocks && last.blocks === later.blocks && last.slices === later.slices && busy.kept < 250, { atClose: atClose.blocks, later, last, busy });
  await page.close();
}

// ---- an editor that closes ends the reading
{
  const page = await newPage(context);
  await open(page, "word-editor", long);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  const atClose = await ev(page, () => { const w = window.__word, e = window.__editors[0]; const s = { blocks: w.getRenderStats(e).Blocks, load: w.isLoadComplete(e) }; window.__setOpen(false); return s; });
  const busy = await ev(page, busyProbe);
  const later = await ev(page, () => { const w = window.__word, e = window.__editors[0]; return { blocks: w.getRenderStats(e).Blocks }; });
  check("editor closed while its document is read: no block is read afterwards (the page is idle)", atClose.load === false && later.blocks === atClose.blocks && busy.kept < 250, { atClose, later, busy });
  await page.close();
}

// ---- the page count at open: provisional while a face is fetched, final at once when its bytes are at hand
{
  const page = await newPage(context);
  await open(page, "word-editor", long);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  // cold: the faces of the later blocks have to be fetched
  const cold = await ev(page, async () => {
    const w = window.__word, e = window.__editors[0]; const notices = []; const stop = w.onChange(e, (c) => notices.push({ layout: c.LayoutChanged, pending: c.FontsPending, complete: c.LayoutComplete }));
    const v0 = w.getRenderStats(e).MeasurerVersion;
    const first = w.getPageCount(e); const pendingAfter = w.getPendingFonts(e).map((f) => `${f.family}${f.bold ? " bold" : ""}${f.italic ? " italic" : ""}`); const size = w.getDocumentSize(e);
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    const sub = document.querySelector('[data-testid="office-editor-status"]').innerText.replace(/\n/g, " | ");
    await new Promise((r) => { const f = () => (w.getPendingFonts(e).length === 0 ? r() : setTimeout(f, 10)); f(); });
    await w.whenLayoutComplete(e);
    stop();
    return { first, pendingAfter, sizePending: size.FontsPending, sub, final: w.getPageCount(e), installs: w.getRenderStats(e).MeasurerVersion - v0, noticesPending: notices.filter((n) => n.pending).length, lastNotice: notices[notices.length - 1], files: window.RedBatWordTextStack.loadedFaceFiles() };
  });
  await raf(page);
  console.log("  cold:", JSON.stringify(cold));
  check("cold: the count asked for at open is known to be provisional while faces are fetched", cold.pendingAfter.length > 0 ? cold.sizePending === true && /about \d+ pages…/.test(cold.sub) : cold.first === LONG_PAGES, { first: cold.first, pending: cold.pendingAfter, sub: cold.sub });
  check("cold: with the faces there the count is the document's and nothing is pending", cold.final === LONG_PAGES && new RegExp(`No changes \\| ${LONG_PAGES} pages$`).test((await status(page)).replace(/\n/g, " | ").trim()), { final: cold.final, status: await status(page) });
  // warm: the same document again in the same session - the faces' bytes are at hand
  await ev(page, () => { window.__setOpen(false); });
  await raf(page);
  await ev(page, () => { window.__editors.length = 0; window.__setOpen(true); });
  await page.waitForFunction(() => window.__editors.length === 1 && document.querySelector('[data-testid="word-toolbar"]') && !document.querySelector(".rb-office-editor-cover"), null, { timeout: 300000 });
  const warm = await ev(page, async () => {
    const w = window.__word, e = window.__editors[0]; const notices = []; const stop = w.onChange(e, (c) => notices.push({ layout: c.LayoutChanged, pending: c.FontsPending }));
    const before = w.isLoadComplete(e); const v0 = w.getRenderStats(e).MeasurerVersion;
    const first = w.getPageCount(e); const pending = w.getPendingFonts(e).length; const size = w.getDocumentSize(e);
    await new Promise((r) => setTimeout(r, 600));
    stop();
    return { before, first, pending, sizePending: size.FontsPending, later: w.getPageCount(e), installs: w.getRenderStats(e).MeasurerVersion - v0, noticesPending: notices.filter((n) => n.pending).length };
  });
  console.log("  warm:", JSON.stringify(warm));
  check("warm: a page count forced at open (document still being read) is the final one the first time", warm.before === false && warm.first === LONG_PAGES && warm.later === LONG_PAGES && warm.pending === 0 && warm.sizePending === false && warm.noticesPending === 0, warm);
  await raf(page);
  check("warm: the status shows it without a provisional mark", new RegExp(`No changes \\| ${LONG_PAGES} pages$`).test((await status(page)).replace(/\n/g, " | ").trim()), await status(page));
  await page.close();
}

// ---- faces arrive one by one: the layout is the one made with all of them from the start
for (const file of [late, long]) {
  const page = await newPage(context);
  await open(page, "word-editor", file);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  const got = await ev(page, async () => {
    const w = window.__word, e = window.__editors[0], m = window.__wordModule;
    const snap = () => ({ pages: w.getPageCount(e), frames: JSON.stringify(w.getPageFrames(e).map((p) => [Math.round(p.Top), Math.round(p.Height)])), height: w.getDocumentSize(e).Height, fp: (() => { const t = JSON.stringify(m.getLayoutFingerprint(e)); let h = 0; for (let i = 0; i < t.length; i++) h = (h * 31 + t.charCodeAt(i)) | 0; return `${t.length}:${h}`; })() });
    await w.whenLoadComplete(e);
    await new Promise((r) => { const f = () => (w.getPendingFonts(e).length === 0 ? r() : setTimeout(f, 10)); f(); });
    await w.whenLayoutComplete(e);
    const installs = w.getRenderStats(e).MeasurerVersion;
    const incremental = snap();
    // the same document laid out from nothing with the font set as it is now
    const stack = await window.__stack();
    stack.clearCaches();
    m.setTextMeasurers(e, stack.measureText, stack.measureRange);
    w.completeLayout(e);
    return { incremental, scratch: snap(), installs, files: stack.loadedFaceFiles() };
  });
  const name = file.split("/").pop();
  console.log(`  ${name}: pages ${got.incremental.pages}, measurer installs while opening ${got.installs}, faces ${got.files.join(", ")}`);
  check(`${name}: with faces added as they arrive the pages are those of a layout from nothing with every face`, got.incremental.pages === got.scratch.pages && got.incremental.frames === got.scratch.frames && got.incremental.height === got.scratch.height && got.incremental.fp === got.scratch.fp, { incremental: [got.incremental.pages, got.incremental.height], scratch: [got.scratch.pages, got.scratch.height], fp: got.incremental.fp === got.scratch.fp });
  await page.close();
}

// ---- a face that cannot be fetched: the document keeps its shaped text, nothing stays pending, no fetch loop
{
  const page = await newPage(context);
  await page.route("**/word-text-stack/fonts/arialbd.ttf", (route) => route.fulfill({ status: 404, body: "gone" }));
  await open(page, "word-editor", long);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  const got = await ev(page, async () => {
    const w = window.__word, e = window.__editors[0];
    await w.whenLoadComplete(e);
    const t = performance.now();
    await new Promise((r) => { const f = () => (w.getPendingFonts(e).length === 0 || performance.now() - t > 15000 ? r() : setTimeout(f, 20)); f(); });
    await w.whenLayoutComplete(e);
    await new Promise((r) => setTimeout(r, 800));
    return { pending: w.getPendingFonts(e).length, pages: w.getPageCount(e), files: window.RedBatWordTextStack.loadedFaceFiles(), notice: [...document.querySelectorAll(".rb-office-notice")].map((n) => n.textContent).join(" | ") };
  });
  await raf(page);
  const asked = page.requests.filter((p) => p.endsWith("/arialbd.ttf")).length;
  console.log("  without arialbd.ttf:", JSON.stringify({ ...got, files: got.files.length, asked }), await status(page));
  check("a face that cannot be fetched: the others are used, nothing stays pending, the status is final", got.pending === 0 && !got.files.includes("arialbd.ttf") && got.files.includes("times.ttf") && got.notice === "" && /\n\d+ pages$/.test((await status(page)).trim()), { got, status: await status(page) });
  check("... and it is not asked for again and again", asked >= 1 && asked <= 3, asked);
  await page.close();
}

// ---- a save of a document that is what was opened again (edit, save, undo, save) is the opened file
{
  const file = `${T}/test-documents/sample-formatted.docx`;
  const original = fs.readFileSync(file);
  const page = await newPage(context);
  await open(page, "word-editor", file);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 120000 });
  await page.click('[data-testid="word-canvas"]', { position: { x: 300, y: 200 } });
  await page.keyboard.type("x");
  await raf(page);
  await page.click('[data-testid="office-editor-save"]');
  const first = await waitUploads(page, 1);
  check("an edited document is written by the writer (not the opened bytes)", !first.equals(original) && first.length > 1000, { bytes: first.length, original: original.length });
  await page.keyboard.press("Control+z");
  await raf(page);
  check("undo after the save is a change of what was saved", /Unsaved changes/.test(await status(page)), await status(page));
  const unchanged = await ev(page, () => window.__word.exportDocumentPackage(window.__editors[0]).Unchanged);
  await page.click('[data-testid="office-editor-save"]');
  const second = await waitUploads(page, 2);
  check("the document as it was opened again is saved as the opened file, byte for byte", unchanged === true && second.equals(original), { unchanged, bytes: second.length, original: original.length });
  check("and the editor is clean", /Saved/.test(await status(page)), await status(page));
  await page.close();
}

await browser.close();
process.exit(summary() ? 1 : 0);
