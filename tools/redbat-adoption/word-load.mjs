// Progressive read (0.1.10): the editor and the preview show a long document while its package is
// still being read; typing and saving meanwhile; pictures and fonts of blocks that are read later.
import fs from "node:fs";
import { launch, newPage, open, S, SHOTS, check, summary, waitUploads, saved, LONG_PAGES } from "./lib.mjs";
const { browser, context } = await launch();
const ev = (page, fn, arg) => page.evaluate(fn, arg);
const long = `${S}/long-generated.docx`, late = `${S}/late.docx`;
const status = (page) => page.locator('[data-testid="office-editor-status"]').innerText();
const raf = (page) => ev(page, () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
const longTasks = () => { window.__lt = []; new PerformanceObserver((l) => { for (const e of l.getEntries()) window.__lt.push([Math.round(e.startTime), Math.round(e.duration)]); }).observe({ entryTypes: ["longtask"] }); };
const state = (page) => ev(page, () => { const w = window.__word, e = window.__editors[0]; return { load: w.isLoadComplete(e), layout: w.isLayoutComplete(e), size: w.getDocumentSize(e), modified: w.isModified(e), canUndo: w.canUndo(e) }; });
const colourful = (page, selector) => ev(page, (selector) => { let n = 0; for (const c of document.querySelectorAll(selector)) { if (!c.width || !c.height) continue; const d = c.getContext("2d").getImageData(0, 0, c.width, c.height).data; for (let i = 0; i < d.length; i += 4) { const mx = Math.max(d[i], d[i + 1], d[i + 2]), mn = Math.min(d[i], d[i + 1], d[i + 2]); if (d[i + 3] > 200 && mx - mn > 60) n++; } } return n; }, selector);

// ---- the reference: the long document's text and pages, read to the end
let reference;
{
  const page = await newPage(context);
  await open(page, "word-editor", long);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  reference = await ev(page, async () => { const w = window.__word, e = window.__editors[0]; await w.whenLoadComplete(e); await w.whenLayoutComplete(e); return { text: w.getText(e), pages: w.getPageCount(e), blocks: w.getRenderStats(e).Blocks }; });
  check("the reference document is long", reference.pages === LONG_PAGES && reference.text.length > 500000, { pages: reference.pages, chars: reference.text.length, blocks: reference.blocks });
  await page.close();
}

// ---- the editor shows the document while it is read; nothing the host does at open reads the rest
{
  const page = await newPage(context);
  await page.addInitScript(longTasks);
  await open(page, "word-editor", long);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  const first = await state(page);
  const firstStatus = await status(page);
  check("the editor is shown while the package is still being read", first.load === false && first.layout === false && first.size.LayoutComplete === false && first.size.PageCount >= 1 && first.size.PageCount < 100, first);
  check("the status says there is more to come, and the document is clean", /No changes\s+\d+\+ pages…/.test(firstStatus) && !first.modified && !first.canUndo, firstStatus);
  // two frames of the host's own bookkeeping (toolbar state, viewport, caret): still not read to the end
  await raf(page); await raf(page);
  const settled = await state(page);
  check("the host's toolbar, status and viewport bookkeeping do not read the rest", settled.load === false, settled);
  // the notices: the rest arrives in slices, then the layout completes
  const notes = await ev(page, () => new Promise((resolve) => { const w = window.__word, e = window.__editors[0]; const out = { load: 0, loadDocChanged: 0, completeAt: null, layoutAt: null, t0: performance.now() }; const stop = w.onChange(e, (c) => { if (c.LoadChanged) { out.load++; if (c.DocumentChanged || c.SelectionChanged) out.loadDocChanged++; if (c.LoadComplete && out.completeAt == null) out.completeAt = Math.round(performance.now() - out.t0); } if (c.LayoutComplete) { out.layoutAt = Math.round(performance.now() - out.t0); stop(); resolve(out); } }); }));
  const lt = await ev(page, () => window.__lt);
  const bytesAt = await ev(page, () => performance.getEntriesByType("resource").find((e) => e.name.includes("/api/fixture/")).responseEnd);
  const longest = Math.max(0, ...lt.filter(([s]) => s > bytesAt).map(([, d]) => d));
  check("the rest is read in slices behind the editor (several notices, none a change of the document), then laid out", notes.load >= 5 && notes.loadDocChanged === 0 && notes.completeAt != null && notes.layoutAt >= notes.completeAt, notes);
  check(`no task holds the page for long while it is read and laid out (longest ${longest} ms)`, longest < 700, lt.filter(([s, d]) => s > bytesAt && d > 100));
  await page.waitForTimeout(200);
  const end = await ev(page, () => { const w = window.__word, e = window.__editors[0]; return { load: w.isLoadComplete(e), pages: w.getDocumentSize(e).PageCount, text: w.getText(e), modified: w.isModified(e) }; });
  check("read and laid out in the background it is the whole document: text and pages", end.load && end.pages === reference.pages && end.text === reference.text && !end.modified, { pages: end.pages, chars: end.text.length });
  check("the status shows the page count once it is there", new RegExp(`No changes\\s+${LONG_PAGES} pages$`).test((await status(page)).trim()), await status(page));
  await page.close();
}

// ---- typing at once, then saving while it is still being read
{
  const page = await newPage(context);
  await open(page, "word-editor", long);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await page.keyboard.type("TYPED-AT-ONCE ");
  const typed = await state(page);
  const head = await ev(page, () => document.querySelector('[data-testid="office-editor-status"]').innerText);
  check("typing right after the first paint is an edit like any other, with the document still being read", typed.modified && typed.canUndo && typed.load === false, { typed, head });
  await page.keyboard.press("Control+s");
  const bytes = await waitUploads(page, 1);
  const after = await ev(page, () => { const w = window.__word, e = window.__editors[0]; return { load: w.isLoadComplete(e), modified: w.isModified(e), text: w.getText(e), pages: w.getPageCount(e) }; });
  check("saving reads the rest: the editor holds the whole document with what was typed, and is clean", after.load && !after.modified && after.text === "TYPED-AT-ONCE " + reference.text, { chars: after.text.length, pages: after.pages });
  check("the save's status", /Saved/.test(await status(page)), await status(page));
  // the saved file, opened again
  saved.set("typed.docx", bytes);
  const again = await newPage(context);
  await open(again, "word-editor", "/api/saved/typed.docx", { name: "typed.docx" });
  await again.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  const reopened = await ev(again, async () => { const w = window.__word, e = window.__editors[0]; await w.whenLoadComplete(e); await w.whenLayoutComplete(e); return { text: w.getText(e), pages: w.getPageCount(e), blocks: w.getRenderStats(e).Blocks }; });
  check("the saved file re-opens as the whole document with what was typed", reopened.text === "TYPED-AT-ONCE " + reference.text && reopened.blocks === reference.blocks, { chars: [reopened.text.length, reference.text.length + 14], blocks: [reopened.blocks, reference.blocks] });
  check("with the pages the editor that saved it has", reopened.pages === after.pages && Math.abs(reopened.pages - reference.pages) <= 1, { reopened: reopened.pages, editor: after.pages, original: reference.pages });
  await again.close();
  // undo back to the loaded document is clean again... of the saved revision: the typing is the saved state now
  await page.keyboard.press("Control+z");
  await raf(page);
  const undone = await ev(page, () => { const w = window.__word, e = window.__editors[0]; return { modified: w.isModified(e), head: w.getText(e).slice(0, 14) }; });
  check("undo after the save takes the typing back (a change against the saved file)", undone.modified && undone.head !== "TYPED-AT-ONCE ", undone);
  await page.close();
}

// ---- typing, letting the rest arrive, then saving
{
  const page = await newPage(context);
  await open(page, "word-editor", long);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await page.keyboard.type("ONE ");
  await ev(page, () => window.__word.whenLoadComplete(window.__editors[0]));
  await page.keyboard.type("TWO ");
  await ev(page, () => window.__word.whenLayoutComplete(window.__editors[0]));
  const now = await ev(page, () => { const w = window.__word, e = window.__editors[0]; return { text: w.getText(e), pages: w.getPageCount(e) }; });
  check("what is typed while it is read stays where it was typed", now.text === "ONE TWO " + reference.text, now.text.slice(0, 40));
  await page.keyboard.press("Control+z"); await page.keyboard.press("Control+z");
  await raf(page);
  let back = await ev(page, () => { const w = window.__word, e = window.__editors[0]; return { text: w.getText(e), modified: w.isModified(e) }; });
  for (let i = 0; i < 12 && back.modified; i++) { await page.keyboard.press("Control+z"); await raf(page); back = await ev(page, () => { const w = window.__word, e = window.__editors[0]; return { text: w.getText(e), modified: w.isModified(e) }; }); }
  check("undo back to what was loaded is the whole document, unchanged", !back.modified && back.text === reference.text, { modified: back.modified, chars: back.text.length });
  check("and the status is clean", /No changes/.test(await status(page)), await status(page));
  await page.close();
}

// ---- the page count asked for at once (reads and lays out the rest): right once the late faces are in
{
  const page = await newPage(context);
  await open(page, "word-editor", long);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  const forced = await ev(page, () => { const w = window.__word, e = window.__editors[0]; const before = w.isLoadComplete(e); return { before, pages: w.getPageCount(e), after: w.isLoadComplete(e) }; });
  const settled = await ev(page, async () => { const w = window.__word, e = window.__editors[0]; await new Promise((r) => { const f = () => (window.__word.getPendingFonts(window.__editors[0]).length === 0 ? r() : setTimeout(f, 20)); f(); }); await w.whenLayoutComplete(e); return { pages: w.getPageCount(e), text: w.getText(e).length }; });
  await raf(page);
  check("a page count forced at open reads the rest; with the later blocks' faces loaded the pages are the document's", forced.before === false && forced.after === true && settled.pages === reference.pages && settled.text === reference.text.length, { forced, settled });
  check("and the status shows them", new RegExp(`No changes\\s+${LONG_PAGES} pages$`).test((await status(page)).trim()), await status(page));
  await page.close();
}

// ---- find while it is read: the whole document is searched
{
  const page = await newPage(context);
  await open(page, "word-editor", long);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  const before = await state(page);
  await page.click('[data-testid="word-find"]');
  const opened = await state(page);
  check("opening the find bar reads nothing", before.load === false && opened.load === false, { before: before.load, opened: opened.load });
  const word = reference.text.slice(reference.text.length - 400).match(/[A-Za-z]{6,}/)?.[0] ?? "the";
  await page.fill('[data-testid="word-find-input"]', word);
  await page.keyboard.press("Enter");
  const expected = reference.text.toLowerCase().split(word.toLowerCase()).length - 1;
  const found = await page.locator('[data-testid="word-find-status"]').innerText();
  check(`find sees the whole document ("${word}")`, new RegExp(`^1 of ${expected}$`).test(found.trim()) || /^1 of \d+$/.test(found.trim()) && Number(found.trim().split(" of ")[1]) >= expected, { found, expected });
  await page.close();
}

// ---- pictures and fonts of blocks that are read after the first screen is shown: the editor
{
  const page = await newPage(context);
  await open(page, "word-editor", late);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  const first = await state(page);
  const facesFirst = await ev(page, () => window.RedBatWordTextStack.loadedFaceFiles());
  // (Whether this short-paragraph document is still being read when the editor shows it depends on how long
  // the fonts take to arrive; the long document above is the one that asserts it.)
  console.log(`  late.docx at first paint: read to its end ${first.load}; faces loaded ${facesFirst.join(", ")}`);
  check("the face only its last paragraph uses is not among those the first screen waited for", first.load === true || !facesFirst.some((f) => /^georgia/.test(f)), facesFirst);
  await ev(page, async () => { const w = window.__word, e = window.__editors[0]; await w.whenLoadComplete(e); });
  await page.waitForFunction(() => window.RedBatWordTextStack.loadedFaceFiles().includes("georgiaz.ttf"), null, { timeout: 30000 }).catch(() => {});
  const faces = await ev(page, () => window.RedBatWordTextStack.loadedFaceFiles());
  check("the face a later block needs is fetched and loaded when the document is read, before the layout asks for it", faces.includes("georgiaz.ttf") && page.requests.some((p) => p.endsWith("/georgiaz.ttf")), faces);
  const debug = await ev(page, () => window.RedBatWordTextStack.getStats?.());
  await ev(page, () => window.__word.whenLayoutComplete(window.__editors[0]));
  await ev(page, () => window.__word.getInputElement(window.__editors[0]).focus());
  await page.keyboard.press("Control+End");
  await raf(page); await page.waitForTimeout(400);
  // the picture is a little above the end: scroll up until it shows
  let seen = 0;
  for (let i = 0; i < 12 && seen < 2000; i++) { seen = await colourful(page, '[data-testid="word-canvas"]'); if (seen < 2000) { await ev(page, () => { const sc = document.querySelector('[data-testid="word-scroll"]'); sc.scrollTop -= 500; }); await raf(page); await page.waitForTimeout(150); } }
  await page.screenshot({ path: `${SHOTS}/word-late-picture-editor.png` });
  check("the picture of a block that was read later is drawn in the editor", seen >= 2000, seen);
  const cached = await ev(page, () => { const c = window.__editors[0].Handler.RenderContext.ImageCache; const m = c && typeof c.size === "number" ? c : c?.fields?.[0] ?? c?.Value; return m ? [...m.keys()].filter((k) => k.startsWith("data:image/")).length : -1; });
  check("it was decoded into the editor's image cache by the host", cached >= 1, cached);
  const text = await ev(page, () => window.__word.getText(window.__editors[0]));
  check("the document is whole", text.includes("LATE-FONT-MARKER") && text.includes("Filler paragraph 9000."), text.length);
  await page.close();
}

// ---- the preview: first pages while it is read, pictures and fonts of the later blocks
{
  const page = await newPage(context);
  await open(page, "word-preview", long);
  await page.waitForFunction(() => window.__wordLoaded != null, null, { timeout: 300000, polling: 10 });
  const first = await ev(page, () => ({ load: window.__wordLoaded.viewer.isLoadComplete(), layout: window.__wordLoaded.layoutComplete, pages: window.__wordFirst }));
  check("the preview shows the first pages while the package is still being read", first.load === false && first.layout === false && first.pages >= 1 && first.pages < 100, first);
  await page.waitForFunction(() => window.__wordPages > 0, null, { timeout: 300000 });
  const end = await ev(page, () => ({ pages: window.__wordPages, load: window.__wordLoaded.loadComplete, canvases: document.querySelectorAll(".rb-file-viewer-canvas").length }));
  check("and all of its pages once it is read and laid out", end.pages === reference.pages && end.load === true && end.canvases === reference.pages, end);
  await page.close();
}
{
  const page = await newPage(context);
  await open(page, "word-preview", late);
  await page.waitForFunction(() => window.__wordLoaded != null, null, { timeout: 300000, polling: 10 });
  const first = await ev(page, () => ({ load: window.__wordLoaded.viewer.isLoadComplete(), faces: window.RedBatWordTextStack.loadedFaceFiles() }));
  await page.waitForFunction(() => window.__wordPages > 0, null, { timeout: 300000 });
  await page.waitForFunction(() => window.RedBatWordTextStack.loadedFaceFiles().includes("georgiaz.ttf"), null, { timeout: 30000 }).catch(() => {});
  const faces = await ev(page, () => window.RedBatWordTextStack.loadedFaceFiles());
  check("preview: the later block's face is loaded when the document is read", (first.load === true || !first.faces.some((f) => /^georgia/.test(f))) && faces.includes("georgiaz.ttf"), { readAtFirstPaint: first.load, first: first.faces, faces });
  await page.waitForFunction(() => window.__wordLoaded.layoutComplete, null, { timeout: 60000 });
  // to the end of the document, then up until the picture's page is on screen
  let seen = 0;
  await ev(page, () => { const sc = document.querySelector('[data-testid="stage-scroll"]'); sc.scrollTop = sc.scrollHeight; });
  for (let i = 0; i < 12 && seen < 2000; i++) { await raf(page); await page.waitForTimeout(250); seen = await colourful(page, ".rb-file-viewer-canvas"); if (seen < 2000) await ev(page, () => { const sc = document.querySelector('[data-testid="stage-scroll"]'); sc.scrollTop -= 600; }); }
  await page.screenshot({ path: `${SHOTS}/word-late-picture-preview.png` });
  check("preview: the picture of a block that was read later is drawn", seen >= 2000, seen);
  check("preview: it was decoded into the viewer's image cache by the host", await ev(page, () => window.__wordLoaded.imageCache.size) >= 1, await ev(page, () => window.__wordLoaded.imageCache.size));
  await page.close();
}
// ---- a package whose later part is unreadable: the beginning is shown, with a notice; nothing is saved
{
  const broken = `${S}/broken-late.docx`;
  const page = await newPage(context);
  await open(page, "word-editor", broken);
  await page.waitForSelector('[data-testid="word-toolbar"]', { timeout: 300000 });
  await page.waitForSelector(".rb-office-notice", { timeout: 30000 }).catch(() => {});
  const notice = await page.locator(".rb-office-notice").innerText().catch(() => null);
  check("editor: a document whose rest is unreadable shows its beginning and says so", /rest of this document could not be read/.test(notice ?? ""), notice);
  await page.keyboard.type("X");
  await page.keyboard.press("Control+s");
  await page.waitForSelector('[data-testid="office-editor-error"]', { timeout: 30000 }).catch(() => {});
  check("editor: saving it fails and uploads nothing (no truncated file replaces the document)", page.uploads.length === 0 && /Not saved/.test(await page.locator('[data-testid="office-editor-error"]').innerText().catch(() => "")) && /Not saved|Unsaved/.test(await status(page)), await status(page));
  await page.close();
  const preview = await newPage(context);
  await open(preview, "chat-viewer", broken, { params: { mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" } });
  await preview.waitForSelector('[data-testid="word-preview-truncated"]', { timeout: 60000 }).catch(() => {});
  check("preview: the beginning is shown, with a notice", (await preview.locator('[data-testid="word-preview-truncated"]').count()) === 1 && (await preview.locator(".rb-file-viewer-canvas").count()) >= 1);
  await preview.close();
}
await browser.close();
process.exit(summary() ? 1 : 0);
