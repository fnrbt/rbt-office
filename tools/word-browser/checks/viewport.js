// Viewport rendering: the canvas is only as large as the visible area, the
// engine paints the band that intersects it at scale x devicePixelRatio, and
// hit-testing, the caret and the input proxy follow the transform. Pixels and
// hit tests are compared with the whole-canvas mode; a 200+ page document is
// painted page by page with renderPage and with the per-page full render that
// viewers had to use before.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';

const makeCanvas = (style) => { const c = document.createElement('canvas'); c.setAttribute('style', style); document.body.appendChild(c); return c; };
// Fraction of pixels that differ visibly between a canvas and a region of another.
const differing = (canvas, reference, offsetX, offsetY) => {
    const w = canvas.width, h = canvas.height;
    const a = canvas.getContext('2d').getImageData(0, 0, w, h).data;
    const b = reference.getContext('2d').getImageData(offsetX, offsetY, w, h).data;
    let bad = 0;
    for (let i = 0; i < a.length; i += 4)
        if (Math.abs(a[i] - b[i]) > 48 || Math.abs(a[i + 1] - b[i + 1]) > 48 || Math.abs(a[i + 2] - b[i + 2]) > 48) bad++;
    return bad / (w * h);
};
// Share of dark (text) and white (paper) pixels: a painted page has both.
const inkAndPaper = (canvas) => {
    const d = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    let ink = 0, paper = 0;
    for (let i = 0; i < d.length; i += 4) { if (d[i] < 90 && d[i + 1] < 90 && d[i + 2] < 90) ink++; else if (d[i] > 250 && d[i + 1] > 250 && d[i + 2] > 250) paper++; }
    return { ink: ink / (d.length / 4), paper: paper / (d.length / 4) };
};

try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const Renderer = await import('/word/Renderer.js');
    const { Editor__get_Doc } = await import('/word/Word.js');
    document.getElementById('c').style.display = 'none';
    const pkg = await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents-features/test-multi-headers.docx'));

    // ---- reference: whole-canvas mode, scaled by the host as hosts do today ----
    const refCanvas = makeCanvas('position:absolute;left:0;top:3000px');
    const ref = Word.initializeCanvas(refCanvas);
    Word.importFromDocxPackage(ref, pkg);
    const paintReference = (k) => {
        const size = Word.getDocumentSize(ref);
        refCanvas.width = Math.round(size.Width * k); refCanvas.height = Math.round(size.Height * k);
        refCanvas.getContext('2d').setTransform(k, 0, 0, k, 0, 0);
        Word.repaint(ref);
    };
    check('whole-canvas mode reports no viewport', Word.getViewport(ref), null);

    // ---- the editor under test, in viewport mode ----
    const canvas = makeCanvas('position:fixed;left:20px;top:30px');
    const editor = Word.initializeCanvas(canvas);
    Word.importFromDocxPackage(editor, pkg);
    const docSize = Word.getDocumentSize(editor);
    note(`document ${docSize.Width} x ${docSize.Height}, ${docSize.PageCount} pages`);
    check('document size covers every page frame', [docSize.PageCount, docSize.Height], (() => { const f = Word.getPageFrames(editor); return [f.length, f[f.length - 1].Top + f[f.length - 1].Height]; })());

    const configs = [
        { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 700, height: 500 },
        { scale: 1, devicePixelRatio: 2, scrollTop: 310, scrollLeft: 40, width: 600, height: 420 },
        { scale: 1.5, devicePixelRatio: 1, scrollTop: 900, scrollLeft: 120, width: 640, height: 480 },
        { scale: 0.75, devicePixelRatio: 2, scrollTop: 1000, scrollLeft: 0, width: 500, height: 600 },
        { scale: 2, devicePixelRatio: 1.5, scrollTop: 1712, scrollLeft: 334, width: 520, height: 380 },
        { scale: 1.25, devicePixelRatio: 1, scrollTop: 260, scrollLeft: 90, width: 560, height: 400, originX: 30, originY: 50 },
    ];
    for (const vp of configs) {
        const label = `scale ${vp.scale} dpr ${vp.devicePixelRatio} scroll ${vp.scrollLeft},${vp.scrollTop}` + (vp.originX ? ` origin ${vp.originX},${vp.originY}` : '');
        const size = Word.setViewport(editor, vp);
        const dpr = vp.devicePixelRatio, ox = vp.originX ?? 0, oy = vp.originY ?? 0;
        check(`${label}: canvas is the viewport, not the document`, [canvas.width, canvas.height, canvas.style.width, canvas.style.height],
            [Math.round(vp.width * dpr), Math.round(vp.height * dpr), vp.width + 'px', vp.height + 'px']);
        check(`${label}: scaled document size`, [Math.round(size.ScaledWidth), Math.round(size.ScaledHeight)], [Math.round(docSize.Width * vp.scale), Math.round(docSize.Height * vp.scale)]);
        paintReference(vp.scale * dpr);
        const fraction = differing(canvas, refCanvas, Math.round((vp.scrollLeft - ox) * dpr), Math.round((vp.scrollTop - oy) * dpr));
        check(`${label}: pixels match the whole-canvas rendering (${(fraction * 100).toFixed(3)}% differ)`, fraction < 0.004, true);
        const content = inkAndPaper(canvas);
        check(`${label}: shows text on paper`, content.ink > 0.002 && content.paper > 0.3, true);
        // Hit-testing through the transform equals hit-testing the layout point in whole-canvas mode.
        let mismatches = 0, hits = 0;
        for (let x = 15; x < vp.width; x += 67) for (let y = 11; y < vp.height; y += 41) {
            const got = Word.hitTestViewport(editor, x, y);
            // The same layout point in whole-canvas mode: the offset from the reference editor, the block
            // from this one (block ids are per import).
            const lx = (x + vp.scrollLeft - ox) / vp.scale, ly = (y + vp.scrollTop - oy) / vp.scale;
            const want = Word.hitTestPosition(ref, lx, ly), own = Word.hitTestPosition(editor, lx, ly);
            if (want) hits++;
            if (JSON.stringify(got && [got.BlockId, got.Offset]) !== JSON.stringify(want && [own.BlockId, want.Offset])) mismatches++;
        }
        check(`${label}: hit tests agree with whole-canvas mode (${hits} positions)`, [mismatches, hits > 20], [0, true]);
        const p = Word.viewportToDocument(editor, 100, 80), back = Word.documentToViewport(editor, p.X, p.Y);
        check(`${label}: viewport <-> document mapping`, [Math.round(p.X * 100), Math.round(p.Y * 100), Math.round(back.X), Math.round(back.Y)],
            [Math.round((100 + vp.scrollLeft - ox) / vp.scale * 100), Math.round((80 + vp.scrollTop - oy) / vp.scale * 100), 100, 80]);
    }

    // ---- an edit repaints only its band, and the result still matches ----
    const editVp = { scale: 1.5, devicePixelRatio: 1, scrollTop: 150, scrollLeft: 60, width: 640, height: 480 };
    Word.setViewport(editor, editVp);
    paintReference(1.5);
    const line = Word.getLayoutLines(ref).find((l) => l.Y > 180 && l.Y < 380 && (l.Text || '').trim().length > 8);
    for (const target of [editor, ref]) {
        const hit = Word.hitTestPosition(target, line.X + 30, line.Y + line.Height / 2);
        Word.setSelection(target, hit.BlockId, hit.Offset, hit.BlockId, hit.Offset);
        Word.insertText(target, 'viewport edit ');
    }
    check('after an edit the viewport still matches whole-canvas mode', differing(canvas, refCanvas, 60, 150) < 0.004, true);
    Word.toggleBold(editor); Word.selectAll(editor); Word.selectAll(ref);
    check('after a selection change too', differing(canvas, refCanvas, 60, 150) < 0.004, true);

    // ---- real pointer input and the caret / input proxy through the transform ----
    const clickVp = { scale: 1.5, devicePixelRatio: 2, scrollTop: 150, scrollLeft: 60, width: 640, height: 480 };
    Word.setViewport(editor, clickVp);
    const refLine = Word.getLayoutLines(editor).find((l) => l.Y > 200 && l.Y < 380 && (l.Text || '').trim().length > 8);
    const clickAt = Word.documentToViewport(editor, refLine.X + 40, refLine.Y + refLine.Height / 2);
    const expected = Word.hitTestPosition(editor, refLine.X + 40, refLine.Y + refLine.Height / 2);
    window.afterClick = () => {
        const sel = Word.getSelection(editor);
        check('a real click lands on the position under the pointer', [sel.FocusBlockId, sel.FocusOffset], [expected.BlockId, expected.Offset]);
        const caret = Word.getCaretRect(editor);
        const proxy = Word.getInputElement(editor).getBoundingClientRect(), box = canvas.getBoundingClientRect();
        check('the input proxy sits at the caret on screen', [Math.round(proxy.left - box.left), Math.round(proxy.top - box.top), Math.round(proxy.height)],
            [Math.round(caret.ViewportX), Math.round(caret.ViewportY), Math.round(caret.ViewportHeight)]);
        check('caret rect in both spaces', [Math.round(caret.ViewportX), Math.round(caret.ViewportY)],
            [Math.round(caret.X * 1.5 - 60), Math.round(caret.Y * 1.5 - 150)]);
        window.textBeforeTyping = Word.getText(editor);
    };
    window.afterTyping = () => {
        const sel = Word.getSelection(editor);
        check('typing goes in at the clicked position', [Word.getText(editor).length, sel.FocusBlockId, sel.FocusOffset], [window.textBeforeTyping.length + 2, expected.BlockId, expected.Offset + 2]);
    };

    window.finish = () => {
        try {
            // ---- back to whole-canvas mode ----
            Word.setViewport(editor, null);
            check('null goes back to whole-canvas mode', Word.getViewport(editor), null);
            const size = Word.getDocumentSize(editor);
            canvas.width = size.Width; canvas.height = Math.min(size.Height, 3000); canvas.style.width = canvas.width + 'px'; canvas.style.height = canvas.height + 'px';
            Word.repaint(editor);
            const one = inkAndPaper(canvas);
            check('whole-canvas mode paints again', one.ink > 0.002 && one.paper > 0.3, true);

            // ---- a long document: canvas stays small, deep scroll paints, pages render fast ----
            const big = Word.initializeCanvas(makeCanvas('position:fixed;left:700px;top:0'));
            const bigCanvas = Word.getInputElement(big).previousSibling;
            const words = 'The quick brown fox jumps over the lazy dog and keeps running through the long document ';
            Word.setParagraphs(big, Array.from({ length: 5000 }, (_, i) => `Paragraph ${i + 1}. ` + words.repeat(2 + (i % 3))));
            Word.setViewport(big, { scale: 1, devicePixelRatio: 2, scrollTop: 0, scrollLeft: 0, width: 600, height: 700 });
            // The deep-page comparison and whole-document benchmark require all page frames.
            Word.completeLayout(big);
            const bigSize = Word.getDocumentSize(big);
            note(`long document: ${bigSize.PageCount} pages, ${Math.round(bigSize.Height)} layout px tall`);
            check('the long document has more than 200 pages', bigSize.PageCount > 200, true);
            check('its canvas is the viewport (1200 x 1400 backing pixels), far below canvas size limits', [bigCanvas.width, bigCanvas.height], [1200, 1400]);
            const frames = Word.getPageFrames(big);
            const deep = frames[180];
            let t0 = performance.now();
            Word.setViewport(big, { scale: 1, devicePixelRatio: 2, scrollTop: deep.Top + 40, scrollLeft: 0, width: 600, height: 700 });
            const scrollMs = performance.now() - t0;
            const deepContent = inkAndPaper(bigCanvas);
            check('page 181 is painted after scrolling to it', deepContent.ink > 0.005 && deepContent.paper > 0.5, true);
            const deepHit = Word.hitTestViewport(big, 150, 200), deepWant = Word.hitTestPosition(big, 150, deep.Top + 40 + 200);
            check('hit test on page 181', [deepHit.BlockId, deepHit.Offset], [deepWant.BlockId, deepWant.Offset]);
            t0 = performance.now();
            for (let i = 0; i < 20; i++) Word.setViewport(big, { scale: 1, devicePixelRatio: 2, scrollTop: deep.Top + 40 + i * 53, scrollLeft: 0, width: 600, height: 700 });
            const perScroll = (performance.now() - t0) / 20;
            note(`scroll repaint on the long document: first ${scrollMs.toFixed(1)} ms, then ${perScroll.toFixed(1)} ms per scroll step`);

            // The viewport at page 181 shows what renderPage paints for page 181.
            const pageCanvas = makeCanvas('position:absolute;left:0;top:9000px');
            Word.setViewport(big, { scale: 1, devicePixelRatio: 1, scrollTop: deep.Top, scrollLeft: 0, width: 600, height: 700 });
            Word.clearSelection(big);
            check('renderPage paints page 181', Word.renderPage(big, pageCanvas, 180, 1), true);
            check('renderPage sizes the canvas to the page', [pageCanvas.width, pageCanvas.height], [Math.round(deep.Width), Math.round(deep.Height)]);
            check('renderPage and the viewport show the same page', differing(bigCanvas, pageCanvas, 0, 0) < 0.004, true);
            check('renderPage refuses a page that does not exist', Word.renderPage(big, pageCanvas, frames.length, 1), false);

            // Page by page: before (a full render through a translated context per page, as viewers did) and after (renderPage).
            const doc = Editor__get_Doc(big), scale = 0.25, sample = 40;
            const before = (pageIndex) => {
                const page = frames[pageIndex];
                pageCanvas.width = Math.round(page.Width * scale); pageCanvas.height = Math.round(page.Height * scale);
                const ctx = Renderer.createContext(pageCanvas);
                ctx.LayoutState = big.Handler.RenderContext.LayoutState;
                ctx.Ctx.setTransform(scale, 0, 0, scale, 0, -page.Top * scale);
                Renderer.render(ctx, doc);
            };
            before(0);
            t0 = performance.now();
            for (let i = 0; i < sample; i++) before(Math.floor(i * frames.length / sample));
            const beforeMs = (performance.now() - t0) / sample * frames.length;
            Word.renderPage(big, pageCanvas, 0, scale);
            t0 = performance.now();
            for (let i = 0; i < frames.length; i++) Word.renderPage(big, pageCanvas, i, scale);
            const afterMs = performance.now() - t0;
            note(`rendering ${frames.length} pages one by one at scale ${scale}: before ${Math.round(beforeMs)} ms (extrapolated from ${sample} pages), after ${Math.round(afterMs)} ms`);
            check('renderPage is at least ten times faster than a full render per page', afterMs * 10 < beforeMs, true);

            // The viewer object does the same without an editor.
            const viewer = Word.createViewer(makeCanvas('display:none'));
            viewer.load(pkg);
            check('viewer page count', viewer.getPageCount(), docSize.PageCount);
            check('viewer renders a page', [viewer.renderPage(pageCanvas, 0, 0.5), pageCanvas.width, pageCanvas.height], [true, Math.round(viewer.getPageFrames()[0].Width * 0.5), Math.round(viewer.getPageFrames()[0].Height * 0.5)]);
            const viewerInk = inkAndPaper(pageCanvas);
            check('viewer page has text on paper', viewerInk.ink > 0.002 && viewerInk.paper > 0.5, true);
            Word.dispose(big); Word.dispose(editor); Word.dispose(ref);
        } catch (e) { fail(e); }
        finish();
    };
    window.drive = [['click', 20 + clickAt.X, 30 + clickAt.Y], ['eval', 'window.afterClick()'], ['type', 'ab'], ['eval', 'window.afterTyping()']];
} catch (e) { fail(e); finish(); }
