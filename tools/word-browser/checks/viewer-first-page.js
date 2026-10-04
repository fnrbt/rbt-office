// The viewer's first page: renderPage builds the paint frame once, and the
// paint commands of plain paragraphs are collected when a page that shows
// them is painted - the first page does not wait for those of the whole
// document. getRenderStats() (the editor's counters) says so. What a page
// shows does not depend on which pages were painted before it.
import { check, note, finish, fail, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const createViewer = (await import('/word/WordViewer.js')).default;
    const editorPackage = await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents-wild/docx4j-getting-started.docx'));
    const viewer = createViewer(document.createElement('canvas'));
    viewer.load(editorPackage);
    const frames = viewer.getPageFrames();
    const before = viewer.getRenderStats();
    check('getRenderStats on the viewer', [typeof before.FrameBuilds, typeof before.CollectedOnPaint, before.Blocks > 500], ['number', 'number', true]);
    const target = document.createElement('canvas');
    const pixels = () => target.getContext('2d').getImageData(0, 0, target.width, target.height).data;
    const same = (a, b) => { if (a.length !== b.length) return false; for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false; return true; };
    let t0 = performance.now();
    viewer.renderPage(target, 0, 1);
    const first = performance.now() - t0;
    const firstPixels = pixels();
    const afterFirst = viewer.getRenderStats();
    check('the first page builds one frame', afterFirst.FrameBuilds - before.FrameBuilds, 1);
    check('and collects the commands of the paragraphs it shows, not of the document', [afterFirst.CollectedOnPaint > 0, afterFirst.CollectedOnPaint < afterFirst.Blocks / 8], [true, true]);
    check('the page is painted (ink on it)', (() => { let ink = 0; for (let i = 0; i < firstPixels.length; i += 4) if (firstPixels[i] < 100) ink++; return ink > 500; })(), true);
    // A page in the middle, then all of them.
    viewer.renderPage(target, 40, 1);
    const middle = pixels();
    const times = [];
    for (let i = 0; i < frames.length; i++) { const t = performance.now(); viewer.renderPage(target, i, 1); times.push(performance.now() - t); }
    const afterAll = viewer.getRenderStats();
    check('every page painted: still one frame, every paragraph collected once', [afterAll.FrameBuilds - before.FrameBuilds, afterAll.CollectedOnPaint > afterFirst.CollectedOnPaint, afterAll.CollectedOnPaint <= afterAll.Blocks], [1, true, true]);
    viewer.renderPage(target, 0, 1);
    check('the first page painted again is the same pixels', same(pixels(), firstPixels), true);
    viewer.renderPage(target, 40, 1);
    check('and so is a page in the middle', same(pixels(), middle), true);
    check('painting them again collects nothing more', viewer.getRenderStats().CollectedOnPaint, afterAll.CollectedOnPaint);
    times.sort((a, b) => a - b);
    note(`${frames.length} pages: the first renderPage ${first.toFixed(1)} ms (${afterFirst.CollectedOnPaint} of ${afterFirst.Blocks} blocks collected), a page after it median ${times[Math.floor(times.length / 2)].toFixed(1)} ms`);
    check('the first page takes less than 150 ms at 75 pages', first < 150, true);
} catch (e) { fail(e); }
finish();
