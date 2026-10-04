// A viewer given the text measurers of an editor breaks lines and pages as
// that editor does (a preview used to paginate with browser fonts while the
// editor used the host's shaped text: 460 pages against 394), and renderPage
// paints text with the painter made for the canvas it paints on.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    document.getElementById('c').style.display = 'none';
    const make = (style) => { const c = document.createElement('canvas'); c.setAttribute('style', style); document.body.appendChild(c); return c; };
    // A deterministic measurer that is not the browser's: every character is 0.9 of the font size wide.
    const k = 0.9;
    const measureText = (text, style) => text.length * style.FontSize * k;
    const measureRange = (text, style, start, length) => Math.max(0, Math.min(length, text.length - start)) * style.FontSize * k;
    for (const file of ['/repo/tests/test-documents/sample-formatted.docx', '/repo/tests/test-documents-features/test-multi-headers.docx']) {
        const short = file.split('/').pop();
        const pkg = await OF.processors.forEditorPackage(await bytesOf(file));
        const editor = Word.initializeCanvas(make('position:fixed;left:0;top:0'));
        Word.importFromDocxPackage(editor, pkg);
        const browserPages = Word.getPageCount(editor);
        Word.setTextMeasurers(editor, measureText, measureRange);
        const editorFrames = Word.getPageFrames(editor).map((f) => [Math.round(f.Top), Math.round(f.Width), Math.round(f.Height)]);
        const viewer = Word.createViewer(make('position:fixed;left:0;top:0;width:10px;height:10px'));
        viewer.load(pkg);
        const plain = viewer.getPageCount();
        check(`${short}: a viewer without measurers paginates like an editor without`, plain, browserPages);
        viewer.setTextMeasurers(measureText, measureRange);
        const viewerFrames = viewer.getPageFrames().map((f) => [Math.round(f.Top), Math.round(f.Width), Math.round(f.Height)]);
        note(`${short}: ${browserPages} pages with browser text, ${editorFrames.length} with the measurer`);
        check(`${short}: the measurer changes the pagination (the check means something)`, editorFrames.length !== browserPages, true);
        check(`${short}: with the same measurers the viewer has the editor's pages`, viewerFrames, editorFrames);
        check(`${short}: getPageCount / getDocumentSize agree`, [viewer.getPageCount(), viewer.getDocumentSize().PageCount, Math.round(viewer.getDocumentSize().Height)], [editorFrames.length, editorFrames.length, Math.round(Word.getDocumentSize(editor).Height)]);
        // Line by line: the same text on the same page.
        const page = make('position:absolute;left:1200px;top:0');
        check(`${short}: every page renders`, editorFrames.map((_, index) => viewer.renderPage(page, index, 0.5)).every(Boolean), true);
        check(`${short}: there is no page after the last`, viewer.renderPage(page, editorFrames.length, 0.5), false);
        // Loading another document keeps the measurers.
        viewer.load(pkg);
        check(`${short}: the measurers stay when a document is loaded`, viewer.getPageCount(), editorFrames.length);
        viewer.clearTextMeasurer();
        check(`${short}: clearTextMeasurer goes back to browser text`, viewer.getPageCount(), browserPages);
        Word.dispose(editor);
    }
    // The text painter: made per canvas context, used by renderPage on its own canvas.
    const pkg = await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents-features/test-multi-headers.docx'));
    const viewer = Word.createViewer(make('position:fixed;left:0;top:0;width:10px;height:10px'));
    viewer.load(pkg);
    const made = [];
    viewer.setTextPathPainter((context) => {
        const entry = { context, calls: 0 };
        made.push(entry);
        return (text, style, x, baseline) => { entry.calls++; context.fillStyle = '#FF00FF'; context.fillRect(x, baseline - 6, Math.max(2, text.length * 3), 4); return true; };
    });
    const target = make('position:absolute;left:1200px;top:1200px');
    viewer.renderPage(target, 0, 1);
    const forTarget = made.find((entry) => entry.context === target.getContext('2d'));
    note('painters ' + JSON.stringify(made.map((entry) => entry.calls)));
    check('renderPage paints text with the painter made for its canvas', [!!forTarget, forTarget && forTarget.calls > 10], [true, true]);
    const data = target.getContext('2d').getImageData(0, 0, target.width, target.height).data;
    let magenta = 0;
    for (let i = 0; i < data.length; i += 4) if (data[i] === 255 && data[i + 1] === 0 && data[i + 2] === 255) magenta++;
    check('and what the painter draws is on the page', magenta > 200, true);
    // The same for an editor's renderPage.
    const editor = Word.initializeCanvas(make('position:fixed;left:0;top:0'));
    Word.importFromDocxPackage(editor, pkg);
    const editorMade = [];
    Word.setTextPathPainterFactory(editor, (context) => { const entry = { context, calls: 0 }; editorMade.push(entry); return () => { entry.calls++; return false; }; });
    const editorTarget = make('position:absolute;left:1200px;top:2400px');
    Word.renderPage(editor, editorTarget, 0, 1);
    check('an editor\'s renderPage uses the factory too', editorMade.some((entry) => entry.context === editorTarget.getContext('2d') && entry.calls > 10), true);
} catch (e) { fail(e); }
finish();
