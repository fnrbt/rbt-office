// A text path painter and text measurers given as plain JavaScript functions:
// the painted frame is built when one is installed or cleared and at no other
// time (the renderer used to compare the painter by reference, and saw a
// different function object at every paint: the whole document's commands
// were collected again for each scroll frame).
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    const editor = Word.initializeCanvas(canvas);
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents-wild/docx4j-getting-started.docx')));
    const pages = Word.getPageCount(editor);
    const size = Word.getDocumentSize(editor);
    const builds = () => Word.getRenderStats(editor).FrameBuilds;
    const view = (scrollTop) => Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop, scrollLeft: 0, width: 1000, height: 800 });
    view(0);
    // Scroll through the document: one frame per step.
    const scroll = (steps) => {
        const before = builds(), t0 = performance.now();
        for (let i = 0; i < steps; i++) view((i * 37) % Math.max(1, size.Height - 800));
        return { built: builds() - before, perFrame: (performance.now() - t0) / steps };
    };
    const plain = scroll(60);
    check('canvas text: scrolling builds no frame', plain.built, 0);

    // A painter that is a plain function of five arguments.
    let calls = 0;
    const painter = (text, style, x, baseline, width) => { calls++; const c = canvas.getContext('2d'); c.fillStyle = '#C000C0'; c.fillRect(x, baseline - 5, Math.max(1, width), 3); return true; };
    let before = builds();
    Word.setTextPathPainter(editor, painter);
    check('installing a painter builds the frame once', [builds() - before, Word.getRenderStats(editor).PainterVersion], [1, 1]);
    check('the painter paints the text', calls > 20, true);
    calls = 0;
    const withPainter = scroll(60);
    check('with a plain-function painter: scrolling builds no frame', withPainter.built, 0);
    check('and every scroll frame is painted with it', calls > 60, true);
    note(`${pages} pages; a scroll frame: ${plain.perFrame.toFixed(2)} ms with canvas text, ${withPainter.perFrame.toFixed(2)} ms with the painter (it draws a bar per run)`);
    before = builds();
    // The same painter handed over again is a new installation (the host said so), not a comparison.
    Word.setTextPathPainter(editor, painter);
    check('installing again counts as a new painter', [builds() - before, Word.getRenderStats(editor).PainterVersion], [1, 2]);

    // The factory form: a painter per canvas context.
    const made = [];
    before = builds();
    Word.setTextPathPainterFactory(editor, (context) => { const entry = { context, calls: 0 }; made.push(entry); return (text, style, x, baseline, width) => { entry.calls++; return false; }; });
    check('a factory: one frame, the painter for the editor\'s canvas asked for once', [builds() - before, made.length, made[0].context === canvas.getContext('2d')], [1, 1, true]);
    const withFactory = scroll(60);
    check('with a factory painter: scrolling builds no frame, and no new painter is made', [withFactory.built, made.length], [0, 1]);
    check('a painter that declines leaves the text to the canvas', made[0].calls > 60, true);
    before = builds();
    Word.clearTextPathPainter(editor);
    check('clearing the painter builds the frame once', builds() - before, 1);
    check('and then scrolling builds none', scroll(30).built, 0);

    // Measurers as plain functions.
    const measureText = (text, style) => text.length * style.FontSize * 0.9;
    const measureRange = (text, style, start, length) => Math.max(0, Math.min(length, text.length - start)) * style.FontSize * 0.9;
    before = builds();
    Word.setTextMeasurers(editor, measureText, measureRange);
    const measured = Word.getPageCount(editor);
    check('plain-function measurers lay the document out again', [builds() - before >= 1, measured !== pages, Word.getRenderStats(editor).MeasurerVersion], [true, true, 1]);
    view(0);
    check('with them: scrolling builds no frame', scroll(60).built, 0);
    Word.setTextPathPainter(editor, painter);
    check('measurers and a painter together: scrolling builds no frame', scroll(60).built, 0);
    // Typing still shows (the frame follows the document, not the painter).
    before = builds();
    Word.insertText(editor, 'x');
    check('an edit builds one frame', builds() - before, 1);
    Word.clearTextMeasurer(editor);
    check('clearTextMeasurer: canvas text and the browser\'s pages again', Word.getPageCount(editor), pages);
    Word.dispose(editor);
} catch (e) { fail(e); }
finish();
