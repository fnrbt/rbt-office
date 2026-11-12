// The placeholder of an object the editor does not draw (a chart) says what it
// is in a label of its own size in the middle of the frame - it used to be
// stretched to the frame's width. And the frame of a content control is an
// editing aid: drawn only around the control the caret is in, never in a
// viewer or on a page rendered for preview.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
// Bounding box of the pixels of a colour (within a tolerance) in a rectangle of a canvas.
const inkBox = (canvas, rect, match) => {
    const { x, y, w, h } = rect;
    const d = canvas.getContext('2d').getImageData(x, y, w, h).data;
    let left = w, right = -1, top = h, bottom = -1, count = 0;
    for (let row = 0; row < h; row++) for (let col = 0; col < w; col++) {
        const i = (row * w + col) * 4;
        if (match(d[i], d[i + 1], d[i + 2])) { count++; if (col < left) left = col; if (col > right) right = col; if (row < top) top = row; if (row > bottom) bottom = row; }
    }
    return count === 0 ? null : { left: x + left, right: x + right, top: y + top, bottom: y + bottom, width: right - left + 1, height: bottom - top + 1, count };
};
const labelInk = (r, g, b) => r < 150 && g < 160 && b < 170 && Math.abs(r - 0x4B) < 70;   // #4B5563 and its antialiasing, not the #9CA3AF frame
// #2563EB, the frame's colour (the caret is #0066FF, links are darker).
const controlBlue = (r, g, b) => Math.abs(r - 0x25) < 14 && Math.abs(g - 0x63) < 14 && Math.abs(b - 0xEB) < 14;
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    canvas.style.display = 'none';
    const make = (style) => { const c = document.createElement('canvas'); c.setAttribute('style', style); document.body.appendChild(c); return c; };

    // ---- the chart placeholder ----
    const pkg = await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents-wild/docx4j-chart.docx'));
    const editorCanvas = make('position:fixed;left:0;top:0');
    const editor = Word.initializeCanvas(editorCanvas);
    Word.importFromDocxPackage(editor, pkg);
    Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 900, height: 700 });
    const segment = Word.getLayoutSegments(editor).find((s) => (s.Kind ?? s.kind) === 'image' || s.Src || s.IsImage);
    const lines = Word.getLayoutLines(editor);
    note('segments ' + JSON.stringify(Word.getLayoutSegments(editor).slice(0, 3)).slice(0, 400));
    // The frame: where the layout put the object (the page's left margin plus the segment's x).
    const setup = Word.getPageSetup(editor);
    const marginLeft = setup.MarginLeft ?? setup.marginLeft ?? 96;
    const frame = { left: marginLeft + segment.X, top: segment.Y, width: segment.Width, height: segment.Height, right: marginLeft + segment.X + segment.Width - 1, bottom: segment.Y + segment.Height - 1 };
    const fill = inkBox(editorCanvas, { x: frame.left + 4, y: frame.top + 4, w: 20, h: 20 }, (r, g, b) => r === 0xF3 && g === 0xF4 && b === 0xF6);
    check('the chart has a placeholder frame', [fill !== null && fill.count === 400, frame.width > 150 && frame.height > 60], [true, true]);
    const naturalWidth = (() => { const ctx = document.createElement('canvas').getContext('2d'); ctx.font = '13px Calibri'; return ctx.measureText('Chart').width; })();
    const checkLabel = (label, target, box) => {
        const ink = inkBox(target, { x: box.left + 2, y: box.top + 2, w: box.width - 4, h: box.height - 4 }, labelInk);
        check(`${label}: the label is drawn`, ink !== null, true);
        if (!ink) return;
        note(`${label}: label ${ink.width} x ${ink.height} in a frame of ${box.width} x ${box.height} (natural width ${naturalWidth.toFixed(1)})`);
        check(`${label}: at its own width, not the frame's`, Math.abs(ink.width - naturalWidth) <= 4 && ink.width < box.width / 2, true);
        note(`${label}: label box ${JSON.stringify(ink)} frame ${JSON.stringify(box)}`);
        check(`${label}: centred in the frame`, [Math.abs((ink.left + ink.right) / 2 - (box.left + box.right) / 2) <= 3, Math.abs((ink.top + ink.bottom) / 2 - (box.top + box.bottom) / 2) <= 6], [true, true]);
        check(`${label}: of a normal height`, ink.height >= 8 && ink.height <= 16, true);
    };
    checkLabel('editor', editorCanvas, frame);
    // The same page in a viewer, painted with renderPage.
    const viewerCanvas = make('position:fixed;left:0;top:720px;width:10px;height:10px');
    const viewer = Word.createViewer(viewerCanvas);
    viewer.load(pkg);
    const pageCanvas = make('position:absolute;left:1000px;top:0');
    check('the viewer paints the page', viewer.renderPage(pageCanvas, 0, 1), true);
    checkLabel('preview', pageCanvas, frame);
    // A frame too narrow for the label shows what fits; nothing is drawn outside it.
    void lines;

    // ---- content-control frames ----
    const second = make('position:fixed;left:0;top:0');
    editorCanvas.style.display = 'none';
    const controls = Word.initializeCanvas(second);
    Word.setParagraphs(controls, ['Before the control.', 'Inside the control.', 'Also inside it.', 'After the control.']);
    const ids = Word.getLayoutLines(controls).map((line) => line.BlockId);
    Word.setSelection(controls, ids[1], 0, ids[2], 5);
    Word.setContentControl(controls, 'cc1', 'tag', 'A control');
    Word.setViewport(controls, { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 900, height: 400 });
    const blue = () => inkBox(second, { x: 0, y: 0, w: 900, h: 400 }, controlBlue);
    Word.setSelection(controls, ids[0], 3, ids[0], 3);
    check('caret outside the control: no frame', blue(), null);
    Word.setSelection(controls, ids[1], 4, ids[1], 4);
    const inside = blue();
    check('caret inside the control: its frame is drawn', inside !== null && inside.width > 300, true);
    const linesNow = Word.getLayoutLines(controls);
    const first = linesNow.find((l) => l.BlockId === ids[1]), last = linesNow.find((l) => l.BlockId === ids[2]), after = linesNow.find((l) => l.BlockId === ids[3]);
    check('around the paragraphs of the control, and only those', inside && [inside.top >= first.Y - 2, inside.bottom <= last.Y + last.Height + 2, inside.bottom < after.Y + 2], [true, true, true]);
    Word.setSelection(controls, ids[3], 2, ids[3], 2);
    check('caret after the control: no frame again', blue(), null);
    Word.setSelection(controls, ids[0], 2, ids[1], 3);
    check('a selection reaching into the control shows it', blue() !== null, true);
    Word.clearSelection(controls);
    check('no selection: no frame', blue(), null);
    // Never on a rendered page or in a viewer.
    Word.setSelection(controls, ids[1], 4, ids[1], 4);
    const preview = make('position:absolute;left:1000px;top:1200px');
    Word.renderPage(controls, preview, 0, 1);
    check('renderPage draws no control frame, even with the caret in the control', inkBox(preview, { x: 0, y: 0, w: preview.width, h: preview.height }, controlBlue), null);
    const controlViewer = Word.createViewer(make('position:fixed;left:0;top:720px;width:10px;height:10px'));
    controlViewer.load(Word.exportDocumentPackage(controls));
    const controlPage = make('position:absolute;left:1000px;top:2400px');
    controlViewer.renderPage(controlPage, 0, 1);
    check('a viewer draws no control frame', inkBox(controlPage, { x: 0, y: 0, w: controlPage.width, h: controlPage.height }, controlBlue), null);
    // Metadata-only edits can reuse text layout, but must refresh control frames.
    Word.removeContentControl(controls, 'cc1');
    check('removing the active control clears its frame without a text edit', blue(), null);
    Word.undo(controls);
    const restored = blue();
    check('undo restores the complete control frame',
        restored && [restored.top >= first.Y - 2, restored.bottom >= last.Y, restored.bottom < after.Y + 2],
        [true, true, true]);
    Word.setParagraphs(controls, ['Replacement document.']);
    const replacementId = Word.getLayoutLines(controls)[0].BlockId;
    Word.setSelection(controls, replacementId, 1, replacementId, 1);
    check('a replacement document does not inherit the old control frame', blue(), null);
    // The document's own TOC and page-number fields: no boxes around them at rest.
    const tocPkg = await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents/sample-formatted.docx'));
    const tocCanvas = make('position:fixed;left:0;top:0');
    second.style.display = 'none';
    const toc = Word.initializeCanvas(tocCanvas);
    Word.importFromDocxPackage(toc, tocPkg);
    Word.clearSelection(toc);
    const size = Word.setViewport(toc, { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 900, height: 1100 });
    let boxes = 0;
    for (let top = 0; top < Math.min(size.Height, 6000); top += 1000) {
        Word.setViewport(toc, { scale: 1, devicePixelRatio: 1, scrollTop: top, scrollLeft: 0, width: 900, height: 1100 });
        if (inkBox(tocCanvas, { x: 0, y: 0, w: 900, h: 1100 }, controlBlue)) boxes++;
    }
    check('a document with content controls shows no frames at rest', boxes, 0);
} catch (e) { fail(e); }
finish();
