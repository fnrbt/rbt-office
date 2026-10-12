// A shape that is selected in the editor itself: selectShape / getSelectedShape,
// its frame and eight handles painted by the editor (turned as the shape is),
// a click into a box's text selects the box, Escape goes back to the body with
// the box still selected and then lets go of it, the arrow keys move it,
// Delete takes it out, a drag of its frame moves it and a drag of a handle
// resizes it - one undo step each, also by a handle of the left or the top
// edge (resizeShapeBy does the same for a host: it was a moveShape and a
// resizeShape, two steps).
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const context = canvas.getContext('2d', { willReadFrequently: true });
    const editor = Word.initializeCanvas(canvas);
    Word.setParagraphs(editor, Array.from({ length: 12 }, (_, i) => `Paragraph ${i + 1} of the body, with a few words in it.`));
    const first = Word.getLayoutLines(editor)[0].BlockId;
    Word.setSelection(editor, first, 0, first, 0);
    Word.insertTextBox(editor, 'boxed');
    const rectOf = () => { const shape = Word.getShapeRects(editor)[0]; return shape ? [shape.X, shape.Y, shape.Width, shape.Height].map((n) => Math.round(n * 10) / 10) : null; };
    const id = Word.getShapeRects(editor)[0].Id;
    const start = rectOf();
    const [x0, y0, w0, h0] = start;
    note(`the box: ${JSON.stringify(start)}`);
    const has = (name) => typeof Word[name] === 'function';
    const selected = () => (has('getSelectedShape') ? Word.getSelectedShape(editor) : 'no getSelectedShape');
    // The blue of the frame and the handles in a square around a point of the canvas.
    const blue = (x, y, reach = 6) => { const data = context.getImageData(Math.round(x - reach), Math.round(y - reach), reach * 2, reach * 2).data; let n = 0; for (let i = 0; i < data.length; i += 4) if (data[i + 2] > 170 && data[i] < 200 && data[i + 2] - data[i] > 40) n++; return n; };
    const depth = () => { let n = 0; while (Word.canUndo(editor) && n < 200) { Word.undo(editor); n++; } for (let i = 0; i < n; i++) Word.redo(editor); return n; };
    const notices = [];
    Word.onChange(editor, (notice) => { if (notice.ShapeSelectionChanged) notices.push(selected()); });
    const state = {};
    window.step = (name) => {
        try {
            const story = Word.getCaretContext(editor).Story;
            // (The undo steps are counted in the body: undoing all of them leaves the text of a box.)
            const steps = story === 'body' ? depth() : -1;
            Word.repaint(editor);
            const rect = rectOf() ?? start;
            state[name] = { story, selected: selected(), rect: rectOf(), depth: steps, corner: blue(rect[0], rect[1]), text: Word.getText(editor) };
        }
        catch (e) { state[name] = { error: String(e) }; }
    };
    window.finish = async () => {
        try {
            const s = state;
            check('nothing is selected at first; no handle is painted', [s.before.selected, s.before.corner], [null, 0]);
            check('a click into the box: its text is edited, the box is selected, its handles are painted', [s.clicked.story, s.clicked.selected === id, s.clicked.corner > 8], ['textbox', true, true]);
            check('Escape: back to the body, the box still selected and its handles painted', [s.escaped.story, s.escaped.selected === id, s.escaped.corner > 8], ['body', true, true]);
            check('ArrowRight three times, ArrowDown twice: the box moves by pixels', [s.nudged.rect, s.nudged.depth - s.escaped.depth, s.nudged.text === s.escaped.text], [[x0 + 3, y0 + 2, w0, h0], 5, true]);
            check('Shift+ArrowLeft: by ten', s.nudgedTen.rect, [x0 - 7, y0 + 2, w0, h0]);
            check('a drag of its frame moves it: one undo step', [s.dragged.rect, s.dragged.depth - s.nudgedTen.depth, s.dragged.selected === id], [[x0 - 7 + 40, y0 + 2 + 30, w0, h0], 1, true]);
            check('a drag of the south-east handle resizes it: one undo step', [s.resized.rect, s.resized.depth - s.dragged.depth], [[x0 + 33, y0 + 32, w0 + 20, h0 + 10], 1]);
            check('a drag of the north-west handle: the other corner stays, one undo step', [s.resizedNw.rect, s.resizedNw.depth - s.resized.depth], [[x0 + 33 - 10, y0 + 32 - 6, w0 + 30, h0 + 16], 1]);
            check('undone once: as before that drag', s.undone.rect, [x0 + 33, y0 + 32, w0 + 20, h0 + 10]);
            check('a click on the body lets go of it', [s.released.selected, s.released.corner, s.released.story], [null, 0, 'body']);
            check('selectShape: selected, its handles painted, the caret stays in the body', [s.api.selected === id, s.api.corner > 8, s.api.story], [true, true, 'body']);
            check('Delete takes it out, its text with it: one undo step', [s.deleted.rect, s.deleted.selected, s.deleted.depth - s.api.depth], [null, null, 1]);
            check('undone: it is back', s.back.rect, [x0 + 33, y0 + 32, w0 + 20, h0 + 10]);
            check('the listeners were told of every change of the selected shape: the click, the click elsewhere, selectShape, Delete', notices.map((entry) => entry === id), [true, false, true, false]);

            // resizeShapeBy: what a drag of a handle does, for a host.
            const before = rectOf(), depthBefore = depth();
            check('resizeShapeBy by the west handle: one undo step', [has('resizeShapeBy') && Word.resizeShapeBy(editor, id, 'w', -15, 4), rectOf(), depth() - depthBefore], [true, [before[0] - 15, before[1], before[2] + 15, before[3]], 1]);
            check('by no handle, by nothing: false', [has('resizeShapeBy') && Word.resizeShapeBy(editor, id, 'x', 5, 5), has('resizeShapeBy') && Word.resizeShapeBy(editor, id, 'e', 0, 0)], [false, false]);
            check('selectShape of no shape', [has('selectShape') && Word.selectShape(editor, 'none'), has('selectShape') && Word.selectShape(editor, null), selected()], [false, true, null]);

            // A turned shape: the frame and the handles are turned with it.
            const turned = Word.initializeCanvas(canvas);
            Word.importFromDocxPackage(turned, await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents-wild/docx4j-drawingml-wps.docx')));
            const shapes = Word.getShapeRects(turned);
            // (The one that is turned: its text box, 101.4 x 86 px, by 13.96 degrees.)
            const shape = shapes.find((entry) => Math.abs(entry.Width - 101.4) < 1 && Math.abs(entry.Height - 86) < 1) ?? shapes[0];
            Word.repaint(turned);
            const angle = 13.96 * Math.PI / 180, cx = shape.X + shape.Width / 2, cy = shape.Y + shape.Height / 2;
            const turn = (x, y) => [cx + (x - cx) * Math.cos(angle) - (y - cy) * Math.sin(angle), cy + (x - cx) * Math.sin(angle) + (y - cy) * Math.cos(angle)];
            const corners = [[shape.X, shape.Y], [shape.X + shape.Width, shape.Y], [shape.X + shape.Width, shape.Y + shape.Height], [shape.X, shape.Y + shape.Height]];
            // The pixels of a square around a point, and how many of them differ between two pictures.
            const grab = ([x, y]) => context.getImageData(Math.round(x - 5), Math.round(y - 5), 10, 10).data;
            const differ = (one, two) => { let n = 0; for (let i = 0; i < one.length; i += 4) if (Math.abs(one[i] - two[i]) + Math.abs(one[i + 1] - two[i + 1]) + Math.abs(one[i + 2] - two[i + 2]) > 24) n++; return n; };
            const turnedCorners = corners.map(([x, y]) => turn(x, y));
            const plainTurned = turnedCorners.map(grab), plainUpright = corners.map(grab);
            check('selectShape of the turned shape', has('selectShape') && Word.selectShape(turned, shape.Id), true);
            Word.repaint(turned);
            const atTurned = turnedCorners.map((point, i) => differ(plainTurned[i], grab(point))), atUpright = corners.map((point, i) => differ(plainUpright[i], grab(point)));
            note(`the turned shape ${JSON.stringify([shape.X, shape.Y, shape.Width, shape.Height].map(Math.round))}: pixels painted at its turned corners ${JSON.stringify(atTurned)}, at the corners of its rectangle before the turn ${JSON.stringify(atUpright)}`);
            check('a handle is painted at each corner where the turn puts it, not where the corner is before the turn', [atTurned.map((n) => n >= 20), atTurned.reduce((sum, n) => sum + n, 0) > 2 * atUpright.reduce((sum, n) => sum + n, 0)], [[true, true, true, true], true]);
        } catch (e) { fail(e); }
        finish();
    };
    const cx = x0 + w0 / 2, cy = y0 + h0 / 2;
    // (Where the box is after each step: the drags go by where it is then.)
    const a = [x0 - 7, y0 + 2], b = [a[0] + 40, a[1] + 30];
    window.drive = [
        ['eval', 'window.step("before")'],
        ['click', cx, cy], ['eval', 'window.step("clicked")'],
        ['press', 'Escape'], ['eval', 'window.step("escaped")'],
        ['press', 'ArrowRight'], ['press', 'ArrowRight'], ['press', 'ArrowRight'], ['press', 'ArrowDown'], ['press', 'ArrowDown'], ['eval', 'window.step("nudged")'],
        ['press', 'Shift+ArrowLeft'], ['eval', 'window.step("nudgedTen")'],
        // Its frame, a quarter along the top edge (no handle there).
        ['down', a[0] + w0 / 4, a[1]], ['move', a[0] + w0 / 4 + 40, a[1] + 30], ['up'], ['eval', 'window.step("dragged")'],
        ['down', b[0] + w0, b[1] + h0], ['move', b[0] + w0 + 20, b[1] + h0 + 10], ['up'], ['eval', 'window.step("resized")'],
        ['down', b[0], b[1]], ['move', b[0] - 10, b[1] - 6], ['up'], ['eval', 'window.step("resizedNw")'],
        ['eval', 'window.undoOnce()'], ['eval', 'window.step("undone")'],
        ['click', 600, 700], ['eval', 'window.step("released")'],
        ['eval', 'window.selectIt()'], ['eval', 'window.step("api")'],
        ['press', 'Delete'], ['eval', 'window.step("deleted")'],
        ['eval', 'window.undoOnce()'], ['eval', 'window.step("back")'],
    ];
    window.undoOnce = () => { Word.undo(editor); };
    window.selectIt = () => { if (has('selectShape')) Word.selectShape(editor, id); };
} catch (e) { fail(e); finish(); }
