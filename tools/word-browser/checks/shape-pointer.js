// What a host no longer does for the pointer and shapes: a click on the
// frame of a shape that is not selected - behind the text, without text -
// selects it and a drag from there moves it (one undo step); the listeners
// hear that there is a step to undo when a drag of a shape ends (they heard
// of the drag as it went, before it was one); getShapeHandles and
// hitTestShapeHandle put a turned shape's handles where the turn paints them,
// and getShapeRects gives its Rotation.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const editor = Word.initializeCanvas(canvas);
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await bytesOf('/repo/tests/repros/word/shapes-behind-text.docx')));
    const shapes = () => Array.from(Word.getShapeRects(editor));
    const all = shapes();
    // A turned oval with text, a rectangle behind the text, a triangle in front without text.
    const oval = all.find((s) => s.Text === 'in the oval');
    const rect = all.find((s) => s.BehindText);
    const triangle = all.find((s) => !s.BehindText && s.Text === '');
    note(`shapes: ${JSON.stringify(all.map((s) => [s.Id, Math.round(s.X), Math.round(s.Y), Math.round(s.Width), Math.round(s.Height), s.Rotation, s.BehindText]))}`);
    const of = (shape) => shapes().find((s) => s.Id === shape.Id);
    const place = (shape) => { const s = of(shape); return s ? [s.X, s.Y, s.Width, s.Height].map((n) => Math.round(n * 10) / 10) : null; };
    const depth = () => { let n = 0; while (Word.canUndo(editor) && n < 200) { Word.undo(editor); n++; } for (let i = 0; i < n; i++) Word.redo(editor); return n; };
    const notices = [];
    Word.onChange(editor, (notice) => notices.push({ CanUndo: notice.CanUndo, DocumentChanged: notice.DocumentChanged }));
    const state = {};
    window.step = (name) => {
        try { state[name] = { story: Word.getCaretContext(editor).Story, selected: Word.getSelectedShape(editor), rect: place(rect), triangle: place(triangle), oval: place(oval), depth: depth(), notices: notices.length, last: notices.at(-1) ?? null }; }
        catch (e) { state[name] = { error: String(e) }; }
    };
    window.finish = () => {
        try {
            const s = state, r = [rect.X, rect.Y, rect.Width, rect.Height].map((n) => Math.round(n * 10) / 10);
            check('nothing is selected at first', s.before.selected, null);
            check('a click on the frame of the shape behind the text selects it; the caret stays in the body', [s.edgeClicked.selected === rect.Id, s.edgeClicked.story, s.edgeClicked.rect], [true, 'body', r]);
            check('a click on the text beside it lets go of it', [s.textClicked.selected, s.textClicked.story], [null, 'body']);
            check('a drag from its frame, not selected before: moved by 30, 20, one undo step, selected', [s.dragged.rect, s.dragged.depth - s.textClicked.depth, s.dragged.selected === rect.Id], [[r[0] + 30, r[1] + 20, r[2], r[3]], 1, true]);
            check('the listeners heard of the step when the drag ended (the notices during it had nothing to undo)', [s.dragged.last?.CanUndo, s.dragged.notices > s.textClicked.notices], [true, true]);
            const t = [triangle.X, triangle.Y, triangle.Width, triangle.Height].map((n) => Math.round(n * 10) / 10);
            check('the triangle (no text) is selected by its frame', [s.triangle.selected === triangle.Id, s.triangle.triangle], [true, t]);
            // A turned shape: its handles are where they are painted.
            const angle = oval.Rotation * Math.PI / 180, cx = oval.X + oval.Width / 2, cy = oval.Y + oval.Height / 2;
            const turn = (x, y) => [cx + (x - cx) * Math.cos(angle) - (y - cy) * Math.sin(angle), cy + (x - cx) * Math.sin(angle) + (y - cy) * Math.cos(angle)];
            const se = turn(oval.X + oval.Width, oval.Y + oval.Height);
            const handle = s.handles.find((h) => h.Name === 'se');
            check('getShapeRects gives the oval its turn', Math.abs(oval.Rotation) > 1, true);
            check('the south-east handle of the turned oval is at its turned corner', [Math.abs(handle.X + handle.Width / 2 - se[0]) < 0.5, Math.abs(handle.Y + handle.Height / 2 - se[1]) < 0.5], [true, true]);
            check('hitTestShapeHandle finds it there', s.hit, 'se');
            const o = [oval.X, oval.Y, oval.Width, oval.Height].map((n) => Math.round(n * 10) / 10);
            check('a drag of that handle resizes the oval', [s.ovalResized.oval[2] > o[2] + 5, s.ovalResized.oval[3] > o[3] + 5, s.ovalResized.depth - s.triangle.depth], [true, true, 1]);
        } catch (e) { fail(e); }
        finish();
    };
    const rx = rect.X + rect.Width / 4;
    window.drive = [
        ['eval', 'window.step("before")'],
        // the top edge of the rectangle, a quarter along (no handle there)
        ['click', rx, rect.Y], ['eval', 'window.step("edgeClicked")'],
        // on the text of the line beside it (inside it, a click would take it: it is selected and behind the text)
        ['click', rect.X - 40, rect.Y + 12], ['eval', 'window.step("textClicked")'],
        ['down', rx, rect.Y + rect.Height], ['move', rx + 15, rect.Y + rect.Height + 10], ['move', rx + 30, rect.Y + rect.Height + 20], ['up'], ['eval', 'window.step("dragged")'],
        ['click', triangle.X + triangle.Width / 4, triangle.Y], ['eval', 'window.step("triangle")'],
        ['eval', `window.selectOval()`],
        ['then', 'window.ovalDrag()'], ['eval', 'window.step("ovalResized")'],
    ];
    window.selectOval = () => {
        Word.selectShape(editor, oval.Id);
        state.handles = Array.from(Word.getShapeHandles(editor, oval.Id));
        const h = state.handles.find((entry) => entry.Name === 'se');
        state.hit = Word.hitTestShapeHandle(editor, oval.Id, h.X + h.Width / 2, h.Y + h.Height / 2);
    };
    window.ovalDrag = () => {
        const h = state.handles.find((entry) => entry.Name === 'se');
        const x = h.X + h.Width / 2, y = h.Y + h.Height / 2;
        return [['down', x, y], ['move', x + 10, y + 10], ['move', x + 20, y + 20], ['up']];
    };
} catch (e) { fail(e); finish(); }
