// A caret move or a selection change keeps the paint frame: the body commands,
// the layouts and their index are reused and only the overlay (selection
// highlight, caret, content-control frame) is made again. getRenderStats()
// .FrameBuilds does not move, and what is on the canvas is what a frame built
// from nothing paints for the same selection.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    const editor = Word.initializeCanvas(canvas);
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents-wild/docx4j-getting-started.docx')));
    const builds = () => Word.getRenderStats(editor).FrameBuilds;
    const pixels = () => canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    const same = (a, b) => { if (a.length !== b.length) return false; for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false; return true; };
    const input = Word.getInputElement(editor);
    const press = (key, shiftKey = false) => input.dispatchEvent(new KeyboardEvent('keydown', { key, shiftKey, bubbles: true, cancelable: true }));
    // What a frame built from nothing shows for the current selection.
    const cold = () => { const shown = pixels(); const before = builds(); Word.repaint(editor); return { shown, rebuilt: builds() - before, fresh: pixels() }; };
    const lines = Word.getLayoutLines(editor);
    const long = lines.filter((line) => line.EndOffset - line.StartOffset > 40 && line.Y > 100 && line.Y < 3000);
    const first = long[0], other = long.filter((line) => line.BlockId !== first.BlockId).pop();
    check('the page has lines of text to move in', long.length >= 2 && !!other, true);

    const run = (mode) => {
        Word.setSelection(editor, first.BlockId, first.StartOffset + 2, first.BlockId, first.StartOffset + 2);
        Word.insertText(editor, 'x');
        let before = builds();
        const rects = [];
        const times = [];
        for (let i = 0; i < 20; i++) {
            const at = first.StartOffset + 3 + i;
            const t0 = performance.now();
            Word.setSelection(editor, first.BlockId, at, first.BlockId, at);
            times.push(performance.now() - t0);
            rects.push(Word.getCaretRect(editor).X);
        }
        check(`${mode}: setSelection builds no frame`, builds() - before, 0);
        check(`${mode}: the caret rectangle follows`, rects.every((x, i) => i === 0 || x > rects[i - 1]), true);
        let state = cold();
        check(`${mode}: the canvas is what a new frame paints (caret moved through the API)`, [state.rebuilt, same(state.shown, state.fresh)], [1, true]);
        before = builds();
        for (let i = 0; i < 6; i++) press('ArrowRight');
        press('ArrowDown'); press('ArrowLeft'); press('Home'); press('End');
        check(`${mode}: arrow keys, Home and End build no frame`, builds() - before, 0);
        state = cold();
        check(`${mode}: the canvas is what a new frame paints (caret moved by keys)`, same(state.shown, state.fresh), true);
        before = builds();
        for (let i = 0; i < 12; i++) press('ArrowRight', true);
        press('ArrowDown', true);
        check(`${mode}: extending the selection builds no frame`, [builds() - before, Word.getSelection(editor).IsCollapsed], [0, false]);
        state = cold();
        check(`${mode}: the canvas is what a new frame paints (a selection in one paragraph)`, same(state.shown, state.fresh), true);
        before = builds();
        Word.setSelection(editor, first.BlockId, first.StartOffset + 4, other.BlockId, other.StartOffset + 9);
        check(`${mode}: a selection over several paragraphs builds no frame`, builds() - before, 0);
        state = cold();
        check(`${mode}: the canvas is what a new frame paints (a selection over paragraphs)`, same(state.shown, state.fresh), true);
        before = builds();
        Word.setSelection(editor, other.BlockId, other.StartOffset + 1, other.BlockId, other.StartOffset + 1);
        state = cold();
        check(`${mode}: collapsing it removes the highlight`, [builds() - before - state.rebuilt, same(state.shown, state.fresh)], [0, true]);
        // An edit still builds one frame, and the caret moves on from it without another.
        before = builds();
        Word.insertText(editor, 'y');
        const afterEdit = builds() - before;
        press('ArrowLeft'); press('ArrowRight');
        check(`${mode}: an edit builds one frame, the moves after it none`, [afterEdit, builds() - before], [1, 1]);
        times.sort((a, b) => a - b);
        return times[10];
    };
    const whole = run('whole canvas');
    Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 1000, height: 800 });
    const through = run('viewport');
    note(`a caret move: median ${whole.toFixed(2)} ms on a whole canvas, ${through.toFixed(2)} ms through a viewport (${Word.getPageCount(editor)} pages)`);
    check('a caret move is far below a frame build (under 12 ms)', whole < 12 && through < 12, true);
    Word.dispose(editor);
} catch (e) { fail(e); }
finish();
