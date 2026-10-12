// The text box of a header reached with the pointer, and as a shape:
// resume-template.docx has its name and contact lines in a text box of the
// first page's header. A double click on it from the body edits the box's
// text (it entered the header and left the caret in the header's own
// paragraph), a click on it while the header is edited does (it did
// nothing), a click elsewhere on the header goes back to the header, a
// double click on the body to the body. moveShape / resizeShape /
// deleteShape / getShapeHandles take the shapes of headers and footers as
// they take the body's, and a moved one is saved in its part.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const bytes = await bytesOf('/repo/tests/test-documents-features/resume-template.docx');
    const canvas = document.getElementById('c');
    const editor = Word.initializeCanvas(canvas);
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
    Word.repaint(editor);
    const boxOf = (target = editor) => Word.getShapeRects(target).find((shape) => shape.Story === 'header');
    const box = boxOf();
    const frame = Word.getPageFrames(editor)[0];
    const rect = canvas.getBoundingClientRect();
    note(`the header's box: ${JSON.stringify([Math.round(box.X), Math.round(box.Y), Math.round(box.Width), Math.round(box.Height)])}, the page's top margin ${Math.round(frame.MarginTop)}`);
    const story = () => { const context = Word.getCaretContext(editor); return [context.Story, context.TextBoxId === box.Id]; };
    const state = {};
    window.step = (name) => { state[name] = { story: story(), text: boxOf() ? boxOf().Text : null, caret: Word.getCaretRect(editor), all: Word.getText(editor) }; };
    const inBox = (caret) => !!caret && caret.X >= box.X - 2 && caret.X <= box.X + box.Width + 2 && caret.Y >= box.Y - 2 && caret.Y <= box.Y + box.Height + 2;
    // A point in the box (in its text), one on the header beside or below the box, one on the body.
    const px = rect.left + box.X + Math.min(60, box.Width / 2), py = rect.top + box.Y + Math.min(14, box.Height / 2);
    const beside = [[frame.Width - 12, 8], [12, 8], [frame.Width - 12, frame.MarginTop - 6], [12, frame.MarginTop - 6]]
        .find(([x, y]) => !(x >= box.X && x <= box.X + box.Width && y >= box.Y && y <= box.Y + box.Height)) ?? [frame.Width - 12, 8];
    const bodyPoint = [rect.left + frame.Width / 2, rect.top + frame.MarginTop + 300];
    check('in the body at first', story(), ['body', false]);

    window.finish = async () => {
        try {
            check('a double click on the header\'s text box from the body: the box\'s text is edited, the caret in it', [state.dbl.story, inBox(state.dbl.caret)], [['textbox', true], true]);
            check('typed there: the box\'s text', [state.typed.text.includes('Q'), state.typed.text.length - box.Text.length], [true, 1]);
            check('Escape: back to the body', state.escaped.story, ['body', false]);
            check('the header entered (enterHeaderFooter): the header\'s own text', state.header.story, ['header', false]);
            check('a click on the box while the header is edited: the box\'s text, the caret in it', [state.clicked.story, inBox(state.clicked.caret)], [['textbox', true], true]);
            check('typed there', [state.typed2.text.includes('W'), state.typed2.text.length - box.Text.length], [true, 2]);
            check('a double click in the box that is edited selects a word (the box stays)', [state.word.story, Word.getSelection(editor) !== null], [['textbox', true], true]);
            check('a click elsewhere on the header: back to the header', state.beside.story, ['header', false]);
            check('a double click on the body: the body', state.body.story, ['body', false]);

            // ---- as a shape ----
            const before = boxOf();
            const handles = Word.getShapeHandles(editor, box.Id);
            check('getShapeHandles of a header\'s shape', [handles.length, handles.map((handle) => handle.Name).join(' '), Math.round(handles[0].X + handles[0].Width / 2 - before.X), Math.round(handles[0].Y + handles[0].Height / 2 - before.Y)], [8, 'nw n ne e se s sw w', 0, 0]);
            check('hitTestShapeHandle', [Word.hitTestShapeHandle(editor, box.Id, before.X + before.Width / 2, before.Y + before.Height / 2), Word.hitTestShapeHandle(editor, box.Id, before.X + before.Width, before.Y + before.Height)], ['move', 'se']);
            check('moveShape of a header\'s shape', Word.moveShape(editor, box.Id, 12, 7), true);
            const moved = boxOf();
            check('it stands 12 px right and 7 px down', [Math.round((moved.X - before.X) * 10) / 10, Math.round((moved.Y - before.Y) * 10) / 10, Math.round(moved.Width - before.Width)], [12, 7, 0]);
            const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
            const zip = await window.JSZip.loadAsync(out), source = await window.JSZip.loadAsync(bytes);
            check('saved: the body is the text that was read', await zip.file('word/document.xml').async('string') === await source.file('word/document.xml').async('string'), true);
            const again = Word.initializeCanvas(document.createElement('canvas'));
            Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(out));
            const reread = boxOf(again);
            check('read again: the box is where it was moved to, with its text', [Math.round((reread.X - before.X) * 10) / 10, Math.round((reread.Y - before.Y) * 10) / 10, reread.Text === moved.Text], [12, 7, true]);
            Word.undo(editor);
            check('one undo: where it was', [Math.round(boxOf().X - before.X), Math.round(boxOf().Y - before.Y), Word.getCaretContext(editor).Story], [0, 0, 'body']);
            check('resizeShape of a header\'s shape', [Word.resizeShape(editor, box.Id, before.Width + 20, before.Height + 10), Math.round(boxOf().Width - before.Width), Math.round(boxOf().Height - before.Height)], [true, 20, 10]);
            Word.undo(editor);
            check('deleteShape of a header\'s shape', [Word.deleteShape(editor, box.Id), boxOf() === undefined], [true, true]);
            Word.undo(editor);
            check('undone: the box is back', [boxOf() !== undefined, boxOf().Text === moved.Text], [true, true]);
        } catch (e) { fail(e); }
        finish();
    };
    window.drive = [
        ['click', bodyPoint[0], bodyPoint[1]],
        ['dblclick', px, py], ['eval', 'window.step("dbl")'],
        ['type', 'Q'], ['eval', 'window.step("typed")'],
        ['press', 'Escape'], ['eval', 'window.step("escaped")'],
        ['eval', 'window.enterHeader()'], ['eval', 'window.step("header")'],
        ['click', px, py], ['eval', 'window.step("clicked")'],
        ['type', 'W'], ['eval', 'window.step("typed2")'],
        ['dblclick', px, py], ['eval', 'window.step("word")'],
        ['click', rect.left + beside[0], rect.top + beside[1]], ['eval', 'window.step("beside")'],
        ['dblclick', bodyPoint[0], bodyPoint[1]], ['eval', 'window.step("body")'],
    ];
    window.enterHeader = () => Word.enterHeaderFooter(editor, 0, 'header');
} catch (e) { fail(e); finish(); }
