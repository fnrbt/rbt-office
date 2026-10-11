// Moving through the body never goes into a text box: the text of a box is
// another story, entered by a click in it (or enterTextBox) and left by
// Escape. Ctrl+End is the end of the body's last paragraph, Ctrl+Home the
// start of its first, the arrow keys go from a paragraph to the next one of
// the body, Ctrl+A selects the body. A position in a box given to
// setSelection enters the box, so that what is found in a box (findText) can
// be shown.
// The document: docx4j-drawingml-wps.docx, whose last blocks are text boxes
// (anchored in the third paragraph).
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const editor = Word.initializeCanvas(canvas);
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents-wild/docx4j-drawingml-wps.docx')));
    Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 1000, height: 900 });
    const boxes = Word.getShapeRects(editor);
    const boxText = () => Word.getShapeRects(editor).map((box) => box.Text).join('|');
    const boxTextBefore = boxText();
    const bodyLines = Word.getLayoutLines(editor);
    const bodyIds = [...new Set(bodyLines.map((line) => line.BlockId))];
    // (The paragraphs of the boxes: where a word of a box's text is found. The body does not have it.)
    const word = boxTextBefore.split(/\s+/).find((part) => part.length >= 3);
    const boxIds = [...new Set(Word.findText(editor, word, true).map((match) => match.BlockId))];
    const inBody = (id) => typeof id === 'string' && !boxIds.includes(id);
    const title = bodyLines.find((line) => (line.Text ?? '').startsWith('Milk glass'));
    note(`boxes: ${boxes.length} (${JSON.stringify(boxTextBefore.slice(0, 60))}); body paragraphs: ${bodyIds.length}`);
    check('the document has a text box with text, which the body\'s lines do not have', [boxes.length >= 1, boxIds.length, bodyLines.some((line) => (line.Text ?? '').includes(word))], [true, 1, false]);
    const state = {};
    const where = () => { const s = Word.getSelection(editor), c = Word.getCaretContext(editor), r = Word.getCaretRect(editor); return { block: s.FocusBlockId, offset: s.FocusOffset, anchor: s.AnchorBlockId, anchorOffset: s.AnchorOffset, story: c.Story, caret: r !== null && r !== undefined }; };
    window.steps = {
        end: () => { state.end = where(); },
        typed: () => { state.typed = [boxText() === boxTextBefore, Word.getText(editor).includes('yy'), where()]; },
        home: () => { state.home = where(); },
        right: () => { state.right = where(); },
        all: () => { state.all = where(); },
        down: () => { state.down = where(); },
    };
    const frame = Word.getPageFrames(editor)[0];
    window.drive = [
        ['click', Math.round((frame.X ?? 0) + title.X + 20), Math.round(title.Y + title.Height / 2)],
        ['press', 'Control+End'], ['eval', 'window.steps.end()'],
        ['type', 'yy'], ['eval', 'window.steps.typed()'],
        // From the end of the body on: Right and Down stay there.
        ['press', 'ArrowRight'], ['press', 'ArrowRight'], ['eval', 'window.steps.right()'],
        ['press', 'ArrowDown'], ['press', 'PageDown'], ['eval', 'window.steps.down()'],
        ['press', 'Control+Home'], ['eval', 'window.steps.home()'],
        ['press', 'Control+a'], ['eval', 'window.steps.all()'],
    ];
    window.finish = () => {
        try {
            note('end ' + JSON.stringify(state.end));
            check('Ctrl+End: the end of the last paragraph of the body, with a caret, in the body', [inBody(state.end.block), state.end.story, state.end.caret], [true, 'body', true]);
            check('typing there goes into the body, not into a box', [state.typed[0], state.typed[1], state.typed[2].block, state.typed[2].offset, state.typed[2].caret], [true, true, state.end.block, state.end.offset + 2, true]);
            check('Right at the end of the body stays there', [state.right.block, state.right.offset, state.right.caret], [state.end.block, state.typed[2].offset, true]);
            check('Down and PageDown too', [state.down.block, state.down.story, state.down.caret], [state.end.block, 'body', true]);
            check('Ctrl+Home: the start of the first paragraph', [state.home.block, state.home.offset, state.home.caret], [bodyIds[0], 0, true]);
            check('Ctrl+A: from the start of the body to its end', [state.all.anchor, state.all.anchorOffset, state.all.block, state.all.story], [bodyIds[0], 0, state.end.block, 'body']);
            // Formatting everything leaves the text of the boxes alone (it is not selected).
            Word.selectAll(editor);
            const all = Word.getSelection(editor);
            check('selectAll: the body', [all.AnchorBlockId, all.FocusBlockId], [bodyIds[0], state.end.block]);
            // What is found in a box is shown in the box.
            const matches = Word.findText(editor, word, true);
            const bodyText = Word.getText(editor);
            note(`looking for ${JSON.stringify(word)}: ${matches.length} match(es)`);
            const inBox = matches.find((match) => boxIds.includes(match.BlockId));
            if (inBox !== undefined) {
                Word.setSelection(editor, inBox.BlockId, inBox.Start, inBox.BlockId, inBox.End);
                const c = Word.getCaretContext(editor), s = Word.getSelection(editor);
                check('a match in a box, selected: the box is entered and the match is selected in it', [c.Story, s.AnchorBlockId, s.AnchorOffset, s.FocusOffset, Word.getCaretRect(editor) !== null], ['textbox', inBox.BlockId, inBox.Start, inBox.End, true]);
                check('and left again', [Word.exitTextBox(editor), Word.getCaretContext(editor).Story], [true, 'body']);
            } else check('the text of a box is found', false, true);
        } catch (e) { fail(e); }
        finish();
    };
} catch (e) { fail(e); finish(); }
