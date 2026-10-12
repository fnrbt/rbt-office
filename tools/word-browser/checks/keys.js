// The editor's own keys beyond typing and moving: Ctrl+Enter is a page break
// (it was Enter: a new paragraph on the same page), Ctrl+Shift+Enter a column
// break, Shift+Enter a line break in the paragraph, Ctrl+Backspace /
// Ctrl+Delete take the word before / after the caret, Tab outside a list and
// a table is a tab character (in a list it moves the item a level down, in a
// table to the next cell; Ctrl+Tab is a tab character there). Each is one
// undo step. insertColumnBreak / insertLineBreak do what the keys do.
import { check, note, finish, fail, loadWord, loadFormats } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    const fresh = () => {
        const editor = Word.initializeCanvas(canvas);
        Word.setParagraphs(editor, ['alpha beta gamma', 'second paragraph here', 'third']);
        return editor;
    };
    const idsOf = (editor) => [...new Set(Word.getLayoutLines(editor).map((line) => line.BlockId).filter((id) => id.length > 0))];
    const key = (editor, name, mods = {}) => Word.getInputElement(editor).dispatchEvent(new KeyboardEvent('keydown', { key: name, code: name, ctrlKey: !!mods.ctrl, shiftKey: !!mods.shift, bubbles: true, cancelable: true }));
    const caret = (editor, block, offset) => { const id = idsOf(editor)[block]; Word.setSelection(editor, id, offset, id, offset); };
    const state = (editor) => [Word.getText(editor), idsOf(editor).length, Word.getPageCount(editor)];
    const pageOf = (editor, block) => { const id = idsOf(editor)[block]; const line = Word.getLayoutLines(editor).find((l) => l.BlockId === id); const frames = Word.getPageFrames(editor); let page = -1; frames.forEach((frame, index) => { if (line.Y >= (frame.Top ?? frame.Y)) page = index; }); return page; };
    const start = ['alpha beta gamma\nsecond paragraph here\nthird', 3, 1];

    // Ctrl+Enter: a page break.
    {
        const editor = fresh();
        check('three paragraphs on one page', state(editor), start);
        caret(editor, 0, 5);
        key(editor, 'Enter', { ctrl: true });
        check('Ctrl+Enter: the paragraph is split at the caret and what follows starts a new page', [state(editor), pageOf(editor, 0), pageOf(editor, 1), Word.getSelection(editor).FocusOffset, Word.getSelection(editor).FocusBlockId === idsOf(editor)[1]],
            [['alpha\n beta gamma\nsecond paragraph here\nthird', 4, 2], 0, 1, 0, true]);
        const saved = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const again = Word.initializeCanvas(document.createElement('canvas'));
        Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(saved));
        check('saved and read again: two pages', state(again), ['alpha\n beta gamma\nsecond paragraph here\nthird', 4, 2]);
        Word.undo(editor);
        check('one undo step', state(editor), start);
        // The API does the same.
        caret(editor, 0, 5);
        check('insertPageBreak is what the key does', [Word.insertPageBreak(editor), state(editor)], [true, ['alpha\n beta gamma\nsecond paragraph here\nthird', 4, 2]]);
    }
    // Ctrl+Shift+Enter: a column break (the next page where the section has one column).
    {
        const editor = fresh();
        caret(editor, 1, 6);
        key(editor, 'Enter', { ctrl: true, shift: true });
        check('Ctrl+Shift+Enter: split, and what follows is on the next page of a section of one column', [state(editor), pageOf(editor, 1), pageOf(editor, 2)], [['alpha beta gamma\nsecond\n paragraph here\nthird', 4, 2], 0, 1]);
        const saved = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const xml = await (await window.JSZip.loadAsync(saved)).file('word/document.xml').async('string');
        check('saved as a column break', [(xml.match(/<w:br w:type="column"\/>/g) ?? []).length, (xml.match(/<w:br w:type="page"\/>/g) ?? []).length], [1, 0]);
        Word.undo(editor);
        check('one undo step', state(editor), start);
        caret(editor, 1, 6);
        check('insertColumnBreak is what the key does', [typeof Word.insertColumnBreak === 'function' && Word.insertColumnBreak(editor), state(editor)], [true, ['alpha beta gamma\nsecond\n paragraph here\nthird', 4, 2]]);
    }
    // Shift+Enter: a line break.
    {
        const editor = fresh();
        caret(editor, 0, 5);
        const linesBefore = Word.getLayoutLines(editor).filter((line) => line.BlockId === idsOf(editor)[0]).length;
        key(editor, 'Enter', { shift: true });
        const lines = Word.getLayoutLines(editor).filter((line) => line.BlockId === idsOf(editor)[0]).length;
        check('Shift+Enter: the same paragraph, on one line more', [idsOf(editor).length, Word.getPageCount(editor), lines - linesBefore, Word.getText(editor).startsWith('alpha\n beta gamma\nsecond'), Word.getSelection(editor).FocusOffset], [3, 1, 1, true, 6]);
        const saved = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const xml = await (await window.JSZip.loadAsync(saved)).file('word/document.xml').async('string');
        check('saved as a break in the paragraph', [(xml.match(/<w:br\/>/g) ?? []).length, (xml.match(/<w:p[ >]/g) ?? []).length], [1, 3]);
        Word.undo(editor);
        check('one undo step', state(editor), start);
        caret(editor, 0, 5);
        if (typeof Word.insertLineBreak === 'function') Word.insertLineBreak(editor);
        check('insertLineBreak is what the key does', [idsOf(editor).length, Word.getText(editor).startsWith('alpha\n beta gamma\nsecond')], [3, true]);
    }
    // Ctrl+Backspace / Ctrl+Delete: a word.
    {
        const editor = fresh();
        caret(editor, 0, 10);
        key(editor, 'Backspace', { ctrl: true });
        check('Ctrl+Backspace: the word before the caret', [Word.getText(editor).split('\n')[0], Word.getSelection(editor).FocusOffset], ['alpha  gamma', 6]);
        Word.undo(editor);
        check('one undo step', state(editor), start);
        caret(editor, 0, 6);
        key(editor, 'Delete', { ctrl: true });
        check('Ctrl+Delete: the word after the caret, with the space after it', [Word.getText(editor).split('\n')[0], Word.getSelection(editor).FocusOffset], ['alpha gamma', 6]);
        Word.undo(editor);
        check('one undo step', state(editor), start);
        caret(editor, 0, 8);
        key(editor, 'Backspace', { ctrl: true });
        check('in a word: its part before the caret', Word.getText(editor).split('\n')[0], 'alpha ta gamma');
        Word.undo(editor);
        caret(editor, 1, 0);
        key(editor, 'Backspace', { ctrl: true });
        check('at the start of a paragraph: joined with the one before', state(editor), ['alpha beta gammasecond paragraph here\nthird', 2, 1]);
        Word.undo(editor);
        const ids = idsOf(editor);
        Word.setSelection(editor, ids[0], 2, ids[0], 4);
        key(editor, 'Delete', { ctrl: true });
        check('with a selection: the selection', Word.getText(editor).split('\n')[0], 'ala beta gamma');
        Word.undo(editor);
        check('undone', state(editor), start);
    }
    // Tab.
    {
        const editor = fresh();
        caret(editor, 0, 5);
        key(editor, 'Tab');
        check('Tab in a paragraph: a tab character', [Word.getText(editor).split('\n')[0], Word.getSelection(editor).FocusOffset], ['alpha\t beta gamma', 6]);
        key(editor, 'Tab', { shift: true });
        check('Shift+Tab there: nothing', Word.getText(editor).split('\n')[0], 'alpha\t beta gamma');
        Word.undo(editor);
        check('one undo step', state(editor), start);
        caret(editor, 1, 0);
        Word.toggleBulletList(editor);
        const before = Word.getSelectionFormat(editor).Paragraph.ListLevel;
        key(editor, 'Tab');
        const after = Word.getSelectionFormat(editor).Paragraph.ListLevel;
        check('Tab in a list item: a level down, no character', [Word.getText(editor), after - before], [start[0], 1]);
        key(editor, 'Tab', { shift: true });
        check('Shift+Tab: a level up', [Word.getText(editor), Word.getSelectionFormat(editor).Paragraph.ListLevel - before], [start[0], 0]);
        // In a table: the next cell; Ctrl+Tab a tab character.
        const table = Word.initializeCanvas(canvas);
        Word.setParagraphs(table, ['one']);
        Word.insertTable(table, 2, 2);
        Word.insertText(table, 'ab');
        const cell = Word.getSelection(table).FocusBlockId;
        key(table, 'Tab', { ctrl: true });
        check('Ctrl+Tab in a table cell: a tab character, the caret stays in the cell', [Word.getSelection(table).FocusBlockId === cell, Word.getText(table).includes('ab\t')], [true, true]);
        key(table, 'Tab');
        check('Tab in a table cell: the next cell', [Word.getSelection(table).FocusBlockId !== cell, Word.getCaretContext(table).Column], [true, 1]);
    }
    // Where a page cannot break Ctrl+Enter does nothing (it is not Enter).
    {
        const editor = fresh();
        caret(editor, 0, 0);
        Word.insertTextBox(editor, 'boxed text');
        const box = Word.getShapeRects(editor)[0];
        Word.enterTextBox(editor, box.Id);
        const before = Word.getText(editor);
        key(editor, 'Enter', { ctrl: true });
        key(editor, 'Enter', { ctrl: true, shift: true });
        check('Ctrl+Enter in a text box: nothing', [Word.getText(editor) === before, Word.getShapeRects(editor)[0].Text], [true, 'boxed text']);
        key(editor, 'Enter');
        check('Enter there: a paragraph', Word.getText(editor) === before, false);
    }
} catch (e) { fail(e); }
finish();
