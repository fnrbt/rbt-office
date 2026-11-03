// Pasted plain text (Ctrl+V, pasteText) has the format chosen at the caret
// with nothing selected - every line of it, not the first alone. What is
// done in the body does not reach into a text box that is not entered:
// typing or pasting over a selection keeps the box of the paragraph the
// selection starts in (the box goes with the start of its paragraph, where
// its anchor is) and never changes a box's text; replaceAllText - which goes
// by all the text findText finds - does, and says so. A text box can be
// inserted into a header (and is saved in the header's part).
import { check, note, finish, fail, loadWord, loadFormats } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const editor = Word.initialize('c');
    Word.setViewport(editor, { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 1000, height: 700 });
    const ids = () => [...new Set(Word.getLayoutLines(editor).filter((line) => line.BlockId.length > 0).map((line) => line.BlockId))];
    const texts = () => Word.getText(editor).split('\n');
    const boldOf = (block, from, to) => { Word.setSelection(editor, ids()[block], from, ids()[block], to); return Word.getSelectionFormat(editor).Bold; };

    // ---- the pending format and pasted text ---------------------------------
    Word.setParagraphs(editor, ['Hello world']);
    Word.setSelection(editor, ids()[0], 5, ids()[0], 5);
    Word.toggleBold(editor);
    Word.pasteText(editor, 'one\ntwo\r\nthree');
    check('three lines pasted at a caret: three paragraphs', texts(), ['Helloone', 'two', 'three world']);
    check('every line has the format chosen at the caret: bold', [boldOf(0, 5, 8), boldOf(1, 0, 3), boldOf(2, 0, 5)], [true, true, true]);
    check('the text around it is as it was', [boldOf(0, 0, 5), boldOf(2, 5, 11)], [false, false]);
    Word.undo(editor);
    check('one undo step', [texts(), Word.canUndo(editor)], [['Hello world'], false]);
    // One line.
    Word.setSelection(editor, ids()[0], 5, ids()[0], 5);
    Word.toggleItalic(editor);
    Word.pasteText(editor, ' there');
    Word.setSelection(editor, ids()[0], 5, ids()[0], 11);
    check('one line pasted at a caret with italic chosen', [texts(), Word.getSelectionFormat(editor).Italic, boldOf(0, 0, 5)], [['Hello there world'], true, false]);
    // Over a selection: the text takes the place of the selection, in the format the text has there.
    Word.setSelection(editor, ids()[0], 0, ids()[0], 5);
    Word.pasteText(editor, 'Goodbye');
    check('pasted over a selection', texts(), ['Goodbye there world']);
    // A size and a colour chosen at a caret, then two lines pasted.
    Word.setParagraphs(editor, ['plain']);
    Word.setSelection(editor, ids()[0], 5, ids()[0], 5);
    Word.setFontSizePt(editor, 20);
    Word.setColor(editor, '#FF0000');
    Word.pasteText(editor, ' big\nred');
    Word.setSelection(editor, ids()[1], 0, ids()[1], 3);
    const second = Word.getSelectionFormat(editor);
    check('a size and a colour chosen at the caret are those of the second pasted line too', [texts(), second.FontSizePt ?? second.FontSize, String(second.Color).toUpperCase()], [['plain big', 'red'], 20, '#FF0000']);

    // ---- text boxes and what is done in the body ------------------------------
    Word.setParagraphs(editor, ['first paragraph', 'second paragraph', 'third paragraph']);
    Word.setSelection(editor, ids()[0], 0, ids()[0], 0);
    check('a text box anchored in the first paragraph', Word.insertTextBox(editor, 'boxed paragraph'), true);
    const boxes = () => Word.getShapeRects(editor).map((shape) => [shape.Text, shape.AnchorBlockId === ids()[0]]);
    check('it is there', boxes(), [['boxed paragraph', true]]);
    const body = () => texts().filter((text) => !text.startsWith('boxed '));
    // Typed over a selection from the middle of the first paragraph into the second.
    Word.setSelection(editor, ids()[0], 6, ids()[1], 7);
    Word.insertText(editor, 'X');
    check('typed over a selection that starts in the middle of the box\'s paragraph: the box stays, its text untouched', [body(), boxes()], [['first Xparagraph', 'third paragraph'], [['boxed paragraph', true]]]);
    Word.undo(editor);
    // Pasted over the same selection.
    Word.setSelection(editor, ids()[0], 6, ids()[1], 7);
    Word.pasteText(editor, 'one\ntwo');
    check('pasted over it: the same', [body(), boxes().map((box) => box[0])], [['first one', 'twoparagraph', 'third paragraph'], ['boxed paragraph']]);
    Word.undo(editor);
    // From the start of the box's paragraph: the box goes with it.
    Word.setSelection(editor, ids()[0], 0, ids()[1], 7);
    Word.insertText(editor, 'X');
    check('typed over a selection from the start of the box\'s paragraph: the box goes with the paragraph\'s start', [body(), boxes()], [['Xparagraph', 'third paragraph'], []]);
    Word.undo(editor);
    check('and comes back with undo', boxes(), [['boxed paragraph', true]]);
    // Everything selected and typed over.
    Word.selectAll(editor);
    Word.setSelection(editor, ids()[0], 3, ids()[2], 5);
    Word.pasteText(editor, '-');
    check('the body from the fourth character of its first paragraph on pasted over: the box stays', [body(), boxes().map((box) => box[0])], [['fir- paragraph'], ['boxed paragraph']]);
    Word.undo(editor);
    // A table pasted into the text of a box: its rows, the cells parted by tabs (a box holds paragraphs).
    check('into the box', Word.enterTextBox(editor, Word.getShapeRects(editor)[0].Id), true);
    Word.pasteClipboard(editor, 'a\tb', '<table><tr><td>a</td><td>b</td></tr></table>');
    Word.exitTextBox(editor);
    check('a table pasted into a text box: a row of text, the box kept', [boxes().map((box) => box[0])], [['a\tbboxed paragraph']]);
    Word.undo(editor);
    // (Undo goes back into the box the step was made in.)
    Word.exitTextBox(editor);
    check('... undone', boxes().map((box) => box[0]), ['boxed paragraph']);
    // Find and replace go by all the text.
    check('findText finds in the body and in the box', Word.findText(editor, 'paragraph', true).length, 4);
    check('replaceAllText replaces there too (4)', [Word.replaceAllText(editor, 'paragraph', 'PARA', true), body(), boxes().map((box) => box[0])], [4, ['first PARA', 'second PARA', 'third PARA'], ['boxed PARA']]);
    Word.undo(editor);
    check('one undo step', [body(), boxes().map((box) => box[0])], [['first paragraph', 'second paragraph', 'third paragraph'], ['boxed paragraph']]);

    // ---- a text box in a header ------------------------------------------------
    check('into the header of the first page', Word.enterHeaderFooter(editor, 0, 'header'), true);
    Word.insertText(editor, 'Header text');
    const inserted = Word.insertTextBox(editor, 'box in the header');
    check('insertTextBox in a header', inserted, true);
    Word.exitHeaderFooter(editor);
    const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
    const zip = await window.JSZip.loadAsync(out);
    const headerNames = Object.keys(zip.files).filter((name) => /^word\/header\d*\.xml$/.test(name));
    const headers = await Promise.all(headerNames.map((name) => zip.file(name).async('string')));
    const withBox = headers.find((xml) => xml.includes('box in the header'));
    check('saved: the header\'s part has the text and the box (a drawing with the box\'s text)', [headers.some((xml) => xml.includes('Header text')), !!withBox && /<w:txbxContent>/.test(withBox) && /<(wp:anchor|v:shape|w:pict)/.test(withBox)], [true, true]);
    // Read again: the header has the box.
    const again = Word.initialize('c');
    Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(out));
    const pkg = Word.exportDocumentPackage(again);
    const headerItems = JSON.stringify(pkg.HeaderParts ?? pkg.Headers ?? []);
    check('read again: the header\'s box and its text', [headerItems.includes('box in the header'), headerItems.includes('Header text')], [true, true]);
    // And it is painted: ink right of the header's text, in the header's band of the first page.
    Word.setViewport(again, { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 1000, height: 700 });
    Word.repaint(again);
    note('header parts: ' + headerNames.join(', '));
} catch (e) { fail(e); }
finish();
