// The text boxes of headers and footers: getShapeRects lists them (Story
// "header" / "footer", the first page that shows the header: PageIndex), and
// enterTextBox edits one in place - on that page; what is typed is the
// header's (the box of every page that shows it) and is saved in its part.
// resume-template.docx: its name and contact lines are a text box of the
// first page's header; getShapeRects used to return nothing for it.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const bytes = await bytesOf('/repo/tests/test-documents-features/resume-template.docx');
    const canvas = document.getElementById('c');
    const editor = Word.initializeCanvas(canvas);
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
    const shapes = Word.getShapeRects(editor);
    const inHeader = shapes.filter((shape) => shape.Story === 'header' || shape.Story === 'footer');
    note('shapes: ' + JSON.stringify(shapes.map((shape) => [shape.Story, shape.PageIndex, Math.round(shape.X), Math.round(shape.Y), Math.round(shape.Width), Math.round(shape.Height), shape.Text.slice(0, 30)])));
    check('the header\'s text box is listed, on the first page', [inHeader.length >= 1, inHeader[0]?.Story, inHeader[0]?.PageIndex, (inHeader[0]?.Text ?? '').length > 0], [true, 'header', 0, true]);
    const box = inHeader[0];
    const frame = Word.getPageFrames(editor)[0];
    check('where the page paints it: inside the first page', [box.X >= 0 && box.X + box.Width <= frame.Width + 1, box.Y >= (frame.Top ?? frame.Y) && box.Y + box.Height <= (frame.Top ?? frame.Y) + frame.Height], [true, true]);
    // The ink of the box's text is inside the rectangle that is reported.
    Word.repaint(editor);
    const context = canvas.getContext('2d');
    const ink = (x, y, width, height) => { const data = context.getImageData(Math.round(x), Math.round(y), Math.max(1, Math.round(width)), Math.max(1, Math.round(height))).data; let count = 0; for (let i = 0; i < data.length; i += 4) if (data[i] < 140 && data[i + 1] < 140 && data[i + 2] < 140) count++; return count; };
    check('its text is painted in that rectangle', ink(box.X, box.Y, box.Width, box.Height) > 20, true);

    // Entered and typed into.
    const before = Word.getText(editor);
    check('enterTextBox takes its id', Word.enterTextBox(editor, box.Id), true);
    const context0 = Word.getCaretContext(editor);
    check('the caret is in a text box', [context0.Story, context0.TextBoxId === box.Id], ['textbox', true]);
    Word.insertText(editor, 'ZZ');
    check('typed into it: the box\'s text, not the body\'s', [Word.getShapeRects(editor).find((shape) => shape.Id === box.Id).Text.startsWith('ZZ'), Word.getText(editor).startsWith('ZZ')], [true, true]);
    const caret = Word.getCaretRect(editor);
    check('the caret stands in the box', [caret.X >= box.X - 2 && caret.X <= box.X + box.Width + 2, caret.Y >= box.Y - 2 && caret.Y <= box.Y + box.Height + 2], [true, true]);
    check('back to the body', [Word.exitTextBox(editor), Word.getText(editor) === before, Word.getCaretContext(editor).Story], [true, true, 'body']);
    // Saved in the header's part; the body's XML is the file's.
    const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
    const zip = await window.JSZip.loadAsync(out), source = await window.JSZip.loadAsync(bytes);
    const headers = await Promise.all(Object.keys(zip.files).filter((name) => /^word\/header\d*\.xml$/.test(name)).map((name) => zip.file(name).async('string')));
    check('saved: the typed text is in a header part, inside the box\'s text', headers.some((xml) => /<w:txbxContent>(?:(?!<\/w:txbxContent>)[^])*ZZ/.test(xml)), true);
    check('the body is the text that was read', await zip.file('word/document.xml').async('string') === await source.file('word/document.xml').async('string'), true);
    // Read again, and undone.
    const again = Word.initializeCanvas(canvas);
    Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(out));
    check('read again: the box has the typed text', Word.getShapeRects(again).filter((shape) => shape.Story === 'header').map((shape) => shape.Text.startsWith('ZZ')), [true]);
    Word.undo(editor);
    check('undone: the box\'s text is what it was', Word.getShapeRects(editor).find((shape) => shape.Id === box.Id).Text, box.Text);
    // A box of the body is listed as before, with Story "body".
    const plain = Word.initializeCanvas(canvas);
    Word.setParagraphs(plain, ['one']);
    Word.insertTextBox(plain, 'boxed');
    check('a text box of the body', Word.getShapeRects(plain).map((shape) => [shape.Story, shape.PageIndex, shape.Text]), [['body', -1, 'boxed']]);
} catch (e) { fail(e); }
finish();
