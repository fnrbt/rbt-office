// Editing as in Word, through the editor's own keys (synthetic events on the
// input proxy):
//  - Enter in the paragraph that ends a section: the document keeps its
//    sections (pages, saved w:sectPr) - the end of the section goes with the
//    second part;
//  - characters typed one after the other are one undo step, so are
//    Backspaces and Deletes one after the other; a caret move, another key,
//    a format, an undo start a new step.
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    const key = (target, k, extra = {}) => {
        const event = new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...extra });
        target.dispatchEvent(event);
        return event.defaultPrevented;
    };
    const type = (target, text) => { for (const ch of text) key(target, ch); };

    // ---- Enter in the last paragraph of a section ----
    {
        const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
        const original = await zip.file('word/document.xml').async('string');
        const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
        const page = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
        const body = '<w:p><w:r><w:t>first section</w:t></w:r></w:p>'
            + `<w:p><w:pPr><w:sectPr>${page}</w:sectPr></w:pPr><w:r><w:t>its last paragraph</w:t></w:r></w:p>`
            + '<w:p><w:r><w:t>second section</w:t></w:r></w:p>'
            + `<w:sectPr><w:type w:val="nextPage"/>${page}</w:sectPr>`;
        zip.file('word/document.xml', head + body + '</w:body></w:document>');
        const editor = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' })));
        const proxy = Word.getInputElement(editor);
        const pages = () => Word.getPageFrames(editor).length;
        check('two sections: two pages', pages(), 2);
        const last = Word.getLayoutLines(editor).find((line) => (line.Text ?? '').startsWith('its last'));
        Word.setSelection(editor, last.BlockId, 3, last.BlockId, 3);
        key(proxy, 'Enter');
        check('Enter in the paragraph that ends the first section: still two pages', [Word.getText(editor).split('\n'), pages()], [['first section', 'its', ' last paragraph', 'second section'], 2]);
        key(proxy, 'End'); key(proxy, 'Enter'); type(proxy, 'more');
        check('Enter at its end, typing: still two pages', [Word.getText(editor).split('\n'), pages()], [['first section', 'its', ' last paragraph', 'more', 'second section'], 2]);
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const xml = await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string');
        check('saved with two sections', (xml.match(/<w:sectPr[ >]/g) ?? []).length, 2);
        const firstEnd = xml.indexOf('<w:sectPr');
        const paragraphOfIt = xml.slice(xml.lastIndexOf('<w:p>', firstEnd), xml.indexOf('</w:p>', firstEnd));
        check('the first one in the paragraph typed last, before the second section', [paragraphOfIt.includes('>more<'), xml.indexOf('second section') > firstEnd], [true, true]);
        Word.dispose(editor);
    }

    // ---- runs of typing are undo steps ----
    {
        const editor = Word.initialize('c');
        const proxy = Word.getInputElement(editor);
        canvas.focus();
        const text = () => Word.getText(editor);
        const steps = () => { let n = 0; while (Word.canUndo(editor)) { Word.undo(editor); n++; } return n; };
        type(proxy, 'Hello world');
        check('typed', text(), 'Hello world');
        check('a run of typing is one undo step', [Word.undo(editor), text(), Word.canUndo(editor)], [true, '', false]);
        check('... redone as one', [Word.redo(editor), text()], [true, 'Hello world']);
        type(proxy, '!');
        check('typing after a redo is a step of its own', [Word.undo(editor), text(), Word.undo(editor), text()], [true, 'Hello world', true, '']);
        type(proxy, 'abc');
        key(proxy, 'ArrowLeft');
        type(proxy, 'X');
        check('a caret move starts a new step', [text(), Word.undo(editor), text()], ['abXc', true, 'abc']);
        key(proxy, 'End');
        type(proxy, 'def');
        key(proxy, 'Backspace'); key(proxy, 'Backspace');
        check('Backspaces after typing: a step of their own', [text(), Word.undo(editor), text(), Word.undo(editor), text()], ['abcd', true, 'abcdef', true, 'abc']);
        key(proxy, 'Home');
        key(proxy, 'Delete'); key(proxy, 'Delete');
        check('Deletes one after the other: one step', [text(), Word.undo(editor), text()], ['c', true, 'abc']);
        key(proxy, 'End');
        type(proxy, 'gh');
        key(proxy, 'Enter');
        type(proxy, 'ij');
        check('Enter is a step between two runs', [text(), Word.undo(editor), text(), Word.undo(editor), text(), Word.undo(editor), text()], ['abcgh\nij', true, 'abcgh\n', true, 'abcgh', true, 'abc']);
        key(proxy, 'End');
        type(proxy, 'k');
        key(proxy, 'b', { ctrlKey: true });
        type(proxy, 'l');
        check('a format chosen at the caret starts a new step', [text(), Word.undo(editor), text()], ['abckl', true, 'abck']);
        // Text from a virtual keyboard (beforeinput) goes on a run as keys do.
        type(proxy, 'm');
        proxy.dispatchEvent(new InputEvent('beforeinput', { data: 'n', inputType: 'insertText', bubbles: true, cancelable: true }));
        check('beforeinput text on a run of keys', [text(), Word.undo(editor), text()], ['abckmn', true, 'abck']);
        check('all undone in steps', steps(), 2);
        Word.dispose(editor);
    }

    // ---- copy, cut and paste with formatting ----
    {
        const editor = Word.initialize('c');
        Word.setParagraphs(editor, ['Hello bold world', 'Second', '']);
        const ids = [...new Set(Word.getLayoutLines(editor).map((line) => line.BlockId))];
        const proxy = Word.getInputElement(editor);
        canvas.focus();
        const text = () => Word.getText(editor);
        const boldIn = (block, from, to) => { Word.setSelection(editor, block, from, block, to); return Word.getSelectionFormat(editor).Bold; };
        const blockOf = (start) => Word.getLayoutLines(editor).find((line) => (line.Text ?? '').startsWith(start)).BlockId;
        const clipboardEvent = (type, data) => { const event = new ClipboardEvent(type, { clipboardData: data, bubbles: true, cancelable: true }); proxy.dispatchEvent(event); return event.defaultPrevented; };
        const paste = (plain, html) => { const data = new DataTransfer(); if (plain) data.setData('text/plain', plain); if (html) data.setData('text/html', html); return clipboardEvent('paste', data); };
        Word.setSelection(editor, ids[0], 6, ids[0], 10);
        Word.toggleBold(editor);
        // Ctrl+C: left to the browser, whose copy event follows; meanwhile the proxy has the selection's text, selected.
        const prevented = key(proxy, 'c', { ctrlKey: true });
        check('Ctrl+C is left to the browser, the selection is in the proxy, selected', [prevented, proxy.value, proxy.selectionStart, proxy.selectionEnd], [false, 'bold', 0, 4]);
        const copied = new DataTransfer();
        check('the copy event is taken', clipboardEvent('copy', copied), true);
        check('copied: text and HTML; the proxy is empty again', [copied.getData('text/plain'), copied.getData('text/html'), proxy.value], ['bold', '<p><span style="font-weight:bold">bold</span></p>', '']);
        Word.setSelection(editor, ids[1], 6, ids[1], 6);
        check('Ctrl+V is left to the browser', key(proxy, 'v', { ctrlKey: true }), false);
        check('the paste event is taken', paste('bold', copied.getData('text/html')), true);
        check('pasted with its formatting', [text(), boldIn(ids[1], 6, 10), boldIn(ids[1], 0, 6)], ['Hello bold world\nSecondbold\n', true, false]);
        check('one undo step', [Word.undo(editor), text()], [true, 'Hello bold world\nSecond\n']);
        // Other text on the clipboard than was copied: that text (the HTML there is not ours).
        Word.setSelection(editor, ids[1], 0, ids[1], 0);
        paste('plain ', '');
        check('other text: as text', [text(), boldIn(ids[1], 0, 6)], ['Hello bold world\nplain Second\n', false]);
        Word.undo(editor);
        // HTML of another program: its paragraphs, list and formatting.
        key(proxy, 'End', { ctrlKey: true });   // (the empty paragraph at the end)
        paste('x Big\nitem\nend', '<html><body><!--StartFragment--><p>x <b>Big</b></p><ul><li>item</li></ul><p>end</p><!--EndFragment--></body></html>');
        check('HTML from elsewhere: its paragraphs', text(), 'Hello bold world\nSecond\nx Big\nitem\nend');
        check('... its bold text', boldIn(blockOf('x Big'), 2, 5), true);
        Word.setSelection(editor, blockOf('item'), 1, blockOf('item'), 1);
        check('... its list', Word.getSelectionFormat(editor).Paragraph.ListKind, 'bullet');
        check('... one undo step', [Word.undo(editor), text()], [true, 'Hello bold world\nSecond\n']);
        // Cut.
        Word.setSelection(editor, ids[0], 0, ids[0], 6);
        const cut = new DataTransfer();
        clipboardEvent('cut', cut);
        check('cut: copied and taken out', [cut.getData('text/plain'), text()], ['Hello ', 'bold world\nSecond\n']);
        check('... one undo step', [Word.undo(editor), text()], [true, 'Hello bold world\nSecond\n']);
        // With nothing selected a copy event is the browser's.
        Word.setSelection(editor, ids[0], 2, ids[0], 2);
        check('nothing selected: the copy event is left alone', clipboardEvent('copy', new DataTransfer()), false);
        // The host's menu: copySelection / cutSelection / pasteClipboard.
        Word.setSelection(editor, ids[0], 6, ids[0], 10);
        const fromApi = Word.copySelection(editor);
        check('copySelection', [fromApi.Text, fromApi.Html], ['bold', '<p><span style="font-weight:bold">bold</span></p>']);
        Word.setSelection(editor, ids[0], 0, ids[0], 0);
        Word.pasteClipboard(editor, 'bold', null);
        check('pasteClipboard: the copy, with its formatting', [text().split('\n')[0], boldIn(ids[0], 0, 4)], ['boldHello bold world', true]);
        Word.setSelection(editor, ids[0], 0, ids[0], 4);
        check('cutSelection', [Word.cutSelection(editor).Text, text().split('\n')[0]], ['bold', 'Hello bold world']);
        check('copySelection with nothing selected', Word.copySelection(editor), null);
        Word.dispose(editor);
    }

    // ---- Word's formatting keys ----
    {
        const editor = Word.initialize('c');
        Word.setParagraphs(editor, ['Some text here', 'Second']);
        const ids = [...new Set(Word.getLayoutLines(editor).map((line) => line.BlockId))];
        const proxy = Word.getInputElement(editor);
        canvas.focus();
        const ctrl = (k, extra = {}) => key(proxy, k, { ctrlKey: true, ...extra });
        const select = (from, to, block = 0) => Word.setSelection(editor, ids[block], from, ids[block], to);
        const format = () => Word.getSelectionFormat(editor);
        select(2, 2);
        check('Ctrl+E centres the paragraph (taken)', [ctrl('e'), format().Paragraph.Alignment], [true, 'center']);
        ctrl('r'); const right = format().Paragraph.Alignment;
        ctrl('j'); const justified = format().Paragraph.Alignment;
        ctrl('l');
        check('Ctrl+R, Ctrl+J, Ctrl+L', [right, justified, format().Paragraph.Alignment], ['right', 'justify', 'left']);
        check('each one undo step', [Word.undo(editor), format().Paragraph.Alignment, Word.undo(editor), format().Paragraph.Alignment], [true, 'justify', true, 'right']);
        ctrl('m');
        check('Ctrl+M: half an inch in', format().Paragraph.IndentLeft, 48);
        ctrl('M', { shiftKey: true });
        check('Ctrl+Shift+M: back', format().Paragraph.IndentLeft, 0);
        select(0, 14);
        Word.setFontSizePt(editor, 11);
        select(0, 4);
        const size = format().FontSizePt;
        ctrl('>', { shiftKey: true });
        check('Ctrl+Shift+>: the next size', [size, format().FontSizePt], [11, 12]);
        ctrl(']');
        check('Ctrl+]: a point larger', format().FontSizePt, 13);
        ctrl('<', { shiftKey: true });
        check('Ctrl+Shift+<: the size below', format().FontSizePt, 12);
        ctrl('[');
        check('Ctrl+[: a point smaller', format().FontSizePt, 11);
        ctrl('=');
        check('Ctrl+=: subscript', format().VertAlign, 'subscript');
        ctrl('+', { shiftKey: true });
        check('Ctrl+Shift+=: superscript', format().VertAlign, 'superscript');
        ctrl('b');
        ctrl(' ');
        check('Ctrl+Space: the character formatting cleared', [format().Bold, format().VertAlign ?? null], [false, null]);
        check('the text is as it was', Word.getText(editor), 'Some text here\nSecond');
        // At a caret: for what is typed next, as Ctrl+B.
        select(14, 14);
        ctrl('>', { shiftKey: true });
        type(proxy, 'X');
        select(14, 15);
        check('at a caret: the size of what is typed next', format().FontSizePt, 12);
        Word.dispose(editor);
    }
} catch (e) { fail(e); }
finish();
