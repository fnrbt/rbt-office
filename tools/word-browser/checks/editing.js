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

    // ---- a run of Backspaces across a paragraph mark or a bullet is one step ----
    {
        const editor = Word.initialize('c');
        const proxy = Word.getInputElement(editor);
        canvas.focus();
        const text = () => Word.getText(editor);
        Word.setParagraphs(editor, ['ab', 'cd']);
        let ids = [...new Set(Word.getLayoutLines(editor).map((line) => line.BlockId))];
        Word.setSelection(editor, ids[1], 2, ids[1], 2);
        for (let i = 0; i < 4; i++) key(proxy, 'Backspace');
        check('Backspaces over a paragraph mark', text(), 'a');
        check('... one undo step with the characters', [Word.undo(editor), text(), Word.canUndo(editor)], [true, 'ab\ncd', false]);
        check('... redone as one', [Word.redo(editor), text()], [true, 'a']);
        Word.undo(editor);
        Word.setSelection(editor, ids[1], 0, ids[1], 0);
        Word.toggleBulletList(editor);
        Word.setSelection(editor, ids[1], 2, ids[1], 2);
        const list = () => Word.getSelectionFormat(editor).Paragraph.ListKind;
        key(proxy, 'Backspace'); key(proxy, 'Backspace'); key(proxy, 'Backspace');
        check('Backspaces to the start of a list item, then the bullet taken off', [text(), list()], ['ab\n', 'none']);
        key(proxy, 'Backspace');
        check('... and the paragraphs joined', text(), 'ab');
        check('... all one undo step', [Word.undo(editor), text(), (Word.setSelection(editor, ids[1], 0, ids[1], 0), list())], [true, 'ab\ncd', 'bullet']);
        check('... the step before it is the bullet', [Word.undo(editor), list()], [true, 'none']);
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
        const clipboardEvent = (type, data) => {
            const event = new ClipboardEvent(type, { clipboardData: data, bubbles: true, cancelable: true });
            // (Firefox gives a synthetic ClipboardEvent a DataTransfer of its own, without the init's data.)
            Object.defineProperty(event, 'clipboardData', { value: data });
            proxy.dispatchEvent(event); return event.defaultPrevented;
        };
        const paste = (plain, html) => { const data = new DataTransfer(); if (plain) data.setData('text/plain', plain); if (html) data.setData('text/html', html); return clipboardEvent('paste', data); };
        Word.setSelection(editor, ids[0], 6, ids[0], 10);
        Word.toggleBold(editor);
        // Ctrl+C: left to the browser, whose copy event follows; meanwhile the proxy has the selection's text, selected.
        const prevented = key(proxy, 'c', { ctrlKey: true });
        check('Ctrl+C is left to the browser, the selection is in the proxy, selected', [prevented, proxy.value, proxy.selectionStart, proxy.selectionEnd], [false, 'bold', 0, 4]);
        const copied = new DataTransfer();
        check('the copy event is taken', clipboardEvent('copy', copied), true);
        check('copied text and marked HTML; the proxy is empty', [copied.getData('text/plain'), copied.getData('text/html').startsWith('<p><span style="font-weight:bold">bold</span></p><!--rbt-word-clip:'), proxy.value], ['bold', true, '']);
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
        // A table of another program is a table.
        Word.setSelection(editor, ids[0], 5, ids[0], 5);
        paste('x\na\tb\nc\td\ny', '<html><body><!--StartFragment--><p>x</p><table border="1"><tbody><tr><td>a</td><td><b>b</b></td></tr><tr><td>c</td><td><p>d</p><p>d2</p></td></tr></tbody></table><p>y</p><!--EndFragment--></body></html>');
        const cellOf = (start) => { Word.setSelection(editor, blockOf(start), 0, blockOf(start), 0); const at = Word.getCaretContext(editor); return [at.InTable, at.Row, at.Column, at.RowCount, at.ColumnCount]; };
        check('HTML table: a table, its cells', [cellOf('a'), cellOf('b'), cellOf('c'), cellOf('d2')], [[true, 0, 0, 2, 2], [true, 0, 1, 2, 2], [true, 1, 0, 2, 2], [true, 1, 1, 2, 2]]);
        check('... no tabs in its text', Word.getLayoutLines(editor).some((line) => (line.Text ?? '').includes('\t')), false);
        check('... the text before and after it', [cellOf('Hellox')[0], cellOf('y bold')[0]], [false, false]);
        check('... the formatting of its cells', boldIn(blockOf('b'), 0, 1), true);
        const savedTable = await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer();
        const tableXml = await (await window.JSZip.loadAsync(savedTable)).file('word/document.xml').async('string');
        check('... saved as a table of two rows', [tableXml.split('<w:tbl>').length - 1, tableXml.split('<w:tr>').length - 1 + tableXml.split('<w:tr ').length - 1], [1, 2]);
        // Pasted into a cell of a table: its rows are paragraphs, the cells parted by tabs.
        Word.setSelection(editor, blockOf('c'), 1, blockOf('c'), 1);
        paste('m\tn', '<table><tr><td>m</td><td>n</td></tr></table>');
        check('a table pasted into a cell: its text, the cells parted by a tab', [Word.getLayoutLines(editor).some((line) => line.Text === 'cm\tn'), cellOf('cm')[0]], [true, true]);
        Word.undo(editor);
        check('... one undo step', [Word.undo(editor), text()], [true, 'Hello bold world\nSecond\n']);
        // A synthetic clipboard transfer that throws cannot authorize a Cut.
        Word.setSelection(editor, ids[0], 0, ids[0], 6);
        const failedTransfer = new DataTransfer();
        failedTransfer.setData = () => { throw new Error('clipboard unavailable'); };
        const originalError = console.error;
        let transferReported = false;
        console.error = () => { transferReported = true; };
        try {
            const beforeCut = text();
            clipboardEvent('cut', failedTransfer);
            check('failed synchronous clipboard transfer cannot cut', [text(), transferReported], [beforeCut, true]);
        } finally { console.error = originalError; }
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
        check('copySelection includes rich provenance', [fromApi.Text, fromApi.Html.startsWith('<p><span style="font-weight:bold">bold</span></p><!--rbt-word-clip:')], ['bold', true]);
        Word.setSelection(editor, ids[0], 0, ids[0], 0);
        Word.pasteClipboard(editor, fromApi.Text, fromApi.Html);
        check('marked host copy preserves original rich formatting', boldIn(ids[0], 0, 4), true);
        Word.undo(editor);
        Word.setSelection(editor, ids[0], 0, ids[0], 0);
        Word.pasteClipboard(editor, 'bold', null);
        check('text-only external overwrite does not reuse rich copy', [text().split('\n')[0], boldIn(ids[0], 0, 4)], ['boldHello bold world', false]);
        Word.setSelection(editor, ids[0], 0, ids[0], 4);
        check('cutSelection', [Word.cutSelection(editor).Text, text().split('\n')[0]], ['bold', 'Hello bold world']);
        check('copySelection with nothing selected', Word.copySelection(editor), null);
        // Even identical plain text may have new HTML formatting or a link.
        Word.setSelection(editor, ids[0], 0, ids[0], 0);
        Word.pasteClipboard(editor, 'bold', '<p><i><a href="https://example.com/new">bold</a></i></p>');
        Word.setSelection(editor, ids[0], 0, ids[0], 4);
        check('identical text from external HTML keeps its italic and link', [Word.getSelectionFormat(editor).Italic, Word.getSelectionFormat(editor).Hyperlink, Word.getSelectionFormat(editor).Bold], [true, 'https://example.com/new', false]);
        Word.undo(editor);
        // A program that carries our comment but edits the HTML is not the
        // original rich clip; its visible formatting takes precedence.
        Word.setSelection(editor, ids[0], 6, ids[0], 10);
        const marked = Word.copySelection(editor);
        Word.setSelection(editor, ids[0], 0, ids[0], 0);
        Word.pasteClipboard(editor, 'bold', marked.Html.replace('font-weight:bold', 'font-style:italic'));
        Word.setSelection(editor, ids[0], 0, ids[0], 4);
        check('modified marked HTML is parsed, not cached', [Word.getSelectionFormat(editor).Italic, Word.getSelectionFormat(editor).Bold], [true, false]);
        Word.undo(editor);
        const secondCanvas = document.createElement('canvas');
        secondCanvas.id = 'clipboard-second';
        secondCanvas.width = canvas.width;
        secondCanvas.height = canvas.height;
        document.body.appendChild(secondCanvas);
        const second = Word.initialize(secondCanvas.id);
        Word.pasteClipboard(second, fromApi.Text, fromApi.Html);
        const secondId = Word.getLayoutLines(second)[0].BlockId;
        Word.setSelection(second, secondId, 0, secondId, 4);
        check('another editor parses marked HTML instead of using foreign clip', [Word.getText(second).includes('bold'), Word.getSelectionFormat(second).Bold], [true, true]);
        Word.dispose(second);
        secondCanvas.remove();
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

    // ---- Shift+F3, Ctrl+Shift+L, Ctrl+Shift+N ----
    {
        const editor = Word.initialize('c');
        Word.setParagraphs(editor, ['the quick fox', 'Second']);
        const ids = [...new Set(Word.getLayoutLines(editor).map((line) => line.BlockId))];
        const proxy = Word.getInputElement(editor);
        canvas.focus();
        const select = (from, to, block = 0) => Word.setSelection(editor, ids[block], from, ids[block], to);
        const first = () => Word.getText(editor).split('\n')[0];
        const format = () => Word.getSelectionFormat(editor);
        select(4, 9);
        Word.toggleBold(editor);
        select(0, 13);
        const cycle = [];
        for (let i = 0; i < 3; i++) { key(proxy, 'F3', { shiftKey: true }); cycle.push(first()); }
        check('Shift+F3 on a selection: Each Word, UPPER, lower (taken)', cycle, ['The Quick Fox', 'THE QUICK FOX', 'the quick fox']);
        const sel = Word.getSelection(editor);
        check('... the selection stays, the formatting too', [sel.AnchorOffset, sel.FocusOffset, (select(4, 9), format().Bold), (select(0, 3), format().Bold)], [0, 13, true, false]);
        check('... one undo step each', [Word.undo(editor), first(), Word.undo(editor), first()], [true, 'THE QUICK FOX', true, 'The Quick Fox']);
        Word.undo(editor);
        select(6, 6);
        key(proxy, 'F3', { shiftKey: true });
        check('with a caret: the word at it', [first(), Word.getSelection(editor).FocusOffset, Word.getSelection(editor).IsCollapsed], ['the Quick fox', 6, true]);
        key(proxy, 'F3', { shiftKey: true });
        check('... again: upper case', first(), 'the QUICK fox');
        select(0, 0, 1);
        check('Ctrl+Shift+L: a bulleted list (taken)', [key(proxy, 'L', { ctrlKey: true, shiftKey: true }), format().Paragraph.ListKind], [true, 'bullet']);
        key(proxy, 'L', { ctrlKey: true, shiftKey: true });
        check('... twice: still one', format().Paragraph.ListKind, 'bullet');
        check('... the alignment is as it was', format().Paragraph.Alignment, 'left');
        key(proxy, '1', { ctrlKey: true });
        check('a heading', format().Paragraph.HeadingLevel, 1);
        check('Ctrl+Shift+N: the Normal style (taken)', [key(proxy, 'N', { ctrlKey: true, shiftKey: true }), format().Paragraph.HeadingLevel ?? 0], [true, 0]);
        Word.setParagraphStyle(editor, 'Quote');
        key(proxy, 'N', { ctrlKey: true, shiftKey: true });
        check('... a named style taken off', format().Paragraph.StyleId ?? null, null);
        Word.dispose(editor);
    }
} catch (e) { fail(e); }
finish();
