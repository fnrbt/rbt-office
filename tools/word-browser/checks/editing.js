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

} catch (e) { fail(e); }
finish();
