// An underline read from a file (w:u w:val="single") and one made in the
// editor are the same formatting in the model: text underlined both ways is
// one run with one format (not "single" beside none), and both are saved as
// w:u w:val="single". Another line style (double) stays what it is.
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/floating-image.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    const run = (value, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${value}</w:t></w:r>`;
    zip.file('word/document.xml', head + `<w:p>${run('from the file ', '<w:u w:val="single"/>')}${run('made here')}${run(' and doubled', '<w:u w:val="double"/>')}</w:p>`
        + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>');
    const editor = Word.initializeCanvas(document.createElement('canvas'));
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' })));
    const blockId = Word.getLayoutLines(editor)[0].BlockId;
    Word.setSelection(editor, blockId, 'from the file '.length, blockId, 'from the file made here'.length);
    Word.toggleUnderline(editor);
    const runs = Word.exportDocumentPackage(editor).Paragraphs[0].Runs.map((r) => [r.Text, r.Underline === true, r.UnderlineStyle ?? null]);
    check('underlined in the file and in the editor: one run, one format', runs, [['from the file made here', true, null], [' and doubled', true, 'double']]);
    Word.setSelection(editor, blockId, 0, blockId, 'from the file made here'.length);
    const format = Word.getSelectionFormat(editor);
    check('the selection\'s format: underlined, not mixed', format.Underline, true);
    const saved = await (await window.JSZip.loadAsync(new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer()))).file('word/document.xml').async('string');
    check('saved: single underlines, and the double one', [(saved.match(/<w:u w:val="single"\/>/g) ?? []).length >= 1, (saved.match(/<w:u w:val="double"\/>/g) ?? []).length, /<w:u w:val="single"\/>[^]*from the file[^]*made here/.test(saved)], [true, 1, true]);
} catch (e) { fail(e); }
finish();
