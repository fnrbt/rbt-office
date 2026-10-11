// The underline of a link can be switched off. A link is shown underlined
// unless its text says that it is not - an underline switched off in the
// editor, <w:u w:val="none"/> in a file - and that is saved and read again.
// (The import made every link underlined and the painter underlined every
// link whatever its text said.)
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px'; canvas.style.top = '0px';
    const context = canvas.getContext('2d', { willReadFrequently: true });
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    const link = (text, rPr) => `<w:hyperlink w:anchor="top"><w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r></w:hyperlink>`;
    const p = (content) => `<w:p><w:r><w:t xml:space="preserve">Go to </w:t></w:r>${content}<w:r><w:t xml:space="preserve"> now.</w:t></w:r></w:p>`;
    const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
    zip.file('word/document.xml', head + p(link('underlinedlink', '<w:color w:val="0000FF"/><w:u w:val="single"/>')) + p(link('plainlink', '<w:color w:val="0000FF"/>')) + p(link('barelink', '<w:color w:val="0000FF"/><w:u w:val="none"/>')) + sect + '</w:body></w:document>');
    const open = async (bytes) => { const editor = Word.initializeCanvas(canvas); Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes)); return editor; };
    const save = async (editor) => new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
    const xmlOf = async (bytes) => (await window.JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
    // The link's segment on its line, selected; whether a rule is painted under it (a row of the line's lower part coloured over most of the word's width).
    const segmentOf = (editor, word) => { for (const line of Word.getLayoutLines(editor)) { const segment = (line.Segments ?? []).find((item) => (item.Text ?? '').includes(word)); if (segment) return { line, segment }; } return null; };
    const select = (editor, word) => { const { line } = segmentOf(editor, word); const at = line.Text.indexOf(word); Word.setSelection(editor, line.BlockId, line.StartOffset + at, line.BlockId, line.StartOffset + at + word.length); };
    const ruled = (editor, word) => {
        Word.clearSelection(editor);
        const { line, segment } = segmentOf(editor, word);
        const x0 = Math.round(segment.X) + 2, width = Math.round(segment.Width) - 4, y0 = Math.round(line.Y + line.Height * 0.55), height = Math.round(line.Height * 0.45) + 3;
        const data = context.getImageData(x0, y0, width, height).data;
        let most = 0;
        for (let row = 0; row < height; row++) { let count = 0; for (let column = 0; column < width; column++) { const i = (row * width + column) * 4; if (data[i] + data[i + 1] + data[i + 2] < 600) count++; } most = Math.max(most, count); }
        note(`${word}: line ${Math.round(line.Y)} + ${Math.round(line.Height)}, rows ${y0}..${y0 + height}, the fullest row ${most} of ${width}`);
        return most >= width * 0.9;
    };
    const underline = (editor, word) => { select(editor, word); return Word.getSelectionFormat(editor).Underline; };
    const editor = await open(await zip.generateAsync({ type: 'uint8array' }));
    check('read: a link is underlined unless its run says it is not', ['underlinedlink', 'plainlink', 'barelink'].map((word) => underline(editor, word)), [true, true, false]);
    check('painted so', ['underlinedlink', 'plainlink', 'barelink'].map((word) => ruled(editor, word)), [true, true, false]);
    // Switched off, and on again.
    select(editor, 'plainlink'); Word.toggleUnderline(editor);
    check('the underline of a link switched off: not underlined, still a link', [Word.getSelectionFormat(editor).Underline, Word.getSelectionFormat(editor).Hyperlink !== null, ruled(editor, 'plainlink')], [false, true, false]);
    const saved = await save(editor);
    const xml = await xmlOf(saved);
    const run = xml.slice(xml.lastIndexOf('<w:r>', xml.indexOf('plainlink')), xml.indexOf('plainlink'));
    check('saved: the run says <w:u w:val="none"/>, in its link', [/<w:u w:val="none"\/>/.test(run), xml.slice(0, xml.indexOf('plainlink')).lastIndexOf('<w:hyperlink') > xml.slice(0, xml.indexOf('plainlink')).lastIndexOf('</w:hyperlink>')], [true, true]);
    const again = await open(saved);
    check('read again: the same three', [['underlinedlink', 'plainlink', 'barelink'].map((word) => underline(again, word)), ['underlinedlink', 'plainlink', 'barelink'].map((word) => ruled(again, word))], [[true, false, false], [true, false, false]]);
    select(again, 'barelink'); Word.toggleUnderline(again);
    check('switched on where it was off', [Word.getSelectionFormat(again).Underline, ruled(again, 'barelink')], [true, true]);
    // A link made in the editor.
    const made = Word.initializeCanvas(canvas);
    Word.setParagraphs(made, ['See the newlink here']);
    select(made, 'newlink'); Word.setHyperlink(made, 'https://example.com/');
    check('a link made in the editor is underlined', [underline(made, 'newlink'), ruled(made, 'newlink')], [true, true]);
    select(made, 'newlink'); Word.toggleUnderline(made);
    check('and its underline is switched off with one toggle', [Word.getSelectionFormat(made).Underline, ruled(made, 'newlink')], [false, false]);
} catch (e) { fail(e); }
finish();
