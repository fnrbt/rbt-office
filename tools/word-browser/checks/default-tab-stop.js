// w:defaultTabStop of the settings part is the distance of the default tab
// stops (708 twips, 1.25 cm, in many files from Europe): the text after a
// tab is at the next of them. It was not read: the stops were every 720
// twips (48 px) whatever the file said. The saved settings part keeps it.
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const opened = async (settings) => {
        const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
        const original = await zip.file('word/document.xml').async('string');
        const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
        const geometry = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
        zip.file('word/document.xml', head + '<w:p><w:r><w:t>a</w:t></w:r><w:r><w:tab/><w:t>b</w:t></w:r><w:r><w:tab/><w:t>c</w:t></w:r></w:p>' + `<w:sectPr>${geometry}</w:sectPr></w:body></w:document>`);
        if (settings !== null) zip.file('word/settings.xml', `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${settings}</w:settings>`);
        const editor = Word.initializeCanvas(document.getElementById('c'));
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' })));
        return editor;
    };
    const xAt = (editor, offset) => {
        const block = Word.getLayoutLines(editor)[0].BlockId;
        Word.setSelection(editor, block, offset, block, offset);
        return Math.round((Word.getCaretRect(editor).X - 96) * 100) / 100;
    };
    const european = await opened('<w:zoom w:percent="100"/><w:defaultTabStop w:val="708"/>');
    check('708 twips: the text after the tabs at 47.2 and 94.4 px from the margin', [xAt(european, 2), xAt(european, 4)], [47.2, 94.4]);
    const saved = await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(european))).arrayBuffer();
    const settingsXml = await (await window.JSZip.loadAsync(saved)).file('word/settings.xml').async('string');
    check('the saved settings part keeps it', settingsXml.includes('<w:defaultTabStop w:val="708"/>'), true);
    const inch = await opened('<w:defaultTabStop w:val="1440"/>');
    check('1440 twips: at 96 and 192 px', [xAt(inch, 2), xAt(inch, 4)], [96, 192]);
    const plain = await opened(null);
    check('none stated: every 720 twips', [xAt(plain, 2), xAt(plain, 4)], [48, 96]);
} catch (e) { fail(e); }
finish();
