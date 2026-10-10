// A PAGE field in the text of the body (not in a header or footer) shows the
// page's number in the number format of its section (w:pgNumType w:fmt):
// roman in a preface, letters where the last section says so, decimal where
// nothing is said - and the field's own format (\* ROMAN) before the
// section's. It used to be decimal always. refreshFields brings the fields up
// to the layout; the text of the document and the saved document have the result.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const zip = await window.JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
    const original = await zip.file('word/document.xml').async('string');
    const head = original.slice(0, original.indexOf('<w:body>') + '<w:body>'.length);
    const p = (text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
    const field = (label, instr) => `<w:p><w:r><w:t xml:space="preserve">${label} </w:t></w:r><w:fldSimple w:instr="${instr}"><w:r><w:t>0</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>`;
    const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    const geometry = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
    const sectionEnd = (pgNumType) => `<w:p><w:pPr><w:sectPr>${geometry}${pgNumType}</w:sectPr></w:pPr></w:p>`;
    const body = p('preface') + pageBreak + field('front', ' PAGE ') + field('own', ' PAGE \\* ROMAN ') + sectionEnd('<w:pgNumType w:fmt="lowerRoman"/>')
        + p('main') + pageBreak + field('main', ' PAGE ') + sectionEnd('<w:pgNumType w:start="1"/>')
        + p('back') + field('back', ' PAGE ');
    zip.file('word/document.xml', head + body + `<w:sectPr>${geometry}<w:pgNumType w:fmt="upperLetter" w:start="4"/></w:sectPr></w:body></w:document>`);
    const editor = Word.initializeCanvas(document.getElementById('c'));
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' })));
    check('the document has several pages', Word.getPageCount(editor) >= 5, true);
    Word.refreshFields(editor);
    const lineOf = (label) => Word.getText(editor).split('\n').find((line) => line.startsWith(label + ' '));
    note(`after refreshFields: ${JSON.stringify(['front', 'own', 'main', 'back'].map(lineOf))}`);
    check('in a section numbered in roman: ii on its second page', lineOf('front'), 'front ii end');
    check('the format the field states itself wins', lineOf('own'), 'own II end');
    check('a section that starts at 1 again, in decimal: 2 on its second page', lineOf('main'), 'main 2 end');
    check('the last section: letters from D', lineOf('back'), 'back D end');
    check('undo takes the refresh back, redo brings it again', [Word.undo(editor), lineOf('front'), Word.redo(editor), lineOf('front')], [true, 'front 0 end', true, 'front ii end']);
    // The saved document has the results.
    const saved = await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer();
    const xml = await (await window.JSZip.loadAsync(saved)).file('word/document.xml').async('string');
    check('the saved fields have the results', [/<w:fldSimple w:instr="\s*PAGE\s*">.*?<w:t[^>]*>ii<\/w:t>/.test(xml), /<w:t[^>]*>D<\/w:t>/.test(xml)], [true, true]);
} catch (e) { fail(e); }
finish();
