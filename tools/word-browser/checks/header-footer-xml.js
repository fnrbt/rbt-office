// A header or footer that was typed in is saved as the XML it was read from with the change made in it, as
// the body is: what the model does not hold of it (later formats' attributes, proofing marks, a smart tag, a
// field, a comment range, a link, a bookmark, the part's own declaration and namespaces) stays; its other
// paragraphs are the text they were. A paragraph split in two there is two paragraphs made of the one that was
// read. (synthetic/kept-header; real keys.)
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
let OF, Word, editor, original, footer;
const saved = {};
try {
    OF = await loadFormats();
    Word = await loadWord();
    const bytes = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/kept-header.docx'));
    const zip = await window.JSZip.loadAsync(bytes);
    original = await zip.file('word/header1.xml').async('string');
    footer = await zip.file('word/footer1.xml').async('string');
    window.place = async (stage) => {
        if (editor) Word.dispose(editor);
        editor = Word.initializeCanvas(document.getElementById('c'));
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
        Word.completeLayout(editor);
        Word.enterHeaderFooter(editor, 0, stage === 'footer' ? 'footer' : 'header');
        Word.getInputElement(editor).focus();
    };
    window.snap = async (stage) => {
        Word.exitHeaderFooter(editor);
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const report = OF.word.getLastSaveReport();
        const savedZip = await window.JSZip.loadAsync(out);
        const part = async (name) => { const file = savedZip.file(name); return file ? file.async('string') : null; };
        const again = Word.initializeCanvas(document.createElement('canvas'));
        Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(out));
        const headerText = (e) => JSON.stringify(Word.exportDocumentPackage(e).HeaderParts.concat(Word.exportDocumentPackage(e).FooterParts).map((p) => (p.Paragraphs ?? []).map((para) => (para.Runs ?? []).map((run) => run.Text ?? '').join(''))));
        saved[stage] = { header: await part('word/header1.xml'), footer: await part('word/footer1.xml'), document: await part('word/document.xml'), rels: await part('word/_rels/header1.xml.rels'), report, readAgain: headerText(again) === headerText(editor), text: headerText(editor) };
        Word.dispose(again);
    };
    window.drive = [
        ['eval', 'window.place("type")'], ['type', 'X'], ['eval', 'window.snap("type")'],
        ['eval', 'window.place("split")'], ...Array(10).fill(['press', 'ArrowRight']), ['press', 'Enter'], ['eval', 'window.snap("split")'],
        ['eval', 'window.place("footer")'], ['type', 'Y'], ['eval', 'window.snap("footer")'],
        ['eval', 'window.place("footer")'], ['eval', 'window.field()'], ['eval', 'window.snap("field")'],
    ];
    // A field put in before the footer's text (it takes the style of the run there): a field, not a run made like that run.
    window.field = () => { Word.insertPageCountField(editor); Word.insertText(editor, ' pages, '); };
} catch (e) { fail(e); finish(); }
window.finish = async () => {
    try {
        const second = original.slice(original.indexOf('<w:p w14:paraId="3A3B3C10"'), original.indexOf('</w:hdr>'));
        // Typed at the start of the header: the character in the first run's text, nothing else.
        let { header, rels, report, readAgain } = saved.type;
        check('typed in the header: the text', saved.type.text.includes('Xalpha beta gamma'), true);
        check('typed in the header: header1.xml is the text that was read with the character in its run',
            header === original.replace('>alpha beta gamma<', '>Xalpha beta gamma<') ? true : header, true);
        check('typed in the header: its relationships are what they were (the link\'s)', rels !== null && rels.includes('Id="rId30"') && (rels.match(/<Relationship /g) ?? []).length === 1, true);
        check('typed in the header: the footer is the text it was', saved.type.footer === footer, true);
        check('typed in the header: read again, the same text', readAgain, true);
        check('typed in the header: no paragraph written from the model', report.GeneratedParagraphs, 0);

        // Split in the header: two paragraphs made of the first; the second paragraph is the text it was.
        ({ header, report, readAgain } = saved.split);
        check('split in the header: the text', saved.split.text.includes('"alpha beta"') && saved.split.text.includes(' gamma delta'), true);
        check('split in the header: the halves keep the paragraph\'s properties and the run\'s', [
            header.includes('w14:paraId="3A3B3C09"'), (header.match(/<w:suppressAutoHyphens\/>/g) ?? []).length, (header.match(/<w14:ligatures w14:val="standard"\/>/g) ?? []).length,
            header.includes('<w:instrText xml:space="preserve"> QUOTE "result" </w:instrText>'), header.includes('<w:commentRangeStart w:id="7"/>'), header.includes('<w:hyperlink r:id="rId30" w:tooltip="a tip" w:history="1">')], [true, 2, 2, true, true, true]);
        check('split in the header: the other paragraph is the text it was', header.includes(second), true);
        check('split in the header: the part\'s declaration and root are what they were', header.slice(0, header.indexOf('>', header.indexOf('<w:hdr')) + 1) === original.slice(0, original.indexOf('>', original.indexOf('<w:hdr')) + 1), true);
        check('split in the header: no paragraph written from the model', report.GeneratedParagraphs, 0);
        check('split in the header: read again, the same text', readAgain, true);

        // Typed in the footer: its proofing marks and field stay.
        ({ report, readAgain } = saved.footer);
        check('typed in the footer: footer1.xml is the text that was read with the character in its run',
            saved.footer.footer === footer.replace('>Page <', '>YPage <') ? true : saved.footer.footer, true);
        check('typed in the footer: the header is the text it was', saved.footer.header === original, true);
        check('typed in the footer: read again, the same text', readAgain, true);

        // A field put in at the start of the footer: a field of its own (no run in a run); the rest as it was.
        const withField = saved.field.footer;
        check('a field put in the footer: it is a field, not text in a run made like the run after it', [withField.includes('NUMPAGES'), withField.includes('<w:r><w:r>'), /<w:r>(<w:rPr>(?:(?!<\/w:rPr>).)*<\/w:rPr>)?<w:r>/.test(withField)], [true, false, false]);
        check('a field put in the footer: the text typed after it is in the run it was typed into; the footer\'s proofing marks, its field and its paragraph id are what they were',
            [withField.includes('<w:proofErr w:type="gramStart"/><w:r><w:t xml:space="preserve"> pages, Page </w:t></w:r>'), withField.includes(footer.slice(footer.indexOf('<w:proofErr w:type="gramEnd"/>'), footer.indexOf('</w:p>'))),
                withField.includes('<w:p w14:paraId="3A3B3C20"><w:pPr><w:jc w:val="center"/></w:pPr>')], [true, true, true]);
        check('a field put in the footer: read again, the same text', saved.field.readAgain, true);
    } catch (e) { fail(e); }
    finish();
};
