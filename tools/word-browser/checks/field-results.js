// Text typed into the result of a field in another format, and the entries of
// a table of contents.
//
// Bold chosen at a caret inside the result of a field and two characters
// typed: the first was of the field and the second was not (text typed at
// the edge of a run is not of the run's field - also where the run after it
// is of the same field), so the paragraph was written from the model. And a
// result a part of which is formatted differently could not be changed in
// its XML, since the reader made one run of a field's result: a table of
// contents - one field over many paragraphs - was written from the model
// whole whenever one of its entries was.
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const bytes = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/field-results.docx'));
    const original = await (await window.JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
    const canvas = document.getElementById('c');
    const open = async (data = bytes) => { const editor = Word.initializeCanvas(canvas); Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(data)); Word.completeLayout(editor); return editor; };
    const layoutText = (e) => JSON.stringify(Word.getLayoutFingerprint(e)).replace(/,single\//g, ',-/');
    const saved = async (editor) => {
        const pkg = Word.exportDocumentPackage(editor);
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(pkg)).arrayBuffer());
        const report = OF.word.getLastSaveReport();
        const xml = await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string');
        const again = await open(out);
        const result = { xml, report, pkg, same: [Word.getText(again) === Word.getText(editor), layoutText(again) === layoutText(editor)] };
        Word.dispose(again);
        return result;
    };
    const paragraphs = (xml) => xml.slice(xml.indexOf('<w:body>') + 8).split('</w:p>').map((piece) => piece + '</w:p>');
    const count = (xml, text) => xml.split(text).length - 1;
    const runsOf = (pkg, index) => pkg.Paragraphs[index].Runs.map((run) => [run.Text, !!run.Bold, !!run.Italic, (run.FieldCode ?? '').trim().split(' ')[0]]);
    {
        const editor = await open();
        const pkg = Word.exportDocumentPackage(editor);
        check('a result of which a part is bold is read as its parts, each with the field\'s code and whole result',
            [runsOf(pkg, 5), pkg.Paragraphs[5].Runs.map((run) => run.FieldResult)], [[['quoted ', false, false, 'QUOTE'], ['bold', true, false, 'QUOTE']], ['quoted bold', 'quoted bold']]);
        const unedited = await saved(editor);
        check('saved unedited: the body is the text that was read', unedited.xml === original, true);
        Word.dispose(editor);
    }
    {
        // "see the| result text after": bold chosen, two characters typed.
        const editor = await open();
        const id = Word.findText(editor, 'see the result', true)[0].BlockId;
        Word.setSelection(editor, id, 7, id, 7);
        Word.toggleBold(editor);
        Word.insertText(editor, 'Q'); Word.insertText(editor, 'Q');
        const result = await saved(editor);
        const para = paragraphs(result.xml)[1];
        check('bold chosen inside the result of a field, two characters typed: both are of the field, one bold run',
            runsOf(result.pkg, 1), [['see ', false, false, ''], ['the', false, true, 'REF'], ['QQ', true, true, 'REF'], [' result text', false, true, 'REF'], [' after', false, false, '']]);
        check('the paragraph is changed in the XML it was read from', [result.report.PatchedParagraphs, result.report.GeneratedParagraphs, result.report.ModelBlocks, result.report.Reasons], [1, 0, 0, []]);
        check('the result\'s run is split, the typed part bold, each part with the run\'s revision id; the field has its marks once',
            [para.includes('<w:r w:rsidR="00CC0001"><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">the</w:t></w:r><w:r w:rsidR="00CC0001"><w:rPr><w:b/><w:i/></w:rPr><w:t xml:space="preserve">QQ</w:t></w:r><w:r w:rsidR="00CC0001"><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> result text</w:t></w:r>')
                || para.includes('<w:t>the</w:t></w:r><w:r w:rsidR="00CC0001"><w:rPr><w:b/><w:i/></w:rPr><w:t>QQ</w:t></w:r><w:r w:rsidR="00CC0001"><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> result text</w:t></w:r>'),
             count(para, 'w:fldCharType="begin"'), count(para, 'w:fldCharType="separate"'), count(para, 'w:fldCharType="end"'), count(para, 'REF target')], [true, 1, 1, 1, 1]);
        check('the field: read again, the text and the layout of the editor that saved it', result.same, [true, true]);
        Word.dispose(editor);
    }
    {
        // The last entry of the table of contents (the field ends in it): bold chosen, two characters typed.
        const editor = await open();
        const id = Word.findText(editor, 'entry three', true)[0].BlockId;
        Word.setSelection(editor, id, 5, id, 5);
        Word.toggleBold(editor);
        Word.insertText(editor, 'Q'); Word.insertText(editor, 'Q');
        const result = await saved(editor);
        const pieces = paragraphs(result.xml), was = paragraphs(original);
        check('text typed in another format into the last entry of a table of contents: the entry is changed in its own XML, the others are the text they were',
            [result.report.PatchedParagraphs, result.report.GeneratedParagraphs, result.report.ModelBlocks, result.report.Reasons, pieces[2] === was[2], pieces[3] === was[3], pieces[4].startsWith('<w:p w:rsidR="00CC0003">'), /<w:b\/><\/w:rPr><w:t[^>]*>QQ<\/w:t>/.test(pieces[4]), count(result.xml, 'w:fldCharType="end"')],
            [1, 0, 0, [], true, true, true, true, 3]);
        check('the last entry: read again, the text and the layout of the editor that saved it', result.same, [true, true]);
        Word.dispose(editor);
    }
    {
        // An entry in the middle is written from the model (a link is put into it): the first and the last stay.
        const editor = await open();
        const id = Word.findText(editor, 'entry two', true)[0].BlockId;
        Word.setSelection(editor, id, 5, id, 5);
        Word.insertHyperlink(editor, 'https://example.com/entry', 'LINK');
        const result = await saved(editor);
        const pieces = paragraphs(result.xml), was = paragraphs(original);
        check('an entry in the middle of a table of contents written from the model: the entries the field begins and ends in are the text they were',
            [result.report.ModelBlocks, pieces[2] === was[2], pieces[4] === was[4], pieces[3].includes('LINK'), count(result.xml, 'w:fldCharType="begin"'), count(result.xml, 'w:fldCharType="end"')], [1, true, true, true, 3, 3]);
        check('the entry in the middle: read again, the text of the editor that saved it', result.same[0], true);
        Word.dispose(editor);
    }
} catch (e) { fail(e); }
finish();
