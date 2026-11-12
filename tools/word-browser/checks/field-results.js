// Field results retain their formatting and field boundaries after editing and
// saving, including a table of contents whose field spans paragraphs.
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
        const xml = await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string');
        const again = await open(out);
        const result = { xml, pkg, reopened: Word.exportDocumentPackage(again), same: [Word.getText(again) === Word.getText(editor), layoutText(again) === layoutText(editor)] };
        Word.dispose(again);
        return result;
    };
    const paragraphs = (xml) => xml.slice(xml.indexOf('<w:body>') + 8).split('</w:p>').map((piece) => piece + '</w:p>');
    const count = (xml, text) => xml.split(text).length - 1;
    const runsOf = (pkg, index) => pkg.Paragraphs[index].Runs.map((run) => [run.Text, !!run.Bold, !!run.Italic, (run.FieldCode ?? '').trim().split(' ')[0]]);
    const wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const sourceRunIds = (xml, texts) => {
        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        const runs = Array.from(doc.getElementsByTagNameNS(wordNs, 'r'));
        return texts.map((text) => {
            const run = runs.find((r) => Array.from(r.getElementsByTagNameNS(wordNs, 't')).some((t) => t.textContent === text));
            return run?.getAttributeNS(wordNs, 'rsidR');
        });
    };
    const boldText = (pkg, index) => pkg.Paragraphs[index].Runs.filter((run) => run.Bold).map((run) => run.Text ?? '').join('');
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
        check('the original result text keeps its source revision id and the field has its marks once',
            [sourceRunIds(result.xml, ['the', ' result text']), count(para, 'w:fldCharType="begin"'), count(para, 'w:fldCharType="separate"'), count(para, 'w:fldCharType="end"'), count(para, 'REF target')],
            [['00CC0001', '00CC0001'], 1, 1, 1, 1]);
        check('the bold field result survives reopening', boldText(result.reopened, 1).includes('QQ'), true);
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
        check('the other table of contents entries retain their source and the field boundaries remain',
            [pieces[2] === was[2], pieces[3] === was[3], count(result.xml, 'w:fldCharType="begin"'), count(result.xml, 'w:fldCharType="separate"'), count(result.xml, 'w:fldCharType="end"')],
            [true, true, 3, 3, 3]);
        check('the typed text in the last entry stays bold after reopening', boldText(result.reopened, 4).includes('QQ'), true);
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
        check('an entry in the middle of a table of contents: the entries the field begins and ends in retain their source',
            [pieces[2] === was[2], pieces[4] === was[4], pieces[3].includes('LINK'), count(result.xml, 'w:fldCharType="begin"'), count(result.xml, 'w:fldCharType="end"')], [true, true, true, 3, 3]);
        check('the entry in the middle: read again, the text of the editor that saved it', result.same[0], true);
        Word.dispose(editor);
    }
} catch (e) { fail(e); }
finish();
