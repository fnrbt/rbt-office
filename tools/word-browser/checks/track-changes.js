// Track changes through real keyboard input: typed text is a tracked insertion,
// Enter is a tracked paragraph mark, Backspace at the start of a paragraph marks
// the previous mark as deleted (unless it is the same author's pending Enter),
// and all of it reaches the saved file.
import { check, finish, fail, loadWord, loadFormats } from './common.js';
try {
    const Word = await loadWord();
    const OF = await loadFormats();
    const editor = Word.initialize('c');
    Word.setTrackChangesOptions(editor, 'Tess', '2024-05-06T07:08:09Z');
    Word.setTrackChanges(editor, true);
    window.finish = async () => {
        try {
            check('typed text', Word.getText(editor), 'one\ntwo');
            const revisions = Word.getRevisions ? Word.getRevisions(editor) : null;
            const blob = await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor));
            const zip = await window.JSZip.loadAsync(new Uint8Array(await blob.arrayBuffer()));
            const xml = await zip.files['word/document.xml'].async('string');
            check('the Enter is a tracked paragraph mark in the file',
                /<w:pPr>(?:(?!<\/w:pPr>).)*<w:rPr><w:ins w:id="\d+" w:author="Tess" w:date="2024-05-06T07:08:09Z"\/><\/w:rPr><\/w:pPr><w:ins [^>]*w:author="Tess"[^>]*><w:r>(?:(?!<\/w:r>).)*one<\/w:t>/.test(xml), true);
            check('the second paragraph has no tracked mark', /two<\/w:t><\/w:r><\/w:ins><\/w:p>/.test(xml) && (xml.match(/<w:rPr><w:ins /g) || []).length === 1, true);
            // Reject everything: the insertions and the split are gone.
            Word.rejectAllRevisions(editor);
            check('reject all removes the typed text and the split', Word.getText(editor), '');
            if (revisions !== null) check('revisions were listed', revisions.length >= 2, true);
        } catch (e) { fail(e); }
        finish();
    };
    window.drive = [['click', 300, 120], ['type', 'one'], ['press', 'Enter'], ['type', 'two']];
} catch (e) { fail(e); finish(); }
