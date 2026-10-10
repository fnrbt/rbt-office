// The mark of a note or a comment has the formatting of its run: the reader
// reads the run's properties (a mark that is bold, or has a colour, is that
// in the editor), and a block written from the model saves what was set on a
// mark - without the colour the editor shows a comment's mark in, which is
// not the file's.
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const bytes = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/rich-notes.docx'));
    const marksOf = (pkg) => pkg.Paragraphs[0].Runs.filter((run) => run.NoteId || run.CommentId).map((run) => [run.NoteId ? 'note' : 'comment', run.Bold === true, run.Italic === true, run.Color ? String(run.Color).toLowerCase() : null]);

    // Read: the properties of the marks' runs.
    const zip = await window.JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml').async('string');
    const formatted = xml
        .replace('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference', '<w:r><w:rPr><w:b/><w:color w:val="FF0000"/><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference')
        .replace('<w:r><w:commentReference', '<w:r><w:rPr><w:i/></w:rPr><w:commentReference');
    check('the fixture has the marks', formatted !== xml && formatted.includes('<w:i/></w:rPr><w:commentReference'), true);
    zip.file('word/document.xml', formatted);
    const read = await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' }));
    check('read: a bold red footnote mark, a plain endnote mark, an italic comment mark in the display colour', marksOf(read),
        [['note', true, false, '#ff0000'], ['note', false, false, null], ['comment', false, true, '#7c3aed']]);

    // Saved from the model: the paragraph made bold, marks and all.
    const editor = Word.initializeCanvas(document.getElementById('c'));
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
    Word.completeLayout(editor);
    const lines = Word.getLayoutLines(editor);
    const id = lines[0].BlockId;
    const end = Math.max(...lines.filter((line) => line.BlockId === id).map((line) => (line.StartOffset ?? 0) + (line.Text ?? '').length));
    Word.setSelection(editor, id, 0, id, end);
    Word.toggleBold(editor);
    const exported = Word.exportDocumentPackage(editor);
    exported.RegenerateAll = true;
    const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(exported)).arrayBuffer());
    const saved = await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string');
    check('written from the model: the comment\'s mark is bold', /<w:r><w:rPr>(?:(?!<\/w:rPr>).)*<w:b\/>(?:(?!<\/w:rPr>).)*<\/w:rPr><w:commentReference w:id="5"\/><\/w:r>/.test(saved), true);
    check('written from the model: the display colour of the comment\'s mark is not in the file', /7c3aed/i.test(saved), false);
    const again = await OF.processors.forEditorPackage(out);
    check('read again: every mark is bold', marksOf(again).map((mark) => mark.slice(0, 2)), [['note', true], ['note', true], ['comment', true]]);
    check('read again: the comment\'s mark is shown in the display colour', marksOf(again)[2][3], '#7c3aed');
} catch (e) { fail(e); }
finish();
