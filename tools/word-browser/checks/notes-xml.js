// A part of notes or comments in which one changed is saved as the XML it was read from with that one written
// anew: the others - what the model does not hold of them (paragraph ids, proofing marks, a bookmark, the
// separators' spacing) - and the part's declaration and root are the text they were; with none changed the part
// is the text that was read. (synthetic/kept-notes.)
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const bytes = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/kept-notes.docx'));
    const zip = await window.JSZip.loadAsync(bytes);
    const footnotes = await zip.file('word/footnotes.xml').async('string');
    const comments = await zip.file('word/comments.xml').async('string');
    const elementOf = (xml, name, id) => { const at = xml.indexOf(`<${name} `, xml.indexOf(`w:id="${id}"`) - name.length - 30); return xml.slice(at, xml.indexOf(`</${name}>`, at) + name.length + 3); };
    const open = async () => {
        const editor = Word.initializeCanvas(document.createElement('canvas'));
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
        Word.completeLayout(editor);
        return editor;
    };
    const save = async (editor) => {
        const out = await window.JSZip.loadAsync(new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer()));
        return { footnotes: await out.file('word/footnotes.xml').async('string'), comments: await out.file('word/comments.xml').async('string') };
    };
    const blockWith = (editor, start) => Word.getLayoutLines(editor).find((line) => (line.Text ?? '').startsWith(start)).BlockId;

    // Nothing changed: both parts are the text that was read.
    let editor = await open();
    let saved = await save(editor);
    check('unchanged: footnotes.xml and comments.xml are the text that was read', [saved.footnotes === footnotes, saved.comments === comments], [true, true]);

    // The second footnote's text changed (setNote on its reference).
    const first = blockWith(editor, 'First note');
    const text = Word.getText(editor);
    const mark = text.indexOf(' second note') + ' second note'.length;
    Word.setSelection(editor, first, mark, first, mark + 1);
    Word.setNote(editor, '2', 'footnote', 'A changed note');
    saved = await save(editor);
    check('a footnote changed: it is written anew', saved.footnotes.includes('A changed note') && !saved.footnotes.includes('The second note'), true);
    check('a footnote changed: the other note and the separators are the text they were',
        [elementOf(footnotes, 'w:footnote', '-1'), elementOf(footnotes, 'w:footnote', '0'), elementOf(footnotes, 'w:footnote', '1')].map((piece) => saved.footnotes.includes(piece)), [true, true, true]);
    check('a footnote changed: the part\'s declaration and root are what they were',
        saved.footnotes.slice(0, saved.footnotes.indexOf('>', saved.footnotes.indexOf('<w:footnotes')) + 1) === footnotes.slice(0, footnotes.indexOf('>', footnotes.indexOf('<w:footnotes')) + 1), true);
    check('a footnote changed: the comments are the text that was read', saved.comments === comments, true);

    // The second comment's text changed (setComment on its range).
    editor = await open();
    const second = blockWith(editor, 'one');
    const line = Word.getLayoutLines(editor).find((candidate) => candidate.BlockId === second).Text;
    Word.setSelection(editor, second, line.indexOf('two'), second, line.indexOf('two') + 3);
    Word.setComment(editor, '4', 'A changed comment', 'Bo', '2020-01-02T00:00:00Z');
    saved = await save(editor);
    check('a comment changed: it is written anew', saved.comments.includes('A changed comment') && !saved.comments.includes('The other comment'), true);
    check('a comment changed: the other comment is the text it was', saved.comments.includes(elementOf(comments, 'w:comment', '3')), true);
    check('a comment changed: the footnotes are the text that was read', saved.footnotes === footnotes, true);
} catch (e) { fail(e); }
finish();
