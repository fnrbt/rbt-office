// A part of notes or comments in which one changed is saved as the XML it was read from with that one written
// anew: the others - what the model does not hold of them (paragraph ids, proofing marks, a bookmark, the
// separators' spacing) - and the part's declaration and root are the text they were; with none changed the part
// is the text that was read. (synthetic/kept-notes.)
// A note or comment that changed is written from its plan, as the body is: its element's start tag and the
// paragraphs of it that are as they were read are their text, those that changed are written in their place;
// the run of its own mark (w:footnoteRef, w:annotationRef) is the run that was read. (The editor changes a note
// only as a whole - setNote - so these are written from a package whose note says which of its blocks changed,
// as the editor's tagging says of a header's.)
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

    // A package in which one paragraph of a note or comment changed: the blocks say which they were when read.
    const writeWith = async (part, id, changed, text) => {
        const editor = await open();
        const pkg = Word.exportDocumentPackage(editor);
        const story = pkg[part].find((item) => String(item.Id) === id);
        story.SourceUnchanged = false;
        story.SourceStoryCounts = story.Blocks.map(() => 1);
        story.Blocks.forEach((block, index) => { block.SourceIndex = index; block.SourceUnchanged = index !== changed; });
        const block = story.Blocks[changed];
        block.Runs = [Object.assign({}, block.Runs[block.Runs.length - 1], { Text: text, SourceRun: undefined })];
        pkg[part === 'Comments' ? 'CommentsUnchanged' : 'FootnotesUnchanged'] = false;
        pkg.Unchanged = false;
        const out = await window.JSZip.loadAsync(new Uint8Array(await (await OF.word.writeFromEditorPackage(pkg)).arrayBuffer()));
        return { footnotes: await out.file('word/footnotes.xml').async('string'), comments: await out.file('word/comments.xml').async('string') };
    };
    const paragraphOf = (xml, paraId) => { const at = xml.lastIndexOf('<w:p ', xml.indexOf(`w14:paraId="${paraId}"`)); return xml.slice(at, xml.indexOf('</w:p>', at) + 6); };
    const count = (text, piece) => text.split(piece).length - 1;
    const footnoteMark = '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>';
    const commentMark = '<w:r><w:annotationRef/></w:r>';
    const note2 = elementOf(footnotes, 'w:footnote', '2');
    const comment4 = elementOf(comments, 'w:comment', '4');

    saved = await writeWith('Footnotes', '2', 1, 'A rewritten paragraph');
    let element = elementOf(saved.footnotes, 'w:footnote', '2');
    check('second paragraph of a note changed: the first (with the mark) is the text it was, the second written anew',
        [element.startsWith('<w:footnote w:id="2">' + paragraphOf(note2, '1A000004')), element.includes('A rewritten paragraph'), element.includes('Its second paragraph')], [true, true, false]);
    check('second paragraph of a note changed: the other notes and the separators are the text they were',
        [elementOf(footnotes, 'w:footnote', '-1'), elementOf(footnotes, 'w:footnote', '0'), elementOf(footnotes, 'w:footnote', '1')].map((piece) => saved.footnotes.includes(piece)), [true, true, true]);

    saved = await writeWith('Footnotes', '2', 0, 'A rewritten first paragraph');
    element = elementOf(saved.footnotes, 'w:footnote', '2');
    check('first paragraph of a note changed: its mark run is the one read, once, before the new text; the second paragraph is the text it was',
        [count(element, footnoteMark), element.indexOf(footnoteMark) >= 0 && element.indexOf(footnoteMark) < element.indexOf('A rewritten first paragraph'), element.includes(paragraphOf(note2, '1A000005')), element.includes('The second note')],
        [1, true, true, false]);

    saved = await writeWith('Comments', '4', 1, 'A rewritten comment paragraph');
    element = elementOf(saved.comments, 'w:comment', '4');
    check('second paragraph of a comment changed: its start tag and first paragraph are the text they were',
        [element.startsWith(comment4.slice(0, comment4.indexOf('>') + 1) + paragraphOf(comment4, '2B000002')), element.includes('A rewritten comment paragraph'), element.includes('And its second paragraph')], [true, true, false]);
    check('second paragraph of a comment changed: the other comment and the footnotes are the text they were',
        [saved.comments.includes(elementOf(comments, 'w:comment', '3')), saved.footnotes === footnotes], [true, true]);

    saved = await writeWith('Comments', '4', 0, 'A rewritten first comment paragraph');
    element = elementOf(saved.comments, 'w:comment', '4');
    check('first paragraph of a comment changed: the annotation mark run is the one read, once, before the new text; the second paragraph is the text it was',
        [count(element, commentMark), element.indexOf(commentMark) >= 0 && element.indexOf(commentMark) < element.indexOf('A rewritten first comment paragraph'), element.includes(paragraphOf(comment4, '2B000003'))],
        [1, true, true]);
} catch (e) { fail(e); }
finish();
