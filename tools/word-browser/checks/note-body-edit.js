// A footnote's or a comment's body edited in place (enterNote / enterComment, a double click on a footnote where
// the page draws it): typing, Ctrl+End, Enter, Ctrl+B, undo work there with real keys; getCaretContext says the
// story; Escape / exitNote go back. What is saved is the part as it was read with the edit made in the note's
// paragraph and nothing else: the other notes, the separators, the note's other paragraphs and its mark run
// (w:footnoteRef, w:annotationRef) are the text they were, and the body is the text it was. (synthetic/kept-notes;
// an endnote, by the API and a double click where the endnotes follow the text: synthetic/rich-notes.)
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
let OF, Word, editor;
const bytes = {};
const original = {};
const richEndnotes = {};
const saved = {};
const seen = {};
const shift = 500;
try {
    OF = await loadFormats();
    Word = await loadWord();
    bytes.kept = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/kept-notes.docx'));
    bytes.rich = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/rich-notes.docx'));
    const zip = await window.JSZip.loadAsync(bytes.kept);
    for (const name of ['footnotes', 'comments', 'document']) original[name] = await zip.file(`word/${name}.xml`).async('string');
    const richZip = await window.JSZip.loadAsync(bytes.rich);
    for (const name of ['endnotes', 'footnotes', 'document']) richEndnotes[name] = await richZip.file(`word/${name}.xml`).async('string');
    const canvas = document.getElementById('c');
    canvas.style.position = 'absolute'; canvas.style.left = '0px';
    const context = () => { const c = Word.getCaretContext(editor); return [c.Story, c.NoteId ?? null, c.CommentId ?? null]; };
    // (What a step throws is a failure of the check, not of the run.)
    const guarded = (step) => async (...args) => { try { await step(...args); } catch (e) { fail(e); } };
    window.place = guarded(async (stage, story, fixture = 'kept') => {
        if (editor) Word.dispose(editor);
        // (Moved up for the first document: the foot of its first page - its footnotes - is in the window.)
        canvas.style.top = fixture === 'kept' ? `-${shift}px` : '0px';
        editor = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes[fixture]));
        Word.completeLayout(editor);
        const entered = story === 'comment' ? Word.enterComment(editor, '4') : story === 'footnote' ? Word.enterNote(editor, '2', 'footnote') : story === 'endnote' ? Word.enterNote(editor, '1', 'endnote') : null;
        seen[stage] = { entered, context: context() };
        Word.getInputElement(editor).focus();
    });
    // Where the page draws the second footnote (its line in the footnote area), the endnote (after the text).
    window.noteLine = () => Word.getLayoutLines(editor).find((line) => String(line.BlockId).startsWith('__footnote__') && (line.Text ?? '').includes('second note'));
    window.endnoteLine = () => Word.getLayoutLines(editor).find((line) => String(line.BlockId).startsWith('__endnote__') && (line.Text ?? '').includes('italic endnote'));
    // Dark pixels of the canvas in a rectangle (layout coordinates: the canvas holds the whole document).
    const inkAt = (x, y, w, h) => { const data = canvas.getContext('2d').getImageData(x, y, w, h).data; let ink = 0; for (let i = 0; i < data.length; i += 4) if (data[i] < 100 && data[i + 1] < 100 && data[i + 2] < 100) ink++; return ink; };
    window.look = (stage) => {
        const caret = Word.getCaretRect(editor);
        seen[stage] = Object.assign(seen[stage] ?? {}, { context: context(), caret, text: Word.getText(editor), ink: caret ? inkAt(Math.round(caret.X), Math.round(caret.Y), 200, Math.round(caret.Height)) : -1 });
    };
    window.snap = guarded(async (stage) => {
        seen[stage] = Object.assign(seen[stage] ?? {}, { before: context(), exited: Word.exitNote(editor), after: context() });
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const savedZip = await window.JSZip.loadAsync(out);
        const part = async (name) => { const file = savedZip.file(`word/${name}.xml`); return file ? file.async('string') : null; };
        const again = Word.initializeCanvas(document.createElement('canvas'));
        Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(out));
        const texts = (e) => { const pkg = Word.exportDocumentPackage(e); return JSON.stringify([...pkg.Footnotes, ...pkg.Comments].map((item) => [String(item.Id), item.Text])); };
        saved[stage] = { footnotes: await part('footnotes'), endnotes: await part('endnotes'), comments: await part('comments'), document: await part('document'), texts: texts(editor), readAgain: texts(again) === texts(editor) };
        Word.dispose(again);
    });
    window.drive = [
        ['eval', 'window.place("type", "footnote")'], ['eval', 'window.look("enterNote")'], ['type', 'X'], ['eval', 'window.look("typed")'], ['eval', 'window.snap("type")'],
        ['eval', 'window.place("end", "footnote")'], ['press', 'Control+End'], ['type', 'Z'], ['eval', 'window.snap("end")'],
        ['eval', 'window.place("comment", "comment")'], ['type', 'Y'], ['eval', 'window.snap("comment")'],
        ['eval', 'window.place("undo", "footnote")'], ['type', 'abc'], ['press', 'Control+z'], ['eval', 'window.look("undone")'], ['press', 'Control+y'], ['press', 'Control+z'], ['eval', 'window.snap("undo")'],
        ['eval', 'window.place("enter", "footnote")'], ...Array(4).fill(['press', 'ArrowRight']), ['press', 'Enter'], ['eval', 'window.snap("enter")'],
        ['eval', 'window.place("bold", "footnote")'], ['press', 'ArrowRight'], ...Array(3).fill(['press', 'Shift+ArrowRight']), ['press', 'Control+b'], ['eval', 'window.snap("bold")'],
        ['eval', 'window.place("dblclick", null)'],
        ['then', `window.noteLine() ? [["dblclick", Math.round(window.noteLine().X + window.noteLine().Width - 4), Math.round(window.noteLine().Y + window.noteLine().Height / 2) - ${shift}]] : []`],
        ['eval', 'window.look("dblclick")'], ['type', 'Q'], ['press', 'Escape'], ['eval', 'window.look("escaped")'], ['eval', 'window.snap("dblclick")'],
        ['eval', 'window.place("endnote", "endnote", "rich")'], ['type', 'E'], ['eval', 'window.snap("endnote")'],
        ['eval', 'window.place("endnoteclick", null, "rich")'],
        ['then', 'window.endnoteLine() ? [["dblclick", Math.round(window.endnoteLine().X + window.endnoteLine().Width - 4), Math.round(window.endnoteLine().Y + window.endnoteLine().Height / 2)]] : []'],
        ['eval', 'window.look("endnoteclick")'], ['press', 'End'], ['type', 'W'], ['eval', 'window.snap("endnoteclick")'],
    ];
} catch (e) { fail(e); finish(); }
const elementOf = (xml, name, id) => { const at = xml.indexOf(`<${name} `, xml.indexOf(`w:id="${id}"`) - name.length - 30); return xml.slice(at, xml.indexOf(`</${name}>`, at) + name.length + 3); };
const paragraphOf = (xml, paraId) => { const at = xml.lastIndexOf('<w:p ', xml.indexOf(`w14:paraId="${paraId}"`)); return xml.slice(at, xml.indexOf('</w:p>', at) + 6); };
const count = (text, piece) => text.split(piece).length - 1;
const footnoteMark = '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>';
window.finish = async () => {
    try {
        const others = (xml) => [elementOf(original.footnotes, 'w:footnote', '-1'), elementOf(original.footnotes, 'w:footnote', '0'), elementOf(original.footnotes, 'w:footnote', '1')].map((piece) => xml.includes(piece));

        // enterNote, a character typed at the start of the note's text.
        check('enterNote: true, the caret in the footnote', seen.type.entered === true && seen.type.context, ['footnote', '2', null]);
        check('enterNote: the caret is at the note (in the footnote area, under the body text)', seen.enterNote.caret !== null && seen.enterNote.caret.Y > 300, true);
        check('typed in the note: its text is painted after the caret', seen.typed.ink > 40 ? true : seen.typed.ink, true);
        check('typed in the note: the editor\'s text is the note\'s', seen.typed.text.includes('X The second note') && !seen.typed.text.includes('First note'), true);
        let { footnotes, comments, document } = saved.type;
        check('typed in the note: exitNote goes back to the body', [seen.type.exited, seen.type.after], [true, ['body', null, null]]);
        check('typed in the note: footnotes.xml is the text that was read with the character in its run',
            footnotes === original.footnotes.replace('> The second note<', '>X The second note<') ? true : footnotes, true);
        check('typed in the note: comments.xml and document.xml are the text that was read', [comments === original.comments, document === original.document], [true, true]);
        check('typed in the note: read again, the same notes and comments', saved.type.readAgain, true);

        // Ctrl+End: the end of the note's second paragraph.
        check('Ctrl+End in the note: the character at the end of its last paragraph, nothing else changed',
            saved.end.footnotes === original.footnotes.replace('> ends<', '> endsZ<') ? true : saved.end.footnotes, true);
        check('Ctrl+End in the note: the body is the text it was', saved.end.document === original.document, true);

        // enterComment.
        check('enterComment: true, the caret in the comment', seen.comment.entered === true && seen.comment.context, ['comment', null, '4']);
        check('typed in the comment: comments.xml is the text that was read with the character in its run',
            saved.comment.comments === original.comments.replace('>The other comment<', '>YThe other comment<') ? true : saved.comment.comments, true);
        check('typed in the comment: footnotes.xml and document.xml are the text that was read', [saved.comment.footnotes === original.footnotes, saved.comment.document === original.document], [true, true]);
        check('typed in the comment: read again, the same notes and comments', saved.comment.readAgain, true);

        // Undo and redo in the note.
        check('undo in the note: the caret stays in the note, its text as it was', [seen.undone.context, seen.undone.text.includes(' The second note') && !seen.undone.text.includes('abc')], [['footnote', '2', null], true]);
        check('undo in the note: the parts are the text that was read', [saved.undo.footnotes === original.footnotes, saved.undo.document === original.document], [true, true]);

        // Enter in the note: two paragraphs made of the first; the mark run once, in the first.
        let element = elementOf(saved.enter.footnotes, 'w:footnote', '2');
        check('Enter in the note: the note is its paragraph split in two (the first with its id and mark run) and its second paragraph as it was',
            element === `<w:footnote w:id="2"><w:p w14:paraId="1A000004">${footnoteMark}<w:r><w:t xml:space="preserve"> The</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve"> second note</w:t></w:r></w:p>${paragraphOf(original.footnotes, '1A000005')}</w:footnote>` ? true : element, true);
        check('Enter in the note: the text is split, the second paragraph and the other notes are the text they were',
            [element.includes('> The<'), element.includes('> second note<'), element.includes(paragraphOf(original.footnotes, '1A000005')), ...others(saved.enter.footnotes)], [true, true, true, true, true, true]);
        check('Enter in the note: the mark run once, before the text', [count(element, footnoteMark), element.indexOf(footnoteMark) < element.indexOf('> The<')], [1, true]);
        check('Enter in the note: the comments and the body are the text they were', [saved.enter.comments === original.comments, saved.enter.document === original.document], [true, true]);

        // Ctrl+B on a word of the note.
        element = elementOf(saved.bold.footnotes, 'w:footnote', '2');
        check('Ctrl+B in the note: the word is a bold run of its own between the halves of the run it was in, nothing else changed',
            saved.bold.footnotes === original.footnotes.replace('<w:r><w:t xml:space="preserve"> The second note</w:t></w:r>', '<w:r><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">The</w:t></w:r><w:r><w:t xml:space="preserve"> second note</w:t></w:r>') ? true : element, true);
        check('Ctrl+B in the note: the word is bold, the mark run and the second paragraph as they were',
            [element.includes('<w:b/>'), count(element, footnoteMark), element.includes(paragraphOf(original.footnotes, '1A000005')), ...others(saved.bold.footnotes)], [true, 1, true, true, true, true]);

        // A double click on the footnote where the page draws it.
        check('a double click on the footnote: the caret in it', seen.dblclick.context, ['footnote', '2', null]);
        check('Escape: back in the body', seen.escaped.context, ['body', null, null]);
        element = elementOf(saved.dblclick.footnotes, 'w:footnote', '2');
        check('typed after the double click: in the note; the rest as it was',
            [element.includes('Q'), element.includes(paragraphOf(original.footnotes, '1A000005')), ...others(saved.dblclick.footnotes), saved.dblclick.document === original.document], [true, true, true, true, true, true]);

        // An endnote: by the API, and by a double click where the endnotes follow the text.
        check('enterNote of an endnote: true, the caret in it', seen.endnote.entered === true && seen.endnote.context, ['endnote', '1', null]);
        check('typed in the endnote: endnotes.xml is the text that was read with the character in its (italic) run; the footnotes and the body as they were',
            [saved.endnote.endnotes === richEndnotes.endnotes.replace('> italic endnote<', '>E italic endnote<') ? true : saved.endnote.endnotes, saved.endnote.footnotes === richEndnotes.footnotes, saved.endnote.document === richEndnotes.document], [true, true, true]);
        check('a double click on the endnote: the caret in it', seen.endnoteclick.context, ['endnote', '1', null]);
        check('typed at the end of the endnote after the double click',
            saved.endnoteclick.endnotes === richEndnotes.endnotes.replace('> italic endnote<', '> italic endnoteW<') ? true : saved.endnoteclick.endnotes, true);
    } catch (e) { fail(e); }
    finish();
};
