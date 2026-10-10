// A paragraph that was edited is saved as the XML it was read from with the
// change of text made in it (DocumentWriter: "a change of text made in the
// XML a block was read from"), not written from the model: what the model
// does not hold of it - a bookmark, a complex field, the anchor and tooltip
// of a hyperlink, a content control's properties, revision ids, w14:paraId,
// noProof / lang - is still there. With real keys: typing inside a run,
// deleting a selection, typing at the end of a paragraph, typing before a
// bookmarked run (a run of its own, before the bookmark), Enter (the first
// half keeps its XML). The saved document reads back with the
// editor's text and layout.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const bytes = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/edited-paragraph-kept.docx'));
    const original = await (await window.JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
    const editor = Word.initializeCanvas(document.getElementById('c'));
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
    const blocks = () => { const ids = []; for (const line of Word.getLayoutLines(editor)) if (!ids.includes(line.BlockId)) ids.push(line.BlockId); return ids; };
    const [first, second] = blocks();
    const caret = (id, from, to = from) => { Word.setSelection(editor, id, from, id, to); Word.getInputElement(editor).focus(); return []; };
    const saved = async () => {
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const xml = await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string');
        const again = Word.initializeCanvas(document.createElement('canvas'));
        Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(out));
        const same = [Word.getText(again) === Word.getText(editor), JSON.stringify(Word.getLayoutFingerprint(again)) === JSON.stringify(Word.getLayoutFingerprint(editor))];
        Word.dispose(again);
        return { xml, same };
    };
    const kept = ['w:rsidR="00AA0001" w:rsidRDefault="00AA0003" w14:paraId="2A2B2C2D"', '<w:pPr><w:suppressAutoHyphens/><w:rPr><w:lang w:val="de-DE"/></w:rPr></w:pPr>', '<w:bookmarkStart w:id="3" w:name="kept"/><w:r w:rsidRPr="00AA0002"><w:rPr><w:noProof/><w:lang w:val="de-DE"/></w:rPr>',
        '<w:bookmarkEnd w:id="3"/><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF kept \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>result</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>',
        '<w:hyperlink w:anchor="kept" w:tooltip="a tip" w:history="1">', '<w:sdtPr><w:alias w:val="An alias"/><w:tag w:val="a tag"/><w:id w:val="5"/><w:lock w:val="sdtLocked"/></w:sdtPr>'];
    const stage = async (label, expectXml) => {
        const { xml, same } = await saved();
        for (const [what, piece] of expectXml) check(`${label}: ${what}`, xml.includes(piece), true);
        check(`${label}: everything the model does not hold is kept`, kept.filter((piece) => !xml.includes(piece)), []);
        check(`${label}: read again, the editor's text and layout`, same, [true, true]);
        return xml;
    };
    window.stages = {
        typed: () => stage('a character typed inside a run', [['in its w:t', '<w:t xml:space="preserve">staXrt </w:t>']]),
        deleted: () => stage('a selection deleted', [['out of its w:t', '<w:t xml:space="preserve">sXrt </w:t>']]),
        atEnd: () => stage('a character typed at the end of the second paragraph', [['in its w:t', 'second paragraph!</w:t></w:r></w:p>']]),
        beforeBookmark: () => stage('a character typed before the bookmarked run', [['a run of its own with that run\'s properties, before the bookmark', '</w:pPr><w:r><w:rPr><w:noProof/><w:lang w:val="de-DE"/></w:rPr><w:t>Q</w:t></w:r><w:bookmarkStart w:id="3" w:name="kept"/>']]),
        entered: async () => {
            const xml = await stage('Enter in the second paragraph', [['the first half is the XML it was, without the text that went', '<w:t xml:space="preserve">second</w:t></w:r></w:p>']]);
            check('Enter: three paragraphs', (xml.match(/<\/w:p>/g) ?? []).length, 3);
        },
    };
    window.run = async (name) => { try { await window.stages[name](); } catch (e) { fail(e); } };
    window.drive = [
        // (One stretch of a paragraph is changed in its XML: the changes of a paragraph that are apart, in
        // runs formatted differently, are written from the model. So the first one is undone before the others.)
        ['eval', `(${caret})("${first}", 0)`], ['type', 'Q'], ['eval', 'window.run("beforeBookmark")'], ['press', 'Control+z'],
        ['then', `(() => { Word.setSelection(window.editor, "${first}", 3, "${first}", 3); return [["type", "X"]]; })()`], ['eval', 'window.run("typed")'],
        ['then', `(() => { Word.setSelection(window.editor, "${first}", 1, "${first}", 3); return [["press", "Delete"]]; })()`], ['eval', 'window.run("deleted")'],
        ['then', `(() => { Word.setSelection(window.editor, "${second}", 16, "${second}", 16); return [["type", "!"]]; })()`], ['eval', 'window.run("atEnd")'],
        ['then', `(() => { Word.setSelection(window.editor, "${second}", 6, "${second}", 6); return [["press", "Enter"]]; })()`], ['eval', 'window.run("entered")'],
    ];
    window.editor = editor; window.Word = Word;
    Word.getInputElement(editor).focus();
    window.finish = async () => {
        try {
            // Undo everything: an unedited save again - document.xml is the text that was read.
            while (Word.canUndo(editor)) Word.undo(editor);
            const { xml } = await saved();
            check('all undone: document.xml is the text that was read', xml === original, true);
        } catch (e) { fail(e); }
        Word.dispose(editor);
        finish();
    };
} catch (e) { fail(e); finish(); }
