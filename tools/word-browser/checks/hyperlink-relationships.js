// A paragraph with a hyperlink that is saved after a change keeps the link's
// relationship: the id the file had for the target, no second relationship
// for it, and none left behind that nothing refers to. A link that is taken
// away takes its relationship with it. A paragraph written from the model has
// no empty w:pPr.
//
// (tests/test-documents-wild/phpword-sample-11-readword2007.docx: the link of
// its first paragraph is rId8.)
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const bytes = await bytesOf('/repo/tests/test-documents-wild/phpword-sample-11-readword2007.docx');
    const partsOf = async (data) => {
        const zip = await window.JSZip.loadAsync(data);
        return { xml: await zip.file('word/document.xml').async('string'), rels: await zip.file('word/_rels/document.xml.rels').async('string') };
    };
    const original = await partsOf(bytes);
    const editor = Word.initializeCanvas(document.getElementById('c'));
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
    Word.completeLayout(editor);
    const lineOf = (text) => Word.getLayoutLines(editor).find((line) => (line.Text ?? '').includes(text));
    const select = (text, word) => {
        const line = lineOf(text);
        const at = (line.StartOffset ?? 0) + line.Text.indexOf(text) + text.indexOf(word);
        Word.setSelection(editor, line.BlockId, at, line.BlockId, at + word.length);
    };
    const save = async () => partsOf(new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer()));
    // The relationships of hyperlinks: [id, target], and the ids the XML does not use.
    const links = (rels) => [...rels.matchAll(/<Relationship\b[^>]*>/g)].map((m) => m[0]).filter((tag) => tag.includes('/relationships/hyperlink"'))
        .map((tag) => [/\bId="([^"]*)"/.exec(tag)[1], /Target="([^"]*)"/.exec(tag)[1]]);
    const orphans = (parts) => links(parts.rels).filter(([id]) => !parts.xml.includes(`"${id}"`)).map(([id]) => id);
    const linkIds = (xml) => [...xml.matchAll(/<w:hyperlink\b[^>]*\br:id="([^"]*)"/g)].map((m) => m[1]);
    const paragraphOf = (xml, text) => { const at = xml.indexOf(text); return xml.slice(Math.max(xml.lastIndexOf('<w:p ', at), xml.lastIndexOf('<w:p>', at)), xml.indexOf('</w:p>', at) + 6); };
    const target = 'https://github.com/PHPOffice/PHPWord';
    check('the fixture: the link of the first paragraph is rId8', [links(original.rels), linkIds(original.xml)], [[['rId8', target]], ['rId8']]);

    // One word of the paragraph made bold.
    select('PHPWord is a', 'is');
    Word.toggleBold(editor);
    let saved = await save();
    check('a word made bold: the link refers to the relationship it had', linkIds(saved.xml), ['rId8']);
    check('a word made bold: the relationships of links are the file\'s', links(saved.rels), [['rId8', target]]);
    check('a word made bold: no relationship without a referrer', orphans(saved), []);
    check('a word made bold: no empty w:pPr', saved.xml.includes('<w:pPr/>'), false);
    check('a word made bold: the paragraph is the XML it was read from, with the change (its ids, its link)', [paragraphOf(saved.xml, 'PHPWord').includes('w14:paraId="21C0B3CD" w14:textId="77777777" w:rsidR="000B6ACB"'), saved.rels === original.rels], [true, true]);
    check('a word made bold: the word is bold', /<w:b\/>(<[^>]*>)*<w:t[^>]*>\s?is\s?<\/w:t>/.test(paragraphOf(saved.xml, 'PHPWord')), true);

    // A bookmark set on a word: the paragraph is written from its source, with the bookmark (docs/save-architecture.md).
    select('PHPWord is a', 'PHPWord');
    Word.setBookmark(editor, 'mark');
    saved = await save();
    check('a bookmark set: the paragraph is the XML it was read from, with the bookmark', [paragraphOf(saved.xml, 'PHPWord').includes('w:name="mark"'), paragraphOf(saved.xml, 'PHPWord').includes('w14:paraId')], [true, true]);
    check('a bookmark set: the link refers to the relationship it had', linkIds(saved.xml), ['rId8']);
    check('a bookmark set: the relationships of links are the file\'s', links(saved.rels), [['rId8', target]]);
    check('a bookmark set: no relationship without a referrer', orphans(saved), []);
    check('a bookmark set: no empty w:pPr', saved.xml.includes('<w:pPr/>'), false);
    const again = Word.initializeCanvas(document.createElement('canvas'));
    Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer())));
    check('a bookmark set: read again, the text', Word.getText(again) === Word.getText(editor), true);
    Word.dispose(again);

    // The link taken away: its relationship goes with it.
    const linkLine = Word.getLayoutLines(editor).find((line) => line.BlockId === lineOf('PHPWord is a').BlockId);
    const all = Word.getLayoutLines(editor).filter((line) => line.BlockId === linkLine.BlockId);
    const end = Math.max(...all.map((line) => (line.StartOffset ?? 0) + (line.Text ?? '').length));
    Word.setSelection(editor, linkLine.BlockId, 0, linkLine.BlockId, end);
    Word.setHyperlink(editor, '');
    saved = await save();
    check('the link taken away: no link in the paragraph, no relationship of a link', [linkIds(saved.xml), links(saved.rels)], [[], []]);
} catch (e) { fail(e); }
finish();
