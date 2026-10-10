// A paragraph whose formatting was changed as a whole - all of it made bold,
// italic, struck through, underlined (or not), its alignment - is saved as
// the XML it was read from with the properties set in it (w:b and its like
// in the w:rPr of its runs, w:jc in its w:pPr, each where the schema puts
// it), not written from the model: what the model does not hold of it - a
// bookmark, a complex field, the anchor and tooltip of a hyperlink, a content
// control's properties, revision ids, w14:paraId, noProof / lang - is still
// there. The saved document reads back as the editor's. A part of a run made
// bold is still written from the model.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const bytes = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/edited-paragraph-kept.docx'));
    const original = await (await window.JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
    const editor = Word.initializeCanvas(document.getElementById('c'));
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
    const ids = []; for (const line of Word.getLayoutLines(editor)) if (!ids.includes(line.BlockId)) ids.push(line.BlockId);
    const [first, second] = ids;
    const firstText = 'start result link control', secondText = 'second paragraph';
    const whole = (id, text) => Word.setSelection(editor, id, 0, id, text.length);
    const modelOf = (e) => JSON.stringify(Word.exportDocumentPackage(e).Paragraphs.map((para) => [para.Alignment, para.ExplicitAlignment ?? null, (para.Runs ?? []).map((run) => [run.Text, run.Bold, run.Italic, run.Underline, run.Strike, run.ExplicitBold ?? null, run.ExplicitItalic ?? null, run.ExplicitUnderline ?? null, run.ExplicitStrike ?? null])]));
    // (The layout, with what it says of each run. An underline the editor made has no style of its own; read back it is "single".)
    const layoutText = (e) => JSON.stringify(Word.getLayoutFingerprint(e)).replace(/,single\//g, ',-/');
    const saved = async () => {
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const xml = await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string');
        const again = Word.initializeCanvas(document.createElement('canvas'));
        Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(out));
        const same = [Word.getText(again) === Word.getText(editor), modelOf(again) === modelOf(editor), layoutText(again) === layoutText(editor)];
        Word.dispose(again);
        return { xml, same };
    };
    // What the model does not hold of the first paragraph (the properties of its runs aside: see each stage).
    const kept = ['<w:p w:rsidR="00AA0001" w:rsidRDefault="00AA0003" w14:paraId="2A2B2C2D">', '<w:suppressAutoHyphens/>', '<w:rPr><w:lang w:val="de-DE"/></w:rPr></w:pPr>', '<w:bookmarkStart w:id="3" w:name="kept"/><w:r w:rsidRPr="00AA0002">', '<w:bookmarkEnd w:id="3"/>',
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF kept \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>', '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
        '<w:hyperlink w:anchor="kept" w:tooltip="a tip" w:history="1">', '<w:sdtPr><w:alias w:val="An alias"/><w:tag w:val="a tag"/><w:id w:val="5"/><w:lock w:val="sdtLocked"/></w:sdtPr>', '<w:noProof/><w:lang w:val="de-DE"/></w:rPr><w:t xml:space="preserve">start </w:t>'];
    const secondXml = original.slice(original.lastIndexOf('<w:p>', original.indexOf(secondText)), original.indexOf('</w:p>', original.indexOf(secondText)) + 6);
    const stage = async (label, expectXml, otherParagraph = true) => {
        const { xml, same } = await saved();
        for (const [what, piece] of expectXml) check(`${label}: ${what}`, xml.includes(piece), true);
        check(`${label}: everything the model does not hold is kept`, kept.filter((piece) => !xml.includes(piece)), []);
        if (otherParagraph) check(`${label}: the other paragraph is the text it was`, xml.includes(secondXml), true);
        check(`${label}: read again, the editor's text, formatting and layout`, same, [true, true, true]);
        return xml;
    };
    check('the fixture', [Word.getText(editor).startsWith(firstText), secondXml.includes(secondText)], [true, true]);

    // All of the paragraph made bold: w:b first in every run's properties, a w:rPr for the run that had none.
    whole(first, firstText);
    Word.toggleBold(editor);
    await stage('the paragraph made bold', [
        ['in the properties of the bookmarked run, before what it had', '<w:r w:rsidRPr="00AA0002"><w:rPr><w:b/><w:noProof/><w:lang w:val="de-DE"/></w:rPr><w:t xml:space="preserve">start </w:t></w:r>'],
        ['in those of the field\'s result', '<w:r><w:rPr><w:b/><w:noProof/></w:rPr><w:t>result</w:t></w:r>'],
        ['before the underline of the link\'s run', '<w:r><w:rPr><w:b/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve"> link</w:t></w:r>'],
        ['properties for the run that had none', '<w:sdtContent><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> control</w:t></w:r></w:sdtContent>'],
    ]);
    // Italic as well, and centred: both in one save, w:i after w:b, w:jc after what the paragraph's properties had before its place.
    Word.toggleItalic(editor);
    Word.alignCenter(editor);
    await stage('italic as well, and centred', [
        ['w:i after w:b', '<w:rPr><w:b/><w:i/><w:noProof/><w:lang w:val="de-DE"/></w:rPr><w:t xml:space="preserve">start </w:t>'],
        ['w:jc in the paragraph\'s properties, before the mark\'s', '<w:pPr><w:suppressAutoHyphens/><w:jc w:val="center"/><w:rPr><w:lang w:val="de-DE"/></w:rPr></w:pPr>'],
    ]);
    // Not bold again: stated, as the writer states it.
    Word.toggleBold(editor);
    await stage('not bold again', [
        ['w:b with the value 0', '<w:rPr><w:b w:val="0"/><w:i/><w:noProof/><w:lang w:val="de-DE"/></w:rPr><w:t xml:space="preserve">start </w:t>'],
        ['in the run that got its properties here', '<w:rPr><w:b w:val="0"/><w:i/></w:rPr><w:t xml:space="preserve"> control</w:t>'],
    ]);
    // Struck through: w:strike where the schema puts it, between w:i and noProof.
    Word.toggleStrike(editor);
    await stage('struck through', [
        ['w:strike after w:i, before noProof', '<w:rPr><w:b w:val="0"/><w:i/><w:strike/><w:noProof/><w:lang w:val="de-DE"/></w:rPr><w:t xml:space="preserve">start </w:t>'],
        ['before the underline the link\'s run has', '<w:rPr><w:b w:val="0"/><w:i/><w:strike/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve"> link</w:t>'],
    ]);
    // Everything undone: the document that was opened.
    while (Word.canUndo(editor)) Word.undo(editor);
    check('undone: the file that was opened', [(await saved()).xml === original], [true]);

    // The second paragraph alone: the first one is the text it was.
    whole(second, secondText);
    Word.alignCenter(editor);
    Word.toggleBold(editor);
    Word.toggleUnderline(editor);
    {
        const { xml, same } = await saved();
        check('a paragraph without properties: w:pPr and w:rPr are made for it', [xml.includes('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">second paragraph</w:t></w:r></w:p>'), xml.split('<w:p>second')[0].includes(original.slice(original.indexOf('<w:p w:rsidR'), original.indexOf('</w:p>') + 6)), same], [true, true, [true, true, true]]);
    }
    while (Word.canUndo(editor)) Word.undo(editor);

    // Both paragraphs selected and made bold: each in its own XML.
    Word.setSelection(editor, first, 0, second, secondText.length);
    Word.toggleBold(editor);
    await stage('both paragraphs made bold', [
        ['the first', '<w:rPr><w:b/><w:noProof/><w:lang w:val="de-DE"/></w:rPr><w:t xml:space="preserve">start </w:t>'],
        ['the second', '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">second paragraph</w:t></w:r></w:p>'],
    ], false);
    while (Word.canUndo(editor)) Word.undo(editor);

    // A part of a run made bold is a change this is not done for: the block is written from the model
    // (and still reads back as the editor's).
    Word.setSelection(editor, first, 1, first, 3);
    Word.toggleBold(editor);
    {
        const { xml, same } = await saved();
        check('a part of a run made bold: written from the model, read back as the editor\'s', [same, xml.includes('w14:paraId="2A2B2C2D"')], [[true, true, true], false]);
        note('a part of a run made bold is written from the model: the paragraph loses what the model does not hold');
    }
} catch (e) { fail(e); }
finish();
