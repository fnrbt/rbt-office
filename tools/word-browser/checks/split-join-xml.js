// A paragraph split in two (Enter) is saved as two paragraphs made of the XML
// that was read: each with the paragraph's properties and the runs, proofing
// marks, smart tag, field, comment range and link of its part of the text;
// the run the caret was in is split; the second half has no paragraph id.
// Paragraphs joined (Backspace at the start of the second) are the first with
// the second's content - its bookmark, its run's revision id - at its end. A
// split undone by hand is the paragraph that was read. In a table the halves
// are made of the paragraph's XML too; the table's properties, its rows and
// its other cells go back as they were read.
// (synthetic/patched-paragraphs; real Enter and Backspace keys.)
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
let OF, Word, editor, original;
const text = 'alpha beta gamma delta\tzeta\netaresult theta link one link two';
const saved = {};
try {
    OF = await loadFormats();
    Word = await loadWord();
    const bytes = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/patched-paragraphs.docx'));
    original = await (await window.JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
    const open = async () => {
        if (editor) Word.dispose(editor);
        editor = Word.initializeCanvas(document.getElementById('c'));
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
        Word.completeLayout(editor);
        Word.getInputElement(editor).focus();
    };
    const blockStarting = (start, nth = 0) => {
        const ids = [];
        for (const line of Word.getLayoutLines(editor)) if (!ids.includes(line.BlockId) && (line.Text ?? '').startsWith(start) && (line.StartOffset ?? 0) === 0) ids.push(line.BlockId);
        return ids[nth];
    };
    const layoutText = (e) => JSON.stringify(Word.getLayoutFingerprint(e)).replace(/,single\//g, ',-/');
    // Where the caret goes before the key, and what is kept of the document after it.
    window.place = async (stage) => {
        if (stage !== 'rejoin') await open();
        if (stage === 'split') { const id = blockStarting('alpha'); Word.setSelection(editor, id, 10, id, 10); }
        if (stage === 'rejoin') { const id = blockStarting(' gamma'); Word.setSelection(editor, id, 0, id, 0); }
        if (stage === 'join') { const id = blockStarting('second of two'); Word.setSelection(editor, id, 0, id, 0); }
        if (stage === 'cell') { const id = blockStarting('alpha', 1); Word.setSelection(editor, id, 10, id, 10); }
        Word.getInputElement(editor).focus();
    };
    window.snap = async (stage) => {
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const report = OF.word.getLastSaveReport();
        const xml = await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string');
        const again = Word.initializeCanvas(document.createElement('canvas'));
        Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(out));
        saved[stage] = { xml, report, text: Word.getText(editor), same: [Word.getText(again) === Word.getText(editor), layoutText(again) === layoutText(editor)] };
        Word.dispose(again);
    };
    window.drive = [
        ['eval', 'window.place("split")'], ['press', 'Enter'], ['eval', 'window.snap("split")'],
        ['eval', 'window.place("rejoin")'], ['press', 'Backspace'], ['eval', 'window.snap("rejoin")'],
        ['eval', 'window.place("join")'], ['press', 'Backspace'], ['eval', 'window.snap("join")'],
        ['eval', 'window.place("cell")'], ['press', 'Enter'], ['eval', 'window.snap("cell")'],
    ];
} catch (e) { fail(e); finish(); }
window.finish = async () => {
    try {
        const props = '<w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:eastAsia="MS Mincho" w:cs="Georgia"/><w:noProof/><w14:ligatures w14:val="standard"/></w:rPr>';
        const pPr = '<w:pPr><w:suppressAutoHyphens/><w:spacing w:beforeAutospacing="1" w:after="120"/><w:rPr><w:lang w:val="de-DE"/></w:rPr></w:pPr>';
        const rest = (xml) => xml.slice(xml.indexOf('<w:smartTag'), xml.indexOf('</w:p>', xml.indexOf('<w:smartTag')) + 6);
        const tableOf = (xml) => xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>') + 8);
        const head = (id) => `<w:p w:rsidR="00BB0001" w:rsidRDefault="00BB0003" w14:paraId="3A3B3C${id}" w14:textId="77777777">${pPr}<w:proofErr w:type="spellStart"/><w:r w:rsidRPr="00BB0002">${props}<w:t xml:space="preserve">alpha beta</w:t></w:r></w:p>`;
        const tail = `<w:p w:rsidR="00BB0001" w:rsidRDefault="00BB0003">${pPr}<w:r w:rsidRPr="00BB0002">${props}<w:t xml:space="preserve"> gamma</w:t></w:r><w:proofErr w:type="spellEnd"/>${rest(original)}`;

        // Split in the body.
        let { xml, report, same } = saved.split;
        check('split: the text is two paragraphs', saved.split.text.startsWith('alpha beta\n gamma delta'), true);
        check('split: the first half is the paragraph with the text before the caret - its ids, its properties, the run split', xml.includes(head('01')) ? true : xml.slice(xml.indexOf('<w:body>'), xml.indexOf('<w:tbl>')), true);
        check('split: the second half is the rest of the paragraph - without the paragraph\'s id, with everything after the caret', xml.includes(tail) ? true : xml.slice(xml.indexOf('<w:body>'), xml.indexOf('<w:tbl>')), true);
        check('split: nothing is written from the model', [report.GeneratedParagraphs, report.PatchedParagraphs, report.DerivedParagraphs], [0, 1, 1]);
        check('split: the table is the text it was', xml.includes(tableOf(original)), true);
        check('split: read again, the editor\'s text and layout', same, [true, true]);

        // Joined again by hand: the paragraph that was read.
        ({ xml, report, same } = saved.rejoin);
        check('split and joined again: document.xml is the text that was read', xml === original, true);

        // Two paragraphs joined.
        ({ xml, report, same } = saved.join);
        check('joined: the text is one paragraph', saved.join.text.includes('first of twosecond of two'), true);
        check('joined: the first paragraph with the content of the second at its end',
            xml.includes('<w:p><w:r><w:t xml:space="preserve">first of two</w:t></w:r><w:bookmarkStart w:id="11" w:name="second"/><w:r w:rsidR="00BB0007"><w:rPr><w:noProof/></w:rPr><w:t>second of two</w:t></w:r><w:bookmarkEnd w:id="11"/></w:p>') ? true : xml.slice(xml.indexOf('first of two') - 60), true);
        check('joined: nothing is written from the model', [report.GeneratedParagraphs, report.PatchedParagraphs], [0, 1]);
        check('joined: the second paragraph is gone', (xml.match(/second of two/g) ?? []).length, 1);
        check('joined: read again, the editor\'s text and layout', same, [true, true]);

        // Split in a cell of a table: the table is written anew, its paragraphs are not.
        ({ xml, report, same } = saved.cell);
        const table = tableOf(xml);
        check('split in a cell: both halves are made of the paragraph that was read', [table.includes(head('02')), table.includes(tail.replace(rest(original), rest(tableOf(original))))], [true, true]);
        check('split in a cell: the other paragraphs of the table are the text they were (a tracked change of a run\'s properties too)',
            [table.includes('<w:p w14:paraId="3A3B3C03"><w:r><w:rPr><w:i/><w:rPrChange w:id="9" w:author="A" w:date="2020-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t xml:space="preserve">tracked format here</w:t></w:r></w:p>'), table.includes('second cell'), table.includes('last cell')], [true, true, true]);
        check('split in a cell: no paragraph is written from the model', report.GeneratedParagraphs, 0);
        check('split in a cell: the table\'s properties, its rows\' and its cells\' are the text they were, and the other row',
            ['<w:tblLook w:val="04A0" w:firstRow="1" w:noHBand="0"/><w:tblCaption w:val="a caption"/>', '<w:tr w:rsidR="00BB0005" w14:paraId="4A4B4C01"><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:noWrap/><w:hideMark/></w:tcPr>',
                tableOf(original).slice(tableOf(original).indexOf('<w:tr w:rsidR="00BB0006">'))].filter((piece) => !table.includes(piece)), []);
        check('split in a cell: the paragraph of the body is the text it was', xml.includes(original.slice(original.indexOf('<w:p '), original.indexOf('<w:tbl>'))), true);
        check('split in a cell: read again, the editor\'s text and layout', same, [true, true]);
    } catch (e) { fail(e); }
    finish();
};
