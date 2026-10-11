// The bookmarks of a paragraph survive what is typed into it. A format chosen
// at a caret (Bold with nothing selected) and two characters typed there used
// to write the title of sample-formatted.docx from the model with its first
// bookmark three times, the other two gone and three ends of one: a run had
// one bookmark start and one end in the model, the typed character took the
// bookmarks of the run it was typed into, and the second character was a run
// of its own. Now: the characters typed with a chosen format are one run
// without the bookmarks of the run around them; the paragraph is changed in
// the XML it was read from; a run has every bookmark that starts before it
// and ends after it, and a paragraph written from the model writes each once
// - with the outline level the paragraph states.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const bytes = await bytesOf('/repo/tests/test-documents/sample-formatted.docx');
    const original = await (await window.JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
    const canvas = document.getElementById('c');
    const open = async () => { const editor = Word.initializeCanvas(canvas); Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes)); return editor; };
    const titleOf = (xml) => xml.slice(xml.indexOf('<w:body>') + 8, xml.indexOf('</w:p>') + 6);
    // Every range of a part is opened once and closed once.
    const ranges = (xml) => {
        const problems = [];
        for (const [open, close] of [['bookmarkStart', 'bookmarkEnd'], ['commentRangeStart', 'commentRangeEnd'], ['permStart', 'permEnd'], ['moveFromRangeStart', 'moveFromRangeEnd'], ['moveToRangeStart', 'moveToRangeEnd']]) {
            const ids = (tag) => [...xml.matchAll(new RegExp(`<w:${tag}\\b[^>]*\\bw:id="([^"]*)"`, 'g'))].map((m) => m[1]);
            const starts = ids(open), ends = ids(close);
            if (new Set(starts).size !== starts.length) problems.push(`${open} twice`);
            if (new Set(ends).size !== ends.length) problems.push(`${close} twice`);
            if (starts.slice().sort().join() !== ends.slice().sort().join()) problems.push(`${starts.length} ${open}, ${ends.length} ${close}`);
        }
        return problems;
    };
    const names = (xml) => [...xml.matchAll(/<w:bookmarkStart [^>]*w:name="([^"]*)"/g)].map((m) => m[1]);
    const saved = async (editor, fromModel) => {
        const pkg = Word.exportDocumentPackage(editor);
        if (fromModel) pkg.RegenerateAll = true;
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(pkg)).arrayBuffer());
        return { xml: await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string'), report: OF.word.getLastSaveReport(), pkg, bytes: out };
    };
    check('the file: its ranges are balanced; the title has three bookmarks and an outline level', [ranges(original), names(titleOf(original)), /<w:outlineLvl w:val="0"\/>/.test(titleOf(original))], [[], ['OLE_LINK1', 'OLE_LINK2', '_Toc359077851'], true]);
    const typed = async (count) => {
        const editor = await open();
        const match = Word.findText(editor, 'Demonstration of DOCX', true)[0];
        Word.setSelection(editor, match.BlockId, 4, match.BlockId, 4);
        Word.toggleBold(editor);
        for (let index = 0; index < count; index++) Word.insertText(editor, 'Q');
        return { editor, id: match.BlockId };
    };
    for (const count of [1, 2, 5]) {
        const { editor, id } = await typed(count);
        const text = 'Q'.repeat(count);
        const result = await saved(editor, false);
        const title = titleOf(result.xml);
        const runs = result.pkg.Paragraphs[0].Runs.map((run) => [run.Text, run.Bold]);
        check(`bold chosen at the caret, ${count} typed: the typed characters are one bold run`, runs, [['Demo', false], [text, true], ['nstration of DOCX support in calibre', false]]);
        check(`${count} typed: the paragraph is changed in the XML it was read from`, [result.report.PatchedParagraphs, result.report.GeneratedParagraphs, result.report.ModelBlocks, result.report.Reasons], [1, 0, 0, []]);
        check(`${count} typed: its bookmarks, once each, its outline level and its rsids are there; the document's ranges are balanced`,
            [names(title), (title.match(/<w:bookmarkEnd /g) ?? []).length, /<w:outlineLvl w:val="0"\/>/.test(title), /w:rsidR="00332BA2"/.test(title), ranges(result.xml), names(result.xml).length],
            [['OLE_LINK1', 'OLE_LINK2', '_Toc359077851'], 1, true, true, [], names(original).length]);
        check(`${count} typed: the typed text is bold in the file`, new RegExp(`<w:r><w:rPr>(?:(?!</w:rPr>).)*<w:b/>(?:(?!</w:rPr>).)*</w:rPr><w:t[^>]*>${text}</w:t></w:r>`).test(title), true);
        // The same document written from the model alone.
        const model = await saved(editor, true);
        const modelTitle = titleOf(model.xml);
        check(`${count} typed, written from the model: every bookmark of the title once, its outline level, balanced ranges`,
            [names(modelTitle), /<w:outlineLvl w:val="0"\/>/.test(modelTitle), ranges(model.xml), names(model.xml).slice().sort()],
            [['OLE_LINK1', 'OLE_LINK2', '_Toc359077851'], true, [], names(original).slice().sort()]);
        // Read again: the bookmarks the editor lists.
        const again = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(model.bytes));
        check(`${count} typed, written from the model and read again: the title's bookmarks`, Word.getBookmarkInfos(again).filter((info) => ['OLE_LINK1', 'OLE_LINK2', '_Toc359077851'].includes(info.Name)).map((info) => info.Name).sort(), ['OLE_LINK1', 'OLE_LINK2', '_Toc359077851']);
        Word.dispose(again); Word.dispose(editor);
    }
    // A part of the title made bold (its run split in three), written from the model: the same.
    {
        const editor = await open();
        const match = Word.findText(editor, 'Demonstration of DOCX', true)[0];
        Word.setSelection(editor, match.BlockId, 4, match.BlockId, 9);
        Word.toggleBold(editor);
        const model = await saved(editor, true);
        const title = titleOf(model.xml);
        check('a part of the title made bold, written from the model: every bookmark once', [names(title), ranges(model.xml)], [['OLE_LINK1', 'OLE_LINK2', '_Toc359077851'], []]);
        // One of the three removed, another renamed: the others stay.
        Word.removeBookmark(editor, 'OLE_LINK2');
        Word.renameBookmark(editor, 'OLE_LINK1', 'First');
        const after = await saved(editor, true);
        check('one bookmark of the three removed and one renamed', [names(titleOf(after.xml)), ranges(after.xml)], [['First', '_Toc359077851'], []]);
    }
} catch (e) { fail(e); }
finish();
