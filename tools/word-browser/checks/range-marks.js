// The marks of ranges the model does not hold stay where the file had them
// when their paragraph is written from the model, and every part's ranges are
// opened once and closed once.
//
// The model has a bookmark as the bookmark of a run: one in a paragraph
// without runs, the end of one before the first run of the next paragraph,
// and whatever stands between two blocks of the body are not in it - nor are
// the ranges of comments and of permissions. A paragraph written from the
// model used to lose them (a start between two blocks went with the block
// after it, and left its end behind); a block that was deleted left the end
// of a bookmark it opened; and a header, the notes and the comments written
// from the model were not checked as the body is.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
let editor, Word, OF, bytes;
const saved = {};
const ranges = (xml) => {
    const problems = [];
    for (const [open, close] of [['bookmarkStart', 'bookmarkEnd'], ['commentRangeStart', 'commentRangeEnd'], ['permStart', 'permEnd'], ['moveFromRangeStart', 'moveFromRangeEnd'], ['moveToRangeStart', 'moveToRangeEnd']]) {
        const ids = (tag) => [...xml.matchAll(new RegExp(`<w:${tag}\\b[^>]*\\bw:id="([^"]*)"`, 'g'))].map((m) => m[1]);
        const starts = ids(open), ends = ids(close);
        if (new Set(starts).size !== starts.length) problems.push(`${open} twice`);
        if (new Set(ends).size !== ends.length) problems.push(`${close} twice`);
        if (starts.slice().sort().join() !== ends.slice().sort().join()) problems.push(`${open} ${starts.join(' ')} / ${close} ${ends.join(' ')}`);
    }
    return problems;
};
try {
    OF = await loadFormats();
    Word = await loadWord();
    bytes = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/range-marks.docx'));
    const open = async () => {
        if (editor) Word.dispose(editor);
        editor = Word.initializeCanvas(document.getElementById('c'));
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
        Word.completeLayout(editor);
        Word.getInputElement(editor).focus();
    };
    const blockOf = (text) => Word.findText(editor, text, true)[0].BlockId;
    const snap = async (stage) => {
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const report = OF.word.getLastSaveReport();
        const zip = await window.JSZip.loadAsync(out);
        const part = async (name) => (zip.file(name) ? zip.file(name).async('string') : '');
        const again = Word.initializeCanvas(document.createElement('canvas'));
        Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(out));
        saved[stage] = { xml: await part('word/document.xml'), header: await part('word/header1.xml'), report, bookmarks: Word.getBookmarkInfos(again).map((info) => info.Name).sort(), text: Word.getText(again) };
        Word.dispose(again);
    };
    // A link put into a paragraph: the paragraph is written from the model.
    const linkInto = async (stage, id, offset) => {
        await open();
        const target = typeof id === 'function' ? id() : id;
        Word.setSelection(editor, target, offset, target, offset);
        Word.insertHyperlink(editor, 'https://example.com/' + stage, 'LINK');
        await snap(stage);
    };
    window.stage = async (name) => {
        if (name === 'unedited') { await open(); await snap('unedited'); }
        if (name === 'between') await linkInto('between', () => blockOf('second, bookmarked'), 3);
        // (The paragraph without text: the one after "second".)
        if (name === 'empty') { await open(); const id = blockOf('second, bookmarked'); Word.setSelection(editor, id, 31, id, 31); }
        if (name === 'ranges') await linkInto('ranges', () => blockOf('fourth'), 2);
        if (name === 'cell') await linkInto('cell', () => blockOf('cell text'), 2);
        // The paragraph a bookmark starts in is deleted, with the empty one before it (before which the end of another bookmark stands): the end of the first is in a paragraph that stays as it was read.
        if (name === 'deleted') { await open(); const from = blockOf('second, bookmarked'), to = blockOf('third'); Word.setSelection(editor, from, 31, to, 5); }
        // The paragraph before which a bookmark starts (between the blocks) is joined to the paragraph before it.
        if (name === 'joined') { await open(); const id = blockOf('second, bookmarked'); Word.setSelection(editor, id, 0, id, 0); }
        // The header's second paragraph, which the header's bookmark ends in, is joined to the first and its text deleted.
        if (name === 'header') { await open(); Word.enterHeaderFooter(editor, 0, 'header'); const a = blockOf('header A'), b = blockOf('header B'); Word.setSelection(editor, a, 8, b, 8); }
        Word.getInputElement(editor).focus();
    };
    window.snap = snap;
    window.linkHere = async (stage) => { Word.insertHyperlink(editor, 'https://example.com/' + stage, 'LINK'); await snap(stage); };
    window.drive = [
        ['eval', 'window.stage("unedited")'], ['eval', 'window.stage("between")'], ['eval', 'window.stage("empty")'], ['press', 'ArrowRight'], ['eval', 'window.linkHere("empty")'], ['eval', 'window.stage("ranges")'], ['eval', 'window.stage("cell")'],
        ['eval', 'window.stage("deleted")'], ['press', 'Delete'], ['eval', 'window.snap("deleted")'],
        ['eval', 'window.stage("joined")'], ['press', 'Backspace'], ['eval', 'window.snap("joined")'],
        ['eval', 'window.stage("header")'], ['press', 'Delete'], ['eval', 'window.snap("header")'],
    ];
} catch (e) { fail(e); finish(); }
window.finish = async () => {
    try {
        const original = await (await window.JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
        check('the file: the ranges of its body and of its header are balanced', [ranges(original), ranges(saved.unedited.header)], [[], []]);
        check('saved unedited: the body and the header are the text that was read', [saved.unedited.xml === original, saved.unedited.header.includes('<w:bookmarkStart w:id="1" w:name="in_header"/>')], [true, true]);
        {
            const s = saved.between;
            check('a paragraph between the start and the end of a bookmark that stand between blocks, written from the model: the start is before it, the end after it',
                [s.report.ModelBlocks, /<w:bookmarkStart w:id="20" w:name="between_blocks"\/><w:p>(?:(?!<\/w:p>).)*LINK(?:(?!<\/w:p>).)*<\/w:p><w:bookmarkEnd w:id="20"\/>/.test(s.xml), ranges(s.xml)], [1, true, []]);
        }
        {
            const s = saved.empty;
            check('a paragraph without runs that holds a bookmark, written from the model: the bookmark is in it',
                [s.report.ModelBlocks, /<w:p><w:pPr><w:jc w:val="center"\/><\/w:pPr><w:bookmarkStart w:id="21" w:name="in_empty"\/><w:bookmarkEnd w:id="21"\/>(?:(?!<\/w:p>).)*LINK/.test(s.xml), ranges(s.xml)], [1, true, []]);
        }
        {
            const s = saved.ranges;
            const para = (s.xml.match(/<w:p>(?:(?!<\/w:p>).)*commentReference(?:(?!<\/w:p>).)*<\/w:p>/) ?? [''])[0];
            check('a paragraph with the end of a bookmark before its first run, the range of a permission and the range of a comment, written from the model: all of them are in it, once',
                [s.report.ModelBlocks, /^<w:p><w:bookmarkEnd w:id="22"\/><w:permStart w:id="30" w:edGrp="everyone"\/>/.test(para), /<w:commentRangeStart w:id="5"\/>/.test(para), /<w:commentRangeEnd w:id="5"\/>(?:(?!<w:r>).)*<w:permEnd w:id="30"\/><\/w:p>$|<w:permEnd w:id="30"\/><w:commentRangeEnd w:id="5"\/><\/w:p>$|<w:commentRangeEnd w:id="5"\/><w:permEnd w:id="30"\/><\/w:p>$/.test(para),
                 para.indexOf('commentRangeStart') < para.indexOf('commented'), ranges(s.xml)], [1, true, true, true, true, []]);
        }
        {
            const s = saved.cell;
            check('a paragraph of a table with the range of a permission, written from the model into the table\'s XML: the range is in it',
                [s.report.ModelBlocks, s.report.ModelParagraphsInTables, /<w:p><w:permStart w:id="31" w:edGrp="everyone"\/>(?:(?!<\/w:p>).)*LINK(?:(?!<\/w:p>).)*<w:permEnd w:id="31"\/><\/w:p>/.test(s.xml), ranges(s.xml)], [0, 1, true, []]);
        }
        for (const stage of ['between', 'empty', 'ranges', 'cell']) check(`${stage}: read again, the editor lists the bookmarks it listed`, [saved[stage].bookmarks, saved[stage].text.includes('LINK')],
            // (The bookmark of the paragraph that had no run is, now that the paragraph has one, a bookmark the editor lists.)
            [stage === 'empty' ? ['in_empty', ...saved.unedited.bookmarks].sort() : saved.unedited.bookmarks, true]);
        {
            const s = saved.deleted;
            check('the paragraph a bookmark starts in is deleted: its end, in a paragraph that is the text it was read from, is not left behind; the end of a bookmark that stood before a deleted block is before the next block',
                [s.text.includes('third'), s.text.includes('fourth commented'), /<w:bookmarkEnd w:id="22"\/>/.test(s.xml), /"spans"/.test(s.xml), /<w:p><w:permStart w:id="30" w:edGrp="everyone"\/><w:r><w:t xml:space="preserve">fourth /.test(s.xml),
                 /<w:bookmarkStart w:id="20" w:name="between_blocks"\/><w:p>(?:(?!<\/w:p>).)*second, bookmarked from outside(?:(?!<\/w:p>).)*<\/w:p><w:bookmarkEnd w:id="20"\/><w:p><w:permStart/.test(s.xml), /in_empty/.test(s.xml), ranges(s.xml)],
                [false, true, false, false, true, true, false, []]);
        }
        {
            const s = saved.joined;
            check('a paragraph joined to the one before it (Backspace): the start of the bookmark that stood between the two is where its text now begins, in the XML that was read; the end is where it was',
                [s.report.PatchedParagraphs, s.report.GeneratedParagraphs, s.report.ModelBlocks, /<w:t xml:space="preserve">first<\/w:t><\/w:r><w:bookmarkStart w:id="20" w:name="between_blocks"\/><w:r><w:t>second, bookmarked from outside<\/w:t><\/w:r><\/w:p><w:bookmarkEnd w:id="20"\/><w:p><w:pPr><w:jc w:val="center"\/>/.test(s.xml), ranges(s.xml)],
                [1, 0, 0, true, []]);
        }
        {
            const s = saved.header;
            check('a header whose bookmark ended in a paragraph that was deleted, written from the model: the bookmark is opened and closed once',
                [s.text !== undefined, /in_header/.test(s.header), ranges(s.header)], [true, true, []]);
        }
    } catch (e) { fail(e); }
    finish();
};
