// The bookmarks of the text go from the reader through the editor to the writer as typed values (a run's
// BookmarkStarts and BookmarkEnds), and the editor's edits keep them where the text is: a bookmark across
// a paragraph split by Enter starts in the first half and ends in the second; paragraphs joined by
// Backspace keep the bookmarks of both, among them one opened after the last run of the first; pasted text
// is of no bookmark (the bookmark stays once, where it was); a split undone is the paragraph as read; a
// bookmark that ends where it starts stays before its run when that run is split. Each saved as the XML it
// was read from and written from the model alone; each read again.
import { check, finish, fail, loadWord, loadFormats } from './common.js';
let OF, Word, editor, bytes, original;
const saved = {};
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const body =
    '<w:p><w:r><w:t xml:space="preserve">alpha </w:t></w:r><w:bookmarkStart w:id="1" w:name="across"/><w:r><w:t>beta gamma</w:t></w:r><w:bookmarkEnd w:id="1"/><w:r><w:t xml:space="preserve"> delta</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>first of two</w:t></w:r></w:p>' +
    '<w:p><w:bookmarkStart w:id="2" w:name="joined"/><w:r><w:t>second of two</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>' +
    '<w:p><w:bookmarkStart w:id="3" w:name="A"/><w:bookmarkStart w:id="4" w:name="B"/><w:r><w:t>several marks</w:t></w:r><w:bookmarkEnd w:id="4"/><w:bookmarkEnd w:id="3"/><w:bookmarkStart w:id="5" w:name="opened"/></w:p>' +
    '<w:p><w:r><w:t>closed here</w:t></w:r><w:bookmarkEnd w:id="5"/></w:p>' +
    '<w:p><w:r><w:t xml:space="preserve">paste </w:t></w:r><w:bookmarkStart w:id="6" w:name="point"/><w:bookmarkEnd w:id="6"/><w:r><w:t>target</w:t></w:r></w:p>';
const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
const makeDocx = async () => {
    const zip = new window.JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId100" Type="${REL}/styles" Target="styles.xml"/></Relationships>`);
    zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`);
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}${sect}</w:body></w:document>`);
    return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
};
// The paragraphs of a part, each as its text with its bookmarks' marks: "{id:name" a start, "}id" an end.
const shape = (xml) => [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map((m) =>
    [...m[0].matchAll(/<w:bookmarkStart [^>]*w:id="([^"]*)" w:name="([^"]*)"\/>|<w:bookmarkEnd w:id="([^"]*)"\/>|<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)]
        .map((t) => (t[1] !== undefined ? `{${t[1]}:${t[2]}` : t[3] !== undefined ? `}${t[3]}` : t[4])).join(''));
// Every range is opened once and closed once.
const balanced = (xml) => {
    const ids = (tag) => [...xml.matchAll(new RegExp(`<w:${tag}\\b[^>]*\\bw:id="([^"]*)"`, 'g'))].map((m) => m[1]);
    const starts = ids('bookmarkStart'), ends = ids('bookmarkEnd');
    return new Set(starts).size === starts.length && new Set(ends).size === ends.length && starts.slice().sort().join() === ends.slice().sort().join();
};
const open = async () => {
    if (editor) Word.dispose(editor);
    editor = Word.initializeCanvas(document.getElementById('c'));
    Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
    Word.completeLayout(editor);
    Word.getInputElement(editor).focus();
};
const blockOf = (text) => Word.findText(editor, text, true)[0];
const caretAt = (text, offset) => { const found = blockOf(text); Word.setSelection(editor, found.BlockId, found.Start + offset, found.BlockId, found.Start + offset); Word.getInputElement(editor).focus(); };
const write = async (pkg) => {
    const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(pkg)).arrayBuffer());
    const xml = await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string');
    const again = Word.initializeCanvas(document.createElement('canvas'));
    Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(out));
    const infos = Word.getBookmarkInfos(again).map((info) => info.Name).sort();
    Word.dispose(again);
    return { xml, infos };
};
const snap = async (stage) => {
    const asRead = await write(Word.exportDocumentPackage(editor));
    const model = Word.exportDocumentPackage(editor);
    model.RegenerateAll = true;
    saved[stage] = { asRead, model: await write(model), text: Word.getText(editor) };
};
try {
    OF = await loadFormats();
    Word = await loadWord();
    bytes = await makeDocx();
    original = await (await window.JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
    window.place = async (stage) => {
        await open();
        if (stage === 'split') caretAt('beta gamma', 4);
        if (stage === 'join') caretAt('second of two', 0);
        if (stage === 'joinOpened') caretAt('closed here', 0);
        if (stage === 'undo') caretAt('several marks', 7);
    };
    window.snap = snap;
    window.paste = async () => {
        await open();
        const found = blockOf('beta gamma');
        Word.setSelection(editor, found.BlockId, found.Start, found.BlockId, found.Start + 4);
        const clip = Word.copySelection(editor);
        caretAt('target', 3);
        Word.pasteClipboard(editor, clip.Text, clip.Html);
        await snap('paste');
    };
    // A bookmark set inside another, the one opened after a run removed, another renamed (the host API).
    window.api = async () => {
        await open();
        const found = blockOf('gamma');
        Word.setSelection(editor, found.BlockId, found.Start, found.BlockId, found.End);
        Word.setBookmark(editor, 'inner');
        Word.removeBookmark(editor, 'opened');
        Word.renameBookmark(editor, 'A', 'First');
        saved.apiInfos = Word.getBookmarkInfos(editor).map((info) => info.Name).sort();
        await snap('api');
    };
    window.drive = [
        ['eval', 'window.place("split")'], ['press', 'Enter'], ['eval', 'window.snap("split")'],
        ['eval', 'window.place("join")'], ['press', 'Backspace'], ['eval', 'window.snap("join")'],
        ['eval', 'window.place("joinOpened")'], ['press', 'Backspace'], ['eval', 'window.snap("joinOpened")'],
        ['eval', 'window.place("undo")'], ['press', 'Enter'], ['press', 'Control+z'], ['eval', 'window.snap("undo")'],
        ['eval', 'window.paste()'],
        ['eval', 'window.api()'],
    ];
} catch (e) { fail(e); finish(); }
window.finish = async () => {
    try {
        const all = ['A', 'B', 'across', 'joined', 'opened', 'point'];
        check('the file as read: its paragraphs and bookmarks', shape(original),
            ['alpha {1:acrossbeta gamma}1 delta', 'first of two', '{2:joinedsecond of two}2', '{3:A{4:Bseveral marks}4}3{5:opened', 'closed here}5', 'paste {6:point}6target']);
        const both = (stage, expected) => {
            for (const [how, result] of [['as read', saved[stage].asRead], ['from the model', saved[stage].model]]) {
                check(`${stage}, ${how}: the paragraphs and their bookmarks`, shape(result.xml), expected);
                check(`${stage}, ${how}: every bookmark opened once and closed once`, balanced(result.xml), true);
                check(`${stage}, ${how}: read again, the bookmarks`, result.infos, all);
            }
        };
        // Enter inside the bookmark: it starts before "beta" in the first half and ends after "gamma" in the second.
        both('split', ['alpha {1:acrossbeta', ' gamma}1 delta', 'first of two', '{2:joinedsecond of two}2', '{3:A{4:Bseveral marks}4}3{5:opened', 'closed here}5', 'paste {6:point}6target']);
        // Backspace at the start of a bookmarked paragraph: the bookmark is around its text in the paragraph joined.
        both('join', ['alpha {1:acrossbeta gamma}1 delta', 'first of two{2:joinedsecond of two}2', '{3:A{4:Bseveral marks}4}3{5:opened', 'closed here}5', 'paste {6:point}6target']);
        // The bookmark opened after the last run of a paragraph, the paragraph after it joined: it is where the text joined begins.
        both('joinOpened', ['alpha {1:acrossbeta gamma}1 delta', 'first of two', '{2:joinedsecond of two}2', '{3:A{4:Bseveral marks}4}3{5:openedclosed here}5', 'paste {6:point}6target']);
        // A split in a run of several bookmarks, undone: the paragraph as read.
        check('undo: the text as read', saved.undo.text, saved.undo.text.replace('several\nmarks', 'several marks'));
        both('undo', ['alpha {1:acrossbeta gamma}1 delta', 'first of two', '{2:joinedsecond of two}2', '{3:A{4:Bseveral marks}4}3{5:opened', 'closed here}5', 'paste {6:point}6target']);
        check('undo, as read: document.xml is the text that was read', saved.undo.asRead.xml === original, true);
        // Text of a bookmark pasted into the run after an empty bookmark: of no bookmark; the empty one stays before the run.
        both('paste', ['alpha {1:acrossbeta gamma}1 delta', 'first of two', '{2:joinedsecond of two}2', '{3:A{4:Bseveral marks}4}3{5:opened', 'closed here}5', 'paste {6:point}6tarbetaget']);
        // The editor lists every bookmark; one set inside another keeps the other around the text; one opened
        // after a run is removed with its end, another renamed (written from the model: the model's bookmarks).
        check('api: the bookmarks the editor lists', saved.apiInfos, ['B', 'First', 'across', 'inner', 'joined', 'point']);
        const api = saved.api.model;
        check('api, from the model: the paragraphs and their bookmarks', shape(api.xml),
            ['alpha {1:acrossbeta {7:innergamma}7}1 delta', 'first of two', '{2:joinedsecond of two}2', '{3:First{4:Bseveral marks}4}3', 'closed here', 'paste {6:point}6target']);
        check('api, from the model: every bookmark opened once and closed once', balanced(api.xml), true);
        check('api, from the model: read again, the bookmarks', api.infos, ['B', 'First', 'across', 'inner', 'joined', 'point']);
    } catch (e) { fail(e); }
    finish();
};
