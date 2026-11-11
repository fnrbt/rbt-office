// A bookmark renamed, removed or added (renameBookmark / removeBookmark / setBookmark) in a paragraph that is
// saved as the XML it was read from: the change is made in the paragraph's source - the name of its
// w:bookmarkStart set, its w:bookmarkStart and w:bookmarkEnd taken out, a new one put in around its text -
// and everything else of the paragraph stays the bytes it was read as (a proofing mark, a later format's
// attribute, the run properties). Also for a bookmark that ends in another paragraph, and for one opened
// after the last run of a paragraph. Saved as read and from the model; each read again.
import { check, finish, fail, loadWord, loadFormats } from './common.js';
let OF, Word, editor, bytes, original;
const saved = {};
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
// (What the model does not hold, in each paragraph: a w14 id, a proofing mark, run properties as written.)
const paragraphs = [
    '<w:p w14:paraId="1A2B3C01"><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">alpha </w:t></w:r><w:proofErr w:type="spellStart"/><w:bookmarkStart w:id="1" w:name="one"/><w:r><w:t>beta</w:t></w:r><w:bookmarkEnd w:id="1"/><w:proofErr w:type="spellEnd"/><w:r><w:t xml:space="preserve"> gamma</w:t></w:r></w:p>',
    '<w:p w14:paraId="1A2B3C02"><w:bookmarkStart w:id="2" w:name="two"/><w:r><w:rPr><w:i/></w:rPr><w:t>remove me</w:t></w:r><w:bookmarkEnd w:id="2"/><w:proofErr w:type="gramStart"/><w:r><w:t xml:space="preserve"> kept</w:t></w:r><w:proofErr w:type="gramEnd"/></w:p>',
    '<w:p w14:paraId="1A2B3C03"><w:r><w:t xml:space="preserve">first </w:t></w:r><w:bookmarkStart w:id="3" w:name="three"/><w:r><w:t>across</w:t></w:r></w:p>',
    '<w:p w14:paraId="1A2B3C04"><w:r><w:t>second</w:t></w:r><w:bookmarkEnd w:id="3"/><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>',
    '<w:p w14:paraId="1A2B3C05"><w:r><w:t>opened at the end</w:t></w:r><w:bookmarkStart w:id="5" w:name="opened"/></w:p>',
    '<w:p w14:paraId="1A2B3C06"><w:r><w:t>closed here</w:t></w:r><w:bookmarkEnd w:id="5"/><w:proofErr w:type="spellStart"/></w:p>',
    '<w:p w14:paraId="1A2B3C07"><w:proofErr w:type="spellStart"/><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">plain text </w:t></w:r><w:r><w:t>here</w:t></w:r><w:proofErr w:type="spellEnd"/></w:p>',
    '<w:p w14:paraId="1A2B3C08"><w:bookmarkStart w:id="8" w:name="untouched"/><w:r><w:t>untouched</w:t></w:r><w:bookmarkEnd w:id="8"/></w:p>'];
const sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
const makeDocx = async () => {
    const zip = new window.JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>');
    zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId100" Type="${REL}/styles" Target="styles.xml"/></Relationships>`);
    zip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`);
    zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${paragraphs.join('')}${sect}</w:body></w:document>`);
    return new Uint8Array(await zip.generateAsync({ type: 'uint8array' }));
};
// The paragraphs of a part, as written.
const paragraphsOf = (xml) => [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map((m) => m[0]);
// Each paragraph as its text with its bookmarks' marks: "{id:name" a start, "}id" an end.
const shape = (xml) => paragraphsOf(xml).map((p) =>
    [...p.matchAll(/<w:bookmarkStart [^>]*w:id="([^"]*)" w:name="([^"]*)"\/>|<w:bookmarkEnd w:id="([^"]*)"\/>|<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)]
        .map((t) => (t[1] !== undefined ? `{${t[1]}:${t[2]}` : t[3] !== undefined ? `}${t[3]}` : t[4])).join(''));
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
};
const write = async (pkg) => {
    const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(pkg)).arrayBuffer());
    const report = OF.word.getLastSaveReport();
    const xml = await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string');
    const again = Word.initializeCanvas(document.createElement('canvas'));
    Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(out));
    const infos = Word.getBookmarkInfos(again).map((info) => info.Name).sort();
    Word.dispose(again);
    return { xml, infos, report };
};
const snap = async (stage) => {
    const asRead = await write(Word.exportDocumentPackage(editor));
    const model = Word.exportDocumentPackage(editor);
    model.RegenerateAll = true;
    saved[stage] = { asRead, model: await write(model), infos: Word.getBookmarkInfos(editor).map((info) => info.Name).sort() };
};
try {
    OF = await loadFormats();
    Word = await loadWord();
    bytes = await makeDocx();
    original = await (await window.JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
    window.stage = async (name) => {
        await open();
        if (name === 'rename') { Word.renameBookmark(editor, 'one', 'Renamed'); Word.renameBookmark(editor, 'opened', 'Moved'); }
        if (name === 'remove') { Word.removeBookmark(editor, 'two'); Word.removeBookmark(editor, 'three'); Word.removeBookmark(editor, 'opened'); }
        if (name === 'add') {
            const found = Word.findText(editor, 'text', true)[0];
            Word.setSelection(editor, found.BlockId, found.Start, found.BlockId, found.End);
            Word.setBookmark(editor, 'added');
        }
        await snap(name);
    };
    window.drive = [['eval', 'window.stage("rename")'], ['eval', 'window.stage("remove")'], ['eval', 'window.stage("add")']];
} catch (e) { fail(e); finish(); }
window.finish = async () => {
    try {
        const all = ['one', 'opened', 'three', 'two', 'untouched'];
        check('the file as read: its paragraphs and bookmarks', shape(original),
            ['alpha {1:onebeta}1 gamma', '{2:tworemove me}2 kept', 'first {3:threeacross', 'second}3 after', 'opened at the end{5:opened', 'closed here}5', 'plain text here', '{8:untoucheduntouched}8']);
        const sameBut = (stage, expectedShape, expectedNames, changed) => {
            for (const [how, result] of [['as read', saved[stage].asRead], ['from the model', saved[stage].model]]) {
                check(`${stage}, ${how}: the paragraphs and their bookmarks`, shape(result.xml), expectedShape);
                check(`${stage}, ${how}: every bookmark opened once and closed once`, balanced(result.xml), true);
                check(`${stage}, ${how}: read again, the bookmarks`, result.infos, expectedNames);
            }
            check(`${stage}: the bookmarks the editor lists`, saved[stage].infos, expectedNames);
            // As read: each paragraph is the bytes it was read as, but for the change.
            const now = paragraphsOf(saved[stage].asRead.xml);
            check(`${stage}, as read: every paragraph the source with only the change made in it`,
                now.map((p, i) => p === (changed[i] ?? paragraphs[i]) ? 'same' : p), paragraphs.map(() => 'same'));
            check(`${stage}, as read: no paragraph written from the model`, saved[stage].asRead.report.GeneratedParagraphs ?? 0, 0);
        };
        // Renamed: the w:name of its start, in place - in a paragraph of its own, and one opened after the last run.
        sameBut('rename', ['alpha {1:Renamedbeta}1 gamma', '{2:tworemove me}2 kept', 'first {3:threeacross', 'second}3 after', 'opened at the end{5:Moved', 'closed here}5', 'plain text here', '{8:untoucheduntouched}8'],
            ['Moved', 'Renamed', 'three', 'two', 'untouched'],
            { 0: paragraphs[0].replace('w:name="one"', 'w:name="Renamed"'), 4: paragraphs[4].replace('w:name="opened"', 'w:name="Moved"') });
        // Removed: its start and its end, wherever they are (an end in the next paragraph too).
        sameBut('remove', ['alpha {1:onebeta}1 gamma', 'remove me kept', 'first across', 'second after', 'opened at the end', 'closed here', 'plain text here', '{8:untoucheduntouched}8'],
            ['one', 'untouched'],
            { 1: paragraphs[1].replace('<w:bookmarkStart w:id="2" w:name="two"/>', '').replace('<w:bookmarkEnd w:id="2"/>', ''),
              2: paragraphs[2].replace('<w:bookmarkStart w:id="3" w:name="three"/>', ''), 3: paragraphs[3].replace('<w:bookmarkEnd w:id="3"/>', ''),
              4: paragraphs[4].replace('<w:bookmarkStart w:id="5" w:name="opened"/>', ''), 5: paragraphs[5].replace('<w:bookmarkEnd w:id="5"/>', '') });
        // Added: around the text it was set on, the rest of its paragraph as read.
        const added = saved.add.asRead;
        check('add, as read: the bookmark around its text', shape(added.xml)[6].replace(/\{\d+:added/, '{N:added').replace(/\}\d+/, '}N'), 'plain {N:addedtext}N here');
        check('add, from the model: the bookmark around its text', shape(saved.add.model.xml)[6].replace(/\{\d+:added/, '{N:added').replace(/\}\d+/, '}N'), 'plain {N:addedtext}N here');
        check('add: read again, the bookmarks', [added.infos, saved.add.model.infos], [[...all, 'added'].sort(), [...all, 'added'].sort()]);
        check('add, as read: every bookmark opened once and closed once', balanced(added.xml), true);
        check('add, as read: the other paragraphs are the bytes they were read as', paragraphsOf(added.xml).filter((_, i) => i !== 6), paragraphs.filter((_, i) => i !== 6));
        check('add, as read: the paragraph keeps its paraId, proofing marks and run properties',
            ['w14:paraId="1A2B3C07"', '<w:proofErr w:type="spellStart"/>', '<w:proofErr w:type="spellEnd"/>', '<w:u w:val="single"/>'].map((piece) => paragraphsOf(added.xml)[6].includes(piece)), [true, true, true, true]);
    } catch (e) { fail(e); }
    finish();
};
