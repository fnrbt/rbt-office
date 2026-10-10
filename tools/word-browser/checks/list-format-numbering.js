// Changing the number format of a list does not write numbering.xml anew
// from the model: the level that was changed is written anew in the list's
// own abstract definition, and the other lists, definitions and levels stay
// the text they were read from. A definition that another list uses too is
// copied for the list that was changed (the other list keeps its numbers).
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const bytes = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/lists.docx'));
    const original = await (await window.JSZip.loadAsync(bytes)).file('word/numbering.xml').async('string');
    const element = (xml, pattern) => { const at = xml.search(pattern); if (at < 0) return null; const name = /^<([\w:]+)/.exec(xml.slice(at))[1]; return xml.slice(at, xml.indexOf(`</${name}>`, at) + name.length + 3); };
    const abstractOf = (xml, id) => element(xml, new RegExp(`<w:abstractNum [^>]*w:abstractNumId="${id}"`));
    const numOf = (xml, id) => element(xml, new RegExp(`<w:num w:numId="${id}"`));
    // (The format is set for every paragraph of the level: the editor keeps a format per paragraph, the file one per level.)
    const run = async (lineText, lastText, format) => {
        const editor = Word.initializeCanvas(document.createElement('canvas'));
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
        const line = Word.getLayoutLines(editor).find((l) => (l.Text ?? '').includes(lineText));
        const last = Word.getLayoutLines(editor).find((l) => (l.Text ?? '').includes(lastText));
        Word.setSelection(editor, line.BlockId, 0, last.BlockId, 1);
        Word.setNumberFormat(editor, format);
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const zip = await window.JSZip.loadAsync(out);
        const again = Word.initializeCanvas(document.createElement('canvas'));
        Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(out));
        const a = Word.getLayoutFingerprint(editor), b = Word.getLayoutFingerprint(again);
        const same = JSON.stringify(a) === JSON.stringify(b);
        if (!same) a.forEach((block, i) => { if (JSON.stringify(block) !== JSON.stringify(b[i])) note(`differs: ${JSON.stringify(block)} | ${JSON.stringify(b[i])}`); });
        Word.dispose(again); Word.dispose(editor);
        return { numbering: await zip.file('word/numbering.xml').async('string'), document: await zip.file('word/document.xml').async('string'), same };
    };
    // A list with a definition of its own (numId 4 -> abstractNum 11, upper roman from 3).
    const own = await run('III', 'IV', 'lowerLetter');
    check('own definition: the level is written anew in it', [/<w:numFmt w:val="lowerLetter"\/>/.test(abstractOf(own.numbering, 11) ?? ''), /upperRoman/.test(abstractOf(own.numbering, 11) ?? '')], [true, false]);
    check('own definition: the other definitions are the text they were', [abstractOf(own.numbering, 10) === abstractOf(original, 10), abstractOf(own.numbering, 12) === abstractOf(original, 12)], [true, true]);
    check('own definition: the lists are the text they were', [3, 4, 5, 6].map((id) => numOf(own.numbering, id) === numOf(original, id)), [true, true, true, true]);
    check('own definition: no definition more', (own.numbering.match(/<w:abstractNum /g) ?? []).length, (original.match(/<w:abstractNum /g) ?? []).length);
    check('own definition: read again, laid out as the editor had it', own.same, true);
    // A list that shares its definition (numId 3 and numId 6 -> abstractNum 10).
    const shared = await run('one', 'two', 'upperLetter');
    const newId = /<w:abstractNumId w:val="(\d+)"\/>/.exec(numOf(shared.numbering, 3) ?? '')?.[1];
    check('shared definition: the list has a definition of its own now', [newId !== undefined && newId !== '10', (shared.numbering.match(/<w:abstractNum /g) ?? []).length], [true, (original.match(/<w:abstractNum /g) ?? []).length + 1]);
    const copy = abstractOf(shared.numbering, newId) ?? '';
    check('shared definition: the copy has the new format at its first level and the second level as it was', [/<w:lvl w:ilvl="0">(?:(?!<\/w:lvl>).)*<w:numFmt w:val="upperLetter"\/>/.test(copy), /<w:lvl w:ilvl="1">(?:(?!<\/w:lvl>).)*<w:numFmt w:val="lowerLetter"\/>/.test(copy)], [true, true]);
    check('shared definition: the definition and the other list are the text they were', [abstractOf(shared.numbering, 10) === abstractOf(original, 10), numOf(shared.numbering, 6) === numOf(original, 6), abstractOf(shared.numbering, 11) === abstractOf(original, 11), abstractOf(shared.numbering, 12) === abstractOf(original, 12)], [true, true, true, true]);
    check('shared definition: read again, laid out as the editor had it', shared.same, true);
    note(`numbering.xml: ${original.length} -> ${own.numbering.length} / ${shared.numbering.length} characters`);
} catch (e) { fail(e); }
finish();
