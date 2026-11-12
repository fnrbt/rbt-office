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
    // A split instance must clone the source, not regenerate its level. Model
    // restart/continue scope is exercised by tests/word/list-levels.fsx; this
    // exercises the DOCX bridge with its actual preserved package.
    const fixture = await window.JSZip.loadAsync(bytes);
    const sourceDefinition = abstractOf(original, 10);
    if (!sourceDefinition) throw new Error('list fixture lacks abstract numbering 10');
    const augmentedDefinition = sourceDefinition
        .replace('<w:lvl w:ilvl="0">', '<w:lvl w:ilvl="0"><w:lvlRestart w:val="0"/>')
        .replace(/(<w:lvl\b[^>]*w:ilvl="1"[^>]*>[\s\S]*?<w:lvlText\b[^>]*w:val=")[^"]*"/,
            (_match, prefix) => prefix + '%1.%2)"');
    if (augmentedDefinition === sourceDefinition || !augmentedDefinition.includes('w:val="%1.%2)"')) throw new Error('cannot construct nested numbering template');
    const augmented = original.replace(sourceDefinition, augmentedDefinition)
        .replace('</w:numbering>', '<w:num w:numId="99"><w:abstractNumId w:val="10"/></w:num></w:numbering>');
    fixture.file('word/numbering.xml', augmented);
    const augmentedBytes = await fixture.generateAsync({ type: 'uint8array' });
    const splitEditor = Word.initializeCanvas(document.createElement('canvas'));
    Word.importFromDocxPackage(splitEditor, await OF.processors.forEditorPackage(augmentedBytes));
    const splitPackage = Word.exportDocumentPackage(splitEditor);
    splitPackage.Unchanged = false;
    const split = splitPackage.Paragraphs.find((p) => (p.Runs ?? []).map((r) => r.Text ?? '').join('') === 'two');
    if (!split) throw new Error('list fixture has no two paragraph');
    Object.assign(split, { ListId: '100', ListSourceId: '3', ListRestart: true, ListStart: 7,
        ListNumberFormat: 'upperLetter', ListLevelText: '%1.', ListFormatEdited: true, SourceUnchanged: false, SourcePara: null, SourceTextChanges: null });
    const splitBytes = new Uint8Array(await (await OF.word.writeFromEditorPackage(splitPackage)).arrayBuffer());
    const splitZip = await window.JSZip.loadAsync(splitBytes);
    const splitNumbering = await splitZip.file('word/numbering.xml').async('string');
    const splitNum = numOf(splitNumbering, 100) ?? '';
    const splitAbstractId = /<w:abstractNumId w:val="(\d+)"\/>/.exec(splitNum)?.[1];
    const splitAbstract = abstractOf(splitNumbering, splitAbstractId) ?? '';
    check('split clone retains other lists and unused IDs',
        [numOf(splitNumbering, 3) === numOf(augmented, 3), numOf(splitNumbering, 6) === numOf(augmented, 6), numOf(splitNumbering, 99) === numOf(augmented, 99)], [true, true, true]);
    check('split clone keeps level metadata and authors marker and restart',
        [/w:lvlRestart w:val="0"/.test(splitAbstract), /w:numFmt w:val="upperLetter"/.test(splitAbstract), /w:startOverride w:val="7"/.test(splitNum)], [true, true, true]);
    const splitAgain = Word.initializeCanvas(document.createElement('canvas'));
    Word.importFromDocxPackage(splitAgain, await OF.processors.forEditorPackage(splitBytes));
    const reopened = Word.exportDocumentPackage(splitAgain).Paragraphs.find((p) => (p.Runs ?? []).map((r) => r.Text ?? '').join('') === 'two');
    check('split restart survives save and reopen', [String(reopened?.ListId), reopened?.ListNumberFormat, reopened?.ListStart], ['100', 'upperLetter', 7]);
    const continuePackage = Word.exportDocumentPackage(splitAgain);
    continuePackage.Unchanged = false;
    const first = continuePackage.Paragraphs.find((p) => (p.Runs ?? []).map((r) => r.Text ?? '').join('') === 'one');
    const continued = continuePackage.Paragraphs.find((p) => (p.Runs ?? []).map((r) => r.Text ?? '').join('') === 'two');
    Object.assign(continued, { ListId: first.ListId, ListNumberFormat: first.ListNumberFormat, ListLevelText: first.ListLevelText,
        ListStart: first.ListStart, ListSourceId: null, ListRestart: false, ListFormatEdited: false, SourceUnchanged: false, SourcePara: null, SourceTextChanges: null });
    const continuedBytes = new Uint8Array(await (await OF.word.writeFromEditorPackage(continuePackage)).arrayBuffer());
    const continuedPackage = await OF.processors.forEditorPackage(continuedBytes);
    const continuedAgain = continuedPackage.Paragraphs.find((p) => (p.Runs ?? []).map((r) => r.Text ?? '').join('') === 'two');
    check('continued list identity survives another save and reopen',
        [String(continuedAgain?.ListId), continuedAgain?.ListNumberFormat], [String(first.ListId), first.ListNumberFormat]);
    const ordinalPackage = Word.exportDocumentPackage(splitEditor);
    ordinalPackage.Unchanged = false;
    const findText = (pkg, text) => pkg.Paragraphs.find((p) => (p.Runs ?? []).map((r) => r.Text ?? '').join('') === text);
    const ordinalTwo = findText(ordinalPackage, 'two');
    const ordinalNested = findText(ordinalPackage, 'nested a');
    if (!ordinalTwo || !ordinalNested) throw new Error('list fixture lacks ordinal paragraphs');
    Object.assign(ordinalTwo, { ListId: '101', ListSourceId: '3', ListRestart: true, ListStart: 2,
        ListNumberFormat: 'lowerLetter', ListLevelText: '%1.', ListFormatEdited: true,
        ListStartOverrides: [{ Level: 1, Start: 2 }], SourceUnchanged: false, SourcePara: null, SourceTextChanges: null });
    Object.assign(ordinalNested, { ListId: '102', ListSourceId: '3', ListRestart: true, ListStart: 1,
        ListStartOverrides: [{ Level: 1, Start: 2 }, { Level: 2, Start: 1 }],
        SourceUnchanged: false, SourcePara: null, SourceTextChanges: null });
    const ordinalBytes = new Uint8Array(await (await OF.word.writeFromEditorPackage(ordinalPackage)).arrayBuffer());
    const ordinalAgain = Word.initializeCanvas(document.createElement('canvas'));
    Word.importFromDocxPackage(ordinalAgain, await OF.processors.forEditorPackage(ordinalBytes));
    const markers = Word.getLayoutLines(ordinalAgain);
    check('physical counter snapshots survive save and reopen including nested parent',
        ['one', 'two', 'nested a'].map((text) => markers.find((line) => line.Text === text)?.ListMarker?.trim()),
        ['1.', 'b.', '2.a)']);
    Word.dispose(ordinalAgain);
    Word.dispose(splitAgain); Word.dispose(splitEditor);
    note(`numbering.xml: ${original.length} -> ${own.numbering.length} / ${shared.numbering.length} characters`);
} catch (e) { fail(e); }
finish();
