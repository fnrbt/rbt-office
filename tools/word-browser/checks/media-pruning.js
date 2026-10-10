// Pictures nothing refers to do not accumulate in the package. A paragraph
// with a picture that is written from the model (its alignment was changed)
// gets a picture file of the writer's own: the file and the relationship it
// was read with go. Saved, opened and changed again, three times over: as
// many picture files as pictures. A picture that was deleted leaves no file;
// the header's picture, and everything of an unedited save, stays.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const original = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/images.docx'));
    const media = async (bytes) => Object.keys((await window.JSZip.loadAsync(bytes)).files).filter((name) => name.startsWith('word/media/') && !name.endsWith('/')).sort();
    const part = async (bytes, name) => (await window.JSZip.loadAsync(bytes)).file(name).async('string');
    const imageRelationships = (rels) => (rels.match(/<Relationship [^>]*relationships\/image"[^>]*\/>/g) ?? []).length;
    const edit = async (bytes, change) => {
        const editor = Word.initializeCanvas(document.createElement('canvas'));
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
        const first = Word.getLayoutLines(editor)[0];
        change(editor, first.BlockId);
        const pictures = Word.getLayoutLines(editor).flatMap((line) => line.Segments ?? []).filter((segment) => segment.Kind === 'image').length;
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        Word.dispose(editor);
        return { out, pictures };
    };
    const before = await media(original);
    check('the fixture: a picture of the body and one of the header', before, ['word/media/body.png', 'word/media/head.png']);
    const unedited = await edit(original, () => { /* nothing */ });
    check('an unedited save: the same picture files', await media(unedited.out), before);
    // The picture's paragraph written from the model, again and again.
    let bytes = original;
    const counts = [];
    for (let round = 0; round < 3; round++) {
        const saved = await edit(bytes, (editor, id) => { Word.setSelection(editor, id, 0, id, 0); if (round % 2 === 0) Word.alignCenter(editor); else Word.alignRight(editor); });
        bytes = saved.out;
        counts.push([(await media(bytes)).length, imageRelationships(await part(bytes, 'word/_rels/document.xml.rels')), saved.pictures]);
    }
    check('written from the model three times over: two picture files, one picture relationship of the body, the picture shown', counts, [[2, 1, 1], [2, 1, 1], [2, 1, 1]]);
    const documentXml = await part(bytes, 'word/document.xml'), rels = await part(bytes, 'word/_rels/document.xml.rels');
    const used = /r:embed="([^"]+)"/.exec(documentXml)?.[1];
    check('the picture\'s relationship is one of the part, to a file of the package', [used !== undefined && rels.includes(`Id="${used}"`), (await media(bytes)).some((name) => rels.includes(name.slice('word/'.length) + '"'))], [true, true]);
    // The picture deleted.
    const deleted = await edit(original, (editor, id) => { Word.setSelection(editor, id, 0, id, 1); Word.insertText(editor, 'gone'); });
    check('the picture deleted: its file and its relationship are gone, the header\'s picture stays', [await media(deleted.out), imageRelationships(await part(deleted.out, 'word/_rels/document.xml.rels')), deleted.pictures], [['word/media/head.png'], 0, 0]);
    const again = await edit(deleted.out, () => { /* reads */ });
    check('and the document reads', (await media(again.out)), ['word/media/head.png']);
} catch (e) { fail(e); }
finish();
