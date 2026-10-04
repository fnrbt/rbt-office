// A click on the placeholder of an object the editor does not model (a chart)
// followed by typing must not cost the object: it is written back as it was
// read unless the user deletes it.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
const opaqueRuns = (pkg) => {
    const out = [];
    const walk = (items) => { for (const item of items || []) {
        for (const run of item.Runs || []) if (run.IsImage && run.OpaqueXml) out.push({ kind: run.OpaqueKind, xml: run.OpaqueXml.length });
        for (const row of item.Rows || []) for (const cell of row.Cells || []) { walk(cell.Paragraphs); walk(cell.NestedTables); }
        walk(item.Paragraphs);
    } };
    walk(pkg.Paragraphs ?? pkg.paragraphs);
    return out;
};
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const pkg = await OF.processors.forEditorPackage(await bytesOf('/repo/tests/test-documents-wild/docx4j-chart.docx'));
    const editor = Word.initialize('c');
    Word.importFromDocxPackage(editor, pkg);
    const before = opaqueRuns(Word.exportDocumentPackage(editor));
    note('opaque runs on import ' + JSON.stringify(before));
    check('the chart is an opaque object', before.length > 0 && before[0].kind, 'chart');
    const textBefore = Word.getText(editor);
    const exported = () => opaqueRuns(Word.exportDocumentPackage(editor));
    // Steps run between the driven gestures (run.mjs "eval").
    window.afterClick = () => {
        check('a click on the chart is not an edit', [Word.canUndo(editor), Word.getText(editor) === textBefore], [false, true]);
        check('the chart survives the click', exported(), before);
    };
    window.afterDragOnItself = () => {
        check('a drag that ends on the chart itself is not a move', Word.canUndo(editor), false);
        check('the chart survives it', exported(), before);
    };
    window.afterTyping = () => {
        check('the typed character is in the text', Word.getText(editor).length, textBefore.length + 1);
        check('the chart survives a click and typing', exported(), before);
    };
    window.finish = () => {
        try {
            // The last gesture dragged the chart off itself: it moved (one undo step) and is still the chart.
            check('a real drag moves the chart', Word.canUndo(editor), true);
            check('the moved chart is still the chart', exported(), before);
            while (Word.canUndo(editor)) Word.undo(editor);
            check('undo restores the text', Word.getText(editor) === textBefore, true);
            check('the chart is still there after undo', exported(), before);
        } catch (e) { fail(e); }
        finish();
    };
    window.drive = [
        ['click', 200, 150], ['eval', 'window.afterClick()'],
        ['down', 200, 150], ['move', 260, 210], ['up'], ['eval', 'window.afterDragOnItself()'],
        ['press', 'Control+Home'], ['type', 'x'], ['eval', 'window.afterTyping()'],
        // Off the object: released in the text line, left of the typed character.
        ['down', 300, 300], ['move', 97, 300], ['up'],
    ];
} catch (e) { fail(e); finish(); }
