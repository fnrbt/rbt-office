// Item: a run can switch its style's formatting off. Uses the round-trip fixture
// synthetic/explicit-off (a bold, centred paragraph style; one run with w:b w:val="0").
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
  const OF = await loadFormats();
  const Word = await loadWord();
  const bytes = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/explicit-off.docx'));
  const pkg = await OF.processors.forEditorPackage(bytes);
  const editor = Word.initialize('c');
  Word.importFromDocxPackage(editor, pkg);
  const runs = () => Word.exportDocumentPackage(editor).Paragraphs[0].Runs.map((r) => [r.Text, r.Bold, r.ExplicitBold ?? null]);
  check('imported: bold through style, nothing explicit', runs().slice(0, 1), [['inherits ', true, null]]);
  check('canUndo before', Word.canUndo(editor), false);
  Word.selectAll(editor);
  Word.toggleBold(editor);   // selection has a non-bold run -> all bold
  check('first toggle makes everything bold', Word.exportDocumentPackage(editor).Paragraphs[0].Runs.every((r) => r.Bold === true), true);
  Word.toggleBold(editor);   // all bold (first run through its style) -> explicit off
  const off = Word.exportDocumentPackage(editor).Paragraphs[0].Runs.map((r) => [r.Bold, r.ExplicitBold ?? null]);
  check('toggle off under bold style is explicit non-bold', off.every(([b, e]) => b === false && e === false), true);
  const blob = await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor));
  const zip = await window.JSZip.loadAsync(await blob.arrayBuffer());
  const xml = await zip.files['word/document.xml'].async('string');
  check('saved xml has w:b w:val=0 on the Loud paragraph', /<w:pStyle w:val="Loud"\/><\/w:pPr><w:r><w:rPr>(<w:rFonts[^>]*\/>)?<w:b w:val="0"\/>/.test(xml), true);
  const reread = await OF.processors.forEditorPackage(new Uint8Array(await blob.arrayBuffer()));
  check('reader sees non-bold after save', reread.Paragraphs[0].Runs.map((r) => r.Bold), reread.Paragraphs[0].Runs.map(() => false));
  check('undo 1', Word.undo(editor), true);
  check('after undo: all bold again', Word.exportDocumentPackage(editor).Paragraphs[0].Runs.every((r) => r.Bold === true), true);
  check('undo 2', Word.undo(editor), true);
  check('after second undo: original runs', runs().map((r) => [r[0], r[1]]), [['inherits ', true], ['quiet', false]]);
  check('nothing left to undo', Word.canUndo(editor), false);
  check('redo twice', [Word.redo(editor), Word.redo(editor)], [true, true]);
  check('after redo: explicit non-bold', Word.exportDocumentPackage(editor).Paragraphs[0].Runs.every((r) => r.Bold === false && r.ExplicitBold === false), true);
  // Explicit left under the centred style.
  Word.alignLeft(editor);
  check('align left is explicit', Word.exportDocumentPackage(editor).Paragraphs.map((p) => p.ExplicitAlignment ?? null), ['left', 'left', 'left']);
} catch (e) { fail(e); }
finish();
