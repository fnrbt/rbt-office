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

  // A real package distinguishes inherited fonts from identical direct
  // overrides, including after applying a different named style.
  const fontZip = await window.JSZip.loadAsync(bytes);
  fontZip.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:style>
      <w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="32"/></w:rPr></w:style>
      <w:style w:type="character" w:styleId="Accent"><w:name w:val="Accent"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>
    </w:styles>`);
  fontZip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>abc</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr><w:t>def</w:t></w:r></w:p>
      <w:p><w:r><w:t>other</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  const fontPackage = await OF.processors.forEditorPackage(await fontZip.generateAsync({ type: 'uint8array' }));
  const fontEditor = Word.initializeCanvas(document.createElement('canvas'));
  Word.importFromDocxPackage(fontEditor, fontPackage);
  const fontIds = [...new Set(Word.getLayoutLines(fontEditor).map(line => line.BlockId))];
  const selectFont = (a, b) => Word.setSelection(fontEditor, fontIds[0], a, fontIds[0], b);
  selectFont(0, 3); Word.setParagraphStyle(fontEditor, 'Quote');
  check('paragraph style replaces inherited font and size but retains bold', [Word.getSelectionFormat(fontEditor).FontFamily, Word.getSelectionFormat(fontEditor).FontSizePt, Word.getSelectionFormat(fontEditor).Bold], ['Georgia', 16, true]);
  selectFont(3, 6);
  check('equal-to-old-style direct font and size remain direct', [Word.getSelectionFormat(fontEditor).FontFamily, Word.getSelectionFormat(fontEditor).FontSizePt], ['Calibri', 11]);
  const changedRuns = Word.exportDocumentPackage(fontEditor).Paragraphs[0].Runs;
  check('saved properties distinguish inherited and direct fonts', changedRuns.map(r => [r.FontFamilyInherited, r.ExplicitFontSize == null]), [[true, true], [false, false]]);
  Word.undo(fontEditor); selectFont(1, 1); Word.setCharacterStyle(fontEditor, 'Accent');
  check('character style at a caret supplies the pending size', Word.getSelectionFormat(fontEditor).FontSizePt, 12);
  Word.insertText(fontEditor, 'x'); selectFont(1, 2);
  check('typed character inherits the named character size', [Word.getSelectionFormat(fontEditor).CharacterStyleId, Word.getSelectionFormat(fontEditor).FontSizePt], ['Accent', 12]);
  const fontBlob = await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(fontEditor));
  const fontReopen = Word.initializeCanvas(document.createElement('canvas'));
  Word.importFromDocxPackage(fontReopen, await OF.processors.forEditorPackage(new Uint8Array(await fontBlob.arrayBuffer())));
  const reopenedId = Word.getLayoutLines(fontReopen)[0].BlockId;
  Word.setSelection(fontReopen, reopenedId, 1, reopenedId, 2);
  check('named character style survives save and reopen', [Word.getSelectionFormat(fontReopen).CharacterStyleId, Word.getSelectionFormat(fontReopen).FontSizePt], ['Accent', 12]);
  Word.dispose(fontReopen); Word.dispose(fontEditor);
} catch (e) { fail(e); }
finish();
