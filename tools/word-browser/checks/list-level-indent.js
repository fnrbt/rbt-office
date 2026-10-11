// The indents of a numbered paragraph: what the paragraph states itself
// counts first, then the level of the numbering the paragraph itself refers
// to (over its paragraph style's indent), then the style.
// file-sample.docx: "Simple Tables" and "Complex Tables" are
// <w:pStyle w:val="ListParagraph"/> (the style: <w:ind w:left="720"/>) with
// <w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>, and level 0 of
// that numbering has <w:ind w:left="1080" w:hanging="360"/>: the text stands
// 1080 twips = 72 px (0.75 in) right of the margin, the bullet 360 twips =
// 24 px before it. They used to stand at the style's 720 twips = 48 px.
import { check, note, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const canvas = document.getElementById('c');
    const bytes = await bytesOf('/repo/tests/test-documents-online/file-sample.docx');
    const open = async (data) => { const editor = Word.initializeCanvas(canvas); Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(data)); return editor; };
    const textX = (editor, text) => {
        const line = Word.getLayoutLines(editor).find((item) => (item.Text ?? '').startsWith(text));
        const segment = line.Segments.find((item) => (item.Text ?? '').startsWith(text.slice(0, 4)));
        const frame = Word.getPageFrames(editor)[line.Page - 1];
        return Math.round(((segment ? segment.X : line.X) - frame.MarginLeft) * 10) / 10;
    };
    const editor = await open(bytes);
    note('segments ' + JSON.stringify(Word.getLayoutLines(editor).find((item) => (item.Text ?? '').startsWith('Simple Tables')).Segments.map((s) => [s.Text, s.X])));
    check('the bullets\' text stands at the level\'s 0.75 in', [textX(editor, 'Simple Tables'), textX(editor, 'Complex Tables')], [72, 72]);
    const pkg = await OF.processors.forEditorPackage(bytes);
    const item = pkg.Paragraphs.find((para) => (para.Runs ?? []).some((run) => (run.Text ?? '').startsWith('Simple Tables')));
    check('read: the level\'s indents (px), none stated by the paragraph', [item.IndentLeft, item.FirstLineIndent, item.ExplicitIndentLeft ?? null, item.ListLevelIndentLeft], [72, -24, null, 72]);
    // Saved as it was read: the paragraph gets no w:ind of its own.
    const saved = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
    const savedXml = await (await window.JSZip.loadAsync(saved)).file('word/document.xml').async('string');
    const at = savedXml.indexOf('Simple Tables');
    const paragraph = savedXml.slice(savedXml.lastIndexOf('<w:p ', at), at);
    check('saved: the paragraph states no indent', [paragraph.includes('<w:numId w:val="2"/>'), paragraph.includes('<w:ind')], [true, false]);
    // A paragraph that states its own left indent keeps it, with the level's hanging indent.
    const zip = await window.JSZip.loadAsync(bytes);
    const xml = await zip.file('word/document.xml').async('string');
    const numbered = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">Simple Tables';
    if (!xml.includes(numbered)) throw new Error('the document changed');
    zip.file('word/document.xml', xml.replace(numbered, numbered.replace('</w:numPr>', '</w:numPr><w:ind w:left="2010"/>')));
    const own = await open(await zip.generateAsync({ type: 'uint8array' }));
    check('a paragraph\'s own left indent counts first; the next paragraph has the level\'s', [textX(own, 'Simple Tables'), textX(own, 'Complex Tables')], [134, 72]);
} catch (e) { fail(e); }
finish();
