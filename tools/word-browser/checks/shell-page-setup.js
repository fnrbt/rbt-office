// Page setup edits only the selected section and patches imported section XML.
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
let dispose;
try {
    const [Word, OF, Host, Shell, Dom] = await Promise.all([
        loadWord(), loadFormats(), import('/word/WordHost.js'), import('/word/WordShell.js'), import('/word/FableUI/Dom.js')
    ]);
    const root = document.createElement('div'); root.id = 'shell-page-setup'; document.body.appendChild(root);
    const host = Host.create('', 'word-shell');
    dispose = Dom.mountApp('#shell-page-setup', Shell.view(host));
    const editor = host.Editor;
    const control = id => document.querySelector(`[data-testid="${id}"]`);
    const settle = () => new Promise(resolve => requestAnimationFrame(resolve));
    const click = async id => { control(id).click(); await settle(); };
    const change = async (id, value, event = 'input') => { const el = control(id); el.value = value; el.dispatchEvent(new Event(event, { bubbles: true })); await settle(); };
    const open = async () => { Host.openPageSetup(host); await settle(); };
    const selectText = text => { const line = Word.getLayoutLines(editor).find(x => (x.Text ?? '').startsWith(text)); Word.setSelection(editor, line.BlockId, 0, line.BlockId, 0); };
    const saveXml = async () => { const blob = await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor)); const zip = await JSZip.loadAsync(await blob.arrayBuffer()); return zip.file('word/document.xml').async('string'); };
    const sections = xml => [...new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'sectPr')];
    const child = (node, name) => [...node.children].find(x => x.localName === name);
    const attr = (node, name) => node.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', name);
    Word.setParagraphs(editor, ['Fresh']);
    selectText('Fresh');
    let revision = Word.getRevision(editor);
    await open(); await click('word-page-setup-apply');
    check('unchanged setup has no history', Word.getRevision(editor), revision);
    await open(); await change('word-page-width', '0'); await click('word-page-setup-apply');
    check('invalid setup stays open without an edit', [!!control('word-page-setup-error').textContent, Word.getRevision(editor)], [true, revision]);
    await click('word-page-setup-cancel');
    await open(); await change('word-page-preset', 'legal', 'change'); await change('word-page-orientation', 'landscape', 'change'); await click('word-page-setup-apply');
    check('preset and orientation affect the page', [Word.getPageFrames(editor)[0].Width, Word.getPageFrames(editor)[0].Height], [1344, 816]);
    await open();
    check('reopening landscape setup preserves orientation selection', control('word-page-orientation').value, 'landscape');
    await click('word-page-setup-cancel');
    Word.undo(editor);
    check('one undo restores fresh setup', Word.getRevision(editor), revision);
    await open(); await change('word-page-width', '14');
    check('custom width updates preset and orientation', [control('word-page-preset').value, control('word-page-orientation').value], ['custom', 'landscape']);
    await change('word-page-orientation', 'portrait', 'change');
    check('orientation swaps custom dimensions', [Number(control('word-page-width').value), Number(control('word-page-height').value)], [11, 14]);
    await click('word-page-setup-cancel');
    await open(); await change('word-margin-top', '1.25'); Word.insertText(editor, 'X'); const staleRevision = Word.getRevision(editor);
    await click('word-page-setup-apply');
    check('stale form cannot mutate the document', [!!control('word-page-setup-error').textContent, Word.getRevision(editor)], [true, staleRevision]);
    await click('word-page-setup-cancel'); Word.undo(editor);
    await open(); await change('word-margin-top', '1.25'); Word.getInputElement(editor).readOnly = true;
    await click('word-page-setup-apply');
    check('readonly prevents application', [!!control('word-page-setup-error').textContent, Word.getRevision(editor)], [true, revision]);
    Word.getInputElement(editor).readOnly = false; await click('word-page-setup-cancel');

    const zip = await JSZip.loadAsync(await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/three-sections.docx')));
    let xml = await zip.file('word/document.xml').async('string');
    xml = xml.replace(/<w:sectPr([^>]*)>/g, '<w:sectPr$1 w:rsidSect="00AABBCC"><w:docGrid w:linePitch="360"/>');
    xml = xml.replace('w:left="1440"', 'w:left="1441"');
    zip.file('word/document.xml', xml);
    Host.importDocumentPackage(host, await OF.processors.forEditorPackage(await zip.generateAsync({ type: 'uint8array' })), 'sections.docx');
    Word.enterHeaderFooter(editor, 0, 'header'); const originalHeaderY = Word.getCaretRect(editor).Y; Word.exitHeaderFooter(editor);
    Word.enterHeaderFooter(editor, 1, 'header'); const otherHeaderY = Word.getCaretRect(editor).Y - Word.getPageFrames(editor)[1].Top; Word.exitHeaderFooter(editor);
    selectText('section one'); revision = Word.getRevision(editor);
    await open(); await change('word-margin-top', '1.25'); await change('word-header-distance', '1'); await click('word-page-setup-apply');
    Word.enterHeaderFooter(editor, 0, 'header'); const changedHeaderY = Word.getCaretRect(editor).Y; Word.exitHeaderFooter(editor);
    Word.enterHeaderFooter(editor, 1, 'header'); const otherHeaderAfter = Word.getCaretRect(editor).Y - Word.getPageFrames(editor)[1].Top; Word.exitHeaderFooter(editor);
    check('editable header follows its section distance only', [Math.round(changedHeaderY - originalHeaderY), Math.round(otherHeaderAfter - otherHeaderY)], [48, 0]);
    let saved = sections(await saveXml());
    check('first section patch retains unknown metadata and untouched precise margin', [attr(saved[0], 'rsidSect'), attr(child(saved[0], 'docGrid'), 'linePitch'), attr(child(saved[0], 'pgMar'), 'left'), attr(child(saved[0], 'pgMar'), 'top')], ['00AABBCC', '360', '1441', '1800']);
    check('other sections retain metadata', saved.slice(1).map(s => attr(child(s, 'docGrid'), 'linePitch')), ['360', '360']);
    Word.undo(editor); check('section edit is one undo', Word.getRevision(editor), revision);
    selectText('section three'); await open(); await change('word-margin-bottom', '1.5'); await click('word-page-setup-apply');
    saved = sections(await saveXml());
    check('final section patch retains unknown metadata', [attr(saved[2], 'rsidSect'), attr(child(saved[2], 'docGrid'), 'linePitch'), attr(child(saved[2], 'pgMar'), 'bottom')], ['00AABBCC', '360', '2160']);
    const savedBlob = await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor));
    Host.importDocumentPackage(host, await OF.processors.forEditorPackage(new Uint8Array(await savedBlob.arrayBuffer())), 'reopened.docx');
    selectText('section three'); await open();
    check('reopened last section prefills the saved distance', Number(control('word-margin-bottom').value), 1.5);
    await click('word-page-setup-cancel');
    dispose(); dispose = null;
} catch (e) { fail(e); }
finally { if (dispose) dispose(); finish(); }
