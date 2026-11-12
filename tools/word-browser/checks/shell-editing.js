// Selection-only shell commands, live chrome, mixed paragraph patches and find.
import { check, finish, fail, loadWord, loadFormats } from './common.js';
let dispose;
try {
    const [Word, OF, Host, Shell, Dom] = await Promise.all([
        loadWord(), loadFormats(), import('/word/WordHost.js'), import('/word/WordShell.js'), import('/word/FableUI/Dom.js')
    ]);
    const root = document.createElement('div'); root.id = 'shell-editing'; document.body.appendChild(root);
    const host = Host.create('', 'word-shell');
    dispose = Dom.mountApp('#shell-editing', Shell.view(host));
    const editor = host.Editor;
    const control = id => document.querySelector(`[data-testid="${id}"]`);
    const settle = () => new Promise(resolve => requestAnimationFrame(resolve));
    const click = async id => { await settle(); control(id).click(); await settle(); };
    const change = async (id, value, event = 'input') => {
        const input = control(id); input.value = value; input.dispatchEvent(new Event(event, { bubbles: true }));
        await settle();
    };
    const title = () => control('word-document-title').textContent;
    const pressed = id => control(id).getAttribute('aria-pressed');
    const blocksOf = () => [...new Set(Word.getLayoutLines(editor).map(line => line.BlockId))];
    Word.setParagraphs(editor, ['Target target', 'other TARGET']);
    Word.selectAll(editor); Word.clearFormatting(editor); Word.alignLeft(editor);
    let blocks = blocksOf();
    const select = (index, start = 0, end = start) => Word.setSelection(editor, blocks[index], start, blocks[index], end);
    const formatAt = (index, start, end) => { select(index, start, end); return Word.getSelectionFormat(editor); };
    const paragraphDialog = async () => { await click('word-format-menu'); await click('word-menu-paragraph-style'); };
    select(0, 0, 6);
    const beforeStylesSearch = Word.getRevision(editor);
    await click('word-styles-toggle');
    check('opening styles focuses the visible search control', document.activeElement === control('word-styles-search'), true);
    await change('word-styles-search', 'heading 2');
    check('styles search filters without editing the document',
        [document.querySelectorAll('[data-testid="word-style-choice"]').length, Word.getRevision(editor)],
        [1, beforeStylesSearch]);
    const headingChoice = document.querySelector('[data-testid="word-style-choice"][data-style-id="heading-2"]');
    headingChoice.click(); await settle();
    check('styles pane applies to the retained selection', [Word.getSelectionFormat(editor).Paragraph.HeadingLevel, Word.getSelection(editor).IsCollapsed], [2, false]);
    Word.undo(editor);
    check('one undo restores the paragraph style', Word.getSelectionFormat(editor).Paragraph.HeadingLevel, 0);
    const search = control('word-styles-search'); search.focus();
    search.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
    await settle();
    check('Escape closes styles and restores focus', [control('word-styles-pane').style.display, document.activeElement === control('word-styles-toggle')], ['none', true]);
    select(1); Word.alignCenter(editor); Word.markSaved(editor, Word.getRevision(editor));
    select(0, 0, 6); await click('word-bold');
    check('selected text formatting updates the live toolbar and dirty state',
        [Word.getSelectionFormat(editor).Bold, pressed('word-bold'), title().endsWith(' *')], [true, 'true', true]);
    check('toolbar bold leaves unselected text unchanged', formatAt(0, 7, 12).Bold, false);
    await settle();
    check('caret and selection changes update the bold button', pressed('word-bold'), 'false');
    Word.undo(editor); select(0, 0, 6);
    await settle();
    check('one undo restores formatting and saved state', [Word.getSelectionFormat(editor).Bold, pressed('word-bold'), Word.isModified(editor)], [false, 'false', false]);
    Word.redo(editor); select(0, 0, 12);
    await settle();
    check('mixed selected runs are shown as mixed', pressed('word-bold'), 'mixed');
    select(0, 0, 6); await click('word-bold');
    check('a checked bold command toggles off rather than becoming a no-op', [Word.getSelectionFormat(editor).Bold, pressed('word-bold')], [false, 'false']);
    Word.undo(editor);

    select(0, 0, 6); const otherSize = formatAt(1, 0, 5).FontSizePt; select(0, 0, 6);
    await change('word-font-size', '18');
    check('font-size control changes only the selected text', [Word.getSelectionFormat(editor).FontSizePt, formatAt(1, 0, 5).FontSizePt], [18, otherSize]);
    select(0, 0, 6);
    await settle();
    check('font-size display follows the current selection', control('word-font-size').value, '18');
    Word.undo(editor); select(0, 0, 6);
    await settle();
    check('undo refreshes the displayed font size', control('word-font-size').value, String(otherSize));

    Word.setSelection(editor, blocks[0], 0, blocks[1], 12);
    await settle();
    check('mixed paragraph alignment is visible', pressed('word-align-center'), 'mixed');
    await paragraphDialog();
    check('mixed paragraph values are not replaced by dialog defaults', control('word-paragraph-align').value, '');
    const beforeNoop = Word.getRevision(editor);
    await click('word-paragraph-apply');
    check('applying an unchanged mixed dialog makes no edit', Word.getRevision(editor), beforeNoop);
    const originalSpacing = [formatAt(0, 0, 6).Paragraph.SpaceBefore, formatAt(1, 0, 5).Paragraph.SpaceBefore];
    Word.setSelection(editor, blocks[0], 0, blocks[1], 12);
    await paragraphDialog(); await change('word-paragraph-spacing-before', '12'); await click('word-paragraph-apply');
    check('one changed paragraph field preserves differing alignment',
        [formatAt(0, 0, 6).Paragraph.Alignment, formatAt(1, 0, 5).Paragraph.Alignment,
         formatAt(0, 0, 6).Paragraph.SpaceBefore, formatAt(1, 0, 5).Paragraph.SpaceBefore], ['left', 'center', 12, 12]);
    Word.undo(editor);
    check('partial paragraph changes undo together', [formatAt(0, 0, 6).Paragraph.SpaceBefore, formatAt(1, 0, 5).Paragraph.SpaceBefore], originalSpacing);
    select(1); await paragraphDialog();
    check('reopening follows the current paragraph', control('word-paragraph-align').value, 'center');
    await click('word-paragraph-cancel');

    const blob = await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor));
    const packageToOpen = await OF.processors.forEditorPackage(new Uint8Array(await blob.arrayBuffer()));
    Word.setParagraphs(editor, ['Replacement']); Word.markSaved(editor, Word.getRevision(editor));
    Host.importDocumentPackage(host, packageToOpen, 'formatting.docx');
    blocks = blocksOf(); select(0, 0, 6);
    await settle();
    check('package replacement refreshes filename, text and selection chrome',
        [title(), Word.getText(editor), pressed('word-bold')], ['formatting.docx', 'Target target\nother TARGET', 'true']);

    const proxy = Word.getInputElement(editor);
    const sendKey = (target, key, options = {}) => {
        const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...options });
        target.dispatchEvent(event); return event.defaultPrevented;
    };
    proxy.focus();
    check('Ctrl+B executes once on the owning editor', [sendKey(proxy, 'b', { ctrlKey: true }), Word.getSelectionFormat(editor).Bold], [true, false]);
    Word.undo(editor); select(0, 0, 6);
    const outside = document.createElement('input'); document.body.appendChild(outside);
    check('unrelated controls retain their shortcuts', [sendKey(outside, 's', { ctrlKey: true }), sendKey(outside, 'f', { ctrlKey: true }), host.Saving, control('word-find-query')], [false, false, false, null]);
    outside.remove();
    check('composition does not dispatch formatting', [sendKey(proxy, 'b', { ctrlKey: true, isComposing: true }), Word.getSelectionFormat(editor).Bold], [false, true]);
    const findHandled = sendKey(proxy, 'f', { ctrlKey: true });
    await settle();
    check('Ctrl+F opens the owned find dialog', [findHandled, control('word-find-query') !== null], [true, true]);
    check('find inputs do not trigger document formatting', [sendKey(control('word-find-query'), 'b', { ctrlKey: true }), Word.getSelectionFormat(editor).Bold], [false, true]);
    await change('word-find-query', 'target');
    const range = () => { const s = Word.getSelection(editor); return [s.AnchorBlockId, s.AnchorOffset, s.FocusOffset]; };
    await click('word-find-next'); check('find starts after the selection', range(), [blocks[0], 7, 13]);
    await click('word-find-next'); check('find advances into the next paragraph', range(), [blocks[1], 6, 12]);
    await click('word-find-next'); check('find wraps to the first match', range(), [blocks[0], 0, 6]);
    Word.setSelection(editor, blocks[1], 0, blocks[0], 0);
    await click('word-find-next'); check('find uses the normalized end of a backwards selection', range(), [blocks[1], 6, 12]);
    select(1, 0, 5); Word.insertText(editor, 'live');
    await change('word-find-query', 'live'); await click('word-find-next');
    check('find sees edits made while its panel is open', range(), [blocks[1], 0, 4]);
    Word.undo(editor); await change('word-find-query', 'target'); await change('word-find-replacement', 'done');
    await click('word-replace-all');
    check('replace-all changes actual document text', Word.getText(editor), 'done done\nother done');
    Word.undo(editor); check('replace-all undoes as one transaction', Word.getText(editor), 'Target target\nother TARGET');
    Word.redo(editor); check('replace-all redo restores its text', Word.getText(editor), 'done done\nother done');
    dispose(); dispose = null;
    check('disposal releases keyboard handling', sendKey(root, 'f', { ctrlKey: true }), false);
} catch (e) { fail(e); }
finally { if (dispose) dispose(); finish(); }
