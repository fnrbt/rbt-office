// The shell's outline view (View > Outline): the document as Word's outline over the pages - a line per
// block, indented by level, with a plus / minus / body-text symbol - and the outlining toolbar:
// promote to Heading 1, promote, a level, demote, demote to body text, move up / down (a collapsed
// heading with its subtext), expand / collapse (the view only), Show level, first line only, text
// formatting; Word's keys (Alt+Shift+arrows, Plus / Minus, 1-9); a double click or Enter back to print
// layout with the caret on the line; Close outline view; a macro's ActiveWindow.View.Type and
// ShowHeading.
import { check, finish, fail, loadWord } from './common.js';
let dispose;
try {
    const [Word, Host, Shell, Dom] = await Promise.all([
        loadWord(), import('/word/WordHost.js'), import('/word/WordShell.js'), import('/word/FableUI/Dom.js')
    ]);
    const root = document.createElement('div'); root.id = 'outline-view'; root.style.height = '700px'; document.body.appendChild(root);
    const host = Host.create('', 'word-shell');
    dispose = Dom.mountApp('#outline-view', Shell.view(host));
    const editor = host.Editor;
    const control = id => root.querySelector(`[data-testid="${id}"]`);
    const settle = () => new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 0)));
    const click = async (element) => { await settle(); element.click(); await settle(); };
    Word.setParagraphs(editor, ['Intro', 'Alpha', 'alpha body', 'Alpha one', 'one body', 'Beta', 'beta body']);
    const blocks = Word.getOutlineLines(editor, false).map(l => l.BlockId);
    const caret = (i) => Word.setSelection(editor, blocks[i], 0, blocks[i], 0);
    for (const [i, n] of [[1, 1], [3, 2], [5, 1]]) { caret(i); Word.setOutlineLevel(editor, n); }
    caret(3);
    check('outline view hidden in print layout', control('word-outline-view').style.display, 'none');
    await click(control('word-view-menu'));
    await click(control('word-menu-view-outline'));
    const view = control('word-outline-view'), lines = control('word-outline-lines');
    const rows = () => [...root.querySelectorAll('[data-testid="word-outline-line"]')];
    const textOf = r => r.querySelector('.rb-word-outline-text').textContent;
    check('View > Outline: outline view over the pages', [view.style.display, Word.getView(editor), view.getBoundingClientRect().top >= control('word-shell-host').getBoundingClientRect().top - 1], ['flex', 2, true]);
    await click(control('word-view-menu'));
    check('the menu checks Outline', [control('word-menu-view-outline').getAttribute('aria-pressed'), control('word-menu-view-print').getAttribute('aria-pressed')], ['true', 'false']);
    await click(control('word-view-menu'));
    check('a line per block', rows().map(textOf), ['Intro', 'Alpha', 'alpha body', 'Alpha one', 'one body', 'Beta', 'beta body']);
    check('Word\'s symbols: plus with subtext, body text', rows().map(r => r.querySelector('[data-testid="word-outline-symbol"]').getAttribute('data-symbol')), ['body', 'plus', 'body', 'plus', 'body', 'plus', 'body']);
    const indent = r => parseInt(r.style.paddingLeft);
    check('indented by level, body text under its heading', rows().map(indent), [30, 30, 54, 54, 78, 30, 54]);
    check('the line of the caret is selected', rows().filter(r => r.getAttribute('aria-selected') === 'true').map(textOf), ['Alpha one']);
    check('the level box says its level', control('word-outline-level').value, '2');
    check('text formatting shown: the heading style\'s size', rows()[1].querySelector('span span').style.fontSize, '16pt');
    check('the lines have the focus', document.activeElement === lines, true);

    const down = (row, extra = {}) => row.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0, ...extra }));
    const key = (k, extra = {}) => lines.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...extra }));
    const selectedTexts = () => rows().filter(r => r.getAttribute('aria-selected') === 'true').map(textOf);
    const docTexts = () => Word.getOutlineLines(editor, false).map(l => [l.Text, l.Level]);
    // Promote and demote: the toolbar, Word's keys (through the shell's commands), the level box.
    down(rows()[3]); await settle();
    key('ArrowLeft', { altKey: true, shiftKey: true, code: 'ArrowLeft' }); await settle();
    check('Alt+Shift+Left promotes the selected line', rows()[3].getAttribute('data-level'), '1');
    key('ArrowRight', { altKey: true, shiftKey: true, code: 'ArrowRight' }); await settle();
    check('Alt+Shift+Right demotes it', rows()[3].getAttribute('data-level'), '2');
    const revision = Word.getRevision(editor);
    await click(control('word-outline-promote-h1'));
    check('Promote to Heading 1', [rows()[3].getAttribute('data-level'), Word.getStyles(editor).find(s => s.Id === 'Heading1').HeadingLevel], ['1', 1]);
    await click(control('word-outline-demote-body'));
    check('Demote to body text: a body line, its symbol', [rows()[3].getAttribute('data-level'), rows()[3].querySelector('[data-testid="word-outline-symbol"]').getAttribute('data-symbol')], ['10', 'body']);
    const level = control('word-outline-level'); level.value = '3'; level.dispatchEvent(new Event('change', { bubbles: true })); await settle();
    check('the level box: Level 3', rows()[3].getAttribute('data-level'), '3');
    await click(control('word-outline-promote'));
    await click(control('word-outline-demote'));
    check('promote, demote from the toolbar', rows()[3].getAttribute('data-level'), '3');
    Word.undo(editor); Word.undo(editor); Word.undo(editor); Word.undo(editor); Word.undo(editor); await settle();
    check('each one undo step', [rows()[3].getAttribute('data-level'), Word.getRevision(editor)], ['2', revision]);

    // Selection of lines: click, Shift+click, the arrows; a range promoted (body text among headings stays).
    down(rows()[1]); down(rows()[4], { shiftKey: true }); await settle();
    check('Shift+click selects the lines between', selectedTexts(), ['Alpha', 'alpha body', 'Alpha one', 'one body']);
    check('the document\'s selection is of them', [Word.getSelection(editor).AnchorBlockId, Word.getSelection(editor).FocusBlockId], [blocks[1], blocks[4]]);
    await click(control('word-outline-demote'));
    check('a range demoted: its headings, not its body text', docTexts().map(t => t[1]), [10, 2, 10, 3, 10, 1, 10]);
    Word.undo(editor); await settle();
    key('ArrowDown'); key('ArrowDown', { shiftKey: true }); await settle();
    check('Down, Shift+Down', selectedTexts(), ['Beta', 'beta body']);

    // Move up / down: a line; a collapsed heading with its subtext.
    down(rows()[3]); await settle();
    key('ArrowUp', { altKey: true, shiftKey: true, code: 'ArrowUp' }); await settle();
    check('Alt+Shift+Up moves the line above the line before it', rows().map(textOf), ['Intro', 'Alpha', 'Alpha one', 'alpha body', 'one body', 'Beta', 'beta body']);
    check('the line stays selected', selectedTexts(), ['Alpha one']);
    Word.undo(editor); await settle();
    down(rows()[1]); await settle();
    key('_', { altKey: true, shiftKey: true, code: 'Minus' }); await settle();
    check('Alt+Shift+Minus collapses: the subtext hidden, the heading marked', [rows().map(textOf), rows()[1].classList.contains('rb-word-outline-collapsed')], [['Intro', 'Alpha', 'Beta', 'beta body'], true]);
    check('collapsing is not a change of the document', Word.getRevision(editor), revision);
    await click(control('word-outline-move-down'));
    // (Past the next line shown - Beta's line, not its body text: as Word moves in outline view.)
    check('a collapsed heading moves down with its subtext past the next line', docTexts().map(t => t[0]), ['Intro', 'Beta', 'Alpha', 'alpha body', 'Alpha one', 'one body', 'beta body']);
    check('still collapsed and selected', [rows().map(textOf), selectedTexts()], [['Intro', 'Beta', 'Alpha'], ['Alpha']]);
    await click(control('word-outline-move-up'));
    // (Beta's body text is under Alpha now, in its subtext: it goes with it.)
    check('and up again, past Beta, with what is under it now', docTexts().map(t => t[0]), ['Intro', 'Alpha', 'alpha body', 'Alpha one', 'one body', 'beta body', 'Beta']);
    Word.undo(editor); await settle();
    Word.undo(editor); await settle();
    key('+', { altKey: true, shiftKey: true, code: 'Equal' }); await settle();
    check('Alt+Shift+Plus expands', rows().length, 7);
    const symbol = rows()[3].querySelector('[data-testid="word-outline-symbol"]');
    symbol.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })); await settle();
    check('the plus of a heading collapses it', rows().map(textOf), ['Intro', 'Alpha', 'alpha body', 'Alpha one', 'Beta', 'beta body']);
    rows()[3].querySelector('[data-testid="word-outline-symbol"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })); await settle();
    check('and expands it', rows().length, 7);

    // Show level, first line only, text formatting.
    const show = control('word-outline-show-level'); show.value = '1'; show.dispatchEvent(new Event('change', { bubbles: true })); await settle();
    check('Show level 1: the Heading 1 lines', [rows().map(textOf), Word.getOutlineShowLevel(editor)], [['Alpha', 'Beta'], 1]);
    down(rows()[0]); await settle(); await click(control('word-outline-expand'));
    check('Expand under a level shown: one level more', rows().map(textOf), ['Alpha', 'alpha body', 'Alpha one', 'Beta']);
    key('3', { altKey: true, shiftKey: true, code: 'Digit3' }); await settle();
    check('Alt+Shift+3: headings to level 3', [rows().map(textOf), show.value], [['Alpha', 'Alpha one', 'Beta'], '3']);
    key('A', { altKey: true, shiftKey: true, code: 'KeyA' }); await settle();
    check('Alt+Shift+A: everything', rows().length, 7);
    const first = control('word-outline-first-line'); first.click(); await settle();
    check('Show first line only', lines.classList.contains('rb-word-outline-first-line'), true);
    first.click();
    const formatting = control('word-outline-formatting'); formatting.click(); await settle();
    check('Show text formatting off: plain text', rows()[1].querySelector('span span'), null);
    formatting.click(); await settle();

    // Back to print layout.
    rows()[4].dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true })); await settle();
    check('a double click: print layout, the caret on the line', [Word.getView(editor), view.style.display, Word.getSelection(editor).FocusBlockId, Word.getSelection(editor).FocusOffset], [3, 'none', blocks[4], 0]);
    check('the editor has the focus', document.activeElement === Word.getInputElement(editor), true);
    const input = Word.getInputElement(editor);
    caret(4);
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowLeft', altKey: true, shiftKey: true, bubbles: true, cancelable: true })); await settle();
    check('Alt+Shift+Left in print layout promotes the caret\'s paragraph (as Word)', docTexts()[4], ['one body', 2]);
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', altKey: true, shiftKey: true, bubbles: true, cancelable: true })); await settle();
    check('Alt+Shift+Up in print layout moves it up', docTexts().map(t => t[0]).slice(3, 5), ['one body', 'Alpha one']);
    Word.undo(editor); Word.undo(editor);
    await click(control('word-view-menu')); await click(control('word-menu-view-outline'));
    key('Enter'); await settle();
    check('Enter: print layout, the caret on the selected line', [Word.getView(editor), Word.getSelection(editor).FocusBlockId], [3, blocks[4]]);
    await click(control('word-view-menu')); await click(control('word-menu-view-outline'));
    await click(control('word-outline-close'));
    check('Close outline view', [Word.getView(editor), view.style.display], [3, 'none']);

    // A macro sets the view.
    await Word.enableMacros(editor);
    await Word.runMacroLine(editor, 'ActiveWindow.View.Type = wdOutlineView');
    await settle();
    check('a macro\'s View.Type = wdOutlineView shows outline view', [view.style.display, Word.getView(editor)], ['flex', 2]);
    await Word.runMacroLine(editor, 'ActiveWindow.View.ShowHeading 1');
    await settle();
    check('View.ShowHeading 1', [rows().map(textOf), control('word-outline-show-level').value], [['Alpha', 'Beta'], '1']);
    await Word.runMacroLine(editor, 'ActiveWindow.View.ShowAllHeadings');
    await Word.runMacroLine(editor, '?ActiveWindow.View.Type');
    await settle();
    check('ShowAllHeadings; View.Type read', [rows().length, Word.getMacroOutput(editor).at(-1).trim()], [7, '2']);
    await Word.runMacroLine(editor, 'ActiveWindow.View.Type = wdPrintView');
    await settle();
    check('View.Type = wdPrintView: back to the pages', view.style.display, 'none');
    check('the view is no change of the document', Word.getRevision(editor), revision);
} catch (e) { fail(e); }
finally { if (dispose) dispose(); finish(); }
