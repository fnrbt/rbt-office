// The shell's navigation pane (View > Navigation pane): the headings as a tree indented by level, the
// heading the caret is under highlighted as the caret moves and as the document changes, a click
// going to a heading, expand / collapse, the keyboard, the context menu (promote, demote, move up /
// down, delete) and a heading dragged onto another - each one undo step.
import { check, finish, fail, loadWord } from './common.js';
let dispose;
try {
    const [Word, Host, Shell, Dom] = await Promise.all([
        loadWord(), import('/word/WordHost.js'), import('/word/WordShell.js'), import('/word/FableUI/Dom.js')
    ]);
    const root = document.createElement('div'); root.id = 'outline-pane'; root.style.height = '700px'; document.body.appendChild(root);
    const host = Host.create('', 'word-shell');
    dispose = Dom.mountApp('#outline-pane', Shell.view(host));
    const editor = host.Editor;
    const control = id => root.querySelector(`[data-testid="${id}"]`);
    const settle = () => new Promise(resolve => requestAnimationFrame(() => setTimeout(resolve, 0)));
    const later = () => new Promise(resolve => setTimeout(resolve, 450));
    const click = async (element) => { await settle(); element.click(); await settle(); };
    Word.setParagraphs(editor, ['Intro', 'Alpha', 'alpha body', 'Alpha one', 'one body', 'Alpha two', 'Beta', 'beta body']);
    const blocks = Word.getOutlineLines(editor, false).map(l => l.BlockId);
    const caret = (i, offset = 0) => Word.setSelection(editor, blocks[i], offset, blocks[i], offset);
    for (const [i, n] of [[1, 1], [3, 2], [5, 2], [6, 1]]) { caret(i); Word.setOutlineLevel(editor, n); }
    caret(0);
    check('the pane is closed at first', control('word-nav-pane').style.display, 'none');
    await click(control('word-view-menu'));
    check('View menu: Print layout checked, the pane not', [control('word-menu-view-print').getAttribute('aria-pressed'), control('word-menu-view-navigation').getAttribute('aria-pressed')], ['true', 'false']);
    await click(control('word-menu-view-navigation'));
    const items = () => [...root.querySelectorAll('[data-testid="word-nav-item"]')];
    const labels = () => items().map(i => i.textContent.replace(/[▸▾]/g, ''));
    check('opened: beside the pages', [control('word-nav-pane').style.display, control('word-nav-pane').getBoundingClientRect().right <= control('word-shell-host').getBoundingClientRect().left + 1], ['flex', true]);
    check('the headings, in order', labels(), ['Alpha', 'Alpha one', 'Alpha two', 'Beta']);
    check('indented by their place in the outline', items().map(i => i.getAttribute('aria-level')), ['1', '2', '2', '1']);
    check('a heading with subheadings can be collapsed', items().map(i => i.getAttribute('aria-expanded')), ['true', null, null, null]);
    const current = () => { const c = root.querySelector('.rb-word-nav-current'); return c && c.getAttribute('data-block-id'); };
    check('no heading before the caret: none highlighted', current(), null);
    caret(4); await settle();
    check('the caret in one body: Alpha one highlighted', current(), blocks[3]);
    caret(7); await settle();
    check('the caret moves: Beta', current(), blocks[6]);
    await click(items()[1]);
    check('a click goes to the heading', [Word.getSelection(editor).FocusBlockId, Word.getSelection(editor).FocusOffset, current()], [blocks[3], 0, blocks[3]]);

    // The document changes: typing into a heading, a new heading.
    caret(3, 9); Word.insertText(editor, ' (first)');
    await later();
    check('typing into a heading: its text in the pane once typing pauses', labels()[1], 'Alpha one (first)');
    caret(2); Word.setOutlineLevel(editor, 2); await later();
    check('a paragraph made a heading appears', labels(), ['Alpha', 'alpha body', 'Alpha one (first)', 'Alpha two', 'Beta']);
    Word.undo(editor); Word.undo(editor); await later();
    check('undone, as it was', labels(), ['Alpha', 'Alpha one', 'Alpha two', 'Beta']);

    // Expand / collapse.
    await click(items()[0].querySelector('[data-testid="word-nav-toggle"]'));
    check('collapsed: its subheadings are not listed', [labels(), items()[0].getAttribute('aria-expanded')], [['Alpha', 'Beta'], 'false']);
    caret(4); await settle();
    check('the caret under a collapsed heading: that heading highlighted', current(), blocks[1]);
    await click(items()[0].querySelector('[data-testid="word-nav-toggle"]'));
    check('expanded again', labels(), ['Alpha', 'Alpha one', 'Alpha two', 'Beta']);

    // The keyboard.
    const key = (target, k, extra = {}) => target.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...extra }));
    items()[0].focus(); key(items()[0], 'ArrowDown');
    check('Down: the next heading', document.activeElement && document.activeElement.getAttribute('data-block-id'), blocks[3]);
    key(document.activeElement, 'ArrowLeft');
    check('Left on a subheading: its heading', document.activeElement.getAttribute('data-block-id'), blocks[1]);
    key(document.activeElement, 'ArrowLeft'); await settle();
    check('Left on an expanded heading collapses it', labels(), ['Alpha', 'Beta']);
    key(document.activeElement, 'ArrowRight'); await settle();
    check('Right expands it', labels().length, 4);
    key(items()[3], 'Enter'); await settle();
    check('Enter goes to the heading', Word.getSelection(editor).FocusBlockId, blocks[6]);

    // The context menu.
    const menuFor = async (index) => {
        const item = items()[index]; const r = item.getBoundingClientRect();
        item.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: r.left + 10, clientY: r.top + 5 }));
        await settle();
        return control('word-nav-menu');
    };
    let menu = await menuFor(1);
    check('the context menu: its actions', [...menu.querySelectorAll('button')].map(b => [b.textContent, b.disabled]),
        [['Promote', false], ['Demote', false], ['Move up', true], ['Move down', false], ['Delete heading and content', false], ['Expand all', true], ['Collapse all', false]]);
    const revision = Word.getRevision(editor);
    await click(control('word-nav-menu-move-down'));
    check('Move down: Alpha one with its body past Alpha two', [labels(), Word.getOutlineLines(editor, false).map(l => l.Text).slice(1, 6)],
        [['Alpha', 'Alpha two', 'Alpha one', 'Beta'], ['Alpha', 'alpha body', 'Alpha two', 'Alpha one', 'one body']]);
    check('the menu closed', control('word-nav-menu'), null);
    Word.undo(editor); await later();
    check('one undo step', [labels(), Word.getRevision(editor)], [['Alpha', 'Alpha one', 'Alpha two', 'Beta'], revision]);
    await click((await menuFor(0)).querySelector('[data-testid="word-nav-menu-demote"]'));
    check('Demote: the heading and its subheadings', items().map(i => i.getAttribute('data-level')), ['2', '3', '3', '1']);
    await click((await menuFor(0)).querySelector('[data-testid="word-nav-menu-promote"]'));
    check('Promote: back', items().map(i => i.getAttribute('data-level')), ['1', '2', '2', '1']);
    await click((await menuFor(3)).querySelector('[data-testid="word-nav-menu-move-up"]'));
    check('Move up: Beta before Alpha and all under it', Word.getOutlineLines(editor, false).map(l => l.Text), ['Intro', 'Beta', 'beta body', 'Alpha', 'alpha body', 'Alpha one', 'one body', 'Alpha two']);
    Word.undo(editor); Word.undo(editor); Word.undo(editor); await later();
    check('undone, step by step', [labels(), Word.getRevision(editor)], [['Alpha', 'Alpha one', 'Alpha two', 'Beta'], revision]);
    menu = await menuFor(1);
    key(menu.querySelector('button'), 'Escape'); await settle();
    check('Escape closes the menu, focus back on the heading', [control('word-nav-menu'), document.activeElement.getAttribute('data-block-id')], [null, blocks[3]]);
    key(items()[1], 'F10', { shiftKey: true }); await settle();
    check('Shift+F10 opens it from the keyboard', !!control('word-nav-menu'), true);
    await click(control('word-nav-menu-delete'));
    check('Delete heading and content', [labels(), Word.getOutlineLines(editor, false).map(l => l.Text)], [['Alpha', 'Alpha two', 'Beta'], ['Intro', 'Alpha', 'alpha body', 'Alpha two', 'Beta', 'beta body']]);
    Word.undo(editor); await later();

    // Drag and drop.
    const drag = async (from, to, below) => {
        const transfer = new DataTransfer();
        const source = items()[from], target = items()[to];
        const r = target.getBoundingClientRect();
        const y = below ? r.bottom - 2 : r.top + 2;
        source.dispatchEvent(new DragEvent('dragstart', { bubbles: true, cancelable: true, dataTransfer: transfer }));
        const over = new DragEvent('dragover', { bubbles: true, cancelable: true, dataTransfer: transfer, clientX: r.left + 10, clientY: y });
        target.dispatchEvent(over);
        const accepted = over.defaultPrevented;
        const marked = target.className;
        target.dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: transfer, clientX: r.left + 10, clientY: y }));
        source.dispatchEvent(new DragEvent('dragend', { bubbles: true, dataTransfer: transfer }));
        await settle();
        return [accepted, marked.includes(below ? 'drop-after' : 'drop-before')];
    };
    check('Beta dragged above Alpha: the drop is shown and accepted', await drag(3, 0, false), [true, true]);
    check('moved with its subtext', Word.getOutlineLines(editor, false).map(l => l.Text), ['Intro', 'Beta', 'beta body', 'Alpha', 'alpha body', 'Alpha one', 'one body', 'Alpha two']);
    Word.undo(editor); await later();
    check('Alpha one dragged below Beta (after its subtext)', [await drag(1, 3, true), Word.getOutlineLines(editor, false).map(l => l.Text)],
        [[true, true], ['Intro', 'Alpha', 'alpha body', 'Alpha two', 'Beta', 'beta body', 'Alpha one', 'one body']]);
    Word.undo(editor); await later();
    check('a heading is not dropped into its own subtext', (await drag(0, 1, false))[0], false);
    check('nothing moved', Word.getRevision(editor), revision);

    await click(control('word-nav-close'));
    check('Close hides the pane', control('word-nav-pane').style.display, 'none');
} catch (e) { fail(e); }
finally { if (dispose) dispose(); finish(); }
