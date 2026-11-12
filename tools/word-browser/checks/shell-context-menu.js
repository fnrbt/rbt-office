// Context targeting and asynchronous clipboard ownership are editing boundaries.
import {check, finish, fail, loadWord} from './common.js';
let dispose;
const clipboardDescriptor = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
try {
    const [Word, Host, Shell, Dom] = await Promise.all([loadWord(), import('/word/WordHost.js'), import('/word/WordShell.js'), import('/word/FableUI/Dom.js')]);
    const root = document.createElement('div'); root.id = 'context-shell'; document.body.append(root);
    const host = Host.create('', 'word-shell'); dispose = Dom.mountApp('#context-shell', Shell.view(host));
    const editor = host.Editor;
    const settle = () => new Promise(resolve => requestAnimationFrame(resolve));
    const input = () => Word.getInputElement(editor);
    const menu = () => root.querySelector('[data-testid="word-context-menu"]');
    const key = (name, extra = {}) => document.activeElement.dispatchEvent(new KeyboardEvent('keydown', {key: name, bubbles: true, cancelable: true, ...extra}));
    const open = async () => {input().focus(); key('F10', {shiftKey: true}); await settle();};
    const act = id => {const button = menu()?.querySelector(`[data-command="${id}"]`); if (!button || button.disabled) throw new Error(`Unavailable ${id}`); button.click();};
    const reset = async () => {Word.setParagraphs(editor, ['alpha beta', 'gamma delta']); await settle(); return Word.getLayoutLines(editor)[0].BlockId;};
    let id = await reset();
    Word.setSelection(editor, id, 0, id, 5);
    await open(); act('format.bold'); await settle();
    check('context formatting applies to the retained selection', Word.getSelectionFormat(editor).Bold, true);
    await open(); act('edit.undo'); await settle();
    check('context formatting is one undo step', Word.getSelectionFormat(editor).Bold, false);
    await open(); key('End');
    check('End focuses the last enabled context action', document.activeElement.dataset.command, 'insert.link');
    key('Escape');
    check('Escape dismisses and restores the editor input', [menu() === null, document.activeElement === input()], [true, true]);
    await open(); Word.setContent(editor, 'replacement'); await settle();
    check('document replacement dismisses a stale context', menu(), null);
    id = await reset(); Word.setSelection(editor, id, 0, id, 5);
    input().readOnly = true; await open();
    check('read-only context cannot undo or format', [...menu().querySelectorAll('button:not(:disabled)')].every(button => button.dataset.command === 'edit.copy'), true);
    key('Escape'); input().readOnly = false;
    input().dispatchEvent(new CompositionEvent('compositionstart', {bubbles: true, data: ''}));
    await open(); check('composition does not open the context menu', menu(), null);
    input().dispatchEvent(new CompositionEvent('compositionend', {bubbles: true, data: ''}));
    const external = document.createElement('input'); document.body.append(external); external.focus();
    const native = new MouseEvent('contextmenu', {bubbles: true, cancelable: true}); external.dispatchEvent(native);
    check('external input retains its native context menu', native.defaultPrevented, false); external.remove();
    // Clipboard promises are deliberately controlled at the browser boundary:
    // failed writes and delayed completion must never delete different text.
    let resolveWrite;
    Object.defineProperty(navigator, 'clipboard', {configurable: true, value: {
        write: () => new Promise(resolve => {resolveWrite = resolve;}),
        read: () => Promise.resolve([{types: ['text/plain'], getType: async () => new Blob(['pasted'])}])
    }});
    id = await reset(); Word.setSelection(editor, id, 0, id, 5);
    await open(); act('edit.cut');
    check('cut waits for a successful clipboard write', Word.getText(editor), 'alpha beta\ngamma delta');
    Word.setSelection(editor, id, 6, id, 10); resolveWrite(); await settle();
    check('a moved selection cancels the pending cut', Word.getText(editor), 'alpha beta\ngamma delta');
    await open(); act('edit.cut'); resolveWrite(); await settle();
    check('successful cut deletes only its selection', Word.getText(editor), 'alpha \ngamma delta');
    Word.undo(editor); await settle();
    check('cut is undoable', Word.getText(editor), 'alpha beta\ngamma delta');
    navigator.clipboard.write = () => Promise.reject(new Error('permission denied'));
    Word.setSelection(editor, id, 0, id, 5); await open(); act('edit.cut'); await settle();
    check('denied clipboard write leaves text intact', Word.getText(editor), 'alpha beta\ngamma delta');
    check('clipboard denial is visible', host.Status.Value.includes('permission denied'), true);
    let resolveRead;
    navigator.clipboard.read = () => new Promise(resolve => {resolveRead = resolve;});
    await open(); act('edit.paste'); Word.setContent(editor, 'new document');
    resolveRead([{types: ['text/plain'], getType: async () => new Blob(['wrong'])}]); await settle();
    check('pending paste cannot overwrite a replacement document', Word.getText(editor), 'new document');
    navigator.clipboard.read = () => Promise.resolve([{types: ['text/plain'], getType: async () => new Blob(['right'])}]);
    id = await reset(); Word.setSelection(editor, id, 0, id, 5);
    await open(); act('edit.paste'); await settle();
    check('paste replaces the retained selection', Word.getText(editor), 'right beta\ngamma delta');
    Word.undo(editor); await settle(); check('paste is one undo step', Word.getText(editor), 'alpha beta\ngamma delta');
    Word.setSelection(editor, id, 0, id, 5); Word.insertHyperlink(editor, 'https://example.com', '');
    Word.setSelection(editor, id, 2, id, 2); await open(); act('link.remove'); await settle();
    check('remove link keeps its text', [Word.getHyperlink(editor), Word.getText(editor)], [null, 'alpha beta\ngamma delta']);
    Word.undo(editor); check('link removal is undoable', Word.getHyperlink(editor).Target, 'https://example.com');
    await open(); dispose(); dispose = null;
    check('disposal removes the popup', document.querySelector('[data-testid="word-context-menu"]'), null);
} catch (error) { fail(error); }
finally {
    if (dispose) dispose();
    if (clipboardDescriptor) Object.defineProperty(navigator, 'clipboard', clipboardDescriptor); else delete navigator.clipboard;
    finish();
}
