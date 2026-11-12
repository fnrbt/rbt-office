// Exercise the shell's real File > Save path, including an in-flight edit,
// writer failure, and replacing or disposing the document during serialization.
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';

const pause = () => new Promise(resolve => setTimeout(resolve, 10));
const until = async (predicate) => {
    for (let i = 0; i < 1000; i++) { if (predicate()) return; await pause(); }
    throw new Error('Timed out waiting for shell save');
};
let dispose;
const originalGenerate = window.JSZip.prototype.generateAsync;
const originalCreate = URL.createObjectURL;
try {
    localStorage.setItem('rb.word.canvasMeasure', '1');
    const [Word, OF, Host, Shell, Dom] = await Promise.all([
        loadWord(), loadFormats(), import('/word/WordHost.js'), import('/word/WordShell.js'), import('/word/FableUI/Dom.js')
    ]);
    const root = document.createElement('div'); root.id = 'shell-test'; document.body.appendChild(root);
    const host = Host.create('', 'word-shell');
    dispose = Dom.mountApp('#shell-test', Shell.view(host));
    const editor = host.Editor;
    const originalBytes = await bytesOf('/repo/tests/test-documents-wild/docx4j-chart.docx');
    const pkg = await OF.processors.forEditorPackage(originalBytes);
    Word.importFromDocxPackage(editor, pkg);
    const line = Word.getLayoutLines(editor).find(l => (l.Text ?? '').length > 0);
    Word.setSelection(editor, line.BlockId, 0, line.BlockId, 0);
    Word.insertText(editor, 'SAVEDMARK ');
    Word.setSelection(editor, line.BlockId, 0, line.BlockId, 9);
    Word.toggleBold(editor);
    Word.setSelection(editor, line.BlockId, 9, line.BlockId, 9);
    const blobs = [];
    URL.createObjectURL = blob => { blobs.push(blob); return originalCreate.call(URL, blob); };
    const downloads = [];
    document.addEventListener('click', event => {
        if (event.target instanceof HTMLAnchorElement && event.target.download)
            downloads.push([event.target.download, event.target.href.startsWith('blob:')]);
    }, true);
    const save = async () => {
        document.querySelector('[data-testid="word-file-menu"]').click();
        await until(() => document.querySelector('[data-testid="word-menu-save"]'));
        document.querySelector('[data-testid="word-menu-save"]').click();
        await pause();
    };
    const gateWriter = () => {
        let release;
        const gate = new Promise(resolve => { release = resolve; });
        let calls = 0;
        window.JSZip.prototype.generateAsync = function (...args) {
            calls++;
            const result = originalGenerate.apply(this, args);
            return gate.then(() => result);
        };
        return { release, calls: () => calls };
    };
    let gate = gateWriter();
    const snapshotText = Word.getText(editor);
    await save();
    await until(() => gate.calls() === 1);
    await save(); // overlapping menu clicks must not start another writer
    Word.insertText(editor, 'LATERMARK');
    gate.release();
    await until(() => !host.Saving);
    check('one serialization for overlapping saves', gate.calls(), 1);
    check('download handoff and later edit stays dirty', [blobs.length, blobs[0] instanceof Blob, Word.isModified(editor)], [1, true, true]);
    Word.undo(editor);
    check('undo to the downloaded snapshot is clean', Word.isModified(editor), false);
    Word.redo(editor);
    check('redo of the later edit is dirty', Word.isModified(editor), true);
    const bytes = new Uint8Array(await blobs[0].arrayBuffer());
    check('browser received a temporary DOCX download link', [downloads, document.querySelectorAll('a[download=\"Document1.docx\"]').length], [[['Document1.docx', true]], 0]);
    const reopened = Word.initializeCanvas(document.createElement('canvas'));
    try {
        Word.importFromDocxPackage(reopened, await OF.processors.forEditorPackage(bytes));
        check('download reopens at captured revision', Word.getText(reopened), snapshotText);
        const savedLine = Word.getLayoutLines(reopened).find(l => (l.Text ?? '').includes('SAVEDMARK'));
        Word.setSelection(reopened, savedLine.BlockId, 0, savedLine.BlockId, 9);
        check('download preserves selected bold formatting', Word.getSelectionFormat(reopened).Bold, true);
        const zip = await window.JSZip.loadAsync(bytes);
        const originalZip = await window.JSZip.loadAsync(originalBytes);
        const chart = Object.keys(originalZip.files).find(name => /^word\/charts\/chart\d+\.xml$/.test(name));
        check('opaque chart part survives the edited download', !!chart && !!zip.file(chart) && await zip.file(chart).async('string') === await originalZip.file(chart).async('string'), true);
    } finally { Word.dispose(reopened); }
    window.JSZip.prototype.generateAsync = function () { return Promise.reject(new Error('intentional writer failure')); };
    await save();
    await until(() => !host.Saving);
    check('writer rejection remains dirty and names the failure', [Word.isModified(editor), host.Status.Value.includes('intentional writer failure'), blobs.length], [true, true, 1]);
    gate = gateWriter();
    await save();
    await until(() => gate.calls() === 1);
    Word.importFromDocxPackage(editor, pkg);
    Word.insertText(editor, 'NEW DOCUMENT EDIT');
    host.Status.Value = 'New document imported';
    gate.release();
    await until(() => !host.Saving);
    check('import during save does not mark or report on the new document', [Word.isModified(editor), host.Status.Value, blobs.length], [true, 'New document imported', 1]);
    gate = gateWriter();
    await save();
    await until(() => gate.calls() === 1);
    host.Status.Value = 'Disposed';
    dispose(); dispose = null;
    check('unmount releases the nested editor input and host ownership',
        [host.Editor == null, Word.getInputElement(editor)], [true, null]);
    gate.release();
    await until(() => !host.Saving);
    check('disposal during save does not report or download', [host.Status.Value, blobs.length], ['Disposed', 1]);
    const host2 = Host.create('', 'word-shell');
    const dispose2 = Dom.mountApp('#shell-test', Shell.view(host2));
    try {
        const editor2 = host2.Editor;
        Word.importFromDocxPackage(editor2, pkg);
        Word.insertText(editor2, 'DIRECT DISPOSAL EDIT');
        gate = gateWriter();
        await save();
        await until(() => gate.calls() === 1);
        host2.Status.Value = 'Editor disposed';
        Word.dispose(editor2);
        gate.release();
        await until(() => !host2.Saving);
        check('direct editor disposal during save does not report or download', [host2.Status.Value, blobs.length], ['Editor disposed', 1]);
    } finally { dispose2(); }
    window.JSZip.prototype.generateAsync = originalGenerate;
    const host3 = Host.create('Shortcut snapshot', 'word-shell');
    const dispose3 = Dom.mountApp('#shell-test', Shell.view(host3));
    try {
        const editor3 = host3.Editor;
        Word.insertText(editor3, ' keyboard edit');
        const expectedText = Word.getText(editor3);
        const input = Word.getInputElement(editor3);
        input.focus();
        const foreign = document.createElement('input');
        document.body.appendChild(foreign);
        try {
            const foreignKey = new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true, cancelable: true });
            foreign.dispatchEvent(foreignKey);
            const composingKey = new KeyboardEvent('keydown', { key: 's', ctrlKey: true, isComposing: true, bubbles: true, cancelable: true });
            input.dispatchEvent(composingKey);
            check('unrelated input and composition do not trigger shell save', [foreignKey.defaultPrevented, composingKey.defaultPrevented, host3.Saving, blobs.length], [false, false, false, 1]);
        } finally { foreign.remove(); }
        const shortcut = new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true, cancelable: true });
        input.dispatchEvent(shortcut);
        await until(() => blobs.length === 2 && !host3.Saving);
        check('editor input shortcut requests download and marks its snapshot saved', [shortcut.defaultPrevented, Word.isModified(editor3)], [true, false]);
        const shortcutEditor = Word.initializeCanvas(document.createElement('canvas'));
        try {
            const shortcutBytes = new Uint8Array(await blobs[1].arrayBuffer());
            Word.importFromDocxPackage(shortcutEditor, await OF.processors.forEditorPackage(shortcutBytes));
            check('shortcut download reopens with the edited snapshot', Word.getText(shortcutEditor), expectedText);
        } finally { Word.dispose(shortcutEditor); }
    } finally { dispose3(); }
} catch (e) { fail(e); }
finally {
    window.JSZip.prototype.generateAsync = originalGenerate;
    URL.createObjectURL = originalCreate;
    if (dispose) dispose();
    finish();
}
