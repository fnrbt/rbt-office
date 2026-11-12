// Exercise the browser file session with real DOCX packages and the shell picker.
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';

const pause = () => new Promise(resolve => setTimeout(resolve, 10));
const until = async predicate => {
    for (let i = 0; i < 1000; i++) { if (predicate()) return; await pause(); }
    throw new Error('Timed out waiting for shell file operation');
};
const originalRead = File.prototype.arrayBuffer;
const originalConfirm = window.confirm;
const originalCreate = URL.createObjectURL;
let dispose;
try {
    localStorage.setItem('rb.word.canvasMeasure', '1');
    const [Word, OF, Host, Shell, Dom] = await Promise.all([
        loadWord(), loadFormats(), import('/word/WordHost.js'), import('/word/WordShell.js'), import('/word/FableUI/Dom.js')
    ]);
    const root = document.createElement('div'); root.id = 'shell-files'; document.body.appendChild(root);
    const host = Host.create('Initial proof text', 'word-shell');
    dispose = Dom.mountApp('#shell-files', Shell.view(host));
    const editor = host.Editor;
    const picker = root.querySelector('[data-testid="word-document-file-input"]');
    const title = () => root.querySelector('[data-testid="word-document-title"]').textContent;
    const caret = () => {
        const line = Word.getLayoutLines(editor).find(item => (item.Text ?? '').length > 0);
        if (!line) throw new Error('No editable text in opened document');
        Word.setSelection(editor, line.BlockId, 0, line.BlockId, 0);
    };
    const unload = () => {
        const event = new Event('beforeunload', { cancelable: true });
        window.dispatchEvent(event);
        return event.defaultPrevented;
    };
    const bytesA = await bytesOf('/repo/tests/test-documents-wild/docx4j-chart.docx');
    const bytesB = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/tables.docx'));
    const file = (bytes, name) => new File([bytes], name, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const select = chosen => {
        const transfer = new DataTransfer(); transfer.items.add(chosen);
        picker.files = transfer.files;
        picker.dispatchEvent(new Event('change', { bubbles: true }));
    };
    const chooseMenu = async () => {
        root.querySelector('[data-testid="word-file-menu"]').click();
        await until(() => root.querySelector('[data-testid="word-menu-open"]'));
        root.querySelector('[data-testid="word-menu-open"]').click();
    };
    let pickerClicks = 0;
    picker.addEventListener('click', event => { pickerClicks++; event.preventDefault(); });
    let confirmations = 0, accept = true;
    window.confirm = () => { confirmations++; return accept; };
    check('initial proof text is a clean untitled session', [title(), Word.isModified(editor), unload()], ['Document1.docx', false, false]);
    await chooseMenu();
    check('File Open invokes the document picker', pickerClicks, 1);
    select(file(bytesA, 'chart.docx'));
    await until(() => host.FileName.Value === 'chart.docx');
    await until(() => title() === 'chart.docx');
    const originalText = Word.getText(editor);
    check('real DOCX becomes the clean named document', [title(), originalText.length > 0, unload()], ['chart.docx', true, false]);

    caret();
    Word.insertText(editor, 'UNSAVED FILE MARK');
    const editedText = Word.getText(editor);
    await until(() => title() === 'chart.docx *');
    check('edit is visibly dirty and protects unload', [title(), unload()], ['chart.docx *', true]);
    accept = false;
    select(file(bytesB, 'cancelled.docx'));
    await pause();
    check('declined replacement preserves text, filename and dirty state', [confirmations, Word.getText(editor), title(), unload()], [1, editedText, 'chart.docx *', true]);
    accept = true;
    select(file(new Uint8Array([1, 2, 3]), 'broken.docx'));
    await until(() => host.Status.Value.includes('DOCX open failed'));
    check('failed file read preserves the edited document', [Word.getText(editor), title(), unload()], [editedText, 'chart.docx *', true]);
    File.prototype.arrayBuffer = function () {
        return this.name === 'read-error.docx' ? Promise.reject(new Error('intentional read failure')) : originalRead.call(this);
    };
    select(file(bytesB, 'read-error.docx'));
    await until(() => host.Status.Value.includes('intentional read failure'));
    check('file read rejection preserves the edited document', [Word.getText(editor), title(), unload()], [editedText, 'chart.docx *', true]);
    File.prototype.arrayBuffer = originalRead;
    const notWord = new window.JSZip(); notWord.file('readme.txt', 'Not a Word document');
    select(file(await notWord.generateAsync({ type: 'uint8array' }), 'not-word.docx'));
    await until(() => host.Status.Value.includes('does not contain a Word document'));
    check('ZIP without a Word document cannot replace the session', [Word.getText(editor), title(), unload()], [editedText, 'chart.docx *', true]);

    select(file(bytesB, 'tables.docx'));
    await until(() => host.FileName.Value === 'tables.docx');
    await until(() => title() === 'tables.docx');
    const tablesText = Word.getText(editor);
    check('confirmed replacement installs and names the real file', [title(), Word.isModified(editor), unload()], ['tables.docx', false, false]);
    caret();
    Word.insertText(editor, 'SAVE NAME MARK');
    const blobs = [], downloads = [];
    URL.createObjectURL = blob => { blobs.push(blob); return originalCreate.call(URL, blob); };
    const captureDownload = event => {
        if (event.target instanceof HTMLAnchorElement && event.target.download) {
            downloads.push(event.target.download); event.preventDefault();
        }
    };
    document.addEventListener('click', captureDownload, true);
    root.querySelector('[data-testid="word-file-menu"]').click();
    await until(() => root.querySelector('[data-testid="word-menu-save"]'));
    root.querySelector('[data-testid="word-menu-save"]').click();
    await until(() => blobs.length === 1 && !host.Saving);
    await until(() => title() === 'tables.docx');
    check('Save downloads the opened filename and clears dirty state', [downloads, title(), unload()], [['tables.docx'], 'tables.docx', false]);
    const reopened = Word.initializeCanvas(document.createElement('canvas'));
    try {
        Word.importFromDocxPackage(reopened, await OF.processors.forEditorPackage(new Uint8Array(await blobs[0].arrayBuffer())));
        check('named download contains the edited document', Word.getText(reopened), Word.getText(editor));
    } finally { Word.dispose(reopened); }
    Word.insertText(editor, 'AFTER SAVE');
    await until(() => title() === 'tables.docx *');
    check('later edit is dirty', [title(), unload()], ['tables.docx *', true]);
    Word.undo(editor);
    await until(() => title() === 'tables.docx');
    check('undo to saved snapshot clears dirty state', [title(), unload()], ['tables.docx', false]);
    Word.redo(editor);
    await until(() => title() === 'tables.docx *');
    check('redo restores dirty state', [title(), unload()], ['tables.docx *', true]);

    let release, started = false;
    const gate = new Promise(resolve => { release = resolve; });
    File.prototype.arrayBuffer = function () {
        if (this.name === 'slow.docx') { started = true; return gate.then(() => originalRead.call(this)); }
        return originalRead.call(this);
    };
    select(file(bytesA, 'slow.docx'));
    await until(() => started);
    Word.insertText(editor, 'WHILE READING');
    const duringRead = Word.getText(editor);
    Word.undo(editor);
    Word.redo(editor);
    release();
    await until(() => host.Status.Value.includes('Open cancelled because'));
    check('edit during parsing cannot be overwritten', [Word.getText(editor), title()], [duringRead, 'tables.docx *']);

    let releaseOld, oldStarted = false;
    const oldGate = new Promise(resolve => { releaseOld = resolve; });
    File.prototype.arrayBuffer = function () {
        if (this.name === 'old.docx') { oldStarted = true; return oldGate.then(() => originalRead.call(this)); }
        return originalRead.call(this);
    };
    select(file(bytesA, 'old.docx'));
    await until(() => oldStarted);
    select(file(bytesB, 'new.docx'));
    await until(() => host.FileName.Value === 'new.docx');
    releaseOld();
    await new Promise(resolve => setTimeout(resolve, 300));
    check('newer open wins over an older pending open', [host.FileName.Value, Word.getText(editor), Word.isModified(editor)], ['new.docx', tablesText, false]);

    const input = Word.getInputElement(editor);
    const foreign = document.createElement('input'); document.body.appendChild(foreign);
    const foreignKey = new KeyboardEvent('keydown', { key: 'o', ctrlKey: true, bubbles: true, cancelable: true });
    const composition = new KeyboardEvent('keydown', { key: 'o', ctrlKey: true, isComposing: true, bubbles: true, cancelable: true });
    foreign.dispatchEvent(foreignKey); input.dispatchEvent(composition);
    const shortcut = new KeyboardEvent('keydown', { key: 'o', metaKey: true, bubbles: true, cancelable: true });
    input.dispatchEvent(shortcut); foreign.remove();
    check('open shortcut is scoped and ignores composition', [foreignKey.defaultPrevented, composition.defaultPrevented, shortcut.defaultPrevented, pickerClicks], [false, false, true, 2]);

    let releaseDisposed, disposedStarted = false;
    const disposedGate = new Promise(resolve => { releaseDisposed = resolve; });
    File.prototype.arrayBuffer = function () {
        if (this.name === 'disposed.docx') { disposedStarted = true; return disposedGate.then(() => originalRead.call(this)); }
        return originalRead.call(this);
    };
    caret();
    Word.insertText(editor, 'DIRTY BEFORE DISPOSE');
    select(file(bytesA, 'disposed.docx'));
    await until(() => disposedStarted);
    dispose(); dispose = null;
    host.Status.Value = 'Disposed';
    releaseDisposed(); await pause(); await pause();
    check('disposal invalidates pending open and removes unload protection', [host.FileName.Value, host.Status.Value, unload()], ['new.docx', 'Disposed', false]);
    document.removeEventListener('click', captureDownload, true);
} catch (e) { fail(e); }
finally {
    File.prototype.arrayBuffer = originalRead;
    window.confirm = originalConfirm;
    URL.createObjectURL = originalCreate;
    if (dispose) dispose();
    finish();
}
