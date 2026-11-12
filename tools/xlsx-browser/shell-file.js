// Browser-side File menu checks: real package import/export, browser picker and download handoff.
import { checks, makeLog, mountGrid, bytesOf } from '/component.js';

const tick = () => new Promise((resolve) => setTimeout(resolve, 0));
const until = async (condition) => {
    for (let i = 0; i < 300; i++) { if (condition()) return; await new Promise((resolve) => setTimeout(resolve, 10)); }
    throw new Error('file operation did not settle');
};
const pick = (host, bytes, name) => {
    const input = host.querySelector('[data-sheet-file-input]');
    const transfer = new DataTransfer();
    transfer.items.add(new File([bytes], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    input.files = transfer.files;
    input.dispatchEvent(new Event('change', { bubbles: true }));
    return transfer.files[0];
};
const menu = (host, action) => {
    host.querySelector('[data-sheet-menu="file"]').click();
    host.querySelector(`[data-sheet-menu-item="${action}"]`).click();
};
const state = (host) => {
    const name = host.querySelector('.sheet-file-name');
    return [name.getAttribute('data-sheet-filename'), name.getAttribute('data-sheet-dirty')];
};
const mount = async () => {
    const opened = mountGrid();
    const host = document.createElement('div');
    document.body.appendChild(host);
    const { createSpreadsheetShell } = await import('/spreadsheet/Toolbar.js');
    const actions = [];
    const shell = createSpreadsheetShell(opened.grid, (action) => actions.push(action));
    shell.Mount(host);
    return { ...opened, host, gridHost: opened.host, shell, actions, closeAll: () => { shell.Dispose(); host.remove(); opened.close(); } };
};

checks['shell-file-session'] = async () => {
    const { lines, check } = makeLog();
    const session = await mount();
    const { grid, host, actions } = session;
    const confirm = window.confirm, anchorClick = HTMLAnchorElement.prototype.click;
    const handed = [];
    window.confirm = () => false;
    HTMLAnchorElement.prototype.click = function () {
        if (this.download) handed.push({ name: this.download, bytes: fetch(this.href).then((r) => r.arrayBuffer()).then((b) => new Uint8Array(b)) });
        else anchorClick.call(this);
    };
    try {
        const original = await bytesOf('tests/real-world/milk-street-office.xlsx');
        check('new workbook has a visible default filename and is clean', state(host), ['Workbook1.xlsx', 'false']);
        grid.SetCell(0, 0, 'keep');
        check('editing shows dirty state', [state(host), grid.IsModified()], [['Workbook1.xlsx', 'true'], true]);
        const warning = new Event('beforeunload', { cancelable: true });
        window.dispatchEvent(warning);
        check('dirty workbook warns on page unload', warning.defaultPrevented, true);
        pick(host, original, 'Quarterly Report.xlsx');
        await tick();
        check('declining replacement keeps workbook, name and dirty state', [grid.GetCell(0, 0).value, state(host)], ['keep', ['Workbook1.xlsx', 'true']]);
        const input = host.querySelector('[data-sheet-file-input]');
        input.value = '';
        input.dispatchEvent(new Event('change'));
        check('cancelled picker leaves workbook intact', [grid.GetCell(0, 0).value, state(host)], ['keep', ['Workbook1.xlsx', 'true']]);
        window.confirm = () => true;
        pick(host, new Uint8Array([1, 2, 3]), 'Broken.xlsx');
        await until(() => /Open failed:/.test(host.querySelector('.sheet-print-status').textContent));
        check('malformed package keeps workbook, undo, filename and dirty state', [grid.GetCell(0, 0).value, grid.CanUndo(), state(host)], ['keep', true, ['Workbook1.xlsx', 'true']]);
        pick(host, original, 'Quarterly Report.xlsx');
        await until(() => state(host)[0] === 'Quarterly Report.xlsx');
        check('real package opens with imported filename and clean state', [state(host), grid.GetSheetNames().length > 0, grid.IsModified()], [['Quarterly Report.xlsx', 'false'], true, false]);
        const cleanWarning = new Event('beforeunload', { cancelable: true });
        window.dispatchEvent(cleanWarning);
        check('clean workbook has no unload warning', cleanWarning.defaultPrevented, false);
        grid.SetCell(0, 0, 'first snapshot');
        const exportReal = grid.ExportXlsx.bind(grid);
        let release;
        grid.ExportXlsx = () => exportReal().then((bytes) => new Promise((resolve) => { release = () => resolve(bytes); }));
        menu(host, 'save');
        await until(() => !!release);
        menu(host, 'save');
        check('coalesced saves do not launch a second export', handed.length, 0);
        grid.SetCell(0, 0, 'later edit');
        release();
        await until(() => handed.length === 1);
        const savedBytes = await handed[0].bytes;
        const another = mountGrid();
        try {
            await another.grid.ImportXlsx(savedBytes);
            check('downloaded XLSX contains earlier captured snapshot, not later edit', [handed[0].name, another.grid.GetCell(0, 0).value, grid.GetCell(0, 0).value, grid.IsModified()], ['Quarterly Report.xlsx', 'first snapshot', 'later edit', true]);
        } finally { another.close(); }
        grid.ExportXlsx = exportReal;
        menu(host, 'save');
        await until(() => handed.length === 2);
        await handed[1].bytes;
        check('successful browser download handoff clears dirty state', state(host), ['Quarterly Report.xlsx', 'false']);
        check('file actions are owned by shell, never action dispatcher', actions, []);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    finally { window.confirm = confirm; HTMLAnchorElement.prototype.click = anchorClick; session.closeAll(); }
    return lines;
};

checks['shell-file-races'] = async () => {
    const { lines, check } = makeLog();
    const session = await mount();
    const { grid, host } = session;
    const originalRead = File.prototype.arrayBuffer, confirm = window.confirm, anchorClick = HTMLAnchorElement.prototype.click;
    const original = await bytesOf('tests/real-world/milk-street-office.xlsx');
    const delayed = [], held = new WeakSet();
    File.prototype.arrayBuffer = function () {
        if (this.name === 'Slow.xlsx' && !held.has(this)) {
            held.add(this);
            return new Promise((resolve, reject) => delayed.push(() => originalRead.call(this).then(resolve, reject)));
        }
        return originalRead.call(this);
    };
    window.confirm = () => true;
    let downloads = 0;
    HTMLAnchorElement.prototype.click = function () {
        if (this.download) downloads++;
        else anchorClick.call(this);
    };
    try {
        grid.SetCell(0, 0, 'original');
        pick(host, original, 'Slow.xlsx');
        await until(() => delayed.length > 0);
        grid.SetCell(0, 0, 'edited during parse');
        delayed.shift()();
        await until(() => /Open cancelled:/.test(host.querySelector('.sheet-print-status').textContent));
        check('edit while parsing vetoes open without losing edits or metadata', [grid.GetCell(0, 0).value, state(host)], ['edited during parse', ['Workbook1.xlsx', 'true']]);
        pick(host, original, 'Slow.xlsx');
        await until(() => delayed.length > 0);
        pick(host, original, 'Newer.xlsx');
        await until(() => state(host)[0] === 'Newer.xlsx');
        delayed.shift()();
        await tick();
        check('late older file cannot replace newer open', [state(host), grid.GetSheetNames().length > 0], [['Newer.xlsx', 'false'], true]);
        grid.SetCell(0, 0, 'saving old workbook');
        const exportReal = grid.ExportXlsx.bind(grid);
        let releaseSave;
        grid.ExportXlsx = () => exportReal().then((bytes) => new Promise((resolve) => { releaseSave = () => resolve(bytes); }));
        menu(host, 'save');
        await until(() => !!releaseSave);
        pick(host, original, 'Replacement.xlsx');
        await until(() => state(host)[0] === 'Replacement.xlsx');
        grid.SetCell(0, 0, 'new workbook edit');
        const newerStatus = host.querySelector('.sheet-print-status').textContent;
        releaseSave();
        await tick();
        check('old save cannot download or mark replaced workbook clean', [downloads, state(host), grid.GetCell(0, 0).value, host.querySelector('.sheet-print-status').textContent], [0, ['Replacement.xlsx', 'true'], 'new workbook edit', newerStatus]);
        grid.ExportXlsx = exportReal;
        pick(host, original, 'Slow.xlsx');
        await until(() => delayed.length > 0);
        const name = host.querySelector('.sheet-file-name');
        const status = host.querySelector('.sheet-print-status');
        const shownBeforeDisposal = name.textContent;
        const pending = status.textContent;
        session.shell.Dispose();
        delayed.shift()();
        await tick();
        check('disposed shell has no stale status or filename update', [name.textContent, status.textContent], [shownBeforeDisposal, pending]);
        const warning = new Event('beforeunload', { cancelable: true });
        grid.SetCell(0, 0, 'dirty after disposal');
        window.dispatchEvent(warning);
        check('dispose removes unload guard', warning.defaultPrevented, false);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    finally { File.prototype.arrayBuffer = originalRead; window.confirm = confirm; HTMLAnchorElement.prototype.click = anchorClick; session.closeAll(); }
    return lines;
};

checks['shell-file-shortcuts'] = async () => {
    const { lines, check } = makeLog();
    const first = await mount(), second = await mount();
    const outer = document.createElement('input');
    document.body.appendChild(outer);
    const createURL = URL.createObjectURL, inputClick = HTMLInputElement.prototype.click, anchorClick = HTMLAnchorElement.prototype.click;
    let pickers = 0, downloads = 0;
    HTMLInputElement.prototype.click = function () {
        if (this.hasAttribute('data-sheet-file-input')) pickers++;
        else inputClick.call(this);
    };
    HTMLAnchorElement.prototype.click = function () {
        if (this.download) downloads++;
        else anchorClick.call(this);
    };
    const key = (target, value, options = {}) => {
        const event = new KeyboardEvent('keydown', { key: value, ctrlKey: true, bubbles: true, cancelable: true, ...options });
        target.dispatchEvent(event);
        return event.defaultPrevented;
    };
    try {
        const gridTarget = first.gridHost.querySelector('.spreadsheet-container');
        const otherTarget = second.gridHost.querySelector('.spreadsheet-container');
        first.grid.SetCell(0, 0, 'not yet saved');
        URL.createObjectURL = () => { throw new Error('download denied'); };
        menu(first.host, 'save');
        await until(() => /Save failed:/.test(first.host.querySelector('.sheet-print-status').textContent));
        check('failed download is visible and leaves workbook dirty', [state(first.host), downloads], [['Workbook1.xlsx', 'true'], 0]);
        URL.createObjectURL = createURL;
        check('shortcut ignores another editor and composition', [key(outer, 's'), key(gridTarget, 's', { isComposing: true })], [false, false]);
        check('shortcut works in owned grid only', [key(gridTarget, 'o'), pickers], [true, 1]);
        check('second grid owns its own shortcut', [key(otherTarget, 'o'), pickers], [true, 2]);
        check('shell shortcut handles both platform modifiers', [key(first.host.querySelector('.sheet-file-name'), 'o', { ctrlKey: false, metaKey: true }), pickers], [true, 3]);
        check('save shortcut prevents browser save and invokes download', key(gridTarget, 's'), true);
        await until(() => downloads === 1);
        check('default filename downloaded and marked clean only after handoff', [downloads, state(first.host), state(second.host)], [1, ['Workbook1.xlsx', 'false'], ['Workbook1.xlsx', 'false']]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    finally {
        URL.createObjectURL = createURL;
        HTMLInputElement.prototype.click = inputClick;
        HTMLAnchorElement.prototype.click = anchorClick;
        outer.remove(); first.closeAll(); second.closeAll();
    }
    return lines;
};

checks['shell-file-edit-buffer'] = async () => {
    const { lines, check } = makeLog();
    const session = await mount();
    const { grid, host, gridHost } = session;
    const originalConfirm = window.confirm, anchorClick = HTMLAnchorElement.prototype.click;
    const handed = [];
    HTMLAnchorElement.prototype.click = function () {
        if (this.download) handed.push(fetch(this.href).then((r) => r.arrayBuffer()).then((buffer) => new Uint8Array(buffer)));
        else anchorClick.call(this);
    };
    try {
        const original = await bytesOf('tests/real-world/milk-street-office.xlsx');
        grid.StartEdit('draft');
        const input = gridHost.querySelector('.spreadsheet-edit-input');
        input.value = 'typed in active editor';
        input.dispatchEvent(new Event('input', { bubbles: true }));
        check('uncommitted entry is visibly unsaved without changing workbook revision',
            [grid.IsModified(), grid.IsEditing(), state(host)], [false, true, ['Workbook1.xlsx', 'true']]);
        const unload = new Event('beforeunload', { cancelable: true });
        window.dispatchEvent(unload);
        check('uncommitted entry warns on unload', unload.defaultPrevented, true);
        window.confirm = () => false;
        pick(host, original, 'Other.xlsx');
        await tick();
        check('declined open leaves edit buffer visible and uncommitted',
            [input.value, grid.IsEditing(), grid.IsModified(), state(host)], ['typed in active editor', true, false, ['Workbook1.xlsx', 'true']]);
        window.confirm = () => true;
        pick(host, new Uint8Array([1, 2, 3]), 'Broken.xlsx');
        await until(() => /Open failed:/.test(host.querySelector('.sheet-print-status').textContent));
        check('failed import leaves edit buffer and dirty state untouched',
            [input.value, grid.IsEditing(), grid.IsModified(), state(host)], ['typed in active editor', true, false, ['Workbook1.xlsx', 'true']]);
        menu(host, 'save');
        await until(() => handed.length === 1);
        const bytes = await handed[0];
        const reopened = mountGrid();
        try {
            await reopened.grid.ImportXlsx(bytes);
            check('downloaded real XLSX contains committed active editor text', reopened.grid.GetCell(0, 0).value, 'typed in active editor');
        } finally { reopened.close(); }
        check('save commits before capturing revision and clears dirty baseline',
            [grid.IsEditing(), grid.GetCell(0, 0).value, grid.IsModified(), state(host)],
            [false, 'typed in active editor', false, ['Workbook1.xlsx', 'false']]);
        grid.SetValidation('A2', { type: 'whole', operator: 'between', formula1: 1, formula2: 5 });
        grid.SelectCell(1, 0);
        grid.StartEdit('99');
        menu(host, 'save');
        check('validation rejection blocks save and preserves active input',
            [handed.length, grid.IsEditing(), input.value, state(host), host.querySelector('.sheet-print-status').textContent],
            [1, true, '99', ['Workbook1.xlsx', 'true'], 'Save blocked: correct the cell entry first.']);
        pick(host, original, 'Other.xlsx');
        await until(() => /Open blocked:/.test(host.querySelector('.sheet-print-status').textContent));
        check('validation rejection blocks replacement without losing input or metadata',
            [grid.IsEditing(), input.value, state(host), grid.GetCell(0, 0).value],
            [true, '99', ['Workbook1.xlsx', 'true'], 'typed in active editor']);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    finally { window.confirm = originalConfirm; HTMLAnchorElement.prototype.click = anchorClick; session.closeAll(); }
    return lines;
};
