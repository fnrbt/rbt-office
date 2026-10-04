// Browser side of run.mjs for the Spreadsheet web component: change events,
// dispose, sheet access and sizing. Each check returns a list of
// "ok ..." / "FAIL ..." lines.
const main = await import('/office-formats/src/Main.js');
export const OF = typeof main.default === 'function' ? main.default() : main.default;
const { default: createSpreadsheet } = await import('/spreadsheet/Spreadsheet.js');
export { createSpreadsheet };

export const bytesOf = async (name) => new Uint8Array(await (await fetch('/fixture?name=' + encodeURIComponent(name))).arrayBuffer());
export const makeLog = () => {
    const lines = [];
    const check = (label, got, want) => {
        const g = JSON.stringify(got), w = JSON.stringify(want);
        lines.push(g === w ? 'ok   ' + label : `FAIL ${label}: got ${g} want ${w}`);
    };
    return { lines, check, note: (text) => lines.push('ok   (' + text + ')') };
};
export const container = (style) => {
    const div = document.createElement('div');
    div.setAttribute('style', style);
    document.body.appendChild(div);
    return div;
};

// Every listener registered on `document` while `run` executes and still registered afterwards.
const trackDocumentListeners = () => {
    const live = new Map();
    const add = document.addEventListener.bind(document), remove = document.removeEventListener.bind(document);
    document.addEventListener = (type, fn, options) => { live.set(fn, type); add(type, fn, options); };
    document.removeEventListener = (type, fn, options) => { live.delete(fn); remove(type, fn, options); };
    return { live, restore: () => { document.addEventListener = add; document.removeEventListener = remove; } };
};

export const checks = {};

// Item: OnChange / OnSelectionChange / Dispose / GetSheets / container height.
checks['component-host-api'] = async () => {
    const { lines, check } = makeLog();
    const tracker = trackDocumentListeners();
    try {
        const sized = container('position:absolute;left:0;top:0;width:900px;height:420px');
        const grid = createSpreadsheet(sized);
        grid.Mount();
        const gridDiv = sized.querySelector('.spreadsheet-container');
        check('the grid fills a container that has a height', gridDiv.getBoundingClientRect().height, 420);
        sized.style.height = '300px';
        await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        check('and follows it when it is resized', gridDiv.getBoundingClientRect().height, 300);
        const rowsAt300 = sized.querySelectorAll('.spreadsheet-content > *').length;
        sized.style.height = '700px';
        await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
        check('more rows are rendered in a taller container', sized.querySelectorAll('.spreadsheet-content > *').length > rowsAt300, true);
        const short = container('position:absolute;left:0;top:800px;width:900px;height:90px');
        const shortGrid = createSpreadsheet(short); shortGrid.Mount();
        check('a sensible minimum height', short.querySelector('.spreadsheet-container').getBoundingClientRect().height, 200);
        shortGrid.Dispose(); short.remove();
        const unsized = container('position:absolute;left:0;top:1100px;width:900px');
        const other = createSpreadsheet(unsized); other.Mount();
        check('a container without a height gets the fixed height', unsized.querySelector('.spreadsheet-container').getBoundingClientRect().height, 600);
        other.Dispose(); unsized.remove();

        const changes = [], selections = [];
        const offChange = grid.OnChange((change) => changes.push(change));
        const offSelection = grid.OnSelectionChange((change) => selections.push(change));
        const takeChanges = () => changes.splice(0).map((c) => `${c.Sheets.join(',')}${c.StructureChanged ? ' structure' : ''}`);
        const takeSelections = () => selections.splice(0).map((c) => `${c.SheetIndex}:${c.Selection}`);

        grid.SetCell(0, 0, '12'); grid.SetCell(1, 0, '=A1*2');
        check('cell edits are reported', takeChanges(), ['0', '0']);
        check('they are not selection changes', takeSelections(), []);
        grid.SelectCell(3, 2);
        check('a selection move is reported as one', [takeChanges(), takeSelections()], [[], ['0:C4']]);
        grid.SelectRange(1, 1, 4, 3);
        check('a range selection too', takeSelections(), ['0:B2:D5']);
        grid.Render();
        check('a render with nothing changed reports nothing', [takeChanges(), takeSelections()], [[], []]);
        grid.SelectCell(0, 0); takeSelections();
        grid.ApplyFormat('bold'); check('formatting is a change', takeChanges(), ['0']);
        grid.MergeCells(6, 0, 7, 2); check('a merge is a change', takeChanges(), ['0']);
        grid.InsertRow(0); check('inserting a row is a change', takeChanges(), ['0']);
        grid.Undo(); check('undo is a change', takeChanges(), ['0']);
        grid.Redo(); takeChanges();
        check('GetSheets returns the live models', [grid.GetSheets().length, grid.GetSheets()[0] === grid.GetModel(), grid.GetSheetNames(), grid.GetActiveSheetIndex()], [1, true, ['Sheet1'], 0]);

        // A second sheet: switching is a selection change, edits name the sheet.
        const first = grid.GetSheets()[0];
        const { createEmpty } = await import('/spreadsheet/Model.js');
        grid.InstallWorkbook([first, createEmpty(100, 10)], ['One', 'Two'], 0, undefined);
        check('installing a workbook is a structure change', takeChanges(), ['1 structure']);
        takeSelections();
        grid.SwitchSheet(1);
        check('switching sheets is a selection change, not a document change', [takeChanges(), takeSelections()], [[], ['1:A1']]);
        grid.SetCell(2, 2, 'on two');
        check('an edit on the second sheet names it', [takeChanges(), grid.GetActiveSheetIndex(), grid.GetSheetNames()], [['1'], 1, ['One', 'Two']]);
        const sheets = grid.GetSheets();
        check('GetSheets has both sheets with their latest edits', [sheets.length, sheets[1] === grid.GetModel(), sheets[0] === first], [2, true, true]);

        offChange();
        grid.SetCell(3, 3, 'unheard');
        check('no change event after unsubscribing', changes.length, 0);
        grid.SelectCell(5, 5);
        check('the other subscription still works', takeSelections(), ['1:F6']);
        offSelection();
        grid.SelectCell(6, 6);
        check('until it is unsubscribed too', selections.length, 0);

        // Dispose with a dialog open: nodes and document listeners all go.
        grid.OnChange((change) => changes.push(change));
        grid.ShowFindReplaceDialog();
        check('listeners are on the document while mounted', tracker.live.size >= 3, true);
        check('the dialog is in the page', !!document.getElementById('find-replace-dialog'), true);
        grid.Dispose();
        check('dispose removes every document listener', [...tracker.live.values()], []);
        check('dispose removes the grid and its dialogs', [sized.children.length, !!document.getElementById('find-replace-dialog'), document.querySelectorAll('.spreadsheet-container').length], [0, false, 0]);
        grid.SetCell(0, 0, 'after dispose');
        check('no events after dispose', changes.length, 0);
        document.dispatchEvent(new MouseEvent('mousemove', { clientX: 5, clientY: 5 }));
        document.dispatchEvent(new MouseEvent('mouseup'));
        sized.remove();
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    tracker.restore();
    return lines;
};

window.componentChecks = checks;
