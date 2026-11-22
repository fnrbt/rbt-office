// Browser side of run.mjs, fifth part: the grid's own chrome that hosts build on (filter
// buttons, the outline gutter, drawings over frozen panes, ...). Input goes through the
// browser's real mouse and keyboard (window.realMouse / window.realKey, see run.mjs).
// Each check returns a list of "ok ..." / "FAIL ..." lines.
import { checks, makeLog, mountGrid, frame, parts, container, createSpreadsheet, everyObjectSeen as seenAll } from '/component.js';
import { reopened, thrown } from '/sheet-api.js';

const near = (a, b, tolerance = 1.01) => Math.abs(a - b) <= tolerance;
const centreOf = (rect) => [rect.left + rect.width / 2, rect.top + rect.height / 2];
export const realClick = async (x, y, options = {}) => { await window.realMouse('click', x, y, options); await frame(); };
// The box of a cell on screen, from the DOM node the grid drew for it.
export const cellBox = (host, row, col) => {
    const node = host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
    return node ? node.getBoundingClientRect() : null;
};
const inside = (rect, box) => !!box && rect.left >= box.left - 0.5 && rect.top >= box.top - 0.5 && rect.left + rect.width <= box.right + 0.5 && rect.top + rect.height <= box.bottom + 0.5;
// The element the browser would hit at the middle of a button's rectangle.
const hitAt = (rect) => { const [x, y] = centreOf(rect); const node = document.elementFromPoint(x, y); return node ? node.className : null; };

// AutoFilter buttons: OnFilterButton / GetFilterButtonRects, drawn in the header cells.
checks['chrome-filter-buttons'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const rows = [['Region', 'Units', 'Price', 'Note'], ['East', 10, 2, 'a'], ['West', 30, 4, 'b'], ['East', 20, 6, 'c'], ['North', 40, 8, 'd']];
        rows.forEach((row, r) => row.forEach((value, c) => grid.SetCell(r, c, String(value))));
        check('no AutoFilter: no buttons', [grid.GetFilterButtonRects().length, host.querySelectorAll('.spreadsheet-filter-button').length], [0, 0]);
        grid.SetAutoFilter('A1:C5');
        await frame();
        let rects = grid.GetFilterButtonRects();
        check('one button per header cell of the filter', rects.map((r) => [r.column, r.row, r.sheetIndex, r.active, r.sorted]), [[0, 0, 0, false, null], [1, 0, 0, false, null], [2, 0, 0, false, null]]);
        check('each inside its header cell, at the right edge', rects.map((r) => { const box = cellBox(host, 0, r.column); return [inside(r.clientRect, box), near(r.clientRect.left + r.clientRect.width, box.right - 1), r.clientRect.width, r.clientRect.height]; }), [0, 1, 2].map(() => [true, true, 16, 16]));
        check('the browser hits the button there', rects.map((r) => hitAt(r.clientRect)), [0, 1, 2].map(() => 'spreadsheet-filter-button'));
        check('the nodes say what they are', [...host.querySelectorAll('.spreadsheet-filter-button')].map((n) => [n.getAttribute('data-row'), n.getAttribute('data-col'), n.getAttribute('data-active'), n.getAttribute('data-sorted')]), [['0', '0', '0', ''], ['0', '1', '0', ''], ['0', '2', '0', '']]);

        // ---- no subscriber: the grid's own checklist ----
        grid.SelectCell(3, 3);
        await realClick(...centreOf(rects[1].clientRect));
        let dropdown = document.getElementById('filter-dropdown');
        check('a click opens the built-in checklist; the selection stays', [!!dropdown, dropdown ? [...dropdown.querySelectorAll('label span')].map((n) => n.textContent) : [], grid.GetSelectionRef()], [true, ['10', '20', '30', '40'], 'D4']);
        if (dropdown) {
            [...dropdown.querySelectorAll('input[type=checkbox]')].forEach((box, i) => { box.checked = i < 2; });
            [...dropdown.querySelectorAll('button')].find((b) => b.textContent === 'Apply').click();
            await frame();
        }
        rects = grid.GetFilterButtonRects();
        check('applied: the column is filtered and its button says so', [grid.GetFilters().hiddenRows, rects.map((r) => r.active), host.querySelector('.spreadsheet-filter-button[data-col="1"]').getAttribute('data-active')], [[2, 4], [false, true, false], '1']);
        grid.Undo(); await frame();
        check('undo', [grid.GetFilters().hiddenRows, grid.GetFilterButtonRects().map((r) => r.active)], [[], [false, false, false]]);
        grid.Redo(); await frame();
        check('redo', [grid.GetFilters().hiddenRows, grid.GetFilterButtonRects().map((r) => r.active)], [[2, 4], [false, true, false]]);

        // ---- a subscriber: the event instead of the checklist ----
        const events = [];
        const shownLegacy = () => [...host.querySelectorAll('.col-filter-btn')].filter((n) => getComputedStyle(n).display !== 'none').length;
        check('the legacy buttons beside the column letters are there without a subscriber', shownLegacy() > 0, true);
        const unsubscribe = grid.OnFilterButton((event) => events.push(event));
        await frame();
        check('... and go with one', shownLegacy(), 0);
        rects = grid.GetFilterButtonRects();
        await realClick(...centreOf(rects[1].clientRect));
        check('a click tells the subscriber, with the rectangle GetFilterButtonRects gives', [events.length, events[0] && [events[0].column, events[0].row, events[0].sheetIndex, events[0].active, events[0].sorted], events[0] && JSON.stringify(events[0].clientRect) === JSON.stringify(rects[1].clientRect)], [1, [1, 0, 0, true, null], true]);
        check('no checklist opens; the selection stays', [document.getElementById('filter-dropdown') == null, grid.GetSelectionRef()], [true, 'D4']);
        // Alt+Down on a header cell.
        events.length = 0;
        await realClick(...centreOf(cellBox(host, 0, 2)).map((v, i) => v - (i === 0 ? 20 : 0)));
        check('(the header cell is selected by a click beside its button)', [grid.GetSelectionRef(), events.length], ['C1', 0]);
        await window.realKey('Alt+ArrowDown'); await frame();
        check('Alt+Down on a header cell raises the same event', [events.length, events[0] && [events[0].column, events[0].row, events[0].active], grid.GetSelectionRef()], [1, [2, 0, false], 'C1']);
        events.length = 0;
        grid.SelectCell(2, 2); parts(host).content.focus();
        await window.realKey('Alt+ArrowDown'); await frame();
        check('Alt+Down elsewhere does not', events.length, 0);

        // ---- state: sorted ----
        grid.ClearFilter();
        grid.SortRange('A1:C5', [{ column: 1, descending: true }], { hasHeader: true });
        await frame();
        check('a sorted column says which way', grid.GetFilterButtonRects().map((r) => r.sorted), [null, 'desc', null]);
        grid.SortRange('A1:C5', [{ column: 0 }], { hasHeader: true });
        await frame();
        check('... ascending', [grid.GetFilterButtonRects().map((r) => r.sorted), host.querySelector('.spreadsheet-filter-button[data-col="0"]').getAttribute('data-sorted')], [['asc', null, null], 'asc']);

        // ---- geometry: column resize, hidden columns, zoom, scroll, frozen panes ----
        const before = grid.GetFilterButtonRects();
        grid.SetColumnWidth(0, 180); await frame();
        rects = grid.GetFilterButtonRects();
        check('a wider column moves its button and the ones after it', rects.map((r, i) => Math.round(r.clientRect.left - before[i].clientRect.left)), [80, 80, 80]);
        check('... still inside their cells', rects.map((r) => [inside(r.clientRect, cellBox(host, 0, r.column)), hitAt(r.clientRect)]), [0, 1, 2].map(() => [true, 'spreadsheet-filter-button']));
        grid.HideColumns(1); await frame();
        rects = grid.GetFilterButtonRects();
        check('a hidden column has no button', [rects.map((r) => r.column), host.querySelectorAll('.spreadsheet-filter-button').length, rects.map((r) => inside(r.clientRect, cellBox(host, 0, r.column)))], [[0, 2], 2, [true, true]]);
        grid.UnhideColumns(1); grid.SetColumnWidth(0, 100);
        grid.SetZoom(1.5); await frame();
        rects = grid.GetFilterButtonRects();
        check('zoomed: the rectangles are screen pixels', rects.map((r) => [r.clientRect.width, inside(r.clientRect, cellBox(host, 0, r.column)), hitAt(r.clientRect)]), [0, 1, 2].map(() => [24, true, 'spreadsheet-filter-button']));
        events.length = 0;
        await realClick(...centreOf(rects[2].clientRect));
        check('... and a click there is that button', events.map((e) => e.column), [2]);
        grid.SetZoom(1); await frame();
        const { scroller } = parts(host);
        for (let r = 5; r < 80; r++) grid.SetCell(r, 0, 'x' + r);
        for (let c = 4; c < 40; c++) grid.SetCell(0, c, 'h' + c);
        await frame();
        scroller.scrollTop = 300; await frame(); await frame();
        check('the header row scrolled away: no visible button', grid.GetFilterButtonRects().length, 0);
        scroller.scrollTop = 0; scroller.scrollLeft = 150; await frame(); await frame();
        rects = grid.GetFilterButtonRects();
        check('scrolled sideways: the buttons that are still in view', [rects.map((r) => r.column), rects.map((r) => hitAt(r.clientRect))], [[1, 2], ['spreadsheet-filter-button', 'spreadsheet-filter-button']]);
        scroller.scrollLeft = 0; await frame();
        grid.FreezeAt(1, 1); await frame();
        const still = grid.GetFilterButtonRects();
        scroller.scrollTop = 300; await frame(); await frame();
        rects = grid.GetFilterButtonRects();
        check('a frozen header row keeps its buttons where they were', [rects.map((r) => r.column), rects.map((r, i) => near(r.clientRect.top, still[i].clientRect.top) && near(r.clientRect.left, still[i].clientRect.left)), rects.map((r) => hitAt(r.clientRect))], [[0, 1, 2], [true, true, true], [0, 1, 2].map(() => 'spreadsheet-filter-button')]);
        scroller.scrollLeft = 150; await frame(); await frame();
        rects = grid.GetFilterButtonRects();
        check('scrolled sideways under a frozen column: its button stays, one goes under it, one moves', [rects.map((r) => r.column), near(rects[0].clientRect.left, still[0].clientRect.left), rects[1] && Math.round(still[2].clientRect.left - rects[1].clientRect.left), rects.map((r) => hitAt(r.clientRect))], [[0, 2], true, 150, ['spreadsheet-filter-button', 'spreadsheet-filter-button']]);
        events.length = 0;
        await realClick(...centreOf(rects[1].clientRect));
        check('... and is clicked there', events.map((e) => e.column), [2]);
        scroller.scrollLeft = 0; scroller.scrollTop = 0; grid.Unfreeze(); await frame();

        // ---- unsubscribe, a viewer ----
        unsubscribe(); await frame();
        check('unsubscribed: the legacy buttons are back', shownLegacy() > 0, true);
        rects = grid.GetFilterButtonRects();
        await realClick(...centreOf(rects[0].clientRect));
        dropdown = document.getElementById('filter-dropdown');
        check('... and a click opens the checklist again', !!dropdown, true);
        if (dropdown) { [...dropdown.querySelectorAll('button')].find((b) => b.textContent === 'Cancel').click(); await frame(); }
        grid.SetReadOnly(true); await frame();
        check('a viewer has no buttons', [grid.GetFilterButtonRects().length, host.querySelectorAll('.spreadsheet-filter-button').length], [0, 0]);
        const off = grid.OnFilterButton((event) => events.push(event)); await frame();
        events.length = 0;
        rects = grid.GetFilterButtonRects();
        if (rects[1]) await realClick(...centreOf(rects[1].clientRect));
        check('... unless a host listens (it decides what to offer)', [rects.length, events.map((e) => e.column)], [3, [1]]);
        off(); grid.SetReadOnly(false); await frame();

        // ---- a table's header row ----
        grid.ClearAutoFilter(); await frame();
        check('the filter removed: no buttons', grid.GetFilterButtonRects().length, 0);
        grid.SetCell(8, 2, 'Item'); grid.SetCell(8, 3, 'Qty'); grid.SetCell(9, 2, 'nut'); grid.SetCell(9, 3, '4'); grid.SetCell(10, 2, 'bolt'); grid.SetCell(10, 3, '6');
        grid.DefineTable('C9:D11', 'Parts');
        await frame();
        rects = grid.GetFilterButtonRects();
        check('a table\'s header row has them too', [rects.map((r) => [r.column, r.row, r.active]), rects.map((r) => inside(r.clientRect, cellBox(host, 8, r.column)))], [[[2, 8, false], [3, 8, false]], [true, true]]);
        grid.SetAutoFilter('C9:D11'); grid.SetFilter(3, { values: ['6'] }); await frame();
        rects = grid.GetFilterButtonRects();
        check('filtering the table lights its button (one button per cell)', rects.map((r) => [r.column, r.row, r.active]), [[2, 8, false], [3, 8, true]]);

        // ---- export, re-import ----
        const again = await reopened(grid);
        try {
            await frame();
            check('opened again: the buttons and their state', [again.grid.GetFilters().range, again.grid.GetFilterButtonRects().map((r) => [r.column, r.row, r.active])], ['C9:D11', [[2, 8, false], [3, 8, true]]]);
        } finally { again.close(); }
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { document.getElementById('filter-dropdown')?.remove(); close(); }
    return lines;
};

// The outline strips: brackets, +/- and level buttons for grouped rows / columns.
checks['chrome-outline'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        for (let r = 0; r < 60; r++) for (let c = 0; c < 6; c++) grid.SetCell(r, c, String(r * 10 + c));
        await frame();
        // (From the inside of the grid's border.)
        const inner = () => { const scroller = parts(host).scroller, box = scroller.getBoundingClientRect(); return [box.left + scroller.clientLeft, box.top + scroller.clientTop]; };
        const origin = () => { const box = cellBox(host, 0, 0), [left, top] = inner(); return [Math.round(box.left - left), Math.round(box.top - top)]; };
        const nodes = (kind) => [...host.querySelectorAll('.spreadsheet-outline-' + kind)];
        const buttons = (axis) => nodes('button').filter((n) => n.getAttribute('data-axis') === axis).map((n) => [Number(n.getAttribute('data-first')), Number(n.getAttribute('data-last')), n.textContent]);
        const buttonOf = (axis, first) => nodes('button').find((n) => n.getAttribute('data-axis') === axis && n.getAttribute('data-first') === String(first));
        const levelOf = (axis, level) => nodes('level').find((n) => n.getAttribute('data-axis') === axis && n.getAttribute('data-level') === String(level));
        const click = async (node) => { const box = node.getBoundingClientRect(); await realClick(box.left + box.width / 2, box.top + box.height / 2); };
        const rowOfButton = (node) => { const box = node.getBoundingClientRect(), y = box.top + box.height / 2; return [...host.querySelectorAll('.spreadsheet-row-header')].filter((h) => { const b = h.getBoundingClientRect(); return b.height > 0 && y >= b.top && y <= b.bottom; }).map((h) => Number(h.getAttribute('data-row')))[0]; };
        const colOfButton = (node) => { const box = node.getBoundingClientRect(), x = box.left + box.width / 2; return [...host.querySelectorAll('.spreadsheet-col-header')].filter((h) => { const b = h.getBoundingClientRect(); return b.width > 0 && x >= b.left && x <= b.right; }).map((h) => Number(h.getAttribute('data-col')))[0]; };
        check('no groups: no strip, the cells begin after the headers', [host.querySelectorAll('.spreadsheet-outline').length, origin(), grid.GetShowOutline()], [0, [50, 30], true]);

        grid.GroupRows('3:8'); grid.GroupRows('4:5');
        await frame();
        check('grouped rows: a strip of three steps left of the row numbers', [origin(), buttons('rows'), nodes('level').map((n) => n.textContent)], [[112, 30], [[2, 7, '−'], [3, 4, '−']], ['1', '2', '3']]);
        check('the +/- buttons are on the summary rows (the row below each group)', [rowOfButton(buttonOf('rows', 2)), rowOfButton(buttonOf('rows', 3))], [8, 5]);
        check('the row numbers moved right with the cells', Math.round(host.querySelector('.spreadsheet-row-header[data-row="0"]').getBoundingClientRect().left - inner()[0]), 62);
        check('the brackets are drawn', nodes('line').length, 2);
        check('a dot on each grouped row but where a button is (Excel 2019)', nodes('dot').length, 5);
        // A cell is still where the pointer says it is.
        const b2 = cellBox(host, 1, 1);
        await realClick(b2.left + b2.width / 2, b2.top + b2.height / 2);
        check('a click on a cell selects that cell', grid.GetSelectionRef(), 'B2');
        await window.realType('77'); await window.realKey('Enter'); await frame();
        check('... and typing goes there', grid.GetCell(1, 1).value, 77);

        await click(buttonOf('rows', 3));
        check('a click on the inner button collapses its group', [grid.GetHiddenRows(), grid.GetOutline().rows.map((g) => g.collapsed), buttons('rows'), grid.GetSelectionRef()], [['4:5'], [false, true], [[2, 7, '−'], [3, 4, '+']], 'B3']);
        grid.Undo(); await frame();
        check('undo', [grid.GetHiddenRows(), buttons('rows').map((b) => b[2])], [[], ['−', '−']]);
        grid.Redo(); await frame();
        check('redo', [grid.GetHiddenRows(), buttons('rows').map((b) => b[2])], [['4:5'], ['−', '+']]);
        await click(buttonOf('rows', 3));
        check('a click on + expands it', [grid.GetHiddenRows(), buttons('rows').map((b) => b[2])], [[], ['−', '−']]);
        await click(levelOf('rows', 1));
        check('level 1: every group collapsed', [grid.GetHiddenRows(), grid.GetOutline().rows.map((g) => g.collapsed), buttons('rows')], [['3:8'], [true, true], [[2, 7, '+']]]);
        await click(levelOf('rows', 2));
        check('level 2: the outer group open', [grid.GetHiddenRows(), buttons('rows').map((b) => b[2])], [['4:5'], ['−', '+']]);
        await click(levelOf('rows', 3));
        check('level 3: all rows', grid.GetHiddenRows(), []);
        grid.ShowOutlineLevel('rows', 1);
        check('ShowOutlineLevel does the same', grid.GetHiddenRows(), ['3:8']);
        grid.Undo();
        check('... as one undo step', grid.GetHiddenRows(), []);

        // ---- columns ----
        grid.GroupColumns('C:D'); await frame();
        check('grouped columns: a strip above the column letters', [origin(), buttons('columns'), colOfButton(buttonOf('columns', 2)), nodes('level').filter((n) => n.getAttribute('data-axis') === 'columns').map((n) => n.textContent)], [[112, 73], [[2, 3, '−']], 4, ['1', '2']]);
        await click(buttonOf('columns', 2));
        check('a click collapses the columns', [grid.GetHiddenColumns(), buttons('columns')], [['C:D'], [[2, 3, '+']]]);
        await click(buttonOf('columns', 2));
        check('... and expands them', grid.GetHiddenColumns(), []);

        // ---- the summary rows above ----
        check('GetOutline says where the summaries are', [grid.GetOutline().summaryBelow, grid.GetOutline().summaryRight], [true, true]);
        grid.SetOutlineSummary(false, true); await frame();
        check('summary above: the buttons are on the row before each group', [grid.GetOutline().summaryBelow, rowOfButton(buttonOf('rows', 2)), rowOfButton(buttonOf('rows', 3))], [false, 1, 2]);
        await click(buttonOf('rows', 3));
        check('... and collapse it', [grid.GetHiddenRows(), grid.GetOutline().rows.map((g) => g.collapsed)], [['4:5'], [false, true]]);

        // ---- zoom, frozen panes, scrolling ----
        grid.SetZoom(1.5); await frame();
        check('zoomed: the strip scales', origin(), [168, 110]);
        await click(buttonOf('rows', 3));
        check('... and its buttons are where they are drawn', grid.GetHiddenRows(), []);
        grid.SetZoom(1); await frame();
        grid.FreezeAt(2, 0); await frame();
        const levelBox = levelOf('rows', 2).getBoundingClientRect();
        parts(host).scroller.scrollTop = 400; await frame(); await frame();
        const after = levelOf('rows', 2).getBoundingClientRect();
        check('scrolled: the level buttons stay; the groups scroll away under the headers', [near(after.top, levelBox.top), near(after.left, levelBox.left), document.elementFromPoint(after.left + 6, after.top + 6) === levelOf('rows', 2)], [true, true, true]);
        const visible = nodes('button').filter((n) => n.getAttribute('data-axis') === 'rows').map((n) => { const b = n.getBoundingClientRect(); return document.elementFromPoint(b.left + 6, b.top + 6) === n; });
        check('... no group button of a row that scrolled away can be hit', visible.includes(true), false);
        parts(host).scroller.scrollTop = 0; await frame(); grid.Unfreeze(); await frame();

        // ---- switching it off; a sheet without headers ----
        grid.SetShowOutline(false); await frame();
        check('SetShowOutline(false): no strip, the cells are back', [host.querySelectorAll('.spreadsheet-outline').length, origin(), grid.GetShowOutline(), grid.GetOutline().rowLevels], [0, [50, 30], false, 2]);
        grid.SetShowOutline(true); await frame();
        check('SetShowOutline(true)', [origin(), nodes('button').length], [[112, 73], 3]);
        // A sheet without headers keeps its strips: a narrow band beside each holds the level buttons of the other axis.
        grid.SetShowHeaders(false); await frame();
        const shownHeaders = () => [...host.querySelectorAll('.spreadsheet-row-header, .spreadsheet-col-header')].filter((n) => n.getClientRects().length > 0).length;
        const inHeaderArea = (node) => { const box = node.getBoundingClientRect(), [left, top] = inner(), [x, y] = origin(); return box.left >= left - 1 && box.top >= top - 1 && (box.right <= left + x + 1 || box.bottom <= top + y + 1); };
        check('no headers: the strips stay, the row numbers and column letters go', [nodes('button').length, nodes('level').length, origin(), shownHeaders(), nodes('button').concat(nodes('level')).every(inHeaderArea)], [3, 5, [82, 63], 0, true]);
        check('... the level buttons are in the corner bands, not over the cells', nodes('level').map((n) => { const box = n.getBoundingClientRect(), [left, top] = inner(), [x, y] = origin(); return box.right <= left + x + 1 && box.bottom <= top + y + 1; }).every(Boolean), true);
        await click(buttonOf('rows', 3));
        check('... and a click collapses', grid.GetHiddenRows(), ['4:5']);
        grid.Undo(); await frame();
        const a1 = cellBox(host, 0, 0);
        await realClick(a1.left + a1.width / 2, a1.top + a1.height / 2);
        check('... a click on a cell still selects that cell', grid.GetSelectionRef(), 'A1');
        grid.SetShowHeaders(true); await frame();
        check('headers again', [origin(), shownHeaders() > 0], [[112, 73], true]);
        // The sheet's own switch (the sheet view's showOutlineSymbols): saved with the sheet, one undo step.
        check('the symbols are shown', [grid.GetShowOutlineSymbols(), grid.GetOutline().showSymbols], [true, true]);
        grid.SetShowOutlineSymbols(false); await frame();
        check('SetShowOutlineSymbols(false): no strip, the groups stay', [host.querySelectorAll('.spreadsheet-outline').length, origin(), grid.GetShowOutlineSymbols(), grid.GetOutline().showSymbols, grid.GetOutline().rowLevels], [0, [50, 30], false, false, 2]);
        {
            const hiddenZip = await window.JSZip.loadAsync(await grid.ExportXlsx());
            check('... written on the sheet view', /<sheetView [^>]*showOutlineSymbols="0"/.test(await hiddenZip.files['xl/worksheets/sheet1.xml'].async('string')), true);
            const hidden = await reopened(grid);
            try { await frame(); check('... read back: grouped, without symbols', [hidden.grid.GetShowOutlineSymbols(), hidden.grid.GetOutline().rowLevels, hidden.host.querySelectorAll('.spreadsheet-outline').length], [false, 2, 0]); } finally { hidden.close(); }
        }
        grid.Undo(); await frame();
        check('... undone', [grid.GetShowOutlineSymbols(), origin()], [true, [112, 73]]);
        // A viewer: the buttons collapse and expand - how the sheet is looked at, not a change of the workbook.
        grid.SetReadOnly(true); await frame();
        {
            let told = 0;
            const off = grid.OnChange(() => { told++; });
            const revision = grid.GetRevision(), canUndo = grid.CanUndo();
            await click(buttonOf('rows', 3)); await frame();
            check('a viewer: a click collapses the group', [grid.GetHiddenRows(), buttons('rows').map((b) => b[2])], [['4:5'], ['−', '+']]);
            check('... without a new revision, a change event or an undo step', [grid.GetRevision() === revision, told, grid.CanUndo() === canUndo], [true, 0, true]);
            await click(levelOf('rows', 1)); await frame();
            check('... the level buttons work', grid.GetHiddenRows(), ['3:8']);
            await click(levelOf('rows', 3)); await frame();
            check('... and expand again without editing the workbook', [grid.GetHiddenRows(), grid.GetRevision() === revision, told, grid.CanUndo() === canUndo], [[], true, 0, true]);
            if (typeof off === 'function') off();
        }
        grid.SetReadOnly(false);

        // ---- export, re-import ----
        grid.SetGroupCollapsed('rows', 3, true); grid.SetGroupCollapsed('columns', 2, true);
        const bytes = await grid.ExportXlsx();
        const zip = await window.JSZip.loadAsync(bytes);
        const sheetXml = await zip.files['xl/worksheets/sheet1.xml'].async('string');
        check('written: outlinePr, and the collapsed mark on the row above', [/<sheetPr><outlinePr summaryBelow="0"\/><\/sheetPr>/.test(sheetXml), /<row r="3"[^>]*collapsed="1"/.test(sheetXml), /<row r="4"[^>]*hidden="1"[^>]*outlineLevel="2"|<row r="4"[^>]*outlineLevel="2"[^>]*hidden="1"/.test(sheetXml)], [true, true, true]);
        const again = mountGrid();
        try {
            await again.grid.ImportXlsx(bytes); await frame();
            check('opened again: the outline', JSON.stringify(again.grid.GetOutline()), JSON.stringify(grid.GetOutline()));
            check('... hidden rows and columns', [again.grid.GetHiddenRows(), again.grid.GetHiddenColumns()], [['4:5'], ['C:D']]);
            const reButton = [...again.host.querySelectorAll('.spreadsheet-outline-button')].find((n) => n.getAttribute('data-axis') === 'rows' && n.getAttribute('data-first') === '3');
            check('... its strip', [!!reButton, reButton && reButton.textContent], [true, '+']);
            if (reButton) { await click(reButton); check('... expands by a click', again.grid.GetHiddenRows(), []); }
            again.grid.Undo();
            const twice = await reopened(again.grid);
            try { check('written again from the file: the same outline', JSON.stringify(twice.grid.GetOutline()), JSON.stringify(grid.GetOutline())); } finally { twice.close(); }
        } finally { again.close(); }
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// One screen pixel as [r, g, b].
const pixelAt = async (x, y) => {
    const image = new Image();
    image.src = 'data:image/png;base64,' + await window.screenshotRegion(Math.round(x), Math.round(y), 1, 1);
    await image.decode();
    const canvas = document.createElement('canvas'); canvas.width = 1; canvas.height = 1;
    const context = canvas.getContext('2d'); context.drawImage(image, 0, 0);
    return [...context.getImageData(0, 0, 1, 1).data.slice(0, 3)];
};
const isRed = (pixel) => pixel[0] > 180 && pixel[1] < 60 && pixel[2] < 60;

// A drawing over frozen panes: the part over the frozen rows / columns stays, the rest scrolls and is cut at the frozen edge.
checks['chrome-frozen-drawings'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const canvas = document.createElement('canvas'); canvas.width = 250; canvas.height = 150;
        const context = canvas.getContext('2d'); context.fillStyle = '#e00000'; context.fillRect(0, 0, 250, 150);
        for (let r = 0; r < 80; r++) grid.SetCell(r, 0, 'r' + r);
        for (let c = 1; c < 40; c++) grid.SetCell(0, c, 'c' + c);
        // Over A1:C6 (columns 100 px, rows 25 px): x 0..250, y 0..150.
        const id = grid.InsertPicture(canvas.toDataURL('image/png'), 'A1');
        await frame();
        const { scroller } = parts(host);
        const box = scroller.getBoundingClientRect();
        const at = (x, y) => pixelAt(box.left + scroller.clientLeft + 50 + x, box.top + scroller.clientTop + 30 + y);   // sheet pixels of the view, after the headers
        const reds = async (points) => { const out = []; for (const [x, y] of points) out.push(isRed(await at(x, y))); return out; };
        // Points of the view: in the frozen corner, over the frozen rows, over the frozen column, in the scrolled part.
        const corner = [50, 25], overRows = [200, 25], overCols = [50, 120], body = [200, 120];
        check('no frozen panes: the picture is one node, drawn over its cells', [host.querySelectorAll('.spreadsheet-picture').length, host.querySelectorAll('.spreadsheet-drawing-part').length, await reds([corner, overRows, overCols, body])], [1, 0, [true, true, true, true]]);
        grid.FreezeAt(2, 1); await frame();
        check('frozen at B3: the picture is drawn over the frozen cells too (it was under them)', await reds([corner, overRows, overCols, body]), [true, true, true, true]);
        check('... as one part per pane it reaches', [...host.querySelectorAll('.spreadsheet-drawing-part')].map((n) => n.getAttribute('data-pane')).sort(), ['0', '1', '2', '3']);
        scroller.scrollTop = 60; await frame(); await frame();
        // The frozen rows (y < 50) keep their part; below them the picture moved up by 60: its bottom edge is at y 150 - 60 = 90.
        check('scrolled down by 60: the part over the frozen rows stays, the rest moved up under them', await reds([corner, overRows, [50, 70], [200, 70], [50, 100], [200, 100]]), [true, true, true, true, false, false]);
        scroller.scrollTop = 200; await frame(); await frame();
        check('scrolled past it: only the part over the frozen rows is left', await reds([corner, overRows, [50, 60], [200, 60], overCols, body]), [true, true, false, false, false, false]);
        scroller.scrollTop = 0; scroller.scrollLeft = 120; await frame(); await frame();
        // The frozen column (x < 100) keeps its part; right of it the picture moved left by 120: its right edge is at x 250 - 120 = 130.
        check('scrolled right by 120: the part over the frozen column stays, the rest is cut at the frozen edge', await reds([corner, overCols, [115, 25], [115, 120], [140, 25], [140, 120]]), [true, true, true, true, false, false]);
        scroller.scrollLeft = 300; await frame(); await frame();
        check('scrolled past it sideways', await reds([corner, overCols, [115, 25], [115, 120]]), [true, true, false, false]);
        scroller.scrollLeft = 0; await frame();

        // A picture beside the frozen panes goes under them.
        grid.MovePicture(id, 'C5'); await frame();   // x 200..450, y 100..250
        check('a picture in the scrolled part is one node', [host.querySelectorAll('.spreadsheet-drawing-part').length, host.querySelectorAll('.spreadsheet-picture').length, await reds([[300, 150]])], [0, 1, [true]]);
        scroller.scrollTop = 80; scroller.scrollLeft = 150; await frame(); await frame();
        // Now at x 50..300, y 20..170 of the view: cut at x 100 and y 50.
        check('scrolled: it goes under the frozen rows and the frozen column', await reds([[150, 100], [75, 100], [150, 35], [75, 35]]), [true, false, false, false]);
        scroller.scrollTop = 0; scroller.scrollLeft = 0; await frame();

        // The mouse: select and drag the picture from the scrolled part to over the frozen rows.
        grid.MovePicture(id, 'A1'); await frame();
        const start = [box.left + scroller.clientLeft + 50 + 50, box.top + scroller.clientTop + 30 + 25];
        await realClick(start[0], start[1]);
        const handles = [...host.querySelectorAll('.spreadsheet-picture-handle')].filter((n) => n.getClientRects().length > 0);
        const se = handles.find((n) => n.getAttribute('data-handle') === 'se').getBoundingClientRect();
        check('a click on the part in the frozen corner selects the picture; its frame is over the frozen panes', [grid.GetSelectedPicture(), handles.length, document.elementFromPoint(se.left + 4, se.top + 4)?.className], [id, 8, 'spreadsheet-picture-handle']);
        await window.realMouse('move', start[0], start[1]); await window.realMouse('down', 0, 0);
        await window.realMouse('move', start[0] + 60, start[1] + 40, { steps: 4 }); await frame();
        check('while it is dragged it is cut where it is', await reds([[80, 55], [280, 60], [40, 20]]), [true, true, false]);
        await window.realMouse('up', 0, 0); await frame();
        const moved = grid.GetPictures()[0];
        check('dragged by (60, 40)', [moved.x, moved.y, moved.width, moved.height], [60, 40, 250, 150]);
        grid.Undo(); await frame();
        check('undo', [grid.GetPictures()[0].x, grid.GetPictures()[0].y, await reds([corner, body])], [0, 0, [true, true]]);
        grid.Redo(); await frame();
        check('redo', [grid.GetPictures()[0].x, grid.GetPictures()[0].y], [60, 40]);
        grid.Undo();

        // Zoom, and the panes going away.
        grid.SetZoom(2); await frame();
        check('zoomed to 200 %: the parts scale with their panes', await reds([[100, 50], [400, 50], [100, 240], [400, 240]]), [true, true, true, true]);
        grid.SetZoom(1); grid.Unfreeze(); await frame();
        check('unfrozen: one node again', [host.querySelectorAll('.spreadsheet-drawing-part').length, host.querySelectorAll('.spreadsheet-picture').length], [0, 1]);

        // Export, re-import (the frozen pane and the picture).
        grid.FreezeAt(2, 1);
        const again = await reopened(grid);
        try {
            await frame();
            check('opened again: frozen, with the picture cut along the panes', [again.grid.GetFrozen(), again.grid.GetPictures().map((p) => [p.x, p.y, p.width, p.height]), [...again.host.querySelectorAll('.spreadsheet-drawing-part')].length], [{ rows: 2, cols: 1 }, [[0, 0, 250, 150]], 4]);
        } finally { again.close(); }

        // A chart's frame is cut the same way.
        const { bytesOf } = await import('/component.js');
        const charts = mountGrid();
        try {
            await charts.grid.ImportXlsx((await bytesOf('tests/real-world/chart-zoo.xlsx')).slice());
            const withChart = charts.grid.GetSheets().findIndex((sheet, index) => { charts.grid.SwitchSheet(index); return charts.grid.GetDrawings().some((d) => d.kind === 'chart'); });
            const chart = charts.grid.GetDrawings().find((d) => d.kind === 'chart');
            await frame();
            const before = charts.host.querySelectorAll('.spreadsheet-object-chart').length;
            charts.grid.FreezeAt(chart.row + 2, chart.col + 1); await frame();
            const chartParts = [...charts.host.querySelectorAll(`.spreadsheet-drawing-part[data-drawing-id="${chart.id}"]`)];
            check('a chart frame over frozen panes is cut into parts as well', [withChart >= 0, before > 0, chartParts.length > 1, chartParts.every((part) => part.querySelector('.spreadsheet-object-chart'))], [true, true, true, true]);
        } finally { charts.close(); }
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// Nothing of the opened package is lost on a save: every part the editor does not model goes back as it was.
checks['export-keeps-parts'] = async () => {
    const { lines, check } = makeLog();
    const { bytesOf } = await import('/component.js');
    const base64 = async (zip, name) => zip.files[name].async('base64');
    const entries = (zip) => Object.keys(zip.files).filter((name) => !zip.files[name].dir);
    // Parts the writer makes from the model (their text may differ); every other part is the file's, byte for byte.
    const rewritten = (name) => /^(\[Content_Types\]\.xml|_rels\/\.rels|xl\/workbook\.xml|xl\/_rels\/workbook\.xml\.rels|xl\/styles\.xml|xl\/sharedStrings\.xml|xl\/metadata\.xml|xl\/worksheets\/|xl\/comments|xl\/threadedComments\/|xl\/persons\/)/.test(name);
    for (const file of ['tests/real-world/dashboard-ecommerce-sales.xlsx', 'tests/test-documents-excel/pandas-test_decimal.xlsx', 'tests/test-documents-excel/poi-55864.xlsx', 'tests/test-documents-excel/phpsp-27template.xlsx', 'tests/real-world/milk-street-office.xlsx']) {
        const short = file.split('/').pop();
        const { grid, close } = mountGrid();
        try {
            const original = await bytesOf(file);
            const before = await window.JSZip.loadAsync(original);
            await grid.ImportXlsx(original.slice());
            const compare = async (label, bytes, skip = () => false) => {
                const after = await window.JSZip.loadAsync(bytes);
                const names = entries(before).filter((name) => name !== 'xl/calcChain.xml' && !skip(name));
                check(`${short}: ${label}: every part is there`, names.filter((name) => !after.files[name]), []);
                const changed = [];
                for (const name of names) if (!rewritten(name) && after.files[name] && await base64(before, name) !== await base64(after, name)) changed.push(name);
                check(`${short}: ${label}: what is not made from the model is byte for byte the file's`, changed, []);
                return after;
            };
            // Nothing changed at all: the bytes that were opened.
            const same = (a, b) => a.length === b.length && a.every((value, index) => value === b[index]);
            const untouched = await grid.ExportXlsx();
            check(`${short}: nothing changed: the export is the opened file, byte for byte`, [same(untouched, original), untouched.buffer !== original.buffer], [true, true]);
            // (From here on the writer writes: another zoom is not the file as it was opened.)
            grid.SetZoom(grid.GetZoom() === 0.9 ? 0.8 : 0.9);
            const zoomed = await grid.ExportXlsx();
            check(`${short}: another zoom: written anew`, same(zoomed, original), false);
            const unedited = await compare('unedited', zoomed);
            if (short === 'dashboard-ecommerce-sales.xlsx') {
                const rels = await unedited.files['xl/_rels/workbook.xml.rels'].async('string'), types = await unedited.files['[Content_Types].xml'].async('string');
                check('the data model and the metadata are related and typed', [/Target="(\/xl\/)?model\/item\.data"/.test(rels), /Target="(\/xl\/)?metadata"/.test(rels), /PartName="\/xl\/model\/item\.data"|Extension="data"/.test(types), /PartName="\/xl\/metadata"/.test(types)], [true, true, true, true]);
                const sheet3 = await unedited.files['xl/worksheets/sheet3.xml'].async('string'), sheet3Rels = await unedited.files['xl/worksheets/_rels/sheet3.xml.rels'].async('string');
                const legacy = /<legacyDrawing r:id="([^"]+)"/.exec(sheet3);
                check('sheet3 names its VML drawing', [!!legacy, legacy && new RegExp(`Id="${legacy[1]}"[^>]*vmlDrawing1\\.vml|vmlDrawing1\\.vml"[^>]*Id="${legacy[1]}"`).test(sheet3Rels)], [true, true]);
                check('the drawings keep their names', entries(unedited).filter((name) => /sheetDrawing/.test(name)), []);
            }
            // Only the sheet on display changed (its zoom): every other worksheet is the part that was read, byte for byte,
            // beside styles and shared strings that are the file's, added to.
            const sheetParts = entries(before).filter((name) => /^xl\/worksheets\/[^/]+\.xml$/.test(name));
            const sameSheets = async (zip) => { const same = []; for (const name of sheetParts) if (zip.files[name] && await base64(before, name) === await base64(zip, name)) same.push(name); return same; };
            const date1904 = /date1904="(1|true)"/.test(await before.files['xl/workbook.xml'].async('string'));
            if (sheetParts.length > 1 && !date1904) {
                // (The sheet the file had on display is not the one shown here: its part no longer says tabSelected, and nothing else.)
                let selectedParts = 0;
                for (const name of sheetParts) {
                    const was = await before.files[name].async('string');
                    if (!/<sheetView\b[^>]*\btabSelected="(1|true)"/.test(was)) continue;
                    const now = unedited.files[name] ? await unedited.files[name].async('string') : '';
                    if (now === was.replace(/(<sheetView\b[^>]*?)\s+tabSelected="(1|true)"/, '$1')) selectedParts++;
                }
                const expected = 1 + selectedParts;
                check(`${short}: another zoom: every other sheet is byte for byte the file's`, sheetParts.length - (await sameSheets(unedited)).length, expected);
                // ... and so with one cell edited on it, outside everything.
                const usedNow = grid.GetUsedRange();
                grid.SetCell((usedNow ? usedNow.lastRow : 0) + 3, (usedNow ? usedNow.lastCol : 0) + 3, 'one edit');
                const one = await compare('one cell edited', await grid.ExportXlsx());
                check(`${short}: one cell edited: every other sheet is byte for byte the file's`, sheetParts.length - (await sameSheets(one)).length, expected);
                // The edited sheet is the file's part with the new cell (and the zoom) in it: nothing else of it is written anew.
                const tally = async (zip) => {
                    const found = { shared: 0, styledRows: 0, rowHeights: 0, sheetPr: [], ignoredErrors: 0, headerFooter: 0, selections: 0, strings: ((await zip.files['xl/sharedStrings.xml']?.async('string')) ?? '').split('<si').length - 1 };
                    for (const name of sheetParts) {
                        const xml = await zip.files[name].async('string');
                        found.shared += (xml.match(/t="shared"/g) || []).length;
                        found.styledRows += (xml.match(/<row\b[^>]*\scustomFormat="1"/g) || []).length;
                        found.rowHeights += (xml.match(/<row\b[^>]*\sht="/g) || []).length;
                        found.sheetPr.push((/<sheetPr\b[^>]*\/>|<sheetPr\b[^>]*>[\s\S]*?<\/sheetPr>/.exec(xml) || [''])[0]);
                        found.ignoredErrors += (xml.match(/<ignoredError\b/g) || []).length;
                        found.headerFooter += (xml.match(/<headerFooter\b/g) || []).length;
                        found.selections += (xml.match(/<selection\b/g) || []).length;
                    }
                    return found;
                };
                check(`${short}: one cell edited: shared formulas, row formats and heights, sheet properties, ignored errors, selections and the shared strings are the file's`, await tally(one), { ...(await tally(before)), strings: (await tally(before)).strings + 1 });
                const plumbingChanged = []; for (const name of ['[Content_Types].xml', '_rels/.rels', 'xl/_rels/workbook.xml.rels']) if (await base64(before, name) !== await base64(one, name)) plumbingChanged.push(name);
                check(`${short}: one cell edited: content types and relationships are the file's`, plumbingChanged, []);
                const styles = await one.files['xl/styles.xml'].async('string'), stylesBefore = await before.files['xl/styles.xml'].async('string');
                // What the editor has no model of stays in the styles and the workbook part; untouched comments are the file's part.
                const element = (xml, name) => { const m = new RegExp(`<${name}(?=[\\s>/])[^>]*/>|<${name}(?=[\\s>])[^>]*>[\\s\\S]*?</${name}>`).exec(xml); return m ? m[0] : ''; };
                const count = (xml, needle) => xml.split(needle).length - 1;
                check(`${short}: one cell edited: table styles, palette and extensions of the styles stay`, ['tableStyles', 'colors', 'extLst', 'cellStyles'].filter((name) => element(styles, name) !== element(stylesBefore, name)), []);
                check(`${short}: one cell edited: protection, quote prefix and pivot buttons stay`, [count(styles, '<protection') >= count(stylesBefore, '<protection'), count(styles, 'quotePrefix="1"') >= count(stylesBefore, 'quotePrefix="1"'), count(styles, 'pivotButton="1"') >= count(stylesBefore, 'pivotButton="1"')], [true, true, true]);
                const bookNow = await one.files['xl/workbook.xml'].async('string'), bookBefore = await before.files['xl/workbook.xml'].async('string');
                const bookTags = (xml) => (xml.match(/<(fileVersion|workbookPr|workbookProtection|calcPr|workbookView)(?=[\s>\/])[^>]*>/g) || []).map((tag) => tag.replace(/\s(activeTab|firstSheet|fullCalcOnLoad)="[^"]*"/g, ''));
                check(`${short}: one cell edited: the workbook part keeps its properties, version, window and calculation settings`, bookTags(bookNow), bookTags(bookBefore));
                const commentParts = entries(before).filter((name) => /^xl\/comments[^/]*\.xml$/.test(name));
                const commentsChanged = []; for (const name of commentParts) if (!one.files[name] || await base64(before, name) !== await base64(one, name)) commentsChanged.push(name);
                check(`${short}: one cell edited: the comments parts are the file's`, commentsChanged, []);
                const xfs = (xml) => { const m = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml); return m ? m[1] : ''; };
                check(`${short}: one cell edited: the file's cell formats keep their places`, xfs(styles).startsWith(xfs(stylesBefore)), true);
                // The written file opens with what the edited sheet had, and the edit.
                const { grid: again, close: closeAgain } = mountGrid();
                try {
                    await again.ImportXlsx(await grid.ExportXlsx());
                    const shown = grid.GetActiveSheetIndex();
                    if (shown > 0) again.SwitchSheet(shown);
                    const ranges = [grid.GetUsedRange(), again.GetUsedRange()];
                    check(`${short}: one cell edited: read again, the used range`, ranges[1], ranges[0]);
                    const sample = (g) => { const u = g.GetUsedRange(); const out = []; if (u) for (let r = u.firstRow; r <= Math.min(u.lastRow, u.firstRow + 40); r++) for (let c = u.firstCol; c <= Math.min(u.lastCol, u.firstCol + 12); c++) { const cell = g.GetCell(r, c); out.push([cell.value, cell.text, cell.format && cell.format.bold, cell.format && cell.format.numberFormat, cell.format && cell.format.backgroundColor]); } return out; };
                    check(`${short}: one cell edited: read again, cells and formats`, sample(again), sample(grid));
                } finally { closeAgain(); }
            }
            // An edit of a cell on every sheet: still nothing lost (drawings and tables of the edited sheets included).
            const sheets = grid.GetSheets();
            for (let index = 0; index < sheets.length; index++) { if (sheets[index].visibility && sheets[index].visibility !== 'visible') continue; grid.SwitchSheet(index); const used = grid.GetUsedRange(); grid.SetCell((used ? used.lastRow : 0) + 3, 0, 'edited'); }
            await compare('a cell edited on each sheet', await grid.ExportXlsx());
        } catch (e) { lines.push(`FAIL ${short} threw: ` + String(e && e.stack || e)); }
        finally { close(); }
    }
    return lines;
};

// A sheet with more than cells changed keeps what its part says (ExportXlsx): a hidden row and another row
// height are put into the file's part; a link makes the rest of the part be written from the sheet, the rows
// still the file's; with a row inserted the sheet is written from its cells and still has the code name, the
// rows' styles, ignoredErrors (where their cells are now), the selection, and its formulas written once.
checks['export-edited-sheet-keeps-part'] = async () => {
    const { lines, check } = makeLog();
    const { bytesOf } = await import('/component.js');
    const file = 'tests/real-world/milk-street-office.xlsx', part = 'xl/worksheets/sheet1.xml';
    const element = (xml, name) => { const m = new RegExp(`<${name}(?=[\\s>/])[^>]*/>|<${name}(?=[\\s>])[^>]*>[\\s\\S]*?</${name}>`).exec(xml); return m ? m[0] : ''; };
    const count = (xml, pattern) => (xml.match(pattern) || []).length;
    const rowTag = (xml, number) => (new RegExp(`<row r="${number}"[^>]*>`).exec(xml) || [''])[0];
    const original = await bytesOf(file);
    const was = await (await window.JSZip.loadAsync(original)).files[part].async('string');
    const edited = async (edit) => {
        const { grid, close } = mountGrid();
        try {
            await grid.ImportXlsx(original.slice());
            grid.SwitchSheet(0);
            edit(grid);
            const bytes = await grid.ExportXlsx();
            const xml = await (await window.JSZip.loadAsync(bytes)).files[part].async('string');
            const { grid: again, close: closeAgain } = mountGrid();
            try {
                await again.ImportXlsx(bytes);
                again.SwitchSheet(0);
                const formulas = (g) => { const out = []; for (let r = 60; r < 130; r++) for (let c = 3; c < 11; c++) out.push(g.GetCell(r, c).formula ?? null); return out; };
                return { xml, same: JSON.stringify(formulas(again)) === JSON.stringify(formulas(grid)), again: { hidden: again.GetHiddenRows(), height: again.GetRowHeight(5) } };
            } finally { closeAgain(); }
        } finally { close(); }
    };
    try {
        check('the file has what the check is about', [element(was, 'sheetPr').includes('codeName'), element(was, 'ignoredErrors') !== '', count(was, /t="shared"/g) > 100, count(was, /<row\b[^>]*\ss="/g) > 10], [true, true, true, true]);
        // A hidden row, another row height: the part is the file's but for those two row tags.
        const rows = await edited((grid) => { grid.HideRows(9); grid.SetRowHeight(5, 53); });
        const withoutRows = (xml) => xml.replace(/<row r="(6|10)"[^>]*>/g, '');
        check('row hidden and row height: nothing but the two row tags differs from the file\'s part', withoutRows(rows.xml) === withoutRows(was), true);
        check('row hidden: the row says so and keeps what it said', rowTag(rows.xml, 10), rowTag(was, 10).replace('>', ' hidden="1">'));
        check('row height: the row states it', [/ ht="39\.\d+"| ht="40"/.test(rowTag(rows.xml, 6)), rowTag(rows.xml, 6).includes('customHeight="1"'), rowTag(rows.xml, 6).includes('x14ac:dyDescent')], [true, true, true]);
        check('row hidden and row height: read again', [rows.again.hidden.includes('10:10'), Math.round(rows.again.height)], [true, 53]);
        // A link: the rows are the file's, and so is what the part says beyond a sheet.
        const linked = await edited((grid) => grid.SetHyperlink('B401', 'https://example.com/a?b=1&c=2'));
        check('link added: the part names it', element(linked.xml, 'hyperlinks').includes('ref="B401"'), true);
        check('link added: shared formulas, styled rows, sheetPr, the views, sheetFormatPr and ignoredErrors are the file\'s',
            [count(linked.xml, /t="shared"/g), count(linked.xml, /<row\b[^>]*\ss="/g), element(linked.xml, 'sheetPr'), element(linked.xml, 'sheetViews'), element(linked.xml, 'sheetFormatPr'), element(linked.xml, 'ignoredErrors')],
            [count(was, /t="shared"/g), count(was, /<row\b[^>]*\ss="/g), element(was, 'sheetPr'), element(was, 'sheetViews'), element(was, 'sheetFormatPr'), element(was, 'ignoredErrors')]);
        check('link added: the formulas read as they did', linked.same, true);
        // A row inserted: written from the cells, with what the part said where the rows are now.
        const inserted = await edited((grid) => grid.InsertRows(100, 1));
        check('row inserted: sheetPr and sheetFormatPr are the file\'s', [element(inserted.xml, 'sheetPr'), element(inserted.xml, 'sheetFormatPr')], [element(was, 'sheetPr'), element(was, 'sheetFormatPr')]);
        check('row inserted: ignoredErrors follow their cells', [element(inserted.xml, 'ignoredErrors').includes('E96:E110'), element(inserted.xml, 'ignoredErrors').includes('F138:K158'), count(element(inserted.xml, 'ignoredErrors'), /<ignoredError /g)],
            [true, true, count(element(was, 'ignoredErrors'), /<ignoredError /g)]);
        check('row inserted: the rows keep their own styles', count(inserted.xml, /<row\b[^>]*\ss="/g), count(was, /<row\b[^>]*\ss="/g));
        check('row inserted: the selection follows', [element(inserted.xml, 'sheetViews').includes('activeCell="B309"'), element(inserted.xml, 'sheetViews').includes('topLeftCell="A309"')], [true, true]);
        check('row inserted: formulas that read alike are written once', count(inserted.xml, /t="shared"/g) * 10 >= count(was, /t="shared"/g) * 8, true);
        check('row inserted: the formulas read as the sheet has them', inserted.same, true);
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    return lines;
};

// A note whose text already starts with its author's name says the name once (`data-comment`); a note
// edited without an author keeps its author.
checks['comment-author-once'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    try {
        const said = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`)?.getAttribute('data-comment');
        grid.SetComment(1, 1, 'BIWS:\n4% tax', 'BIWS');
        grid.SetComment(2, 1, 'a remark', 'BIWS');
        await frame();
        check('a text that starts with the author: the author is said once', said(1, 1), 'BIWS:\n4% tax');
        check('a text that does not: the author, then the text', said(2, 1), 'BIWS:\na remark');
        grid.SetComment(2, 1, 'another remark');
        await frame();
        check('edited without an author: the note keeps its author', [grid.GetComment(2, 1).author, said(2, 1)], ['BIWS', 'BIWS:\nanother remark']);
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// A sheet of a file can be extended below its last used row and right of its last used column: the arrow
// keys leave the used range, typing lands in the new cell, a cell beyond the grid can be written (SetCell)
// and pasted into; the saved file has the new cells.
checks['sheet-grows-beyond-used-range'] = async () => {
    const { lines, check } = makeLog();
    const { bytesOf } = await import('/component.js');
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    for (const file of ['tests/real-world/milk-street-office.xlsx']) {
        const { grid, host, close } = mountGrid();
        try {
            await grid.ImportXlsx((await bytesOf(file)).slice());
            grid.SwitchSheet(0);
            await frame();
            const used = grid.GetUsedRange();
            check('the grid has rows below the used range', grid.GetModel().RowCount >= used.lastRow + 1 + 50, true);
            grid.SelectCell(used.lastRow, 1);
            await frame();
            // (A click, so that the keys go to the grid.)
            const clickCell = async (row, col) => { grid.SelectCell(row, col); await frame(); const box = cellBox(host, row, col); await window.realMouse('click', box.left + box.width / 2, box.top + box.height / 2); await frame(); };
            await clickCell(used.lastRow, 1);
            const before = grid.GetCell(used.lastRow, 1).value;
            for (let i = 0; i < 5; i++) { await window.realKey('ArrowDown'); await frame(); }
            check('ArrowDown five times from the last used row: the caret is five rows below', grid.GetActiveCell().Row, used.lastRow + 5);
            // (The Normal font is Calibri 12: rows of 15.75 pt, 21 px, and digits of 8 px, so defaultColWidth="11" is 88 px as in
            // Excel 2019 - also for the rows and columns the grid adds.)
            const added = cellBox(host, used.lastRow + 5, 1), wide = grid.GetModel().ColCount - 1;
            check('the rows below the used range have the height the sheet gives a row, the columns past it its width', [Math.round(added.height), Math.round(grid.GetRowHeight(used.lastRow + 5)), Math.round(grid.GetRowHeight(used.lastRow - 150)), Math.round(grid.GetColumnWidth(wide))], [21, 21, 21, 88]);
            await window.realType('below'); await window.realKey('Enter'); await frame();
            check('typing lands in the new row, the last used row is as it was', [grid.GetCell(used.lastRow + 5, 1).value, grid.GetCell(used.lastRow, 1).value], ['below', before]);
            // Far down, step by step: the grid keeps growing ahead of the caret.
            grid.SelectCell(grid.GetModel().RowCount - 1, 1); await frame();
            const rows = grid.GetModel().RowCount;
            check('the caret on the last row of the grid: more rows below it', rows > grid.GetActiveCell().Row + 50, true);
            // Columns.
            const lastCol = grid.GetModel().ColCount - 1;
            await clickCell(3, lastCol);
            const cols = lastCol + 1;
            for (let i = 0; i < 3; i++) { await window.realKey('ArrowRight'); await frame(); }
            check('ArrowRight three times from the last column of the grid: the caret is three columns right', grid.GetActiveCell().Col, cols - 1 + 3);
            // A cell beyond the grid written by the host.
            const far = grid.GetModel().RowCount + 500;
            grid.SetCell(far, 2, '42');
            check('SetCell beyond the grid: the cell is a cell of the sheet', [grid.GetCell(far, 2).value, grid.GetModel().RowCount > far], [42, true]);
            const zip = await window.JSZip.loadAsync(await grid.ExportXlsx());
            const sheet = await zip.files['xl/worksheets/sheet1.xml'].async('string');
            check('saved: the new cells are in the sheet', [new RegExp(`<c r="B${used.lastRow + 6}"`).test(sheet), new RegExp(`<c r="C${far + 1}"`).test(sheet)], [true, true]);
        } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
        finally { close(); }
    }
    return lines;
};

// A recalculation of many formulas reads cells through a hash index of each sheet's cells, handed on from
// edit to edit: what it gives is what a workbook opened afresh gives for the same entry, edit after edit,
// and after undo and redo.
checks['recalc-many-formulas-same-values'] = async () => {
    const { lines, check } = makeLog();
    const { bytesOf } = await import('/component.js');
    const original = await bytesOf('tests/real-world/milk-street-office.xlsx');
    const values = (g) => { const out = []; for (let sheet = 0; sheet < 2; sheet++) { g.SwitchSheet(sheet); const used = g.GetUsedRange(); for (let r = used.firstRow; r <= used.lastRow; r++) for (let c = used.firstCol; c <= used.lastCol; c++) { const v = g.GetCell(r, c).value; if (v !== null && v !== undefined && v !== '') out.push(`${sheet}!${r},${c}=${typeof v === 'number' ? v.toPrecision(12) : v}`); } } g.SwitchSheet(0); return out; };
    const differing = (a, b) => { const out = []; for (let i = 0; i < Math.max(a.length, b.length) && out.length < 3; i++) if (a[i] !== b[i]) out.push(`${a[i]} / ${b[i]}`); return out; };
    const fresh = async (text) => { const { grid, close } = mountGrid(); try { await grid.ImportXlsx(original.slice()); grid.SwitchSheet(0); grid.EnterText(11, 4, text); return { values: values(grid), stats: grid.GetRecalcStats() }; } finally { close(); } };
    const { grid, close } = mountGrid();
    try {
        await grid.ImportXlsx(original.slice());
        grid.SwitchSheet(0);
        const was = grid.GetCell(11, 4).value;
        const first = await fresh('0.07'), second = await fresh('0.05');
        check('the entry reaches most of the workbook\'s formulas', first.stats.EvaluatedFormulas > 10000, true);
        check('the two entries give different workbooks', differing(first.values, second.values).length > 0, true);
        grid.EnterText(11, 4, '0.07');
        check('first entry: as in a workbook opened afresh', differing(values(grid), first.values), []);
        grid.EnterText(11, 4, '0.05');
        check('second entry: as in a workbook opened afresh', differing(values(grid), second.values), []);
        grid.EnterText(11, 4, '0.07');
        check('third entry (the first again): the same values', differing(values(grid), first.values), []);
        grid.Undo();
        check('undo: the values of the second entry', differing(values(grid), second.values), []);
        grid.Redo();
        check('redo: the values of the first entry', differing(values(grid), first.values), []);
        // An edit of a cell few formulas read, then the many again.
        grid.EnterText(0, 13, '5');
        grid.EnterText(11, 4, '0.05');
        grid.EnterText(0, 13, '');
        check('after other edits: the values of the second entry', differing(values(grid), second.values), []);
        grid.EnterText(11, 4, String(was));
        const back = values(grid);
        const { grid: opened, close: closeOpened } = mountGrid();
        try { await opened.ImportXlsx(original.slice()); opened.SwitchSheet(0); opened.EnterText(11, 4, String(was)); check('the file\'s own value typed again: as in a workbook opened afresh', differing(back, values(opened)), []); } finally { closeOpened(); }
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// A sheet with a pivot table copied, deleted, renamed (DuplicateSheet, DeleteSheet, RenameSheet, ExportXlsx): the
// copy has a pivot table of its own; the slicers no longer list a deleted sheet's pivot tables, and a chart
// drawn from one names none; a renamed sheet is named in the charts of its pivot table.
checks['pivot-sheets-copied-deleted-renamed'] = async () => {
    const { lines, check } = makeLog();
    const { bytesOf } = await import('/component.js');
    const { grid, close } = mountGrid();
    try {
        const original = await bytesOf('tests/real-world/dashboard-sales-analysis.xlsx');
        const before = await window.JSZip.loadAsync(original);
        const pivots = (zip) => Object.keys(zip.files).filter((name) => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(name)).sort();
        const tabIds = async (zip) => { const out = []; for (const name of Object.keys(zip.files).sort()) if (/^xl\/slicerCaches\/slicerCache\d+\.xml$/.test(name)) for (const m of (await zip.files[name].async('string')).matchAll(/<pivotTable tabId="(\d+)"/g)) out.push(m[1]); return out; };
        const sources = async (zip) => { const out = []; for (const name of Object.keys(zip.files).sort()) if (/^xl\/charts\/chart\d+\.xml$/.test(name)) { const m = /<c:pivotSource><c:name>([^<]*)<\/c:name>/.exec(await zip.files[name].async('string')); if (m) out.push(m[1]); } return out; };
        await grid.ImportXlsx(original.slice());
        const indexOf = (name) => grid.GetSheetNames().indexOf(name);
        // -- copied
        const copyIndex = grid.DuplicateSheet(indexOf('Category Pivot'));
        const copied = await window.JSZip.loadAsync(await grid.ExportXlsx());
        const added = pivots(copied).filter((name) => !before.files[name]);
        check('copied: one pivot table part more', [pivots(copied).length, added.length], [pivots(before).length + 1, 1]);
        const root = (xml) => /<pivotTableDefinition\b[^>]*>/.exec(xml)[0];
        const copyRoot = root(await copied.files[added[0]].async('string')), sourceRoot = root(await before.files['xl/pivotTables/pivotTable2.xml'].async('string'));
        const attr = (tag, name) => (new RegExp(`\\s${name}="([^"]*)"`).exec(tag) || [])[1];
        check('copied: another name and uid, the same cache', [attr(copyRoot, 'name'), attr(copyRoot, 'cacheId'), attr(copyRoot, 'xr:uid') !== attr(sourceRoot, 'xr:uid')], [attr(sourceRoot, 'name') + '_2', attr(sourceRoot, 'cacheId'), true]);
        const book = await copied.files['xl/workbook.xml'].async('string');
        check('copied: the copy is a sheet of the workbook', new RegExp(`<sheet name="${grid.GetSheetNames()[copyIndex].replace(/[()]/g, '\\$&')}"`).test(book), true);
        check('copied: the slicers list what they listed', await tabIds(copied), await tabIds(before));
        // -- deleted (the copy first: back to the file's sheets)
        grid.DeleteSheet(copyIndex);
        const lineId = /<sheet name="Line Pivot" sheetId="(\d+)"/.exec(await before.files['xl/workbook.xml'].async('string'))[1];
        grid.DeleteSheet(indexOf('Line Pivot'));
        const deleted = await window.JSZip.loadAsync(await grid.ExportXlsx());
        check('deleted: the slicers list no pivot table of the sheet', await tabIds(deleted), (await tabIds(before)).filter((id) => id !== lineId));
        check('deleted: the charts of its pivot table name none, the others theirs', await sources(deleted), (await sources(before)).filter((name) => !name.includes(']Line Pivot!')));
        check('deleted: the pivot table parts are the file\'s less the sheet\'s', pivots(deleted).length, pivots(before).length - 1);
        // -- renamed
        grid.RenameSheet(indexOf('Category Pivot'), 'Categories');
        const renamed = await window.JSZip.loadAsync(await grid.ExportXlsx());
        check('renamed: the charts name the pivot table on the renamed sheet', await sources(renamed), (await sources(before)).filter((name) => !name.includes(']Line Pivot!')).map((name) => name.replace(']Category Pivot!', ']Categories!')));
        const { grid: again, close: closeAgain } = mountGrid();
        try { await again.ImportXlsx(await grid.ExportXlsx()); check('read again: the sheets', again.GetSheetNames(), grid.GetSheetNames()); } finally { closeAgain(); }
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// A workbook in the 1904 date system stays in it: GetDateSystem, the file's serials, typed dates, the texts
// of dates, and the save - beside a grid with a workbook in the 1900 system on the same page.
checks['date-system-1904'] = async () => {
    const { lines, check } = makeLog();
    const { bytesOf } = await import('/component.js');
    const { grid, close } = mountGrid();
    const { grid: other, close: closeOther } = mountGrid();
    try {
        check('a new workbook is in the 1900 system', grid.GetDateSystem(), 1900);
        const original = await bytesOf('tests/test-documents-excel/pandas-times_1904.xlsx');
        await grid.ImportXlsx(original.slice());
        await other.ImportXlsx((await bytesOf('tests/real-world/milk-street-office.xlsx')).slice());
        check('the date systems of the two workbooks', [grid.GetDateSystem(), other.GetDateSystem()], [1904, 1900]);
        check('a time of the 1904 file is the file\'s number', Math.abs(grid.GetCell(1, 0).value - 0.043090277777777776) < 1e-12, true);
        check('... shown as the time', grid.GetCell(1, 0).text, '01:02:03.000');
        grid.SetCell(20, 2, '2018-12-28');
        other.SwitchSheet(0);
        const free = 9;         // (N10 and N11: nothing there.)
        other.SetCell(free, 13, '2018-12-28');
        check('a date typed into each: the serial of its workbook\'s system', [grid.GetCell(20, 2).value, other.GetCell(free, 13).value], [42000, 43462]);
        check('... and both show the date', [/2018/.test(grid.GetCell(20, 2).text), /2018/.test(other.GetCell(free, 13).text), grid.GetCell(20, 2).text], [true, true, other.GetCell(free, 13).text]);
        grid.SetCell(21, 2, '=YEAR(C21)*100+MONTH(C21)');
        other.SetCell(free + 1, 13, `=YEAR(N${free + 1})*100+MONTH(N${free + 1})`);
        check('the date functions read the serial in the workbook\'s system', [grid.GetCell(21, 2).value, other.GetCell(free + 1, 13).value], [201812, 201812]);
        grid.SetCell(22, 2, '=DATE(1904,1,2)');
        check('DATE gives the workbook\'s serial', grid.GetCell(22, 2).value, 1);
        const bytes = await grid.ExportXlsx();
        const zip = await window.JSZip.loadAsync(bytes);
        check('saved: the workbook says date1904', /<workbookPr\b[^>]*\sdate1904="(1|true)"/.test(await zip.files['xl/workbook.xml'].async('string')), true);
        const sheet = await zip.files['xl/worksheets/sheet1.xml'].async('string');
        check('saved: the typed serial, and the file\'s numbers as they were', [/<c r="C21"[^>]*><v>42000<\/v><\/c>/.test(sheet), sheet.includes('<c r="A2" s="1"><v>4.3090277777777776E-2</v></c>')], [true, true]);
        const otherBytes = await other.ExportXlsx();
        check('the 1900 workbook saved: no date1904', /date1904="(1|true)"/.test(await (await window.JSZip.loadAsync(otherBytes)).files['xl/workbook.xml'].async('string')), false);
        await other.ImportXlsx(bytes);
        check('read again: the system, the serial and the date', [other.GetDateSystem(), other.GetCell(20, 2).value, /2018/.test(other.GetCell(20, 2).text), other.GetCell(21, 2).value], [1904, 42000, true, 201812]);
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); closeOther(); }
    return lines;
};

// Two workbooks side by side, one in each date system: whatever one grid is asked, the other's
// dates stay its own - there is no date system "at work" that the grid that acted last sets.
checks['date-system-side-by-side'] = async () => {
    const { lines, check } = makeLog();
    const { bytesOf } = await import('/component.js');
    const { grid, close } = mountGrid();
    const { grid: other, close: closeOther } = mountGrid();
    try {
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/pandas-times_1904.xlsx')).slice());
        check('the date systems', [grid.GetDateSystem(), other.GetDateSystem()], [1904, 1900]);
        // C21:C23 of the 1904 workbook and A1:A3 of the 1900 one hold the same three days.
        const days = ['2018-12-28', '2019-01-15', '2019-02-01'];
        days.forEach((day, i) => { grid.SetCell(20 + i, 2, day); other.SetCell(i, 0, day); });
        grid.SetCell(19, 2, 'Day'); other.SetCell(3, 0, 'x');
        check('the serials are each workbook\'s', [grid.GetCell(20, 2).value, other.GetCell(0, 0).value], [42000, 43462]);
        const shown = grid.GetCell(20, 2).text;
        check('both show the day', [/2018/.test(shown), other.GetCell(0, 0).text], [true, shown]);
        // Every question below is put to one grid right after the other was the last to act.
        const afterOther = (ask) => { other.GetCell(0, 0); return ask(); };
        const afterGrid = (ask) => { grid.GetCell(20, 2); return ask(); };
        check('1904: GetRange reads the dates', afterOther(() => grid.GetRange('C21:C23').map((row) => row[0].text)), [shown, grid.GetCell(21, 2).text, grid.GetCell(22, 2).text]);
        check('1900: GetRange reads the dates', afterGrid(() => other.GetRange('A1:A3').map((row) => row[0].text)), [shown, other.GetCell(1, 0).text, other.GetCell(2, 0).text]);
        check('1904: Find finds the day by its year', afterOther(() => grid.Find('2018', {}).map((match) => match.address)), ['C21']);
        check('1900: Find finds the day by its year', afterGrid(() => other.Find('2018', {}).map((match) => match.address)), ['A1']);
        check('1904: Find in formulas reads the date as typed', afterOther(() => grid.Find('12/28/2018', { lookIn: 'formulas' }).map((match) => match.address)), ['C21']);
        check('1900: Find in formulas reads the date as typed', afterGrid(() => other.Find('12/28/2018', { lookIn: 'formulas' }).map((match) => match.address)), ['A1']);
        grid.SetAutoFilter('C20:C23');
        check('1904: a filter\'s list of values', afterOther(() => grid.GetFilterValues(2).values), afterGrid(() => grid.GetFilterValues(2).values));
        check('... are the days', afterOther(() => grid.GetFilterValues(2).values.map((text) => /20(18|19)/.test(text))), [true, true, true]);
        afterOther(() => grid.SetFilter(2, { conditions: [{ operator: 'greaterOrEqual', value: '2019-01-15' }] }));
        check('1904: a filter by a date hides the earlier day only', grid.GetFilters().hiddenRows, [20]);
        grid.ClearAutoFilter();
        // A formula entered in one grid while the other was the last to act; TEXT, DATE and the calendar functions.
        afterOther(() => grid.SetCell(24, 2, '=TEXT(C21,"yyyy-mm-dd")&"|"&YEAR(C22)&"|"&DATE(2018,12,28)&"|"&WEEKDAY(C21)'));
        afterGrid(() => other.SetCell(5, 0, '=TEXT(A1,"yyyy-mm-dd")&"|"&YEAR(A2)&"|"&DATE(2018,12,28)&"|"&WEEKDAY(A1)'));
        check('formulas of both', [afterOther(() => grid.GetCell(24, 2).value), afterGrid(() => other.GetCell(5, 0).value)], ['2018-12-28|2019|42000|6', '2018-12-28|2019|43462|6']);
        // A recalculation of one workbook, the other having acted last.
        afterOther(() => grid.SetCell(21, 2, '2020-03-01'));
        afterGrid(() => other.SetCell(1, 0, '2020-03-01'));
        check('recalculated: both', [afterOther(() => grid.GetCell(24, 2).value), afterGrid(() => other.GetCell(5, 0).value)], ['2018-12-28|2020|42000|6', '2018-12-28|2020|43462|6']);
        // A sheet added to the 1904 workbook is in its system.
        afterOther(() => grid.AddSheet('More'));
        grid.SwitchSheet(grid.GetSheetNames().indexOf('More'));
        afterOther(() => grid.SetCell(0, 0, '2018-12-28'));
        check('a sheet added to the 1904 workbook: a typed date', [grid.GetDateSystem(), grid.GetCell(0, 0).value, grid.GetCell(0, 0).text], [1904, 42000, shown]);
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); closeOther(); }
    return lines;
};

// Spill ranges for hosts: GetCell(...).spill, GetSpillRange, and the outline around the active cell's array.
checks['chrome-spill'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        grid.SetCell(1, 1, '=SEQUENCE(4,3)'); grid.SetCell(0, 5, '=SEQUENCE(2)'); grid.SetCell(8, 0, '=1+1'); grid.SetCell(9, 0, 'text');
        await frame();
        const outlines = () => [...host.querySelectorAll('.spreadsheet-spill-outline')].filter((n) => n.getClientRects().length > 0);
        check('the anchor: its range', [grid.GetCell(1, 1).spill, grid.GetSpillRange(1, 1)], [{ anchor: 'B2', range: 'B2:D5' }, { anchor: 'B2', range: 'B2:D5' }]);
        check('a cell the array filled', [grid.GetCell(4, 3).spill, grid.GetCell(4, 3).formula, grid.GetCell(4, 3).value], [{ anchor: 'B2', range: 'B2:D5' }, null, 12]);
        check('another array', grid.GetSpillRange(1, 5), { anchor: 'F1', range: 'F1:F2' });
        check('a cell beside it, a plain formula, text', [grid.GetCell(5, 1).spill, grid.GetSpillRange(8, 0), grid.GetCell(9, 0).spill], [null, null, null]);
        check('GetRange gives it per cell', grid.GetRange('A1:B2').map((row) => row.map((cell) => cell.spill && cell.spill.range)), [[null, null], [null, 'B2:D5']]);
        check('the active cell is outside: no outline', [grid.GetSelectionRef(), outlines().length], ['A1', 0]);
        // A real click inside the array.
        const c3 = cellBox(host, 2, 2);
        await realClick(c3.left + c3.width / 2, c3.top + c3.height / 2);
        const first = cellBox(host, 1, 1), last = cellBox(host, 4, 3);
        const outline = outlines()[0], box = outline && outline.getBoundingClientRect();
        check('a click in the array: a thin outline around its range', [grid.GetSelectionRef(), outlines().length, outline && outline.getAttribute('data-range'), box && [near(box.left, first.left), near(box.top, first.top), near(box.right, last.right, 1.5), near(box.bottom, last.bottom, 1.5)], outline && getComputedStyle(outline).borderTopWidth], ['C3', 1, 'B2:D5', [true, true, true, true], '1px']);
        const edge = await pixelAt(box.left + 40, box.top);
        check('... drawn in blue', edge[2] > 150 && edge[0] < 120, true);
        await window.realKey('ArrowUp'); await window.realKey('ArrowUp'); await frame();
        check('the keyboard leaves the array: the outline goes', [grid.GetSelectionRef(), outlines().length], ['C1', 0]);
        // Something in the way: #SPILL!, no range.
        grid.EnterText(3, 2, 'blocked'); await frame();
        check('blocked: #SPILL! and no range', [grid.GetCell(1, 1).value, grid.GetSpillRange(1, 1), grid.GetCell(2, 1).spill], ['#SPILL!', null, null]);
        grid.Undo(); await frame();
        check('undo: the array is back', [grid.GetSpillRange(4, 3), grid.GetCell(4, 3).value], [{ anchor: 'B2', range: 'B2:D5' }, 12]);
        grid.Redo(); grid.Undo();
        // The array changes size.
        grid.EnterText(1, 1, '=SEQUENCE(2,2)'); await frame();
        check('a smaller array: its range, and the cells it left', [grid.GetSpillRange(1, 1), grid.GetSpillRange(4, 3)], [{ anchor: 'B2', range: 'B2:C3' }, null]);
        grid.SelectCell(2, 2); await frame();
        check('the outline follows', outlines().map((n) => n.getAttribute('data-range')), ['B2:C3']);
        // Frozen panes: one outline piece per pane.
        grid.FreezeAt(2, 2); await frame();
        check('over frozen panes: a piece in each pane the array reaches', outlines().length, 4);
        grid.Unfreeze();
        // Export, re-import.
        const again = await reopened(grid);
        try {
            check('opened again: the ranges', [again.grid.GetSpillRange(2, 2), again.grid.GetSpillRange(1, 5), again.grid.GetCell(1, 1).formula], [{ anchor: 'B2', range: 'B2:C3' }, { anchor: 'F1', range: 'F1:F2' }, '=SEQUENCE(2,2)']);
        } finally { again.close(); }
        check('row / col outside the sheet throw', (() => { try { grid.GetSpillRange(-1, 0); return null; } catch (e) { return /GetSpillRange/.test(String(e.message)); } })(), true);
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// An export of a workbook nothing was changed in is the file that was opened; any change writes it anew.
checks['export-verbatim'] = async () => {
    const { lines, check } = makeLog();
    const { bytesOf } = await import('/component.js');
    const { grid, close } = mountGrid();
    try {
        const original = await bytesOf('tests/real-world/milk-street-office.xlsx');
        const same = (a, b) => a.length === b.length && a.every((value, index) => value === b[index]);
        const given = original.slice();
        await grid.ImportXlsx(given);
        given.fill(0);   // (the host's buffer is its own: the grid keeps a copy)
        check('opened, nothing done: the opened bytes', same(await grid.ExportXlsx(), original), true);
        grid.SelectCell(5, 3); grid.SelectRange(2, 2, 6, 4);
        check('a selection is not a change', same(await grid.ExportXlsx(), original), true);
        const revision = grid.GetRevision();
        grid.EnterText(40, 12, 'typed');
        const edited = await grid.ExportXlsx();
        check('an edit: written anew, with the edit', [same(edited, original), grid.GetRevision() !== revision], [false, true]);
        const zip = await window.JSZip.loadAsync(edited);
        check('... a package of its own', Object.keys(zip.files).length > 5, true);
        grid.Undo();
        check('the edit undone: the opened bytes again', [grid.GetRevision() === revision, same(await grid.ExportXlsx(), original)], [true, true]);
        grid.Redo(); grid.Undo();
        grid.SwitchSheet(1);
        check('another sheet on display: written anew (the file says which sheet is shown)', same(await grid.ExportXlsx(), original), false);
        grid.SwitchSheet(0);
        check('back on the first: the opened bytes', [same(await grid.ExportXlsx(), original), grid.GetRevision() === revision, grid.IsModified(), grid.GetActiveSheetIndex(), grid.GetZoom()], [true, true, false, 0, 1]);
        grid.SetShowGridLines(!grid.GetShowGridLines());
        check('grid lines switched: written anew', same(await grid.ExportXlsx(), original), false);
        grid.SetShowGridLines(!grid.GetShowGridLines());
        // (Switched back by hand it is a new revision - the grid does not know it is where it was -; undone it is the old one.)
        check('... and back by hand: as the revision says', same(await grid.ExportXlsx(), original), grid.GetRevision() === revision);
        grid.Undo(); grid.Undo();
        check('... undone: the opened bytes', [grid.GetRevision() === revision, same(await grid.ExportXlsx(), original)], [true, true]);
        grid.SetDefinedName('Fresh', '=Annual!$A$1');
        check('a defined name: written anew', same(await grid.ExportXlsx(), original), false);
        // A workbook that was not opened from a file is always written.
        const made = mountGrid();
        try { made.grid.SetCell(0, 0, '1'); check('a workbook made here is written', (await made.grid.ExportXlsx()).length > 1000, true); } finally { made.close(); }
        // The same grid opening another file forgets the first.
        const other = await bytesOf('tests/real-world/chart-zoo.xlsx');
        await grid.ImportXlsx(other.slice());
        check('another file opened: its bytes', same(await grid.ExportXlsx(), other), true);
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// The strips' elements are reused from one drawing to the next: a bracket or a strip never keeps a button's digit.
checks['chrome-outline-reuse'] = async () => {
    const { lines, check } = makeLog();
    const stray = (host) => [...host.querySelectorAll('.spreadsheet-outline-line, .spreadsheet-outline-strip')].filter((n) => n.textContent !== '' || n.hasAttribute('data-level') || n.hasAttribute('title')).map((n) => [n.className.replace('spreadsheet-outline ', ''), n.textContent]);
    const labels = (host) => [[...host.querySelectorAll('.spreadsheet-outline-level')].every((n) => /^[1-8]$/.test(n.textContent)), [...host.querySelectorAll('.spreadsheet-outline-button')].every((n) => n.textContent === '+' || n.textContent === '\u2212')];
    {
        const { grid, host, close } = mountGrid();
        try {
            for (let r = 0; r < 40; r++) grid.SetCell(r, 0, String(r));
            grid.GroupRows('22:30'); await frame();
            grid.GroupRows('24:26'); await frame();
            check('rows grouped, then a group inside: no text in a bracket or strip', [stray(host), labels(host)], [[], [true, true]]);
            grid.UngroupRows('24:26'); await frame();
            check('... ungrouped again', [stray(host), labels(host)], [[], [true, true]]);
        } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
        finally { close(); }
    }
    {
        const { grid, host, close } = mountGrid();
        try {
            for (let r = 0; r < 12; r++) grid.SetCell(r, 0, String(r));
            grid.GroupColumns('C:D'); await frame();
            grid.GroupRows('4:6'); await frame();
            grid.SetOutlineSummary(false, false); await frame();
            grid.GroupRows('5:5'); await frame();
            check('columns, rows, summaries moved, a group inside: no text in a bracket or strip', [stray(host), labels(host)], [[], [true, true]]);
            grid.Undo(); await frame(); grid.Undo(); await frame(); grid.Undo(); await frame();
            check('... undone step by step', [stray(host), labels(host)], [[], [true, true]]);
        } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
        finally { close(); }
    }
    return lines;
};

// An entry in a cell an array spilled into starts empty; the formula bar's text for it is the anchor's formula, as a ghost.
checks['spill-edit'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        grid.SetCell(1, 1, '=SEQUENCE(3,2)'); grid.SetCell(0, 0, '=1+1'); await frame();
        const states = []; const off = grid.OnEditStateChange ? grid.OnEditStateChange((state) => states.push(state)) : null;
        grid.SelectCell(2, 2); await frame();
        check('a spilled cell: its value', grid.GetCell(2, 2).value, 4);
        check('... the formula of the cell is empty, the ghost is the anchor\'s', [grid.GetCurrentCellFormula(), grid.GetGhostFormula()], ['', '=SEQUENCE(3,2)']);
        await window.realKey('F2'); await frame();
        const box = host.querySelector('.spreadsheet-edit-input');
        check('F2: the entry starts empty', [grid.IsEditing ? grid.IsEditing() : true, box ? box.value : null], [true, '']);
        await window.realKey('Escape'); await frame();
        check('Escape: the array is as it was', [grid.GetCell(2, 2).value, grid.GetCell(1, 1).formula], [4, '=SEQUENCE(3,2)']);
        grid.StartEdit(); await frame();
        check('StartEdit(): empty too', (host.querySelector('.spreadsheet-edit-input') || {}).value, '');
        grid.CancelEdit(); await frame();
        grid.SelectCell(1, 1); await frame();
        check('the anchor: its formula, no ghost', [grid.GetCurrentCellFormula(), grid.GetGhostFormula()], ['=SEQUENCE(3,2)', null]);
        await window.realKey('F2'); await frame();
        check('F2 on the anchor: the formula', (host.querySelector('.spreadsheet-edit-input') || {}).value, '=SEQUENCE(3,2)');
        await window.realKey('Escape'); await frame();
        grid.SelectCell(0, 0); await frame();
        check('an ordinary formula cell', [grid.GetCurrentCellFormula(), grid.GetGhostFormula()], ['=1+1', null]);
        if (typeof off === 'function') off();
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// A list validation whose source is a defined name that is a formula (a dynamic range), or a formula itself.
checks['validation-list-formula'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    try {
        const lists = grid.AddSheet('Lists');
        grid.SwitchSheet(lists);
        ['North', 'South', 'East'].forEach((text, row) => grid.SetCell(row, 0, text));
        grid.SwitchSheet(0);
        grid.SetDefinedName('Regions', '=OFFSET(Lists!$A$1,0,0,COUNTA(Lists!$A:$A),1)');
        grid.SetDefinedName('Ends', '=Lists!$A$1:INDEX(Lists!$A:$A,COUNTA(Lists!$A:$A))');
        grid.SetDefinedName('Plain', '=Lists!$A$1:$A$2');
        grid.SetValidation('B2:B5', { type: 'list', source: 'Regions' });
        grid.SetValidation('C2', { type: 'list', source: 'Ends' });
        grid.SetValidation('D2', { type: 'list', source: 'Plain' });
        grid.SetValidation('E2', { type: 'list', source: 'OFFSET(Lists!$A$1,1,0,2,1)' });
        check('a name that is an OFFSET formula', grid.GetValidationList(1, 1), ['North', 'South', 'East']);
        check('... on another cell of the rule', grid.GetValidationList(4, 1), ['North', 'South', 'East']);
        check('a name that ends in INDEX', grid.GetValidationList(1, 2), ['North', 'South', 'East']);
        check('a name that is a range of another sheet', grid.GetValidationList(1, 3), ['North', 'South']);
        check('a formula as the source', grid.GetValidationList(1, 4), ['South', 'East']);
        grid.SwitchSheet(lists); grid.SetCell(3, 0, 'West'); grid.SwitchSheet(0);
        check('the list follows the cells', grid.GetValidationList(1, 1), ['North', 'South', 'East', 'West']);
        grid.SetCell(1, 1, 'West');
        check('an item of it is accepted', grid.GetCell(1, 1).value, 'West');
        check('nothing is circled', grid.CircleInvalid(false).length, 0);
        grid.SetCell(2, 1, 'Nowhere');
        check('... and something else is refused or flagged', [grid.GetCell(2, 1).value === 'Nowhere' ? grid.CircleInvalid(false).map((c) => c.address) : 'refused'].flat(), grid.GetCell(2, 1).value === 'Nowhere' ? ['B3'] : ['refused']);
        // Saved and opened again: the rules say what they said, and the lists are read from the (not yet shown) sheet.
        const bytes = await grid.ExportXlsx();
        const zip = await window.JSZip.loadAsync(bytes);
        const sheetXml = await zip.files['xl/worksheets/sheet1.xml'].async('string');
        check('written: the sources as they are', [/<formula1>Regions<\/formula1>/.test(sheetXml), /<formula1>Ends<\/formula1>/.test(sheetXml), /OFFSET\(Lists!\$A\$1,1,0,2,1\)/.test(sheetXml)], [true, true, true]);
        const workbookXml = await zip.files['xl/workbook.xml'].async('string');
        check('written: the names', [/>OFFSET\(Lists!\$A\$1,0,0,COUNTA\(Lists!\$A:\$A\),1\)</.test(workbookXml), /Lists!\$A\$1:INDEX\(Lists!\$A:\$A,COUNTA\(Lists!\$A:\$A\)\)/.test(workbookXml)], [true, true]);
        const again = mountGrid();
        try {
            await again.grid.ImportXlsx(bytes);
            check('opened again: the evaluated lists', [again.grid.GetValidationList(1, 1), again.grid.GetValidationList(1, 2), again.grid.GetValidationList(1, 4)], [['North', 'South', 'East', 'West'], ['North', 'South', 'East', 'West'], ['South', 'East']]);
            check('... the rule reads back with its source', again.grid.GetValidationAt(1, 1).source, 'Regions');
            const twice = await again.grid.ExportXlsx();
            const twiceXml = await (await window.JSZip.loadAsync(twice)).files['xl/worksheets/sheet1.xml'].async('string');
            const rules = (xml) => (/<dataValidations[\s\S]*<\/dataValidations>/.exec(xml) || [''])[0];
            check('... and an unchanged save writes the same rules', rules(twiceXml), rules(sheetXml));
        } finally { again.close(); }
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};
// Cells of one row meet edge to edge whatever the column widths: no gap of a pixel (a white line through filled cells and the text running over them).
checks['cells-meet'] = async () => {
    const { lines, check } = makeLog();
    const { bytesOf } = await import('/component.js');
    const gaps = (host, rows) => {
        const found = [];
        for (const r of rows) {
            const nodes = [...host.querySelectorAll(`.spreadsheet-cell[data-row="${r}"]`)].map((n) => ({ c: Number(n.getAttribute('data-col')), box: n.getBoundingClientRect() })).filter((n) => n.box.width > 0).sort((a, b) => a.c - b.c);
            for (let i = 1; i < nodes.length; i++) if (nodes[i].c === nodes[i - 1].c + 1 && Math.abs(nodes[i].box.left - nodes[i - 1].box.right) > 0.01) found.push([r, nodes[i - 1].c, nodes[i].box.left - nodes[i - 1].box.right]);
        }
        return found;
    };
    {
        const { grid, host, close } = mountGrid();
        try {
            await grid.ImportXlsx(await bytesOf('tests/real-world/milk-street-office.xlsx')); await frame(); await frame();
            const b5 = host.querySelector('.spreadsheet-cell[data-row="4"][data-col="1"]'), c5 = host.querySelector('.spreadsheet-cell[data-row="4"][data-col="2"]');
            check('milk-street Annual B5 runs over into C5, both filled', [b5.textContent, getComputedStyle(b5).backgroundColor, getComputedStyle(c5).backgroundColor, getComputedStyle(b5).overflow], ['Acquisition Overview:', 'rgb(31, 73, 125)', 'rgb(31, 73, 125)', 'visible']);
            check('... no pixel between B5 and C5', c5.getBoundingClientRect().left - b5.getBoundingClientRect().right, 0);
            check('... nor between any two cells of rows 1 to 30', gaps(host, [...Array(30).keys()]), []);
        } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
        finally { close(); }
    }
    {
        const { grid, host, close } = mountGrid();
        try {
            for (let c = 0; c < 12; c++) { grid.SetColumnWidth(c, 17.4 + c * 3.3); for (let r = 0; r < 12; r++) grid.SetCell(r, c, 'x'); }
            for (let r = 0; r < 12; r++) grid.SetRowHeight(r, 18.6 + r * 0.7);
            await frame();
            check('columns of fractional widths: the cells meet', gaps(host, [...Array(12).keys()]), []);
            const down = [];
            for (let c = 0; c < 3; c++) for (let r = 1; r < 12; r++) { const a = host.querySelector(`.spreadsheet-cell[data-row="${r - 1}"][data-col="${c}"]`).getBoundingClientRect(), b = host.querySelector(`.spreadsheet-cell[data-row="${r}"][data-col="${c}"]`).getBoundingClientRect(); if (Math.abs(b.top - a.bottom) > 0.01) down.push([r, c, b.top - a.bottom]); }
            check('rows of fractional heights: the cells meet', down, []);
            const header = (c) => host.querySelector(`.spreadsheet-col-header[data-col="${c}"]`).getBoundingClientRect(), cell = (c) => host.querySelector(`.spreadsheet-cell[data-row="0"][data-col="${c}"]`).getBoundingClientRect();
            check('the column letters stand over their columns', [...Array(12).keys()].filter((c) => Math.abs(header(c).left - cell(c).left) > 0.01 || Math.abs(header(c).right - cell(c).right) > 1.01), []);
        } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
        finally { close(); }
    }
    return lines;
};

// A print area of several areas through the host API and a save.
checks['print-areas'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    try {
        for (let r = 0; r < 30; r++) for (let c = 0; c < 8; c++) grid.SetCell(r, c, String(r * 8 + c));
        grid.SetPageSetup({ printArea: 'A1:C10,E1:G10' });
        check('set as a comma list', [grid.GetPageSetup().printArea, grid.GetPageSetup().printAreas], ['A1:C10,E1:G10', ['A1:C10', 'E1:G10']]);
        grid.SetPageSetup(grid.GetPageSetup());
        check('what GetPageSetup gives can be given back', grid.GetPageSetup().printAreas, ['A1:C10', 'E1:G10']);
        grid.SetPageSetup({ printAreas: ['B2:C3', 'E5:F6', 'A20:B22'] });
        check('set as a list', grid.GetPageSetup().printArea, 'B2:C3,E5:F6,A20:B22');
        grid.Undo();
        check('undo', grid.GetPageSetup().printAreas, ['A1:C10', 'E1:G10']);
        const bytes = await grid.ExportXlsx();
        const book = await (await window.JSZip.loadAsync(bytes)).files['xl/workbook.xml'].async('string');
        check('written: each area with its sheet', /_xlnm\.Print_Area"[^>]*>Sheet1!\$A\$1:\$C\$10,Sheet1!\$E\$1:\$G\$10</.test(book.replace(/&apos;|'/g, '')), true);
        const again = await reopened(grid);
        try { check('read back', again.grid.GetPageSetup().printAreas, ['A1:C10', 'E1:G10']); } finally { again.close(); }
        grid.InsertRows(0, 2);
        check('inserted rows move every area', grid.GetPageSetup().printArea, 'A3:C12,E3:G12');
        grid.SetPageSetup({ printArea: 'A1:B2' });
        check('one area again', [grid.GetPageSetup().printArea, grid.GetPageSetup().printAreas], ['A1:B2', ['A1:B2']]);
        grid.SetPageSetup({ printAreas: [] });
        check('none', [grid.GetPageSetup().printArea, grid.GetPageSetup().printAreas], [null, []]);
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// A viewer: filter buttons are drawn for a host that listens, and filtering there is how the sheet is looked at.
checks['viewer-filter'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        ['Name', 'Ann', 'Bob', 'Cy', 'Ann'].forEach((text, row) => { grid.SetCell(row, 0, text); grid.SetCell(row, 1, row === 0 ? 'N' : String(row)); });
        grid.SetAutoFilter('A1:B5'); await frame();
        const bytes = await grid.ExportXlsx();
        const viewer = mountGrid();
        try {
            await viewer.grid.ImportXlsx(bytes); viewer.grid.SetReadOnly(true); await frame();
            const buttons = () => [...viewer.host.querySelectorAll('.spreadsheet-filter-button')].filter((n) => n.getClientRects().length > 0);
            check('a viewer no host listens to: no filter buttons', [buttons().length, viewer.grid.GetFilterButtonRects().length], [0, 0]);
            const pressed = [];
            const off = viewer.grid.OnFilterButton((button) => pressed.push(button.row));
            const shownRows = () => [...viewer.host.querySelectorAll('.spreadsheet-row-header')].filter((n) => n.getBoundingClientRect().height > 0).map((n) => Number(n.getAttribute('data-row'))).filter((r) => r < 5).sort((a, b) => a - b);
            await frame();
            check('a host listens: the buttons are drawn', [buttons().length, viewer.grid.GetFilterButtonRects().length], [2, 2]);
            const box = buttons()[0].getBoundingClientRect();
            await realClick(box.left + box.width / 2, box.top + box.height / 2);
            check('a press reaches the host', pressed, [0]);
            let told = 0;
            const offChange = viewer.grid.OnChange(() => { told++; });
            const revision = viewer.grid.GetRevision(), canUndo = viewer.grid.CanUndo();
            viewer.grid.SetFilter(0, { values: ['Ann'] }); await frame();
            check('SetFilter in the viewer hides the rows', shownRows(), [0, 1, 4]);
            check('... as view state: no revision, no change event, no undo step', [viewer.grid.GetRevision() === revision, told, viewer.grid.CanUndo() === canUndo], [true, 0, true]);
            const out = await viewer.grid.ExportXlsx();
            check('... and an export is still the file that was opened', [out.length === bytes.length, out.every((value, index) => value === bytes[index])], [true, true]);
            viewer.grid.ClearFilter(); await frame();
            check('ClearFilter shows them again, still no revision', [shownRows(), viewer.grid.GetRevision() === revision, told], [[0, 1, 2, 3, 4], true, 0]);
            if (typeof offChange === 'function') offChange();
            off(); await frame();
            check('the host stops listening: the buttons go', buttons().length, 0);
        } finally { viewer.close(); }
        // In the editor the same call is an edit.
        const before = grid.GetRevision();
        grid.SetFilter(0, { values: ['Ann'] });
        check('in the editor SetFilter is an edit', [grid.GetRevision() !== before, grid.CanUndo()], [true, true]);
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// The end of the sheet, as in Excel. What can be scrolled through is the used range and room past
// it: scrolling to the end adds nothing. The caret going further does: Ctrl+Down / Ctrl+Right with
// nothing ahead go to row 1 048 576 / column XFD, Ctrl+Up / Ctrl+Left back to the last cell that
// holds something, and the sheet is then as long as it was. A drag of the selection or of the fill
// handle held at the edge of the grid scrolls the sheet under it. A fill and a paste past the last
// row make the sheet that large.
checks['grid-end-of-sheet'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const { scroller } = parts(host);
        const size = () => [grid.GetModel().RowCount, grid.GetModel().ColCount];
        const [rows, cols] = size();
        for (let row = 0; row < 40; row++) grid.SetCell(row, 0, String(row + 1));
        await frame();
        check('entries near the top leave the sheet as large as it is', size(), [rows, cols]);
        scroller.scrollTop = 4000; await frame();
        scroller.scrollTop = 8000; await frame();
        check('scrolled about in the middle: no more rows', size(), [rows, cols]);
        const toEnd = async () => { scroller.scrollTop = scroller.scrollHeight; scroller.dispatchEvent(new WheelEvent('wheel', { deltaY: 300, bubbles: true })); await frame(); };
        await toEnd();
        const height = scroller.scrollHeight;
        for (let i = 0; i < 12; i++) await toEnd();
        check('scrolled to the end, again and again: the sheet is as long as it was', [size(), scroller.scrollHeight], [[rows, cols], height]);
        scroller.scrollLeft = scroller.scrollWidth; await frame(); scroller.scrollLeft = scroller.scrollWidth; await frame();
        check('... and as wide', size(), [rows, cols]);
        scroller.scrollLeft = 0; scroller.scrollTop = 0; await frame();

        // Ctrl+arrows (the browser's own keys, on a cell that was clicked).
        const active = () => [grid.GetActiveCell().Row, grid.GetActiveCell().Col];
        const clickCell = async (row, col) => { grid.SelectCell(row, col); await frame(); const box = cellBox(host, row, col); await window.realMouse('click', box.left + box.width / 2, box.top + box.height / 2); await frame(); };
        await clickCell(2, 0);
        await window.realKey('Control+ArrowDown'); await frame();
        check('Ctrl+Down in a column of data: its last cell', active(), [39, 0]);
        let start = performance.now();
        await window.realKey('Control+ArrowDown'); await frame();
        const down = performance.now() - start;
        check('Ctrl+Down with nothing below: row 1 048 576', [active(), grid.GetSelectionRef(), size()[0]], [[1048575, 0], 'A1048576', 1048576]);
        const last = cellBox(host, 1048575, 0);
        check('... which is in view, drawn', [last !== null && last.bottom <= scroller.getBoundingClientRect().bottom + 1 && last.top >= scroller.getBoundingClientRect().top, host.querySelectorAll('.spreadsheet-cell').length < 3000], [true, true]);
        const layout = await import('/spreadsheet/Layout.js'), metrics = layout.metricsOf(grid.GetModel());
        check('... and no row of the sheet has an entry of its own in memory', [layout.AxisOffsets__get_Known(metrics.RowPrefix).length < 200, layout.AxisOffsets__get_Length(metrics.RowPrefix)], [true, 1048577]);
        start = performance.now();
        grid.EnterText(1048575, 0, 'the last row'); await frame();
        note(`Ctrl+Down to row 1 048 576: ${down.toFixed(1)} ms; an entry there and its frame: ${(performance.now() - start).toFixed(1)} ms`);
        grid.Undo(); await frame();
        await window.realKey('Control+ArrowDown'); await frame();
        check('Ctrl+Down on the last row stays there', active(), [1048575, 0]);
        start = performance.now();
        await window.realKey('Control+ArrowUp'); await frame();
        check('Ctrl+Up from there: the last cell that holds something', active(), [39, 0]);
        await frame();
        check('... and the sheet is as long as it was', [size(), scroller.scrollHeight], [[rows, cols], height]);
        note(`Ctrl+Up back: ${(performance.now() - start).toFixed(1)} ms`);
        await window.realKey('Control+ArrowRight'); await frame();
        check('Ctrl+Right with nothing to the right: column XFD', [active(), grid.GetSelectionRef(), size()[1]], [[39, 16383], 'XFD40', 16384]);
        await window.realKey('Control+ArrowLeft'); await frame(); await frame();
        check('Ctrl+Left back: the cell with data, the sheet as wide as it was', [active(), size()], [[39, 0], [rows, cols]]);
        await clickCell(0, 3);
        await window.realKey('Control+ArrowUp'); await window.realKey('Control+ArrowLeft'); await frame();
        check('Ctrl+Up in the first row stays, Ctrl+Left goes to the cell with data', active(), [0, 0]);
        // Shift+Ctrl+Down selects to the end; a cell written there keeps its row.
        await clickCell(1, 2);
        await window.realKey('Control+Shift+ArrowDown'); await frame();
        check('Shift+Ctrl+Down in an empty column: selected to the last row', grid.GetSelectionRef(), 'C2:C1048576');
        grid.SelectCell(0, 0); await frame(); await frame();
        check('... and back at the top the sheet is as long as it was', size(), [rows, cols]);
        grid.SelectCell(5000, 1); await frame();
        grid.EnterText(5000, 1, 'far'); grid.SelectCell(0, 0); await frame(); await frame();
        check('a cell entered far down stays a row of the sheet', [size()[0], grid.GetCell(5000, 1).value], [5101, 'far']);
        grid.Undo(); grid.SelectCell(0, 0); await frame(); await frame();

        // A drag of the selection held below the grid scrolls the sheet under the pointer.
        const pause = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
        const rect = scroller.getBoundingClientRect();
        let box = cellBox(host, 3, 1);
        await window.realMouse('move', box.left + 10, box.top + 5);
        await window.realMouse('down');
        await window.realMouse('move', box.left + 10, rect.bottom - 40);
        await frame();
        const inView = Number(/:B(\d+)$/.exec(grid.GetSelectionRef())?.[1]);
        await window.realMouse('move', box.left + 10, rect.bottom + 60);
        await pause(700);
        const scrolled = scroller.scrollTop, reached = Number(/:B(\d+)$/.exec(grid.GetSelectionRef())?.[1]);
        await window.realMouse('up');
        check('a selection dragged below the grid: the sheet scrolls, the selection goes on', [scrolled > 200, reached > inView + 8, grid.GetSelectionRef().startsWith('B4:B')], [true, true, true]);
        const afterUp = scroller.scrollTop;
        await pause(200);
        check('... until the button is released', scroller.scrollTop, afterUp);
        // ... and held there it goes past the last row the sheet had.
        grid.SelectCell(rows - 30, 1); await frame(); await frame();
        box = cellBox(host, rows - 30, 1);
        await window.realMouse('move', box.left + 10, box.top + 5);
        await window.realMouse('down');
        await window.realMouse('move', box.left + 10, rect.bottom + 250);
        await pause(1500);
        await window.realMouse('up');
        const draggedTo = Number(/:B(\d+)$/.exec(grid.GetSelectionRef())?.[1]);
        check('held there: the selection goes past the last row the sheet had, which grows with it', [draggedTo > rows + 20, size()[0] > draggedTo, grid.GetSelectionRef().startsWith(`B${rows - 29}:B`)], [true, true, true]);
        // The fill handle, the same: a fill past what the grid showed in one drag.
        grid.SelectCell(0, 0); await frame(); await frame();
        grid.SetCell(rows - 31, 4, '1'); grid.SetCell(rows - 30, 4, '2'); grid.SelectRange(rows - 31, 4, rows - 30, 4); grid.ScrollToCell(rows - 20, 4); await frame(); await frame();
        const handle = host.querySelector('.spreadsheet-fill-handle').getBoundingClientRect();
        await window.realMouse('move', handle.left + handle.width / 2, handle.top + handle.height / 2);
        await window.realMouse('down');
        await window.realMouse('move', handle.left + handle.width / 2, rect.bottom + 250);
        await pause(1500);
        await window.realMouse('up'); await frame();
        const filledTo = Number(/:E(\d+)$/.exec(grid.GetSelectionRef())?.[1]);
        check('the fill handle dragged below the grid and held: a series past the last row the sheet had', [filledTo > rows + 20, grid.GetCell(filledTo - 1, 4).value, grid.GetCell(rows + 10, 4).value, size()[0] > filledTo], [true, filledTo - (rows - 31), 42, true]);
        note(`the selection was dragged to row ${draggedTo}, the fill to row ${filledTo} (the sheet had ${rows} rows)`);
        grid.Undo(); grid.SelectCell(0, 0); await frame(); await frame();
        scroller.scrollLeft = 0; scroller.scrollTop = 0; await frame();

        // A fill down past the last row.
        const lastRow = size()[0] - 1;
        grid.SetCell(lastRow, 1, '7'); grid.SelectCell(lastRow, 1); await frame();
        const before = size()[0];
        grid.AutoFill(before + 20, 1); await frame();
        check('a fill past the last row: the sheet holds the cells', [size()[0] > before + 20, typeof grid.GetCell(before + 20, 1).value, grid.GetCell(lastRow, 1).value], [true, 'number', 7]);
        // A paste past the last row and the last column.
        const [rowsNow, colsNow] = size();
        grid.SelectCell(rowsNow - 1, colsNow - 1); await frame();
        grid.PasteText('a\tb\tc\n1\t2\t3\n4\t5\t6', ''); await frame();
        check('a paste past the last row and column: the sheet holds the cells', [size()[0] > rowsNow + 1, size()[1] > colsNow + 1, grid.GetCell(rowsNow + 1, colsNow + 1).value], [true, true, 6]);
        // A viewer's sheet is as long as it is, and its keys stay in it.
        grid.SetReadOnly(true); scroller.scrollTop = 0; await frame();
        const fixed = size();
        await toEnd(); await toEnd();
        check('read-only: scrolling to the end adds nothing', size(), fixed);
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// The box of a comment stays inside the grid: beside a cell at the right edge it is on the cell's
// left, and for a cell in the last row in view it is moved up.
checks['comment-box-in-view'] = async () => {
    const { lines, check, note } = makeLog();
    const { bytesOf } = await import('/component.js');
    const { grid, host, close } = mountGrid();
    try {
        const { scroller } = parts(host);
        const box = () => { const el = host.querySelector('.spreadsheet-comment-box'); return el && el.style.display !== 'none' ? el : null; };
        const nodeOf = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
        const hover = (row, col) => {
            const rect = nodeOf(row, col).getBoundingClientRect();
            const x = rect.left + Math.min(rect.width / 2, 20), y = rect.top + rect.height / 2;
            (document.elementFromPoint(x, y) ?? nodeOf(row, col)).dispatchEvent(new MouseEvent('mousemove', { bubbles: true, cancelable: true, clientX: x, clientY: y, buttons: 0 }));
        };
        const leave = () => host.querySelector('.spreadsheet-content').dispatchEvent(new MouseEvent('mouseleave', { bubbles: false }));
        // The view of the scroller without its scroll bars.
        const view = () => { const rect = scroller.getBoundingClientRect(); return { left: rect.left, top: rect.top, right: rect.left + scroller.clientWidth, bottom: rect.top + scroller.clientHeight }; };
        const inside = (el) => { const rect = el.getBoundingClientRect(), v = view(); return [rect.left >= v.left - 1, rect.top >= v.top - 1, rect.right <= v.right + 1, rect.bottom <= v.bottom + 1]; };
        // The cell as the last row in view.
        const asLastRow = async (row, col) => {
            grid.ScrollToCell(row, col); await frame();
            const rect = nodeOf(row, col).getBoundingClientRect();
            scroller.scrollTop += rect.bottom - view().bottom + 1; await frame();
        };
        await grid.ImportXlsx((await bytesOf('tests/real-world/jumeirah-hotel.xlsx')).slice());
        const sheet = grid.GetSheetNames().findIndex((_, index) => { grid.SwitchSheet(index); return grid.GetComment(48, 4) !== null; });
        check('the workbook has a comment on E49', sheet >= 0, true);
        await frame();
        await asLastRow(48, 4);
        hover(48, 4); await frame();
        const shown = box();
        check('E49 as the last row in view: its box is shown', shown !== null && shown.getBoundingClientRect().height > 20, true);
        if (shown) {
            note(`the box: ${Math.round(shown.getBoundingClientRect().height)} px high, its cell ${Math.round(nodeOf(48, 4).getBoundingClientRect().height)} px`);
            check('... inside the grid, on all sides', inside(shown), [true, true, true, true]);
        }
        leave(); await frame();
        // A cell further up: the box is beside the cell's top, as before.
        scroller.scrollTop += 200; await frame();
        hover(48, 4); await frame();
        const beside = box();
        check('with room below: the box starts at the top of its cell', beside !== null && Math.abs(beside.getBoundingClientRect().top - nodeOf(48, 4).getBoundingClientRect().top) <= 2, true);
        leave(); await frame();
        // A note on the last column in view, in a new workbook: the box is on the left of the cell.
        const { grid: other, host: otherHost, close: closeOther } = mountGrid();
        try {
            const scrollerOther = parts(otherHost).scroller;
            const cells = [...otherHost.querySelectorAll('.spreadsheet-cell[data-row="3"]')].filter((node) => node.style.display !== 'none');
            const edge = scrollerOther.getBoundingClientRect().left + scrollerOther.clientWidth;
            const lastCol = Math.max(...cells.filter((node) => node.getBoundingClientRect().right <= edge + 1).map((node) => Number(node.getAttribute('data-col'))));
            other.SetCommentRuns(3, lastCol, [{ text: 'A note of some length, long enough to need the width a box of a note has, and then some more words.' }], 'Ann');
            await frame();
            const cell = otherHost.querySelector(`.spreadsheet-cell[data-row="3"][data-col="${lastCol}"]`).getBoundingClientRect();
            otherHost.querySelector(`.spreadsheet-cell[data-row="3"][data-col="${lastCol}"]`).dispatchEvent(new MouseEvent('mousemove', { bubbles: true, cancelable: true, clientX: cell.left + 10, clientY: cell.top + cell.height / 2, buttons: 0 }));
            await frame();
            const el = otherHost.querySelector('.spreadsheet-comment-box');
            const rect = el.getBoundingClientRect();
            check('a note on the last column in view: its box is shown, inside the grid', [el.style.display !== 'none', rect.right <= edge + 1, rect.left >= scrollerOther.getBoundingClientRect().left], [true, true, true]);
            check('... on the left of its cell', rect.right <= cell.left + 1, true);
        } finally { closeOther(); }
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// The label of a chart, shape or slicer the grid does not draw is one line inside its frame, cut with an ellipsis.
checks['object-label-fits'] = async () => {
    const { lines, check, note } = makeLog();
    const { bytesOf } = await import('/component.js');
    const { grid, host, close } = mountGrid();
    try {
        let frames = 0, cut = 0, drawnObjects = 0;
        for (const file of ['tests/real-world/chart-zoo.xlsx', 'tests/real-world/dashboard-sales-analysis.xlsx']) {
            await grid.ImportXlsx((await bytesOf(file)).slice());
            for (let index = 0; index < grid.GetSheetNames().length; index++) {
                grid.SwitchSheet(index); await frame();
                await seenAll(host);
                const wrong = [];
                for (const node of host.querySelectorAll('.spreadsheet-object')) {
                    frames++;
                    const label = node.querySelector('.spreadsheet-object-label');
                    // (An object that is drawn has no label: its picture is in its frame.)
                    if (!label && node.classList.contains('spreadsheet-object-drawn') && node.querySelector('svg')) { drawnObjects++; continue; }
                    if (!label) { wrong.push(`${node.title}: no label`); continue; }
                    if (getComputedStyle(label).display === 'none') continue;
                    const outer = node.getBoundingClientRect(), inner = label.getBoundingClientRect(), style = getComputedStyle(label);
                    if (label.scrollWidth > label.clientWidth) cut++;
                    if (inner.left < outer.left - 0.5 || inner.right > outer.right + 0.5 || inner.top < outer.top - 0.5 || inner.bottom > outer.bottom + 0.5) wrong.push(`${node.title}: the label reaches outside its frame`);
                    if (style.whiteSpace !== 'nowrap' || style.textOverflow !== 'ellipsis' || inner.height > parseFloat(style.fontSize) * 1.6) wrong.push(`${node.title}: more than one line, or no ellipsis`);
                    if (label.textContent !== node.title) wrong.push(`${node.title}: the label says "${label.textContent}"`);
                }
                check(`${file.split('/').pop()} / ${grid.GetSheetNames()[index]}: every label is one line inside its frame`, wrong, []);
            }
        }
        note(`${frames} frames, ${drawnObjects} of them drawn, ${cut} labels cut with an ellipsis`);
        // (Since slicers are drawn these workbooks have no labelled frame: what is checked of a label above holds for any.)
        check('the workbooks have frames, drawn ones among them', [frames > 5, drawnObjects > 5], [true, true]);
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// A note is a change of the sheet's comments part and of the box in its VML drawing: the sheet's
// own part goes back as the file's, byte for byte (a sheet that had no note gets a
// <legacyDrawing>), the cell keeps the cellXfs entry it had and the styles are the file's; a note
// of the file that is edited keeps shapeId, xr:uid and what its runs say beyond what a sheet holds.
checks['note-leaves-sheet-part'] = async () => {
    const { lines, check } = makeLog();
    const { bytesOf } = await import('/component.js');
    const textOf = async (zip, path) => zip.files[path] ? zip.files[path].async('string') : null;
    const { grid, close } = mountGrid();
    try {
        const original = await bytesOf('tests/real-world/milk-street-office.xlsx');
        const zip0 = await window.JSZip.loadAsync(original);
        await grid.ImportXlsx(original.slice());
        grid.SetComment(2, 1, 'a note', 'me');
        grid.SetComment(112, 4, 'another text', 'BIWS');
        const zip1 = await window.JSZip.loadAsync(await grid.ExportXlsx());
        check('milk-street, a note added in B3 and the note of E113 edited: sheet, styles, relationships and content types are the file\'s',
            await Promise.all(['xl/worksheets/sheet1.xml', 'xl/styles.xml', 'xl/worksheets/_rels/sheet1.xml.rels', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels', 'xl/worksheets/sheet2.xml'].map(async (path) => (await textOf(zip1, path)) === (await textOf(zip0, path)))),
            [true, true, true, true, true, true]);
        const comments = await textOf(zip1, 'xl/comments1.xml');
        const edited = /<comment ref="E113"[^>]*>[\s\S]*?<\/comment>/.exec(comments)?.[0] ?? '';
        check('the edited note keeps shapeId, xr:uid and its run\'s <color indexed> and <family>',
            [/shapeId="0"/.test(edited), /xr:uid="\{00000000-0006-0000-0000-000001000000\}"/.test(edited), edited.includes('<r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t xml:space="preserve">another text</t></r>')], [true, true, true]);
        check('the new note is in the part, by its author', [/<comment ref="B3" authorId="1">/.test(comments), comments.includes('<author>BIWS</author><author>me</author>')], [true, true]);
        check('... with a box in the file\'s drawing', ((await textOf(zip1, 'xl/drawings/vmlDrawing1.vml')).match(/ObjectType="Note"/g) ?? []).length, ((await textOf(zip0, 'xl/drawings/vmlDrawing1.vml')).match(/ObjectType="Note"/g) ?? []).length + 1);
        // Undone, the workbook is the file again.
        grid.Undo(); grid.Undo();
        check('undone: the opened bytes', (await grid.ExportXlsx()).length === original.length && !grid.IsModified(), true);
        // A sheet without a note (it has a VML drawing for other things): G5 keeps s="31".
        const dashboard = await bytesOf('tests/real-world/dashboard-ecommerce-sales.xlsx');
        const dash0 = await window.JSZip.loadAsync(dashboard);
        await grid.ImportXlsx(dashboard.slice());
        grid.SwitchSheet(2);
        grid.SetComment(4, 6, 'a first note', 'me');
        const dash1 = await window.JSZip.loadAsync(await grid.ExportXlsx());
        check('dashboard-ecommerce, a first note in G5 of the third sheet: the sheet part and the styles are the file\'s',
            [(await textOf(dash1, 'xl/worksheets/sheet3.xml')).replace(/ tabSelected="[01]"/g, '') === (await textOf(dash0, 'xl/worksheets/sheet3.xml')).replace(/ tabSelected="[01]"/g, ''), /<c r="G5" s="31"\/>/.test(await textOf(dash1, 'xl/worksheets/sheet3.xml')), (await textOf(dash1, 'xl/styles.xml')) === (await textOf(dash0, 'xl/styles.xml'))], [true, true, true]);
        const rels = await textOf(dash1, 'xl/worksheets/_rels/sheet3.xml.rels'), rels0 = await textOf(dash0, 'xl/worksheets/_rels/sheet3.xml.rels');
        check('... its relationships the file\'s with the comments part behind them', [rels.startsWith(rels0.slice(0, rels0.lastIndexOf('</Relationships>'))), (rels.match(/<Relationship /g) ?? []).length - (rels0.match(/<Relationship /g) ?? []).length, /relationships\/comments" Target="\.\.\/comments\d+\.xml"/.test(rels)], [true, 1, true]);
        const again = mountGrid();
        try {
            await again.grid.ImportXlsx(await grid.ExportXlsx());
            again.grid.SwitchSheet(2);
            check('... and it reads back', again.grid.GetComment(4, 6)?.text, 'a first note');
        } finally { again.close(); }
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// The page break preview: SetViewMode('pageBreak') outlines every page the sheet prints on over
// its cells (the pages of GetPrintPages), numbered; an edge is solid at a manual break and at the
// ends of the print area, dashed where the paper ends a page. Breaks are added, moved
// (MovePageBreak, one undo step) and removed through the API, and saved as rowBreaks / colBreaks.
checks['page-break-preview'] = async () => {
    const { lines, check } = makeLog();
    const host = container('position:absolute;left:0;top:0;width:1100px;height:800px');
    const grid = createSpreadsheet(host);
    grid.Mount();
    const frame = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    const pages = () => [...host.querySelectorAll('.spreadsheet-page')];
    const cellBox = (row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`)?.getBoundingClientRect();
    try {
        for (let row = 0; row < 120; row++) for (let col = 0; col < 4; col++) grid.SetCell(row, col, String(row * 10 + col));
        await frame();
        check('the normal view: no page is outlined', [grid.GetViewMode(), pages().length, 'viewMode' in grid.GetViewOptions()], ['normal', 0, false]);
        const revision = grid.GetRevision();
        grid.SetViewMode('pageBreak'); await frame();
        const printed = grid.GetPrintPages().pages;
        check('the preview: more than one page, the first ones in view are outlined', [grid.GetViewMode(), printed.length > 1, pages().length >= 1, pages().length <= printed.length], ['pageBreak', true, true, true]);
        const first = pages()[0];
        check('a page says its number and its cells, as GetPrintPages has them', [first.dataset.page, first.dataset.range, first.querySelector('.spreadsheet-page-label').textContent], ['1', printed[0].range, 'Page 1']);
        const box = first.getBoundingClientRect(), topLeft = cellBox(printed[0].firstRow, printed[0].firstCol), right = cellBox(printed[0].firstRow, printed[0].lastCol);
        check('its frame begins at its first cell and is as wide as its columns', [Math.abs(box.left - topLeft.left) < 1.5, Math.abs(box.top - topLeft.top) < 1.5, Math.abs(box.right - right.right) < 1.5], [true, true, true]);
        check('the print area\'s edge is solid, and the frame takes no click', [first.dataset.breakTop, getComputedStyle(first).borderTopStyle, getComputedStyle(first).pointerEvents], ['area', 'solid', 'none']);
        check('switching the view is no change of the workbook', [grid.GetRevision(), grid.IsModified()], [revision, true]);
        // The second page begins where the paper ends the first: a dashed edge.
        host.querySelector('.spreadsheet-container').scrollTop = Math.max(0, cellBox(0, 0) ? 20 * (printed[1].firstRow - 5) : 0); await frame();
        const second = pages().find((node) => node.dataset.page === '2');
        check('the second page begins at an automatic break: dashed', [second?.dataset.firstRow, second?.dataset.breakTop, second ? getComputedStyle(second).borderTopStyle : null], [String(printed[1].firstRow), 'automatic', 'dashed']);
        // A manual break: the pages follow, its edge is solid.
        check('a break before row 11', grid.AddPageBreak(10), true); await frame();
        host.querySelector('.spreadsheet-container').scrollTop = 0; await frame();
        const after = grid.GetPrintPages().pages;
        const manual = pages().find((node) => node.dataset.firstRow === '10');
        check('... the first page ends before it, the next begins at it with a solid edge', [after[0].lastRow, after[1].firstRow, manual?.dataset.page, manual?.dataset.breakTop, manual ? getComputedStyle(manual).borderTopStyle : null], [9, 10, '2', 'manual', 'solid']);
        // Dragged to row 21: one undo step.
        check('moved to before row 21', [grid.MovePageBreak('row', 10, 20), grid.GetPageBreaks().rows], [true, [20]]); await frame();
        check('... the frames follow', [pages().some((node) => node.dataset.firstRow === '10'), pages().find((node) => node.dataset.firstRow === '20')?.dataset.breakTop], [false, 'manual']);
        check('a break that is not there does not move; a place no break can be at throws', [grid.MovePageBreak('row', 10, 30), thrown(() => grid.MovePageBreak('row', 20, 0)) !== null, thrown(() => grid.MovePageBreak('diagonal', 20, 5)) !== null], [false, true, true]);
        grid.Undo(); await frame();
        check('undone: the break is before row 11 again', [grid.GetPageBreaks().rows, pages().find((node) => node.dataset.firstRow === '10')?.dataset.breakTop], [[10], 'manual']);
        grid.Redo(); await frame();
        // A column break, and what is saved.
        grid.AddPageBreak(null, 2); await frame();
        check('a break before column C: pages side by side', [grid.GetPageBreaks().columns, pages().some((node) => node.dataset.firstCol === '2' && node.dataset.breakLeft === 'manual')], [[2], true]);
        const zip = await window.JSZip.loadAsync(await grid.ExportXlsx());
        const sheet = await zip.files['xl/worksheets/sheet1.xml'].async('string');
        check('saved as rowBreaks and colBreaks', [/<rowBreaks count="1" manualBreakCount="1"><brk id="20" [^>]*man="1"\/><\/rowBreaks>/.test(sheet), /<colBreaks count="1" manualBreakCount="1"><brk id="2" [^>]*man="1"\/><\/colBreaks>/.test(sheet)], [true, true]);
        // Editing goes on as before; a row made taller moves the breaks the paper makes.
        const lastOfFirst = grid.GetPrintPages().pages[0].lastRow;
        grid.EnterText(3, 1, 'typed in the preview');
        check('a cell is typed into as in the normal view', grid.GetCell(3, 1).value, 'typed in the preview');
        grid.RemovePageBreak(20); grid.SetRowHeight(2, 400); await frame();
        check('a taller row: the first page ends earlier, and its frame with it', [grid.GetPrintPages().pages[0].lastRow < Math.max(lastOfFirst, 30), pages()[0].dataset.range, pages()[0].dataset.range === grid.GetPrintPages().pages[0].range], [true, grid.GetPrintPages().pages[0].range, true]);
        check('another mode throws', thrown(() => grid.SetViewMode('layout')) !== null, true);
        grid.SetViewMode('normal'); await frame();
        check('back in the normal view: no frame is left', [grid.GetViewMode(), pages().length], ['normal', 0]);
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { grid.Dispose(); host.remove(); }
    return lines;
};

// A link is a change of the list of links in the sheet's part and - an external one - of the
// part's relationships; the last note taken off a sheet takes its comments part and the drawing
// of the boxes with it. The part is otherwise the file's text: its columns (`bestFit`), its page
// setup (what it says of the printer), its views, its extension list.
checks['link-and-last-note-leave-sheet-part'] = async () => {
    const { lines, check } = makeLog();
    const { bytesOf } = await import('/component.js');
    const textOf = async (zip, path) => zip.files[path] ? zip.files[path].async('string') : null;
    const element = (name, xml) => (new RegExp(`<${name}(?=[\\s>/])[^>]*/>|<${name}(?=[\\s>])[^>]*>[\\s\\S]*?</${name}>`).exec(xml) ?? [''])[0];
    const { grid, close } = mountGrid();
    try {
        const original = await bytesOf('tests/real-world/milk-street-office.xlsx');
        const zip0 = await window.JSZip.loadAsync(original);
        const sheet0 = await textOf(zip0, 'xl/worksheets/sheet1.xml'), rels0 = await textOf(zip0, 'xl/worksheets/_rels/sheet1.xml.rels');
        await grid.ImportXlsx(original.slice());
        grid.SetHyperlink('B402', 'https://example.com/a?b=1&c=2', 'a link', 'a tip');
        const zip1 = await window.JSZip.loadAsync(await grid.ExportXlsx());
        const sheet1 = await textOf(zip1, 'xl/worksheets/sheet1.xml'), rels1 = await textOf(zip1, 'xl/worksheets/_rels/sheet1.xml.rels');
        const kept = ['sheetPr', 'sheetViews', 'sheetFormatPr', 'cols', 'mergeCells', 'dataValidations', 'pageMargins', 'pageSetup', 'rowBreaks', 'ignoredErrors', 'legacyDrawing', 'extLst'];
        check('milk-street, a link in B402: every other element of the part is the file\'s text', kept.filter((name) => element(name, sheet1) !== element(name, sheet0)), []);
        check('... the file\'s part says what a part made from the sheet would not', [/bestFit="1"/.test(element('cols', sheet0)) || /horizontalDpi/.test(element('pageSetup', sheet0)), sheet0.slice(0, sheet0.indexOf('<sheetData')).replace(element('dimension', sheet0), '') === sheet1.slice(0, sheet1.indexOf('<sheetData')).replace(element('dimension', sheet1), '')], [true, true]);
        const added = (rels1.match(/<Relationship [^>]*\/>/g) ?? []).filter((entry) => !rels0.includes(entry));
        check('... its relationships are the file\'s with the link\'s address behind them', [rels1.startsWith(rels0.slice(0, rels0.lastIndexOf('</Relationships>'))), added.length, /Target="https:\/\/example\.com\/a\?b=1&amp;c=2" TargetMode="External"/.test(added[0] ?? '')], [true, 1, true]);
        const id = /Id="([^"]*)"/.exec(added[0] ?? '')?.[1];
        check('... and the part names it', new RegExp(`<hyperlink ref="B402" r:id="${id}" tooltip="a tip" display="a link"/>`).test(element('hyperlinks', sheet1)), true);
        // (The link's text is typed into the cell, in the look of a link: a string and a format are added.)
        check('... styles and strings aside, every other part of the file is in the package as it was',
            (await Promise.all(Object.keys(zip0.files).filter((path) => !['xl/worksheets/sheet1.xml', 'xl/worksheets/_rels/sheet1.xml.rels', 'xl/styles.xml', 'xl/sharedStrings.xml', 'docProps/app.xml'].includes(path)).map(async (path) => (await textOf(zip1, path)) === (await textOf(zip0, path)) ? null : path))).filter(Boolean), []);
        const again = mountGrid();
        try {
            await again.grid.ImportXlsx(await grid.ExportXlsx());
            check('... it reads back', again.grid.GetHyperlink(401, 1)?.target, 'https://example.com/a?b=1&c=2');
        } finally { again.close(); }
        grid.Undo();
        check('undone: the opened bytes', (await grid.ExportXlsx()).length === original.length && !grid.IsModified(), true);
        // Every note of the sheet deleted.
        const notes = grid.GetComments().filter((note) => !note.thread);
        for (const note of notes) grid.DeleteComment(note.row, note.col);
        const zip2 = await window.JSZip.loadAsync(await grid.ExportXlsx());
        const sheet2 = await textOf(zip2, 'xl/worksheets/sheet1.xml'), rels2 = await textOf(zip2, 'xl/worksheets/_rels/sheet1.xml.rels');
        check(`its ${notes.length} notes deleted: the part is the file\'s without its <legacyDrawing>`, [notes.length > 0, sheet2 === sheet0.replace(element('legacyDrawing', sheet0), '')], [true, true]);
        check('... the comments part and the drawing of the boxes are gone, with their relationships and the content type',
            [zip2.files['xl/comments1.xml'] === undefined, zip2.files['xl/drawings/vmlDrawing1.vml'] === undefined, /comments|vmlDrawing/.test(rels2), /comments1\.xml/.test(await textOf(zip2, '[Content_Types].xml')), /comments1\.xml/.test(await textOf(zip0, '[Content_Types].xml'))],
            [true, true, false, false, true]);
        const last = mountGrid();
        try {
            await last.grid.ImportXlsx(await grid.ExportXlsx());
            check('... it reads back without a note, the cells as they were', [last.grid.GetComments().length, last.grid.GetCell(11, 4).value === grid.GetCell(11, 4).value], [0, true]);
        } finally { last.close(); }
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// A formula entered into a General cell shows its result in the format of what it reads, as in
// Excel: `=A1+1` under a date is a date, a sum of amounts an amount. The format is given when the
// formula is entered and is the cell's own from then on; a cell that has a format keeps it.
checks['formula-result-format'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    try {
        const at = (row, col) => { const cell = grid.GetCell(row, col); return [cell.text, cell.format.numberFormat ?? null]; };
        grid.EnterText(0, 0, '1/15/2026');
        const dateFormat = grid.GetCell(0, 0).format.numberFormat;
        check('a typed date is a date cell', [typeof dateFormat, grid.GetCell(0, 0).text], ['string', '1/15/2026']);
        grid.EnterText(1, 0, '=A1+1');
        check('=A1+1 under it shows a date, in the format of A1', at(1, 0), ['1/16/2026', dateFormat]);
        grid.EnterText(2, 0, '=A2-A1');
        check('the difference of two dates is a number', at(2, 0), ['1', null]);
        grid.EnterText(0, 1, '12.5%'); grid.EnterText(1, 1, '=B1*2');
        check('=B1*2 on a percentage is a percentage', [at(1, 1)[0], at(1, 1)[1] === grid.GetCell(0, 1).format.numberFormat], ['25.00%', true]);
        grid.EnterText(0, 2, '$1,200.50'); grid.EnterText(1, 2, '$300'); grid.SetCell(2, 2, '=SUM(C1:C2)');
        // (Excel's format for a typed amount keeps room for a negative's parenthesis: "$1,500.50 ".)
        check('a sum of amounts is an amount (SetCell too)', [at(2, 2)[0], at(2, 2)[1] === grid.GetCell(0, 2).format.numberFormat], ['$1,500.50 ', true]);
        grid.EnterText(3, 2, '=C1*B1');
        check('an amount times a rate is an amount', at(3, 2)[1] === grid.GetCell(0, 2).format.numberFormat, true);
        grid.EnterText(0, 3, '=DATE(2026,3,31)'); grid.EnterText(1, 3, '=COUNT(C1:C2)');
        check('DATE() is a date, COUNT() a number', [at(0, 3), at(1, 3)], [['3/31/2026', 'm/d/yyyy'], ['2', null]]);
        // Only a General cell takes a format, and only when the formula is entered.
        grid.SelectCell(4, 0); grid.SetNumberFormat('0.00'); grid.EnterText(4, 0, '=A1+1');
        check('a cell with a format of its own keeps it', at(4, 0)[1], '0.00');
        grid.EnterText(5, 0, '=C1'); grid.SelectCell(0, 2); grid.SetNumberFormat('0.0');
        check('the format is the cell\'s own: it does not follow the cell it came from', at(5, 0)[1] === '0.0', false);
        // One undo step: the formula and its format go together.
        grid.EnterText(6, 0, '=A1+7'); grid.Undo();
        check('undo takes the formula and its format', [grid.GetCell(6, 0).text, grid.GetCell(6, 0).format.numberFormat ?? null], ['', null]);
        // Across sheets, and written to a file.
        grid.AddSheet('Other'); grid.SwitchSheet(1); grid.EnterText(0, 0, '=Sheet1!A1+30');
        check('a date of another sheet', at(0, 0), ['2/14/2026', dateFormat]);
        const again = mountGrid();
        try {
            await again.grid.ImportXlsx(await grid.ExportXlsx());
            check('saved and opened: the cells are date cells', [again.grid.GetRange('Sheet1!A2')[0][0].text, again.grid.GetRange('Other!A1')[0][0].text], ['1/16/2026', '2/14/2026']);
        } finally { again.close(); }
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// A row inserted above a sheet's charts and shapes: the drawing part keeps its name and its
// text but for where the anchors are; its relationships, the sheet's relationships, the content
// types and the workbook's relationships are the file's (less the calculation chain, which goes).
checks['row-insert-keeps-drawing-part'] = async () => {
    const { lines, check } = makeLog();
    const { bytesOf } = await import('/component.js');
    const { grid, close } = mountGrid();
    try {
        for (const [file, sheetIndex, sheetPart, drawingPart] of [['tests/real-world/dashboard-ecommerce-sales.xlsx', 2, 'xl/worksheets/sheet3.xml', 'xl/drawings/drawing2.xml'], ['tests/real-world/chart-zoo.xlsx', 0, 'xl/worksheets/sheet1.xml', 'xl/drawings/drawing1.xml']]) {
            const short = file.split('/').pop();
            const original = await bytesOf(file);
            const zip0 = await window.JSZip.loadAsync(original);
            await grid.ImportXlsx(original.slice());
            grid.SwitchSheet(sheetIndex);
            const before = grid.GetDrawings().map((drawing) => drawing.anchor?.fromRow ?? drawing.fromRow ?? null);
            grid.InsertRows(0, 1);
            const zip1 = await window.JSZip.loadAsync(await grid.ExportXlsx());
            const text = async (zip, path) => zip.files[path] ? zip.files[path].async('string') : null;
            const names0 = Object.keys(zip0.files).filter((name) => !name.endsWith('/')), names1 = Object.keys(zip1.files).filter((name) => !name.endsWith('/'));
            check(`${short}: no part is added, only the calculation chain goes`, [names1.filter((name) => !names0.includes(name)), names0.filter((name) => !names1.includes(name) && name !== 'xl/calcChain.xml')], [[], []]);
            const markers = (xml) => xml.replace(/<xdr:(from|to)>[\s\S]*?<\/xdr:\1>/g, '<xdr:$1/>');
            const drawing0 = await text(zip0, drawingPart), drawing1 = await text(zip1, drawingPart);
            check(`${short}: the drawing part is the file's but for where its anchors are`, [drawing1 !== null && markers(drawing1) === markers(drawing0), drawing1 !== drawing0], [true, true]);
            const rels = (path) => path.replace(/\/([^/]+)$/, '/_rels/$1.rels');
            const withoutChain = (xml) => xml.replace(/<(Override|Relationship)\b[^>]*calcChain[^>]*\/>/g, '');
            check(`${short}: the relationships of the drawing and of the sheet, the content types and the workbook's relationships are the file's`,
                [(await text(zip1, rels(drawingPart))) === (await text(zip0, rels(drawingPart))), (await text(zip1, rels(sheetPart))) === (await text(zip0, rels(sheetPart))),
                 (await text(zip1, '[Content_Types].xml')) === withoutChain(await text(zip0, '[Content_Types].xml')), (await text(zip1, 'xl/_rels/workbook.xml.rels')) === withoutChain(await text(zip0, 'xl/_rels/workbook.xml.rels'))],
                [true, true, true, true]);
            const again = mountGrid();
            try {
                await again.grid.ImportXlsx(await grid.ExportXlsx());
                again.grid.SwitchSheet(sheetIndex);
                const after = again.grid.GetDrawings().map((drawing) => drawing.anchor?.fromRow ?? drawing.fromRow ?? null);
                check(`${short}: read back, ${before.length} objects one row down`, after, before.map((row) => row === null ? null : row + 1));
            } finally { again.close(); }
        }
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { close(); }
    return lines;
};

// Structural edits must preserve pivot lookup geometry through the Fable
// reader/writer path, including a gap within its data rows.
checks['pivot-structure-roundtrip'] = async () => {
    const { lines, check } = makeLog();
    const { bytesOf } = await import('/component.js');
    const { grid, close } = mountGrid();
    const again = mountGrid();
    const geometry = (p) => [p.Top, p.Left, p.Bottom, p.Right, p.FirstHeaderRow, p.FirstDataRow, p.FirstDataCol,
        [...p.RowLines].map((line) => [line.Kind, line.Repeat, [...line.Items]]),
        [...p.ColumnLines].map((line) => [line.Kind, line.Repeat, [...line.Items]])];
    try {
        await grid.ImportXlsx((await bytesOf('tests/test-documents-excel/synthetic-slicers.xlsx')).slice());
        const sheetIndex = grid.GetSheetNames().findIndex((_, index) => {
            grid.SwitchSheet(index);
            return [...grid.GetModel().Pivots].length > 0;
        });
        if (sheetIndex < 0) throw new Error('pivot fixture has no pivot');
        const initial = [...grid.GetModel().Pivots][0];
        const { formatCellRef } = await import('/spreadsheet/Model.js');
        const queryRow = initial.Bottom + 10;
        const field = [...initial.DataFields][0].replaceAll('"', '""');
        grid.SetCell(queryRow, 0, `=GETPIVOTDATA("${field}",${formatCellRef(initial.Top, initial.Left)})`);
        const total = grid.GetCell(queryRow, 0).value;
        check('fixture pivot lookup has a numeric total', typeof total, 'number');
        grid.InsertRows(initial.Top + initial.FirstDataRow + 1, 1);
        grid.InsertColumns(initial.Left + initial.FirstDataCol, 1);
        const moved = [...grid.GetModel().Pivots][0];
        check('row and column insertion expand pivot bounds', [moved.Bottom, moved.Right], [initial.Bottom + 1, initial.Right + 1]);
        check('a gap does not change the pivot total', grid.GetCell(queryRow + 1, 0).value, total);
        grid.Undo(); grid.Undo();
        check('undo restores pivot geometry', geometry([...grid.GetModel().Pivots][0]), geometry(initial));
        grid.Redo(); grid.Redo();
        await again.grid.ImportXlsx(await grid.ExportXlsx());
        again.grid.SwitchSheet(sheetIndex);
        check('saved pivot geometry reopens intact', geometry([...again.grid.GetModel().Pivots][0]), geometry(moved));
        check('saved lookup cache and formula remain usable', again.grid.GetCell(queryRow + 1, 0).value, total);
        again.grid.SetCell(queryRow + 1, 0, again.grid.GetCell(queryRow + 1, 0).formula);
        check('reopened pivot lookup recalculates', again.grid.GetCell(queryRow + 1, 0).value, total);
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { again.close(); close(); }
    return lines;
};

// Charts, shapes and groups of an opened file are drawn in the grid from the file's parts: an
// <svg> in the object's frame, with the values of the cells the series read as they are now (a
// chart is drawn again when one of them changes), the theme's colours, title, axes and legend.
// What is not drawn (a slicer, a chart of a kind the grid does not know) keeps the frame with its
// label. Drawing changes nothing: the workbook is unmodified and is exported as the file.
checks['charts-in-grid'] = async () => {
    const { lines, check, note } = makeLog();
    const { bytesOf } = await import('/component.js');
    const host = container('position:absolute;left:0;top:0;width:1180px;height:880px');
    const grid = createSpreadsheet(host);
    grid.Mount();
    const errors = [];
    grid.OnError((error) => errors.push(error.source + ': ' + error.message));
    const shoot = async (label) => { await frame(); if (window.screenshotRegion) await window.screenshotRegion(0, 0, 1180, 880); note(`screenshot: ${label}`); };
    const drawn = () => [...host.querySelectorAll('.spreadsheet-object-drawn')];
    const everyObjectSeen = () => seenAll(host);
    try {
        const original = await bytesOf('tests/real-world/chart-zoo.xlsx');
        await grid.ImportXlsx(original.slice());
        // The charts are laid out once the grid is on screen: opening the workbook lays out none.
        check('opened: the grid is there, no chart is laid out yet', [host.querySelectorAll('.spreadsheet-cell').length > 0, grid.GetChartLayouts ? grid.GetChartLayouts() : -1, host.querySelectorAll('.spreadsheet-object-pending').length], [true, 0, 6]);
        await frame();
        const inViewFirst = drawn().length;
        check('a frame later those in view are, the others when they come into view', [inViewFirst >= 1, inViewFirst + host.querySelectorAll('.spreadsheet-object-pending').length, grid.GetChartLayouts()], [true, 6, inViewFirst]);
        await everyObjectSeen();
        check('scrolled over all of them: every chart is laid out, once', [host.querySelectorAll('.spreadsheet-object-pending').length, grid.GetChartLayouts()], [0, 6]);
        const frames = drawn();
        check('chart-zoo: its six charts are drawn, none is a labelled frame', [frames.length, frames.every((node) => node.querySelector('svg') !== null), host.querySelectorAll('.spreadsheet-object-label').length], [6, true, 0]);
        const titles = frames.map((node) => node.querySelector('svg text')?.textContent ?? '');
        check('each with its title', titles.slice().sort(), ['Area', 'Bar', 'Column', 'Line', 'Pie', 'Scatter']);
        const column = frames[titles.indexOf('Column')];
        const bars = () => { const rects = [...column.querySelectorAll('svg rect')].filter((rect) => rect.getAttribute('fill') !== '#FFFFFF' && Number(rect.getAttribute('height')) < 250); return rects.map((rect) => Number(rect.getAttribute('height'))); };
        const before = bars();
        check('the column chart: four bars as high as their values (10, 25, 15, 30)', [before.length, before.map((height) => Math.round(height / before[3] * 30))], [4, [10, 25, 15, 30]]);
        check('... a frame as large as the anchor says, that takes no click', [Math.abs(column.getBoundingClientRect().width - column.querySelector('svg').getBoundingClientRect().width) < 2, getComputedStyle(column).pointerEvents], [true, 'none']);
        check('drawing the charts changes nothing: unmodified, exported as the file', [grid.IsModified(), (await grid.ExportXlsx()).length === original.length], [false, true]);
        await shoot('chart-zoo');
        // A cell a series reads changes: the chart is drawn again, with its values and its scale.
        const start = performance.now();
        grid.EnterText(2, 1, '60'); await frame();
        const after = bars();
        check('B3 = 60: the second bar is the highest now', [after.length, after[1] > after[3] * 1.9, drawn()[titles.indexOf('Column')] === column], [4, true, true]);
        note(`an entry that six charts read, with its frame: ${(performance.now() - start).toFixed(1)} ms`);
        grid.Undo(); await frame();
        check('undone: the bars as they were', bars().map((height) => Math.round(height)), before.map((height) => Math.round(height)));
        // A cell no chart reads: nothing is laid out again - every chart keeps the node it has.
        const svgsBefore = drawn().map((node) => node.querySelector('svg'));
        // (A build before GetChartRedraws counts none: the nodes say what was drawn again.)
        const redraws = () => (grid.GetChartRedraws ? grid.GetChartRedraws() : 0);
        const redrawsBefore = redraws();
        grid.EnterText(30, 12, 'x'); await frame();
        check('an entry no chart reads: no chart is laid out again', [drawn().map((node, index) => node.querySelector('svg') === svgsBefore[index]), redraws() - redrawsBefore, bars().map((height) => Math.round(height))], [[true, true, true, true, true, true], 0, before.map((height) => Math.round(height))]);
        grid.Undo(); await frame();
        // A cell a series reads typed again with the value it has: the cells the charts read are as they were.
        grid.EnterText(2, 1, String(grid.GetCell(2, 1).value)); await frame();
        check('... nor by an entry that leaves a cell they read as it was', [drawn().map((node, index) => node.querySelector('svg') === svgsBefore[index]).every(Boolean), redraws() - redrawsBefore], [true, 0]);
        // A cell they read changes: the charts that read it are laid out again, and only those.
        grid.EnterText(2, 1, '61'); await frame();
        const redrawn = drawn().map((node, index) => node.querySelector('svg') !== svgsBefore[index]);
        check('an entry a series reads: the charts that read the cell are laid out again', [redrawn[titles.indexOf('Column')], redraws() - redrawsBefore, redrawn.filter(Boolean).length], [true, redrawn.filter(Boolean).length, redraws() - redrawsBefore]);
        note(`charts laid out again by an entry in B3: ${redrawn.filter(Boolean).length} of ${redrawn.length}`);
        grid.Undo(); await frame(); grid.Undo(); await frame();
        // (A waterfall: its bars joined by level lines, light grey as the gridlines, of one length - a fraction of the chart's width.)
        const joined = (chart) => {
            const width = chart.getBBox().width;
            const lengths = [...chart.querySelectorAll('line')]
                .filter((line) => line.getAttribute('stroke') === '#D9D9D9' && line.getAttribute('y1') === line.getAttribute('y2'))
                .map((line) => Math.round(Math.abs(Number(line.getAttribute('x2')) - Number(line.getAttribute('x1')))))
                .filter((length) => length > 0 && length < width / 3);
            return lengths.some((length) => lengths.filter((other) => other === length).length >= 2);
        };
        // The dashboards: charts inside group shapes, shapes with text, slicers as labelled frames.
        for (const [file, sheet] of [['tests/real-world/dashboard-ecommerce-sales.xlsx', 'Dashboard'], ['tests/real-world/dashboard-sales-analysis.xlsx', 'Dashboard']]) {
            const bytes = await bytesOf(file);
            await grid.ImportXlsx(bytes.slice());
            const index = grid.GetSheetNames().indexOf(sheet);
            grid.SwitchSheet(index); await frame();
            await everyObjectSeen();
            const objects = grid.GetDrawings();
            const inGroups = drawn().reduce((count, node) => count + node.querySelectorAll('svg g[clip-path]').length, 0);
            const whole = drawn().filter((node) => node.querySelector('svg > rect, svg > text, svg > path, svg > g') !== null).length;
            note(`${file.split('/').pop()} "${sheet}": ${objects.length} objects, ${drawn().length} drawn (${inGroups} charts inside groups), ${host.querySelectorAll('.spreadsheet-object-label').length} labelled frames`);
            check(`${file.split('/').pop()}: objects are drawn, charts inside groups among them`, [drawn().length > 0, whole > 0, inGroups + drawn().filter((node) => node.className.includes('spreadsheet-object-chart')).length >= 3], [true, true, true]);
            check('... and the workbook is unmodified', grid.IsModified(), false);
            // (A chartEx part keeps no copy of its data: the e-commerce waterfall reads a sheet not shown yet,
            // and has its points - the grand total its last bar, its bars joined by lines - when first shown.)
            if (file.includes('ecommerce')) {
                const texts = drawn().flatMap((node) => [...node.querySelectorAll('svg text')].map((text) => text.textContent));
                check('the waterfall, its data on a sheet not shown yet: its points on first display', [texts.includes('Grand Total'), drawn().some((node) => [...node.querySelectorAll('svg g[clip-path]'), ...node.querySelectorAll('svg')].some(joined)), grid.IsModified()], [true, true, false]);
            }
            await shoot(`${file.split('/').pop()} ${sheet}`);
            // Another sheet and back: the objects are as they were drawn - nothing is laid out again.
            if (grid.GetChartLayouts) {
                // (Every sheet is shown once first: one that was not converted yet gives a chart that reads it its cells.)
                for (const info of grid.GetSheetInfos()) if (info.visibility === 'visible') { grid.SwitchSheet(info.index); await frame(); }
                grid.SwitchSheet(index); await frame();
                await everyObjectSeen();
                const layouts = grid.GetChartLayouts(), svgs = drawn().map((node) => node.innerHTML);
                const other = grid.GetSheetNames().findIndex((name, at) => at !== index && grid.GetSheetInfos()[at].visibility === 'visible');
                grid.SwitchSheet(other); await frame();
                const away = grid.GetChartLayouts();
                grid.SwitchSheet(index); await frame();
                // (Every sheet was shown: a pivot chart draws the points its cache has, not the pivot table's grand total.)
                // (A waterfall has the grand total as its total point, as in Excel: told by the lines that join its bars.
                // Each chart is a clipped group of the drawing's SVG, the innermost, or the SVG itself.)
                const charts = drawn().flatMap((node) => { const groups = [...node.querySelectorAll('svg g[clip-path]')].filter((group) => !group.querySelector('g[clip-path]')); return groups.length ? groups : [...node.querySelectorAll('svg')].slice(0, 1); });
                const textsOf = (chart) => [...chart.querySelectorAll('text')].map((text) => text.textContent);
                const isWaterfall = joined;
                check('no chart has a "Grand Total" among its points, its source sheet shown or not', charts.filter((chart) => !isWaterfall(chart)).flatMap(textsOf).filter((text) => text === 'Grand Total'), []);
                check('shown again after another sheet: the same drawings, none laid out again', [grid.GetChartLayouts() - away, away - layouts, drawn().length === svgs.length && drawn().every((node, at) => node.innerHTML === svgs[at])], [0, 0, true]);
                // Scrolling and an entry no object reads lay nothing out either.
                host.querySelector('.spreadsheet-container').scrollTop = 400; await frame();
                host.querySelector('.spreadsheet-container').scrollTop = 0; await frame();
                check('... nor by scrolling', grid.GetChartLayouts() - away, 0);
            } else check('GetChartLayouts is there', false, true);
        }
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { grid.Dispose(); host.remove(); }
    return lines;
};

// Slicers and timelines: drawn where their frames are (synthetic-slicers.xlsx is synthetic,
// tools/excel-parity/make-slicers-fixture.py; the dashboards' slicers are Excel's).
checks['slicers-in-grid'] = async () => {
    const { lines, check, note } = makeLog();
    const { bytesOf } = await import('/component.js');
    const host = container('position:absolute;left:0;top:0;width:1180px;height:880px');
    const grid = createSpreadsheet(host);
    grid.Mount();
    const errors = [];
    grid.OnError((error) => errors.push(error.source + ': ' + error.message));
    const drawn = () => [...host.querySelectorAll('.spreadsheet-object-drawn')];
    const textsOf = (node) => [...node.querySelectorAll('svg text')].map((text) => text.textContent);
    const frameWith = (text) => drawn().find((node) => textsOf(node).includes(text));
    const labels = () => [...host.querySelectorAll('.spreadsheet-object-label')].map((node) => node.textContent);
    const shoot = async (label) => { await frame(); if (window.screenshotRegion) await window.screenshotRegion(0, 0, 1180, 880); note(`screenshot: ${label}`); };
    try {
        const original = await bytesOf('tests/test-documents-excel/synthetic-slicers.xlsx');
        await grid.ImportXlsx(original.slice());
        grid.SwitchSheet(grid.GetSheetNames().indexOf('Report')); await frame();
        await seenAll(host);
        const region = frameWith('North'), product = frameWith('Product (pivot)'), timeline = frameWith('Mar - Dec 2012');
        check('Report: the two slicers and the timeline are drawn, none a labelled frame', [!!region, !!product, !!timeline, labels()], [true, true, true, []]);
        if (region) check('Region: its caption and a button an item, no data last', textsOf(region), ['Region', 'North', 'South', 'East', 'West']);
        if (product) check('Product 2: in its custom style (orange buttons)', [...product.querySelectorAll('svg rect')].filter((rect) => rect.getAttribute('fill') === '#FFC000').length, 3);
        if (timeline) check('the timeline: its caption, level, months', ['Date', 'MONTHS ▾', 'JAN', 'MAR'].every((text) => textsOf(timeline).includes(text)), true);
        await shoot('synthetic-slicers Report');
        check('drawing them changes nothing: unmodified, exported as the file', [grid.IsModified(), (await grid.ExportXlsx()).length === original.length], [false, true]);
        grid.SwitchSheet(grid.GetSheetNames().indexOf('Data')); await frame();
        await seenAll(host);
        const buttonsOf = () => { const node = frameWith('Product'); return node ? textsOf(node).slice(1) : null; };
        check('Data: the table slicer, the column\'s values', buttonsOf(), ['Apples', 'Pears', 'Plums']);
        // (The table's slicer reads its column's cells: an entry there draws it again.)
        grid.EnterText(1, 2, 'Kiwis'); await frame();
        check('... a value typed into the column is a button', buttonsOf(), ['Apples', 'Kiwis', 'Pears', 'Plums']);
        grid.Undo(); await frame();
        check('... undone', buttonsOf(), ['Apples', 'Pears', 'Plums']);
        await shoot('synthetic-slicers Data');
        for (const [file, slicers] of [['tests/real-world/dashboard-ecommerce-sales.xlsx', ['Select Region', 'Select Shipment Mode']], ['tests/real-world/dashboard-sales-analysis.xlsx', ['Financial Year', 'State', 'Category']]]) {
            const bytes = await bytesOf(file);
            await grid.ImportXlsx(bytes.slice());
            grid.SwitchSheet(grid.GetSheetNames().indexOf('Dashboard')); await frame();
            await seenAll(host);
            const name = file.split('/').pop();
            // (A caption stops before the header's buttons, cut as Excel cuts it: "Financial..." in a slicer of 116 px.)
            const captioned = (caption) => drawn().find((node) => textsOf(node).some((text) => text === caption || (text.length > 4 && caption.startsWith(text.replace(/(\.\.\.|\u2026)$/, '')))));
            check(`${name}: its slicers are drawn with their captions`, slicers.map((caption) => !!captioned(caption)), slicers.map(() => true));
            check('... no slicer is a labelled frame', labels().filter((label) => /Region|Ship|Year|State|Category/.test(label)), []);
            await shoot(`${name} Dashboard`);
        }
        check('no error reported', errors, []);
    } catch (e) { lines.push('FAIL threw: ' + String(e && e.stack || e)); }
    finally { grid.Dispose(); host.remove(); }
    return lines;
};

checks['chart-reference-edits'] = async () => {
    const { lines, check } = makeLog();
    const { bytesOf } = await import('/component.js');
    const host = container('position:absolute;left:0;top:0;width:1180px;height:880px');
    const grid = createSpreadsheet(host);
    grid.Mount();
    let seriesFill;
    const values = async () => {
        await frame(); await seenAll(host);
        const column = [...host.querySelectorAll('.spreadsheet-object-drawn')]
            .find(node => node.querySelector('svg text')?.textContent === 'Column');
        const rects = [...column.querySelectorAll('svg rect')];
        seriesFill ??= rects.at(-1).getAttribute('fill');
        const heights = rects.filter(rect => rect.getAttribute('fill') === seriesFill)
            .map(rect => Number(rect.getAttribute('height')));
        return heights.map(height => Math.round(height / heights.at(-1) * 30));
    };
    const formulas = async () => {
        const zip = await window.JSZip.loadAsync(await grid.ExportXlsx());
        const xml = new DOMParser().parseFromString(await zip.file('xl/charts/chart1.xml').async('string'), 'application/xml');
        return [...xml.getElementsByTagNameNS('*', 'f')].map(node => node.textContent);
    };
    try {
        await grid.ImportXlsx(await bytesOf('tests/real-world/chart-zoo.xlsx'));
        check('original plotted values', await values(), [10, 25, 15, 30]);
        grid.InsertRows(3, 1);
        grid.EnterText(3, 1, '20');
        check('inserting inside a series expands plotted data', await values(), [10, 25, 20, 15, 30]);
        grid.Undo(); grid.Undo();
        grid.InsertRows(0, 2);
        grid.InsertColumns(0, 1);
        grid.RenameSheet(0, 'Zoo renamed');
        check('insertions and rename retain plotted values', await values(), [10, 25, 15, 30]);
        check('series name, categories and values follow edits', await formulas(),
            ["'Zoo renamed'!$C$3", "'Zoo renamed'!$B$4:$B$7", "'Zoo renamed'!$C$4:$C$7"]);
        grid.EnterText(4, 2, '65');
        check('chart reads moved cells', await values(), [10, 65, 15, 30]);
        for (let i = 0; i < 4; i++) grid.Undo();
        check('undo restores bindings and values', [await formulas(), await values()],
            [['Zoo!$B$1', 'Zoo!$A$2:$A$5', 'Zoo!$B$2:$B$5'], [10, 25, 15, 30]]);
        for (let i = 0; i < 4; i++) grid.Redo();
        check('redo restores moved values', await values(), [10, 65, 15, 30]);
        grid.DeleteRows(4, 1);
        grid.DeleteColumns(0, 1);
        check('deletion shrinks data and shifts bindings', [await formulas(), await values()],
            [["'Zoo renamed'!$B$3", "'Zoo renamed'!$A$4:$A$6", "'Zoo renamed'!$B$4:$B$6"], [10, 15, 30]]);
        const saved = await grid.ExportXlsx();
        await grid.ImportXlsx(saved);
        check('saved chart reopens on current cells', await values(), [10, 15, 30]);
        grid.InsertRows(5, 2);
        grid.EnterText(5, 1, '20');
        grid.EnterText(6, 1, '25');
        await grid.ImportXlsx(await grid.ExportXlsx());
        check('expanded series survives save and reopen beyond original cache count', await values(), [10, 15, 20, 25, 30]);
        await grid.ImportXlsx(await bytesOf('tests/test-documents-excel/synthetic-chartex.xlsx'));
        const funnelValues = async () => {
            await frame(); await seenAll(host);
            const funnel = [...host.querySelectorAll('.spreadsheet-object-drawn')]
                .find(node => node.querySelector('svg text')?.textContent === 'Funnel');
            const widths = [...funnel.querySelectorAll('svg rect')]
                .filter(rect => rect.getAttribute('fill') !== '#FFFFFF' && Number(rect.getAttribute('height')) < 100)
                .map(rect => Number(rect.getAttribute('width')));
            return widths.map(width => Math.round(width / Math.max(...widths) * 100));
        };
        const initialFunnel = await funnelValues();
        for (let row = 1; row <= 5; row++) grid.EnterText(row, 7, String(row * 10));
        grid.SetDefinedName('_xlchart.v2.5', '=ChartEx!$H$2:$H$6');
        check('chartEx resolves a changed workbook name without a cell edit', await funnelValues(), [20, 40, 60, 80, 100]);
        grid.Undo();
        check('undo restores chartEx named binding', await funnelValues(), initialFunnel);
        grid.Redo();
        grid.InsertRows(0, 2);
        check('chartEx named binding follows structural edits', await funnelValues(), [20, 40, 60, 80, 100]);
        await grid.ImportXlsx(await grid.ExportXlsx());
        check('chartEx name survives save and reopen', await funnelValues(), [20, 40, 60, 80, 100]);
    } catch (e) { lines.push('FAIL threw: ' + String(e?.stack || e)); }
    finally { grid.Dispose(); host.remove(); }
    return lines;
};
