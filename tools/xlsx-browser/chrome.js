// Browser side of run.mjs, fifth part: the grid's own chrome that hosts build on (filter
// buttons, the outline gutter, drawings over frozen panes, ...). Input goes through the
// browser's real mouse and keyboard (window.realMouse / window.realKey, see run.mjs).
// Each check returns a list of "ok ..." / "FAIL ..." lines.
import { checks, makeLog, mountGrid, frame, parts } from '/component.js';
import { reopened } from '/sheet-api.js';

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
