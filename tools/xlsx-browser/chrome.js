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
        check('grouped rows: a strip of three steps left of the row numbers', [origin(), buttons('rows'), nodes('level').map((n) => n.textContent)], [[98, 30], [[2, 7, '−'], [3, 4, '−']], ['1', '2', '3']]);
        check('the +/- buttons are on the summary rows (the row below each group)', [rowOfButton(buttonOf('rows', 2)), rowOfButton(buttonOf('rows', 3))], [8, 5]);
        check('the row numbers moved right with the cells', Math.round(host.querySelector('.spreadsheet-row-header[data-row="0"]').getBoundingClientRect().left - inner()[0]), 48);
        check('the brackets are drawn', nodes('line').length, 2);
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
        check('grouped columns: a strip above the column letters', [origin(), buttons('columns'), colOfButton(buttonOf('columns', 2)), nodes('level').filter((n) => n.getAttribute('data-axis') === 'columns').map((n) => n.textContent)], [[98, 63], [[2, 3, '−']], 4, ['1', '2']]);
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
        check('zoomed: the strip scales', origin(), [147, 95]);
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
        check('SetShowOutline(true)', [origin(), nodes('button').length], [[98, 63], 3]);
        grid.SetShowHeaders(false); await frame();
        check('no headers: no strip', [host.querySelectorAll('.spreadsheet-outline').length, origin()], [0, [0, 0]]);
        grid.SetShowHeaders(true); await frame();
        check('headers again', origin(), [98, 63]);
        grid.SetReadOnly(true); await frame();
        await click(buttonOf('rows', 3));
        check('a viewer draws the strip; its buttons do nothing', [nodes('button').length, grid.GetHiddenRows()], [3, []]);
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
            // An edit of a cell on every sheet: still nothing lost (drawings and tables of the edited sheets included).
            const sheets = grid.GetSheets();
            for (let index = 0; index < sheets.length; index++) { if (sheets[index].visibility && sheets[index].visibility !== 'visible') continue; grid.SwitchSheet(index); const used = grid.GetUsedRange(); grid.SetCell((used ? used.lastRow : 0) + 3, 0, 'edited'); }
            await compare('a cell edited on each sheet', await grid.ExportXlsx());
        } catch (e) { lines.push(`FAIL ${short} threw: ` + String(e && e.stack || e)); }
        finally { close(); }
    }
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
