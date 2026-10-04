// Browser side of run.mjs, second part: what host applications rely on.
// Each check returns a list of "ok ..." / "FAIL ..." lines.
import { checks, makeLog, mountGrid, bytesOf, frame, parts, press, valueAt, formatAt, OF, createSpreadsheet, container } from '/component.js';

const cellNode = (host, row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
const borderOf = (host, row, col, side) => {
    const style = getComputedStyle(cellNode(host, row, col));
    return `${style.getPropertyValue(`border-${side}-width`)} ${style.getPropertyValue(`border-${side}-style`)} ${style.getPropertyValue(`border-${side}-color`)}`;
};
const GRIDLINE = '1px solid rgb(224, 224, 224)';

// Cell borders are drawn: each Excel border style gives a complete CSS border
// (width, line style, colour); ApplyBorder's "outside" / "inside" / "none";
// an imported workbook's borders; the HTML flavour of a copy carries them.
checks['borders'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const sides = (row, col) => ['top', 'right', 'bottom', 'left'].map((side) => borderOf(host, row, col, side));
        grid.SelectRange(1, 1, 3, 3);
        grid.ApplyBorder('all', 'thin', '#000000');
        const black = 'rgb(0, 0, 0)';
        check('all / thin: a 1px solid border on every side of every cell', [sides(1, 1), sides(2, 2), sides(3, 3)].flat().every((v) => v === `1px solid ${black}`), true);
        check('a cell outside the selection keeps its grid lines', sides(5, 5), [GRIDLINE, GRIDLINE, GRIDLINE, GRIDLINE]);
        const kinds = { thin: '1px solid', medium: '2px solid', thick: '3px solid', dashed: '1px dashed', dotted: '1px dotted', double: '3px double', hair: '1px dotted',
                        mediumDashed: '2px dashed', dashDot: '1px dashed', mediumDashDot: '2px dashed', dashDotDot: '1px dashed', mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed' };
        grid.SelectCell(6, 1);
        for (const [style, css] of Object.entries(kinds)) {
            grid.ApplyBorder('top', style, '#ff0000');
            check(`${style}: computed border-top`, borderOf(host, 6, 1, 'top'), `${css} rgb(255, 0, 0)`);
        }
        check('the other sides of that cell are untouched', sides(6, 1).slice(1), [GRIDLINE, GRIDLINE, GRIDLINE]);
        grid.ApplyBorder('top', 'none', '#000000');
        check('a style of none clears the side', borderOf(host, 6, 1, 'top'), GRIDLINE);

        // none: one step, the whole selection.
        grid.SelectRange(1, 1, 3, 3);
        grid.ApplyBorder('none', 'thin', '#000000');
        check('none: every border of the selection cleared', [sides(1, 1), sides(2, 2), sides(3, 3)].flat().every((v) => v === GRIDLINE), true);
        grid.Undo();
        check('none: one undo step brings them all back', [sides(1, 1), sides(3, 3)].flat().every((v) => v === `1px solid ${black}`), true);
        grid.Redo();
        // outside: the outline only.
        grid.ApplyBorder('outside', 'medium', '#0000ff');
        const blue = '2px solid rgb(0, 0, 255)';
        check('outside: the corner cell has its two outer sides', sides(1, 1), [blue, GRIDLINE, GRIDLINE, blue]);
        check('outside: an edge cell has one', sides(1, 2), [blue, GRIDLINE, GRIDLINE, GRIDLINE]);
        check('outside: the middle cell has none', sides(2, 2), [GRIDLINE, GRIDLINE, GRIDLINE, GRIDLINE]);
        check('outside: the far corner', sides(3, 3), [GRIDLINE, blue, blue, GRIDLINE]);
        grid.Undo();
        check('outside: one undo step', [sides(1, 1), sides(3, 3)].flat().every((v) => v === GRIDLINE), true);
        grid.ApplyBorder('inside', 'thin', '#000000');
        const line = `1px solid ${black}`;
        check('inside: the corner cell has its two inner sides', sides(1, 1), [GRIDLINE, line, line, GRIDLINE]);
        check('inside: the middle cell has all four', sides(2, 2), [line, line, line, line]);
        grid.Undo();

        // The HTML flavour of a copy.
        grid.SelectRange(1, 1, 1, 2);
        grid.ApplyBorder('bottom', 'double', '#00aa00');
        grid.SetCell(1, 1, 'x');
        const html = grid.GetSelectionHtml();
        check('GetSelectionHtml carries the border', /border-bottom:3px double #00aa00/.test(html), true);
        const probe = document.createElement('div'); probe.innerHTML = html; document.body.appendChild(probe);
        check('and a browser reads it as a border', getComputedStyle(probe.querySelector('td')).borderBottomStyle, 'double');
        probe.remove();

        // The portable painter draws them too (a thick red box on a canvas).
        grid.SelectRange(9, 1, 10, 2);
        grid.ApplyBorder('outside', 'thick', '#ff0000');
        grid.SelectCell(0, 0);   // the selection outline is drawn over the cells
        const canvas = document.createElement('canvas');
        grid.RenderPortableCanvas(canvas);
        const { columnLeft, rowTop } = await import('/spreadsheet/Layout.js');
        const model = grid.GetModel();
        const pixel = canvas.getContext('2d').getImageData(Math.round(50 + (columnLeft(model, 1) + columnLeft(model, 2)) / 2), Math.round(30 + rowTop(model, 9)), 1, 1).data;
        check('the portable painter draws the border', [pixel[0] > 200, pixel[1] < 80, pixel[2] < 80], [true, true, true]);

        // An imported workbook: every bordered cell of the sheet on display has a drawn border.
        await grid.ImportXlsx(await bytesOf('tests/real-world/milk-street-office.xlsx'));
        await frame();
        const { getCell, CellAddress } = await import('/spreadsheet/Model.js');
        let bordered = 0, drawn = 0;
        for (const node of host.querySelectorAll('.spreadsheet-cell')) {
            if (node.style.display === 'none') continue;
            const row = Number(node.getAttribute('data-row')), col = Number(node.getAttribute('data-col'));
            const format = getCell(new CellAddress(row, col), grid.GetModel())?.Format;
            for (const [side, border] of [['top', format?.BorderTop], ['left', format?.BorderLeft]]) {
                if (!border) continue;
                bordered++;
                const style = getComputedStyle(node);
                if (style.getPropertyValue(`border-${side}-style`) !== 'none' && style.getPropertyValue(`border-${side}-color`) !== 'rgb(224, 224, 224)' && parseFloat(style.getPropertyValue(`border-${side}-width`)) >= 1) drawn++;
            }
        }
        check('imported borders are drawn (bordered sides in view > 0, all drawn)', [bordered > 0, drawn], [true, bordered]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};
