// Browser side of run.mjs, second part: what host applications rely on.
// Each check returns a list of "ok ..." / "FAIL ..." lines.
import { checks, makeLog, mountGrid, bytesOf, frame, parts, press, valueAt, formatAt, OF, createSpreadsheet, container } from '/component.js';

const cellNode = (host, row, col) => host.querySelector(`.spreadsheet-cell[data-row="${row}"][data-col="${col}"]`);
// What is on screen, pixel by pixel: run.mjs takes a screenshot of a region of the page.
const screenPixels = async (x, y, width, height) => {
    const image = new Image();
    image.src = 'data:image/png;base64,' + await window.screenshotRegion(Math.round(x), Math.round(y), Math.round(width), Math.round(height));
    await image.decode();
    const canvas = document.createElement('canvas');
    canvas.width = image.width; canvas.height = image.height;
    const context = canvas.getContext('2d');
    context.drawImage(image, 0, 0);
    const data = context.getImageData(0, 0, canvas.width, canvas.height).data;
    // (Outside the region: no colour that any check looks for.)
    return { width: canvas.width, height: canvas.height, at: (px, py) => (px < 0 || py < 0 || px >= canvas.width || py >= canvas.height) ? [-999, -999, -999] : [data[(py * canvas.width + px) * 4], data[(py * canvas.width + px) * 4 + 1], data[(py * canvas.width + px) * 4 + 2]] };
};
const near = (pixel, color) => Math.abs(pixel[0] - color[0]) + Math.abs(pixel[1] - color[1]) + Math.abs(pixel[2] - color[2]) < 60;
const GRID = [224, 224, 224], BLACK = [0, 0, 0], WHITE = [255, 255, 255];
// The runs of a colour along a line of pixels: [[start, length], ...].
const runsOf = (pixels, color) => {
    const runs = [];
    pixels.forEach((pixel, i) => { if (near(pixel, color)) { const last = runs[runs.length - 1]; if (last && last[0] + last[1] === i) last[1]++; else runs.push([i, 1]); } });
    return runs;
};
// What a source of pixels (the screen, a canvas) shows around the cells of a grid: lines across
// an edge and along it. `origin` is where the sheet's cell (0, 0) is in the source.
const edgeProbe = (source, origin, model, layout) => {
    const left = (col) => origin.x + layout.columnLeft(model, col), top = (row) => origin.y + layout.rowTop(model, row);
    const column = (x, y0, y1) => { const out = []; for (let y = Math.round(y0); y < Math.round(y1); y++) out.push(source.at(Math.round(x), y)); return out; };
    const line = (y, x0, x1) => { const out = []; for (let x = Math.round(x0); x < Math.round(x1); x++) out.push(source.at(x, Math.round(y))); return out; };
    return {
        left, top, at: (x, y) => source.at(x, y),
        // The thicknesses of the lines of a colour met going down through the middle of a column, from the middle of one row to the middle of another.
        across: (col, fromRow, toRow, color) => runsOf(column((left(col) + left(col + 1)) / 2, (top(fromRow) + top(fromRow + 1)) / 2, (top(toRow) + top(toRow + 1)) / 2), color).map((run) => run[1]),
        // ... and going right through the middle of a row.
        along: (row, fromCol, toCol, color) => runsOf(line((top(row) + top(row + 1)) / 2, (left(fromCol) + left(fromCol + 1)) / 2, (left(toCol) + left(toCol + 1)) / 2), color).map((run) => run[1]),
        // Whether every pixel of a horizontal / vertical stretch is of a colour.
        solidRow: (y, x0, x1, color) => line(y, x0, x1).every((pixel) => near(pixel, color)),
        solidColumn: (x, y0, y1, color) => column(x, y0, y1).every((pixel) => near(pixel, color)),
        // The rows of pixels around an edge between two rows that are of the colour all the way from x0 to x1.
        rowsOf: (row, x0, x1, color) => { const out = []; for (let y = Math.round(top(row)) - 4; y <= Math.round(top(row)) + 4; y++) if (line(y, x0, x1).every((pixel) => near(pixel, color))) out.push(y - Math.round(top(row))); return out; },
        columnsOf: (col, y0, y1, color) => { const out = []; for (let x = Math.round(left(col)) - 4; x <= Math.round(left(col)) + 4; x++) if (column(x, y0, y1).every((pixel) => near(pixel, color))) out.push(x - Math.round(left(col))); return out; },
    };
};

// Cell borders, as Excel draws them: one line per edge two cells share - the stronger of the two
// cells' borders, the later cell's on a tie - at the thickness of its style whether it is inside a
// block or its outline; outlines continuous through corners and cell boundaries; grid lines one
// pixel wide, replaced by a border rather than beside it. Checked in the pixels of the DOM grid
// (a screenshot) and of the portable painter (a canvas), which draw the same edges.
checks['borders'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const layout = await import('/spreadsheet/Layout.js');
        const screen = async () => {
            await frame();
            const box = host.getBoundingClientRect();
            const shot = await screenPixels(box.left, box.top, Math.min(1090, box.width), Math.min(590, box.height));
            const first = cellNode(host, 0, 0).getBoundingClientRect();
            return edgeProbe(shot, { x: first.left - box.left, y: first.top - box.top }, grid.GetModel(), layout);
        };
        const painted = () => {
            const canvas = document.createElement('canvas');
            grid.RenderPortableCanvas(canvas);
            const context = canvas.getContext('2d');
            const data = context.getImageData(0, 0, canvas.width, canvas.height).data;
            // (Where the painter drew nothing the canvas is transparent: the page's white shows.)
            const at = (x, y) => { const i = (y * canvas.width + x) * 4; return (x < 0 || y < 0 || x >= canvas.width || y >= canvas.height) ? [-999, -999, -999] : data[i + 3] < 8 ? [255, 255, 255] : [data[i], data[i + 1], data[i + 2]]; };
            return edgeProbe({ at }, { x: 50, y: 30 }, grid.GetModel(), layout);
        };
        grid.SelectCell(12, 8);   // (the selection outline is drawn over the cells: away from what is measured)

        // Grid lines: one pixel between two cells, both ways.
        let probe = await screen();
        check('grid lines are 1 px wide (three rows and three columns crossed)', [probe.across(5, 5, 8, GRID), probe.along(5, 5, 8, GRID)], [[1, 1, 1], [1, 1, 1]]);

        // "all": every edge of a 3 x 3 block, inner lines as thick as the outline.
        const widths = { thin: 1, medium: 2, thick: 3 };
        for (const [style, width] of Object.entries(widths)) {
            grid.SelectRange(1, 1, 3, 3);
            grid.ApplyBorder('all', style, '#000000');
            grid.SelectCell(12, 8);
            probe = await screen();
            // Down the middle of column 2 from row 0 to row 4: the outline, two inner lines, the outline. Likewise across.
            check(`all / ${style}: the outline and the inner lines are ${width} px, one line per edge`, [probe.across(2, 0, 4, BLACK), probe.along(2, 0, 4, BLACK)], [[width, width, width, width], [width, width, width, width]]);
            check(`all / ${style}: no grid line beside a border`, [probe.across(2, 0, 4, GRID), probe.along(2, 0, 4, GRID)], [[], []]);
            const canvas = painted();
            check(`all / ${style}: the portable painter draws the same lines`, [canvas.across(2, 0, 4, BLACK), canvas.along(2, 0, 4, BLACK)], [[width, width, width, width], [width, width, width, width]]);
            check(`all / ${style}: the edges the component reports`, [grid.GetCellEdges(2, 2).Top, grid.GetCellEdges(2, 2).Right, grid.GetCellEdges(0, 2).Bottom, grid.GetCellEdges(4, 4).Top],
                [{ Style: style, Color: '#000000', Width: width }, { Style: style, Color: '#000000', Width: width }, { Style: style, Color: '#000000', Width: width }, 'grid']);
            grid.SelectRange(1, 1, 3, 3);
            grid.ApplyBorder('none', 'thin', '#000000');
        }
        // double: two rules and a gap, once per edge.
        grid.SelectRange(1, 1, 3, 3);
        grid.ApplyBorder('all', 'double', '#000000');
        grid.SelectCell(12, 8);
        probe = await screen();
        // The outer rule of a double outline goes round the block unbroken: along the top, down the
        // right side (it was two pixels off at the right corners), along the bottom and up the left.
        for (const kind of ['all', 'outside']) {
            if (kind === 'outside') { grid.SelectRange(1, 1, 3, 3); grid.ApplyBorder('none', 'thin', '#000000'); grid.SelectRange(1, 1, 3, 3); grid.ApplyBorder('outside', 'double', '#000000'); grid.SelectCell(12, 8); probe = await screen(); }
            const x0 = Math.round(probe.left(1)), x1 = Math.round(probe.left(4)), y0 = Math.round(probe.top(1)), y1 = Math.round(probe.top(4));
            check(`${kind} / double: the outer rule runs round the block (top, right, bottom, left)`,
                [probe.solidRow(y0, x0, x1 + 3, BLACK), probe.solidColumn(x1 + 2, y0, y1 + 3, BLACK), probe.solidRow(y1 + 2, x0, x1 + 3, BLACK), probe.solidColumn(x0, y0, y1 + 3, BLACK)], [true, true, true, true]);
            check(`${kind} / double: the inner rule runs round inside it, and the gap between the two is clear at the corners`,
                [probe.solidRow(y0 + 2, x0 + 2, x1 + 1, BLACK), probe.solidColumn(x1, y0 + 2, y1 + 1, BLACK), probe.solidRow(y1, x0 + 2, x1 + 1, BLACK), probe.solidColumn(x0 + 2, y0 + 2, y1 + 1, BLACK),
                 [[x0 + 1, y0 + 1], [x1 + 1, y0 + 1], [x1 + 1, y1 + 1], [x0 + 1, y1 + 1]].every(([x, y]) => !near(probe.shot ? probe.shot.at(x, y) : probe.at(x, y), BLACK))], [true, true, true, true, true]);
            if (kind === 'outside') { grid.SelectRange(1, 1, 3, 3); grid.ApplyBorder('none', 'thin', '#000000'); grid.SelectRange(1, 1, 3, 3); grid.ApplyBorder('all', 'double', '#000000'); grid.SelectCell(12, 8); probe = await screen(); }
        }
        check('all / double: each edge is two 1 px rules (not a box of its own per cell)', [probe.across(2, 0, 4, BLACK), probe.along(2, 0, 4, BLACK)], [[1, 1, 1, 1, 1, 1, 1, 1], [1, 1, 1, 1, 1, 1, 1, 1]]);
        grid.SelectRange(1, 1, 3, 3);
        grid.ApplyBorder('none', 'thin', '#000000');
        check('none: every border of the selection cleared', [grid.GetCellEdges(1, 1), grid.GetCellEdges(3, 3)].map((edges) => Object.values(edges).every((edge) => edge === 'grid')), [true, true]);
        grid.Undo();
        check('none: one undo step brings them all back', [grid.GetCellEdges(1, 1).Top.Style, grid.GetCellEdges(3, 3).Bottom.Style], ['double', 'double']);
        grid.Redo();

        // outside: an outline, continuous along its sides (no break at a cell boundary) and through its corners.
        for (const [style, width] of [['medium', 2], ['thick', 3]]) {
            grid.SelectRange(1, 1, 3, 3);
            grid.ApplyBorder('outside', style, '#000000');
            grid.SelectCell(12, 8);
            for (const [name, source] of [['the grid', await screen()], ['the portable painter', painted()]]) {
                const x0 = source.left(1), x1 = source.left(4), y0 = source.top(1), y1 = source.top(4);
                // The rows of pixels that are black all the way along the top and the bottom side, the columns along the left and the right side.
                const sides = [source.rowsOf(1, x0 + 1, x1 - 1, BLACK), source.rowsOf(4, x0 + 1, x1 - 1, BLACK), source.columnsOf(1, y0 + 1, y1 - 1, BLACK), source.columnsOf(4, y0 + 1, y1 - 1, BLACK)];
                check(`outside / ${style} in ${name}: each side is ${width} unbroken rows / columns of pixels`, sides.map((rows) => rows.length), [width, width, width, width]);
                // Through the corners: the top and bottom sides, at their rows, run on over the columns of the left and right sides.
                check(`outside / ${style} in ${name}: the top side runs through both corners`, sides[0].every((dy) => source.solidRow(source.top(1) + dy, x0 + Math.min(...sides[2]), x1 + Math.max(...sides[3]) + 1, BLACK)), true);
                check(`outside / ${style} in ${name}: the bottom side runs through both corners`, sides[1].every((dy) => source.solidRow(source.top(4) + dy, x0 + Math.min(...sides[2]), x1 + Math.max(...sides[3]) + 1, BLACK)), true);
                check(`outside / ${style} in ${name}: the left and right sides run from the top side to the bottom side`,
                    [sides[2].every((dx) => source.solidColumn(x0 + dx, y0 + Math.min(...sides[0]), y1 + Math.max(...sides[1]) + 1, BLACK)), sides[3].every((dx) => source.solidColumn(x1 + dx, y0 + Math.min(...sides[0]), y1 + Math.max(...sides[1]) + 1, BLACK))], [true, true]);
                // (The painter's grid lines are a darker grey.)
                const grey = name === 'the grid' ? GRID : [166, 166, 166];
                check(`outside / ${style} in ${name}: the inner edges keep their grid lines, 1 px`, [source.across(2, 1, 3, grey), source.along(2, 1, 3, grey), source.across(2, 1, 3, BLACK)], [[1, 1], [1, 1], []]);
            }
            check(`outside / ${style}: the corner cell has its two outer edges, the middle cell none`, [grid.GetCellEdges(1, 1).Top?.Style, grid.GetCellEdges(1, 1).Left?.Style, grid.GetCellEdges(1, 1).Bottom, grid.GetCellEdges(2, 2).Top, grid.GetCellEdges(3, 3).Bottom?.Style, grid.GetCellEdges(3, 3).Right?.Style],
                [style, style, 'grid', 'grid', style, style]);
            grid.Undo();
            check(`outside / ${style}: one undo step`, [grid.GetCellEdges(1, 1).Top, grid.GetCellEdges(3, 3).Bottom], ['grid', 'grid']);
        }
        // A single side over several cells: one unbroken line.
        grid.SelectRange(6, 1, 6, 4);
        grid.ApplyBorder('bottom', 'thin', '#000000');
        grid.SelectCell(12, 8);
        probe = await screen();
        check('bottom / thin over four cells: one unbroken 1 px line, in place of the grid line', [probe.rowsOf(7, probe.left(1), probe.left(5), BLACK).length, probe.across(2, 6, 7, GRID)], [1, []]);
        check('inside: the inner edges of a block only', (() => { grid.SelectRange(1, 1, 3, 3); grid.ApplyBorder('inside', 'thin', '#000000'); const e = [grid.GetCellEdges(1, 1), grid.GetCellEdges(2, 2)]; grid.Undo(); return [e[0].Top, e[0].Right.Style, e[0].Bottom.Style, e[0].Left, e[1].Top.Style, e[1].Left.Style]; })(), ['grid', 'thin', 'thin', 'grid', 'thin', 'thin']);

        // Two cells that state different borders for the edge they share: the stronger one, the later cell's on a tie.
        grid.SelectCell(9, 1); grid.ApplyBorder('bottom', 'thin', '#ff0000');
        grid.SelectCell(10, 1); grid.ApplyBorder('top', 'thick', '#0000ff');
        grid.SelectCell(9, 2); grid.ApplyBorder('bottom', 'thick', '#ff0000');
        grid.SelectCell(10, 2); grid.ApplyBorder('top', 'thin', '#0000ff');
        grid.SelectCell(9, 3); grid.ApplyBorder('bottom', 'medium', '#ff0000');
        grid.SelectCell(10, 3); grid.ApplyBorder('top', 'medium', '#0000ff');
        grid.SelectCell(9, 4); grid.ApplyBorder('right', 'thin', '#ff0000');
        grid.SelectCell(9, 5); grid.ApplyBorder('left', 'thin', '#0000ff');
        grid.SelectCell(12, 8);
        check('a shared edge: the stronger border (below), the stronger border (above), the later cell on a tie (below, right)',
            [grid.GetCellEdges(10, 1).Top, grid.GetCellEdges(10, 2).Top, grid.GetCellEdges(9, 3).Bottom, grid.GetCellEdges(9, 4).Right].map((edge) => `${edge.Style} ${edge.Color}`),
            ['thick #0000ff', 'thick #ff0000', 'medium #0000ff', 'thin #0000ff']);
        probe = await screen();
        const RED = [255, 0, 0], BLUE = [0, 0, 255];
        check('drawn so: one line per edge, of the winning colour', [probe.across(1, 9, 10, BLUE), probe.across(1, 9, 10, RED), probe.across(2, 9, 10, RED), probe.across(2, 9, 10, BLUE), probe.across(3, 9, 10, BLUE), probe.across(3, 9, 10, RED), probe.along(9, 4, 5, BLUE), probe.along(9, 4, 5, RED)],
            [[3], [], [3], [], [2], [], [1], []]);
        const canvas = painted();
        check('and by the portable painter', [canvas.across(1, 9, 10, BLUE), canvas.across(1, 9, 10, RED), canvas.across(2, 9, 10, RED), canvas.across(2, 9, 10, BLUE), canvas.across(3, 9, 10, BLUE), canvas.across(3, 9, 10, RED), canvas.along(9, 4, 5, BLUE), canvas.along(9, 4, 5, RED)],
            [[3], [], [3], [], [2], [], [1], []]);

        // Every style has a line of its width and CSS line style.
        const kinds = { thin: '1px solid', medium: '2px solid', thick: '3px solid', dashed: '1px dashed', dotted: '1px dotted', double: '3px double', hair: '1px dotted',
                        mediumDashed: '2px dashed', dashDot: '1px dashed', mediumDashDot: '2px dashed', dashDotDot: '1px dashed', mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed' };
        const drawnTop = (row, col) => { const edge = [...cellNode(host, row, col).querySelectorAll('.spreadsheet-edge')].map((node) => getComputedStyle(node)).find((style) => parseFloat(style.borderTopWidth) > 0); return edge ? `${edge.borderTopWidth} ${edge.borderTopStyle} ${edge.borderTopColor}` : null; };
        grid.SelectCell(12, 1);
        for (const [style, css] of Object.entries(kinds)) {
            grid.ApplyBorder('top', style, '#ff0000');
            await frame();
            check(`${style}: the edge element's border`, drawnTop(12, 1), `${css} rgb(255, 0, 0)`);
        }
        grid.ApplyBorder('top', 'none', '#000000');
        await frame();
        check('a style of none clears the side', [drawnTop(12, 1), grid.GetCellEdges(12, 1).Top], [null, 'grid']);

        // A filled cell hides the grid lines around it; a sheet without grid lines shows none.
        grid.SelectRange(14, 1, 15, 2);
        grid.ApplyBgColor('#ffeeaa');
        grid.SelectCell(12, 8);
        probe = await screen();
        check('a fill hides the grid lines of its cells', [grid.GetCellEdges(14, 1).Top, grid.GetCellEdges(15, 2).Left, probe.across(1, 14, 15, GRID), probe.across(1, 16, 17, GRID)], [null, null, [], [1]]);

        // The HTML flavour of a copy.
        grid.SelectRange(17, 1, 17, 2);
        grid.ApplyBorder('bottom', 'double', '#00aa00');
        grid.SetCell(17, 1, 'x');
        const html = grid.GetSelectionHtml();
        check('GetSelectionHtml carries the border', /border-bottom:3px double #00aa00/.test(html), true);
        const sample = document.createElement('div'); sample.innerHTML = html; document.body.appendChild(sample);
        check('and a browser reads it as a border', getComputedStyle(sample.querySelector('td')).borderBottomStyle, 'double');
        sample.remove();

        // A sheet of a file with grid lines on: 1 px too.
        await grid.ImportXlsx(await bytesOf('tests/test-documents-excel/synthetic-drawing-extras.xlsx'));
        grid.SelectCell(30, 12);
        probe = await screen();
        check('synthetic-drawing-extras.xlsx: grid lines 1 px', [grid.GetModel().ShowGridLines, probe.across(6, 9, 12, GRID), probe.along(10, 6, 9, GRID)], [true, [1, 1, 1], [1, 1, 1]]);

        // An imported workbook: every edge with a border in view is drawn as one edge element in the cell that owns it.
        await grid.ImportXlsx(await bytesOf('tests/real-world/milk-street-office.xlsx'));
        await frame();
        let bordered = 0, drawn = 0;
        for (const node of host.querySelectorAll('.spreadsheet-cell')) {
            if (node.style.display === 'none') continue;
            const row = Number(node.getAttribute('data-row')), col = Number(node.getAttribute('data-col'));
            const edges = grid.GetCellEdges(row, col);
            const own = [...node.querySelectorAll('.spreadsheet-edge')].map((edge) => getComputedStyle(edge));
            for (const [edge, side] of [[edges.Top, 'Top'], [edges.Left, 'Left']]) {
                if (!edge || edge === 'grid') continue;
                bordered++;
                if (own.some((style) => parseFloat(style[`border${side}Width`]) === edge.Width && style[`border${side}Style`] !== 'none')) drawn++;
            }
        }
        check('imported borders are drawn (bordered edges in view > 0, all drawn)', [bordered > 0, drawn], [true, bordered]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// InsertRow / DeleteRow / InsertColumn / DeleteColumn take a count; InsertRows / DeleteRows /
// InsertColumns / DeleteColumns act on the selection when called without arguments. Each is
// one change, one undo step, and formulas are rewritten once.
checks['row-column-counts'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    try {
        const { getCell, CellAddress } = await import('/spreadsheet/Model.js');
        const formula = (row, col) => { const v = getCell(new CellAddress(row, col), grid.GetModel())?.Value; return v && v.fields ? String(v.fields[0]) : null; };
        const fill = () => { for (let row = 0; row < 10; row++) { grid.SetCell(row, 0, String(row + 1)); grid.SetCell(row, 1, String((row + 1) * 10)); } grid.SetCell(12, 0, '=SUM(A1:A10)'); grid.SetCell(0, 5, '=B10+A9'); };
        fill();
        const column = async (col, n) => { const out = []; for (let row = 0; row < n; row++) out.push(await valueAt(grid, row, col)); return out; };
        const changes = [];
        grid.OnChange((change) => changes.push(change));
        const rows0 = grid.GetModel().RowCount, cols0 = grid.GetModel().ColCount;
        grid.InsertRow(2, 3);
        check('InsertRow(2, 3): three rows, one change', [await column(0, 7), changes.length, grid.GetModel().RowCount - rows0], [['1', '2', '', '', '', '3', '4'], 1, 3]);
        check('formulas follow, rewritten once', [formula(15, 0), await valueAt(grid, 15, 0), formula(0, 5)], ['=SUM(A1:A13)', '55', '=B13+A12']);
        grid.Undo();
        check('one undo step', [await column(0, 4), formula(12, 0), grid.GetModel().RowCount - rows0, grid.CanUndo()], [['1', '2', '3', '4'], '=SUM(A1:A10)', 0, true]);   // (the cells the check set with SetCell are steps too)
        grid.Redo(); grid.Undo();
        changes.length = 0;
        grid.DeleteRow(1, 4);
        check('DeleteRow(1, 4)', [await column(0, 6), changes.length, formula(8, 0), await valueAt(grid, 8, 0)], [['1', '6', '7', '8', '9', '10'], 1, '=SUM(A1:A6)', '41']);
        grid.Undo();
        check('undone in one step', [await column(0, 3), await valueAt(grid, 12, 0)], [['1', '2', '3'], '55']);
        grid.DeleteRow(0);
        check('DeleteRow(0): one row as before', [await column(0, 2), await valueAt(grid, 11, 0)], [['2', '3'], '54']);
        grid.Undo();
        grid.InsertColumn(1, 2);
        check('InsertColumn(1, 2)', [await valueAt(grid, 0, 1), await valueAt(grid, 0, 3), formula(0, 7), grid.GetModel().ColCount - cols0], ['', '10', '=D10+A9', 2]);
        grid.Undo();
        grid.DeleteColumn(1, 1);
        check('DeleteColumn(1, 1)', [await valueAt(grid, 0, 1), grid.GetModel().ColCount - cols0], ['', -1]);
        grid.Undo();
        // On the selection.
        grid.SelectRange(3, 0, 5, 1);
        changes.length = 0;
        grid.InsertRows();
        check('InsertRows(): as many rows as selected, above them', [await column(0, 8), changes.length], [['1', '2', '3', '', '', '', '4', '5'], 1]);
        grid.Undo();
        grid.SelectRange(3, 0, 5, 1);
        grid.DeleteRows();
        check('DeleteRows(): the selected rows', [await column(0, 5), await valueAt(grid, 9, 0)], [['1', '2', '3', '7', '8'], '40']);
        grid.Undo();
        check('and back in one step', await valueAt(grid, 12, 0), '55');
        grid.SelectRange(0, 1, 2, 2);
        grid.InsertColumns();
        check('InsertColumns(): two columns before the selection', [await valueAt(grid, 0, 1), await valueAt(grid, 0, 2), await valueAt(grid, 0, 3)], ['', '', '10']);
        grid.Undo();
        grid.SelectRange(0, 1, 2, 1);
        grid.DeleteColumns();
        check('DeleteColumns(): the selected column', [await valueAt(grid, 0, 1), grid.GetModel().ColCount - cols0], ['', -1]);
        grid.Undo();
        grid.InsertRows(4, 2); grid.DeleteRows(4, 2);
        check('InsertRows(at, count) / DeleteRows(at, count)', [await column(0, 6), await valueAt(grid, 12, 0)], [['1', '2', '3', '4', '5', '6'], '55']);
        grid.InsertColumns(0, 2); grid.DeleteColumns(0, 2);
        check('InsertColumns(at, count) / DeleteColumns(at, count)', [await valueAt(grid, 0, 0), await valueAt(grid, 0, 1)], ['1', '10']);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// GetRevision / MarkSaved / IsModified: undo back to the saved state is clean again; a save
// that finishes late marks the revision it wrote; OnChange carries Revision and Modified.
checks['revision'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        check('a new grid is unmodified', [grid.IsModified(), grid.GetRevision()], [false, 0]);
        const changes = [];
        grid.OnChange((change) => changes.push([change.Revision, change.Modified]));
        grid.EnterText(0, 0, 'one');
        const r1 = grid.GetRevision();
        check('an edit is a new revision, reported with the change', [grid.IsModified(), r1 !== 0, changes], [true, true, [[r1, true]]]);
        grid.SelectCell(4, 4); grid.Render();
        check('a selection move is not', grid.GetRevision(), r1);
        grid.Undo();
        check('undo restores the revision', [grid.GetRevision(), grid.IsModified(), changes[changes.length - 1]], [0, false, [0, false]]);
        grid.Redo();
        check('redo too', [grid.GetRevision(), grid.IsModified()], [r1, true]);
        grid.MarkSaved();
        check('MarkSaved: unmodified', grid.IsModified(), false);
        grid.EnterText(1, 0, 'two'); grid.SelectRange(0, 0, 1, 0); grid.ApplyFormat('bold'); grid.InsertRow(0, 2);
        const r4 = grid.GetRevision();
        check('three more steps: modified', [grid.IsModified(), new Set([r1, r4]).size], [true, 2]);
        grid.Undo(); grid.Undo();
        check('part of the way back: still modified', grid.IsModified(), true);
        grid.Undo();
        check('undone back to the saved state: unmodified', [grid.IsModified(), grid.GetRevision()], [false, r1]);
        grid.Undo();
        check('undone past it: modified', grid.IsModified(), true);
        grid.Redo();
        check('and redone to it: unmodified', grid.IsModified(), false);
        grid.Redo(); grid.Redo(); grid.Redo();
        check('redone to the end: the revision it had', grid.GetRevision(), r4);
        // A new edit after an undo is a revision of its own (not the one that was undone).
        grid.Undo();
        const undone = grid.GetRevision();
        grid.EnterText(5, 5, 'x');
        check('a new edit after an undo is a new revision', [grid.GetRevision() !== r4, grid.GetRevision() !== undone], [true, true]);
        // The save protocol: the revision that was written is the one marked.
        const written = grid.GetRevision();
        const bytes = await grid.ExportXlsx();
        grid.EnterText(6, 6, 'typed while saving');
        grid.MarkSaved(written);
        check('a save that finishes after more typing leaves the workbook modified', [grid.IsModified(), bytes.length > 0], [true, true]);
        grid.Undo();
        check('and undoing that typing gives the saved state', grid.IsModified(), false);
        // A host's write is an undo step like any other: undone, the workbook is the saved one again.
        grid.SetCell(8, 8, 'host write');
        const hostRevision = grid.GetRevision();
        check('a host write is a new revision', [grid.IsModified(), hostRevision !== written], [true, true]);
        grid.Undo();
        check('undo after a host write: the saved revision', [grid.IsModified(), grid.GetRevision(), grid.GetCell(8, 8).text], [false, written, '']);
        // Opening a workbook: saved.
        await grid.ImportXlsx(await bytesOf('tests/real-world/milk-street-office.xlsx'));
        check('an opened workbook is unmodified', grid.IsModified(), false);
        const opened = grid.GetRevision();
        grid.SwitchSheet(1); grid.SwitchSheet(0);
        check('switching sheets changes nothing', [grid.IsModified(), grid.GetRevision()], [false, opened]);
        grid.EnterText(0, 0, 'edit'); grid.Undo();
        check('edit and undo in an opened workbook: unmodified', [grid.IsModified(), grid.GetRevision()], [false, opened]);
        // Unmounted: the revision still follows.
        const bare = createSpreadsheet(document.createElement('div'));
        bare.EnterText(0, 0, 'a'); const b1 = bare.GetRevision(); bare.EnterText(0, 1, 'b');
        bare.Undo();
        check('not mounted: edits and undo move the revision', [bare.IsModified(), bare.GetRevision()], [true, b1]);
        bare.Undo();
        check('not mounted: all undone is unmodified', bare.IsModified(), false);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// ExportXlsx keeps the document properties, printer settings (with the pageSetup r:id), the
// calculation chain (only while the formula cells are where they were) and the x14 extensions
// of sheets that were not restructured; what belonged to a removed sheet is not left behind.
// (Every export here also goes through the package-integrity check of component.js.)
checks['export-package-parts'] = async () => {
    const { lines, check } = makeLog();
    const { grid, close } = mountGrid();
    const partsIn = async (bytes) => { const zip = await window.JSZip.loadAsync(bytes); const out = {}; for (const [path, file] of Object.entries(zip.files)) if (!file.dir && !/^xl\/(media|pivotCache\/pivotCacheRecords|worksheets\/sheet8)/.test(path)) out[path] = await file.async('string'); return out; };
    try {
        const { packageProblems } = await import('/component.js');
        const original = await bytesOf('tests/real-world/dashboard-sales-analysis.xlsx');
        const before = await partsIn(original);
        check('the fixture itself is whole', await packageProblems(original), []);
        await grid.ImportXlsx(original);
        grid.EnterText(40, 0, 'edited');
        let after = await partsIn(await grid.ExportXlsx());
        for (const part of ['docProps/core.xml', 'docProps/app.xml', 'xl/printerSettings/printerSettings1.bin', 'xl/calcChain.xml'])
            check(`a value edit: ${part} is carried over`, after[part] === before[part] && after[part] !== undefined, true);
        check('the package relationships name the properties', ['docProps/core.xml', 'docProps/app.xml', 'xl/workbook.xml'].every((target) => after['_rels/.rels'].includes(target)), true);
        const dataIndex = grid.GetSheetNames().indexOf('Data') + 1;
        const relId = (/<pageSetup[^>]*\sr:id="([^"]*)"/.exec(await (await window.JSZip.loadAsync(await grid.ExportXlsx())).files[`xl/worksheets/sheet${dataIndex}.xml`].async('string')) ?? [])[1];
        check('pageSetup names the printer settings by a relationship of the sheet', new RegExp(`Id="${relId}"[^>]*printerSettings1\\.bin`).test(after[`xl/worksheets/_rels/sheet${dataIndex}.xml.rels`]), true);
        check('x14 conditional formats of an untouched sheet are kept', [(after['xl/worksheets/sheet1.xml'].match(/<x14:cfRule\b/g) ?? []).length > 0, (after['xl/worksheets/sheet1.xml'].match(/<x14:cfRule\b/g) ?? []).length], [true, (before['xl/worksheets/sheet1.xml'].match(/<x14:cfRule\b/g) ?? []).length]);
        // A formula added: Excel rebuilds the chain; its relationship and content type go with it.
        grid.EnterText(41, 0, '=1+1');
        after = await partsIn(await grid.ExportXlsx());
        check('a formula added: no calcChain part, relationship or content type', [after['xl/calcChain.xml'] === undefined, /calcChain/.test(after['xl/_rels/workbook.xml.rels']), /calcChain/.test(after['[Content_Types].xml'])], [true, false, false]);
        check('the properties and printer settings stay', [after['docProps/core.xml'] === before['docProps/core.xml'], after['xl/printerSettings/printerSettings1.bin'] !== undefined], [true, true]);
        grid.Undo();
        after = await partsIn(await grid.ExportXlsx());
        check('undone: the chain is kept again', after['xl/calcChain.xml'] === before['xl/calcChain.xml'], true);
        // Rows inserted in the first sheet: its x14 data bar (a rule of the model) moves with its cells.
        grid.InsertRows(2, 2);
        after = await partsIn(await grid.ExportXlsx());
        check('rows inserted: the x14 data bar of that sheet follows its cells', [/<xm:sqref>E17:E24<\/xm:sqref>/.test(after['xl/worksheets/sheet1.xml']), /<xm:sqref>E15:E22<\/xm:sqref>/.test(after['xl/worksheets/sheet1.xml'])], [true, false]);
        grid.Undo();
        // The Data sheet removed by the host.
        const sheets = grid.GetSheets(), names = grid.GetSheetNames();
        const keep = names.map((_, i) => i).filter((i) => names[i] !== 'Data');
        grid.InstallWorkbook(keep.map((i) => sheets[i]), keep.map((i) => names[i]), 0, undefined);
        after = await partsIn(await grid.ExportXlsx());
        check('a removed sheet: its printer settings are not left in the package', [after['xl/printerSettings/printerSettings1.bin'] === undefined, /printerSettings/.test(after['[Content_Types].xml'].replace(/<Default[^>]*>/g, ''))], [true, false]);
        check('the document properties stay', after['docProps/app.xml'] === before['docProps/app.xml'], true);
        // Custom properties.
        const custom = await bytesOf('tests/real-world/milk-street-office.xlsx');
        await grid.ImportXlsx(custom);
        const customBefore = await partsIn(custom);
        after = await partsIn(await grid.ExportXlsx());
        check('custom properties and several printer settings', ['docProps/custom.xml', 'docProps/core.xml', 'docProps/app.xml', 'xl/printerSettings/printerSettings1.bin', 'xl/printerSettings/printerSettings2.bin', 'xl/calcChain.xml'].map((part) => after[part] === customBefore[part] && after[part] !== undefined), [true, true, true, true, true, true]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Text wider than its cell runs over the empty cells beside it, up to the first filled one
// (measured from the DOM); numbers never do - one that does not fit shows ####.
checks['text-overflow'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        // How far a cell's text is visible past its own box: [left, right] in px, from its clip.
        const let_out = (row, col) => {
            const node = cellNode(host, row, col), style = getComputedStyle(node);
            // (A browser gives the shortest form: "inset(0px -200px)" when left and right are the same.)
            const inset = (/inset\(([^)]*)\)/.exec(style.clipPath)?.[1] ?? '').split(' ').filter(Boolean).map((v) => -parseFloat(v));
            return style.overflow === 'visible' && inset.length >= 2 ? [inset[3] ?? inset[1], inset[1]] : [0, 0];
        };
        // The ink of a cell's text, relative to the cell's box: [left overhang, right overhang].
        const ink = (row, col) => {
            const node = cellNode(host, row, col), range = document.createRange();
            range.selectNodeContents(node);
            const text = [...range.getClientRects()].reduce((box, r) => ({ left: Math.min(box.left, r.left), right: Math.max(box.right, r.right) }), { left: Infinity, right: -Infinity });
            const box = node.getBoundingClientRect();
            return [Math.max(0, Math.round(box.left - text.left)), Math.max(0, Math.round(text.right - box.right))];
        };
        // What is drawn at a point of the grid (the topmost cell there).
        const topCellAt = (row, col) => {
            const box = cellNode(host, row, col).getBoundingClientRect();
            const node = document.elementFromPoint(box.left + box.width / 2, box.top + box.height / 2).closest('.spreadsheet-cell');
            return [Number(node.getAttribute('data-row')), Number(node.getAttribute('data-col'))];
        };
        const title = 'A title that is much longer than one column of the sheet';
        grid.SetCell(1, 1, title);
        let over = ink(1, 1), allowed = let_out(1, 1);
        check('long text: the cell shows all of it', cellNode(host, 1, 1).textContent, title);
        check('its ink runs past the cell to the right, not to the left', [over[0], over[1] > 100], [0, true]);
        check('and the cell lets it out that far (whole columns)', [allowed[0], allowed[1] >= over[1], allowed[1] % 100], [0, true, 0]);
        check('the text is drawn over the empty neighbour', topCellAt(1, 2), [1, 1]);
        check('no ellipsis', getComputedStyle(cellNode(host, 1, 1)).textOverflow, 'clip');
        check('the grid lines it crosses are covered', cellNode(host, 1, 1).querySelectorAll('.spreadsheet-overflow-gap').length, allowed[1] / 100);
        // A filled cell stops it.
        grid.SetCell(1, 3, 'stop');
        check('a filled cell two columns on: let out over one column only', let_out(1, 1), [0, 100]);
        check('that cell is drawn itself', topCellAt(1, 3), [1, 3]);
        grid.SetCell(1, 2, 'x');
        check('a filled neighbour: the text stays in its cell', [let_out(1, 1), getComputedStyle(cellNode(host, 1, 1)).overflow], [[0, 0], 'hidden']);
        grid.SetCell(1, 2, ''); grid.SetCell(1, 3, '');
        check('cleared again: let out again', let_out(1, 1)[1] >= over[1], true);
        // Right-aligned and centred text.
        grid.SetCell(3, 6, title); grid.SelectCell(3, 6); grid.ApplyFormat('alignRight');
        over = ink(3, 6); allowed = let_out(3, 6);
        check('right-aligned: runs to the left', [over[0] > 100, over[1], allowed[0] >= over[0], allowed[1]], [true, 0, true, 0]);
        grid.SetCell(5, 5, title); grid.SelectCell(5, 5); grid.ApplyFormat('alignCenter');
        over = ink(5, 5); allowed = let_out(5, 5);
        check('centred: runs both ways, evenly', [over[0] > 50, over[1] > 50, Math.abs(over[0] - over[1]) <= 2, allowed[0] >= over[0], allowed[1] >= over[1]], [true, true, true, true, true]);
        // Numbers.
        grid.SetCell(7, 1, '123456789012345678');
        grid.SetCell(8, 1, '12.5');
        const { setColumnWidth } = await import('/spreadsheet/Model.js');
        grid.SetWorkbook([setColumnWidth(1, 40, grid.GetModel())], ['Sheet1'], 0);
        await frame();
        check('a number never overflows', [let_out(7, 1), getComputedStyle(cellNode(host, 7, 1)).overflow], [[0, 0], 'hidden']);
        check('a number too wide for its cell shows ####', [/^#+$/.test(cellNode(host, 7, 1).textContent), cellNode(host, 8, 1).textContent], [true, '12.5']);
        check('and the hashes fit the cell', ink(7, 1), [0, 0]);
        check('text in the narrow column still runs over', let_out(1, 1)[1] > 100, true);
        // Wrapped text and merged cells stay where they are.
        const { getCell, CellAddress, setCell } = await import('/spreadsheet/Model.js');
        grid.SetCell(10, 1, title);
        grid.MergeCells(10, 1, 10, 2);
        check('a merged cell keeps its text inside', let_out(10, 1), [0, 0]);
        // ---- what is on screen (pixels) ----
        {
            host.style.display = 'none';   // (the grid of the checks above is in the same place)
            const pixelGrid = mountGrid();
            const g = pixelGrid.grid, h = pixelGrid.host;
            const layout = await import('/spreadsheet/Layout.js');
            const { setCell: put, getCell: get, CellAddress: Addr, setColumnWidth: widen } = await import('/spreadsheet/Model.js');
            const shoot = async () => {
                await frame();
                const box = h.getBoundingClientRect();
                const shot = await screenPixels(box.left, box.top, Math.min(1090, box.width), Math.min(590, box.height));
                const first = cellNode(h, g.GetModel().FrozenRows > 0 || g.GetModel().FrozenCols > 0 ? 0 : 0, 0).getBoundingClientRect();
                return { shot, probe: edgeProbe(shot, { x: first.left - box.left, y: first.top - box.top }, g.GetModel(), layout) };
            };
            // The colours down the column of pixels of a vertical grid line (between a row's top line and the next row's), and of the row's top line over some columns.
            const edgeColumn = (view, row, col) => { const out = []; for (let y = Math.round(view.probe.top(row)) + 1; y < Math.round(view.probe.top(row + 1)); y++) out.push(view.shot.at(Math.round(view.probe.left(col)), y)); return out; };
            const lineOver = (view, row, fromCol, toCol) => view.probe.solidRow(view.probe.top(row), view.probe.left(fromCol) + 1, view.probe.left(toCol), GRID);
            // (A grid line is grey all the way down; the soft edge of a glyph is grey in a few places.)
            const noGrid = (pixels) => pixels.filter((pixel) => near(pixel, GRID)).length < pixels.length / 2;
            const INK = (pixel) => pixel[0] < 110 && pixel[1] < 110 && pixel[2] < 110;
            const inkIn = (view, row, col) => { let n = 0; for (let y = Math.round(view.probe.top(row)) + 2; y < Math.round(view.probe.top(row + 1)) - 1; y++) for (let x = Math.round(view.probe.left(col)) + 2; x < Math.round(view.probe.left(col + 1)) - 1; x++) if (INK(view.shot.at(x, y))) n++; return n; };
            const long = 'Overflowing text that is a good deal wider than one cell';
            g.SetCell(1, 1, long);                                         // left-aligned, to the right
            g.SetCell(3, 6, long); g.SelectCell(3, 6); g.ApplyFormat('alignRight');   // to the left
            g.SetCell(5, 4, long); g.SelectCell(5, 4); g.ApplyFormat('alignCenter');  // both ways
            g.SetCell(7, 1, long); g.SetCell(7, 3, 'stop');                // stopped by a cell
            g.SetCell(9, 1, long); g.SelectCell(9, 1); g.ApplyBgColor('#ffeeaa');        // a filled cell's text
            g.SetCell(11, 1, long); g.SelectRange(11, 2, 11, 3); g.ApplyBgColor('#cce5ff'); // over filled empty cells
            g.SelectCell(14, 9);
            let view = await shoot();
            check('pixels: the text runs over the cells to its right', [inkIn(view, 1, 2) > 20, inkIn(view, 1, 3) > 20], [true, true]);
            check('pixels: the grid lines it crosses are not drawn (between B|C and C|D), the one past its end is', [noGrid(edgeColumn(view, 1, 2)), noGrid(edgeColumn(view, 1, 3)), edgeColumn(view, 1, 6).every((pixel) => near(pixel, GRID))], [true, true, true]);
            check('pixels: the lines above and below the row are whole where it crosses', [lineOver(view, 1, 1, 6), lineOver(view, 2, 1, 6)], [true, true]);
            check('pixels: right-aligned text runs to the left, its crossed lines gone, the lines above and below whole', [inkIn(view, 3, 5) > 20, inkIn(view, 3, 7), noGrid(edgeColumn(view, 3, 6)), noGrid(edgeColumn(view, 3, 5)), lineOver(view, 3, 2, 7), lineOver(view, 4, 2, 7)], [true, 0, true, true, true, true]);
            check('pixels: centred text runs both ways', [inkIn(view, 5, 3) > 20, inkIn(view, 5, 5) > 20, noGrid(edgeColumn(view, 5, 4)), noGrid(edgeColumn(view, 5, 5)), lineOver(view, 5, 2, 7), lineOver(view, 6, 2, 7)], [true, true, true, true, true, true]);
            check('pixels: text stops at the first cell that holds something (its line stays, nothing drawn past it)', [inkIn(view, 7, 2) > 20, noGrid(edgeColumn(view, 7, 2)), edgeColumn(view, 7, 3).every((pixel) => near(pixel, GRID) || INK(pixel) === false && near(pixel, GRID)), inkIn(view, 7, 4)], [true, true, true, 0]);
            const FILL = [255, 238, 170], BLUE = [204, 229, 255];
            check('pixels: a filled cell: its fill stays in its cell, the text runs on over white', [near(view.shot.at(Math.round(view.probe.left(1)) + 3, Math.round(view.probe.top(9)) + 2), FILL), near(view.shot.at(Math.round(view.probe.left(2)) + 3, Math.round(view.probe.top(9)) + 2), WHITE), inkIn(view, 9, 2) > 20, noGrid(edgeColumn(view, 9, 2)), noGrid(edgeColumn(view, 9, 3))], [true, true, true, true, true]);
            check('pixels: over filled cells the text is on their fill - no white strip at the edge between them', [inkIn(view, 11, 2) > 20, near(view.shot.at(Math.round(view.probe.left(2)) + 3, Math.round(view.probe.top(11)) + 2), BLUE), edgeColumn(view, 11, 3).filter((pixel) => near(pixel, WHITE)).length, edgeColumn(view, 11, 3).filter((pixel) => near(pixel, BLUE)).length > 5], [true, true, 0, true]);
            // Wrapped and rotated text stays in its cell.
            const model = g.GetModel();
            const withFormat = (m, row, col, change) => { const cell = get(new Addr(row, col), m); return put(new Addr(row, col), Object.assign(Object.create(Object.getPrototypeOf(cell)), cell, { Format: Object.assign(Object.create(Object.getPrototypeOf(cell.Format)), cell.Format, change) }), m); };
            g.SetCell(13, 1, long); g.SetCell(15, 1, long);
            g.SetWorkbook([withFormat(withFormat(g.GetModel(), 13, 1, { WrapText: true }), 15, 1, { TextRotation: 45 })], ['Sheet1'], 0);
            g.SelectCell(18, 9);
            view = await shoot();
            check('pixels: wrapped text does not run over', [inkIn(view, 13, 2), edgeColumn(view, 13, 2).every((pixel) => near(pixel, GRID))], [0, true]);
            check('pixels: rotated text does not run over', [inkIn(view, 15, 2), edgeColumn(view, 15, 2).every((pixel) => near(pixel, GRID))], [0, true]);
            // Frozen columns: text of a frozen cell runs on into the columns beside it while they are
            // there, and is cut at the frozen edge once they are scrolled away; text of a scrolled
            // cell does not show through the frozen columns.
            g.SetWorkbook([Object.assign(Object.create(Object.getPrototypeOf(g.GetModel())), g.GetModel(), { FrozenCols: 2 })], ['Sheet1'], 0);
            g.SelectCell(18, 9);
            view = await shoot();
            check('pixels: frozen columns, not scrolled: the text of a frozen cell runs on', inkIn(view, 1, 2) > 20, true);
            const scroller = parts(h).scroller;
            scroller.scrollLeft = 250; scroller.dispatchEvent(new Event('scroll'));
            view = await shoot();
            // On screen now: columns A, B frozen, then the scrolled columns from 250 px on. The 100 px right of the frozen edge show columns D/E.
            const frozenEdge = Math.round(view.probe.left(2));
            let inkPastEdge = 0;
            for (let y = Math.round(view.probe.top(1)) + 2; y < Math.round(view.probe.top(2)) - 1; y++) for (let x = frozenEdge + 2; x < frozenEdge + 60; x++) if (INK(view.shot.at(x, y))) inkPastEdge++;
            let inkInFrozen = 0;
            for (let y = Math.round(view.probe.top(3)) + 2; y < Math.round(view.probe.top(4)) - 1; y++) for (let x = Math.round(view.probe.left(0)) + 2; x < frozenEdge - 1; x++) if (INK(view.shot.at(x, y))) inkInFrozen++;
            check('pixels: scrolled: the frozen cell B2 is cut at the frozen edge (the cells it ran over are not there)', inkPastEdge, 0);
            check('pixels: scrolled: text of a scrolled cell (row 4, running left) does not show in the frozen columns', inkInFrozen, 0);
            pixelGrid.close();
            host.style.display = '';
        }
        // The workbook the hosts reported: B2 of the first sheet shows its long title.
        await grid.ImportXlsx(await bytesOf('tests/real-world/milk-street-office.xlsx'));
        await frame();
        const b2 = cellNode(host, 1, 1);
        over = ink(1, 1); allowed = let_out(1, 1);
        check('milk-street-office B2: the whole title is shown', [b2.textContent.length > 10, over[1] > 0, allowed[1] >= over[1], getComputedStyle(b2).textOverflow], [true, true, true, 'clip']);
        const right = b2.getBoundingClientRect().right + Math.min(over[1], 60) - 5, middle = b2.getBoundingClientRect().top + b2.getBoundingClientRect().height / 2;
        check('and drawn over the cells beside it', document.elementFromPoint(right, middle).closest('.spreadsheet-cell') === b2, true);
        // Scrolled so that column B is out of view: its text still shows in the columns it runs over.
        const { scroller } = parts(host);
        const { columnLeft } = await import('/spreadsheet/Layout.js');
        scroller.scrollLeft = columnLeft(grid.GetModel(), 2) + 30; scroller.dispatchEvent(new Event('scroll'));
        await frame();
        check('scrolled past its column: the cell is still drawn', [cellNode(host, 1, 1) !== null && cellNode(host, 1, 1).style.display !== 'none', let_out(1, 1)[1] >= over[1]], [true, true]);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Dispose removes every listener the grid registered - also those on its own (now detached)
// nodes and on the nodes of its dialogs - and a disposed grid is collected.
checks['dispose-releases'] = async () => {
    const { lines, check } = makeLog();
    // Every listener registered anywhere while this runs and not removed since.
    const live = new Map();   // target -> Map(type -> Set(handler))
    const add = EventTarget.prototype.addEventListener, remove = EventTarget.prototype.removeEventListener;
    EventTarget.prototype.addEventListener = function (type, handler, options) {
        if (!live.has(this)) live.set(this, new Map());
        const types = live.get(this);
        if (!types.has(type)) types.set(type, new Set());
        types.get(type).add(handler);
        return add.call(this, type, handler, options);
    };
    EventTarget.prototype.removeEventListener = function (type, handler, options) {
        live.get(this)?.get(type)?.delete(handler);
        return remove.call(this, type, handler, options);
    };
    const count = (where) => [...live].reduce((n, [target, types]) => n + (where(target) ? [...types.values()].reduce((m, set) => m + set.size, 0) : 0), 0);
    const describe = () => [...live].flatMap(([target, types]) => [...types].filter(([, set]) => set.size > 0).map(([type, set]) => `${target === document ? 'document' : target === window ? 'window' : (target.className || target.nodeName)}:${type}x${set.size}`));
    try {
        const run = async () => {
            const host = container('position:absolute;left:0;top:0;width:1100px;height:600px');
            const grid = createSpreadsheet(host);
            grid.Mount();
            await grid.ImportXlsx(await bytesOf('tests/real-world/milk-street-office.xlsx'));
            grid.OnChange(() => {}); grid.OnSelectionChange(() => {}); grid.OnEditStateChange(() => {}); grid.OnValidationNotice(() => {});
            grid.SelectRange(2, 1, 6, 4); grid.EnterText(3, 1, 'x'); grid.StartEdit('abc'); grid.CancelEdit();
            // Dialogs and a dropdown: one closed again, the others left open.
            grid.ShowFindReplaceDialog();
            [...document.querySelectorAll('button')].find((button) => button.textContent === 'Close')?.click();
            grid.ShowFindReplaceDialog();
            grid.ShowConditionalFormattingDialog();
            host.querySelector('.col-filter-btn')?.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
            await new Promise((resolve) => setTimeout(resolve, 150));
            const registered = count(() => true);
            const onGrid = count((target) => target instanceof Node && host.contains(target));
            const nodes = [...host.querySelectorAll('*')];
            const bodyChildren = document.body.children.length;
            grid.Dispose();
            return { registered, onGrid, weak: new WeakRef(grid), model: new WeakRef(grid.GetModel()), hostNode: host, nodes: nodes.length };
        };
        const result = await run();
        check('the grid had listeners on the document and on its own nodes', [result.registered > 10, result.onGrid > 5], [true, true]);
        check('after Dispose none is left anywhere', describe(), []);
        check('its nodes are out of the page', [result.hostNode.querySelector('.spreadsheet-container'), document.querySelectorAll('.spreadsheet-container').length], [null, 0]);
        result.hostNode.remove();
        delete result.hostNode;
        // A disposed grid nobody holds is collected, with its workbook.
        let collected = false;
        for (let attempt = 0; attempt < 5 && !collected; attempt++) {
            await new Promise((resolve) => setTimeout(resolve, 50));
            await window.collectGarbage();
            collected = result.weak.deref() === undefined && result.model.deref() === undefined;
        }
        check('a disposed grid and its workbook are collected', [result.weak.deref() === undefined, result.model.deref() === undefined], [true, true]);
        // A disposed grid can be mounted again and works.
        const host = container('position:absolute;left:0;top:0;width:900px;height:400px');
        const grid = createSpreadsheet(host);
        grid.Mount(); grid.SetCell(0, 0, 'kept'); grid.Dispose();
        check('disposed: no listeners', describe(), []);
        grid.Mount();
        check('mounted again: the workbook is still there and drawn', [cellNode(host, 0, 0)?.textContent, count(() => true) > 5], ['kept', true]);
        grid.Dispose(); host.remove();
        check('and disposed again: no listeners', describe(), []);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    EventTarget.prototype.addEventListener = add; EventTarget.prototype.removeEventListener = remove;
    return lines;
};

// ImportXlsx converts sheets to models when they are needed: opening a workbook with a large
// data sheet does not pay for that sheet until it is shown or read. Formulas into a sheet that
// is not converted, recalculation, export and the revision behave as with every sheet converted.
// Number of formula cells per sheet of a workbook file, in workbook order.
const sheetFormulas = async (bytes) => {
    const zip = await window.JSZip.loadAsync(bytes.slice());
    const workbook = await zip.file('xl/workbook.xml').async('string');
    const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const target = new Map([...rels.matchAll(/<Relationship\b[^>]*>/g)].map((m) => [/Id="([^"]+)"/.exec(m[0])[1], /Target="([^"]+)"/.exec(m[0])[1]]));
    const out = [];
    for (const m of workbook.matchAll(/<sheet\b[^>]*>/g)) {
        const part = target.get(/r:id="([^"]+)"/.exec(m[0])[1]);
        const file = zip.file(part.startsWith('/') ? part.slice(1) : 'xl/' + part);
        out.push(file ? ((await file.async('string')).match(/<f[\s>]/g) ?? []).length : 0);
    }
    return out;
};

checks['lazy-sheets'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, close } = mountGrid();
    const eager = mountGrid();
    try {
        const importer = await import('/spreadsheet/XlsxImporter.js');
        const { displayValue, getCellValue, CellAddress } = await import('/spreadsheet/Model.js');
        const shown = (model, row, col) => displayValue(getCellValue(new CellAddress(row, col), model));
        const dashboard = await bytesOf('tests/real-world/dashboard-sales-analysis.xlsx');
        // Every sheet converted up front (what ImportXlsx did before), for the time and for comparison.
        let t0 = performance.now();
        eager.grid.InstallImportedWorkbook(await importer.importWorkbookPromise(dashboard.buffer.slice(0)), -1);
        const eagerTime = performance.now() - t0;
        t0 = performance.now();
        await grid.ImportXlsx(dashboard);
        const lazyTime = performance.now() - t0;
        note(`dashboard (11 MB): ImportXlsx ${lazyTime.toFixed(0)} ms, every sheet converted up front ${eagerTime.toFixed(0)} ms`);
        check('opening converts less than everything (faster than converting every sheet)', lazyTime < eagerTime * 0.8, true);
        const names = grid.GetSheetNames(), data = names.indexOf('Data'), shownIndex = grid.GetActiveSheetIndex();
        check('only the sheet on display is converted', names.map((_, i) => grid.IsSheetLoaded(i)), names.map((_, i) => i === shownIndex));
        check('an opened workbook is unmodified', grid.IsModified(), false);
        // An edit converts the sheets that hold formulas (they are recalculated) and the sheets
        // the formulas that are evaluated read - no other.
        {
            const probe = mountGrid();
            await probe.grid.ImportXlsx(dashboard.slice());
            const formulas = await sheetFormulas(dashboard);
            probe.grid.EnterText(60, 30, 'x');
            const loaded = names.map((_, i) => probe.grid.IsSheetLoaded(i));
            note('sheets: ' + names.map((name, i) => `${i} ${name}${formulas[i] ? ' (' + formulas[i] + ' formulas)' : ''}${loaded[i] ? ' [loaded]' : ''}`).join(', '));
            check('a text typed where nothing reads it converts no sheet (the dependency graph reads the formulas of unconverted sheets off the file)',
                names.filter((_, i) => loaded[i] && i !== shownIndex), []);
            check('and it was recalculated incrementally, with the formulas of every sheet in the graph',
                [probe.grid.GetRecalcStats().Kind, probe.grid.GetRecalcStats().Formulas], ['incremental', formulas.reduce((a, b) => a + b, 0)]);
            probe.close();
        }
        const revision = grid.GetRevision();
        const changes = [];
        grid.OnChange((change) => changes.push(change.Sheets.join(',')));
        // A formula into the sheet that is not converted.
        const dataModel = eager.grid.GetSheets()[data];
        const row = grid.GetModel().RowCount - 1;
        grid.EnterText(row, 0, '=Data!A2');
        check('a formula into an unconverted sheet reads its cell', await valueAt(grid, row, 0), shown(dataModel, 1, 0));
        check('that converted the sheet', grid.IsSheetLoaded(data), true);
        check('the edit is one change, of the edited sheet only', changes, [String(shownIndex)]);
        grid.EnterText(row, 1, '=COUNTA(Data!A1:A200)');
        check('a range of it', await valueAt(grid, row, 1), '200');
        grid.Undo(); grid.Undo();
        check('undone: the revision the workbook was opened with', [grid.GetRevision(), grid.IsModified()], [revision, false]);
        // Showing a sheet converts it; that is not a change.
        const other = names.findIndex((_, i) => !grid.IsSheetLoaded(i));
        if (other >= 0) {
            changes.length = 0;
            grid.SwitchSheet(other);
            check('showing a sheet converts it, without a change to the workbook', [grid.IsSheetLoaded(other), changes, grid.IsModified()], [true, [], false]);
            grid.SwitchSheet(shownIndex);
        }

        // The same edits with sheets converted on demand and with all converted: the same workbook.
        const compare = async (file, edit) => {
            const bytes = await bytesOf(file);
            const a = mountGrid(), b = mountGrid();
            try {
                await a.grid.ImportXlsx(bytes);
                // (On the sheet ImportXlsx shows: the one the file was saved on.)
                b.grid.InstallImportedWorkbook(await importer.importWorkbookPromise(bytes.buffer.slice(0)), a.grid.GetActiveSheetIndex());
                const loadedAtOpen = a.grid.GetSheetNames().filter((_, i) => a.grid.IsSheetLoaded(i)).length;
                edit(a.grid); edit(b.grid);
                // Saved with what is converted so far, then read back.
                const saved = await a.grid.ExportXlsx(), savedEager = await b.grid.ExportXlsx();
                const stillDeferred = a.grid.GetSheetNames().filter((_, i) => !a.grid.IsSheetLoaded(i)).length;
                const reread = async (data) => { const wb = await importer.importWorkbookPromise(data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength)); return wb.Sheets.map((sheet) => sheet.Model); };
                const differences = (left, right) => {
                    const out = [];
                    left.forEach((model, index) => {
                        const rows = Math.max(model.RowCount, right[index].RowCount), cols = Math.max(model.ColCount, right[index].ColCount);
                        for (let r = 0; r < rows && out.length < 5; r++) for (let c = 0; c < cols && out.length < 5; c++) {
                            const x = shown(model, r, c), y = shown(right[index], r, c);
                            if (x !== y) out.push(`sheet ${index} R${r + 1}C${c + 1}: ${x} / ${y}`);
                        }
                    });
                    return out;
                };
                const label = file.split('/').pop();
                check(`${label}: saved files read the same`, differences(await reread(saved), await reread(savedEager)), []);
                check(`${label}: every sheet shows the same`, differences(a.grid.GetSheets(), b.grid.GetSheets()), []);
                return { loadedAtOpen, stillDeferred, sheets: a.grid.GetSheetNames().length };
            } finally { a.close(); b.close(); }
        };
        // No edit: nothing but the first sheet is ever converted, and the file is written from the source.
        let result = await compare('tests/real-world/milk-street-office.xlsx', () => {});
        check('milk-street unedited: one sheet converted at open, the rest never', [result.loadedAtOpen, result.stillDeferred], [1, result.sheets - 1]);
        // An edit that formulas of other sheets depend on: they are converted and recalculated.
        result = await compare('tests/real-world/milk-street-office.xlsx', (g) => { g.EnterText(4, 2, '123456'); g.EnterText(6, 3, '0.5'); g.InsertRows(3, 2); g.Undo(); });
        result = await compare('tests/real-world/lyric-multifamily.xlsx', (g) => { g.EnterText(5, 2, '77'); g.SelectRange(2, 1, 4, 3); g.ApplyFormat('bold'); g.SwitchSheet(1); g.EnterText(3, 3, '5'); g.DeleteRows(8, 1); });
        result = await compare('tests/real-world/dcf-stock-analysis.xlsx', (g) => { g.EnterText(3, 1, '9'); });
        // The dashboard: an edit on the first sheet leaves the data sheet unconverted, and the save does too.
        const fresh = mountGrid();
        try {
            await fresh.grid.ImportXlsx(dashboard);
            fresh.grid.EnterText(row, 0, 'edited');
            t0 = performance.now();
            const bytes = await fresh.grid.ExportXlsx();
            note(`dashboard: ExportXlsx after an edit ${(performance.now() - t0).toFixed(0)} ms`);
            check('dashboard: an edit and a save leave the data sheet unconverted', fresh.grid.IsSheetLoaded(data), false);
            const zip = await window.JSZip.loadAsync(bytes), original = await window.JSZip.loadAsync(dashboard);
            const sheetXml = async (z) => { const xml = await z.files[`xl/worksheets/sheet${data + 1}.xml`].async('string'); return [(xml.match(/<row /g) ?? []).length, (xml.match(/<c /g) ?? []).length]; };
            check('and its rows and cells are all written', await sheetXml(zip), await sheetXml(original));
            const back = await importer.importWorkbookPromise(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
            check('and read back the same', [0, 1, 500, 76929].map((r) => [0, 3, 7].map((c) => shown(back.Sheets[data].Model, r, c))), [0, 1, 500, 76929].map((r) => [0, 3, 7].map((c) => shown(dataModel, r, c))));
        } finally { fresh.close(); }
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close(); eager.close();
    return lines;
};

// The picture of tests/test-documents-excel/synthetic-drawing-extras.xlsx is a real PNG
// (24 x 16: a blue field, a white border, an orange diagonal): it decodes and is drawn, in the
// DOM grid and by the portable painter, and survives an export.
checks['fixture-picture'] = async () => {
    const { lines, check } = makeLog();
    const { grid, host, close } = mountGrid();
    try {
        const original = await bytesOf('tests/test-documents-excel/synthetic-drawing-extras.xlsx');
        await grid.ImportXlsx(original);
        const picture = grid.GetDrawings().find((drawing) => drawing.kind === 'picture');
        check('the sheet has the picture', [picture?.name, picture?.width > 20, picture?.height > 20], ['Logo', true, true]);
        const image = await new Promise((resolve) => {
            const poll = () => { const node = host.querySelector('img'); if (node && node.complete && node.naturalWidth > 0) resolve(node); else if (performance.now() - t0 > 5000) resolve(node); else setTimeout(poll, 20); };
            const t0 = performance.now(); poll();
        });
        check('the grid shows it as a decoded image of 24 x 16', [image?.naturalWidth, image?.naturalHeight], [24, 16]);
        const box = image.getBoundingClientRect();
        check('stretched over its anchor', [Math.round(box.width), Math.round(box.height)], [Math.round(picture.width), Math.round(picture.height)]);
        // Its pixels, read back from a canvas the image is drawn on.
        const probe = document.createElement('canvas'); probe.width = 24; probe.height = 16;
        probe.getContext('2d').drawImage(image, 0, 0);
        const at = (x, y) => [...probe.getContext('2d').getImageData(x, y, 1, 1).data.slice(0, 3)];
        check('its pixels: the border, the field, the diagonal', [at(0, 0), at(20, 3), at(12, 8)], [[255, 255, 255], [32, 96, 200], [240, 140, 20]]);
        // The portable painter draws the picture too (not a labelled frame).
        const canvas = document.createElement('canvas');
        grid.SelectCell(20, 8);
        grid.RenderPortableCanvas(canvas);
        const centre = canvas.getContext('2d').getImageData(Math.round(50 + picture.x + picture.width * 0.8), Math.round(30 + picture.y + picture.height * 0.25), 1, 1).data;
        check('the portable painter draws its pixels', [Math.abs(centre[0] - 32) < 30, Math.abs(centre[1] - 96) < 30, Math.abs(centre[2] - 200) < 40], [true, true, true]);
        // And it is in the exported file, byte for byte.
        const media = async (bytes) => { const zip = await window.JSZip.loadAsync(bytes); const names = Object.keys(zip.files).filter((name) => name.startsWith('xl/media/') && !zip.files[name].dir); return Promise.all(names.map(async (name) => (await zip.files[name].async('uint8array')).length)); };
        check('the exported file holds the same picture', await media(await grid.ExportXlsx()), await media(original));
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    close();
    return lines;
};

// Every call that changes the workbook is one undo step: after a typed entry and the call, one
// Undo() takes back the call alone (what the workbook holds and its revision are those after the
// entry), a second one the entry (the saved revision: unmodified), and two Redo() give the
// revision the call left.
const PIXEL = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const UNDO_STEPS = [
    ['SetCell', (g) => g.SetCell(8, 8, 'host write')],
    ['SetCell (a formula)', (g) => g.SetCell(8, 8, '=SUM(B2:B5)')],
    ['SetCell (over a value)', (g) => g.SetCell(1, 1, '99')],
    ['SetCell (clearing)', (g) => g.SetCell(1, 1, '')],
    ['EnterText', (g) => g.EnterText(8, 8, 'typed')],
    ['SetCurrentCellFromFormula', (g) => { g.SelectCell(8, 8); g.SetCurrentCellFromFormula('=B2*2'); }],
    ['FillDown', (g) => { g.SelectRange(4, 1, 7, 1); g.FillDown(); }],
    ['FillRight', (g) => { g.SelectRange(1, 2, 1, 5); g.FillRight(); }],
    ['AutoFill', (g) => { g.SelectRange(1, 1, 4, 1); g.AutoFill(8, 1); }],
    ['ApplyFormat', (g) => { g.SelectRange(1, 0, 2, 2); g.ApplyFormat('bold'); }],
    ['ApplyFormatPatch', (g) => { g.SelectRange(1, 0, 2, 2); g.ApplyFormatPatch({ italic: true, fontSize: 14 }); }],
    ['SetFontFamily', (g) => { g.SelectRange(1, 0, 2, 2); g.SetFontFamily('Georgia'); }],
    ['SetFontSize', (g) => { g.SelectRange(1, 0, 2, 2); g.SetFontSize(15); }],
    ['SetStrikethrough', (g) => { g.SelectRange(1, 0, 2, 2); g.SetStrikethrough(true); }],
    ['SetHorizontalAlign', (g) => { g.SelectRange(1, 0, 2, 2); g.SetHorizontalAlign('center'); }],
    ['SetVerticalAlign', (g) => { g.SelectRange(1, 0, 2, 2); g.SetVerticalAlign('top'); }],
    ['SetWrapText', (g) => { g.SelectRange(1, 0, 2, 2); g.SetWrapText(true); }],
    ['SetIndent', (g) => { g.SelectRange(1, 0, 2, 2); g.SetIndent(2); }],
    ['IncreaseIndent', (g) => { g.SelectRange(1, 0, 2, 2); g.IncreaseIndent(); }],
    ['SetTextRotation', (g) => { g.SelectRange(1, 0, 2, 2); g.SetTextRotation(45); }],
    ['SetShrinkToFit', (g) => { g.SelectRange(1, 0, 2, 2); g.SetShrinkToFit(true); }],
    ['SetNumberFormat', (g) => { g.SelectRange(1, 1, 4, 1); g.SetNumberFormat('0.00'); }],
    ['IncreaseDecimals', (g) => { g.SelectRange(1, 1, 4, 1); g.IncreaseDecimals(); }],
    ['ApplyFormatTo', (g) => g.ApplyFormatTo('B2:C3', { bold: true, fill: '#ffee00' })],
    ['ClearFormats', (g) => { g.ApplyFormatTo('B2:C3', { bold: true }); }, (g) => g.ClearFormats('B2:C3')],
    ['ClearContents', (g) => g.ClearContents('B2:C3')],
    ['ClearAll', (g) => g.ClearAll('B2:C3')],
    ['ApplyTextColor', (g) => { g.SelectRange(1, 0, 2, 2); g.ApplyTextColor('#ff0000'); }],
    ['ApplyBgColor', (g) => { g.SelectRange(1, 0, 2, 2); g.ApplyBgColor('#00ff00'); }],
    ['ApplyNumberFormat', (g) => { g.SelectRange(1, 1, 2, 1); g.ApplyNumberFormat('#,##0.00'); }],
    ['ApplyBorder', (g) => { g.SelectRange(1, 0, 2, 2); g.ApplyBorder('all', 'thin', '#000000'); }],
    ['MergeCells', (g) => g.MergeCells(7, 0, 8, 2)],
    ['UnmergeCells', (g) => g.MergeCells(7, 0, 8, 2), (g) => g.UnmergeCells(7, 0)],
    ['SetHyperlink', (g) => g.SetHyperlink('A2', 'https://example.com/')],
    ['RemoveHyperlink', (g) => g.SetHyperlink('A2', 'https://example.com/'), (g) => g.RemoveHyperlink('A2')],
    ['SetComment', (g) => g.SetComment(2, 2, 'a note', 'Tester')],
    ['SetComment (changing one)', (g) => g.SetComment(2, 2, 'a note', 'Tester'), (g) => g.SetComment(2, 2, 'another note', 'Tester')],
    ['SetCommentRuns', (g) => g.SetCommentRuns(2, 2, [{ text: 'bold', bold: true }, { text: ' and plain' }], 'Tester')],
    ['DeleteComment', (g) => g.SetComment(2, 2, 'a note', 'Tester'), (g) => g.DeleteComment(2, 2)],
    ['RemoveComment', (g) => g.SetComment(2, 2, 'a note', 'Tester'), (g) => g.RemoveComment(2, 2)],
    ['AddCommentReply', (g) => g.SetComment(2, 2, 'a note', 'Tester'), (g) => g.AddCommentReply(2, 2, 'a reply', 'Bob')],
    ['ResolveComment', (g) => { g.SetComment(2, 2, 'a note', 'Tester'); g.AddCommentReply(2, 2, 'a reply', 'Bob'); }, (g) => g.ResolveComment(2, 2, true)],
    ['DefineTable', (g) => g.DefineTable('A1:C5', 'Sales', true, 'TableStyleMedium2')],
    ['ResizeTable', (g) => g.DefineTable('A1:C4', 'Sales', true), (g) => g.ResizeTable('Sales', 'A1:C5')],
    ['RenameTable', (g) => g.DefineTable('A1:C5', 'Sales', true), (g) => g.RenameTable('Sales', 'Deals')],
    ['DeleteTable', (g) => g.DefineTable('A1:C5', 'Sales', true), (g) => g.DeleteTable('Sales', true)],
    ['SetTableTotalRow', (g) => g.DefineTable('A1:C5', 'Sales', true), (g) => g.SetTableTotalRow('Sales', true)],
    ['SetTableStyle', (g) => g.DefineTable('A1:C5', 'Sales', true), (g) => g.SetTableStyle('Sales', 'TableStyleLight9')],
    ['InsertRow', (g) => g.InsertRow(2, 2)],
    ['DeleteRow', (g) => g.DeleteRow(2, 1)],
    ['InsertColumn', (g) => g.InsertColumn(1, 1)],
    ['DeleteColumn', (g) => g.DeleteColumn(1, 1)],
    ['InsertRows', (g) => g.InsertRows(2, 2)],
    ['DeleteRows', (g) => g.DeleteRows(2, 1)],
    ['InsertColumns', (g) => g.InsertColumns(1, 2)],
    ['DeleteColumns', (g) => g.DeleteColumns(1, 1)],
    ['ReplaceMatch', (g) => { g.Find('north', {}); g.FindNext(); g.ReplaceMatch('NORD'); }],
    ['Replace', (g) => { g.SelectCell(1, 0); g.Replace('north', 'NORD', false); }],
    ['ReplaceAll', (g) => g.ReplaceAll('o', '0', {})],
    ['ReplaceAll (the old call)', (g) => g.ReplaceAll('o', '0', false)],
    ['AddConditionalFormat', (g) => g.AddConditionalFormat('B2:B5', { type: 'dataBar', color: '#63C384' })],
    ['UpdateConditionalFormat', (g) => g.AddConditionalFormat('B2:B5', { type: 'dataBar', color: '#63C384' }), (g) => g.UpdateConditionalFormat(g.GetConditionalFormats()[0].id, { range: 'B2:B4' })],
    ['RemoveConditionalFormat', (g) => g.AddConditionalFormat('B2:B5', { type: 'dataBar', color: '#63C384' }), (g) => g.RemoveConditionalFormat(g.GetConditionalFormats()[0].id)],
    ['ReorderConditionalFormats', (g) => { g.AddConditionalFormat('B2:B5', { type: 'dataBar', color: '#63C384' }); g.AddConditionalFormat('C2:C5', { type: 'iconSet', iconSet: '3Stars' }); },
        (g) => g.ReorderConditionalFormats(g.GetConditionalFormats().map((format) => format.id).reverse())],
    ['ClearConditionalFormats', (g) => g.AddConditionalFormat('B2:B5', { type: 'dataBar', color: '#63C384' }), (g) => g.ClearConditionalFormats('B2:B5')],
    ['RemoveConditionalFormats', (g) => g.AddConditionalFormat('B2:B5', { type: 'dataBar', color: '#63C384' }), (g) => g.RemoveConditionalFormats(2, 1)],
    ['SetValidation', (g) => g.SetValidation('B2:B5', { type: 'list', source: 'a,b,c' })],
    ['ClearValidation', (g) => g.SetValidation('B2:B5', { type: 'list', source: 'a,b,c' }), (g) => g.ClearValidation('B2:B5')],
    ['FreezeAt', (g) => g.FreezeAt(1, 1)],
    ['Unfreeze', (g) => g.FreezeAt(1, 1), (g) => g.Unfreeze()],
    ['SetColumnWidth', (g) => g.SetColumnWidth(0, 180)],
    ['SetRowHeight', (g) => g.SetRowHeight(1, 44)],
    ['AutoFitColumns', (g) => g.SetColumnWidth(0, 300), (g) => g.AutoFitColumns('A:A')],
    ['AutoFitRows', (g) => g.SetRowHeight(1, 80), (g) => g.AutoFitRows('2:2')],
    ['HideRows', (g) => g.HideRows(2, 2)],
    ['UnhideRows', (g) => g.HideRows(2, 2), (g) => g.UnhideRows(2, 2)],
    ['HideColumns', (g) => g.HideColumns(1, 1)],
    ['UnhideColumns', (g) => g.HideColumns(1, 1), (g) => g.UnhideColumns(1, 1)],
    ['GroupRows', (g) => g.GroupRows('3:4')],
    ['UngroupRows', (g) => g.GroupRows('3:4'), (g) => g.UngroupRows('3:4')],
    ['GroupColumns', (g) => g.GroupColumns('B:C')],
    ['UngroupColumns', (g) => g.GroupColumns('B:C'), (g) => g.UngroupColumns('B:C')],
    ['SetOutlineSummary', (g) => g.SetOutlineSummary(false, false)],
    ['SetShowGridLines', (g) => g.SetShowGridLines(false)],
    ['SetShowHeaders', (g) => g.SetShowHeaders(false)],
    ['SetShowFormulas', (g) => g.SetShowFormulas(true)],
    ['SetShowZeros', (g) => g.SetShowZeros(false)],
    ['SetDefinedName', (g) => g.SetDefinedName('Fresh', '=Sheet1!$A$1')],
    ['DeleteDefinedName', (g) => g.SetDefinedName('Fresh', '=Sheet1!$A$1'), (g) => g.DeleteDefinedName('Fresh')],
    ['SetIterativeCalculation', (g) => g.SetIterativeCalculation(true, 50, 0.0001)],
    ['ImportCSV', (g) => g.ImportCSV('p,q\n1,2', 7, 0)],
    ['SortCurrentColumnAsc', (g) => { g.SelectCell(2, 1); g.SortCurrentColumnAsc(); }],
    ['SortCurrentColumnDesc', (g) => { g.SelectCell(2, 0); g.SortCurrentColumnDesc(); }],
    ['SortRange', (g) => g.SortRange('A1:C5', [{ column: 1, descending: true }], { hasHeader: true })],
    ['PasteText', (g) => { g.SelectCell(7, 0); g.PasteText('x\ty\n1\t2', ''); }],
    ['SetAutoFilter', (g) => g.SetAutoFilter('A1:C5')],
    ['ClearAutoFilter', (g) => g.SetAutoFilter('A1:C5'), (g) => g.ClearAutoFilter()],
    ['SetFilter', (g) => g.SetAutoFilter('A1:C5'), (g) => g.SetFilter(1, { values: ['10'] })],
    ['ClearFilter', (g) => { g.SetAutoFilter('A1:C5'); g.SetFilter(1, { values: ['10'] }); }, (g) => g.ClearFilter()],
    ['SetPageSetup', (g) => g.SetPageSetup({ printArea: 'A1:C5', orientation: 'landscape' })],
    ['SetPageBreaks', (g) => g.SetPageBreaks({ rows: [3], columns: [2] })],
    ['AddPageBreak', (g) => g.AddPageBreak(3)],
    ['RemovePageBreak', (g) => g.AddPageBreak(3), (g) => g.RemovePageBreak(3)],
    ['ClearPageBreaks', (g) => g.AddPageBreak(3), (g) => g.ClearPageBreaks()],
    ['InsertPicture', (g) => g.InsertPicture(PIXEL, 'E2')],
    ['MovePicture', (g) => g.InsertPicture(PIXEL, 'E2'), (g) => g.MovePicture(g.GetPictures()[0].id, 'F7')],
    ['ResizePicture', (g) => g.InsertPicture(PIXEL, 'E2'), (g) => g.ResizePicture(g.GetPictures()[0].id, { width: 40, height: 30 })],
    ['SetPictureAltText', (g) => g.InsertPicture(PIXEL, 'E2'), (g) => g.SetPictureAltText(g.GetPictures()[0].id, 'a dot', 'Dot')],
    ['DeletePicture', (g) => g.InsertPicture(PIXEL, 'E2'), (g) => g.DeletePicture(g.GetPictures()[0].id)],
    ['AddSheet', (g) => g.AddSheet('Other')],
    ['RenameSheet', (g) => g.RenameSheet(0, 'First')],
    ['DuplicateSheet', (g) => g.DuplicateSheet(0)],
    ['DeleteSheet', (g) => { g.AddSheet('Other'); g.SwitchSheet(0); }, (g) => g.DeleteSheet(1)],
    ['MoveSheet', (g) => { g.AddSheet('Other'); g.SwitchSheet(0); }, (g) => g.MoveSheet(1, 0)],
    ['SetSheetVisibility', (g) => { g.AddSheet('Other'); g.SwitchSheet(0); }, (g) => g.SetSheetVisibility(1, 'hidden')],
    ['SetSheetTabColor', (g) => g.SetSheetTabColor(0, '#ff8800')],
];
const workbookState = (grid) => {
    const quiet = (read) => { try { return read(); } catch (e) { return 'threw ' + String(e && e.message); } };
    const model = grid.GetModel();
    return JSON.stringify({
        sheets: grid.GetSheetInfos(), index: grid.GetActiveSheetIndex(),
        cells: grid.GetRange('A1:L14').map((row) => row.map((cell) => [cell.formula ?? null, cell.value ?? null, cell.text ?? null, cell.format ?? null])),
        comments: quiet(() => grid.GetComments()), links: quiet(() => grid.GetHyperlinks()), rules: quiet(() => grid.GetConditionalFormats()),
        validations: quiet(() => grid.GetValidations()), tables: quiet(() => grid.GetTables()), names: quiet(() => grid.GetDefinedNames()),
        view: quiet(() => grid.GetViewOptions()), hiddenRows: grid.GetHiddenRows(), hiddenCols: grid.GetHiddenColumns(), outline: quiet(() => grid.GetOutline()),
        setup: quiet(() => grid.GetPageSetup()), breaks: quiet(() => grid.GetPageBreaks()), filters: quiet(() => grid.GetFilters()),
        pictures: quiet(() => grid.GetPictures()),
        widths: [0, 1, 2, 3].map((col) => grid.GetColumnWidth(col)), heights: [0, 1, 2, 3].map((row) => grid.GetRowHeight(row)),
        merges: Array.from(model.MergedCells ?? []), iterative: quiet(() => grid.GetIterativeCalculation()),
    });
};
checks['undo-steps'] = async () => {
    const { lines, check } = makeLog();
    const only = new URLSearchParams(location.search).get('step');
    for (const [name, first, second] of UNDO_STEPS) {
        if (only && only !== name) continue;
        const { grid, close } = mountGrid();
        try {
            const table = [['Region', 'Units', 'Price'], ['north', '10', '2.5'], ['south', '30', '1.5'], ['east', '20', '4'], ['west', '5', '3']];
            table.forEach((row, r) => row.forEach((text, c) => grid.EnterText(r, c, text)));
            grid.EnterText(0, 4, '=SUM(B2:B5)');
            if (second) first(grid);
            grid.SelectCell(0, 0);
            grid.MarkSaved();
            const saved = grid.GetRevision(), s0 = workbookState(grid);
            grid.EnterText(10, 10, 'typed');
            const r1 = grid.GetRevision(), s1 = workbookState(grid);
            (second ?? first)(grid);
            const r2 = grid.GetRevision(), s2 = workbookState(grid);
            const problems = [];
            if (s2 === s1) problems.push('the call changed nothing the check sees');
            if (r2 === r1 || r2 === saved) problems.push(`no new revision (${saved}, ${r1}, ${r2})`);
            grid.Undo();
            if (workbookState(grid) !== s1) problems.push('one Undo does not give the workbook as it was before the call' + (workbookState(grid) === s2 ? ' (nothing was undone)' : workbookState(grid) === s0 ? ' (the entry before it was undone too)' : ''));
            if (grid.GetRevision() !== r1) problems.push(`one Undo: revision ${grid.GetRevision()}, not the ${r1} before the call`);
            grid.Undo();
            if (workbookState(grid) !== s0) problems.push('two Undo do not give the saved workbook');
            if (grid.GetRevision() !== saved || grid.IsModified()) problems.push(`two Undo: revision ${grid.GetRevision()} (saved ${saved}), modified ${grid.IsModified()}`);
            grid.Redo(); grid.Redo();
            if (workbookState(grid) !== s2) problems.push('two Redo do not give the workbook the call left');
            if (grid.GetRevision() !== r2) problems.push(`two Redo: revision ${grid.GetRevision()}, not ${r2}`);
            check(`${name}: one undo step`, problems, []);
        } catch (e) { lines.push(`FAIL ${name} threw ` + String(e && e.stack || e).split('\n').slice(0, 2).join(' | ')); }
        close();
    }
    return lines;
};
