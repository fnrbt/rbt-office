// Browser side of run.mjs: the function table (Formula.fs, "The function table").
// Function modules are fetched only when a formula or an opened workbook needs them
// (#BUSY! meanwhile, then the value), provider functions answer through the host, and a
// formula of a provider function nobody provides keeps the value the file cached.
import { checks, makeLog, mountGrid } from '/component.js';
import { excelPackage } from '/features.js';

const fetched = (module) => performance.getEntriesByType('resource').some((entry) => entry.name.includes('/spreadsheet/Functions/' + module + '.js'));
const valueAt = (grid, row, col) => grid.GetCell(row, col).value;
// The value of a cell once it is no longer `#BUSY!` (at most `ms`).
const settled = async (grid, row, col, ms = 5000) => {
    const until = performance.now() + ms;
    while (valueAt(grid, row, col) === '#BUSY!' && performance.now() < until) await new Promise((resolve) => setTimeout(resolve, 20));
    return valueAt(grid, row, col);
};

// The value of a cell once it is `want` (at most `ms`): a recalculation after an arrival runs in a task of its own.
const becomes = async (grid, row, col, want, ms = 5000) => {
    const until = performance.now() + ms;
    while (valueAt(grid, row, col) !== want && performance.now() < until) await new Promise((resolve) => setTimeout(resolve, 20));
    return valueAt(grid, row, col);
};

checks['function-modules'] = async () => {
    const { lines, check, note } = makeLog();
    const { grid, close } = mountGrid();
    try {
        // Nothing of the modules comes with the grid.
        const before = ['FnMath', 'FnEngineering', 'FnFinancial', 'Kit'].filter(fetched);
        if (before.length === 0) check('no function module is fetched with the grid', before, []);
        else note('modules an earlier check fetched: ' + before.join(', '));
        // A formula typed: #BUSY! until its module is there, then its value.
        const mathBefore = fetched('FnMath');
        grid.SetCell(0, 0, '=ROMAN(499)');
        if (!mathBefore) check('ROMAN before its module is there', valueAt(grid, 0, 0), '#BUSY!');
        check('ROMAN once its module is there', await settled(grid, 0, 0), 'CDXCIX');
        check('the Math module was fetched', fetched('FnMath'), true);
        grid.SetCell(1, 0, '=ROMAN(A3)');
        grid.SetCell(2, 0, '1999');
        check('a loaded function calculates at once', valueAt(grid, 1, 0), 'MCMXCIX');
        grid.SetCell(3, 0, '=NOSUCHFUNCTION(1)');
        check('a name no module provides', valueAt(grid, 3, 0), '#NAME?');
        grid.SetCell(4, 0, '=#CONNECT!');
        check('a newer error value as a literal', valueAt(grid, 4, 0), '#CONNECT!');
        grid.SetCell(5, 0, '=ERROR.TYPE(#BLOCKED!)');
        check('ERROR.TYPE of #BLOCKED!', valueAt(grid, 5, 0), 11);

        // A workbook opened: its modules are loaded before it is calculated (BIN2HEX has no
        // cached value in the file, so the first calculation computes it).
        const engineeringBefore = fetched('FnEngineering');
        const book = await excelPackage('<sheetData><row r="1"><c r="A1"><f>BIN2HEX(1110)</f></c><c r="B1"><f>_xlfn.BITLSHIFT(4,2)</f></c></row>'
            + '<row r="2"><c r="A2" t="str"><f>_xlfn.WEBSERVICE("http://example.com/a")</f><v>cached answer</v></c><c r="B2"><f>_xlfn.WEBSERVICE("http://example.com/b")</f></c>'
            + '<c r="C2"><f>LEN(A2)</f><v>13</v></c></row></sheetData>');
        await grid.ImportXlsx(book);
        if (!engineeringBefore) note('the Engineering module was not loaded before the workbook was opened');
        check('BIN2HEX right after the workbook is opened', valueAt(grid, 0, 0), 'E');
        check('BITLSHIFT (stored as _xlfn.BITLSHIFT)', valueAt(grid, 0, 1), 16);
        check('the Engineering module was fetched', fetched('FnEngineering'), true);
        check('a module nothing needs is not fetched', fetched('FnFinancial'), false);
        // Provider functions without a provider: the cached value is kept, none is the error.
        check('WEBSERVICE without a provider keeps its cached value', valueAt(grid, 1, 0), 'cached answer');
        check('WEBSERVICE without a provider or cached value', valueAt(grid, 1, 1), '#CONNECT!');
        grid.SetCell(9, 9, '5');
        check('an edit keeps the cached value', valueAt(grid, 1, 0), 'cached answer');
        check('what reads it', valueAt(grid, 1, 2), 13);
        check('the provider functions are listed', grid.GetProviderFunctions().filter((f) => f.name === 'WEBSERVICE').map((f) => [f.provided, f.errorWithout]), [[false, '#CONNECT!']]);

        // A provider: #BUSY! while its Promise is pending, then its answer; the workbook is calculated again.
        const asked = [];
        grid.RegisterFunctionProvider('WEBSERVICE', (call) => {
            asked.push(call.args[0]);
            return new Promise((resolve) => setTimeout(() => resolve('got ' + call.args[0]), 30));
        });
        check('with a provider: its answer replaces the cached value', await becomes(grid, 1, 0, 'got http://example.com/a'), 'got http://example.com/a');
        check('with a provider: the formula without a value', await becomes(grid, 1, 1, 'got http://example.com/b'), 'got http://example.com/b');
        check('what reads it is calculated again', await becomes(grid, 1, 2, 24), 'got http://example.com/a'.length);
        check('each call asked once', [...asked].sort(), ['http://example.com/a', 'http://example.com/b']);
        grid.RegisterFunctionProvider('STOCKHISTORY', () => [['Date', 'Close'], [45000, 250.5]]);
        grid.SetCell(10, 0, '=STOCKHISTORY("MSFT",45000)');
        check('a provider answering at once, spilled', [valueAt(grid, 10, 0), valueAt(grid, 10, 1), valueAt(grid, 11, 0), valueAt(grid, 11, 1)], ['Date', 'Close', 45000, 250.5]);
        grid.RegisterFunctionProvider('TRANSLATE', () => Promise.reject(new Error('offline')));
        grid.SetCell(12, 0, '=TRANSLATE("Hallo","de","en")');
        check('a rejected Promise is #CONNECT!', await settled(grid, 12, 0), '#CONNECT!');
        let threw = '';
        try { grid.RegisterFunctionProvider('SUM', () => 1); } catch (e) { threw = String(e.message).split(':')[0]; }
        check('SUM is no provider function', threw, 'RegisterFunctionProvider');
        for (const name of ['WEBSERVICE', 'STOCKHISTORY', 'TRANSLATE']) grid.RegisterFunctionProvider(name, null);

        // INFO and CELL("filename") from the host.
        grid.SetHostInfo('filename', 'Plan.xlsx');
        grid.SetHostInfo('directory', 'C:\\Books\\');
        grid.SetCell(13, 0, '=CELL("filename",A1)');
        check('CELL("filename") with the host\'s file name', await settled(grid, 13, 0), 'C:\\Books\\[Plan.xlsx]Sheet1');
        grid.SetHostInfo('filename', null);
        grid.SetHostInfo('directory', null);
        // LoadFunctions ahead of use.
        await grid.LoadFunctions(['DOLLARDE']);
        check('LoadFunctions fetched the Financial module', fetched('FnFinancial'), true);
        grid.SetCell(14, 0, '=DOLLARDE(1.02,16)');
        check('a function loaded ahead calculates at once', valueAt(grid, 14, 0), 1.125);
    } catch (e) { lines.push('FAIL threw ' + String(e && e.stack || e).split('\n').slice(0, 3).join(' | ')); }
    finally { close(); }
    return lines;
};
