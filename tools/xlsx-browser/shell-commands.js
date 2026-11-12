// Selection-driven shell behavior against the mounted spreadsheet component.
import { checks, makeLog, mountGrid, frame } from '/component.js';

const mount = async () => {
    const opened = mountGrid();
    const host = document.createElement('div');
    document.body.appendChild(host);
    const { createSpreadsheetShell } = await import('/spreadsheet/Toolbar.js');
    const callbacks = [];
    const shell = createSpreadsheetShell(opened.grid, (action) => callbacks.push(action));
    shell.Mount(host);
    return { ...opened, host, gridHost: opened.host, shell, callbacks, closeAll: () => { shell.Dispose(); host.remove(); opened.close(); } };
};
const key = (target, text, options = {}) => {
    const event = new KeyboardEvent('keydown', { key: text, ctrlKey: true, bubbles: true, cancelable: true, ...options });
    target.dispatchEvent(event);
    return event.defaultPrevented;
};
// Shared dialog renderer can mount outside the shell's own chrome subtree.
const dialog = () => document.querySelector('[data-testid="xl-dlg-format-cells"]');
const field = (id) => dialog()?.querySelector(`[data-testid="dlg-field-${id}"]`);
const select = (id, value) => {
    const control = field(id);
    control.value = value;
    control.dispatchEvent(new Event('change', { bubbles: true }));
};
const type = (id, value) => {
    const control = field(id);
    control.value = value;
    control.dispatchEvent(new Event('input', { bubbles: true }));
};
const menu = (host, section, id) => {
    host.querySelector(`[data-sheet-menu="${section}"]`).click();
    host.querySelector(`[data-sheet-menu-item="${id}"]`).click();
};

checks['shell-selection-commands'] = async () => {
    const { lines, check } = makeLog();
    const first = await mount(), other = await mount();
    const outside = document.createElement('input');
    document.body.appendChild(outside);
    const { grid, host, gridHost, shell, callbacks } = first;
    const button = (id) => host.querySelector(`[data-sheet-icon="${id}"]`);
    const gridTarget = first.host === first.gridHost ? first.gridHost : first.gridHost.querySelector('.spreadsheet-container');
    try {
        grid.SetCell(0, 0, 'one');
        grid.SetCell(0, 1, 'two');
        grid.SetCell(0, 2, 'three');
        grid.SelectRange(0, 0, 0, 1);
        button('bold').click();
        check('bold button changes selected cells only', [0, 1, 2].map((col) => grid.GetCell(0, col).format.bold), [true, true, false]);
        check('selection makes bold button checked', button('bold').getAttribute('aria-pressed'), 'true');
        menu(host, 'edit', 'undo');
        check('menu undo restores both selected cells in one step and button state',
            [[0, 1, 2].map((col) => grid.GetCell(0, col).format.bold), button('bold').getAttribute('aria-pressed')], [[false, false, false], 'false']);
        grid.SelectCell(0, 2);
        check('caret moves state off independently of history', button('bold').getAttribute('aria-pressed'), 'false');
        const notifications = [];
        const offCommands = shell.Commands.OnChange(() => notifications.push(shell.Commands.State('format.bold').Enabled));
        check('registry accepts a parameterized color on the active cell',
            shell.Commands.Run('format.textColor', '#123456'), true);
        check('parameterized command changes only its selected cell and not the host callback',
            [grid.GetCell(0, 0).format.textColor, grid.GetCell(0, 2).format.textColor, callbacks, notifications.length > 0],
            [null, '#123456', [], true]);
        offCommands();
        grid.Undo();
        grid.SelectRange(0, 0, 0, 1);
        grid.ApplyFormatPatch({ italic: true, fill: '#ff0000', numberFormat: '0.00' });
        grid.SelectCell(0, 1);
        grid.ApplyFormatPatch({ fill: '#0000ff', fontFamily: 'Georgia', bold: true });
        grid.SelectRange(0, 0, 0, 1);
        check('mixed selection is announced on toolbar', button('bold').getAttribute('aria-pressed'), 'mixed');
        check('registry reports selected mixed and enabled state',
            [shell.Commands.State('format.bold').Mixed, shell.Commands.State('format.bold').Enabled], [true, true]);
        check('format shortcut scoped to owned grid only', [key(outside, '1'), key(other.gridHost.querySelector('.spreadsheet-container'), '1')], [false, true]);
        await frame();
        check('foreign shortcut opens the other grid’s selection state, not the first grid’s',
            [!!dialog(), field('bold')?.value, grid.GetSelectionFormat().bold, other.grid.GetSelectionFormat().bold],
            [true, 'false', 'mixed', false]);
        dialog().querySelector('[data-testid="dlg-cancel"]').click();
        await frame();
        check('composition does not open format', key(gridTarget, '1', { isComposing: true }), false);
        check('Ctrl+1 opens one selection-aware dialog', key(gridTarget, '1'), true);
        await frame();
        check('mixed fill, family and bold show explicitly', [field('fill').value, field('fontFamily').value, field('bold').value], ['<mixed>', '<mixed>', '<mixed>']);
        select('bold', 'true');
        dialog().querySelector('[data-testid="dlg-format-cells-ok"]').click();
        await frame();
        check('edited field affects selection but mixed fill/family and existing italic/number stay',
            [0, 1, 2].map((col) => { const f = grid.GetCell(0, col).format; return [f.bold, f.fill, f.fontFamily, f.italic, f.numberFormat]; }),
            [[true, '#FF0000', null, true, '0.00'], [true, '#0000FF', 'Georgia', true, '0.00'], [false, null, null, false, null]]);
        grid.Undo();
        check('dialog patch undoes as one transaction without undoing earlier edits',
            [grid.GetCell(0, 0).format.bold, grid.GetCell(0, 1).format.bold, grid.GetCell(0, 1).format.fill], [false, true, '#0000FF']);
        const revision = grid.GetRevision();
        check('Ctrl+B is handled once and not routed through host', key(gridTarget, 'b'), true);
        check('single keystroke toggles whole selected range once',
            [[0, 1].map((col) => grid.GetCell(0, col).format.bold), grid.GetRevision() !== revision, callbacks], [[true, true], true, []]);
        grid.Undo();
        check('undo restores mixed selection state', [button('bold').getAttribute('aria-pressed'), grid.GetSelectionFormat().bold], ['mixed', 'mixed']);
        check('Ctrl+I toggles once across selected cells', key(gridTarget, 'i'), true);
        check('italic keyboard route and toolbar state agree',
            [[0, 1].map((col) => grid.GetCell(0, col).format.italic), button('italic').getAttribute('aria-pressed')], [[false, false], 'false']);
        grid.Undo();
        check('Ctrl+U toggles once across selected cells', key(gridTarget, 'u'), true);
        check('underline keyboard route and toolbar state agree',
            [[0, 1].map((col) => grid.GetCell(0, col).format.underline), button('underline').getAttribute('aria-pressed')], [[true, true], 'true']);
        grid.Undo();
        const input = document.createElement('input');
        host.appendChild(input);
        check('shortcut does not commandeer shell text inputs', [key(input, 'b'), grid.GetCell(0, 0).format.bold], [false, false]);
        input.remove();
        grid.SetValidation('B2', { type: 'whole', operator: 'between', formula1: 1, formula2: 5 });
        grid.SelectCell(1, 1);
        grid.StartEdit('99');
        check('format keyboard does not intercept pending cell entry',
            key(gridHost.querySelector('.spreadsheet-edit-input'), '1'), false);
        menu(host, 'format', 'format-cells');
        await frame();
        check('failed validation keeps pending edit and prevents Format Cells',
            [grid.IsEditing(), !!dialog()], [true, false]);
        grid.CancelEdit();
        grid.SelectRange(0, 0, 0, 1);
        shell.Dispose();
        check('disposed shell leaves owned shortcuts alone', key(gridTarget, '1'), false);
        shell.Mount(host);
        check('remounted shell has one shortcut listener', key(gridTarget, 'b'), true);
        check('remounted handler changes bold exactly once', grid.GetCell(0, 0).format.bold, true);
        shell.Dispose();
        grid.Undo();
        await grid.ImportXlsx(await grid.ExportXlsx());
        shell.Mount(host);
        grid.SelectCell(0, 1);
        check('after import toolbar reads actual cell font', host.querySelector('[data-sheet-format="fontFamily"]').value, grid.GetCell(0, 1).format.fontFamily ?? '');
        const before = [grid.GetCell(0, 0).format, grid.GetCell(0, 1).format];
        menu(host, 'format', 'format-cells');
        await frame();
        type('numberFormat', '0.000');
        dialog().querySelector('[data-dlg-tab="Alignment"]').click();
        select('horizontalAlign', 'center');
        select('wrap', 'true');
        dialog().querySelector('[data-dlg-tab="Font"]').click();
        type('fontFamily', 'Arial');
        type('fontSize', '16');
        type('textColor', '#abcdef');
        select('underline', 'double');
        dialog().querySelector('[data-dlg-tab="Border"]').click();
        select('top-style', 'medium');
        type('top-color', '#123456');
        dialog().querySelector('[data-dlg-tab="Fill"]').click();
        type('fill', '#112233');
        dialog().querySelector('[data-testid="dlg-format-cells-ok"]').click();
        await frame();
        const changed = grid.GetCell(0, 1).format;
        check('all supported dialog tabs patch only selected cell and chosen properties',
            [changed.numberFormat, changed.horizontalAlign, changed.wrap, changed.fontFamily, changed.fontSize,
                changed.textColor, changed.underline, changed.doubleUnderline, changed.borders.top, changed.fill,
                grid.GetCell(0, 0).format],
            ['0.000', 'center', true, 'Arial', 16, '#ABCDEF', true, true,
                { style: 'medium', color: '#123456' }, '#112233', before[0]]);
        grid.Undo();
        check('all tab edits undo together without touching other cells',
            [grid.GetCell(0, 0).format, grid.GetCell(0, 1).format], before);
        const modeRevision = grid.GetRevision();
        grid.SetReadOnly(true);
        check('viewer mode immediately disables visible editing and sheet controls without an edit',
            [button('bold').disabled, host.querySelector('[data-sheet-tab-add]').disabled,
                shell.Commands.State('axis.rowSize').Enabled, grid.GetRevision()],
            [true, true, false, modeRevision]);
        grid.SetReadOnly(false);
        check('leaving viewer mode restores visible controls without moving selection or editing',
            [button('bold').disabled, host.querySelector('[data-sheet-tab-add]').disabled,
                shell.Commands.State('axis.rowSize').Enabled, grid.GetRevision()],
            [false, false, true, modeRevision]);
    } catch (error) { lines.push('FAIL threw ' + String(error?.stack ?? error).split('\n').slice(0, 3).join(' | ')); }
    finally { outside.remove(); first.closeAll(); other.closeAll(); }
    return lines;
};

checks['shell-find-selection'] = async () => {
    const { lines, check } = makeLog();
    const session = await mount();
    const { grid, host, gridHost, callbacks, shell } = session;
    try {
        grid.SetCell(0, 0, 'apple');
        grid.SetCell(0, 1, 'pear');
        grid.SetCell(0, 2, 'apple');
        const target = gridHost.querySelector('.spreadsheet-container');
        check('Ctrl+F opens shell search once', key(target, 'f'), true);
        check('search bar visible, legacy grid dialog absent',
            [host.querySelector('.sheet-find-bar').dataset.hidden, !!document.getElementById('find-replace-dialog')], ['false', false]);
        const find = host.querySelector('.sheet-find-input'), replace = host.querySelector('.sheet-replace-input');
        const foreign = document.createElement('input');
        document.body.appendChild(foreign);
        check('Ctrl+F ignores foreign and search inputs', [key(foreign, 'f'), key(find, 'f')], [false, false]);
        foreign.remove();
        menu(host, 'edit', 'find-replace');
        check('Edit menu uses same command to close search', host.querySelector('.sheet-find-bar').dataset.hidden, 'true');
        menu(host, 'edit', 'find-replace');
        check('Edit menu reopens owned search', host.querySelector('.sheet-find-bar').dataset.hidden, 'false');
        find.value = 'apple'; find.dispatchEvent(new Event('input', { bubbles: true }));
        check('live search counts actual sheet cells', host.querySelector('.sheet-find-result').textContent, '2 matches');
        host.querySelector('[data-sheet-icon="find-run"]').click();
        check('next match selects next actual cell', grid.GetSelectionRef(), 'C1');
        host.querySelector('[data-sheet-icon="find-run"]').click();
        check('next match wraps to first actual cell', grid.GetSelectionRef(), 'A1');
        replace.value = 'plum';
        host.querySelector('[data-sheet-icon="replace-run"]').click();
        check('replace all changes both matches, not neighbor', [0, 1, 2].map((col) => grid.GetCell(0, col).value), ['plum', 'pear', 'plum']);
        check('live result no longer retains stale matches after replacement', host.querySelector('.sheet-find-result').textContent, 'Replaced 2 cells');
        find.dispatchEvent(new Event('input', { bubbles: true }));
        check('requery sees zero after replace', host.querySelector('.sheet-find-result').textContent, '0 matches');
        grid.Undo();
        check('replacement is one undo step and results refresh after undo',
            [[0, 1, 2].map((col) => grid.GetCell(0, col).value), host.querySelector('.sheet-find-result').textContent], [['apple', 'pear', 'apple'], '2 matches']);
        grid.SetCell(0, 2, 'orange');
        check('live edit invalidates previous result count', host.querySelector('.sheet-find-result').textContent, '1 matches');
        grid.SetValidation('B2', { type: 'whole', operator: 'between', formula1: 1, formula2: 5 });
        grid.SelectCell(1, 1);
        grid.StartEdit('99');
        check('Ctrl+F during cell editing does not hijack the buffer',
            key(gridHost.querySelector('.spreadsheet-edit-input'), 'f'), false);
        host.querySelector('[data-sheet-icon="find-run"]').click();
        check('failed validation blocks navigation, preserves entry',
            [grid.IsEditing(), grid.GetSelectionRef(), host.querySelector('.sheet-find-result').textContent], [true, 'B2', 'Find blocked: correct the cell entry first.']);
        grid.CancelEdit();
        const saved = await grid.ExportXlsx();
        grid.SetCell(0, 0, 'unrelated');
        await grid.ImportXlsx(saved);
        check('import rebuilds current result count', host.querySelector('.sheet-find-result').textContent, '1 matches');
        check('find/replace never call host callbacks', callbacks, []);
        shell.Dispose();
        check('disposed Ctrl+F no longer steals shell keys', key(host, 'f'), false);
    } catch (error) { lines.push('FAIL threw ' + String(error?.stack ?? error).split('\n').slice(0, 3).join(' | ')); }
    finally { session.closeAll(); }
    return lines;
};

checks['shell-find-shared-session'] = async () => {
    const { lines, check } = makeLog();
    const first = await mount(), second = await mount(), bare = mountGrid();
    const input = (host, selector) => host.querySelector(selector);
    const typeSearch = (host, text) => {
        const field = input(host, '.sheet-find-input');
        field.value = text;
        field.dispatchEvent(new Event('input', { bubbles: true }));
    };
    try {
        first.grid.SetCell(0, 0, 'rose');
        first.grid.SetCell(0, 1, 'rose');
        second.grid.SetCell(0, 0, 'tulip');
        first.grid.ShowFindReplaceDialog();
        first.shell.ShowFindReplace();
        check('grid and shell entrypoints share one visible search bar',
            [first.host.querySelectorAll('.sheet-find-bar').length, first.gridHost.querySelectorAll('.sheet-find-bar').length,
                first.host.querySelector('.sheet-find-bar').dataset.hidden], [1, 0, 'false']);
        check('shell Ctrl+H reuses the open grid search',
            [key(first.host.querySelector('.sheet-toolbar'), 'h'), first.host.querySelectorAll('.sheet-find-bar').length],
            [true, 1]);
        typeSearch(first.host, 'rose');
        check('first grid has two matches while the second grid has no open search',
            [input(first.host, '.sheet-find-result').textContent, second.host.querySelector('.sheet-find-bar').dataset.hidden],
            ['2 matches', 'true']);
        second.shell.ShowFindReplace();
        typeSearch(second.host, 'tulip');
        check('independent grid sessions keep their own query and result',
            [input(first.host, '.sheet-find-input').value, input(first.host, '.sheet-find-result').textContent,
                input(second.host, '.sheet-find-result').textContent], ['rose', '2 matches', '1 matches']);
        first.shell.Commands.Run('edit.findNext');
        check('next selects one match and reports its position',
            [first.grid.GetSelectionRef(), input(first.host, '.sheet-find-result').textContent], ['B1', '2 of 2 matches']);
        first.host.querySelector('[data-sheet-icon="find-previous"]').click();
        check('previous returns to the earlier match',
            [first.grid.GetSelectionRef(), input(first.host, '.sheet-find-result').textContent], ['A1', '1 of 2 matches']);
        first.grid.SetReadOnly(true);
        check('read-only search retains matches while replacements cannot be clicked',
            [input(first.host, '.sheet-find-result').textContent, first.host.querySelector('[data-sheet-icon="replace-one"]').disabled,
                first.host.querySelector('[data-sheet-icon="replace-run"]').disabled],
            ['1 of 2 matches', true, true]);
        first.grid.SetReadOnly(false);
        const replacement = input(first.host, '.sheet-replace-input');
        replacement.value = 'lily';
        first.shell.Commands.Run('edit.replaceAll');
        check('replace all updates only the first grid',
            [first.grid.GetCell(0, 0).value, first.grid.GetCell(0, 1).value, second.grid.GetCell(0, 0).value],
            ['lily', 'lily', 'tulip']);
        first.grid.Undo();
        check('undo restores both replacements together',
            [first.grid.GetCell(0, 0).value, first.grid.GetCell(0, 1).value, input(first.host, '.sheet-find-result').textContent],
            ['rose', 'rose', '2 matches']);
        first.grid.SetCell(0, 1, 'Rose');
        const regex = input(first.host, '.sheet-find-regex'), matchCase = input(first.host, '.sheet-find-case');
        regex.checked = true;
        regex.dispatchEvent(new Event('change', { bubbles: true }));
        matchCase.checked = true;
        matchCase.dispatchEvent(new Event('change', { bubbles: true }));
        typeSearch(first.host, '^rose$');
        check('case-sensitive regular expression searches only matching values',
            input(first.host, '.sheet-find-result').textContent, '1 matches');
        matchCase.checked = false;
        matchCase.dispatchEvent(new Event('change', { bubbles: true }));
        check('turning off match case includes both cells',
            input(first.host, '.sheet-find-result').textContent, '2 matches');
        regex.checked = false;
        regex.dispatchEvent(new Event('change', { bubbles: true }));
        first.grid.SetCell(0, 1, 'rose');
        first.grid.SetCell(1, 0, '=1+2');
        typeSearch(first.host, '=1+2');
        const lookIn = input(first.host, '.sheet-find-look-in');
        lookIn.value = 'formulas';
        lookIn.dispatchEvent(new Event('change', { bubbles: true }));
        first.grid.SelectCell(1, 0);
        replacement.value = '=1+4';
        first.host.querySelector('[data-sheet-icon="replace-one"]').click();
        check('replace current edits formula text and recalculates',
            [first.grid.GetRange('A2')[0][0].formula, first.grid.GetCell(1, 0).value], ['=1+4', 5]);
        first.grid.Undo();
        check('formula replacement has one undo step',
            [first.grid.GetRange('A2')[0][0].formula, first.grid.GetCell(1, 0).value], ['=1+2', 3]);
        const range = input(first.host, '.sheet-find-range'), scope = input(first.host, '.sheet-find-scope');
        scope.value = 'workbook';
        scope.dispatchEvent(new Event('change', { bubbles: true }));
        range.value = 'A1:A2';
        range.dispatchEvent(new Event('input', { bubbles: true }));
        check('invalid workbook range reports an error instead of replacing',
            input(first.host, '.sheet-find-result').textContent.startsWith('Find failed:'), true);
        range.value = '';
        range.dispatchEvent(new Event('input', { bubbles: true }));
        first.grid.AddSheet('Second');
        first.grid.SetCell(0, 0, '=1+2');
        check('workbook-scope formula search tracks matches on both sheets',
            input(first.host, '.sheet-find-result').textContent, '2 matches');
        first.host.querySelector('[data-sheet-icon="find-run"]').click();
        check('workbook next navigates to the other sheet',
            [first.grid.GetActiveSheetIndex(), first.grid.GetSelectionRef()], [0, 'A2']);
        replacement.value = '=1+4';
        first.host.querySelector('[data-sheet-icon="replace-run"]').click();
        first.grid.SwitchSheet(1);
        check('workbook replace all reaches both formula sheets',
            [first.grid.GetCell(0, 0).value, (first.grid.SwitchSheet(0), first.grid.GetCell(1, 0).value)], [5, 5]);
        first.grid.Undo();
        first.grid.SwitchSheet(1);
        check('one undo restores formulas across sheets',
            [first.grid.GetCell(0, 0).value, (first.grid.SwitchSheet(0), first.grid.GetCell(1, 0).value)], [3, 3]);
        const saved = await first.grid.ExportXlsx();
        first.grid.SetCell(1, 0, 'changed since save');
        await first.grid.ImportXlsx(saved);
        check('reopened workbook rebuilds search on imported formulas',
            input(first.host, '.sheet-find-result').textContent, '2 matches');
        first.grid.SetCell(1, 0, 'no match');
        first.grid.SwitchSheet(1);
        first.grid.SetCell(0, 0, 'no match');
        const beforeNoop = first.grid.GetRevision();
        first.host.querySelector('[data-sheet-icon="replace-run"]').click();
        check('stale matches cannot replace after both cells change',
            [first.grid.GetRevision(), first.grid.GetCell(0, 0).value], [beforeNoop, 'no match']);
        first.shell.Dispose();
        check('unmounting a shell leaves its one bar available on the bare grid',
            [first.gridHost.querySelectorAll('.sheet-find-bar').length, first.host.querySelectorAll('.sheet-find-bar').length],
            [1, 0]);
        first.grid.ShowFindReplaceDialog();
        check('grid entrypoint reopens same query after shell unmount',
            [first.gridHost.querySelector('.sheet-find-input').value, first.gridHost.querySelectorAll('.sheet-find-bar').length],
            ['=1+2', 1]);
        bare.grid.SetCell(0, 0, 'standalone');
        const target = bare.host.querySelector('.spreadsheet-content');
        check('bare grid Ctrl+H opens search', key(target, 'h'), true);
        typeSearch(bare.host, 'standalone');
        check('bare grid uses the same search behavior',
            [bare.host.querySelectorAll('.sheet-find-bar').length, input(bare.host, '.sheet-find-result').textContent],
            [1, '1 matches']);
        bare.grid.SetCell(0, 1, 'standalone');
        const searchField = input(bare.host, '.sheet-find-input');
        searchField.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
        check('Enter from search advances and keeps focus in the bar',
            [bare.grid.GetSelectionRef(), document.activeElement === searchField], ['B1', true]);
        searchField.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', shiftKey: true, bubbles: true, cancelable: true }));
        check('Shift+Enter goes to the previous match', bare.grid.GetSelectionRef(), 'A1');
        searchField.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }));
        check('Escape closes only this search', bare.host.querySelector('.sheet-find-bar').dataset.hidden, 'true');
        check('bare Ctrl+F reopens without duplicating the bar',
            [key(target, 'f'), bare.host.querySelectorAll('.sheet-find-bar').length], [true, 1]);
        input(bare.host, '.sheet-replace-input').value = 'updated';
        bare.host.querySelector('[data-sheet-icon="replace-run"]').click();
        check('bare grid replacement changes both matches through one bar',
            [bare.grid.GetCell(0, 0).value, bare.grid.GetCell(0, 1).value], ['updated', 'updated']);
        bare.grid.Undo();
        check('bare grid replacement undoes in one step',
            [bare.grid.GetCell(0, 0).value, bare.grid.GetCell(0, 1).value], ['standalone', 'standalone']);
    } catch (error) { lines.push('FAIL threw ' + String(error?.stack ?? error).split('\n').slice(0, 3).join(' | ')); }
    finally { bare.close(); first.closeAll(); second.closeAll(); }
    return lines;
};

checks['shell-function-help'] = async () => {
    const { lines, check } = makeLog();
    const session = await mount();
    const { grid, host, gridHost, shell } = session;
    const input = gridHost.querySelector('.spreadsheet-edit-input');
    const hint = gridHost.querySelector('.spreadsheet-formula-help');
    const chooser = () => host.querySelector('[data-sheet-function-chooser]');
    const put = (text, caret = text.length) => {
        input.value = text;
        input.setSelectionRange(caret, caret);
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new KeyboardEvent('keyup', { key: 'ArrowLeft', bubbles: true }));
    };
    const press = (text) => {
        const event = new KeyboardEvent('keydown', { key: text, bubbles: true, cancelable: true });
        input.dispatchEvent(event);
        return event.defaultPrevented;
    };
    try {
        grid.SelectCell(4, 4);
        const revision = grid.GetRevision();
        host.querySelector('[data-sheet-icon="insert-function"]').click();
        const sumChoice = chooser().querySelector('[data-sheet-function="SUM"]');
        check('function choice leaves its description and argument labels readable', sumChoice.scrollHeight <= sumChoice.clientHeight, true);
        const search = chooser().querySelector('.sheet-function-search');
        search.value = 'numbers';
        search.dispatchEvent(new Event('input', { bubbles: true }));
        check('search finds the supported SUM function by description',
            chooser().querySelector('[data-sheet-function="SUM"]').style.display, '');
        search.value = 'sUm';
        search.dispatchEvent(new Event('input', { bubbles: true }));
        check('function-name search is case-insensitive', chooser().querySelector('[data-sheet-function="SUM"]').style.display, '');
        search.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
        check('chooser inserts SUM into blank active edit with caret inside arguments, not committed',
            [input.value, input.selectionStart, grid.IsEditing(), grid.GetCell(4, 4).value, grid.GetRevision()],
            ['=SUM()', 5, true, null, revision]);
        put('=SUM(2,3)', 7);
        check('active SUM argument help follows caret after comma',
            [hint.style.display !== 'none', hint.textContent.includes('number2')], [true, true]);
        press('Enter');
        check('normal Enter calculates selected function once', [grid.GetCell(4, 4).value, grid.IsEditing()], [5, false]);
        grid.Undo();
        check('undo restores blank cell after function entry', grid.GetCell(4, 4).value, null);
        grid.SetCell(4, 4, '=SUM(1,2)');
        grid.SelectCell(4, 4);
        grid.StartEdit();
        input.setSelectionRange(5, 6);
        const before = grid.GetRevision();
        host.querySelector('[data-sheet-icon="insert-function"]').click();
        chooser().querySelector('[data-sheet-function="IF"]').click();
        check('insertion replaces only selected text in existing formula, without committing',
            [input.value, input.selectionStart, grid.GetCell(4, 4).formula, grid.GetRevision()],
            ['=SUM(IF(),2)', 8, '=SUM(1,2)', before]);
        grid.CancelEdit();
        check('cancel preserves previous formula', grid.GetCell(4, 4).formula, '=SUM(1,2)');
        grid.StartEdit();
        const quoted = '=IF(SUM(1,2)>0,"a,""b"",c",SUM(3,4))';
        put(quoted, quoted.indexOf('SUM(3,4)') + 'SUM(3,'.length);
        check('nested call selects inner SUM argument', hint.textContent.includes('SUM — number2'), true);
        put(quoted, quoted.indexOf('b'));
        check('quoted and escaped commas do not show function help inside text', hint.style.display, 'none');
        put(quoted, quoted.indexOf('""b') + 1);
        check('caret between escaped quote characters remains inside quoted text', hint.style.display, 'none');
        put(quoted, quoted.lastIndexOf('SUM(3'));
        check('outer IF returns after string with escaped quote and commas',
            hint.textContent.includes('IF — value_if_false'), true);
        const structured = '=SUM(Table1[[A,B]],2)';
        put(structured, structured.lastIndexOf('2'));
        check('structured-reference comma leaves SUM at its second argument',
            hint.textContent.includes('SUM — number2'), true);
        put('=IF({1,2},3,4)', 8);
        check('an array constant comma does not advance the function argument', hint.textContent.includes('IF — test:'), true);
        put('=IF({1,2},3,4)', 10);
        check('the comma after an array constant advances exactly one argument', hint.textContent.includes('IF — value_if_true:'), true);
        const partialStructured = '=SUM(Table1[[A,SU';
        put(partialStructured);
        press('Tab');
        check('Tab does not replace a structured-reference column prefix with a function', grid.GetCell(4, 4).formula, partialStructured);
        grid.Undo();
        grid.SelectCell(4, 4);
        grid.StartEdit();
        put('=LEF');
        check('Tab completes unique function token without committing', [press('Tab'), input.value, grid.IsEditing()], [true, '=LEFT(', true]);
        put('=SUM');
        check('Tab completes exact function even with SUMIF candidates', [press('Tab'), input.value], [true, '=SUM(']);
        put('=SEQUEN');
        check('Tab completes a unique dynamic-array function token', [press('Tab'), input.value], [true, '=SEQUENCE(']);
        put('=TODAY');
        check('Tab completes an exact no-argument function', [press('Tab'), input.value, hint.textContent.includes('No arguments')],
            [true, '=TODAY(', true]);
        put('=SUMIFS(A1:A2,B1:B2,">0",A1:A2,');
        check('SUMIFS later pair requests a criterion, not another range',
            hint.textContent.includes('SUMIFS — criteria...'), true);
        put('=IFS(FALSE,1,TRUE,');
        check('IFS later pair requests its result', hint.textContent.includes('IFS — value...'), true);
        grid.CancelEdit();
        grid.StartEdit('=SUM(7,8)');
        host.querySelector('[data-sheet-icon="insert-function"]').click();
        chooser().querySelector('[data-sheet-function-cancel]').click();
        check('Cancel keeps the existing pending entry untouched',
            [grid.IsEditing(), input.value, grid.GetCell(4, 4).formula], [true, '=SUM(7,8)', '=SUM(1,2)']);
        host.querySelector('[data-sheet-icon="insert-function"]').click();
        grid.SetEditText('=SUM(9,9)');
        chooser().querySelector('[data-sheet-function="IF"]').click();
        check('stale chooser cannot overwrite changed pending entry', input.value, '=SUM(9,9)');
        grid.CancelEdit();
        host.querySelector('[data-sheet-icon="insert-function"]').click();
        const saved = await grid.ExportXlsx();
        await grid.ImportXlsx(saved);
        chooser().querySelector('[data-sheet-function="IF"]').click();
        check('workbook replacement invalidates the chooser without changing imported formula',
            [grid.GetCell(4, 4).formula, grid.IsEditing()], ['=SUM(1,2)', false]);
        grid.SetCell(0, 0, '1');
        grid.SetCell(1, 0, '2');
        grid.SetCell(0, 1, '3');
        grid.SetCell(1, 1, '0');
        grid.SelectCell(6, 6);
        host.querySelector('[data-sheet-icon="insert-function"]').click();
        const criteriaSearch = chooser().querySelector('.sheet-function-search');
        criteriaSearch.value = 'COUNTIFS';
        criteriaSearch.dispatchEvent(new Event('input', { bubbles: true }));
        const countifsChoice = chooser().querySelector('[data-sheet-function="COUNTIFS"]');
        check('new criteria function is searchable and shows the repeated pair', 
            [countifsChoice.style.display, countifsChoice.querySelector('strong').textContent.includes('criteria_range..., criteria...')],
            ['', true]);
        countifsChoice.click();
        check('COUNTIFS insertion stays pending with caret inside', 
            [input.value, input.selectionStart, grid.GetCell(6, 6).value], ['=COUNTIFS()', 10, null]);
        put('=COUNTIFS(A1:A2,">0",B1:B2,">0",');
        check('fifth COUNTIFS argument cycles to criteria range', hint.textContent.includes('COUNTIFS — criteria_range...'), true);
        put('=COUNTIFS(A1:A2,">0",B1:B2,">0",A1:A2,');
        check('sixth COUNTIFS argument cycles to criteria', hint.textContent.includes('COUNTIFS — criteria...'), true);
        const nestedCriteria = '=IF(COUNTIFS(A1:A2,">0",B1:B2,">0"),1,0)';
        put(nestedCriteria, nestedCriteria.indexOf('B1:B2'));
        check('caret inside nested COUNTIFS sees the second range', hint.textContent.includes('COUNTIFS — criteria_range...'), true);
        put('=COUNTIFS(A1:A2,">0",B1:B2,">0")');
        press('Enter');
        check('chooser-inserted criteria formula evaluates', 
            [grid.GetCell(6, 6).formula, grid.GetCell(6, 6).value], ['=COUNTIFS(A1:A2,">0",B1:B2,">0")', 1]);
        grid.Undo();
        check('criteria formula insertion is one undo step', grid.GetCell(6, 6).value, null);
        grid.Redo();
        const criteriaSaved = await grid.ExportXlsx();
        await grid.ImportXlsx(criteriaSaved);
        check('criteria formula survives save and reopen',
            [grid.GetCell(6, 6).formula, grid.GetCell(6, 6).value], ['=COUNTIFS(A1:A2,">0",B1:B2,">0")', 1]);
        grid.SelectCell(6, 7);
        host.querySelector('[data-sheet-icon="insert-function"]').click();
        chooser().querySelector('[data-sheet-function="TODAY"]').click();
        check('no-argument function insertion shows no-argument hint',
            [input.value, hint.textContent.includes('TODAY — No arguments')], ['=TODAY()', true]);
        press('Enter');
        check('no-argument function evaluates after insertion', 
            [grid.GetCell(6, 7).formula, typeof grid.GetCell(6, 7).value], ['=TODAY()', 'number']);
        grid.SelectCell(7, 8);
        grid.StartEdit('=ROU');
        press('Tab');
        check('ambiguous ROUND prefix is not silently completed',
            [grid.IsEditing(), grid.GetCell(7, 8).formula], [false, '=ROU']);
        grid.Undo();
        const { createFormulaBar } = await import('/spreadsheet/Toolbar.js');
        const barHost = document.createElement('div');
        host.appendChild(barHost);
        const changed = [];
        const bar = createFormulaBar(() => {}, (value) => changed.push(value), () => {});
        bar.Mount(barHost);
        const barInput = barHost.querySelector('.formula-input');
        const barHint = barHost.querySelector('.sheet-formula-argument-help');
        const barText = '=IF(1,"a,b",SUM(3,4))';
        bar.SetFormula(barText);
        barInput.setSelectionRange(barText.indexOf('4'), barText.indexOf('4'));
        barInput.dispatchEvent(new KeyboardEvent('keyup', { key: 'ArrowLeft', bubbles: true }));
        check('standalone formula bar shows nested argument help without changing host callback',
            [barHint.textContent.includes('SUM — number2'), changed], [true, []]);
        const barCriteria = '=COUNTIFS(A1:A2,">0",B1:B2,">0",A1:A2,';
        bar.SetFormula(barCriteria);
        barInput.setSelectionRange(barCriteria.length, barCriteria.length);
        barInput.dispatchEvent(new KeyboardEvent('keyup', { key: 'ArrowLeft', bubbles: true }));
        check('standalone formula bar cycles repeated criteria labels without invoking host callback',
            [barHint.textContent.includes('COUNTIFS — criteria...'), changed], [true, []]);
        barInput.value = '=SUM(8,9)';
        barInput.dispatchEvent(new Event('input', { bubbles: true }));
        check('formula bar still owns its original edit callback', changed, ['=SUM(8,9)']);
        barHost.remove();
        grid.SetReadOnly(true);
        check('read-only disables chooser command and button',
            [shell.Commands.State('insert.function').Enabled, host.querySelector('[data-sheet-icon="insert-function"]').disabled], [false, true]);
        grid.SetReadOnly(false);
        host.querySelector('[data-sheet-icon="insert-function"]').click();
        shell.Dispose();
        check('dispose removes open function chooser', !!chooser(), false);
    } catch (error) { lines.push('FAIL threw ' + String(error?.stack ?? error).split('\n').slice(0, 3).join(' | ')); }
    finally { session.closeAll(); }
    return lines;
};

checks['shell-selection-status'] = async () => {
    const { lines, check } = makeLog();
    const session = await mount(), other = await mount();
    const { grid, host, shell } = session;
    const fields = ['count', 'numericCount', 'sum', 'average', 'min', 'max'];
    const field = (key) => host.querySelector(`[data-sheet-selection-stat="${key}"]`);
    const shown = () => fields.map((key) => field(key)?.textContent);
    try {
        check('selection status starts empty without fabricated numeric values',
            shown(), ['Count: 0', 'Numeric count: 0', '', '', '', '']);
        grid.SetCell(0, 0, '10');
        grid.SetCell(1, 0, '-4');
        grid.SetCell(2, 0, '=A1*2');
        grid.SetCell(0, 1, 'text');
        grid.SetCell(1, 1, 'TRUE');
        grid.SetCell(2, 1, '=1/0');
        grid.SelectRange(0, 0, 2, 1);
        check('mixed numbers, formulas, text, boolean and error use GetStatus semantics',
            shown(), ['Count: 6', 'Numeric count: 3', 'Sum: 26', 'Average: ' + String(26 / 3), 'Min: -4', 'Max: 20']);
        grid.SetCell(0, 2, 'Group');
        grid.SetCell(1, 2, 'Show');
        grid.SetCell(2, 2, 'Hide');
        grid.SetAutoFilter('A1:C3');
        grid.SetFilter(2, { values: ['Show'] });
        check('AutoFilter-hidden rows are excluded from shell statistics',
            [field('count').textContent, field('numericCount').textContent, field('sum').textContent],
            ['Count: 4', 'Numeric count: 2', 'Sum: 6']);
        grid.ClearAutoFilter();
        host.style.width = '390px';
        const region = host.querySelector('.sheet-selection-status');
        check('selection statistics wrap inside a 390px shell',
            region.scrollWidth <= region.clientWidth && [...region.children].every((node) => node.getBoundingClientRect().right <= region.getBoundingClientRect().right + 1), true);
        check('another shell remains on its own empty grid',
            other.host.querySelector('[data-sheet-selection-stat="sum"]').textContent, '');
        grid.SetCell(0, 0, '5');
        check('committed edit recalculates formulas and status',
            shown(), ['Count: 6', 'Numeric count: 3', 'Sum: 11', 'Average: ' + String(11 / 3), 'Min: -4', 'Max: 10']);
        grid.Undo();
        check('undo refreshes the unchanged selection', [field('sum').textContent, field('max').textContent], ['Sum: 26', 'Max: 20']);
        grid.Redo();
        check('redo refreshes the unchanged selection', field('sum').textContent, 'Sum: 11');
        grid.SelectCell(0, 0);
        grid.StartEdit('99');
        check('pending edit does not show the uncommitted value',
            [field('sum').textContent, grid.GetStatus().sum, grid.GetCell(0, 0).value], ['Sum: 5', 5, 5]);
        grid.CancelEdit();
        grid.SetReadOnly(true);
        check('read-only keeps status visible', [field('sum').textContent, field('count').textContent], ['Sum: 5', 'Count: 1']);
        grid.SetReadOnly(false);
        grid.SelectRange(0, 1, 2, 1);
        check('text/boolean/error-only range hides undefined numeric results',
            [shown(), fields.slice(2).every((key) => getComputedStyle(field(key)).display === 'none')],
            [['Count: 3', 'Numeric count: 0', '', '', '', ''], true]);
        grid.SelectCell(20, 20);
        check('empty selection counts zero and has no sum', shown(), ['Count: 0', 'Numeric count: 0', '', '', '', '']);
        grid.AddSheet('Other');
        grid.SetCell(0, 0, '7');
        grid.SelectCell(0, 0);
        check('new active sheet reports its own values', field('sum').textContent, 'Sum: 7');
        grid.SwitchSheet(0);
        grid.SelectCell(0, 0);
        check('switching back restores original sheet statistics', field('sum').textContent, 'Sum: 5');
        const saved = await grid.ExportXlsx();
        grid.SetCell(0, 0, '100');
        await grid.ImportXlsx(saved);
        grid.SwitchSheet(0);
        grid.SelectCell(0, 0);
        check('import replaces the workbook stats, not a cached pre-import sum', field('sum').textContent, 'Sum: 5');
        const oldRegion = host.querySelector('.sheet-selection-status');
        shell.Dispose();
        grid.SetCell(0, 0, '12');
        check('disposed shell does not update its detached status', oldRegion.querySelector('[data-sheet-selection-stat="sum"]').textContent, 'Sum: 5');
        shell.Mount(host);
        check('remounted shell reads current committed stats', field('sum').textContent, 'Sum: 12');
        grid.GoTo('A60000');
        grid.SetCell(59999, 0, '3');
        grid.GoTo('A:A');
        check('sparse whole-column selection uses cached GetStatus aggregation',
            [grid.GetModel().RowCount >= 60000, field('count').textContent, field('numericCount').textContent, field('sum').textContent],
            [true, 'Count: 4', 'Numeric count: 4', 'Sum: 35']);
        check('foreign shell still has no status from this workbook', other.host.querySelector('[data-sheet-selection-stat="count"]').textContent, 'Count: 0');
    } catch (error) { lines.push('FAIL threw ' + String(error?.stack ?? error).split('\n').slice(0, 3).join(' | ')); }
    finally { session.closeAll(); other.closeAll(); }
    return lines;
};

checks['shell-insert-chart'] = async () => {
    const { lines, check } = makeLog();
    const session = await mount();
    const { grid, host, gridHost, shell } = session;
    const chartDialog = () => document.querySelector('[data-testid="xl-dlg-insert-chart"]');
    const field = (name) => chartDialog()?.querySelector(`[data-testid="dlg-field-${name}"]`);
    try {
        for (const [row, cells] of [[0, ['Quarter', 'Revenue']], [1, ['Q1', '12']], [2, ['Q2', '25']]])
            cells.forEach((text, col) => grid.SetCell(row, col, text));
        grid.SelectRange(0, 0, 2, 1);
        menu(host, 'insert', 'insert-chart');
        await frame();
        check('Insert menu opens a source-aware shared chart dialog',
            [!!chartDialog(), field('range')?.value, field('headers')?.checked, field('categories')?.checked],
            [true, 'A1:B3', true, true]);
        const title = field('title');
        title.value = 'Revenue by quarter';
        title.dispatchEvent(new Event('input', { bubbles: true }));
        const kind = field('type');
        kind.value = 'pie';
        kind.dispatchEvent(new Event('change', { bubbles: true }));
        chartDialog().querySelector('[data-testid="xl-dlg-insert-chart-apply"]').click();
        await frame(); await frame();
        check('dialog inserts a rendered one-series pie without routing to an unwired host action',
            [chartDialog(), grid.GetDrawings().filter((item) => item.kind === 'chart').length,
                !!gridHost.querySelector('.spreadsheet-object-chart.spreadsheet-object-drawn svg')],
            [null, 1, true]);
        grid.Undo();
        check('shell insertion undoes in one transaction', grid.GetDrawings().length, 0);
        grid.Redo();
        check('shell insertion redoes in one transaction', grid.GetDrawings().length, 1);
        menu(host, 'insert', 'insert-chart');
        await frame();
        grid.SelectCell(1, 1);
        chartDialog().querySelector('[data-testid="xl-dlg-insert-chart-apply"]').click();
        await frame();
        check('stale chart dialog remains open and explains selection change',
            [!!chartDialog(), /changed/.test(field('error')?.textContent ?? ''), grid.GetDrawings().length],
            [true, true, 1]);
        chartDialog().querySelector('[data-testid="dlg-cancel"]').click();
        grid.SetReadOnly(true);
        check('viewer disables the insert-chart command', shell.Commands.State('insert.chart').Enabled, false);
        grid.SetReadOnly(false);
    } catch (error) { lines.push('FAIL threw ' + String(error?.stack ?? error).split('\n').slice(0, 3).join(' | ')); }
    finally { session.closeAll(); }
    return lines;
};
