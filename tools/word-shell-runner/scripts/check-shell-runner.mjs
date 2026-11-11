// Browser regression for the shared OfficeUi renderer and genuine WordUi dialogs.
// Default: build Fable into a private temporary directory, serve via Vite, use
// one Chromium page. For a separately built/running surface set SHELL_RUNNER_URL
// (and SHELL_RUNNER_OUTDIR for the Vite entry when building it yourself).
// SHELL_RUNNER_OUTDIR may point to an existing private output directory; unlike
// the automatically created one, it is never removed by this script.
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import os from 'node:os';
import fs from 'node:fs/promises';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { chromium } = createRequire(path.join(root, '../browser-runner/package.json'))('playwright');
const outDir = path.resolve(process.env.SHELL_RUNNER_ARTIFACTS ?? path.join(root, '../../artifacts/shell-runner'));
const externalUrl = process.env.SHELL_RUNNER_URL;
const port = Number(process.env.SHELL_RUNNER_PORT ?? 5180);
const outputIsTemporary = !externalUrl && !process.env.SHELL_RUNNER_OUTDIR;
const fableOutDir = outputIsTemporary
    ? await fs.mkdtemp(path.join(os.tmpdir(), 'rbt-shell-fable-'))
    : path.resolve(process.env.SHELL_RUNNER_OUTDIR ?? path.join(root, 'src'));
const cases = [];
const pageErrors = [];
let server;
let browser;
await fs.mkdir(outDir, { recursive: true });

function equal(actual, expected, label) {
    if (actual !== expected) throw new Error(`${label}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
}
function ok(value, label) {
    if (!value) throw new Error(label);
}
function run(command, args) {
    return new Promise((resolve, reject) => {
        const child = spawn(command, args, { cwd: root, stdio: 'inherit' });
        child.once('error', reject);
        child.once('exit', code => code === 0 ? resolve() : reject(new Error(`${command} exited ${code}`)));
    });
}
async function awaitServer(url) {
    for (let i = 0; i < 80; i++) {
        if (server.exitCode !== null) throw new Error(`vite exited ${server.exitCode}`);
        try {
            const response = await fetch(url);
            if (response.ok) return;
        } catch { /* server not listening yet */ }
        await new Promise(resolve => setTimeout(resolve, 250));
    }
    throw new Error(`vite did not start at ${url}`);
}

try {
    let url = externalUrl;
    if (!url) {
        await run('dotnet', ['tool', 'restore']);
        await run('dotnet', ['fable', '.', '--outDir', fableOutDir, '--lang', 'JavaScript']);
        server = spawn(path.join(root, '../browser-runner/node_modules/.bin/vite'),
            ['--host', '127.0.0.1', '--port', String(port), '--strictPort'],
            { cwd: root, env: { ...process.env, SHELL_RUNNER_OUTDIR: fableOutDir }, stdio: ['ignore', 'pipe', 'pipe'] });
        server.stdout.on('data', chunk => process.stdout.write(chunk));
        server.stderr.on('data', chunk => process.stderr.write(chunk));
        url = `http://127.0.0.1:${port}/`;
        await awaitServer(url);
    }

    browser = await chromium.launch();
    const page = await browser.newPage({ viewport: { width: 900, height: 850 } });
    page.on('pageerror', error => pageErrors.push(error.message));
    await page.goto(url);
    await page.waitForFunction(() => typeof window.rbDialog?.editorText === 'function', null, { timeout: 15000 });
    const byId = id => page.locator(`[data-testid="${id}"]`);
    const result = byId('result');
    const card = page.locator('.rb-dlg-card');
    async function closed(label) {
        await card.waitFor({ state: 'detached', timeout: 3000 });
        equal(await page.getByRole('dialog').count(), 0, `${label}: modal removed`);
    }
    async function openAndSnap(testId, snapshotName, fill, expectedResult) {
        await byId(testId).click();
        await card.waitFor({ state: 'visible' });
        if (fill) await fill();
        await page.screenshot({ path: path.join(outDir, `${snapshotName}.png`), fullPage: true });
        await page.locator('.rb-dlg-button-primary').click();
        await closed(snapshotName);
        const actual = await result.textContent();
        cases.push({ testId, snapshotName, result: actual });
        equal(actual, expectedResult, snapshotName);
    }

    // Preserve the preexisting real WordUi spec scenarios.
    await openAndSnap('btn-find', 'find-replace', async () => {
        await byId('dlg-field-find').fill('hello');
        await byId('dlg-field-replace').fill('world');
        await byId('dlg-field-case').check();
    }, 'replace:hello->world');
    await openAndSnap('btn-table', 'insert-table', async () => {
        await byId('dlg-field-rows').fill('4');
        await byId('dlg-field-cols').fill('5');
    }, 'table:4x5');
    await openAndSnap('btn-link', 'insert-link', async () => {
        await byId('dlg-field-url').fill('https://example.com');
    }, 'link:https://example.com');
    await openAndSnap('btn-page-setup', 'page-setup', async () => {
        await byId('dlg-field-orientation').selectOption('landscape');
    }, 'page:landscape ml=72 mr=72 mt=72 mb=72');
    await openAndSnap('btn-image', 'insert-image', async () => {
        await byId('dlg-field-src').fill('/img/cover.png');
    }, 'image:/img/cover.png');

    const attempts = () => page.evaluate(() => window.rbDialog.attempts());
    const selected = label => page.getByRole('tab', { name: label }).getAttribute('aria-selected');
    const focused = id => page.evaluate(testId => document.activeElement?.getAttribute('data-testid') === testId, id);
    const waitFocus = id => page.waitForFunction(testId =>
        document.activeElement?.getAttribute('data-testid') === testId, id, { timeout: 2000 });
    const name = byId('dlg-field-name');
    const firstTab = page.getByRole('tab', { name: 'Basics' });
    const detailsTab = page.getByRole('tab', { name: 'Details' });

    await byId('btn-link').click();
    await waitFocus('dlg-field-url');
    await page.keyboard.press('Escape');
    await closed('non-tab Escape');
    await waitFocus('btn-link');

    await byId('btn-fixture').click();
    await card.waitFor({ state: 'visible' });
    equal(await page.getByRole('dialog', { name: 'All field kinds' }).count(), 1, 'labelled accessible dialog');
    equal(await page.getByRole('dialog').getAttribute('aria-modal'), 'true', 'modal ARIA');
    ok((await page.getByRole('dialog').getAttribute('aria-describedby')) !== null, 'subtitle described by ARIA');
    equal(await page.getByRole('tablist').count(), 1, 'tab navigation has tablist role');
    equal(await page.getByRole('textbox', { name: 'Name' }).count(), 1, 'text field has accessible label');
    equal(await page.getByRole('spinbutton', { name: 'Quantity' }).count(), 1, 'number field has accessible label');
    equal(await page.getByText('Identity', { exact: true }).count(), 1, 'section header renders');
    equal(await byId('dlg-field-ignored').count(), 0, 'top-level Fields ignored when Tabs exist');
    equal(await selected('Basics'), 'true', 'initial tab selected');
    equal(await byId('dlg-field-size').inputValue(), 'large', 'non-first initial select option visible');
    equal(await byId('dlg-field-status').textContent(), 'Ready', 'status field renders');
    equal(await byId('dlg-field-name').inputValue(), 'Initial', 'text field initial value');
    await page.waitForFunction(() =>
        document.activeElement?.getAttribute('role') === 'tab'
        && document.activeElement?.textContent === 'Basics');
    await page.evaluate(() => document.querySelector('[data-testid="btn-image"]').focus());
    await page.waitForFunction(() => document.activeElement?.closest('[role="dialog"]') !== null);

    // The selected tab precedes the fields in tab order. From that first
    // focusable control, reverse Tab wraps to the final button and back.
    await firstTab.click();
    await page.keyboard.press('Shift+Tab');
    ok(await focused('fixture-apply'), 'reverse Tab wraps within modal');
    await page.keyboard.press('Tab');
    ok(await firstTab.evaluate(node => document.activeElement === node), 'forward Tab wraps within modal');
    await page.keyboard.press('Escape');
    await closed('Escape');
    equal((await attempts()).length, 0, 'Escape never applies');
    await waitFocus('btn-fixture');
    // Focused secondary Enter must activate Cancel, never the primary action.
    await byId('btn-fixture').click();
    await firstTab.click();
    await page.keyboard.press('Shift+Tab');
    await page.keyboard.press('Shift+Tab');
    await page.keyboard.press('Shift+Tab');
    await waitFocus('dlg-cancel');
    await page.keyboard.press('Enter');
    await closed('secondary Enter');
    equal((await attempts()).length, 0, 'secondary Enter never submits primary');
    await waitFocus('btn-fixture');

    await byId('btn-fixture').click();
    await name.fill('Reject');
    await byId('dlg-field-quantity').fill('9');
    await byId('dlg-field-enabled').check();
    // Native select keyboard navigation, not selectOption or direct model writes.
    const size = byId('dlg-field-size');
    equal(await size.inputValue(), 'large', 'initial choice survives reopen');
    await size.focus();
    await page.keyboard.press('Home');
    equal(await size.inputValue(), 'small', 'native select Home');
    await page.keyboard.press('ArrowDown');
    equal(await size.inputValue(), 'large', 'native select ArrowDown');
    await firstTab.click();
    await page.keyboard.press('ArrowRight');
    equal(await selected('Details'), 'true', 'right arrow selects next tab');
    await page.keyboard.press('Home');
    equal(await selected('Basics'), 'true', 'Home selects first tab');
    await page.keyboard.press('End');
    equal(await selected('Details'), 'true', 'End selects last tab');
    equal(await byId('dlg-field-name').isVisible(), false, 'inactive tab field hidden');
    equal(await byId('dlg-field-color').isVisible(), true, 'active tab field visible');
    await page.waitForFunction(() => document.querySelector('[data-testid="dlg-field-preview"]')?.textContent === 'Reject / Original / plain');
    equal(await byId('dlg-field-color').inputValue(), '#336699', 'color initial hex');
    equal(await byId('dlg-field-strength').inputValue(), '2', 'slider initial position');
    equal(await page.getByRole('radiogroup').count(), 1, 'radio group accessible role');
    ok(await page.getByRole('radio', { name: 'Plain' }).isChecked(), 'radio initial option');
    await byId('dlg-field-color').fill('#c00000');
    await byId('dlg-field-strength').click();
    await byId('dlg-field-strength').press('End');
    equal(await byId('dlg-field-strength').inputValue(), '10', 'range keyboard End reaches max');
    const notes = byId('dlg-field-notes');
    await notes.fill('first line');
    await notes.press('Enter');
    await notes.type('second line');
    equal(await notes.inputValue(), 'first line\nsecond line', 'textarea Enter inserts newline');
    equal((await attempts()).length, 0, 'textarea Enter does not submit');
    ok(await card.isVisible(), 'textarea Enter keeps dialog open');
    const plain = page.getByRole('radio', { name: 'Plain' });
    const rich = page.getByRole('radio', { name: 'Rich' });
    await plain.click();
    await page.keyboard.press('ArrowRight');
    ok(await rich.isChecked(), 'native radio arrow selects next option');
    await page.waitForFunction(() => document.querySelector('[data-testid="dlg-field-preview"]')?.textContent === 'Reject / first line\nsecond line / rich');
    await detailsTab.click();
    await page.keyboard.press('ArrowLeft');
    equal(await selected('Basics'), 'true', 'left arrow selects previous tab');
    equal(await size.inputValue(), 'large', 'first tab retains select value');
    equal(await name.inputValue(), 'Reject', 'first tab retains text value');
    await detailsTab.click();
    equal(await notes.inputValue(), 'first line\nsecond line', 'second tab retains multiline value');
    ok(await rich.isChecked(), 'second tab retains native radio selection');

    await byId('fixture-apply').click();
    ok(await card.isVisible(), 'false callback keeps dialog open');
    equal((await attempts()).length, 1, 'first callback invoked once');
    const parse = snapshot => Object.fromEntries(snapshot.split('|').map(entry => {
        const split = entry.indexOf('=');
        return [entry.slice(0, split), entry.slice(split + 1)];
    }));
    const first = parse((await attempts())[0]);
    equal(first.ignored, undefined, 'ignored field excluded from callback');
    equal(first.name, 'Reject', 'callback sees edited first tab');
    equal(first.quantity, '9', 'callback sees edited number');
    equal(first.enabled, 'true', 'callback sees edited checkbox');
    equal(first.size, 'large', 'callback select matches visible non-first choice');
    equal(first.color, '#c00000', 'callback sees edited color');
    equal(first.strength, '10', 'slider keyboard max reaches callback');
    equal(first.notes, 'first line\nsecond line', 'callback sees multiline text');
    equal(first.mode, 'rich', 'callback sees keyboard-selected radio');
    await firstTab.click();
    equal(await name.inputValue(), 'Reject', 'false callback retains edits');
    await name.fill('Accept');
    await name.press('Enter');
    await closed('primary Enter');
    const submitted = await attempts();
    equal(submitted.length, 2, 'primary Enter invokes callback exactly once');
    const second = parse(submitted[1]);
    equal(second.name, 'Accept', 'primary Enter submits latest edit');
    equal(second.notes, first.notes, 'primary Enter includes hidden tab values');
    equal(second.size, first.size, 'primary Enter preserves initial non-first select');
    cases.push({ testId: 'btn-fixture', snapshotName: 'all-fields', first, second });
    await waitFocus('btn-fixture');

    // A real Word editor receives the unchanged WordUi link dialog's callback.
    // Cancel must not mutate document/history; submit is exactly one API undo.
    const editor = () => page.evaluate(() => ({
        text: window.rbDialog.editorText(), canUndo: window.rbDialog.editorCanUndo(),
        canRedo: window.rbDialog.editorCanRedo()
    }));
    const before = await editor();
    equal(before.canUndo, false, 'fresh editor has no undo step');
    await byId('btn-editor-link').click();
    await byId('dlg-field-url').fill('https://cancel.example');
    await byId('dlg-cancel').click();
    await closed('editor Cancel');
    equal(JSON.stringify(await editor()), JSON.stringify(before), 'Cancel does not alter Word document or undo history');
    await byId('btn-editor-link').click();
    await byId('dlg-field-url').fill('https://applied.example');
    await byId('dlg-insert-link-ok').click();
    await closed('editor link apply');
    const applied = await editor();
    ok(applied.text.includes('linked') && applied.text !== before.text, 'Word document contains inserted hyperlink text');
    equal(applied.canUndo, true, 'Word edit entered undo history');
    equal(await page.evaluate(() => window.rbDialog.editorUndo()), true, 'Word undo succeeds');
    const undone = await editor();
    equal(undone.text, before.text, 'one undo restores exact original Word document');
    equal(undone.canUndo, false, 'link application created exactly one undo step');
    equal(undone.canRedo, true, 'undo preserves redo');
    cases.push({ testId: 'btn-editor-link', before, applied, undone });

    // A callback can synchronously replace its own dialog; the old primary
    // state must not close or steal focus from the newly opened WordUi spec.
    await byId('btn-fixture').click();
    await byId('fixture-replace').click();
    await byId('word-dlg-insert-link').waitFor({ state: 'visible' });
    await waitFocus('dlg-field-url');
    await byId('dlg-field-url').fill('https://replacement.example');
    await byId('dlg-insert-link-ok').click();
    await closed('callback replacement');
    equal(await result.textContent(), 'replacement-link:https://replacement.example', 'replacement callback remains active');
    equal((await attempts()).length, 2, 'replacement does not submit original dialog');

    // Replacing/closing before deferred focus runs must not focus removed nodes.
    await byId('btn-fixture').click();
    await page.evaluate(() => { window.rbDialog.openFind(); });
    await byId('word-dlg-find-replace').waitFor({ state: 'visible' });
    await waitFocus('dlg-field-find');
    await page.keyboard.press('Escape');
    await closed('replacement Escape');
    await page.evaluate(() => { window.rbDialog.openFixture(); window.rbDialog.close(); });
    await page.waitForTimeout(60);
    await closed('immediate close');
    ok(await page.evaluate(() => document.activeElement?.isConnected), 'pending focus never targets detached control');

    // The editor is outside the shell mount: disposal must restore its focus
    // and remove the modal's document-level focus listener.
    await page.locator('#word-fixture').focus();
    const editorFocus = await page.evaluateHandle(() => document.activeElement);
    await page.evaluate(() => window.rbDialog.openFixture());
    await card.waitFor({ state: 'visible' });
    await page.waitForFunction(() => document.activeElement?.closest('[role="dialog"]') !== null);
    await page.evaluate(() => window.rbDialog.disposeShell());
    await closed('shell disposal');
    ok(await page.evaluate(node => document.activeElement === node, editorFocus), 'disposal restores editor focus');
    await page.keyboard.press('ArrowRight');
    equal(await page.getByRole('dialog').count(), 0, 'disposed dialog does not intercept editor keyboard input');
    await editorFocus.dispose();
    equal(pageErrors.length, 0, `no uncaught browser errors: ${pageErrors.join('; ')}`);

    await fs.writeFile(path.join(outDir, 'summary.json'), JSON.stringify({ ok: true, cases }, null, 2));
    console.log(`shell runner ok: ${cases.length} dialog cases + keyboard, focus, ARIA and Word undo`);
} catch (error) {
    console.error('shell runner FAILED:', error);
    await fs.writeFile(path.join(outDir, 'summary.json'), JSON.stringify({ ok: false, cases, pageErrors, error: String(error) }, null, 2));
    process.exitCode = 1;
} finally {
    if (browser) await browser.close();
    if (server) server.kill('SIGTERM');
    if (outputIsTemporary) await fs.rm(fableOutDir, { recursive: true, force: true });
}
