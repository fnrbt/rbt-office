import { chromium } from 'playwright';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import fs from 'node:fs/promises';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const officeRoot = path.resolve(root, '../..');
const outDir = path.resolve(officeRoot, 'artifacts/browser-perf');
const skipBuild = process.argv.includes('--skip-build');
const port = Number(process.env.RB_BROWSER_PERF_PORT ?? 5183);
const wordFixture = path.resolve(officeRoot, process.env.RB_BROWSER_PERF_WORD_DOC ?? 'tests/test-documents-business/docx4j-invoice.docx');
const xlsxFixture = path.resolve(officeRoot, process.env.RB_BROWSER_PERF_XLSX ?? 'tests/real-world/dashboard-sales-analysis.xlsx');

await fs.mkdir(outDir, { recursive: true });

function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd: options.cwd ?? root, stdio: 'inherit', env: process.env });
    child.on('exit', code => code === 0 ? resolve() : reject(new Error(`${command} exited ${code}`)));
  });
}

function rel(file) {
  return path.relative(officeRoot, file);
}

function fsUrl(file) {
  return `/@fs/${file}`;
}

async function fileExists(file) {
  try {
    await fs.access(file);
    return true;
  } catch {
    return false;
  }
}

async function startVite() {
  const viteBin = path.join(root, 'node_modules/.bin/vite');
  const server = spawn(viteBin, ['--host', '127.0.0.1', '--port', String(port)], {
    cwd: root,
    stdio: ['ignore', 'pipe', 'pipe']
  });

  let serverReady = false;
  let startupOutput = '';
  server.stdout.on('data', chunk => {
    const text = chunk.toString();
    startupOutput += text;
    process.stdout.write(text);
    if (text.includes('Local:')) serverReady = true;
  });
  server.stderr.on('data', chunk => {
    const text = chunk.toString();
    startupOutput += text;
    process.stderr.write(text);
  });

  for (let i = 0; i < 100 && !serverReady; i++) {
    await new Promise(resolve => setTimeout(resolve, 250));
  }
  if (!serverReady) {
    server.kill('SIGKILL');
    throw new Error(`vite server did not start:\n${startupOutput}`);
  }
  return server;
}

async function forceGc(page, cdp) {
  await page.evaluate(() => {
    if (typeof globalThis.gc === 'function') globalThis.gc();
  }).catch(() => {});
  await cdp.send('HeapProfiler.collectGarbage').catch(() => {});
}

async function heapUsage(cdp) {
  const usage = await cdp.send('Runtime.getHeapUsage');
  return {
    usedBytes: Math.round(usage.usedSize),
    totalBytes: Math.round(usage.totalSize)
  };
}

async function measure(page, cdp, name, body) {
  await forceGc(page, cdp);
  const before = await heapUsage(cdp);
  const result = await body();
  await forceGc(page, cdp);
  const after = await heapUsage(cdp);
  return {
    name,
    ...result,
    heapBeforeBytes: before.usedBytes,
    heapAfterBytes: after.usedBytes,
    heapDeltaBytes: after.usedBytes - before.usedBytes
  };
}

async function loadWordFixture(page, cdp) {
  await forceGc(page, cdp);
  const before = await heapUsage(cdp);
  const params = new URLSearchParams({ doc: fsUrl(wordFixture) });
  const started = performance.now();
  await page.goto(`http://127.0.0.1:${port}/?${params.toString()}`);
  await page.waitForFunction(
    () => {
      const status = document.querySelector('#status')?.textContent || '';
      return window.rbWordRunner && (status.startsWith('Loaded ') || status.startsWith('Load failed:'));
    },
    null,
    { timeout: 120000 }
  );
  const statusText = await page.locator('#status').textContent();
  if (!statusText?.startsWith('Loaded ')) {
    throw new Error(`Word fixture did not load: ${statusText}`);
  }
  await forceGc(page, cdp);
  const after = await heapUsage(cdp);
  const layoutSegments = await page.evaluate(() => window.rbWordRunner.getLayoutSegments().length);
  const layoutLines = await page.evaluate(() => window.rbWordRunner.getVbaLayoutLines().length);
  return {
    name: 'word.load-docx',
    fixture: rel(wordFixture),
    durationMs: performance.now() - started,
    layoutSegments,
    layoutLines,
    statusText,
    heapBeforeBytes: before.usedBytes,
    heapAfterBytes: after.usedBytes,
    heapDeltaBytes: after.usedBytes - before.usedBytes
  };
}

async function runWordWorkflows(page, cdp) {
  const rangeQueries = await measure(page, cdp, 'word.range-info-stats-cycle', () =>
    page.evaluate(({ iterations }) => {
      const lines = window.rbWordRunner.getVbaLayoutLines()
        .filter(line => line.BlockId && line.EndOffset >= line.StartOffset)
        .slice(0, 96);
      if (!lines.length) throw new Error('Word range workflow has no layout lines');
      let checksum = 0;
      const started = performance.now();
      for (let i = 0; i < iterations; i++) {
        const line = lines[i % lines.length];
        const info = window.rbWordRunner.getRangeInformation(line.BlockId, line.StartOffset, line.EndOffset);
        const stats = window.rbWordRunner.computeRangeStatistics(line.BlockId, line.StartOffset, line.EndOffset);
        checksum += (info?.wdActiveEndPageNumber ?? 0) + (stats?.Words ?? 0) + (stats?.Characters ?? 0);
      }
      return {
        iterations,
        sampleLines: lines.length,
        durationMs: performance.now() - started,
        checksum
      };
    }, { iterations: 5000 }));

  const findReplace = await measure(page, cdp, 'word.find-replace-noop-cycle', () =>
    page.evaluate(({ iterations }) => {
      let checksum = 0;
      const started = performance.now();
      for (let i = 0; i < iterations; i++) {
        checksum += window.rbWordRunner.findText('invoice', false).length;
        checksum += window.rbWordRunner.findText('total', false).length;
        checksum += window.rbWordRunner.replaceAllText('__missing_browser_perf__', 'x', false);
      }
      return {
        iterations,
        durationMs: performance.now() - started,
        checksum
      };
    }, { iterations: 1000 }));

  const editCycle = await measure(page, cdp, 'word.synthetic-edit-cycle', () =>
    page.evaluate(({ iterations }) => {
      const paragraph = 'browser perf paragraph alpha beta gamma delta epsilon ';
      const content = Array.from({ length: 180 }, (_, i) => `${i}: ${paragraph}${i % 7}`).join('\n');
      window.rbWordRunner.setContent(content);
      let checksum = 0;
      let findTextMs = 0;
      let selectMatchMs = 0;
      let insertTextMs = 0;
      let getTextMs = 0;
      const started = performance.now();
      for (let i = 0; i < iterations; i++) {
        let t = performance.now();
        const matches = window.rbWordRunner.findText('gamma', false);
        findTextMs += performance.now() - t;
        if (!matches.length) throw new Error('Word edit workflow could not find insertion target');
        const match = matches[i % matches.length];
        t = performance.now();
        window.rbWordRunner.selectFindMatch(match);
        selectMatchMs += performance.now() - t;
        t = performance.now();
        window.rbWordRunner.insertTextAt('gamma');
        insertTextMs += performance.now() - t;
        t = performance.now();
        checksum += window.rbWordRunner.getText().length;
        getTextMs += performance.now() - t;
      }
      return {
        iterations,
        durationMs: performance.now() - started,
        findTextMs,
        selectMatchMs,
        insertTextMs,
        getTextMs,
        checksum,
        finalTextLength: window.rbWordRunner.getText().length
      };
    }, { iterations: 240 }));

  return [rangeQueries, findReplace, editCycle];
}

async function runSpreadsheetWorkflows(page, cdp) {
  const moduleBase = fsUrl(path.join(officeRoot, 'dist/spreadsheet'));
  const synthetic = await measure(page, cdp, 'spreadsheet.synthetic-workflows', () =>
    page.evaluate(async ({ moduleBase }) => {
      const Model = await import(`${moduleBase}/Model.js`);
      const Formula = await import(`${moduleBase}/Formula.js`);

      const visibleValues = ['Group 1', 'Group 4', 'Group 7'];
      const visibleGroupNames = new Set(visibleValues);
      let expectedVisibleRows = 0;
      let expectedHiddenRows = 0;
      const rows = [['Group', 'Name', 'Amount', 'Tax', 'Formula', 'Note']];
      for (let row = 1; row <= 1800; row++) {
        const group = `Group ${row % 9}`;
        if (visibleGroupNames.has(group)) expectedVisibleRows++;
        else expectedHiddenRows++;
        rows.push([
          group,
          `Item ${row}`,
          String(row * 3),
          String(row % 17),
          `=C${row + 1}+D${row + 1}`,
          row % 11 === 0 ? 'needle value' : 'plain value'
        ]);
      }
      const csv = rows.map(row => row.join(',')).join('\n');
      const started = performance.now();
      let model = Model.importCSV(csv, 0, 0, Model.createEmpty(2200, 32));
      const importMs = performance.now() - started;

      let checksum = 0;
      let usedBoundsMs = 0;
      let searchCellsMs = 0;
      let uniqueValueCountMs = 0;
      let t = performance.now();
      for (let i = 0; i < 160; i++) {
        let step = performance.now();
        checksum += Model.usedBounds(model)?.[0] ?? 0;
        usedBoundsMs += performance.now() - step;
        step = performance.now();
        checksum += Model.searchCells('needle', false, model).length;
        searchCellsMs += performance.now() - step;
        step = performance.now();
        checksum += Model.columnUniqueValueCount(0, model);
        uniqueValueCountMs += performance.now() - step;
      }
      const queryMs = performance.now() - t;

      t = performance.now();
      const filtered = Model.applyFilterValues(0, visibleValues, model);
      const filterApplyFirstMs = performance.now() - t;

      t = performance.now();
      const filteredRowCount = Model.visibleRowCount(filtered);
      checksum += filteredRowCount;
      const visibleRowsFirstMs = performance.now() - t;

      t = performance.now();
      const projected = Model.applyFilterRowHeights(filtered);
      const hiddenRowCount = projected.FilterHiddenRows.length;
      if (filteredRowCount !== expectedVisibleRows || hiddenRowCount !== expectedHiddenRows) {
        throw new Error(
          `unexpected filter result: visible ${filteredRowCount}/${expectedVisibleRows}, hidden ${hiddenRowCount}/${expectedHiddenRows}`
        );
      }
      checksum += hiddenRowCount;
      const filterProjectionFirstMs = performance.now() - t;

      const hotFilterReadIterations = 400;
      t = performance.now();
      for (let i = 0; i < hotFilterReadIterations; i++) {
        checksum += Model.visibleRowCount(filtered);
      }
      const visibleRowsHotMs = performance.now() - t;

      t = performance.now();
      for (let i = 0; i < hotFilterReadIterations; i++) {
        checksum += Model.applyFilterRowHeights(filtered).FilterHiddenRows.length;
      }
      const filterProjectionHotMs = performance.now() - t;

      const filterToggleIterations = 30;
      let filterToggleApplyMs = 0;
      let filterToggleVisibleRowsMs = 0;
      let filterToggleRowHeightsMs = 0;
      t = performance.now();
      for (let i = 0; i < filterToggleIterations; i++) {
        let step = performance.now();
        const toggled = Model.applyFilterValues(0, visibleValues, model);
        filterToggleApplyMs += performance.now() - step;
        step = performance.now();
        checksum += Model.visibleRowCount(toggled);
        filterToggleVisibleRowsMs += performance.now() - step;
        step = performance.now();
        checksum += Model.applyFilterRowHeights(toggled).FilterHiddenRows.length;
        filterToggleRowHeightsMs += performance.now() - step;
      }
      const filterToggleProjectionMs = performance.now() - t;

      const filterMs =
        filterApplyFirstMs
        + visibleRowsFirstMs
        + filterProjectionFirstMs
        + visibleRowsHotMs
        + filterProjectionHotMs
        + filterToggleProjectionMs;

      let clearSelectionMs = 0;
      let exportCsvMs = 0;
      t = performance.now();
      for (let i = 0; i < 24; i++) {
        const selection = new Model.Selection$(1, [50, 0, 900, 5]);
        let step = performance.now();
        const cleared = Model.clearSelectionCells(selection, model);
        clearSelectionMs += performance.now() - step;
        step = performance.now();
        checksum += Model.exportCSV(cleared).length;
        exportCsvMs += performance.now() - step;
      }
      const clearExportMs = performance.now() - t;

      let recalcAndSpillMs = 0;
      let recalcSearchMs = 0;
      t = performance.now();
      for (let i = 0; i < 30; i++) {
        let step = performance.now();
        const recalculated = Formula.recalcAndSpill(undefined, undefined, model);
        recalcAndSpillMs += performance.now() - step;
        step = performance.now();
        checksum += Model.searchCells('5400', false, recalculated).length;
        recalcSearchMs += performance.now() - step;
      }
      const recalcMs = performance.now() - t;

      return {
        durationMs: performance.now() - started,
        rows: rows.length,
        importMs,
        queryMs,
        usedBoundsMs,
        searchCellsMs,
        uniqueValueCountMs,
        filterMs,
        filterApplyFirstMs,
        visibleRowsFirstMs,
        filterProjectionFirstMs,
        visibleRowsHotMs,
        filterProjectionHotMs,
        filterToggleProjectionMs,
        filterToggleApplyMs,
        filterToggleVisibleRowsMs,
        filterToggleRowHeightsMs,
        hotFilterReadIterations,
        filterToggleIterations,
        filteredRowCount,
        hiddenRowCount,
        clearExportMs,
        clearSelectionMs,
        exportCsvMs,
        recalcMs,
        recalcAndSpillMs,
        recalcSearchMs,
        checksum
      };
    }, { moduleBase }));

  const realXlsx = await measure(page, cdp, 'spreadsheet.real-xlsx-import-query', () =>
    page.evaluate(async ({ moduleBase, xlsxUrl }) => {
      const Importer = await import(`${moduleBase}/XlsxImporter.js`);
      const OfficeFormats = await import(`${moduleBase}/../office-formats/src/Main.js`);
      const Model = await import(`${moduleBase}/Model.js`);

      const started = performance.now();
      const fetchStarted = performance.now();
      const buffer = await fetch(xlsxUrl).then(response => {
        if (!response.ok) throw new Error(`fetch failed ${response.status}: ${xlsxUrl}`);
        return response.arrayBuffer();
      });
      const fetchMs = performance.now() - fetchStarted;
      const loadStarted = performance.now();
      const workbook = await OfficeFormats.default().excel.load(buffer);
      const workbookLoadMs = performance.now() - loadStarted;
      const sheetImportStarted = performance.now();
      const model = Importer.importSheet(workbook, undefined);
      const sheetImportMs = performance.now() - sheetImportStarted;
      const importFromBufferMs = workbookLoadMs + sheetImportMs;
      const importMs = performance.now() - started;

      let checksum = 0;
      let realUsedBoundsMs = 0;
      let realSearchCellsMs = 0;
      let realExportCsvMs = 0;
      const queryStarted = performance.now();
      for (let i = 0; i < 80; i++) {
        let step = performance.now();
        checksum += Model.usedBounds(model)?.[0] ?? 0;
        realUsedBoundsMs += performance.now() - step;
        step = performance.now();
        checksum += Model.searchCells('total', false, model).length;
        realSearchCellsMs += performance.now() - step;
        step = performance.now();
        checksum += Model.exportCSV(model).length;
        realExportCsvMs += performance.now() - step;
      }
      const queryMs = performance.now() - queryStarted;

      return {
        durationMs: performance.now() - started,
        importMs,
        fetchMs,
        importFromBufferMs,
        workbookLoadMs,
        sheetImportMs,
        queryMs,
        usedBoundsMs: realUsedBoundsMs,
        searchCellsMs: realSearchCellsMs,
        exportCsvMs: realExportCsvMs,
        sourceCellCount: Model.cellCount(model),
        rowCount: model.RowCount,
        colCount: model.ColCount,
        checksum
      };
    }, { moduleBase, xlsxUrl: fsUrl(xlsxFixture) }));

  realXlsx.fixture = rel(xlsxFixture);
  return [synthetic, realXlsx];
}

if (!(await fileExists(wordFixture))) {
  throw new Error(`Word fixture not found: ${wordFixture}`);
}
if (!(await fileExists(xlsxFixture))) {
  throw new Error(`XLSX fixture not found: ${xlsxFixture}`);
}

if (!skipBuild) {
  await run('npm', ['run', 'build:fable'], { cwd: officeRoot });
  await run('dotnet', ['tool', 'restore'], { cwd: root });
  await run('dotnet', ['fable', '.', '--outDir', 'src', '--lang', 'JavaScript'], { cwd: root });
}

const server = await startVite();
const report = {
  timestamp: new Date().toISOString(),
  skipBuild,
  wordFixture: rel(wordFixture),
  xlsxFixture: rel(xlsxFixture),
  userAgent: null,
  word: [],
  spreadsheet: []
};

let browser;
try {
  browser = await chromium.launch({
    args: ['--js-flags=--expose-gc', '--enable-precise-memory-info']
  });
  const page = await browser.newPage({ viewport: { width: 1200, height: 1800 } });
  page.on('console', msg => {
    if (msg.type() !== 'debug') console.log(`[browser:${msg.type()}] ${msg.text()}`);
  });
  page.on('pageerror', error => {
    console.error(`[browser:pageerror] ${error?.stack || error?.message || String(error)}`);
  });
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Runtime.enable');
  await cdp.send('HeapProfiler.enable');
  report.userAgent = await page.evaluate(() => navigator.userAgent);

  const loadMetric = await loadWordFixture(page, cdp);
  report.word.push(loadMetric);
  report.word.push(...await runWordWorkflows(page, cdp));
  report.spreadsheet.push(...await runSpreadsheetWorkflows(page, cdp));

  const outFile = path.join(outDir, 'latest.json');
  await fs.writeFile(outFile, JSON.stringify(report, null, 2));
  console.log(`browser perf report: ${path.relative(officeRoot, outFile)}`);
  for (const metric of [...report.word, ...report.spreadsheet]) {
    const ms = Number(metric.durationMs).toFixed(3);
    const heap = metric.heapDeltaBytes >= 0 ? `+${metric.heapDeltaBytes}` : `${metric.heapDeltaBytes}`;
    console.log(`${metric.name}: ${ms} ms, heap ${heap} B`);
  }
} finally {
  if (browser) await browser.close();
  server.kill('SIGKILL');
}
