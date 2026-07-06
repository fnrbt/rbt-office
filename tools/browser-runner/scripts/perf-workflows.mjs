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
  // Use Vite's normalized /@fs/home/... form so page-level dynamic imports
  // match dist modules' rewritten relative import URLs and stay single-instance.
  return `/@fs${file}`;
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
        const content = Array.from({ length: 300 }, (_, i) => `${i}: ${paragraph}${i % 7}`).join('\n');
        const replacements = ['theta', 'lambda', 'sigma', 'omega', 'kappa', 'zeta'];
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
          window.rbWordRunner.insertTextAt(replacements[i % replacements.length]);
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

    // Realistic typing shape: 300 separate paragraphs (setContent puts the
    // whole fixture in ONE paragraph, which per-block layout reuse cannot
    // help; real documents are many blocks).
    const multiParagraphEditCycle = await measure(page, cdp, 'word.multi-paragraph-edit-cycle', () =>
      page.evaluate(({ iterations }) => {
        const paragraph = 'browser perf paragraph alpha beta gamma delta epsilon ';
        const paragraphs = Array.from({ length: 300 }, (_, i) => `${i}: ${paragraph}${i % 7}`);
        const replacements = ['theta', 'lambda', 'sigma', 'omega', 'kappa', 'zeta'];
        window.rbWordRunner.setParagraphs(paragraphs);
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
          if (!matches.length) throw new Error('Word multi-paragraph edit workflow could not find insertion target');
          const match = matches[i % matches.length];
          t = performance.now();
          window.rbWordRunner.selectFindMatch(match);
          selectMatchMs += performance.now() - t;
          t = performance.now();
          window.rbWordRunner.insertTextAt(replacements[i % replacements.length]);
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

  return [rangeQueries, findReplace, editCycle, multiParagraphEditCycle];
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
      let queryEditMs = 0;
      let usedBoundsAfterEditMs = 0;
      let searchCellsAfterEditMs = 0;
      let uniqueValueCountAfterEditMs = 0;
      let t = performance.now();
      let queryModel = model;
      for (let i = 0; i < 160; i++) {
          let step = performance.now();
          queryModel = Model.importCSV(`query-scratch-${i}`, 1, 5, queryModel);
          queryEditMs += performance.now() - step;
          step = performance.now();
          checksum += Model.usedBounds(queryModel)?.[0] ?? 0;
          usedBoundsAfterEditMs += performance.now() - step;
          step = performance.now();
          checksum += Model.searchCells('needle', false, queryModel).length;
          searchCellsAfterEditMs += performance.now() - step;
          step = performance.now();
          checksum += Model.columnUniqueValueCount(0, queryModel);
          uniqueValueCountAfterEditMs += performance.now() - step;
      }
      const queryAfterEditMs = performance.now() - t;

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

      const filterAfterEditIterations = 30;
      let filterEditMs = 0;
      let filterApplyAfterEditMs = 0;
      let filterVisibleRowsAfterEditMs = 0;
      let filterProjectionAfterEditMs = 0;
      t = performance.now();
      let filterModel = model;
      for (let i = 0; i < filterAfterEditIterations; i++) {
          let step = performance.now();
          filterModel = Model.importCSV(`filter-scratch-${i}`, 1, 5, filterModel);
          filterEditMs += performance.now() - step;
          step = performance.now();
          const filteredAfterEdit = Model.applyFilterValues(0, visibleValues, filterModel);
          filterApplyAfterEditMs += performance.now() - step;
          step = performance.now();
          const visibleRowsAfterEdit = Model.visibleRowCount(filteredAfterEdit);
          if (visibleRowsAfterEdit !== expectedVisibleRows) {
            throw new Error(`unexpected edited filter result: visible ${visibleRowsAfterEdit}/${expectedVisibleRows}`);
          }
          checksum += visibleRowsAfterEdit;
          filterVisibleRowsAfterEditMs += performance.now() - step;
          step = performance.now();
          const hiddenRowsAfterEdit = Model.applyFilterRowHeights(filteredAfterEdit).FilterHiddenRows.length;
          if (hiddenRowsAfterEdit !== expectedHiddenRows) {
            throw new Error(`unexpected edited filter projection: hidden ${hiddenRowsAfterEdit}/${expectedHiddenRows}`);
          }
          checksum += hiddenRowsAfterEdit;
          filterProjectionAfterEditMs += performance.now() - step;
      }
      const filterAfterEditMs = performance.now() - t;

      const filterReadMs =
          filterApplyFirstMs
          + visibleRowsFirstMs
          + filterProjectionFirstMs
          + visibleRowsHotMs
          + filterProjectionHotMs;

      let exportEditMs = 0;
      let clearSelectionAfterEditMs = 0;
      let exportCsvAfterEditMs = 0;
      t = performance.now();
      let exportModel = model;
      for (let i = 0; i < 24; i++) {
          const selection = new Model.Selection$(1, [50, 0, 900, 5]);
          let step = performance.now();
          exportModel = Model.importCSV(`export-scratch-${i}`, 1, 5, exportModel);
          exportEditMs += performance.now() - step;
          step = performance.now();
          const cleared = Model.clearSelectionCells(selection, exportModel);
          clearSelectionAfterEditMs += performance.now() - step;
          step = performance.now();
          checksum += Model.exportCSV(cleared).length;
          exportCsvAfterEditMs += performance.now() - step;
      }
      const clearExportAfterEditMs = performance.now() - t;

      let formulaEditMs = 0;
      let recalcAndSpillAfterEditMs = 0;
      let recalcSearchAfterEditMs = 0;
      t = performance.now();
      let recalcModel = model;
      for (let i = 0; i < 30; i++) {
          let step = performance.now();
          const formula = i % 2 === 0 ? '=C3+D3' : '=C2+D2';
          recalcModel = Model.importCSV(formula, 1, 4, recalcModel);
          formulaEditMs += performance.now() - step;
          step = performance.now();
          const recalculated = Formula.recalcAndSpill(undefined, undefined, recalcModel);
          recalcAndSpillAfterEditMs += performance.now() - step;
          step = performance.now();
          checksum += Model.searchCells('5400', false, recalculated).length;
          recalcSearchAfterEditMs += performance.now() - step;
          recalcModel = recalculated;
      }
      const recalcAfterEditMs = performance.now() - t;

      return {
          durationMs: performance.now() - started,
          rows: rows.length,
          importMs,
          queryAfterEditMs,
          queryEditMs,
          usedBoundsAfterEditMs,
          searchCellsAfterEditMs,
          uniqueValueCountAfterEditMs,
          filterReadMs,
          filterApplyFirstMs,
          visibleRowsFirstMs,
          filterProjectionFirstMs,
          visibleRowsHotMs,
          filterProjectionHotMs,
          filterAfterEditMs,
          filterEditMs,
          filterApplyAfterEditMs,
          filterVisibleRowsAfterEditMs,
          filterProjectionAfterEditMs,
          hotFilterReadIterations,
          filterAfterEditIterations,
          filteredRowCount,
          hiddenRowCount,
          clearExportAfterEditMs,
          exportEditMs,
          clearSelectionAfterEditMs,
          exportCsvAfterEditMs,
          recalcAfterEditMs,
          formulaEditMs,
          recalcAndSpillAfterEditMs,
          recalcSearchAfterEditMs,
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
      const excelApi = OfficeFormats.default().excel;
      const workbook =
        typeof excelApi.loadSelected === 'function'
          ? await excelApi.loadSelected(buffer, undefined)
          : await excelApi.load(buffer);
      const workbookLoadMs = performance.now() - loadStarted;
      const sheetImportStarted = performance.now();
      const model = Importer.importSheet(workbook, undefined);
      const sheetImportMs = performance.now() - sheetImportStarted;
      const importFromBufferMs = workbookLoadMs + sheetImportMs;
      const importMs = performance.now() - started;

      let checksum = 0;
      let realQueryEditMs = 0;
      let realUsedBoundsAfterEditMs = 0;
      let realSearchCellsAfterEditMs = 0;
      let realExportCsvAfterEditMs = 0;
      const queryStarted = performance.now();
      let queryModel = model;
      for (let i = 0; i < 80; i++) {
          let step = performance.now();
          queryModel = Model.importCSV(`real-scratch-${i}`, 0, 0, queryModel);
          realQueryEditMs += performance.now() - step;
          step = performance.now();
          checksum += Model.usedBounds(queryModel)?.[0] ?? 0;
          realUsedBoundsAfterEditMs += performance.now() - step;
          step = performance.now();
          checksum += Model.searchCells('total', false, queryModel).length;
          realSearchCellsAfterEditMs += performance.now() - step;
          step = performance.now();
          checksum += Model.exportCSV(queryModel).length;
          realExportCsvAfterEditMs += performance.now() - step;
      }
      const queryAfterEditMs = performance.now() - queryStarted;

      return {
          durationMs: performance.now() - started,
          importMs,
          fetchMs,
          importFromBufferMs,
          workbookLoadMs,
          sheetImportMs,
          queryAfterEditMs,
          queryEditMs: realQueryEditMs,
          usedBoundsAfterEditMs: realUsedBoundsAfterEditMs,
          searchCellsAfterEditMs: realSearchCellsAfterEditMs,
          exportCsvAfterEditMs: realExportCsvAfterEditMs,
          loadedSheetCount: excelApi.sheetNames(workbook).length,
          sourceCellCount: Model.cellCount(model),
          rowCount: model.RowCount,
        colCount: model.ColCount,
      checksum
      };
    }, { moduleBase, xlsxUrl: fsUrl(xlsxFixture) }));

  realXlsx.fixture = rel(xlsxFixture);
  return [synthetic, realXlsx];
}

async function runSpreadsheetIncrementalEditWorkflow(page, cdp) {
  const moduleBase = fsUrl(path.join(officeRoot, 'dist/spreadsheet'));
  return await measure(page, cdp, 'spreadsheet.incremental-edit', () =>
    page.evaluate(async ({ moduleBase }) => {
      const Model = await import(`${moduleBase}/Model.js`);
      const Formula = await import(`${moduleBase}/Formula.js`);
      const Spreadsheet = await import(`${moduleBase}/Spreadsheet.js`);

      const workflowStarted = performance.now();
      const dataRows = 2000;
      const editIterations = 200;
      const rows = [['Item', 'Group', 'Amount', 'Tax', 'Total', 'Note', 'Aggregate']];
      let initialAggregate = 0;
      for (let row = 1; row <= dataRows; row++) {
        const amount = row * 3;
        const tax = row % 17;
        initialAggregate += amount + tax;
        rows.push([
          `Item ${row}`,
          `Group ${row % 11}`,
          String(amount),
          String(tax),
          `=C${row + 1}+D${row + 1}`,
          row % 13 === 0 ? 'flagged' : 'plain',
          row === 1 ? `=SUM(E2:E${dataRows + 1})` : ''
        ]);
      }

      const csv = rows.map(row => row.join(',')).join('\n');
      const createFormulaModel = () =>
        Formula.recalcAndSpill(
          undefined,
          undefined,
          Model.importCSV(csv, 0, 0, Model.createEmpty(dataRows + 20, 12))
        );

      const edits = Array.from({ length: editIterations }, (_, i) => {
        const row = (i % dataRows) + 1;
        const value = 10000 + i;
        const oldAmount = row * 3;
        const tax = row % 17;
        return { row, col: 2, value, oldAmount, tax };
      });
      const expectedAggregate = edits.reduce(
        (sum, edit) => sum - edit.oldAmount + edit.value,
        initialAggregate
      );
      const lastEdit = edits[edits.length - 1];
      const expectedDependent = lastEdit.value + lastEdit.tax;
      const dependentAddr = new Model.CellAddress(lastEdit.row, 4);
      const aggregateAddr = new Model.CellAddress(1, 6);

      const cellNumber = (model, addr, label) => {
        const value = Model.getCellValue(addr, model);
        const text = Model.displayValue(value);
        const number = Number(text);
        if (!Number.isFinite(number)) {
          throw new Error(`${label}: expected numeric value, saw "${text}"`);
        }
        return number;
      };
      const assertClose = (label, actual, expected) => {
        if (Math.abs(actual - expected) > 0.000001) {
          throw new Error(`${label}: expected ${expected}, saw ${actual}`);
        }
      };

      let directModel = createFormulaModel();
      let directSetCellMs = 0;
      let directRecalcAndSpillMs = 0;
      let directChecksum = 0;
      const directStarted = performance.now();
      for (const edit of edits) {
        const addr = new Model.CellAddress(edit.row, edit.col);
        const cellData = new Model.CellData(
          new Model.CellValue(1, [edit.value]),
          Model.defaultFormat,
          undefined
        );
        let step = performance.now();
        directModel = Model.setCell(addr, cellData, directModel);
        directSetCellMs += performance.now() - step;
        step = performance.now();
        directModel = Formula.recalcAndSpill(undefined, undefined, directModel);
        directRecalcAndSpillMs += performance.now() - step;
        directChecksum += edit.value;
      }
      const directEditMs = performance.now() - directStarted;
      const directDependent = cellNumber(directModel, dependentAddr, 'direct dependent formula');
      const directAggregate = cellNumber(directModel, aggregateAddr, 'direct aggregate formula');
      assertClose('direct dependent formula', directDependent, expectedDependent);
      assertClose('direct aggregate formula', directAggregate, expectedAggregate);

      const host = document.createElement('div');
      host.style.width = '1000px';
      host.style.height = '700px';
      host.style.position = 'relative';
      document.body.appendChild(host);
      const doubleRaf = () =>
        new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));

      let realSetCellMs = 0;
      let realChecksum = 0;
      let realDependent = 0;
      let realAggregate = 0;
      try {
        const spreadsheet = Spreadsheet.default(host);
        spreadsheet.Mount();
        spreadsheet.SetWorkbook([createFormulaModel()], ['Sheet1'], 0);
        await doubleRaf();

        const realStarted = performance.now();
        for (const edit of edits) {
          const step = performance.now();
          spreadsheet.SetCell(edit.row, edit.col, String(edit.value));
          realSetCellMs += performance.now() - step;
          realChecksum += edit.value;
        }
        const realEditMs = performance.now() - realStarted;
        await doubleRaf();

        const realModel = spreadsheet.GetModel();
        realDependent = cellNumber(realModel, dependentAddr, 'real dependent formula');
        realAggregate = cellNumber(realModel, aggregateAddr, 'real aggregate formula');
        assertClose('real dependent formula', realDependent, expectedDependent);
        assertClose('real aggregate formula', realAggregate, expectedAggregate);

        return {
          durationMs: performance.now() - workflowStarted,
          rows: dataRows,
          editIterations,
          realEditMs,
          realSetCellMs,
          directEditMs,
          directSetCellMs,
          directRecalcAndSpillMs,
          expectedDependent,
          realDependent,
          directDependent,
          expectedAggregate,
          realAggregate,
          directAggregate,
          checksum: realChecksum + directChecksum + realDependent + directDependent + realAggregate + directAggregate
        };
      } finally {
        host.remove();
      }
    }, { moduleBase }));
}

async function runSpreadsheetScrollWorkflow(page, cdp) {
  const moduleBase = fsUrl(path.join(officeRoot, 'dist/spreadsheet'));
  return await measure(page, cdp, 'spreadsheet.scroll-dom', () =>
    page.evaluate(async ({ moduleBase }) => {
      const Model = await import(`${moduleBase}/Model.js`);
      const Spreadsheet = await import(`${moduleBase}/Spreadsheet.js`);

      const workflowStarted = performance.now();
      const originalCreateElement = Document.prototype.createElement;
      let createdElementCount = 0;
      Document.prototype.createElement = function (...args) {
        createdElementCount++;
        return originalCreateElement.apply(this, args);
      };

      const resetCreateElementCount = () => {
        createdElementCount = 0;
      };
      const doubleRaf = () =>
        new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      const expectedText = (row, col) => `r${row}c${col}-${row * 31 + col}`;

      const host = document.createElement('div');
      host.style.width = '1000px';
      host.style.height = '700px';
      host.style.position = 'relative';
      document.body.appendChild(host);

      try {
        const spreadsheet = Spreadsheet.default(host);
        spreadsheet.Mount();

        const dataRows = 3000;
        const dataCols = 26;
        const rows = Array.from({ length: dataRows }, (_, row) =>
          Array.from({ length: dataCols }, (_, col) => expectedText(row, col)).join(','));
        const model = Model.importCSV(rows.join('\n'), 0, 0, Model.createEmpty(3200, 30));
        spreadsheet.SetWorkbook([model], ['Sheet1'], 0);
        await doubleRaf();

        const grid = host.querySelector('.spreadsheet-container');
        if (!grid) throw new Error('spreadsheet scroll workflow could not find grid');

        const headerHeight = 30;
        const rowHeaderWidth = 50;
        const rowHeight = 25;
        const colWidth = 100;
        const cellSelector = '.spreadsheet-content .spreadsheet-cell[data-row][data-col]';
        let checksum = 0;

        const mountedCells = () => Array.from(host.querySelectorAll(cellSelector));
        const visibleRowFromScroll = () => Math.max(0, Math.floor((grid.scrollTop - headerHeight) / rowHeight));
        const visibleColFromScroll = () => Math.max(0, Math.floor((grid.scrollLeft - rowHeaderWidth) / colWidth));
        const validateVisibleCells = label => {
          const cells = mountedCells();
          if (!cells.length) throw new Error(`${label}: no mounted spreadsheet cells`);

          const rows = cells.map(cell => Number(cell.getAttribute('data-row'))).filter(Number.isFinite);
          if (!rows.length) throw new Error(`${label}: no mounted cell row attributes`);

          const expectedFirstVisibleRow = visibleRowFromScroll();
          if (!rows.some(row => Math.abs(row - expectedFirstVisibleRow) <= 2)) {
            throw new Error(`${label}: expected row ${expectedFirstVisibleRow} near viewport, saw ${rows.slice(0, 8).join(',')}`);
          }

          const viewportRows = Math.max(1, Math.floor((grid.clientHeight - headerHeight) / rowHeight));
          const viewportCols = Math.max(1, Math.floor((grid.clientWidth - rowHeaderWidth) / colWidth));
          const targetRow = Math.min(dataRows - 1, expectedFirstVisibleRow + Math.floor(viewportRows / 2));
          const targetCol = Math.min(dataCols - 1, visibleColFromScroll() + Math.floor(viewportCols / 2));
          const target = host.querySelector(
            `.spreadsheet-content .spreadsheet-cell[data-row="${targetRow}"][data-col="${targetCol}"]`
          );
          if (!target) {
            throw new Error(`${label}: missing middle viewport cell ${targetRow},${targetCol}`);
          }

          const text = target.textContent ?? '';
          const expected = expectedText(targetRow, targetCol);
          if (text !== expected) {
            throw new Error(`${label}: cell ${targetRow},${targetCol} text "${text}" did not match "${expected}"`);
          }

          const minRow = Math.min(...rows);
          const maxRow = Math.max(...rows);
          if (minRow < expectedFirstVisibleRow - 10 || maxRow > expectedFirstVisibleRow + 60) {
            throw new Error(`${label}: stale offscreen rows mounted; expected around ${expectedFirstVisibleRow}, saw ${minRow}-${maxRow}`);
          }
          if (cells.length >= 5000) {
            throw new Error(`${label}: mounted cell leak guard failed with ${cells.length} cells`);
          }

          checksum += cells.length + minRow + maxRow + targetRow + targetCol + text.length + grid.scrollTop + grid.scrollLeft;
          return cells.length;
        };

        resetCreateElementCount();
        const scrollSteps = 150;
        const scrollStarted = performance.now();
        for (let step = 0; step < scrollSteps; step++) {
          grid.scrollTop += 75;
          if ((step + 1) % 5 === 0) grid.scrollLeft += 60;
          grid.dispatchEvent(new Event('scroll'));
          await doubleRaf();
          checksum += grid.scrollTop + grid.scrollLeft;
        }
        const scrollPassMs = performance.now() - scrollStarted;
        const nodesCreatedDuringScroll = createdElementCount;
        let mountedCellCount = validateVisibleCells('pass A');

        grid.scrollTop = headerHeight + (rowHeight * 600);
        grid.scrollLeft = rowHeaderWidth + (colWidth * 8);
        grid.dispatchEvent(new Event('scroll'));
        await doubleRaf();

        const burstStarted = performance.now();
        for (let i = 0; i < 50; i++) {
          grid.scrollTop += 25;
          grid.dispatchEvent(new Event('scroll'));
          checksum += grid.scrollTop;
        }
        const burstDispatchMs = performance.now() - burstStarted;
        await doubleRaf();
        mountedCellCount = validateVisibleCells('pass B');

        return {
          durationMs: performance.now() - workflowStarted,
          scrollPassMs,
          avgStepMs: scrollPassMs / scrollSteps,
          nodesCreatedDuringScroll,
          burstDispatchMs,
          mountedCellCount,
          checksum
        };
      } finally {
        host.remove();
        Document.prototype.createElement = originalCreateElement;
      }
    }, { moduleBase }));
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
  report.spreadsheet.push(await runSpreadsheetIncrementalEditWorkflow(page, cdp));
  report.spreadsheet.push(await runSpreadsheetScrollWorkflow(page, cdp));

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
