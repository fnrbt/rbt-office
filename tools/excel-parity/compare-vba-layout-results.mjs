#!/usr/bin/env node
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

// Mirrors tools/word-parity/compare-vba-layout-results.mjs. The browser
// side currently exposes only a paint snapshot; once the planned
// getLayoutCells/getRangeInformation/computeRangeStatistics calls land on
// Rbt.Spreadsheet.API, swap the placeholder browser capture below for a
// real Playwright-driven call analogous to capture-browser-docx.mjs.

const officeRoot = path.resolve(process.env.OFFICE_ROOT ?? process.cwd());
const args = process.argv.slice(2);

const options = {
  outDir: path.join(officeRoot, 'artifacts/vba-layout-compare-excel'),
  limit: Infinity,
  maxCells: 5000,
  coordThreshold: 6,
  browserCommand: process.env.EXCEL_BROWSER_CAPTURE,
};

const inputs = [];
for (let i = 0; i < args.length; i++) {
  const arg = args[i];
  if (arg === '--out') options.outDir = path.resolve(args[++i]);
  else if (arg === '--limit') options.limit = Number(args[++i]);
  else if (arg === '--max-cells') options.maxCells = Number(args[++i]);
  else if (arg === '--coord-threshold') options.coordThreshold = Number(args[++i]);
  else if (arg === '--browser') options.browserCommand = args[++i];
  else if (arg === '-h' || arg === '--help') usage(0);
  else inputs.push(arg);
}

if (inputs.length === 0) usage(2);
fs.mkdirSync(options.outDir, { recursive: true });

const docs = collectXlsx(inputs).slice(0, options.limit);
if (docs.length === 0) {
  console.error('no XLSX files found');
  process.exit(2);
}

const summaryRows = [];
for (const xlsx of docs) {
  const rel = path.relative(officeRoot, xlsx);
  const stem = sanitize(rel.replace(/\.xlsx$/i, ''));
  const docOut = path.join(options.outDir, stem);
  fs.mkdirSync(docOut, { recursive: true });

  const excelTsv = path.join(docOut, 'excel-layout.tsv');
  const browserJson = path.join(docOut, 'browser-cell-info.json');
  const browserPng = path.join(docOut, 'browser.png');
  const reportTsv = path.join(docOut, 'vba-diff.tsv');

  console.log(`compare excel: ${rel}`);
  const excelOk = run('tools/excel-parity/excel-layout-tsv.sh', [xlsx, excelTsv]);
  let browserOk = false;
  if (options.browserCommand) {
    browserOk = run(options.browserCommand, [xlsx, browserPng, browserJson]);
  } else {
    // No browser capture is wired up yet — write an empty JSON so the diff
    // step still produces a valid file and the missing-side is obvious in
    // the summary.
    fs.writeFileSync(browserJson, '[]');
    browserOk = true;
  }

  if (!excelOk || !browserOk) {
    summaryRows.push({
      doc: rel,
      status: excelOk ? 'browser-failed' : browserOk ? 'excel-failed' : 'excel-and-browser-failed',
      score: '',
      detail: '',
    });
    continue;
  }

  const excelCells = parseExcelTsv(excelTsv).slice(0, options.maxCells);
  const browserCells = JSON.parse(fs.readFileSync(browserJson, 'utf8'))
    .map((cell, index) => normalizeBrowserCell(cell, index))
    .slice(0, options.maxCells);

  const pairs = alignCells(excelCells, browserCells);
  const diffs = pairs.map(buildDiff);
  writeReport(reportTsv, diffs);

  if (pairs.length === 0) {
    summaryRows.push({
      doc: rel,
      status: 'no-matches',
      score: '',
      detail: `0 matched cells: excel=${excelCells.length} browser=${browserCells.length}`,
    });
    continue;
  }

  const mismatches = diffs
    .filter(row => row.mismatchCount > 0 || row.coordinateScore >= options.coordThreshold)
    .sort((a, b) => b.score - a.score)
    .slice(0, 12);

  if (mismatches.length === 0) {
    summaryRows.push({
      doc: rel,
      status: 'ok',
      score: '',
      detail: `${pairs.length} matched cells`,
    });
  } else {
    for (const row of mismatches) {
      summaryRows.push({
        doc: rel,
        status: 'delta',
        score: row.score.toFixed(2),
        detail: `${row.excelAddress}/${row.browserAddress} ${row.mismatches || 'coordinate'} dx=${fmt(row.dx)} dy=${fmt(row.dy)} value="${clip(row.excelValue)}"`,
      });
    }
  }
}

writeSummary(path.join(options.outDir, 'summary.tsv'), summaryRows);
console.log(`summary: ${path.join(options.outDir, 'summary.tsv')}`);

function usage(code) {
  console.error(`usage: node tools/excel-parity/compare-vba-layout-results.mjs [options] XLSX_OR_DIR...

Options:
  --out DIR             artifact output directory (default: artifacts/vba-layout-compare-excel)
  --limit N             max number of discovered XLSX files
  --max-cells N         compare first N non-empty cells per document (default: 5000)
  --coord-threshold N   coordinate mismatch threshold in CSS pixels (default: 6)
  --browser CMD         command that captures browser cell info; receives XLSX, PNG, JSON args

Environment:
  EXCEL_BROWSER_CAPTURE same as --browser if set.

Outputs per document:
  excel-layout.tsv, browser-cell-info.json, browser.png, vba-diff.tsv
`);
  process.exit(code);
}

function collectXlsx(paths) {
  const found = [];
  for (const input of paths) {
    const full = path.resolve(input);
    if (!fs.existsSync(full)) continue;
    const stat = fs.statSync(full);
    if (stat.isDirectory()) {
      for (const name of fs.readdirSync(full).sort()) {
        found.push(...collectXlsx([path.join(full, name)]));
      }
    } else if (/\.xlsx$/i.test(full)) {
      found.push(full);
    }
  }
  return found;
}

function run(command, args, extraEnv = {}) {
  const result = spawnSync(command, args, {
    cwd: officeRoot,
    env: { ...process.env, ...extraEnv },
    stdio: 'inherit',
  });
  return result.status === 0;
}

function sanitize(value) {
  return value.replace(/[^A-Za-z0-9._-]+/g, '_');
}

function parseExcelTsv(file) {
  const text = fs.readFileSync(file, 'utf8');
  const lines = text.split(/\r?\n/).filter(Boolean);
  if (lines.length === 0) return [];
  const header = lines[0].split('\t');
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = lines[i].split('\t');
    if (cols.length < 4) continue;
    const record = {};
    for (let j = 0; j < header.length; j++) record[header[j]] = cols[j] ?? '';
    rows.push(record);
  }
  // Convert Excel point coordinates into CSS pixels so the diff matches the
  // browser CSS pixel space.
  for (const row of rows) {
    row.norm = {
      sheet: row.sheet,
      address: row.address,
      row: Number(row.row),
      col: Number(row.column),
      xPx: Number(row.x) * (96 / 72),
      yPx: Number(row.y) * (96 / 72),
      widthPx: Number(row.width) * (96 / 72),
      heightPx: Number(row.height) * (96 / 72),
      value: row.value,
      text: row.text,
    };
  }
  return rows;
}

function normalizeBrowserCell(cell, index) {
  return {
    norm: {
      sheet: cell.Sheet ?? cell.sheet ?? '',
      address: cell.Address ?? cell.address ?? '',
      row: Number(cell.Row ?? cell.row ?? -1),
      col: Number(cell.Column ?? cell.column ?? -1),
      xPx: Number(cell.X ?? cell.x ?? NaN),
      yPx: Number(cell.Y ?? cell.y ?? NaN),
      widthPx: Number(cell.Width ?? cell.width ?? NaN),
      heightPx: Number(cell.Height ?? cell.height ?? NaN),
      value: String(cell.Value ?? cell.value ?? ''),
      text: String(cell.Text ?? cell.text ?? ''),
    },
    index,
  };
}

function alignCells(excelCells, browserCells) {
  // Excel addresses are sheet+A1 references; browsers should report the same.
  const browserMap = new Map();
  for (const cell of browserCells) {
    const key = `${cell.norm.sheet}!${cell.norm.address}`;
    if (!browserMap.has(key)) browserMap.set(key, cell);
  }
  const pairs = [];
  for (const cell of excelCells) {
    const key = `${cell.norm.sheet}!${cell.norm.address}`;
    if (browserMap.has(key)) {
      pairs.push({ excel: cell, browser: browserMap.get(key) });
    }
  }
  return pairs;
}

function buildDiff(pair) {
  const e = pair.excel.norm;
  const b = pair.browser.norm;
  const dx = b.xPx - e.xPx;
  const dy = b.yPx - e.yPx;
  const coordinateScore = Math.hypot(dx, dy);
  const mismatches = [];
  if (e.value !== b.value) mismatches.push('value');
  if (e.text !== b.text) mismatches.push('text');
  if (Math.abs(e.widthPx - b.widthPx) > 1) mismatches.push('width');
  if (Math.abs(e.heightPx - b.heightPx) > 1) mismatches.push('height');
  return {
    excelAddress: `${e.sheet}!${e.address}`,
    browserAddress: `${b.sheet}!${b.address}`,
    excelValue: e.value,
    dx,
    dy,
    mismatches: mismatches.join(','),
    mismatchCount: mismatches.length,
    coordinateScore,
    score: coordinateScore + mismatches.length * 10,
  };
}

function writeReport(file, diffs) {
  const header = 'excelAddress\tbrowserAddress\tdx\tdy\tmismatches\tcoordScore\tscore';
  const lines = [header];
  for (const diff of diffs) {
    lines.push(`${diff.excelAddress}\t${diff.browserAddress}\t${diff.dx.toFixed(2)}\t${diff.dy.toFixed(2)}\t${diff.mismatches || ''}\t${diff.coordinateScore.toFixed(2)}\t${diff.score.toFixed(2)}`);
  }
  fs.writeFileSync(file, lines.join('\n'));
}

function writeSummary(file, rows) {
  const header = 'doc\tstatus\tscore\tdetail';
  const lines = [header];
  for (const row of rows) lines.push(`${row.doc}\t${row.status}\t${row.score}\t${row.detail}`);
  fs.writeFileSync(file, lines.join('\n'));
}

function fmt(value) {
  if (!Number.isFinite(value)) return '';
  return value.toFixed(2);
}

function clip(value) {
  if (typeof value !== 'string') return '';
  return value.length > 40 ? value.slice(0, 37) + '...' : value;
}
