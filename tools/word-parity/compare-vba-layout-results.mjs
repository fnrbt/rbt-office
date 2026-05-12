#!/usr/bin/env node
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const officeRoot = path.resolve(process.env.OFFICE_ROOT ?? process.cwd());
const args = process.argv.slice(2);

const options = {
  outDir: path.join(officeRoot, 'artifacts/vba-layout-compare'),
  limit: Infinity,
  maxLines: 500,
  coordThreshold: 6,
};

const inputs = [];
for (let i = 0; i < args.length; i++) {
  const arg = args[i];
  if (arg === '--out') options.outDir = path.resolve(args[++i]);
  else if (arg === '--limit') options.limit = Number(args[++i]);
  else if (arg === '--max-lines') options.maxLines = Number(args[++i]);
  else if (arg === '--coord-threshold') options.coordThreshold = Number(args[++i]);
  else if (arg === '-h' || arg === '--help') usage(0);
  else inputs.push(arg);
}

if (inputs.length === 0) usage(2);
fs.mkdirSync(options.outDir, { recursive: true });

const docs = collectDocx(inputs).slice(0, options.limit);
if (docs.length === 0) {
  console.error('no DOCX files found');
  process.exit(2);
}

const summaryRows = [];
for (const docx of docs) {
  const rel = path.relative(officeRoot, docx);
  const stem = sanitize(rel.replace(/\.docx$/i, ''));
  const docOut = path.join(options.outDir, stem);
  fs.mkdirSync(docOut, { recursive: true });

  const wordTsv = path.join(docOut, 'word-layout.tsv');
  const browserJson = path.join(docOut, 'browser-range-info.json');
  const browserPng = path.join(docOut, 'browser.png');
  const reportTsv = path.join(docOut, 'vba-diff.tsv');

  console.log(`compare vba: ${rel}`);
  const wordOk = run('tools/word-parity/word-layout-tsv.sh', [docx, wordTsv]);
  const browserOk = run('node', ['tools/word-parity/capture-browser-docx.mjs', docx, browserPng, browserJson], {
    BROWSER_CAPTURE_LAYOUT: 'range-info',
  });

  if (!wordOk || !browserOk) {
    summaryRows.push({
      doc: rel,
      status: wordOk ? 'browser-failed' : browserOk ? 'word-failed' : 'word-and-browser-failed',
      score: '',
      detail: '',
    });
    continue;
  }

  const wordLines = parseWordTsv(wordTsv).filter(line => line.norm).slice(0, options.maxLines);
  const browserLines = JSON.parse(fs.readFileSync(browserJson, 'utf8'))
    .map((line, index) => normalizeBrowserRangeInfo(line, index))
    .filter(line => line.norm)
    .slice(0, options.maxLines);

  const pairs = alignLines(wordLines, browserLines);
  const diffs = pairs.map(buildDiff);
  writeReport(reportTsv, diffs);

  if (pairs.length === 0) {
    summaryRows.push({
      doc: rel,
      status: 'no-matches',
      score: '',
      detail: `0 matched lines: word=${wordLines.length} browser=${browserLines.length}`,
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
      detail: `${pairs.length} matched ranges`,
    });
  } else {
    for (const row of mismatches) {
      summaryRows.push({
        doc: rel,
        status: 'delta',
        score: row.score.toFixed(2),
        detail: `line ${row.wordIndex}/${row.browserIndex} ${row.mismatches || 'coordinate'} dx=${fmt(row.dx)} dy=${fmt(row.dy)} text="${clip(row.wordText)}"`,
      });
    }
  }
}

writeSummary(path.join(options.outDir, 'summary.tsv'), summaryRows);
console.log(`summary: ${path.join(options.outDir, 'summary.tsv')}`);

function usage(code) {
  console.error(`usage: node tools/word-parity/compare-vba-layout-results.mjs [options] DOCX_OR_DIR...

Options:
  --out DIR             artifact output directory (default: artifacts/vba-layout-compare)
  --limit N             max number of discovered DOCX files
  --max-lines N         compare first N non-empty lines per document (default: 500)
  --coord-threshold N   coordinate mismatch threshold in CSS pixels (default: 6)

Outputs per document:
  word-layout.tsv, browser-range-info.json, browser.png, vba-diff.tsv
`);
  process.exit(code);
}

function collectDocx(paths) {
  const found = [];
  for (const input of paths) {
    const full = path.resolve(input);
    if (!fs.existsSync(full)) continue;
    const stat = fs.statSync(full);
    if (stat.isDirectory()) {
      for (const name of fs.readdirSync(full).sort()) {
        found.push(...collectDocx([path.join(full, name)]));
      }
    } else if (/\.docx$/i.test(full) && !/\.capture-docx\//.test(full)) {
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

function parseWordTsv(file) {
  const text = fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, '');
  const rows = text.split(/\r?\n/).filter(Boolean);
  if (rows.length === 0) return [];
  const headers = rows[0].split('\t');
  return rows.slice(1).map((row, index) => {
    const cells = row.split('\t');
    const item = Object.fromEntries(headers.map((header, i) => [header, cells[i] ?? '']));
    const text = item.text ?? '';
    return {
      index,
      page: int(item.page),
      adjustedPage: int(item.adjustedPage || item.page),
      section: int(item.section),
      pageCount: int(item.pageCount),
      x: cssPx(item.x),
      y: cssPx(item.y),
      textBoundaryX: cssPx(item.textBoundaryX),
      textBoundaryY: cssPx(item.textBoundaryY),
      firstCharColumn: int(item.firstCharColumn),
      startRow: tableCellValue(item.startRow),
      endRow: tableCellValue(item.endRow),
      maxRows: tableCellValue(item.maxRows),
      startColumn: tableCellValue(item.startColumn),
      endColumn: tableCellValue(item.endColumn),
      maxColumns: tableCellValue(item.maxColumns),
      atEndOfRowMarker: boolText(item.atEndOfRowMarker),
      inContentControl: boolText(item.inContentControl),
      inFieldCode: boolText(item.inFieldCode),
      inFieldResult: boolText(item.inFieldResult),
      rangeLines: int(item.rangeLines),
      rangePages: int(item.rangePages),
      inTable: boolText(item.inTable),
      text,
      norm: normalizeText(text),
    };
  });
}

function normalizeBrowserRangeInfo(line, index) {
  const info = line.Information ?? {};
  const stats = line.Statistics ?? {};
  const text = line.Text ?? '';
  return {
    index,
    page: int(info.wdActiveEndPageNumber),
    adjustedPage: int(info.wdActiveEndAdjustedPageNumber),
    section: int(info.wdActiveEndSectionNumber),
    pageCount: int(info.wdNumberOfPagesInDocument),
    x: num(info.wdHorizontalPositionRelativeToPage),
    y: num(info.wdVerticalPositionRelativeToPage),
    textBoundaryX: num(info.wdHorizontalPositionRelativeToTextBoundary),
    textBoundaryY: num(info.wdVerticalPositionRelativeToTextBoundary),
    firstCharColumn: int(info.wdFirstCharacterColumnNumber),
    startRow: scalar(info.wdStartOfRangeRowNumber),
    endRow: scalar(info.wdEndOfRangeRowNumber),
    maxRows: scalar(info.wdMaximumNumberOfRows),
    startColumn: scalar(info.wdStartOfRangeColumnNumber),
    endColumn: scalar(info.wdEndOfRangeColumnNumber),
    maxColumns: scalar(info.wdMaximumNumberOfColumns),
    atEndOfRowMarker: boolText(info.wdAtEndOfRowMarker),
    inContentControl: boolText(info.wdInContentControl),
    inFieldCode: boolText(info.wdInFieldCode),
    inFieldResult: boolText(info.wdInFieldResult),
    rangeLines: int(stats.Lines),
    rangePages: int(stats.Pages),
    inTable: boolText(info.wdWithInTable),
    blockId: line.BlockId ?? '',
    startOffset: int(line.StartOffset),
    endOffset: int(line.EndOffset),
    text,
    norm: normalizeText(text),
  };
}

function alignLines(wordLines, browserLines) {
  if (wordLines.length * browserLines.length > 250000) {
    const count = Math.min(wordLines.length, browserLines.length);
    return Array.from({ length: count }, (_, index) => ({ word: wordLines[index], browser: browserLines[index] }));
  }

  const n = wordLines.length;
  const m = browserLines.length;
  const dp = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      dp[i][j] = wordLines[i].norm === browserLines[j].norm
        ? dp[i + 1][j + 1] + 1
        : Math.max(dp[i + 1][j], dp[i][j + 1]);
    }
  }

  const pairs = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (wordLines[i].norm === browserLines[j].norm) {
      pairs.push({ word: wordLines[i], browser: browserLines[j] });
      i++;
      j++;
    } else if (dp[i + 1][j] >= dp[i][j + 1]) {
      i++;
    } else {
      j++;
    }
  }

  if (pairs.length < Math.min(n, m) * 0.25) {
    const count = Math.min(n, m);
    return Array.from({ length: count }, (_, index) => ({ word: wordLines[index], browser: browserLines[index] }));
  }
  return pairs;
}

function buildDiff({ word, browser }) {
  const checks = [
    exact('page', word.page, browser.page),
    exact('adjustedPage', word.adjustedPage, browser.adjustedPage),
    exact('section', word.section, browser.section),
    exact('pageCount', word.pageCount, browser.pageCount),
    exact('firstCharColumn', word.firstCharColumn, browser.firstCharColumn),
    exact('inTable', word.inTable, browser.inTable),
    exact('startRow', word.startRow, browser.startRow),
    exact('endRow', word.endRow, browser.endRow),
    exact('maxRows', word.maxRows, browser.maxRows),
    exact('startColumn', word.startColumn, browser.startColumn),
    exact('endColumn', word.endColumn, browser.endColumn),
    exact('maxColumns', word.maxColumns, browser.maxColumns),
    exact('atEndOfRowMarker', word.atEndOfRowMarker, browser.atEndOfRowMarker),
    exact('inContentControl', word.inContentControl, browser.inContentControl),
    exact('inFieldCode', word.inFieldCode, browser.inFieldCode),
    exact('inFieldResult', word.inFieldResult, browser.inFieldResult),
    exact('rangeLines', word.rangeLines, browser.rangeLines),
    exact('rangePages', word.rangePages, browser.rangePages),
  ].filter(Boolean);

  const dx = browser.x - word.x;
  const dy = browser.y - word.y;
  const textBoundaryDx = browser.textBoundaryX - word.textBoundaryX;
  const textBoundaryDy = browser.textBoundaryY - word.textBoundaryY;
  const coordinateScore = Math.max(abs(dx), abs(dy), abs(textBoundaryDx), abs(textBoundaryDy));
  const score = Math.max(coordinateScore, checks.length * 10);

  return {
    wordIndex: word.index,
    browserIndex: browser.index,
    blockId: browser.blockId,
    startOffset: browser.startOffset,
    endOffset: browser.endOffset,
    wordPage: word.page,
    browserPage: browser.page,
    wordAdjustedPage: word.adjustedPage,
    browserAdjustedPage: browser.adjustedPage,
    wordSection: word.section,
    browserSection: browser.section,
    wordPageCount: word.pageCount,
    browserPageCount: browser.pageCount,
    wordX: word.x,
    browserX: browser.x,
    dx,
    wordY: word.y,
    browserY: browser.y,
    dy,
    wordTextBoundaryX: word.textBoundaryX,
    browserTextBoundaryX: browser.textBoundaryX,
    textBoundaryDx,
    wordTextBoundaryY: word.textBoundaryY,
    browserTextBoundaryY: browser.textBoundaryY,
    textBoundaryDy,
    wordFirstCharColumn: word.firstCharColumn,
    browserFirstCharColumn: browser.firstCharColumn,
    wordInTable: word.inTable,
    browserInTable: browser.inTable,
    wordStartRow: word.startRow,
    browserStartRow: browser.startRow,
    wordEndRow: word.endRow,
    browserEndRow: browser.endRow,
    wordMaxRows: word.maxRows,
    browserMaxRows: browser.maxRows,
    wordStartColumn: word.startColumn,
    browserStartColumn: browser.startColumn,
    wordEndColumn: word.endColumn,
    browserEndColumn: browser.endColumn,
    wordMaxColumns: word.maxColumns,
    browserMaxColumns: browser.maxColumns,
    wordAtEndOfRowMarker: word.atEndOfRowMarker,
    browserAtEndOfRowMarker: browser.atEndOfRowMarker,
    wordInContentControl: word.inContentControl,
    browserInContentControl: browser.inContentControl,
    wordInFieldCode: word.inFieldCode,
    browserInFieldCode: browser.inFieldCode,
    wordInFieldResult: word.inFieldResult,
    browserInFieldResult: browser.inFieldResult,
    wordRangeLines: word.rangeLines,
    browserRangeLines: browser.rangeLines,
    wordRangePages: word.rangePages,
    browserRangePages: browser.rangePages,
    coordinateScore,
    mismatchCount: checks.length,
    mismatches: checks.map(check => check.field).join(','),
    score,
    wordText: word.text,
    browserText: browser.text,
  };
}

function exact(field, word, browser) {
  return comparable(word) !== comparable(browser) ? { field, word, browser } : null;
}

function writeReport(file, rows) {
  const headers = [
    'wordIndex', 'browserIndex', 'blockId', 'startOffset', 'endOffset',
    'wordPage', 'browserPage', 'wordAdjustedPage', 'browserAdjustedPage',
    'wordSection', 'browserSection', 'wordPageCount', 'browserPageCount',
    'wordX', 'browserX', 'dx', 'wordY', 'browserY', 'dy',
    'wordTextBoundaryX', 'browserTextBoundaryX', 'textBoundaryDx',
    'wordTextBoundaryY', 'browserTextBoundaryY', 'textBoundaryDy',
    'wordFirstCharColumn', 'browserFirstCharColumn',
    'wordInTable', 'browserInTable', 'wordStartRow', 'browserStartRow',
    'wordEndRow', 'browserEndRow', 'wordMaxRows', 'browserMaxRows',
    'wordStartColumn', 'browserStartColumn', 'wordEndColumn', 'browserEndColumn',
    'wordMaxColumns', 'browserMaxColumns', 'wordAtEndOfRowMarker',
    'browserAtEndOfRowMarker', 'wordInContentControl', 'browserInContentControl',
    'wordInFieldCode', 'browserInFieldCode', 'wordInFieldResult',
    'browserInFieldResult', 'wordRangeLines', 'browserRangeLines',
    'wordRangePages', 'browserRangePages', 'coordinateScore', 'mismatchCount',
    'mismatches', 'score', 'wordText', 'browserText',
  ];
  fs.writeFileSync(file, toTsv(headers, rows), 'utf8');
}

function writeSummary(file, rows) {
  const headers = ['doc', 'status', 'score', 'detail'];
  fs.writeFileSync(file, toTsv(headers, rows), 'utf8');
}

function toTsv(headers, rows) {
  return [
    headers.join('\t'),
    ...rows.map(row => headers.map(header => cleanCell(formatCell(row[header]))).join('\t')),
    '',
  ].join('\n');
}

function formatCell(value) {
  return typeof value === 'number' ? fmt(value) : value ?? '';
}

function normalizeText(text) {
  return String(text)
    .replace(/\u0007/g, '')
    .replace(/\uFFFC/g, '[image]')
    .replace(/\s+/g, ' ')
    .trim();
}

function sanitize(name) {
  return name.replace(/[^A-Za-z0-9._-]+/g, '_').replace(/^_+|_+$/g, '') || 'doc';
}

function cleanCell(value) {
  return String(value).replace(/[\t\r\n]+/g, ' ');
}

function tableCellValue(value) {
  const text = String(value ?? '');
  return text === '-1' ? '' : text;
}

function scalar(value) {
  if (value === null || value === undefined) return '';
  return String(value);
}

function boolText(value) {
  if (value === true) return 'true';
  if (value === false) return 'false';
  const text = String(value ?? '').trim().toLowerCase();
  if (text === 'true') return 'true';
  if (text === 'false') return 'false';
  return text;
}

function comparable(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? String(value) : '';
  return String(value ?? '');
}

function cssPx(value) {
  const n = num(value);
  return Number.isFinite(n) ? n * 96 / 72 : NaN;
}

function num(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : NaN;
}

function int(value) {
  const n = Number.parseInt(value, 10);
  return Number.isFinite(n) ? n : 0;
}

function abs(value) {
  return Number.isFinite(value) ? Math.abs(value) : 0;
}

function fmt(value) {
  return Number.isFinite(value) ? value.toFixed(2) : '';
}

function clip(text) {
  const value = normalizeText(text);
  return value.length > 80 ? `${value.slice(0, 77)}...` : value;
}
