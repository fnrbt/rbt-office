#!/usr/bin/env node
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const officeRoot = path.resolve(process.env.OFFICE_ROOT ?? process.cwd());
const args = process.argv.slice(2);

const options = {
  outDir: path.join(officeRoot, 'artifacts/layout-compare'),
  limit: Infinity,
  maxLines: 500,
  threshold: 6,
  wordSource: 'vbs',
};

const inputs = [];
for (let i = 0; i < args.length; i++) {
  const arg = args[i];
  if (arg === '--out') options.outDir = path.resolve(args[++i]);
  else if (arg === '--limit') options.limit = Number(args[++i]);
  else if (arg === '--max-lines') options.maxLines = Number(args[++i]);
  else if (arg === '--threshold') options.threshold = Number(args[++i]);
  else if (arg === '--word-source') options.wordSource = args[++i];
  else if (arg === '-h' || arg === '--help') usage(0);
  else inputs.push(arg);
}

if (inputs.length === 0) usage(2);
if (!['vbs', 'pdf'].includes(options.wordSource)) {
  console.error(`unsupported --word-source ${options.wordSource}; expected vbs or pdf`);
  process.exit(2);
}

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
  const wordPdf = path.join(docOut, 'word.pdf');
  const browserJson = path.join(docOut, 'browser-lines.json');
  const browserPng = path.join(docOut, 'browser.png');
  const reportTsv = path.join(docOut, 'layout-diff.tsv');

  console.log(`compare: ${rel}`);
  const wordOk = captureWordLayout(docx, wordTsv, wordPdf);
  const browserOk = run('node', ['tools/word-parity/capture-browser-docx.mjs', docx, browserPng, browserJson], {
    BROWSER_CAPTURE_LAYOUT: 'lines',
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
    .map((line, index) => normalizeBrowserLine(line, index))
    .filter(line => line.norm)
    .slice(0, options.maxLines);

  const pairs = alignLines(wordLines, browserLines);
  const diffs = buildDiffs(pairs);
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

  const suspicious = diffs
    .filter(row => row.score >= options.threshold)
    .sort((a, b) => b.score - a.score)
    .slice(0, 12);

  if (suspicious.length === 0) {
    summaryRows.push({
      doc: rel,
      status: 'ok',
      score: '',
      detail: `${pairs.length} matched lines`,
    });
  } else {
    for (const row of suspicious) {
      summaryRows.push({
        doc: rel,
        status: 'delta',
        score: row.score.toFixed(2),
        detail: `line ${row.wordIndex}/${row.browserIndex} dx=${fmt(row.dx)} dy=${fmt(row.dy)} gapDy=${fmt(row.gapDy)} word="${clip(row.wordText)}" browser="${clip(row.browserText)}"`,
      });
    }
  }
}

writeSummary(path.join(options.outDir, 'summary.tsv'), summaryRows);
console.log(`summary: ${path.join(options.outDir, 'summary.tsv')}`);

function usage(code) {
  console.error(`usage: node tools/word-parity/compare-word-browser-layouts.mjs [options] DOCX_OR_DIR...

Options:
  --out DIR       artifact output directory (default: artifacts/layout-compare)
  --limit N       max number of discovered DOCX files
  --max-lines N   compare first N non-empty lines per document (default: 500)
  --threshold N   include rows at or above this CSS-pixel delta (default: 6)
  --word-source S extract Word reference positions from vbs or printed pdf (default: vbs)

Outputs per document:
  word-layout.tsv, browser-lines.json, browser.png, layout-diff.tsv; pdf mode also keeps word.pdf
`);
  process.exit(code);
}

function captureWordLayout(docx, wordTsv, wordPdf) {
  if (options.wordSource === 'pdf') {
    return run('tools/word-parity/word-print-pdf.sh', [docx, wordPdf])
      && run('python3', ['tools/word-parity/extract-pdf-text-layout.py', wordPdf, wordTsv]);
  }
  return run('tools/word-parity/word-layout-tsv.sh', [docx, wordTsv]);
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
    const x = num(item.x);
    const y = num(item.y);
    const lastX = num(item.lastX);
    const lastY = num(item.lastY);
    const textBoundaryX = num(item.textBoundaryX);
    const textBoundaryY = num(item.textBoundaryY);
    const text = item.text ?? '';
    return {
      index,
      page: int(item.page),
      adjustedPage: int(item.adjustedPage || item.page),
      section: int(item.section),
      pageCount: int(item.pageCount),
      paragraph: int(item.paragraph),
      line: int(item.line),
      xPt: x,
      yPt: y,
      x: x * 96 / 72,
      y: y * 96 / 72,
      lastX: Number.isFinite(lastX) ? lastX * 96 / 72 : NaN,
      lastY: Number.isFinite(lastY) ? lastY * 96 / 72 : NaN,
      textBoundaryX: Number.isFinite(textBoundaryX) ? textBoundaryX * 96 / 72 : NaN,
      textBoundaryY: Number.isFinite(textBoundaryY) ? textBoundaryY * 96 / 72 : NaN,
      firstCharColumn: int(item.firstCharColumn),
      startRow: tableCellValue(item.startRow),
      endRow: tableCellValue(item.endRow),
      maxRows: tableCellValue(item.maxRows),
      startColumn: tableCellValue(item.startColumn),
      endColumn: tableCellValue(item.endColumn),
      maxColumns: tableCellValue(item.maxColumns),
      atEndOfRowMarker: item.atEndOfRowMarker,
      inContentControl: item.inContentControl,
      inFieldCode: item.inFieldCode,
      inFieldResult: item.inFieldResult,
      rangeLines: int(item.rangeLines),
      rangePages: int(item.rangePages),
      inTable: item.inTable,
      table: item.table,
      row: item.row,
      column: item.column,
      style: item.style,
      text,
      norm: normalizeText(text),
    };
  });
}

function normalizeBrowserLine(line, index) {
  const text = line.Text ?? '';
  return {
    index,
    page: int(line.Page),
    adjustedPage: int(line.AdjustedPage ?? line.Page),
    section: int(line.Section),
    pageCount: int(line.PageCount),
    blockId: line.BlockId ?? '',
    blockIndex: int(line.BlockIndex),
    lineIndex: int(line.LineIndex),
    x: num(line.X),
    y: num(line.PageY ?? line.Y),
    absoluteY: num(line.Y),
    width: num(line.Width),
    height: num(line.Height),
    textBoundaryX: num(line.HorizontalPositionRelativeToTextBoundary),
    textBoundaryY: num(line.VerticalPositionRelativeToTextBoundary),
    firstCharLine: int(line.FirstCharacterLineNumber),
    firstCharColumn: int(line.FirstCharacterColumnNumber),
    inTable: line.InTable,
    table: line.Table ?? '',
    row: line.Row ?? '',
    column: line.Column ?? '',
    startRow: line.StartRow ?? '',
    endRow: line.EndRow ?? '',
    maxRows: line.MaximumRows ?? '',
    startColumn: line.StartColumn ?? '',
    endColumn: line.EndColumn ?? '',
    maxColumns: line.MaximumColumns ?? '',
    atEndOfRowMarker: line.AtEndOfRowMarker,
    inContentControl: line.InContentControl,
    inFieldCode: line.InFieldCode,
    inFieldResult: line.InFieldResult,
    rangeLines: int(line.RangeLines),
    rangePages: int(line.RangePages),
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

function buildDiffs(pairs) {
  return pairs.map((pair, i) => {
    const next = pairs[i + 1];
    const wordGap = next ? next.word.y - pair.word.y : NaN;
    const browserGap = next ? next.browser.y - pair.browser.y : NaN;
    const dx = pair.browser.x - pair.word.x;
    const dy = pair.browser.y - pair.word.y;
    const gapDy = browserGap - wordGap;
    const textBoundaryDx = pair.browser.textBoundaryX - pair.word.textBoundaryX;
    const textBoundaryDy = pair.browser.textBoundaryY - pair.word.textBoundaryY;
    const score = Math.max(abs(dx), abs(dy), Number.isFinite(gapDy) ? abs(gapDy) : 0);
    return {
      wordIndex: pair.word.index,
      browserIndex: pair.browser.index,
      wordPage: pair.word.page,
      browserPage: pair.browser.page,
      wordAdjustedPage: pair.word.adjustedPage,
      browserAdjustedPage: pair.browser.adjustedPage,
      wordSection: pair.word.section,
      browserSection: pair.browser.section,
      wordPageCount: pair.word.pageCount,
      browserPageCount: pair.browser.pageCount,
      wordParagraph: pair.word.paragraph,
      wordLine: pair.word.line,
      browserFirstCharLine: pair.browser.firstCharLine,
      browserBlock: pair.browser.blockId,
      browserLine: pair.browser.lineIndex,
      wordX: pair.word.x,
      wordY: pair.word.y,
      browserX: pair.browser.x,
      browserY: pair.browser.y,
      dx,
      dy,
      wordTextBoundaryX: pair.word.textBoundaryX,
      wordTextBoundaryY: pair.word.textBoundaryY,
      browserTextBoundaryX: pair.browser.textBoundaryX,
      browserTextBoundaryY: pair.browser.textBoundaryY,
      textBoundaryDx,
      textBoundaryDy,
      wordFirstCharColumn: pair.word.firstCharColumn,
      browserFirstCharColumn: pair.browser.firstCharColumn,
      wordGap,
      browserGap,
      gapDy,
      score,
      inTable: pair.word.inTable,
      browserInTable: pair.browser.inTable,
      table: pair.word.table,
      browserTable: pair.browser.table,
      row: pair.word.row,
      browserRow: pair.browser.row,
      column: pair.word.column,
      browserColumn: pair.browser.column,
      wordStartRow: pair.word.startRow,
      browserStartRow: pair.browser.startRow,
      wordEndRow: pair.word.endRow,
      browserEndRow: pair.browser.endRow,
      wordMaxRows: pair.word.maxRows,
      browserMaxRows: pair.browser.maxRows,
      wordStartColumn: pair.word.startColumn,
      browserStartColumn: pair.browser.startColumn,
      wordEndColumn: pair.word.endColumn,
      browserEndColumn: pair.browser.endColumn,
      wordMaxColumns: pair.word.maxColumns,
      browserMaxColumns: pair.browser.maxColumns,
      wordInContentControl: pair.word.inContentControl,
      browserInContentControl: pair.browser.inContentControl,
      wordInFieldCode: pair.word.inFieldCode,
      browserInFieldCode: pair.browser.inFieldCode,
      wordInFieldResult: pair.word.inFieldResult,
      browserInFieldResult: pair.browser.inFieldResult,
      wordRangeLines: pair.word.rangeLines,
      browserRangeLines: pair.browser.rangeLines,
      wordRangePages: pair.word.rangePages,
      browserRangePages: pair.browser.rangePages,
      style: pair.word.style,
      wordText: pair.word.text,
      browserText: pair.browser.text,
    };
  });
}

function writeReport(file, rows) {
  const headers = [
    'wordIndex', 'browserIndex', 'wordPage', 'browserPage', 'wordParagraph', 'wordLine',
    'browserBlock', 'browserLine', 'wordX', 'wordY', 'browserX', 'browserY',
    'dx', 'dy', 'wordTextBoundaryX', 'wordTextBoundaryY', 'browserTextBoundaryX',
    'browserTextBoundaryY', 'textBoundaryDx', 'textBoundaryDy',
    'wordFirstCharColumn', 'browserFirstCharColumn', 'wordGap', 'browserGap', 'gapDy', 'score',
    'wordAdjustedPage', 'browserAdjustedPage', 'wordSection', 'browserSection',
    'wordPageCount', 'browserPageCount', 'browserFirstCharLine',
    'inTable', 'browserInTable', 'table', 'browserTable', 'row', 'browserRow',
    'column', 'browserColumn', 'wordStartRow', 'browserStartRow', 'wordEndRow',
    'browserEndRow', 'wordMaxRows', 'browserMaxRows', 'wordStartColumn',
    'browserStartColumn', 'wordEndColumn', 'browserEndColumn', 'wordMaxColumns',
    'browserMaxColumns', 'wordInContentControl', 'browserInContentControl',
    'wordInFieldCode', 'browserInFieldCode', 'wordInFieldResult', 'browserInFieldResult',
    'wordRangeLines', 'browserRangeLines', 'wordRangePages', 'browserRangePages',
    'style', 'wordText', 'browserText',
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
