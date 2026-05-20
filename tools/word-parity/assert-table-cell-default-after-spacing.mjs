#!/usr/bin/env node

import fs from 'node:fs';
import path from 'node:path';

const input = process.argv[2];

if (!input) {
  console.error('Usage: node tools/word-parity/assert-table-cell-default-after-spacing.mjs layout-diff.tsv');
  process.exit(2);
}

const rows = parseTsv(path.resolve(input));
const byText = new Map(rows.map((row) => [textKey(row.wordText), row]));

for (const text of ['Alpha', 'Beta', 'Gamma', 'Delta', 'After table']) {
  if (!byText.has(text)) fail(`missing row for ${JSON.stringify(text)}`);
}

const suspicious = rows.filter((row) => num(row.score) >= 6.0);
if (suspicious.length > 0) {
  fail(`layout drift >=6px: ${suspicious.map((row) => `${row.wordText}:${num(row.score).toFixed(2)}`).join(', ')}`);
}

const beta = byText.get('Beta');
const delta = byText.get('Delta');
for (const row of [beta, delta]) {
  if (Math.abs(num(row.gapDy)) > 1.0) {
    fail(`${row.wordText} row pitch gapDy=${num(row.gapDy).toFixed(2)}px; expected within 1px`);
  }
}

const after = byText.get('After table');
if (Math.abs(num(after.dy)) > 5.0) {
  fail(`After table dy=${num(after.dy).toFixed(2)}px; expected within 5px`);
}

console.log('table-cell-default-after-spacing assertion ok');
console.log(`row pitch gapDy: Beta=${num(beta.gapDy).toFixed(2)}px, Delta=${num(delta.gapDy).toFixed(2)}px; After dy=${num(after.dy).toFixed(2)}px`);

function parseTsv(file) {
  const lines = fs.readFileSync(file, 'utf8').trim().split(/\r?\n/);
  const headers = lines[0].split('\t');
  return lines.slice(1).map((line) => {
    const cells = line.split('\t');
    return Object.fromEntries(headers.map((header, index) => [header, cells[index] ?? '']));
  });
}

function num(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function textKey(value) {
  return String(value ?? '').trimEnd();
}

function fail(message) {
  console.error(`table-cell-default-after-spacing assertion failed: ${message}`);
  process.exit(1);
}
