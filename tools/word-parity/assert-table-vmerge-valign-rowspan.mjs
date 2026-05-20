#!/usr/bin/env node

import fs from 'node:fs';
import path from 'node:path';

const input = process.argv[2];

if (!input) {
  console.error('Usage: node tools/word-parity/assert-table-vmerge-valign-rowspan.mjs layout-diff.tsv');
  process.exit(2);
}

const rows = parseTsv(path.resolve(input));
const byText = new Map(rows.map((row) => [row.wordText, row]));
const a = byText.get('A');
const b = byText.get('B');
const c = byText.get('C');

if (!a || !b || !c) fail('expected A/B/C rows in table-vmerge-valign-rowspan layout diff');

const bVsA = num(b.dy) - num(a.dy);
const bVsC = num(b.dy) - num(c.dy);

if (Math.abs(bVsA) > 2.0 || Math.abs(bVsC) > 2.0) {
  fail(`merged B row-specific y drift: B-A=${bVsA.toFixed(2)}, B-C=${bVsC.toFixed(2)}; expected within 2px`);
}

if (Math.abs(num(b.dx)) > 6.0) {
  fail(`merged B x drift is ${num(b.dx).toFixed(2)}px; expected within 6px`);
}

console.log('table-vmerge-valign-rowspan assertion ok');
console.log(`B relative y: B-A=${bVsA.toFixed(2)}px, B-C=${bVsC.toFixed(2)}px`);

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

function fail(message) {
  console.error(`table-vmerge-valign-rowspan assertion failed: ${message}`);
  process.exit(1);
}
