#!/usr/bin/env node

import { execFileSync } from 'node:child_process';
import path from 'node:path';

const input = process.argv[2];

if (!input) {
  console.error('Usage: node tools/word-parity/assert-docx4j-toc-pdf-wrap.mjs ours.pdf');
  process.exit(2);
}

function fail(message) {
  console.error(`docx4j-toc PDF wrap assertion failed: ${message}`);
  process.exit(1);
}

const pdfPath = path.resolve(input);
let extracted;

try {
  extracted = execFileSync(
    'pdftotext',
    ['-f', '2', '-l', '2', '-layout', pdfPath, '-'],
    { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] },
  );
} catch (error) {
  fail(`pdftotext failed for ${pdfPath}: ${error.message}`);
}

const lines = extracted
  .split(/\r?\n/)
  .map((line) => line.replace(/\s+/g, ' ').trim())
  .filter(Boolean);

const expectedFirstLine =
  '1. Solemnly affirms the necessity of speedily eliminating racial discrimination throughout the world,';
const expectedSecondLinePrefix =
  'in all its forms and manifestations, and of securing understanding of and respect for the';

if (!lines.includes(expectedFirstLine)) {
  const nearby = lines.filter((line) => /Solemnly|world,|in all its forms/.test(line)).join(' / ');
  fail(`body paragraph wrap does not match Word. Nearby extracted lines: ${nearby || '<none>'}`);
}

const worldLead = lines.find((line) => line.startsWith('world, in all its forms'));
if (worldLead) {
  fail(`line starts with "world,"; Word keeps "world," on the prior line: ${worldLead}`);
}

const secondLine = lines.find((line) => line.startsWith(expectedSecondLinePrefix));
if (!secondLine) {
  const nearby = lines.filter((line) => /in all its forms|dignity/.test(line)).join(' / ');
  fail(`missing expected continuation line after "world,". Nearby extracted lines: ${nearby || '<none>'}`);
}

console.log('docx4j-toc PDF wrap assertion ok');
