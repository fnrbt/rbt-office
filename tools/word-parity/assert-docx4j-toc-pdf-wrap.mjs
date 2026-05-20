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
    ['-f', '2', '-l', '4', '-layout', pdfPath, '-'],
    { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] },
  );
} catch (error) {
  fail(`pdftotext failed for ${pdfPath}: ${error.message}`);
}

const lines = extracted
  .split(/\r?\n/)
  .map((line) => line.replace(/\s+/g, ' ').trim())
  .filter(Boolean);

const expectedLines = [
  'the prejudice of racial superiority or on racial hatred, besides constituting a violation of fundamental',
  '1. Solemnly affirms the necessity of speedily eliminating racial discrimination throughout the world,',
  'Discrimination between human beings on the ground of race, colour or ethnic origin is an offence to',
  '3. Special concrete measures shall be taken in appropriate circumstances in order to secure adequate',
  'All States shall take effective measures to revise governmental and other public policies and to',
  'rescind laws and regulations which have the effect of creating and perpetuating racial',
  'discrimination wherever it still exists. They should pass legislation for prohibiting such',
  'discrimination and should take all appropriate measures to combat those prejudices which lead to',
  'racial discrimination.',
  '2. Everyone shall have the right to an effective remedy and protection against any discrimination he',
  'may suffer on the ground of race, colour or ethnic origin with respect to his fundamental rights and',
];

for (const expectedLine of expectedLines) {
  if (!lines.includes(expectedLine)) {
    const pattern = expectedLine.split(' ').slice(0, 6).join('|');
    const nearby = lines.filter((line) => new RegExp(pattern).test(line)).join(' / ');
    fail(`missing Word-matching line: ${expectedLine}. Nearby extracted lines: ${nearby || '<none>'}`);
  }
}

const worldLead = lines.find((line) => line.startsWith('world, in all its forms'));
if (worldLead) {
  fail(`line starts with "world,"; Word keeps "world," on the prior line: ${worldLead}`);
}

const badStarts = [
  'fundamental human rights,',
  'to human dignity and shall',
  'adequate development or protection',
  'rescind laws and regulations which have the effect of creating and perpetuating racial discrimination',
  'wherever it still exists. They should pass legislation for prohibiting such discrimination and should',
  'he may suffer on the ground',
  'and freedoms through independent',
];

for (const badStart of badStarts) {
  const line = lines.find((candidate) => candidate.startsWith(badStart));
  if (line) fail(`line starts with "${badStart}"; Word keeps that text on the prior line: ${line}`);
}

const expectedSecondLinePrefix =
  'in all its forms and manifestations, and of securing understanding of and respect for the';
const secondLine = lines.find((line) => line.startsWith(expectedSecondLinePrefix));
if (!secondLine) {
  const nearby = lines.filter((line) => /in all its forms|dignity/.test(line)).join(' / ');
  fail(`missing expected continuation line after "world,". Nearby extracted lines: ${nearby || '<none>'}`);
}

console.log('docx4j-toc PDF wrap assertion ok');
