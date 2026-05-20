#!/usr/bin/env node

import fs from 'node:fs';
import path from 'node:path';

const input = process.argv[2];

if (!input) {
  console.error('Usage: node tools/word-parity/assert-docx4j-toc-parity.mjs browser-segments.json');
  process.exit(2);
}

const segments = JSON.parse(fs.readFileSync(path.resolve(input), 'utf8'));
const pageStride = Number(process.env.WORD_TOC_PAGE_STRIDE ?? 1152);

const expectedTocEntries = [
  [1, 2],
  [2, 2],
  [3, 3],
  [4, 3],
  [5, 3],
  [6, 3],
  [7, 3],
  [8, 4],
  [9, 4],
  [10, 4],
  [11, 4],
];

const expectedBodyHeadings = new Map(expectedTocEntries);

function fail(message) {
  console.error(`docx4j-toc parity assertion failed: ${message}`);
  process.exit(1);
}

function pageForY(y) {
  return Math.floor(Math.max(0, y) / pageStride) + 1;
}

function assertEqualJson(label, actual, expected) {
  const actualJson = JSON.stringify(actual);
  const expectedJson = JSON.stringify(expected);
  if (actualJson !== expectedJson) {
    fail(`${label} expected ${expectedJson}, got ${actualJson}`);
  }
}

function tocEntries() {
  const entries = [];
  for (let i = 0; i < segments.length; i += 1) {
    if ((segments[i].Text ?? '') !== 'Article ' || i + 3 >= segments.length) continue;

    const article = segments[i + 1].Text ?? '';
    const tab = segments[i + 2].Text ?? '';
    const page = segments[i + 3].Text ?? '';
    const sameBlock =
      segments[i].BlockId === segments[i + 1].BlockId
      && segments[i].BlockId === segments[i + 2].BlockId
      && segments[i].BlockId === segments[i + 3].BlockId;

    if (sameBlock && /^\d+$/.test(article) && tab === '\t' && /^\d+$/.test(page)) {
      entries.push({
        article: Number(article),
        page: Number(page),
        y: segments[i].Y,
        tabWidth: segments[i + 2].Width,
        pageX: segments[i + 3].X,
      });
    }
  }
  return entries.slice(0, expectedTocEntries.length);
}

function lineTexts() {
  const byLine = new Map();
  for (const segment of segments) {
    if (segment.Kind !== 'text') continue;
    const key = `${segment.BlockId}:${segment.Y.toFixed(1)}`;
    const line = byLine.get(key) ?? { y: segment.Y, page: pageForY(segment.Y), parts: [] };
    line.parts.push({ x: segment.X, text: segment.Text ?? '' });
    byLine.set(key, line);
  }

  return [...byLine.values()]
    .map((line) => ({
      y: line.y,
      page: line.page,
      text: line.parts
        .sort((a, b) => a.x - b.x)
        .map((part) => part.text)
        .join('')
        .replace(/\s+/g, ' ')
        .trim(),
    }))
    .filter((line) => line.text.length > 0);
}

function bodyHeadings() {
  const byBlock = new Map();
  for (const segment of segments) {
    if (segment.Kind !== 'text' || segment.Y < pageStride) continue;
    const block = byBlock.get(segment.BlockId) ?? { y: segment.Y, parts: [] };
    block.y = Math.min(block.y, segment.Y);
    block.parts.push({ x: segment.X, text: segment.Text ?? '' });
    byBlock.set(segment.BlockId, block);
  }

  const headings = [];
  for (const block of byBlock.values()) {
    const text = block.parts
      .sort((a, b) => a.x - b.x)
      .map((part) => part.text)
      .join('')
      .trim();
    const match = /^Article\s+(\d+)\s*$/.exec(text);
    if (match) {
      headings.push({ article: Number(match[1]), page: pageForY(block.y), y: block.y });
    }
  }
  return headings.sort((a, b) => a.article - b.article);
}

const actualTocEntries = tocEntries();
assertEqualJson(
  'TOC page map',
  actualTocEntries.map((entry) => [entry.article, entry.page]),
  expectedTocEntries,
);

const tocHeading = segments.find((segment) => segment.Kind === 'text' && segment.Text === 'Contents');
if (!tocHeading) fail('missing TOC heading segment: "Contents"');
if (tocHeading.Bold !== true) fail('TOC heading is not bold');
if (tocHeading.Italic !== false) fail('TOC heading is italic; expected TOCHeading w:i val="0" to override inherited Heading1 italic');
if (!String(tocHeading.FontFamily ?? '').includes('Cambria')) {
  fail(`TOC heading font is ${tocHeading.FontFamily ?? '<missing>'}, expected Cambria`);
}
if (Math.abs(tocHeading.Y - 289) > 1.5) {
  fail(`TOC heading y drifted to ${tocHeading.Y}; expected about 289`);
}

const titleStart = segments.find((segment) => segment.Kind === 'text' && segment.Text === 'United ');
if (!titleStart) fail('missing title segment: "United "');
const titleLine = segments
  .filter((segment) => segment.Kind === 'text' && segment.BlockId === titleStart.BlockId && Math.abs(segment.Y - titleStart.Y) < 0.1)
  .sort((a, b) => a.X - b.X);
const titleText = titleLine.map((segment) => segment.Text ?? '').join('').trimEnd();
if (titleText !== 'United Nations Declaration on the') {
  fail(`title first line text changed to ${JSON.stringify(titleText)}`);
}
const titleLineWidth = titleLine.reduce((right, segment) => Math.max(right, segment.X + segment.Width), 0) - titleLine[0].X;
if (Math.abs(titleLineWidth - 520.0) > 2.0) {
  fail(`title first line width is ${titleLineWidth.toFixed(1)}; expected Word-matching Cambria title metrics`);
}
if (Math.abs(titleStart.X - 65.6) > 2.0) {
  fail(`title first line x drifted to ${titleStart.X.toFixed(1)}; expected centered visible title text`);
}

for (const entry of actualTocEntries) {
  if (entry.tabWidth < 500) fail(`TOC article ${entry.article} tab leader segment is too narrow (${entry.tabWidth})`);
  if (entry.pageX < 630) fail(`TOC article ${entry.article} page number is not right-aligned (x=${entry.pageX})`);
}

const actualBodyHeadings = bodyHeadings().filter((entry) => expectedBodyHeadings.has(entry.article));
assertEqualJson(
  'body heading page map',
  actualBodyHeadings.map((entry) => [entry.article, entry.page]),
  expectedTocEntries,
);

const lines = lineTexts();
const colonialLine = lines.find((line) => line.text.startsWith('to Colonial Countries and Peoples'));
if (!colonialLine) fail('missing rendered-page-break continuation line: "to Colonial Countries..."');
if (colonialLine.page !== 2) fail(`rendered-page-break continuation starts on page ${colonialLine.page}, expected page 2`);

const expectedBodyLines = [
  ['Discrimination between human beings on the ground of race, colour or ethnic origin is an offence to', 2],
  ['3. Special concrete measures shall be taken in appropriate circumstances in order to secure adequate', 3],
  ['All States shall take effective measures to revise governmental and other public policies and to', 3],
  ['rescind laws and regulations which have the effect of creating and perpetuating racial', 3],
  ['discrimination wherever it still exists. They should pass legislation for prohibiting such', 3],
  ['discrimination and should take all appropriate measures to combat those prejudices which lead to', 3],
  ['racial discrimination.', 3],
  ['2. Everyone shall have the right to an effective remedy and protection against any discrimination he', 4],
];

for (const [expectedText, expectedPage] of expectedBodyLines) {
  const line = lines.find((candidate) => candidate.text === expectedText);
  if (!line) fail(`missing Word-matching body wrap line: ${expectedText}`);
  if (line.page !== expectedPage) {
    fail(`body wrap line "${expectedText}" is on page ${line.page}, expected page ${expectedPage}`);
  }
}

const article2 = actualBodyHeadings.find((entry) => entry.article === 2);
const article3 = actualBodyHeadings.find((entry) => entry.article === 3);
const article5 = actualBodyHeadings.find((entry) => entry.article === 5);
const article8 = actualBodyHeadings.find((entry) => entry.article === 8);
if (!article2 || article2.page !== 2) fail('Article 2 heading is not on page 2');
if (!article3 || article3.page !== 3) fail('Article 3 heading is not on page 3');
if (!article5 || article5.page !== 3) fail('Article 5 heading is not on page 3');
if (Math.abs(article5.y - (pageStride * 2) - 685) > 6) {
  fail(`Article 5 y drifted to ${(article5.y - (pageStride * 2)).toFixed(1)} on page 3; expected about 685`);
}
if (!article8 || article8.page !== 4) fail('Article 8 heading is not on page 4');

console.log('docx4j-toc parity assertion ok');
console.log(`TOC: ${actualTocEntries.map((entry) => `${entry.article}->${entry.page}`).join(', ')}`);
