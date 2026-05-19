#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(process.env.OFFICE_ROOT ?? process.cwd());
const htmlPath = path.resolve(root, process.argv[2] ?? 'docs/excel-parity-fixes.html');
const html = fs.readFileSync(htmlPath, 'utf8');
const htmlDir = path.dirname(htmlPath);

const tbodyMatch = html.match(/<tbody\b[^>]*>([\s\S]*?)<\/tbody>/i);
if (!tbodyMatch) fail('missing <tbody>');

const rows = tbodyMatch[1].match(/<tr\b[^>]*>[\s\S]*?<\/tr>/gi) ?? [];
const errors = [];
let imageCount = 0;

rows.forEach((row, index) => {
  const rowNum = index + 1;
  const title = plainText((row.match(/<strong\b[^>]*>([\s\S]*?)<\/strong>/i) ?? [null, `row ${rowNum}`])[1]);
  const cells = row.match(/<td\b[^>]*>[\s\S]*?<\/td>/gi) ?? [];
  if (cells.length !== 5) {
    errors.push(`row ${rowNum} (${title}): expected 5 cells, found ${cells.length}`);
    return;
  }

  [
    ['Excel Reference', cells[1]],
    ['Ours After', cells[2]],
    ['Ours Before', cells[3]],
  ].forEach(([label, cell]) => {
    const imgs = [...cell.matchAll(/<img\b[^>]*\bsrc="([^"]+)"/gi)];
    const hasPlaceholder = /class="[^"]*\bplaceholder\b/i.test(cell);
    if (imgs.length === 0 && !hasPlaceholder) {
      errors.push(`row ${rowNum} (${title}): ${label} cell has no image or placeholder`);
    }
    imgs.forEach(match => {
      imageCount += 1;
      const src = match[1];
      const target = path.resolve(htmlDir, src);
      if (!fs.existsSync(target)) {
        errors.push(`row ${rowNum} (${title}): missing image ${src}`);
      }
    });
  });
});

if (errors.length > 0) {
  for (const error of errors) console.error(error);
  process.exit(1);
}

console.log(`gallery ok: rows=${rows.length} images=${imageCount} file=${path.relative(root, htmlPath)}`);

function plainText(value) {
  return value
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function fail(message) {
  console.error(message);
  process.exit(1);
}
