// Opens saved documents in LibreOffice, headless, and compares what it reads
// with what it reads from the originals: a second reader for the files the
// editors write, where every other check re-reads them with this repository's
// own reader.
//
//   node tools/docx-roundtrip/roundtrip.mjs --out DIR      (keeps the documents)
//   node tools/libreoffice-check/check.mjs DIR [--verbose]
//
// DIR holds <name>.original.<ext> and its saves, <name>.<save>.<ext>
// (roundtrip.mjs writes .pass1, .pass2 and .model), ext docx or xlsx. Each
// file is converted to text (a workbook: its first sheet, as CSV). A save
// fails when LibreOffice does not open it, or when it opens the original and
// reads other text from the save. A .model save is written from the model
// alone and may differ: its differences are listed, not counted, unless
// --strict.
//
// Environment: SOFFICE (default "soffice").
// Exit code is non-zero when any save fails.

import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';

const args = process.argv.slice(2);
const verbose = args.includes('--verbose');
const strict = args.includes('--strict');
const dir = args.find(a => !a.startsWith('--'));
if (!dir || !fs.existsSync(dir)) { console.error('usage: check.mjs DIR [--verbose] [--strict]'); process.exit(2); }
const soffice = process.env.SOFFICE || 'soffice';

const filters = { docx: ['txt:Text (encoded):UTF8', 'txt'], xlsx: ['csv:Text - txt - csv (StarCalc):44,34,76,1', 'csv'] };
const work = fs.mkdtempSync(path.join(os.tmpdir(), 'libreoffice-check-'));

// One soffice run converts a batch; a file it cannot open leaves no output.
// A document LibreOffice hangs on takes its batch with it, so what a batch
// left unconverted is tried again alone, each with its own time limit.
let runs = 0;
function soffice1(ext, out, paths, timeout) {
    const [filter] = filters[ext];
    const profile = 'file://' + path.join(work, 'profile-' + runs++);
    const run = spawnSync(soffice, ['--headless', '--norestore', `-env:UserInstallation=${profile}`,
        '--convert-to', filter, '--outdir', out, ...paths], { encoding: 'utf8', timeout, killSignal: 'SIGKILL' });
    if (run.error && run.error.code !== 'ETIMEDOUT') { console.error(`${soffice}: ${run.error.message}`); process.exit(2); }
    return !run.error;
}

const timedOut = new Set();
function convert(ext, files) {
    const outExt = filters[ext][1];
    const out = path.join(work, ext);
    fs.mkdirSync(out, { recursive: true });
    const produced = f => path.join(out, f.slice(0, -ext.length) + outExt);
    for (let i = 0; i < files.length; i += 40) {
        const batch = files.slice(i, i + 40);
        soffice1(ext, out, batch.map(f => path.join(dir, f)), 240000);
        for (const f of batch.filter(f => !fs.existsSync(produced(f))))
            if (!soffice1(ext, out, [path.join(dir, f)], 90000)) timedOut.add(f);
    }
    const read = new Map();
    for (const f of files) if (fs.existsSync(produced(f))) read.set(f, fs.readFileSync(produced(f), 'utf8'));
    return read;
}

// Runs of white space are one space: line ends and tabs between cells differ
// with how a paragraph or a table is written, not with what it says.
const normal = text => text.replace(/﻿/g, '').replace(/\s+/g, ' ').trim();

function firstDifference(a, b) {
    let i = 0;
    while (i < a.length && i < b.length && a[i] === b[i]) i++;
    const from = Math.max(0, i - 30);
    return `at ${i}: original "${a.slice(from, i + 40)}" / save "${b.slice(from, i + 40)}"`;
}

let failed = 0, passed = 0, listed = 0, skipped = 0;
for (const ext of Object.keys(filters)) {
    const files = fs.readdirSync(dir).filter(f => f.endsWith('.' + ext)).sort();
    if (files.length === 0) continue;
    const read = convert(ext, files);
    const originals = files.filter(f => f.endsWith(`.original.${ext}`));
    for (const original of originals) {
        const name = original.slice(0, -`.original.${ext}`.length);
        const saves = files.filter(f => f.startsWith(name + '.') && f !== original
            && !f.slice(name.length + 1, -ext.length - 1).includes('.'));
        if (!read.has(original)) {
            // Nothing to compare with; the saves must still open.
            if (verbose) console.log(`  ${original}: LibreOffice ${timedOut.has(original) ? 'hangs on' : 'does not open'} the original`);
        }
        for (const save of saves) {
            const kind = save.slice(name.length + 1, -ext.length - 1);
            if (!read.has(save)) {
                if (read.has(original)) { failed++; console.log(`FAIL ${save}: LibreOffice ${timedOut.has(save) ? 'hangs on it' : 'does not open it'}`); }
                else { skipped++; if (verbose) console.log(`  ${save}: not opened, nor is its original`); }
                continue;
            }
            if (!read.has(original)) { passed++; continue; }
            const a = normal(read.get(original)), b = normal(read.get(save));
            if (a === b) { passed++; if (verbose) console.log(`  ok ${save}`); continue; }
            if (kind === 'model' && !strict) { listed++; console.log(`  differs ${save}: ${firstDifference(a, b)}`); }
            else { failed++; console.log(`FAIL ${save}: other text ${firstDifference(a, b)}`); }
        }
    }
}
fs.rmSync(work, { recursive: true, force: true });
console.log(`LibreOffice check: ${passed} passed, ${failed} failed, ${listed} model saves differ, ${skipped} skipped`);
if (passed + failed + listed === 0) { console.error('no <name>.original.docx / .xlsx with saves in ' + dir); process.exit(2); }
console.log(failed === 0 ? 'ALL PASSED' : 'FAILED');
process.exit(failed === 0 ? 0 : 1);
