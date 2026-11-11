// oracle.json is created by the human reviewer after running the existing
// Word and Excel open-all.vbs scripts on word/docs and excel/docs respectively.
// Required schema:
// { "commit": "<commit file>", "packageSha256": "<tarball sha256>",
//   "manifestSha256": "<SHA256SUMS sha256>", "reviewer": "<name>",
//   "reviewedAt": "<ISO date>", "wordVersion": "<Office version>",
//   "excelVersion": "<Office version>", "comparison": "<review notes>",
//   "wordAndExcelCompared": true }
// The reviewer must compare original, saved and edited outputs in the real
// applications (including intended changes and preservation), not merely open
// them. The logs and corpus are retained with the attestation in the bundle.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';

const dir = path.resolve(process.argv[2]);
const read = (name) => fs.readFileSync(path.join(dir, name));
const sha = (bytes) => crypto.createHash('sha256').update(bytes).digest('hex');
const fail = (message) => { throw new Error(`release oracle: ${message}`); };
const evidence = JSON.parse(read('oracle.json'));
const version = read('version').toString().trim();
if (evidence.commit !== read('commit').toString().trim() ||
    evidence.packageSha256 !== sha(read(`rbt-office-${version}.tgz`)) ||
    evidence.manifestSha256 !== sha(read('SHA256SUMS'))) fail('commit, package or corpus manifest binding mismatch');
for (const field of ['reviewer', 'reviewedAt', 'wordVersion', 'excelVersion', 'comparison'])
    if (typeof evidence[field] !== 'string' || !evidence[field].trim()) fail(`missing ${field}`);
if (!Number.isFinite(Date.parse(evidence.reviewedAt)) || evidence.wordAndExcelCompared !== true)
    fail('missing review date or comparison attestation');
const manifest = read('SHA256SUMS').toString();
for (const [app, ext] of [['word', '.docx'], ['excel', '.xlsx']]) {
    const names = fs.readdirSync(path.join(dir, app, 'docs')).sort();
    if (!names.length || names.some((n) => !n.toLowerCase().endsWith(ext))) fail(`${app} corpus is empty or contains unexpected files`);
    const normalized = names.map((n) => n.toLowerCase());
    if (new Set(normalized).size !== names.length) fail(`${app} corpus has case-colliding names`);
    const originals = names.filter((n) => n.endsWith(`.original${ext}`));
    if (!originals.length) fail(`${app} corpus has no originals`);
    for (const original of originals) {
        const base = original.slice(0, -`.original${ext}`.length);
        for (const variant of app === 'word' ? ['pass1', 'pass2', 'model'] : ['save', 'edit', 'insert'])
            if (!names.includes(`${base}.${variant}${ext}`)) fail(`${app} missing ${variant} for ${original}`);
    }
    for (const name of names) {
        const line = `${sha(read(`${app}/docs/${name}`))}  ${app}/docs/${name}`;
        if (!manifest.split('\n').includes(line)) fail(`${app} corpus is not completely bound to manifest: ${name}`);
    }
    const lines = read(`${app}/opened.tsv`).toString().replace(/^\uFEFF/, '').trim().split(/\r?\n/);
    if (lines.pop() !== 'DONE') fail(`${app} log is incomplete`);
    const seen = new Set();
    for (const line of lines) {
        const fields = line.split('\t');
        const name = fields[0].toLowerCase();
        if (!normalized.includes(name) || seen.has(name) || fields[1] !== 'ok' || fields.length < (app === 'word' ? 4 : 7))
            fail(`${app} log contains a failure, duplicate or unknown file: ${line}`);
        seen.add(name);
    }
    if (seen.size !== names.length) fail(`${app} log does not cover the complete corpus`);
}
console.log('Word and Excel oracle evidence matches the release bundle');
