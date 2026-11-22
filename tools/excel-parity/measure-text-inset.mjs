// Where a cell's text stands between black marker cells, in pictures of the grid (make-text-inset-fixture.py's
// workbooks, shot by shoot-sheets.ps1 and shoot-browser-sheets.mjs):
//   node tools/excel-parity/measure-text-inset.mjs A.png [B.png]
// Each band of marker cells (rows of black runs at least 12 px wide) is a row of cases; each pair of runs a case:
// the cell between them, its ink's first and last column (grey < 160) from the left run's last black pixel and to
// the right run's first. With two pictures the cases are listed side by side, with the differences (B - A).
import fs from 'node:fs';
import zlib from 'node:zlib';

export function readPng(file) {
    const buf = fs.readFileSync(file);
    let pos = 8, width = 0, height = 0, depth = 0, type = 0;
    const idat = [];
    while (pos < buf.length) {
        const len = buf.readUInt32BE(pos), kind = buf.toString('ascii', pos + 4, pos + 8), data = buf.subarray(pos + 8, pos + 8 + len);
        if (kind === 'IHDR') { width = data.readUInt32BE(0); height = data.readUInt32BE(4); depth = data[8]; type = data[9]; }
        else if (kind === 'IDAT') idat.push(data);
        else if (kind === 'IEND') break;
        pos += 12 + len;
    }
    if (depth !== 8) throw new Error(`${file}: bit depth ${depth}`);
    const channels = { 0: 1, 2: 3, 4: 2, 6: 4 }[type];
    const raw = zlib.inflateSync(Buffer.concat(idat));
    const stride = width * channels, out = new Uint8Array(width * height * channels);
    let prev = new Uint8Array(stride);
    for (let y = 0; y < height; y++) {
        const f = raw[y * (stride + 1)], line = raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1)), cur = out.subarray(y * stride, (y + 1) * stride);
        for (let x = 0; x < stride; x++) {
            const a = x >= channels ? cur[x - channels] : 0, b = prev[x], c = x >= channels ? prev[x - channels] : 0;
            let v = line[x];
            if (f === 1) v += a; else if (f === 2) v += b; else if (f === 3) v += (a + b) >> 1;
            else if (f === 4) { const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c); v += pa <= pb && pa <= pc ? a : pb <= pc ? b : c; }
            cur[x] = v & 255;
        }
        prev = cur;
    }
    const gray = new Uint8Array(width * height);
    for (let i = 0; i < width * height; i++) {
        const o = i * channels;
        gray[i] = channels >= 3 ? Math.round((out[o] * 299 + out[o + 1] * 587 + out[o + 2] * 114) / 1000) : out[o];
    }
    return { width, height, gray, rgb: out, channels };
}

const runsOf = (img, y, dark = 30, min = 12) => {
    const runs = [];
    let start = -1;
    for (let x = 0; x <= img.width; x++) {
        const black = x < img.width && img.gray[y * img.width + x] < dark;
        if (black && start < 0) start = x;
        if (!black && start >= 0) { if (x - start >= min) runs.push([start, x - 1]); start = -1; }
    }
    return runs;
};

export function measure(file) {
    const img = readPng(file);
    const bands = [];
    let band = null;
    for (let y = 0; y < img.height; y++) {
        const runs = runsOf(img, y);
        if (runs.length >= 2) { if (!band) { band = { top: y, bottom: y, runs }; bands.push(band); } else band.bottom = y; }
        else band = null;
    }
    const cases = [];
    for (const b of bands) {
        if (b.bottom - b.top < 6) continue;
        for (let i = 0; i + 1 < b.runs.length; i += 2) {
            const left = b.runs[i][1], right = b.runs[i + 1][0];
            let inkL = Infinity, inkR = -Infinity, inkT = Infinity, inkB = -Infinity;
            for (let y = b.top + 1; y < b.bottom; y++)
                for (let x = left + 1; x < right; x++)
                    if (img.gray[y * img.width + x] < 160) { inkL = Math.min(inkL, x); inkR = Math.max(inkR, x); inkT = Math.min(inkT, y); inkB = Math.max(inkB, y); }
            const profile = new Float64Array(right - left - 1);
            for (let y = b.top + 1; y < b.bottom; y++)
                for (let x = left + 1; x < right; x++) profile[x - left - 1] += 255 - img.gray[y * img.width + x];
            cases.push({ col: i / 2, top: b.top, profile, width: right - left - 1, height: b.bottom - b.top + 1,
                l: Number.isFinite(inkL) ? inkL - left : null, r: Number.isFinite(inkR) ? right - inkR : null,
                ink: Number.isFinite(inkL) ? inkR - inkL + 1 : 0, b: Number.isFinite(inkB) ? b.bottom - inkB : null });
        }
    }
    // (Down the first column of cases, then the second.)
    return cases.sort((p, q) => p.col - q.col || p.top - q.top);
}

if (import.meta.url === `file://${process.argv[1]}`) {
    const [a, b] = process.argv.slice(2);
    const ca = measure(a), cb = b ? measure(b) : null;
    // (The shift of B's ink that best matches A's, from the left marker and from the right one.)
    const shift = (p, q, fromRight) => {
        let best = 0, bestErr = Infinity;
        for (let s = -20; s <= 20; s++) {
            let err = 0;
            for (let x = 0; x < p.length; x++) {
                const xq = fromRight ? x + (q.length - p.length) + s : x + s;
                err += Math.abs(p[x] - (xq >= 0 && xq < q.length ? q[xq] : 0));
            }
            if (err < bestErr) { bestErr = err; best = s; }
        }
        return best;
    };
    const fmt = (c) => c ? `w${c.width} h${c.height} L${c.l} R${c.r} ink${c.ink} B${c.b}` : '-';
    for (let i = 0; i < Math.max(ca.length, cb?.length ?? 0); i++) {
        const p = ca[i], q = cb?.[i];
        const d = p && q && p.l !== null && q.l !== null ? `  dL ${q.l - p.l} dR ${q.r - p.r} dInk ${q.ink - p.ink} dW ${q.width - p.width} sL ${shift(p.profile, q.profile, false)} sR ${shift(p.profile, q.profile, true)}` : '';
        console.log(`${i}\t${fmt(p)}${cb ? '\t' + fmt(q) + d : ''}`);
    }
}
