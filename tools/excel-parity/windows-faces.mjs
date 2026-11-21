// The Windows fonts the parity tools set text in (artifacts/fonts/windows and windows-extra: Calibri, Tahoma,
// Georgia ... - tools/word-parity/fetch-windows-fonts.sh), as faces a page registers (FontFace) under their
// families, so that a browser on Linux measures and draws as with the fonts Excel and Word use.
import fs from 'node:fs';
import path from 'node:path';

// Family, weight and style of a TrueType / OpenType file (its first face) from its name and OS/2 tables.
export const faceOf = (file) => {
    const b = fs.readFileSync(file);
    let base = 0;
    if (b.toString('latin1', 0, 4) === 'ttcf') base = b.readUInt32BE(12);
    const numTables = b.readUInt16BE(base + 4);
    const tables = {};
    for (let i = 0; i < numTables; i++) { const o = base + 12 + i * 16; tables[b.toString('latin1', o, o + 4)] = b.readUInt32BE(o + 8); }
    const names = {};
    const n = tables.name;
    if (n === undefined) return null;
    const count = b.readUInt16BE(n + 2), strings = n + b.readUInt16BE(n + 4);
    for (let i = 0; i < count; i++) {
        const r = n + 6 + i * 12;
        const platform = b.readUInt16BE(r), lang = b.readUInt16BE(r + 4), id = b.readUInt16BE(r + 6), len = b.readUInt16BE(r + 8), off = b.readUInt16BE(r + 10);
        if (platform === 3 && (lang === 0x409 || names[id] === undefined)) {
            const s = b.subarray(strings + off, strings + off + len);
            let t = ''; for (let k = 0; k + 1 < s.length; k += 2) t += String.fromCharCode(s.readUInt16BE(k));
            names[id] = t;
        }
    }
    const os2 = tables['OS/2'];
    const weight = os2 !== undefined ? b.readUInt16BE(os2 + 4) : 400;
    const fsSelection = os2 !== undefined ? b.readUInt16BE(os2 + 62) : 0;
    const sub = (names[2] ?? '').toLowerCase();
    return { family: names[1], weight: weight >= 600 || sub.includes('bold') ? 'bold' : 'normal', style: (fsSelection & 1) || sub.includes('italic') ? 'italic' : 'normal' };
};

/// Every face of the directories: { family, weight, style, dir, file }.
export const facesIn = (dirs) => {
    const faces = [];
    for (const dir of dirs.filter((d) => fs.existsSync(d)))
        for (const file of fs.readdirSync(dir).filter((f) => /\.(ttf|ttc|otf)$/i.test(f)).sort()) {
            try { const f = faceOf(path.join(dir, file)); if (f?.family) faces.push({ ...f, dir, file }); } catch { }
        }
    return faces;
};
