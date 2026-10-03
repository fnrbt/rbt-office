// Browser side of roundtrip.mjs. Served next to the Fable output.
let OF, Docx;
try {
    const main = await import('/office-formats/src/Main.js');
    OF = typeof main.default === 'function' ? main.default() : main.default;
    Docx = await import('/word/Docx.js');
} catch (e) {
    window.loadError = String(e && e.stack || e);
    throw e;
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// ---- package summary -------------------------------------------------------

const norm = (v) => {
    if (v === undefined || v === null || v === false || v === '') return null;
    if (typeof v === 'number') return Math.round(v * 1000) / 1000;
    if (Array.isArray(v)) return v.map(norm);
    if (typeof v === 'object') { const o = {}; for (const k of Object.keys(v).sort()) { const n = norm(v[k]); if (n !== null) o[k] = n; } return Object.keys(o).length ? o : null; }
    return v;
};
const pick = (o, keys) => { const r = {}; for (const k of keys) r[k] = norm(o[k]); return r; };

const RUN_KEYS = ['Bold', 'Italic', 'Underline', 'UnderlineStyle', 'UnderlineColor', 'Strike', 'DoubleStrike', 'Hidden', 'AllCaps', 'SmallCaps',
    'FontSize', 'CharacterSpacing', 'FontFamily', 'Color', 'Highlight', 'VertAlign', 'Position', 'EmphasisMark', 'Hyperlink',
    'RevisionType', 'RevisionAuthor', 'FieldCode', 'NoteId', 'NoteType', 'CommentId'];
const IMAGE_KEYS = ['ImageWidth', 'ImageHeight', 'ImageFloating', 'ImageWrap', 'ImagePositionX', 'ImagePositionY', 'ImageDistanceLeft', 'ImageDistanceTop', 'ImageAlt'];
const PARA_KEYS = ['Type', 'Style', 'Alignment', 'ListStyle', 'ListLevel', 'ListId', 'ListNumberFormat', 'ListLevelText', 'ListStart',
    'IndentLeft', 'IndentRight', 'FirstLineIndent', 'SpacingBefore', 'SpacingAfter', 'LineSpacing', 'LineSpacingRule',
    'TabStops', 'Rtl', 'BackgroundColor', 'ParagraphBorders', 'PageBreakBefore', 'PageBreakAfter', 'SectionBreak',
    'SectionPageNumberStart', 'SectionPageSetup', 'SectionColumns', 'StyleContextualSpacing',
    'ContentControlTag', 'ContentControlAlias'];
const TABLE_KEYS = ['PreferredWidth', 'PreferredWidthType', 'TableAlignment', 'LayoutMode', 'GridColumns', 'Borders', 'BackgroundColor',
    'CellSpacing', 'PaddingTop', 'PaddingRight', 'PaddingBottom', 'PaddingLeft'];
const ROW_KEYS = ['IsHeader', 'Height', 'HeightRule', 'CanSplit', 'CellSpacing', 'PaddingTop', 'PaddingRight', 'PaddingBottom', 'PaddingLeft'];
const CELL_KEYS = ['ColSpan', 'BackgroundColor', 'Width', 'PreferredWidth', 'PreferredWidthType', 'Borders', 'VerticalAlignment', 'VerticalMerge',
    'PaddingTop', 'PaddingRight', 'PaddingBottom', 'PaddingLeft'];
const SETUP_KEYS = ['PageWidth', 'PageHeight', 'MarginLeft', 'MarginRight', 'MarginTop', 'MarginBottom', 'HeaderDistance', 'FooterDistance',
    'DifferentFirstPage', 'EvenAndOddHeaders', 'ColumnCount', 'ColumnSpace', 'ColumnSeparator', 'ColumnWidths'];

function runSummary(run) {
    if (run.IsImage) {
        const s = { image: (run.ImageSrc || '').length, ...pick(run, IMAGE_KEYS) };
        // No offset (the image is placed by wp:align, which the model does not hold) is written as offset 0.
        if (s.ImageFloating) { s.ImagePositionX ??= 0; s.ImagePositionY ??= 0; }
        return s;
    }
    const s = { text: run.Text ?? '', ...pick(run, RUN_KEYS) };
    // The reader reports w:u val="none" as UnderlineStyle "none" with Underline off.
    if (!s.Underline) { s.UnderlineStyle = null; s.UnderlineColor = null; }
    return s;
}

// The editor buffer merges neighbouring runs with the same formatting and
// drops empty ones, so compare runs in that normal form.
function runsSummary(runs) {
    const out = [];
    for (const run of runs || []) {
        const s = runSummary(run);
        if (s.text === '') continue;
        const prev = out[out.length - 1];
        if (prev && prev.text !== undefined && s.text !== undefined) {
            const a = { ...prev, text: '' }, b = { ...s, text: '' };
            if (JSON.stringify(a) === JSON.stringify(b)) { prev.text += s.text; continue; }
        }
        out.push(s);
    }
    return out;
}

function summarise(items, prefix, out) {
    (items || []).forEach((item, i) => {
        const path = `${prefix}[${i}]`;
        const type = item.Type || 'Paragraph';
        if (type === 'Table') {
            out.push({ path, kind: 'table', props: pick(item, TABLE_KEYS) });
            (item.Rows || []).forEach((row, r) => {
                out.push({ path: `${path}.row${r}`, kind: 'row', props: pick(row, ROW_KEYS) });
                (row.Cells || []).forEach((cell, c) => {
                    const cp = `${path}.row${r}.cell${c}`;
                    out.push({ path: cp, kind: 'cell', props: { ...pick(cell, CELL_KEYS), NestedTableAnchors: norm(cell.NestedTableAnchors) } });
                    summarise(cell.Paragraphs, cp + '.p', out);
                    summarise(cell.NestedTables, cp + '.nested', out);
                });
            });
        } else if (type === 'TextBox') {
            out.push({ path, kind: 'textbox', props: pick(item, ['Width', 'Height', 'X', 'Y']) });
            summarise(item.Paragraphs, path + '.p', out);
        } else {
            const props = pick(item, PARA_KEYS);
            // "No spacing anywhere" and an explicit 0 are the same paragraph; the reader
            // reports the former as 0 once the paragraph has direct formatting, which a
            // saved paragraph always has.
            props.SpacingBefore ??= 0; props.SpacingAfter ??= 0;
            // A numbered paragraph whose list has no definition is shown, and saved, as decimal.
            if (props.ListStyle === 'numbered') { props.ListNumberFormat ??= 'decimal'; props.ListLevelText ??= `%${props.ListLevel}.`; }
            out.push({ path, kind: 'para', props, runs: runsSummary(item.Runs) });
        }
    });
    return out;
}

function summarisePackage(pkg) {
    const out = [];
    // The editor drops empty / whitespace-only paragraphs at the very end of the body on import.
    const body = [...(pkg.Paragraphs || [])];
    const blank = (item) => (item.Type || 'Paragraph') !== 'Table' && (item.Type || 'Paragraph') !== 'TextBox' && runsSummary(item.Runs).every((r) => r.text !== undefined && r.text.trim() === '')
        && !item.SectionBreak && !item.PageBreakAfter && !item.ParagraphBorders && !(item.HeaderReferences || []).length && !(item.FooterReferences || []).length;
    while (body.length && blank(body[body.length - 1])) body.pop();
    summarise(body, 'body', out);
    for (const [key, label] of [['HeaderParts', 'header'], ['FooterParts', 'footer']])
        (pkg[key] || []).forEach((part, i) => {
            out.push({ path: `${label}${i}`, kind: 'part', props: pick(part, ['Type']) });
            summarise(part.Paragraphs, `${label}${i}`, out);
        });
    // No sectPr at all means the default page; the writer then spells the default out.
    const setup = pick(pkg.PageSetup || {}, SETUP_KEYS);
    const DEFAULTS = { PageWidth: 816, PageHeight: 1056, MarginLeft: 96, MarginRight: 96, MarginTop: 96, MarginBottom: 96, HeaderDistance: 48, FooterDistance: 48, ColumnCount: 1, ColumnSpace: 0 };
    if (setup.PageWidth === null) Object.assign(setup, DEFAULTS);
    setup.ColumnWidths ??= []; setup.ColumnCount ??= 1; setup.ColumnSpace ??= 0;
    out.push({ path: 'pageSetup', kind: 'setup', props: setup });
    return out;
}

function diff(label, a, b, problems) {
    if (a.length !== b.length) problems.push(`${label}: ${a.length} items became ${b.length}`);
    const byField = new Map();
    const note = (field, text) => { const e = byField.get(field) ?? { count: 0, first: text }; e.count++; byField.set(field, e); };
    for (let i = 0; i < Math.min(a.length, b.length); i++) {
        const x = a[i], y = b[i];
        if (x.kind !== y.kind) { note('kind', `${x.path}: ${x.kind} -> ${y.kind}`); continue; }
        for (const k of Object.keys(x.props)) {
            const p = JSON.stringify(x.props[k]), q = JSON.stringify(y.props[k]);
            if (p !== q) note(`${x.kind}.${k}`, `${x.path}: ${p} -> ${q}`);
        }
        if (x.runs) {
            const tx = x.runs.map((r) => r.text ?? '￼').join(''), ty = y.runs.map((r) => r.text ?? '￼').join('');
            if (tx !== ty) note('para.text', `${x.path}: ${JSON.stringify(tx)} -> ${JSON.stringify(ty)}`);
            else if (x.runs.length !== y.runs.length) note('run.count', `${x.path}: ${x.runs.length} -> ${y.runs.length} ${JSON.stringify(x.runs)} -> ${JSON.stringify(y.runs)}`);
            else x.runs.forEach((r, j) => {
                for (const k of Object.keys(r)) {
                    const p = JSON.stringify(r[k]), q = JSON.stringify(y.runs[j][k]);
                    if (p !== q) note(`run.${k}`, `${x.path}.run${j} ${JSON.stringify(r.text ?? 'image')}: ${p} -> ${q}`);
                }
            });
        }
    }
    for (const [field, e] of byField) problems.push(`${label}: ${field} x${e.count}, first ${e.first}`);
}

// ---- written XML checks ----------------------------------------------------

const ORDER = {
    'w:pPr': ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
        'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
        'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
        'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'],
    'w:rPr': ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof',
        'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd',
        'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'],
    'w:tblPr': ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing',
        'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription'],
    'w:tblPrEx': ['tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook'],
    'w:tcPr': ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark'],
    'w:sectPr': ['headerReference', 'footerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType',
        'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid'],
    'w:tblBorders': ['top', 'left', 'start', 'bottom', 'right', 'end', 'insideH', 'insideV'],
    'w:tcBorders': ['top', 'left', 'start', 'bottom', 'right', 'end', 'insideH', 'insideV', 'tl2br', 'tr2bl'],
    'w:pBdr': ['top', 'left', 'bottom', 'right', 'between', 'bar'],
    'w:tcMar': ['top', 'left', 'start', 'bottom', 'right', 'end'],
    'w:tblCellMar': ['top', 'left', 'start', 'bottom', 'right', 'end'],
    'w:tbl': ['tblPr', 'tblGrid', 'tr', 'sdt', 'bookmarkStart', 'bookmarkEnd'],
    'w:tr': ['tblPrEx', 'trPr', 'tc', 'sdt'],
    'w:lvl': ['start', 'numFmt', 'lvlRestart', 'pStyle', 'isLgl', 'suff', 'lvlText', 'lvlPicBulletId', 'legacy', 'lvlJc', 'pPr', 'rPr'],
    'w:style': ['name', 'aliases', 'basedOn', 'next', 'link', 'autoRedefine', 'hidden', 'uiPriority', 'semiHidden', 'unhideWhenUsed', 'qFormat',
        'locked', 'personal', 'personalCompose', 'personalReply', 'rsid', 'pPr', 'rPr', 'tblPr', 'trPr', 'tcPr', 'tblStylePr'],
    'w:numbering': ['numPicBullet', 'abstractNum', 'num'],
    'w:num': ['abstractNumId', 'lvlOverride'],
    'w:lvlOverride': ['startOverride', 'lvl'],
    'w:docDefaults': ['rPrDefault', 'pPrDefault'],
};
// Same group may repeat (headerReference/footerReference interleave, tr/sdt).
const SAME_RANK = [['headerReference', 'footerReference'], ['tr', 'sdt', 'bookmarkStart', 'bookmarkEnd'], ['tc', 'sdt'], ['left', 'start'], ['right', 'end']];
const INT_ATTRS = {
    'w:sz': ['w:val'], 'w:szCs': ['w:val'], 'w:spacing': ['w:val', 'w:before', 'w:after', 'w:line'], 'w:position': ['w:val'],
    'w:ind': ['w:left', 'w:right', 'w:firstLine', 'w:hanging', 'w:start', 'w:end'], 'w:tab': ['w:pos'],
    'w:pgSz': ['w:w', 'w:h'], 'w:pgMar': ['w:top', 'w:right', 'w:bottom', 'w:left', 'w:header', 'w:footer', 'w:gutter'],
    'w:gridCol': ['w:w'], 'w:tcW': ['w:w'], 'w:tblW': ['w:w'], 'w:trHeight': ['w:val'], 'w:cols': ['w:num', 'w:space'], 'w:col': ['w:w', 'w:space'],
    'w:tblCellSpacing': ['w:w'], 'w:gridSpan': ['w:val'], 'w:ilvl': ['w:val'], 'w:numId': ['w:val'], 'w:start': ['w:val'],
    'wp:extent': ['cx', 'cy'],
};

function rank(order, local) {
    for (const group of SAME_RANK) if (group.includes(local)) { const idx = group.map((n) => order.indexOf(n)).filter((n) => n >= 0); if (idx.length) return Math.min(...idx); }
    return order.indexOf(local);
}

function checkXml(part, text, problems) {
    const bad = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/.exec(text);
    if (bad) problems.push(`${part}: contains the illegal XML character U+${bad[0].charCodeAt(0).toString(16).padStart(4, '0')}`);
    const dom = new DOMParser().parseFromString(text, 'application/xml');
    const err = dom.getElementsByTagName('parsererror')[0];
    if (err) { problems.push(`${part}: not well-formed XML: ${err.textContent.replace(/\s+/g, ' ').slice(0, 160)}`); return null; }
    const seen = new Set();
    const once = (msg) => { if (!seen.has(msg)) { seen.add(msg); problems.push(`${part}: ${msg}`); } };
    const walk = (el) => {
        const order = ORDER[el.nodeName];
        if (order) {
            let last = -1, lastName = '';
            for (const child of el.children) {
                if (child.namespaceURI !== W) continue;
                const r = rank(order, child.localName);
                if (r < 0) { once(`unexpected <${child.nodeName}> in <${el.nodeName}>`); continue; }
                if (r < last) once(`<${el.nodeName}> children out of schema order: <w:${child.localName}> after <w:${lastName}>`);
                else { last = r; lastName = child.localName; }
            }
        }
        if (el.nodeName === 'w:p' || el.nodeName === 'w:r' || el.nodeName === 'w:tc') {
            const propName = { 'w:p': 'pPr', 'w:r': 'rPr', 'w:tc': 'tcPr' }[el.nodeName];
            [...el.children].forEach((child, i) => { if (child.namespaceURI === W && child.localName === propName && i !== 0) once(`<w:${propName}> is not the first child of <${el.nodeName}>`); });
        }
        if (el.nodeName === 'w:tc' && ![...el.children].some((c) => c.nodeName === 'w:p')) once('<w:tc> without a paragraph');
        if (el.nodeName === 'w:tc' && el.lastElementChild && el.lastElementChild.nodeName === 'w:tbl') once('<w:tc> ends with a table (needs a trailing paragraph)');
        if (el.nodeName === 'w:t' && el.closest && inside(el, 'w:del')) once('<w:t> inside <w:del> (must be <w:delText>)');
        if (el.nodeName === 'w:delText' && !inside(el, 'w:del')) once('<w:delText> outside <w:del>');
        if (el.nodeName === 'w:t' && /[\t\n\r]/.test(el.textContent)) once(`<w:t> holds a literal ${JSON.stringify(/[\t\n\r]/.exec(el.textContent)[0])} (needs <w:tab/>/<w:br/>)`);
        const ints = INT_ATTRS[el.nodeName];
        if (ints) for (const name of ints) {
            const v = el.getAttribute(name);
            if (v !== null && !/^-?\d+$/.test(v) && !(el.nodeName === 'w:tblW' || el.nodeName === 'w:tcW' ? /^\d+(\.\d+)?%$/.test(v) : false)) once(`<${el.nodeName} ${name}="${v}"> is not an integer`);
        }
        for (const child of el.children) walk(child);
    };
    walk(dom.documentElement);
    return dom;
}
const inside = (el, name) => { for (let p = el.parentNode; p; p = p.parentNode) if (p.nodeName === name) return true; return false; };

// ---- pipeline --------------------------------------------------------------

function exportPackage(doc) {
    if (typeof Docx.exportDocumentPackage === 'function') return Docx.exportDocumentPackage(doc);
    // Mirror of Word.fs exportDocumentPackage for builds that predate Docx.exportDocumentPackage.
    const s = doc.PageSetup;
    return {
        Paragraphs: Docx.exportToEditor(doc),
        PageSetup: { PageWidth: s.PageWidth, PageHeight: s.PageHeight, MarginLeft: s.MarginLeft, MarginRight: s.MarginRight, MarginTop: s.MarginTop,
            MarginBottom: s.MarginBottom, HeaderDistance: s.HeaderDistance, FooterDistance: s.FooterDistance,
            DifferentFirstPage: s.DifferentFirstPage, EvenAndOddHeaders: s.EvenAndOddHeaders },
        Styles: Docx.exportStylesToEditor(doc.Styles),
        HeaderParts: Docx.exportHeaderFooterPartsToEditor(doc.HeaderParts),
        FooterParts: Docx.exportHeaderFooterPartsToEditor(doc.FooterParts),
    };
}

async function open(bytes) {
    const pkg = await OF.processors.forEditorPackage(bytes);
    return { pkg, doc: Docx.importFromEditorPackage(pkg) };
}

async function save(doc, mutate) {
    const exported = exportPackage(doc);
    if (mutate) mutate(exported);
    const blob = await OF.word.writeFromEditorPackage(exported);
    return new Uint8Array(await blob.arrayBuffer());
}

const toBase64 = (bytes) => { let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000)); return btoa(s); };

async function parts(bytes) {
    const zip = await window.JSZip.loadAsync(bytes);
    const out = {};
    for (const name of Object.keys(zip.files)) if (/\.(xml|rels)$/.test(name)) out[name] = await zip.files[name].async('string');
    return out;
}

window.roundTrip = async (name, expect) => {
    const problems = [];
    let stats = '', saved = [], xml = '';
    try {
        const original = new Uint8Array(await (await fetch('/fixture?name=' + encodeURIComponent(name))).arrayBuffer());
        const first = await open(original);
        const s0 = summarisePackage(first.pkg);
        const mutate = expect.find((e) => e.mutate);
        const bytes1 = await save(first.doc, mutate ? new Function('pkg', mutate.mutate) : null);
        const second = await open(bytes1);
        const s1 = summarisePackage(second.pkg);
        const bytes2 = await save(second.doc, null);
        const third = await open(bytes2);
        const s2 = summarisePackage(third.pkg);
        const bytes3 = await save(third.doc, null);
        saved = [toBase64(bytes1), toBase64(bytes2)];

        if (!mutate) diff('save 1 vs original', s0, s1, problems);
        diff('save 2 vs save 1 (not a fixed point)', s1, s2, problems);

        const p1 = await parts(bytes1), p2 = await parts(bytes2), p3 = await parts(bytes3);
        for (const part of Object.keys(p1)) checkXml(part, p1[part], problems);
        // Byte-level fixed point from the second save on (the first save may still
        // normalise what the reader reports differently for a written file).
        for (const part of Object.keys(p2)) if (p2[part] !== p3[part]) problems.push(`${part} differs between save 2 and save 3`);
        if (Object.keys(p2).length !== Object.keys(p3).length) problems.push('part list differs between save 2 and save 3');
        xml = p1['word/document.xml'];

        const flatText = (s) => s.filter((e) => e.runs).map((e) => e.runs.map((r) => r.text ?? '￼').join('')).join('\n');
        for (const e of expect) {
            if (e.has && !new RegExp(e.has).test(p1[e.part] ?? '')) problems.push(`${e.part}: expected /${e.has}/`);
            if (e.lacks && new RegExp(e.lacks).test(p1[e.part] ?? '')) problems.push(`${e.part}: unexpected /${e.lacks}/`);
            if (e.readText && !flatText(s0).includes(e.readText)) problems.push(`reader: expected text ${JSON.stringify(e.readText)} in ${JSON.stringify(flatText(s0).slice(0, 300))}`);
            if (e.savedText && !flatText(s1).includes(e.savedText)) problems.push(`after save: expected text ${JSON.stringify(e.savedText)} in ${JSON.stringify(flatText(s1).slice(0, 300))}`);
            if (e.read) { const got = e.read.split('.').reduce((o, k) => (o == null ? o : o[k]), { pkg: first.pkg, items: s0 }); if (JSON.stringify(norm(got)) !== JSON.stringify(e.equals)) problems.push(`reader: ${e.read} = ${JSON.stringify(norm(got))}, expected ${JSON.stringify(e.equals)}`); }
        }
        stats = `${s0.filter((e) => e.kind === 'para').length} paragraphs, ${s0.filter((e) => e.kind === 'table').length} tables, ${bytes1.length} bytes`;
    } catch (e) {
        problems.push('threw: ' + String(e && e.stack || e).split('\n').slice(0, 4).join(' | '));
    }
    return { problems, stats, saved, xml };
};
