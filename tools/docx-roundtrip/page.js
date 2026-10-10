// Browser side of roundtrip.mjs. Served next to the Fable output.
let OF, Docx, Word;
try {
    const main = await import('/office-formats/src/Main.js');
    OF = typeof main.default === 'function' ? main.default() : main.default;
    Docx = await import('/word/Docx.js');
    Word = (await import('/word/Word.js')).default;
} catch (e) {
    window.loadError = String(e && e.stack || e);
    throw e;
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// ---- package summary -------------------------------------------------------

const norm = (v) => {
    if (v === undefined || v === null || v === false || v === '') return null;
    if (typeof v === 'number') return Math.round(v * 1000) / 1000;
    if (Array.isArray(v)) return v.length ? v.map(norm) : null;
    if (typeof v === 'object') { const o = {}; for (const k of Object.keys(v).sort()) { const n = norm(v[k]); if (n !== null) o[k] = n; } return Object.keys(o).length ? o : null; }
    return v;
};
const pick = (o, keys) => { const r = {}; for (const k of keys) r[k] = norm(o[k]); return r; };

const RUN_KEYS = ['Bold', 'Italic', 'Underline', 'UnderlineStyle', 'UnderlineColor', 'Strike', 'DoubleStrike', 'Hidden', 'AllCaps', 'SmallCaps',
    'FontSize', 'CharacterSpacing', 'FontFamily', 'Color', 'Highlight', 'VertAlign', 'Position', 'EmphasisMark', 'Hyperlink', 'InactiveHyperlink',
    'RevisionType', 'RevisionAuthor', 'RevisionDate', 'InsertedRevisionId', 'InsertedRevisionAuthor', 'InsertedRevisionDate', 'RevisionMoveName',
    'FieldCode', 'NoteId', 'NoteType', 'CommentId', 'ExplicitColorAuto'];
const IMAGE_KEYS = ['ImageLinked', 'ImageWidth', 'ImageHeight', 'ImageFloating', 'ImageWrap', 'ImageWrapText', 'ImageWrapPolygon', 'ImagePositionX', 'ImagePositionY',
    'ImagePositionHRelativeFrom', 'ImagePositionVRelativeFrom', 'ImagePositionHAlign', 'ImagePositionVAlign', 'ImageBehindDoc', 'ImageRelativeHeight',
    'ImageDistanceLeft', 'ImageDistanceTop', 'ImageDistanceRight', 'ImageDistanceBottom', 'ImageCropLeft', 'ImageCropTop', 'ImageCropRight', 'ImageCropBottom',
    'ImageAllowOverlap', 'ImageLayoutInCell', 'ImageAlt', 'ImageTitle'];
const SHAPE_KEYS = ['Width', 'Height', 'X', 'Y', 'Geometry', 'CustomPath', 'CustomW', 'CustomH', 'Gradient', 'Rotation', 'FillColor', 'BorderColor', 'BorderWidth',
    'NoFill', 'NoLine', 'BorderDash', 'Inline', 'Wrap', 'PositionHRelativeFrom', 'PositionVRelativeFrom', 'PositionHAlign', 'PositionVAlign', 'BehindDoc',
    'RelativeHeight', 'AnchorBlock', 'DistanceTop', 'DistanceBottom', 'DistanceLeft', 'DistanceRight'];
const PARA_KEYS = ['Type', 'Style', 'Alignment', 'ListStyle', 'ListLevel', 'ListId', 'ListNumberFormat', 'ListLevelText', 'ListStart',
    'IndentLeft', 'IndentRight', 'FirstLineIndent', 'SpacingBefore', 'SpacingAfter', 'LineSpacing', 'LineSpacingRule',
    'TabStops', 'Rtl', 'BackgroundColor', 'ParagraphBorders', 'PageBreakBefore', 'PageBreakBeforeDerived', 'PageBreakAfter', 'SectionBreak',
    'SectionPageNumberStart', 'SectionPageNumberFormat', 'SectionPageSetup', 'SectionColumns', 'StyleContextualSpacing',
    'ContentControlTag', 'ContentControlAlias', 'KeepNext', 'KeepLines', 'HeaderReferences', 'FooterReferences',
    // What the paragraph itself states, as opposed to what it inherits.
    'ExplicitSpacingBefore', 'ExplicitSpacingAfter', 'ExplicitLineSpacing', 'ExplicitIndentLeft', 'ExplicitIndentRight', 'ExplicitFirstLineIndent',
    'ExplicitContextualSpacing', 'ExplicitBackgroundColor', 'ExplicitTabs', 'HasMarkProperties', 'ExplicitMarkFontSize',
    'MarkRevisionType', 'MarkRevisionId', 'MarkRevisionAuthor', 'MarkRevisionDate'];
const TABLE_KEYS = ['Floating', 'FloatingXml', 'TableStyleId', 'TableLook', 'PreferredWidth', 'PreferredWidthType', 'TableAlignment', 'LayoutMode', 'GridColumns', 'Borders', 'BackgroundColor',
    'CellSpacing', 'PaddingTop', 'PaddingRight', 'PaddingBottom', 'PaddingLeft'];
const ROW_KEYS = ['IsHeader', 'Height', 'HeightRule', 'CanSplit', 'CellSpacing', 'PaddingTop', 'PaddingRight', 'PaddingBottom', 'PaddingLeft'];
const CELL_KEYS = ['ColSpan', 'BackgroundColor', 'Width', 'PreferredWidth', 'PreferredWidthType', 'Borders', 'VerticalAlignment', 'VerticalMerge',
    'PaddingTop', 'PaddingRight', 'PaddingBottom', 'PaddingLeft'];
const SETUP_KEYS = ['PageWidth', 'PageHeight', 'MarginLeft', 'MarginRight', 'MarginTop', 'MarginBottom', 'HeaderDistance', 'FooterDistance',
    'DifferentFirstPage', 'EvenAndOddHeaders', 'ColumnCount', 'ColumnSpace', 'ColumnSeparator', 'ColumnWidths',
    'HeaderReferences', 'FooterReferences', 'PageNumberStart', 'PageNumberFormat', 'SectionStart'];

function runSummary(run) {
    if (run.IsImage) {
        const s = { image: (run.ImageSrc || '').length, ...pick(run, IMAGE_KEYS), opaque: run.OpaqueXml || null, opaqueKind: run.OpaqueKind || null };
        // What an anchor without the attribute means, which the writer spells out.
        if (s.ImageFloating) {
            if (!s.ImagePositionHAlign) s.ImagePositionX ??= 0;
            if (!s.ImagePositionVAlign) s.ImagePositionY ??= 0;
            s.ImageWrap ??= 'square';
            if (s.ImageWrap === 'square' || s.ImageWrap === 'tight' || s.ImageWrap === 'through') s.ImageWrapText ??= 'bothSides';
            s.ImagePositionHRelativeFrom ??= 'column'; s.ImagePositionVRelativeFrom ??= 'paragraph'; s.ImageRelativeHeight ??= '0';
            s.ImageDistanceLeft ??= 12; s.ImageDistanceRight ??= 12; s.ImageAllowOverlap ??= true; s.ImageLayoutInCell ??= true;
        }
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
            const props = pick(item, SHAPE_KEYS);
            // A shape that says nothing about these is written with what it then means.
            props.Geometry ??= 'rect'; props.Rotation ??= 0; props.BorderWidth ??= 1;
            if (!props.NoFill && !props.Gradient) props.FillColor = (props.FillColor ?? '#FFFFFF').toUpperCase();
            if (!props.NoLine) props.BorderColor = (props.BorderColor ?? '#808080').toUpperCase();
            if (!props.Inline) {
                props.Wrap ??= 'none'; props.RelativeHeight ??= '251658240';
                props.PositionHRelativeFrom ??= 'column'; props.PositionVRelativeFrom ??= 'paragraph';
                if (!props.PositionHAlign) props.X ??= 0;
                if (!props.PositionVAlign) props.Y ??= 0;
            } else { props.X = null; props.Y = null; }
            out.push({ path, kind: 'textbox', props });
            summarise(item.Paragraphs, path + '.p', out);
        } else {
            const props = pick(item, PARA_KEYS);
            // A numbered paragraph whose list has no definition is shown, and saved, as decimal.
            if (props.ListStyle === 'numbered') { props.ListNumberFormat ??= 'decimal'; props.ListLevelText ??= `%${props.ListLevel}.`; }
            out.push({ path, kind: 'para', props, runs: runsSummary(item.Runs) });
        }
    });
    return out;
}

function summarisePackage(pkg) {
    const out = [];
    // Every body block counts, empty paragraphs at the very end included.
    const body = [...(pkg.Paragraphs || [])];
    summarise(body, 'body', out);
    for (const [key, label] of [['HeaderParts', 'header'], ['FooterParts', 'footer']])
        (pkg[key] || []).forEach((part, i) => {
            out.push({ path: `${label}${i}`, kind: 'part', props: pick(part, ['Type']) });
            summarise(part.Paragraphs, `${label}${i}`, out);
        });
    // Footnote / endnote / comment bodies are blocks like the body's.
    for (const [key, label] of [['Footnotes', 'footnote'], ['Endnotes', 'endnote'], ['Comments', 'comment']])
        (pkg[key] || []).forEach((note) => {
            const path = `${label}#${note.Id}`;
            out.push({ path, kind: 'note', props: pick(note, ['Id', 'Kind', 'HasMark', 'MarkStyle', 'Author', 'Date', 'Initials', 'Text']) });
            summarise(note.Blocks, path, out);
        });
    // No sectPr at all means the default page; the writer then spells the default out.
    const setup = pick(pkg.PageSetup || {}, SETUP_KEYS);
    const DEFAULTS = { PageWidth: 816, PageHeight: 1056, MarginLeft: 96, MarginRight: 96, MarginTop: 96, MarginBottom: 96, HeaderDistance: 48, FooterDistance: 48, ColumnCount: 1, ColumnSpace: 0 };
    if (setup.PageWidth === null) Object.assign(setup, DEFAULTS);
    setup.ColumnCount ??= 1; setup.ColumnSpace ??= 0; setup.HeaderDistance ??= 48; setup.FooterDistance ??= 48;
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
    'w:rPr': ['ins', 'del', 'moveFrom', 'moveTo', 'rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof',
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

// ---- layout ----------------------------------------------------------------

// The document in an editor, laid out to the end: the page count and one
// fingerprint per laid-out block (Word.getLayoutFingerprint: page, y, height,
// lines, text, resolved paragraph and run formatting).
const VIEW = { scale: 1, devicePixelRatio: 1, scrollTop: 0, scrollLeft: 0, width: 900, height: 600 };
function openEditor(pkg) {
    const canvas = document.createElement('canvas');
    canvas.style.cssText = 'position:fixed;left:0;top:0';
    document.body.appendChild(canvas);
    const editor = Word.initializeCanvas(canvas);
    Word.setViewport(editor, VIEW);
    Word.importFromDocxPackage(editor, pkg);
    return { editor, canvas };
}
function closeEditor(handle) { try { Word.dispose(handle.editor); } catch { /* nothing to release */ } handle.canvas.remove(); }
function layoutOf(editor) {
    Word.completeLayout(editor);
    return { pages: Word.getPageCount(editor), blocks: Word.getLayoutFingerprint(editor) };
}
function layoutOfPackage(pkg) { const handle = openEditor(pkg); try { return layoutOf(handle.editor); } finally { closeEditor(handle); } }

const LAYOUT_TOLERANCE = 0.5; // px, for y and height
function diffLayout(label, a, b, problems) {
    if (a.pages !== b.pages) problems.push(`${label}: ${a.pages} pages became ${b.pages}`);
    if (a.blocks.length !== b.blocks.length) problems.push(`${label}: ${a.blocks.length} laid-out blocks became ${b.blocks.length}`);
    const byKind = new Map();
    const note = (kind, text) => { const e = byKind.get(kind) ?? { count: 0, first: text }; e.count++; byKind.set(kind, e); };
    for (let i = 0; i < Math.min(a.blocks.length, b.blocks.length); i++) {
        const x = a.blocks[i], y = b.blocks[i];
        const where = `block ${i} ${JSON.stringify(x.Text.slice(0, 30))}`;
        if (x.Text !== y.Text) { note('text', `${where} -> ${JSON.stringify(y.Text.slice(0, 30))}`); continue; }
        if (x.Page !== y.Page) note('page', `${where}: page ${x.Page} -> ${y.Page}`);
        if (Math.abs(x.Y - y.Y) > LAYOUT_TOLERANCE) note('y', `${where}: y ${x.Y.toFixed(2)} -> ${y.Y.toFixed(2)}`);
        if (Math.abs(x.Height - y.Height) > LAYOUT_TOLERANCE) note('height', `${where}: height ${x.Height.toFixed(2)} -> ${y.Height.toFixed(2)}`);
        if (x.LineCount !== y.LineCount) note('line count', `${where}: ${x.LineCount} lines -> ${y.LineCount}`);
        else if (x.LineHeights !== y.LineHeights) note('line heights', `${where}: ${x.LineHeights} -> ${y.LineHeights}`);
        if (x.Paragraph !== y.Paragraph) {
            const p = x.Paragraph.split(';'), q = y.Paragraph.split(';');
            p.forEach((field, j) => { if (field !== q[j]) note('paragraph ' + field.split('=')[0], `${where}: ${field} -> ${q[j]}`); });
        }
        if (x.Runs !== y.Runs) note('run format', `${where}: ${x.Runs} -> ${y.Runs}`);
    }
    for (const [kind, e] of byKind) problems.push(`${label}: ${kind} x${e.count}, first ${e.first}`);
}

// ---- document.xml, element by element ----------------------------------------

// How often each element (under its parent) and each attribute occurs.
function xmlCensus(text) {
    const counts = new Map();
    const dom = new DOMParser().parseFromString(text, 'application/xml');
    if (dom.getElementsByTagName('parsererror')[0]) return counts;
    const bump = (key) => counts.set(key, (counts.get(key) ?? 0) + 1);
    const walk = (el) => {
        bump(`<${el.nodeName}>`);
        for (const attr of el.attributes) if (!attr.name.startsWith('xmlns')) bump(`<${el.nodeName} ${attr.name}>`);
        for (const child of el.children) walk(child);
    };
    walk(dom.documentElement);
    return counts;
}
// { kind: [count before, count after] } for every kind whose count changed.
function xmlCensusDiff(before, after) {
    const a = xmlCensus(before), b = xmlCensus(after), out = {};
    for (const key of new Set([...a.keys(), ...b.keys()])) if ((a.get(key) ?? 0) !== (b.get(key) ?? 0)) out[key] = [a.get(key) ?? 0, b.get(key) ?? 0];
    return out;
}

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

// Every part as base64, for byte-for-byte comparison of the parts that are kept.
async function rawParts(bytes) {
    const zip = await window.JSZip.loadAsync(bytes);
    const out = {};
    for (const name of Object.keys(zip.files)) if (!zip.files[name].dir) out[name] = await zip.files[name].async('base64');
    return out;
}
// Parts of an object the editor does not model: they must come back unchanged.
const OPAQUE_PART = /^word\/(charts|embeddings|diagrams)\//;

// One way of saving, checked against the original: the saved document reads
// back as the original did and is laid out as it was, a second save reproduces
// the first, every written part is well formed and in schema order, the
// opaque parts and every relationship target are there.
//   regenerate: every part is written from the model alone (package.RegenerateAll),
//   as a document made in the editor is. Otherwise the save is the editor's own:
//   what was not changed goes back as the text it was read from.
async function checkSave(label, regenerate, ctx, problems) {
    const { first, s0, fp0, original, p0, raw0, mutate, expect } = ctx;
    const tag = (text) => (label ? `${label}: ${text}` : text);
    const prepare = (run) => (pkg) => { if (run) run(pkg); if (regenerate) pkg.RegenerateAll = true; };
    const bytes1 = await save(first.doc, prepare(mutate ? new Function('pkg', mutate.mutate) : null));
    const second = await open(bytes1);
    const s1 = summarisePackage(second.pkg);
    const bytes2 = await save(second.doc, prepare(null));
    const third = await open(bytes2);
    const s2 = summarisePackage(third.pkg);
    const bytes3 = await save(third.doc, prepare(null));

    if (!mutate) diff(tag('save 1 vs original'), s0, s1, problems);
    diff(tag('save 2 vs save 1 (not a fixed point)'), s1, s2, problems);
    // The saved document is laid out as the original was: the same pages, every
    // block at the same place with the same lines and the same resolved formatting.
    if (!mutate) diffLayout(tag('layout after save 1 vs original'), fp0, layoutOfPackage(second.pkg), problems);

    const p1 = await parts(bytes1), p2 = await parts(bytes2), p3 = await parts(bytes3);
    // A schema-order slip that was already in the original part is not this writer's.
    for (const part of Object.keys(p1)) {
        const found = [];
        checkXml(part, p1[part], found);
        const inherited = [];
        if (p0[part]) checkXml(part, p0[part], inherited);
        for (const problem of found) if (!inherited.includes(problem)) problems.push(tag(problem));
    }
    // Byte-level fixed point from the second save on (the first save may still
    // normalise what the reader reports differently for a written file).
    for (const part of Object.keys(p2)) if (p2[part] !== p3[part]) problems.push(tag(`${part} differs between save 2 and save 3`));
    if (Object.keys(p2).length !== Object.keys(p3).length) problems.push(tag('part list differs between save 2 and save 3'));
    const raw1 = await rawParts(bytes1);
    for (const name of Object.keys(raw0)) if (OPAQUE_PART.test(name) && raw0[name] !== raw1[name]) problems.push(tag(`${name}: ${raw1[name] === undefined ? 'missing after save' : 'changed by save'}`));
    // Every relationship target of the saved package exists, and every part has a content type.
    const types = p1['[Content_Types].xml'] || '';
    for (const [relsName, relsXml] of Object.entries(p1)) {
        if (!relsName.endsWith('.rels')) continue;
        const owner = relsName.replace(/_rels\/([^/]*)\.rels$/, '$1');
        const base = owner.includes('/') ? owner.slice(0, owner.lastIndexOf('/')) : '';
        const ids = new Set();
        for (const m of relsXml.matchAll(/<Relationship\b[^>]*>/g)) {
            const id = /\bId="([^"]*)"/.exec(m[0])?.[1];
            if (ids.has(id)) problems.push(tag(`${relsName}: relationship id ${id} twice`));
            ids.add(id);
            if (/TargetMode="External"/.test(m[0])) continue;
            const target = /Target="([^"]*)"/.exec(m[0])?.[1] ?? '';
            const resolved = target.startsWith('/') ? target.slice(1) : (base ? base + '/' : '') + target;
            const norm = []; for (const seg of resolved.split('/')) { if (seg === '..') norm.pop(); else if (seg && seg !== '.') norm.push(seg); }
            if (raw1[norm.join('/')] === undefined) problems.push(tag(`${relsName}: relationship target ${target} is not in the package`));
        }
        // Every relationship the part's XML uses is one the part has.
        const partXml = p1[owner];
        if (partXml && /\.xml$/.test(owner) && owner.startsWith('word/'))
            for (const m of partXml.matchAll(/<[^<>]*?\sr:(?:id|embed|link|pict)="([^"]*)"/g))
                if (m[1] && !ids.has(m[1])) { problems.push(tag(`${owner}: refers to relationship ${m[1]}, which the part does not have`)); break; }
    }
    for (const name of Object.keys(raw1)) {
        if (name === '[Content_Types].xml') continue;
        const ext = name.slice(name.lastIndexOf('.') + 1).toLowerCase();
        if (!new RegExp(`PartName="/${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`, 'i').test(types) && !new RegExp(`Extension="${ext}"`, 'i').test(types))
            problems.push(tag(`${name}: no content type`));
    }

    const flatText = (s) => s.filter((e) => e.runs).map((e) => e.runs.map((r) => r.text ?? '￼').join('')).join('\n');
    for (const e of expect) {
        if (e.has && !new RegExp(e.has).test(p1[e.part] ?? '')) problems.push(tag(`${e.part}: expected /${e.has}/`));
        if (e.lacks && new RegExp(e.lacks).test(p1[e.part] ?? '')) problems.push(tag(`${e.part}: unexpected /${e.lacks}/`));
        if (e.readText && !flatText(s0).includes(e.readText)) problems.push(tag(`reader: expected text ${JSON.stringify(e.readText)} in ${JSON.stringify(flatText(s0).slice(0, 300))}`));
        if (e.savedText && !flatText(s1).includes(e.savedText)) problems.push(tag(`after save: expected text ${JSON.stringify(e.savedText)} in ${JSON.stringify(flatText(s1).slice(0, 300))}`));
        if (e.same && raw0[e.same] !== raw1[e.same]) problems.push(tag(`${e.same}: ${raw1[e.same] === undefined ? 'missing after save' : 'not byte-identical after save'}`));
        if (e.absent && raw1[e.absent] !== undefined) problems.push(tag(`${e.absent}: should not be in the saved package`));
        if (e.kept && !regenerate && !new RegExp(e.kept).test(p1[e.part] ?? '')) problems.push(tag(`${e.part}: an unedited save lost /${e.kept}/`));
        if (e.read) { const got = e.read.split('.').reduce((o, k) => (o == null ? o : o[k]), { pkg: first.pkg, items: s0 }); if (JSON.stringify(norm(got)) !== JSON.stringify(e.equals)) problems.push(tag(`reader: ${e.read} = ${JSON.stringify(norm(got))}, expected ${JSON.stringify(e.equals)}`)); }
    }
    return { bytes1, bytes2, p1 };
}

// Parts an unedited save writes back exactly as they were read.
const SOURCE_PART = /^word\/(document|numbering|footnotes|endnotes|comments|styles|settings)\.xml$/;
const HEADER_FOOTER_PART = /^word\/(header|footer)\d*\.xml$/;
// The parts the document's relationships lead to.
const referencedParts = (allParts) => {
    const found = new Set();
    for (const m of (allParts['word/_rels/document.xml.rels'] ?? '').matchAll(/Target="([^"]*)"/g)) found.add(m[1].startsWith('/') ? m[1].slice(1) : 'word/' + m[1]);
    return found;
};

window.roundTrip = async (name, expect) => {
    const problems = [];
    let stats = '', saved = [], xml = '', xmlDiff = {}, xmlIdentical = false, pages = 0, changedParts = [];
    window.roundTripState = undefined;
    try {
        let original = new Uint8Array(await (await fetch('/fixture?name=' + encodeURIComponent(name))).arrayBuffer());
        // An empty fixture stands for a document made in the editor, with nothing
        // kept from a file: what it saves is the document the round trip starts from.
        if (original.length === 0) {
            const run = (text, extra = {}) => ({ Text: text, FontSize: 14.667, FontFamily: 'Calibri', ...extra });
            const fresh = Docx.importFromEditorPackage({ Paragraphs: [
                { Type: 'Heading', HeadingLevel: 1, Runs: [run('New document')], Alignment: 'left' },
                { Runs: [run('typed '), run('bold', { Bold: true, ExplicitBold: true }), run(' text')], Alignment: 'left', ListStyle: 'bullet', ListLevel: 1 },
                { Runs: [run('second paragraph')], Alignment: 'center' } ] });
            original = await save(fresh, null);
        }
        const first = await open(original);
        const s0 = summarisePackage(first.pkg);
        // Import then export with no edit at all: the same number of body blocks
        // (and of blocks in every header / footer part) comes out as went in.
        {
            const untouched = exportPackage(first.doc);
            const count = (items) => (items || []).length;
            if (count(untouched.Paragraphs) !== count(first.pkg.Paragraphs))
                problems.push(`import -> export: ${count(first.pkg.Paragraphs)} body blocks became ${count(untouched.Paragraphs)}`);
            for (const key of ['HeaderParts', 'FooterParts'])
                (first.pkg[key] || []).forEach((part, i) => {
                    const got = count(((untouched[key] || [])[i] || {}).Paragraphs);
                    if (got !== count(part.Paragraphs)) problems.push(`import -> export: ${key}[${i}] ${count(part.Paragraphs)} blocks became ${got}`);
                });
        }
        const fp0 = layoutOfPackage(first.pkg);
        const p0 = await parts(original), raw0 = await rawParts(original);
        const mutate = expect.find((e) => e.mutate);
        const ctx = { first, s0, fp0, original, p0, raw0, mutate, expect };
        // Everything from the model: what a changed block, header or note is written as.
        const model = await checkSave(mutate ? '' : 'written from the model', true, ctx, problems);
        let shown = model;
        if (!mutate) {
            // The editor's save of a document nobody changed.
            shown = await checkSave('', false, { ...ctx, expect: expect.filter((e) => e.same || e.absent || e.kept) }, problems);
            if (p0['word/document.xml'] !== undefined) {
                xmlDiff = xmlCensusDiff(p0['word/document.xml'], model.p1['word/document.xml']);
                // (A byte order mark is not written back. Headers and footers are numbered
                // in the order the document refers to them, so they are looked for by their text.)
                const text = (value) => (value ?? '').replace(/^\uFEFF/, '');
                xmlIdentical = text(p0['word/document.xml']) === text(shown.p1['word/document.xml']);
                const savedHeadersFooters = new Set(Object.keys(shown.p1).filter((part) => HEADER_FOOTER_PART.test(part)).map((part) => text(shown.p1[part])));
                // (Section properties that are not the last element of the body - not a
                // document Word writes - are put there: the body is then written from the model.)
                const bodySectionLast = /<\/w:sectPr>\s*<\/w:body>/.test(p0['word/document.xml']) || !/<w:sectPr[\s>]/.test(p0['word/document.xml'].replace(/<w:p[\s>][\s\S]*?<\/w:p>/g, ''));
                for (const part of Object.keys(p0)) {
                    if (part === 'word/document.xml' && !bodySectionLast) continue;
                    // (A style the document uses and its style part does not have is added to it.)
                    const styleIds = (value) => [...(value ?? '').matchAll(/w:styleId="([^"]*)"/g)].map((m) => m[1]).sort().join(' ');
                    if (part === 'word/styles.xml' && styleIds(p0[part]) !== styleIds(shown.p1[part])) continue;
                    if (SOURCE_PART.test(part) && shown.p1[part] !== undefined && text(shown.p1[part]) !== text(p0[part])) changedParts.push(part);
                    if (HEADER_FOOTER_PART.test(part) && referencedParts(p0).has(part) && !savedHeadersFooters.has(text(p0[part]))) changedParts.push(part);
                }
            }
        }
        saved = [toBase64(shown.bytes1), toBase64(shown.bytes2), toBase64(model.bytes1)];
        xml = shown.p1['word/document.xml'];
        window.roundTripState = { pkg: first.pkg, fp0, s0 };
        pages = fp0.pages;
        stats = `${fp0.pages} pages, ${s0.filter((e) => e.kind === 'para').length} paragraphs, ${s0.filter((e) => e.kind === 'table').length} tables, ${shown.bytes1.length} bytes`;
    } catch (e) {
        problems.push('threw: ' + String(e && e.stack || e).split('\n').slice(0, 4).join(' | '));
    }
    return { problems, stats, saved, xml, xmlDiff, xmlIdentical, changedParts, pages };
};

// A trivial edit, made with real key events by roundtrip.mjs between these two
// calls: one character typed at the start of the first line of text, then
// deleted. The document saved after it is laid out as the original.
window.editOpen = (outsideTables) => {
    const state = window.roundTripState;
    if (!state) return false;
    const handle = openEditor(state.pkg);
    state.handle = handle;
    const lines = Word.getLayoutLines(handle.editor);
    // (A line of text the document's text has: not one of a header, a footnote or a text box.)
    const bodyText = Word.getText(handle.editor);
    const line = outsideTables ? lines.find((l) => !l.InTable && (l.Text || '').trim().length > 0 && bodyText.includes(l.Text.trim()))
        : lines.find((l) => (l.Text || '').trim().length > 0) ?? lines[0];
    if (!line) { closeEditor(handle); return false; }
    Word.setSelection(handle.editor, line.BlockId, line.StartOffset, line.BlockId, line.StartOffset);
    Word.getInputElement(handle.editor).focus();
    state.text = Word.getText(handle.editor);
    state.revision = Word.getRevision(handle.editor);
    return true;
};
window.editFinish = async () => {
    const state = window.roundTripState, problems = [];
    try {
        const editor = state.handle.editor;
        if (Word.getRevision(editor) === state.revision) problems.push('trivial edit: the key events did not reach the editor');
        if (Word.getText(editor) !== state.text) problems.push('trivial edit: the text is not what it was after typing and deleting a character');
        const exported = Word.exportDocumentPackage(editor);
        const bytes = new Uint8Array(await (await OF.word.writeFromEditorPackage(exported)).arrayBuffer());
        const reopened = await open(bytes);
        diff('trivial edit, saved, vs original', state.s0, summarisePackage(reopened.pkg), problems);
        diffLayout('layout after a trivial edit and a save vs original', state.fp0, layoutOfPackage(reopened.pkg), problems);
    } catch (e) {
        problems.push('trivial edit threw: ' + String(e && e.stack || e).split('\n').slice(0, 4).join(' | '));
    }
    closeEditor(state.handle);
    return problems;
};

// A real edit (editOpen(true), then one character typed by roundtrip.mjs): the
// saved document has the character, its parts are well formed, and every
// paragraph but the one that was typed in is the text it was read from.
window.editKeptFinish = async (original, expect) => {
    const state = window.roundTripState, problems = [];
    try {
        const editor = state.handle.editor;
        const typed = Word.getText(editor);
        if (typed.length !== state.text.length + 1) problems.push(`edit: the typed character is not in the text (${JSON.stringify(state.text.slice(0, 80))} -> ${JSON.stringify(typed.slice(0, 80))}, selection ${JSON.stringify(Word.getSelection(editor))})`);
        const bytes = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const reopened = await open(bytes);
        const handle = openEditor(reopened.pkg);
        if (Word.getText(handle.editor) !== typed) problems.push('edit: the saved document does not read back with the text that was saved');
        closeEditor(handle);
        const saved = await parts(bytes);
        const originalBytes = new Uint8Array(await (await fetch('/fixture?name=' + encodeURIComponent(original))).arrayBuffer());
        const before = originalBytes.length ? await parts(originalBytes) : {};
        for (const part of Object.keys(saved)) {
            const found = [], inherited = [];
            checkXml(part, saved[part], found);
            if (before[part]) checkXml(part, before[part], inherited);
            for (const problem of found) if (!inherited.includes(problem)) problems.push('edit: ' + problem);
        }
        const was = before['word/document.xml'], is = saved['word/document.xml'];
        for (const e of expect ?? []) {
            if (e.editKeeps && !new RegExp(e.editKeeps).test(is)) problems.push(`edit: word/document.xml lost /${e.editKeeps}/`);
            if (e.editLacks && new RegExp(e.editLacks).test(is)) problems.push(`edit: word/document.xml still has /${e.editLacks}/`);
        }
        if (was !== undefined && /<\/w:sectPr>\s*<\/w:body>/.test(was)) {
            // Paragraph by paragraph (the text up to each end tag): how many of the original's are gone.
            const count = new Map();
            for (const piece of is.split('</w:p>')) count.set(piece, (count.get(piece) ?? 0) + 1);
            let lost = 0;
            for (const piece of was.replace(/^\uFEFF/, '').split('</w:p>')) { const left = count.get(piece) ?? 0; if (left > 0) count.set(piece, left - 1); else lost++; }
            // (The paragraph typed in, and the piece the root's start tag is in when the writer adds a namespace to it.)
            if (lost > 2) problems.push(`edit: ${lost} paragraphs of document.xml are not the text they were read from (one was changed)`);
        }
    } catch (e) {
        problems.push('edit threw: ' + String(e && e.stack || e).split('\n').slice(0, 4).join(' | '));
    }
    closeEditor(state.handle);
    window.roundTripState = undefined;
    return problems;
};
