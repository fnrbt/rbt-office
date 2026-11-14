// A paragraph whose formatting was changed - a part of it made bold, given a
// font, a size, a colour; its spacing, alignment, list - is saved as the XML
// it was read from with the change made in it: runs split where the selection
// begins and ends (each part with the run's start tag and properties), the
// properties set in w:rPr / w:pPr where the schema puts them, attribute by
// attribute where an element is there already. What the model does not hold
// of the paragraph is still there: w14 attributes and elements, proofing
// marks, a smart tag, a field's code, a comment range, a hyperlink around its
// runs, attributes the model does not write. In a table the paragraph is
// changed in the table's own XML, and a paragraph that cannot be (a tracked
// change of its run's properties) is alone written from the model: the
// table, its rows, cells and other paragraphs stay the text they were.
// (synthetic/patched-paragraphs, tools/docx-roundtrip/fixtures.mjs.)
import { check, finish, fail, loadWord, loadFormats, bytesOf } from './common.js';
try {
    const OF = await loadFormats();
    const Word = await loadWord();
    const bytes = await bytesOf('/fixture?name=' + encodeURIComponent('synthetic/patched-paragraphs.docx'));
    const original = await (await window.JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
    const text = 'alpha beta gamma delta\tzeta\netaresult theta link one link two';
    const paragraphAt = (xml, from) => xml.slice(xml.lastIndexOf('<w:p ', xml.indexOf('alpha', from)), xml.indexOf('</w:p>', xml.indexOf('alpha', from)) + 6);
    const bodyOf = (xml) => paragraphAt(xml, 0);
    const tableOf = (xml) => xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>') + 8);
    const cellOf = (xml) => paragraphAt(xml, xml.indexOf('<w:tbl>'));
    const modelOf = (e) => JSON.stringify(Word.exportDocumentPackage(e).Paragraphs.map(function item(para) {
        if (para.Type === 'Table') return para.Rows.map((row) => row.Cells.map((c) => c.Paragraphs.map(item)));
        // (Character by character: the editor and the reader need not cut a paragraph into the same runs.)
        const characters = [];
        for (const run of para.Runs ?? []) {
            const format = JSON.stringify([run.Bold, run.Italic, run.Underline, run.FontFamily, run.FontSize, run.Color ?? null, run.Highlight ?? null, run.VertAlign ?? null, run.AllCaps, run.Hyperlink ?? null]);
            for (const c of run.IsImage ? '\uFFFC' : run.Text) characters.push(c + format);
        }
        return [para.Alignment, para.ExplicitAlignment ?? null, para.SpacingAfter ?? null, para.ListStyle, characters];
    }));
    const layoutText = (e) => JSON.stringify(Word.getLayoutFingerprint(e)).replace(/,single\//g, ',-/');
    // What the model does not hold of the paragraph (in the body: with its own id; in the cell its comment range's is 72).
    const kept = (id, comment = 7) => [`<w:p w:rsidR="00BB0001" w:rsidRDefault="00BB0003" w14:paraId="3A3B3C${id}" w14:textId="77777777">`, '<w:suppressAutoHyphens/>', '<w:rPr><w:lang w:val="de-DE"/></w:rPr></w:pPr>',
        '<w:proofErr w:type="spellStart"/>', '<w:proofErr w:type="spellEnd"/>', '<w:smartTag w:uri="urn:example" w:element="place">', '<w:instrText xml:space="preserve"> QUOTE "result" </w:instrText>',
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>', '<w:r><w:fldChar w:fldCharType="end"/></w:r>', `<w:commentRangeStart w:id="${comment}"/>`, `<w:commentRangeEnd w:id="${comment}"/>`, '<w:hyperlink r:id="rId30" w:tooltip="a tip" w:history="1">',
        '<w14:ligatures w14:val="standard"/>', 'w:eastAsia="MS Mincho"', 'w:rsidRPr="00BB0002"', 'w:rsidR="00BB0004"'];
    const tableKept = ['<w:tblLook w:val="04A0" w:firstRow="1" w:noHBand="0"/><w:tblCaption w:val="a caption"/>', '<w:tr w:rsidR="00BB0005" w14:paraId="4A4B4C01"><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>', '<w:noWrap/><w:hideMark/>', '<w:tcFitText/>', '<w:tr w:rsidR="00BB0006">'];
    // One edit of a document opened anew: saved XML and the editor model read back from it.
    const stage = async (label, edit, where, expectXml) => {
        const canvas = document.createElement('canvas');
        const editor = Word.initializeCanvas(canvas);
        Word.importFromDocxPackage(editor, await OF.processors.forEditorPackage(bytes));
        Word.completeLayout(editor);
        const ids = [];
        for (const line of Word.getLayoutLines(editor)) if (!ids.includes(line.BlockId)) ids.push(line.BlockId);
        const starts = ids.filter((id) => Word.getLayoutLines(editor).some((line) => line.BlockId === id && (line.Text ?? '').startsWith('alpha')));
        const tracked = ids.find((id) => Word.getLayoutLines(editor).some((line) => line.BlockId === id && (line.Text ?? '').startsWith('tracked')));
        const startsWith = (start) => ids.find((id) => Word.getLayoutLines(editor).some((line) => line.BlockId === id && (line.Text ?? '').startsWith(start)));
        const id = where === 'body' ? starts[0] : where === 'cell' || where === 'table' ? starts[1] : where === 'list' ? startsWith('a list item') : where === 'after' ? startsWith('after the table') : tracked;
        if (where === 'list' || where === 'after') Word.setSelection(editor, id, 2, id, 2);
        const select = (word, source = text) => { const at = source.indexOf(word); Word.setSelection(editor, id, at, id, at + word.length); };
        edit(editor, select, id);
        const out = new Uint8Array(await (await OF.word.writeFromEditorPackage(Word.exportDocumentPackage(editor))).arrayBuffer());
        const xml = await (await window.JSZip.loadAsync(out)).file('word/document.xml').async('string');
        const again = Word.initializeCanvas(document.createElement('canvas'));
        Word.importFromDocxPackage(again, await OF.processors.forEditorPackage(out));
        const same = [Word.getText(again) === Word.getText(editor), modelOf(again) === modelOf(editor), layoutText(again) === layoutText(editor)];
        if (!same[1] && window.location.search.includes('patched')) { const a = JSON.parse(modelOf(editor)), b = JSON.parse(modelOf(again)); const flat = (v) => JSON.stringify(v).split('","'); const fa = flat(a), fb = flat(b); const at = fa.findIndex((v, i) => v !== fb[i]); check(`${label}: (the first difference of the model)`, [fa[at], fb[at]], []); }
        if (!same[2] && where !== 'list' && where !== 'after') { const a = layoutText(editor), b = layoutText(again); let at = 0; while (at < a.length && a[at] === b[at]) at++; check(`${label}: (the first difference of the layout)`, [a.slice(Math.max(0, at - 200), at + 200), b.slice(Math.max(0, at - 200), at + 200)], []); }
        Word.dispose(again); Word.dispose(editor);
        const own = where === 'body' ? bodyOf(xml) : where === 'cell' ? cellOf(xml) : where === 'tracked' || where === 'table' ? tableOf(xml) : xml.slice(xml.indexOf('</w:tbl>') + 8);
        for (const [what, piece] of expectXml) {
            const found = piece instanceof RegExp ? piece.test(own) : own.includes(piece);
            // (What is there instead, when it is not.)
            check(`${label}: ${what}`, found ? true : own, true);
        }
        if (where === 'body' || where === 'cell') check(`${label}: everything the model does not hold of the paragraph is kept`, (where === 'body' ? kept('01') : kept('02', 72)).filter((piece) => !own.includes(piece)), []);
        if (where !== 'table') check(`${label}: the table's own XML is kept`, tableKept.filter((piece) => !xml.includes(piece)), []);
        if (where === 'body') check(`${label}: the table is the text it was`, xml.includes(tableOf(original)), true);
        else check(`${label}: the paragraph of the body is the text it was`, xml.includes(bodyOf(original)), true);
        if (where === 'tracked') check(`${label}: the other paragraph of the table is the text it was`, xml.includes(cellOf(original)), true);
        // (A list item whose level or list was changed is laid out with the level's indents and number once it is read again.)
        if (where === 'list' || where === 'after') check(`${label}: read again, the editor's text`, same[0], true);
        else check(`${label}: read again, the editor's text, formatting and layout`, same, [true, true, true]);
        return xml;
    };
    check('the fixture', [original.includes(kept('01')[0]), original.includes(kept('02')[0]), tableKept.filter((piece) => !original.includes(piece))], [true, true, []]);
    const regexOf = (piece) => piece.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const props = '<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:eastAsia="MS Mincho" w:cs="Georgia"/>';

    // A word made bold: its run is split in three, each with the run's start tag and properties.
    await stage('a word made bold', (editor, select) => { select('beta'); Word.toggleBold(editor); }, 'body', [
        ['the run before the word', `<w:r w:rsidRPr="00BB0002"><w:rPr>${props}<w:noProof/><w14:ligatures w14:val="standard"/></w:rPr><w:t xml:space="preserve">alpha </w:t></w:r>`],
        ['the word, bold where the schema puts it', `<w:r w:rsidRPr="00BB0002"><w:rPr>${props}<w:b/><w:noProof/><w14:ligatures w14:val="standard"/></w:rPr><w:t xml:space="preserve">beta</w:t></w:r>`],
        ['the run after the word', `<w:r w:rsidRPr="00BB0002"><w:rPr>${props}<w:noProof/><w14:ligatures w14:val="standard"/></w:rPr><w:t xml:space="preserve"> gamma</w:t></w:r><w:proofErr w:type="spellEnd"/>`],
    ]);
    // A font: the attributes the model writes are set in the element that is there, the others stay.
    await stage('a word given a font', (editor, select) => { select('alpha'); Word.setFontFamily(editor, 'Arial'); }, 'body', [
        ['w:rFonts with its w:eastAsia', '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="MS Mincho" w:cs="Arial"/><w:noProof/><w14:ligatures w14:val="standard"/></w:rPr><w:t xml:space="preserve">alpha</w:t></w:r>'],
        ['the rest of the run as it was', `<w:rPr>${props}<w:noProof/><w14:ligatures w14:val="standard"/></w:rPr><w:t xml:space="preserve"> beta gamma</w:t></w:r>`],
    ]);
    // Size, colour, highlight, capitals, superscript of a run inside a comment range.
    await stage('a run given a size, a colour, a highlight, capitals and a raised position', (editor, select) => {
        select(' theta'); Word.setFontSizePt(editor, 13); Word.setColor(editor, '#C00000'); Word.setHighlight(editor, 'yellow'); Word.toggleAllCaps(editor); Word.setVertAlign(editor, 'superscript');
    }, 'body', [
        ['all of them in the schema\'s order, in the range', '<w:commentRangeStart w:id="7"/><w:r><w:rPr><w:caps/><w:smallCaps w:val="0"/><w:color w:val="C00000"/><w:sz w:val="26"/><w:szCs w:val="26"/><w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/></w:rPr><w:t xml:space="preserve"> theta</w:t></w:r><w:commentRangeEnd w:id="7"/>'],
    ]);
    // A selection that begins after a run's tab and ends inside its second text: the tab, the break and the texts are shared out.
    await stage('a part of a run with a tab and a break made bold', (editor, select) => { select('zeta\ne'); Word.toggleBold(editor); }, 'body', [
        ['the tab, the bold part with the break, the rest', '<w:r w:rsidR="00BB0004"><w:tab/></w:r><w:r w:rsidR="00BB0004"><w:rPr><w:b/></w:rPr><w:t>zeta</w:t><w:br/><w:t>e</w:t></w:r><w:r w:rsidR="00BB0004"><w:t>ta</w:t></w:r>'],
    ]);
    // All of it: every run that shows text, also the runs in the link; not the runs of the field's code.
    await stage('the paragraph made italic', (editor, select) => { select(text); Word.toggleItalic(editor); }, 'body', [
        ['the runs of the link', '<w:hyperlink r:id="rId30" w:tooltip="a tip" w:history="1"><w:r><w:rPr><w:i/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve"> link one</w:t></w:r><w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/></w:rPr><w:t xml:space="preserve"> link two</w:t></w:r></w:hyperlink>'],
        ['the run with the tab and the break', '<w:r w:rsidR="00BB0004"><w:rPr><w:i/></w:rPr><w:tab/><w:t>zeta</w:t><w:br/><w:t>eta</w:t></w:r>'],
        ['the field\'s result', '<w:r><w:rPr><w:i/><w:noProof/></w:rPr><w:t>result</w:t></w:r>'],
    ]);
    // Properties of the paragraph: an attribute set in the element that is there, a new element where the schema puts it.
    await stage('the paragraph spaced and aligned', (editor, select) => { select('beta'); Word.setSpacingAfter(editor, 9); Word.alignRight(editor); }, 'body', [
        ['w:spacing with the measure set in it, w:jc after what was there', /<w:pPr><w:suppressAutoHyphens\/><w:spacing w:beforeAutospacing="[01]" w:after="\d+"[^>]*\/><w:jc w:val="right"\/><w:rPr><w:lang w:val="de-DE"\/><\/w:rPr><\/w:pPr>/],
    ]);
    // The level of a list item: w:ilvl in the w:numPr that is there.
    await stage('a list item moved a level in', (editor) => { Word.increaseIndent(editor); }, 'list', [
        ['the level set, the rest of the properties as they were', /<w:p w14:paraId="3A3B3C04"><w:pPr><w:keepNext\/><w:numPr><w:ilvl w:val="1"\/><w:numId w:val="3"\/><\/w:numPr><w:suppressAutoHyphens\/>(<w:ind [^>]*\/>)?<\/w:pPr><w:r><w:t>a list item<\/w:t><\/w:r><\/w:p>/],
    ]);
    // A paragraph made an item of a list the file does not have: the list is the writer's; the paragraph is its own
    // XML with the list in its properties.
    await stage('a paragraph made a bullet list item', (editor) => { Word.toggleBulletList(editor); }, 'after', [
        ['written with the list', /<w:pPr><w:numPr><w:ilvl w:val="0"\/><w:numId w:val="\d+"\/><\/w:numPr>/],
    ]);
    // Text typed and a word formatted before one save.
    await stage('text typed and a word made bold', (editor, select, id) => { Word.setSelection(editor, id, 5, id, 5); Word.insertText(editor, ' more'); select('gamma', text.replace('alpha', 'alpha more')); Word.toggleBold(editor); }, 'body', [
        ['the text typed into its run', '<w:t xml:space="preserve">alpha more beta </w:t></w:r>'],
        ['the word bold', `<w:rPr>${props}<w:b/><w:noProof/><w14:ligatures w14:val="standard"/></w:rPr><w:t xml:space="preserve">gamma</w:t></w:r><w:proofErr w:type="spellEnd"/>`],
    ]);
    // In a table: the paragraph in the table's XML.
    await stage('in a table cell, a word made bold', (editor, select) => { select('gamma'); Word.toggleBold(editor); }, 'cell', [
        ['the word bold', `<w:rPr>${props}<w:b/><w:noProof/><w14:ligatures w14:val="standard"/></w:rPr><w:t xml:space="preserve">gamma</w:t></w:r>`],
    ]);
    // A paragraph of the table whose run has a tracked change of its properties: changed in its XML, the bold
    // put in at its place among the run's properties, the tracked change (what they were before it) kept.
    await stage('in a table cell, a paragraph with a tracked change of its properties', (editor, select) => { select('format', 'tracked format here'); Word.toggleBold(editor); }, 'tracked', [
        ['the word bold, the tracked change kept', /<w:rPr><w:b\/><w:i\/><w:rPrChange w:id="9" w:author="A" w:date="2020-01-01T00:00:00Z"><w:rPr\/><\/w:rPrChange><\/w:rPr><w:t[^>]*>format<\/w:t>/],
        ['the paragraph is the one that was read', '<w:tcFitText/></w:tcPr><w:p w14:paraId="3A3B3C03">'],
    ]);

    // ---- the table's rows, columns and cells: what was read goes back as it was read ----
    const table = tableOf(original);
    const rowsOf = (xml) => xml.split('</w:tr>').slice(0, -1).map((row) => row.slice(row.indexOf('<w:tr')) + '</w:tr>');
    const [firstRow, secondRow] = rowsOf(table);
    const cellsOf = (row) => row.split('</w:tc>').slice(0, -1).map((cell) => cell.slice(cell.indexOf('<w:tc>')) + '</w:tc>');
    const tableHead = table.slice(0, table.indexOf('<w:tr'));
    check('the fixture\'s table', [rowsOf(table).length, cellsOf(firstRow).length, tableHead.includes('<w:tblCaption')], [2, 2, true]);
    // A row inserted: the rows that were read are the text they were, the new row between them - made of the row it
    // was put in after (its start tag without the row's own ids, its properties but that it is not a header row, its
    // cells' properties), with an empty paragraph in each cell made like that cell's first (its start tag without its
    // ids, its properties): none written from the model.
    await stage('a row inserted', (editor, select) => { select('gamma'); Word.insertTableRowAfter(editor); }, 'table', [
        ['the table\'s properties and grid as they were read', tableHead],
        ['the first row as it was read', firstRow], ['the second row as it was read', secondRow],
        ['the new row between them', new RegExp(regexOf(firstRow) + '<w:tr[ >].*?</w:tr>' + regexOf(secondRow))],
        ['the new row: the first row\'s start tag without its ids, its properties without the header', '<w:tr w:rsidR="00BB0005"><w:trPr><w:cantSplit/></w:trPr>'],
        ['its first cell: the first row\'s first cell\'s properties, an empty paragraph like its first', '</w:trPr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:noWrap/><w:hideMark/></w:tcPr><w:p w:rsidR="00BB0001" w:rsidRDefault="00BB0003"><w:pPr><w:suppressAutoHyphens/><w:spacing w:beforeAutospacing="1" w:after="120"/><w:rPr><w:lang w:val="de-DE"/></w:rPr></w:pPr></w:p></w:tc>'],
    ]);
    // A row deleted: the other row and the table's properties are the text they were.
    await stage('a row deleted', (editor, select) => { select('gamma'); Word.deleteTableRow(editor); }, 'table', [
        ['the table is its properties and the other row, as they were read', tableHead + secondRow + '</w:tbl>'],
    ]);
    // A column inserted: every cell that was read is the text it was, in its row's own start tag and properties.
    await stage('a column inserted', (editor, select) => { select('gamma'); Word.insertTableColumnAfter(editor); }, 'table', [
        ['the table\'s properties as they were read', table.slice(0, table.indexOf('<w:tblGrid>'))],
        ['the first row: its start tag, its properties, its cells with a new one between', new RegExp(regexOf(firstRow.slice(0, firstRow.indexOf('<w:tc>')) + cellsOf(firstRow)[0]) + '<w:tc>.*?</w:tc>' + regexOf(cellsOf(firstRow)[1] + '</w:tr>'))],
        ['the second row likewise', new RegExp(regexOf(secondRow.slice(0, secondRow.indexOf('<w:tc>')) + cellsOf(secondRow)[0]) + '<w:tc>.*?</w:tc>' + regexOf(cellsOf(secondRow)[1] + '</w:tr>'))],
    ]);
    // A cell shaded: the shading in the cell's properties, which keep what the model does not hold; the rest as it was read.
    await stage('a cell shaded', (editor, select) => { select('gamma'); Word.setSelectedTableCellStyleWithShading(editor, '#FFEEAA', '', '', 0, '', 0, '', -1, -1, -1, -1); }, 'table', [
        ['the shading where the schema puts it, beside what was there', /<w:tr w:rsidR="00BB0005" w14:paraId="4A4B4C01"><w:trPr><w:cantSplit\/><w:tblHeader\/><\/w:trPr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"\/><w:shd [^>]*w:fill="FFEEAA"[^>]*\/><w:noWrap\/><w:hideMark\/><\/w:tcPr><w:p w:rsidR="00BB0001"/],
        ['the cell\'s paragraph as it was read', cellOf(original)],
        ['the cell beside it, and the other row, as they were read', cellsOf(firstRow)[1] + '</w:tr>' + secondRow],
        ['the table\'s properties and grid as they were read', tableHead],
    ]);
} catch (e) { fail(e); }
finish();
