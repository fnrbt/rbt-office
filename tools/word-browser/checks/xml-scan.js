// The XML scanner of the DOCX reader (XmlDsl.tryScan, used for a document's
// main part) against the browser's own parser (XmlDsl.parse, DOMParser):
// every XML part of the test documents scans to the same nodes, and what the
// scanner does not read is handed back. (tests/word/xml-scan.fsx is the same
// check against .NET's parser.)
import { check, note, finish, fail, bytesOf } from './common.js';
try {
    const Xml = await import('/office-formats/src/XmlDsl.js');
    const { equals } = await import('/office-formats/fable_modules/fable-library-js.4.28.0/Util.js');
    const scan = (xml) => Xml.tryScan(xml, new Set());
    const cases = [
        ['elements and attributes', '<a x="1" y=\'2\'><b/><c z="q">text</c></a>'],
        ['declaration, comments, instructions', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<!-- c --><a><?pi data?><b>t</b><!-- inside --></a>\n<!-- after -->'],
        ['references', '<a v="&lt;&gt;&amp;&quot;&apos;&#65;&#x42;&#x1F600;">&lt;x&gt; &amp; &#169; &#x1F600;</a>'],
        ['whitespace is trimmed', '<a>  <b>  padded  </b>\n  <c> </c></a>'],
        ['xml:space keeps it', '<a><t xml:space="preserve">  kept  </t><t xml:space="preserve"> </t><t xml:space="default">  not  </t></a>'],
        ['xml:space is not inherited', '<a xml:space="preserve"><b><c>  x  </c></b></a>'],
        ['CDATA', '<a><![CDATA[ <raw> & ]]></a>'],
        ['CDATA next to text', '<a>before <![CDATA[mid]]> after</a>'],
        ['comment splits text', '<a xml:space="preserve">one<!-- c -->two</a>'],
        ['line ends', '<a v="l1\r\nl2\tx">r1\r\nr2\rr3</a>'],
        ['namespaces as written', '<w:document xmlns:w="urn:w" xmlns="urn:d"><w:body w:val="1"><plain/></w:body></w:document>'],
        ['space before the end of a tag', '<a  x = "1" ><b ></b ></a >'],
        ['gt in an attribute', '<a v="1 > 0">x</a>'],
    ];
    for (const [name, xml] of cases) {
        const scanned = scan(xml);
        check(name, scanned != null && equals(scanned, Xml.parse(xml)), true);
    }
    for (const [name, xml] of [['a DTD', '<!DOCTYPE a [<!ENTITY e "x">]><a>&e;</a>'], ['an unknown entity', '<a>&nbsp;</a>'], ['tags that do not match', '<a><b></a></b>'], ['two roots', '<a/><b/>'], ['an attribute twice', '<a x="1" x="2"/>'], ['nothing', '']])
        check(`left to the parser: ${name}`, scan(xml) == null, true);
    const listing = await (await fetch('/corpus-list')).json();
    let parts = 0, scannedParts = 0, different = [];
    for (const file of listing) {
        let zip;
        try { zip = await window.JSZip.loadAsync(await bytesOf('/repo/' + file)); } catch { continue; }
        for (const name of Object.keys(zip.files)) {
            if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue;
            const xml = await zip.file(name).async('string');
            let parsed;
            try { parsed = Xml.parse(xml); } catch { if (scan(xml) != null) different.push(file + ' ' + name + ' (the parser rejects it)'); continue; }
            parts++;
            const scanned = scan(xml);
            if (scanned == null) continue;
            scannedParts++;
            if (!equals(scanned, parsed)) different.push(file + ' ' + name);
        }
    }
    check('the test documents were found', listing.length > 50 && parts > 500, true);
    check('every part scans to the nodes the parser gives', different.slice(0, 8), []);
    check('the scanner reads nearly every part', scannedParts * 100 >= parts * 99, true);
    note(`${listing.length} documents, ${scannedParts} of ${parts} parts scanned`);
} catch (e) { fail(e); }
finish();
