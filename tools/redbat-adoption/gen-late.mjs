// late.docx: 9000 plain paragraphs, then the body of a document with a picture, then a paragraph in a
// font nothing before it uses (Georgia bold italic): the picture and the face belong to blocks that
// are read long after the first screen is shown.
import fs from "node:fs";
import { JSZip, T, S } from "./lib.mjs";
const zip = await JSZip.loadAsync(fs.readFileSync(`${T}/test-documents-wild/phpword-sample-11-readword2007.docx`));
const xml = await zip.file("word/document.xml").async("string");
const open = xml.indexOf(">", xml.indexOf("<w:body")) + 1, close = xml.lastIndexOf("</w:body>"), sect = xml.lastIndexOf("<w:sectPr", close);
console.log("pictures in the body:", (xml.match(/<w:drawing|<w:pict/g) ?? []).length, "body length", close - open);
let filler = "";
for (let i = 0; i < 9000; i++) filler += `<w:p><w:r><w:t xml:space="preserve">Filler paragraph ${i + 1}. The quick brown fox jumps over the lazy dog, again and again, so that this paragraph wraps onto a second line of the page it is on.</w:t></w:r></w:p>`;
const late = `<w:p><w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:cs="Georgia"/><w:b/><w:i/></w:rPr><w:t>LATE-FONT-MARKER in Georgia bold italic</w:t></w:r></w:p>`;
zip.file("word/document.xml", xml.slice(0, open) + filler + xml.slice(open, sect) + late + xml.slice(sect));
fs.writeFileSync(`${S}/late.docx`, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
console.log("written", fs.statSync(`${S}/late.docx`).size);
