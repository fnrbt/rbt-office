// broken-late.docx: late.docx with malformed XML far down the body.
import fs from "node:fs";
import { JSZip, S } from "./lib.mjs";
const zip = await JSZip.loadAsync(fs.readFileSync(`${S}/late.docx`));
let xml = await zip.file("word/document.xml").async("string");
const at = xml.indexOf("Filler paragraph 7000.");
const p = xml.lastIndexOf("<w:p>", at);
xml = xml.slice(0, p) + `<w:p><w:r><w:t>broken here</w:r></w:p>` + xml.slice(p);
zip.file("word/document.xml", xml);
fs.writeFileSync(`${S}/broken-late.docx`, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
