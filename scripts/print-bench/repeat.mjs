// A longer fixture: the body's content repeated k times before its own sectPr.
//   node repeat.mjs <in.docx> <k> <out.docx>
import { readFileSync, writeFileSync } from "node:fs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const [input, k, output] = process.argv.slice(2);
const files = unzipSync(new Uint8Array(readFileSync(input)));
const xml = strFromU8(files["word/document.xml"]);
const open = xml.indexOf("<w:body");
const bodyStart = xml.indexOf(">", open) + 1;
const bodyEnd = xml.lastIndexOf("</w:body>");
const body = xml.slice(bodyStart, bodyEnd);
const sectPrAt = body.lastIndexOf("<w:sectPr");
const content = sectPrAt === -1 ? body : body.slice(0, sectPrAt);
const sectPr = sectPrAt === -1 ? "" : body.slice(sectPrAt);
const repeated = xml.slice(0, bodyStart) + content.repeat(Number(k)) + sectPr + xml.slice(bodyEnd);
files["word/document.xml"] = strToU8(repeated);
writeFileSync(output, zipSync(files, { level: 6 }));
console.log(`${output}: document.xml ${Math.round(repeated.length / 1048576)} MB`);
