import { existsSync, readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname } from "node:path";
import { Writable } from "node:stream";
import { unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { deobfuscateFont } from "../src/pdf/fonts.js";
import { RenderPdfResult, renderDocxToPdf } from "../src/pdf/render.js";
import { Face } from "../src/pdf/sfnt.js";
import { bytesSource } from "../src/stream/source.js";
import { W_NS, makeDocxWithMedia } from "./helpers.js";

const require = createRequire(import.meta.url);
const carlito = require.resolve("@fontsource/carlito/files/carlito-latin-400-normal.woff");
const caladea = require.resolve("@fontsource/caladea/files/caladea-latin-400-normal.woff");
const carlitoDir = dirname(carlito);
const caladeaDir = dirname(caladea);

const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const KEY = "{01234567-89AB-CDEF-0123-456789ABCDEF}";

/** A document whose text asks for `family`, which fontTable.xml embeds in
 * one style as Carlito's sfnt, obfuscated with KEY. */
function docxEmbedding(family: string, slot: "embedRegular" | "embedBold"): Uint8Array {
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W_NS}><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/></w:rPr><w:t>Embedded face</w:t></w:r></w:p></w:body></w:document>`;
  const fontTable = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts ${W_NS} ${R_NS}><w:font w:name="${family}"><w:${slot} r:id="rId1" w:fontKey="${KEY}"/></w:font></w:fonts>`;
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.odttf"/></Relationships>`;
  const contentTypes = `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
  // Obfuscation is its own inverse.
  const obfuscated = deobfuscateFont(new Face(carlito).sfnt, KEY);
  return makeDocxWithMedia(
    { "word/document.xml": document, "word/fontTable.xml": fontTable, "word/_rels/fontTable.xml.rels": rels, "[Content_Types].xml": contentTypes },
    { "word/fonts/font1.odttf": obfuscated },
  );
}

async function render(docx: Uint8Array, fontDirs: string[]): Promise<{ text: string; result: RenderPdfResult }> {
  const chunks: Buffer[] = [];
  const out = new Writable({
    write(chunk: Buffer, _encoding, callback) {
      chunks.push(chunk);
      callback();
    },
  });
  const result = await renderDocxToPdf(bytesSource(docx), out, { fontDirs });
  return { text: Buffer.concat(chunks).toString("latin1"), result };
}

describe("embedded fonts", () => {
  it("deobfuscates as Word obfuscates: the first 32 bytes against the fontKey GUID's bytes in reverse order", () => {
    // The first 32 bytes of word/fonts/font1.odttf in a Word-written
    // document (wild3-template-nps-science-report) and its w:fontKey.
    const obfuscated = Buffer.from("7a7535711d9b02b53f4d256c2564b7c812f06c4b1d9023413f4938482673bbc9ff00", "hex");
    const plain = deobfuscateFont(obfuscated, "{8FFE3761-FC25-493F-B503-821D7135747A}");
    expect(Buffer.from(plain).toString("hex")).toBe("000100000019010000040090445349476884593a001220f400001db447444546ff00");
    expect(() => deobfuscateFont(obfuscated, "{not-a-guid}")).toThrow(/malformed font key/);
  });

  it("lists the files fontTable.xml embeds, by family, style and part", () => {
    const doc = DocxDocument.fromFiles(unzipSync(docxEmbedding("Parity Embedded Sans", "embedBold")));
    expect(doc.embeddedFonts).toEqual([{ family: "Parity Embedded Sans", bold: true, italic: false, part: "word/fonts/font1.odttf", key: KEY }]);
  });

  it("draws a family no directory holds with the face the document embeds, subset and embedded as usual", async () => {
    const { text, result } = await render(docxEmbedding("Parity Embedded Sans", "embedRegular"), [caladeaDir]);
    expect(result.fonts).toEqual(["Carlito Regular"]);
    expect(result.missingFonts).toEqual([]);
    expect(text).toMatch(/\/BaseFont \/[A-Z]{6}\+Carlito/);
    expect(text).toContain("/FontFile2");
  });

  it("prefers the embedded face to a directory's face of the same family and style", async () => {
    const { result } = await render(docxEmbedding("Caladea", "embedRegular"), [caladeaDir]);
    expect(result.fonts).toEqual(["Carlito Regular"]);
  });

  it("takes a directory's face in the asked style over an embedded face in another style, and the embedded face over the fallbacks", async () => {
    const fromDirectory = await render(docxEmbedding("Caladea", "embedBold"), [caladeaDir]);
    expect(fromDirectory.result.fonts).toEqual(["Caladea Regular"]);
    const embeddedOnly = await render(docxEmbedding("Parity Embedded Sans", "embedBold"), [caladeaDir]);
    expect(embeddedOnly.result.fonts).toEqual(["Carlito Regular"]);
    expect(embeddedOnly.result.missingFonts).toEqual([]);
  });

  const wordWritten = "/Users/alexpickett/Desktop/Projects/wordinweb-parity/apps/demo/public/fixtures/wild3-template-nps-science-report.docx";
  it.skipIf(!existsSync(wordWritten))("draws a Word-written document's text with the fonts it embeds when no directory has them", async () => {
    const { result } = await render(new Uint8Array(readFileSync(wordWritten)), [carlitoDir]);
    expect(result.fonts).toEqual(expect.arrayContaining(["Times New Roman Regular", "Times New Roman Bold", "Arial Regular", "Courier New Regular"]));
    expect(result.missingFonts).toEqual([]);
    expect(result.missingCharacters).toEqual([]);
  });
});
