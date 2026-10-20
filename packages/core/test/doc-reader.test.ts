import { readFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { PassThrough } from "node:stream";
import * as CFB from "cfb";
import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { docToDocxFiles, openDoc } from "../src/doc/convert.js";
import { docToDocx } from "../src/doc/hooks.js";
import { isWordDocumentFile, openDocStreams } from "../src/doc/streams.js";
import { parseFib } from "../src/doc/fib.js";
import { PdfRenderError, renderDocxToPdf } from "../src/pdf/render.js";

// Word 97-2003 .doc files desktop Word saved from the .docx beside each
// (scripts/render-eval's legacy eval), all under 40 KB.
const fixtures = join(fileURLToPath(new URL(".", import.meta.url)), "fixtures", "doc");
const bytes = (name: string): Uint8Array => new Uint8Array(readFileSync(join(fixtures, name)));

/** The package's text parts, by name. */
function parts(doc: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, data] of Object.entries(unzipSync(doc))) if (/\.(xml|rels)$/.test(name)) out[name] = strFromU8(data);
  return out;
}

/** The text of a document's body paragraphs, in order. */
function paragraphTexts(doc: DocxDocument): string[] {
  const texts: string[] = [];
  const walk = (blocks: typeof doc.sections[number]["blocks"]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        let text = "";
        for (const child of block.children) for (const run of child.type === "run" ? [child] : child.runs) for (const c of run.content) if (c.kind === "text") text += c.text;
        texts.push(text);
      } else {
        for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks);
      }
    }
  };
  for (const section of doc.sections) walk(section.blocks);
  return texts;
}

async function pages(source: Uint8Array): Promise<number> {
  const out = new PassThrough();
  out.resume();
  return (await renderDocxToPdf({ size: async () => source.length, read: async (a, b) => source.subarray(a, b) }, out)).pages;
}

describe("the .doc reader", () => {
  it("opens the streams and the FIB", () => {
    const streams = openDocStreams(bytes("parity-lists--v00.doc"));
    const fib = parseFib(streams.wordDocument);
    expect(fib.nFib).toBeGreaterThanOrEqual(0x00c1);
    expect(fib.ccpText).toBeGreaterThan(100);
    expect(isWordDocumentFile(bytes("parity-lists--v00.doc"))).toBe(true);
    expect(isWordDocumentFile(bytes("parity-lists--v00.docx"))).toBe(false);
    const reader = openDoc(bytes("parity-lists--v00.doc"));
    expect(reader.pieces.text(0, 40)).toContain("Lists parity");
    expect(reader.fonts.map((f) => f.name)).toContain("Calibri");
    expect(reader.styles.id(0)).toBe("Normal");
  });

  it("writes the text with its paragraphs, styles and runs", () => {
    const doc = DocxDocument.load(docToDocx(bytes("parity-lists--v00.doc")));
    const texts = paragraphTexts(doc);
    expect(texts[0]).toMatch(/^Lists parity/);
    expect(texts).toContain("First numbered item thistle ravine 37");
    const original = DocxDocument.load(bytes("parity-lists--v00.docx"));
    // The original's heading keeps its style through Word's .doc save.
    const first = doc.sections[0].blocks[0];
    const originalFirst = original.sections[0].blocks[0];
    expect(first.type === "paragraph" && first.props.styleId).toBe("Heading1");
    expect(originalFirst.type === "paragraph" && originalFirst.props.styleId).toBe("Heading1");
  });

  it("writes the lists as numbering.xml with their levels", () => {
    const xml = parts(docToDocx(bytes("parity-lists--v00.doc")));
    expect(xml["word/numbering.xml"]).toMatch(/<w:abstractNum w:abstractNumId="0">.*<w:lvl w:ilvl="0">.*<w:numFmt w:val="decimal"\/>.*<w:lvlText w:val="%1\."\/>/s);
    expect(xml["word/numbering.xml"]).toMatch(/<w:numFmt w:val="bullet"\/>/);
    expect(xml["word/numbering.xml"]).toMatch(/<w:num w:numId="1"><w:abstractNumId w:val="\d+"\/>/);
    expect(xml["word/document.xml"]).toMatch(/<w:numPr><w:ilvl w:val="1"\/><w:numId w:val="\d+"\/><\/w:numPr>/);
    const doc = DocxDocument.load(docToDocx(bytes("parity-lists--v00.doc")));
    const numbered = doc.sections[0].blocks.filter((b) => b.type === "paragraph" && b.props.numbering);
    expect(numbered.length).toBeGreaterThanOrEqual(6);
  });

  it("writes tables with their grid, header row, widths and shading", () => {
    const xml = parts(docToDocx(bytes("parity-tables--v00.doc")));
    const body = xml["word/document.xml"];
    // Two tables, as Word's own conversion gives: the .doc joined two that touched.
    expect(body.match(/<w:tbl>/g)?.length).toBe(2);
    expect(body).toMatch(/<w:tblGrid><w:gridCol w:w="\d+"\/><w:gridCol w:w="\d+"\/><w:gridCol w:w="\d+"\/><\/w:tblGrid>/);
    expect(body).toContain("<w:trPr><w:tblHeader/></w:trPr>");
    expect(body).toMatch(/<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"\/>/);
    expect(body).toMatch(/<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"\/>/);
    const doc = DocxDocument.load(docToDocx(bytes("parity-tables--v00.doc")));
    const tables = doc.sections[0].blocks.filter((b) => b.type === "table");
    expect(tables.length).toBe(2);
    expect(tables[0].type === "table" && tables[0].rows[0].cells.length).toBe(3);
    expect(paragraphTexts(doc)).toContain("Key");
  });

  it("writes the footnotes with their references", () => {
    const xml = parts(docToDocx(bytes("footnotes--v03.doc")));
    const notes = xml["word/footnotes.xml"].match(/<w:footnote w:id="\d+">/g) ?? [];
    expect(notes.length).toBe(DocxDocument.load(bytes("footnotes--v03.docx")).footnotes.size);
    const refs = xml["word/document.xml"].match(/<w:footnoteReference w:id="\d+"\/>/g) ?? [];
    expect(refs.length).toBe(notes.length);
    expect(xml["word/footnotes.xml"]).toContain("<w:footnoteRef/>");
    const doc = DocxDocument.load(docToDocx(bytes("footnotes--v03.doc")));
    expect(doc.footnotes.size).toBe(notes.length);
  });

  it("writes the sections with their headers and footers", () => {
    const xml = parts(docToDocx(bytes("parity-headerfooter--v07.doc")));
    expect(xml["word/document.xml"]).toMatch(/<w:headerReference w:type="default" r:id="rId\d+"\/>/);
    expect(xml["word/document.xml"]).toMatch(/<w:footerReference w:type="default" r:id="rId\d+"\/>/);
    const original = DocxDocument.load(bytes("parity-headerfooter--v07.docx"));
    const doc = DocxDocument.load(docToDocx(bytes("parity-headerfooter--v07.doc")));
    expect([doc.headers.size, doc.footers.size]).toEqual([original.headers.size, original.footers.size]);
    expect(xml["word/header1.xml"]).toContain("Header/Footer Parity");
    expect(xml["word/footer1.xml"]).toMatch(/<w:instrText xml:space="preserve"> ?PAGE/);
    expect(xml["word/settings.xml"]).toContain('w:name="compatibilityMode"');
    const columns = parts(docToDocx(bytes("parity-columns--v03.doc")))["word/document.xml"];
    expect(columns).toMatch(/<w:cols w:num="2"/);
    expect(columns.match(/<w:sectPr>/g)?.length).toBeGreaterThanOrEqual(2);
  });

  it("renders a .doc to as many pages as its .docx original, from a path or bytes", async () => {
    for (const name of ["parity-lists--v00", "parity-tables--v00", "parity-columns--v03"]) {
      expect(await pages(bytes(`${name}.doc`))).toBe(await pages(bytes(`${name}.docx`)));
    }
    const out = new PassThrough();
    out.resume();
    expect((await renderDocxToPdf(join(fixtures, "parity-lists--v00.doc"), out)).pages).toBe(1);
  });

  it("loads a .doc as a DocxDocument", () => {
    const doc = DocxDocument.load(bytes("parity-columns--v03.doc"));
    expect(paragraphTexts(doc).join(" ")).toContain("One-column intro paragraph before the continuous break.");
    expect(Object.keys(docToDocxFiles(bytes("parity-columns--v03.doc")))).toContain("word/styles.xml");
  });

  it("writes legacy form fields with their data, as Word converts them", () => {
    const body = parts(docToDocx(bytes("probe2-form-checkboxes--v02.doc")))["word/document.xml"];
    expect(body).toContain('<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Name"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar>');
    expect(body).toContain('<w:ddList><w:result w:val="1"/><w:listEntry w:val="Sales"/><w:listEntry w:val="Engineering"/><w:listEntry w:val="Finance"/></w:ddList>');
    expect(body).toContain('<w:name w:val="Sub"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="1"/><w:checked/></w:checkBox>');
    expect(body).toContain('<w:checkBox><w:sizeAuto/><w:default w:val="0"/><w:checked w:val="0"/></w:checkBox>');
  });

  it("writes an EQ ruby field as a w:ruby, keeping the comment its code holds", () => {
    const body = parts(docToDocx(bytes("probe2-ruby-vertical--v00.doc")))["word/document.xml"];
    expect(body.match(/<w:ruby>/g)?.length).toBe(6);
    expect(body).toMatch(/<w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"\/><w:hps w:val="8"\/><w:hpsRaise w:val="10"\/><\/w:rubyPr><w:rt><w:r>.*?<w:t xml:space="preserve">かんじ<\/w:t><\/w:r><\/w:rt><w:rubyBase><w:r>.*?<w:t xml:space="preserve">漢字<\/w:t><\/w:r><\/w:rubyBase><\/w:ruby>/);
    expect(body).not.toContain("EQ \\* jc2");
    // Ten comment references, as in Word's conversion: one sits inside a ruby field's code.
    expect(body.match(/<w:commentReference /g)?.length).toBe(10);
    // The section's vertical text flow and its grid.
    expect(body).toContain('<w:textDirection w:val="tbRl"/>');
    expect(body).toContain('<w:docGrid w:type="lines" w:linePitch="360"/>');
  });

  it("refuses what is not a Word 97-2003 document, naming it", async () => {
    expect(() => docToDocx(new TextEncoder().encode("{\\rtf1 hello}"))).toThrow(/unsupported format: not an OLE compound file/);
    // A compound file with another stream (an Excel 97-2003 workbook), and a password-protected .docx.
    const ole = (name: string): Uint8Array => {
      const cfb = CFB.utils.cfb_new();
      CFB.utils.cfb_add(cfb, name, new Uint8Array(64));
      return new Uint8Array(CFB.write(cfb, { type: "array" }) as number[]);
    };
    expect(() => docToDocx(ole("Workbook"))).toThrow(/unsupported format: an Excel 97-2003 workbook/);
    expect(() => docToDocx(ole("PowerPoint Document"))).toThrow(/unsupported format: a PowerPoint 97-2003 presentation/);
    expect(() => docToDocx(ole("Contents"))).toThrow(/unsupported format: an OLE compound file with no Word document/);
    expect(() => docToDocx(ole("EncryptedPackage"))).toThrow(/unsupported format: a password-protected document/);
    // An encrypted Word 95 file: its FIB starts with 0xA5DC and nFib 0x68.
    const word6 = bytes("parity-lists--v00.doc").slice();
    const streams = openDocStreams(word6);
    const at = findStream(word6, streams.wordDocument);
    expect(at).toBeGreaterThan(0);
    word6[at] = 0xdc;
    word6[at + 2] = 0x68;
    word6[at + 3] = 0;
    word6[at + 0x0b] |= 0x01;
    expect(() => docToDocx(word6)).toThrow(/unsupported format: a password-protected document \(Word 6\.0\/95 encryption\)/);
    // An encrypted one: fEncrypted in the FIB's flags.
    const encrypted = bytes("parity-lists--v00.doc").slice();
    encrypted[at + 0x0b] |= 0x01;
    expect(() => docToDocx(encrypted)).toThrow(/unsupported format: a password-protected document/);
    const out = new PassThrough();
    out.resume();
    const error = await renderDocxToPdf({ size: async () => encrypted.length, read: async (a, b) => encrypted.subarray(a, b) }, out).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(PdfRenderError);
    expect((error as PdfRenderError).code).toBe("unsupported_format");
  });
});

// Word 6.0/95 files from GovDocs1 (public domain US government documents):
// 000114 a full-saved Word 95 file with a table, a header and footers;
// 025814 a fast-saved one (its text in a piece table); 004072 one with
// tracked changes; 021219 one with bulleted lists; 022984 one with drawn
// lines; 043221 one with a comment.
describe("the .doc reader with a Word 6.0/95 file", () => {
  it("reads the FIB and the 8-bit text of a full-saved file", () => {
    const reader = openDoc(bytes("word95-000114.doc"));
    expect(reader.fib.nFib).toBe(0x68);
    expect(reader.fib.word6?.complex).toBe(false);
    expect(reader.streams.table).toBe(reader.streams.wordDocument);
    expect(reader.fonts.map((f) => f.name)).toContain("Times New Roman");
    expect(reader.styles.id(0)).toBe("Normal");
    const doc = DocxDocument.load(docToDocx(bytes("word95-000114.doc")));
    const texts = paragraphTexts(doc);
    expect(texts[0]).toBe("FIELD PROTOCOLS FOR NRDABiota/Tarball Stranding Documentation");
    expect(texts).toContain("Objectives");
    expect(doc.sections[0].blocks.some((block) => block.type === "table")).toBe(true);
    // The header the section's grpfIhdt names.
    expect(Object.values(parts(docToDocx(bytes("word95-000114.doc")))).some((xml) => xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr') && xml.includes(">12/98<"))).toBe(true);
  });

  it("reads a fast-saved file's pieces as 8-bit text, in order", () => {
    const reader = openDoc(bytes("word95-025814.doc"));
    expect(reader.fib.word6?.complex).toBe(true);
    const texts = paragraphTexts(DocxDocument.load(docToDocx(bytes("word95-025814.doc"))));
    expect(texts).toContain("STS-84 Crew Debrief Captions");
    expect(texts).toContain("Center of Tropical Cyclone #01B in the Bay of Bengal.  It later reached landfall near the Burma/ Bangladesh boarder.");
    // Text pasted from a Macintosh (sprmCChse) keeps its Macintosh quotes.
    expect(texts.join("\n")).toContain("the basin’s responses to past climate change");
    // Read as 16-bit, the 8-bit text would turn into CJK characters.
    expect(texts.join("")).not.toMatch(/[　-￿]/);
  });

  it("writes the tracked changes with their author and date", () => {
    const body = parts(docToDocx(bytes("word95-004072.doc")))["word/document.xml"];
    expect(body).toContain('<w:ins w:id="9" w:author="CAE Department" w:date="1994-07-15T16:02:00Z">');
    expect(body).toMatch(/<w:del w:id="\d+" w:author="Unknown">/);
  });

  it("writes the autonumbered paragraphs as lists", () => {
    const xml = parts(docToDocx(bytes("word95-021219.doc")));
    expect(xml["word/numbering.xml"]).toMatch(/<w:numFmt w:val="bullet"\/><w:lvlText w:val=""\/>.*<w:rFonts w:ascii="Symbol"/);
    expect(xml["word/document.xml"]).toContain('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
  });

  it("writes the drawing objects as anchored shapes", () => {
    // 022984: horizontal rules drawn as lines, 1/4 inch under their paragraphs, 1.75 points wide, gray.
    const body = parts(docToDocx(bytes("word95-022984.doc")))["word/document.xml"];
    expect(body.match(/<a:prstGeom prst="line">/g)?.length).toBeGreaterThanOrEqual(6);
    expect(body).toContain('<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>152400</wp:posOffset></wp:positionV>');
    expect(body).toContain('<a:ln w="22225"><a:solidFill><a:srgbClr val="D4D4D4"></a:srgbClr></a:solidFill>');
  });

  it("writes the comments with their 8-bit authors and initials", () => {
    const xml = parts(docToDocx(bytes("word95-043221.doc")));
    expect(xml["word/comments.xml"]).toContain('<w:comment w:id="0" w:author="kevin l mitchell" w:initials="klm">');
    expect(xml["word/document.xml"]).toContain('<w:commentReference w:id="0"/>');
  });

  it("renders each file", async () => {
    for (const name of ["word95-000114.doc", "word95-025814.doc", "word95-004072.doc", "word95-021219.doc", "word95-022984.doc", "word95-043221.doc"]) {
      expect(await pages(bytes(name)), name).toBeGreaterThan(0);
    }
  });
});

/** Where the WordDocument stream's bytes start in the file (the FIB's wIdent). */
function findStream(file: Uint8Array, stream: Uint8Array): number {
  const head = stream.subarray(0, 32);
  outer: for (let i = 512; i + 32 <= file.length; i += 64) {
    for (let k = 0; k < 32; k++) if (file[i + k] !== head[k]) continue outer;
    return i;
  }
  return -1;
}
