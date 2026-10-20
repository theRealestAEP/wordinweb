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

  it("refuses what is not a Word 97-2003 document, naming it", async () => {
    expect(() => docToDocx(new TextEncoder().encode("{\\rtf1 hello}"))).toThrow(/unsupported format: not an OLE compound file/);
    // A compound file with another stream (an Excel 97-2003 workbook), and a password-protected .docx.
    const ole = (name: string): Uint8Array => {
      const cfb = CFB.utils.cfb_new();
      CFB.utils.cfb_add(cfb, name, new Uint8Array(64));
      return new Uint8Array(CFB.write(cfb, { type: "array" }) as number[]);
    };
    expect(() => docToDocx(ole("Workbook"))).toThrow(/unsupported format: an OLE compound file with no Word document/);
    expect(() => docToDocx(ole("EncryptedPackage"))).toThrow(/unsupported format: a password-protected document/);
    // A Word 6.0/95 file: its FIB starts with 0xA5DC.
    const word6 = bytes("parity-lists--v00.doc").slice();
    const streams = openDocStreams(word6);
    const at = findStream(word6, streams.wordDocument);
    expect(at).toBeGreaterThan(0);
    word6[at] = 0xdc;
    expect(() => docToDocx(word6)).toThrow(/unsupported format: a Word 6\.0\/95 document/);
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

/** Where the WordDocument stream's bytes start in the file (the FIB's wIdent). */
function findStream(file: Uint8Array, stream: Uint8Array): number {
  const head = stream.subarray(0, 32);
  outer: for (let i = 512; i + 32 <= file.length; i += 64) {
    for (let k = 0; k < 32; k++) if (file[i + k] !== head[k]) continue outer;
    return i;
  }
  return -1;
}
