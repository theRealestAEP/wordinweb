import { readFileSync } from "node:fs";
import { join } from "node:path";
import { PassThrough } from "node:stream";
import { fileURLToPath } from "node:url";
import { strFromU8 } from "fflate";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { isOldWordFile, oldWordPackageFiles } from "../src/oldword/hooks.js";
import { PdfRenderError, renderDocxToPdf } from "../src/pdf/render.js";

// Documents in the oldest Word formats from GovDocs1 (US government works,
// public domain), each under 100 KB: two Word for the Macintosh 5.0 files
// (the second fast-saved, its text in pieces) and Word for Windows 2.0 and
// 1.x files.
const fixtures = join(fileURLToPath(new URL(".", import.meta.url)), "fixtures", "oldword");
const bytes = (name: string): Uint8Array => new Uint8Array(readFileSync(join(fixtures, name)));

/** The text of a document's body paragraphs, tables' too, in order. */
function paragraphTexts(doc: DocxDocument): string[] {
  const texts: string[] = [];
  const walk = (blocks: (typeof doc.sections)[number]["blocks"]): void => {
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

async function render(source: Uint8Array): Promise<number> {
  const out = new PassThrough();
  out.resume();
  return (await renderDocxToPdf({ size: async () => source.length, read: async (a, b) => source.subarray(a, b) }, out)).pages;
}

describe("the oldest Word formats", () => {
  it("tells them from their first bytes", () => {
    for (const name of ["macword5-govdocs-000271.doc", "winword2-govdocs-474644.doc", "winword1-govdocs-307072.doc"]) expect(isOldWordFile(bytes(name).subarray(0, 1024))).toBe(true);
    // A Mac Word file of a version this does not read, a Word for Windows 1.0 beta, a .docx.
    expect(isOldWordFile(Uint8Array.of(0xfe, 0x37, 0x00, 0x1e))).toBe(false);
    expect(isOldWordFile(Uint8Array.of(0x9b, 0xa5, 0x1c, 0x00))).toBe(false);
    expect(isOldWordFile(Uint8Array.of(0x50, 0x4b, 0x03, 0x04))).toBe(false);
  });

  it("reads a Word for the Macintosh 5.0 document: Mac OS Roman text, its styles, its PICT", () => {
    const files = oldWordPackageFiles(bytes("macword5-govdocs-000271.doc"));
    const document = strFromU8(files["word/document.xml"]);
    // The title's em dash is Mac OS Roman 0xD1.
    expect(document).toContain("Giant porphyry-related metal camps of the world — a database");
    expect(strFromU8(files["word/styles.xml"])).toContain('<w:name w:val="SECTION_HDR"/>');
    const pict = files["word/media/image1.pict"];
    expect(pict).toBeDefined();
    // A version 2 PICT: after its size and frame, the version opcode.
    expect([...pict.subarray(10, 14)]).toEqual([0x00, 0x11, 0x02, 0xff]);
  });

  it("reads a fast-saved Word for the Macintosh 5.0 document's pieces in order", () => {
    const texts = paragraphTexts(DocxDocument.load(bytes("macword5-govdocs-000790.doc")));
    const title = ["GEOLOGIC MAP ", "OF THE", "NEAR SIDE ", "OF THE MOON "].map((line) => texts.indexOf(line));
    expect(title.every((at, i) => at >= 0 && (i === 0 || at > title[i - 1]))).toBe(true);
    expect(texts).toContain("Prepared for the National Aeronautics and Space Administration");
  });

  it("reads Word for Windows 2.0 and 1.x documents: Windows ANSI text", () => {
    const word2 = paragraphTexts(DocxDocument.load(bytes("winword2-govdocs-474644.doc"))).join("\n");
    expect(word2).toContain("Haike Hölzen");
    expect(word2).toContain("Bettina Löscher");
    const word1 = paragraphTexts(DocxDocument.load(bytes("winword1-govdocs-307072.doc"))).join("\n");
    expect(word1).toContain("Sampling and laboratory procedures and mathematical equations for dust and climate data");
  });

  it("renders each, and calls a cut-short file damaged", async () => {
    for (const name of ["macword5-govdocs-000271.doc", "winword2-govdocs-474644.doc", "winword1-govdocs-307072.doc"]) expect(await render(bytes(name))).toBeGreaterThan(0);
    const cut = bytes("macword5-govdocs-000271.doc").slice(0, 4096);
    const error = await render(cut).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(PdfRenderError);
    expect((error as PdfRenderError).code).toBe("damaged");
  });
});
