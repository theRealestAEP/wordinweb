import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, acceptAllRevisions, collectRevisions, rejectAllRevisions } from "@wordinweb/core";
import { comparedName, compareWithFile } from "../src/compare.js";

/**
 * The host side of Compare Documents: the browser plumbing between the file
 * picker and the engine call. The alignment and the revision output are core's
 * and are tested there (core/test/compare-corpus, core/test/compare-roundtrip);
 * what matters here is that the wiring hands the right document to the right
 * side of the comparison — swapping original and revised silently inverts every
 * insertion and deletion, and nothing downstream would notice.
 */

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function docxBytes(text: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/document.xml": strToU8(
      `<?xml version="1.0"?><w:document ${W}><w:body>` +
        `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`,
    ),
  });
}

function textOf(doc: DocxDocument): string {
  return doc.sections
    .flatMap((s) => s.blocks)
    .map((b) => (b.type === "paragraph" ? b.children : []))
    .flat()
    .map((child) => (child.type === "run" ? child.content : []))
    .flat()
    .map((c) => (c.kind === "text" ? c.text : ""))
    .join("");
}

describe("compareWithFile", () => {
  it("treats the open document as the original and the picked file as the revised one", async () => {
    const file = new File([docxBytes("the revised wording") as unknown as BlobPart], "revised.docx");
    const result = await compareWithFile(docxBytes("the original wording"), file, { author: "Ada", date: "2026-08-12T00:00:00Z" });

    const compared = DocxDocument.load(result.bytes);
    expect(collectRevisions(compared).map((r) => r.author)).toEqual(["Ada", "Ada"]);

    // Accepting gives the PICKED file; rejecting gives what was on screen.
    const accepted = DocxDocument.load(result.bytes);
    acceptAllRevisions(accepted);
    expect(textOf(accepted)).toBe("the revised wording");

    const rejected = DocxDocument.load(result.bytes);
    rejectAllRevisions(rejected);
    expect(textOf(rejected)).toBe("the original wording");

    expect(result.revisedName).toBe("revised.docx");
    expect(result.notes).toEqual([]);
  });

  it("names the result after the file it was compared against", () => {
    expect(comparedName("report.docx")).toBe("report (compared).docx");
    expect(comparedName("no extension")).toBe("no extension (compared).docx");
  });
});
