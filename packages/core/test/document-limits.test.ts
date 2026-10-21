import { Writable } from "node:stream";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { PdfRenderError } from "../src/pdf/errors.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { bytesSource } from "../src/stream/source.js";
import { countElements } from "../src/xml.js";
import { DOCUMENT_LIMITS } from "../src/zip.js";
import { W_NS, makeDocx, wrapDocument } from "./helpers.js";

// Each document is just over one of DOCUMENT_LIMITS. The render refuses it
// with too_large, and DocxDocument.load with a "too large:" error, from
// counts made before the XML is parsed. The probes that measured each limit
// (at it and over it, in a 1024 MB heap) are in the commit that set them.

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const R_NS = `xmlns:r="${REL}"`;
const rels = (more: string) => `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${more}</Relationships>`;
const run = "<w:r><w:t>a</w:t></w:r>";

function withHeader(header: string): Uint8Array {
  return makeDocx({
    "word/document.xml": wrapDocument(`<w:p><w:r><w:t>Body.</w:t></w:r></w:p><w:sectPr><w:headerReference ${R_NS} w:type="default" r:id="rIdH"/></w:sectPr>`),
    "word/_rels/document.xml.rels": rels(`<Relationship Id="rIdH" Type="${REL}/header" Target="header1.xml"/>`),
    "word/header1.xml": `<?xml version="1.0"?><w:hdr ${W_NS}>${header}</w:hdr>`,
  });
}

async function renderError(docx: Uint8Array): Promise<PdfRenderError> {
  try {
    await renderDocxToPdf(bytesSource(docx), new Writable({ write: (_chunk, _encoding, done) => done() }));
  } catch (error) {
    return error as PdfRenderError;
  }
  throw new Error("rendered without an error");
}

const over = (limit: number) => limit + 10;

/** [document, what the render says, whether DocxDocument.load refuses it too]. */
const cases: Record<string, [() => Uint8Array, string, boolean]> = {
  "the XML parts other than the document, in bytes": [
    () => makeDocx({ "word/document.xml": wrapDocument("<w:p/>"), "word/styles.xml": `<w:styles ${W_NS}>${" ".repeat(DOCUMENT_LIMITS.heldBytes)}</w:styles>` }),
    "the XML parts other than word/document.xml come to",
    true,
  ],
  "the XML parts other than the document, in elements": [
    () => makeDocx({ "word/document.xml": wrapDocument("<w:p/>"), "customXml/item1.xml": `<root>${"<a/>".repeat(over(DOCUMENT_LIMITS.heldElements))}</root>` }),
    "the XML parts other than word/document.xml hold",
    true,
  ],
  "the XML parts other than the document, in paragraphs": [() => withHeader("<w:p/>".repeat(over(DOCUMENT_LIMITS.heldBlocks))), "paragraphs and table cells, over the 10000", true],
  "the document, in elements": [
    () => makeDocx({ "word/document.xml": wrapDocument(`<w:p>${run.repeat(99_990)}</w:p>`.repeat(5) + "<w:p/>".repeat(100)) }),
    "elements",
    true,
  ],
  "the document, in paragraphs": [() => makeDocx({ "word/document.xml": wrapDocument("<w:p/>".repeat(over(DOCUMENT_LIMITS.documentBlocks))) }), "paragraphs and table cells", true],
  "a block, in elements": [() => makeDocx({ "word/document.xml": wrapDocument(`<w:p>${run.repeat(over(DOCUMENT_LIMITS.blockElements / 2))}</w:p>`) }), "the body holds a block of over 250000 elements", false],
  "a block, in text": [
    () => makeDocx({ "word/document.xml": wrapDocument(`<w:p><w:r><w:t>${"word ".repeat(over(DOCUMENT_LIMITS.blockText / 5))}</w:t></w:r></w:p>`) }),
    "the body holds a block of over 1048576 bytes of text",
    false,
  ],
  "a block, in bytes": [
    () => makeDocx({ "word/document.xml": wrapDocument(`<w:p>${`<w:r><w:rPr><w:rFonts w:ascii="${"x".repeat(200)}"/></w:rPr>${run.slice(5)}`.repeat(40_000)}</w:p>`) }),
    "the body holds a block over 8388608 bytes",
    false,
  ],
  "a table, in rows": [
    () => makeDocx({ "word/document.xml": wrapDocument(`<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>${"<w:tr><w:tc/></w:tr>".repeat(over(DOCUMENT_LIMITS.tableRows))}</w:tbl>`) }),
    "the body holds a table of over 20000 rows",
    false,
  ],
};

describe("DOCUMENT_LIMITS", () => {
  it("counts a part's elements, and its paragraphs and table cells, from its bytes", () => {
    const xml = `<?xml version="1.0"?><!-- <x/> --><w:body ${W_NS}><w:p><w:r><w:t>a &lt;b/&gt;</w:t></w:r></w:p><p/><x:tc xmlns:x="x"/><w:tbl/><w:pPr/></w:body>`;
    expect(countElements(new TextEncoder().encode(xml))).toEqual({ elements: 9, blocks: 3 });
  });

  for (const [what, [make, message, load]] of Object.entries(cases)) {
    it(`refuses a document over the limit for ${what}`, async () => {
      const docx = make();
      const error = await renderError(docx);
      expect(error).toBeInstanceOf(PdfRenderError);
      expect(error.code).toBe("too_large");
      expect(error.message).toContain(message);
      if (load) expect(() => DocxDocument.load(docx)).toThrow(/^too large: /);
      else expect(() => DocxDocument.load(docx)).not.toThrow();
    }, 60_000);
  }
});
