import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument, layoutDocumentStream } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { LaidOutPage } from "../src/layout/types.js";
import { openStreamingDocument } from "../src/stream/body.js";
import { bytesSource } from "../src/stream/source.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

/** The pages of a whole layout, as JSON a streamed layout must reproduce. */
function wholePages(bytes: Uint8Array): string[] {
  const doc = DocxDocument.load(bytes);
  return layoutDocument(doc, { measurer: new ApproxMeasurer() }).pages.map(pageJson);
}

/** The pages of the streamed layout: a counting pass, then a rendering pass
 * with its page total and bookmark pages, as openPrintStream runs them. */
async function streamedPages(bytes: Uint8Array): Promise<{ pages: string[]; counted: number }> {
  const opened = await openStreamingDocument(bytesSource(bytes));
  const measurer = new ApproxMeasurer();
  let final = await layoutDocumentStream(opened.doc, { measurer, prescan: opened.prescan, onPage: () => {} });
  let knownTotalPages: number | undefined;
  if (opened.bodyHasPageTotal) {
    for (let pass = 0; pass < 2; pass++) {
      knownTotalPages = final.totalPages;
      final = await layoutDocumentStream(opened.doc, { measurer, prescan: opened.prescan, knownTotalPages, onPage: () => {} });
      if (final.totalPages === knownTotalPages) break;
    }
  }
  const pages: string[] = [];
  const rendered = await layoutDocumentStream(opened.doc, {
    measurer,
    prescan: opened.prescan,
    knownTotalPages,
    final,
    onPage: (page) => {
      pages.push(pageJson(page));
    },
  });
  expect(rendered.totalPages).toBe(final.totalPages);
  return { pages, counted: final.totalPages };
}

function pageJson(page: LaidOutPage): string {
  return JSON.stringify(page, (key, value: unknown) => (key === "src" || key === "srcParent" ? undefined : value));
}

async function expectSame(bytes: Uint8Array): Promise<void> {
  const whole = wholePages(bytes);
  const streamed = await streamedPages(bytes);
  expect(streamed.counted).toBe(whole.length);
  expect(streamed.pages.length).toBe(whole.length);
  for (let i = 0; i < whole.length; i++) expect(streamed.pages[i]).toBe(whole[i]);
}

const pageBreak = `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
const prose = (n: number, text: string) => Array.from({ length: n }, (_, i) => p(`${text} ${i + 1}. ${"Words that fill a line and wrap onto the next. ".repeat(4)}`)).join("");

describe("streamed layout", () => {
  it("reproduces a whole layout, page for page", async () => {
    await expectSame(makeDocx({ "word/document.xml": wrapDocument(prose(120, "Paragraph")) }));
  });

  it("keeps sections, content controls and deleted paragraphs as parseBody does", async () => {
    const landscape = `<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>`;
    const body =
      prose(30, "First") +
      `<w:p><w:pPr><w:sectPr>${landscape}</w:sectPr></w:pPr><w:r><w:t>Closes the first section</w:t></w:r></w:p>` +
      `<w:sdt><w:sdtPr><w:alias w:val="group"/></w:sdtPr><w:sdtContent>${prose(10, "Inside a control")}<w:sdt><w:sdtContent>${p("Nested control")}</w:sdtContent></w:sdt></w:sdtContent></w:sdt>` +
      `<w:p><w:pPr><w:rPr><w:del w:id="1" w:author="a" w:date="2024-01-01T00:00:00Z"/></w:rPr></w:pPr><w:del w:id="2" w:author="a" w:date="2024-01-01T00:00:00Z"><w:r><w:delText>Deleted whole</w:delText></w:r></w:del></w:p>` +
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p("In a cell")}</w:tc></w:tr></w:tbl>` +
      prose(40, "Last") +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
    await expectSame(makeDocx({ "word/document.xml": wrapDocument(body) }));
  });

  it("shows NUMPAGES and PAGEREF from the counting pass", async () => {
    const field = (instr: string, cached: string) =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${cached}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const body =
      `<w:p><w:r><w:t xml:space="preserve">Pages: </w:t></w:r>${field("NUMPAGES", "1")}<w:r><w:t xml:space="preserve">; the end is on page </w:t></w:r>${field("PAGEREF end \\h", "1")}</w:p>` +
      prose(50, "Body") +
      pageBreak +
      `<w:p><w:bookmarkStart w:id="0" w:name="end"/><w:r><w:t>The end</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
    const bytes = makeDocx({ "word/document.xml": wrapDocument(body) });
    await expectSame(bytes);
    const streamed = await streamedPages(bytes);
    const texts = (JSON.parse(streamed.pages[0]) as LaidOutPage).items.flatMap((item) => (item.kind === "text" ? [item.text] : []));
    expect(texts).toContain(String(streamed.counted));
  });

  it("numbers footnotes and SEQ fields in document order", async () => {
    const seq = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Figure \\* ARABIC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>9</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const note = (id: number) => `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
    const body =
      prose(20, "Before") +
      `<w:p><w:r><w:t xml:space="preserve">Figure </w:t></w:r>${seq}${note(1)}</w:p>` +
      prose(40, "Between") +
      `<w:p><w:r><w:t xml:space="preserve">Figure </w:t></w:r>${seq}${note(2)}</w:p>` +
      prose(20, "After") +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
    const footnotes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> First note</w:t></w:r></w:p></w:footnote>
<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Second note</w:t></w:r></w:p></w:footnote>
</w:footnotes>`;
    const rels = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>`;
    await expectSame(makeDocx({ "word/document.xml": wrapDocument(body), "word/footnotes.xml": footnotes, "word/_rels/document.xml.rels": rels }));
  });
});
