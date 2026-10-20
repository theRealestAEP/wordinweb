import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import type { LaidOutPage } from "../src/layout/types.js";
import { makeDocxWithMedia } from "./helpers.js";

/**
 * A paragraph that ends with a page break arms the paragraph after it to open
 * at the top of the new page. When a table came next, the table opened there,
 * but the flag stayed armed. The first paragraph after the table then went
 * back to the page top and was drawn over the table's first row (a Word 95
 * sample garbled pages 2 to 4 this way; repro r1.docx).
 *
 * Each fixture below puts a break of some kind next to a table, a picture or
 * a text box, and checks that every word is on a page, in reading order, and
 * that no two pieces of content are drawn over each other.
 */

const measurer = new ApproxMeasurer();

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const PGMAR =
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>';
const SECT = `<w:sectPr>${PGMAR}</w:sectPr>`;

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const PAGE_BREAK_ONLY = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
const LONG =
  "one two three four five six seven eight nine ten eleven twelve thirteen fourteen fifteen " +
  "sixteen seventeen eighteen nineteen twenty twentyone twentytwo twentythree twentyfour twentyfive";

/** A two-cell table, as repro r1.docx has it: "A" and a five-line text. */
const table = (a = "A", b = LONG) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>' +
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${a.startsWith("<") ? a : p(a)}</w:tc>` +
  `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${p(b)}</w:tc>` +
  "</w:tr></w:tbl>";

/** A one-cell table whose cell holds the given block XML. */
const oneCell = (cellXml: string) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr>' +
  `<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr>${cellXml}</w:tc>` +
  "</w:tr></w:tbl>";

const picture =
  '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImg"/></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></pic:spPr>' +
  "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

/** A paragraph that anchors a 1in text box, wrapped top and bottom, holding "BOX". */
const textBox =
  '<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"' +
  ' relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1828800" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
  '<wp:wrapTopAndBottom/><wp:docPr id="2" name="Text Box 2"/><wp:cNvGraphicFramePr/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>' +
  '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  `<wps:txbx><w:txbxContent>${p("BOX")}</w:txbxContent></wps:txbx>` +
  '<wps:bodyPr rot="0" anchor="t"><a:noAutofit/></wps:bodyPr>' +
  "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>";

/** A paragraph that anchors a 1in text box at a fixed place on its page, holding "PAGEBOX". */
const pageBox = (topPx: number) =>
  '<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"' +
  ' relativeHeight="3" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="page"><wp:posOffset>1143000</wp:posOffset></wp:positionH>' +
  `<wp:positionV relativeFrom="page"><wp:posOffset>${topPx * 9525}</wp:posOffset></wp:positionV>` +
  '<wp:extent cx="1828800" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
  '<wp:wrapTopAndBottom/><wp:docPr id="3" name="Text Box 3"/><wp:cNvGraphicFramePr/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>' +
  '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  `<wps:txbx><w:txbxContent>${p("PAGEBOX")}</w:txbxContent></wps:txbx>` +
  '<wps:bodyPr rot="0" anchor="t"><a:noAutofit/></wps:bodyPr>' +
  "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>";

/** A run holding a red 1in-wide anchored rectangle, as validator doc97-check-d made them. */
const redBox = (wrap: "wrapNone" | "wrapTopAndBottom", offsetEmu: number, heightEmu: number) =>
  '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"' +
  ' relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${offsetEmu}</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="914400" cy="${heightEmu}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
  `<wp:${wrap}/><wp:docPr id="4" name="R"/><wp:cNvGraphicFramePr/>` +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp><wps:cNvSpPr/><wps:spPr>' +
  `<a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="${heightEmu}"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
  "</wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>";

/** The pages (1-based) that draw a red shape. */
const redPages = (pages: LaidOutPage[]) =>
  pages.flatMap((page, n) =>
    page.items.some((it) => it.kind === "path" && it.fill?.toLowerCase() === "#ff0000") ? [n + 1] : [],
  );

function layout(body: string, parts: Record<string, string> = {}): LaidOutPage[] {
  const doc = DocxDocument.load(
    makeDocxWithMedia(
      {
        "word/document.xml":
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}>` +
          `<w:body>${body}${SECT}</w:body></w:document>`,
        "word/_rels/document.xml.rels":
          '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
          "</Relationships>",
        ...parts,
      },
      { "word/media/image1.png": PNG },
    ),
  );
  return layoutDocument(doc, { measurer }).pages;
}

interface Box {
  label: string;
  page: number;
  x0: number;
  x1: number;
  y0: number;
  y1: number;
}

/** Every word, picture and table drawn, as boxes in reading order. */
function boxesOf(pages: LaidOutPage[]): Box[] {
  const out: Box[] = [];
  pages.forEach((page, n) => {
    for (const it of page.items) {
      if (it.kind === "text" && it.text.trim() !== "") {
        out.push({ label: it.text, page: n, x0: it.x, x1: it.x + it.width, y0: it.lineTop, y1: it.lineTop + it.lineHeight });
      } else if (it.kind === "image") {
        out.push({ label: "[picture]", page: n, x0: it.x, x1: it.x + it.width, y0: it.y, y1: it.y + it.height });
      }
    }
  });
  return out;
}

/** The tables' outer boxes on each page, from their move grips. */
function tablesOf(pages: LaidOutPage[]): Box[] {
  const out: Box[] = [];
  pages.forEach((page, n) => {
    for (const it of page.items) {
      if (it.kind === "grip" && it.axis === "move") {
        out.push({ label: "[table]", page: n, x0: it.x, x1: it.x2 ?? it.x, y0: it.y1, y1: it.y2 });
      }
    }
  });
  return out;
}

const EPS = 0.5;
const meets = (a: Box, b: Box) =>
  a.page === b.page && a.x0 < b.x1 - EPS && b.x0 < a.x1 - EPS && a.y0 < b.y1 - EPS && b.y0 < a.y1 - EPS;

/** Pairs of drawn boxes that cover each other. */
function overlaps(boxes: Box[]): string[] {
  const out: string[] = [];
  for (let i = 0; i < boxes.length; i++) {
    for (let j = i + 1; j < boxes.length; j++) {
      if (meets(boxes[i], boxes[j])) out.push(`${boxes[i].label} / ${boxes[j].label} on page ${boxes[i].page + 1}`);
    }
  }
  return out;
}

/** Body paragraphs (named by their text) that are drawn inside a table's box. */
function bodyTextInTables(pages: LaidOutPage[], bodyWords: string[]): string[] {
  const tables = tablesOf(pages);
  return boxesOf(pages)
    .filter((b) => bodyWords.includes(b.label) && tables.some((t) => meets(b, t)))
    .map((b) => `${b.label} on page ${b.page + 1}`);
}

/** The page (1-based) of each listed word, in the order given. */
const pagesOfWords = (pages: LaidOutPage[], words: string[]) =>
  words.map((w) => {
    const hit = boxesOf(pages).find((b) => b.label === w);
    return hit ? hit.page + 1 : 0;
  });

/** True when the words appear in this order in the drawn reading order (page, then y). */
function inOrder(pages: LaidOutPage[], words: string[]): boolean {
  const boxes = boxesOf(pages);
  const at = words.map((w) => boxes.find((b) => b.label === w));
  if (at.some((b) => !b)) return false;
  for (let i = 1; i < at.length; i++) {
    const a = at[i - 1]!;
    const b = at[i]!;
    if (b.page < a.page || (b.page === a.page && b.y0 < a.y0 - EPS)) return false;
  }
  return true;
}

function expectClean(pages: LaidOutPage[], order: string[], body: string[]) {
  expect(pagesOfWords(pages, order).every((n) => n > 0)).toBe(true);
  expect(inOrder(pages, order)).toBe(true);
  expect(overlaps(boxesOf(pages))).toEqual([]);
  expect(bodyTextInTables(pages, body)).toEqual([]);
}

describe("a break next to a table, a picture or a text box", () => {
  it("lays out repro r1: a page-break-only paragraph and then a table", () => {
    const pages = layout(p("Before") + PAGE_BREAK_ONLY + table() + p("After"));
    expect(pagesOfWords(pages, ["Before", "A", "one", "twentyfive", "After"])).toEqual([1, 2, 2, 2, 2]);
    expectClean(pages, ["Before", "A", "twentyfive", "After"], ["Before", "After"]);
    // The paragraph after the table starts below the table's bottom edge.
    const [tbl] = tablesOf(pages);
    const after = boxesOf(pages).find((b) => b.label === "After")!;
    expect(after.y0).toBeGreaterThanOrEqual(tbl.y1 - EPS);
  });

  it("lays out the same document with w:pageBreakBefore in place of the break run", () => {
    const pages = layout(
      p("Before") + '<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>' + table() + p("After"),
    );
    expectClean(pages, ["Before", "A", "twentyfive", "After"], ["Before", "After"]);
  });

  it("lays out a paragraph with text that ends in a page break, then a table", () => {
    const pages = layout(
      '<w:p><w:r><w:t>Before</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>' + table() + p("After"),
    );
    expect(pagesOfWords(pages, ["Before", "A", "After"])).toEqual([1, 2, 2]);
    expectClean(pages, ["Before", "A", "twentyfive", "After"], ["Before", "After"]);
  });

  it("lays out two tables after a page-break-only paragraph", () => {
    const pages = layout(p("Before") + PAGE_BREAK_ONLY + table("A") + p("Mid") + table("B") + p("After"));
    expectClean(pages, ["Before", "A", "Mid", "B", "After"], ["Before", "Mid", "After"]);
  });

  it("lays out a table directly after another table after a page break", () => {
    const pages = layout(p("Before") + PAGE_BREAK_ONLY + table("A") + table("B") + p("After"));
    expectClean(pages, ["Before", "A", "B", "After"], ["Before", "After"]);
  });

  it("lays out a column-break-only paragraph and then a table in two columns", () => {
    const cols = `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/>${PGMAR}<w:cols w:num="2" w:space="720"/></w:sectPr></w:pPr></w:p>`;
    const pages = layout(
      `<w:p><w:pPr><w:sectPr>${PGMAR}</w:sectPr></w:pPr><w:r><w:t>Intro</w:t></w:r></w:p>` +
        p("Before") +
        '<w:p><w:r><w:br w:type="column"/></w:r></w:p>' +
        table("A", "one two three four five six seven eight") +
        p("After") +
        cols,
    );
    expectClean(pages, ["Before", "A", "After"], ["Intro", "Before", "After"]);
  });

  it("lays out a page-break-only paragraph and then a picture", () => {
    const pages = layout(p("Before") + PAGE_BREAK_ONLY + picture + p("After"));
    expect(pagesOfWords(pages, ["Before", "[picture]", "After"])).toEqual([1, 2, 2]);
    expectClean(pages, ["Before", "[picture]", "After"], ["Before", "After"]);
  });

  it("lays out a page-break-only paragraph and then a text box", () => {
    const pages = layout(p("Before") + PAGE_BREAK_ONLY + textBox + p("After"));
    expect(pagesOfWords(pages, ["Before", "BOX", "After"])).toEqual([1, 2, 2]);
    expectClean(pages, ["Before", "BOX", "After"], ["Before", "After"]);
  });

  it("keeps a page-anchored text box on the page of its anchor, after a page break", () => {
    const pages = layout(p("Before") + PAGE_BREAK_ONLY + pageBox(100) + p("After"));
    expect(pagesOfWords(pages, ["Before", "PAGEBOX", "After"])).toEqual([1, 2, 2]);
    expectClean(pages, ["Before", "PAGEBOX", "After"], ["Before", "After"]);
  });

  it("lays out a page break inside a table cell", () => {
    const cell = '<w:p><w:r><w:t>C1</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>C2</w:t></w:r></w:p>';
    const pages = layout(p("Before") + oneCell(cell) + p("After"));
    expectClean(pages, ["Before", "C1", "C2", "After"], ["Before", "After"]);
  });

  it("lays out a page-break-only paragraph inside a cell, then a nested table", () => {
    const cell = p("Outer") + PAGE_BREAK_ONLY + table("A", "nested text") + p("Tail");
    const pages = layout(p("Before") + oneCell(cell) + p("After"));
    expectClean(pages, ["Before", "Outer", "A", "nested", "Tail", "After"], ["Before", "After"]);
  });

  it("lays out a page-break-only paragraph and then a table that opens with a nested table", () => {
    const pages = layout(p("Before") + PAGE_BREAK_ONLY + oneCell(table("A", "nested text") + p("Tail")) + p("After"));
    expect(pagesOfWords(pages, ["Before", "A", "Tail", "After"])).toEqual([1, 2, 2, 2]);
    expectClean(pages, ["Before", "A", "nested", "Tail", "After"], ["Before", "After"]);
  });

  it("lays out a next-page section break and then a table", () => {
    const pages = layout(
      `<w:p><w:pPr><w:sectPr>${PGMAR}</w:sectPr></w:pPr><w:r><w:t>Before</w:t></w:r></w:p>` + table() + p("After"),
    );
    expect(pagesOfWords(pages, ["Before", "A", "After"])).toEqual([1, 2, 2]);
    expectClean(pages, ["Before", "A", "twentyfive", "After"], ["Before", "After"]);
  });

  it("lays out a page-break-only paragraph that also ends a section, then a table", () => {
    const pages = layout(
      p("Before") +
        `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/>${PGMAR}</w:sectPr></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>` +
        table() +
        p("After"),
    );
    expectClean(pages, ["Before", "A", "twentyfive", "After"], ["Before", "After"]);
  });

  // Validator doc97-check-d, minimal file anchor_then_br.docx: the break that
  // follows only an anchored shape was dropped, so "ZZZ" stayed on page 1.
  // The shape stays on page 1 with the line that holds its anchor: 000677.doc
  // anchors its cover-page banner this way, 435pt below a heading that opens
  // with [anchor][page break], and the banner belongs at the cover's foot.
  it("keeps a page break that follows only an anchored shape", () => {
    const pages = layout(
      p("AAA") + `<w:p>${redBox("wrapNone", 0, 914400)}<w:r><w:br w:type="page"/></w:r><w:r><w:t>ZZZ</w:t></w:r></w:p>` + p("BBB"),
    );
    expect(pagesOfWords(pages, ["AAA", "ZZZ", "BBB"])).toEqual([1, 2, 2]);
    expect(redPages(pages)).toEqual([1]);
  });

  it("keeps a page break that follows an anchored shape and text", () => {
    const pages = layout(
      p("AAA") +
        `<w:p>${redBox("wrapNone", 0, 914400)}<w:r><w:t>Q</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>ZZZ</w:t></w:r></w:p>` +
        p("BBB"),
    );
    expect(pagesOfWords(pages, ["AAA", "Q", "ZZZ", "BBB"])).toEqual([1, 1, 2, 2]);
    expect(redPages(pages)).toEqual([1]);
  });

  // Validator doc97-check-d, 085236.doc: a SHAPE field whose result is only an
  // anchored drawing, then a page break. The field draws no text, so the line
  // breaker took the break as a leading break, but placeParagraph did not.
  it("keeps a page break that follows a field that draws no text", () => {
    const shapeField =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SHAPE  \\* MERGEFORMAT </w:instrText></w:r>' +
      `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${redBox("wrapNone", 0, 914400)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const pages = layout(
      p("AAA") + `<w:p>${shapeField}<w:r><w:br w:type="page"/></w:r><w:r><w:t>ZZZ</w:t></w:r></w:p>` + p("BBB"),
    );
    expect(pagesOfWords(pages, ["AAA", "ZZZ", "BBB"])).toEqual([1, 2, 2]);
  });

  // word95-check-a R1, 109561.doc: a numbered heading whose first run is a
  // page break. The break moved the paragraph to a new page, and the label
  // then went alone onto that page with the heading text one page later.
  it("keeps a numbered heading's label with its text after a leading page break", () => {
    const numbering =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NS}>` +
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/>' +
      '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr></w:lvl></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
    const heading = (text: string) =>
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
      `<w:r><w:br w:type="page"/></w:r><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const pages = layout(p("AAA") + heading("CITE") + p("Body1") + heading("SECOND") + p("Body2"), {
      "word/numbering.xml": numbering,
    });
    expect(pages.length).toBe(3);
    expect(pagesOfWords(pages, ["AAA", "I.", "CITE", "Body1", "II.", "SECOND", "Body2"])).toEqual([1, 2, 2, 2, 3, 3, 3]);
    const label = boxesOf(pages).find((b) => b.label === "I.")!;
    const text = boxesOf(pages).find((b) => b.label === "CITE")!;
    expect(label.y0).toBeCloseTo(text.y0, 1);
  });

  // Validator doc97-check-d, minimal files br_smallTB_br.docx and
  // br_tallTB_br.docx: the shape was drawn on page 1, before the break.
  for (const [name, height] of [["small", 914400], ["tall", 7772400]] as const) {
    it(`draws a ${name} top-and-bottom shape after a page break on the next page`, () => {
      const pages = layout(
        p("AAA") + PAGE_BREAK_ONLY + `<w:p>${redBox("wrapTopAndBottom", -152400, height)}</w:p>` + PAGE_BREAK_ONLY + p("ZZZ"),
      );
      expect(pages.length).toBe(3);
      expect(pagesOfWords(pages, ["AAA", "ZZZ"])).toEqual([1, 3]);
      expect(redPages(pages)).toEqual([2]);
    });
  }
});
