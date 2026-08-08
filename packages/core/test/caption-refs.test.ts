import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { serializeXml } from "../src/xml.js";
import {
  ensureRefBookmark,
  insertCaptionAt,
  insertCrossReference,
  listCrossRefTargets,
  nextRefBookmarkName,
} from "../src/edit/references.js";
import { insertToc, rebuildToc, findTocFields, tocEntryCount } from "../src/edit/toc.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";
import type { Paragraph, Run } from "../src/model.js";

const STYLES = `<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`;

const CONTENT_TYPES = `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const DOC_RELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

function loadDoc(bodyXml: string): DocxDocument {
  return DocxDocument.load(makeDocx({
    "word/document.xml": wrapDocument(bodyXml),
    "word/styles.xml": STYLES,
    "word/_rels/document.xml.rels": DOC_RELS,
    "[Content_Types].xml": CONTENT_TYPES,
  }));
}

const captionP = (label: string, n: number, tail: string) =>
  `<w:p><w:r><w:t xml:space="preserve">${label} </w:t></w:r>` +
  `<w:fldSimple w:instr=" SEQ ${label} \\* ARABIC "><w:r><w:t>${n}</w:t></w:r></w:fldSimple>` +
  `<w:r><w:t xml:space="preserve"> ${tail}</w:t></w:r></w:p>`;

const firstT = (doc: DocxDocument, blockIdx = 0) =>
  ((doc.sections[0].blocks[blockIdx] as Paragraph).children[0] as Run).content.find((c) => c.kind === "text")!.srcT!;

describe("insertCaptionAt (SEQ captions)", () => {
  it("inserts a Caption-styled SEQ paragraph below the anchor, numbered from document order", () => {
    const doc = loadDoc(captionP("Figure", 1, "old chart") + p("the picture paragraph"));
    expect(insertCaptionAt(doc, firstT(doc, 1), "Figure", "new chart", "below")).toBe(true);
    const xml = serializeXml(doc.docRoot);
    expect(xml).toContain(`<w:fldSimple w:instr=" SEQ Figure \\* ARABIC "><w:r><w:t xml:space="preserve">2</w:t></w:r></w:fldSimple>`);
    expect(xml).toContain(`<w:pStyle w:val="Caption"/>`);
    expect(xml).toContain("new chart");
    // The Caption style definition was injected for the file.
    expect(doc.styles.byId.has("Caption")).toBe(true);
    // Order: the caption paragraph FOLLOWS the anchor.
    expect(doc.sections[0].blocks).toHaveLength(3);
    const cap = doc.sections[0].blocks[2] as Paragraph;
    expect(cap.props.styleId ?? (cap.src && serializeXml(cap.src).includes("Caption"))).toBeTruthy();
  });

  it("an above-insert excludes the anchor's own captions from the seed", () => {
    const doc = loadDoc(p("target") + captionP("Figure", 1, "later"));
    expect(insertCaptionAt(doc, firstT(doc, 0), "Figure", "", "above")).toBe(true);
    // First caption in document order — numbered 1 at insert time.
    const xml = serializeXml(doc.docRoot);
    const first = xml.indexOf("SEQ Figure");
    expect(xml.slice(first, first + 120)).toContain(">1<");
  });

  it("a caret inside a table cell captions the TABLE (hoists to the tbl)", () => {
    const doc = loadDoc(
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>` +
      `<w:tr><w:tc><w:p><w:r><w:t xml:space="preserve">cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
      p("after"),
    );
    const cellT = (() => {
      const tbl = doc.sections[0].blocks[0];
      if (tbl.type !== "table") throw new Error("expected table");
      const para = tbl.rows[0].cells[0].blocks[0] as Paragraph;
      return (para.children[0] as Run).content.find((c) => c.kind === "text")!.srcT!;
    })();
    expect(insertCaptionAt(doc, cellT, "Table", "results", "above")).toBe(true);
    // The caption is a TOP-LEVEL block before the table, not inside the cell.
    expect(doc.sections[0].blocks[0].type).toBe("paragraph");
    expect(doc.sections[0].blocks[1].type).toBe("table");
    expect(serializeXml(doc.docRoot)).toContain("SEQ Table");
  });

  it("rejects a label that could smuggle switches", () => {
    const doc = loadDoc(p("x"));
    expect(insertCaptionAt(doc, firstT(doc), 'Fig" \\h', "", "below")).toBe(false);
  });
});

describe("cross-reference targets (headings / captions / numbered items)", () => {
  const fixture = () => loadDoc(
    `<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t xml:space="preserve">Chapter One</w:t></w:r></w:p>` +
    captionP("Figure", 1, "a chart") +
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">first item</w:t></w:r></w:p>` +
    p("plain paragraph"),
  );

  it("lists them in document order with resolved text", () => {
    const doc = fixture();
    const targets = listCrossRefTargets(doc);
    expect(targets.map((t) => t.kind)).toEqual(["heading", "caption", "numberedItem"]);
    expect(targets[0].text).toBe("Chapter One");
    expect(targets[1].text).toBe("Figure 1 a chart"); // SEQ result included
    expect(targets[2].text).toBe("first item");
  });

  it("ensureRefBookmark wraps the paragraph once and a REF resolves it", () => {
    const doc = fixture();
    const targets = listCrossRefTargets(doc);
    const name = nextRefBookmarkName(doc);
    expect(name).toMatch(/^_Ref\d+$/);
    expect(ensureRefBookmark(doc, targets[0].paragraph, name)).toBe(true);
    // Idempotent: a second wrap is a clean no-op.
    expect(ensureRefBookmark(doc, targets[0].paragraph, nextRefBookmarkName(doc))).toBe(false);
    // The listing now reports the existing bookmark for reuse.
    expect(listCrossRefTargets(doc)[0].bookmark).toBe(name);
    // A REF to it inserts through the ordinary cross-reference machinery.
    const at = firstT(doc, 3);
    expect(insertCrossReference(doc, at, at.text.length, name, "text")).toBe(true);
    expect(serializeXml(doc.docRoot)).toContain(`REF ${name}`);
  });
});

describe("table of figures (TOC \\c)", () => {
  const fixture = () => loadDoc(
    captionP("Figure", 1, "first chart") +
    p("middle text") +
    captionP("Figure", 2, "second chart") +
    p("caret here"),
  );

  it("insertToc with captionLabel builds a \\c field from the captions", () => {
    const doc = fixture();
    expect(tocEntryCount(doc, { captionLabel: "Figure" })).toBe(2);
    expect(insertToc(doc, firstT(doc, 3), { captionLabel: "Figure" })).toBe(true);
    const xml = serializeXml(doc.docRoot);
    expect(xml).toContain(` TOC \\h \\z \\c "Figure" `);
    expect(xml).toContain("Figure 1 first chart");
    expect(xml).toContain("Figure 2 second chart");
    expect(xml).toContain(`<w:pStyle w:val="TableofFigures"/>`);
    expect(doc.styles.byId.has("TableofFigures")).toBe(true);
    // Entries reference _Toc bookmarks wrapped around the captions.
    expect(xml).toMatch(/PAGEREF _Toc\d+/);
  });

  it("rebuildToc keeps a \\c field a FIGURES table (label read off the instruction)", () => {
    const doc = fixture();
    insertToc(doc, firstT(doc, 3), { captionLabel: "Figure" });
    // A new caption appears; rebuild must pick it up and stay caption-based.
    insertCaptionAt(doc, firstT(doc, 1), "Figure", "third chart", "below");
    const toc = findTocFields(doc)[0];
    expect(toc).toBeTruthy();
    expect(rebuildToc(doc, toc)).toBe(true);
    const xml = serializeXml(doc.docRoot);
    expect(xml).toContain(` TOC \\h \\z \\c "Figure" `);
    expect(xml).toContain("third chart");
    // Not rebuilt as a heading TOC.
    expect(xml).not.toContain(`\\o "1-3"`);
  });

  it("a bad label refuses", () => {
    const doc = fixture();
    expect(insertToc(doc, firstT(doc, 3), { captionLabel: 'Fig"ure' })).toBe(false);
    expect(tocEntryCount(doc, { captionLabel: 'Fig"ure' })).toBe(0);
  });
});
