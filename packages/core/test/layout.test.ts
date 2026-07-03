import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { formatNumber } from "../src/parse/numbering.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";

const measurer = new ApproxMeasurer();

function layout(parts: Record<string, string>) {
  const doc = DocxDocument.load(makeDocx(parts));
  return { doc, result: layoutDocument(doc, { measurer }) };
}

function pageText(result: ReturnType<typeof layoutDocument>, pageIdx: number): string {
  return result.pages[pageIdx].items
    .filter((i) => i.kind === "text")
    .map((i) => (i.kind === "text" ? i.text : ""))
    .join("");
}

describe("number formatting", () => {
  it("formats roman and letters", () => {
    expect(formatNumber(4, "lowerRoman")).toBe("iv");
    expect(formatNumber(1949, "upperRoman")).toBe("MCMXLIX");
    expect(formatNumber(1, "upperLetter")).toBe("A");
    expect(formatNumber(27, "lowerLetter")).toBe("aa");
  });
});

describe("layout engine", () => {
  it("produces a single page for a short document", () => {
    const { result } = layout({ "word/document.xml": wrapDocument(p("Hello world")) });
    expect(result.totalPages).toBe(1);
    expect(pageText(result, 0)).toContain("Hello");
  });

  it("paginates long content onto multiple pages", () => {
    const paras = Array.from({ length: 120 }, (_, i) => p(`Paragraph number ${i}`)).join("");
    const { result } = layout({ "word/document.xml": wrapDocument(paras) });
    expect(result.totalPages).toBeGreaterThan(1);
    // Content must stay inside the body box (1in margins on US Letter).
    for (const page of result.pages) {
      for (const item of page.items) {
        if (item.kind === "text") {
          expect(item.lineTop).toBeGreaterThanOrEqual(95); // ~96px margin
          expect(item.lineTop + item.lineHeight).toBeLessThanOrEqual(page.height - 95);
        }
      }
    }
  });

  it("honors explicit page breaks", () => {
    const body =
      p("first page") +
      `<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>second page</w:t></w:r></w:p>`;
    const { result } = layout({ "word/document.xml": wrapDocument(body) });
    expect(result.totalPages).toBe(2);
    expect(pageText(result, 0)).toContain("first page");
    expect(pageText(result, 0)).not.toContain("second page");
    expect(pageText(result, 1)).toContain("second page");
  });

  it("resolves PAGE and NUMPAGES fields in footers per page", () => {
    const paras = Array.from({ length: 120 }, (_, i) => p(`Paragraph ${i}`)).join("");
    const { result } = layout({
      "word/document.xml": wrapDocument(
        paras +
          `<w:sectPr>
            <w:footerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdF"/>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
          </w:sectPr>`,
      ),
      "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`,
      "word/footer1.xml": `<?xml version="1.0"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Page </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:t xml:space="preserve"> of </w:t></w:r>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:p>
</w:ftr>`,
    });
    const total = result.totalPages;
    expect(total).toBeGreaterThan(1);
    for (let i = 0; i < total; i++) {
      const text = pageText(result, i);
      expect(text).toContain(`Page ${i + 1} of ${total}`);
    }
  });

  it("computes numbering labels with restarts", () => {
    const numberingXml = `<?xml version="1.0"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;
    const numPara = (text: string, ilvl: number) =>
      `<w:p><w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(
        numPara("one", 0) + numPara("two", 0) + numPara("sub-a", 1) + numPara("sub-b", 1) + numPara("three", 0) + numPara("sub-restart", 1),
      ),
      "word/numbering.xml": numberingXml,
    });
    const text = pageText(result, 0);
    expect(text).toContain("1.");
    expect(text).toContain("2.");
    expect(text).toContain("a)");
    expect(text).toContain("b)");
    expect(text).toContain("3.");
    // After returning to level 0, level 1 restarts at "a)"
    const idxThree = text.indexOf("3.");
    expect(text.slice(idxThree)).toContain("a)");
  });

  it("lays out table rows and repeats content within page", () => {
    const rows = Array.from(
      { length: 3 },
      (_, i) => `<w:tr>
        <w:tc><w:p><w:r><w:t>R${i}C0</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>R${i}C1</w:t></w:r></w:p></w:tc>
      </w:tr>`,
    ).join("");
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${rows}</w:tbl>` + p("after table"),
      ),
    });
    const text = pageText(result, 0);
    for (let i = 0; i < 3; i++) {
      expect(text).toContain(`R${i}C0`);
      expect(text).toContain(`R${i}C1`);
    }
    // Cells in the same row share a top; columns are offset horizontally.
    const c0 = result.pages[0].items.find((it) => it.kind === "text" && it.text.includes("R0C0"));
    const c1 = result.pages[0].items.find((it) => it.kind === "text" && it.text.includes("R0C1"));
    if (c0?.kind !== "text" || c1?.kind !== "text") throw new Error("cells not found");
    expect(Math.abs(c0.lineTop - c1.lineTop)).toBeLessThan(0.5);
    expect(c1.x).toBeGreaterThan(c0.x + 100);
  });

  it("applies pgNumType start to display numbers", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        p("content") + `<w:sectPr><w:pgNumType w:start="5"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`,
      ),
    });
    expect(result.pages[0].number).toBe(5);
  });

  it("draws a paragraph bottom border as a divider line", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr><w:r><w:t>above the line</w:t></w:r></w:p>`,
      ),
    });
    const edges = result.pages[0].items.filter((i) => i.kind === "edge");
    expect(edges.length).toBe(1);
    if (edges[0].kind !== "edge") return;
    expect(Math.abs(edges[0].y1 - edges[0].y2)).toBeLessThan(0.01); // horizontal
  });

  it("keeps header text in the header zone and pushes body below it", () => {
    const { result, doc } = layout({
      "word/document.xml": wrapDocument(
        p("body text") +
          `<w:sectPr>
            <w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdH"/>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
          </w:sectPr>`,
      ),
      "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`,
      "word/header1.xml": `<?xml version="1.0"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${p("HEADER")}</w:hdr>`,
    });
    const items = result.pages[0].items.filter((i) => i.kind === "text");
    const headerItem = items.find((i) => i.kind === "text" && i.text === "HEADER");
    const bodyItem = items.find((i) => i.kind === "text" && i.text.includes("body"));
    if (headerItem?.kind !== "text" || bodyItem?.kind !== "text") throw new Error("items missing");
    // Header sits at headerDistance (720 twips = 48px), above the body top (96px).
    expect(headerItem.lineTop).toBeGreaterThanOrEqual(47);
    expect(headerItem.lineTop).toBeLessThan(96);
    expect(bodyItem.lineTop).toBeGreaterThanOrEqual(95);
  });
});

describe("pagination robustness", () => {
  it("terminates when orphan control cannot fit two lines on any page", () => {
    // Two lines of 400pt exact spacing (533px) on a US Letter body (~848px):
    // a lone first line at the bottom triggers the orphan push; after the
    // push the pair still cannot share a page. Must not loop forever.
    const body =
      p("filler before") +
      `<w:p><w:pPr><w:spacing w:line="8000" w:lineRule="exact"/></w:pPr>
        <w:r><w:t>first tall line</w:t><w:br/><w:t>second tall line</w:t></w:r>
      </w:p>` +
      p("after");
    const { result } = layout({ "word/document.xml": wrapDocument(body) });
    expect(result.totalPages).toBeGreaterThanOrEqual(2);
    const all = result.pages.map((_, i) => pageText(result, i)).join("");
    expect(all).toContain("first tall line");
    expect(all).toContain("second tall line");
    expect(all).toContain("after");
  });

  it("lays out the chronology-style header (exact small line + colored border)", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:color="7A6E73" w:sz="4" w:space="2"/></w:pBdr><w:spacing w:after="0" w:before="0" w:line="170" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:color w:val="7A6E73"/><w:sz w:val="17"/></w:rPr><w:t>Created by Cobbery</w:t></w:r></w:p>` +
          p("body"),
      ),
    });
    const txt = result.pages[0].items.find((i) => i.kind === "text" && i.text.includes("Created"));
    if (txt?.kind !== "text") throw new Error("missing header text");
    expect(txt.props.color).toBe("#7A6E73");
    const edge = result.pages[0].items.find((i) => i.kind === "edge");
    if (edge?.kind !== "edge") throw new Error("missing border edge");
    expect(edge.border.color).toBe("#7A6E73");
  });
});

describe("floating images", () => {
  it("wraps text beside a square-anchored image", () => {
    const anchor = `<w:p><w:r><w:drawing>
      <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
        <wp:extent cx="1905000" cy="1905000"/>
        <wp:wrapSquare wrapText="bothSides"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdImg"/></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1905000"/></a:xfrm></pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing></w:r>
    <w:r><w:t>${"words flow beside the floating image ".repeat(40)}</w:t></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(anchor),
      "word/_rels/document.xml.rels": `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/x.png"/>
</Relationships>`,
      "word/media/x.png": "PNGDATA",
    });
    const items = result.pages[0].items;
    const img = items.find((i) => i.kind === "image");
    expect(img).toBeDefined();
    if (img?.kind !== "image") return;
    expect(Math.round(img.width)).toBe(200); // 1905000 EMU = 200px
    // Text beside the image starts to its right; text below spans full width.
    const beside = items.filter(
      (i) => i.kind === "text" && i.text.trim() && i.lineTop < img.y + img.height,
    );
    const below = items.filter(
      (i) => i.kind === "text" && i.text.trim() && i.lineTop > img.y + img.height + 2,
    );
    expect(beside.length).toBeGreaterThan(0);
    expect(below.length).toBeGreaterThan(0);
    for (const t of beside) {
      if (t.kind !== "text") continue;
      expect(t.x).toBeGreaterThanOrEqual(img.x + img.width);
    }
    expect(below.some((t) => t.kind === "text" && t.x < img.x + img.width)).toBe(true);
  });
});
