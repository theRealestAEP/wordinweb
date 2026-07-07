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
    // No tcW anywhere: Word ignores the authored grid and autofits columns
    // to content (probe-tablegrid), so the second column hugs the first.
    expect(c1.x - c0.x).toBeGreaterThan(20);
    expect(c1.x - c0.x).toBeLessThan(90);
  });

  it("autofits columns to content when the grid is a placeholder", () => {
    // Junk grid (100 twips per column, like generator output) + 100% width:
    // Word ignores the grid and sizes columns by content.
    const tbl = `<w:tbl>
      <w:tblPr><w:tblW w:type="pct" w:w="100%"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Tiny</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Mid col</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>A distinctly longer notes column with plenty of words</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(tbl + p("after")) });
    const cell = (t: string) => {
      const it = result.pages[0].items.find((i) => i.kind === "text" && i.text.includes(t));
      if (it?.kind !== "text") throw new Error(`cell ${t} not found`);
      return it;
    };
    const x0 = cell("Tiny").x;
    const x1 = cell("Mid").x;
    const x2 = cell("distinctly").x;
    const w0 = x1 - x0;
    const w1 = x2 - x1;
    // An even split would give each column ~1/3 of the content width
    // (~200px). Content-proportional sizing keeps the short columns narrow,
    // leaving the notes column with the bulk of the width.
    expect(w0).toBeLessThan(120);
    expect(w1).toBeLessThan(160);
    expect(w0 + w1).toBeLessThan(300);
    expect(w0).toBeLessThan(w1 + 40); // both stay near content size
  });

  it("autofits a tcW-less grid to content like Word", () => {
    // Word only trusts a grid it wrote itself, and it writes tcW on every
    // cell. A plausible-looking grid with no tcW is ignored: Word sizes the
    // "x" column to its content (5.75pt in the probe-tablegrid export).
    const tbl = `<w:tbl>
      <w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>a much much much longer cell body here</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(tbl + p("after")) });
    const a = result.pages[0].items.find((i) => i.kind === "text" && i.text.includes("x"));
    const b = result.pages[0].items.find((i) => i.kind === "text" && i.text === "a");
    if (a?.kind !== "text" || b?.kind !== "text") throw new Error("cells not found");
    expect(b.x - a.x).toBeLessThan(40);
  });

  it("honors a Word-authored grid (cells carry tcW)", () => {
    const tbl = `<w:tbl>
      <w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4680"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4680"/></w:tcPr><w:p><w:r><w:t>a much much much longer cell body here</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`;
    const { result } = layout({ "word/document.xml": wrapDocument(tbl + p("after")) });
    const a = result.pages[0].items.find((i) => i.kind === "text" && i.text.includes("x"));
    const b = result.pages[0].items.find((i) => i.kind === "text" && i.text === "a");
    if (a?.kind !== "text" || b?.kind !== "text") throw new Error("cells not found");
    // Equal 4680-twip columns: the second cell starts at the halfway point.
    expect(b.x - a.x).toBeGreaterThan(280);
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

describe("tab leaders", () => {
  it("fills dotted leaders up to the tab stop (TOC pattern)", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr>
          <w:r><w:t>Chapter One</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>4</w:t></w:r>
        </w:p>`,
      ),
    });
    const dots = result.pages[0].items.find(
      (i) => i.kind === "text" && /^\.{10,}$/.test(i.text),
    );
    expect(dots).toBeDefined();
    if (dots?.kind !== "text") return;
    const num = result.pages[0].items.find((i) => i.kind === "text" && i.text === "4");
    if (num?.kind !== "text") throw new Error("page number missing");
    expect(dots.x + dots.width).toBeLessThanOrEqual(num.x + 4);
  });
});

describe("footnotes and endnotes", () => {
  const FN_RELS = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rIdEn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>`;

  const footnotesXml = (notes: string) => `<?xml version="1.0"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  ${notes}
</w:footnotes>`;

  const endnotesXml = (notes: string) => `<?xml version="1.0"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  ${notes}
</w:endnotes>`;

  const note = (tag: "footnote" | "endnote", id: number, text: string) =>
    `<w:${tag} w:id="${id}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:${tag}Ref/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:${tag}>`;

  it("renders a footnote at the bottom of the referencing page with a separator", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:r><w:t>Body text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`,
      ),
      "word/_rels/document.xml.rels": FN_RELS,
      "word/footnotes.xml": footnotesXml(note("footnote", 1, "alpha note text")),
    });
    expect(result.totalPages).toBe(1);
    const page = result.pages[0];
    expect(pageText(result, 0)).toContain("alpha note text");

    // Reference mark and the note's own mark are both superscript "1"s.
    const marks = page.items.filter(
      (i) => i.kind === "text" && i.text === "1" && i.props.verticalAlign === "superscript",
    );
    expect(marks.length).toBe(2);

    // Separator rule: a short (2in) line above the note, below the body text.
    const body = page.items.find((i) => i.kind === "text" && i.text.includes("Body"));
    if (body?.kind !== "text") throw new Error("body text missing");
    const sep = page.items.find(
      (i) => i.kind === "edge" && Math.abs(i.x2 - i.x1 - 192) < 1 && i.y1 === i.y2,
    );
    expect(sep).toBeDefined();
    if (sep?.kind !== "edge") return;
    expect(sep.y1).toBeGreaterThan(body.lineTop);

    // Note text sits at the bottom of the body box (1in margins on US Letter).
    const noteText = page.items.find((i) => i.kind === "text" && i.text.includes("alpha"));
    if (noteText?.kind !== "text") throw new Error("note text missing");
    expect(noteText.lineTop).toBeGreaterThan(sep.y1);
    expect(noteText.lineTop).toBeGreaterThan(880);
    expect(noteText.lineTop + noteText.lineHeight).toBeLessThanOrEqual(page.bodyBottom + 0.5);
  });

  it("binds a footnote in a split table row to the page painting its partition", () => {
    // A single row taller than a page: the split's rest-partition carries a
    // footnote reference and must register the note on page 2.
    const cellParas =
      Array.from({ length: 70 }, (_, i) => `<w:p><w:r><w:t>Row line ${i}</w:t></w:r></w:p>`).join("") +
      `<w:p><w:r><w:t xml:space="preserve">noted line</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`;
    const tbl =
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="9026"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="9026" w:type="dxa"/></w:tcPr>${cellParas}</w:tc></w:tr></w:tbl>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(tbl + `<w:p><w:r><w:t>after table</w:t></w:r></w:p>`),
      "word/_rels/document.xml.rels": FN_RELS,
      "word/footnotes.xml": footnotesXml(note("footnote", 1, "split row note")),
    });
    expect(result.totalPages).toBeGreaterThan(1);
    // The reference line and the note text must share a page.
    const refPage = result.pages.findIndex((pg) => pg.items.some((i) => i.kind === "text" && i.text.includes("noted")));
    const notePage = result.pages.findIndex((pg) => pg.items.some((i) => i.kind === "text" && i.text.includes("split")));
    expect(refPage).toBeGreaterThan(0);
    expect(notePage).toBe(refPage);
  });

  it("keeps each footnote on the same page as its reference", () => {
    const filler = (n: number, from = 0) =>
      Array.from({ length: n }, (_, i) => p(`Filler paragraph ${from + i}`)).join("");
    const refPara = (word: string, id: number) =>
      `<w:p><w:r><w:t xml:space="preserve">${word}</w:t></w:r><w:r><w:footnoteReference w:id="${id}"/></w:r></w:p>`;
    const { result } = layout({
      "word/document.xml": wrapDocument(
        refPara("HEADREF", 1) + filler(70) + refPara("TAILREF", 2) + filler(5, 70),
      ),
      "word/_rels/document.xml.rels": FN_RELS,
      "word/footnotes.xml": footnotesXml(
        note("footnote", 1, "alphanote") + note("footnote", 2, "betanote"),
      ),
    });
    expect(result.totalPages).toBeGreaterThan(1);
    const pageOf = (needle: string) =>
      result.pages.findIndex((pg) =>
        pg.items.some((i) => i.kind === "text" && i.text.includes(needle)),
      );
    expect(pageOf("HEADREF")).toBe(0);
    expect(pageOf("alphanote")).toBe(0);
    const tailPage = pageOf("TAILREF");
    expect(tailPage).toBeGreaterThan(0);
    expect(pageOf("betanote")).toBe(tailPage);
    // Numbering follows document order.
    expect(pageText(result, 0)).toContain("alphanote");
    const tailMarks = result.pages[tailPage].items.filter(
      (i) => i.kind === "text" && i.text === "2" && i.props.verticalAlign === "superscript",
    );
    expect(tailMarks.length).toBe(2);
  });

  it("shrinks the body so footnote content never overlaps the footer margin", () => {
    const longNote = Array.from({ length: 4 }, (_, i) => p(`NOTETOKEN${i} with a good amount of text to wrap around`)).join("");
    const filler = Array.from({ length: 60 }, (_, i) => p(`Body ${i}`)).join("");
    const { result } = layout({
      "word/document.xml": wrapDocument(
        filler + `<w:p><w:r><w:t>REFHERE</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`,
      ),
      "word/_rels/document.xml.rels": FN_RELS,
      "word/footnotes.xml": footnotesXml(
        `<w:footnote w:id="1">${longNote}</w:footnote>`,
      ),
    });
    for (const pg of result.pages) {
      for (const item of pg.items.slice(0, pg.hfStart)) {
        if (item.kind === "text") {
          expect(item.lineTop + item.lineHeight).toBeLessThanOrEqual(pg.bodyBottom + 0.5);
        }
      }
    }
    const refPage = result.pages.findIndex((pg) =>
      pg.items.some((i) => i.kind === "text" && i.text.includes("REFHERE")),
    );
    const notePage = result.pages.findIndex((pg) =>
      pg.items.some((i) => i.kind === "text" && i.text.includes("NOTETOKEN0")),
    );
    expect(notePage).toBe(refPage);
  });

  it("flows endnotes after the last body block with lowerRoman marks", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:r><w:t>Body start</w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>` +
          p("Body end"),
      ),
      "word/_rels/document.xml.rels": FN_RELS,
      "word/endnotes.xml": endnotesXml(note("endnote", 1, "closing remark")),
    });
    expect(result.totalPages).toBe(1);
    const page = result.pages[0];
    expect(pageText(result, 0)).toContain("closing remark");
    const marks = page.items.filter(
      (i) => i.kind === "text" && i.text === "i" && i.props.verticalAlign === "superscript",
    );
    expect(marks.length).toBe(2);
    // Endnote content comes after the last body paragraph, not at page bottom.
    const bodyEnd = page.items.find((i) => i.kind === "text" && i.text === "end");
    const noteText = page.items.find((i) => i.kind === "text" && i.text.includes("closing"));
    if (bodyEnd?.kind !== "text" || noteText?.kind !== "text") throw new Error("items missing");
    expect(noteText.lineTop).toBeGreaterThan(bodyEnd.lineTop);
    expect(noteText.lineTop).toBeLessThan(400);
  });

  it("honors sectPr footnote number format", () => {
    const { result } = layout({
      "word/document.xml": wrapDocument(
        `<w:p><w:r><w:t>Text</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>` +
          `<w:sectPr><w:footnotePr><w:numFmt w:val="chicago"/></w:footnotePr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>
          </w:sectPr>`,
      ),
      "word/_rels/document.xml.rels": FN_RELS,
      "word/footnotes.xml": footnotesXml(note("footnote", 1, "starred note")),
    });
    const marks = result.pages[0].items.filter((i) => i.kind === "text" && i.text === "*");
    expect(marks.length).toBe(2);
  });
});

describe("justified line breaking (Word pack-vs-break rule)", () => {
  // ApproxMeasurer: n/o = 0.5em, i = 0.28em, m = 0.85em, space = 0.25em,
  // em = 14.666px. Ten "no" fillers + 9 spaces = 179.6585px of line.
  const fillers = Array(10).fill("no").join(" ");
  const jp = (text: string) =>
    `<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const sect = (pgW: number) =>
    `<w:sectPr><w:pgSz w:w="${pgW}" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`;
  const linesOf = (result: ReturnType<typeof layoutDocument>) => {
    const tops = new Map<number, string[]>();
    for (const i of result.pages[0].items) {
      if (i.kind !== "text") continue;
      const key = Math.round(i.lineTop);
      if (!tops.has(key)) tops.set(key, []);
      tops.get(key)!.push(i.text);
    }
    return [...tops.entries()].sort((a, b) => a[0] - b[0]).map(([, t]) => t.join("").replace(/\s+/g, " ").trim());
  };

  it("packs a wide final word at ~20% compression (break would leave a gaping line)", () => {
    // content 225.87px; "mmmm" needs 20% space compression, the break
    // alternative a 140% stretch: Word packs (compress <= stretch/2, <= 25%).
    const { result } = layout({
      "word/document.xml": wrapDocument(jp(`${fillers} mmmm`) + sect(6268)),
    });
    const lines = linesOf(result);
    expect(lines[0].endsWith("mmmm")).toBe(true);
  });

  it("breaks before a narrow final word at ~16% compression (stretch is cheap)", () => {
    // content 188.87px; "in" needs only 16% compression but breaking costs a
    // mere 28% stretch: Word breaks (16% > 28%/2).
    const { result } = layout({
      "word/document.xml": wrapDocument(jp(`${fillers} in`) + sect(5713)),
    });
    const lines = linesOf(result);
    expect(lines[0].endsWith("no")).toBe(true);
    expect(lines[1]).toBe("in");
  });

  it("never splits a word at a formatting-run boundary", () => {
    // "c" (run 1) + "cc" (bold run 2) form one word; "c" alone fits on line 1
    // but the word must move down as a unit.
    const para =
      '<w:p><w:r><w:t xml:space="preserve">aa bb c</w:t></w:r>' +
      '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">cc</w:t></w:r></w:p>';
    const { result } = layout({ "word/document.xml": wrapDocument(para + sect(3650)) });
    const lines = linesOf(result);
    expect(lines[0]).toBe("aa bb");
    expect(lines[1]).toBe("ccc");
  });
});

describe("superscript / subscript", () => {
  it("raises superscript and drops subscript from the baseline like Word", () => {
    const para =
      '<w:p><w:r><w:t xml:space="preserve">base </w:t></w:r>' +
      '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>sup</w:t></w:r>' +
      '<w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>sub</w:t></w:r></w:p>';
    const { result } = layout({ "word/document.xml": wrapDocument(para) });
    const items = result.pages[0].items.filter((i) => i.kind === "text");
    const base = items.find((i) => i.kind === "text" && i.text === "base")!;
    const sup = items.find((i) => i.kind === "text" && i.text === "sup")!;
    const sub = items.find((i) => i.kind === "text" && i.text === "sub")!;
    if (base.kind !== "text" || sup.kind !== "text" || sub.kind !== "text") throw new Error();
    // Word (probe-vertalign): raise 7/22 and drop 1/11 of the UNSCALED size;
    // scaled size is 65% rounded to half-points (14.666px -> 9.333px).
    expect(sup.font.size).toBeCloseTo(9.3333, 3);
    expect(base.baseline - sup.baseline).toBeCloseTo(14.666 * (7 / 22), 2);
    expect(sub.baseline - base.baseline).toBeCloseTo(14.666 / 11, 2);
    // Renderer anchor: explicit glyph box, baseline-aligned.
    expect(sup.glyphTop).toBeCloseTo(sup.baseline - 0.9 * sup.font.size, 2);
    expect(sup.glyphBoxH).toBeCloseTo(1.15 * sup.font.size, 2);
    expect(base.glyphTop).toBeUndefined();
  });
});

describe("sections & page borders", () => {
  it("continuous section shares the page; new-page section does not", () => {
    const cont = wrapDocument(
      p("first section text") +
        `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr></w:p>` +
        p("second section text") +
        `<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`,
    );
    const { result } = layout({ "word/document.xml": cont });
    expect(result.totalPages).toBe(1);
    // and without continuous: two pages
    const hard = cont.replace('<w:type w:val="continuous"/>', "");
    const { result: r2 } = layout({ "word/document.xml": hard });
    expect(r2.totalPages).toBe(2);
  });

  it("renders page borders inset from the text margins", () => {
    const docXml = wrapDocument(
      p("bordered page") +
        `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>` +
        `<w:pgBorders w:offsetFrom="text"><w:top w:val="single" w:sz="8" w:space="24" w:color="FF0000"/>` +
        `<w:left w:val="single" w:sz="8" w:space="24"/><w:bottom w:val="single" w:sz="8" w:space="24"/>` +
        `<w:right w:val="single" w:sz="8" w:space="24"/></w:pgBorders></w:sectPr>`,
    );
    const { result } = layout({ "word/document.xml": docXml });
    const edges = result.pages[0].items.filter((i) => i.kind === "edge");
    expect(edges.length).toBe(4);
    const top = edges.find((e) => e.kind === "edge" && e.y1 === e.y2 && e.y1 < 100);
    if (!top || top.kind !== "edge") throw new Error();
    expect(top.border.color).toBe("#FF0000");
    // 1in margin (96px) minus 24pt (32px) offset = 64px
    expect(top.y1).toBeCloseTo(64, 0);
  });
});

describe("table row splitting", () => {
  const bigRow = (n: number, extra = "") => {
    const paras = Array.from({ length: n }, (_, i) => `<w:p><w:r><w:t>cell line ${i}</w:t></w:r></w:p>`).join("");
    return `<w:tbl><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid>
      <w:tr>${extra}<w:tc>${paras}</w:tc></w:tr></w:tbl>`;
  };

  it("splits a too-tall row across pages by default", () => {
    // ~90 lines at ~17px in a ~700px body: must span pages, content intact.
    const { result } = layout({ "word/document.xml": wrapDocument(bigRow(90)) });
    expect(result.totalPages).toBeGreaterThan(1);
    const all = result.pages.flatMap((pg) => pg.items.filter((i) => i.kind === "text"));
    expect(all.some((i) => i.kind === "text" && i.text.includes("0"))).toBe(true);
    const last = result.pages[result.pages.length - 1].items.filter((i) => i.kind === "text");
    expect(last.length).toBeGreaterThan(0);
    // every page's text stays inside the body box
    for (const pg of result.pages) {
      for (const it of pg.items) {
        if (it.kind === "text") expect(it.lineTop + it.lineHeight).toBeLessThanOrEqual(pg.height - 90);
      }
    }
  });

  it("honors cantSplit by moving the row whole", () => {
    const xml = wrapDocument(p("lead") + bigRow(20, '<w:trPr><w:cantSplit/></w:trPr>') );
    // Fill most of page 1 first so the row doesn't fit
    const filler = Array.from({ length: 42 }, (_, i) => p(`filler ${i}`)).join("");
    const { result } = layout({ "word/document.xml": wrapDocument(filler + bigRow(25, "<w:trPr><w:cantSplit/></w:trPr>")) });
    void xml;
    // The row starts on page 2 (its first cell line is not on page 1).
    const p1 = result.pages[0].items.filter((i) => i.kind === "text").map((i) => (i.kind === "text" ? i.text : ""));
    expect(p1.some((t) => t.includes("cell"))).toBe(false);
  });
});
