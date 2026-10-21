// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { renderToDom } from "../src/render/dom.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

/** The rendered pages of a document in the editor's DOM. */
function domPages(bodyXml: string): HTMLElement[] {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(bodyXml) }));
  const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
  const container = document.createElement("div");
  document.body.appendChild(container);
  renderToDom(doc, result, container);
  return [...container.querySelectorAll<HTMLElement>(".dxw-page")];
}

describe("DOM renderer: tables that break across pages", () => {
  it("shows every row of a floating table taller than the page", () => {
    const num = (i: number) => String(i).padStart(3, "0");
    const rows = Array.from(
      { length: 120 },
      (_, i) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${p(`row${num(i + 1)}`)}</w:tc></w:tr>`,
    ).join("");
    const table =
      `<w:tbl><w:tblPr><w:tblpPr w:horzAnchor="margin" w:vertAnchor="text" w:tblpY="300"/>` +
      `<w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>` +
      `${rows}</w:tbl>`;
    const pages = domPages(p("Before") + table + p("After")).map((el) => el.textContent ?? "");
    expect(pages.length).toBeGreaterThanOrEqual(3);
    const all = pages.join("");
    for (let i = 1; i <= 120; i++) expect(all).toContain(`row${num(i)}`);
    expect(pages[pages.length - 1]).toContain("row120");
    expect(pages[pages.length - 1]).toContain("After");
  });
});

describe("DOM renderer: overlapping floating tables", () => {
  it("puts a picture in the first table above the second table", () => {
    const floating = (y: number, body: string) =>
      `<w:tbl><w:tblPr><w:tblpPr w:horzAnchor="page" w:vertAnchor="page" w:tblpX="1440" w:tblpY="${y}"/>` +
      `<w:tblW w:w="3000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${body}</w:tc></w:tr></w:tbl>`;
    const picture =
      `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
      `<wp:extent cx="762000" cy="762000"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
      `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rIdImg"/></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="762000" cy="762000"/></a:xfrm></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
    const body = p("Intro") + floating(1440, picture) + floating(1700, p("B")) + p("Tail");
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(body),
        "word/_rels/document.xml.rels":
          `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/x.png"/>` +
          `</Relationships>`,
        "word/media/x.png": "PNGDATA",
      }),
    );
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });
    const container = document.createElement("div");
    document.body.appendChild(container);
    // jsdom has no blob URLs; the renderer makes one for the picture's bytes.
    URL.createObjectURL = () => "blob:picture";
    renderToDom(doc, result, container);
    const img = container.querySelector<HTMLElement>(".dxw-page img");
    if (!img) throw new Error("missing picture");
    // Table B's opaque sheet has no z-index; the picture sits on a layer above it.
    expect(Number(img.style.zIndex)).toBeGreaterThan(0);
  });
});

describe("DOM renderer: vertical text in a merged cell", () => {
  it("lays the text along the merged height, not the first row", () => {
    for (const dir of ["btLr", "tbRl"]) {
      const table =
        `<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="600"/><w:gridCol w:w="1400"/></w:tblGrid>` +
        `<w:tr><w:trPr><w:trHeight w:hRule="exact" w:val="240"/></w:trPr>` +
        `<w:tc><w:tcPr><w:tcW w:w="600" w:type="dxa"/><w:vMerge w:val="restart"/><w:textDirection w:val="${dir}"/></w:tcPr>` +
        `${p("Pickup practices conducted")}</w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="1400" w:type="dxa"/></w:tcPr>${p("x")}</w:tc></w:tr>` +
        `<w:tr><w:trPr><w:trHeight w:val="4000"/></w:trPr>` +
        `<w:tc><w:tcPr><w:tcW w:w="600" w:type="dxa"/><w:vMerge/><w:textDirection w:val="${dir}"/></w:tcPr><w:p/></w:tc>` +
        `<w:tc><w:tcPr><w:tcW w:w="1400" w:type="dxa"/></w:tcPr>${p("y")}</w:tc></w:tr></w:tbl>`;
      const [page] = domPages(table + p(""));
      const rotated = [...page.querySelectorAll<HTMLElement>("*")].filter(
        (el) => el.style.transform.includes("rotate(") && el.textContent,
      );
      expect(rotated.map((el) => el.textContent).join("").replace(/\s/g, "")).toBe("Pickuppracticesconducted");
      // One rotated line holds the whole heading; the old layout cut it into
      // "Pi", "ck", "up", ... on lines of their own.
      expect(new Set(rotated.map((el) => el.style.top)).size).toBe(1);
    }
  });
});
