import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { insertCoverPage } from "../src/edit/sections.js";
import {
  insertHeaderFooterPreset,
  insertPageNumberPosition,
  removePageNumberFields,
} from "../src/edit/hf-gallery.js";
import { XmlElement, attr, localName, serializeXml } from "../src/xml.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";

const SECT = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>`;

function loadDoc(body: string, extra: Record<string, string> = {}) {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body), ...extra }));
}

function fldKeywords(root: XmlElement): string[] {
  const out: string[] = [];
  const walk = (e: XmlElement): void => {
    if (localName(e.name) === "fldSimple") {
      out.push((attr(e, "instr") ?? "").trim().split(/\s+/)[0]?.toUpperCase());
    }
    for (const c of e.children) walk(c);
  };
  walk(root);
  return out;
}

describe("insertPageNumberPosition (page-number position gallery)", () => {
  it("creates the header part and inserts a centered PAGE field", () => {
    const doc = loadDoc(p("Body text") + SECT);
    expect(insertPageNumberPosition(doc, "top", "center")).toBe(true);
    const headers = doc.headerRoots();
    expect(headers.length).toBe(1);
    expect(fldKeywords(headers[0])).toEqual(["PAGE"]);
    expect(serializeXml(headers[0])).toContain(`w:val="center"`);
  });

  it("targets the footer for 'bottom' and honors alignment", () => {
    const doc = loadDoc(p("Body text") + SECT);
    expect(insertPageNumberPosition(doc, "bottom", "right")).toBe(true);
    expect(doc.headerRoots().length).toBe(0);
    const footers = doc.footerRoots();
    expect(footers.length).toBe(1);
    expect(serializeXml(footers[0])).toContain(`w:val="right"`);
  });

  it("a second pick REPLACES the part's content, matching Word's gallery", () => {
    const doc = loadDoc(p("Body text") + SECT);
    insertPageNumberPosition(doc, "top", "left");
    insertPageNumberPosition(doc, "top", "right");
    const headers = doc.headerRoots();
    expect(headers.length).toBe(1);
    expect(headers[0].children.length).toBe(1);
    expect(serializeXml(headers[0])).toContain(`w:val="right"`);
    expect(serializeXml(headers[0])).not.toContain(`w:val="left"`);
  });
});

describe("removePageNumberFields (Remove Page Numbers)", () => {
  it("removes a fldSimple PAGE field the gallery inserted, from header and footer", () => {
    const doc = loadDoc(p("Body text") + SECT);
    insertPageNumberPosition(doc, "top", "center");
    insertPageNumberPosition(doc, "bottom", "center");
    expect(removePageNumberFields(doc)).toBe(true);
    expect(fldKeywords(doc.headerRoots()[0])).toEqual([]);
    expect(fldKeywords(doc.footerRoots()[0])).toEqual([]);
    // The band itself stays, just empty of page-number content.
    expect(doc.headerRoots()[0].children.some((c) => localName(c.name) === "p")).toBe(true);
  });

  it("removes a complex-field PAGE span (fldChar begin/instrText/separate/end) and its 'Page X of Y' literals", () => {
    const doc = loadDoc(p("Body text") + SECT);
    const footer = doc.ensureHfPart("footer");
    footer.children = [{
      name: "w:p",
      attrs: {},
      text: "",
      children: [
        { name: "w:r", attrs: {}, text: "", children: [{ name: "w:t", attrs: { "xml:space": "preserve" }, children: [], text: "Page " }] },
        { name: "w:r", attrs: {}, text: "", children: [{ name: "w:fldChar", attrs: { "w:fldCharType": "begin" }, children: [], text: "" }] },
        { name: "w:r", attrs: {}, text: "", children: [{ name: "w:instrText", attrs: {}, children: [], text: " PAGE " }] },
        { name: "w:r", attrs: {}, text: "", children: [{ name: "w:fldChar", attrs: { "w:fldCharType": "separate" }, children: [], text: "" }] },
        { name: "w:r", attrs: {}, text: "", children: [{ name: "w:t", attrs: {}, children: [], text: "1" }] },
        { name: "w:r", attrs: {}, text: "", children: [{ name: "w:fldChar", attrs: { "w:fldCharType": "end" }, children: [], text: "" }] },
        { name: "w:r", attrs: {}, text: "", children: [{ name: "w:t", attrs: { "xml:space": "preserve" }, children: [], text: " of " }] },
        { name: "w:r", attrs: {}, text: "", children: [{ name: "w:fldChar", attrs: { "w:fldCharType": "begin" }, children: [], text: "" }] },
        { name: "w:r", attrs: {}, text: "", children: [{ name: "w:instrText", attrs: {}, children: [], text: " NUMPAGES " }] },
        { name: "w:r", attrs: {}, text: "", children: [{ name: "w:fldChar", attrs: { "w:fldCharType": "separate" }, children: [], text: "" }] },
        { name: "w:r", attrs: {}, text: "", children: [{ name: "w:t", attrs: {}, children: [], text: "3" }] },
        { name: "w:r", attrs: {}, text: "", children: [{ name: "w:fldChar", attrs: { "w:fldCharType": "end" }, children: [], text: "" }] },
      ],
    }];
    doc.refresh();
    expect(removePageNumberFields(doc)).toBe(true);
    // Every field and literal is gone; refresh's own caret-anchor repair
    // (parse/document.ts ensureCaretAnchor) may leave a single empty,
    // omitWhenEmpty run behind — invisible on save, so serializeXml (which
    // honors omitWhenEmpty) is the honest check here, not raw children.length.
    expect(fldKeywords(footer)).toEqual([]);
    expect(serializeXml(footer)).toBe(
      `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p></w:p></w:ftr>`,
    );
  });

  it("false when no header/footer PAGE fields exist", () => {
    const doc = loadDoc(p("Body text") + SECT);
    expect(removePageNumberFields(doc)).toBe(false);
    doc.ensureHfPart("header"); // an ordinary header with no page number
    expect(removePageNumberFields(doc)).toBe(false);
  });
});

describe("insertHeaderFooterPreset (Header & Footer preset gallery)", () => {
  it("blank leaves a single empty paragraph", () => {
    const doc = loadDoc(p("Body text") + SECT);
    expect(insertHeaderFooterPreset(doc, "header", "blank")).toBe(true);
    const header = doc.headerRoots()[0];
    expect(header.children.length).toBe(1);
    expect(localName(header.children[0].name)).toBe("p");
  });

  it("centeredTitle writes one centered placeholder paragraph", () => {
    const doc = loadDoc(p("Body text") + SECT);
    expect(insertHeaderFooterPreset(doc, "header", "centeredTitle")).toBe(true);
    const xml = serializeXml(doc.headerRoots()[0]);
    expect(xml).toContain("[Document Title]");
    expect(xml).toContain(`w:val="center"`);
  });

  it("titleAndDate writes a title paragraph and a live DATE field", () => {
    const doc = loadDoc(p("Body text") + SECT);
    expect(insertHeaderFooterPreset(doc, "footer", "titleAndDate")).toBe(true);
    const footer = doc.footerRoots()[0];
    expect(footer.children.filter((c) => localName(c.name) === "p").length).toBe(2);
    expect(fldKeywords(footer)).toEqual(["DATE"]);
  });

  it("threeColumn sets paragraph tab stops via setTabStops and ends with a live PAGE field", () => {
    const doc = loadDoc(p("Body text") + SECT);
    expect(insertHeaderFooterPreset(doc, "footer", "threeColumn")).toBe(true);
    const footer = doc.footerRoots()[0];
    const para = footer.children.find((c) => localName(c.name) === "p")!;
    const pPr = para.children.find((c) => localName(c.name) === "pPr");
    const tabs = pPr?.children.find((c) => localName(c.name) === "tabs");
    expect(tabs?.children.length).toBe(2);
    expect(fldKeywords(footer)).toEqual(["PAGE"]);
  });

  it("a second pick replaces the part's content", () => {
    const doc = loadDoc(p("Body text") + SECT);
    insertHeaderFooterPreset(doc, "header", "titleAndDate");
    insertHeaderFooterPreset(doc, "header", "blank");
    const header = doc.headerRoots()[0];
    expect(header.children.length).toBe(1);
    expect(serializeXml(header)).not.toContain("[Document Title]");
  });
});

describe("insertCoverPage layouts", () => {
  it("defaults to the original centered 'title' layout", () => {
    const doc = loadDoc(p("Body text") + SECT);
    expect(insertCoverPage(doc, { title: "Report" })).toBe(true);
    const xml = serializeXml(doc.docRoot);
    expect(xml).toContain(`w:val="center"`);
    expect(xml).not.toContain("w:shd");
    expect(xml).not.toContain("w:pBdr");
  });

  it("'banner' shades the title paragraph", () => {
    const doc = loadDoc(p("Body text") + SECT);
    expect(insertCoverPage(doc, { title: "Report", layout: "banner" })).toBe(true);
    const xml = serializeXml(doc.docRoot);
    expect(xml).toContain("w:shd");
    expect(xml).toContain("2F5597");
  });

  it("'sidebar' left-aligns the block and adds a left accent border", () => {
    const doc = loadDoc(p("Body text") + SECT);
    expect(insertCoverPage(doc, { title: "Report", subtitle: "Q1", layout: "sidebar" })).toBe(true);
    const xml = serializeXml(doc.docRoot);
    expect(xml).toContain("w:pBdr");
    expect(xml).toContain(`w:val="left"`);
    expect(xml).not.toContain(`w:val="center"`);
  });
});
