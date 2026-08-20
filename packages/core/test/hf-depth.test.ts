import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { DocxDocument } from "../src/docx.js";
import {
  endnoteOptionsAt,
  footnoteOptionsAt,
  pageNumberFormatAt,
  sectPrAt,
  setEndnoteOptions,
  setEvenOddHeaders,
  setFootnoteOptions,
  setPageNumberFormat,
  setTitlePage,
  titlePageEnabled,
} from "../src/edit/sections.js";
import { XmlElement, attr, localName } from "../src/xml.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";

const SECT = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr>`;

function loadDoc(body: string, extra: Record<string, string> = {}) {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body), ...extra }));
}

function bodySectPr(doc: DocxDocument): XmlElement {
  const found: XmlElement[] = [];
  const walk = (e: XmlElement): void => {
    if (localName(e.name) === "sectPr") found.push(e);
    else for (const c of e.children) walk(c);
  };
  walk(doc.docRoot);
  expect(found.length).toBeGreaterThan(0);
  return found[found.length - 1];
}

function refsOf(sectPr: XmlElement, refLocal: string): { type: string | undefined; id: string | undefined }[] {
  return sectPr.children
    .filter((c) => localName(c.name) === refLocal)
    .map((c) => ({ type: attr(c, "type"), id: attr(c, "id") }));
}

describe("setTitlePage (different first page)", () => {
  it("adds w:titlePg and creates empty first-page header and footer parts", () => {
    const doc = loadDoc(p("Body text") + SECT);
    expect(titlePageEnabled(doc)).toBe(false);
    expect(setTitlePage(doc, true)).toBe(true);
    expect(titlePageEnabled(doc)).toBe(true);
    expect(doc.sections[0].props.titlePage).toBe(true);

    const sectPr = bodySectPr(doc);
    // Schema order: titlePg comes after pgMar/cols, never before the refs.
    expect(sectPr.children.some((c) => localName(c.name) === "titlePg")).toBe(true);
    const headerRefs = refsOf(sectPr, "headerReference");
    const footerRefs = refsOf(sectPr, "footerReference");
    expect(headerRefs.some((r) => r.type === "first")).toBe(true);
    expect(footerRefs.some((r) => r.type === "first")).toBe(true);
    // Layout resolves the first-page variant.
    expect(doc.sections[0].props.headerRefs.first).toBeTruthy();
    expect(doc.sections[0].props.footerRefs.first).toBeTruthy();
    expect(doc.headers.get(doc.sections[0].props.headerRefs.first!)).toBeTruthy();

    // The parts survive save with content-type overrides.
    const files = unzipSync(doc.save());
    const names = Object.keys(files);
    expect(names).toContain("word/header1.xml");
    expect(names).toContain("word/footer1.xml");
    const types = strFromU8(files["[Content_Types].xml"]);
    expect(types).toContain("/word/header1.xml");
    expect(types).toContain("/word/footer1.xml");
  });

  it("is an honest no-op when already enabled, and disable keeps the parts", () => {
    const doc = loadDoc(p("Body text") + SECT);
    expect(setTitlePage(doc, true)).toBe(true);
    expect(setTitlePage(doc, true)).toBe(false);
    expect(setTitlePage(doc, false)).toBe(true);
    expect(setTitlePage(doc, false)).toBe(false);
    expect(titlePageEnabled(doc)).toBe(false);
    // Word keeps the parts and references so re-enable restores content.
    const sectPr = bodySectPr(doc);
    expect(refsOf(sectPr, "headerReference").some((r) => r.type === "first")).toBe(true);
    expect(setTitlePage(doc, true)).toBe(true);
    // Re-enable creates nothing new: still exactly one first ref per band.
    expect(refsOf(sectPr, "headerReference").filter((r) => r.type === "first")).toHaveLength(1);
  });

  it("keeps an existing first-page part and its content", () => {
    const header = `<?xml version="1.0"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>First page header</w:t></w:r></w:p></w:hdr>`;
    const rels = `<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`;
    const body = p("Body text") +
      `<w:sectPr><w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="first" r:id="rId5"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
    const doc = loadDoc(body, { "word/header1.xml": header, "word/_rels/document.xml.rels": rels });
    expect(setTitlePage(doc, true)).toBe(true);
    const sectPr = bodySectPr(doc);
    // The existing first header ref is untouched; only the footer was missing.
    expect(refsOf(sectPr, "headerReference").filter((r) => r.type === "first")).toHaveLength(1);
    expect(refsOf(sectPr, "footerReference").filter((r) => r.type === "first")).toHaveLength(1);
    const hf = doc.headers.get("rId5");
    expect(JSON.stringify(hf)).toContain("First page header");
  });

  it("materializes a sectPr on a minimal document", () => {
    const doc = loadDoc(p("Just text"));
    expect(setTitlePage(doc, true)).toBe(true);
    expect(titlePageEnabled(doc)).toBe(true);
    expect(doc.sections[0].props.headerRefs.first).toBeTruthy();
  });
});

describe("setEvenOddHeaders (different odd & even pages)", () => {
  it("writes w:evenAndOddHeaders and creates the even-page parts", () => {
    const doc = loadDoc(p("Body text") + SECT);
    expect(doc.evenAndOddHeaders).toBe(false);
    expect(setEvenOddHeaders(doc, true)).toBe(true);
    expect(doc.evenAndOddHeaders).toBe(true);
    const sectPr = bodySectPr(doc);
    expect(refsOf(sectPr, "headerReference").some((r) => r.type === "even")).toBe(true);
    expect(refsOf(sectPr, "footerReference").some((r) => r.type === "even")).toBe(true);
    expect(doc.sections[0].props.headerRefs.even).toBeTruthy();
    const files = unzipSync(doc.save());
    expect(strFromU8(files["word/settings.xml"])).toContain("evenAndOddHeaders");
  });

  it("honest no-ops, and disable removes only the settings switch", () => {
    const doc = loadDoc(p("Body text") + SECT);
    expect(setEvenOddHeaders(doc, true)).toBe(true);
    expect(setEvenOddHeaders(doc, true)).toBe(false);
    expect(setEvenOddHeaders(doc, false)).toBe(true);
    expect(setEvenOddHeaders(doc, false)).toBe(false);
    expect(doc.evenAndOddHeaders).toBe(false);
    // Parts and refs stay for re-enable.
    expect(refsOf(bodySectPr(doc), "headerReference").some((r) => r.type === "even")).toBe(true);
  });
});

describe("setPageNumberFormat (w:pgNumType)", () => {
  it("writes fmt and start, readable via pageNumberFormatAt and the section props", () => {
    const doc = loadDoc(p("Body text") + SECT);
    const t = firstT(doc);
    expect(setPageNumberFormat(doc, { fmt: "lowerRoman", start: 1 })).toBe(true);
    expect(pageNumberFormatAt(doc, t)).toEqual({ fmt: "lowerRoman", start: 1 });
    expect(doc.sections[0].props.pageNumberFormat).toBe("lowerRoman");
    expect(doc.sections[0].props.pageNumberStart).toBe(1);
    // Schema-ordered: pgNumType sits after lnNumType/pgBorders, before cols.
    const sectPr = sectPrAt(doc, t)!;
    const names = sectPr.children.map((c) => localName(c.name));
    expect(names.indexOf("pgNumType")).toBeGreaterThan(names.indexOf("pgMar"));
  });

  it("null clears one attribute; an empty pgNumType is removed; unchanged patches reject", () => {
    const doc = loadDoc(p("Body text") + SECT);
    const t = firstT(doc);
    expect(setPageNumberFormat(doc, { fmt: "upperRoman", start: 5 })).toBe(true);
    expect(setPageNumberFormat(doc, { fmt: "upperRoman" })).toBe(false); // unchanged
    expect(setPageNumberFormat(doc, { start: null })).toBe(true);
    expect(pageNumberFormatAt(doc, t)).toEqual({ fmt: "upperRoman", start: null });
    // fmt "decimal" is the default and clears the attribute — pgNumType goes away.
    expect(setPageNumberFormat(doc, { fmt: "decimal" })).toBe(true);
    const sectPr = sectPrAt(doc, t)!;
    expect(sectPr.children.some((c) => localName(c.name) === "pgNumType")).toBe(false);
    expect(setPageNumberFormat(doc, { fmt: null })).toBe(false); // nothing to clear
  });
});

describe("setFootnoteOptions (w:footnotePr)", () => {
  it("writes pos/numFmt/numStart/numRestart in schema order, readable via footnoteOptionsAt", () => {
    const doc = loadDoc(p("Body text") + SECT);
    const t = firstT(doc);
    expect(setFootnoteOptions(doc, { fmt: "upperRoman", start: 3, restart: "eachSect", pos: "beneathText" })).toBe(true);
    expect(footnoteOptionsAt(doc, t)).toEqual({ fmt: "upperRoman", start: 3, restart: "eachSect", pos: "beneathText" });
    expect(doc.sections[0].props.footnoteNumFmt).toBe("upperRoman");
    expect(doc.sections[0].props.footnoteNumStart).toBe(3);
    expect(doc.sections[0].props.footnoteNumRestart).toBe("eachSect");
    expect(doc.sections[0].props.footnotePos).toBe("beneathText");
    const sectPr = sectPrAt(doc, t)!;
    const pr = sectPr.children.find((c) => localName(c.name) === "footnotePr")!;
    expect(pr).toBeTruthy();
    const names = pr.children.map((c) => localName(c.name));
    // CT_FtnProps order: pos, numFmt, numStart, numRestart.
    expect(names).toEqual(["pos", "numFmt", "numStart", "numRestart"]);
    // footnotePr itself sits after headerReference/footerReference, before endnotePr/type/pgMar.
    const sectNames = sectPr.children.map((c) => localName(c.name));
    expect(sectNames.indexOf("footnotePr")).toBeLessThan(sectNames.indexOf("pgMar"));
  });

  it("defaults are decimal/continuous/pageBottom; null clears an attribute back to default", () => {
    const doc = loadDoc(p("Body text") + SECT);
    const t = firstT(doc);
    expect(footnoteOptionsAt(doc, t)).toEqual({ fmt: "decimal", start: null, restart: "continuous", pos: "pageBottom" });
    expect(setFootnoteOptions(doc, { fmt: "lowerLetter", pos: "beneathText" })).toBe(true);
    expect(setFootnoteOptions(doc, { fmt: "lowerLetter" })).toBe(false); // unchanged
    expect(setFootnoteOptions(doc, { pos: null })).toBe(true);
    expect(footnoteOptionsAt(doc, t)).toEqual({ fmt: "lowerLetter", start: null, restart: "continuous", pos: "pageBottom" });
    // fmt back to its default clears the child; footnotePr goes away entirely once empty.
    expect(setFootnoteOptions(doc, { fmt: "decimal" })).toBe(true);
    const sectPr = sectPrAt(doc, t)!;
    expect(sectPr.children.some((c) => localName(c.name) === "footnotePr")).toBe(false);
    expect(setFootnoteOptions(doc, { fmt: null })).toBe(false); // nothing to clear
  });

  it("accepts chicago (Word's symbol mark style)", () => {
    const doc = loadDoc(p("Body text") + SECT);
    const t = firstT(doc);
    expect(setFootnoteOptions(doc, { fmt: "chicago" })).toBe(true);
    expect(footnoteOptionsAt(doc, t).fmt).toBe("chicago");
  });
});

describe("setEndnoteOptions (w:endnotePr)", () => {
  it("writes pos/numFmt/numStart/numRestart, readable via endnoteOptionsAt", () => {
    const doc = loadDoc(p("Body text") + SECT);
    const t = firstT(doc);
    expect(setEndnoteOptions(doc, { fmt: "lowerLetter", start: 2, restart: "eachSect", pos: "sectEnd" })).toBe(true);
    expect(endnoteOptionsAt(doc, t)).toEqual({ fmt: "lowerLetter", start: 2, restart: "eachSect", pos: "sectEnd" });
    expect(doc.sections[0].props.endnoteNumFmt).toBe("lowerLetter");
    expect(doc.sections[0].props.endnoteNumRestart).toBe("eachSect");
    expect(doc.sections[0].props.endnotePos).toBe("sectEnd");
  });

  it("endnotes default to lowerRoman/continuous/docEnd, distinct from footnotes' defaults", () => {
    const doc = loadDoc(p("Body text") + SECT);
    const t = firstT(doc);
    expect(endnoteOptionsAt(doc, t)).toEqual({ fmt: "lowerRoman", start: null, restart: "continuous", pos: "docEnd" });
    // Setting the endnote default explicitly is a no-op write (matches the default, so nothing to add).
    expect(setEndnoteOptions(doc, { fmt: "lowerRoman", pos: "docEnd" })).toBe(false);
  });
});

function firstT(doc: DocxDocument): XmlElement {
  const find = (e: XmlElement): XmlElement | null => {
    if (localName(e.name) === "t") return e;
    for (const c of e.children) {
      const f = find(c);
      if (f) return f;
    }
    return null;
  };
  const t = find(doc.docRoot);
  expect(t).toBeTruthy();
  return t!;
}
