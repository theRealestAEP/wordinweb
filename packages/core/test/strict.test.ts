import { strFromU8, strToU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { compareDocuments } from "../src/edit/compare/index.js";
import { collectRevisions } from "../src/edit/suggest.js";
import { StrictRewriter, normalizeStrictPackage, normalizeStrictXml, transitionalUri } from "../src/strict.js";
import { parseXml, serializeXml } from "../src/xml.js";
import { AUTHOR, DATE } from "./compare-helpers.js";
import { STRICT_DOCUMENT, STRICT_RELS, TRANSITIONAL_DOCUMENT, TRANSITIONAL_RELS, strictPackage } from "./strict-helpers.js";

const W = 'xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"';

/** A Strict document part around `body`, normalized, as its body's XML. */
function normalizedBody(body: string): string {
  const out = normalizeStrictXml(`<w:document ${W}><w:body>${body}</w:body></w:document>`);
  return out.slice(out.indexOf("<w:body>") + "<w:body>".length, out.indexOf("</w:body>"));
}

describe("Strict Open XML measures", () => {
  it("converts each unit to twips", () => {
    expect(normalizedBody('<w:ind w:left="72pt" w:right="1in" w:hanging="2.54cm" w:firstLine="25.4mm"/>')).toBe('<w:ind w:left="1440" w:right="1440" w:hanging="1440" w:firstLine="1440"/>');
    expect(normalizedBody('<w:pgMar w:top="6pc" w:bottom="6pi" w:left="0.5pt" w:right="-5.40pt"/>')).toBe('<w:pgMar w:top="1440" w:bottom="1440" w:left="10" w:right="-108"/>');
  });

  it("converts half-point, eighth-point and point measures by their attribute", () => {
    expect(normalizedBody('<w:sz w:val="12pt"/><w:szCs w:val="0.5in"/><w:position w:val="-3pt"/><w:kern w:val="1pc"/>')).toBe('<w:sz w:val="24"/><w:szCs w:val="72"/><w:position w:val="-6"/><w:kern w:val="24"/>');
    expect(normalizedBody('<w:top w:val="single" w:sz="0.5pt" w:space="4pt"/><w:cols w:space="36pt"/>')).toBe('<w:top w:val="single" w:sz="4" w:space="4"/><w:cols w:space="720"/>');
  });

  it("converts percentages by their namespace", () => {
    expect(normalizedBody('<w:w w:val="150%"/><w:zoom w:percent="109%"/><w:tblW w:w="90.0%" w:type="pct"/><w:tcW w:w="33.33%" w:type="pct"/>')).toBe('<w:w w:val="150"/><w:zoom w:percent="109"/><w:tblW w:w="4500" w:type="pct"/><w:tcW w:w="1667" w:type="pct"/>');
    expect(normalizedBody('<a:lumMod val="75%"/><a:fillToRect t="-80%"/><a:normAutofit fontScale="62.5%"/><c:gapWidth val="150%"/>')).toBe('<a:lumMod val="75000"/><a:fillToRect t="-80000"/><a:normAutofit fontScale="62500"/><c:gapWidth val="150"/>');
  });

  it("converts DrawingML measures to EMUs", () => {
    expect(normalizedBody('<wp:extent cx="1in" cy="2.54cm"/><a:ln w="0.75pt"/>')).toBe('<wp:extent cx="914400" cy="914400"/><a:ln w="9525"/>');
  });

  it("renames Strict values and border sides", () => {
    expect(normalizedBody('<w:jc w:val="start"/><w:lvlJc w:val="end"/><w:tab w:val="end" w:pos="1in"/><w:textDirection w:val="rl"/><w:textDirection w:val="lrV"/>')).toBe(
      '<w:jc w:val="left"/><w:lvlJc w:val="right"/><w:tab w:val="right" w:pos="1440"/><w:textDirection w:val="tbRl"/><w:textDirection w:val="tbLrV"/>',
    );
    expect(normalizedBody('<w:pBdr><w:start w:val="single"/><w:end w:val="single"/></w:pBdr><w:tcBorders><w:start w:val="nil"></w:start></w:tcBorders><w:ind w:start="1pt" w:end="1pt"/>')).toBe(
      '<w:pBdr><w:left w:val="single"/><w:right w:val="single"/></w:pBdr><w:tcBorders><w:left w:val="nil"></w:left></w:tcBorders><w:ind w:start="20" w:end="20"/>',
    );
  });

  it("leaves text content alone and drops w:conformance", () => {
    const out = normalizeStrictXml(`<w:document ${W} w:conformance="strict"><w:body><w:p><w:r><w:t>72pt &amp; w:sz="1in"</w:t></w:r></w:p><!-- 1in --></w:body></w:document>`);
    expect(out).toBe('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>72pt &amp; w:sz="1in"</w:t></w:r></w:p><!-- 1in --></w:body></w:document>');
  });

  it("rewrites the sample document to what Word writes in Transitional", () => {
    expect(serializeXml(parseXml(normalizeStrictXml(STRICT_DOCUMENT)))).toBe(serializeXml(parseXml(TRANSITIONAL_DOCUMENT)));
    expect(normalizeStrictXml(STRICT_RELS)).toBe(TRANSITIONAL_RELS);
  });

  it("is the identity on a Transitional part", () => {
    expect(normalizeStrictXml(TRANSITIONAL_DOCUMENT)).toBe(TRANSITIONAL_DOCUMENT);
    // A Transitional part keeps its Transitional values, unit or not.
    const transitional = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tblW w:w="50%" w:type="pct"/></w:body></w:document>';
    expect(normalizeStrictXml(transitional)).toBe(transitional);
  });
});

describe("Strict Open XML namespaces", () => {
  it("maps each Strict URI to its Transitional one", () => {
    expect(transitionalUri("http://purl.oclc.org/ooxml/wordprocessingml/main")).toBe("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
    expect(transitionalUri("http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing")).toBe("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
    expect(transitionalUri("http://purl.oclc.org/ooxml/officeDocument/math")).toBe("http://schemas.openxmlformats.org/officeDocument/2006/math");
    expect(transitionalUri("http://purl.oclc.org/ooxml/officeDocument/relationships/image")).toBe("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
    expect(transitionalUri("http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties")).toBe("http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
    expect(transitionalUri("http://purl.oclc.org/ooxml/officeDocument/customProperties")).toBe("http://schemas.openxmlformats.org/officeDocument/2006/custom-properties");
    expect(transitionalUri("http://schemas.microsoft.com/office/word/2010/wordml")).toBe("http://schemas.microsoft.com/office/word/2010/wordml");
  });
});

describe("StrictRewriter", () => {
  const bytes = strToU8(STRICT_DOCUMENT);
  const expected = normalizeStrictXml(STRICT_DOCUMENT);

  it.each([1, 7, 64, 1000, bytes.length])("rewrites a part in pieces of %i bytes", (size) => {
    const rewriter = new StrictRewriter();
    const pieces: Uint8Array[] = [];
    for (let at = 0; at < bytes.length; at += size) pieces.push(rewriter.push(bytes.subarray(at, at + size)));
    pieces.push(rewriter.end());
    const out = new Uint8Array(pieces.reduce((n, piece) => n + piece.length, 0));
    let at = 0;
    for (const piece of pieces) {
      out.set(piece, at);
      at += piece.length;
    }
    expect(rewriter.strict).toBe(true);
    expect(strFromU8(out)).toBe(expected);
  });

  it("passes a Transitional part through untouched", () => {
    const rewriter = new StrictRewriter();
    const piece = strToU8(TRANSITIONAL_DOCUMENT);
    const out = rewriter.push(piece);
    expect(rewriter.strict).toBe(false);
    expect(out).toBe(piece);
    expect(rewriter.end().length).toBe(0);
  });
});

describe("a Strict package", () => {
  it("is normalized in place, a Transitional one is left as it is", () => {
    const strict = unzipSync(strictPackage(true));
    expect(normalizeStrictPackage(strict)).toBe(true);
    expect(strFromU8(strict["_rels/.rels"])).toBe(TRANSITIONAL_RELS);
    expect(strFromU8(strict["word/document.xml"])).not.toContain("purl.oclc.org");
    const transitional = unzipSync(strictPackage(false));
    const before = strFromU8(transitional["word/document.xml"]);
    expect(normalizeStrictPackage(transitional)).toBe(false);
    expect(strFromU8(transitional["word/document.xml"])).toBe(before);
  });

  it("loads as its Transitional twin", () => {
    const strict = DocxDocument.load(strictPackage(true));
    const transitional = DocxDocument.load(strictPackage(false));
    expect(strict.sections.length).toBe(1);
    expect(strict.sections[0].props).toEqual(transitional.sections[0].props);
    expect(strict.sections[0].blocks).toEqual(transitional.sections[0].blocks);
    expect(strict.docRoot.attrs["w:conformance"]).toBeUndefined();
    const saved = unzipSync(strict.save());
    for (const name in saved) if (/\.(xml|rels)$/.test(name)) expect(strFromU8(saved[name])).not.toContain("purl.oclc.org");
  });

  it("compares as its Transitional twin", () => {
    const signatures = (doc: DocxDocument): string[] => collectRevisions(doc).map((r) => `${r.kind}:${serializeXml(r.el).replace(/ w:id="\d+"/g, "")}`);
    const options = { author: AUTHOR, date: DATE };
    const transitional = compareDocuments(DocxDocument.load(strictPackage(false)), DocxDocument.load(strictPackage(false, "Hello there")), options);
    const strictOriginal = compareDocuments(DocxDocument.load(strictPackage(true)), DocxDocument.load(strictPackage(false, "Hello there")), options);
    const both = compareDocuments(DocxDocument.load(strictPackage(true)), DocxDocument.load(strictPackage(true, "Hello there")), options);
    expect(signatures(transitional).length).toBeGreaterThan(0);
    expect(signatures(strictOriginal)).toEqual(signatures(transitional));
    expect(signatures(both)).toEqual(signatures(transitional));
    const saved = unzipSync(both.save());
    expect(strFromU8(saved["word/document.xml"])).not.toContain("purl.oclc.org");
    expect(DocxDocument.load(both.save()).sections[0].props).toEqual(transitional.sections[0].props);
  });
});
