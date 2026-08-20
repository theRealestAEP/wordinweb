import { describe, expect, it } from "vitest";
import { parseXml, serializeXml } from "../src/xml.js";
import { pruneToPastedSubset, validatePastedOoxml } from "../src/ooxml-validate.js";

function blocks(xml: string) {
  // Wrap so parseXml has a single root, then take its children as the block list.
  const root = parseXml(`<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${xml}</root>`);
  return root.children;
}

describe("validatePastedOoxml (gate 2 / F3)", () => {
  it("accepts a plain formatted paragraph", () => {
    const r = validatePastedOoxml(blocks(`<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">hi</w:t></w:r></w:p>`));
    expect(r.ok).toBe(true);
  });

  it("rejects a hyperlink (authored URL until the scheme gate)", () => {
    const r = validatePastedOoxml(blocks(`<w:p><w:hyperlink r:id="rId5"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p>`));
    expect(r.ok).toBe(false);
    expect(r.reason).toMatch(/hyperlink/);
  });

  it("rejects field codes (HYPERLINK/INCLUDETEXT etc.)", () => {
    expect(validatePastedOoxml(blocks(`<w:p><w:r><w:instrText>HYPERLINK "http://x"</w:instrText></w:r></w:p>`)).ok).toBe(false);
    expect(validatePastedOoxml(blocks(`<w:p><w:fldSimple w:instr=" DATE "/></w:p>`)).ok).toBe(false);
  });

  it("rejects altChunk (embedded sub-documents)", () => {
    expect(validatePastedOoxml(blocks(`<w:altChunk r:id="rId9"/>`)).ok).toBe(false);
  });

  it("rejects drawings/OLE (media is out-of-band)", () => {
    expect(validatePastedOoxml(blocks(`<w:p><w:r><w:drawing/></w:r></w:p>`)).ok).toBe(false);
    expect(validatePastedOoxml(blocks(`<w:p><w:r><w:object/></w:r></w:p>`)).ok).toBe(false);
  });

  it("rejects any relationship reference attribute", () => {
    const r = validatePastedOoxml(blocks(`<w:p><w:r r:embed="rId2"><w:t>x</w:t></w:r></w:p>`));
    expect(r.ok).toBe(false);
    expect(r.reason).toMatch(/relationship/);
  });

  it("rejects a non-allowlisted element", () => {
    expect(validatePastedOoxml(blocks(`<w:p><w:customEvil/></w:p>`)).ok).toBe(false);
  });

  it("enforces depth and node caps", () => {
    const deep = "<w:p>".repeat(30) + "</w:p>".repeat(30);
    expect(validatePastedOoxml(blocks(deep), { maxNodes: 5000, maxDepth: 8 }).ok).toBe(false);
    const many = "<w:p>" + "<w:r><w:t>x</w:t></w:r>".repeat(20) + "</w:p>";
    expect(validatePastedOoxml(blocks(many), { maxNodes: 10, maxDepth: 24 }).ok).toBe(false);
  });
});

describe("pruneToPastedSubset keeps table cell formatting whole", () => {
  it("carries a cell margin override without spilling its sides into tcPr", () => {
    // Off the allowlist tcMar was UNWRAPPED, not dropped, which left
    // <w:top w:w="120"/> as a direct child of w:tcPr — an element that is not
    // a valid tcPr child in any schema.
    const [tbl] = pruneToPastedSubset(
      blocks(
        `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr>` +
          `<w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="240" w:type="dxa"/></w:tcMar>` +
          `<w:tcBorders><w:tl2br w:val="single" w:sz="8"/><w:top w:val="single" w:sz="8"/></w:tcBorders>` +
          `</w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      ),
    );
    const xml = serializeXml(tbl);
    expect(xml).toContain(`<w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="240" w:type="dxa"/></w:tcMar>`);
    expect(xml).toContain(`<w:tl2br w:val="single" w:sz="8"/>`);
  });
});
