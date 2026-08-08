import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { DocxDocument } from "../src/docx.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function load(settingsXml?: string): DocxDocument {
  return DocxDocument.load(
    makeDocx({
      "word/document.xml": wrapDocument(p("hello")),
      ...(settingsXml !== undefined ? { "word/settings.xml": settingsXml } : {}),
      ...(settingsXml !== undefined
        ? {
            "word/_rels/document.xml.rels":
              `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
              `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`,
          }
        : {}),
    }),
  );
}

function savedSettings(doc: DocxDocument): string {
  const files = unzipSync(doc.save());
  return files["word/settings.xml"] ? strFromU8(files["word/settings.xml"]) : "";
}

describe("hyphenation settings (w:autoHyphenation cluster)", () => {
  it("parses the cluster from settings.xml", () => {
    const doc = load(
      `<?xml version="1.0"?><w:settings ${W}>` +
        `<w:autoHyphenation/><w:hyphenationZone w:val="360"/><w:doNotHyphenateCaps/>` +
        `</w:settings>`,
    );
    expect(doc.autoHyphenation).toBe(true);
    expect(doc.hyphenationZoneTwips).toBe(360);
    expect(doc.doNotHyphenateCaps).toBe(true);
  });

  it("writes the cluster in schema order, after defaultTabStop", () => {
    const doc = load(
      `<?xml version="1.0"?><w:settings ${W}>` +
        `<w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:defaultTableStyle w:val="X"/>` +
        `</w:settings>`,
    );
    expect(doc.setHyphenation({ auto: true, zoneTwips: 360, noCaps: true })).toBe(true);
    const xml = savedSettings(doc);
    const order = ["w:zoom", "w:defaultTabStop", "w:autoHyphenation", "w:hyphenationZone", "w:doNotHyphenateCaps", "w:defaultTableStyle"];
    const positions = order.map((name) => xml.indexOf(`<${name}`));
    expect(positions.every((n) => n >= 0)).toBe(true);
    expect([...positions].sort((a, b) => a - b)).toEqual(positions);
    expect(xml).toContain('<w:hyphenationZone w:val="360"/>');
  });

  it("round-trips through parse, clears, and reports the honest no-op", () => {
    const doc = load();
    expect(doc.setHyphenation({})).toBe(false);
    expect(doc.setHyphenation({ auto: true, zoneTwips: 425 })).toBe(true);
    doc.refresh();
    expect(doc.autoHyphenation).toBe(true);
    expect(doc.hyphenationZoneTwips).toBe(425);
    // Same values again: nothing changes.
    expect(doc.setHyphenation({ auto: true, zoneTwips: 425 })).toBe(false);
    // Clearing removes the elements.
    expect(doc.setHyphenation({ auto: false, zoneTwips: null })).toBe(true);
    const xml = savedSettings(doc);
    expect(xml).not.toContain("autoHyphenation");
    expect(xml).not.toContain("hyphenationZone");
  });

  it("creates and registers settings.xml when the package was born without it", () => {
    const doc = load();
    expect(doc.setHyphenation({ auto: true })).toBe(true);
    const files = unzipSync(doc.save());
    expect(strFromU8(files["word/settings.xml"])).toContain("<w:autoHyphenation/>");
    expect(strFromU8(files["[Content_Types].xml"])).toContain("settings+xml");
  });
});
