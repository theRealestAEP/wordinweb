import { describe, expect, it } from "vitest";
import { unzipSync } from "fflate";
import { DocxDocument } from "../src/index.js";
import { makeDocx, wrapDocument } from "./helpers.js";

/**
 * A save must never MATERIALIZE the compatibility mode it merely assumed.
 *
 * `compatibilityMode` defaults to 15 so the layout has a number to branch on,
 * but `declaredCompatibilityMode` stays undefined unless settings.xml really
 * carries `<w:compatSetting w:name="compatibilityMode">` (docx.ts, the parse
 * at "for (const cs of children(compat, ...))"). That distinction is now
 * load-bearing: the percentage-table margin allowance keys on the DECLARED
 * value, so a save that wrote the defaulted 15 into settings.xml would flip a
 * legacy document's table widths on the next load — silently, and only after
 * a round trip.
 *
 * Two independent things stop it, and this pins both. settings.xml is written
 * only when `settingsDirty` is set, which one method does (setMirrorMargins);
 * and even then it re-serializes the PARSED settings tree, which no code path
 * ever adds a compatSetting to.
 */

const LEGACY_SETTINGS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:defaultTabStop w:val="720"/></w:settings>`;

const CONTENT_TYPES = `<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`;

/** A legacy package: settings.xml present, but declaring NO compatibilityMode. */
function legacyDocx(): Uint8Array {
  return makeDocx({
    "[Content_Types].xml": CONTENT_TYPES,
    "word/settings.xml": LEGACY_SETTINGS,
    "word/document.xml": wrapDocument(
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid>` +
        `<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    ),
  });
}

function settingsBytes(docx: Uint8Array): Uint8Array {
  const part = unzipSync(docx)["word/settings.xml"];
  if (!part) throw new Error("no word/settings.xml");
  return part;
}

describe("compatibilityMode survives a save round trip undeclared", () => {
  it("loads a legacy document with no declared mode", () => {
    const doc = DocxDocument.load(legacyDocx());
    expect(doc.declaredCompatibilityMode).toBeUndefined();
    // The layout still gets a number; only the DECLARED one is absent.
    expect(doc.compatibilityMode).toBe(15);
  });

  it("keeps it undeclared after save and reload", () => {
    const source = legacyDocx();
    const reloaded = DocxDocument.load(DocxDocument.load(source).save());
    expect(reloaded.declaredCompatibilityMode).toBeUndefined();
  });

  it("writes settings.xml back byte for byte", () => {
    const source = legacyDocx();
    const saved = DocxDocument.load(source).save();
    expect(Buffer.from(settingsBytes(saved)).equals(Buffer.from(settingsBytes(source)))).toBe(true);
  });

  it("keeps it undeclared even when an edit dirties settings.xml", () => {
    // setMirrorMargins is the one path that re-serializes settings.xml, so it
    // is the one that could write a synthesized compatSetting.
    const doc = DocxDocument.load(legacyDocx());
    doc.setMirrorMargins(true);
    const reloaded = DocxDocument.load(doc.save());
    expect(reloaded.declaredCompatibilityMode).toBeUndefined();
    expect(reloaded.mirrorMargins).toBe(true);
    expect(new TextDecoder().decode(settingsBytes(doc.save()))).not.toContain("compatibilityMode");
  });

  it("still reads a mode the document really declares", () => {
    const declared = makeDocx({
      "[Content_Types].xml": CONTENT_TYPES,
      "word/settings.xml":
        `<?xml version="1.0"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
        `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>` +
        `</w:settings>`,
      "word/document.xml": wrapDocument(`<w:p><w:r><w:t>legacy</w:t></w:r></w:p>`),
    });
    const reloaded = DocxDocument.load(DocxDocument.load(declared).save());
    expect(reloaded.declaredCompatibilityMode).toBe(14);
    expect(reloaded.compatibilityMode).toBe(14);
  });
});
