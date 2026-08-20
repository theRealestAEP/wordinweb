import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { DocxDocument } from "../src/docx.js";
import { encodeClipboardOoxml } from "../src/edit/clipboard.js";
import {
  buildingBlockNodeCount,
  createBuildingBlock,
  deleteBuildingBlock,
  insertBuildingBlock,
  listBuildingBlocks,
} from "../src/edit/quick-parts.js";
import { makeDocx, wrapDocument, W_NS, p } from "./helpers.js";

/**
 * Quick Parts / Building Blocks: writing the glossary part (word/glossary/
 * document.xml, ECMA-376 §17.12) — creating it when the package has none,
 * reading one that arrived from Word, and the byte-stability discipline: only
 * the glossary part is ever written, and an untouched part keeps its original
 * bytes through save().
 */

function blocksXmlOf(...paragraphsXml: string[]): string {
  const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(paragraphsXml.join("")) }));
  // Reuse the clipboard OOXML encoder exactly as the toolbar's "Save
  // Selection as Quick Part" command will: the paragraphs are already this
  // engine's own retained XML, so no selection walk is needed for the test.
  const body = doc.docRoot.children.find((c) => c.name === "w:body")!;
  return encodeClipboardOoxml(body.children.filter((c) => c.name === "w:p" || c.name === "w:tbl"));
}

function blankDoc(): DocxDocument {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(p("anchor")) }));
}

function savedParts(doc: DocxDocument): Record<string, string> {
  const files = unzipSync(doc.save());
  const out: Record<string, string> = {};
  for (const [name, bytes] of Object.entries(files)) out[name] = strFromU8(bytes);
  return out;
}

const GLOSSARY_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:glossaryDocument ${W_NS}><w:docParts>` +
  `<w:docPart><w:docPartPr><w:name w:val="Signature Block"/>` +
  `<w:category><w:name w:val="Legal"/><w:gallery w:val="docParts"/></w:category>` +
  `</w:docPartPr><w:docPartBody>` +
  `<w:p><w:r><w:t xml:space="preserve">Best regards,</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t xml:space="preserve">Jane Doe</w:t></w:r></w:p>` +
  `</w:docPartBody></w:docPart>` +
  `</w:docParts></w:glossaryDocument>`;

const DOC_RELS_WITH_GLOSSARY =
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" Target="glossary/document.xml"/>` +
  `</Relationships>`;

function docWithGlossary(bodyXml = p("anchor")): DocxDocument {
  return DocxDocument.load(
    makeDocx({
      "word/document.xml": wrapDocument(bodyXml),
      "word/_rels/document.xml.rels": DOC_RELS_WITH_GLOSSARY,
      "word/glossary/document.xml": GLOSSARY_XML,
    }),
  );
}

describe("glossary part creation (blank document)", () => {
  it("creates the part, the relationship, and the content-type override the way Word lays them out", () => {
    const doc = blankDoc();
    expect(listBuildingBlocks(doc)).toEqual([]);
    expect(
      createBuildingBlock(doc, { name: "Greeting", category: "General", blocksXml: blocksXmlOf(p("Hello there")) }),
    ).toBe(true);

    const parts = savedParts(doc);
    expect(parts["word/glossary/document.xml"]).toContain("w:glossaryDocument");
    expect(parts["word/glossary/document.xml"]).toContain(`<w:name w:val="Greeting"/>`);
    expect(parts["word/glossary/document.xml"]).toContain(`<w:name w:val="General"/><w:gallery w:val="docParts"/>`);
    expect(parts["word/glossary/document.xml"]).toContain("Hello there");
    expect(parts["word/_rels/document.xml.rels"]).toContain(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument",
    );
    expect(parts["word/_rels/document.xml.rels"]).toContain(`Target="glossary/document.xml"`);
    expect(parts["[Content_Types].xml"]).toContain(
      `PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"`,
    );

    // Immediately listable and reloadable.
    expect(listBuildingBlocks(doc)).toEqual([{ name: "Greeting", category: "General" }]);
    const reloaded = DocxDocument.load(doc.save());
    expect(listBuildingBlocks(reloaded)).toEqual([{ name: "Greeting", category: "General" }]);
  });

  it("defaults an empty category to General, Word's own default", () => {
    const doc = blankDoc();
    expect(createBuildingBlock(doc, { name: "X", blocksXml: blocksXmlOf(p("x")) })).toBe(true);
    expect(listBuildingBlocks(doc)).toEqual([{ name: "X", category: "General" }]);
  });

  it("rejects a duplicate name without touching the part", () => {
    const doc = blankDoc();
    expect(createBuildingBlock(doc, { name: "T1", blocksXml: blocksXmlOf(p("one")) })).toBe(true);
    const before = savedParts(doc)["word/glossary/document.xml"];
    expect(createBuildingBlock(doc, { name: "T1", blocksXml: blocksXmlOf(p("two")) })).toBe(false);
    expect(savedParts(doc)["word/glossary/document.xml"]).toBe(before);
  });

  it("rejects malformed specs and unusable OOXML fragments", () => {
    expect(createBuildingBlock(blankDoc(), { name: "", blocksXml: blocksXmlOf(p("x")) })).toBe(false);
    expect(createBuildingBlock(blankDoc(), { name: "x".repeat(65), blocksXml: blocksXmlOf(p("x")) })).toBe(false);
    expect(createBuildingBlock(blankDoc(), { name: "Ok", blocksXml: "" })).toBe(false);
    expect(createBuildingBlock(blankDoc(), { name: "Ok", blocksXml: "<w:document/>" })).toBe(false);
    expect(
      createBuildingBlock(blankDoc(), {
        name: "Ok",
        blocksXml: `<w:document ${W_NS}><w:body><w:p><w:fldSimple w:instr=" DATE "><w:r><w:t>x</w:t></w:r></w:fldSimple></w:p></w:body></w:document>`,
      }),
    ).toBe(false);
  });
});

describe("an arriving Word glossary part", () => {
  it("keeps original bytes through save() while untouched", () => {
    const doc = docWithGlossary();
    expect(savedParts(doc)["word/glossary/document.xml"]).toBe(GLOSSARY_XML);
  });

  it("lists the arriving building block", () => {
    const doc = docWithGlossary();
    expect(listBuildingBlocks(doc)).toEqual([{ name: "Signature Block", category: "Legal" }]);
  });

  it("loads cleanly with DocxDocument.load and re-round-trips", () => {
    const doc = docWithGlossary();
    const reloaded = DocxDocument.load(doc.save());
    expect(listBuildingBlocks(reloaded)).toEqual([{ name: "Signature Block", category: "Legal" }]);
    expect(savedParts(reloaded)["word/glossary/document.xml"]).toBe(GLOSSARY_XML);
  });

  it("word interop: a full Word-shaped glossary part (own styles/settings/fontTable/rels) loads cleanly, lists, inserts, and every companion part survives byte-stably", () => {
    // Real Word writes a whole part family for the glossary, mirroring the
    // main document's own — this engine reads only word/glossary/document.xml
    // (DocxDocument.glossaryTree's doc comment) and never touches the rest,
    // so they must round-trip through the generic byte-preservation path
    // (buildPackageFiles starts from the untouched raw package) rather than
    // any glossary-specific code.
    const glossaryStyles =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:styles ${W_NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;
    const glossarySettings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${W_NS}/>`;
    const glossaryFontTable = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts ${W_NS}/>`;
    const glossaryRels =
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>` +
      `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>` +
      `</Relationships>`;
    const doc = DocxDocument.load(
      makeDocx({
        "word/document.xml": wrapDocument(p("anchor")),
        "word/_rels/document.xml.rels": DOC_RELS_WITH_GLOSSARY,
        "word/glossary/document.xml": GLOSSARY_XML,
        "word/glossary/styles.xml": glossaryStyles,
        "word/glossary/settings.xml": glossarySettings,
        "word/glossary/fontTable.xml": glossaryFontTable,
        "word/glossary/_rels/document.xml.rels": glossaryRels,
      }),
    );
    expect(listBuildingBlocks(doc)).toEqual([{ name: "Signature Block", category: "Legal" }]);
    const anchor = doc.docRoot.children.find((c) => c.name === "w:body")!.children[0].children[0].children[0];
    expect(insertBuildingBlock(doc, anchor, "Signature Block")).toBe(true);

    const reloaded = DocxDocument.load(doc.save());
    expect(listBuildingBlocks(reloaded)).toEqual([{ name: "Signature Block", category: "Legal" }]);
    const parts = savedParts(doc);
    expect(parts["word/glossary/styles.xml"]).toBe(glossaryStyles);
    expect(parts["word/glossary/settings.xml"]).toBe(glossarySettings);
    expect(parts["word/glossary/fontTable.xml"]).toBe(glossaryFontTable);
    expect(parts["word/glossary/_rels/document.xml.rels"]).toBe(glossaryRels);
  });
});

describe("insertBuildingBlock", () => {
  it("clones the stored blocks after the caret's paragraph, deep-copied", () => {
    const doc = docWithGlossary(p("anchor"));
    const anchor = doc.docRoot.children.find((c) => c.name === "w:body")!.children[0].children[0].children[0];
    expect(buildingBlockNodeCount(doc, "Signature Block")).toBe(4); // 2 paragraphs + 2 runs
    expect(insertBuildingBlock(doc, anchor, "Signature Block")).toBe(true);

    const body = doc.docRoot.children.find((c) => c.name === "w:body")!;
    const texts = body.children
      .filter((c) => c.name === "w:p")
      .map((pEl) => pEl.children.filter((c) => c.name === "w:r").map((r) => r.children.find((c) => c.name === "w:t")?.text).join(""));
    expect(texts).toEqual(["anchor", "Best regards,", "Jane Doe"]);

    // The glossary part's own tree is untouched (deep clone, not a move).
    expect(listBuildingBlocks(doc)).toEqual([{ name: "Signature Block", category: "Legal" }]);
    const savedGlossary = savedParts(doc)["word/glossary/document.xml"];
    expect(savedGlossary).toBe(GLOSSARY_XML);
  });

  it("is an honest no-op for an unknown name", () => {
    const doc = docWithGlossary();
    const anchor = doc.docRoot.children.find((c) => c.name === "w:body")!.children[0].children[0].children[0];
    expect(insertBuildingBlock(doc, anchor, "Nope")).toBe(false);
    expect(buildingBlockNodeCount(doc, "Nope")).toBe(0);
  });
});

describe("deleteBuildingBlock", () => {
  it("removes a stored building block", () => {
    const doc = docWithGlossary();
    expect(deleteBuildingBlock(doc, "Signature Block")).toBe(true);
    expect(listBuildingBlocks(doc)).toEqual([]);
    expect(savedParts(doc)["word/glossary/document.xml"]).not.toContain("Signature Block");
  });

  it("is an honest no-op when the name is absent, including on a package with no glossary part at all", () => {
    const doc = docWithGlossary();
    expect(deleteBuildingBlock(doc, "Nope")).toBe(false);
    expect(savedParts(doc)["word/glossary/document.xml"]).toBe(GLOSSARY_XML);

    expect(deleteBuildingBlock(blankDoc(), "Anything")).toBe(false);
  });
});
