import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { DocxDocument, encodeClipboardOoxml, listBuildingBlocks, localName, type XmlElement } from "@wordinweb/core";
import { AgentDocument, validateAgentOperationShape } from "../src/index.js";
import { body } from "./helpers.js";

/**
 * Quick Parts / Building Blocks through the agent tool surface: the
 * registered-operation machinery (capabilities, schema, id budget, apply)
 * needs no per-kind agent code, only the capability rows and field schemas
 * declared in capabilities.ts — this exercises that wiring end to end.
 */

function textOf(doc: DocxDocument): string {
  const collect = (element: XmlElement): string =>
    (localName(element.name) === "t" ? element.text : "") + element.children.map(collect).join("");
  return collect(doc.docRoot);
}

const GLOSSARY_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts>` +
  `<w:docPart><w:docPartPr><w:name w:val="Signature Block"/>` +
  `<w:category><w:name w:val="Legal"/><w:gallery w:val="docParts"/></w:category>` +
  `</w:docPartPr><w:docPartBody>` +
  `<w:p><w:r><w:t xml:space="preserve">Best regards,</w:t></w:r></w:p>` +
  `</w:docPartBody></w:docPart>` +
  `</w:docParts></w:glossaryDocument>`;

function agentWithGlossary(): AgentDocument {
  return AgentDocument.load(
    zipSync({
      "[Content_Types].xml": strToU8(
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
      ),
      "word/glossary/document.xml": strToU8(GLOSSARY_XML),
      "word/document.xml": strToU8(body(`<w:p><w:r><w:t xml:space="preserve">anchor</w:t></w:r></w:p>`)),
    }),
  );
}

describe("Quick Parts through the agent tool surface", () => {
  it("exposes createBuildingBlock, insertBuildingBlock, deleteBuildingBlock with their registry-declared capabilities", () => {
    const capabilities = AgentDocument.create().capabilities();
    const byKind = new Map(capabilities.map((capability) => [capability.kind, capability]));
    expect(byKind.get("createBuildingBlock")).toMatchObject({
      category: "document",
      required: ["name", "blocksXml"],
      optional: ["category"],
    });
    expect(byKind.get("insertBuildingBlock")).toMatchObject({
      category: "insert",
      required: ["runRef", "name", "blockCount"],
    });
    expect(byKind.get("deleteBuildingBlock")).toMatchObject({ category: "document", required: ["name"] });
  });

  it("rejects a malformed name/blockCount before the wire", () => {
    expect(validateAgentOperationShape({ kind: "deleteBuildingBlock", name: "Ok" })).toBeNull();
    expect(validateAgentOperationShape({ kind: "deleteBuildingBlock", name: "" })).not.toBeNull();
    expect(
      validateAgentOperationShape({ kind: "insertBuildingBlock", runRef: "run:1", name: "Ok", blockCount: -1 }),
    ).not.toBeNull();
  });

  it("saves a selection as a building block, then inserts it at a run", async () => {
    const agent = AgentDocument.load(
      zipSync({
        "[Content_Types].xml": strToU8(
          `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
            `<Default Extension="xml" ContentType="application/xml"/>` +
            `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
            `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
        ),
        "_rels/.rels": strToU8(
          `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
        ),
        "word/document.xml": strToU8(body(`<w:p><w:r><w:t xml:space="preserve">anchor</w:t></w:r></w:p>`)),
      }),
    );
    const blocksXml = encodeClipboardOoxml([
      { name: "w:p", attrs: {}, children: [{ name: "w:r", attrs: {}, children: [{ name: "w:t", attrs: { "xml:space": "preserve" }, children: [], text: "Sincerely," }], text: "" }], text: "" },
    ]);
    const created = await agent.edit({
      revision: agent.revision,
      operations: [{ kind: "createBuildingBlock", name: "Sign-off", category: "Letters", blocksXml }],
    });
    expect(created.status).toBe("applied");
    expect(listBuildingBlocks(DocxDocument.load(agent.save()))).toEqual([{ name: "Sign-off", category: "Letters" }]);

    const read = agent.inspect({ kind: "read" });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
    const runRef = read.blocks[0].runs[0].ref;
    const inserted = await agent.edit({
      revision: agent.revision,
      operations: [{ kind: "insertBuildingBlock", runRef, name: "Sign-off", blockCount: 2 }],
    });
    expect(inserted.status).toBe("applied");
    expect(textOf(DocxDocument.load(agent.save()))).toContain("Sincerely,");
  });

  it("lists and deletes an arriving Word building block", async () => {
    const agent = agentWithGlossary();
    expect(listBuildingBlocks(DocxDocument.load(agent.save()))).toEqual([{ name: "Signature Block", category: "Legal" }]);

    const read = agent.inspect({ kind: "read" });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
    const runRef = read.blocks[0].runs[0].ref;
    const inserted = await agent.edit({
      revision: agent.revision,
      operations: [{ kind: "insertBuildingBlock", runRef, name: "Signature Block", blockCount: 2 }],
    });
    expect(inserted.status).toBe("applied");
    expect(textOf(DocxDocument.load(agent.save()))).toContain("Best regards,");

    const deleted = await agent.edit({
      revision: agent.revision,
      operations: [{ kind: "deleteBuildingBlock", name: "Signature Block" }],
    });
    expect(deleted.status).toBe("applied");
    expect(listBuildingBlocks(DocxDocument.load(agent.save()))).toEqual([]);
  });
});
