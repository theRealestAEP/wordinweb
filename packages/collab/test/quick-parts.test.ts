import { describe, expect, it } from "vitest";
import { zipSync, strToU8, strFromU8, unzipSync } from "fflate";
import { DocxDocument, documentOperationBody, operationBody, serializeXml, encodeClipboardOoxml, type Paragraph, type Run } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { validateIntent } from "../src/validate.js";

/**
 * Quick Parts / Building Blocks OVER THE WIRE.
 *
 * createBuildingBlock/deleteBuildingBlock are document-scoped like the
 * source-management ops, and the convergence story is the same: the
 * glossary part is sequenced state, every value comes from the payload, and
 * the rejection predicate (the building-block NAME) is decided from that
 * shared state, so every replica decides alike. Convergence is checked on
 * the SAVED PACKAGE BYTES, not just the XML trees, because these operations
 * write a part save() serializes separately — including the part-creation
 * path. insertBuildingBlock is run-scoped like insertBibliography: its
 * content is DERIVED from the already-synced glossary part rather than
 * carried, so convergence also covers the carried stable-id assignment.
 */

const GLOSSARY_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:glossaryDocument xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docParts>` +
  `<w:docPart><w:docPartPr><w:name w:val="Signature Block"/>` +
  `<w:category><w:name w:val="Legal"/><w:gallery w:val="docParts"/></w:category>` +
  `</w:docPartPr><w:docPartBody>` +
  `<w:p><w:r><w:t xml:space="preserve">Best regards,</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t xml:space="preserve">Jane Doe</w:t></w:r></w:p>` +
  `</w:docPartBody></w:docPart>` +
  `</w:docParts></w:glossaryDocument>`;

function makeDoc(withGlossary: boolean): DocxDocument {
  const documentXml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body><w:p><w:r><w:t xml:space="preserve">anchor</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
  return DocxDocument.load(
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
      "word/document.xml": strToU8(documentXml),
      ...(withGlossary
        ? {
            "word/_rels/document.xml.rels": strToU8(
              `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
                `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" Target="glossary/document.xml"/></Relationships>`,
            ),
            "word/glossary/document.xml": strToU8(GLOSSARY_XML),
          }
        : {}),
    }),
  );
}

const base = { clientId: "a", clientSeq: 1, base: 0 } as const;

/** Replicated trees PLUS the saved package: the part these operations move is
 * serialized by save(), so byte-convergence is asserted there — the
 * citation-sources.test.ts pattern. */
function stateOf(s: DocumentSession): string {
  const xml = s.doc.editableRoots().map((r) => serializeXml(r)).join("|");
  return xml + "||" + Buffer.from(s.doc.save()).toString("base64");
}

function anchorRunId(s: DocumentSession): number {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  return s.ids.idOf((para.children[0] as Run).src!)!;
}

function runIds(s: DocumentSession): (number | undefined)[] {
  const out: (number | undefined)[] = [];
  const walk = (el: { name: string; children: unknown[] }): void => {
    if (el.name === "w:r" || el.name.endsWith(":r")) out.push(s.ids.idOf(el as never));
    for (const c of el.children) walk(c as never);
  };
  for (const root of s.doc.editableRoots()) walk(root as never);
  return out;
}

const GREETING_BLOCKS_XML = encodeClipboardOoxml([
  { name: "w:p", attrs: {}, children: [{ name: "w:r", attrs: {}, children: [{ name: "w:t", attrs: { "xml:space": "preserve" }, children: [], text: "Hello there" }], text: "" }], text: "" },
]);

describe("createBuildingBlock over the wire", () => {
  it("creates the glossary part and the docPart identically on two replicas of a blank document", () => {
    const a = new DocumentSession(makeDoc(false));
    const b = new DocumentSession(makeDoc(false));
    const intent = {
      ...base,
      ...documentOperationBody("createBuildingBlock", { name: "Greeting", category: "General", blocksXml: GREETING_BLOCKS_XML }),
    };
    expect(a.submit(intent as never).kind).toBe("applied");
    expect(b.submit(intent as never).kind).toBe("applied");
    expect(stateOf(a)).toBe(stateOf(b));
    const saved = Buffer.from(a.doc.save());
    expect(saved.toString("latin1")).toContain("word/glossary/document.xml");
    const glossary = strFromU8(unzipSync(saved)["word/glossary/document.xml"]);
    expect(glossary).toContain("Hello there");
    expect(glossary).toContain(`<w:name w:val="Greeting"/>`);
  });

  it("rejects a duplicate name and a malformed payload before sequencing", () => {
    const a = new DocumentSession(makeDoc(true));
    const intent = {
      ...base,
      ...documentOperationBody("createBuildingBlock", { name: "Signature Block", category: "Legal", blocksXml: GREETING_BLOCKS_XML }),
    };
    const before = stateOf(a);
    expect(a.submit(intent as never).kind).toBe("rejected");
    expect(stateOf(a)).toBe(before);

    for (const body of [
      { name: "", blocksXml: GREETING_BLOCKS_XML },
      { name: "x".repeat(65), blocksXml: GREETING_BLOCKS_XML },
      { name: "Ok", blocksXml: "" },
      { name: "Ok", category: "c".repeat(65), blocksXml: GREETING_BLOCKS_XML },
    ]) {
      expect(
        validateIntent({ ...base, ...documentOperationBody("createBuildingBlock", body as never) } as never),
        JSON.stringify(body),
      ).not.toBeNull();
    }
  });
});

describe("insertBuildingBlock over the wire", () => {
  it("derives the same cloned blocks and carried ids on two replicas from the shared glossary part", () => {
    const a = new DocumentSession(makeDoc(true));
    const b = new DocumentSession(makeDoc(true));
    let next = 900;
    const intent = {
      ...base,
      ...operationBody(
        "insertBuildingBlock",
        anchorRunId(a),
        { name: "Signature Block", blockCount: 4 },
        (n) => Array.from({ length: n }, () => next++),
      ),
    };
    expect(a.submit(intent as never).kind).toBe("applied");
    expect(b.submit(intent as never).kind).toBe("applied");
    expect(stateOf(a)).toBe(stateOf(b));
    expect(runIds(a)).toEqual(runIds(b));
    expect(stateOf(a)).toContain("Best regards,");
    expect(stateOf(a)).toContain("Jane Doe");
  });

  it("is an honest no-op for an unknown name, and rejects a malformed blockCount before sequencing", () => {
    const s = new DocumentSession(makeDoc(true));
    const before = stateOf(s);
    let next = 900;
    const alloc = (n: number) => Array.from({ length: n }, () => next++);
    const intent = { ...base, ...operationBody("insertBuildingBlock", anchorRunId(s), { name: "Nope", blockCount: 4 }, alloc) };
    expect(s.submit(intent as never).kind).toBe("rejected");
    expect(stateOf(s)).toBe(before);

    for (const blockCount of [-1, 1.5, 5001]) {
      const bad = { ...base, ...operationBody("insertBuildingBlock", anchorRunId(s), { name: "Signature Block", blockCount }, alloc) };
      expect(validateIntent(bad as never), String(blockCount)).not.toBeNull();
    }
  });
});

describe("deleteBuildingBlock over the wire", () => {
  it("removes the docPart convergently, and rejects for an unknown name", () => {
    const a = new DocumentSession(makeDoc(true));
    const b = new DocumentSession(makeDoc(true));
    const del = { ...base, ...documentOperationBody("deleteBuildingBlock", { name: "Signature Block" }) };
    expect(a.submit(del as never).kind).toBe("applied");
    expect(b.submit(del as never).kind).toBe("applied");
    expect(stateOf(a)).toBe(stateOf(b));
    expect(stateOf(a)).not.toContain("Signature Block");

    const before = stateOf(a);
    const again = { ...base, clientSeq: 2, ...documentOperationBody("deleteBuildingBlock", { name: "Signature Block" }) };
    expect(a.submit(again as never).kind).toBe("rejected");
    expect(stateOf(a)).toBe(before);
  });
});
