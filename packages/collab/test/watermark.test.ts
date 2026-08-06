import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, documentOperationBody, headerWatermarks, serializeXml } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { validateIntent } from "../src/validate.js";

/**
 * insertWatermark / removeWatermark OVER THE WIRE.
 *
 * Watermark editing used to be a pure local mutation against header-part VML.
 * Header parts DO replicate — they are editable roots, and their runs carry
 * stable ids — so every watermark click in a room forked the document.
 *
 * Both operations are document-scoped, which means the address gives them no
 * honest-no-op protection and their payload has to carry its own rejection
 * predicate. For insertWatermark that is `headerCount`, which does double duty:
 * it sizes the carried ids (the number of nodes created depends on how many
 * header parts the document has, which `nodeIds` cannot see) and it rejects a
 * replica whose header set has moved. For removeWatermark it is the
 * watermark's own existence.
 *
 * What these tests pin is the property that matters: two replicas applying the
 * identical intent end up byte-identical, ids included.
 */

const HEADER_XML =
  `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:p><w:r><w:t xml:space="preserve">Chapter one</w:t></w:r></w:p></w:hdr>`;

/** A document with `count` header parts, all referenced from the one section. */
function makeDoc(count = 2): DocxDocument {
  const types = ["default", "first", "even"];
  const parts: Record<string, Uint8Array> = {};
  const refs: string[] = [];
  const rels: string[] = [];
  const overrides: string[] = [];
  for (let i = 0; i < count; i++) {
    parts[`word/header${i + 1}.xml`] = strToU8(HEADER_XML);
    refs.push(`<w:headerReference w:type="${types[i]}" r:id="rIdH${i + 1}"/>`);
    rels.push(
      `<Relationship Id="rIdH${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header${i + 1}.xml"/>`,
    );
    overrides.push(
      `<Override PartName="/word/header${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`,
    );
  }
  const documentXml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>` +
    `<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>` +
    `<w:sectPr>${refs.join("")}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
  return DocxDocument.load(
    zipSync({
      "[Content_Types].xml": strToU8(
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
          `${overrides.join("")}</Types>`,
      ),
      "_rels/.rels": strToU8(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join("")}</Relationships>`,
      ),
      "word/document.xml": strToU8(documentXml),
      ...parts,
    }),
  );
}

const base = { clientId: "a", clientSeq: 1, base: 0 } as const;

/** The intent a client would send, sized the way the React host sizes it. */
function insertIntent(s: DocumentSession, args: Record<string, unknown> = {}) {
  let next = 900;
  return {
    ...base,
    ...documentOperationBody(
      "insertWatermark",
      { text: "CONFIDENTIAL", headerCount: s.doc.headerRoots().length, ...args } as never,
      (n) => Array.from({ length: n }, () => next++),
    ),
  };
}

const removeIntent = { ...base, ...documentOperationBody("removeWatermark", {} as never) };

function xmlOf(s: DocumentSession): string {
  return s.doc.editableRoots().map((r) => serializeXml(r)).join("|");
}

/** Every run's stable id, in document order across every editable root. */
function runIds(s: DocumentSession): (number | undefined)[] {
  const out: (number | undefined)[] = [];
  const walk = (el: { name: string; children: unknown[] }): void => {
    if (el.name === "w:r" || el.name.endsWith(":r")) out.push(s.ids.idOf(el as never));
    for (const c of el.children) walk(c as never);
  };
  for (const root of s.doc.editableRoots()) walk(root as never);
  return out;
}

describe("insertWatermark over the wire", () => {
  it("stamps the same watermark on two replicas, byte for byte", () => {
    const a = new DocumentSession(makeDoc());
    const b = new DocumentSession(makeDoc());
    const intent = insertIntent(a);
    expect(a.submit(intent as never).kind).toBe("applied");
    expect(b.submit(intent as never).kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(headerWatermarks(a.doc)).toHaveLength(2);
    // The full _x0000_t136 guide path is what makes Word fit the glyphs to the
    // box instead of painting the literal 1pt text.
    expect(xmlOf(a)).toContain('path="m@7,0l@8,0m@5,21600l@6,21600e"');
    expect(xmlOf(a)).toContain('string="CONFIDENTIAL"');
  });

  it("gives every new run the SAME stable id on both replicas", () => {
    const a = new DocumentSession(makeDoc());
    const b = new DocumentSession(makeDoc());
    const intent = insertIntent(a);
    a.submit(intent as never);
    b.submit(intent as never);
    expect(runIds(a)).toEqual(runIds(b));
    // …and the ids came from the intent rather than a local counter, so they
    // are the ones every other replica was told about.
    const carried = new Set((intent as { nodeIds: number[] }).nodeIds);
    expect(runIds(a).filter((id) => id !== undefined && carried.has(id))).toHaveLength(2);
  });

  it("rejects when the replica's header set has moved", () => {
    // The rejection predicate standing in for the stable id a document-scoped
    // operation does not have. Both replicas in this position reject alike.
    const s = new DocumentSession(makeDoc(2));
    const stale = { ...base, ...documentOperationBody("insertWatermark", { text: "DRAFT", headerCount: 3, nodeIds: [] } as never) };
    expect(s.submit(stale as never).kind).toBe("rejected");
    expect(headerWatermarks(s.doc)).toHaveLength(0);
  });

  it("replaces rather than stacking, so a second stamp converges too", () => {
    const a = new DocumentSession(makeDoc());
    const b = new DocumentSession(makeDoc());
    for (const intent of [insertIntent(a), { ...insertIntent(a, { text: "FINAL" }), clientSeq: 2 }]) {
      a.submit(intent as never);
      b.submit(intent as never);
    }
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(headerWatermarks(a.doc)).toHaveLength(2);
    expect(xmlOf(a)).not.toContain('string="CONFIDENTIAL"');
  });

  it("leaves the header's own text alone", () => {
    const s = new DocumentSession(makeDoc());
    s.submit(insertIntent(s) as never);
    expect(xmlOf(s)).toContain("Chapter one");
  });

  it("refuses a malformed payload before it reaches the document", () => {
    const cases: [string, Record<string, unknown>][] = [
      ["empty text", { text: "", headerCount: 1 }],
      ["a header count of zero", { text: "X", headerCount: 0 }],
      ["a non-integer header count", { text: "X", headerCount: 1.5 }],
      ["a colour with a hash", { text: "X", headerCount: 1, color: "#C0C0C0" }],
      ["an opacity above 1", { text: "X", headerCount: 1, opacity: 2 }],
      ["a non-boolean diagonal", { text: "X", headerCount: 1, diagonal: "yes" }],
    ];
    for (const [why, args] of cases) {
      const intent = { ...base, kind: "insertWatermark", nodeIds: [], ...args };
      expect(validateIntent(intent as never), why).not.toBeNull();
    }
  });
});

describe("removeWatermark over the wire", () => {
  it("takes the watermark off both replicas identically", () => {
    const a = new DocumentSession(makeDoc());
    const b = new DocumentSession(makeDoc());
    const stamp = insertIntent(a);
    a.submit(stamp as never);
    b.submit(stamp as never);
    const remove = { ...removeIntent, clientSeq: 2, base: a.seq };
    expect(a.submit(remove as never).kind).toBe("applied");
    expect(b.submit(remove as never).kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(headerWatermarks(a.doc)).toHaveLength(0);
    expect(xmlOf(a)).toContain("Chapter one");
  });

  it("rejects cleanly when there is no watermark to remove", () => {
    const s = new DocumentSession(makeDoc());
    expect(s.submit({ ...removeIntent } as never).kind).toBe("rejected");
  });

  it("retires the stable ids of the runs it removed", () => {
    // Deletion and pruning are paired everywhere else in the apply path; an id
    // left pointing at a detached run is a later intent aimed at nothing.
    const s = new DocumentSession(makeDoc());
    const stamp = insertIntent(s);
    s.submit(stamp as never);
    const carried = (stamp as { nodeIds: number[] }).nodeIds;
    const used = carried.filter((id) => s.ids.elOf(id) !== undefined);
    expect(used.length).toBeGreaterThan(0);
    s.submit({ ...removeIntent, clientSeq: 2, base: s.seq } as never);
    for (const id of used) expect(s.ids.elOf(id), `id ${id}`).toBeUndefined();
  });
});
