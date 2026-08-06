import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import {
  DocxDocument,
  operationBody,
  serializeXml,
  tocEntryCount,
  type Paragraph,
  type Run,
} from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { validateIntent } from "../src/validate.js";

/**
 * insertToc OVER THE WIRE.
 *
 * A TOC used to refuse in a room: it adds one paragraph per heading, and the
 * carried id allocation that keeps two replicas naming the same new nodes is
 * sized from the PAYLOAD, which knew nothing about how many headings the
 * document had. `entryCount` is that missing number — a budget the originator
 * computes with tocEntryCount, never an instruction, since the mutation still
 * builds its entries from the replica's own headings.
 *
 * What these tests pin is the property the budget exists for: two replicas
 * applying the identical intent end up byte-identical AND agree on the stable
 * id of every paragraph the insert created.
 */

const STYLES_XML =
  `<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
  [1, 2, 3]
    .map(
      (n) =>
        `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/>` +
        `<w:pPr><w:outlineLvl w:val="${n - 1}"/></w:pPr></w:style>`,
    )
    .join("") +
  `</w:styles>`;

const heading = (level: number, text: string) =>
  `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

function makeDoc(): DocxDocument {
  const body =
    `<w:p><w:r><w:t xml:space="preserve">anchor</w:t></w:r></w:p>` +
    heading(1, "Introduction") +
    heading(2, "Background") +
    heading(1, "Method");
  const documentXml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
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
      "word/styles.xml": strToU8(STYLES_XML),
    }),
  );
}

/** The stable id of the run in the first (anchor) paragraph. */
function anchorRunId(s: DocumentSession): number {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  return s.ids.idOf((para.children[0] as Run).src!)!;
}

const base = { clientId: "a", clientSeq: 1, base: 0 } as const;

/** The intent a client would send, sized the way the React host sizes it. */
function tocIntent(s: DocumentSession, options: Record<string, unknown> = {}) {
  let next = 900;
  return {
    ...base,
    ...operationBody(
      "insertToc",
      anchorRunId(s),
      { entryCount: tocEntryCount(s.doc, options), ...options } as never,
      (n) => Array.from({ length: n }, () => next++),
    ),
  };
}

function xmlOf(s: DocumentSession): string {
  return [...s.doc.editableRoots(), s.doc.stylesTree()!].map((r) => serializeXml(r)).join("|");
}

/** Every paragraph's stable id, in document order. */
function paragraphIds(s: DocumentSession): (number | undefined)[] {
  const out: (number | undefined)[] = [];
  const walk = (el: { name: string; children: unknown[] }): void => {
    if (el.name.endsWith("p") || el.name === "p") out.push(s.ids.idOf(el as never));
    for (const c of el.children) walk(c as never);
  };
  for (const root of s.doc.editableRoots()) walk(root as never);
  return out;
}

describe("insertToc over the wire", () => {
  it("builds the same TOC on two replicas, byte for byte", () => {
    const a = new DocumentSession(makeDoc());
    const b = new DocumentSession(makeDoc());
    const intent = tocIntent(a);
    expect(a.submit(intent as never).kind).toBe("applied");
    expect(b.submit(intent as never).kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    // A real TOC field, with one entry per heading.
    expect(xmlOf(a)).toContain(`<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>`);
    expect((xmlOf(a).match(/PAGEREF/g) ?? []).length).toBe(3);
  });

  it("gives every new paragraph the SAME stable id on both replicas", () => {
    // This is what the carried allocation is FOR. Without it each replica
    // would name the new paragraphs from its own counter, and the next intent
    // addressing one of them would land in a different place on each.
    const a = new DocumentSession(makeDoc());
    const b = new DocumentSession(makeDoc());
    const intent = tocIntent(a);
    a.submit(intent as never);
    b.submit(intent as never);
    expect(paragraphIds(a)).toEqual(paragraphIds(b));
    // …and the ids actually came from the intent rather than from a local
    // counter, so they are the ones every other replica was told about.
    const carried = new Set((intent as { nodeIds: number[] }).nodeIds);
    const fresh = paragraphIds(a).filter((id): id is number => id !== undefined && carried.has(id));
    expect(fresh.length, "TOC paragraphs carrying wire ids").toBe(4); // 3 entries + the field's end
  });

  it("sizes the allocation for the entries the document actually has", () => {
    const s = new DocumentSession(makeDoc());
    // Three headings within levels 1-3; restricting to level 1 leaves two.
    expect(tocEntryCount(s.doc)).toBe(3);
    expect(tocEntryCount(s.doc, { levels: [1, 1] })).toBe(2);
    // A document with no qualifying heading still gets one paragraph — the
    // "no entries" text — so the budget is never zero.
    expect(tocEntryCount(s.doc, { levels: [8, 9] })).toBe(1);
  });

  it("honours the levels and leader the caller asked for", () => {
    const s = new DocumentSession(makeDoc());
    s.submit(tocIntent(s, { levels: [1, 1], leader: "hyphen" }) as never);
    const xml = xmlOf(s);
    expect(xml).toContain(`> TOC \\o "1-1" \\h \\z \\u <`);
    expect(xml).toContain(`w:leader="hyphen"`);
    // Only the two level-1 headings became entries.
    expect((xml.match(/PAGEREF/g) ?? []).length).toBe(2);
  });

  it("leaves page numbers as placeholders for the update pass to carry", () => {
    // Page numbers come from a layout, and a layout depends on the host's
    // font metrics — the exact value updateFields exists to send as data.
    const s = new DocumentSession(makeDoc());
    s.submit(tocIntent(s) as never);
    expect(xmlOf(s)).toContain(`<w:t xml:space="preserve">1</w:t>`);
  });

  it("rejects a malformed payload before it is sequenced", () => {
    const s = new DocumentSession(makeDoc());
    const runId = anchorRunId(s);
    const cases: [string, Record<string, unknown>][] = [
      ["no entry budget", { entryCount: 0 }],
      ["a non-integer budget", { entryCount: 2.5 }],
      ["an absurd budget", { entryCount: 100000 }],
      ["a backwards level range", { entryCount: 3, levels: [3, 1] }],
      ["a level above 9", { entryCount: 3, levels: [1, 10] }],
      ["a one-ended level range", { entryCount: 3, levels: [1] }],
      ["an unknown leader", { entryCount: 3, leader: "squiggle" }],
    ];
    for (const [what, payload] of cases) {
      expect(
        validateIntent({ ...base, kind: "insertToc", runId, nodeIds: [], ...payload } as never),
        what,
      ).not.toBeNull();
    }
    expect(
      validateIntent({ ...base, kind: "insertToc", runId, nodeIds: [], entryCount: 3, levels: [1, 3], leader: "dot" } as never),
    ).toBeNull();
  });

  it("is an honest no-op when the anchor run is gone", () => {
    // The address is the whole rejection predicate: a replica that no longer
    // has the run applies nothing rather than guessing a position.
    const s = new DocumentSession(makeDoc());
    const before = xmlOf(s);
    const outcome = s.submit({
      ...base,
      ...operationBody("insertToc", 999999, { entryCount: 3 } as never, (n) =>
        Array.from({ length: n }, (_, i) => 900 + i),
      ),
    } as never);
    expect(outcome.kind).not.toBe("applied");
    expect(xmlOf(s)).toBe(before);
  });
});
