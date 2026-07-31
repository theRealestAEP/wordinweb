import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientReplica } from "../src/replica.js";

/**
 * The suggest-intent lifecycle (doc 14 §3 L2): suggested inserts (w:ins),
 * suggestRevision strikes (w:del wraps) and paragraph-glyph marks — all
 * converging byte-identically across replicas + server, with the carried
 * author/date landing in the OOXML (the durable attribution that outlives
 * zero custody and opens attributed in Word), and the already-wired
 * accept-all review flow resolving everything identically.
 */

function docBytes(text: string): Uint8Array {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`;
  return zipSync({
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
  });
}

function harness() {
  const server = new DocumentSession(DocxDocument.load(docBytes("hello world")));
  const a = new ClientReplica(docBytes("hello world"));
  const b = new ClientReplica(docBytes("hello world"));
  const roundTrip = (intent: never) => {
    a.submitLocal(intent);
    const entry = server.submit(intent);
    a.receive([entry]);
    b.receive([entry]);
    return entry;
  };
  const xmlOf = (r: ClientReplica) => serializeXml(r.doc.docRoot);
  const serverXml = () => serializeXml(server.doc.docRoot);
  return { server, a, b, roundTrip, xmlOf, serverXml };
}

const SUG = { author: "Priya", date: "2026-07-23T12:00:00Z" };

describe("suggest lifecycle (doc 14 §3 L2)", () => {
  it("a suggested insert converges byte-identically and lands as w:ins with the carried author", () => {
    const { a, b, roundTrip, xmlOf, serverXml } = harness();
    const e = roundTrip({
      kind: "insertText", clientId: "a", clientSeq: 1, base: 0,
      at: { blockId: 1, runId: 2, offset: 5 }, text: " BRAVE", suggest: SUG,
    } as never);
    expect(e.kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(xmlOf(a)).toBe(serverXml());
    expect(xmlOf(a)).toContain("w:ins");
    expect(xmlOf(a)).toContain(`w:author="Priya"`); // durable attribution
    expect(xmlOf(a)).toContain(`w:date="2026-07-23T12:00:00Z"`); // carried, not re-derived
    expect(xmlOf(a)).toContain("BRAVE");
  });

  it("suggestRevision strikes wrap in w:del (text preserved) and paragraph marks record glyph revisions", () => {
    const { a, b, roundTrip, xmlOf, serverXml } = harness();
    const e = roundTrip({
      kind: "suggestRevision", clientId: "a", clientSeq: 1, base: 0,
      ranges: [{ blockId: 1, runId: 2, start: 0, end: 5 }], // strike "hello"
      marks: [{ blockId: 1, glyph: "del" }], // suggest merging this paragraph away
      suggest: SUG,
    } as never);
    expect(e.kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(xmlOf(a)).toBe(serverXml());
    expect(xmlOf(a)).toContain("w:del");
    expect(xmlOf(a)).toContain("hello"); // struck, NOT removed — Word keeps both sides
    expect(xmlOf(a)).toContain(`w:author="Priya"`);
  });

  it("accept-all resolves suggested content identically everywhere (the review flow was already wired)", () => {
    const { a, b, roundTrip, xmlOf, serverXml } = harness();
    roundTrip({
      kind: "insertText", clientId: "a", clientSeq: 1, base: 0,
      at: { blockId: 1, runId: 2, offset: 5 }, text: " BRAVE", suggest: SUG,
    } as never);
    // The suggested insert SPLIT the paragraph's runs — old run ids retire
    // and fresh ones are assigned deterministically (identical walk +
    // counter on every replica). The editor always encodes against its
    // CURRENT doc, so a sequential gesture addresses the new id — mirror
    // that here by resolving "hello"'s run id from the live table. (A
    // truly-concurrent stale-id strike no-ops identically everywhere —
    // the documented OT-lite degradation, same class as formatRange.)
    const helloRunId = [3, 4, 5].find((id) => {
      const el = a.ids.elOf(id);
      return el?.children.some((c) => c.text === "hello");
    })!;
    roundTrip({
      kind: "suggestRevision", clientId: "a", clientSeq: 2, base: 1,
      ranges: [{ blockId: 1, runId: helloRunId, start: 0, end: 5 }], suggest: SUG,
    } as never);
    const e = roundTrip({ kind: "acceptAllRevisions", clientId: "a", clientSeq: 3, base: 2 } as never);
    expect(e.kind).toBe("applied");
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(xmlOf(a)).toBe(serverXml());
    expect(xmlOf(a)).not.toContain("w:ins"); // resolved
    expect(xmlOf(a)).not.toContain("w:del");
    expect(xmlOf(a)).toContain("BRAVE"); // accepted insert stays
    expect(xmlOf(a)).not.toContain("hello"); // accepted deletion goes
  });

  it("transform: a strike's offsets shift over a concurrent earlier insert in the same run", () => {
    const { a, b, server, xmlOf, serverXml } = harness();
    // TRUE concurrency: both authored against base 0; b's insert wins the
    // sequencing race, so a's strike of "hello" [0,5) must transform to
    // [2,7) — still the word "hello", never "XXhel".
    const bIns = { kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 0 }, text: "XX" } as never;
    const strike = {
      kind: "suggestRevision", clientId: "a", clientSeq: 1, base: 0,
      ranges: [{ blockId: 1, runId: 2, start: 0, end: 5 }], suggest: SUG,
    } as never;
    a.submitLocal(strike);
    b.submitLocal(bIns);
    const e1 = server.submit(bIns);
    const e2 = server.submit(strike);
    a.receive([e1, e2]); // rollback-replay: a's optimistic strike rebases
    b.receive([e1, e2]);
    expect(xmlOf(a)).toBe(xmlOf(b));
    expect(xmlOf(a)).toBe(serverXml());
    // The struck region is exactly "hello": XX stays outside the w:del.
    const xml = xmlOf(a);
    const delIdx = xml.indexOf("<w:del ");
    expect(delIdx).toBeGreaterThan(-1);
    const delChunk = xml.slice(delIdx, xml.indexOf("</w:del>"));
    expect(delChunk).toContain("hello");
    expect(delChunk).not.toContain("XX");
  });

  it("hostile bounds reject cleanly (blocker-3 discipline extends to the new kind)", () => {
    const { server } = harness();
    const huge = server.submit({
      kind: "suggestRevision", clientId: "m", clientSeq: 1, base: 0,
      ranges: Array.from({ length: 500 }, () => ({ blockId: 1, runId: 2, start: 0, end: 1 })),
      suggest: SUG,
    } as never);
    expect(huge.kind).toBe("rejected");
    const empty = server.submit({
      kind: "suggestRevision", clientId: "m", clientSeq: 2, base: 0, suggest: SUG,
    } as never);
    expect(empty.kind).toBe("rejected");
  });
});
