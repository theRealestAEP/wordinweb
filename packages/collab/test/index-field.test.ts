import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, indexEntryCount, operationBody, documentOperationBody, serializeXml, localName } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientReplica } from "../src/replica.js";
import type { Intent } from "../src/intents.js";

function docBytes(paras: string[]): Uint8Array {
  const body = paras.map((t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`).join("");
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
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

function runIdOf(session: DocumentSession, paraIdx: number): number {
  const body = session.doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const p = body.children.filter((c) => localName(c.name) === "p")[paraIdx];
  const r = p.children.find((c) => localName(c.name) === "r")!;
  return session.ids.idOf(r)!;
}

describe("the index cluster on the wire (registered operations)", () => {
  it("marks entries, builds, and rebuilds byte-identically on every replica", () => {
    const initial = docBytes(["about widgets", "about anchors", "the index goes here"]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);
    let next = 1000;
    const alloc = (n: number) => Array.from({ length: n }, () => next++);
    let seq = 0;
    const submit = (intent: Intent) => {
      a.submitLocal(intent);
      const e = server.submit(intent);
      a.receive([e]);
      b.receive([e]);
      seq = server.seq;
    };
    const agree = () => {
      const serverXml = serializeXml(server.doc.docRoot);
      expect(serializeXml(a.doc.docRoot)).toBe(serverXml);
      expect(serializeXml(b.doc.docRoot)).toBe(serverXml);
      return serverXml;
    };

    submit({
      ...operationBody("insertIndexEntry", runIdOf(server, 0), { entry: "Widgets:assembly" }, alloc),
      clientId: "a", clientSeq: 1, base: seq,
    } as Intent);
    submit({
      ...operationBody("insertIndexEntry", runIdOf(server, 1), { entry: "anchors" }, alloc),
      clientId: "a", clientSeq: 2, base: seq,
    } as Intent);
    let xml = agree();
    expect(xml).toContain(` XE "Widgets:assembly" `);

    submit({
      ...operationBody("insertIndex", runIdOf(server, 2), { entryCount: indexEntryCount(server.doc) }, alloc),
      clientId: "a", clientSeq: 3, base: seq,
    } as Intent);
    xml = agree();
    expect(xml).toContain(" INDEX ");
    expect(xml).toContain(">anchors<");
    expect(xml).toContain(">assembly<");
    expect(xml).toContain("PAGEREF _Idx");

    // A refresh with an unchanged entry set is an honest no-op everywhere.
    const before = agree();
    submit({
      ...documentOperationBody("refreshIndex", { entryCount: indexEntryCount(server.doc) }, alloc),
      clientId: "b", clientSeq: 1, base: seq,
    } as Intent);
    expect(agree()).toBe(before);

    // A new mark, then a refresh: the rebuilt index converges with the new
    // entry sorted in.
    submit({
      ...operationBody("insertIndexEntry", runIdOf(server, 2), { entry: "bolts" }, alloc),
      clientId: "a", clientSeq: 4, base: seq,
    } as Intent);
    submit({
      ...documentOperationBody("refreshIndex", { entryCount: indexEntryCount(server.doc) }, alloc),
      clientId: "a", clientSeq: 5, base: seq,
    } as Intent);
    xml = agree();
    expect(xml).toContain(">bolts<");
    expect(xml.indexOf(">anchors<")).toBeLessThan(xml.indexOf(">bolts<"));
  });

  it("rejects malformed payloads at validation", async () => {
    const { validateIntent } = await import("../src/validate.js");
    const base = { clientId: "a", clientSeq: 1, base: 0 };
    expect(validateIntent({ kind: "insertIndexEntry", runId: 1, entry: 'a"b', nodeIds: [], ...base } as Intent)).toContain("bad entry");
    expect(validateIntent({ kind: "insertIndexEntry", runId: 1, entry: "Widgets", nodeIds: [], ...base } as Intent)).toBeNull();
    expect(validateIntent({ kind: "insertIndex", runId: 1, entryCount: 0, nodeIds: [], ...base } as Intent)).toContain("bad entryCount");
    expect(validateIntent({ kind: "refreshIndex", entryCount: 20000, nodeIds: [], ...base } as Intent)).toContain("bad entryCount");
  });
});
