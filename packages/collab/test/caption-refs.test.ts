import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, operationBody, serializeXml, localName, tocEntryCount } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientReplica } from "../src/replica.js";
import type { Intent } from "../src/intents.js";

function docBytes(body: string): Uint8Array {
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

const para = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const captionP = (label: string, n: number, tail: string) =>
  `<w:p><w:r><w:t xml:space="preserve">${label} </w:t></w:r>` +
  `<w:fldSimple w:instr=" SEQ ${label} \\* ARABIC "><w:r><w:t>${n}</w:t></w:r></w:fldSimple>` +
  `<w:r><w:t xml:space="preserve"> ${tail}</w:t></w:r></w:p>`;

function blockIdOf(session: DocumentSession, idx: number): number {
  const body = session.doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const p = body.children.filter((c) => localName(c.name) === "p")[idx];
  return session.ids.idOf(p)!;
}

function firstRunIdOf(session: DocumentSession, paraIdx: number): number {
  const body = session.doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const p = body.children.filter((c) => localName(c.name) === "p")[paraIdx];
  const r = p.children.find((c) => localName(c.name) === "r")!;
  return session.ids.idOf(r)!;
}

function roundTrip(server: DocumentSession, replicas: ClientReplica[], author: ClientReplica, intent: Intent): void {
  author.submitLocal(intent);
  const entry = server.submit(intent);
  for (const c of replicas) c.receive([entry]);
}

let nextCarried = 700000;
const alloc = (n: number) => Array.from({ length: n }, () => nextCarried++);

describe("captions and cross-reference plumbing on the wire", () => {
  it("insertCaption converges and numbers deterministically", () => {
    const initial = docBytes(captionP("Figure", 1, "old") + para("anchor"));
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);

    const intent: Intent = {
      ...operationBody("insertCaption", blockIdOf(server, 1), { label: "Figure", text: "fresh", position: "below" as const }, alloc),
      clientId: "a", clientSeq: 1, base: 0,
    } as Intent;
    roundTrip(server, [a, b], a, intent);

    const serverXml = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(b.doc.docRoot)).toBe(serverXml);
    expect(serverXml).toContain(`SEQ Figure`);
    expect(serverXml).toContain("fresh");
    expect(serverXml).toContain(`<w:pStyle w:val="Caption"/>`);
  });

  it("ensureRefBookmark + insertCrossRef land a live REF on every replica", () => {
    const initial = docBytes(para("Chapter One") + para("see the chapter"));
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);

    const wrap: Intent = {
      ...operationBody("ensureRefBookmark", blockIdOf(server, 0), { name: "_Ref100000000" }),
      clientId: "a", clientSeq: 1, base: 0,
    } as Intent;
    roundTrip(server, [a, b], a, wrap);

    const ref: Intent = {
      kind: "insertCrossRef", runId: firstRunIdOf(server, 1), bookmark: "_Ref100000000",
      refKind: "text", nodeIds: alloc(8),
      clientId: "a", clientSeq: 2, base: server.seq,
    } as Intent;
    roundTrip(server, [a, b], a, ref);

    const serverXml = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(b.doc.docRoot)).toBe(serverXml);
    expect(serverXml).toContain(`w:name="_Ref100000000"`);
    expect(serverXml).toContain("REF _Ref100000000");

    // A second wrap of the same paragraph is a clean rejection everywhere.
    const again: Intent = {
      ...operationBody("ensureRefBookmark", blockIdOf(server, 0), { name: "_Ref100000001" }),
      clientId: "b", clientSeq: 1, base: server.seq,
    } as Intent;
    roundTrip(server, [a, b], b, again);
    expect(serializeXml(a.doc.docRoot)).toBe(serializeXml(server.doc.docRoot));
    expect(serializeXml(server.doc.docRoot)).not.toContain("_Ref100000001");
  });

  it("insertToc with captionLabel builds the same table of figures everywhere", () => {
    const initial = docBytes(captionP("Figure", 1, "alpha") + captionP("Figure", 2, "beta") + para("caret"));
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);

    const entryCount = tocEntryCount(a.doc, { captionLabel: "Figure" });
    expect(entryCount).toBe(2);
    const intent: Intent = {
      ...operationBody("insertToc", firstRunIdOf(server, 2), { entryCount, captionLabel: "Figure" }, alloc),
      clientId: "a", clientSeq: 1, base: 0,
    } as Intent;
    roundTrip(server, [a, b], a, intent);

    const serverXml = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(b.doc.docRoot)).toBe(serverXml);
    expect(serverXml).toContain(`TOC \\h \\z \\c "Figure"`);
    expect(serverXml).toContain("Figure 1 alpha");
    expect(serverXml).toContain("Figure 2 beta");
  });

  it("rejects malformed payloads at validation", async () => {
    const { validateIntent } = await import("../src/validate.js");
    const base = { clientId: "a", clientSeq: 1, base: 0 };
    expect(validateIntent({ kind: "insertCaption", blockId: 1, label: 'Fig" \\h', nodeIds: [], ...base } as Intent)).toContain("bad label");
    expect(validateIntent({ kind: "insertCaption", blockId: 1, label: "Figure", position: "left", nodeIds: [], ...base } as Intent)).toContain("bad position");
    expect(validateIntent({ kind: "ensureRefBookmark", blockId: 1, name: "NotHidden", ...base } as Intent)).toContain("bad name");
    expect(validateIntent({ kind: "insertToc", runId: 1, entryCount: 1, captionLabel: 'Fig"', nodeIds: [], ...base } as Intent)).toContain("bad captionLabel");
  });
});
