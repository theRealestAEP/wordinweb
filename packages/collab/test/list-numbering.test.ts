import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, layoutDocument, serializeXml, localName, attr, type Paragraph, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientReplica } from "../src/replica.js";
import { validateIntent } from "../src/validate.js";
import type { Intent } from "../src/intents.js";

/**
 * Multi-paragraph setListType (task #104): a toolbar toggle over a selection
 * is ONE intent carrying every target paragraph (blockId + moreBlockIds), so
 * a single apply mints ONE shared numbering definition on every replica.
 *
 * The old emission — one intent per paragraph — minted a FRESH definition per
 * apply: the originating client (one local mutation over all targets) had
 * numId 1,1,1 while the server and every peer had 1,2,3. Byte divergence AND
 * wrong semantics (numbered lists restarted at 1 on every paragraph).
 */

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

function blockIdOf(session: DocumentSession, paraIdx: number): number {
  const para = session.doc.sections[0].blocks[paraIdx] as Paragraph;
  return session.ids.idOf(para.src!)!;
}

/** numId of each body paragraph (null when the paragraph is not a list item),
 * read straight from the XML tree. */
function numIdsOf(doc: DocxDocument): (number | null)[] {
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const find = (el: XmlElement, name: string): XmlElement | undefined =>
    el.children.find((c) => localName(c.name) === name);
  return body.children
    .filter((c) => localName(c.name) === "p")
    .map((p) => {
      const numId = find(find(find(p, "pPr") ?? p, "numPr") ?? p, "numId");
      if (!numId) return null;
      const v = parseInt(attr(numId, "val") ?? "", 10);
      return Number.isFinite(v) ? v : null;
    });
}

/** Count of w:num instances in the numbering part (one per minted list). */
function numDefinitionCount(doc: DocxDocument): number {
  const tree = doc.numberingTree();
  if (!tree) return 0;
  return tree.children.filter((c) => localName(c.name) === "num").length;
}

const toggle = (
  clientId: string, clientSeq: number, base: number,
  blockId: number, moreBlockIds: number[], listKind: "bullet" | "number" | null,
): Intent => ({
  kind: "setListType", clientId, clientSeq, base, blockId, listKind,
  ...(moreBlockIds.length ? { moreBlockIds } : {}),
} as Intent);

describe("multi-paragraph setListType mints one shared numbering definition", () => {
  it("one intent numbers three paragraphs continuously (1. 2. 3., single numId)", () => {
    const session = new DocumentSession(DocxDocument.load(docBytes(["alpha", "beta", "gamma"])));
    const [b0, b1, b2] = [0, 1, 2].map((i) => blockIdOf(session, i));
    const entry = session.submit(toggle("a", 1, 0, b0, [b1, b2], "number"));
    expect(entry.kind).toBe("applied");

    // One shared definition, not one per paragraph.
    const ids = numIdsOf(session.doc);
    expect(ids[0]).not.toBeNull();
    expect(ids).toEqual([ids[0], ids[0], ids[0]]);
    expect(numDefinitionCount(session.doc)).toBe(1);

    // Word's continuity: the labels run 1. 2. 3. — no per-paragraph restart.
    const labels = layoutDocument(session.doc).pages.flatMap((page) => page.items)
      .filter((item) => item.kind === "text" && /^\d+\.$/.test(item.text))
      .map((item) => (item.kind === "text" ? item.text : ""));
    expect(labels).toEqual(["1.", "2.", "3."]);
  });

  it("originating client, server, and a second replica converge byte-for-byte", () => {
    const initial = docBytes(["alpha", "beta", "gamma"]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const author = new ClientReplica(initial);
    const peer = new ClientReplica(initial);
    const [b0, b1, b2] = [0, 1, 2].map((i) => blockIdOf(server, i));

    const intent = toggle("a", 1, 0, b0, [b1, b2], "number");
    author.submitLocal(intent);
    const entry = server.submit(intent);
    author.receive([entry]);
    peer.receive([entry]);

    const serverXml = serializeXml(server.doc.docRoot);
    expect(serializeXml(author.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(peer.doc.docRoot)).toBe(serverXml);
    expect(numIdsOf(server.doc)).toEqual([numIdsOf(server.doc)[0], numIdsOf(server.doc)[0], numIdsOf(server.doc)[0]]);
  });

  it("interleaved with a remote mint: the pending toggle rebases and still converges", () => {
    // B's concurrent toggle (sequenced FIRST) mints a definition, shifting the
    // numId A's optimistic apply chose. A's reconciliation must restore, apply
    // B's canonical entry, and replay its own (identity-transformed) intent so
    // the re-mint lands on the same numId every other replica computes.
    const initial = docBytes(["alpha", "beta", "gamma", "delta"]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);
    const [b0, b1, b2, b3] = [0, 1, 2, 3].map((i) => blockIdOf(server, i));

    const aToggle = toggle("a", 1, 0, b0, [b1, b2], "number");
    const bToggle = toggle("b", 1, 0, b3, [], "bullet");
    a.submitLocal(aToggle);
    b.submitLocal(bToggle);

    // Server sequences B first: A's echo applies onto a tree whose numbering
    // part already holds B's definition.
    const eB = server.submit(bToggle);
    const eA = server.submit(aToggle);
    a.receive([eB]); // remote interleaves A's pending -> rollback-replay
    a.receive([eA]);
    b.receive([eB, eA]);

    const serverXml = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(b.doc.docRoot)).toBe(serverXml);

    // Two mints total (A's shared one + B's), and A's three paragraphs still
    // share ONE numId distinct from B's.
    expect(numDefinitionCount(server.doc)).toBe(2);
    const ids = numIdsOf(server.doc);
    expect(ids[0]).not.toBeNull();
    expect(ids[1]).toBe(ids[0]);
    expect(ids[2]).toBe(ids[0]);
    expect(ids[3]).not.toBeNull();
    expect(ids[3]).not.toBe(ids[0]);
  });

  it("the old single-paragraph wire form (no moreBlockIds) still applies unchanged", () => {
    const session = new DocumentSession(DocxDocument.load(docBytes(["alpha", "beta"])));
    const entry = session.submit(toggle("a", 1, 0, blockIdOf(session, 0), [], "bullet"));
    expect(entry.kind).toBe("applied");
    const ids = numIdsOf(session.doc);
    expect(ids[0]).not.toBeNull();
    expect(ids[1]).toBeNull();
    expect(numDefinitionCount(session.doc)).toBe(1);
  });

  it("an unresolvable extra id is skipped; the resolvable targets still format", () => {
    const session = new DocumentSession(DocxDocument.load(docBytes(["alpha", "beta"])));
    const entry = session.submit(toggle("a", 1, 0, blockIdOf(session, 0), [99999, blockIdOf(session, 1)], "number"));
    expect(entry.kind).toBe("applied");
    const ids = numIdsOf(session.doc);
    expect(ids[0]).not.toBeNull();
    expect(ids[1]).toBe(ids[0]);
  });

  it("rejects malformed moreBlockIds before sequencing work", () => {
    expect(validateIntent(toggle("a", 1, 0, 1, [-1], "number"))).toBe("setListType: bad moreBlockIds");
    expect(validateIntent({ kind: "setListType", clientId: "a", clientSeq: 1, base: 0, blockId: 1, listKind: "number", moreBlockIds: "nope" } as unknown as Intent))
      .toBe("setListType: bad moreBlockIds");
    expect(validateIntent(toggle("a", 1, 0, 1, Array.from({ length: 10_001 }, (_, i) => i), "number")))
      .toBe("setListType: bad moreBlockIds");
    expect(validateIntent(toggle("a", 1, 0, 1, [2, 3], "number"))).toBeNull();
  });

  it("clearing list formatting across a selection converges too (listKind null)", () => {
    const initial = docBytes(["alpha", "beta"]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const author = new ClientReplica(initial);
    const [b0, b1] = [0, 1].map((i) => blockIdOf(server, i));

    const on = toggle("a", 1, 0, b0, [b1], "bullet");
    author.submitLocal(on);
    const e1 = server.submit(on);
    author.receive([e1]);

    const off = toggle("a", 2, server.seq, b0, [b1], null);
    author.submitLocal(off);
    const e2 = server.submit(off);
    author.receive([e2]);

    expect(serializeXml(author.doc.docRoot)).toBe(serializeXml(server.doc.docRoot));
    expect(numIdsOf(server.doc)).toEqual([null, null]);
  });
});
