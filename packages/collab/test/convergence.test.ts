import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type Run, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientReplica } from "../src/replica.js";
import { InsertTextIntent } from "../src/intents.js";

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

function paraText(doc: DocxDocument): string {
  const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
  const collectT = (el: XmlElement): string => {
    let s = localName(el.name) === "t" ? el.text : "";
    for (const c of el.children) s += collectT(c);
    return s;
  };
  return body.children.filter((c) => localName(c.name) === "p").map(collectT).join("\n");
}

/** Ids as the server currently sees them for paragraph 0's first run. */
function runAddr(session: DocumentSession, paraIdx = 0): { blockId: number; runId: number; len: number } {
  const para = session.doc.sections[paraIdx === 0 ? 0 : 0].blocks[paraIdx] as Paragraph;
  const run = para.children[0] as Run;
  const t = run.content.find((c) => c.kind === "text")!.srcT!;
  return { blockId: session.ids.idOf(para.src!)!, runId: session.ids.idOf(run.src!)!, len: t.text.length };
}

describe("multi-client convergence (optimistic apply + reconciliation)", () => {
  it("three clients editing the same run from the same base all converge to the server state", () => {
    const initial = docBytes(["seed"]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const clients = [new ClientReplica(initial), new ClientReplica(initial), new ClientReplica(initial)];
    const addr = runAddr(server); // same ids on every replica (deterministic parse order)

    // Each client optimistically types a distinct marker at the run start,
    // all from base 0, before any broadcast is delivered (true concurrency).
    const intents: InsertTextIntent[] = clients.map((_, i) => ({
      kind: "insertText", clientId: String.fromCharCode(97 + i), clientSeq: 1, base: 0,
      at: { blockId: addr.blockId, runId: addr.runId, offset: 0 }, text: `${i}`,
    }));
    clients.forEach((c, i) => c.submitLocal(intents[i]));

    // Server sequences them (order a,b,c); collect the canonical broadcasts.
    const broadcasts = intents.map((i) => server.submit(i));

    // Every client receives the full canonical log and reconciles.
    for (const c of clients) c.receive(broadcasts);

    // All replicas converge to identical XML == server state.
    const serverXml = serializeXml(server.doc.docRoot);
    for (const c of clients) {
      expect(serializeXml(c.doc.docRoot)).toBe(serverXml);
    }
    // And each client's own marker survived somewhere in the text.
    const text = paraText(server.doc);
    expect(text).toContain("0");
    expect(text).toContain("1");
    expect(text).toContain("2");
    expect(text.length).toBe("seed".length + 3);
  });

  it("converges across multiple rounds of concurrent single edits", () => {
    const initial = docBytes(["x"]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const clients = [new ClientReplica(initial), new ClientReplica(initial)];

    for (let round = 0; round < 4; round++) {
      const addr = runAddr(server);
      const intents: InsertTextIntent[] = clients.map((_, i) => ({
        kind: "insertText", clientId: `c${i}`, clientSeq: round + 1, base: server.seq,
        at: { blockId: addr.blockId, runId: addr.runId, offset: i === 0 ? 0 : addr.len }, text: i === 0 ? "L" : "R",
      }));
      clients.forEach((c, i) => c.submitLocal(intents[i]));
      const broadcasts = intents.map((i) => server.submit(i));
      for (const c of clients) c.receive(broadcasts);
    }

    const serverXml = serializeXml(server.doc.docRoot);
    for (const c of clients) expect(serializeXml(c.doc.docRoot)).toBe(serverXml);
  });
});

import { SplitParagraphIntent } from "../src/intents.js";

describe("sidecar: split-created ids survive a reconciliation restore (round-2 F1)", () => {
  it("a client that split then hits a remote interleave still converges", () => {
    const initial = docBytes(["HelloWorld"]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);
    const addr = runAddr(server);

    // Client A splits "HelloWorld" -> "Hello" | "World" (carried ids 6000/6001).
    const split: SplitParagraphIntent = {
      kind: "splitParagraph", clientId: "a", clientSeq: 1, base: 0,
      at: { blockId: addr.blockId, runId: addr.runId, offset: 5 }, newBlockId: 6000, newRunId: 6001,
    };
    a.submitLocal(split);
    // Client B concurrently inserts "!" at the end of the (still whole) run.
    const bIns = { kind: "insertText" as const, clientId: "b", clientSeq: 1, base: 0, at: { blockId: addr.blockId, runId: addr.runId, offset: addr.len }, text: "!" };
    b.submitLocal(bIns);

    // Server order: split then insert. B's insert offset 10 remaps past the
    // split point into the new run 6001.
    const e1 = server.submit(split);
    const e2 = server.submit(bIns);
    a.receive([e1, e2]);
    b.receive([e1, e2]);

    // A had a local split in its history; its confirmed-restore during
    // reconciliation must not lose the carried ids. All three converge.
    const serverXml = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot)).toBe(serverXml);
    expect(serializeXml(b.doc.docRoot)).toBe(serverXml);
  });
});
