import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type Paragraph, type Run, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { ClientReplica } from "../src/replica.js";
import { InsertTextIntent } from "../src/intents.js";

// --- helpers (mirrored from convergence.test.ts) ---
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
function runAddr(session: DocumentSession): { blockId: number; runId: number; len: number } {
  const para = session.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  const t = run.content.find((c) => c.kind === "text")!.srcT!;
  return { blockId: session.ids.idOf(para.src!)!, runId: session.ids.idOf(run.src!)!, len: t.text.length };
}

/**
 * The gap: real typing produces MANY unconfirmed (pending) ops before the
 * server confirms any — you type "HELLO" and 5 ops fly with the same base
 * before a single echo returns. Every existing convergence test reconciles
 * after ONE op (the "one-in-flight" model), so the multi-pending rebase — the
 * path real typing actually takes — was never exercised. These reproduce the
 * scramble/divergence users hit when two people type at once.
 */
describe("typing stress: many pending ops per client (real-typing model)", () => {
  it("one client typing several chars with NO confirmations converges to the typed word", () => {
    const initial = docBytes(["S"]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const addr = runAddr(server);
    // Type "ABC" appended after S: offsets 1,2,3 in the client's optimistic
    // view, ALL based on seq 0 (no echo yet) — three pending ops.
    const ins = [..."ABC"].map((ch, i): InsertTextIntent => ({
      kind: "insertText", clientId: "a", clientSeq: i + 1, base: 0,
      at: { blockId: addr.blockId, runId: addr.runId, offset: 1 + i }, text: ch,
    }));
    ins.forEach((x) => a.submitLocal(x));
    const bc = ins.map((x) => server.submit(x));
    a.receive(bc);
    expect(paraText(server.doc)).toBe("SABC");
    expect(serializeXml(a.doc.docRoot)).toBe(serializeXml(server.doc.docRoot));
  });

  it("TWO clients each type a word (many pending), interleaved arrival — all converge, no data loss", () => {
    const initial = docBytes(["S"]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);
    const addr = runAddr(server);

    // Each appends its 3 chars after "S": offsets 1,2,3, all base 0. Neither has
    // seen the other's edits (true simultaneous typing).
    const aI = [..."AAA"].map((ch, i): InsertTextIntent => ({
      kind: "insertText", clientId: "a", clientSeq: i + 1, base: 0,
      at: { blockId: addr.blockId, runId: addr.runId, offset: 1 + i }, text: ch,
    }));
    const bI = [..."BBB"].map((ch, i): InsertTextIntent => ({
      kind: "insertText", clientId: "b", clientSeq: i + 1, base: 0,
      at: { blockId: addr.blockId, runId: addr.runId, offset: 1 + i }, text: ch,
    }));
    aI.forEach((x) => a.submitLocal(x));
    bI.forEach((x) => b.submitLocal(x));

    // Server sees them INTERLEAVED as they arrive over the wire.
    const order = [aI[0], bI[0], aI[1], bI[1], aI[2], bI[2]];
    const bc = order.map((x) => server.submit(x));
    // ONE envelope per broadcast message — exactly how the live connection
    // ingests (enc-connection.ingest → replica.receive([entry]) per frame),
    // NOT a single batch. This is the path real typing takes.
    for (const e of bc) { a.receive([e]); b.receive([e]); }

    const sx = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot), "client A must match the server").toBe(sx);
    expect(serializeXml(b.doc.docRoot), "client B must match the server").toBe(sx);
    const text = paraText(server.doc);
    // No data loss: the seed + all six typed chars survive.
    expect(text.length, `expected 7 chars, got "${text}"`).toBe(7);
    expect((text.match(/A/g) ?? []).length, `A's chars lost: "${text}"`).toBe(3);
    expect((text.match(/B/g) ?? []).length, `B's chars lost: "${text}"`).toBe(3);
  });

  it("randomized: many clients, deep pending queues, random interleaving — always converge, never lose a char", () => {
    // Deterministic LCG (no Math.random — reproducible).
    const lcg = (seed: number) => { let s = seed >>> 0; return () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 0x100000000; }; };
    for (const seed of [1, 7, 42, 99, 2024, 31337]) {
      const rand = lcg(seed);
      const initial = docBytes(["S"]);
      const server = new DocumentSession(DocxDocument.load(initial));
      const clients = [new ClientReplica(initial), new ClientReplica(initial), new ClientReplica(initial)];
      const seqs = [0, 0, 0];
      const perClientChar = ["a", "b", "c"];
      const expected = { a: 0, b: 0, c: 0 } as Record<string, number>;
      const inflight: ReturnType<DocumentSession["submit"]>[] = [];

      for (let step = 0; step < 40; step++) {
        // A random client types a char at its OWN current end (append) — often
        // BEFORE its previous ops are confirmed, so pending queues run deep.
        const ci = Math.floor(rand() * clients.length);
        const c = clients[ci];
        const ch = perClientChar[ci];
        const end = paraText(c.doc).split("\n")[0].length;
        const intent: InsertTextIntent = {
          kind: "insertText", clientId: `c${ci}`, clientSeq: ++seqs[ci], base: c.confirmedSeq,
          at: { blockId: runAddr(server).blockId, runId: runAddr(server).runId, offset: end }, text: ch,
        };
        c.submitLocal(intent);
        expected[ch]++;
        inflight.push(server.submit(intent));
        // Flush the server's broadcast backlog to ALL clients at random times,
        // one envelope at a time (as the wire delivers), so pending drains
        // partially and interleaves with fresh local ops.
        if (rand() < 0.5 && inflight.length) {
          const batch = inflight.splice(0, 1 + Math.floor(rand() * inflight.length));
          for (const e of batch) for (const cl of clients) cl.receive([e]);
        }
      }
      for (const e of inflight) for (const cl of clients) cl.receive([e]);

      const sx = serializeXml(server.doc.docRoot);
      for (let i = 0; i < clients.length; i++) {
        expect(serializeXml(clients[i].doc.docRoot), `seed ${seed}: client c${i} diverged`).toBe(sx);
      }
      const text = paraText(server.doc);
      for (const ch of ["a", "b", "c"]) {
        expect((text.match(new RegExp(ch, "g")) ?? []).length, `seed ${seed}: '${ch}' lost/duplicated in "${text}"`).toBe(expected[ch]);
      }
    }
  });

  it("CONCURRENT STRUCTURAL: one client presses Enter (split) mid-burst while another types", () => {
    // The user's real case: multiple paragraphs (Enter = splitParagraph, a
    // STRUCTURAL edit) typed concurrently. Reproduce: A types "AB", Enter,
    // "CD" (a split with inserts on both sides, all pending); B types "XY"
    // concurrently. One-at-a-time receive. Must converge with no loss.
    const initial = docBytes(["S"]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);
    const A = runAddr(server);

    const aOps: Intent[] = [
      { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: A.blockId, runId: A.runId, offset: 1 }, text: "A" },
      { kind: "insertText", clientId: "a", clientSeq: 2, base: 0, at: { blockId: A.blockId, runId: A.runId, offset: 2 }, text: "B" },
      { kind: "splitParagraph", clientId: "a", clientSeq: 3, base: 0, at: { blockId: A.blockId, runId: A.runId, offset: 3 }, newBlockId: 9000, newRunId: 9001 },
      { kind: "insertText", clientId: "a", clientSeq: 4, base: 0, at: { blockId: 9000, runId: 9001, offset: 0 }, text: "C" },
      { kind: "insertText", clientId: "a", clientSeq: 5, base: 0, at: { blockId: 9000, runId: 9001, offset: 1 }, text: "D" },
    ];
    const bOps: Intent[] = [
      { kind: "insertText", clientId: "b", clientSeq: 1, base: 0, at: { blockId: A.blockId, runId: A.runId, offset: 1 }, text: "X" },
      { kind: "insertText", clientId: "b", clientSeq: 2, base: 0, at: { blockId: A.blockId, runId: A.runId, offset: 2 }, text: "Y" },
    ];
    aOps.forEach((o) => a.submitLocal(o));
    bOps.forEach((o) => b.submitLocal(o));

    const order = [aOps[0], bOps[0], aOps[1], bOps[1], aOps[2], aOps[3], aOps[4]];
    const bc = order.map((o) => server.submit(o));
    for (const e of bc) { a.receive([e]); b.receive([e]); }

    const sx = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot), "A vs server").toBe(sx);
    expect(serializeXml(b.doc.docRoot), "B vs server").toBe(sx);
    const text = paraText(server.doc);
    for (const ch of ["A", "B", "C", "D", "X", "Y"]) {
      expect((text.match(new RegExp(ch, "g")) ?? []).length, `'${ch}' lost/dup in "${text}"`).toBe(1);
    }
  });

  it("rapid back-and-forth: clients keep typing while echoes stream back", () => {
    const initial = docBytes(["S"]);
    const server = new DocumentSession(DocxDocument.load(initial));
    const a = new ClientReplica(initial);
    const b = new ClientReplica(initial);
    let aSeq = 0;
    let bSeq = 0;
    const pendingBroadcasts: ReturnType<DocumentSession["submit"]>[] = [];

    // 8 rounds; each round BOTH clients type a char based on the last confirmed
    // seq, but the server only flushes broadcasts to clients every other round —
    // so pending queues grow past one. Append at the current end each time.
    for (let round = 0; round < 8; round++) {
      const endA = (paraText(a.doc).split("\n")[0]).length;
      const endB = (paraText(b.doc).split("\n")[0]).length;
      const ai: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: ++aSeq, base: a.confirmedSeq, at: { blockId: runAddr(server).blockId, runId: runAddr(server).runId, offset: endA }, text: "a" };
      const bi: InsertTextIntent = { kind: "insertText", clientId: "b", clientSeq: ++bSeq, base: b.confirmedSeq, at: { blockId: runAddr(server).blockId, runId: runAddr(server).runId, offset: endB }, text: "b" };
      a.submitLocal(ai);
      b.submitLocal(bi);
      pendingBroadcasts.push(server.submit(ai), server.submit(bi));
      if (round % 2 === 1) {
        a.receive(pendingBroadcasts);
        b.receive(pendingBroadcasts);
        pendingBroadcasts.length = 0;
      }
    }
    if (pendingBroadcasts.length) { a.receive(pendingBroadcasts); b.receive(pendingBroadcasts); }

    const sx = serializeXml(server.doc.docRoot);
    expect(serializeXml(a.doc.docRoot)).toBe(sx);
    expect(serializeXml(b.doc.docRoot)).toBe(sx);
    const text = paraText(server.doc);
    expect((text.match(/a/g) ?? []).length, `A's chars lost: "${text}"`).toBe(8);
    expect((text.match(/b/g) ?? []).length, `B's chars lost: "${text}"`).toBe(8);
  });
});
