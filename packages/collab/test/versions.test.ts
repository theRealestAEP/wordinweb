import { describe, expect, it } from "vitest";
import { VersionRing, InMemoryVersionStore } from "../src/versions.js";
import { CollabConnection } from "../src/connection.js";
import { DocumentSession } from "../src/session.js";
import { DocxDocument } from "@wordinweb/core";
import { zipSync, strToU8 } from "fflate";
import type { ClientMessage, ServerMessage } from "../src/protocol.js";

/** Version ring retention + capture (plan doc 14 §1, round-4 F15). */

function blankDocx(): Uint8Array {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
    `<w:p><w:r><w:t xml:space="preserve">hi</w:t></w:r></w:p></w:body></w:document>`;
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

function connectedConn(): { conn: CollabConnection; session: DocumentSession } {
  const session = new DocumentSession(DocxDocument.load(blankDocx()));
  const peer = { deliver: (_m: ServerMessage) => {} };
  const transport = {
    send: (msg: ClientMessage) => {
      if (msg.t === "hello") {
        const cp = session.checkpoint();
        peer.deliver({ t: "welcome", docId: "d", seq: cp.seq, snapshot: Buffer.from(cp.docx).toString("base64"),
          sidecar: cp.sidecar, tail: [], genesisId: "g1", mode: "plaintext" });
      } else if (msg.t === "submit") {
        peer.deliver({ t: "broadcast", entries: [session.submit(msg.intent)] });
      }
    },
    onMessage: (cb: (m: ServerMessage) => void) => { peer.deliver = cb; },
  };
  const conn = new CollabConnection(transport, "alice");
  conn.join("d");
  return { conn, session };
}

describe("VersionRing (doc 14 §1)", () => {
  it("auto ring keeps the newest N; labeled versions are pinned outside the ring", async () => {
    const { conn } = connectedConn();
    const store = new InMemoryVersionStore();
    let now = 0;
    const ring = new VersionRing(store, { autoCap: 3, now: () => ++now });
    await ring.capture(conn, "d", "milestone"); // labeled — never ring-evicted
    for (let i = 0; i < 6; i++) await ring.capture(conn, "d"); // 6 autos through a 3-ring
    const all = await store.list("d");
    expect(all.filter((v) => v.auto)).toHaveLength(3); // ring held
    expect(all.filter((v) => !v.auto)).toHaveLength(1); // label survived 6 auto churns
    // Ring kept the NEWEST three.
    const autoTimes = all.filter((v) => v.auto).map((v) => v.savedAt);
    expect(Math.min(...autoTimes)).toBeGreaterThan(4);
  });

  it("labeled cap evicts oldest-labeled AND reports it (no silent caps)", async () => {
    const { conn } = connectedConn();
    const store = new InMemoryVersionStore();
    let now = 0;
    const ring = new VersionRing(store, { labeledCap: 2, now: () => ++now });
    await ring.capture(conn, "d", "first");
    await ring.capture(conn, "d", "second");
    const third = await ring.capture(conn, "d", "third");
    expect(third.evicted).toHaveLength(1);
    expect(third.evicted[0].label).toBe("first"); // oldest labeled went, reported
    expect((await store.list("d")).filter((v) => !v.auto)).toHaveLength(2);
  });

  it("captures CONFIRMED state with the sidecar (restorable as a working seed)", async () => {
    const { conn, session } = connectedConn();
    const store = new InMemoryVersionStore();
    const ring = new VersionRing(store, { now: () => 1 });
    const { saved } = await ring.capture(conn, "d", "restore-point");
    expect(saved).toBeTruthy();
    expect(saved!.genesisId).toBe("g1");
    expect(saved!.sidecar).toBeTruthy();
    // The frozen docx equals the session's canonical bytes (F1-safe seed).
    expect(Buffer.from(saved!.docx).equals(Buffer.from(session.checkpoint().docx))).toBe(true);
  });

  it("quota relief evicts oldest-auto before labeled (F15 ordering)", async () => {
    const { conn } = connectedConn();
    const store = new InMemoryVersionStore();
    let now = 0;
    const ring = new VersionRing(store, { now: () => ++now });
    await ring.capture(conn, "d", "precious");
    await ring.capture(conn, "d");
    await ring.capture(conn, "d");
    const dropped = await ring.evictOldest("d");
    expect(dropped[0].auto).toBe(true); // an auto went first
    const dropped2 = await ring.evictOldest("d");
    expect(dropped2[0].auto).toBe(true); // the other auto next
    const dropped3 = await ring.evictOldest("d");
    expect(dropped3[0].label).toBe("precious"); // labeled only when nothing else left
  });
});

describe("activity attribution L1 (doc 14 §3)", () => {
  it("derives {seq, clientId, kind} from applied entries, bounded, in order", () => {
    const { conn } = connectedConn();
    conn.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "a" } as never);
    conn.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 1 }, text: "b" } as never);
    expect(conn.activity).toEqual([
      { seq: 1, clientId: "alice", kind: "insertText" },
      { seq: 2, clientId: "alice", kind: "insertText" },
    ]);
  });
});
