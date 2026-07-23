import { describe, expect, it } from "vitest";
import { CollabHub, Connection } from "../src/hub.js";
import { PROTOCOL_VERSION, ENGINE_VERSION } from "@wordinweb/collab/server";
import type { ServerMessage, IntentEnvelope } from "@wordinweb/collab/server";

/** Hub as blind sequencer (doc 13 §2): encrypted rooms, engine fence,
 * cross-mode refusals, dedup on plaintext bookkeeping. */

class FakeConn implements Connection {
  received: ServerMessage[] = [];
  constructor(public id: string) {}
  send(msg: ServerMessage): void {
    this.received.push(msg);
  }
  last(): ServerMessage {
    for (let i = this.received.length - 1; i >= 0; i--) {
      if (this.received[i].t !== "roster") return this.received[i];
    }
    return this.received[this.received.length - 1];
  }
}

const CP = { seq: 0, iv: "aXY=", ciphertext: "b3BhcXVl" }; // opaque to the server — any bytes do
const env = (clientId: string, clientSeq: number): IntentEnvelope => ({
  clientId,
  clientSeq,
  base: 0,
  iv: "aXY=",
  ciphertext: "Y2lwaGVy",
});
const hello = (clientId: string, over: Record<string, unknown> = {}) =>
  ({ t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId, sinceSeq: 0, engineVersion: ENGINE_VERSION, ...over }) as never;

const encHub = () => {
  const hub = new CollabHub(null); // zero-custody: encrypted docs are magic-link tier
  expect(hub.seedEncrypted("d", "g1", CP).ok).toBe(true);
  return hub;
};

describe("CollabHub encrypted rooms (doc 13 §2/§3)", () => {
  it("seedEncrypted is first-wins like plaintext seed", () => {
    const hub = encHub();
    const second = hub.seedEncrypted("d", "g2", CP);
    expect(second.ok).toBe(false);
    if (!second.ok) expect(second.genesisId).toBe("g1");
  });

  it("welcome-enc carries the sealed checkpoint + tail; the server never parses anything", async () => {
    const hub = encHub();
    const a = new FakeConn("a");
    await hub.handle(a, hello("alice"));
    const w = a.last();
    expect(w.t).toBe("welcome-enc");
    if (w.t !== "welcome-enc") return;
    expect(w.genesisId).toBe("g1");
    expect(w.checkpoint).toEqual(CP); // byte-for-byte what the seeder sent
    expect(w.tail).toEqual([]);
    expect(w.mode).toBe("encrypted");
  });

  it("engine fence: first joiner registers; a mismatched second is refused; missing is refused", async () => {
    const hub = encHub();
    const a = new FakeConn("a");
    await hub.handle(a, hello("alice"));
    expect(a.last().t).toBe("welcome-enc");
    const b = new FakeConn("b");
    await hub.handle(b, hello("bob", { engineVersion: "e999" }));
    expect(b.last()).toEqual({ t: "refused", reason: "engine-version-mismatch" });
    const c = new FakeConn("c");
    await hub.handle(c, hello("carol", { engineVersion: undefined }));
    expect(c.last()).toEqual({ t: "refused", reason: "engine-version-required" });
  });

  it("sequences envelopes in order, dedups by plaintext bookkeeping, fans out", async () => {
    const hub = encHub();
    const a = new FakeConn("a");
    const b = new FakeConn("b");
    await hub.handle(a, hello("alice"));
    await hub.handle(b, hello("bob"));
    await hub.handle(a, { t: "submit-enc", envelope: env("alice", 1) });
    await hub.handle(b, { t: "submit-enc", envelope: env("bob", 1) });
    await hub.handle(a, { t: "submit-enc", envelope: env("alice", 1) }); // re-send
    const bcasts = b.received.filter((m) => m.t === "broadcast-enc");
    expect(bcasts).toHaveLength(3);
    const seqs = bcasts.flatMap((m) => (m.t === "broadcast-enc" ? m.entries.map((e) => e.seq) : []));
    expect(seqs).toEqual([1, 2, 1]); // the re-send returns the ORIGINAL entry (dedup)
  });

  it("identity binding holds in encrypted mode (forged envelope clientId refused)", async () => {
    const hub = encHub();
    const a = new FakeConn("a");
    await hub.handle(a, hello("mallory"));
    await hub.handle(a, { t: "submit-enc", envelope: env("alice", 1) });
    expect(a.last()).toEqual({ t: "refused", reason: "client-id-mismatch" });
  });

  it("no mixed-mode documents: plaintext submit into an encrypted room (and vice versa) refused", async () => {
    const hub = encHub();
    const a = new FakeConn("a");
    await hub.handle(a, hello("alice"));
    await hub.handle(a, {
      t: "submit",
      intent: { kind: "insertText", clientId: "alice", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 0 }, text: "x" } as never,
    });
    expect(a.last()).toEqual({ t: "refused", reason: "wrong-mode" });

    // Vice versa on a plaintext room.
    const hub2 = new CollabHub(null);
    hub2.seed("p", plaintextDoc());
    const c = new FakeConn("c");
    await hub2.handle(c, hello("carol", { docId: "p" }));
    await hub2.handle(c, { t: "submit-enc", envelope: env("carol", 1) });
    expect(c.last()).toEqual({ t: "refused", reason: "wrong-mode" });
  });

  it("oversized ciphertext is refused (the size cap a blind server CAN enforce)", async () => {
    const hub = encHub();
    const a = new FakeConn("a");
    await hub.handle(a, hello("alice"));
    await hub.handle(a, { t: "submit-enc", envelope: { ...env("alice", 1), ciphertext: "A".repeat(300 * 1024) } });
    expect(a.last()).toEqual({ t: "refused", reason: "too-large" });
  });
});

import { zipSync, strToU8 } from "fflate";
function plaintextDoc(): Uint8Array {
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
