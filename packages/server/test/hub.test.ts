import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub, Connection, DocProvider } from "../src/hub.js";
import { ServerMessage, PROTOCOL_VERSION } from "@wordinweb/collab/server";
import type { InsertTextIntent } from "@wordinweb/collab/server";

function blankDoc(text: string): Uint8Array {
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

const provider: DocProvider = { load: () => blankDoc("hi") };

class FakeConn implements Connection {
  received: ServerMessage[] = [];
  constructor(public id: string) {}
  send(msg: ServerMessage): void {
    this.received.push(msg);
  }
  last(): ServerMessage {
    return this.received[this.received.length - 1];
  }
}

describe("CollabHub", () => {
  it("refuses a version-mismatched client at hello", async () => {
    const hub = new CollabHub(provider);
    const c = new FakeConn("c1");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION + 1, docId: "d", clientId: "c1", sinceSeq: 0 });
    expect(c.last()).toEqual({ t: "refused", reason: "version-mismatch" });
  });

  it("welcomes a joining client with a snapshot and the current seq", async () => {
    const hub = new CollabHub(provider);
    const c = new FakeConn("c1");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "c1", sinceSeq: 0 });
    const w = c.last();
    expect(w.t).toBe("welcome");
    if (w.t === "welcome") {
      expect(w.seq).toBe(0);
      expect(w.snapshot.length).toBeGreaterThan(0);
      expect(w.tail).toEqual([]);
    }
    expect(hub.activeDocs()).toEqual(["d"]);
  });

  it("refuses submit from a connection that has not joined", async () => {
    const hub = new CollabHub(provider);
    const c = new FakeConn("c1");
    const intent: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" };
    await hub.handle(c, { t: "submit", intent });
    expect(c.last()).toEqual({ t: "refused", reason: "not-joined" });
  });

  it("broadcasts a submitted intent to every participant in the room", async () => {
    const hub = new CollabHub(provider);
    const a = new FakeConn("a");
    const b = new FakeConn("b");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    await hub.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "b", sinceSeq: 0 });

    // Address paragraph 0's run using ids the session assigned (1 = para, 2 = run
    // in deterministic parse order for a single-paragraph doc).
    const intent: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" };
    await hub.handle(a, { t: "submit", intent });

    for (const c of [a, b]) {
      const msg = c.last();
      expect(msg.t).toBe("broadcast");
      if (msg.t === "broadcast") {
        expect(msg.entries).toHaveLength(1);
        expect(msg.entries[0].kind).toBe("applied");
        expect(msg.entries[0].seq).toBe(1);
      }
    }
  });

  it("isolates rooms: a submit on one doc does not reach another", async () => {
    const hub = new CollabHub(provider);
    const a = new FakeConn("a");
    const b = new FakeConn("b");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "docA", clientId: "a", sinceSeq: 0 });
    await hub.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "docB", clientId: "b", sinceSeq: 0 });
    const before = b.received.length;
    const intent: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" };
    await hub.handle(a, { t: "submit", intent });
    expect(b.received.length).toBe(before); // b got nothing new
    expect(hub.activeDocs().sort()).toEqual(["docA", "docB"]);
  });

  it("drops a connection on disconnect", async () => {
    const hub = new CollabHub(provider);
    const a = new FakeConn("a");
    const b = new FakeConn("b");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    await hub.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "b", sinceSeq: 0 });
    hub.disconnect(b);
    const beforeB = b.received.length;
    const intent: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" };
    await hub.handle(a, { t: "submit", intent });
    expect(b.received.length).toBe(beforeB); // disconnected b receives nothing
  });
});

import { attachWebSocketServer, WsServer, WsSocket } from "../src/ws.js";

describe("attachWebSocketServer", () => {
  it("bridges ws-style frames to the hub and serializes responses", async () => {
    let onConn: ((s: WsSocket) => void) | null = null;
    const wss: WsServer = { on: (_e, cb) => { onConn = cb as (s: WsSocket) => void; } };
    const hub = new CollabHub(provider);
    attachWebSocketServer(wss, hub);

    const sent: string[] = [];
    let onMsg: ((d: unknown) => void) | null = null;
    const socket: WsSocket = {
      send: (d) => sent.push(d),
      on: (e: "message" | "close", cb: ((d: unknown) => void) & (() => void)) => {
        if (e === "message") onMsg = cb;
      },
    };
    onConn!(socket);
    onMsg!(JSON.stringify({ t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "w", sinceSeq: 0 }));
    await new Promise((r) => setTimeout(r, 0)); // flush the fire-and-forget handle()
    expect(sent).toHaveLength(1);
    expect(JSON.parse(sent[0]).t).toBe("welcome");
  });

  it("ignores a malformed frame without crashing", async () => {
    let onConn: ((s: WsSocket) => void) | null = null;
    const wss: WsServer = { on: (_e, cb) => { onConn = cb as (s: WsSocket) => void; } };
    attachWebSocketServer(wss, new CollabHub(provider));
    let onMsg: ((d: unknown) => void) | null = null;
    const socket: WsSocket = { send: () => {}, on: (e, cb: ((d: unknown) => void) & (() => void)) => { if (e === "message") onMsg = cb; } };
    onConn!(socket);
    expect(() => onMsg!("{ not json")).not.toThrow();
  });
});

import { InMemoryStorage } from "../src/storage.js";

describe("CollabHub persistence + rehydration", () => {
  it("persists sequenced entries and rehydrates a new hub from storage", async () => {
    const storage = new InMemoryStorage();
    const hub1 = new CollabHub(provider, storage);
    const a = new FakeConn("a");
    await hub1.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    const intent: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" };
    await hub1.handle(a, { t: "submit", intent });

    // The entry is durably stored.
    const log = await storage.readLog("d", 0);
    expect(log).toHaveLength(1);
    expect(log[0].kind).toBe("applied");

    // A fresh hub (simulating a restart) rehydrates the doc from the log tail:
    // the joining client's welcome reports the persisted seq.
    const hub2 = new CollabHub(provider, storage);
    const b = new FakeConn("b");
    await hub2.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "b", sinceSeq: 0 });
    const w = b.last();
    expect(w.t).toBe("welcome");
    if (w.t === "welcome") expect(w.seq).toBe(1);
  });

  it("deduplicates a resent intent across the persisted log", async () => {
    const storage = new InMemoryStorage();
    const hub = new CollabHub(provider, storage);
    const a = new FakeConn("a");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    const intent: InsertTextIntent = { kind: "insertText", clientId: "a", clientSeq: 5, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "X" };
    await hub.handle(a, { t: "submit", intent });
    await hub.handle(a, { t: "submit", intent }); // resend
    const log = await storage.readLog("d", 0);
    expect(log).toHaveLength(1); // deduped
  });
});

import type { PresencePosition } from "@wordinweb/collab/server";

describe("CollabHub presence", () => {
  const pos: PresencePosition = { anchor: { blockId: 1, runId: 2, offset: 3 } };

  it("fans a presence update out to other participants but not the sender", async () => {
    const hub = new CollabHub(provider);
    const a = new FakeConn("a");
    const b = new FakeConn("b");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    await hub.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "b", sinceSeq: 0 });
    const aBefore = a.received.length;
    await hub.handle(a, { t: "presence", position: pos });
    expect(a.received.length).toBe(aBefore); // sender does not echo
    const msg = b.last();
    expect(msg.t).toBe("presence");
    if (msg.t === "presence") {
      expect(msg.participant).toBe("a");
      expect(msg.position).toEqual(pos);
    }
  });

  it("ignores presence from a connection that has not joined", async () => {
    const hub = new CollabHub(provider);
    const c = new FakeConn("c");
    await hub.handle(c, { t: "presence", position: pos });
    expect(c.received).toHaveLength(0); // ignored, not refused
  });

  it("does not persist presence (storage untouched)", async () => {
    const storage = new InMemoryStorage();
    const hub = new CollabHub(provider, storage);
    const a = new FakeConn("a");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    await hub.handle(a, { t: "presence", position: pos });
    expect(await storage.readLog("d", 0)).toHaveLength(0);
  });
});

import { blankDocxBytes, blankProvider } from "../src/blank.js";
import { DocxDocument } from "@wordinweb/core";

describe("blank document provider", () => {
  it("produces a loadable one-paragraph docx", () => {
    const doc = DocxDocument.load(blankDocxBytes());
    expect(doc.sections[0].blocks.length).toBe(1);
    // Round-trips (save is side-effect-free / valid).
    expect(doc.save().length).toBeGreaterThan(0);
  });

  it("blankProvider serves fresh bytes per docId", () => {
    const a = blankProvider.load("x");
    const b = blankProvider.load("y");
    expect(a.length).toBeGreaterThan(0);
    expect(b.length).toBeGreaterThan(0);
  });

  it("a hub using blankProvider welcomes a client with a blank doc", async () => {
    const hub = new CollabHub(blankProvider, new InMemoryStorage());
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "fresh", clientId: "c", sinceSeq: 0 });
    const w = c.last();
    expect(w.t).toBe("welcome");
    if (w.t === "welcome") expect(w.seq).toBe(0);
  });
});

import type { TokenVerifier } from "../src/hub.js";

describe("CollabHub auth (token verifier)", () => {
  const verifier: TokenVerifier = {
    verify: (token, docId) => {
      if (token === `edit-${docId}`) return { userId: "u", role: "editor" };
      if (token === `view-${docId}`) return { userId: "v", role: "viewer" };
      return null; // wrong doc / no token / bad token
    },
  };

  it("refuses a hello with no/invalid token when a verifier is set", async () => {
    const hub = new CollabHub(provider, undefined, verifier);
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "c", sinceSeq: 0 });
    expect(c.last()).toEqual({ t: "refused", reason: "unauthorized" });
  });

  it("refuses a token minted for a different doc (cross-tenant isolation)", async () => {
    const hub = new CollabHub(provider, undefined, verifier);
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "c", token: "edit-OTHER", sinceSeq: 0 });
    expect(c.last()).toEqual({ t: "refused", reason: "unauthorized" });
  });

  it("admits an editor and accepts their edits", async () => {
    const hub = new CollabHub(provider, undefined, verifier);
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "u", token: "edit-d", sinceSeq: 0 });
    expect(c.last().t).toBe("welcome");
    await hub.handle(c, { t: "submit", intent: { kind: "insertText", clientId: "u", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" } });
    expect(c.last().t).toBe("broadcast");
  });

  it("refuses a viewer's edit at the sequencer (read-only enforced server-side)", async () => {
    const hub = new CollabHub(provider, undefined, verifier);
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "v", token: "view-d", sinceSeq: 0 });
    expect(c.last().t).toBe("welcome");
    await hub.handle(c, { t: "submit", intent: { kind: "insertText", clientId: "v", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" } });
    expect(c.last()).toEqual({ t: "refused", reason: "read-only" });
  });
});

describe("CollabHub JWT live-expiry + revocation (gate 5)", () => {
  it("cuts a session whose token has expired at the next action", async () => {
    let now = 1000;
    const verifier: TokenVerifier = { verify: () => ({ userId: "u", role: "editor", expiresAt: 5000 }) };
    const hub = new CollabHub(provider, undefined, verifier, () => now);
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "u", token: "t", sinceSeq: 0 });
    expect(c.last().t).toBe("welcome");
    now = 6000; // token now expired
    await hub.handle(c, { t: "submit", intent: { kind: "insertText", clientId: "u", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" } });
    expect(c.last()).toEqual({ t: "refused", reason: "token-expired" });
  });

  it("sweepExpired kicks expired sessions proactively", async () => {
    let now = 1000;
    const verifier: TokenVerifier = { verify: () => ({ userId: "u", role: "editor", expiresAt: 2000 }) };
    const hub = new CollabHub(provider, undefined, verifier, () => now);
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "c", token: "t", sinceSeq: 0 });
    now = 3000;
    hub.sweepExpired();
    expect(c.last()).toEqual({ t: "refused", reason: "token-expired" });
  });

  it("revoke(userId) disconnects that user's live sessions immediately", async () => {
    const verifier: TokenVerifier = { verify: (tok) => tok === "alice" ? { userId: "alice", role: "editor" } : { userId: "bob", role: "editor" } };
    const hub = new CollabHub(provider, undefined, verifier);
    const a = new FakeConn("a");
    const b = new FakeConn("b");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "alice", token: "alice", sinceSeq: 0 });
    await hub.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "bob", token: "bob", sinceSeq: 0 });
    hub.revoke("alice");
    expect(a.last()).toEqual({ t: "refused", reason: "revoked" });
    // bob is untouched; his edit still broadcasts.
    const beforeB = b.received.length;
    await hub.handle(b, { t: "submit", intent: { kind: "insertText", clientId: "bob", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" } });
    expect(b.received.length).toBeGreaterThan(beforeB);
  });
});

import { ClientReplica } from "@wordinweb/collab/client";
import type { SplitParagraphIntent } from "@wordinweb/collab/server";
import { EVICTION_GRACE_MS } from "../src/hub.js";

function b64ToBytes(b64: string): Uint8Array {
  return new Uint8Array(Buffer.from(b64, "base64"));
}

describe("CollabHub identity binding (doc 11 decision 8 / round-4 F4)", () => {
  it("refuses a submit whose intent.clientId differs from the bound identity", async () => {
    const hub = new CollabHub(provider);
    const attacker = new FakeConn("atk");
    await hub.handle(attacker, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "mallory", sinceSeq: 0 });
    // Forgery: an intent claiming the victim's identity. Pre-fix this entered
    // the dedup set as (alice, 1) and swallowed alice's real next intent.
    const forged: InsertTextIntent = { kind: "insertText", clientId: "alice", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" };
    await hub.handle(attacker, { t: "submit", intent: forged });
    expect(attacker.last()).toEqual({ t: "refused", reason: "client-id-mismatch" });
    // The forgery never reached the session: alice's real (alice, 1) applies.
    const alice = new FakeConn("al");
    await hub.handle(alice, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "alice", sinceSeq: 0 });
    await hub.handle(alice, { t: "submit", intent: { ...forged, text: "hi" } });
    const msg = alice.last();
    expect(msg.t).toBe("broadcast");
    if (msg.t === "broadcast") expect(msg.entries[0].kind).toBe("applied");
  });

  it("refuses a hello with no clientId", async () => {
    const hub = new CollabHub(provider);
    const c = new FakeConn("c");
    // Cast: simulates an out-of-spec client omitting the field at runtime.
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", sinceSeq: 0 } as never);
    expect(c.last()).toEqual({ t: "refused", reason: "client-id-required" });
  });

  it("presence is keyed by the bound clientId, not the socket id (round-4 F14)", async () => {
    const hub = new CollabHub(provider);
    const a = new FakeConn("socket-1");
    const b = new FakeConn("socket-2");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "alice", sinceSeq: 0 });
    await hub.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "bob", sinceSeq: 0 });
    await hub.handle(a, { t: "presence", position: { anchor: { blockId: 1, runId: 2, offset: 0 } } });
    const msg = b.last();
    expect(msg.t).toBe("presence");
    if (msg.t === "presence") expect(msg.participant).toBe("alice"); // not "socket-1"
  });
});

describe("CollabHub single-live-connection rule (doc 12 §7)", () => {
  it("refuses a second connection binding the same (docId, clientId)", async () => {
    const hub = new CollabHub(provider);
    const tab1 = new FakeConn("t1");
    const tab2 = new FakeConn("t2");
    await hub.handle(tab1, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "alice", sinceSeq: 0 });
    await hub.handle(tab2, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "alice", sinceSeq: 0 });
    expect(tab2.last()).toEqual({ t: "refused", reason: "already-open" });
    // tab1 is untouched and still live.
    await hub.handle(tab1, { t: "submit", intent: { kind: "insertText", clientId: "alice", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" } });
    expect(tab1.last().t).toBe("broadcast");
  });

  it("takeover kicks the incumbent and admits the new connection", async () => {
    const hub = new CollabHub(provider);
    const zombie = new FakeConn("t1");
    const fresh = new FakeConn("t2");
    await hub.handle(zombie, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "alice", sinceSeq: 0 });
    await hub.handle(fresh, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "alice", takeover: true, sinceSeq: 0 });
    expect(zombie.received.some((m) => m.t === "refused" && m.reason === "taken-over")).toBe(true);
    expect(fresh.last().t).toBe("welcome");
    // The identity transferred: the fresh tab edits, the zombie is out.
    await hub.handle(fresh, { t: "submit", intent: { kind: "insertText", clientId: "alice", clientSeq: 1, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" } });
    expect(fresh.last().t).toBe("broadcast");
  });

  it("the same clientId on a DIFFERENT doc is not blocked (per-doc rule)", async () => {
    const hub = new CollabHub(provider);
    const c1 = new FakeConn("t1");
    const c2 = new FakeConn("t2");
    await hub.handle(c1, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "docA", clientId: "alice", sinceSeq: 0 });
    await hub.handle(c2, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "docB", clientId: "alice", sinceSeq: 0 });
    expect(c2.last().t).toBe("welcome");
  });
});

describe("CollabHub zero-custody eviction (doc 12 §2)", () => {
  it("deletes a room after the last disconnect + grace; occupied rooms never evict", async () => {
    let now = 0;
    const hub = new CollabHub(provider, undefined, undefined, () => now);
    const a = new FakeConn("a");
    const b = new FakeConn("b");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    await hub.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "b", sinceSeq: 0 });

    // Idle-with-people: hours pass, nobody types — the room survives (F7).
    now = 10 * 60 * 60 * 1000;
    expect(hub.sweepRooms()).toEqual([]);
    expect(hub.activeDocs()).toEqual(["d"]);

    // One leaves: still occupied, still immune.
    hub.disconnect(a);
    now += EVICTION_GRACE_MS * 2;
    expect(hub.sweepRooms()).toEqual([]);

    // Last one leaves: the grace clock starts. Just before grace: alive.
    hub.disconnect(b);
    now += EVICTION_GRACE_MS - 1;
    expect(hub.sweepRooms()).toEqual([]);
    // At grace: gone — the server holds NOTHING for this doc anymore.
    now += 1;
    expect(hub.sweepRooms()).toEqual(["d"]);
    expect(hub.activeDocs()).toEqual([]);
  });

  it("a rejoin during grace stops the eviction clock", async () => {
    let now = 0;
    const hub = new CollabHub(provider, undefined, undefined, () => now);
    const a = new FakeConn("a");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    hub.disconnect(a);
    now += EVICTION_GRACE_MS - 1000; // refresh lands just inside grace
    const a2 = new FakeConn("a2");
    await hub.handle(a2, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    now += EVICTION_GRACE_MS * 10; // long after — but the room is occupied
    expect(hub.sweepRooms()).toEqual([]);
    expect(hub.activeDocs()).toEqual(["d"]);
  });
});

describe("CollabHub welcome sidecar (round-4 F10 / round-2 F1)", () => {
  it("a late joiner into a session with split-created carried ids converges byte-identically", async () => {
    const hub = new CollabHub(provider);
    const a = new FakeConn("a");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });

    // "a" splits the only paragraph ("hi" → "h" | "i"): the new block/run get
    // CARRIED ids (7000/7001) that parse order can never re-derive — this is
    // exactly the state that used to poison late joiners.
    const split: SplitParagraphIntent = {
      kind: "splitParagraph", clientId: "a", clientSeq: 1, base: 0,
      at: { blockId: 1, runId: 2, offset: 1 }, newBlockId: 7000, newRunId: 7001,
    };
    await hub.handle(a, { t: "submit", intent: split });

    // Late joiner: builds its replica from the welcome bundle (snapshot +
    // sidecar), exactly as CollabConnection does.
    const b = new FakeConn("b");
    await hub.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "b", sinceSeq: 0 });
    const w = b.last();
    expect(w.t).toBe("welcome");
    if (w.t !== "welcome") return;
    expect(w.sidecar).toBeTruthy();
    const replica = new ClientReplica(b64ToBytes(w.snapshot), w.sidecar);
    replica.confirmedSeq = w.seq;

    // "a" now edits INSIDE the split-created run, addressed by its carried id.
    const intoSplit: InsertTextIntent = {
      kind: "insertText", clientId: "a", clientSeq: 2, base: w.seq,
      at: { blockId: 7000, runId: 7001, offset: 1 }, text: "X",
    };
    await hub.handle(a, { t: "submit", intent: intoSplit });
    const bc = b.last();
    expect(bc.t).toBe("broadcast");
    if (bc.t !== "broadcast") return;
    expect(bc.entries[0].kind).toBe("applied");
    replica.receive(bc.entries);

    // The pin: the joiner's document is byte-identical to the server's.
    // Without the sidecar its id table would misplace runId 7001 and the
    // insert would silently no-op or land in the wrong run.
    const serverXml = hub["rooms"].get("d")!.session.doc.save();
    expect(Buffer.from(replica.doc.save()).equals(Buffer.from(serverXml))).toBe(true);
  });
});

describe("CollabHub seed / re-seed / no-session (doc 12 §3/§5)", () => {
  // Deterministic epoch ids for assertions; production uses crypto randomness.
  let epochN = 0;
  const nextEpoch = () => `g_test_${epochN++}`;
  const zeroCustodyHub = (now: () => number = () => 0) =>
    new CollabHub(/*provider*/ null, undefined, undefined, now, nextEpoch);

  it("hello to an unseeded doc on a zero-custody hub is refused no-session", async () => {
    const hub = zeroCustodyHub();
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "ghost", clientId: "a", sinceSeq: 0 });
    expect(c.last()).toEqual({ t: "refused", reason: "no-session" });
    expect(hub.activeDocs()).toEqual([]); // the refused hello created nothing
  });

  it("seed creates the session; joiners get the seeded bytes, its genesisId, and plaintext mode", async () => {
    const hub = zeroCustodyHub();
    const seeded = hub.seed("d", blankDoc("seeded"));
    expect(seeded.ok).toBe(true);
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    const w = c.last();
    expect(w.t).toBe("welcome");
    if (w.t !== "welcome") return;
    expect(w.genesisId).toBe(seeded.ok ? seeded.genesisId : "");
    expect(w.mode).toBe("plaintext");
    // The session IS the seeded document (contains the seeded text).
    const replica = new ClientReplica(b64ToBytes(w.snapshot), w.sidecar);
    expect(new TextDecoder().decode(replica.doc.save()).length).toBeGreaterThan(0);
    const xml = serializeSnapshotText(replica);
    expect(xml).toContain("seeded");
  });

  it("first re-seeder wins; the loser gets the incumbent genesisId (HTTP 409 semantics)", async () => {
    const hub = zeroCustodyHub();
    const first = hub.seed("d", blankDoc("mine"));
    const second = hub.seed("d", blankDoc("theirs"));
    expect(first.ok).toBe(true);
    expect(second.ok).toBe(false);
    if (first.ok && !second.ok) {
      expect(second.reason).toBe("exists");
      // The loser learns the WINNER's epoch — its client lands in resume
      // case 2 (join the winner; local copy becomes a lineage decision).
      expect(second.genesisId).toBe(first.genesisId);
    }
  });

  it("session end → no-session → re-seed mints a NEW epoch (doc 12 lifecycle end-to-end)", async () => {
    let now = 0;
    const hub = zeroCustodyHub(() => now);
    const g1 = hub.seed("d", blankDoc("v1"));
    const a = new FakeConn("a");
    await hub.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    // Everyone leaves; grace passes; the room is deleted — server holds nothing.
    hub.disconnect(a);
    now += EVICTION_GRACE_MS;
    expect(hub.sweepRooms()).toEqual(["d"]);
    // The same link now resolves to no-session…
    const b = new FakeConn("b");
    await hub.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    expect(b.last()).toEqual({ t: "refused", reason: "no-session" });
    // …until any bundle-holder brings it back: same docId, NEW epoch.
    const g2 = hub.seed("d", blankDoc("v1"));
    expect(g2.ok && g1.ok && g2.genesisId !== g1.genesisId).toBe(true);
  });

  it("a seeded room nobody joins evicts after the grace (no RAM accretion via seed)", () => {
    let now = 0;
    const hub = zeroCustodyHub(() => now);
    hub.seed("d", blankDoc("x"));
    now += EVICTION_GRACE_MS - 1;
    expect(hub.sweepRooms()).toEqual([]);
    now += 1;
    expect(hub.sweepRooms()).toEqual(["d"]);
  });

  it("re-seed fidelity: a checkpoint bundle with split-created carried ids revives byte-identically (doc 12 test plan)", async () => {
    // Session 1 (old server): a split creates carried ids, then the "server
    // dies" — we keep only what a CLIENT would hold: the checkpoint bundle.
    const hub1 = zeroCustodyHub();
    hub1.seed("d", blankDoc("hi"));
    const a = new FakeConn("a");
    await hub1.handle(a, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0 });
    const split: SplitParagraphIntent = {
      kind: "splitParagraph", clientId: "a", clientSeq: 1, base: 0,
      at: { blockId: 1, runId: 2, offset: 1 }, newBlockId: 7000, newRunId: 7001,
    };
    await hub1.handle(a, { t: "submit", intent: split });
    const bundle = hub1["rooms"].get("d")!.session.checkpoint();

    // Session 2 (fresh server process): re-seed from the bundle — docx AND
    // sidecar (round-4 F10: the sidecar is not optional for split history).
    const hub2 = zeroCustodyHub();
    const g = hub2.seed("d", bundle.docx, bundle.sidecar);
    expect(g.ok).toBe(true);
    const b = new FakeConn("b");
    await hub2.handle(b, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "b", sinceSeq: 0 });
    const w = b.last();
    expect(w.t).toBe("welcome");
    if (w.t !== "welcome") return;
    const replica = new ClientReplica(b64ToBytes(w.snapshot), w.sidecar);
    replica.confirmedSeq = w.seq;

    // An edit addressed INTO the split-created run (carried id 7001) must
    // apply in the revived session — the pin that the id table survived the
    // server's death via the bundle, not via any server-side state.
    const intoSplit: InsertTextIntent = {
      kind: "insertText", clientId: "b", clientSeq: 1, base: w.seq,
      at: { blockId: 7000, runId: 7001, offset: 1 }, text: "X",
    };
    await hub2.handle(b, { t: "submit", intent: intoSplit });
    const bc = b.last();
    expect(bc.t).toBe("broadcast");
    if (bc.t !== "broadcast") return;
    expect(bc.entries[0].kind).toBe("applied");
    replica.receive(bc.entries);
    const serverXml = hub2["rooms"].get("d")!.session.doc.save();
    expect(Buffer.from(replica.doc.save()).equals(Buffer.from(serverXml))).toBe(true);
  });

  it("provider-created rooms (party/dev mode) also carry an epoch id", async () => {
    const hub = new CollabHub(provider, undefined, undefined, () => 0, nextEpoch);
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "party_0", clientId: "a", sinceSeq: 0 });
    const w = c.last();
    expect(w.t).toBe("welcome");
    if (w.t === "welcome") expect(w.genesisId).toMatch(/^g_test_/);
  });
});

/** Concatenated w:t text of a replica's doc (assert helper). */
function serializeSnapshotText(replica: ClientReplica): string {
  const walk = (el: { name: string; text: string; children: unknown[] }): string =>
    (el.name.endsWith(":t") ? el.text : "") + (el.children as never[]).map(walk).join("");
  return walk(replica.doc.docRoot as never);
}
