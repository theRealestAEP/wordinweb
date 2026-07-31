import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { EncryptedCollabConnection } from "../src/enc-connection.js";
import { CollabConnection } from "../src/connection.js";
import { mintDocKey, deriveEpochKeys, sealCheckpoint, sealPresence, bytesToB64 } from "../src/e2ee.js";
import type { ClientMessage, ServerMessage, PresencePosition } from "../src/protocol.js";

/**
 * SEALED PRESENCE AT THE CONNECTION (#20) — the tolerance and blast-radius
 * half. The crypto pins prove a bad blob cannot be opened; these prove what
 * the connection DOES about it, which is where a hostile participant either
 * costs the room nothing or takes it down.
 *
 * Three behaviours, each a deliberate decision rather than a fallout:
 *  - garbage costs the sender its own cursor and nobody else theirs;
 *  - a PLAINTEXT position arriving in an encrypted room is ignored, not
 *    rendered (rendering it would quietly re-open the leak this closes: the
 *    peer that sent it also handed it to the server in the clear);
 *  - a SEALED blob arriving in a plaintext connection is ignored, not handed
 *    to the renderer as coordinates.
 * The e4 engine fence should make the last two unreachable; they exist so a
 * mixed room degrades to "no remote carets" rather than to a crash.
 */

function docxBytes(text: string): Uint8Array {
  const documentXml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`;
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

async function sealedSeed(genesisId: string, docKey: string) {
  const keys = await deriveEpochKeys(docKey, genesisId);
  const session = new DocumentSession(DocxDocument.load(docxBytes("hi")));
  const cp = session.checkpoint();
  const sealed = await sealCheckpoint(keys.kContent, "d", genesisId, 0, {
    docx: bytesToB64(cp.docx),
    sidecar: cp.sidecar,
    docHash: "seed",
  });
  return { keys, checkpoint: { seq: 0, ...sealed } };
}

/** A transport whose inbound side the test drives directly. */
function fakeTransport(onSend?: (m: ClientMessage) => void) {
  let deliver: (m: ServerMessage) => void = () => {};
  return {
    transport: {
      send: (m: ClientMessage) => onSend?.(m),
      onMessage: (cb: (m: ServerMessage) => void) => {
        deliver = cb;
      },
    },
    deliver: (m: ServerMessage) => deliver(m),
  };
}

async function until(cond: () => boolean, label: string): Promise<void> {
  for (let i = 0; i < 200; i++) {
    if (cond()) return;
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error(`timed out waiting for ${label}`);
}

const GOOD: PresencePosition = { anchor: { blockId: 1, runId: 2, offset: 3 } };

describe("encrypted connection, inbound presence", () => {
  async function connected() {
    const docKey = mintDocKey();
    const { keys, checkpoint } = await sealedSeed("g1", docKey);
    const seen: { participant: string; position: PresencePosition | null }[] = [];
    const t = fakeTransport();
    const conn = new EncryptedCollabConnection(t.transport, "me", docKey, {
      onPresence: (participant, position) => seen.push({ participant, position }),
    });
    conn.join("d");
    t.deliver({ t: "welcome-enc", docId: "d", genesisId: "g1", checkpoint, tail: [], mode: "encrypted" } as never);
    await until(() => conn.ready, "the encrypted connection to rehydrate");
    return { conn, t, keys, seen };
  }

  it("opens a peer's sealed caret and reports it", async () => {
    const { t, keys, seen } = await connected();
    const sealed = await sealPresence(keys.kPresence, "d", "g1", "alice", GOOD);
    t.deliver({ t: "presence", participant: "alice", position: sealed } as never);
    await until(() => seen.length > 0, "the opened caret");
    expect(seen[0]).toEqual({ participant: "alice", position: GOOD });
  });

  it("drops a hostile participant's garbage without disturbing anyone else", async () => {
    const { t, keys, seen } = await connected();

    // mallory sends a blob that is not openable in this room at all.
    t.deliver({
      t: "presence",
      participant: "mallory",
      position: { iv: "AAAAAAAAAAAAAAAA", ciphertext: "3q2+7w==" },
    } as never);
    // …and a VALID blob of alice's, re-attributed by a hostile relay to bob.
    const alices = await sealPresence(keys.kPresence, "d", "g1", "alice", GOOD);
    t.deliver({ t: "presence", participant: "bob", position: alices } as never);
    // An honest peer, after both.
    const carols = await sealPresence(keys.kPresence, "d", "g1", "carol", GOOD);
    t.deliver({ t: "presence", participant: "carol", position: carols } as never);

    await until(() => seen.some((s) => s.participant === "carol"), "carol's caret");
    // The honest caret arrived; neither bad payload produced one, and neither
    // wedged the chain that carol's arrived on.
    expect(seen.map((s) => s.participant)).toEqual(["carol"]);
  });

  it("ignores a PLAINTEXT position arriving in an encrypted room", async () => {
    const { t, seen } = await connected();
    t.deliver({ t: "presence", participant: "legacy", position: GOOD } as never);
    // Give it the same number of turns a real open would have taken.
    await new Promise((r) => setTimeout(r, 40));
    expect(seen, "an unsealed position is not rendered").toEqual([]);
  });

  it("clamps a hostile peer's oversized range list AFTER opening it", async () => {
    const { t, keys, seen } = await connected();
    // Sealing proves who sent it, never that it is sane: the server cannot
    // clamp what it cannot read, so the receiver must.
    const flood = {
      anchor: { blockId: 1, runId: 2, offset: 0 },
      ranges: Array.from({ length: 500 }, (_, i) => ({ blockId: 1, runId: 2, start: i, end: i + 1 })),
    };
    const sealed = await sealPresence(keys.kPresence, "d", "g1", "alice", flood);
    t.deliver({ t: "presence", participant: "alice", position: sealed } as never);
    await until(() => seen.length > 0, "the opened caret");
    expect(seen[0].position?.ranges?.length ?? 0).toBeLessThanOrEqual(64);
  });
});

describe("encrypted connection, outbound presence", () => {
  it("coalesces a burst of caret moves instead of queueing one seal each", async () => {
    const docKey = mintDocKey();
    const { checkpoint } = await sealedSeed("g1", docKey);
    const sent: ClientMessage[] = [];
    const t = fakeTransport((m) => sent.push(m));
    const conn = new EncryptedCollabConnection(t.transport, "me", docKey, {});
    conn.join("d");
    t.deliver({ t: "welcome-enc", docId: "d", genesisId: "g1", checkpoint, tail: [], mode: "encrypted" } as never);
    await until(() => conn.ready, "the connection to rehydrate");
    sent.length = 0;

    // 50 caret moves in one synchronous burst — a drag selection, or a user
    // holding an arrow key. Sealing is async, so all 50 arrive while the
    // first seal is still running.
    for (let i = 0; i < 50; i++) conn.setPresence({ anchor: { blockId: 1, runId: 2, offset: i } });
    await new Promise((r) => setTimeout(r, 120));

    const presence = sent.filter((m) => m.t === "presence");
    // The point of single-flight coalescing: superseded carets have no reader,
    // so they are dropped rather than sealed and shipped. Far fewer than 50
    // sends, and never zero — the LAST position must always get out, or the
    // peer's cursor freezes at a stale spot.
    expect(presence.length, "a burst does not become 50 seals").toBeLessThan(10);
    expect(presence.length, "but the latest position is always sent").toBeGreaterThan(0);
  });

  it("sends nothing at all before the welcome (no plaintext leak, no queue)", async () => {
    const sent: ClientMessage[] = [];
    const t = fakeTransport((m) => sent.push(m));
    const conn = new EncryptedCollabConnection(t.transport, "me", mintDocKey(), {});
    conn.join("d");
    conn.setPresence(GOOD);
    await new Promise((r) => setTimeout(r, 40));
    // There is no key yet. Dropping is correct — the alternative is shipping
    // the coordinates in the clear, which is the leak this whole change
    // closes — and the next caret move re-sends anyway.
    expect(sent.filter((m) => m.t === "presence"), "no presence before keys exist").toEqual([]);
  });
});

describe("edit-loss accounting partitions", () => {
  it("a submit whose send throws is counted as sendFailures, distinctly from droppedPreReady", async () => {
    const docKey = mintDocKey();
    const { checkpoint } = await sealedSeed("g1", docKey);
    const where: string[] = [];
    let failSends = false;
    let deliver: (m: ServerMessage) => void = () => {};
    const transport = {
      send: (_m: ClientMessage) => {
        if (failSends) throw new Error("socket closed");
      },
      onMessage: (cb: (m: ServerMessage) => void) => {
        deliver = cb;
      },
    };
    const conn = new EncryptedCollabConnection(transport, "me", docKey, {
      onError: (info) => where.push(info.where),
    });

    // BEFORE the welcome: refused outright, and counted on the pre-ready side.
    conn.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "x" } as never);
    expect(conn.droppedPreReady, "refused before it was ever applied").toBe(1);
    expect(conn.sendFailures, "…and NOT counted as a send failure").toBe(0);

    conn.join("d");
    deliver({ t: "welcome-enc", docId: "d", genesisId: "g1", checkpoint, tail: [], mode: "encrypted" } as never);
    await until(() => conn.ready, "the connection to rehydrate");

    // AFTER the welcome, with the wire broken: applied optimistically, then
    // lost on the way out. The opposite side of the partition.
    failSends = true;
    conn.submit({ kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "y" } as never);
    await until(() => conn.sendFailures > 0, "the failed send to be counted");

    expect(conn.droppedPreReady, "still just the one pre-ready drop").toBe(1);
    expect(conn.sendFailures, "the broken send counted separately").toBe(1);
    expect(where, "reported under a label that names the submit path").toContain("enc.submit");
    // The whole point: two disjoint numbers. A run that loses N edits can
    // subtract both and see what is left unexplained — the arithmetic B13
    // spent a night without.
  });
});

describe("plaintext connection, inbound presence", () => {
  it("ignores a SEALED blob rather than handing a blob to the renderer", () => {
    const seen: unknown[] = [];
    const t = fakeTransport();
    const conn = new CollabConnection(t.transport, "me", {
      onPresence: (_p, position) => seen.push(position),
    });
    conn.join("d");
    // No welcome needed: the presence branch is independent of the replica,
    // which is the point — a blob must be rejected on shape alone, before
    // anything tries to treat its fields as coordinates.
    t.deliver({ t: "presence", participant: "enc-peer", position: { iv: "AAAA", ciphertext: "BBBB" } } as never);
    expect(seen, "a sealed payload is not renderable here").toEqual([]);
    // A normal plaintext position still works.
    t.deliver({ t: "presence", participant: "plain-peer", position: GOOD } as never);
    expect(seen).toEqual([GOOD]);
  });
});
