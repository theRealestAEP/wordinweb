import { describe, expect, it } from "vitest";
import { CollabHub, type Connection, type DocProvider, type ServerMessage } from "../src/index.js";
import { PROTOCOL_VERSION, ENGINE_VERSION } from "@wordinweb/collab/server";
import { mintDocKey, deriveEpochKeys, sealPresence, sealCheckpoint } from "@wordinweb/collab/client";
import { zipSync, strToU8 } from "fflate";

/**
 * SERVER BLINDNESS FOR PRESENCE (#20).
 *
 * The blind sequencer's claim used to stop at content: it could not read a
 * character anyone typed, but every caret and every selection range crossed
 * it in the clear, so it could watch where each participant was working and
 * how much they had highlighted. This asserts the claim now covers position
 * too — the coordinates appear NOWHERE in anything the server holds or
 * forwards, only in what the peers can decrypt.
 *
 * The test drives the REAL hub with sealed payloads (no plaintext shortcut)
 * and then greps everything the server touched: the relayed message, and the
 * whole room object the hub keeps in RAM.
 */

function docBytes(): Uint8Array {
  const xml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body><w:p><w:r><w:t xml:space="preserve">hello</w:t></w:r></w:p></w:body></w:document>`;
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
    "word/document.xml": strToU8(xml),
  });
}

/** A distinctive position: every number in it is rare enough that finding it
 * anywhere in server state is unambiguous evidence of a leak. */
const SECRET_POSITION = {
  anchor: { blockId: 987651, runId: 987652, offset: 987653 },
  ranges: [{ blockId: 987651, runId: 987652, start: 987654, end: 987655 }],
};
const SECRET_NUMBERS = ["987651", "987652", "987653", "987654", "987655"];

describe("the hub is blind to presence in encrypted rooms", () => {
  it("relays a sealed caret without holding or forwarding a single coordinate", async () => {
    const provider: DocProvider = { load: () => docBytes() };
    const hub = new CollabHub(provider);
    const docKey = mintDocKey();
    const genesisId = "genesis-1";
    const keys = await deriveEpochKeys(docKey, genesisId);

    const mk = (id: string): { conn: Connection; got: ServerMessage[] } => {
      const got: ServerMessage[] = [];
      return { conn: { id, send: (m) => got.push(m) }, got };
    };
    const alice = mk("cA");
    const bob = mk("cB");

    // Seed a real ENCRYPTED room, then join both participants to it.
    const checkpoint = await sealCheckpoint(keys.kContent, "d1", genesisId, 0, {
      docx: "",
      sidecar: null,
    } as never);
    await hub.handle(alice.conn, {
      t: "hello", protocolVersion: PROTOCOL_VERSION, engineVersion: ENGINE_VERSION,
      docId: "d1", clientId: "alice", sinceSeq: 0, enc: { genesisId, checkpoint },
    } as never);
    await hub.handle(bob.conn, {
      t: "hello", protocolVersion: PROTOCOL_VERSION, engineVersion: ENGINE_VERSION,
      docId: "d1", clientId: "bob", sinceSeq: 0,
    } as never);

    const sealed = await sealPresence(keys.kPresence, "d1", genesisId, "alice", SECRET_POSITION);
    await hub.handle(alice.conn, { t: "presence", position: sealed } as never);

    // 1. The peer received it, opaque and attributed.
    const relayed = bob.got.filter((m) => (m as { t: string }).t === "presence") as Array<{
      participant: string;
      position: { iv: string; ciphertext: string } | null;
    }>;
    expect(relayed.length, "bob got exactly one presence message").toBe(1);
    expect(relayed[0].participant, "attributed to the BOUND sender id").toBe("alice");
    expect(relayed[0].position, "relayed verbatim, still sealed").toEqual(sealed);

    // 2. THE BLINDNESS ASSERTION: no coordinate anywhere the server touched.
    // Serializing the whole relayed stream and the hub's own room state and
    // grepping is deliberately blunt — a narrower assertion would only prove
    // the field we thought to check.
    const everythingForwarded = JSON.stringify(bob.got) + JSON.stringify(alice.got);
    const everythingHeld = JSON.stringify(
      (hub as unknown as { rooms: Map<string, unknown> }).rooms.get("d1") ?? {},
      (_k, v) => (v instanceof Map ? [...v.entries()] : v instanceof Set ? [...v] : v),
    );
    for (const n of SECRET_NUMBERS) {
      expect(everythingForwarded, `coordinate ${n} must not appear in anything forwarded`).not.toContain(n);
      expect(everythingHeld, `coordinate ${n} must not appear in server-held room state`).not.toContain(n);
    }
    for (const field of ["anchor", "ranges", "blockId", "offset"]) {
      expect(everythingForwarded, `"${field}" must not appear in the clear`).not.toContain(field);
    }
  });

  it("does not clamp a sealed payload (the server cannot parse it, and must not try)", async () => {
    const provider: DocProvider = { load: () => docBytes() };
    const hub = new CollabHub(provider);
    const docKey = mintDocKey();
    const genesisId = "genesis-2";
    const keys = await deriveEpochKeys(docKey, genesisId);
    const mk = (id: string): { conn: Connection; got: ServerMessage[] } => {
      const got: ServerMessage[] = [];
      return { conn: { id, send: (m) => got.push(m) }, got };
    };
    const a = mk("cA");
    const b = mk("cB");
    const checkpoint = await sealCheckpoint(keys.kContent, "d2", genesisId, 0, { docx: "", sidecar: null } as never);
    await hub.handle(a.conn, {
      t: "hello", protocolVersion: PROTOCOL_VERSION, engineVersion: ENGINE_VERSION,
      docId: "d2", clientId: "alice", sinceSeq: 0, enc: { genesisId, checkpoint },
    } as never);
    await hub.handle(b.conn, {
      t: "hello", protocolVersion: PROTOCOL_VERSION, engineVersion: ENGINE_VERSION,
      docId: "d2", clientId: "bob", sinceSeq: 0,
    } as never);

    // A payload that WOULD be clamped if the server could read it: 500 ranges,
    // far over the 64 wire cap. Sealed, the hub must pass it through untouched
    // — the receiving client clamps after opening, where the ranges exist.
    const flood = {
      anchor: { blockId: 1, runId: 2, offset: 0 },
      ranges: Array.from({ length: 500 }, (_, i) => ({ blockId: 1, runId: 2, start: i, end: i + 1 })),
    };
    const sealed = await sealPresence(keys.kPresence, "d2", genesisId, "alice", flood);
    await hub.handle(a.conn, { t: "presence", position: sealed } as never);

    const relayed = b.got.filter((m) => (m as { t: string }).t === "presence") as Array<{ position: unknown }>;
    expect(relayed.length).toBe(1);
    expect(relayed[0].position, "byte-identical to what was sent").toEqual(sealed);
  });
});
