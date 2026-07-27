import { describe, expect, it } from "vitest";
import { mintDocKey, deriveEpochKeys, sealPresence, openPresence, isSealedPresence } from "../src/e2ee.js";
import type { PresencePosition } from "../src/protocol.js";

/**
 * SEALED PRESENCE (#20) — the server must learn WHO is pointing, never WHERE.
 *
 * Before this, presence rode the wire in the clear even in encrypted rooms:
 * the blind sequencer could not read a single character anyone wrote, but it
 * could watch every caret and, once selection ranges landed, how much each
 * participant had selected. Content confidentiality without position
 * confidentiality is a weaker claim than the room's design implies.
 *
 * The properties pinned here are the ones an attacker would probe:
 * round-trip fidelity, key separation, and each of the three AAD bindings
 * (document, epoch, sender) — because an AAD that binds nothing is a blob
 * that can be lifted from one room and replayed into another.
 */

/** Coordinates are LONG AND DISTINCTIVE on purpose: the leak assertion
 * greps the base64 ciphertext for them, and a short needle ("11") matches
 * base64 entropy by chance — measured at ~2% per seal, which made the leak
 * pin flake ~1 full-suite run in 4. A 9-digit needle makes a chance match
 * astronomically unlikely while proving exactly the same property. An
 * intermittently-failing leak assertion is worse than none: people learn
 * to re-run it. */
const position: PresencePosition = {
  anchor: { blockId: 987654321, runId: 918273645, offset: 123454321 },
  focus: { blockId: 987654321, runId: 918273645, offset: 543212345 },
  ranges: [
    { blockId: 987654321, runId: 918273645, start: 123454321, end: 543212345 },
    { blockId: 192837465, runId: 564738291, start: 0, end: 111222333 },
  ],
};

async function keysFor(docKey: string, genesisId: string) {
  return deriveEpochKeys(docKey, genesisId);
}

describe("sealed presence round-trips", () => {
  it("carries a caret and its selection ranges intact", async () => {
    const docKey = mintDocKey();
    const k = await keysFor(docKey, "g1");
    const sealed = await sealPresence(k.kPresence, "doc1", "g1", "alice", position);
    expect(isSealedPresence(sealed)).toBe(true);
    const opened = await openPresence(k.kPresence, "doc1", "g1", "alice", sealed);
    expect(opened).toEqual(position);
  });

  it("carries a null position (caret cleared) as well as a caret", async () => {
    const k = await keysFor(mintDocKey(), "g1");
    const sealed = await sealPresence(k.kPresence, "doc1", "g1", "alice", null);
    expect(await openPresence(k.kPresence, "doc1", "g1", "alice", sealed)).toBeNull();
  });

  it("leaks no coordinate into the blob the relay sees", async () => {
    const k = await keysFor(mintDocKey(), "g1");
    const sealed = await sealPresence(k.kPresence, "doc1", "g1", "alice", position);
    // The server holds exactly these two strings plus the sender's id. If any
    // coordinate were recoverable from them, the whole change is theatre.
    const onTheWire = JSON.stringify(sealed);
    for (const needle of ["anchor", "ranges", "blockId", "987654321", "918273645", "123454321", "offset"]) {
      expect(onTheWire, `"${needle}" must not appear in the sealed payload`).not.toContain(needle);
    }
    expect(Object.keys(sealed).sort()).toEqual(["ciphertext", "iv"]);
  });

  it("uses a fresh IV per seal, so identical positions do not produce identical blobs", async () => {
    const k = await keysFor(mintDocKey(), "g1");
    const a = await sealPresence(k.kPresence, "doc1", "g1", "alice", position);
    const b = await sealPresence(k.kPresence, "doc1", "g1", "alice", position);
    expect(a.iv).not.toBe(b.iv);
    expect(a.ciphertext).not.toBe(b.ciphertext);
  });
});

describe("sealed presence is bound to its room, epoch and sender", () => {
  it("rejects a blob replayed into a DIFFERENT document", async () => {
    const k = await keysFor(mintDocKey(), "g1");
    const sealed = await sealPresence(k.kPresence, "doc1", "g1", "alice", position);
    await expect(openPresence(k.kPresence, "doc2", "g1", "alice", sealed)).rejects.toThrow();
  });

  it("rejects a blob replayed into a DIFFERENT epoch", async () => {
    const docKey = mintDocKey();
    const k1 = await keysFor(docKey, "g1");
    const sealed = await sealPresence(k1.kPresence, "doc1", "g1", "alice", position);
    // Same document key, later epoch: the derived key differs AND the AAD
    // differs, so a re-seeded room cannot inherit stale carets.
    const k2 = await keysFor(docKey, "g2");
    await expect(openPresence(k2.kPresence, "doc1", "g2", "alice", sealed)).rejects.toThrow();
  });

  it("rejects RE-ATTRIBUTION of a blob to another participant", async () => {
    const k = await keysFor(mintDocKey(), "g1");
    const sealed = await sealPresence(k.kPresence, "doc1", "g1", "alice", position);
    // The receiver builds the AAD from the server's `participant` stamp, so a
    // relay that relabels alice's caret as bob's produces a decryption
    // failure — not bob's cursor sitting where alice is working.
    await expect(openPresence(k.kPresence, "doc1", "g1", "bob", sealed)).rejects.toThrow();
  });

  it("rejects a tampered ciphertext", async () => {
    const k = await keysFor(mintDocKey(), "g1");
    const sealed = await sealPresence(k.kPresence, "doc1", "g1", "alice", position);
    const flipped = {
      ...sealed,
      ciphertext: sealed.ciphertext.slice(0, -2) + (sealed.ciphertext.endsWith("A") ? "B" : "A") + "=",
    };
    await expect(openPresence(k.kPresence, "doc1", "g1", "alice", flipped)).rejects.toThrow();
  });

  it("is sealed under a key DISTINCT from content and media", async () => {
    const k = await keysFor(mintDocKey(), "g1");
    const sealed = await sealPresence(k.kPresence, "doc1", "g1", "alice", position);
    // Key separation is the point: a presence-side mistake must not be a
    // content-side compromise, and vice versa.
    await expect(openPresence(k.kContent, "doc1", "g1", "alice", sealed)).rejects.toThrow();
    await expect(openPresence(k.kMedia, "doc1", "g1", "alice", sealed)).rejects.toThrow();
  });

  it("derives the same presence key for everyone holding the link", async () => {
    // Nobody redistributes keys: two participants derive identically from the
    // link plus the public epoch id, which is what makes sealing free.
    const docKey = mintDocKey();
    const alice = await keysFor(docKey, "g1");
    const bob = await keysFor(docKey, "g1");
    const sealed = await sealPresence(alice.kPresence, "doc1", "g1", "alice", position);
    expect(await openPresence(bob.kPresence, "doc1", "g1", "alice", sealed)).toEqual(position);
  });
});

describe("isSealedPresence discriminates the two wire shapes", () => {
  it("tells a sealed blob from a plaintext position and from junk", () => {
    expect(isSealedPresence({ iv: "a", ciphertext: "b" })).toBe(true);
    expect(isSealedPresence(position)).toBe(false);
    expect(isSealedPresence(null)).toBe(false);
    expect(isSealedPresence(undefined)).toBe(false);
    expect(isSealedPresence({ iv: 1, ciphertext: 2 })).toBe(false);
    expect(isSealedPresence({ iv: "a" })).toBe(false);
  });
});
