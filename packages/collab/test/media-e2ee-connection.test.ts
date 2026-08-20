import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument } from "@wordinweb/core";
import { EncryptedCollabConnection } from "../src/enc-connection.js";
import { DocumentSession } from "../src/session.js";
import { mintDocKey, deriveEpochKeys, sealCheckpoint, bytesToB64 } from "../src/e2ee.js";
import { sha256Hex, mediaAddressesOf, type FetchLike } from "../src/media.js";
import type { ClientMessage, ServerMessage, EnvelopeEntry, SealedCheckpoint } from "../src/protocol.js";

/**
 * Media over the ENCRYPTED connection — the demo's actual path, and the one
 * that had no media code at all before this arc (every duty lived only in the
 * plaintext connection).
 *
 * The property under test is the one a blind relay makes possible: the server
 * orders opaque envelopes and stores opaque blobs, yet two clients that never
 * exchanged anything but ciphertext end up holding identical pixels. Nothing
 * the relay holds is ever the image.
 */

function docxBytes(text: string): Uint8Array {
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

/** The blind sequencer, plus the media control frames it relays. */
function blindServer(genesisId: string, checkpoint: SealedCheckpoint) {
  const log: EnvelopeEntry[] = [];
  const seen = new Set<string>();
  const peers: { deliver: (m: ServerMessage) => void }[] = [];
  const attach = () => {
    const peer = { deliver: (_m: ServerMessage) => {} };
    peers.push(peer);
    return {
      send: (msg: ClientMessage) => {
        if (msg.t === "hello") {
          peer.deliver({ t: "welcome-enc", docId: "d", genesisId, checkpoint, tail: [...log], mode: "encrypted" });
        } else if (msg.t === "submit-enc") {
          const key = `${msg.envelope.clientId}:${msg.envelope.clientSeq}`;
          let entry = seen.has(key) ? log.find((e) => `${e.clientId}:${e.clientSeq}` === key) : undefined;
          if (!entry) {
            seen.add(key);
            entry = { ...msg.envelope, seq: log.length === 0 ? checkpoint.seq + 1 : log[log.length - 1].seq + 1 };
            log.push(entry);
          }
          for (const p of peers) p.deliver({ t: "broadcast-enc", entries: [entry!] });
        }
      },
      onMessage: (cb: (m: ServerMessage) => void) => { peer.deliver = cb; },
    };
  };
  return { attach, log: () => log };
}

async function seedEncrypted(text: string, genesisId: string, docKey: string) {
  const keys = await deriveEpochKeys(docKey, genesisId);
  const session = new DocumentSession(DocxDocument.load(docxBytes(text)));
  const cp = session.checkpoint();
  const sealed = await sealCheckpoint(keys.kContent, "d", genesisId, 0, {
    docx: bytesToB64(cp.docx), sidecar: cp.sidecar, docHash: "seed",
  });
  return { checkpoint: { seq: 0, ...sealed } };
}

/** The relay: hashes every PUT itself, serves bytes at their address. */
function relay() {
  const blobs = new Map<string, Uint8Array>();
  const calls = { put: 0, get: 0 };
  const fetchImpl: FetchLike = async (url, init) => {
    const sha = url.slice(url.lastIndexOf("/") + 1);
    if (init?.method === "PUT") {
      calls.put++;
      if ((await sha256Hex(init.body!)) !== sha) return { ok: false, status: 400, arrayBuffer: async () => new ArrayBuffer(0) };
      blobs.set(sha, init.body!);
      return { ok: true, status: 201, arrayBuffer: async () => new ArrayBuffer(0) };
    }
    calls.get++;
    const hit = blobs.get(sha);
    if (!hit) return { ok: false, status: 404, arrayBuffer: async () => new ArrayBuffer(0) };
    return { ok: true, status: 200, arrayBuffer: async () => hit.slice().buffer as ArrayBuffer };
  };
  return { blobs, calls, opts: { httpBase: "http://relay", fetchImpl } };
}

async function until(cond: () => boolean, label: string): Promise<void> {
  for (let i = 0; i < 300; i++) {
    if (cond()) return;
    await new Promise((r) => setTimeout(r, 10));
  }
  throw new Error(`timed out waiting for ${label}`);
}

const PIXELS = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 42, 43, 44, 45]);

describe("media over the encrypted connection (doc 16 §6)", () => {
  it("one client uploads ciphertext; another decrypts it to identical pixels", async () => {
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("hi", "g1", docKey);
    const srv = blindServer("g1", checkpoint);
    const store = relay();

    const alice = new EncryptedCollabConnection(srv.attach(), "alice", docKey, {}, undefined, 400, store.opts);
    const bob = new EncryptedCollabConnection(srv.attach(), "bob", docKey, {}, undefined, 400, store.opts);
    alice.join("d");
    bob.join("d");
    await until(() => alice.ready && bob.ready, "both clients to rehydrate");

    // PLACER: seal + upload, then reserve with the returned address.
    const fields = await alice.uploadMedia(PIXELS);
    expect(fields).not.toBeNull();
    expect(fields!.iv).toBeTruthy();
    // The relay holds CIPHERTEXT, and its address is that ciphertext's hash.
    const stored = store.blobs.get(fields!.blobSha)!;
    expect(stored).toBeTruthy();
    expect(Buffer.from(stored)).not.toEqual(Buffer.from(PIXELS));
    expect(await sha256Hex(stored)).toBe(fields!.blobSha);

    alice.submit({
      kind: "insertImage", runId: 2, blobSha: fields!.blobSha, bytesLen: fields!.bytesLen,
      ext: "png", iv: fields!.iv, widthPx: 8, heightPx: 8, nodeIds: [900],
    } as never);
    await until(() => (bob.doc?.pendingMedia.size ?? 0) > 0 || (bob.doc?.mediaMeta.size ?? 0) > 0,
      "the reservation to reach bob");

    // RECEIVER: the eager fetch pulled the blob, verified the address, and
    // opened it under K_media — all without a nudge from the test.
    const part = [...bob.doc!.mediaMeta.keys()][0];
    await until(() => bob.doc!.mediaStatus(part) === "ready", "bob to install the pixels");
    expect(Buffer.from(bob.doc!.media(part)!)).toEqual(Buffer.from(PIXELS));

    // And the server never held the image in any form.
    const everything = JSON.stringify(srv.log()) + JSON.stringify([...store.blobs.values()].map((b) => [...b]));
    expect(everything).not.toContain([...PIXELS].join(","));
  });

  it("the placer installs its own pixels without a round trip", async () => {
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("hi", "g1", docKey);
    const srv = blindServer("g1", checkpoint);
    const store = relay();
    const alice = new EncryptedCollabConnection(srv.attach(), "alice", docKey, {}, undefined, 400, store.opts);
    alice.join("d");
    await until(() => alice.ready, "alice to rehydrate");

    const fields = (await alice.uploadMedia(PIXELS))!;
    alice.submit({
      kind: "insertImage", runId: 2, blobSha: fields.blobSha, bytesLen: fields.bytesLen,
      ext: "png", iv: fields.iv, widthPx: 8, heightPx: 8, nodeIds: [900],
    } as never);
    await until(() => (alice.doc?.mediaMeta.size ?? 0) > 0, "alice's own reservation");
    const part = [...alice.doc!.mediaMeta.keys()][0];
    await until(() => alice.doc!.mediaStatus(part) === "ready", "alice to hold pixels");
    expect(Buffer.from(alice.doc!.media(part)!)).toEqual(Buffer.from(PIXELS));
    // …and she never downloaded her own image back (doc 16 §5.1: "bytes
    // install into the placer's own doc immediately"). This matters in an
    // encrypted room, where the reservation applies on an async queue well
    // after the upload returns — the naive version fetched and decrypted a
    // blob it had just encrypted.
    expect(store.calls.get).toBe(0);
  });

  it("re-supply after eviction restores the exact ciphertext address", async () => {
    // The end-to-end version of the recorded-IV rule: evict the relay copy,
    // ask the holder to re-supply, and the blob that comes back must hash to
    // the same address the reservation committed to.
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("hi", "g1", docKey);
    const srv = blindServer("g1", checkpoint);
    const store = relay();
    const alice = new EncryptedCollabConnection(srv.attach(), "alice", docKey, {}, undefined, 400, store.opts);
    alice.join("d");
    await until(() => alice.ready, "alice to rehydrate");

    const fields = (await alice.uploadMedia(PIXELS))!;
    alice.submit({
      kind: "insertImage", runId: 2, blobSha: fields.blobSha, bytesLen: fields.bytesLen,
      ext: "png", iv: fields.iv, widthPx: 8, heightPx: 8, nodeIds: [900],
    } as never);
    await until(() => (alice.doc?.mediaMeta.size ?? 0) > 0, "the reservation");
    const part = [...alice.doc!.mediaMeta.keys()][0];
    await until(() => alice.doc!.mediaStatus(part) === "ready", "alice to hold pixels");

    store.blobs.delete(fields.blobSha); // TTL eviction
    expect(await alice.media!.resupply("d", fields.blobSha)).toBe(true);
    const restored = store.blobs.get(fields.blobSha)!;
    expect(await sha256Hex(restored)).toBe(fields.blobSha);
  });
});

describe("E2EE late join through the sealed checkpoint (doc 16 §6)", () => {
  it("a joiner rehydrating from a checkpoint FETCHES media it never saw inserted", async () => {
    // The encrypted half of the late-join fix. A joiner's checkpoint contains
    // the image's REGISTRATION but not its declared sha — that address lives
    // only in intents already folded into the snapshot. The address map now
    // rides INSIDE the sealed body, so the joiner can fetch while the server
    // still learns nothing about the document's part structure.
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("hi", "g1", docKey);
    const srv = blindServer("g1", checkpoint);
    const store = relay();

    const alice = new EncryptedCollabConnection(srv.attach(), "alice", docKey, {}, undefined, 400, store.opts);
    alice.join("d");
    await until(() => alice.ready, "alice to rehydrate");
    const fields = (await alice.uploadMedia(PIXELS))!;
    alice.submit({
      kind: "insertImage", runId: 2, blobSha: fields.blobSha, bytesLen: fields.bytesLen,
      ext: "png", iv: fields.iv, widthPx: 8, heightPx: 8, nodeIds: [900],
    } as never);
    await until(() => (alice.doc?.mediaMeta.size ?? 0) > 0, "the reservation");
    const part = [...alice.doc!.mediaMeta.keys()][0];
    await until(() => alice.doc!.mediaStatus(part) === "ready", "alice to hold pixels");

    // Seal a checkpoint the way the designated checkpointer does — from a
    // MIRROR, which never installs pixels but does hold the registrations.
    const keys = await deriveEpochKeys(docKey, "g1");
    const mirror = new DocumentSession(DocxDocument.load(docxBytes("hi")));
    mirror.submit({
      kind: "insertImage", clientId: "alice", clientSeq: 1, base: 0, runId: 2,
      blobSha: fields.blobSha, bytesLen: fields.bytesLen, ext: "png", iv: fields.iv,
      widthPx: 8, heightPx: 8, nodeIds: [900],
    } as never);
    expect(mirror.doc.mediaStatus(part)).toBe("pending"); // no pixels, as expected
    const cp = mirror.checkpoint();
    const sealed = await sealCheckpoint(keys.kContent, "d", "g1", cp.seq, {
      docx: bytesToB64(cp.docx),
      sidecar: cp.sidecar,
      docHash: "x",
      mediaMeta: mediaAddressesOf(mirror.doc),
    });
    // The addresses are SEALED: the server sees only opaque ciphertext.
    expect(JSON.stringify(sealed)).not.toContain(fields.blobSha);

    const late = blindServer("g1", { seq: cp.seq, ...sealed });
    const carol = new EncryptedCollabConnection(late.attach(), "carol", docKey, {}, undefined, 400, store.opts);
    carol.join("d");
    await until(() => carol.ready, "carol to rehydrate");
    // She reserved the box AND had the address to fill it.
    await until(() => carol.doc!.mediaStatus(part) === "ready", "carol to fetch the image");
    expect(Buffer.from(carol.doc!.media(part)!)).toEqual(Buffer.from(PIXELS));
  });

  it("a checkpoint WITHOUT the map degrades to reserve-but-cannot-fill", async () => {
    // Tolerant read: an older build's checkpoint has no mediaMeta, and such a
    // joiner must still load cleanly — reserving the image's box and showing
    // it as unavailable rather than erroring or drawing nothing.
    const docKey = mintDocKey();
    const store = relay();
    const keys = await deriveEpochKeys(docKey, "g1");
    const mirror = new DocumentSession(DocxDocument.load(docxBytes("hi")));
    mirror.submit({
      kind: "insertImage", clientId: "alice", clientSeq: 1, base: 0, runId: 2,
      blobSha: "a".repeat(64), bytesLen: 4, ext: "png", widthPx: 8, heightPx: 8, nodeIds: [900],
    } as never);
    const cp = mirror.checkpoint();
    const sealed = await sealCheckpoint(keys.kContent, "d", "g1", cp.seq, {
      docx: bytesToB64(cp.docx), sidecar: cp.sidecar, docHash: "x", // no mediaMeta
    });
    const srv = blindServer("g1", { seq: cp.seq, ...sealed });
    const carol = new EncryptedCollabConnection(srv.attach(), "carol", docKey, {}, undefined, 400, store.opts);
    carol.join("d");
    await until(() => carol.ready, "carol to rehydrate");

    const holes = [...carol.doc!.pendingMedia.entries()];
    expect(holes.length).toBe(1);
    expect(holes[0][1].sha).toBe("");   // addressless: nothing to fetch
    expect(store.calls.get).toBe(0);    // and we never asked pointlessly
  });
});

describe("collaborative undo over two encrypted clients (doc 03 Phase 8)", () => {
  /**
   * The user's own specification, end to end: "undo the last intent that user
   * did, treat it like a new intent; if the target has already been deleted or
   * changed, the intent gets made invalid."
   *
   * No new wire message and no server involvement: the mirror is a real
   * DocumentSession fed the canonical log, so it computes the inverse locally
   * and the inverse rides as an ordinary sealed intent.
   */
  const bodyText = (c: EncryptedCollabConnection): string => {
    const walk = (e: { name: string; text: string; children: unknown[] }): string =>
      (e.name.endsWith(":t") ? e.text : "") + (e.children as never[]).map(walk).join("");
    return walk(c.doc!.docRoot as never);
  };
  const ins = (at: number, t: string) =>
    ({ kind: "insertText", at: { blockId: 1, runId: 2, offset: at }, text: t }) as never;

  async function room() {
    const docKey = mintDocKey();
    const { checkpoint } = await seedEncrypted("ab", "g1", docKey);
    const srv = blindServer("g1", checkpoint);
    const a = new EncryptedCollabConnection(srv.attach(), "alice", docKey);
    const b = new EncryptedCollabConnection(srv.attach(), "bob", docKey);
    a.join("d");
    b.join("d");
    await until(() => a.ready && b.ready, "both clients to rehydrate");
    return { a, b };
  }

  it("A's undo reverses A's edit on BOTH clients", async () => {
    const { a, b } = await room();
    a.submit(ins(2, "HELLO"));
    await until(() => bodyText(b).includes("HELLO"), "B to see A's text");

    expect(a.undoLast()).toBe("undone");
    await until(() => !bodyText(b).includes("HELLO"), "B to see the undo");
    expect(bodyText(a)).toBe("ab");
    expect(bodyText(b)).toBe("ab");
  });

  it("undo is REBASED through another client's concurrent edit", async () => {
    // B types in front of A's text; A's undo must still remove exactly A's
    // characters, not five characters at the old offset.
    const { a, b } = await room();
    a.submit(ins(2, "HELLO"));
    await until(() => bodyText(b).includes("HELLO"), "B to see it");
    b.submit(ins(0, "ZZ"));
    await until(() => bodyText(a).includes("ZZ"), "A to see B's text");

    expect(a.undoLast()).toBe("undone");
    // THE PAINT IS CORRECT IMMEDIATELY — this is what rebasing against the
    // mirror buys, and the only assertion that can tell the two designs
    // apart. Submitting the inverse raw at its original base also CONVERGES
    // (the canonical transform rebases it on every replica), so a
    // convergence-only assertion passes either way; what it does not do is
    // paint correctly in the meantime. Without the local rebase this reads
    // "ZZLO" — A watches five characters of someone else's text disappear and
    // come back.
    expect(bodyText(a)).toBe("ZZab");
    // Then wait for CONVERGENCE, not for A's own paint. An earlier version of this
    // test waited on `!bodyText(a).includes("HELLO")`, which the optimistic
    // apply satisfies instantly — and it passed while the canonical state had
    // no undo at all and A's local text read "ZZLO". Waiting on both clients
    // agreeing is what makes this a statement about the room.
    await until(() => bodyText(a) === "ZZab" && bodyText(b) === "ZZab", "the room to converge on ZZab");
    expect(bodyText(a)).toBe("ZZab");   // B's edit survives, A's is gone
    expect(bodyText(b)).toBe("ZZab");
  });

  it("selective: A's undo does NOT touch B's edit", async () => {
    const { a, b } = await room();
    a.submit(ins(2, "AAA"));
    await until(() => bodyText(b).includes("AAA"), "B to see A's text");
    b.submit(ins(2, "BBB"));
    await until(() => bodyText(a).includes("BBB"), "A to see B's text");

    expect(a.undoLast()).toBe("undone");
    await until(() => !bodyText(a).includes("AAA"), "A's undo to land");
    expect(bodyText(a)).toContain("BBB"); // B's work untouched
    await until(() => bodyText(b) === bodyText(a), "both to converge");
  });

  it("reports nothing-to-undo rather than doing something", async () => {
    const { a } = await room();
    expect(a.undoLast()).toBe("nothing-to-undo");
    expect(bodyText(a)).toBe("ab");
  });

  it("undo twice walks back two actions", async () => {
    const { a, b } = await room();
    a.submit(ins(2, "ONE"));
    await until(() => bodyText(a).includes("ONE"), "first edit");
    a.submit(ins(5, "TWO"));
    await until(() => bodyText(b).includes("TWO"), "B to see both");
    // undoLast() takes from A's MIRROR, which advances only by ingesting
    // sequenced envelopes, and A paints her own edits optimistically before
    // that. B seeing TWO says nothing about A's mirror: the two clients drain
    // independent async chains and A's also carries her own sealIntent work,
    // so A routinely finishes last. Without this wait a slow A undoes ONE —
    // the only entry her stack holds — and the wait below can never pass.
    await until(() => a.pendingCount === 0, "A's mirror to confirm both edits");

    expect(a.undoLast()).toBe("undone");
    await until(() => !bodyText(a).includes("TWO"), "second undo target gone");
    // The first undo is itself an intent in flight now, so the second press
    // would be HELD by the pending gate (see undoLast, and the accumulation
    // case in undo-pending-gate.test.ts). Waiting for the drain keeps this
    // test on its own subject: that undo walks BACK TWO ACTIONS, not that a
    // press mid-flight is deferred.
    await until(() => a.pendingCount === 0, "the first undo to confirm");
    expect(a.undoLast()).toBe("undone");
    await until(() => !bodyText(a).includes("ONE"), "first undo target gone");
    expect(bodyText(a)).toBe("ab");
    await until(() => bodyText(b) === "ab", "B to converge");
  });

  it("declines up front when the target is already gone — no flash", async () => {
    // The pre-check earning its keep: B deletes A's text, then A presses undo.
    // Painting it optimistically would show the text return and vanish again.
    const { a, b } = await room();
    a.submit(ins(2, "GONE"));
    await until(() => bodyText(b).includes("GONE"), "B to see it");
    b.submit({ kind: "deleteText", blockId: 1, runId: 2, start: 2, end: 6 } as never);
    await until(() => !bodyText(a).includes("GONE"), "A to see the delete");

    const before = bodyText(a);
    const outcome = a.undoLast();
    // Either provably-dead (declined, nothing painted) or a genuine attempt
    // that the canonical apply rejects — never a silent local divergence.
    expect(["changed-since", "undone"]).toContain(outcome);
    await until(() => bodyText(b) === bodyText(a), "the room to agree");
    if (outcome === "changed-since") expect(bodyText(a)).toBe(before); // nothing painted
  });
});
