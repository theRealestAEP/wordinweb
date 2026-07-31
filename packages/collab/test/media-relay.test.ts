import { describe, expect, it } from "vitest";
import { DocxDocument } from "@wordinweb/core";
import { zipSync, strToU8 } from "fflate";
import { MediaClient, sha256Hex, type FetchLike, type MediaState } from "../src/media.js";
import { deriveEpochKeys, mintDocKey, sealMediaBlob, openMediaBlob } from "../src/e2ee.js";
import { DocumentSession } from "../src/session.js";

/**
 * The doc-16 §5 CLIENT loop, which was specced and left unbuilt while the
 * server relay shipped: upload-then-intent, eager fetch, verify-before-
 * install, holder re-supply.
 *
 * The single rule every test here is really about (doc 16 §1.1): NOBODY IS
 * EVER TOLD A HASH. The placer commits `blobSha` inside the sequenced
 * intent; after that, bytes are only ever accepted because the receiver
 * re-derived sha256 itself and matched that commitment. A relay that lies,
 * a peer that substitutes, a corrupted local copy — all fail the same check.
 */

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

const PIXELS = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 9, 8, 7, 6]);

/**
 * An in-memory stand-in for the relay's HTTP surface, with the ONE behavior
 * that matters for the security argument: it hashes every PUT body itself
 * and refuses a body that doesn't match the URL's address. That is what
 * makes occupying an address with other bytes a sha256-preimage problem
 * rather than a race (doc 16 §1.1 / §3).
 */
function fakeRelay() {
  const blobs = new Map<string, Uint8Array>();
  let puts = 0;
  let gets = 0;
  const fetchImpl: FetchLike = async (url, init) => {
    const sha = url.slice(url.lastIndexOf("/") + 1);
    if (init?.method === "PUT") {
      puts++;
      const body = init.body!;
      if ((await sha256Hex(body)) !== sha) return { ok: false, status: 400, arrayBuffer: async () => new ArrayBuffer(0) };
      const already = blobs.has(sha);
      blobs.set(sha, body);
      return { ok: true, status: already ? 200 : 201, arrayBuffer: async () => new ArrayBuffer(0) };
    }
    gets++;
    const hit = blobs.get(sha);
    if (!hit) return { ok: false, status: 404, arrayBuffer: async () => new ArrayBuffer(0) };
    return { ok: true, status: 200, arrayBuffer: async () => hit.slice().buffer as ArrayBuffer };
  };
  return {
    fetchImpl, blobs,
    counts: () => ({ puts, gets }),
    evict: (sha: string) => blobs.delete(sha),
    /** A hostile relay: serve bytes that are NOT what the address names. */
    poison: (sha: string, bytes: Uint8Array) => blobs.set(sha, bytes),
  };
}

/** A doc with one registered-but-empty image part, as insertImage leaves it. */
function docWithPendingImage(sha: string, iv?: string): DocxDocument {
  const doc = DocxDocument.load(blankDocx());
  doc.registerPendingImage(sha, "png", { iv });
  return doc;
}

interface Harness {
  client: MediaClient;
  needs: string[];
  haves: string[][];
  states: [string, MediaState][];
  changes: () => number;
}
function harness(doc: () => DocxDocument | null, relay: ReturnType<typeof fakeRelay>, crypto?: {
  seal: (p: Uint8Array, iv?: string) => Promise<{ blob: Uint8Array; iv: string }>;
  open: (b: Uint8Array, iv: string) => Promise<Uint8Array>;
}): Harness {
  const needs: string[] = [];
  const haves: string[][] = [];
  const states: [string, MediaState][] = [];
  let changes = 0;
  const client = new MediaClient(
    { httpBase: "http://relay", fetchImpl: relay.fetchImpl },
    doc,
    {
      onChange: () => { changes++; },
      onState: (part, state) => states.push([part, state]),
      need: (sha) => needs.push(sha),
      have: (shas) => haves.push(shas),
    },
    crypto,
  );
  return { client, needs, haves, states, changes: () => changes };
}

/** Wait for a condition rather than a fixed delay: the fetch path is a
 * multi-await promise chain, and a fixed sleep is exactly the kind of
 * load-sensitive flake that passes alone and fails in a full run. */
async function until(cond: () => boolean, label = "condition"): Promise<void> {
  for (let i = 0; i < 200; i++) {
    if (cond()) return;
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error(`timed out waiting for ${label}`);
}
const settle = () => new Promise((r) => setTimeout(r, 5));

describe("media client — plaintext relay round trip", () => {
  it("placer uploads at the blob's own address, then a receiver eager-fetches and installs", async () => {
    const relay = fakeRelay();
    // PLACER: upload BEFORE any intent (doc 16 §5.1 ordering).
    const placerDoc = DocxDocument.load(blankDocx());
    const placer = harness(() => placerDoc, relay);
    const fields = await placer.client.upload("d", PIXELS);
    expect(fields).not.toBeNull();
    expect(fields!.blobSha).toBe(await sha256Hex(PIXELS));
    expect(fields!.bytesLen).toBe(PIXELS.length);
    expect(fields!.iv).toBeUndefined(); // plaintext: no envelope

    // RECEIVER: the intent applied, leaving a registered hole.
    const rxDoc = docWithPendingImage(fields!.blobSha);
    const part = [...rxDoc.pendingMedia.keys()][0];
    expect(rxDoc.mediaStatus(part)).toBe("pending");
    const rx = harness(() => rxDoc, relay);
    await rx.client.fetchPending("d");
    await until(() => rxDoc.mediaStatus(part) === "ready", "eager fetch to install");

    expect(rxDoc.mediaStatus(part)).toBe("ready");
    expect(Buffer.from(rxDoc.media(part)!)).toEqual(Buffer.from(PIXELS));
    expect(rx.needs).toEqual([]);   // relay had it: no re-supply flow needed
    expect(rx.changes()).toBe(1);   // one repaint
  });

  it("a relay refusal means NO intent fields — the reservation is never made", async () => {
    // The ordering rule exists for this: an insertImage whose blob was
    // refused would leave every replica with a skeleton nobody can ever fill.
    const relay = fakeRelay();
    const rejecting = { ...relay, fetchImpl: (async () => ({ ok: false, status: 507, arrayBuffer: async () => new ArrayBuffer(0) })) as FetchLike };
    const doc = DocxDocument.load(blankDocx());
    const h = harness(() => doc, rejecting as never);
    expect(await h.client.upload("d", PIXELS)).toBeNull();
  });

  it("a relay miss registers a waiter, and media-ready completes the install", async () => {
    const relay = fakeRelay();
    const sha = await sha256Hex(PIXELS);
    const rxDoc = docWithPendingImage(sha);
    const part = [...rxDoc.pendingMedia.keys()][0];
    const rx = harness(() => rxDoc, relay);

    await rx.client.fetchPending("d");  // relay is empty
    await settle();
    expect(rx.needs).toEqual([sha]);            // asked the room
    expect(rx.client.stateOf(part)).toBe("waiting");
    expect(rxDoc.mediaStatus(part)).toBe("pending");

    // A holder re-supplied; the server says it's fetchable now.
    relay.blobs.set(sha, PIXELS);
    await rx.client.onReady("d", sha);
    await until(() => rxDoc.mediaStatus(part) === "ready", "media-ready refetch");
    expect(rx.client.stateOf(part)).toBeUndefined();
  });

  it("media-unavailable marks the skeleton, and a later ready still recovers it", async () => {
    const relay = fakeRelay();
    const sha = await sha256Hex(PIXELS);
    const rxDoc = docWithPendingImage(sha);
    const part = [...rxDoc.pendingMedia.keys()][0];
    const rx = harness(() => rxDoc, relay);
    await rx.client.fetchPending("d");
    await until(() => rx.client.stateOf(part) === "waiting", "the relay miss");

    rx.client.onUnavailable(sha);
    expect(rx.client.stateOf(part)).toBe("unavailable");
    // The REGISTRATION survives — this is a display state, not a deletion.
    expect(rxDoc.pendingMedia.has(part)).toBe(true);

    relay.blobs.set(sha, PIXELS);
    await rx.client.onReady("d", sha);
    await until(() => rxDoc.mediaStatus(part) === "ready", "recovery after unavailable");
  });
});

describe("media client — verification (nobody is ever told a hash)", () => {
  it("a hostile relay serving different bytes at the address installs NOTHING", async () => {
    const relay = fakeRelay();
    const sha = await sha256Hex(PIXELS);
    // The relay itself is compromised — its own admission check bypassed.
    relay.poison(sha, new Uint8Array([1, 1, 1, 1]));
    const rxDoc = docWithPendingImage(sha);
    const part = [...rxDoc.pendingMedia.keys()][0];
    const rx = harness(() => rxDoc, relay);

    await rx.client.fetchPending("d");
    await until(() => rx.client.stateOf(part) === "waiting", "the rejection");
    // The receiver re-derived the hash and refused: skeleton intact, no
    // wrong pixels anywhere in the document.
    expect(rxDoc.mediaStatus(part)).toBe("pending");
    expect(rx.changes()).toBe(0);
  });

  it("the relay refuses a PUT whose body does not hash to the address", async () => {
    const relay = fakeRelay();
    const doc = DocxDocument.load(blankDocx());
    const h = harness(() => doc, relay);
    // Address the upload at someone else's sha by hand.
    const wrongSha = await sha256Hex(new Uint8Array([9, 9, 9]));
    const res = await (await import("../src/media.js")).putBlob(
      { httpBase: "http://relay", fetchImpl: relay.fetchImpl }, "d", wrongSha, PIXELS,
    );
    expect(res.status).toBe(400);
    expect(relay.blobs.has(wrongSha)).toBe(false);
    void h;
  });
});

describe("media client — E2EE envelope and deterministic re-supply", () => {
  async function e2eeCrypto() {
    const docKey = await mintDocKey();
    const keys = await deriveEpochKeys(docKey, "genesis-1");
    return {
      seal: (p: Uint8Array, iv?: string) => sealMediaBlob(keys.kMedia, p, iv),
      open: (b: Uint8Array, iv: string) => openMediaBlob(keys.kMedia, b, iv),
    };
  }

  it("the sha addresses the CIPHERTEXT, and a receiver verifies before decrypting", async () => {
    const relay = fakeRelay();
    const crypto = await e2eeCrypto();
    const placerDoc = DocxDocument.load(blankDocx());
    const placer = harness(() => placerDoc, relay, crypto);
    const fields = await placer.client.upload("d", PIXELS);
    expect(fields).not.toBeNull();
    expect(fields!.iv).toBeTruthy();               // the IV rides in the intent
    expect(fields!.blobSha).not.toBe(await sha256Hex(PIXELS)); // NOT the plaintext's hash
    expect(fields!.blobSha).toBe(await sha256Hex(relay.blobs.get(fields!.blobSha)!));
    // The relay holds ciphertext only — it never sees the image.
    expect(Buffer.from(relay.blobs.get(fields!.blobSha)!)).not.toEqual(Buffer.from(PIXELS));

    const rxDoc = docWithPendingImage(fields!.blobSha, fields!.iv);
    const part = [...rxDoc.pendingMedia.keys()][0];
    const rx = harness(() => rxDoc, relay, crypto);
    await rx.client.fetchPending("d");
    await until(() => rxDoc.mediaStatus(part) === "ready", "decrypt+install");
    // Decrypted back to the exact original pixels.
    expect(Buffer.from(rxDoc.media(part)!)).toEqual(Buffer.from(PIXELS));
  });

  it("a tampered ciphertext fails the GCM tag and keeps the skeleton", async () => {
    const relay = fakeRelay();
    const crypto = await e2eeCrypto();
    const sealed = await crypto.seal(PIXELS);
    const sha = await sha256Hex(sealed.blob);
    // Flip a bit but keep the ADDRESS honest, so the sha check passes and
    // only the authenticated-decryption step can catch it.
    const tampered = sealed.blob.slice();
    tampered[0] ^= 0x01;
    relay.blobs.set(sha, tampered);
    const rxDoc = docWithPendingImage(sha, sealed.iv);
    const part = [...rxDoc.pendingMedia.keys()][0];
    const rx = harness(() => rxDoc, relay, crypto);

    await rx.client.fetchPending("d");
    await until(() => rx.client.stateOf(part) === "waiting", "the GCM rejection");
    expect(rxDoc.mediaStatus(part)).toBe("pending");
    expect(rx.changes()).toBe(0);
  });

  it("HOLDER RE-SUPPLY re-seals with the RECORDED iv and reproduces the address", async () => {
    // The heart of doc 16 §5.3. A fresh IV would produce a perfectly valid
    // blob at the WRONG address; only reusing the recorded IV reproduces the
    // ciphertext the reservation committed to.
    const relay = fakeRelay();
    const crypto = await e2eeCrypto();
    const holderDoc = DocxDocument.load(blankDocx());
    const holder = harness(() => holderDoc, relay, crypto);
    const fields = (await holder.client.upload("d", PIXELS))!;
    // The holder has the part READY with the recorded metadata.
    holderDoc.registerPendingImage(fields.blobSha, "png", { iv: fields.iv });
    const part = [...holderDoc.pendingMedia.keys()][0];
    holderDoc.installMedia(part, PIXELS);
    expect(holderDoc.mediaStatus(part)).toBe("ready");

    // The relay evicted it (TTL) and the holder is chosen to re-supply.
    relay.evict(fields.blobSha);
    expect(await holder.client.resupply("d", fields.blobSha)).toBe(true);
    // Byte-identical ciphertext is back at the same address.
    expect(await sha256Hex(relay.blobs.get(fields.blobSha)!)).toBe(fields.blobSha);

    // And the counter-proof for the warning: a FRESH iv changes the address.
    const fresh = await crypto.seal(PIXELS);
    expect(await sha256Hex(fresh.blob)).not.toBe(fields.blobSha);
  });

  it("re-supply refuses to upload locally-corrupt pixels", async () => {
    const relay = fakeRelay();
    const crypto = await e2eeCrypto();
    const doc = DocxDocument.load(blankDocx());
    const h = harness(() => doc, relay, crypto);
    const fields = (await h.client.upload("d", PIXELS))!;
    doc.registerPendingImage(fields.blobSha, "png", { iv: fields.iv });
    const part = [...doc.pendingMedia.keys()][0];
    doc.installMedia(part, new Uint8Array([4, 4, 4, 4])); // not the original image
    relay.evict(fields.blobSha);

    const before = relay.counts().puts;
    expect(await h.client.resupply("d", fields.blobSha)).toBe(false);
    expect(relay.counts().puts).toBe(before); // never uploaded garbage
  });
});

describe("media client — holder duties", () => {
  it("volunteers only the intersection of the room's needs and local holdings", async () => {
    const relay = fakeRelay();
    const sha = await sha256Hex(PIXELS);
    const doc = DocxDocument.load(blankDocx());
    doc.registerPendingImage(sha, "png");
    const part = [...doc.pendingMedia.keys()][0];
    doc.installMedia(part, PIXELS);
    const h = harness(() => doc, relay);

    h.client.volunteer([sha, "f".repeat(64)]);
    expect(h.haves).toEqual([[sha]]);         // only what we actually hold
    h.client.answerRequest("e".repeat(64));
    expect(h.haves.length).toBe(1);           // silent about what we don't
    h.client.answerRequest(sha);
    expect(h.haves[1]).toEqual([sha]);
  });

  it("a part still PENDING is never volunteered (a hole cannot serve a hole)", async () => {
    const relay = fakeRelay();
    const sha = await sha256Hex(PIXELS);
    const doc = docWithPendingImage(sha);
    const h = harness(() => doc, relay);
    expect(h.client.heldShas()).toEqual([]);
    h.client.volunteer([sha]);
    expect(h.haves).toEqual([]);
  });
});

describe("media consistency (doc 16 §6)", () => {
  it("replicas agree on the document while one still shows a skeleton", async () => {
    // The convergence claim media rests on: byte arrival is NOT part of the
    // document's identity, so a filled replica and a skeleton replica hold
    // the same document.
    const sha = await sha256Hex(PIXELS);
    const filled = docWithPendingImage(sha);
    const skeleton = docWithPendingImage(sha);
    const part = [...filled.pendingMedia.keys()][0];
    filled.installMedia(part, PIXELS);

    const session = new DocumentSession(filled); // exercises the real id table
    void session;
    const roots = (d: DocxDocument) => d.editableRoots().map((r) => r.name).join("|");
    expect(roots(filled)).toBe(roots(skeleton));
    expect(filled.mediaStatus(part)).toBe("ready");
    expect(skeleton.mediaStatus(part)).toBe("pending");
    // Same registration, same declared address — the divergence-relevant state.
    expect(skeleton.mediaMeta.get(part)!.sha).toBe(filled.mediaMeta.get(part)!.sha);
  });
});
