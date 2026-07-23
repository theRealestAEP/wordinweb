import { describe, expect, it } from "vitest";
import { DocumentSession } from "../src/session.js";
import { CollabConnection } from "../src/connection.js";
import { deriveEpochKeys, mintDocKey, sealMediaBlob, openMediaBlob, bytesToB64 } from "../src/e2ee.js";
import { DocxDocument } from "@wordinweb/core";
import { zipSync, strToU8 } from "fflate";
import type { ClientMessage, ServerMessage } from "../src/protocol.js";

/** Client media duties (doc 16 §5) + E2EE media envelopes (§6). */

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

const SHA = "b".repeat(64);
const insertImage = {
  kind: "insertImage", clientId: "a", clientSeq: 1, base: 0,
  runId: 2, blobSha: SHA, bytesLen: 5, ext: "png", widthPx: 10, heightPx: 10, nodeIds: [700],
} as never;

function miniServer(opts: { mediaNeeded?: string[] } = {}) {
  const session = new DocumentSession(DocxDocument.load(blankDocx()));
  const sent: ClientMessage[] = [];
  const peers: { deliver: (m: ServerMessage) => void }[] = [];
  const attach = () => {
    const peer = { deliver: (_m: ServerMessage) => {} };
    peers.push(peer);
    return {
      send: (msg: ClientMessage) => {
        sent.push(msg);
        if (msg.t === "hello") {
          const cp = session.checkpoint();
          peer.deliver({ t: "welcome", docId: "d", seq: cp.seq, snapshot: Buffer.from(cp.docx).toString("base64"),
            sidecar: cp.sidecar, tail: [], genesisId: "g1", mode: "plaintext", mediaNeeded: opts.mediaNeeded });
        } else if (msg.t === "submit") {
          const entry = session.submit(msg.intent);
          for (const p of peers) p.deliver({ t: "broadcast", entries: [entry] });
        }
      },
      onMessage: (cb: (m: ServerMessage) => void) => { peer.deliver = cb; },
    };
  };
  return { session, attach, sent, push: (m: ServerMessage) => peers.forEach((p) => p.deliver(m)) };
}

describe("client media duties (doc 16 §5)", () => {
  it("an applied insertImage registers pending metadata; install flips it to a held (re-suppliable) sha", () => {
    const srv = miniServer();
    const conn = new CollabConnection(srv.attach(), "alice");
    conn.join("d");
    conn.submit(insertImage);
    const doc = conn.doc!;
    const [part, meta] = [...doc.pendingMedia][0];
    expect(meta.sha).toBe(SHA);
    expect(conn.heldMediaShas()).toEqual([]); // pending ≠ holdable
    doc.installMedia(part, new Uint8Array([1, 2, 3, 4, 5]));
    expect(conn.heldMediaShas()).toEqual([SHA]); // mediaMeta persists past install
  });

  it("media control fan-in: request/ready/unavailable surface; media-upload surfaces as a request", () => {
    const srv = miniServer();
    const seen: string[] = [];
    const conn = new CollabConnection(srv.attach(), "alice", {
      onMediaRequest: (sha) => seen.push(`req:${sha}`),
      onMediaReady: (sha) => seen.push(`ready:${sha}`),
      onMediaUnavailable: (sha) => seen.push(`gone:${sha}`),
    });
    conn.join("d");
    srv.push({ t: "media-request", sha: SHA });
    srv.push({ t: "media-upload", sha: SHA });
    srv.push({ t: "media-ready", sha: SHA });
    srv.push({ t: "media-unavailable", sha: SHA });
    expect(seen).toEqual([`req:${SHA}`, `req:${SHA}`, `ready:${SHA}`, `gone:${SHA}`]);
  });

  it("a RESUMING holder volunteers the intersection of welcome.mediaNeeded with its bundle's media metadata (§5.3+§5.4)", () => {
    // Session 1: become a holder (insert + install), then "close the tab" —
    // the bundle carries the docx pixels AND the media ADDRESSES (§5.3:
    // metadata is in-memory on the doc; without the bundle field a resumed
    // holder could neither volunteer nor verify).
    const srv1 = miniServer();
    const conn1 = new CollabConnection(srv1.attach(), "alice");
    conn1.join("d");
    conn1.submit(insertImage);
    const [part] = [...conn1.doc!.pendingMedia][0];
    conn1.doc!.installMedia(part, new Uint8Array([1, 2, 3, 4, 5]));
    // Any later confirmed advance re-snapshots — pixels enter the bundle.
    conn1.submit({ kind: "insertText", clientId: "alice", clientSeq: 2, base: 1, at: { blockId: 1, runId: 2, offset: 0 }, text: "x" } as never);
    const bundle = conn1.exportBundle("d")!;
    expect(bundle.mediaMeta?.map(([, m]) => m.sha)).toEqual([SHA]);

    // Later: the room (same epoch) has waiters on SHA + an unrelated sha.
    const srv2 = miniServer({ mediaNeeded: [SHA, "f".repeat(64)] });
    const conn2 = new CollabConnection(srv2.attach(), "alice");
    conn2.resume(bundle);
    const have = srv2.sent.find((m) => m.t === "media-have");
    expect(have && have.t === "media-have" ? have.shas : []).toEqual([SHA]); // intersection only
  });
});

describe("E2EE media envelopes (doc 16 §6)", () => {
  it("seal/open round-trips; re-sealing with the RECORDED iv is byte-identical (re-supply determinism)", async () => {
    const keys = await deriveEpochKeys(mintDocKey(), "g1");
    const pixels = new Uint8Array([9, 8, 7, 6, 5]);
    const first = await sealMediaBlob(keys.kMedia, pixels);
    expect(await openMediaBlob(keys.kMedia, first.blob, first.iv)).toEqual(pixels);
    // A holder re-seals the same plaintext with the recorded iv: the blob
    // is BYTE-IDENTICAL, so its sha matches the committed address.
    const again = await sealMediaBlob(keys.kMedia, pixels, first.iv);
    expect(bytesToB64(again.blob)).toBe(bytesToB64(first.blob));
    // A fresh iv (the naive "fix") breaks the address — pinned.
    const fresh = await sealMediaBlob(keys.kMedia, pixels);
    expect(bytesToB64(fresh.blob)).not.toBe(bytesToB64(first.blob));
    // Tampered blob refuses to open (GCM tag).
    const flipped = new Uint8Array(first.blob);
    flipped[0] ^= 1;
    await expect(openMediaBlob(keys.kMedia, flipped, first.iv)).rejects.toThrow();
  });
});
