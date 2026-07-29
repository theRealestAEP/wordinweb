import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub, Connection } from "../src/hub.js";
import { CollabConnection } from "@wordinweb/collab/client";
import type { ClientMessage, ServerMessage } from "@wordinweb/collab/server";
import { PROTOCOL_VERSION } from "@wordinweb/collab/server";

/**
 * Cross-layer media round (doc 16 item 7): the full life of one image —
 * reservation intent through the REAL hub (skeleton state on every
 * replica), placer upload (hash-verified), peer need→ready→download→
 * client-side verify→install — ending with both replicas holding
 * byte-identical pixels AND byte-identical documents.
 */

function blankDoc(): Uint8Array {
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

/** Synchronous hub↔connection bridge (the loopback pattern). */
function bridge(hub: CollabHub, connId: string) {
  const listeners: ((m: ServerMessage) => void)[] = [];
  const conn: Connection = { id: connId, send: (m) => listeners.forEach((l) => l(m)) };
  return {
    transport: {
      send: (m: ClientMessage) => void hub.handle(conn, m),
      onMessage: (cb: (m: ServerMessage) => void) => listeners.push(cb),
    },
  };
}

const PIXELS = new Uint8Array([137, 80, 78, 71, 1, 2, 3, 4]);
async function shaOf(bytes: Uint8Array): Promise<string> {
  const d = await crypto.subtle.digest("SHA-256", bytes as unknown as ArrayBuffer);
  let hex = "";
  for (const b of new Uint8Array(d)) hex += b.toString(16).padStart(2, "0");
  return hex;
}
const flush = () => new Promise((r) => setTimeout(r, 10));

describe("media round through the real hub (doc 16 item 7)", () => {
  it("reservation → skeleton everywhere → upload → need/ready → verified install → byte-identical docs + pixels", async () => {
    const hub = new CollabHub(null);
    hub.seed("d", blankDoc());
    const sha = await shaOf(PIXELS);

    const alice = new CollabConnection(bridge(hub, "s1").transport, "alice");
    let bobReady: string | null = null;
    const bob = new CollabConnection(bridge(hub, "s2").transport, "bob", {
      onMediaReady: (s) => (bobReady = s),
    });
    alice.join("d");
    bob.join("d");
    await flush(); // joins are async through the hub bridge

    // 1. Reservation: the insert intent sequences; BOTH replicas register
    //    the part pending with the committed sha (skeleton state).
    alice.submit({
      kind: "insertImage", runId: 2, blobSha: sha, bytesLen: PIXELS.length,
      ext: "png", widthPx: 8, heightPx: 8, nodeIds: [900],
    } as never);
    await flush();
    const alicePart = [...alice.doc!.pendingMedia.keys()][0];
    const bobPart = [...bob.doc!.pendingMedia.keys()][0];
    expect(alicePart).toBe(bobPart);
    expect(bob.doc!.pendingMedia.get(bobPart)!.sha).toBe(sha);

    // 2. The placer uploads (hash-verified by the relay) and installs its
    //    own copy directly (it always has the bytes).
    expect(await hub.mediaUpload("d", sha, PIXELS)).toBe(201);
    alice.doc!.installMedia(alicePart, PIXELS);

    // 3. Bob needs it: need → ready → HTTP download → CLIENT-side verify
    //    against the committed sha (defense in depth behind the relay's
    //    check) → install.
    bob.mediaNeed(sha);
    await flush();
    expect(bobReady).toBe(sha);
    // RAM tier (spill disabled here), so the hit is the bytes themselves.
    const fetched = hub.mediaDownload("d", sha) as Uint8Array;
    expect(fetched).toBeInstanceOf(Uint8Array);
    expect(await shaOf(fetched)).toBe(sha); // the reservation IS the verifier
    bob.doc!.installMedia(bobPart, fetched);

    // 4. End state: both ready, byte-identical pixels, byte-identical docs.
    expect(alice.doc!.mediaStatus(alicePart)).toBe("ready");
    expect(bob.doc!.mediaStatus(bobPart)).toBe("ready");
    expect(Buffer.from(bob.doc!.media(bobPart)!).equals(Buffer.from(alice.doc!.media(alicePart)!))).toBe(true);
    expect(Buffer.from(bob.doc!.save()).equals(Buffer.from(alice.doc!.save()))).toBe(true);
  });
});
