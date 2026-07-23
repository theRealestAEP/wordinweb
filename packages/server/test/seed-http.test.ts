import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub, EVICTION_GRACE_MS } from "../src/hub.js";
import { handleSeedRequest } from "../src/seed-http.js";
import { PROTOCOL_VERSION } from "@wordinweb/collab/server";
import type { ServerMessage } from "@wordinweb/collab/server";
import type { Connection } from "../src/hub.js";

/** Minimal valid docx (same fixture shape as hub.test.ts). */
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
const b64 = (b: Uint8Array) => Buffer.from(b).toString("base64");

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

const zeroCustodyHub = (now: () => number = () => 0) => {
  let n = 0;
  return new CollabHub(null, undefined, undefined, now, () => `g_${n++}`);
};

describe("handleSeedRequest (doc 12 §3 go-live / bring-back)", () => {
  it("POST mints an unguessable docId, seeds, and returns 201 with the epoch", async () => {
    const hub = zeroCustodyHub();
    const res = handleSeedRequest(hub, { method: "POST", body: { docx: b64(docxBytes("hello")) } });
    expect(res.status).toBe(201);
    const docId = res.body.docId as string;
    expect(docId).toMatch(/^d_[0-9a-f]{32}$/); // ≥128-bit random (doc 11)
    expect(res.body.genesisId).toBe("g_0");
    // The link now works: a hello reaches a live session with the bytes.
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId, clientId: "a", sinceSeq: 0 });
    expect(c.last().t).toBe("welcome");
  });

  it("PUT revives a known docId; a concurrent second PUT gets 409 with the winner's epoch", () => {
    const hub = zeroCustodyHub();
    const first = handleSeedRequest(hub, { method: "PUT", docId: "d_link", body: { docx: b64(docxBytes("mine")) } });
    expect(first.status).toBe(200);
    const second = handleSeedRequest(hub, { method: "PUT", docId: "d_link", body: { docx: b64(docxBytes("theirs")) } });
    expect(second.status).toBe(409);
    expect(second.body.genesisId).toBe(first.body.genesisId); // join the winner (case 2)
  });

  it("PUT without a docId, missing docx, bad base64: 400s, and no room is created", () => {
    const hub = zeroCustodyHub();
    expect(handleSeedRequest(hub, { method: "PUT", body: { docx: b64(docxBytes("x")) } }).status).toBe(400);
    expect(handleSeedRequest(hub, { method: "POST", body: {} }).status).toBe(400);
    expect(handleSeedRequest(hub, { method: "POST", body: { docx: "!!!not-base64!!!" } }).status).toBe(400);
    expect(hub.activeDocs()).toEqual([]);
  });

  it("rejects an oversize seed with 413 BEFORE parsing (doc 11 size caps)", () => {
    const hub = zeroCustodyHub();
    const res = handleSeedRequest(
      hub,
      { method: "POST", body: { docx: b64(docxBytes("x".repeat(2000))) } },
      { maxDocxBytes: 100 },
    );
    expect(res.status).toBe(413);
    expect(hub.activeDocs()).toEqual([]);
  });

  it("rejects unparseable bytes with 400 and creates nothing (validation gate)", () => {
    const hub = zeroCustodyHub();
    const res = handleSeedRequest(hub, { method: "POST", body: { docx: b64(new Uint8Array([1, 2, 3, 4])) } });
    expect(res.status).toBe(400);
    expect(res.body.error).toBe("invalid-docx");
    expect(hub.activeDocs()).toEqual([]);
  });

  it("full lifecycle over HTTP semantics: seed → evict → 404-equivalent no-session → PUT revives with a fresh epoch", async () => {
    let now = 0;
    const hub = zeroCustodyHub(() => now);
    const created = handleSeedRequest(hub, { method: "POST", body: { docx: b64(docxBytes("v1")) } });
    const docId = created.body.docId as string;
    // Nobody joins; grace passes; zero custody erases it.
    now += EVICTION_GRACE_MS;
    expect(hub.sweepRooms()).toEqual([docId]);
    // Revival by a bundle-holder: same docId, NEW epoch.
    const revived = handleSeedRequest(hub, { method: "PUT", docId, body: { docx: b64(docxBytes("v1")) } });
    expect(revived.status).toBe(200);
    expect(revived.body.genesisId).not.toBe(created.body.genesisId);
  });
});
