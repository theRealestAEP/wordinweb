import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub, EVICTION_GRACE_MS } from "../src/hub.js";
import { handleSeedRequest } from "../src/seed-http.js";
import { PROTOCOL_VERSION, ENGINE_VERSION } from "@wordinweb/collab/server";
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
    // Roster fan-outs are ambient (fired on every join/leave/rename) —
    // skip them so assertions target the response to the acted-on message.
    for (let i = this.received.length - 1; i >= 0; i--) {
      if (this.received[i].t !== "roster") return this.received[i];
    }
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

describe("blank go-live (the demo's New document flow)", () => {
  it("POST {blank:true} seeds a working blank session", async () => {
    const hub = zeroCustodyHub();
    const res = handleSeedRequest(hub, { method: "POST", body: { blank: true } });
    expect(res.status).toBe(201);
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: res.body.docId as string, clientId: "a", sinceSeq: 0 });
    expect(c.last().t).toBe("welcome");
  });

  it("blank is POST-only (PUT revival must supply real bytes)", () => {
    const hub = zeroCustodyHub();
    const res = handleSeedRequest(hub, { method: "PUT", docId: "d", body: { blank: true } });
    expect(res.status).toBe(400);
  });
});

describe("encrypted seed over HTTP (doc 13)", () => {
  const CP = { seq: 0, iv: "aXY=", ciphertext: "b3BhcXVl" };
  it("PUT with an encrypted body seeds a blind room; second PUT gets 409 with the winner's epoch", () => {
    const hub = zeroCustodyHub();
    const first = handleSeedRequest(hub, { method: "PUT", docId: "d_x", body: { encrypted: { genesisId: "g_a", checkpoint: CP } } });
    expect(first.status).toBe(200);
    expect(first.body.genesisId).toBe("g_a"); // client-minted epoch honored
    const second = handleSeedRequest(hub, { method: "PUT", docId: "d_x", body: { encrypted: { genesisId: "g_b", checkpoint: CP } } });
    expect(second.status).toBe(409);
    expect(second.body.genesisId).toBe("g_a");
  });

  it("malformed encrypted bodies 400; oversized ciphertext 413; nothing is created", () => {
    const hub = zeroCustodyHub();
    expect(handleSeedRequest(hub, { method: "PUT", docId: "d", body: { encrypted: { genesisId: "", checkpoint: CP } } }).status).toBe(400);
    expect(handleSeedRequest(hub, { method: "PUT", docId: "d", body: { encrypted: { genesisId: "g", checkpoint: { seq: 0, iv: "x", ciphertext: "" } } } }).status).toBe(400);
    expect(
      handleSeedRequest(
        hub,
        { method: "PUT", docId: "d", body: { encrypted: { genesisId: "g", checkpoint: { seq: 0, iv: "x", ciphertext: "A".repeat(200) } } } },
        { maxDocxBytes: 100 },
      ).status,
    ).toBe(413);
    expect(hub.activeDocs()).toEqual([]);
  });

  it("encrypted seed registers a share-code verifier too", async () => {
    const hub = zeroCustodyHub();
    handleSeedRequest(hub, { method: "PUT", docId: "d", body: { encrypted: { genesisId: "g", checkpoint: CP }, codeVerifier: "V" } });
    const c = new FakeConn("c");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "a", sinceSeq: 0, engineVersion: ENGINE_VERSION });
    expect(c.last()).toEqual({ t: "refused", reason: "code-required" });
  });
});

describe("owner token in seed responses (doc 14 §2.5)", () => {
  it("POST and PUT return an owner token ONLY to the seeder", () => {
    const hub = zeroCustodyHub();
    const post = handleSeedRequest(hub, { method: "POST", body: { blank: true } });
    expect(post.status).toBe(201);
    expect(post.body.ownerToken).toMatch(/^o_[0-9a-f]{32}$/);
    const put = handleSeedRequest(hub, { method: "PUT", docId: "d2", body: { docx: b64(docxBytes("x")) } });
    expect(put.body.ownerToken).toMatch(/^o_/);
    // A 409 (someone else won) returns NO owner token — the winner holds it.
    const race = handleSeedRequest(hub, { method: "PUT", docId: "d2", body: { docx: b64(docxBytes("y")) } });
    expect(race.status).toBe(409);
    expect(race.body.ownerToken).toBeUndefined();
  });
  // OPT-IN, not the default. A plaintext room means the server holds and
  // parses the real document; a deployment whose claim is that it never does
  // that sets this. The client UI cannot enforce it — this is the same route,
  // and plain bytes posted to it create a plaintext room.
  it("encryptedOnly refuses a plaintext seed and creates no room, but still takes an encrypted one", () => {
    const hub = new CollabHub(null);

    const plain = handleSeedRequest(
      hub,
      { method: "POST", body: { docx: b64(docxBytes("readable")) } },
      { encryptedOnly: true },
    );
    expect(plain.status).toBe(400);
    expect((plain.body as { error?: string }).error).toBe("encrypted-only");
    // The refusal must be real, not merely a status code: no room may exist.
    expect(hub.roomsSnapshot()).toHaveLength(0);

    // `blank: true` takes a different branch to `docx` and must also refuse —
    // it seeds a plaintext room from server-side template bytes.
    const blank = handleSeedRequest(hub, { method: "POST", body: { blank: true } }, { encryptedOnly: true });
    expect(blank.status).toBe(400);
    expect(hub.roomsSnapshot()).toHaveLength(0);

    // The mode the demo actually uses is unaffected.
    const enc = handleSeedRequest(
      hub,
      { method: "POST", body: { encrypted: { genesisId: "g1", checkpoint: { seq: 0, iv: "aXY=", ciphertext: "Y3Q=" } } } },
      { encryptedOnly: true },
    );
    expect(enc.status).toBe(201);
    expect(hub.roomsSnapshot()).toHaveLength(1);
  });

  it("plaintext seeding is the DEFAULT — the flag is opt-in, not opt-out", () => {
    const hub = new CollabHub(null);
    const res = handleSeedRequest(hub, { method: "POST", body: { docx: b64(docxBytes("hello")) } });
    expect(res.status).toBe(201);
    expect(hub.roomsSnapshot()).toHaveLength(1);
  });
});
