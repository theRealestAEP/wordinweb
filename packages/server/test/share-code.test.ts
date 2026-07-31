import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub, Connection, EVICTION_GRACE_MS } from "../src/hub.js";
import { PROTOCOL_VERSION } from "@wordinweb/collab/server";
import type { ServerMessage } from "@wordinweb/collab/server";

/** Share-code gate (plan doc 13 §7): verifier at seed, attempt budget,
 * lockout, rotation-by-re-seed. The verifier is any opaque string here —
 * the crypto half (PBKDF2 stretch + derivation mix-in) is pinned in
 * collab's e2ee tests; this file pins the server's online-guessing wall. */

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

class FakeConn implements Connection {
  received: ServerMessage[] = [];
  constructor(public id: string) {}
  send(msg: ServerMessage): void {
    this.received.push(msg);
  }
  last(): ServerMessage {
    for (let i = this.received.length - 1; i >= 0; i--) {
      if (this.received[i].t !== "roster") return this.received[i];
    }
    return this.received[this.received.length - 1];
  }
}

const hello = (clientId: string, codeProof?: string) =>
  ({ t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId, sinceSeq: 0, codeProof }) as never;

describe("share-code gate (doc 13 §7)", () => {
  it("no proof → code-required; wrong → code-invalid; right → welcome", async () => {
    let now = 0;
    const hub = new CollabHub(null, undefined, undefined, () => now);
    hub.seed("d", blankDoc(), undefined, "VERIFIER");
    const c = new FakeConn("c1");
    await hub.handle(c, hello("alice"));
    expect(c.last()).toEqual({ t: "refused", reason: "code-required" });
    await hub.handle(c, hello("alice", "WRONG"));
    expect(c.last()).toEqual({ t: "refused", reason: "code-invalid" });
    await hub.handle(c, hello("alice", "VERIFIER"));
    expect(c.last().t).toBe("welcome");
  });

  it("5 consecutive failures lock the doc for a window; a correct code works after it expires", async () => {
    let now = 0;
    const hub = new CollabHub(null, undefined, undefined, () => now);
    hub.seed("d", blankDoc(), undefined, "VERIFIER");
    const attacker = new FakeConn("atk");
    for (let i = 0; i < 5; i++) await hub.handle(attacker, hello("mallory", `guess-${i}`));
    // Locked: even the CORRECT code is refused during the window (that is
    // the wall — an online brute force gets 5 tries per minute, not 10^6).
    const honest = new FakeConn("h");
    await hub.handle(honest, hello("alice", "VERIFIER"));
    expect(honest.last()).toEqual({ t: "refused", reason: "code-locked" });
    now += 60_001; // window expires
    await hub.handle(honest, hello("alice", "VERIFIER"));
    expect(honest.last().t).toBe("welcome");
  });

  it("a success resets the failure budget", async () => {
    const hub = new CollabHub(null);
    hub.seed("d", blankDoc(), undefined, "VERIFIER");
    const c = new FakeConn("c1");
    for (let i = 0; i < 4; i++) await hub.handle(c, hello("alice", "bad"));
    await hub.handle(c, hello("alice", "VERIFIER")); // success at 4 failures
    expect(c.last().t).toBe("welcome");
    // Budget reset: four MORE bad guesses from someone else don't lock yet.
    const d = new FakeConn("c2");
    for (let i = 0; i < 4; i++) await hub.handle(d, hello("bob", "bad"));
    await hub.handle(d, hello("bob", "VERIFIER"));
    expect(d.last().t).toBe("welcome");
  });

  it("rotation: a re-seed after eviction registers a NEW verifier (or none)", async () => {
    let now = 0;
    const hub = new CollabHub(null, undefined, undefined, () => now);
    hub.seed("d", blankDoc(), undefined, "OLD-CODE");
    now += EVICTION_GRACE_MS;
    hub.sweepRooms();
    hub.seed("d", blankDoc(), undefined, "NEW-CODE");
    const c = new FakeConn("c1");
    await hub.handle(c, hello("alice", "OLD-CODE"));
    expect(c.last()).toEqual({ t: "refused", reason: "code-invalid" });
    await hub.handle(c, hello("alice", "NEW-CODE"));
    expect(c.last().t).toBe("welcome");
  });

  it("codeless docs are unaffected (the gate only exists when registered)", async () => {
    const hub = new CollabHub(null);
    hub.seed("d", blankDoc());
    const c = new FakeConn("c1");
    await hub.handle(c, hello("alice"));
    expect(c.last().t).toBe("welcome");
  });
});
