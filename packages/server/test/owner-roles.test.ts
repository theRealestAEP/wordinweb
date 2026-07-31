import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub, Connection } from "../src/hub.js";
import { PROTOCOL_VERSION } from "@wordinweb/collab/server";
import type { ServerMessage, InsertTextIntent } from "@wordinweb/collab/server";

/** Owner / editor / viewer roles as capabilities (plan doc 14 §2.5): the
 * seed-minted owner token unlocks the admin channel; read-only, demote,
 * and kick are integrity controls a blind server enforces on writes. */

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
      const t = this.received[i].t;
      if (t !== "roster") return this.received[i];
    }
    return this.received[this.received.length - 1];
  }
}

const edit = (clientId: string, clientSeq: number): InsertTextIntent =>
  ({ kind: "insertText", clientId, clientSeq, base: 0, at: { blockId: 1, runId: 2, offset: 2 }, text: "!" }) as never;
const hello = (clientId: string, over: Record<string, unknown> = {}) =>
  ({ t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId, sinceSeq: 0, ...over }) as never;

function ownedHub() {
  const hub = new CollabHub(null);
  const r = hub.seed("d", blankDoc());
  if (!r.ok) throw new Error("seed failed");
  return { hub, ownerToken: r.ownerToken };
}

describe("owner roles (doc 14 §2.5)", () => {
  it("seed returns an owner token; a matching hello unlocks admin, a non-owner is refused", async () => {
    const { hub, ownerToken } = ownedHub();
    expect(ownerToken).toMatch(/^o_[0-9a-f]{32}$/);
    const stranger = new FakeConn("s");
    await hub.handle(stranger, hello("alice"));
    await hub.handle(stranger, { t: "admin", action: { op: "readOnly", on: true } });
    expect(stranger.last()).toEqual({ t: "refused", reason: "not-owner" });

    const owner = new FakeConn("o");
    await hub.handle(owner, hello("owner", { ownerToken }));
    await hub.handle(owner, { t: "admin", action: { op: "readOnly", on: true } });
    expect(owner.received.some((m) => m.t === "refused")).toBe(false);
  });

  it("read-only: editors are refused, the owner keeps writing, then it lifts", async () => {
    const { hub, ownerToken } = ownedHub();
    const owner = new FakeConn("o");
    const alice = new FakeConn("a");
    await hub.handle(owner, hello("owner", { ownerToken }));
    await hub.handle(alice, hello("alice"));

    await hub.handle(owner, { t: "admin", action: { op: "readOnly", on: true } });
    await hub.handle(alice, { t: "submit", intent: edit("alice", 1) });
    expect(alice.last()).toEqual({ t: "refused", reason: "read-only" });
    // The owner bypasses their own read-only lock.
    await hub.handle(owner, { t: "submit", intent: edit("owner", 1) });
    expect(owner.last().t).toBe("broadcast");
    // Lifted: alice writes again.
    await hub.handle(owner, { t: "admin", action: { op: "readOnly", on: false } });
    await hub.handle(alice, { t: "submit", intent: edit("alice", 2) });
    expect(alice.last().t).toBe("broadcast");
  });

  it("demote to viewer: that clientId's edits are refused; promote restores", async () => {
    const { hub, ownerToken } = ownedHub();
    const owner = new FakeConn("o");
    const bob = new FakeConn("b");
    await hub.handle(owner, hello("owner", { ownerToken }));
    await hub.handle(bob, hello("bob"));
    await hub.handle(owner, { t: "admin", action: { op: "setRole", clientId: "bob", role: "viewer" } });
    await hub.handle(bob, { t: "submit", intent: edit("bob", 1) });
    expect(bob.last()).toEqual({ t: "refused", reason: "read-only" });
    await hub.handle(owner, { t: "admin", action: { op: "setRole", clientId: "bob", role: "editor" } });
    await hub.handle(bob, { t: "submit", intent: edit("bob", 2) });
    expect(bob.last().t).toBe("broadcast");
  });

  it("kick: the target is disconnected and cannot rejoin under the same clientId", async () => {
    const { hub, ownerToken } = ownedHub();
    const owner = new FakeConn("o");
    const mallory = new FakeConn("m");
    await hub.handle(owner, hello("owner", { ownerToken }));
    await hub.handle(mallory, hello("mallory"));
    await hub.handle(owner, { t: "admin", action: { op: "kick", clientId: "mallory" } });
    expect(mallory.received.some((m) => m.t === "refused" && m.reason === "kicked")).toBe(true);
    // Rejoin under the banned clientId is refused at hello.
    const m2 = new FakeConn("m2");
    await hub.handle(m2, hello("mallory"));
    expect(m2.last()).toEqual({ t: "refused", reason: "kicked" });
  });

  it("ownership follows the seed: a re-seeder gets a token, the old one doesn't work on the new epoch", async () => {
    const { hub, ownerToken } = ownedHub();
    // A different session/epoch (fresh hub simulates a revival elsewhere).
    const hub2 = new CollabHub(null);
    const r2 = hub2.seed("d", blankDoc());
    expect(r2.ok && r2.ownerToken !== ownerToken).toBe(true);
    // The OLD token does not unlock admin on the new epoch.
    if (!r2.ok) return;
    const impostor = new FakeConn("x");
    await hub2.handle(impostor, hello("x", { ownerToken })); // old token
    await hub2.handle(impostor, { t: "admin", action: { op: "readOnly", on: true } });
    expect(impostor.last()).toEqual({ t: "refused", reason: "not-owner" });
  });
});
