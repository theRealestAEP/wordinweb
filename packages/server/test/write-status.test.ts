import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub, type Connection, type DocProvider, type TokenVerifier } from "../src/hub.js";
import { PROTOCOL_VERSION, type ServerMessage, type RosterEntry } from "@wordinweb/collab/server";

/**
 * WRITE STATUS ON THE ROSTER — a POSITIVE signal about whether this
 * participant's edits will be sequenced.
 *
 * Before it, a client could only learn it was blocked by making an edit and
 * being refused, which cost the user their first keystrokes every time
 * (optimistic paint, refusal, heal — a burst of text that appears and
 * vanishes) and left them stuck in viewer mode after a lock was LIFTED,
 * because nothing announced the good news either. The e2e "owner lifts
 * read-only, editor writes again" scenario failed for exactly that reason.
 *
 * THE ASSERTION THAT MATTERS THROUGHOUT: the status arrives WITHOUT the
 * client attempting a write. Every test below inspects the roster after a
 * transition and submits nothing.
 */

function blankDoc(text: string): Uint8Array {
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

const provider: DocProvider = { load: () => blankDoc("hi") };

class FakeConn implements Connection {
  received: ServerMessage[] = [];
  constructor(public id: string) {}
  send(msg: ServerMessage): void {
    this.received.push(msg);
  }
  /** The most recent roster snapshot this connection was sent. */
  roster(): RosterEntry[] {
    for (let i = this.received.length - 1; i >= 0; i--) {
      const m = this.received[i];
      if (m.t === "roster") return m.roster;
    }
    throw new Error("no roster received");
  }
  writeOf(clientId: string): string | undefined {
    return this.roster().find((r) => r.clientId === clientId)?.write;
  }
  rosterCount(): number {
    return this.received.filter((m) => m.t === "roster").length;
  }
  last(): ServerMessage {
    return this.received[this.received.length - 1];
  }
}

async function join(hub: CollabHub, conn: FakeConn, clientId: string, extra: Record<string, unknown> = {}): Promise<void> {
  await hub.handle(conn, {
    t: "hello",
    protocolVersion: PROTOCOL_VERSION,
    docId: "d",
    clientId,
    sinceSeq: 0,
    ...extra,
  } as never);
}

function edit(clientId: string, clientSeq = 1) {
  return {
    t: "submit" as const,
    intent: {
      kind: "insertText" as const,
      clientId,
      clientSeq,
      base: 0,
      at: { blockId: 1, runId: 2, offset: 2 },
      text: "x",
    },
  };
}

/** A seeded room whose owner token is known, so ownership can be proven. */
async function seededRoom() {
  const hub = new CollabHub(provider);
  const seed = hub.seed("d", blankDoc("hi"));
  if (!seed.ok) throw new Error("seed failed");
  return { hub, ownerToken: seed.ownerToken };
}

describe("write status at join", () => {
  it("reports `allowed` in an ordinary room", async () => {
    const hub = new CollabHub(provider);
    const a = new FakeConn("a");
    await join(hub, a, "alice");
    expect(a.writeOf("alice")).toBe("allowed");
  });

  it("a JOINER INTO AN ALREADY-LOCKED ROOM arrives knowing", async () => {
    // The case that has no second chance: without this the newcomer's very
    // first keystrokes are the thing that tells them they cannot type.
    const { hub, ownerToken } = await seededRoom();
    const owner = new FakeConn("o");
    await join(hub, owner, "owner", { ownerToken });
    await hub.handle(owner, { t: "admin", action: { op: "readOnly", on: true } });

    const late = new FakeConn("l");
    await join(hub, late, "latecomer");
    expect(late.writeOf("latecomer")).toBe("owner-lock");
    expect(late.writeOf("owner")).toBe("owner-lock");
  });

  it("reports `viewer-role` when the TOKEN grants read only", async () => {
    // Distinct from the owner's lock because nothing the owner does in-session
    // changes it — the UI can say "you're a viewer in this document" instead
    // of implying the lock might lift.
    const verifier: TokenVerifier = {
      verify: (token) => (token ? { userId: token, role: token === "ed" ? "editor" : "viewer" } : null),
    };
    const hub = new CollabHub(provider, undefined, verifier);
    const viewer = new FakeConn("v");
    await join(hub, viewer, "vic", { token: "vi" });
    expect(viewer.writeOf("vic")).toBe("viewer-role");
  });

  it("reports `demoted` for a per-client demotion", async () => {
    const { hub, ownerToken } = await seededRoom();
    const owner = new FakeConn("o");
    const ed = new FakeConn("e");
    await join(hub, owner, "owner", { ownerToken });
    await join(hub, ed, "editor");
    await hub.handle(owner, { t: "admin", action: { op: "setRole", clientId: "editor", role: "viewer" } });
    expect(ed.writeOf("editor")).toBe("demoted");
  });
});

describe("write status on transitions — the signal arrives unprompted", () => {
  it("LIFTING a lock re-fans the roster, with no write attempted", async () => {
    // THE BUG THIS ARC EXISTS FOR. A lift used to be undiscoverable: the only
    // evidence of the lock was a per-edit refusal, and nothing announced its
    // removal, so a blocked client sat in viewer mode until reload.
    const { hub, ownerToken } = await seededRoom();
    const owner = new FakeConn("o");
    const ed = new FakeConn("e");
    await join(hub, owner, "owner", { ownerToken });
    await join(hub, ed, "editor");
    expect(ed.writeOf("editor")).toBe("allowed");

    await hub.handle(owner, { t: "admin", action: { op: "readOnly", on: true } });
    expect(ed.writeOf("editor")).toBe("owner-lock");
    expect(ed.writeOf("owner")).toBe("owner-lock");

    const before = ed.rosterCount();
    await hub.handle(owner, { t: "admin", action: { op: "readOnly", on: false } });
    // A NEW roster arrived, and the blocked client never submitted anything.
    expect(ed.rosterCount()).toBeGreaterThan(before);
    expect(ed.writeOf("editor")).toBe("allowed");
    expect(ed.received.some((m) => m.t === "refused")).toBe(false);
  });

  it("a demotion reaches the demoted client, and the others' rows are unchanged", async () => {
    const { hub, ownerToken } = await seededRoom();
    const owner = new FakeConn("o");
    const a = new FakeConn("a");
    const b = new FakeConn("b");
    await join(hub, owner, "owner", { ownerToken });
    await join(hub, a, "alice");
    await join(hub, b, "bob");

    await hub.handle(owner, { t: "admin", action: { op: "setRole", clientId: "alice", role: "viewer" } });
    expect(a.writeOf("alice")).toBe("demoted");
    // Bob sees alice's demotion in the shared snapshot but is unaffected —
    // demotion is per-client, not a room state.
    expect(b.writeOf("alice")).toBe("demoted");
    expect(b.writeOf("bob")).toBe("allowed");

    // And promotion back is equally announced.
    await hub.handle(owner, { t: "admin", action: { op: "setRole", clientId: "alice", role: "editor" } });
    expect(a.writeOf("alice")).toBe("allowed");
  });
});

describe("the advertised status matches what the sequencer enforces", () => {
  /**
   * The property that makes the signal trustworthy. A client told "you may
   * write" while the server refuses is worse off than one told nothing, so
   * this drives every condition through BOTH the roster and a real submit and
   * requires them to agree. Enforcement and advertisement share one predicate;
   * this is what proves it rather than assuming it.
   */
  it("agrees for allowed, owner-lock, and demoted", async () => {
    const { hub, ownerToken } = await seededRoom();
    const owner = new FakeConn("o");
    const ed = new FakeConn("e");
    await join(hub, owner, "owner", { ownerToken });
    await join(hub, ed, "editor");

    const accepts = async (c: FakeConn, clientId: string, seq: number) => {
      await hub.handle(c, edit(clientId, seq));
      return c.last().t !== "refused";
    };

    // Unlocked: advertised allowed, and the edit lands.
    expect(ed.writeOf("editor")).toBe("allowed");
    expect(await accepts(ed, "editor", 1)).toBe(true);

    // Locked: advertised owner-lock, and the edit is refused.
    await hub.handle(owner, { t: "admin", action: { op: "readOnly", on: true } });
    expect(ed.writeOf("editor")).toBe("owner-lock");
    expect(await accepts(ed, "editor", 2)).toBe(false);

    // The same room lock blocks the owner, while the admin channel stays live.
    expect(ed.writeOf("owner")).toBe("owner-lock");
    expect(await accepts(owner, "owner", 1)).toBe(false);

    // Demotion, with the room lock lifted, is still a block both ways.
    await hub.handle(owner, { t: "admin", action: { op: "readOnly", on: false } });
    await hub.handle(owner, { t: "admin", action: { op: "setRole", clientId: "editor", role: "viewer" } });
    expect(ed.writeOf("editor")).toBe("demoted");
    expect(await accepts(ed, "editor", 3)).toBe(false);
  });

  it("agrees for a viewer TOKEN, which no owner action can override", async () => {
    const verifier: TokenVerifier = {
      verify: (token) => (token ? { userId: token, role: token === "ed" ? "editor" : "viewer" } : null),
    };
    const hub = new CollabHub(provider, undefined, verifier);
    const viewer = new FakeConn("v");
    await join(hub, viewer, "vic", { token: "vi" });
    expect(viewer.writeOf("vic")).toBe("viewer-role");
    await hub.handle(viewer, edit("vic", 1));
    expect(viewer.last()).toEqual({ t: "refused", reason: "read-only" });
  });
});
