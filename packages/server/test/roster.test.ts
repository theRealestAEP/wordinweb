import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub, Connection, DocProvider } from "../src/hub.js";
import { InMemoryStorage } from "../src/storage.js";
import { PROTOCOL_VERSION, type ServerMessage } from "@wordinweb/collab/server";

/** Roster behavior (plan doc 14 §2): sanitized fan-out keyed by the bound
 * clientId, greyed-not-removed on disconnect, resumed on reconnect, never
 * persisted. */

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
const provider: DocProvider = { load: () => blankDoc() };

class FakeConn implements Connection {
  received: ServerMessage[] = [];
  constructor(public id: string) {}
  send(msg: ServerMessage): void {
    this.received.push(msg);
  }
}

const hello = (docId: string, clientId: string, profile?: { name: string; color: string }) =>
  ({ t: "hello", protocolVersion: PROTOCOL_VERSION, docId, clientId, sinceSeq: 0, profile }) as const;
const rosterOf = (c: FakeConn) => {
  const m = [...c.received].reverse().find((m) => m.t === "roster");
  return m && m.t === "roster" ? m.roster : [];
};

describe("CollabHub roster (doc 14 §2)", () => {
  it("join fans out a sanitized roster to everyone", async () => {
    const hub = new CollabHub(provider);
    const a = new FakeConn("s1");
    const b = new FakeConn("s2");
    // "" is a BEL control character — must be stripped server-side.
    await hub.handle(a, hello("d", "alice", { name: "  Alice ", color: "#FF0000" }));
    await hub.handle(b, hello("d", "bob", { name: "x".repeat(60), color: "javascript:alert(1)" }));
    const roster = rosterOf(a);
    expect(roster).toHaveLength(2);
    const alice = roster.find((r) => r.clientId === "alice")!;
    const bob = roster.find((r) => r.clientId === "bob")!;
    expect(alice.profile.name).toBe("Alice"); // trimmed + control char stripped
    expect(alice.profile.color).toBe("#ff0000"); // valid hex, normalized case
    expect(bob.profile.name).toBe("x".repeat(40)); // length-capped at 40
    expect(bob.profile.color).toMatch(/^#[0-9a-f]{6}$/); // hostile color replaced from palette
  });

  it("no profile: a generated default name, never an empty roster entry", async () => {
    const hub = new CollabHub(provider);
    const a = new FakeConn("s1");
    await hub.handle(a, hello("d", "anon"));
    expect(rosterOf(a)[0].profile.name).toMatch(/^Guest \d{3}$/);
  });

  it("rename fans out; disconnect greys the entry; reconnect under the same clientId resumes it", async () => {
    const hub = new CollabHub(provider);
    const a = new FakeConn("s1");
    const b = new FakeConn("s2");
    await hub.handle(a, hello("d", "alice", { name: "Alice", color: "#ff0000" }));
    await hub.handle(b, hello("d", "bob", { name: "Bob", color: "#00ff00" }));
    await hub.handle(a, { t: "profile", profile: { name: "Alicia", color: "#ff0000" } });
    expect(rosterOf(b).find((r) => r.clientId === "alice")!.profile.name).toBe("Alicia");
    await hub.handle(a, {
      t: "presence",
      position: { anchor: { blockId: 1, runId: 2, offset: 0 } },
    });

    hub.disconnect(a);
    const afterLeave = rosterOf(b).find((r) => r.clientId === "alice")!;
    expect(afterLeave.connected).toBe(false); // greyed, NOT removed — attribution keeps a name
    expect(afterLeave.profile.name).toBe("Alicia");
    expect(
      [...b.received].reverse().find((m) => m.t === "presence"),
      "disconnect must clear the participant's last caret",
    ).toEqual({ t: "presence", participant: "alice", position: null });

    const a2 = new FakeConn("s3"); // new socket, same bound identity
    await hub.handle(a2, hello("d", "alice", { name: "Alicia", color: "#ff0000" }));
    const resumed = rosterOf(b).find((r) => r.clientId === "alice")!;
    expect(resumed.connected).toBe(true);
    expect(resumed.profile.name).toBe("Alicia");
  });

  it("takeover does not grey the entry when the old socket tears down", async () => {
    const hub = new CollabHub(provider);
    const zombie = new FakeConn("s1");
    const fresh = new FakeConn("s2");
    const watcher = new FakeConn("s3");
    await hub.handle(watcher, hello("d", "bob"));
    await hub.handle(zombie, hello("d", "alice", { name: "Alice", color: "#ff0000" }));
    await hub.handle(fresh, {
      ...hello("d", "alice", { name: "Alice", color: "#ff0000" }),
      takeover: true,
    });
    // The transport will call disconnect() for the kicked zombie — the NEW
    // socket still holds the identity, so alice must stay connected.
    hub.disconnect(zombie);
    expect(rosterOf(watcher).find((r) => r.clientId === "alice")!.connected).toBe(true);
  });

  it("roster is ephemeral: nothing reaches the storage driver (zero custody extends to identity)", async () => {
    const storage = new InMemoryStorage();
    const hub = new CollabHub(provider, storage);
    const a = new FakeConn("s1");
    await hub.handle(a, hello("d", "alice", { name: "Alice", color: "#ff0000" }));
    await hub.handle(a, { t: "profile", profile: { name: "A2", color: "#ff0000" } });
    expect(await storage.readLog("d", 0)).toHaveLength(0);
  });

  it("binds an invited AI to its inviter and keeps chat private", async () => {
    const hub = new CollabHub(provider);
    const inviter = new FakeConn("s1");
    const watcher = new FakeConn("s2");
    const agent = new FakeConn("s3");
    await hub.handle(inviter, hello("d", "alice", { name: "Alice", color: "#ff0000" }));
    await hub.handle(watcher, hello("d", "bob", { name: "Bob", color: "#00ff00" }));
    await hub.handle(inviter, {
      t: "agent-invite",
      inviteId: "invite_1234567890",
      token: "token_12345678901234567890123456789012",
      expiresAt: Date.now() + 60_000,
    });
    await hub.handle(agent, {
      ...hello("d", "agent-one", { name: "Drafting agent", color: "#0000ff" }),
      agentInvite: {
        inviteId: "invite_1234567890",
        token: "token_12345678901234567890123456789012",
      },
    });

    expect(rosterOf(watcher).find((entry) => entry.clientId === "agent-one")?.participantType).toBe("agent");
    expect(inviter.received.some((message) => message.t === "agent-connected")).toBe(true);
    expect(watcher.received.some((message) => message.t === "agent-connected")).toBe(false);
    const secondAgent = new FakeConn("s4");
    await hub.handle(secondAgent, {
      ...hello("d", "agent-two", { name: "Other agent", color: "#0000ff" }),
      agentInvite: {
        inviteId: "invite_1234567890",
        token: "token_12345678901234567890123456789012",
      },
    });
    expect(secondAgent.received).toContainEqual({ t: "refused", reason: "agent-invite-invalid" });

    inviter.received.length = 0;
    watcher.received.length = 0;
    agent.received.length = 0;
    await hub.handle(watcher, {
      t: "agent-chat",
      agentClientId: "agent-one",
      messageId: "message_1234",
      iv: "iv_12345678",
      ciphertext: "ciphertext_1234",
    });
    expect(inviter.received).toHaveLength(0);
    expect(agent.received).toHaveLength(0);

    await hub.handle(inviter, {
      t: "agent-chat",
      agentClientId: "agent-one",
      messageId: "message_5678",
      iv: "iv_12345678",
      ciphertext: "ciphertext_5678",
    });
    expect(agent.received).toContainEqual(expect.objectContaining({
      t: "agent-chat",
      sender: "inviter",
      agentClientId: "agent-one",
    }));
    expect(inviter.received).toContainEqual(expect.objectContaining({ t: "agent-chat", sender: "inviter" }));
    expect(watcher.received).toHaveLength(0);

    hub.disconnect(agent);
    const retriedAgent = new FakeConn("s5");
    await hub.handle(retriedAgent, {
      ...hello("d", "agent-retry", { name: "Drafting agent", color: "#0000ff" }),
      agentInvite: {
        inviteId: "invite_1234567890",
        token: "token_12345678901234567890123456789012",
      },
    });
    expect(retriedAgent.received.some((message) => message.t === "welcome")).toBe(true);
    expect(rosterOf(inviter).find((entry) => entry.clientId === "agent-retry")?.participantType).toBe("agent");
  });

  it("refuses an AI retry after its invitation expires", async () => {
    let now = 1_000;
    const hub = new CollabHub(provider, undefined, undefined, () => now);
    const inviter = new FakeConn("s1");
    const agent = new FakeConn("s2");
    await hub.handle(inviter, hello("d", "alice"));
    await hub.handle(inviter, {
      t: "agent-invite",
      inviteId: "invite_1234567890",
      token: "token_12345678901234567890123456789012",
      expiresAt: now + 100,
    });
    await hub.handle(agent, {
      ...hello("d", "agent-one"),
      agentInvite: {
        inviteId: "invite_1234567890",
        token: "token_12345678901234567890123456789012",
      },
    });
    hub.disconnect(agent);
    now += 101;

    const retry = new FakeConn("s3");
    await hub.handle(retry, {
      ...hello("d", "agent-retry"),
      agentInvite: {
        inviteId: "invite_1234567890",
        token: "token_12345678901234567890123456789012",
      },
    });
    expect(retry.received).toContainEqual({ t: "refused", reason: "agent-invite-invalid" });
  });
});
