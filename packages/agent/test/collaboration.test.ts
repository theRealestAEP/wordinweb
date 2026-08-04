import { describe, expect, it } from "vitest";
import { DocxDocument } from "@wordinweb/core";
import { CollabConnection, applyIntentScoped, resyncScope, type ClientTransport, type Intent, type IntentBody } from "@wordinweb/collab/client";
import { DocumentSession, PROTOCOL_VERSION, type ClientMessage, type ServerMessage } from "@wordinweb/collab/server";
import { AgentDocument, type AgentCollaborativeTarget } from "../src/index.js";
import { body, makeDocx } from "./helpers.js";

function base64(bytes: Uint8Array): string {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

class LoopbackHub {
  readonly session: DocumentSession;
  private peers: Array<{ deliver(message: ServerMessage): void; joined: boolean }> = [];

  constructor(bytes: Uint8Array) {
    this.session = new DocumentSession(DocxDocument.load(bytes));
  }

  connect(): ClientTransport {
    const peer = { deliver: (_message: ServerMessage) => {}, joined: false };
    this.peers.push(peer);
    return {
      onMessage: (handler) => { peer.deliver = handler; },
      send: (message) => this.handle(peer, message),
    };
  }

  private handle(peer: (typeof this.peers)[number], message: ClientMessage): void {
    if (message.t === "hello") {
      if (message.protocolVersion !== PROTOCOL_VERSION) return peer.deliver({ t: "refused", reason: "version-mismatch" });
      peer.joined = true;
      const checkpoint = this.session.checkpoint();
      peer.deliver({ t: "welcome", docId: message.docId, seq: checkpoint.seq, snapshot: base64(checkpoint.docx), sidecar: checkpoint.sidecar, tail: [], genesisId: "agent-test", mode: "plaintext" });
    } else if (message.t === "submit" && peer.joined) {
      const entry = this.session.submit(message.intent);
      for (const recipient of this.peers) if (recipient.joined) recipient.deliver({ t: "broadcast", entries: [entry] });
    }
  }
}

function firstPosition(agent: AgentDocument): { blockRef: string; runRef: string } {
  const read = agent.inspect({ kind: "read" });
  if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
  return { blockRef: read.blocks[0].ref, runRef: read.blocks[0].runs[0].ref };
}

function paragraphPositions(agent: AgentDocument): Array<{ blockRef: string; runRef: string; offset: number }> {
  const read = agent.inspect({ kind: "read", maxBlocks: 10 });
  if (!("blocks" in read)) throw new Error("missing paragraphs");
  return read.blocks.map((block) => {
    if (block.type !== "paragraph" || !block.runs[0]) throw new Error("missing paragraph run");
    return { blockRef: block.ref, runRef: block.runs[0].ref, offset: block.text.length };
  });
}

function applyingTarget(bytes: Uint8Array): { target: AgentCollaborativeTarget; doc: DocxDocument } {
  const doc = DocxDocument.load(bytes);
  const ids = doc.enableStableIds();
  let revision = 0;
  let clientSeq = 0;
  let nextId = 100_000;
  return {
    doc,
    target: {
      getDocument: () => doc,
      getRevision: () => revision,
      allocateIds: (count) => Array.from({ length: count }, () => nextId++),
      submit: (operation) => {
        const result = applyIntentScoped(doc, ids, { ...operation, clientId: "shared-agent", clientSeq: ++clientSeq, base: revision } as Intent);
        if (!result.applied) throw new Error("shared apply failed");
        resyncScope(doc, ids, result);
        revision++;
      },
      getConnectionState: () => "live",
    },
  };
}

describe("AgentDocument collaboration adapters", () => {
  it("submits through a real collaborative connection and reaches a second replica", async () => {
    const hub = new LoopbackHub(makeDocx(body(`<w:p><w:r><w:t>Shared</w:t></w:r></w:p>`)));
    let aVersion = 0;
    let bVersion = 0;
    const a = new CollabConnection(hub.connect(), "agent", { onChange: () => { aVersion++; } });
    const b = new CollabConnection(hub.connect(), "reviewer", { onChange: () => { bVersion++; } });
    a.join("document");
    b.join("document");
    const target = (connection: CollabConnection, version: () => number): AgentCollaborativeTarget => ({
      getDocument: () => connection.doc,
      getRevision: version,
      allocateIds: (count) => connection.allocIds(count),
      submit: (operation) => connection.submit(operation),
      getConnectionState: () => "live",
    });
    const agent = AgentDocument.connect(target(a, () => aVersion));
    const reviewer = AgentDocument.connect(target(b, () => bVersion));
    const position = firstPosition(agent);
    const result = await agent.edit({
      revision: agent.revision,
      operations: [{ kind: "insertText", at: { ...position, offset: 6 }, text: " by AI" }],
    });
    expect(result.status).toBe("submitted");
    const search = reviewer.inspect({ kind: "search", query: "by AI" });
    expect("matches" in search && search.matches).toHaveLength(1);
    expect(a.doc?.save()).toEqual(b.doc?.save());
  });

  it("uses the same operation surface for an offline tail", async () => {
    const doc = DocxDocument.load(makeDocx(body(`<w:p><w:r><w:t>Offline</w:t></w:r></w:p>`)));
    const ids = doc.enableStableIds();
    let revision = 0;
    let nextId = 100_000;
    const tail: IntentBody[] = [];
    const target: AgentCollaborativeTarget = {
      getDocument: () => doc,
      getRevision: () => revision,
      allocateIds: (count) => Array.from({ length: count }, () => nextId++),
      submit: (operation) => {
        const result = applyIntentScoped(doc, ids, { ...operation, clientId: "offline-agent", clientSeq: tail.length + 1, base: 0 } as Intent);
        if (!result.applied) throw new Error("offline apply failed");
        resyncScope(doc, ids, result);
        tail.push(operation);
        revision++;
      },
      getConnectionState: () => "offline",
    };
    const agent = AgentDocument.connect(target);
    const position = firstPosition(agent);
    const result = await agent.edit({ revision: "0", operations: [{ kind: "insertText", at: { ...position, offset: 7 }, text: " edit" }] });
    expect(result).toMatchObject({ status: "submitted", connection: "offline", revision: "1" });
    expect(tail).toHaveLength(1);
    expect(tail[0]).toMatchObject({ kind: "insertText", text: " edit" });
    const search = agent.inspect({ kind: "search", query: "Offline edit" });
    expect("matches" in search && search.matches).toHaveLength(1);
  });

  it("allows agents to edit different inspected targets across revisions", async () => {
    const { target } = applyingTarget(makeDocx(body(
      `<w:p><w:r><w:t>Alpha</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>Beta</w:t></w:r></w:p>`,
    )));
    const firstAgent = AgentDocument.connect(target);
    const secondAgent = AgentDocument.connect(target);
    const firstTargets = paragraphPositions(firstAgent);
    const secondTargets = paragraphPositions(secondAgent);

    await firstAgent.edit({
      revision: "0",
      operations: [{ kind: "insertText", at: firstTargets[0], text: " one" }],
    });
    const result = await secondAgent.edit({
      revision: "0",
      operations: [{ kind: "insertText", at: secondTargets[1], text: " two" }],
    });

    expect(result).toMatchObject({ status: "submitted", revision: "2" });
    const context = secondAgent.inspect({ kind: "context" });
    if (!("contents" in context)) throw new Error("missing context");
    const text = context.contents.flatMap((story) => story.blocks).flatMap((block) => block.type === "paragraph" ? [block.text] : []);
    expect(text).toEqual(["Alpha one", "Beta two"]);
  });

  it("rejects a stale edit when another agent changed the inspected target", async () => {
    const { target } = applyingTarget(makeDocx(body(`<w:p><w:r><w:t>Shared</w:t></w:r></w:p>`)));
    const firstAgent = AgentDocument.connect(target);
    const secondAgent = AgentDocument.connect(target);
    const firstTarget = paragraphPositions(firstAgent)[0];
    const secondTarget = paragraphPositions(secondAgent)[0];

    await firstAgent.edit({
      revision: "0",
      operations: [{ kind: "insertText", at: firstTarget, text: " first" }],
    });
    await expect(secondAgent.edit({
      revision: "0",
      operations: [{ kind: "insertText", at: secondTarget, text: " second" }],
    })).rejects.toThrow("revision is stale");
  });
});
