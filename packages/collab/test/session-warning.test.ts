import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { CollabConnection } from "../src/connection.js";
import { EncryptedCollabConnection } from "../src/enc-connection.js";
import { mintDocKey, deriveEpochKeys, sealCheckpoint, bytesToB64 } from "../src/e2ee.js";
import type { ClientMessage, ServerMessage } from "../src/protocol.js";

/**
 * SESSION LIFECYCLE WARNINGS at the client, and the property the whole
 * additive-message design rests on: an UNKNOWN message type must be a no-op.
 *
 * The second half matters more than it looks. "Old clients ignore unknown
 * `t`s" is the reason this arc needed no PROTOCOL_VERSION bump, and it was
 * asserted by reading a switch statement — which is exactly the kind of claim
 * that stays true until someone adds a `default:` that throws, at which point
 * a future additive message bricks every deployed client instead of being
 * ignored by it. Pinned in BOTH connections, because they dispatch separately.
 */

function docxBytes(text: string): Uint8Array {
  const documentXml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`;
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

function fakeTransport(onSend?: (m: ClientMessage) => void) {
  let deliver: (m: ServerMessage) => void = () => {};
  return {
    transport: {
      send: (m: ClientMessage) => onSend?.(m),
      onMessage: (cb: (m: ServerMessage) => void) => {
        deliver = cb;
      },
    },
    deliver: (m: ServerMessage) => deliver(m),
  };
}

async function until(cond: () => boolean, label: string): Promise<void> {
  for (let i = 0; i < 200; i++) {
    if (cond()) return;
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error(`timed out waiting for ${label}`);
}

type Warning = { reason: "idle" | "lifetime"; inMs: number };

describe("plaintext connection — session warnings", () => {
  function connected() {
    const warnings: Warning[] = [];
    const cleared: { reason: "idle" }[] = [];
    const refusals: string[] = [];
    const t = fakeTransport();
    const conn = new CollabConnection(t.transport, "me", {
      onSessionWarning: (info) => warnings.push(info),
      onSessionWarningCleared: (info) => cleared.push(info),
      onRefused: (reason) => refusals.push(reason),
    });
    conn.join("d");
    const session = new DocumentSession(DocxDocument.load(docxBytes("hi")));
    const cp = session.checkpoint();
    t.deliver({
      t: "welcome",
      docId: "d",
      seq: 0,
      snapshot: bytesToB64(cp.docx),
      sidecar: cp.sidecar,
      tail: [],
      genesisId: "g1",
      mode: "plaintext",
    });
    return { conn, t, warnings, cleared, refusals };
  }

  it("surfaces a countdown and its cancellation as distinct events", () => {
    const { t, warnings, cleared } = connected();
    t.deliver({ t: "session-warning", reason: "idle", inMs: 57_000 });
    expect(warnings).toEqual([{ reason: "idle", inMs: 57_000 }]);
    t.deliver({ t: "session-warning-cleared", reason: "idle" });
    expect(cleared).toEqual([{ reason: "idle" }]);
    // A warning is NOT a refusal: the session stays fully usable through the
    // grace period, and conflating the two would tear down a live session.
    expect(warnings).toHaveLength(1);
  });

  it("the kick that follows arrives as an ordinary refusal", () => {
    const { t, refusals } = connected();
    t.deliver({ t: "session-warning", reason: "idle", inMs: 60_000 });
    t.deliver({ t: "refused", reason: "idle-timeout" });
    expect(refusals).toEqual(["idle-timeout"]);
  });

  it("IGNORES an unknown message type — the property additive messages rely on", () => {
    const { conn, t, warnings, cleared, refusals } = connected();
    // A message from a NEWER server than this client. Must be a silent no-op:
    // not a throw, not a refusal, not a state change.
    expect(() => t.deliver({ t: "some-future-message", whatever: 1 } as never)).not.toThrow();
    expect(() => t.deliver({ t: "another-unknown", nested: { deep: true } } as never)).not.toThrow();
    expect(warnings).toEqual([]);
    expect(cleared).toEqual([]);
    expect(refusals).toEqual([]);
    // The connection is still alive and still speaks the protocol it knows.
    t.deliver({ t: "session-warning", reason: "lifetime", inMs: 300_000 });
    expect(warnings).toEqual([{ reason: "lifetime", inMs: 300_000 }]);
    expect(conn.ready).toBe(true);
  });
});

describe("encrypted connection — session warnings", () => {
  async function connected() {
    const docKey = mintDocKey();
    const keys = await deriveEpochKeys(docKey, "g1");
    const session = new DocumentSession(DocxDocument.load(docxBytes("hi")));
    const cp = session.checkpoint();
    const sealed = await sealCheckpoint(keys.kContent, "d", "g1", 0, {
      docx: bytesToB64(cp.docx),
      sidecar: cp.sidecar,
      docHash: "seed",
    });
    const warnings: Warning[] = [];
    const cleared: { reason: "idle" }[] = [];
    const errors: string[] = [];
    const t = fakeTransport();
    const conn = new EncryptedCollabConnection(t.transport, "me", docKey, {
      onSessionWarning: (info) => warnings.push(info),
      onSessionWarningCleared: (info) => cleared.push(info),
      onError: (info) => errors.push(info.where),
    });
    conn.join("d");
    t.deliver({
      t: "welcome-enc",
      docId: "d",
      genesisId: "g1",
      checkpoint: { seq: 0, ...sealed },
      tail: [],
      mode: "encrypted",
    });
    await until(() => conn.ready, "the encrypted connection to rehydrate");
    return { conn, t, warnings, cleared, errors };
  }

  it("delivers warnings in encrypted rooms too — the deadline is server policy, not content", async () => {
    const { t, warnings, cleared } = await connected();
    t.deliver({ t: "session-warning", reason: "lifetime", inMs: 240_000 });
    t.deliver({ t: "session-warning", reason: "idle", inMs: 45_000 });
    t.deliver({ t: "session-warning-cleared", reason: "idle" });
    expect(warnings).toEqual([
      { reason: "lifetime", inMs: 240_000 },
      { reason: "idle", inMs: 45_000 },
    ]);
    expect(cleared).toEqual([{ reason: "idle" }]);
  });

  it("IGNORES an unknown message type without wedging the serial chain", async () => {
    const { conn, t, warnings, errors } = await connected();
    expect(() => t.deliver({ t: "some-future-message", whatever: 1 } as never)).not.toThrow();
    // The encrypted connection runs its work on a serial promise chain, so
    // "ignored" has to mean the chain is untouched — an unknown message that
    // rejected it would take every later submit and ingest down with it.
    await new Promise((r) => setTimeout(r, 20));
    expect(errors).toEqual([]);
    expect(conn.ready).toBe(true);
    t.deliver({ t: "session-warning", reason: "idle", inMs: 1_000 });
    expect(warnings).toEqual([{ reason: "idle", inMs: 1_000 }]);
  });
});
