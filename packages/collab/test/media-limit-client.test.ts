import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { CollabConnection } from "../src/connection.js";
import { EncryptedCollabConnection } from "../src/enc-connection.js";
import { mintDocKey, deriveEpochKeys, sealCheckpoint, bytesToB64 } from "../src/e2ee.js";
import type { ClientMessage, ServerMessage } from "../src/protocol.js";

/**
 * THE CLIENT SIDE of the published media limit (#25): both connections expose
 * the server's configured per-blob cap, so a file can be checked before it is
 * uploaded rather than after a 413.
 *
 * The contract that needs pinning is the ABSENT case. A server that publishes
 * nothing must leave the field `null`, and callers must read that as "skip the
 * pre-check" — not as "no limit", and not as a default they invent. A client
 * that guessed a larger number would tell the user an upload is fine and then
 * watch the server refuse it, which is worse than never having checked.
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

function fakeTransport() {
  let deliver: (m: ServerMessage) => void = () => {};
  return {
    transport: {
      send: (_m: ClientMessage) => {},
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

const ODD_LIMIT = 1234567;

function plaintextWelcome(extra: Record<string, unknown>): ServerMessage {
  const session = new DocumentSession(DocxDocument.load(docxBytes("hi")));
  const cp = session.checkpoint();
  return {
    t: "welcome",
    docId: "d",
    seq: 0,
    snapshot: bytesToB64(cp.docx),
    sidecar: cp.sidecar,
    tail: [],
    genesisId: "g1",
    mode: "plaintext",
    ...extra,
  } as ServerMessage;
}

describe("plaintext connection exposes the published media limit", () => {
  it("reads the welcome's value", () => {
    const t = fakeTransport();
    const conn = new CollabConnection(t.transport, "me");
    conn.join("d");
    expect(conn.mediaMaxBlobBytes).toBeNull(); // nothing known before the welcome
    t.deliver(plaintextWelcome({ mediaMaxBlobBytes: ODD_LIMIT }));
    expect(conn.mediaMaxBlobBytes).toBe(ODD_LIMIT);
  });

  it("stays NULL when the server publishes nothing", () => {
    // An older host. `null` is the signal to skip the pre-check entirely —
    // inventing a limit here would either block uploads the server would have
    // accepted or promise ones it will refuse.
    const t = fakeTransport();
    const conn = new CollabConnection(t.transport, "me");
    conn.join("d");
    t.deliver(plaintextWelcome({}));
    expect(conn.mediaMaxBlobBytes).toBeNull();
  });
});

describe("encrypted connection exposes the published media limit", () => {
  async function connected(extra: Record<string, unknown>) {
    const docKey = mintDocKey();
    const keys = await deriveEpochKeys(docKey, "g1");
    const session = new DocumentSession(DocxDocument.load(docxBytes("hi")));
    const cp = session.checkpoint();
    const sealed = await sealCheckpoint(keys.kContent, "d", "g1", 0, {
      docx: bytesToB64(cp.docx),
      sidecar: cp.sidecar,
      docHash: "seed",
    });
    const t = fakeTransport();
    const conn = new EncryptedCollabConnection(t.transport, "me", docKey, {});
    conn.join("d");
    t.deliver({
      t: "welcome-enc",
      docId: "d",
      genesisId: "g1",
      checkpoint: { seq: 0, ...sealed },
      tail: [],
      mode: "encrypted",
      ...extra,
    } as ServerMessage);
    return { conn };
  }

  it("is readable IMMEDIATELY, without waiting on the rehydrate chain", async () => {
    // Deliberate: the value is a plain number with no dependency on keys or
    // the mirror, and a caller asking "how big may this file be" should not
    // have to wait for a decrypt to find out.
    const { conn } = await connected({ mediaMaxBlobBytes: ODD_LIMIT });
    expect(conn.mediaMaxBlobBytes).toBe(ODD_LIMIT);
    // …and it survives the rehydrate that follows.
    await until(() => conn.ready, "the encrypted connection to rehydrate");
    expect(conn.mediaMaxBlobBytes).toBe(ODD_LIMIT);
  });

  it("stays NULL when the server publishes nothing", async () => {
    const { conn } = await connected({});
    await until(() => conn.ready, "the encrypted connection to rehydrate");
    expect(conn.mediaMaxBlobBytes).toBeNull();
  });
});
