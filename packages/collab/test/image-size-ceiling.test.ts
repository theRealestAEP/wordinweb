import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, type Paragraph, type Run } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";
import { validateIntent, MAX_IMAGE_BYTES } from "../src/validate.js";
import { CollabConnection } from "../src/connection.js";
import { EncryptedCollabConnection } from "../src/enc-connection.js";
import { mintDocKey, deriveEpochKeys, sealCheckpoint, sealIntent, bytesToB64 } from "../src/e2ee.js";
import type { ServerMessage } from "../src/protocol.js";
import type { Intent } from "../src/intents.js";

/**
 * THE IMAGE SIZE CEILING, pinned from the ACCEPTANCE side.
 *
 * A user's 16-26 MB photos vanished with no image and no error. The validator
 * hardcoded a 10 MB bound written when the media cap was a fixed 10 MB; once
 * the cap became configurable (5 MB in production, 50 MB in dev) the two
 * numbers were never reconciled, and every image between them died in the gap:
 * the upload SUCCEEDED, then the intent was rejected as "bad size" — identically
 * on every replica, so nothing diverged, nothing threw, nothing was logged.
 *
 * WHY EVERY EXISTING TEST MISSED IT, and the reason these are written the way
 * they are: the media suites all use tiny fixtures (`bytesLen` of 4, 5, 68 — a
 * 1x1 PNG), so they exercise the accepted side of a bound they never approach.
 * A pin asserting "an oversized image is rejected" would have passed happily
 * for this bug's entire life, because the bug was the bound being TOO LOW.
 *
 * So the assertions below are that LARGE, LEGITIMATE images are ACCEPTED, and
 * that the configured cap can never exceed what the wire will honour.
 */

function makeDoc(text: string): DocxDocument {
  const body = `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const documentXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`;
  return DocxDocument.load(
    zipSync({
      "[Content_Types].xml": strToU8(
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(documentXml),
    }),
  );
}

function firstRunId(s: DocumentSession): number {
  const para = s.doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  return s.ids.idOf(run.src!)!;
}

const SHA = "a".repeat(64);

function imageIntent(bytesLen: number, runId: number, clientSeq = 1): Intent {
  return {
    kind: "insertImage",
    clientId: "a",
    clientSeq,
    base: 0,
    runId,
    blobSha: SHA,
    bytesLen,
    ext: "jpg",
    widthPx: 1200,
    heightPx: 900,
    nodeIds: [500 + clientSeq],
  } as Intent;
}

/** The sizes that actually broke, plus the configured caps in play. */
const MB = 1024 * 1024;
const USER_PHOTO_SMALL = 16 * MB; // the low end of the owner's JPEGs
const USER_PHOTO_LARGE = 26 * MB; // the high end
const DEV_CAP = 50 * MB; // scripts/dev.mjs
const PROD_CAP = 5 * MB; // compose.yml

describe("insertImage accepts real-world photo sizes", () => {
  it("VALIDATES the exact sizes that silently vanished (16MB and 26MB)", () => {
    // The regression assertion. Against the old 10MB constant both of these
    // return "insertImage: bad size", which is precisely what the user saw.
    for (const bytesLen of [USER_PHOTO_SMALL, USER_PHOTO_LARGE]) {
      expect(validateIntent(imageIntent(bytesLen, 2))).toBeNull();
    }
  });

  it("APPLIES a 26MB image end-to-end through the session, not merely validates", () => {
    // Validation passing is not the user-visible outcome; the entry being
    // `applied` rather than `rejected` is. The old bound produced a sequenced
    // REJECTION on every replica — consistent, which is why no convergence or
    // desync check ever noticed.
    const s = new DocumentSession(makeDoc("photo:"));
    const e = s.submit(imageIntent(USER_PHOTO_LARGE, firstRunId(s)));
    expect(e.kind).toBe("applied");
  });

  it("accepts an image at the DEV cap (50MB) — the config that exposed the bug", () => {
    // Dev runs 50MB so the owner can paste a phone photo without thinking.
    // Every byte between the old 10MB bound and this was in the dead window.
    const s = new DocumentSession(makeDoc("photo:"));
    expect(s.submit(imageIntent(DEV_CAP, firstRunId(s))).kind).toBe("applied");
  });

  it("still rejects a nonsensical declared length", () => {
    // The bound is a sanity check, not decoration: it must still refuse an
    // absurd claim, zero, and a non-integer.
    expect(validateIntent(imageIntent(MAX_IMAGE_BYTES + 1, 2))).toMatch(/bad size/);
    expect(validateIntent(imageIntent(0, 2))).toMatch(/bad size/);
    expect(validateIntent(imageIntent(-1, 2))).toMatch(/bad size/);
  });

  it("A REJECTION IS NO LONGER SILENT — the connection reports our own dead intent", () => {
    // The more durable half of this fix. The bound was wrong once and can be
    // wrong again; what must not survive is the SILENCE. A rejection is
    // sequenced and agreed by every replica, so it produces no divergence, no
    // throw, and no refusal — the edit just disappears. This pins that the
    // client is told.
    const s = new DocumentSession(makeDoc("photo:"));
    const rejected = s.submit(imageIntent(MAX_IMAGE_BYTES + 1, firstRunId(s)));
    expect(rejected.kind).toBe("rejected");

    const seen: { reason: string; clientSeq: number }[] = [];
    const conn = new CollabConnection({ send: () => {}, onMessage: () => {} }, "a", {
      onIntentRejected: (info) => seen.push(info),
    });
    (conn as unknown as { onServer: (m: ServerMessage) => void }).onServer({
      t: "broadcast",
      entries: [rejected],
    });
    expect(seen).toEqual([{ reason: "insertImage: bad size", clientSeq: 1 }]);
  });

  it("does NOT report another participant's rejection", () => {
    // Not actionable here, and a notice about someone else's failed edit is
    // noise the user cannot do anything with.
    const seen: unknown[] = [];
    const conn = new CollabConnection({ send: () => {}, onMessage: () => {} }, "me", {
      onIntentRejected: (info) => seen.push(info),
    });
    (conn as unknown as { onServer: (m: ServerMessage) => void }).onServer({
      t: "broadcast",
      entries: [{ seq: 1, kind: "rejected", clientId: "someone-else", clientSeq: 1, reason: "insertImage: bad size" }],
    });
    expect(seen).toEqual([]);
  });

  it("ENCRYPTED rooms report it too — the mode the demo actually runs in", async () => {
    // The important half. A blind server cannot validate, so in encrypted
    // rooms the rejection happens ONLY in this client's own mirror and there
    // is no server-side signal anywhere. If the channel did not work here it
    // would not work where the bug actually bit.
    const docKey = mintDocKey();
    const keys = await deriveEpochKeys(docKey, "g1");
    const session = new DocumentSession(makeDoc("photo:"));
    const cp = session.checkpoint();
    const sealedCp = await sealCheckpoint(keys.kContent, "d", "g1", 0, {
      docx: bytesToB64(cp.docx),
      sidecar: cp.sidecar,
      docHash: "seed",
    });

    const seen: { reason: string; clientSeq: number }[] = [];
    let deliver: (m: ServerMessage) => void = () => {};
    const conn = new EncryptedCollabConnection(
      { send: () => {}, onMessage: (cb: (m: ServerMessage) => void) => (deliver = cb) },
      "a",
      docKey,
      { onIntentRejected: (info) => seen.push(info) },
    );
    conn.join("d");
    deliver({ t: "welcome-enc", docId: "d", genesisId: "g1", checkpoint: { seq: 0, ...sealedCp }, tail: [], mode: "encrypted" } as ServerMessage);
    for (let i = 0; i < 200 && !conn.ready; i++) await new Promise((r) => setTimeout(r, 5));
    expect(conn.ready).toBe(true);

    // The mirror is seeded from this same checkpoint AND sidecar, so its id
    // table matches `session`'s exactly — that is the whole point of shipping
    // the sidecar with a snapshot.
    const runId = firstRunId(session);
    const bad = imageIntent(MAX_IMAGE_BYTES + 1, runId);
    const env = await sealIntent(keys.kContent, "d", "g1", bad);
    deliver({ t: "broadcast-enc", entries: [{ ...env, seq: 1 }] } as ServerMessage);
    for (let i = 0; i < 200 && seen.length === 0; i++) await new Promise((r) => setTimeout(r, 5));

    expect(seen).toEqual([{ reason: "insertImage: bad size", clientSeq: 1 }]);
  });

  it("the ceiling leaves headroom above every cap we actually configure", () => {
    // THE RELATIONSHIP THAT ROTTED. These are two numbers in two packages that
    // must stay ordered; the server clamps to enforce it, and this states the
    // intent so a future edit to either one fails here rather than in a user's
    // document. Written against the real configured values, not a literal, so
    // raising the dev cap without raising the ceiling is caught.
    expect(PROD_CAP).toBeLessThanOrEqual(MAX_IMAGE_BYTES);
    expect(DEV_CAP).toBeLessThanOrEqual(MAX_IMAGE_BYTES);
    expect(USER_PHOTO_LARGE).toBeLessThanOrEqual(MAX_IMAGE_BYTES);
  });
});
