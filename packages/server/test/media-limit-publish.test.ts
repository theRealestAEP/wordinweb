import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { CollabHub, type Connection, type DocProvider } from "../src/hub.js";
import { envLimits, normalizeLimits, DEFAULT_LIMITS } from "../src/limits.js";
import { PROTOCOL_VERSION, type ServerMessage } from "@wordinweb/collab/server";

/**
 * PUBLISHING THE MEDIA SIZE LIMIT (#25, server half).
 *
 * The client checks a file's size BEFORE uploading, which means the number has
 * to flow FORWARD — server to client, once per session — instead of backward
 * out of a 413. A refusal has to climb through every layer between the relay
 * and the file picker, each with its own idea of what failure looks like; a
 * value published at join time has no such path to get lost on.
 *
 * THE PIN THAT MATTERS is the env-driven one. Asserting the welcome carries
 * the DEFAULT would pass just as happily against a hardcoded constant — the
 * exact class of "test agrees with itself" bug this codebase keeps catching —
 * so the limit is set to a number nobody would ever pick as a default and the
 * assertion is that precise value.
 */

const ODD_LIMIT = 1234567;

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
  welcome(): Extract<ServerMessage, { t: "welcome" | "welcome-enc" }> {
    const w = this.received.find((m) => m.t === "welcome" || m.t === "welcome-enc");
    if (!w) throw new Error("no welcome received");
    return w as Extract<ServerMessage, { t: "welcome" | "welcome-enc" }>;
  }
}

/** Run a body with an env var set, restoring the environment afterwards. */
function withEnv<T>(name: string, value: string, body: () => T): T {
  const saved = { ...process.env };
  try {
    process.env[name] = value;
    return body();
  } finally {
    process.env = saved;
  }
}

describe("welcome publishes the configured media blob limit", () => {
  it("carries the value WW_MEDIA_MAX_BLOB_BYTES actually set — not the default", async () => {
    // The whole chain: env → envLimits → hub → welcome. A hub publishing a
    // hardcoded constant, or the default, or the wrong limit entirely, fails
    // here; a pin written against the default would not have noticed any of
    // those.
    const limits = withEnv("WW_MEDIA_MAX_BLOB_BYTES", String(ODD_LIMIT), () => envLimits());
    expect(limits.media.maxBlobBytes).toBe(ODD_LIMIT);

    const hub = new CollabHub(provider, undefined, undefined, undefined, undefined, undefined, limits);
    const c = new FakeConn("c1");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "alice", sinceSeq: 0 });

    const w = c.welcome();
    expect(w.t).toBe("welcome");
    expect(w.mediaMaxBlobBytes).toBe(ODD_LIMIT);
    // Stated explicitly, because this is the assertion that has teeth: the
    // published number must not be the compiled-in default.
    expect(w.mediaMaxBlobBytes).not.toBe(DEFAULT_LIMITS.media.maxBlobBytes);
  });

  it("publishes it identically in ENCRYPTED rooms", async () => {
    // A blind sequencer can state its own configuration in the clear: this is
    // server config, not room content. Publishing it only in plaintext rooms
    // would leave the pre-check dead in the demo's DEFAULT mode.
    const limits = withEnv("WW_MEDIA_MAX_BLOB_BYTES", String(ODD_LIMIT), () => envLimits());
    const hub = new CollabHub(null, undefined, undefined, undefined, undefined, undefined, limits);
    hub.seedEncrypted("d", "g1", { seq: 0, iv: "aXY=", ciphertext: "Y2lwaGVy" });
    const c = new FakeConn("c1");
    await hub.handle(c, {
      t: "hello",
      protocolVersion: PROTOCOL_VERSION,
      docId: "d",
      clientId: "alice",
      sinceSeq: 0,
      engineVersion: "e4",
    });

    const w = c.welcome();
    expect(w.t).toBe("welcome-enc");
    expect(w.mediaMaxBlobBytes).toBe(ODD_LIMIT);
  });

  it("tracks a raised limit, so the published number and the enforced one cannot drift", async () => {
    // The failure this guards: an operator raises the cap, the relay accepts
    // bigger files, and the client keeps refusing them locally against a stale
    // number — an upload path that is broken in the direction nobody tests.
    const raised = normalizeLimits({ media: { maxBlobBytes: 64 * 1024 * 1024 } });
    const hub = new CollabHub(provider, undefined, undefined, undefined, undefined, undefined, raised);
    const c = new FakeConn("c1");
    await hub.handle(c, { t: "hello", protocolVersion: PROTOCOL_VERSION, docId: "d", clientId: "alice", sinceSeq: 0 });
    expect(c.welcome().mediaMaxBlobBytes).toBe(64 * 1024 * 1024);

    // …and the relay genuinely enforces the same number, which is what makes
    // the published value trustworthy rather than decorative.
    const justOver = new Uint8Array(64 * 1024 * 1024 + 1);
    expect(await hub.mediaUpload("d", "0".repeat(64), justOver)).toBe(413);
  });
});
