import { describe, expect, it } from "vitest";
import { CollabConnection } from "../src/connection.js";
import { EncryptedCollabConnection } from "../src/enc-connection.js";

/**
 * SURFACE PARITY between the two connection classes. The react layer treats
 * an EncryptedCollabConnection `as unknown as CollabConnection`, which
 * silences the typechecker for the whole public surface — that cast hid a
 * MISSING `admin()` on the encrypted connection for weeks: the demo's owner
 * controls (read-only / kick / role) were silent TypeErrors on encrypted
 * docs (the default mode), and the symptom was misattributed to a Vite
 * prebundle race. This pin makes the cast honest mechanically: every public
 * method of the plaintext connection must exist on the encrypted one,
 * except a REASONED allowlist of genuinely plaintext-only surface.
 */
describe("connection surface parity (the react cast must be honest)", () => {
  it("every plaintext public method exists on the encrypted connection", () => {
    // Plaintext-only, with reasons — additions here need a reason, not a shrug.
    const PLAINTEXT_ONLY = new Set<string>([
      // Callbacks are a constructor argument on the encrypted connection;
      // no react/demo code calls setCallbacks through the cast.
      "setCallbacks",
      // TS-private helper on the plaintext class — runtime-visible only
      // because TS privacy is compile-time; not part of the public surface.
      "recordActivity",
      // Same shape as recordActivity: a private helper that walks a BATCH of
      // broadcast entries looking for this client's rejections. The encrypted
      // connection needs no equivalent because it ingests one envelope at a
      // time and reports the rejection inline in `ingest` — so the BEHAVIOUR
      // is present on both (pinned in image-size-ceiling.test.ts and
      // media-e2ee-connection), only the helper's shape differs. Listed rather
      // than mirrored, because adding a batch method to the encrypted class
      // that nothing calls would be worse than this line.
      "reportRejections",
      // Media-duty plumbing the plaintext connection's own message handler
      // invokes on itself; the encrypted connection delegates these to its
      // internal MediaClient. No external caller goes through the cast.
      "mediaNeed",
      "mediaHave",
      "heldMediaShas",
    ]);
    const methodsOf = (proto: object): string[] =>
      Object.getOwnPropertyNames(proto).filter(
        (n) => n !== "constructor" && typeof (proto as Record<string, unknown>)[n] === "function" && !n.startsWith("_"),
      );
    const plain = methodsOf(CollabConnection.prototype);
    const enc = new Set(methodsOf(EncryptedCollabConnection.prototype));
    const missing = plain.filter((m) => !enc.has(m) && !PLAINTEXT_ONLY.has(m));
    expect(missing, `encrypted connection is missing public methods the react cast promises: ${missing.join(", ")}`).toEqual([]);
  });
});
