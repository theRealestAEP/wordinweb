import { describe, expect, it } from "vitest";
import {
  mintDocKey,
  docKeyFromFragment,
  deriveEpochKeys,
  stretchShareCode,
  sealIntent,
  openIntent,
  sealCheckpoint,
  openCheckpoint,
} from "../src/e2ee.js";
import type { Intent } from "../src/intents.js";

/** Crypto unit tests (plan doc 13 §8): round-trips, AAD binding (epoch/
 * base/clientSeq/doc), IV uniqueness, fragment parsing, share-code mixing. */

const intent: Intent = {
  kind: "insertText",
  clientId: "alice",
  clientSeq: 7,
  base: 3,
  at: { blockId: 1, runId: 2, offset: 0 },
  text: "secret",
} as Intent;

describe("e2ee — keys and fragments", () => {
  it("mints url-safe 256-bit keys and parses them back from fragments", () => {
    const k = mintDocKey();
    expect(k).toMatch(/^[A-Za-z0-9_-]{43}$/); // 32 bytes base64url, unpadded
    expect(docKeyFromFragment(`#k=${k}`)).toBe(k);
    expect(docKeyFromFragment(`#view=print&k=${k}`)).toBe(k);
    expect(docKeyFromFragment("#nothing")).toBeNull();
    expect(docKeyFromFragment("")).toBeNull();
  });

  it("epoch derivation: same inputs agree; different epochs derive different keys", async () => {
    const kDoc = mintDocKey();
    const a = await deriveEpochKeys(kDoc, "g1");
    const env = await sealIntent(a.kContent, "d", "g1", intent);
    // A second derivation from the same link + epoch opens it (derivation
    // replaces distribution — doc 13 §1).
    const b = await deriveEpochKeys(kDoc, "g1");
    expect(await openIntent(b.kContent, "d", "g1", env)).toEqual(intent);
    // A different epoch's keys cannot (kills keyless cross-epoch replay, F5a).
    const c = await deriveEpochKeys(kDoc, "g2");
    await expect(openIntent(c.kContent, "d", "g2", env)).rejects.toThrow();
  });

  it("share code mixes into derivation: without it (or with the wrong one) nothing opens", async () => {
    const kDoc = mintDocKey();
    const code = await stretchShareCode("123456", "d");
    const withCode = await deriveEpochKeys(kDoc, "g1", code);
    const env = await sealIntent(withCode.kContent, "d", "g1", intent);
    // Link alone (no code): undecryptable — the doc-13 §7 guarantee that a
    // leaked link without the out-of-band code is not merely bounced but blind.
    const linkOnly = await deriveEpochKeys(kDoc, "g1");
    await expect(openIntent(linkOnly.kContent, "d", "g1", env)).rejects.toThrow();
    const wrongCode = await deriveEpochKeys(kDoc, "g1", await stretchShareCode("654321", "d"));
    await expect(openIntent(wrongCode.kContent, "d", "g1", env)).rejects.toThrow();
    // Right code independently derived: opens.
    const again = await deriveEpochKeys(kDoc, "g1", await stretchShareCode("123456", "d"));
    expect(await openIntent(again.kContent, "d", "g1", env)).toEqual(intent);
  }, 30000);
});

describe("e2ee — envelope AAD binding (round-4 F5)", () => {
  it("round-trips an intent and refuses every tampered bookkeeping field", async () => {
    const keys = await deriveEpochKeys(mintDocKey(), "g1");
    const env = await sealIntent(keys.kContent, "d", "g1", intent);
    expect(await openIntent(keys.kContent, "d", "g1", env)).toEqual(intent);

    // A keyless server altering `base` must be caught — base is a transform
    // input; shifting it would move WHERE the edit applies (F5b).
    await expect(openIntent(keys.kContent, "d", "g1", { ...env, base: env.base + 1 })).rejects.toThrow();
    // clientSeq splice (dedup poisoning by repositioning a blob).
    await expect(openIntent(keys.kContent, "d", "g1", { ...env, clientSeq: 99 })).rejects.toThrow();
    // Attribution swap.
    await expect(openIntent(keys.kContent, "d", "g1", { ...env, clientId: "bob" })).rejects.toThrow();
    // Cross-document replay.
    await expect(openIntent(keys.kContent, "other-doc", "g1", env)).rejects.toThrow();
    // Ciphertext bit-flip.
    const flipped = env.ciphertext.slice(0, 5) + (env.ciphertext[5] === "A" ? "B" : "A") + env.ciphertext.slice(6);
    await expect(openIntent(keys.kContent, "d", "g1", { ...env, ciphertext: flipped })).rejects.toThrow();
  });

  it("fresh IV per envelope (same intent sealed twice differs everywhere)", async () => {
    const keys = await deriveEpochKeys(mintDocKey(), "g1");
    const a = await sealIntent(keys.kContent, "d", "g1", intent);
    const b = await sealIntent(keys.kContent, "d", "g1", intent);
    expect(a.iv).not.toBe(b.iv);
    expect(a.ciphertext).not.toBe(b.ciphertext);
    // Both open to the same plaintext regardless.
    expect(await openIntent(keys.kContent, "d", "g1", a)).toEqual(intent);
    expect(await openIntent(keys.kContent, "d", "g1", b)).toEqual(intent);
  });
});

describe("e2ee — checkpoints (doc 13 §3)", () => {
  it("round-trips and refuses seq/epoch repositioning (rollback defense, F5c)", async () => {
    const keys = await deriveEpochKeys(mintDocKey(), "g1");
    const body = { docx: "UEsDBA==", sidecar: { ids: [1, 2, 3] }, docHash: "abc123" };
    const sealed = await sealCheckpoint(keys.kContent, "d", "g1", 42, body);
    expect(await openCheckpoint(keys.kContent, "d", "g1", 42, sealed)).toEqual(body);
    // Serving an OLD checkpoint at a different seq is refused by AAD.
    await expect(openCheckpoint(keys.kContent, "d", "g1", 41, sealed)).rejects.toThrow();
    await expect(openCheckpoint(keys.kContent, "d", "g2", 42, sealed)).rejects.toThrow();
  });
});
