import { describe, expect, it } from "vitest";
import {
  deriveEpochKeys,
  mintDocKey,
  sealIntent,
  openIntent,
  sealPresence,
  openPresence,
  sealCheckpoint,
  openCheckpoint,
  type EpochKeys,
} from "../src/e2ee.js";
import type { Intent } from "../src/intents.js";

/**
 * Sealed-body padding (traffic-analysis hardening). AES-GCM is length-
 * preserving, so before padding the envelope size tracked the edit exactly:
 * a keystroke, a backspace, an Enter, a suggesting keystroke and a paste
 * each had a distinct sealed size — a paste-length oracle (±3 chars) plus a
 * "suggesting mode" signature for the blind server. These tests pin the
 * property the padding buys: every small edit seals to ONE size class.
 *
 * Size arithmetic used throughout: ciphertext = plaintext + 16 (GCM tag);
 * base64 chars = 4 * ceil(bytes / 3). Intent rung 1 is a 384-byte plaintext
 * (4-byte prefix + JSON + zero fill) → 400 ct bytes → 536 base64 chars.
 */

const b64len = (bytes: number) => 4 * Math.ceil(bytes / 3);
const RUNG1_B64 = b64len(384 + 16); // 536
const RUNG2_B64 = b64len(1024 + 16); // 1388
const RUNG3_B64 = b64len(4096 + 16); // 5484
const STEP2_B64 = b64len(8192 + 16); // 10944

const DOC = "d-abc";
const GEN = "g-8f3a2b1c";
let keysPromise: Promise<EpochKeys> | null = null;
const getKeys = () => (keysPromise ??= deriveEpochKeys(mintDocKey(), GEN));

const base = { clientId: "c-4f9e8d7a6b5c", clientSeq: 42, base: 137 };
const at = { blockId: 12, runId: 34, offset: 56 };
const suggest = { author: "Alex Pickett", date: "2026-07-29T12:34:56Z" };

/** A census of realistic single-edit operations — every row of the leak
 * measurement table plus the other common editing intents. */
const smallIntents: Record<string, Intent> = {
  "backspace (deleteText)": { kind: "deleteText", ...base, blockId: 12, runId: 34, start: 55, end: 56 } as Intent,
  "bold toggle (formatRun)": { kind: "formatRun", ...base, blockId: 12, runId: 34, patch: { bold: true } } as Intent,
  "keystroke (insertText, 1 char)": { kind: "insertText", ...base, at, text: "a" } as Intent,
  "enter (splitParagraph)": { kind: "splitParagraph", ...base, at, newBlockId: 1201, newRunId: 1202 } as Intent,
  "keystroke while suggesting": { kind: "insertText", ...base, at, text: "a", suggest } as Intent,
  "sentence paste (92 chars)": { kind: "insertText", ...base, at, text: "x".repeat(92) } as Intent,
  "table row insert": { kind: "tableOp", ...base, cellParagraphId: 77, op: "rowBelow", nodeIds: [901, 902, 903, 904] } as Intent,
  "insertImage": {
    kind: "insertImage",
    ...base,
    runId: 34,
    blobSha: "a".repeat(64),
    bytesLen: 48213,
    ext: "png",
    iv: "AAAAAAAAAAAAAAAA",
    widthPx: 640,
    heightPx: 480,
    nodeIds: [911, 912],
  } as Intent,
  "backspace at paragraph start (mergeParagraph)": { kind: "mergeParagraph", ...base, blockId: 12 } as Intent,
};

describe("padding — intents", () => {
  it("HEADLINE: every small edit seals to the same size (no keystroke/paste/suggest signature)", async () => {
    const keys = await getKeys();
    const sizes = new Map<string, number>();
    for (const [name, intent] of Object.entries(smallIntents)) {
      const env = await sealIntent(keys.kContent, DOC, GEN, intent);
      sizes.set(name, env.ciphertext.length);
      // Round-trip too: equal sizes bought by corrupting the body would be
      // worthless.
      expect(await openIntent(keys.kContent, DOC, GEN, env)).toEqual(intent);
    }
    // One size class across ALL of them — this is the property the padding
    // buys. Pinned to the exact rung-1 value so shrinking the rung (or
    // removing the fill) fails here rather than shifting silently.
    for (const [name, size] of sizes) {
      expect(size, `sealed size of "${name}"`).toBe(RUNG1_B64);
    }
  });

  /** Build an insertText whose serialized JSON is exactly `jsonLen` bytes. */
  const intentOfJsonLength = (jsonLen: number): Intent => {
    const probe = { kind: "insertText", ...base, at, text: "" } as Intent;
    const fixed = new TextEncoder().encode(JSON.stringify(probe)).length;
    expect(jsonLen).toBeGreaterThanOrEqual(fixed);
    const it2 = { kind: "insertText", ...base, at, text: "a".repeat(jsonLen - fixed) } as Intent;
    expect(new TextEncoder().encode(JSON.stringify(it2)).length).toBe(jsonLen);
    return it2;
  };

  it("round-trips at the bucket edges (padded length 383 / 384 / 385)", async () => {
    const keys = await getKeys();
    // padded length = 4 (prefix) + jsonLen; the rung boundary is at 384.
    for (const [jsonLen, want] of [
      [379, RUNG1_B64], // padded 383: inside rung 1
      [380, RUNG1_B64], // padded 384: exactly rung 1
      [381, RUNG2_B64], // padded 385: first byte of rung 2
      [1020, RUNG2_B64], // padded 1024: exactly rung 2
      [1021, RUNG3_B64], // padded 1025: rung 3
      [4092, RUNG3_B64], // padded 4096: exactly rung 3
      [4093, STEP2_B64], // padded 4097: next 4 KiB multiple (8192)
    ] as const) {
      const intent = intentOfJsonLength(jsonLen);
      const env = await sealIntent(keys.kContent, DOC, GEN, intent);
      expect(env.ciphertext.length, `sealed size at jsonLen=${jsonLen}`).toBe(want);
      // Exact round-trip proves the length prefix slices precisely (an
      // off-by-one either truncates the text or leaks a fill byte in).
      expect(await openIntent(keys.kContent, DOC, GEN, env)).toEqual(intent);
    }
  });

  it("near-cap paste: bucket rounding clamps to the hub cap instead of pushing past it", async () => {
    const keys = await getKeys();
    // Hub refuses ciphertext > 256 KiB base64 chars ⇒ 196 608 raw ct bytes
    // ⇒ 196 592 plaintext bytes. A 196 000-byte JSON rounds to the next
    // 4 KiB multiple (196 608) which is PAST the cap — the guard must clamp
    // to exactly the cap, never refuse-by-padding.
    const intent = intentOfJsonLength(196_000);
    const env = await sealIntent(keys.kContent, DOC, GEN, intent);
    expect(env.ciphertext.length).toBe(256 * 1024); // clamped: exactly at the cap
    expect(env.ciphertext.length).toBeLessThanOrEqual(256 * 1024); // hub admits it
    expect(await openIntent(keys.kContent, DOC, GEN, env)).toEqual(intent);
  });

  it("an unpadded (pre-e6) body is rejected by openIntent — the e5/e6 fork the ENGINE_VERSION bump fences", async () => {
    const keys = await getKeys();
    const intent = smallIntents["keystroke (insertText, 1 char)"];
    // Seal the raw JSON with the exact same AAD but NO wrapper, the way an
    // e5 client did.
    const iv = new Uint8Array(12);
    crypto.getRandomValues(iv);
    const aad = new TextEncoder().encode(`in:${DOC}:${GEN}:${intent.clientId}:${intent.clientSeq}:${intent.base}`);
    const ct = await crypto.subtle.encrypt(
      { name: "AES-GCM", iv, additionalData: aad },
      keys.kContent,
      new TextEncoder().encode(JSON.stringify(intent)),
    );
    const env = {
      clientId: intent.clientId,
      clientSeq: intent.clientSeq,
      base: intent.base,
      iv: btoa(String.fromCharCode(...iv)),
      ciphertext: btoa(String.fromCharCode(...new Uint8Array(ct))),
    };
    // `{"ki…` read as a 4-byte big-endian length is astronomically large —
    // the prefix check throws, a deterministic no-op (never a mis-parse).
    await expect(openIntent(keys.kContent, DOC, GEN, env)).rejects.toThrow();
  });
});

describe("padding — presence", () => {
  const caret = { anchor: at };
  const range = (i: number) => ({ blockId: 12 + i, runId: 34 + i, start: 0, end: 40 + i });
  const sel = (n: number) => ({ anchor: at, focus: { blockId: 18, runId: 40, offset: 3 }, ranges: Array.from({ length: n }, (_, i) => range(i)) });

  it("caret-only pads to the 128 bucket; selections of any size pad to the 1024 bucket", async () => {
    const keys = await getKeys();
    const clientId = base.clientId;
    const sealedCaret = await sealPresence(keys.kPresence, DOC, GEN, clientId, caret);
    expect(sealedCaret.ciphertext.length).toBe(b64len(128 + 16)); // 192
    // 1, 5 and 12 ranges are indistinguishable — selection SIZE is hidden.
    // (The caret-versus-selection bit intentionally stays visible.)
    for (const n of [1, 5, 12]) {
      const sealed = await sealPresence(keys.kPresence, DOC, GEN, clientId, sel(n));
      expect(sealed.ciphertext.length, `selection with ${n} ranges`).toBe(b64len(1024 + 16)); // 1388
      expect(await openPresence(keys.kPresence, DOC, GEN, clientId, sealed)).toEqual(sel(n));
    }
    expect(await openPresence(keys.kPresence, DOC, GEN, clientId, sealedCaret)).toEqual(caret);
  });

  it("a clamp-limit 64-range selection still round-trips (spills to the next KiB, never truncates)", async () => {
    const keys = await getKeys();
    const sealed = await sealPresence(keys.kPresence, DOC, GEN, base.clientId, sel(64));
    expect(sealed.ciphertext.length % 4).toBe(0);
    expect(await openPresence(keys.kPresence, DOC, GEN, base.clientId, sealed)).toEqual(sel(64));
  });
});

describe("padding — checkpoints", () => {
  it("pads to 64 KiB multiples and round-trips", async () => {
    const keys = await getKeys();
    const body = { docx: "A".repeat(2000), sidecar: { ids: [1, 2, 3] }, docHash: "h".repeat(64) };
    const sealed = await sealCheckpoint(keys.kContent, DOC, GEN, 7, body);
    expect(sealed.ciphertext.length).toBe(b64len(64 * 1024 + 16)); // 87 404
    expect(await openCheckpoint(keys.kContent, DOC, GEN, 7, sealed)).toEqual(body);
    // A body past one rung lands on the next multiple, not a tracking size.
    const big = { ...body, docx: "A".repeat(70 * 1024) };
    const sealedBig = await sealCheckpoint(keys.kContent, DOC, GEN, 8, big);
    expect(sealedBig.ciphertext.length).toBe(b64len(128 * 1024 + 16));
    expect(await openCheckpoint(keys.kContent, DOC, GEN, 8, sealedBig)).toEqual(big);
  });
});
