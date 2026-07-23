import { describe, expect, it } from "vitest";
import { makeDocId, PartyPool, RateLimiter, intentAllowedInDemo, DEMO_INTENT_ALLOWLIST } from "../src/demo.js";

describe("makeDocId", () => {
  it("produces a 128-bit hex id with the d_ prefix", () => {
    const id = makeDocId((n) => new Uint8Array(n).fill(0xab));
    expect(id).toBe("d_" + "ab".repeat(16));
    expect(id.length).toBe(2 + 32);
  });
  it("differs when the random source differs (unguessable)", () => {
    expect(makeDocId((n) => new Uint8Array(n).fill(0x01))).not.toBe(
      makeDocId((n) => new Uint8Array(n).fill(0x02)),
    );
  });
});

describe("PartyPool", () => {
  it("round-robins visitors over a fixed set of docs", () => {
    const pool = new PartyPool(3);
    const got = [pool.assign(), pool.assign(), pool.assign(), pool.assign()];
    expect(got).toEqual(["party_0", "party_1", "party_2", "party_0"]);
    expect(pool.docIds).toHaveLength(3);
  });
  it("rejects an empty pool", () => {
    expect(() => new PartyPool(0)).toThrow();
  });
});

describe("demo intent allowlist (doc 11)", () => {
  it("party mode is text-only (structurally XSS-free: no authored-URL/media)", () => {
    expect(intentAllowedInDemo("party", "insertText")).toBe(true);
    expect(intentAllowedInDemo("party", "formatRun")).toBe(false);
    expect(intentAllowedInDemo("party", "setListType")).toBe(false);
    // No authored-URL intents in either mode until the scheme-allowlist gate.
    expect(DEMO_INTENT_ALLOWLIST.magicLink).not.toContain("insertHyperlink");
    expect(DEMO_INTENT_ALLOWLIST.party).not.toContain("insertHyperlink");
  });
  it("magic-link mode allows the full implemented feature set", () => {
    for (const k of ["insertText", "deleteText", "splitParagraph", "formatRun", "formatParagraph", "setListType"]) {
      expect(intentAllowedInDemo("magicLink", k)).toBe(true);
    }
  });
});

describe("RateLimiter (abuse limits, doc 11)", () => {
  it("allows up to capacity then blocks until refill", () => {
    const rl = new RateLimiter(2, 0); // 2 tokens, no refill
    expect(rl.allow("ip1", 0)).toBe(true);
    expect(rl.allow("ip1", 0)).toBe(true);
    expect(rl.allow("ip1", 0)).toBe(false); // exhausted
    // A different key has its own bucket.
    expect(rl.allow("ip2", 0)).toBe(true);
  });
  it("refills over time", () => {
    const rl = new RateLimiter(1, 0.001); // 1 token/1000ms
    expect(rl.allow("k", 0)).toBe(true);
    expect(rl.allow("k", 500)).toBe(false); // half a token
    expect(rl.allow("k", 1000)).toBe(true); // refilled
  });
});
